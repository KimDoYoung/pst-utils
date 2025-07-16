import pypff
import sys
import os
import argparse
from pathlib import Path
from datetime import datetime, timezone, timedelta
from db_actions import create_db_tables, save_email_data_to_db, create_db_path
from helper import  get_message_class, get_sender_from_address,  byte_decode,extract_attachments,convert_to_kst, safe_get_attachment_count, get_property_from_record_sets, get_recipients_info, determine_message_kind
from logger import get_logger
from config import settings


logger = get_logger()

# MAPI 속성 상수들
# PR_MESSAGE_CLASS = 0x001A  # 26
# # PR_SENDER_EMAIL_ADDRESS = 0x0C1F  # 3103
# # PR_FROM_EMAIL_ADDRESS = 0x0051
# PR_SENDER_NAME = 0x0C1A  # 3098
# PR_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065  # 101
# PR_SENT_REPRESENTING_NAME = 0x0042  # 66
# PR_DISPLAY_TO = 0x0E04  # 3588
# PR_DISPLAY_CC = 0x0E03  # 3587
# # PR_RECEIVED_BY_EMAIL_ADDRESS = 0x0076  # 118
# # PR_RECEIVED_BY_NAME = 0x0040  # 64
# PR_MESSAGE_FLAGS   = 0x0E07  # 3591
# MSGFLAG_FROMME     = 0x00000040

# def get_message_class(msg: pypff.message) -> str:
#     """message_class 추출"""
#     try:
#         if hasattr(msg, 'message_class') and msg.message_class:
#             return msg.message_class
#     except Exception:
#         pass
    
#     try:
#         if hasattr(msg, 'get_message_class'):
#             mc = msg.get_message_class()
#             if mc:
#                 return mc
#     except Exception:
#         pass
    
#     try:
#         if hasattr(msg, 'record_sets') and msg.record_sets:
#             for record_set in msg.record_sets:
#                 if hasattr(record_set, 'entries'):
#                     for entry in record_set.entries:
#                         if (hasattr(entry, 'entry_type') and 
#                             entry.entry_type == PR_MESSAGE_CLASS):
#                             if hasattr(entry, 'data_as_string'):
#                                 try:
#                                     return entry.data_as_string
#                                 except Exception:
#                                     pass
#     except Exception:
#         pass
    
#     return ""





def extract_email_content(msg) -> str:
    """이메일 본문 추출 — plain text > html > rtf 순서, 안전한 디코딩 포함"""
    # 1. plain text
    if hasattr(msg, "plain_text_body") and msg.plain_text_body:
        body = msg.plain_text_body
        if isinstance(body, bytes):
            return byte_decode(body)
        elif str(body).strip():
            return str(body)

    # 2. HTML
    if hasattr(msg, "html_body") and msg.html_body:
        body = msg.html_body
        if isinstance(body, bytes):
            return byte_decode(body)
        elif str(body).strip():
            return str(body)

    # 3. RTF — RTF는 \uNNNN 제어어로 유니코드가 들어있으므로
    #    `striprtf` 등으로 평문 변환 후 디코딩하는 편이 좋다.
    if hasattr(msg, "rtf_body") and msg.rtf_body:
        try:
            from striprtf import rtf_to_text
            body = msg.rtf_body
            text = rtf_to_text(body if isinstance(body, str) else byte_decode(body))
            if text.strip():
                return text
        except Exception:
            pass

    return ""  # 아무것도 없을 때

def build_folder_path(folder: pypff.folder, path_list: list = None) -> str:
    """폴더의 전체 경로 구성"""
    if path_list is None:
        path_list = []
    
    try:
        folder_name = getattr(folder, 'name', 'Unknown')
        if folder_name and folder_name != 'Root':
            path_list.insert(0, folder_name)
        
        # 상위 폴더가 있으면 재귀적으로 처리
        if hasattr(folder, 'parent') and folder.parent:
            return build_folder_path(folder.parent, path_list)
    except Exception:
        pass
    
    return "/".join(path_list) if path_list else "Root"


def extract_email_data(msg: pypff.message, folder_path: str) -> dict:
    """
    pypff.message에서 fund_mail 테이블에 맞는 데이터를 추출합니다.
    """
    email_id = str(getattr(msg, 'identifier', ''))
    # 기본 정보 추출
    email_data = {
        'email_id': email_id,
        'subject': '',
        'sender_address': '',
        'sender_name': '',
        'from_address': '',
        'from_name': '',
        'to_recipients': '',
        'cc_recipients': '',
        'email_time': '',
        'kst_time': '',
        'content': '',
        # 새로 추가된 필드들
        'msg_kind': '',
        'folder_path': folder_path,
        'attachments': [],
    }
    email_data['note'] = None
    # 제목
    try:
        email_data['subject'] = getattr(msg, 'subject', '') or ''
    except Exception:
        pass
    
    # 발신자 정보
    try:
        email_data['sender_name'] = getattr(msg, 'sender_name', '') or ''
        email_data['from_name'] = email_data['sender_name']
    except Exception:
        pass
    
    
    sender_address, from_address = get_sender_from_address(msg)
    email_data['sender_address'] = sender_address
    email_data['from_address'] = from_address

    # 수신자 정보
    to_recipients, cc_recipients = get_recipients_info(msg)
    email_data['to_recipients'] = to_recipients
    email_data['cc_recipients'] = cc_recipients
    
    # 시간 정보
    try:
        delivery_time = getattr(msg, 'delivery_time', None)
        if delivery_time:
            email_data['email_time'] = str(delivery_time)
            email_data['kst_time'] = convert_to_kst(delivery_time)
    except Exception:
        pass
    
    # 이메일 본문
    email_data['content'] = extract_email_content(msg)
    
    # 메시지 종류 판단
    email_data['msg_kind'] = determine_message_kind(msg, folder_path)
    
    # 첨부파일 추출
    ymd = email_data['kst_time'][:10] if email_data['kst_time'] else ''
    # attach_dir = f"/home/kdy987/data/{ymd}"
    attach_dir = settings.ATTATCH_BASE_DIR / ymd
    # if not os.path.exists(attach_dir):
    #     os.makedirs(attach_dir, exist_ok=True)
    attach_count = safe_get_attachment_count(msg)
    if attach_count < 0:
        logger.error(f"{email_data['email_id']} : 첨부파일 갯수를 가져오는 데 실패했습니다.")
        email_data['note'] = "첨부파일 갯수를 가져오는 데 실패했습니다. (PST 손상 by ChatGPT)"
    else:
        logger.info(f"{email_data['email_id']}, 첨부파일 수: {attach_count}")        
        email_data['attach_files'] = extract_attachments(msg, attach_dir, email_id=email_id, attach_count=attach_count)
    logger.info(f"메일데이터 추출 : {email_data['email_id']}  {email_data['subject'][:50]}... 폴더: {folder_path}")
    return email_data

def walk_and_extract_emails(db_path:str, folder: pypff.folder, folder_path: str = "", depth: int = 0):
    """폴더를 순회하며 이메일 데이터를 추출"""
    
    # 현재 폴더명 추가
    current_folder_name = getattr(folder, 'name', 'Unknown')
    if folder_path:
        current_path = f"{folder_path}/{current_folder_name}"
    else:
        current_path = current_folder_name if current_folder_name != 'Root' else ""
    
    try:
        if hasattr(folder, 'sub_messages'):
            for msg in folder.sub_messages:
                
                try:
                    msg_class = get_message_class(msg)
                    if msg_class.upper().startswith("IPM.NOTE"):
                        email_data = extract_email_data(msg, current_path)
                        save_email_data_to_db([email_data], db_path)                    
                        # 진행상황 출력
                        logger.info(f"{'  ' * depth} [{email_data['msg_kind']}] {email_data['kst_time']} {email_data['subject'][:50]}...")
                        
                except Exception as e:
                    logger.warning(f"Warning: Error processing message: {e}")
                    continue
        
        # 하위 폴더 처리
        if hasattr(folder, 'sub_folders'):
            for sub_folder in folder.sub_folders:
            
                try:
                    sub_folder_name = getattr(sub_folder, 'name', 'Unknown')
                    logger.info(f"{'  ' * depth}📁 {sub_folder_name}")
                    
                    walk_and_extract_emails(db_path=db_path, folder=sub_folder, folder_path = current_path, depth =  depth + 1)
                    
                except Exception as e:
                    logger.warning(f"Warning: Error processing subfolder: {e}")
                    continue
    
    except Exception as e:
        logger.error(f"Error walking folder: {e}")

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="PST 파일을 읽어 메일/첨부를 추출합니다."
    )

    # 위치(필수) 인자 ────────────────────────────────
    parser.add_argument(
        "pst_path",
        type=Path,
        help="분석할 PST 파일 경로"
    )
    return parser.parse_args()

def main() -> None:
    args = parse_args()

    # pst_path = "/mnt/c/tmp/2021.pst"
    pst_path = args.pst_path

    if not os.path.exists(pst_path):
        logger.error(f"❌ PST file not found: {pst_path}")
        sys.exit(1)

    logger.info("="*60)
    logger.info(f"🔴 PST파일 추출  시작: {pst_path}")
    logger.info("="*60)
    db_path = create_db_path(pst_path)
    create_db_tables(db_path)
    
    try:
        pf = pypff.file()
        pf.open(str(pst_path))
        
        root_folder = pf.get_root_folder()
        if not root_folder:
            logger.error("❌ root folder에 접근 할 수 없습니다")
            sys.exit(1)
        
        # 이메일 추출
        walk_and_extract_emails(db_path=db_path, folder = root_folder)
        
        pf.close()
        
    except Exception as e:
        logger.error(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        logger.info("="*60)
        logger.info(f"🔴 PST파일 추출 종료")
        logger.info("="*60)


# 메인 실행
if __name__ == "__main__":
    main()
