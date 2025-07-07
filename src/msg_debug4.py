import pypff
import sys
import os
from datetime import datetime, timezone, timedelta
import json
from helper import  recipients_from_headers, byte_decode
from helper import extract_attachments

# MAPI 속성 상수들
PR_MESSAGE_CLASS = 0x001A  # 26
PR_SENDER_EMAIL_ADDRESS = 0x0C1F  # 3103
PR_SENDER_NAME = 0x0C1A  # 3098
PR_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065  # 101
PR_SENT_REPRESENTING_NAME = 0x0042  # 66
PR_DISPLAY_TO = 0x0E04  # 3588
PR_DISPLAY_CC = 0x0E03  # 3587
PR_RECEIVED_BY_EMAIL_ADDRESS = 0x0076  # 118
PR_RECEIVED_BY_NAME = 0x0040  # 64
PR_MESSAGE_FLAGS   = 0x0E07  # 3591
MSGFLAG_FROMME     = 0x00000040

def get_message_class(msg: pypff.message) -> str:
    """message_class 추출"""
    try:
        if hasattr(msg, 'message_class') and msg.message_class:
            return msg.message_class
    except Exception:
        pass
    
    try:
        if hasattr(msg, 'get_message_class'):
            mc = msg.get_message_class()
            if mc:
                return mc
    except Exception:
        pass
    
    try:
        if hasattr(msg, 'record_sets') and msg.record_sets:
            for record_set in msg.record_sets:
                if hasattr(record_set, 'entries'):
                    for entry in record_set.entries:
                        if (hasattr(entry, 'entry_type') and 
                            entry.entry_type == PR_MESSAGE_CLASS):
                            if hasattr(entry, 'data_as_string'):
                                try:
                                    return entry.data_as_string
                                except Exception:
                                    pass
    except Exception:
        pass
    
    return ""

def get_property_from_record_sets(msg: pypff.message, property_id: int) -> str:
    """record_sets에서 특정 MAPI 속성 값을 추출"""
    try:
        if hasattr(msg, 'record_sets') and msg.record_sets:
            for record_set in msg.record_sets:
                if hasattr(record_set, 'entries'):
                    for entry in record_set.entries:
                        if (hasattr(entry, 'entry_type') and 
                            entry.entry_type == property_id):
                            if hasattr(entry, 'data_as_string'):
                                try:
                                    return entry.data_as_string
                                except Exception:
                                    pass
                            # 백업: 직접 디코딩
                            if hasattr(entry, 'data') and entry.data:
                                try:
                                    if hasattr(entry, 'value_type'):
                                        if entry.value_type == 31:  # PT_UNICODE
                                            return entry.data.decode('utf-16-le', 'replace').rstrip('\x00')
                                        elif entry.value_type == 30:  # PT_STRING8
                                            return entry.data.decode('cp1252', 'replace').rstrip('\x00')
                                except Exception:
                                    pass
    except Exception:
        pass
    return ""

def determine_message_kind(msg: pypff.message, folder_path: str) -> str:
    # 1) 폴더명으로 빠르게 판별
    if folder_path.lower() in ("sent items", "sent", "보낸 편지함", "outbox"):
        return "sent"

    # 2) MAPI 플래그 확인 (보다 확실)
    flags = get_property_from_record_sets(msg, PR_MESSAGE_FLAGS)
    try:
        flags_int = int(flags) if flags else 0
        if flags_int & MSGFLAG_FROMME:
            return "sent"
    except ValueError:
        pass
    return "receive"

def get_receiver_info(msg: pypff.message) -> tuple:
    """수신자 정보 추출 (이메일 주소, 이름)"""
    receiver_addresses = []
    receiver_names = []
    
    try:
        if hasattr(msg, 'recipients'):
            for recipient in msg.recipients:
                try:
                    # 수신자 타입 확인 (1=TO, 2=CC, 3=BCC)
                    recipient_type = getattr(recipient, 'type', 1)
                    
                    # TO 수신자만 처리
                    if recipient_type == 1:
                        email = getattr(recipient, 'email_address', '') or ''
                        name = getattr(recipient, 'name', '') or ''
                        
                        if email:
                            receiver_addresses.append(email)
                        if name:
                            receiver_names.append(name)
                            
                except Exception:
                    continue
    except Exception:
        pass
    
    # record_sets에서도 시도
    if not receiver_addresses:
        received_by_email = get_property_from_record_sets(msg, PR_RECEIVED_BY_EMAIL_ADDRESS)
        received_by_name = get_property_from_record_sets(msg, PR_RECEIVED_BY_NAME)
        
        if received_by_email:
            receiver_addresses.append(received_by_email)
        if received_by_name:
            receiver_names.append(received_by_name)
    
    return "; ".join(receiver_addresses), "; ".join(receiver_names)

def get_recipients_info(msg: pypff.message) -> tuple:
    """수신자와 참조자 정보를 추출"""
    to_recipients = []
    cc_recipients = []
    
    # transport_headers에서 추출
    tr_header = getattr(msg, 'transport_headers', b'')
    to1, cc1 = recipients_from_headers(tr_header)
    
    return to1, cc1

def convert_to_kst(dt: datetime) -> str:
    """UTC datetime을 KST로 변환"""
    if dt is None:
        return ""
    
    try:
        # pypff에서 반환하는 datetime이 UTC라고 가정
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        
        # KST로 변환 (UTC+9)
        kst_dt = dt.astimezone(timezone(timedelta(hours=9)))
        return kst_dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return str(dt)

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
    # 기본 정보 추출
    email_data = {
        'email_id': str(getattr(msg, 'identifier', '')),
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
    
    # 발신자 이메일 주소
    sender_email = get_property_from_record_sets(msg, PR_SENDER_EMAIL_ADDRESS)
    if sender_email:
        email_data['sender_address'] = sender_email
        email_data['from_address'] = sender_email
    
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
    attach_dir = f"/home/kdy987/data/{ymd}"
    # if not os.path.exists(attach_dir):
    #     os.makedirs(attach_dir, exist_ok=True)
    email_data['attachments'] = extract_attachments(msg, attach_dir)
    
    return email_data

def debug_message_properties(msg: pypff.message, max_entries: int = 20) -> None:
    """메시지의 모든 속성을 디버깅용으로 출력"""
    print(f"\n=== Message Properties Debug ===")
    print(f"Identifier: {getattr(msg, 'identifier', 'N/A')}")
    print(f"Subject: {getattr(msg, 'subject', 'N/A')}")
    
    # 기본 속성들
    attrs = [
        'sender_name', 'creation_time', 'delivery_time', 'client_submit_time',
        'conversation_topic', 'transport_headers', 'number_of_entries'
    ]
    
    for attr in attrs:
        try:
            value = getattr(msg, attr, None)
            print(f"{attr}: {value}")
        except Exception as e:
            print(f"{attr}: ERROR - {e}")
    
    # Recipients 정보
    print(f"\n--- Recipients ---")
    try:
        if hasattr(msg, 'recipients'):
            for i, recipient in enumerate(msg.recipients):
                print(f"  Recipient {i}:")
                for r_attr in ['name', 'email_address', 'type']:
                    try:
                        value = getattr(recipient, r_attr, 'N/A')
                        print(f"    {r_attr}: {value}")
                    except Exception as e:
                        print(f"    {r_attr}: ERROR - {e}")
    except Exception as e:
        print(f"Recipients: ERROR - {e}")
    
    # Record sets에서 중요한 속성들 찾기
    print(f"\n--- Important Properties from Record Sets ---")
    important_props = [
        (PR_MESSAGE_CLASS, "MESSAGE_CLASS"),
        (PR_SENDER_EMAIL_ADDRESS, "SENDER_EMAIL"),
        (PR_SENDER_NAME, "SENDER_NAME"),
        (PR_SENT_REPRESENTING_EMAIL_ADDRESS, "FROM_EMAIL"),
        (PR_SENT_REPRESENTING_NAME, "FROM_NAME"),
        (PR_DISPLAY_TO, "TO_RECIPIENTS"),
        (PR_DISPLAY_CC, "CC_RECIPIENTS"),
        (PR_RECEIVED_BY_EMAIL_ADDRESS, "RECEIVED_BY_EMAIL"),
        (PR_RECEIVED_BY_NAME, "RECEIVED_BY_NAME")
    ]
    
    for prop_id, prop_name in important_props:
        value = get_property_from_record_sets(msg, prop_id)
        print(f"{prop_name} ({prop_id}): {value}")

def walk_and_extract_emails(folder: pypff.folder, folder_path: str = "", max_emails: int = 5, depth: int = 0) -> list:
    """폴더를 순회하며 이메일 데이터를 추출"""
    emails = []
    
    # 현재 폴더명 추가
    current_folder_name = getattr(folder, 'name', 'Unknown')
    if folder_path:
        current_path = f"{folder_path}/{current_folder_name}"
    else:
        current_path = current_folder_name if current_folder_name != 'Root' else ""
    
    try:
        if hasattr(folder, 'sub_messages'):
            for msg in folder.sub_messages:
                if len(emails) >= max_emails:
                    break
                
                try:
                    msg_class = get_message_class(msg)
                    if msg_class.upper().startswith("IPM.NOTE"):
                        email_data = extract_email_data(msg, current_path)
                        emails.append(email_data)
                        
                        # 진행상황 출력
                        print(f"{'  ' * depth}📧 [{email_data['msg_kind']}] {email_data['subject'][:50]}...")
                        
                except Exception as e:
                    print(f"Warning: Error processing message: {e}")
                    continue
        
        # 하위 폴더 처리
        if hasattr(folder, 'sub_folders') and len(emails) < max_emails:
            for sub_folder in folder.sub_folders:
                if len(emails) >= max_emails:
                    break
                
                try:
                    sub_folder_name = getattr(sub_folder, 'name', 'Unknown')
                    print(f"{'  ' * depth}📁 {sub_folder_name}")
                    
                    sub_emails = walk_and_extract_emails(sub_folder, current_path, max_emails - len(emails), depth + 1)
                    emails.extend(sub_emails)
                    
                except Exception as e:
                    print(f"Warning: Error processing subfolder: {e}")
                    continue
    
    except Exception as e:
        print(f"Error walking folder: {e}")
    
    return emails

# 메인 실행
if __name__ == "__main__":
    pst_path = "/mnt/c/tmp/2021.pst"
    
    if not os.path.exists(pst_path):
        print(f"❌ PST file not found: {pst_path}")
        sys.exit(1)
    
    try:
        pf = pypff.file()
        pf.open(pst_path)
        
        root_folder = pf.get_root_folder()
        if not root_folder:
            print("❌ Cannot access root folder")
            sys.exit(1)
        
        # 첫 번째 메시지 디버깅
        print("🔍 Debugging first message...")
        for msg in root_folder.sub_messages:
            msg_class = get_message_class(msg)
            if msg_class.upper().startswith("IPM.NOTE"):
                debug_message_properties(msg)
                break
        
        print("\n" + "="*60)
        print("🔍 Extracting email data...")
        
        # 이메일 추출
        emails = walk_and_extract_emails(root_folder, max_emails=100)
        
        print(f"\n✅ Extracted {len(emails)} emails")
        
        # 첫 번째 이메일 데이터 출력
        if emails:
            print("\n📧 First email data:")
            first_email = emails[0]
            for key, value in first_email.items():
                if key == 'content':
                    print(f"{key}: {str(value)[:100]}...")
                else:
                    print(f"{key}: {value}")
        
        # 통계 정보 출력
        sent_count = sum(1 for email in emails if email['msg_kind'] == 'sent')
        receiv_count = sum(1 for email in emails if email['msg_kind'] == 'receiv')
        print(f"\n📊 Statistics:")
        print(f"  Sent emails: {sent_count}")
        print(f"  Received emails: {receiv_count}")
        
        # JSON으로 저장
        output_file = "extracted_emails_enhanced.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(emails, f, ensure_ascii=False, indent=2)
        print(f"\n💾 Data saved to: {output_file}")
        
        pf.close()
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)