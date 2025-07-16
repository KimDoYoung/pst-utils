import pypff
import email,  email.policy
import re
from typing import Sequence

from charset_normalizer import from_bytes
from pathlib import Path
from uuid import uuid4
from datetime import datetime, timezone, timedelta
import time

from logger import get_logger
from config import settings


# MAPI 속성 상수
PR_MESSAGE_CLASS = 0x001A  # 26 in decimal
PR_ATTACH_LONG_FILENAME = 0x3707
PR_ATTACH_FILENAME = 0x3704
PR_DISPLAY_NAME = 0x3001
PR_ATTACH_EXTENSION = 0x3703

# MAPI property IDs (hex without the 8-bit type suffix)
PR_ATTACHMENT_HIDDEN   = 0x7FFE  # PT_BOOLEAN
PR_RENDERING_POSITION  = 0x370B  # PT_LONG
PR_ATTACH_FLAGS        = 0x3714  # PT_LONG
PR_ATTACH_CONTENT_ID   = 0x3712  # PT_TSTRING/PT_BINARY

# bits inside PR_ATTACH_FLAGS
ATT_RENDERED_IN_BODY   = 0x00000004
ATT_MHTML_REF          = 0x00002000

PR_RECEIVED_BY_EMAIL_ADDRESS = 0x0076  # 118
PR_RECEIVED_BY_NAME = 0x0040  # 64

PR_SENDER_EMAIL_ADDRESS = 0x0C1F  # 3103
PR_FROM_EMAIL_ADDRESS = 0x0051

PR_MESSAGE_FLAGS   = 0x0E07  # 3591
MSGFLAG_FROMME     = 0x00000040

logger = get_logger()


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

def get_sender_from_address(msg: pypff.message): 
    # sender_address와 from_address를 추출

    sender_email = get_property_from_record_sets(msg, PR_SENDER_EMAIL_ADDRESS)
    from_email = get_property_from_record_sets(msg, PR_FROM_EMAIL_ADDRESS)

    sender_address = ''
    from_address = ''
    if sender_email:
        if '@' in sender_email:
            sender_address = sender_email
        else:
            # Legacy Exchange 주소인 경우, 실제 SMTP 주소를 추정
            sender_address = resolve_sender_address(msg, sender_email)

    if from_email:
        if '@' in from_email:
            from_address = from_email
        else:
            # Legacy Exchange 주소인 경우, 실제 SMTP 주소를 추정
            from_address = resolve_sender_address(msg, from_email)
    else:
        from_address = sender_email  # fallback

    return sender_address, from_address

def resolve_sender_address(msg: pypff.message, original_dn: str) -> str:
    """
    Legacy Exchange 주소(`/O=...`)인 경우, 실제 SMTP 주소를 추정 시도
    """
    possible_props = [
        0x0C1F,  # PR_SENDER_EMAIL_ADDRESS
        0x0C1E,  # PR_SENDER_EMAIL_ADDRESS_TYPE
        0x5D01,  # PR_SENDER_SMTP_ADDRESS
        0x0076,  # PR_RECEIVED_BY_EMAIL_ADDRESS
        0x0065,  # PR_SENT_REPRESENTING_EMAIL_ADDRESS
        0x3003,  # PR_EMAIL_ADDRESS
        0x39FE,  # PR_SMTP_ADDRESS
    ]
    
    for prop_id in possible_props:
        email = get_property_from_record_sets(msg, prop_id)
        if email and '@' in email:
            return email

    # 2. 메시지 헤더에서 From 주소 추출 시도
    try:
        if hasattr(msg, 'transport_headers') and msg.transport_headers:
            headers = msg.transport_headers
            # From 헤더에서 이메일 추출
            from_match = re.search(r'From:.*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', headers, re.IGNORECASE)
            if from_match:
                address1 = from_match.group(1)
                logger.debug(f"From 헤더에서 추출된 주소: {address1}")
                return address1
                
            # Reply-To 헤더에서 이메일 추출
            reply_to_match = re.search(r'Reply-To:.*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', headers, re.IGNORECASE)
            if reply_to_match:
                address1 = reply_to_match.group(1)
                logger.debug(f"Reply-To 헤더에서 추출된 주소: {address1}")
                return address1
    except Exception:
        pass

    # 3. Exchange DN에서 유용한 정보 추출 
    try:
        # CN= 부분에서 사용자 정보 추출
        cn_match = re.search(r'CN=([^/]+)(?:/|$)', original_dn, re.IGNORECASE)
        if cn_match:
            cn_value = cn_match.group(1)
            
            # GUID 형식이 아닌 경우 (실제 사용자명인 경우)
            if not re.match(r'^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}', cn_value, re.IGNORECASE):
                # 도메인 정보 추출 시도
                domain = extract_domain_from_dn(original_dn)
                if domain:
                    return f"{cn_value}@{domain}"
                else:
                    return cn_value  # 적어도 사용자명은 반환
    except Exception:
        pass

    # 4. 추가 MAPI 속성들 검색
    additional_properties = [
        0x3001,  # PR_DISPLAY_NAME
        0x3002,  # PR_ADDRTYPE
        0x0E08,  # PR_MESSAGE_DELIVERY_TIME
        0x007D,  # PR_TRANSPORT_MESSAGE_HEADERS
    ]
    
    for prop_id in additional_properties:
        prop_value = get_property_from_record_sets(msg, prop_id)
        if prop_value and '@' in prop_value:
            # 문자열에서 이메일 주소 패턴 찾기
            email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', prop_value)
            if email_match:
                return email_match.group(1)
    
    # 5. 모든 방법이 실패한 경우, 원본 DN을 정리해서 반환
    return clean_exchange_dn(original_dn)


def extract_domain_from_dn(dn: str) -> str:
    """Exchange DN에서 도메인 정보 추출"""
    try:
        # OU= 부분에서 도메인 정보 추출 시도
        ou_match = re.search(r'OU=([^/]+)', dn, re.IGNORECASE)
        if ou_match:
            ou_value = ou_match.group(1).lower()
            # 일반적인 도메인 패턴 매칭
            if 'exchange' in ou_value:
                # ExchangeLabs의 경우 outlook.com이나 기본 도메인 추정
                return "outlook.com"
    except Exception:
        pass
    return ""

def clean_exchange_dn(dn: str) -> str:
    """Exchange DN을 읽기 쉬운 형태로 정리"""
    try:
        # CN= 부분만 추출
        cn_match = re.search(r'CN=([^/]+)(?:/|$)', dn, re.IGNORECASE)
        if cn_match:
            cn_value = cn_match.group(1)
            # GUID 형식이 아닌 경우 그대로 반환
            if not re.match(r'^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}', cn_value, re.IGNORECASE):
                return cn_value
            # GUID 형식인 경우 앞 부분만 추출
            else:
                prefix_match = re.search(r'^([^-]+)', cn_value)
                if prefix_match:
                    return f"User_{prefix_match.group(1)[:8]}"
        
        # OU= 부분에서 조직 정보 추출
        ou_match = re.search(r'OU=([^/]+)', dn, re.IGNORECASE)
        if ou_match:
            return f"Exchange_{ou_match.group(1)}"
            
    except Exception:
        pass
    
    return dn  # 모든 처리가 실패하면 원본 반환

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

def get_message_class(msg: pypff.message) -> str:
    """
    pypff 3.12 휠에서 message_class를 올바르게 추출합니다.
    """
    # 방법 1: 직접 속성 접근 (일부 버전에서 작동)
    try:
        if hasattr(msg, 'message_class') and msg.message_class:
            return msg.message_class
    except Exception:
        pass
    
    # 방법 2: get_message_class 메서드 (구 버전)
    try:
        if hasattr(msg, 'get_message_class'):
            mc = msg.get_message_class()
            if mc:
                return mc
    except Exception:
        pass
    
    # 방법 3: record_sets를 통한 접근 (3.12 휠에서 작동하는 방법)
    try:
        if hasattr(msg, 'record_sets') and msg.record_sets:
            # record_sets는 pypff._record_sets 타입
            for record_set in msg.record_sets:
                if hasattr(record_set, 'entries'):
                    # entries는 pypff.record_entries 타입
                    for entry in record_set.entries:
                        # entry_type이 26 (PR_MESSAGE_CLASS)인 것을 찾기
                        if (hasattr(entry, 'entry_type') and 
                            entry.entry_type == PR_MESSAGE_CLASS):
                            
                            # data_as_string() 메서드 사용
                            if hasattr(entry, 'data_as_string'):
                                try:
                                    return entry.data_as_string
                                except Exception:
                                    pass
                            
                            # 백업: 직접 data 디코딩
                            if hasattr(entry, 'data') and entry.data:
                                try:
                                    # value_type이 31이면 UTF-16LE, 30이면 CP1252
                                    if hasattr(entry, 'value_type'):
                                        if entry.value_type == 31:  # PT_UNICODE
                                            return entry.data.decode('utf-16-le', 'replace').rstrip('\x00')
                                        elif entry.value_type == 30:  # PT_STRING8
                                            return entry.data.decode('cp1252', 'replace').rstrip('\x00')
                                except Exception:
                                    pass
    except Exception as e:
        logger.debug(f"Warning: Error in record_sets access: {e}")
    
    return ""

def byte_decode(data: bytes, trial_encs: Sequence[str] = ("utf-8", "cp949", "euc-kr",
                                                      "iso-8859-1", "windows-1252")) -> str:
    """
    bytes → str 인코딩을 최대한 안전하게 수행
    1) meta charset, XML decl 등을 먼저 찾고
    2) 지정된 trial_encs 순서대로 시도
    3) 그래도 안 되면 charset-normalizer에게 맡김
    """
    # 1) <meta charset='...'> 혹은 XML 선언의 encoding 값 추출
    try:
        head = data[:2048].decode("ascii", "ignore")  # ASCII 영역만 미리
        m = re.search(r"charset\s*=\s*['\"]?\s*([-A-Za-z0-9_]+)", head, re.I)
        if m:
            enc = m.group(1).lower()
            return data.decode(enc, "replace")
    except Exception:
        pass

    # 2) 후보 인코딩 순차 시도
    for enc in trial_encs:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue

    # 3) 통계적 추정
    return from_bytes(data).best().output()

def _decode_name(raw_name: str) -> str:
    """=?ks_c_5601-1987?B?...?= 같은 RFC 2047 인코딩을 사람이 읽을 수 있게"""
    try:
        parts = email.header.decode_header(raw_name)
        return "".join(
            fragment.decode(charset or "utf-8", "replace") if isinstance(fragment, bytes)
            else fragment
            for fragment, charset in parts
        )
    except Exception:
        return raw_name            # 디코딩 실패해도 원본 반환

def recipients_from_headers(raw_headers: bytes | str) -> tuple[str, str]:
    """
    transport_headers → (to_recipients, cc_recipients)
      • 결과: "이름 <주소>; 이름2 <주소2>" 문자열 2개
      • BCC는 제외
    """
    if isinstance(raw_headers, bytes):
        try:
            raw_headers = raw_headers.decode("utf-8")
        except UnicodeDecodeError:
            raw_headers = raw_headers.decode("latin-1", "replace")

    try:
        hdr = email.message_from_string(raw_headers, policy=email.policy.default)
        to_pairs = email.utils.getaddresses(hdr.get_all("To",  []))
        cc_pairs = email.utils.getaddresses(hdr.get_all("Cc",  []))
    except Exception as e:
        logger.debug("header-parse error:", e)
        to_pairs, cc_pairs = [], []

    def fmt(pairs):
        out = []
        for name, addr in pairs:
            name = _decode_name(name).strip()
            text = f"{name} <{addr}>" if name else addr
            out.append(text)
        # 중복 제거 후 정렬
        return "; ".join(sorted(set(out)))

    return fmt(to_pairs), fmt(cc_pairs)

def list_pypff_attrs(obj):
    # IDE가 못 잡는 속성까지 강제로 훑어보기
    for name in (
        "identifier", "size", "filename", "long_filename",
        "mime_type", "attach_method", "read_buffer",
        "has_sub_message", "get_sub_message",
    ):
        try:
            val = getattr(obj, name)
            logger.debug(f"{name:<15} ➜ {val!r}")
        except AttributeError:
            logger.debug(f"{name:<15} ✗  (없는 속성)")

INVALID = re.compile(r'[<>:"/\\|?*]')        # 윈도우·리눅스 공통 금지문자

def sanitize_filename(name: str) -> str:
    """OS 금지 문자를 제거하고 비어 있으면 UUID를 반환"""
    name = INVALID.sub('_', name).strip().rstrip('.')
    return name or f'unknown_{uuid4().hex[:8]}'

# def extract_attachments(msg: pypff.message, base_dir: Path) -> list[dict]:
#     results = []
#     total = msg.number_of_attachments

#     for i in range(total):
#         try:
#             at = msg.get_attachment(i)
#         except Exception as e:
#             logger.debug(f"[{i}] attachment 객체 접근 실패 ➜ {e}")
#             continue

#         # ── 1) 사이즈 필수 ─────────────────────────
#         try:
#             size = at.get_size()           # or at.size
#             data = at.read_buffer(size)    # ★ 여기!
#         except Exception as e:
#             logger.debug(f"[{i}] read_buffer 실패 ➜ {e}")
#             continue

#         # ── 2) 파일명: 최신 pypff에선 이름 프로퍼티가 없다 ──
#         name = f"attach_{i}_{msg.identifier}"
#         # 필요하면 magic / mimetypes 로 확장자 추측

#         # base_dir을 Path 객체로 변환하고 안전한 파일명 사용
#         base_dir_path = Path(base_dir) if not isinstance(base_dir, Path) else base_dir
#         safe_name = sanitize_filename(name)
#         save_path = base_dir_path / safe_name
#         save_path.parent.mkdir(parents=True, exist_ok=True)

#         with open(save_path, "wb") as fp:
#             fp.write(data)

#         results.append({
#             "filename": name,
#             "size": size,
#             "mime_type": "",          # pypff 에서 직접 못 받음
#             "save_path": str(save_path)
#         })
#     return results


# import re
# from pathlib import Path
# from uuid import uuid4
# import pypff

# ───── MAPI property IDs we care about ─────────────────────────
PR_ATTACH_LONG_FILENAME = 0x3707  # unicode, 최대 256자
PR_ATTACH_FILENAME      = 0x3704  # 8.3 짧은 이름 (fallback)

INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
def safe_name(name: str) -> str:
    """OS 불가 문자 제거 + 공백 정리"""
    name = INVALID.sub('_', name).strip().rstrip('. ')
    return name or f'file_{uuid4().hex[:8]}'

# ─── 보조 함수 ──────────────────────────────────────────────
def decode_mapi_string(raw: bytes | str) -> str:
    """
    MAPI 바이너리 문자열을 보기 좋은 파이썬 str 로 변환
    1) bytes → UTF-16LE 시도
    2) bytes → UTF-8 시도
    3) 실패 시 latin-1 로 강제 디코드
    4) 이미 str 이면 그대로
    """
    if isinstance(raw, bytes):
        for enc in ("utf-16-le", "utf-8" ):
            try:
                return raw.decode(enc).rstrip("\x00")
            except UnicodeDecodeError:
                pass
        # 마지막 보루
        return raw.decode("latin-1", errors="replace").rstrip("\x00")
    return raw.rstrip("\x00")  # str 인 경우

# ─── 수정된 find_attach_name ───────────────────────────────
def find_attach_name(att: pypff.attachment) -> str | None:
    """
    attachment 객체 내부 record set에서
    0x3707 → 0x3704 순서로 파일명을 찾아 UTF-8 문자열로 반환
    """
    wanted = (PR_ATTACH_LONG_FILENAME, PR_ATTACH_FILENAME)

    for rs_idx in range(getattr(att, "number_of_record_sets", 0)):
        rs = att.get_record_set(rs_idx)
        for ent_idx in range(rs.number_of_entries):
            ent = rs.get_entry(ent_idx)
            if ent.entry_type in wanted:
                try:
                    raw = ent.get_data()
                    return decode_mapi_string(raw)
                except Exception:
                    pass
    return None

def find_attach_name(att: pypff.attachment) -> str | None:
    """
    attachment 객체 내부 record set에서 파일명을 찾아 UTF-8 문자열로 반환
    여러 MAPI 속성을 확인하고 디코딩 방법을 개선
    """
    # 우선순위: 긴 파일명 → 짧은 파일명 → 표시명
    wanted = (
        PR_ATTACH_LONG_FILENAME,    # 0x3707
        PR_ATTACH_FILENAME,         # 0x3704  
        PR_DISPLAY_NAME,            # 0x3001 (추가)
        PR_ATTACH_EXTENSION,        # 0x3703 (확장자만)
    )
    
    # 디버깅을 위한 로그 추가
    logger.debug(f" 첨부파일 record_sets 수: {getattr(att, 'number_of_record_sets', 0)}")
    
    for rs_idx in range(getattr(att, "number_of_record_sets", 0)):
        rs = att.get_record_set(rs_idx)
        logger.debug(f" Record Set {rs_idx}: entries={rs.number_of_entries}")
        
        for ent_idx in range(rs.number_of_entries):
            ent = rs.get_entry(ent_idx)
            entry_type = ent.entry_type
            
            # 모든 엔트리 타입 로그 출력
            logger.debug(f" Entry {ent_idx}: type=0x{entry_type:04X}")
            
            if entry_type in wanted:
                try:
                    raw = ent.get_data()
                    logger.debug(f" Found entry type 0x{entry_type:04X}, raw data: {raw[:50]}...")
                    
                    # 다양한 디코딩 방법 시도
                    decoded = decode_mapi_string_enhanced(raw, entry_type)
                    if decoded:
                        logger.debug(f" Successfully decoded: '{decoded}'")
                        return decoded
                        
                except Exception as e:
                    logger.debug(f" Entry 0x{entry_type:04X} decode failed: {e}")
                    
    return None


def decode_mapi_string_enhanced(raw_data: bytes, entry_type: int) -> str | None:
    """
    MAPI 문자열을 다양한 방법으로 디코딩 시도
    """
    if not raw_data:
        return None
    
    logger.debug(f" Raw data length: {len(raw_data)}")
    logger.debug(f" Raw data hex: {raw_data.hex()}")
    
    # 방법 1: 기존 decode_mapi_string 사용
    try:
        result = decode_mapi_string(raw_data)
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f" decode_mapi_string failed: {e}")
    
    # 방법 2: UTF-16LE 디코딩 (Windows 기본)
    try:
        # null 종료 문자 제거
        if len(raw_data) >= 2 and raw_data[-2:] == b'\x00\x00':
            raw_data = raw_data[:-2]
        result = raw_data.decode('utf-16le').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f" UTF-16LE decode failed: {e}")
    
    # 방법 3: UTF-8 디코딩
    try:
        # null 종료 문자 제거
        if raw_data.endswith(b'\x00'):
            raw_data = raw_data[:-1]
        result = raw_data.decode('utf-8').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f" UTF-8 decode failed: {e}")
    
    # 방법 4: CP949 (한국어 환경)
    try:
        result = raw_data.decode('cp949').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f" CP949 decode failed: {e}")
    
    # 방법 5: Latin-1 (바이트 그대로)
    try:
        result = raw_data.decode('latin-1').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f" Latin-1 decode failed: {e}")
    
    return None


def is_inline_attachment(att) -> bool:
    """
    pypff 첨부파일이 인라인 첨부파일(명함, 서명 이미지 등)인지 판별
    True: 인라인 첨부파일 (저장하지 않음)
    False: 실제 첨부파일 (저장함)
    """
    try:
        # 1. 파일명 기반 판별
        filename = att.name or ""
        display_name = getattr(att, 'display_name', '') or ""
        
        # 일반적인 인라인 이미지 파일명 패턴
        inline_patterns = [
            'image001.', 'image002.', 'image003.',  # Outlook 자동 생성
            'oledata.mso',  # OLE 객체
            'filelist.xml',  # 메타데이터
            'themedata.thmx',  # 테마 데이터
            'colorschememapping.xml',  # 색상 스키마
            'header.xml', 'footer.xml',  # 헤더/푸터
            'atl00001.gif', 'atl00002.gif',  # ATL 이미지
        ]
        
        filename_lower = filename.lower()
        for pattern in inline_patterns:
            if pattern in filename_lower:
                return True
        
        # 2. 파일 크기 기반 판별 (매우 작은 파일은 보통 인라인)
        try:
            size = att.size
            if size < 1024:  # 1KB 미만의 매우 작은 파일
                return True
        except:
            pass
        
        # 3. MIME 타입 기반 판별
        try:
            # pypff에서 mime_type이 있는지 확인
            if hasattr(att, 'mime_type'):
                mime_type = att.mime_type or ""
                if mime_type.startswith('image/') and size < 50 * 1024:  # 50KB 미만 이미지
                    return True
        except:
            pass
        
        # 4. 첨부파일 속성 기반 판별 (원본 로직 개선)
        try:
            # pypff attachment 객체에서 직접 속성 확인
            if hasattr(att, 'is_inline') and att.is_inline:
                return True
                
            # 숨김 속성 확인
            if hasattr(att, 'is_hidden') and att.is_hidden:
                return True
                
        except:
            pass
        
        # 5. 확장자 기반 추가 판별
        image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.wmf', '.emf']
        if any(filename_lower.endswith(ext) for ext in image_extensions):
            # 이미지 파일이면서 크기가 작거나 특정 명명 패턴이면 인라인으로 판별
            try:
                size = att.size
                if size < 100 * 1024:  # 100KB 미만
                    # 파일명이 자동생성 패턴이면 인라인
                    if any(x in filename_lower for x in ['image', 'ole', 'atl']):
                        return True
            except:
                pass
        
        # 6. 특수 케이스: 빈 파일명이면서 작은 크기
        if not filename.strip():
            try:
                if att.size < 10 * 1024:  # 10KB 미만
                    return True
            except:
                pass
        
        return False
        
    except Exception as e:
        logger.debug(f"첨부파일 판별 중 오류: {e}")
        return False

def make_physical_file_name(prefix: str = "", ext: str = "") -> str:
    now = datetime.now()
    # 앞부분: YYYYMMDD_HHMMSS
    # 뒷부분: microseconds(000000 ~ 999999)
    micros = f"{now.microsecond:06d}"
    return f"{prefix}_{micros}{ext}"

def extract_path(path: str, base_dir: str) -> str:
    """
    주어진 경로에서 base_dir을 제거하고 상대 경로를 반환
    """
    path = Path(path)
    base_dir = Path(base_dir)
    
    # base_dir이 path의 상위 디렉토리인지 확인
    if not path.is_relative_to(base_dir):
        return str(path)  # base_dir이 포함되지 않은 경우 원본 경로 반환
    
    # 상대 경로 추출
    relative_path = path.relative_to(base_dir)
    return str(relative_path)

def safe_get_attachment_count(msg):
    """안전하게 첨부파일 수를 가져오는 함수"""
    try:
        return msg.number_of_attachments
    except Exception as e:
        logger.warning(f"첨부파일갯수를 가져오는데 실패했습니다. 오류: {e}")
        return -1

def extract_attachments(msg: pypff.message,
                        base_dir: Path, email_id, attach_count) -> list[dict]:
    """
    첨부파일을 base_dir 하위에 저장하고 디버깅 정보 출력
    """
    results = []

    for i in range(attach_count):
        logger.debug(f"\n === 첨부파일 {i+1} 처리 시작 ===")
        
        att = msg.get_attachment(i)
        
        # ── (1) 원본 파일명 찾기 ───────────────────────────────
        raw_name = find_attach_name(att)
        logger.debug(f" 추출된 파일명: '{raw_name}'")
        
        # 파일명이 제대로 추출되지 않은 경우 대체 방법 시도
        if not raw_name or raw_name.startswith('___'):
            logger.debug(f" 파일명 추출 실패, 대체 방법 시도")
            # pypff의 다른 속성들 확인
            try:
                if hasattr(att, 'get_name'):
                    alt_name = att.get_name()
                    logger.debug(f" att.get_name(): '{alt_name}'")
                    if alt_name and not alt_name.startswith('___'):
                        raw_name = alt_name
            except:
                pass
        
        name = safe_name(raw_name) if raw_name else f'attach_{i}'
        logger.debug(f" 최종 파일명: '{name}'")
        # image001.png, image002.jpg 등 자동 생성 인라인 이미지 패턴 여부 체크
        inline_img_pattern = re.compile(r"^image\d{3}\.(png|jpg|jpeg|gif|bmp|tiff|wmf|emf)$", re.I)
        is_auto_inline = bool(inline_img_pattern.match(name))
        if is_auto_inline:
            logger.debug(f" 자동 생성 인라인 이미지로 판단되어 저장하지 않습니다.")
            continue
        # ── (2) 실제 데이터 읽기 ───────────────────────────────
        try:
            size = att.get_size()
            data = att.read_buffer(size)
        except Exception as e:
            logger.error(f" 데이터 읽기 실패: {e}")
            continue
        
        # ── (3) 파일 저장 ─────────────────────────────────────
        base_dir_path = Path(base_dir) if not isinstance(base_dir, Path) else base_dir
        ext = Path(name).suffix  # name에서 확장자 추출 (. 포함)
        physical_file_name = make_physical_file_name(email_id, ext)

        save_path = base_dir_path / physical_file_name
        save_path.parent.mkdir(parents=True, exist_ok=True)

        # 중복 시 뒤에 (_1), (_2) … 덧붙임
        dup = 1
        while save_path.exists():
            save_path = save_path.with_stem(f"{save_path.stem}_{dup}")
            dup += 1
        
        with open(save_path, "wb") as fp:
            fp.write(data)
        
        logger.debug(f" 파일 저장 완료: {save_path}")
        save_folder = extract_path( str(save_path.parent), settings.ATTATCH_BASE_DIR )

        results.append({
            "email_id": email_id,
            "org_file_name": name,
            "phy_file_name": physical_file_name,
            "save_folder": save_folder,
            "file_size": size
        })
    
    return results


