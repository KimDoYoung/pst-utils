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
# logger.debug(email.__file__)
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

logger = get_logger(__name__)

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
    logger.debug(f"[DEBUG] 첨부파일 record_sets 수: {getattr(att, 'number_of_record_sets', 0)}")
    
    for rs_idx in range(getattr(att, "number_of_record_sets", 0)):
        rs = att.get_record_set(rs_idx)
        logger.debug(f"[DEBUG] Record Set {rs_idx}: entries={rs.number_of_entries}")
        
        for ent_idx in range(rs.number_of_entries):
            ent = rs.get_entry(ent_idx)
            entry_type = ent.entry_type
            
            # 모든 엔트리 타입 로그 출력
            logger.debug(f"[DEBUG] Entry {ent_idx}: type=0x{entry_type:04X}")
            
            if entry_type in wanted:
                try:
                    raw = ent.get_data()
                    logger.debug(f"[DEBUG] Found entry type 0x{entry_type:04X}, raw data: {raw[:50]}...")
                    
                    # 다양한 디코딩 방법 시도
                    decoded = decode_mapi_string_enhanced(raw, entry_type)
                    if decoded:
                        logger.debug(f"[DEBUG] Successfully decoded: '{decoded}'")
                        return decoded
                        
                except Exception as e:
                    logger.debug(f"[DEBUG] Entry 0x{entry_type:04X} decode failed: {e}")
                    
    return None


def decode_mapi_string_enhanced(raw_data: bytes, entry_type: int) -> str | None:
    """
    MAPI 문자열을 다양한 방법으로 디코딩 시도
    """
    if not raw_data:
        return None
    
    logger.debug(f"[DEBUG] Raw data length: {len(raw_data)}")
    logger.debug(f"[DEBUG] Raw data hex: {raw_data.hex()}")
    
    # 방법 1: 기존 decode_mapi_string 사용
    try:
        result = decode_mapi_string(raw_data)
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f"[DEBUG] decode_mapi_string failed: {e}")
    
    # 방법 2: UTF-16LE 디코딩 (Windows 기본)
    try:
        # null 종료 문자 제거
        if len(raw_data) >= 2 and raw_data[-2:] == b'\x00\x00':
            raw_data = raw_data[:-2]
        result = raw_data.decode('utf-16le').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f"[DEBUG] UTF-16LE decode failed: {e}")
    
    # 방법 3: UTF-8 디코딩
    try:
        # null 종료 문자 제거
        if raw_data.endswith(b'\x00'):
            raw_data = raw_data[:-1]
        result = raw_data.decode('utf-8').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f"[DEBUG] UTF-8 decode failed: {e}")
    
    # 방법 4: CP949 (한국어 환경)
    try:
        result = raw_data.decode('cp949').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f"[DEBUG] CP949 decode failed: {e}")
    
    # 방법 5: Latin-1 (바이트 그대로)
    try:
        result = raw_data.decode('latin-1').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        logger.debug(f"[DEBUG] Latin-1 decode failed: {e}")
    
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

def make_physical_file_name(prefix: str = "", ext="") -> str:
    now = datetime.now()
    # 앞부분: YYYYMMDD_HHMMSS
    # 뒷부분: microseconds(000000 ~ 999999)
    micros = f"{now.microsecond:06d}"
    return f"{prefix}_{micros}{ext}"

def extract_attachments(msg: pypff.message,
                        base_dir: Path, email_id) -> list[dict]:
    """
    첨부파일을 base_dir 하위에 저장하고 디버깅 정보 출력
    """
    results = []
    
    logger.debug(f"[DEBUG] 전체 첨부파일 수: {msg.number_of_attachments}")
    
    for i in range(msg.number_of_attachments):
        logger.debug(f"\n[DEBUG] === 첨부파일 {i} 처리 시작 ===")
        
        att = msg.get_attachment(i)
        # if is_inline_attachment(att):
        #     logger.debug(f"[DEBUG] 첨부파일 {i}은 본문에 삽입된 가짜 첨부입니다. 건너뜁니다.")
        #     continue
        
        # ── (1) 원본 파일명 찾기 ───────────────────────────────
        raw_name = find_attach_name(att)
        logger.debug(f"[DEBUG] 추출된 파일명: '{raw_name}'")
        
        # 파일명이 제대로 추출되지 않은 경우 대체 방법 시도
        if not raw_name or raw_name.startswith('___'):
            logger.debug(f"[DEBUG] 파일명 추출 실패, 대체 방법 시도")
            # pypff의 다른 속성들 확인
            try:
                if hasattr(att, 'get_name'):
                    alt_name = att.get_name()
                    logger.debug(f"[DEBUG] att.get_name(): '{alt_name}'")
                    if alt_name and not alt_name.startswith('___'):
                        raw_name = alt_name
            except:
                pass
        
        name = safe_name(raw_name) if raw_name else f'attach_{i}'
        logger.debug(f"[DEBUG] 최종 파일명: '{name}'")
        # image001.png, image002.jpg 등 자동 생성 인라인 이미지 패턴 여부 체크
        inline_img_pattern = re.compile(r"^image\d{3}\.(png|jpg|jpeg|gif|bmp|tiff|wmf|emf)$", re.I)
        is_auto_inline = bool(inline_img_pattern.match(name))
        logger.debug(f"[DEBUG] 자동 생성 인라인 이미지 패턴 여부: {is_auto_inline}")
        if is_auto_inline:
            logger.debug(f"[DEBUG] 자동 생성 인라인 이미지로 판단되어 저장하지 않습니다.")
            continue
        # ── (2) 실제 데이터 읽기 ───────────────────────────────
        try:
            size = att.get_size()
            logger.debug(f"[DEBUG] 첨부파일 크기: {size} bytes")
            data = att.read_buffer(size)
            logger.debug(f"[DEBUG] 데이터 읽기 성공")
        except Exception as e:
            logger.debug(f"[DEBUG] 데이터 읽기 실패: {e}")
            continue
        
        # ── (3) 파일 저장 ─────────────────────────────────────
        base_dir_path = Path(base_dir) if not isinstance(base_dir, Path) else base_dir
        save_path = base_dir_path / name
        save_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 중복 시 뒤에 (_1), (_2) … 덧붙임
        dup = 1
        while save_path.exists():
            save_path = save_path.with_stem(f"{save_path.stem}_{dup}")
            dup += 1
        
        with open(save_path, "wb") as fp:
            fp.write(data)
        
        logger.debug(f"[DEBUG] 파일 저장 완료: {save_path}")
                    #          attach["email_id"],
                    #  attach["save_folder"],
                    #  attach["org_file_name"],
                    #  attach["phy_file_name"])
        save_folder = str(save_path.parent)
        # email_id = str(getattr(msg, 'identifier'))
        ext = Path(name).suffix  # name에서 확장자 추출 (. 포함)
        physical_file_name = make_physical_file_name(email_id, ext)
        results.append({
            "email_id": email_id,
            "org_file_name": name,
            "phy_file_name": physical_file_name,
            "save_folder": save_folder,
            "file_size": size
        })
    
    return results


