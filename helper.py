import pypff
import email,  email.policy
import re
from typing import Sequence

from charset_normalizer import from_bytes
# print(email.__file__)
# MAPI 속성 상수
PR_MESSAGE_CLASS = 0x001A  # 26 in decimal

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
        print(f"Warning: Error in record_sets access: {e}")
    
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
        print("header-parse error:", e)
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