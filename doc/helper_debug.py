import pypff
from pathlib import Path

from helper import decode_mapi_string, safe_name

# 추가로 MAPI 속성 상수들 정의 (혹시 없다면)
PR_ATTACH_LONG_FILENAME = 0x3707
PR_ATTACH_FILENAME = 0x3704
PR_DISPLAY_NAME = 0x3001
PR_ATTACH_EXTENSION = 0x3703

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
    print(f"[DEBUG] 첨부파일 record_sets 수: {getattr(att, 'number_of_record_sets', 0)}")
    
    for rs_idx in range(getattr(att, "number_of_record_sets", 0)):
        rs = att.get_record_set(rs_idx)
        print(f"[DEBUG] Record Set {rs_idx}: entries={rs.number_of_entries}")
        
        for ent_idx in range(rs.number_of_entries):
            ent = rs.get_entry(ent_idx)
            entry_type = ent.entry_type
            
            # 모든 엔트리 타입 로그 출력
            print(f"[DEBUG] Entry {ent_idx}: type=0x{entry_type:04X}")
            
            if entry_type in wanted:
                try:
                    raw = ent.get_data()
                    print(f"[DEBUG] Found entry type 0x{entry_type:04X}, raw data: {raw[:50]}...")
                    
                    # 다양한 디코딩 방법 시도
                    decoded = decode_mapi_string_enhanced(raw, entry_type)
                    if decoded:
                        print(f"[DEBUG] Successfully decoded: '{decoded}'")
                        return decoded
                        
                except Exception as e:
                    print(f"[DEBUG] Entry 0x{entry_type:04X} decode failed: {e}")
                    
    return None


def decode_mapi_string_enhanced(raw_data: bytes, entry_type: int) -> str | None:
    """
    MAPI 문자열을 다양한 방법으로 디코딩 시도
    """
    if not raw_data:
        return None
    
    print(f"[DEBUG] Raw data length: {len(raw_data)}")
    print(f"[DEBUG] Raw data hex: {raw_data.hex()}")
    
    # 방법 1: 기존 decode_mapi_string 사용
    try:
        result = decode_mapi_string(raw_data)
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        print(f"[DEBUG] decode_mapi_string failed: {e}")
    
    # 방법 2: UTF-16LE 디코딩 (Windows 기본)
    try:
        # null 종료 문자 제거
        if len(raw_data) >= 2 and raw_data[-2:] == b'\x00\x00':
            raw_data = raw_data[:-2]
        result = raw_data.decode('utf-16le').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        print(f"[DEBUG] UTF-16LE decode failed: {e}")
    
    # 방법 3: UTF-8 디코딩
    try:
        # null 종료 문자 제거
        if raw_data.endswith(b'\x00'):
            raw_data = raw_data[:-1]
        result = raw_data.decode('utf-8').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        print(f"[DEBUG] UTF-8 decode failed: {e}")
    
    # 방법 4: CP949 (한국어 환경)
    try:
        result = raw_data.decode('cp949').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        print(f"[DEBUG] CP949 decode failed: {e}")
    
    # 방법 5: Latin-1 (바이트 그대로)
    try:
        result = raw_data.decode('latin-1').strip()
        if result and not result.startswith('___'):
            return result
    except Exception as e:
        print(f"[DEBUG] Latin-1 decode failed: {e}")
    
    return None


def extract_attachments_debug(msg: pypff.message, base_dir: Path) -> list[dict]:
    """
    첨부파일을 base_dir 하위에 저장하고 디버깅 정보 출력
    """
    results = []
    
    print(f"[DEBUG] 전체 첨부파일 수: {msg.number_of_attachments}")
    
    for i in range(msg.number_of_attachments):
        print(f"\n[DEBUG] === 첨부파일 {i} 처리 시작 ===")
        
        try:
            att = msg.get_attachment(i)
            print(f"[DEBUG] 첨부파일 객체 생성 성공")
        except Exception as e:
            print(f"[DEBUG] 첨부파일 {i} 객체 접근 실패: {e}")
            continue
        
        # ── (1) 원본 파일명 찾기 ───────────────────────────────
        raw_name = find_attach_name(att)
        print(f"[DEBUG] 추출된 파일명: '{raw_name}'")
        
        # 파일명이 제대로 추출되지 않은 경우 대체 방법 시도
        if not raw_name or raw_name.startswith('___'):
            print(f"[DEBUG] 파일명 추출 실패, 대체 방법 시도")
            # pypff의 다른 속성들 확인
            try:
                if hasattr(att, 'get_name'):
                    alt_name = att.get_name()
                    print(f"[DEBUG] att.get_name(): '{alt_name}'")
                    if alt_name and not alt_name.startswith('___'):
                        raw_name = alt_name
            except:
                pass
        
        name = safe_name(raw_name) if raw_name else f'attach_{i}'
        print(f"[DEBUG] 최종 파일명: '{name}'")
        
        # ── (2) 실제 데이터 읽기 ───────────────────────────────
        try:
            size = att.get_size()
            print(f"[DEBUG] 첨부파일 크기: {size} bytes")
            data = att.read_buffer(size)
            print(f"[DEBUG] 데이터 읽기 성공")
        except Exception as e:
            print(f"[DEBUG] 데이터 읽기 실패: {e}")
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
        
        print(f"[DEBUG] 파일 저장 완료: {save_path}")
        
        results.append({
            "filename": name,
            "size": size,
            "save_path": str(save_path),
        })
    
    return results


