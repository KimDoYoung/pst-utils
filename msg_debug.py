import pypff
import sys
import os

# MAPI 속성 상수들
PR_MESSAGE_CLASS = 0x001A
PR_MESSAGE_CLASS_W = 0x001A001F
PR_SUBJECT = 0x0037
PR_SENDER_NAME = 0x0C1A

def debug_message_structure(msg: pypff.message, msg_index: int = 0) -> None:
    """
    메시지의 모든 가능한 구조를 디버깅 출력합니다.
    """
    print(f"\n=== MESSAGE {msg_index} DEBUG ===")
    print(f"Message ID: {getattr(msg, 'identifier', 'N/A')}")
    print(f"Message Type: {type(msg)}")
    
    # 1. 기본 속성들 확인
    print("\n--- Basic Attributes ---")
    basic_attrs = ['message_class', 'subject', 'sender_name', 'size', 'creation_time', 'delivery_time']
    for attr in basic_attrs:
        try:
            value = getattr(msg, attr, None)
            print(f"{attr}: {value} (type: {type(value)})")
        except Exception as e:
            print(f"{attr}: ERROR - {e}")
    
    # 2. 모든 속성 나열
    print("\n--- All Attributes ---")
    all_attrs = [attr for attr in dir(msg) if not attr.startswith('_')]
    for attr in all_attrs:
        try:
            value = getattr(msg, attr)
            if callable(value):
                print(f"{attr}: <method>")
            else:
                print(f"{attr}: {str(value)[:100]}... (type: {type(value)})")
        except Exception as e:
            print(f"{attr}: ERROR - {e}")
    
    # 3. record_sets 상세 분석
    print("\n--- Record Sets Analysis ---")
    try:
        if hasattr(msg, 'record_sets'):
            record_sets = msg.record_sets
            print(f"record_sets type: {type(record_sets)}")
            print(f"record_sets length: {len(record_sets) if hasattr(record_sets, '__len__') else 'N/A'}")
            
            for i, record_set in enumerate(record_sets):
                print(f"\n  Record Set {i}:")
                print(f"    Type: {type(record_set)}")
                
                # record_set의 모든 속성 확인
                rs_attrs = [attr for attr in dir(record_set) if not attr.startswith('_')]
                for attr in rs_attrs:
                    try:
                        value = getattr(record_set, attr)
                        if callable(value):
                            print(f"    {attr}: <method>")
                        else:
                            print(f"    {attr}: {value} (type: {type(value)})")
                    except Exception as e:
                        print(f"    {attr}: ERROR - {e}")
                
                # entries가 있는지 확인
                if hasattr(record_set, 'entries'):
                    print(f"    entries type: {type(record_set.entries)}")
                    print(f"    entries length: {len(record_set.entries) if hasattr(record_set.entries, '__len__') else 'N/A'}")
                    
                    for j, entry in enumerate(record_set.entries):
                        if j >= 5:  # 처음 5개만 출력
                            print(f"    ... ({len(record_set.entries) - 5} more entries)")
                            break
                        print(f"      Entry {j}:")
                        print(f"        Type: {type(entry)}")
                        
                        entry_attrs = [attr for attr in dir(entry) if not attr.startswith('_')]
                        for attr in entry_attrs:
                            try:
                                value = getattr(entry, attr)
                                if callable(value):
                                    print(f"        {attr}: <method>")
                                else:
                                    if attr == 'data' and hasattr(value, '__len__') and len(value) > 50:
                                        print(f"        {attr}: {value[:50]}... (len: {len(value)})")
                                    else:
                                        print(f"        {attr}: {value}")
                            except Exception as e:
                                print(f"        {attr}: ERROR - {e}")
                
                if i >= 3:  # 처음 몇 개만 상세히 출력
                    print(f"  ... ({len(record_sets) - 4} more record sets)")
                    break
        else:
            print("record_sets attribute not found")
    except Exception as e:
        print(f"Error analyzing record_sets: {e}")

def find_message_class_alternatives(msg: pypff.message) -> list:
    """
    message_class를 찾기 위한 모든 가능한 방법을 시도합니다.
    """
    results = []
    
    # 방법 1: 직접 속성
    try:
        if hasattr(msg, 'message_class'):
            mc = msg.message_class
            results.append(f"Direct attribute: {mc}")
    except Exception as e:
        results.append(f"Direct attribute: ERROR - {e}")
    
    # 방법 2: get_message_class 메서드
    try:
        if hasattr(msg, 'get_message_class'):
            mc = msg.get_message_class()
            results.append(f"get_message_class(): {mc}")
    except Exception as e:
        results.append(f"get_message_class(): ERROR - {e}")
    
    # 방법 3: MAPI 속성들
    mapi_methods = ['get_mapi_property', 'get_property', 'get_attribute']
    for method_name in mapi_methods:
        if hasattr(msg, method_name):
            method = getattr(msg, method_name)
            try:
                mc = method(PR_MESSAGE_CLASS)
                results.append(f"{method_name}(PR_MESSAGE_CLASS): {mc}")
            except Exception as e:
                results.append(f"{method_name}(PR_MESSAGE_CLASS): ERROR - {e}")
            
            try:
                mc = method(PR_MESSAGE_CLASS_W)
                results.append(f"{method_name}(PR_MESSAGE_CLASS_W): {mc}")
            except Exception as e:
                results.append(f"{method_name}(PR_MESSAGE_CLASS_W): ERROR - {e}")
    
    # 방법 4: 다른 이름의 속성들
    other_attrs = ['msg_class', 'item_type', 'object_type', 'message_type']
    for attr in other_attrs:
        try:
            if hasattr(msg, attr):
                value = getattr(msg, attr)
                results.append(f"{attr}: {value}")
        except Exception as e:
            results.append(f"{attr}: ERROR - {e}")
    
    return results

def simple_walk(folder: pypff.folder, max_debug: int = 3) -> None:
    """
    간단한 폴더 순회로 디버깅
    """
    debug_count = 0
    
    try:
        if hasattr(folder, 'sub_messages'):
            for i, msg in enumerate(folder.sub_messages):
                if debug_count >= max_debug:
                    break
                
                print(f"\n{'='*60}")
                debug_message_structure(msg, i)
                
                print(f"\n--- Message Class Search Results ---")
                alternatives = find_message_class_alternatives(msg)
                for alt in alternatives:
                    print(f"  {alt}")
                
                debug_count += 1
        
        # 하위 폴더도 확인
        if hasattr(folder, 'sub_folders') and debug_count < max_debug:
            for sub_folder in folder.sub_folders:
                if debug_count >= max_debug:
                    break
                simple_walk(sub_folder, max_debug - debug_count)
                
    except Exception as e:
        print(f"Error in simple_walk: {e}")

# 메인 실행
if __name__ == "__main__":
    pst_path = "/mnt/c/tmp/2021.pst"
    
    if not os.path.exists(pst_path):
        print(f"PST file not found: {pst_path}")
        sys.exit(1)
    
    try:
        print(f"Opening PST file: {pst_path}")
        pf = pypff.file()
        pf.open(pst_path)
        
        root_folder = pf.get_root_folder()
        if not root_folder:
            print("Cannot access root folder")
            sys.exit(1)
        
        print(f"pypff version info:")
        print(f"  pypff module: {pypff}")
        print(f"  pypff file type: {type(pf)}")
        
        # 디버깅 시작
        simple_walk(root_folder, max_debug=5)
        
        pf.close()
        print("\nDebugging completed!")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)