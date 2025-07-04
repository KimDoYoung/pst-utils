import pypff, os, sys
count = 0

def walk(folder, depth=0):
    global count
    indent = "  " * depth

    for m in folder.sub_messages:
        # 이제 바로 속성 사용!
        if not (m.message_class or "").startswith("IPM.Note"):
            continue

        print(f"{indent}- [{m.identifier}] {m.subject} / {m.sender_name}")
        count += 1
        if count >= 1000:
            return                      # 1,000건까지만

    for sub in folder.sub_folders:
        if count >= 1000:
            break
        walk(sub, depth + 1)

def main(pst_path: str):
    pst = pypff.file()
    pst.open(pst_path)
    walk(pst.get_root_folder())
    pst.close()

if __name__ == "__main__":
    pst_path = "/mnt/c/tmp/2021.pst"
    if not os.path.exists(pst_path):
        print(f"File not found: {pst_path}")
        sys.exit(1)
    main(pst_path)