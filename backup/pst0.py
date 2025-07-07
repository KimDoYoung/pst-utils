import pypff, sys, os
from helper import get_message_class

MAX, cnt = 1000, 0

def walk(f, depth=0):
    global cnt
    for m in f.sub_messages:
        if not get_message_class(m).startswith("IPM.Note"):
            continue
        print(f"{'  '*depth}- [{m.identifier}] {m.subject}")
        cnt += 1
        if cnt >= MAX:
            return
    for sub in f.sub_folders:
        if cnt >= MAX:
            break
        walk(sub, depth+1)

pst = "/mnt/c/tmp/2021.pst"
if not os.path.exists(pst):
    sys.exit("file not found")
pf = pypff.file(); pf.open(pst)
walk(pf.get_root_folder())
pf.close()
