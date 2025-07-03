#!/usr/bin/env python3
# pst_extract.py
from __future__ import annotations
import argparse, os, shutil, sys
from datetime import timezone, datetime
from pathlib import Path

import pypff  # pip install pypff
from logger import get_logger
import db_actions

BATCH_SIZE = 500
logger = get_logger("pst_extract")

def iso_kst(utc_dt: datetime) -> str:
    """UTC datetime → KST ISO 문자열(+09:00)"""
    return utc_dt.astimezone(timezone.utc)\
                 .astimezone(timezone.utc)\
                 .astimezone(timezone(timedelta(hours=9)))\
                 .isoformat()

def message_to_dict(msg, attach_dir: Path) -> dict:
    """pypff.Message → fund_mail + attach 리스트(dict)"""
    email_id = msg.identifier
    plain = msg.plain_text_body.decode(errors="replace") if msg.plain_text_body else ""
    utc_dt = msg.client_submit_time or datetime.utcnow().replace(tzinfo=timezone.utc)
    email = {
        "email_id": email_id,
        "subject": msg.subject,
        "sender": msg.sender_name or msg.sender_email_address,
        "to_recipients": ", ".join(filter(None, msg.display_to.split(";"))) if msg.display_to else "",
        "cc_recipients": ", ".join(filter(None, msg.display_cc.split(";"))) if msg.display_cc else "",
        "email_time": utc_dt.isoformat(),
        "kst_time": iso_kst(utc_dt),
        "content": plain,
        "attach_files": [],
    }

    # ---- 첨부파일 ----
    for a in msg.attachments:
        name = a.long_filename or a.file_name or f"{a.identifier}.bin"
        data = a.read_buffer()
        save_path = attach_dir / name
        save_path.write_bytes(data)
        email["attach_files"].append({
            "email_id": email_id,
            "save_folder": str(attach_dir),
            "file_name": name,
        })
    return email

def walk_folder(folder, attach_dir: Path):
    """yield Message objects depth-first"""
    for item in folder.sub_messages:
        yield item
    for sub in folder.sub_folders:
        yield from walk_folder(sub, attach_dir)

def main():
    ap = argparse.ArgumentParser(description="PST → SQLite extractor")
    ap.add_argument("--pst", required=True, type=Path, help="원본 .pst 파일 경로")
    ap.add_argument("--target", required=True, type=Path, help="내보낼 폴더")
    args = ap.parse_args()

    pst_path: Path = args.pst.expanduser().resolve()
    out_dir: Path = args.target.expanduser().resolve()
    db_path = out_dir / f"{pst_path.stem}.db"
    attach_dir = out_dir / "attach"
    attach_dir.mkdir(parents=True, exist_ok=True)

    logger.info("🔗 PST 열기: %s", pst_path)
    pst = pypff.file()
    pst.open(str(pst_path))

    logger.info("📂 DB 초기화: %s", db_path)
    db_actions.create_db_tables(db_path)

    batch = []
    root = pst.get_root_folder()
    total = 0
    for msg in walk_folder(root, attach_dir):
        batch.append(message_to_dict(msg, attach_dir))
        if len(batch) >= BATCH_SIZE:
            db_actions.save_email_data_to_db(batch, db_path)
            total += len(batch)
            logger.info("… 누적 %d 건 저장", total)
            batch.clear()

    # flush 남은 데이터
    if batch:
        db_actions.save_email_data_to_db(batch, db_path)
        total += len(batch)

    logger.info("✅ 완료! 총 %d 건 → %s", total, db_path)
    pst.close()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("⏹️ 사용자 중단")
        sys.exit(1)
