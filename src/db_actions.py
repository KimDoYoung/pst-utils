
import sqlite3

from exceptions import DBCreateError, DBWriteError
from logger import get_logger

logger = get_logger(__name__)

def create_db_tables(db_path):
    """
    fund_mail 테이블과 fund_mail_attach 테이블을 생성합니다.
    """
    if db_path is None:
        logger.error("❌ DB 경로가 지정되지 않았습니다.")
        raise DBCreateError("❌ DB 경로가 지정되지 않았습니다.")

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS fund_mail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,  
            email_id TEXT ,  -- office365의 email_id
            subject TEXT,
            sender_address TEXT,
            sender_name TEXT,  
            from_address TEXT,  
            from_name TEXT,  
            to_recipients TEXT,  -- 수신자 목록
            cc_recipients TEXT,  -- 참조자 목록
            email_time TEXT,
            kst_time TEXT,
            content TEXT,
            msg_kind TEXT,
            folder_path TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS fund_mail_attach (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_id INTEGER,
            email_id TEXT,  -- fund_mail 테이블의 id
            save_folder TEXT,
            org_file_name TEXT,
            phy_file_name TEXT
        )
    """)        
    conn.commit()
    conn.close()
    logger.info(f"✅ DB 테이블이 생성되었습니다: {db_path}")


def save_email_data_to_db(email_data_list, db_path):
    """
    이메일 + 첨부파일을 **트랜잭션**으로 저장.
    실패 시 전체 롤백 → 데이터 일관성 보장
    """
    if not db_path:
        raise ValueError("db_path가 None입니다")

    try:
        # 1) with 블록 = 자동 BEGIN / COMMIT / (예외 시) ROLLBACK
        with sqlite3.connect(db_path) as conn:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()
            attach_count = 0
            for email in email_data_list:
                # --- 1) 메일 INSERT ---------------------------------
                cur.execute("""
                    INSERT INTO fund_mail
                          (email_id, subject, sender_address, sender_name, from_address, from_name,
                           to_recipients, cc_recipients,
                           email_time, kst_time, content)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    email["email_id"],
                    email["subject"],
                    email["sender_address"],
                    email["sender_name"],
                    email["from_address"],
                    email["from_name"],
                    email["to_recipients"],
                    email["cc_recipients"],
                    email["email_time"],
                    email["kst_time"],
                    email["content"],
                ))
                parent_id = cur.lastrowid

                # --- 2) 첨부파일 bulk INSERT ------------------------
                attach_rows = [
                    (parent_id,
                     attach["email_id"],
                     attach["save_folder"],
                     attach["file_name"])
                    for attach in email.get("attach_files", [])
                ]
                cur.executemany("""
                    INSERT INTO fund_mail_attach
                          (parent_id, email_id, save_folder, file_name)
                    VALUES (?, ?, ?, ?)
                """, attach_rows)
                attach_count = attach_count + len(attach_rows)
            # with-블록을 무사히 통과해야만 COMMIT 발생
            logger.info("✅ 이메일 %d건, 첨부파일 %d개 트랜잭션 저장 완료", len(email_data_list), attach_count)
        return db_path
    except sqlite3.Error as e:
        # 예외 발생 시 자동 ROLLBACK
        raise DBWriteError("❌ DB 저장 실패 - 전체 롤백됨")
