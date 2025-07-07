# logger.py
from __future__ import annotations

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv          # NEW (1)

# 한 번만 시도; 이미 로드돼 있으면 noop
load_dotenv(override=False)             # NEW (2)

# ─────────────────────────────── 설정 상수 ────────────────────────────────
DEFAULT_MAX_BYTES: int = 10 * 1024 * 1024   # 10 MB
DEFAULT_BACKUP_COUNT: int = 10
DEFAULT_LOG_LEVEL = "INFO"
DEFAULT_LOG_DIR = "./logs"


def _make_handlers(log_file: Path,
                   max_bytes: int,
                   backup_count: int) -> list[logging.Handler]:
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname).1s] %(name)s: %(message)s")

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setFormatter(fmt)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt)

    return [file_handler, console_handler]


def get_logger(name: str = "fund_mail") -> logging.Logger:
    """
    콘솔과 회전 로그 파일에 동시에 기록하는 로거를 반환한다.
    동일한 이름으로 여러 번 호출해도 중복 핸들러 없이 같은 로거를 돌려준다.
    """
    logger = logging.getLogger(name)
    if logger.handlers:          # 이미 세팅 완료
        return logger

    # 환경변수에서 읽기 ― .env → load_dotenv 로 이미 올라와 있음
    log_dir = Path(os.getenv("LOG_DIR", DEFAULT_LOG_DIR)).expanduser()
    log_level = os.getenv("LOG_LEVEL", DEFAULT_LOG_LEVEL).upper()
    max_bytes = int(os.getenv("LOG_MAX_BYTES", DEFAULT_MAX_BYTES))
    backup_count = int(os.getenv("LOG_BACKUP_COUNT", DEFAULT_BACKUP_COUNT))

    # 디렉터리 준비
    log_dir.mkdir(parents=True, exist_ok=True)

    # 핸들러 준비
    handlers = _make_handlers(log_dir / f"{name}.log",
                              max_bytes, backup_count)
    for h in handlers:
        logger.addHandler(h)

    # 레벨·전파 설정
    logger.setLevel(getattr(logging, log_level, logging.INFO))
    logger.propagate = False
    return logger


# 모듈 전역 로거
logger = get_logger()
