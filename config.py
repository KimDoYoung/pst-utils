"""
config.py  -- Pydantic v2 버전
────────────────────────────────────────────────────────────
1) .env → OS 환경변수 반영 (python-dotenv)
2) BaseSettings + field_validator 로 타입·검증·기본값 관리
3) settings.LOG_DIR 등 속성으로 전역 재사용
────────────────────────────────────────────────────────────
pip install python-dotenv pydantic>=2.0 pydantic_settings
# 사용 예시
# ---------------------------------------------------------
# from config import settings
# log_path = settings.LOG_DIR / "app.log"
# db_root  = settings.DB_DIR

"""

from __future__ import annotations

from pathlib import Path

from dotenv import load_dotenv
from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


# 1) .env → 환경변수 (한 번만 실행, 이미 로드돼 있으면 no-op)
load_dotenv(override=False)


class Settings(BaseSettings):
    # ───────────── Logging ─────────────
    LOG_DIR: Path = Field(default=Path("./logs"))
    LOG_LEVEL: str = Field(default="INFO")
    LOG_MAX_BYTES: int = Field(default=10 * 1024 * 1024)
    LOG_BACKUP_COUNT: int = Field(default=10)

    # ─────────── Base Directories ───────────
    DB_BASE_DIR: Path = Field(default=Path("./db"))
    ATTATCH_BASE_DIR: Path = Field(default=Path("./attachments"))  # 원본 표기 유지

    # ──────────── 공통 검증 & 폴더 자동 생성 ────────────
    @field_validator("LOG_DIR", "DB_BASE_DIR", "ATTATCH_BASE_DIR", mode="before")
    @classmethod
    def expand_and_ensure_dir(cls, v: str | Path) -> Path:
        """~ 확장 + 존재하지 않으면 폴더 생성 후 절대경로 반환"""
        p = Path(v).expanduser()
        p.mkdir(parents=True, exist_ok=True)
        return p.resolve()

    # ──────────── 모델 설정 (v2 스타일) ────────────
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,   # .env 키 대소문자 무시
    )


# 전역 싱글턴 인스턴스
settings = Settings()

