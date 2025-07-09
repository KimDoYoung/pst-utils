#!/usr/bin/env bash
# -----------------------------------------------------------------------------
# make-for-linux.sh – Build a standalone ELF executable for the PST‑utils project
# -----------------------------------------------------------------------------
#  요구사항
#    1) 프로젝트 최상위 트리에  ⬇  구조가 있다고 가정합니다.
#       pst-utils/
#       ├─ src/               # 모든 파이썬 소스 (main.py, config.py …)
#       ├─ .env               # (선택) 실행 시 필요한 환경 변수 파일
#       └─ .venv/             # Poetry, uv, pip … 가 만든 가상환경
#    2) 가상환경에 PyInstaller 가 이미 설치돼 있어야 합니다.
#       (미리  `uv add pyinstaller`  또는  `pip install pyinstaller` 실행)
# -----------------------------------------------------------------------------
set -euo pipefail

# ─── 경로 계산 ---------------------------------------------------------------
PROJECT_ROOT=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
SRC_DIR="${PROJECT_ROOT}/src"
ENTRY_POINT="${SRC_DIR}/main.py"
VENV_BIN="${PROJECT_ROOT}/.venv/bin"
PYTHON="${VENV_BIN}/python"
PYINSTALLER="${VENV_BIN}/pyinstaller"

# ─── 확인 단계 ---------------------------------------------------------------
if [[ ! -x "${PYTHON}" ]]; then
  echo "[✖] 가상환경이 없습니다: ${VENV_BIN}" >&2
  echo "    →  uv venv 또는 python -m venv .venv  로 먼저 만들고 의존성을 설치하세요." >&2
  exit 1
fi

if ! "${PYTHON}" - <<'PY' >/dev/null 2>&1
import importlib.util, sys; sys.exit(importlib.util.find_spec('PyInstaller') is None)
PY
then
  echo "[✖] PyInstaller 가 가상환경에 없습니다.\n    →  uv add pyinstaller  또는  pip install pyinstaller" >&2
  exit 1
fi

# ─── pypff .so  경로 자동 탐색 ------------------------------------------------
PYPFF_BINARY=$("${PYTHON}" - <<'PY'
import pathlib, pypff, sys
so = pathlib.Path(pypff.__file__).with_suffix('.so')
print(so)
PY
)
if [[ ! -f "${PYPFF_BINARY}" ]]; then
  echo "[✖] pypff 공유라이브러리를 찾을 수 없습니다: ${PYPFF_BINARY}" >&2
  exit 1
fi

# ─── 정리: 이전 빌드 산출물 제거 -------------------------------------------
rm -rf "${PROJECT_ROOT}"/{build,dist} "${PROJECT_ROOT}"/*.spec || true

# ─── 빌드 -------------------------------------------------------------------
"${PYINSTALLER}" \
  --onefile \
  --name pst-extract \
  --paths "${SRC_DIR}" \
  --hidden-import=config \
  --add-binary "${PYPFF_BINARY}:." \
  "${ENTRY_POINT}"

# ─── .env 복사 ---------------------------------------------------------------
if [[ -f "${PROJECT_ROOT}/.env" ]]; then
  cp "${PROJECT_ROOT}/.env" "${PROJECT_ROOT}/dist/.env"
  echo "[+] .env 파일을 dist 디렉터리로 복사했습니다."
fi

echo "[✓] 빌드 완료  →  dist/pst-extract"
