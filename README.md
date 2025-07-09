# pst-utils

## 개요

- KFS의 fund계정 outlook 메일, 백업받은 파일로부터 메일내용을을 뽑아서 sqlite3 db에 저장
- 첨부파일을 일자별로 폴더를 만들어서 저장

## 기술스택

- python
- pfflib 사용, 윈도우에서 잘 되지 않아서 wsl & ubuntu에서 작업
- fedora에서도 uv sync로 설치 가능

## 빌드

- make-for-linux.sh 
- dist에 pst-extract, .env  2개의 파일 만들어짐

## 실행

- 만들어지는 db와 첨부파일의 저장폴더를 .env에서 설정
- pst-extract <pst파일명>
- 일부 첨부파일이 저장되지 않음. db 의 비고에 기술, chatgpt에 의하면 pst파일이 깨졌다고 함.
- scanpst.exe (ms가 제공하는 pst파일 복원 프로그램)으로 복원해도 역시 첨부파일이 저장이 되지 않음.
- scanpst위치 (C:\\Program Files\\Microsoft Office\\root\\Office16\\scanpst.exe)

## 설치 참조
```bash
sudo apt update
sudo apt-get install -y libpff-dev python3-dev build-essential  # Debian/Ubuntu
sudo apt install python3-pypff libpff-dev pff-tools
uv add libpff-python

uv pip install \
     --no-binary libpff-python \
     --no-cache-dir \
     --force-reinstall \
     libpff-python==20231205
```