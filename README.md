# pst-utils

## 설치
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