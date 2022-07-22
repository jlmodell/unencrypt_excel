python3 -m venv .
source bin/activate
pip install -r requirements.txt
pyinstaller -F -n decrypt_xls main.py