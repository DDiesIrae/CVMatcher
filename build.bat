@echo off
py -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --noconfirm --clean --onefile --windowed --name CVMatcher --collect-all keyring app.py
echo.
echo READY: dist\CVMatcher.exe
pause
