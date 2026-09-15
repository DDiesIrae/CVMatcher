@echo off
py -m venv .venv 2>nul
call .venv\Scripts\activate
pip install -r requirements.txt
python app.py
