@echo off
python -m pip install -r requirements.txt
copy .env.example .env >nul 2>nul
python app.py
pause
