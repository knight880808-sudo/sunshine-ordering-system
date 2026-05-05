@echo off
cd /d "%~dp0"
echo Starting POS API on http://127.0.0.1:5055  (proxy from pos-web /api)
python -m uvicorn pos_api:app --host 127.0.0.1 --port 5055 --reload
pause
