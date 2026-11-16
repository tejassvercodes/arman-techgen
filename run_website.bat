@echo off
cd /d "%~dp0"
python -c "import flask, lxml, pymupdf, waitress" 2>nul || python -m pip install -r requirements.txt
echo Starting Paper Desk at http://127.0.0.1:5000  (close this window to stop it)
start "" http://127.0.0.1:5000
python webapp\app.py
pause
