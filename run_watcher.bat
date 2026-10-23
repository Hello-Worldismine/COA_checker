@echo off
rem COA 메일 감시 실행: 가상환경을 켜고 mail_watcher.py를 실행한다. 종료는 Ctrl+C.
chcp 65001 > nul
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python mail_watcher.py %*
pause
