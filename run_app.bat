@echo off
rem COA 검증 화면 실행: 가상환경을 켜고 Streamlit을 띄운다. 브라우저가 자동으로 열린다. 종료는 Ctrl+C.
chcp 65001 > nul
cd /d "%~dp0"
call .venv\Scripts\activate.bat
streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
