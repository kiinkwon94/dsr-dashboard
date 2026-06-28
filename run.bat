@echo off
cd /d "%~dp0"
echo DSR Dashboard 시작 중...
echo 브라우저에서 http://localhost:8501 이 자동으로 열립니다.
echo 종료: 이 창을 닫거나 Ctrl+C
echo.
"C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe" -m streamlit run app.py --server.headless false
pause
