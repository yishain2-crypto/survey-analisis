@echo off
chcp 65001 > nul
title דשבורד שאלוני מילואים

echo ========================================
echo    דשבורד שאלוני מילואים
echo ========================================
echo.

:: Kill any existing streamlit process on port 8501
echo מחפש תהליכים ישנים בפורט 8501...
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":8501 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
echo.

:: Check if streamlit is installed
where streamlit >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo מתקין תלויות...
    pip install -r requirements.txt
    echo.
)

echo מפעיל את הדשבורד...
echo.
echo  ^>^> http://localhost:8501 ^<^<
echo.
echo לעצירה: לחץ Ctrl+C בחלון זה
echo ========================================
echo.

:: Open browser after 3 seconds (in background)
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:8501"

:: Run the app
python -m streamlit run app.py --server.port 8501 --server.headless false --browser.gatherUsageStats false --browser.serverAddress localhost

echo.
echo הדשבורד נסגר.
pause
