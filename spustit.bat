@echo off
setlocal
chcp 65001 >nul
title MailAI - Lokalni aplikace

:: Set working directory to the folder where the script is located
cd /d "%~dp0"

echo.
echo  ====================================
echo    MailAI - start
echo  ====================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo CHYBA: Python neni nainstalovan nebo neni v PATH.
    pause
    exit /b 1
)

:: Create Virtual Environment if it doesn't exist
if not exist ".venv\Scripts\python.exe" (
    echo [1/2] Prvni spusteni - instaluji zavislosti...
    python -m venv .venv
    if errorlevel 1 (
        echo CHYBA: Nepodarilo se vytvorit virtualni prostredi.
        pause
        exit /b 1
    )
    .venv\Scripts\python -m pip install --upgrade pip
    .venv\Scripts\pip install streamlit openai requests
)

:: Check if the app file actually exists
if not exist "local_app.py" (
    echo CHYBA: Soubor local_app.py nebyl nalezen v tomto adresari.
    pause
    exit /b 1
)

:: Start browser in background
start "" cmd /c "timeout /t 3 >nul && start http://127.0.0.1:8502"

:: Run App
echo [2/2] Spoustim aplikaci...
.venv\Scripts\streamlit run local_app.py --server.address 127.0.0.1 --server.headless true --server.port 8502

pause