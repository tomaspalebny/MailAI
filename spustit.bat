@echo off
chcp 65001 >nul
title MailAI – Lokalni aplikace

cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════╗
echo  ║        MailAI – start            ║
echo  ╚══════════════════════════════════╝
echo.

REM ── Zkontroluj Python ────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  CHYBA: Python neni nainstalovan nebo neni v PATH.
    echo  Stahni Python z https://www.python.org/downloads/
    echo  Pri instalaci zapamatuj zaskrtnout "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

REM ── Vytvoř virtuální prostředí (jen při prvním spuštění) ─────────
if not exist ".venv\Scripts\python.exe" (
    echo  [1/2] Prvni spusteni – instaluji zavislosti, chvilku strpeni...
    echo.
    python -m venv .venv
    .venv\Scripts\python -m pip install --quiet --upgrade pip
    .venv\Scripts\pip install --quiet streamlit openai requests
    echo.
    echo  Instalace dokoncena!
    echo.
)

REM ── Otevři prohlížeč za 3 sekundy ───────────────────────────────
start "" cmd /c "timeout /t 3 >nul && start http://127.0.0.1:8502"

REM ── Spusť aplikaci ───────────────────────────────────────────────
echo  [2/2] Spoustim aplikaci... (zavri toto okno pro ukonceni)
echo.
.venv\Scripts\streamlit run local_app.py --server.address 127.0.0.1 --server.headless true --server.port 8502

pause
