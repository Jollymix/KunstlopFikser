@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "kunstlopassistenten.py"
    goto :finished
)

where py >nul 2>&1
if not errorlevel 1 (
    py -3 "kunstlopassistenten.py"
    goto :finished
)

where python >nul 2>&1
if not errorlevel 1 (
    python "kunstlopassistenten.py"
    goto :finished
)

echo Fant ikke Python. Installer Python 3 og prov igjen.
pause
exit /b 1

:finished
set "exit_code=%errorlevel%"
if not "%exit_code%"=="0" (
    echo.
    echo Programmet avsluttet med feil. Installer eventuelt avhengigheter med:
    echo py -3 -m pip install -r requirements.txt
    pause
)
exit /b %exit_code%
