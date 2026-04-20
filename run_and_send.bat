@echo off
:: Mergermarket Newsletter – Task Scheduler wrapper with auto-send
::
:: Place this file next to mergermarket_newsletter.py.
::
:: ── Windows Task Scheduler setup ────────────────────────────────────────────
:: 1. Open Task Scheduler → Create Basic Task
:: 2. Name:    Mergermarket Newsletter
:: 3. Trigger: Daily – repeat Mon–Fri only
::             Start time: 09:00
::             Advanced: check "Run task as soon as possible after a scheduled
::             start is missed" (for days the machine was off at 09:00)
:: 4. Action:  Start a program
::             Program:   C:\Windows\System32\cmd.exe
::             Arguments: /c "C:\path\to\run_and_send.bat"
:: 5. Settings: "Run only when user is logged on"
:: ────────────────────────────────────────────────────────────────────────────

setlocal
set SCRIPT_DIR=%~dp0
set LOG_FILE=%USERPROFILE%\Downloads\mergermarket_log.txt

:: Log start time
echo. >> "%LOG_FILE%"
echo ========================================== >> "%LOG_FILE%"
echo %DATE% %TIME% – Task Scheduler triggered   >> "%LOG_FILE%"
echo ========================================== >> "%LOG_FILE%"

py "%SCRIPT_DIR%mergermarket_newsletter.py" --headless --send >> "%LOG_FILE%" 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Script exited with code %ERRORLEVEL% >> "%LOG_FILE%"
)

endlocal
