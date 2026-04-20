@echo off
:: Mergermarket Newsletter – Windows Task Scheduler wrapper
:: Place this file next to mergermarket_newsletter.py and point the scheduler at it.

setlocal
set SCRIPT_DIR=%~dp0

:: Log start time
echo. >> "C:\Temp\mergermarket_log.txt"
echo ========================================== >> "C:\Temp\mergermarket_log.txt"
echo %DATE% %TIME% – Task Scheduler triggered >> "C:\Temp\mergermarket_log.txt"
echo ========================================== >> "C:\Temp\mergermarket_log.txt"

python "%SCRIPT_DIR%mergermarket_newsletter.py" --headless >> "C:\Temp\mergermarket_log.txt" 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Script exited with code %ERRORLEVEL% >> "C:\Temp\mergermarket_log.txt"
)

endlocal
