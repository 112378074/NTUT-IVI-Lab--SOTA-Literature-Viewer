@echo off
REM ----------------------------------------------------------------
REM run_update.bat
REM Triggered by Windows Task Scheduler every Wed and Fri at 02:00.
REM Runs update_papers.py in the project root, logs to update_log.txt.
REM ----------------------------------------------------------------

set "PROJ=C:\Users\user\Desktop\Mypaper"
cd /d "%PROJ%"

REM Use the Python on PATH (Python 3.12 detected). Override here if needed.
set "PY=python"

echo. >> "%PROJ%\scripts\update_log.txt"
echo ====== Task Scheduler invoke at %date% %time% ====== >> "%PROJ%\scripts\update_log.txt"

%PY% "%PROJ%\scripts\update_papers.py" >> "%PROJ%\scripts\update_log.txt" 2>&1

exit /b %errorlevel%
