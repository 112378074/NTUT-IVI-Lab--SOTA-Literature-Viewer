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

REM IMPORTANT: redirect stdout/stderr to a SEPARATE file (run_update_console.log)
REM so we don't lock update_log.txt — update_papers.py opens that file directly
REM via its log() function, and Windows can't share the handle across both.

echo. >> "%PROJ%\scripts\run_update_console.log"
echo ====== Task Scheduler invoke at %date% %time% ====== >> "%PROJ%\scripts\run_update_console.log"

%PY% "%PROJ%\scripts\update_papers.py" >> "%PROJ%\scripts\run_update_console.log" 2>&1

exit /b %errorlevel%
