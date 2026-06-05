@echo off
REM Adobe Sign Widget Monitor - Task Scheduler Runner
REM This batch file activates the virtual environment and runs the widget monitor

REM Change to the script directory
cd /d "C:\Users\Public\RPA\code\PSTN Migration\adobe_sign_voice"

REM Activate Python virtual environment
call "C:\Users\Public\RPA\code\PSTN Migration\.venv\Scripts\activate.bat"

REM Run the widget monitor script
python widget_monitor.py

REM Capture exit code
set EXIT_CODE=%ERRORLEVEL%

REM Deactivate virtual environment
call deactivate

REM Return exit code to Task Scheduler
exit /b %EXIT_CODE%
