@echo off
REM Batch wrapper for Bulk Check Regrades script
REM This ensures proper environment activation for Task Scheduler

cd /d "C:\Users\Public\RPA\code\PSTN Migration"

REM Activate Python virtual environment
call "C:\Users\Public\RPA\code\.venv\Scripts\activate.bat"

REM Run the bulk check regrades script
python bulk_check_regrades.py

REM Capture exit code
set EXIT_CODE=%ERRORLEVEL%

REM Deactivate virtual environment
call deactivate

REM Return exit code to Task Scheduler
exit /b %EXIT_CODE%
