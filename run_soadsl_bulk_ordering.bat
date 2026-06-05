@echo off
REM SOADSL Bulk Ordering - Task Scheduler Runner
REM This batch file activates the virtual environment and runs the script

REM Change to the script directory
cd /d "C:\Users\Public\RPA\code\PSTN Migration"

REM Activate virtual environment and run script
call "C:\Users\Public\RPA\code\.venv\Scripts\activate.bat"
python soadsl_bulk_ordering.py

REM Deactivate virtual environment
deactivate

REM Exit with the Python script's exit code
exit /b %ERRORLEVEL%
