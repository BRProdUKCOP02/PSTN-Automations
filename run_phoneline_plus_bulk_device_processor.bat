@echo off
REM Batch file to run Phoneline+ Bulk Device Processor
REM Used by Task Scheduler for automated processing

cd /d "C:\Users\Public\RPA\code\PSTN Migration"
call "C:\Users\Public\RPA\code\.venv\Scripts\activate.bat"
python phoneline_plus_bulk_device_processor.py
exit /b %ERRORLEVEL%
