# Design Document: Automation Script Execution with Batch Files and Windows Task Scheduler

## Overview
The PSTN Migration automation system uses Windows batch (.bat) files as wrappers to execute Python scripts via Windows Task Scheduler. This architecture enables unattended, scheduled execution of bulk processing tasks for SoGEA orders, Phoneline+ customer provisioning, device ordering, and number allocation.

## Architecture Components

### 1. Batch File Wrapper Layer
Each Python automation script has a corresponding .bat file that serves as an execution wrapper:
- `run_bulk_check_orders.bat` → `bulk_check_orders.py`
- `run_sogea_bulk_ordering.bat` → `sogea_bulk_ordering.py`
- `run_phoneline_plus_bulk_processor.bat` → `phoneline_plus_bulk_processor.py`
- `run_phoneline_plus_bulk_device_processor.bat` → `phoneline_plus_bulk_device_processor.py`
- `run_phoneline_plus_bulk_number_processor.bat` → `phoneline_plus_bulk_number_processor.py`

**Batch File Structure:**
```batch
@echo off
REM Change to script directory
cd /d "C:\Users\Public\RPA\code\PSTN Migration"

REM Activate Python virtual environment
call "C:\Users\Public\RPA\code\.venv\Scripts\activate.bat"

REM Execute Python script
python <script_name>.py

REM Exit with Python's exit code
exit /b %ERRORLEVEL%
```

**Key Functions:**
- **Directory Navigation:** Uses `cd /d` to ensure correct working directory regardless of where Task Scheduler launches from
- **Virtual Environment Activation:** Activates the Python venv to ensure all dependencies are available
- **Exit Code Propagation:** Captures and returns Python script exit codes (`%ERRORLEVEL%`) to Task Scheduler for monitoring
- **Task Scheduler Compatibility:** Provides a stable entry point that Task Scheduler can reliably execute

### 2. Windows Task Scheduler Configuration
Tasks are defined in XML files and imported into Windows Task Scheduler. Each task configuration includes:

**Task Definition (XML Structure):**
- **RegistrationInfo:** Task metadata, author, description, and URI path
- **Triggers:** Schedule definition (start time, repetition interval)
- **Principals:** User account and privilege level for execution
- **Settings:** Execution policies and constraints
- **Actions:** The batch file to execute and working directory

**Common Trigger Configuration:**
```xml
<CalendarTrigger>
  <Repetition>
    <Interval>PT1M</Interval>  <!-- Run every 1 minute -->
    <StopAtDurationEnd>false</StopAtDurationEnd>
  </Repetition>
  <StartBoundary>2026-02-13T08:00:00</StartBoundary>
  <Enabled>true</Enabled>
  <ScheduleByDay>
    <DaysInterval>1</DaysInterval>  <!-- Every day -->
  </ScheduleByDay>
</CalendarTrigger>
```

**Critical Settings:**
- **MultipleInstancesPolicy:** `IgnoreNew` - Prevents concurrent executions if previous instance is still running
- **ExecutionTimeLimit:** `PT1H` or `PT72H` - Maximum runtime (1 hour or 72 hours)
- **StartWhenAvailable:** `true` - Runs task as soon as possible if scheduled time is missed
- **Priority:** `7` - Below normal priority to avoid system resource contention

### 3. Execution Flow

```
Windows Task Scheduler (Trigger fires) 
    ↓
Launches .bat file in specified WorkingDirectory
    ↓
Batch script changes to project directory
    ↓
Activates Python virtual environment
    ↓
Executes Python automation script
    ↓
Python script processes files from input directories
    ↓
Python script returns exit code (0=success, non-zero=error)
    ↓
Batch file propagates exit code to Task Scheduler
    ↓
Task Scheduler logs result in Task History
```

## Task Configurations

| Task Name | Batch File | Schedule | Timeout | Purpose |
|-----------|-----------|----------|---------|---------|
| Bulk Check Orders Processor | run_bulk_check_orders.bat | Every 1 minute | 1 hour | Monitors SoGEA order status |
| SoGEA Bulk Order Processor | run_sogea_bulk_ordering.bat | Every 1 minute | 1 hour | Processes SoGEA order files |
| Phoneline+ Customer Creation | run_phoneline_plus_bulk_processor.bat | Every 1 minute | 72 hours | Creates Phoneline+ customers |
| Phoneline+ Device Ordering | run_phoneline_plus_bulk_device_processor.bat | Every 1 minute | 72 hours | Processes device orders |
| Phoneline+ Number Allocation | run_phoneline_plus_bulk_number_processor.bat | Every 1 minute | 72 hours | Allocates phone numbers |

## Key Design Benefits

1. **Separation of Concerns:** Batch files handle environment setup; Python scripts handle business logic
2. **Task Scheduler Integration:** Native Windows scheduling with monitoring, logging, and alerting capabilities
3. **Reliability:** Automatic retry via `StartWhenAvailable` and multiple instance prevention
4. **Maintainability:** XML task definitions can be version controlled and easily imported/exported
5. **Error Handling:** Exit codes enable Task Scheduler to detect failures and trigger alerts
6. **Environment Isolation:** Virtual environment ensures consistent Python dependencies across executions
7. **Unattended Operation:** Runs 24/7 without user intervention, processes files as they arrive

## Deployment Process

1. Place batch files and Python scripts in `C:\Users\Public\RPA\code\PSTN Migration\`
2. Import XML task definitions: `schtasks /create /xml <task_file>.xml /tn "<TaskName>"`
3. Verify task appears in Task Scheduler Library
4. Test manually: Right-click task → Run
5. Monitor Task History for successful executions

## Monitoring and Troubleshooting

- **Task History:** Task Scheduler → Task → History tab shows all executions, exit codes, and errors
- **Last Run Result:** `0x0` indicates success; other codes indicate failures
- **Log Files:** Python scripts write detailed logs for troubleshooting
- **Manual Testing:** Batch files can be double-clicked for manual execution during development
