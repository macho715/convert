@echo off
setlocal
echo [AGI_TR6] Local Python Test Runner
echo.

REM Check if READY.xlsm exists, otherwise use base xlsx
set "INPUT_FILE=AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm"
if not exist "%INPUT_FILE%" (
    set "INPUT_FILE=AGI_TR6_VBA_Enhanced_AUTOMATION.xlsx"
    echo Using base xlsx (READY.xlsm not found)
)

REM Setup venv if not exists
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
)

REM Activate and install
call .venv\Scripts\activate
if errorlevel 1 (
    echo ERROR: Failed to activate venv
    pause
    exit /b 1
)

echo Installing requirements...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install requirements
    pause
    exit /b 1
)

REM Create output directory
if not exist ".\out" mkdir ".\out"

REM Run update
echo.
echo Running agi_tr_runner.py...
python agi_tr_runner.py --in "%INPUT_FILE%" --out ".\out" --mode update
if errorlevel 1 (
    echo ERROR: Python script failed
    pause
    exit /b 1
)

REM Run pipeline
echo.
echo Running tr6_pipeline.py...
python tr6_pipeline.py --in "%INPUT_FILE%" --out ".\out" --log ".\out\tr6_ops.log"
if errorlevel 1 (
    echo ERROR: Pipeline script failed
    pause
    exit /b 1
)

echo.
echo OK: All tests completed. Check .\out\ folder for results.
pause
endlocal
