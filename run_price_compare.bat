@echo off
setlocal

REM 1. Check if python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python from https://www.python.org/downloads/
    start https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 2. Check if pip is installed
python -m pip --version >nul 2>nul
if %errorlevel% neq 0 (
    echo pip is not installed. Attempting to install pip...
    python -m ensurepip --upgrade
    if %errorlevel% neq 0 (
        echo Failed to install pip. Please install pip manually.
        pause
        exit /b 1
    )
)

REM 3. Check if requirements are installed
python -m pip show pandas >nul 2>nul
if %errorlevel% neq 0 (
    echo Installing required packages...
    python -m pip install -r requirements.txt
)

REM 4. Run the application
python price_compare_gui.py

endlocal
pause 