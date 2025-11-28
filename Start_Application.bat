@echo off
REM IB Question Entry System - Windows Launcher Script
REM Double-click this file to start the application

setlocal enabledelayedexpansion

REM Change to the script's directory
cd /d "%~dp0"

echo =========================================
echo IB Question Entry System
echo =========================================
echo.

REM Check if Python is installed (try both python and python3)
set PYTHON_CMD=
python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=python
) else (
    python3 --version >nul 2>&1
    if not errorlevel 1 (
        set PYTHON_CMD=python3
    )
)

if "!PYTHON_CMD!"=="" (
    echo [ERROR] Python is not installed on your computer.
    echo.
    echo Please install Python 3:
    echo   1. Visit https://www.python.org/downloads/
    echo   2. Download Python 3 for Windows
    echo   3. Run the installer
    echo   4. Make sure to check "Add Python to PATH" during installation
    echo   5. Come back and try again!
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('!PYTHON_CMD! --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [OK] Python found: !PYTHON_VERSION!
echo [OK] Using command: !PYTHON_CMD!
echo.

REM Check if app.py exists
if not exist "app.py" (
    echo [ERROR] Cannot find app.py in this folder.
    echo Current directory: %CD%
    echo.
    pause
    exit /b 1
)

REM Check if required packages are installed
echo Checking for required packages...
!PYTHON_CMD! -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [WARNING] Required packages are not installed.
    echo Installing required packages (this may take a minute)...
    echo.
    
    REM Install requirements
    !PYTHON_CMD! -m pip install -r requirements.txt
    
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install required packages.
        echo.
        echo Please try installing manually:
        echo   1. Open Command Prompt
        echo   2. Navigate to this folder
        echo   3. Run: !PYTHON_CMD! -m pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    
    echo.
    echo [OK] Packages installed successfully!
    echo.
)

echo [OK] All requirements are ready!
echo.
echo =========================================
echo Starting application...
echo =========================================
echo.
echo NOTE: This window will stay open while the application is running.
echo       The application window should appear shortly.
echo       Close the application window when you're done.
echo.
echo Launching...
echo.

REM Run the application
!PYTHON_CMD! app.py

REM Check the exit code
set EXIT_CODE=!ERRORLEVEL!

REM Keep window open to see any errors
echo.
if !EXIT_CODE! NEQ 0 (
    echo =========================================
    echo [ERROR] The application encountered an error.
    echo Exit code: !EXIT_CODE!
    echo =========================================
    echo.
    echo Common issues:
    echo   - Python might not be properly installed
    echo   - PyQt6 might not be installed correctly
    echo   - There might be an error in the application
    echo.
    echo Please check any error messages above.
    echo If the problem persists, contact technical support.
    echo.
) else (
    echo Application closed normally.
    echo.
)

echo Press any key to close this window...
pause >nul

