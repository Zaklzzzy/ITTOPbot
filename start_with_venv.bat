@echo off
cd /d %~dp0

REM Check venv
if not exist "venv\Scripts\python.exe" (
    echo Virtual environment not found. Creating a new one...
    python -m venv venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        echo Run "start_global_install.bat" to install dependencies globally.
        pause
        exit /b
    )
)

REM Activate venv and update pip
echo Activating virtual environment...
venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to update pip.
    echo Run "start_global_install.bat" to install dependencies globally.
    pause
    exit /b
)

REM Download dependecies to venv
echo Installing dependencies...
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies.
    echo Run "start_global_install.bat" to install dependencies globally.
    pause
    exit /b
)

REM Start bot
echo Starting application...
venv\Scripts\python.exe scripts\main.py
if errorlevel 1 (
    echo Application failed to start.
    pause
    exit /b
)

pause
