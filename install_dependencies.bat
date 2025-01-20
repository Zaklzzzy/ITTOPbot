@echo off
cd /d %~dp0

REM Check venv
if not exist "venv\Scripts\python.exe" (
    echo Virtual environment not found. Creating a new one...
    python -m venv venv
    if errorlevel 1 (
        echo Failed to create virtual environment. Please check your Python installation.
        pause
        exit /b
    )
)

REM Activate venv and update pip
echo Activating virtual environment...
venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to update pip in virtual environment.
    pause
    exit /b
)

REM Download dependencies to venv
echo Installing dependencies from requirements.txt...
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies. Please check the requirements file.
    pause
    exit /b
)

echo Dependencies successfully installed in virtual environment.
pause
