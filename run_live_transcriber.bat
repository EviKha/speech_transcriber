@echo off
cd /d %~dp0

if not exist .venv (
    echo Creating virtual environment...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to upgrade pip.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies. Check the error messages above.
    pause
    exit /b 1
)

python live_transcriber.py
if errorlevel 1 (
    echo Application exited with an error.
    pause
    exit /b 1
)
