@echo off
cd /d %~dp0

if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name LiveMeetingTranscriber ^
  live_transcriber.py

echo.
echo EXE created in: %~dp0dist\LiveMeetingTranscriber
pause
