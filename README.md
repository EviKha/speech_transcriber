# Speech Transcriber for Windows 11

This folder now contains two tools:

- `transcribe.py`: transcription of a ready audio or video file
- `live_transcriber.py`: live transcription during a meeting with an on-screen text window and TXT autosave

## Live meeting mode

`live_transcriber.py` is the main app for browser meetings.

It provides:

- a live window where text appears while people speak
- `Meeting` lines from system output, for example Zoom or Pачка in the browser
- `You` lines from your default microphone
- autosave into `txt`
- audio autosave into `wav` for replay
- export into `docx`
- search inside the live transcript
- line markers: `Important`, `Task`, `Question`
- optional meeting speaker split into `Speaker 1`, `Speaker 2`, ...
- one-click summary window after or during the meeting
- replay of a saved audio fragment by clicking the transcript line
- log file and error report export for debugging
- faster live updates every few seconds

Important:

- to capture other participants, the meeting audio must play through this PC
- to capture your own speech, the correct default microphone must be selected in Windows
- this is a practical MVP, not a perfect courtroom-grade transcript system

## What the tools do

- Takes `mp3`, `wav`, `m4a`, `mp4` and other media files.
- Recognizes speech locally on your PC.
- Saves the transcript into `txt` or `docx`.
- Can optionally add timestamps.

## 1. Install Python

If Python is not installed, install `Python 3.11+` from:

- https://www.python.org/downloads/windows/

During installation, enable `Add Python to PATH`.

## 2. Install ffmpeg

`ffmpeg` is required for most audio and video formats.

On Windows 11, the easiest way is:

```powershell
winget install Gyan.FFmpeg
```

Then close and reopen PowerShell.

Check that it works:

```powershell
ffmpeg -version
```

## 3. Install dependencies

Open PowerShell in this folder and run:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If `soundcard` fails to install on your machine, first upgrade packaging tools:

```powershell
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

If you see an error like `fromstring is removed, use frombuffer instead`, reinstall `numpy` in the current virtual environment:

```powershell
pip uninstall -y numpy
pip install "numpy<2"
pip install -r requirements.txt
```

`resemblyzer` is included in `requirements.txt`. It is used only if you enable `Split meeting speakers`.

## 4. Run live transcription

Fastest option:

```powershell
.\run_live_transcriber.bat
```

Or manually:

```powershell
.venv\Scripts\Activate.ps1
python .\live_transcriber.py
```

How to use it:

1. Start your browser meeting.
2. Make sure you hear the meeting through your Windows audio output.
3. Open the app.
4. Pick the TXT file path if needed.
5. Click `Start`.
6. Watch the live text in the app window.
7. If you want the app to split the meeting audio into `Speaker 1`, `Speaker 2`, ..., enable `Split meeting speakers`. Leave it off if you want the old `Meeting` mode.
8. Click any transcript line once to select it.
9. Click the same line again, or click `Play Selected`, to hear the saved audio fragment for that line.
10. Use `Important`, `Task`, `Question` to mark selected lines.
11. Use `Search` to find words in the transcript.
12. Click `Summary` to open a short meeting summary.
13. Click `Stop` when the meeting ends.
14. Click `Export DOCX` if you want a Word document with the transcript and summary.

The app writes:

- `Meeting`: speech captured from system output
- `You`: speech captured from your microphone
- `live_audio\*.wav`: saved meeting audio for transcript replay

If something fails:

- the app shows the log file path
- click `Save Error Report`
- send me the saved report and the latest file from the `logs` folder

Notes about speaker split:

- the feature is optional and can be disabled with the `Split meeting speakers` checkbox
- it only applies to the `Meeting` stream, not to `You`
- the split is approximate and depends on audio quality and speaker overlap
- if you do not need it, leave it off and the app will keep the simpler `Meeting` / `You` mode

Notes about audio replay:

- the app tries several playback methods on Windows
- if one method fails, it automatically falls back to another one
- `ffmpeg`/`ffplay` should stay installed because replay may use `ffplay` as a fallback

## 5. Build EXE

To build a regular Windows executable:

```powershell
.\build_exe.bat
```

After that, the app will be here:

```text
dist\LiveMeetingTranscriber\LiveMeetingTranscriber.exe
```

## 6. Run file transcription

Basic example:

```powershell
python .\transcribe.py "C:\Meetings\call.mp3"
```

This will create:

```text
C:\Meetings\call.txt
```

Save to DOCX:

```powershell
python .\transcribe.py "C:\Meetings\call.mp3" --format docx
```

Force Russian language:

```powershell
python .\transcribe.py "C:\Meetings\call.mp3" --language ru
```

Add timestamps:

```powershell
python .\transcribe.py "C:\Meetings\call.mp3" --with-timestamps
```

Choose a bigger model for better quality:

```powershell
python .\transcribe.py "C:\Meetings\call.mp3" --model medium
```

## Recommended models

- `small`: good default balance of speed and quality
- `medium`: better quality, slower
- `large-v3`: best quality, much heavier

For an average Windows laptop, start with `small`.

## Notes for browser meetings

For live meetings, prefer `live_transcriber.py`.

If you want maximum accuracy after the call, also save the meeting recording and run `transcribe.py` on the recorded file later.

If you do not have a recording yet, the simplest flow is:

1. Record the meeting audio.
2. Save to `mp3` or `m4a`.
3. Run `transcribe.py`.

## Typical command

```powershell
python .\transcribe.py "C:\Meetings\zoom_recording.m4a" --language ru --format docx --model small
```
