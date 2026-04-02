@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d %~dp0

if not exist .env (
    echo .env file not found.
    echo Copy .env.example to .env and fill in GITHUB_USERNAME and GITHUB_TOKEN.
    pause
    exit /b 1
)

for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    if /I "%%A"=="GITHUB_USERNAME" set "GITHUB_USERNAME=%%B"
    if /I "%%A"=="GITHUB_TOKEN" set "GITHUB_TOKEN=%%B"
)

if "%GITHUB_USERNAME%"=="" (
    echo GITHUB_USERNAME is empty in .env
    pause
    exit /b 1
)

if "%GITHUB_TOKEN%"=="" (
    echo GITHUB_TOKEN is empty in .env
    pause
    exit /b 1
)

git remote get-url origin >nul 2>nul
if errorlevel 1 (
    echo Git remote origin is not configured.
    pause
    exit /b 1
)

git push https://%GITHUB_USERNAME%:%GITHUB_TOKEN%@github.com/EviKha/speech_transcriber.git main
if errorlevel 1 (
    echo Push failed.
    pause
    exit /b 1
)

echo Push completed.
pause
