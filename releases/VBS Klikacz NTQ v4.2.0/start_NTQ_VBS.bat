@echo off
setlocal enabledelayedexpansion
title NTQ Intermodal: VBS

REM === KONFIGURACJA ===
set "PROFILE=%USERPROFILE%\chrome-vbs"
set "URL=https://ebrama.baltichub.com/login"
set "DEBUG_PORT=9222"

REM EXE w tym samym folderze co BAT
set "BASEDIR=%~dp0"
set "EXE=VBS Klikacz NTQ v4.2.0.exe"
set "EXEPATH=%BASEDIR%%EXE%"

echo [1/3] Uruchamiam Chrome...
start "" chrome ^
  --remote-debugging-port=%DEBUG_PORT% ^
  --user-data-dir="%PROFILE%" ^
  "%URL%"

timeout /t 5 /nobreak >nul

echo [2/3] Uruchamiam program...
if not exist "%EXEPATH%" (
  echo [BLAD] Nie znaleziono pliku: "%EXEPATH%"
  echo Upewnij sie, ze EXE jest w tym samym folderze co ten BAT.
  pause
  exit /b 1
)

start "" "%EXEPATH%"

echo [3/3] Gotowe.
pause
