@echo off
echo ================================
echo Przygotowanie srodowiska VBS NTQ by Jurek
echo ================================

REM 1. Instalacja Python Launcher + Python 3.11 (Microsoft Store)
echo Instalacja Python 3.11...
winget install --id 9NQ7512CXL7T --source msstore --accept-package-agreements --accept-source-agreements

REM Odczekaj chwile az Python bedzie widoczny
timeout /t 5 >nul

REM 2. Aktualizacja pip
echo Aktualizacja pip...
py -3.11 -m pip install --upgrade pip

REM 3. Instalacja Playwright
echo Instalacja Playwright...
py -3.11 -m pip install playwright

REM 4. Instalacja Chromium dla Playwright
echo Instalacja Chromium...
py -3.11 -m playwright install chromium

echo ================================
echo Srodowisko gotowe. Miłej zabawy życzy Jurek =)
echo ================================
pause
