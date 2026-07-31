@echo off
setlocal
cd /d "%~dp0"

:: Create virtual environment if it doesn't exist
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    py -3 -m venv .venv
)

:: Activate virtual environment
call ".venv\Scripts\activate.bat"

:: Set Playwright path and install dependencies
set PLAYWRIGHT_BROWSERS_PATH=0
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m playwright install chromium

:: Package application with PyInstaller. Uses EcoKingWebRunner.spec (not
:: inline flags) because it needs `excludes` -- ecoking/, ecoking_daily.py
:: and web/ must ship as loose files next to the exe, not baked into the
:: frozen archive, so ecoking.selfupdate can refresh them from GitHub.
pyinstaller --noconfirm --clean EcoKingWebRunner.spec

:: PyInstaller puts `datas` inside dist\EcoKingWebRunner\_internal\, but the
:: app looks for them next to the exe (app_root() = the exe's own folder,
:: same place ecoking.selfupdate writes updates to). Move them out so both
:: agree on one location.
set OUT=dist\EcoKingWebRunner
for %%F in (ecoking ecoking_daily.py web requirements.txt ".env" stations.json "ECO KING BLANKO TABLICA.xlsx") do (
    if exist "%OUT%\_internal\%%~F" (
        if exist "%OUT%\%%~F" rmdir /s /q "%OUT%\%%~F" 2>nul
        if exist "%OUT%\%%~F" del /q "%OUT%\%%~F" 2>nul
        move /y "%OUT%\_internal\%%~F" "%OUT%\%%~F" >nul
    )
)

echo.
echo Build complete: dist\EcoKingWebRunner\EcoKingWebRunner.exe
echo Zip the entire dist\EcoKingWebRunner folder before sending -- the app
echo needs everything in it (loose source files, Chromium, .env).
pause
