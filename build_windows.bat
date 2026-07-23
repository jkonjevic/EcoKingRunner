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

:: Package application with PyInstaller
pyinstaller --noconfirm --clean --onedir --windowed ^
    --name EcoKingRunner ^
    --add-data ".env;." ^
    --add-data "herceg_novi_stations.json;." ^
    --hidden-import ecoking_daily ^
    --collect-all playwright ^
    ecoking_launcher.py

:: Backup step: Ensure .env and json files are explicitly copied into dist\EcoKingRunner
if exist ".env" copy /Y ".env" "dist\EcoKingRunner\.env"
if exist "herceg_novi_stations.json" copy /Y "herceg_novi_stations.json" "dist\EcoKingRunner\herceg_novi_stations.json"

echo Build complete! Make sure to zip the entire dist\EcoKingRunner folder before sending.
pause