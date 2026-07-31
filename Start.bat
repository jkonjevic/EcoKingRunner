@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Prvo pokretanje: instaliram aplikaciju.
    call Update.bat
)

echo Provjeravam nove izmjene...
git pull --ff-only >nul 2>&1

call ".venv\Scripts\activate.bat"
echo Pokrecem EcoKing...
echo Ako nesto ne radi nakon azuriranja, zatvori ovaj prozor i pokreni Update.bat.
echo.
python -m ecoking.webapp

pause
