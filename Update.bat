@echo off
setlocal
cd /d "%~dp0"

echo Provjeravam nove izmjene...
git pull --ff-only
if errorlevel 1 (
    echo.
    echo Automatsko azuriranje nije uspjelo. Ako imas lokalne izmjene u ovom
    echo folderu, sacuvaj ih ili ih odbaci, pa pokreni Update.bat ponovo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Kreiram virtuelno okruzenje po prvi put...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Python 3 nije pronadjen. Instaliraj ga sa python.org i pokreni ponovo.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt
python -m playwright install chromium

echo.
echo Azuriranje zavrseno.
pause
