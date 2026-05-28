@echo off
REM ============================================================================
REM Lance l'app Streamlit en local (Windows)
REM ============================================================================
REM Usage : double-clic sur ce fichier
REM ============================================================================

cd /d "%~dp0"

set ROOT_DIR=%CD%\..
set VENV_DIR=%ROOT_DIR%\venv
set PORT=8501

REM Si un ancien serveur Streamlit traine sur le port, on le tue
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING 2^>nul') do (
    echo ^>^>^> Arret de l'ancien serveur Streamlit ^(PID %%a^) ...
    taskkill /F /PID %%a 2>nul
)

REM Crée le venv si inexistant
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo ^>^>^> Creation du venv local ^(1ere fois, ~30s^) ...
    python -m venv "%VENV_DIR%"
)

REM Active le venv
call "%VENV_DIR%\Scripts\activate.bat"

REM Installe / met a jour les dependances
echo ^>^>^> Verification des dependances ...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt

echo.
echo ============================================================
echo   L'app va s'ouvrir dans ton navigateur sur localhost:%PORT%
echo   Pour arreter : Ctrl+C dans cette fenetre
echo ============================================================
echo.
streamlit run app.py --server.port %PORT%

pause
