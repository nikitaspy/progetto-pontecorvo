@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python non e' installato o non e' disponibile nel PATH.
  echo Installa Python 3.11 o successivo da python.org e riprova.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Preparazione iniziale dell'ambiente...
  py -m venv .venv
  if errorlevel 1 goto :errore
)

echo Verifica delle dipendenze...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :errore

echo Avvio del gestionale...
start "Magazzino Cannizzaro" ".venv\Scripts\python.exe" app.py
timeout /t 3 /nobreak >nul
start "" "http://localhost:3000/setup"
exit /b 0

:errore
echo.
echo Installazione non riuscita. Controlla i messaggi sopra.
pause
exit /b 1
