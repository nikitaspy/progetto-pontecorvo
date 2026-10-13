@echo off
setlocal
cd /d "%~dp0"
set "PYTHONHOME="
set "PYTHONPATH="

if not exist "requirements.txt" (
  echo ERRORE: mancano i file del progetto.
  echo.
  echo Probabilmente avvia-windows.bat e' stato aperto direttamente dal file ZIP.
  echo Chiudi questa finestra, fai clic destro sullo ZIP, scegli "Estrai tutto"
  echo e avvia nuovamente il file dalla cartella estratta.
  pause
  exit /b 1
)

powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }" >nul 2>nul
if not errorlevel 1 (
  echo ERRORE: un gestionale e' gia' attivo sulla porta 3000.
  echo.
  echo Chiudi la precedente finestra "Magazzino Cannizzaro" per caricare
  echo gli aggiornamenti, poi avvia nuovamente questo file.
  pause
  exit /b 1
)

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

echo Avvio del gestionale versione 1.1.0...
echo Non chiudere questa finestra durante l'uso.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process 'http://localhost:3000/setup'"
".venv\Scripts\python.exe" app.py
if errorlevel 1 goto :errore
exit /b 0

:errore
echo.
echo Installazione non riuscita. Controlla i messaggi sopra.
pause
exit /b 1
