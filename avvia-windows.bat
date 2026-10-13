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
  echo ERRORE: la porta 3000 e' gia' utilizzata da un programma.
  echo.
  echo Se e' aperta una precedente finestra del gestionale, chiudila per
  echo caricare gli aggiornamenti, poi avvia nuovamente questo file.
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

echo Avvio del gestionale versione 1.3.0...
echo Non chiudere questa finestra durante l'uso.
start "" powershell -NoProfile -WindowStyle Hidden -Command "$fine=(Get-Date).AddSeconds(30); while ((Get-Date) -lt $fine) { try { Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/' -TimeoutSec 1 | Out-Null; Start-Process 'http://localhost:3000/'; exit 0 } catch { Start-Sleep -Milliseconds 500 } }"
".venv\Scripts\python.exe" app.py
if errorlevel 1 goto :errore
exit /b 0

:errore
echo.
echo Installazione o avvio non riusciti. Controlla i messaggi sopra.
pause
exit /b 1
