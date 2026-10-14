@echo off
setlocal
cd /d "%~dp0"
set "PYTHONHOME="
set "PYTHONPATH="
set "SILENT=0"
if /I "%~1"=="--silent" set "SILENT=1"

if not exist "requirements.txt" (
  echo ERRORE: mancano i file del progetto.
  echo.
  echo Probabilmente avvia-windows.bat e' stato aperto direttamente dal file ZIP.
  echo Chiudi questa finestra, fai clic destro sullo ZIP, scegli "Estrai tutto"
  echo e avvia nuovamente il file dalla cartella estratta.
  if "%SILENT%"=="0" pause
  exit /b 1
)

powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }" >nul 2>nul
if not errorlevel 1 (
  powershell -NoProfile -Command "try { $pagina=Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/' -TimeoutSec 2; if ($pagina.Content -match 'Magazzino Elettronica') { exit 0 } } catch {}; exit 1" >nul 2>nul
  if not errorlevel 1 (
    echo Il gestionale e' gia' in esecuzione. Apertura del browser...
    start "" "http://localhost:3000/"
    exit /b 0
  )
  echo ERRORE: la porta 3000 e' gia' utilizzata da un programma.
  echo.
  echo Chiudere il programma che usa la porta 3000 e riprovare.
  if "%SILENT%"=="0" pause
  exit /b 1
)

where py >nul 2>nul
if errorlevel 1 (
  echo Python non e' installato o non e' disponibile nel PATH.
  echo Installa Python 3.11 o successivo da python.org e riprova.
  if "%SILENT%"=="0" pause
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

echo Avvio del gestionale versione 1.4.0...
if "%SILENT%"=="0" echo Non chiudere questa finestra durante l'uso.
start "" powershell -NoProfile -WindowStyle Hidden -Command "$fine=(Get-Date).AddSeconds(30); while ((Get-Date) -lt $fine) { try { Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/' -TimeoutSec 1 | Out-Null; Start-Process 'http://localhost:3000/'; exit 0 } catch { Start-Sleep -Milliseconds 500 } }"
".venv\Scripts\python.exe" serve.py
if errorlevel 1 goto :errore
exit /b 0

:errore
echo.
echo Installazione o avvio non riusciti. Controlla i messaggi sopra.
if "%SILENT%"=="0" pause
exit /b 1
