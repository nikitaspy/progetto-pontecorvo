@echo off
setlocal
cd /d "%~dp0"

powershell -NoProfile -Command "$connessioni=Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue; if (-not $connessioni) { Write-Host 'Il gestionale non risulta in esecuzione.'; exit 0 }; try { $pagina=Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/' -TimeoutSec 2 } catch { Write-Error 'La porta 3000 e occupata, ma non risponde il gestionale.'; exit 1 }; if ($pagina.Content -notmatch 'Magazzino Elettronica') { Write-Error 'La porta 3000 appartiene a un altro programma: arresto annullato.'; exit 1 }; $connessioni.OwningProcess | Sort-Object -Unique | ForEach-Object { Stop-Process -Id $_ -Force }; Write-Host 'Gestionale arrestato.'"

if errorlevel 1 (
  echo.
  echo Arresto non riuscito. Controlla il messaggio sopra.
)
pause
