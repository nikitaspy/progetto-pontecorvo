@echo off
cd /d "%~dp0"
if not exist "instance" mkdir "instance"
call avvia-windows.bat --silent >> "instance\avvio.log" 2>&1
