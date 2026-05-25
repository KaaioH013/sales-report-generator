@echo off
setlocal
set ROOT=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%tools\backup_projeto.ps1" -ProjectRoot "%ROOT%" -ZipOnly
echo.
echo Backup ZIP finalizado.
pause
