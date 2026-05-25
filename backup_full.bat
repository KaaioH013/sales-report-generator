@echo off
setlocal
set ROOT=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%tools\backup_projeto.ps1" -ProjectRoot "%ROOT%" -WipCommit
echo.
echo Backup ZIP + Git (branch/tag) finalizado.
pause
