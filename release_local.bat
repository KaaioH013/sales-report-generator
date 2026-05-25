@echo off
setlocal

REM Gera o EXE e empacota um zip versionado (uso local; não faz push).

call gerar_exe.bat
if errorlevel 1 (
  echo.
  echo Build falhou.
  exit /b 1
)

for /f "delims=" %%v in ('powershell -NoProfile -Command "(Select-String -Path gerar_relatorio.py -Pattern '^__version__\s*=\s*\"(.+)\"' | ForEach-Object { $_.Matches[0].Groups[1].Value })"') do set VER=%%v
if "%VER%"=="" set VER=dev

set OUT=release_gerar_relatorio_v%VER%.zip

powershell -NoProfile -Command "if (Test-Path '%OUT%') { Remove-Item '%OUT%' -Force }; Compress-Archive -Path 'dist\gerar_relatorio\*' -DestinationPath '%OUT%'"

echo.
echo Release local gerado: %OUT%
endlocal
pause
