@echo off
:: Este script limpa, cria/ativa um venv local e constrói o programa (modo diretório).

echo.
echo [PASSO 1 de 4] Limpando arquivos e pastas da versao anterior...

:: Apaga a pasta 'dist' se ela existir, em modo silencioso.
IF EXIST dist rd /s /q dist
IF EXIST build rd /s /q build

echo Limpeza concluida!
echo.
echo [PASSO 2 de 4] Verificando se os dados do mapa simplificado existem...
IF NOT EXIST "mapa_brasil_dados.pkl" (
    echo.
    echo ATENCAO: O arquivo de mapa simplificado 'mapa_brasil_dados.pkl' nao foi encontrado.
    echo Execute o script 'converter_mapa.py' primeiro para cria-lo.
    echo.
    pause
    exit /b
)
echo Dados do mapa simplificado encontrados!
echo.
echo [PASSO 3 de 4] Criando/ativando venv e instalando dependencias...

IF NOT EXIST ".venv\Scripts\python.exe" (
    echo Criando ambiente virtual em .venv...
    py -3 -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements-dev.txt

echo.
echo [PASSO 4 de 4] Construindo o EXE com PyInstaller (spec)...

pyinstaller --clean gerar_relatorio.spec

echo.
echo Processo finalizado!
echo O novo programa esta na pasta 'dist\gerar_relatorio'.
echo Para distribuir, compacte a pasta 'gerar_relatorio' em um arquivo .zip.
echo.

:: Mantem a janela aberta para voce poder ler o resultado.
pause
# --- FIM DO CÓDIGO ---