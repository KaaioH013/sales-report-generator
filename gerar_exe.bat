@echo off
:: Este script limpa, ativa o ambiente e constroi o programa (modo diretório).

echo.
echo [PASSO 1 de 4] Limpando arquivos e pastas da versao anterior...

:: Apaga a pasta 'dist' se ela existir, em modo silencioso.
IF EXIST dist rd /s /q dist
IF EXIST build rd /s /q build
IF EXIST gerar_relatorio.spec del gerar_relatorio.spec

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
echo [PASSO 3 de 4] Ativando ambiente virtual e construindo o programa (modo diretorio)...

:: Ativa o ambiente virtual e executa o PyInstaller com a flag --onedir para inicializacao rapida
call venv_stable\Scripts\activate.bat && pyinstaller --windowed --onedir --add-data "template.html;." --add-data "template_comparativo.html;." --add-data "template_anual.html;." --add-data "template_comparativo_anual.html;." --add-data "Logo.png;." --add-data "mapa_brasil_dados.pkl;." --splash "splash.png" gerar_relatorio.py

echo.
echo [PASSO 4 de 4] Processo finalizado!
echo O novo programa esta na pasta 'dist\gerar_relatorio'.
echo Para distribuir, compacte a pasta 'gerar_relatorio' em um arquivo .zip.
echo.

:: Mantem a janela aberta para voce poder ler o resultado.
pause
# --- FIM DO CÓDIGO ---