# Gerador de Relatórios (2026)

Aplicação em Python (Tkinter) para gerar relatórios em PDF a partir de um Excel consolidado.

## O que este projeto faz

- Relatório de período (KPIs, gráficos, mapa de calor por UF)
- Relatório comparativo (período A vs período B, churn/inativos)
- Análise anual
- Comparativo anual

## Requisitos

- Windows 10/11
- Python 3.11+ recomendado

> Observação: o PDF é gerado via `weasyprint`. Em alguns ambientes Windows, podem ser necessárias dependências de sistema. Se sua instalação falhar, me avise que eu te passo o checklist específico.

## Rodando pelo código (mais simples)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python gerar_relatorio.py
```

## Gerando o EXE (PyInstaller)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
pyinstaller --clean gerar_relatorio.spec
```

Ou use o script pronto:

```bat
gerar_exe.bat
```

O resultado fica em `dist\gerar_relatorio\gerar_relatorio.exe`.

## Arquivos importantes

- `gerar_relatorio.py`: aplicação
- `template*.html`: templates do relatório
- `Logo.png` e `splash.png`: imagens
- `mapa_brasil_dados.pkl`: dados do mapa (necessário para o mapa de calor)

## Notas sobre o EXE

- O programa salva os gráficos em uma pasta temporária durante a geração do PDF (isso evita tentar escrever dentro do bundle do PyInstaller).
- Os gráficos não ficam persistidos após o PDF ser gerado.
