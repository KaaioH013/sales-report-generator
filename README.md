# Gerador de Relatórios (2026)

Aplicação em Python com **CustomTkinter** para gerar relatórios em **PDF** e/ou **HTML** a partir de Excel (`.xlsx`).

> Histórico e ideias futuras: [ROADMAP.md](ROADMAP.md).

## O que este projeto faz

- Relatório de **período** (KPIs, gráficos, mapa de calor por UF)
- Relatório **comparativo** (período A vs período B, churn/inativos)
- **Análise anual**
- **Comparativo anual**
- **Qualidade de dados** (somente HTML — checagens automáticas)
- Resumo executivo no relatório de **período** (opcional): meta mensal fixa + comparação **YoY**
- Produtos agregados por **`Cod.Material`** (descrição usada só para exibição; falta de código bloqueia geração quando produtos são exigidos)

Extras: validação prévia do Excel, cancelamento durante a geração, backup em ZIP pelo app e tempos por etapa no log.

## Estrutura do código

| Arquivo | Função |
|--------|--------|
| `gerar_relatorio.py` | Entrada da aplicação: janela, abas e fluxo da UI (`App`). |
| `relatorio_servicos.py` | Geração de relatórios, PDF/HTML, leitura de Excel, templates, configurações em `%APPDATA%`, etc. (sem Tkinter). |
| `template*.html` | Modelos Jinja2 dos relatórios |
| `Logo.png`, `splash.png` | Marca e splash do PyInstaller |
| `mapa_brasil_dados.pkl` | Geometrias simplificadas para o mapa de calor |

## Requisitos

- **Windows** 10 ou 11
- **Python 3.12** recomendado (evita problemas de build com dependências como `pandas` em versões muito novas)

### PDF no Windows (WeasyPrint)

O PDF é gerado com **WeasyPrint**, que no Windows depende de bibliotecas nativas (**GTK / Pango / Cairo**) e, em geral, do **Microsoft Visual C++ Redistributable (2015–2022)**.

1. Instale um **GTK3 Runtime** para Windows (ex.: pacote que disponibilize as DLLs em `Program Files\GTK3-Runtime Win64\bin` ou equivalente).
2. Garanta o **VC++ Redistributable** instalado.

Se algo faltar, o app continua utilizável: na geração, o código tenta o PDF e, **se falhar**, grava/atualiza **HTML** ao lado (`pdf_html`/modo combinado ou arquivo `*_fallback.html`), para abrir no navegador.

### Rodando a partir do código

Na pasta `projeto_relatorio`:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python gerar_relatorio.py
```

Opcionalmente, sem novo console persistente:

```powershell
pythonw gerar_relatorio.py
```

> Use sempre o Python do **`.venv`** do projeto ao rodar ou instalar pacotes, para não misturar com o Python do sistema.

### Testes automatizados

Com dependências de desenvolvimento:

```powershell
pip install -r requirements-dev.txt
py -3.12 -m pytest tests -q
```

## Gerando o EXE (PyInstaller)

O build usa o arquivo **`gerar_relatorio.spec`** (modo diretório: uma pasta com o `.exe` e DLLs).

### Opção A — script `gerar_exe.bat`

```bat
gerar_exe.bat
```

(Limpa `build`/`dist`, verifica o `.pkl` do mapa, cria/atualiza `.venv`, instala `requirements-dev.txt` e roda o PyInstaller.)

### Opção B — manual

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
py -3.12 -m PyInstaller gerar_relatorio.spec --noconfirm
```

**Saída:** `dist\gerar_relatorio\gerar_relatorio.exe` — **distribua a pasta inteira** `dist\gerar_relatorio\` (zipada), não só o `.exe`.

### Notas de build

- Em `requirements-dev.txt` o **setuptools** fica restrito a versões **abaixo da 82** porque o PyInstaller 6.x ainda depende de `pkg_resources`; com setuptools mais novo o build pode falhar até o ecossistema atualizar.
- No **PC de destino**, PDF pode continuar exigindo **GTK / VC++** instalados como no desenvolvimento; se não estiverem disponíveis, use modo de saída **somente HTML** ou confie no **fallback HTML** quando o PDF falhar.

## Indicadores (checkboxes)

- No relatório de **período**, o resumo executivo (meta + YoY) pode ser ligado ou desligado.
- **Qualidade de dados** (HTML): severidade por problema e seção “Resumo”.

## Cancelamento

- O botão **Cancelar** interrompe na próxima etapa segura.
- A geração de PDF roda em **subprocesso** (UI responsiva, timeout e cancelamento).

## Backup (antes de mudanças grandes)

- ZIP rápido (`backups\` + export de `settings.json` do `%APPDATA%` se existir): `backup_zip.bat`
- Snapshot Git + tag: `backup_full.bat`

## Notas sobre o EXE em execução

- Gráficos e assets temporários vão para pastas temporárias durante a geração (não se grava dentro do bundle do PyInstaller).
- Após fechar o app, esses arquivos temporários não precisam ficar no disco; o PDF/HTML final é o que você salva no diálogo.
