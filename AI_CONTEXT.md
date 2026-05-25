# AI CONTEXT — Gerador de Relatórios (2026)

Data deste contexto: 2026-05-05 (alinhado a `README.md` / `ROADMAP.md`)

## 1) CONTEXTO GERAL
- Projeto: app desktop **Windows** (Python + **CustomTkinter**)
- Finalidade: gerar relatórios a partir de Excel consolidado
  - Relatórios: Período / Comparativo / Anual / Comparativo Anual (**PDF**, **PDF+HTML** ou **somente HTML**)
  - Relatório **Qualidade de Dados**: **somente HTML**
- Ambiente recomendado: **Python 3.12**, execução local

## 2) OBJETIVO ATUAL
- Evolução incremental e segura (mudanças pequenas + backup antes)
- Resumo executivo de **Período**: **meta + YoY** (sem “projeção do mês” no relatório atual)
- Manter relatório de **Qualidade** útil sem pesar a UI

## 3) DECISÕES TÉCNICAS (NÃO REAVALIAR AUTOMATICAMENTE)
- UI: **CustomTkinter** (`gerar_relatorio.py`)
- Lógica: **`relatorio_servicos.py`** (sem Tkinter — PDF/HTML, pandas, templates, Excel, settings)
- Performance: **lazy import** de pandas/numpy/matplotlib/seaborn no módulo de serviços
- Validação do Excel: openpyxl em thread (UI com barra indeterminada)
- Saídas: `pdf`, `pdf_html`, `html` persistidas em settings
- PDF: **WeasyPrint** via subprocesso (`_write_pdf_safe`); falha → **fallback HTML** (ver serviços)
- Windows PDF: GTK/Pango/Cairo (+ VC++ Redistributable) — ver `README.md`
- Produtos: agregação por **`Cod.Material`**; descrição para label; sem código onde obrigatório → bloqueio
- Persistência: `%APPDATA%\relatorios_2026\settings.json` (`last_dir`, `last_output_path`, `output_mode`, `indicators`)
- Build EXE: PyInstaller + `gerar_relatorio.spec`; setuptools **abaixo da versão 82** por causa de `pkg_resources` (ver `requirements-dev.txt`)
- Segurança operacional: `backup_zip.bat` / botão **Backup**

## 4) ESTRUTURA RELEVANTE
- `gerar_relatorio.py`: UI (`App`), entry `__main__`
- `relatorio_servicos.py`: geração e utilitários
- `template*.html`: Jinja2
- `backup_zip.bat`, `backup_full.bat`, `tools/backup_projeto.ps1`
- `tests/`, `pytest.ini`, `requirements-dev.txt`
- `README.md`, `ROADMAP.md`

## 5) PREMISSAS IMPORTANTES
- Mudanças pequenas; backup antes de alterações grandes
- Não carregar libs pesadas na import da UI além do necessário
- Qualidade de dados: colunas mínimas conforme relatório (ex.: `Dt.Pedido`, `Vlr.Total`, `Nro.Pedido` onde aplicável)

## 6) ESTADO ATUAL (IMPLEMENTADO) — RESUMO
- Abas e fluxos descritos no `README.md`
- Fallback PDF→HTML quando WeasyPrint/ambiente falha
- Testes pytest em `tests/` (ex.: `_produto_label_map`)
- EXE: pasta `dist\gerar_relatorio\` (distribuir inteira)

### Histórico compacto (não usar como fonte única)
- Evoluções antigas: ttk/`ttkthemes`, projeção do mês no resumo executivo, app monolítico em um único `.py`, notas Python 3.14 — **substituídos** pela arquitetura e decisões das seções 1–4 acima.
- Backups nomeados em `backups/` podem existir; datas e ZIPs antigos não precisam refletir o ambiente atual.

## 7) COMO RODAR (REFERÊNCIA)
```powershell
cd <pasta>\projeto_relatorio
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python gerar_relatorio.py
```

Rebuild EXE: `gerar_exe.bat` ou `pip install -r requirements-dev.txt` + `py -3.12 -m PyInstaller gerar_relatorio.spec --noconfirm`

## 8) PRÓXIMOS PASSOS (NÃO INICIADOS / PARCIAIS)
- Exportar linhas problemáticas do relatório de Qualidade (CSV/XLSX)
- Validação prévia ainda mais rica (já há Validar + colunas; pode evoluir resumo)
- Timeout do PDF configurável no `settings.json`
- Cache de agregações na geração
