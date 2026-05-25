# Roadmap / Histórico (Gerador de Relatórios 2026)

Este arquivo existe para não depender do chat.
Aqui fica o **histórico do que já foi implementado**, decisões importantes e **para onde vamos** com mudanças incrementais (sempre com backup).

## Como trabalhar com segurança (regra do projeto)

1) Antes de qualquer mudança: rode `backup_zip.bat` (ou o botão **Backup** na barra superior).
2) Mudanças pequenas e isoladas (1 feature por vez).
3) Sempre validar pelo menos:
   - Abrir a UI sem erro
   - Gerar **Somente HTML** (modo mais estável para depurar)
   - Depois gerar **PDF** se necessário

## Estado atual (resumo)

- App: Python + **CustomTkinter** (interface escura/padrão “blue”; sem `ttkthemes`)
- Camadas: `gerar_relatorio.py` (UI) + `relatorio_servicos.py` (lógica PDF/HTML/Excel/templates/config — sem Tkinter)
- Dados/gráficos: pandas/numpy/matplotlib (+ seaborn em alguns gráficos); **lazy import** dos módulos pesados no pacote de serviços (UI abre leve)
- PDF: **WeasyPrint** (Windows: depende de GTK/Pango/Cairo no sistema; mensagens quando bibliotecas faltam)
- **Fallback:** se PDF falhar, mantém/suplementa **HTML** utilizável (`*_fallback.html` ou modo **PDF + HTML**, conforme o fluxo)
- Templates: `template*.html` com condicionais para mostrar/ocultar seções
- Produtos: agregação por **`Cod.Material`**; descrição só para label; relatórios com produtos exigem código + descrição
- Config: `%APPDATA%\relatorios_2026\settings.json`
  - `last_dir`
  - `indicators` (flags por tipo de relatório)
  - `last_output_path` (último artefato principal gerado — PDF ou HTML conforme modo)
  - `output_mode` (`pdf`, `pdf_html`, `html`)
- Testes: `tests/` + `pytest` (`pytest.ini`, `requirements-dev.txt`)
- Distribuição: PyInstaller (`gerar_relatorio.spec`), pasta `dist\gerar_relatorio\`; build dev exige setuptools **abaixo da 82** (PyInstaller/`pkg_resources`) — ver `README.md`

## Funcionalidades implementadas (marcos)

### UX / Segurança
- Backup do projeto:
  - Scripts: `tools/backup_projeto.ps1`, `backup_zip.bat`, `backup_full.bat`
  - UI: botão **Backup** + **Ver backups**
- Utilidades na UI:
  - **Copiar log**
  - **Abrir relatório** (último) + **Abrir pasta**
- Modo de saída:
  - **PDF** / **PDF + HTML** / **Somente HTML**
  - Export HTML com pasta `*_assets` para imagens e `src=` reescrito para caminho relativo

- Validação prévia do Excel (sem travar a UI):
  - Botões **Validar** + checagens automáticas antes de gerar
  - Barra indeterminada durante a validação (thread)
  - Mensagens quando faltar coluna (com “motivo” / indicador que exige)

- Cancelamento mais robusto:
  - PDF gerado em subprocesso (evita travar e permite cancelar/timeout)
  - Cancelamento checado entre etapas

- Status/UX:
  - Mostra **tempo da última geração** ao concluir
  - UI com **CustomTkinter** (abas, botões, progresso)
  - Nome exibido: **Gerador de Relatórios** (versão fica no código, não aparece na UI)

### Performance
- Lazy import de `pandas/numpy/matplotlib/seaborn` no módulo de serviços: a UI não carrega essas libs na importação inicial.

### Observabilidade (debug / performance)
- Log com **tempos por etapa** (Leitura Excel / Limpeza / Análises / Gráficos / HTML / Salvar HTML / PDF / Total).

### Conteúdo
- Padronização de rankings: **Top 20** onde aplicável
- Ajustes de apresentação (ex.: pizza por UF com % + valor)
- Resumo executivo (período): **meta + YoY** (a **projeção do mês** foi removida do template/fluxo)

### Qualidade de Dados
- Aba **Qualidade de Dados**: relatório **somente HTML** com checagens automáticas (UF, vendedor, campos vazios, inconsistências por pedido, valores/quantidades).
- Severidade por tipo (Alta/Média/Baixa) + seção “Resumo”.

### Indicadores opcionais
- Checkboxes por aba (Período / Comparativo / Anual / Comparativo Anual)
- Persistência em `settings.json`
- Preset **Padrão** (reseta para `_default_indicator_settings()` em `relatorio_servicos.py`)
- Presets **Completo** e **Enxuto**
- Período: `resumo_exec` liga/desliga o **Resumo Executivo** (**meta + YoY**)

## Defaults oficiais (referência)

Os valores padrão ficam em **`_default_indicator_settings()`** em `relatorio_servicos.py`.

## Próximas evoluções (ordem mais segura)

1) Validação prévia do Excel (antes de gerar)  
   Status: **implementado** (Validar + colunas obrigatórias + mensagens por indicador).

2) Performance na abertura (lazy import)  
   Status: **implementado** no serviços; manter esse padrão em novos imports pesados.

3) Performance na geração  
   - Cache/evitar recomputar agregações repetidas  
   - Profiling leve

4) Configs avançadas (sem pesar a UI)  
   - Timeout do PDF configurável (`settings.json`)  
   - “Modo rápido” já parcialmente coberto pelo preset Enxuto

5) Robustez  
   - Padronizar cancelamento em todos os pontos críticos  
   - Tratar avisos/deprecações do pandas quando conveniente

6) Qualidade de dados  
   - Exportar linhas problemáticas (CSV/XLSX) — ainda não iniciado

## Changelog (curto)

- 2026-05-05
  - Documentação: `README.md` e `ROADMAP.md` alinhados ao estado atual (CustomTkinter, split UI/serviços, PDF/GTK, fallback HTML, EXE, pytest).
  - Referência técnica: refatoração `gerar_relatorio.py` + `relatorio_servicos.py`; testes `tests/`; pin setuptools para build PyInstaller.

- 2026-02-06
  - Período: “Resumo Executivo” opcional (meta mensal fixa R$ 2.530.000 + YoY); indicador `periodo.resumo_exec`
  - *Nota posterior (2026-05): card/texto de **projeção do mês** removido do relatório de período.*
  - Backup ZIP antes da mudança: `backups/backup_projeto_20260206-155029.zip`

- 2026-02-06
  - Qualidade de Dados (HTML): severidade por problema + seção “Resumo” (ordenada por severidade e volume)
  - Backup ZIP antes da mudança: `backups/backup_projeto_20260206-152146.zip`

- 2026-02-04
  - UI: aba **Qualidade de Dados**; `Nro.Pedido` obrigatório na validação prévia
  - Tech debt: pandas `resample('M'/'Q')` → `ME/QE`

- 2026-02-04
  - PDF em subprocesso + cancelamento mais cedo (incl. Comparativo Anual)

- 2026-02-03
  - Presets **Completo** / **Enxuto** / **Padrão**; lazy import; validação Excel; spinner; tempos por etapa; UI ttk (histórico — hoje CustomTkinter)

---

Notas:
- Sempre que uma mudança for feita, adicionar 1–3 bullets em **Changelog** com data e, se aplicável, nome do backup gerado.
