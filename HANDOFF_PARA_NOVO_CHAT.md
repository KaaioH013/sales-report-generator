# Handoff para novo chat — Gerador de Relatórios

Data: 2026-02-06

## O que você deve colar no novo chat (copie/cole)

Estou num projeto Python/Tkinter no Windows em `C:\Users\caio.santana\Desktop\Dash\projeto_relatorio`.

Objetivo do app: gerar relatórios (Período / Comparativo / Anual / Comparativo Anual) em PDF/HTML e agora também um relatório **Qualidade de Dados** (somente HTML).

O que já foi feito (importante):
- UI Tkinter/ttk modernizada; nome exibido “Gerador de Relatórios”.
- Lazy import de pandas/numpy/matplotlib/seaborn pra abrir rápido.
- Validação prévia do Excel via openpyxl com spinner em thread.
- Modos de saída: PDF / PDF+HTML / Somente HTML.
- PDF via WeasyPrint em subprocesso (`_write_pdf_safe`) com cancelamento/timeout (resolveu travar em ~85%).
- Persistência em `%APPDATA%\\relatorios_2026\\settings.json` (last_dir, last_output_path, output_mode, indicators).
- Nova aba **Qualidade de Dados**: gera HTML com checagens (UF inválida, vendedor ausente/não encontrado, campos vazios, inconsistência por `Nro.Pedido`, valores/qtde 0/negativo). Função: `gerar_relatorio_qualidade_dados`.
- Para validação base, agora exigimos `Dt.Pedido`, `Vlr.Total`, `Nro.Pedido`.

Teste já feito pelo usuário: relatório de qualidade gerou OK.

Arquivos de contexto no repo:
- `HISTORICO_DO_CHAT_2026-02-04.md`
- `ROADMAP.md`
- `README.md`

Próximo passo desejado: continuar evoluindo “aos poucos”, sempre com backup antes.

## Como rodar o app agora

Pelo venv estável:
```powershell
cd C:\Users\caio.santana\Desktop\Dash\projeto_relatorio
.\venv_stable\Scripts\python.exe .\gerar_relatorio.py
```

## Onde mexer
- Feature principal está em `gerar_relatorio.py` (UI + geração de relatórios).

## Preferências/regras
- Mudanças pequenas, sem quebrar, sempre com backup ZIP antes.
- Manter o app leve.
