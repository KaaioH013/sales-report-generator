# --- INÍCIO DO CÓDIGO (interface) ---
# --------------------------------------------------------------------------
# PROGRAMA: GERADOR DE RELATÓRIOS (v21.0)
# --------------------------------------------------------------------------
# -*- coding: utf-8 -*-

import os
import sys
os.environ.setdefault("MPLBACKEND", "Agg")

import multiprocessing
import queue
import subprocess
import threading
import time
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
except ImportError:
    ctk = None

try:
    import pyi_splash
except ImportError:
    pyi_splash = None

from relatorio_servicos import (
    APP_NAME,
    CancelledError,
    _default_indicator_settings,
    create_project_backup_zip,
    excel_collect_summary,
    excel_quick_missing_columns,
    gerar_relatorio_anual,
    gerar_relatorio_comparativo,
    gerar_relatorio_comparativo_anual,
    gerar_relatorio_completo,
    gerar_relatorio_qualidade_dados,
    get_default_backups_dir,
    get_indicator_settings,
    get_last_dir,
    get_last_output_path,
    get_output_mode,
    set_indicator_settings,
    set_last_dir,
    set_last_output_path,
    set_output_mode,
)

# --- CLASSE DA APLICAÇÃO (ATUALIZADA) ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.state('zoomed')

        self.cancel_event = threading.Event()
        self._run_started_ts = None
        self._progress_value = 0.0
        self.progress_label_var = tk.StringVar(value="0%")
        self.last_status_var = tk.StringVar(value="")
        self.logs_visible = tk.BooleanVar(value=False)
        self.last_output_path = get_last_output_path()
        self.output_mode_var = tk.StringVar(value=get_output_mode())
        self._output_mode_label_map = {
            'pdf': 'PDF',
            'pdf_html': 'PDF + HTML',
            'html': 'Somente HTML',
        }
        self._output_mode_value_map = {v: k for k, v in self._output_mode_label_map.items()}

        self._indicators = get_indicator_settings()

        # ── Header ──────────────────────────────────────────────────────────
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill='x', padx=14, pady=(12, 4))

        ctk.CTkLabel(header, text=APP_NAME, font=ctk.CTkFont(size=22, weight="bold")).pack(side='left')

        self.open_last_report_button = ctk.CTkButton(header, text="Abrir relatório", width=130, command=self.open_last_report)
        self.open_last_report_button.pack(side='right', padx=(4, 0))
        self.open_last_folder_button = ctk.CTkButton(header, text="Abrir pasta", width=110, command=self.open_last_report_folder)
        self.open_last_folder_button.pack(side='right', padx=(4, 0))
        self.backup_button = ctk.CTkButton(header, text="Backup", width=90, command=self.start_backup_thread)
        self.backup_button.pack(side='right', padx=(4, 0))
        self.open_backups_button = ctk.CTkButton(header, text="Ver backups", width=110, command=self.open_backups_folder)
        self.open_backups_button.pack(side='right', padx=(4, 0))

        output_frame = ctk.CTkFrame(header, fg_color="transparent")
        output_frame.pack(side='right', padx=(0, 12))
        ctk.CTkLabel(output_frame, text="Saída:", font=ctk.CTkFont(weight="bold")).pack(side='left', padx=(0, 6))
        self.output_mode_combo = ctk.CTkComboBox(
            output_frame,
            width=140,
            values=list(self._output_mode_label_map.values()),
            command=self._on_output_mode_change,
        )
        self.output_mode_combo.set(self._output_mode_label_map.get(self.output_mode_var.get(), 'PDF'))
        self.output_mode_combo.pack(side='left')

        ctk.CTkLabel(header, text=datetime.now().strftime("%d/%m/%Y"), font=ctk.CTkFont(weight="bold")).pack(side='right', padx=(0, 12))

        # separador
        sep = ctk.CTkFrame(root, height=2, fg_color=("gray70", "gray30"))
        sep.pack(fill='x', padx=12, pady=(4, 8))

        # ── Notebook (abas) ─────────────────────────────────────────────────
        self.notebook = ctk.CTkTabview(root)
        self.notebook.pack(pady=(0, 6), padx=12, fill="both", expand=True)

        for tab_name in ['Relatório de Período', 'Comparativo & Inativos', 'Análise Anual', 'Comparativo Anual', 'Qualidade de Dados']:
            self.notebook.add(tab_name)

        self.tab1 = self.notebook.tab('Relatório de Período')
        self.tab2 = self.notebook.tab('Comparativo & Inativos')
        self.tab3 = self.notebook.tab('Análise Anual')
        self.tab4 = self.notebook.tab('Comparativo Anual')
        self.tab5 = self.notebook.tab('Qualidade de Dados')

        # ── Status ───────────────────────────────────────────────────────────
        status_outer = ctk.CTkFrame(root)
        status_outer.pack(padx=12, pady=(0, 6), fill='x')

        status_header_frame = ctk.CTkFrame(status_outer, fg_color="transparent")
        status_header_frame.pack(fill='x', padx=10, pady=(8, 4))

        ctk.CTkLabel(status_header_frame, text="Status do Processamento", font=ctk.CTkFont(weight="bold")).pack(side='left')

        self.toggle_logs_button = ctk.CTkButton(status_header_frame, text="Ver detalhes", width=110, command=self.toggle_logs)
        self.toggle_logs_button.pack(side='right')
        self.copy_logs_button = ctk.CTkButton(status_header_frame, text="Copiar log", width=100, command=self.copy_logs_to_clipboard)
        self.copy_logs_button.pack(side='right', padx=(0, 8))
        self.validation_progress = ctk.CTkProgressBar(status_header_frame, mode='indeterminate', width=120)

        self.last_status_label = ctk.CTkLabel(status_outer, textvariable=self.last_status_var, anchor='w', font=ctk.CTkFont(size=11))
        self.last_status_label.pack(fill='x', padx=10, pady=(0, 4))

        self.status_box = ctk.CTkTextbox(status_outer, height=130, font=("Consolas", 9), state='disabled')
        self.status_box.pack_forget()

        # ── Barra de progresso ────────────────────────────────────────────────
        progress_frame = ctk.CTkFrame(root, fg_color="transparent")
        progress_frame.pack(padx=12, pady=(0, 12), fill='x')

        self.progress = ctk.CTkProgressBar(progress_frame)
        self.progress.set(0)
        self.progress.pack(side='left', fill='x', expand=True)

        ctk.CTkLabel(progress_frame, textvariable=self.progress_label_var, font=ctk.CTkFont(weight="bold"), width=50).pack(side='left', padx=(10, 0))
        self.cancel_button = ctk.CTkButton(progress_frame, text="Cancelar", width=100, command=self.cancel_run, state='disabled',
                                           fg_color="#c0392b", hover_color="#a93226")
        self.cancel_button.pack(side='right', padx=(10, 0))

        self.create_tab1_widgets()
        self.create_tab2_widgets()
        self.create_tab3_widgets()
        self.create_tab4_widgets()
        self.create_tab5_widgets()

        self.log_queue = queue.Queue()
        self.ui_queue = queue.Queue()
        self.root.after(100, self.process_log_queue)
        self.root.after(50, self.process_ui_queue)
        self.update_status(f"Bem-vindo ao {APP_NAME}!")
        self._apply_logs_visibility(initial=True)
        self._refresh_last_report_buttons()

    def _make_section(self, parent, title, **grid_kwargs):
        """Cria um frame com título estilo LabelFrame usando customtkinter."""
        outer = ctk.CTkFrame(parent)
        ctk.CTkLabel(outer, text=f"  {title}  ", font=ctk.CTkFont(weight="bold", size=12)).pack(anchor='nw', padx=10, pady=(6, 0))
        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.pack(fill='both', expand=True, padx=8, pady=(2, 8))
        if grid_kwargs:
            outer.grid(**grid_kwargs)
        return outer, inner

    def _on_output_mode_change(self, choice=None):
        try:
            label = (choice or self.output_mode_combo.get() or '').strip()
            mode = self._output_mode_value_map.get(label, 'pdf')
            self.output_mode_var.set(mode)
            set_output_mode(mode)
            self.update_status(f"Modo de saída: {label}")
        except Exception:
            pass

    def _refresh_last_report_buttons(self):
        try:
            has_last = bool(self.last_output_path and os.path.isfile(self.last_output_path))
            state = 'normal' if has_last else 'disabled'
            self.open_last_report_button.configure(state=state)
            self.open_last_folder_button.configure(state=state)
        except Exception:
            pass

    def _set_last_output_path(self, path):
        try:
            if path and os.path.isfile(path):
                self.last_output_path = path
                set_last_output_path(path)
        except Exception:
            pass
        self._refresh_last_report_buttons()

    def open_last_report(self):
        try:
            p = self.last_output_path
            if not p or not os.path.isfile(p):
                messagebox.showinfo("Último relatório", "Nenhum relatório recente encontrado.")
                self._refresh_last_report_buttons()
                return
            os.startfile(p)
            self.update_status(f"Abrindo relatório: {p}")
        except Exception as e:
            self.update_status(f"Erro ao abrir relatório: {e}")
            try:
                messagebox.showerror("Último relatório", f"Não foi possível abrir o relatório.\n\nDetalhes: {e}")
            except Exception:
                pass

    def open_last_report_folder(self):
        try:
            p = self.last_output_path
            if not p or not os.path.isfile(p):
                messagebox.showinfo("Último relatório", "Nenhum relatório recente encontrado.")
                self._refresh_last_report_buttons()
                return
            # Abre o Explorer selecionando o arquivo
            subprocess.run(['explorer', '/select,', p], check=False)
            self.update_status(f"Abrindo pasta do relatório: {os.path.dirname(p)}")
        except Exception as e:
            self.update_status(f"Erro ao abrir pasta do relatório: {e}")
            try:
                messagebox.showerror("Último relatório", f"Não foi possível abrir a pasta do relatório.\n\nDetalhes: {e}")
            except Exception:
                pass

    def copy_logs_to_clipboard(self):
        try:
            self.status_box.configure(state='normal')
            text = self.status_box.get('1.0', tk.END)
            self.status_box.configure(state='disabled')
            if not text.strip():
                messagebox.showinfo("Log", "O log está vazio.")
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.update_status("Log copiado para a área de transferência.")
        except Exception as e:
            self.update_status(f"Erro ao copiar log: {e}")

    def start_backup_thread(self):
        try:
            self._ui_call(self.backup_button.config, state='disabled')
        except Exception:
            pass

        def worker():
            try:
                self._ui_call(self.update_status, "Criando backup do projeto (ZIP)...")
                zip_path = create_project_backup_zip()
                self._ui_call(self.update_status, f"Backup criado com sucesso: {zip_path}")
            except Exception as e:
                self._ui_call(self.update_status, f"Erro ao criar backup: {e}")
                try:
                    self._ui_call(messagebox.showerror, "Backup", f"Não foi possível criar o backup.\n\nDetalhes: {e}")
                except Exception:
                    pass
            finally:
                try:
                    self._ui_call(self.backup_button.configure, state='normal')
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def open_backups_folder(self):
        try:
            backups_dir = get_default_backups_dir()
            try:
                os.startfile(backups_dir)  # Windows
            except Exception:
                # Fallback genérico
                subprocess.run(['explorer', backups_dir], check=False)
            self.update_status(f"Abrindo pasta de backups: {backups_dir}")
        except Exception as e:
            self.update_status(f"Erro ao abrir pasta de backups: {e}")
            try:
                messagebox.showerror("Backups", f"Não foi possível abrir a pasta de backups.\n\nDetalhes: {e}")
            except Exception:
                pass

    def process_ui_queue(self):
        try:
            while True:
                fn, args, kwargs = self.ui_queue.get_nowait()
                try:
                    fn(*args, **kwargs)
                except Exception:
                    pass
        except queue.Empty:
            pass
        self.root.after(50, self.process_ui_queue)

    def _ui_call(self, fn, *args, **kwargs):
        try:
            self.ui_queue.put((fn, args, kwargs))
        except Exception:
            pass

    def _set_boolvars(self, vars_list, value: bool):
        for var in vars_list:
            try:
                var.set(bool(value))
            except Exception:
                pass

    def _apply_default_indicators(self, group: str, keys, vars_list):
        try:
            defaults = _default_indicator_settings()
            group_defaults = defaults.get(group, {}) if isinstance(defaults, dict) else {}
            for key, var in zip(list(keys), list(vars_list)):
                try:
                    var.set(bool(group_defaults.get(key, False)))
                except Exception:
                    pass
            try:
                self._persist_indicator_settings()
            except Exception:
                pass
            try:
                self.update_status("Indicadores resetados para o padrão.")
            except Exception:
                pass
        except Exception:
            pass

    def _apply_enxuto_indicators(self, group: str, keys, vars_list):
        try:
            enxuto = {
                'periodo': {
                    'resumo_exec': True,
                    'vendas_tempo': True,
                    'vendedores': True,
                    'clientes': True,
                    'produtos': False,
                    'mapa_calor': False,
                    'geo_detalhada': False,
                },
                'comparativo': {
                    'graf_vendedor': True,
                    'graf_produtos': False,
                    'graf_clientes': False,
                    'churn': False,
                },
                'anual': {
                    'evolucao': True,
                    'pedidos_ticket': True,
                    'rank_produtos': False,
                    'rank_clientes': False,
                    'vendedores': False,
                    'uf': False,
                },
                'comparativo_anual': {
                    'evolucao_mensal': True,
                    'tabela_yoy_mensal': True,
                    'top_produtos': False,
                    'top_clientes': False,
                    'churn': False,
                },
            }

            group_preset = enxuto.get(group, {})
            for key, var in zip(list(keys), list(vars_list)):
                try:
                    var.set(bool(group_preset.get(key, False)))
                except Exception:
                    pass
            try:
                self._persist_indicator_settings()
            except Exception:
                pass
            try:
                self.update_status("Preset 'Enxuto' aplicado (mais rápido para gerar).")
            except Exception:
                pass
        except Exception:
            pass

    def _apply_logs_visibility(self, initial=False):
        show = bool(self.logs_visible.get())
        try:
            if show:
                self.status_box.pack(fill='both', expand=True, padx=8, pady=(0, 8))
                self.toggle_logs_button.configure(text="Ocultar detalhes")
            else:
                self.status_box.pack_forget()
                self.toggle_logs_button.configure(text="Ver detalhes")
        except Exception:
            pass
        if not initial:
            self.root.update_idletasks()

    def toggle_logs(self):
        self.logs_visible.set(not bool(self.logs_visible.get()))
        self._apply_logs_visibility()

    def _setup_styles(self):
        pass  # customtkinter gerencia estilos via set_appearance_mode / set_default_color_theme

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.update_status(message)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)
        
    def create_tab1_widgets(self):
        self.path_consolidado = tk.StringVar()

        self.include_periodo_resumo_exec = tk.BooleanVar(value=bool(self._indicators['periodo']['resumo_exec']))
        self.include_periodo_vendas_tempo = tk.BooleanVar(value=bool(self._indicators['periodo']['vendas_tempo']))
        self.include_periodo_vendedores = tk.BooleanVar(value=bool(self._indicators['periodo']['vendedores']))
        self.include_periodo_clientes = tk.BooleanVar(value=bool(self._indicators['periodo']['clientes']))
        self.include_periodo_produtos = tk.BooleanVar(value=bool(self._indicators['periodo']['produtos']))
        self.include_periodo_mapa_calor = tk.BooleanVar(value=bool(self._indicators['periodo']['mapa_calor']))
        self.include_periodo_geo_detalhada = tk.BooleanVar(value=bool(self._indicators['periodo']['geo_detalhada']))

        container = ctk.CTkFrame(self.tab1, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=6, pady=6)
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        _, file_inner = self._make_section(container, "1. Selecionar Arquivo de Dados", row=0, column=0, sticky='new', padx=(0, 8), pady=(0, 8))
        file_inner.columnconfigure(1, weight=1)
        ctk.CTkLabel(file_inner, text="Arquivo de Dados:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(file_inner, textvariable=self.path_consolidado).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(file_inner, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_consolidado)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(file_inner, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_consolidado.get(), group='periodo', title='Validar Excel (Período)')).grid(row=0, column=3, padx=5, pady=6)

        _, ind_inner = self._make_section(container, "Indicadores (opcional)", row=0, column=1, sticky='nsew', pady=(0, 8))
        ind_inner.columnconfigure(0, weight=1)
        ind_inner.columnconfigure(1, weight=1)

        tab1_keys = ['resumo_exec', 'vendas_tempo', 'vendedores', 'clientes', 'produtos', 'mapa_calor', 'geo_detalhada']
        tab1_vars = [
            self.include_periodo_resumo_exec, self.include_periodo_vendas_tempo,
            self.include_periodo_vendedores, self.include_periodo_clientes,
            self.include_periodo_produtos, self.include_periodo_mapa_calor,
            self.include_periodo_geo_detalhada,
        ]
        acts = ctk.CTkFrame(ind_inner, fg_color="transparent")
        acts.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        for lbl, fn in [("Completo", lambda: self._set_boolvars(tab1_vars, True)),
                        ("Limpar", lambda: self._set_boolvars(tab1_vars, False)),
                        ("Padrão", lambda: self._apply_default_indicators('periodo', tab1_keys, tab1_vars)),
                        ("Enxuto", lambda: self._apply_enxuto_indicators('periodo', tab1_keys, tab1_vars))]:
            ctk.CTkButton(acts, text=lbl, width=72, command=fn).pack(side='left', padx=(0, 4))

        checks = [
            ("Resumo executivo (meta + YoY)", self.include_periodo_resumo_exec, 1, 0, 2),
            ("Evolução das vendas no período",           self.include_periodo_vendas_tempo,  2, 0, 1),
            ("Análise de vendedores",                    self.include_periodo_vendedores,    2, 1, 1),
            ("Análise de clientes (Pareto + Top 20)",    self.include_periodo_clientes,      3, 0, 1),
            ("Análise de produtos",                      self.include_periodo_produtos,      3, 1, 1),
            ("Mapa de calor (UF)",                       self.include_periodo_mapa_calor,    4, 0, 1),
            ("Análise geográfica detalhada",             self.include_periodo_geo_detalhada, 4, 1, 1),
        ]
        for text, var, row, col, span in checks:
            ctk.CTkCheckBox(ind_inner, text=text, variable=var).grid(row=row, column=col, columnspan=span, sticky='w', padx=5, pady=3)

        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(8, 0))
        self.generate_button_A = ctk.CTkButton(button_frame, text="GERAR RELATÓRIO DE PERÍODO", command=self.start_report_thread,
                                               font=ctk.CTkFont(weight="bold", size=13), height=42)
        self.generate_button_A.pack(side='left', expand=True, fill='x', padx=(0, 6))
        self.clear_button_A = ctk.CTkButton(button_frame, text="Limpar", width=90, command=lambda: self.path_consolidado.set(""),
                                            fg_color="gray40", hover_color="gray30")
        self.clear_button_A.pack(side='right')

    def create_tab2_widgets(self):
        self.path_consolidado_A = tk.StringVar()
        self.path_consolidado_B = tk.StringVar()
        self.include_comp_vendedor = tk.BooleanVar(value=bool(self._indicators['comparativo']['graf_vendedor']))
        self.include_comp_produtos = tk.BooleanVar(value=bool(self._indicators['comparativo']['graf_produtos']))
        self.include_comp_clientes = tk.BooleanVar(value=bool(self._indicators['comparativo']['graf_clientes']))
        self.include_churn_monthly = tk.BooleanVar(value=bool(self._indicators['comparativo']['churn']))

        container = ctk.CTkFrame(self.tab2, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=6, pady=6)
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        _, fA = self._make_section(container, "1. Período A (o mais antigo)", row=0, column=0, sticky='new', padx=(0, 8), pady=(0, 6))
        fA.columnconfigure(1, weight=1)
        ctk.CTkLabel(fA, text="Arquivo A:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(fA, textvariable=self.path_consolidado_A).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(fA, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_consolidado_A)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(fA, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_consolidado_A.get(), group='comparativo', title='Validar Excel (Período A)')).grid(row=0, column=3, padx=5, pady=6)

        _, fB = self._make_section(container, "2. Período B (o mais recente)", row=1, column=0, sticky='new', padx=(0, 8), pady=(0, 6))
        fB.columnconfigure(1, weight=1)
        ctk.CTkLabel(fB, text="Arquivo B:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(fB, textvariable=self.path_consolidado_B).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(fB, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_consolidado_B)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(fB, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_consolidado_B.get(), group='comparativo', title='Validar Excel (Período B)')).grid(row=0, column=3, padx=5, pady=6)

        _, opt_inner = self._make_section(container, "Indicadores (opcional)", row=0, column=1, rowspan=2, sticky='nsew', pady=(0, 6))
        opt_inner.columnconfigure(0, weight=1)
        opt_inner.columnconfigure(1, weight=1)

        tab2_keys = ['graf_vendedor', 'graf_produtos', 'graf_clientes', 'churn']
        tab2_vars = [self.include_comp_vendedor, self.include_comp_produtos, self.include_comp_clientes, self.include_churn_monthly]
        acts2 = ctk.CTkFrame(opt_inner, fg_color="transparent")
        acts2.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        for lbl, fn in [("Completo", lambda: self._set_boolvars(tab2_vars, True)),
                        ("Limpar", lambda: self._set_boolvars(tab2_vars, False)),
                        ("Padrão", lambda: self._apply_default_indicators('comparativo', tab2_keys, tab2_vars)),
                        ("Enxuto", lambda: self._apply_enxuto_indicators('comparativo', tab2_keys, tab2_vars))]:
            ctk.CTkButton(acts2, text=lbl, width=72, command=fn).pack(side='left', padx=(0, 4))

        checks2 = [
            ("Comparativo por vendedor",        self.include_comp_vendedor,  1, 0),
            ("Comparativo Top 20 produtos",     self.include_comp_produtos,  1, 1),
            ("Comparativo Top 20 clientes",     self.include_comp_clientes,  2, 0),
            ("Churn (clientes inativos)",        self.include_churn_monthly,  2, 1),
        ]
        for text, var, row, col in checks2:
            ctk.CTkCheckBox(opt_inner, text=text, variable=var).grid(row=row, column=col, sticky='w', padx=5, pady=4)

        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(8, 0))
        self.generate_button_B = ctk.CTkButton(button_frame, text="GERAR ANÁLISE COMPARATIVA", command=self.start_churn_report_thread,
                                               font=ctk.CTkFont(weight="bold", size=13), height=42)
        self.generate_button_B.pack(side='left', expand=True, fill='x', padx=(0, 6))
        self.clear_button_B = ctk.CTkButton(button_frame, text="Limpar", width=90,
                                            command=lambda: [self.path_consolidado_A.set(""), self.path_consolidado_B.set("")],
                                            fg_color="gray40", hover_color="gray30")
        self.clear_button_B.pack(side='right')

    def create_tab3_widgets(self):
        self.path_anual = tk.StringVar()

        self.include_anual_evolucao = tk.BooleanVar(value=bool(self._indicators['anual']['evolucao']))
        self.include_anual_pedidos_ticket = tk.BooleanVar(value=bool(self._indicators['anual'].get('pedidos_ticket', True)))
        self.include_anual_rank_produtos = tk.BooleanVar(value=bool(self._indicators['anual']['rank_produtos']))
        self.include_anual_rank_clientes = tk.BooleanVar(value=bool(self._indicators['anual']['rank_clientes']))
        self.include_anual_vendedores = tk.BooleanVar(value=bool(self._indicators['anual'].get('vendedores', True)))
        self.include_anual_uf = tk.BooleanVar(value=bool(self._indicators['anual'].get('uf', True)))

        container = ctk.CTkFrame(self.tab3, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=6, pady=6)
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        _, fa = self._make_section(container, "1. Selecionar Arquivo de Dados Anuais", row=0, column=0, sticky='new', padx=(0, 8), pady=(0, 8))
        fa.columnconfigure(1, weight=1)
        ctk.CTkLabel(fa, text="Arquivo Anual:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(fa, textvariable=self.path_anual).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(fa, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_anual)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(fa, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_anual.get(), group='anual', title='Validar Excel (Anual)')).grid(row=0, column=3, padx=5, pady=6)

        _, ind3 = self._make_section(container, "Indicadores (opcional)", row=0, column=1, sticky='nsew', pady=(0, 8))
        ind3.columnconfigure(0, weight=1)
        ind3.columnconfigure(1, weight=1)

        tab3_keys = ['evolucao', 'pedidos_ticket', 'rank_produtos', 'rank_clientes', 'vendedores', 'uf']
        tab3_vars = [
            self.include_anual_evolucao, self.include_anual_pedidos_ticket,
            self.include_anual_rank_produtos, self.include_anual_rank_clientes,
            self.include_anual_vendedores, self.include_anual_uf,
        ]
        acts3 = ctk.CTkFrame(ind3, fg_color="transparent")
        acts3.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        for lbl, fn in [("Completo", lambda: self._set_boolvars(tab3_vars, True)),
                        ("Limpar", lambda: self._set_boolvars(tab3_vars, False)),
                        ("Padrão", lambda: self._apply_default_indicators('anual', tab3_keys, tab3_vars)),
                        ("Enxuto", lambda: self._apply_enxuto_indicators('anual', tab3_keys, tab3_vars))]:
            ctk.CTkButton(acts3, text=lbl, width=72, command=fn).pack(side='left', padx=(0, 4))

        checks3 = [
            ("Evolução (mensal + trimestral)",   self.include_anual_evolucao,        1, 0),
            ("Pedidos e ticket médio (mensal)",  self.include_anual_pedidos_ticket,  1, 1),
            ("Rankings de produtos",             self.include_anual_rank_produtos,   2, 0),
            ("Ranking de clientes",              self.include_anual_rank_clientes,   2, 1),
            ("Vendas por vendedor",              self.include_anual_vendedores,      3, 0),
            ("Vendas por UF",                    self.include_anual_uf,              3, 1),
        ]
        for text, var, row, col in checks3:
            ctk.CTkCheckBox(ind3, text=text, variable=var).grid(row=row, column=col, sticky='w', padx=5, pady=4)

        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(8, 0))
        self.generate_button_C = ctk.CTkButton(button_frame, text="GERAR ANÁLISE ANUAL", command=self.start_annual_report_thread,
                                               font=ctk.CTkFont(weight="bold", size=13), height=42)
        self.generate_button_C.pack(side='left', expand=True, fill='x', padx=(0, 6))
        self.clear_button_C = ctk.CTkButton(button_frame, text="Limpar", width=90, command=lambda: self.path_anual.set(""),
                                            fg_color="gray40", hover_color="gray30")
        self.clear_button_C.pack(side='right')

    def create_tab4_widgets(self):
        self.path_comp_anual_A = tk.StringVar()
        self.path_comp_anual_B = tk.StringVar()
        self.include_comp_anual_evolucao = tk.BooleanVar(value=bool(self._indicators['comparativo_anual']['evolucao_mensal']))
        self.include_comp_anual_tabela_yoy = tk.BooleanVar(value=bool(self._indicators['comparativo_anual'].get('tabela_yoy_mensal', True)))
        self.include_comp_anual_top_produtos = tk.BooleanVar(value=bool(self._indicators['comparativo_anual']['top_produtos']))
        self.include_comp_anual_top_clientes = tk.BooleanVar(value=bool(self._indicators['comparativo_anual']['top_clientes']))
        self.include_churn_annual = tk.BooleanVar(value=bool(self._indicators['comparativo_anual']['churn']))

        container = ctk.CTkFrame(self.tab4, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=6, pady=6)
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        _, f4A = self._make_section(container, "1. Ano A (o mais antigo)", row=0, column=0, sticky='new', padx=(0, 8), pady=(0, 6))
        f4A.columnconfigure(1, weight=1)
        ctk.CTkLabel(f4A, text="Arquivo Ano A:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(f4A, textvariable=self.path_comp_anual_A).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(f4A, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_comp_anual_A)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(f4A, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_comp_anual_A.get(), group='comparativo_anual', title='Validar Excel (Ano A)')).grid(row=0, column=3, padx=5, pady=6)

        _, f4B = self._make_section(container, "2. Ano B (o mais recente)", row=1, column=0, sticky='new', padx=(0, 8), pady=(0, 6))
        f4B.columnconfigure(1, weight=1)
        ctk.CTkLabel(f4B, text="Arquivo Ano B:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(f4B, textvariable=self.path_comp_anual_B).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(f4B, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_comp_anual_B)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(f4B, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_comp_anual_B.get(), group='comparativo_anual', title='Validar Excel (Ano B)')).grid(row=0, column=3, padx=5, pady=6)

        _, opt4 = self._make_section(container, "Indicadores (opcional)", row=0, column=1, rowspan=2, sticky='nsew', pady=(0, 6))
        opt4.columnconfigure(0, weight=1)
        opt4.columnconfigure(1, weight=1)

        tab4_keys = ['evolucao_mensal', 'tabela_yoy_mensal', 'top_produtos', 'top_clientes', 'churn']
        tab4_vars = [
            self.include_comp_anual_evolucao, self.include_comp_anual_tabela_yoy,
            self.include_comp_anual_top_produtos, self.include_comp_anual_top_clientes,
            self.include_churn_annual,
        ]
        acts4 = ctk.CTkFrame(opt4, fg_color="transparent")
        acts4.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        for lbl, fn in [("Completo", lambda: self._set_boolvars(tab4_vars, True)),
                        ("Limpar", lambda: self._set_boolvars(tab4_vars, False)),
                        ("Padrão", lambda: self._apply_default_indicators('comparativo_anual', tab4_keys, tab4_vars)),
                        ("Enxuto", lambda: self._apply_enxuto_indicators('comparativo_anual', tab4_keys, tab4_vars))]:
            ctk.CTkButton(acts4, text=lbl, width=72, command=fn).pack(side='left', padx=(0, 4))

        checks4 = [
            ("Evolução mensal comparativa",      self.include_comp_anual_evolucao,       1, 0),
            ("Top 20 produtos (base Ano B)",     self.include_comp_anual_top_produtos,   1, 1),
            ("Top 20 clientes (base Ano B)",     self.include_comp_anual_top_clientes,   2, 0),
            ("Churn (clientes inativos)",         self.include_churn_annual,              2, 1),
            ("Tabela YoY por mês",               self.include_comp_anual_tabela_yoy,     3, 0),
        ]
        for text, var, row, col in checks4:
            span = 2 if col == 0 and row == 3 else 1
            ctk.CTkCheckBox(opt4, text=text, variable=var).grid(row=row, column=col, columnspan=span, sticky='w', padx=5, pady=4)

        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(8, 0))
        self.generate_button_D = ctk.CTkButton(button_frame, text="GERAR COMPARATIVO ANUAL", command=self.start_annual_comp_report_thread,
                                               font=ctk.CTkFont(weight="bold", size=13), height=42)
        self.generate_button_D.pack(side='left', expand=True, fill='x', padx=(0, 6))
        self.clear_button_D = ctk.CTkButton(button_frame, text="Limpar", width=90,
                                            command=lambda: [self.path_comp_anual_A.set(""), self.path_comp_anual_B.set("")],
                                            fg_color="gray40", hover_color="gray30")
        self.clear_button_D.pack(side='right')

    def create_tab5_widgets(self):
        self.path_qualidade = tk.StringVar()

        container = ctk.CTkFrame(self.tab5, fg_color="transparent")
        container.pack(fill='both', expand=True, padx=6, pady=6)
        container.columnconfigure(0, weight=1)

        _, f5 = self._make_section(container, "1. Selecionar Arquivo para Checagem", row=0, column=0, sticky='ew', pady=(0, 8))
        f5.columnconfigure(1, weight=1)
        ctk.CTkLabel(f5, text="Arquivo Excel:").grid(row=0, column=0, sticky='w', padx=5, pady=6)
        ctk.CTkEntry(f5, textvariable=self.path_qualidade).grid(row=0, column=1, sticky='ew', padx=5, pady=6)
        ctk.CTkButton(f5, text="Procurar...", width=100, command=lambda: self.select_file_for_var(self.path_qualidade)).grid(row=0, column=2, padx=5, pady=6)
        ctk.CTkButton(f5, text="Validar", width=80, command=lambda: self.start_excel_validation_thread(self.path_qualidade.get(), group='qualidade', title='Validar Excel (Qualidade)')).grid(row=0, column=3, padx=5, pady=6)

        _, info_inner = self._make_section(container, "Sobre", row=1, column=0, sticky='ew', pady=(0, 8))
        ctk.CTkLabel(
            info_inner,
            text=(
                "Gera um relatório SOMENTE em HTML com checagens automáticas de qualidade: "
                "UF inválida, vendedor ausente, campos vazios, inconsistências por pedido, "
                "valores/quantidades zeradas ou negativas."
            ),
            wraplength=800,
            justify='left',
            anchor='w',
        ).pack(fill='x', padx=4, pady=4)

        button_frame = ctk.CTkFrame(container, fg_color="transparent")
        button_frame.grid(row=2, column=0, sticky='ew', pady=(8, 0))
        self.generate_button_E = ctk.CTkButton(button_frame, text="GERAR RELATÓRIO DE QUALIDADE (HTML)", command=self.start_quality_report_thread,
                                               font=ctk.CTkFont(weight="bold", size=13), height=42)
        self.generate_button_E.pack(side='left', expand=True, fill='x', padx=(0, 6))
        self.clear_button_E = ctk.CTkButton(button_frame, text="Limpar", width=90, command=lambda: self.path_qualidade.set(""),
                                            fg_color="gray40", hover_color="gray30")
        self.clear_button_E.pack(side='right')

    def _persist_indicator_settings(self):
        indicators = {
            'periodo': {
                'resumo_exec': bool(self.include_periodo_resumo_exec.get()),
                'vendas_tempo': bool(self.include_periodo_vendas_tempo.get()),
                'vendedores': bool(self.include_periodo_vendedores.get()),
                'clientes': bool(self.include_periodo_clientes.get()),
                'produtos': bool(self.include_periodo_produtos.get()),
                'mapa_calor': bool(self.include_periodo_mapa_calor.get()),
                'geo_detalhada': bool(self.include_periodo_geo_detalhada.get()),
            },
            'comparativo': {
                'graf_vendedor': bool(self.include_comp_vendedor.get()),
                'graf_produtos': bool(self.include_comp_produtos.get()),
                'graf_clientes': bool(self.include_comp_clientes.get()),
                'churn': bool(self.include_churn_monthly.get()),
            },
            'anual': {
                'evolucao': bool(self.include_anual_evolucao.get()),
                'pedidos_ticket': bool(self.include_anual_pedidos_ticket.get()),
                'rank_produtos': bool(self.include_anual_rank_produtos.get()),
                'rank_clientes': bool(self.include_anual_rank_clientes.get()),
                'vendedores': bool(self.include_anual_vendedores.get()),
                'uf': bool(self.include_anual_uf.get()),
            },
            'comparativo_anual': {
                'evolucao_mensal': bool(self.include_comp_anual_evolucao.get()),
                'tabela_yoy_mensal': bool(self.include_comp_anual_tabela_yoy.get()),
                'top_produtos': bool(self.include_comp_anual_top_produtos.get()),
                'top_clientes': bool(self.include_comp_anual_top_clientes.get()),
                'churn': bool(self.include_churn_annual.get()),
            },
        }
        set_indicator_settings(indicators)
        self._indicators = indicators

    def select_file_for_var(self, path_variable):
        filepath = filedialog.askopenfilename(
            title="Selecione o ficheiro de dados consolidados",
            initialdir=get_last_dir(),
            filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            path_variable.set(filepath)
            set_last_dir(filepath)

    def _current_opts_for_group(self, group: str) -> dict:
        try:
            if group == 'periodo':
                return {
                    'resumo_exec': bool(self.include_periodo_resumo_exec.get()),
                    'vendas_tempo': bool(self.include_periodo_vendas_tempo.get()),
                    'vendedores': bool(self.include_periodo_vendedores.get()),
                    'clientes': bool(self.include_periodo_clientes.get()),
                    'produtos': bool(self.include_periodo_produtos.get()),
                    'mapa_calor': bool(self.include_periodo_mapa_calor.get()),
                    'geo_detalhada': bool(self.include_periodo_geo_detalhada.get()),
                }
            if group == 'comparativo':
                return {
                    'graf_vendedor': bool(self.include_comp_vendedor.get()),
                    'graf_produtos': bool(self.include_comp_produtos.get()),
                    'graf_clientes': bool(self.include_comp_clientes.get()),
                    'churn': bool(self.include_churn_monthly.get()),
                }
            if group == 'anual':
                return {
                    'evolucao': bool(self.include_anual_evolucao.get()),
                    'pedidos_ticket': bool(self.include_anual_pedidos_ticket.get()),
                    'rank_produtos': bool(self.include_anual_rank_produtos.get()),
                    'rank_clientes': bool(self.include_anual_rank_clientes.get()),
                    'vendedores': bool(self.include_anual_vendedores.get()),
                    'uf': bool(self.include_anual_uf.get()),
                }
            if group == 'comparativo_anual':
                return {
                    'evolucao_mensal': bool(self.include_comp_anual_evolucao.get()),
                    'tabela_yoy_mensal': bool(self.include_comp_anual_tabela_yoy.get()),
                    'top_produtos': bool(self.include_comp_anual_top_produtos.get()),
                    'top_clientes': bool(self.include_comp_anual_top_clientes.get()),
                    'churn': bool(self.include_churn_annual.get()),
                }
        except Exception:
            pass
        return {}

    def _required_columns_for_group(self, group: str, opts: dict):
        required = ['Dt.Pedido', 'Vlr.Total', 'Nro.Pedido']

        def add(col):
            if col not in required:
                required.append(col)

        if group == 'periodo':
            if opts.get('vendedores'):
                add('Vendedor')
            if opts.get('clientes'):
                add('Cliente')
            if opts.get('produtos'):
                add('Cod.Material')
                add('Descrição')
                add('Qtde')
            if opts.get('mapa_calor') or opts.get('geo_detalhada'):
                add('UF')
        elif group == 'comparativo':
            if opts.get('graf_vendedor'):
                add('Vendedor')
            if opts.get('graf_produtos'):
                add('Cod.Material')
                add('Descrição')
            if opts.get('graf_clientes'):
                add('Cliente')
            if opts.get('churn'):
                add('Cliente')
        elif group == 'anual':
            if opts.get('rank_produtos'):
                add('Cod.Material')
                add('Descrição')
                add('Qtde')
            if opts.get('rank_clientes'):
                add('Cliente')
            if opts.get('vendedores'):
                add('Vendedor')
            if opts.get('uf'):
                add('UF')
        elif group == 'comparativo_anual':
            if opts.get('top_produtos'):
                add('Cod.Material')
                add('Descrição')
            if opts.get('top_clientes'):
                add('Cliente')
            if opts.get('churn'):
                add('Cliente')
        elif group == 'qualidade':
            pass

        return required

    def _column_reasons_for_group(self, group: str, opts: dict):
        reasons = {
            'Dt.Pedido': {'Base do relatório'},
            'Vlr.Total': {'Base do relatório'},
            'Nro.Pedido': {'Base do relatório'},
        }

        def add(col: str, reason: str):
            reasons.setdefault(col, set()).add(reason)

        if group == 'periodo':
            if opts.get('vendedores'):
                add('Vendedor', 'Indicador: Vendedores')
            if opts.get('clientes'):
                add('Cliente', 'Indicador: Clientes')
            if opts.get('produtos'):
                add('Cod.Material', 'Indicador: Produtos (chave de consolidação)')
                add('Descrição', 'Indicador: Produtos (label de exibição)')
                add('Qtde', 'Indicador: Produtos')
            if opts.get('mapa_calor'):
                add('UF', 'Indicador: Mapa de calor')
            if opts.get('geo_detalhada'):
                add('UF', 'Indicador: Geo detalhada')
        elif group == 'comparativo':
            if opts.get('graf_vendedor'):
                add('Vendedor', 'Gráfico: Vendedor')
            if opts.get('graf_produtos'):
                add('Cod.Material', 'Gráfico: Produtos (chave de consolidação)')
                add('Descrição', 'Gráfico: Produtos (label de exibição)')
            if opts.get('graf_clientes'):
                add('Cliente', 'Gráfico: Clientes')
            if opts.get('churn'):
                add('Cliente', 'Indicador: Churn')
        elif group == 'anual':
            if opts.get('rank_produtos'):
                add('Cod.Material', 'Ranking: Produtos (chave de consolidação)')
                add('Descrição', 'Ranking: Produtos (label de exibição)')
                add('Qtde', 'Ranking: Produtos')
            if opts.get('rank_clientes'):
                add('Cliente', 'Ranking: Clientes')
            if opts.get('vendedores'):
                add('Vendedor', 'Indicador: Vendedores')
            if opts.get('uf'):
                add('UF', 'Indicador: UF')
        elif group == 'comparativo_anual':
            if opts.get('top_produtos'):
                add('Cod.Material', 'Top: Produtos (chave de consolidação)')
                add('Descrição', 'Top: Produtos (label de exibição)')
            if opts.get('top_clientes'):
                add('Cliente', 'Top: Clientes')
            if opts.get('churn'):
                add('Cliente', 'Indicador: Churn')
        elif group == 'qualidade':
            pass

        return reasons

    def _format_missing_columns_block(self, missing: list, reasons_map: dict):
        if not missing:
            return ""

        lines = []
        for col in missing:
            reasons = reasons_map.get(col) or set()
            if reasons:
                lines.append(f"- {col}  ({'; '.join(sorted(reasons))})")
            else:
                lines.append(f"- {col}")
        return "\n".join(lines)

    def _show_excel_validation(self, filepath: str, *, group: str, title: str):
        try:
            if not filepath or not os.path.isfile(filepath):
                messagebox.showerror(title, "Arquivo não encontrado. Selecione um Excel válido.")
                return

            opts = self._current_opts_for_group(group)
            required = self._required_columns_for_group(group, opts)
            missing, _ = excel_quick_missing_columns(filepath, required)
            summary = excel_collect_summary(filepath, want_date=True, want_uf=('UF' in required))
            reasons_map = self._column_reasons_for_group(group, opts)

            rows = summary.get('rows')
            dmin = summary.get('date_min')
            dmax = summary.get('date_max')
            uf_top = summary.get('uf_top') or []

            lines = []
            lines.append(f"Arquivo: {filepath}")
            if summary.get('sheet'):
                lines.append(f"Aba: {summary.get('sheet')}")
            if rows is not None:
                lines.append(f"Linhas (aprox.): {rows}")

            if dmin and dmax:
                lines.append(f"Período (amostra): {dmin.strftime('%d/%m/%Y')}  →  {dmax.strftime('%d/%m/%Y')}")
            else:
                lines.append("Período: não foi possível detectar (verifique 'Dt.Pedido')")

            if uf_top:
                uf_str = ", ".join([f"{uf} ({q})" for uf, q in uf_top])
                lines.append(f"UFs (amostra): {uf_str}")

            lines.append("")
            lines.append("Colunas obrigatórias (pelas opções atuais):")
            lines.extend([f"- {c}" for c in required])

            if missing:
                lines.append("")
                lines.append("⚠️ Colunas ausentes:")
                lines.append(self._format_missing_columns_block(missing, reasons_map))
                messagebox.showwarning(title, "\n".join(lines))
            else:
                lines.append("")
                lines.append("✅ Colunas OK.")
                messagebox.showinfo(title, "\n".join(lines))

            try:
                self.update_status(f"Validação do Excel: {os.path.basename(filepath)} | faltando: {len(missing)}")
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror(title, f"Falha ao validar o Excel.\n\nDetalhes: {e}")
            except Exception:
                pass

    def _set_validation_busy(self, busy: bool, message: str = None):
        try:
            if message:
                self.update_status(message)
        except Exception:
            pass

        try:
            if busy:
                self.validation_progress.pack(side='right', padx=(0, 8))
                self.validation_progress.start()
            else:
                try:
                    self.validation_progress.stop()
                except Exception:
                    pass
                self.validation_progress.pack_forget()
        except Exception:
            pass

    def start_excel_validation_thread(self, filepath: str, *, group: str, title: str):
        try:
            if not filepath or not os.path.isfile(filepath):
                messagebox.showerror(title, "Arquivo não encontrado. Selecione um Excel válido.")
                return

            opts = self._current_opts_for_group(group)
            required = self._required_columns_for_group(group, opts)
            want_uf = ('UF' in required)
            reasons_map = self._column_reasons_for_group(group, opts)

            self._set_validation_busy(True, "Validando Excel...")

            def worker():
                try:
                    missing, _ = excel_quick_missing_columns(filepath, required)
                    summary = excel_collect_summary(filepath, want_date=True, want_uf=want_uf)

                    def show_result():
                        try:
                            rows = summary.get('rows')
                            dmin = summary.get('date_min')
                            dmax = summary.get('date_max')
                            uf_top = summary.get('uf_top') or []

                            lines = []
                            lines.append(f"Arquivo: {filepath}")
                            if summary.get('sheet'):
                                lines.append(f"Aba: {summary.get('sheet')}")
                            if rows is not None:
                                lines.append(f"Linhas (aprox.): {rows}")

                            if dmin and dmax:
                                lines.append(f"Período (amostra): {dmin.strftime('%d/%m/%Y')}  →  {dmax.strftime('%d/%m/%Y')}")
                            else:
                                lines.append("Período: não foi possível detectar (verifique 'Dt.Pedido')")

                            if uf_top:
                                uf_str = ", ".join([f"{uf} ({q})" for uf, q in uf_top])
                                lines.append(f"UFs (amostra): {uf_str}")

                            lines.append("")
                            lines.append("Colunas obrigatórias (pelas opções atuais):")
                            lines.extend([f"- {c}" for c in required])

                            if missing:
                                lines.append("")
                                lines.append("⚠️ Colunas ausentes:")
                                lines.append(self._format_missing_columns_block(missing, reasons_map))
                                messagebox.showwarning(title, "\n".join(lines))
                            else:
                                lines.append("")
                                lines.append("✅ Colunas OK.")
                                messagebox.showinfo(title, "\n".join(lines))

                            try:
                                self.update_status(f"Validação do Excel: {os.path.basename(filepath)} | faltando: {len(missing)}")
                            except Exception:
                                pass
                        finally:
                            self._set_validation_busy(False)

                    self._ui_call(show_result)
                except Exception as e:
                    def show_err():
                        try:
                            messagebox.showerror(title, f"Falha ao validar o Excel.\n\nDetalhes: {e}")
                        finally:
                            self._set_validation_busy(False)
                    self._ui_call(show_err)

            threading.Thread(target=worker, daemon=True).start()
        except Exception as e:
            try:
                messagebox.showerror(title, f"Falha ao iniciar validação.\n\nDetalhes: {e}")
            except Exception:
                pass
            self._set_validation_busy(False)

    def cancel_run(self):
        self.cancel_event.set()
        self.log_queue.put("Solicitação de cancelamento enviada...")

    def update_status(self, message):
        try:
            self.last_status_var.set(message)
        except Exception:
            pass
        try:
            self.status_box.configure(state='normal')
            self.status_box.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
            self.status_box.see(tk.END)
            self.status_box.configure(state='disabled')
        except Exception:
            pass
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def _format_duration(self, seconds: float) -> str:
        try:
            s = int(round(float(seconds)))
        except Exception:
            return "-"

        if s < 0:
            s = 0

        h = s // 3600
        m = (s % 3600) // 60
        sec = s % 60

        if h > 0:
            return f"{h}h {m:02d}m {sec:02d}s"
        if m > 0:
            return f"{m}m {sec:02d}s"
        return f"{sec}s"

    def report_finished(self, success=True, save_path=None, cancelled=False):
        self.generate_button_A.configure(state='normal')
        self.generate_button_B.configure(state='normal')
        self.generate_button_C.configure(state='normal')
        self.generate_button_D.configure(state='normal')
        try:
            self.generate_button_E.configure(state='normal')
        except Exception:
            pass

        elapsed = None
        try:
            if self._run_started_ts is not None:
                elapsed = max(0.0, time.time() - float(self._run_started_ts))
        except Exception:
            elapsed = None
        finally:
            self._run_started_ts = None

        if success and not cancelled:
            self.log_queue.put(">>>> PROCESSO CONCLUÍDO COM SUCESSO! <<<<")
            try:
                if save_path and os.path.isfile(save_path):
                    self._set_last_output_path(save_path)
            except Exception:
                pass
        elif success and cancelled:
            self.log_queue.put(">>>> PROCESSO CANCELADO PELO USUÁRIO. <<<<")
        else:
            self.log_queue.put(">>>> PROCESSO INTERROMPIDO POR ERRO. <<<<")

        try:
            if elapsed is not None:
                if success and not cancelled:
                    self.update_status(f"Concluído. Última geração: {self._format_duration(elapsed)}")
                elif success and cancelled:
                    self.update_status(f"Cancelado. Tempo até cancelar: {self._format_duration(elapsed)}")
                else:
                    self.update_status(f"Erro. Tempo até falhar: {self._format_duration(elapsed)}")
        except Exception:
            pass

        self.cancel_button.configure(state='disabled')
        self._progress_value = 0.0
        self.progress.set(0)
        self.progress_label_var.set("0%")

    def start_report_thread(self):
        path_consolidado = self.path_consolidado.get()
        if not path_consolidado:
            messagebox.showerror("Ficheiro em Falta", "Por favor, selecione o ficheiro de dados consolidados.")
            return

        output_mode = (self.output_mode_var.get() or 'pdf').strip().lower()
        
        base_name = os.path.basename(path_consolidado).replace('dados_consolidados_', '').replace('.xlsx', '')
        default_filename = f"Relatorio_Vendas_{base_name}.pdf" if output_mode != 'html' else f"Relatorio_Vendas_{base_name}.html"
        
        if output_mode == 'html':
            save_path = filedialog.asksaveasfilename(
                defaultextension=".html", filetypes=[("HTML files", "*.html")],
                title="Salvar Relatório de Período (HTML)", initialfile=default_filename, initialdir=get_last_dir())
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                title="Salvar Relatório de Período", initialfile=default_filename, initialdir=get_last_dir())
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        set_last_dir(save_path)

        self._persist_indicator_settings()
        opts = self._indicators['periodo']
        required = self._required_columns_for_group('periodo', opts)
        missing, _ = excel_quick_missing_columns(path_consolidado, required)
        if missing:
            reasons_map = self._column_reasons_for_group('periodo', opts)
            messagebox.showerror(
                "Validação do Excel",
                "O arquivo está sem colunas necessárias para gerar este relatório com as opções atuais.\n\n"
                + "Colunas ausentes:\n"
                + self._format_missing_columns_block(missing, reasons_map)
            )
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_report, args=(path_consolidado, save_path, opts, output_mode))
        thread.start()
        
    def start_churn_report_thread(self):
        path_A = self.path_consolidado_A.get()
        path_B = self.path_consolidado_B.get()
        if not path_A or not path_B:
            messagebox.showerror("Ficheiros em Falta", "Por favor, selecione os dois ficheiros de dados para a comparação.")
            return
        
        output_mode = (self.output_mode_var.get() or 'pdf').strip().lower()
        default_filename = "Relatorio_Comparativo.pdf" if output_mode != 'html' else "Relatorio_Comparativo.html"
        if output_mode == 'html':
            save_path = filedialog.asksaveasfilename(
                defaultextension=".html", filetypes=[("HTML files", "*.html")],
                title="Salvar Análise Comparativa (HTML)", initialfile=default_filename, initialdir=get_last_dir())
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                title="Salvar Análise Comparativa", initialfile=default_filename, initialdir=get_last_dir())
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        set_last_dir(save_path)

        self._persist_indicator_settings()
        opts = self._indicators['comparativo']
        required = self._required_columns_for_group('comparativo', opts)
        missing_A, _ = excel_quick_missing_columns(path_A, required)
        missing_B, _ = excel_quick_missing_columns(path_B, required)
        if missing_A or missing_B:
            reasons_map = self._column_reasons_for_group('comparativo', opts)
            msg = []
            if missing_A:
                msg.append("Período A - colunas ausentes:\n" + self._format_missing_columns_block(missing_A, reasons_map))
            if missing_B:
                msg.append("Período B - colunas ausentes:\n" + self._format_missing_columns_block(missing_B, reasons_map))
            messagebox.showerror("Validação do Excel", "\n\n".join(msg))
            return

        self.prepare_to_run()
        thread = threading.Thread(
            target=self.run_churn_report,
            args=(path_A, path_B, save_path, opts, output_mode)
        )
        thread.start()

    def start_annual_report_thread(self):
        path_anual = self.path_anual.get()
        if not path_anual:
            messagebox.showerror("Ficheiro em Falta", "Por favor, selecione o ficheiro de dados anuais.")
            return
        
        output_mode = (self.output_mode_var.get() or 'pdf').strip().lower()
        default_filename = "Relatorio_Analise_Anual.pdf" if output_mode != 'html' else "Relatorio_Analise_Anual.html"
        if output_mode == 'html':
            save_path = filedialog.asksaveasfilename(
                defaultextension=".html", filetypes=[("HTML files", "*.html")],
                title="Salvar Análise Anual (HTML)", initialfile=default_filename, initialdir=get_last_dir())
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                title="Salvar Análise Anual", initialfile=default_filename, initialdir=get_last_dir())
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        set_last_dir(save_path)

        self._persist_indicator_settings()
        opts = self._indicators['anual']
        required = self._required_columns_for_group('anual', opts)
        missing, _ = excel_quick_missing_columns(path_anual, required)
        if missing:
            reasons_map = self._column_reasons_for_group('anual', opts)
            messagebox.showerror(
                "Validação do Excel",
                "O arquivo anual está sem colunas necessárias para gerar este relatório com as opções atuais.\n\n"
                + "Colunas ausentes:\n"
                + self._format_missing_columns_block(missing, reasons_map)
            )
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_annual_report, args=(path_anual, save_path, opts, output_mode))
        thread.start()

    def start_annual_comp_report_thread(self):
        path_A = self.path_comp_anual_A.get()
        path_B = self.path_comp_anual_B.get()
        if not path_A or not path_B:
            messagebox.showerror("Ficheiros em Falta", "Por favor, selecione os dois ficheiros de dados anuais para a comparação.")
            return
        
        output_mode = (self.output_mode_var.get() or 'pdf').strip().lower()
        default_filename = "Relatorio_Comparativo_Anual.pdf" if output_mode != 'html' else "Relatorio_Comparativo_Anual.html"
        if output_mode == 'html':
            save_path = filedialog.asksaveasfilename(
                defaultextension=".html", filetypes=[("HTML files", "*.html")],
                title="Salvar Comparativo Anual (HTML)", initialfile=default_filename, initialdir=get_last_dir())
        else:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                title="Salvar Comparativo Anual", initialfile=default_filename, initialdir=get_last_dir())
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        set_last_dir(save_path)

        self._persist_indicator_settings()
        opts = self._indicators['comparativo_anual']
        required = self._required_columns_for_group('comparativo_anual', opts)
        missing_A, _ = excel_quick_missing_columns(path_A, required)
        missing_B, _ = excel_quick_missing_columns(path_B, required)
        if missing_A or missing_B:
            reasons_map = self._column_reasons_for_group('comparativo_anual', opts)
            msg = []
            if missing_A:
                msg.append("Ano A - colunas ausentes:\n" + self._format_missing_columns_block(missing_A, reasons_map))
            if missing_B:
                msg.append("Ano B - colunas ausentes:\n" + self._format_missing_columns_block(missing_B, reasons_map))
            messagebox.showerror("Validação do Excel", "\n\n".join(msg))
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_annual_comp_report, args=(path_A, path_B, save_path, opts, output_mode))
        thread.start()

    def start_quality_report_thread(self):
        path_excel = self.path_qualidade.get()
        if not path_excel:
            messagebox.showerror("Ficheiro em Falta", "Por favor, selecione um ficheiro Excel para checar a qualidade.")
            return

        try:
            base_name = os.path.basename(path_excel).replace('dados_consolidados_', '').replace('.xlsx', '')
        except Exception:
            base_name = "dados"
        default_filename = f"Relatorio_Qualidade_{base_name}.html"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html")],
            title="Salvar Relatório de Qualidade (HTML)",
            initialfile=default_filename,
            initialdir=get_last_dir(),
        )
        if not save_path:
            self.update_status("Operação de salvar cancelada pelo usuário.")
            return

        set_last_dir(save_path)

        required = self._required_columns_for_group('qualidade', {})
        missing, _ = excel_quick_missing_columns(path_excel, required)
        if missing:
            reasons_map = self._column_reasons_for_group('qualidade', {})
            messagebox.showerror(
                "Validação do Excel",
                "O arquivo está sem colunas mínimas para gerar o relatório de qualidade.\n\n"
                + "Colunas ausentes:\n"
                + self._format_missing_columns_block(missing, reasons_map)
            )
            return

        self.prepare_to_run()
        thread = threading.Thread(target=self.run_quality_report, args=(path_excel, save_path))
        thread.start()

    def prepare_to_run(self):
        self.generate_button_A.configure(state='disabled')
        self.generate_button_B.configure(state='disabled')
        self.generate_button_C.configure(state='disabled')
        self.generate_button_D.configure(state='disabled')
        try:
            self.generate_button_E.configure(state='disabled')
        except Exception:
            pass
        try:
            self._run_started_ts = time.time()
        except Exception:
            self._run_started_ts = None
        self.cancel_event.clear()
        self.cancel_button.configure(state='normal')
        self._progress_value = 0.0
        self.progress.set(0)
        self.progress_label_var.set("0%")
        try:
            self.status_box.configure(state='normal')
            self.status_box.delete('1.0', tk.END)
            self.status_box.configure(state='disabled')
        except Exception:
            pass

    def _progress_cb(self, value):
        try:
            v = float(value)
        except Exception:
            return

        def _apply_progress():
            try:
                self._progress_value = v
                self.progress.set(v / 100.0)
                self.progress_label_var.set(f"{int(round(v))}%")
                self.root.update_idletasks()
            except Exception:
                pass

        self._ui_call(_apply_progress)
        
    def run_report(self, path_consolidado, save_path, opts, output_mode):
        try:
            out_path = gerar_relatorio_completo(
                path_consolidado,
                self.log_queue,
                save_path,
                cancel_event=self.cancel_event,
                progress_cb=self._progress_cb,
                output_mode=output_mode,
                include_resumo_exec=bool(opts.get('resumo_exec', True)),
                include_vendas_tempo=bool(opts.get('vendas_tempo', True)),
                include_vendedores=bool(opts.get('vendedores', True)),
                include_clientes=bool(opts.get('clientes', True)),
                include_produtos=bool(opts.get('produtos', True)),
                include_mapa_calor=bool(opts.get('mapa_calor', True)),
                include_geo_detalhada=bool(opts.get('geo_detalhada', True)),
            )
            self._ui_call(self.report_finished, True, out_path or save_path, False)
        except CancelledError:
            self._ui_call(self.report_finished, True, None, True)
        except Exception as e:
            self._ui_call(self.report_finished, False, None, False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")
            
    def run_churn_report(self, path_A, path_B, save_path, opts, output_mode):
        try:
            out_path = gerar_relatorio_comparativo(
                path_A,
                path_B,
                self.log_queue,
                save_path,
                cancel_event=self.cancel_event,
                progress_cb=self._progress_cb,
                output_mode=output_mode,
                include_churn=bool(opts.get('churn', False)),
                include_graf_vendedor=bool(opts.get('graf_vendedor', True)),
                include_graf_produtos=bool(opts.get('graf_produtos', True)),
                include_graf_clientes=bool(opts.get('graf_clientes', True)),
            )
            self._ui_call(self.report_finished, True, out_path or save_path, False)
        except CancelledError:
            self._ui_call(self.report_finished, True, None, True)
        except Exception as e:
            self._ui_call(self.report_finished, False, None, False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")
            
    def run_annual_report(self, path_anual, save_path, opts, output_mode):
        try:
            out_path = gerar_relatorio_anual(
                path_anual,
                self.log_queue,
                save_path,
                cancel_event=self.cancel_event,
                progress_cb=self._progress_cb,
                output_mode=output_mode,
                include_evolucao=bool(opts.get('evolucao', True)),
                include_pedidos_ticket=bool(opts.get('pedidos_ticket', True)),
                include_rank_produtos=bool(opts.get('rank_produtos', True)),
                include_rank_clientes=bool(opts.get('rank_clientes', True)),
                include_vendedores=bool(opts.get('vendedores', True)),
                include_uf=bool(opts.get('uf', True)),
            )
            self._ui_call(self.report_finished, True, out_path or save_path, False)
        except CancelledError:
            self._ui_call(self.report_finished, True, None, True)
        except Exception as e:
            self._ui_call(self.report_finished, False, None, False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")

    def run_annual_comp_report(self, path_A, path_B, save_path, opts, output_mode):
        try:
            out_path = gerar_relatorio_comparativo_anual(
                path_A,
                path_B,
                self.log_queue,
                save_path,
                cancel_event=self.cancel_event,
                progress_cb=self._progress_cb,
                output_mode=output_mode,
                include_churn=bool(opts.get('churn', True)),
                include_evolucao_mensal=bool(opts.get('evolucao_mensal', True)),
                include_tabela_yoy_mensal=bool(opts.get('tabela_yoy_mensal', True)),
                include_top_produtos=bool(opts.get('top_produtos', True)),
                include_top_clientes=bool(opts.get('top_clientes', True)),
            )
            self._ui_call(self.report_finished, True, out_path or save_path, False)
        except CancelledError:
            self._ui_call(self.report_finished, True, None, True)
        except Exception as e:
            self._ui_call(self.report_finished, False, None, False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")

    def run_quality_report(self, path_excel, save_path):
        try:
            gerar_relatorio_qualidade_dados(
                path_excel,
                self.log_queue,
                save_path,
                cancel_event=self.cancel_event,
                progress_cb=self._progress_cb,
            )
            self._ui_call(self.report_finished, True, save_path, False)
        except CancelledError:
            self._ui_call(self.report_finished, True, None, True)
        except Exception as e:
            self._ui_call(self.report_finished, False, None, False)
            print(f"Erro capturado pela thread da GUI: {e}")
            self.log_queue.put(f"ERRO CRÍTICO: {e}")

if __name__ == "__main__":
    log_file_path = "crash_log.txt"
    try:
        multiprocessing.freeze_support()
        root = ctk.CTk()
        app = App(root)

        if pyi_splash:
            pyi_splash.close()

        root.mainloop()

    except Exception as e:
        import traceback
        with open(log_file_path, 'w', encoding='utf-8') as f:
            f.write(f"Ocorreu um erro fatal:\n")
            f.write(str(e) + "\n\n")
            f.write(traceback.format_exc())
# --- FIM DO CÓDIGO ---