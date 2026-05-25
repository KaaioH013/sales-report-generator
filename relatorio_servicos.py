# -*- coding: utf-8 -*-
"""Geração de relatórios, exportação e utilitários (sem interface gráfica)."""
import os
import sys
os.environ.setdefault('MPLBACKEND', 'Agg')
import importlib
import threading
import multiprocessing
import locale
import queue
import time
from collections import Counter
from datetime import datetime, date
import pathlib
import tempfile
import shutil
import json
import zipfile
import fnmatch
import subprocess
from contextlib import contextmanager


class LazyModule:
    """Proxy leve para adiar imports pesados até o primeiro uso."""

    def __init__(self, module_name: str):
        self._module_name = module_name
        self._module = None

    def _load(self):
        if self._module is None:
            self._module = importlib.import_module(self._module_name)
        return self._module

    def __getattr__(self, item):
        return getattr(self._load(), item)


# Imports pesados (pandas/numpy/matplotlib/seaborn) somente quando necessário.
pd = LazyModule('pandas')
np = LazyModule('numpy')
plt = LazyModule('matplotlib.pyplot')
sns = LazyModule('seaborn')


def _get_mpl_FuncFormatter():
    from matplotlib.ticker import FuncFormatter
    return FuncFormatter


def _get_mpl_PercentFormatter():
    from matplotlib.ticker import PercentFormatter
    return PercentFormatter

try:
    from jinja2 import Environment, BaseLoader, select_autoescape
    from markupsafe import Markup
except Exception:
    Environment = None
    BaseLoader = None
    Markup = None

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
    except locale.Error:
        print("AVISO: Locale pt_BR não encontrado.")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

APP_NAME = "Gerador de Relatórios"
__version__ = "21.0"


class CancelledError(Exception):
    pass


def _raise_if_cancelled(cancel_event):
    try:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
    except CancelledError:
        raise
    except Exception:
        return


def _detect_project_root_for_backup():
    """Tenta encontrar a raiz do projeto para backups.

    Prioriza o diretório atual quando estiver rodando a partir do projeto.
    """
    try:
        cwd = os.getcwd()
        if os.path.isfile(os.path.join(cwd, 'gerar_relatorio.py')) or os.path.isfile(
            os.path.join(cwd, 'relatorio_servicos.py')
        ):
            return cwd
    except Exception:
        pass

    try:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(os.path.abspath(sys.executable))
    except Exception:
        pass

    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.path.abspath('.')


def _ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def _is_writable_dir(path):
    try:
        _ensure_dir(path)
        test_file = os.path.join(path, f".write_test_{int(time.time())}.tmp")
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write('ok')
        os.remove(test_file)
        return True
    except Exception:
        return False


def get_default_backups_dir(project_root=None):
    if not project_root:
        project_root = _detect_project_root_for_backup()
    project_root = os.path.abspath(project_root)

    preferred = os.path.join(project_root, 'backups')
    if _is_writable_dir(preferred):
        return preferred

    fallback = os.path.join(os.environ.get('APPDATA', project_root), 'relatorios_2026', 'backups')
    _ensure_dir(fallback)
    return fallback


def create_project_backup_zip(*, project_root=None, backups_dir=None, keep_last=10):
    """Cria um backup ZIP com código/templates/recursos (não inclui dados grandes).

    Retorna o caminho do ZIP gerado.
    """
    if not project_root:
        project_root = _detect_project_root_for_backup()
    project_root = os.path.abspath(project_root)

    if not backups_dir:
        backups_dir = get_default_backups_dir(project_root)

    ts = datetime.now().strftime('%Y%m%d-%H%M%S')
    zip_path = os.path.join(backups_dir, f"backup_projeto_{ts}.zip")

    include_exts = {'.py', '.html', '.spec', '.bat', '.md', '.txt', '.json'}
    include_names = {'Logo.png', 'splash.png', 'mapa_brasil_dados.pkl', '.gitignore'}
    include_dirs = {'geo_data'}
    exclude_dirs = {'.git', 'venv_stable', '.venv', 'build', 'dist', '__pycache__', 'backups'}
    exclude_globs = ['dados_consolidados_*.xlsx', '*.xlsx', '*.pdf', 'grafico_*.png']

    def is_excluded_name(name):
        for pat in exclude_globs:
            if fnmatch.fnmatch(name, pat):
                return True
        return False

    def should_include(rel_path, filename):
        if is_excluded_name(filename):
            return False
        ext = os.path.splitext(filename)[1].lower()
        if filename in include_names:
            return True
        if ext in include_exts:
            return True
        top_dir = rel_path.split(os.sep, 1)[0] if rel_path else ''
        if top_dir in include_dirs:
            return True
        return False

    # Git info (se houver)
    git_info_lines = [
        f"Backup timestamp: {ts}",
        f"Project root    : {project_root}",
    ]
    try:
        result = subprocess.run(
            ['git', 'rev-parse', '--is-inside-work-tree'],
            cwd=project_root,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.stdout.strip() == 'true':
            head = subprocess.run(['git', 'rev-parse', 'HEAD'], cwd=project_root, capture_output=True, text=True, check=False).stdout.strip()
            branch = subprocess.run(['git', 'branch', '--show-current'], cwd=project_root, capture_output=True, text=True, check=False).stdout.strip()
            status = subprocess.run(['git', 'status', '--porcelain'], cwd=project_root, capture_output=True, text=True, check=False).stdout.strip()
            git_info_lines += [
                f"Git branch      : {branch}",
                f"Git HEAD        : {head}",
                f"Git dirty       : {bool(status)}",
                "Git status:",
                status if status else "(clean)",
            ]
        else:
            git_info_lines.append('Git            : not a repo')
    except Exception:
        git_info_lines.append('Git            : unavailable')

    # Settings export (se existir)
    settings_path = os.path.join(os.environ.get('APPDATA', ''), 'relatorios_2026', 'settings.json')

    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('git_info.txt', "\n".join(git_info_lines) + "\n")

        if settings_path and os.path.isfile(settings_path):
            try:
                zf.write(settings_path, arcname='settings.json')
            except Exception:
                pass

        for dirpath, dirs, files in os.walk(project_root):
            # prune excluded dirs
            dirs[:] = [d for d in dirs if d not in exclude_dirs]
            for filename in files:
                if is_excluded_name(filename):
                    continue
                full_path = os.path.join(dirpath, filename)
                rel = os.path.relpath(full_path, project_root)
                rel_dir = os.path.dirname(rel)
                if not should_include(rel_dir, filename):
                    continue
                arcname = rel.replace('\\', '/')
                try:
                    zf.write(full_path, arcname=arcname)
                except Exception:
                    pass

    # Retenção simples: manter só os N mais recentes
    try:
        backups = [
            os.path.join(backups_dir, f)
            for f in os.listdir(backups_dir)
            if f.startswith('backup_projeto_') and f.lower().endswith('.zip')
        ]
        backups.sort(reverse=True)
        for old in backups[keep_last:]:
            try:
                os.remove(old)
            except Exception:
                pass
    except Exception:
        pass

    return zip_path


def _get_weasyprint_html():
    """Importa WeasyPrint somente quando necessário.

    No Windows, o WeasyPrint depende de bibliotecas externas (GTK/Pango/Cairo).
    Se não estiverem instaladas, o app ainda deve abrir e avisar na hora de gerar PDF.
    """
    try:
        # Windows: tente localizar GTK runtime e adicionar o diretório de DLLs dinamicamente
        # para evitar depender do PATH do sistema.
        if os.name == 'nt' and hasattr(os, 'add_dll_directory'):
            possíveis = [
                r"C:\Program Files\GTK3-Runtime Win64\bin",
                r"C:\Program Files\GTK3-Runtime\bin",
                r"C:\Program Files (x86)\GTK3-Runtime Win64\bin",
                r"C:\Program Files (x86)\GTK3-Runtime\bin",
            ]
            for pasta in possíveis:
                try:
                    if os.path.isdir(pasta):
                        os.add_dll_directory(pasta)
                except Exception:
                    pass

        from weasyprint import HTML  # type: ignore
        return HTML
    except Exception as e:
        raise RuntimeError(
            "WeasyPrint não conseguiu carregar as bibliotecas externas (GTK/Pango/Cairo).\n\n"
            "Para corrigir no Windows, instale um runtime GTK3 (que forneça libgobject/pango/cairo) "
            "e o Microsoft Visual C++ Redistributable (2015-2022).\n\n"
            "Depois, reabra o programa e tente gerar o PDF novamente.\n\n"
            f"Detalhes técnicos: {e}"
        )


def _write_pdf(html_string, base_url, save_path):
    HTML = _get_weasyprint_html()
    HTML(string=html_string, base_url=base_url).write_pdf(save_path)


def _write_pdf_worker(html_string, base_url, save_path, result_queue):
    try:
        _write_pdf(html_string, base_url=base_url, save_path=save_path)
        result_queue.put({"ok": True})
    except Exception as e:
        import traceback
        result_queue.put({
            "ok": False,
            "error": str(e),
            "trace": traceback.format_exc(),
        })


def _write_pdf_safe(html_string, base_url, save_path, *, cancel_event=None, timeout_seconds=600):
    """Gera PDF em subprocesso para evitar travar a UI e permitir cancelamento/timeout."""
    q = multiprocessing.Queue()
    p = multiprocessing.Process(
        target=_write_pdf_worker,
        args=(html_string, base_url, save_path, q),
    )
    p.start()
    start = time.time()

    while p.is_alive():
        if cancel_event is not None and cancel_event.is_set():
            try:
                p.terminate()
            except Exception:
                pass
            p.join(timeout=5)
            try:
                if os.path.exists(save_path):
                    os.remove(save_path)
            except Exception:
                pass
            raise CancelledError()

        if timeout_seconds and (time.time() - start) > float(timeout_seconds):
            try:
                p.terminate()
            except Exception:
                pass
            p.join(timeout=5)
            try:
                if os.path.exists(save_path):
                    os.remove(save_path)
            except Exception:
                pass
            raise RuntimeError("Timeout ao gerar PDF (WeasyPrint demorou demais).")

        time.sleep(0.2)

    p.join()

    result = None
    try:
        result = q.get_nowait()
    except Exception:
        result = None

    if isinstance(result, dict) and result.get("ok") is True:
        return
    if isinstance(result, dict) and result.get("ok") is False:
        raise RuntimeError(f"Falha ao gerar PDF: {result.get('error')}\n\n{result.get('trace')}")
    # Se não houve retorno no queue, assume falha genérica
    raise RuntimeError("Falha ao gerar PDF (processo encerrou sem retorno).")


def _resolve_output_paths(save_path, output_mode):
    """Retorna (pdf_path, html_path) com base no output_mode.

    - pdf:        salva apenas PDF em save_path
    - pdf_html:   salva PDF em save_path e HTML ao lado (mesmo nome, .html)
    - html:       salva apenas HTML em save_path
    """
    mode = (output_mode or 'pdf').strip().lower()
    if mode == 'html':
        return None, save_path
    if mode == 'pdf_html':
        base, _ = os.path.splitext(save_path)
        return save_path, base + '.html'
    # default: pdf
    return save_path, None


def _export_html_with_assets(html_string, assets_dir, html_path):
    """Exporta HTML para html_path e copia os assets (imagens) para uma pasta ao lado.

    Isso permite abrir o HTML no navegador sem depender da pasta temporária.
    """
    if not html_path:
        return None

    out_dir = os.path.dirname(os.path.abspath(html_path))
    os.makedirs(out_dir, exist_ok=True)

    assets_subdir = os.path.splitext(os.path.basename(html_path))[0] + '_assets'
    assets_out_dir = os.path.join(out_dir, assets_subdir)
    os.makedirs(assets_out_dir, exist_ok=True)

    rewritten = html_string
    try:
        if assets_dir and os.path.isdir(assets_dir):
            import re
            for name in os.listdir(assets_dir):
                src = os.path.join(assets_dir, name)
                if not os.path.isfile(src):
                    continue
                dst = os.path.join(assets_out_dir, name)
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    continue

                # Reescreve apenas src='name' ou src="name" -> src="<subdir>/name"
                escaped = re.escape(name)
                pattern = re.compile(r'(src\s*=\s*)(["\"])' + escaped + r'\2', flags=re.IGNORECASE)
                rewritten = pattern.sub(r'\1\2' + assets_subdir.replace('\\', '/') + '/' + name + r'\2', rewritten)
    except Exception:
        # Se falhar, ainda tenta salvar o HTML sem rewrite.
        rewritten = html_string

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(rewritten)

    return html_path


def _html_fallback_beside_pdf(pdf_path: str) -> str:
    base, _ = os.path.splitext(os.path.abspath(pdf_path))
    return base + '_fallback.html'


def _export_pdf_with_html_fallback(
    *,
    html_final: str,
    assets_dir,
    save_path: str,
    output_mode,
    timings,
    update_status,
    cancel_event=None,
    timeout_pdf: int = 900,
):
    """Salva HTML auxiliar quando `pdf_html`, tenta gerar PDF; se falhar, mantém/suplementa HTML útil.

    Retorna o caminho do artefato preferencial para o usuário abrir por último
    (prioriza PDF se existiu; caso contrário o HTML lado a lado ou o *_fallback.html).
    """
    pdf_path, html_aux_path = _resolve_output_paths(save_path, output_mode)
    mode = (output_mode or 'pdf').strip().lower()

    if html_aux_path:
        update_status("Salvando HTML...")
        t_h = time.perf_counter()
        _export_html_with_assets(html_final, assets_dir, html_aux_path)
        try:
            timings.add("Salvar HTML", time.perf_counter() - t_h)
        except Exception:
            pass

    if not pdf_path:
        effective = html_aux_path
        return effective

    update_status("Gerando PDF (isso pode levar alguns instantes)...")
    t_pdf = time.perf_counter()
    try:
        _write_pdf_safe(
            html_final,
            base_url=assets_dir,
            save_path=pdf_path,
            cancel_event=cancel_event,
            timeout_seconds=timeout_pdf,
        )
        try:
            timings.add("PDF", time.perf_counter() - t_pdf)
        except Exception:
            pass
        return pdf_path
    except CancelledError:
        try:
            if os.path.isfile(pdf_path):
                os.remove(pdf_path)
        except Exception:
            pass
        raise
    except Exception as e:
        try:
            timings.add("PDF", time.perf_counter() - t_pdf)
        except Exception:
            pass
        err_one = str(e).strip().splitlines()[0][:220] if str(e).strip() else str(type(e).__name__)
        update_status(f"AVISO: geração de PDF falhou ({err_one}).")
        update_status(
            "Contingência: usando ou gerando relatório HTML (abra no navegador). "
            "Para PDF no Windows, confira GTK/WeasyPrint instalados conforme README do projeto."
        )
        if html_aux_path and os.path.isfile(html_aux_path):
            return html_aux_path
        fb = _html_fallback_beside_pdf(pdf_path)
        _export_html_with_assets(html_final, assets_dir, fb)
        return fb


def _format_seconds_short(seconds: float) -> str:
    try:
        s = int(round(float(seconds)))
    except Exception:
        return "-"

    if s < 0:
        s = 0

    if s < 60:
        return f"{s}s"

    m, sec = divmod(s, 60)
    if m < 60:
        return f"{m}m {sec:02d}s"

    h, m = divmod(m, 60)
    return f"{h}h {m:02d}m {sec:02d}s"


class TimingCollector:
    """Coleta tempos por etapa sem afetar a lógica de geração."""

    def __init__(self):
        self._times = {}

    @contextmanager
    def stage(self, name: str):
        t0 = time.perf_counter()
        try:
            yield
        finally:
            self.add(name, time.perf_counter() - t0)

    def add(self, name: str, seconds: float):
        try:
            if not name:
                return
            v = float(seconds)
            if v < 0:
                v = 0.0
            self._times[name] = float(self._times.get(name, 0.0)) + v
        except Exception:
            pass

    def format_summary(self) -> str:
        try:
            if not self._times:
                return ""
            total = 0.0
            parts = []
            for k, v in self._times.items():
                total += float(v)
                parts.append(f"{k}: {_format_seconds_short(v)}")
            parts.append(f"Total: {_format_seconds_short(total)}")
            return " | ".join(parts)
        except Exception:
            return ""

    def log_to(self, update_status, prefix: str = "Tempos por etapa"):
        try:
            summary = self.format_summary()
            if summary:
                update_status(f"{prefix}: {summary}")
        except Exception:
            pass


def _get_settings_path():
    appdata = os.environ.get('APPDATA') or os.path.expanduser('~')
    pasta = os.path.join(appdata, 'relatorios_2026')
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, 'settings.json')


def load_settings():
    try:
        with open(_get_settings_path(), 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(settings):
    try:
        with open(_get_settings_path(), 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _default_indicator_settings():
    # Flags por relatório: mantemos KPIs sempre ligados e deixamos o resto opcional.
    return {
        'periodo': {
            'resumo_exec': True,
            'vendas_tempo': True,
            'vendedores': True,
            'clientes': True,
            'produtos': True,
            'mapa_calor': True,
            'geo_detalhada': True,
        },
        'comparativo': {
            'graf_vendedor': True,
            'graf_produtos': True,
            'graf_clientes': True,
            'churn': False,  # padrão: mensal OFF
        },
        'anual': {
            'evolucao': True,
            'pedidos_ticket': True,
            'rank_produtos': True,
            'rank_clientes': True,
            'vendedores': True,
            'uf': True,
        },
        'comparativo_anual': {
            'evolucao_mensal': True,
            'tabela_yoy_mensal': True,
            'top_produtos': True,
            'top_clientes': True,
            'churn': True,  # padrão: anual ON
        },
    }


def get_indicator_settings():
    settings = load_settings()
    defaults = _default_indicator_settings()
    data = settings.get('indicators')
    if not isinstance(data, dict):
        return defaults

    merged = {}
    for group, group_defaults in defaults.items():
        group_data = data.get(group)
        if not isinstance(group_data, dict):
            merged[group] = group_defaults
            continue
        out = dict(group_defaults)
        for k in group_defaults.keys():
            if k in group_data:
                out[k] = bool(group_data.get(k))
        merged[group] = out
    return merged


def set_indicator_settings(indicators):
    if not isinstance(indicators, dict):
        return
    settings = load_settings()
    settings['indicators'] = indicators
    save_settings(settings)


def get_last_dir():
    settings = load_settings()
    last_dir = settings.get('last_dir')
    if last_dir and os.path.isdir(last_dir):
        return last_dir
    return os.path.abspath('.')


def set_last_dir(path):
    try:
        pasta = path if os.path.isdir(path) else os.path.dirname(path)
        if pasta and os.path.isdir(pasta):
            settings = load_settings()
            settings['last_dir'] = pasta
            save_settings(settings)
    except Exception:
        pass


def get_last_output_path():
    settings = load_settings()
    p = settings.get('last_output_path')
    if p and isinstance(p, str) and os.path.isfile(p):
        return p
    return None


def set_last_output_path(path):
    try:
        if not path or not isinstance(path, str):
            return
        if not os.path.isfile(path):
            return
        settings = load_settings()
        settings['last_output_path'] = path
        save_settings(settings)
    except Exception:
        pass


def get_output_mode():
    settings = load_settings()
    mode = settings.get('output_mode')
    if mode in ('pdf', 'pdf_html', 'html'):
        return mode
    return 'pdf'


def set_output_mode(mode):
    try:
        mode = (mode or '').strip().lower()
        if mode not in ('pdf', 'pdf_html', 'html'):
            return
        settings = load_settings()
        settings['output_mode'] = mode
        save_settings(settings)
    except Exception:
        pass


def _norm_header_value(v):
    try:
        if v is None:
            return ''
        return str(v).strip()
    except Exception:
        return ''


def _norm_key(s: str) -> str:
    try:
        return (s or '').strip().casefold()
    except Exception:
        return ''


def _excel_get_header(filepath: str):
    """Lê apenas o cabeçalho (linha 1) do Excel via openpyxl."""
    try:
        from openpyxl import load_workbook  # type: ignore
        wb = load_workbook(filepath, read_only=True, data_only=True)
        try:
            ws = wb.active
            rows = ws.iter_rows(min_row=1, max_row=1, values_only=True)
            first = next(rows, None)
            if not first:
                return []
            return [_norm_header_value(v) for v in list(first)]
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        return []


def _excel_find_col_index(header, col_name: str):
    target = _norm_key(col_name)
    for idx, h in enumerate(header or []):
        if _norm_key(h) == target:
            return idx + 1  # 1-based (openpyxl)
    return None


def excel_quick_missing_columns(filepath: str, required_cols):
    header = _excel_get_header(filepath)
    header_norm = {_norm_key(h) for h in header if h}
    missing = []
    for c in (required_cols or []):
        if _norm_key(c) not in header_norm:
            missing.append(c)
    return missing, header


def _try_parse_excel_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    # Excel serial number
    if isinstance(v, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel  # type: ignore
            return from_excel(v)
        except Exception:
            return None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        fmts = (
            '%d/%m/%Y',
            '%d/%m/%Y %H:%M',
            '%d/%m/%Y %H:%M:%S',
            '%Y-%m-%d',
            '%Y-%m-%d %H:%M',
            '%Y-%m-%d %H:%M:%S',
        )
        for fmt in fmts:
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                continue
        return None
    return None


def excel_collect_summary(filepath: str, *, want_date=True, want_uf=True, sample_rows=4000):
    """Coleta um resumo leve do arquivo sem carregar pandas.

    - linhas: usa ws.max_row
    - período: amostra começo/fim (aproximado) pela coluna Dt.Pedido
    - UFs: contagem simples por amostragem
    """
    out = {
        'sheet': None,
        'rows': None,
        'date_min': None,
        'date_max': None,
        'uf_top': [],
        'header': [],
    }
    try:
        from openpyxl import load_workbook  # type: ignore
        wb = load_workbook(filepath, read_only=True, data_only=True)
        try:
            ws = wb.active
            out['sheet'] = getattr(ws, 'title', None)
            max_row = getattr(ws, 'max_row', None)
            if isinstance(max_row, int) and max_row >= 1:
                out['rows'] = max(0, max_row - 1)

            header = _excel_get_header(filepath)
            out['header'] = header

            dt_col = _excel_find_col_index(header, 'Dt.Pedido') if want_date else None
            uf_col = _excel_find_col_index(header, 'UF') if want_uf else None

            if want_date and dt_col:
                # amostra começo
                start_end = min(max_row or 2, 1 + sample_rows)
                for row in ws.iter_rows(min_row=2, max_row=start_end, min_col=dt_col, max_col=dt_col, values_only=True):
                    d = _try_parse_excel_date(row[0] if row else None)
                    if d:
                        out['date_min'] = d if not out['date_min'] else min(out['date_min'], d)
                        out['date_max'] = d if not out['date_max'] else max(out['date_max'], d)

                # amostra fim
                if isinstance(max_row, int) and max_row > (2 + sample_rows):
                    start_last = max(2, max_row - sample_rows + 1)
                    for row in ws.iter_rows(min_row=start_last, max_row=max_row, min_col=dt_col, max_col=dt_col, values_only=True):
                        d = _try_parse_excel_date(row[0] if row else None)
                        if d:
                            out['date_min'] = d if not out['date_min'] else min(out['date_min'], d)
                            out['date_max'] = d if not out['date_max'] else max(out['date_max'], d)

            if want_uf and uf_col:
                counts = {}
                end_row = min(max_row or 2, 1 + min(sample_rows, 2500))
                for row in ws.iter_rows(min_row=2, max_row=end_row, min_col=uf_col, max_col=uf_col, values_only=True):
                    v = row[0] if row else None
                    s = _norm_header_value(v)
                    if not s:
                        continue
                    s = s.strip().upper()
                    if s in ('N/D', 'ND', 'N\\D'):
                        continue
                    counts[s] = counts.get(s, 0) + 1
                out['uf_top'] = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:10]

            return out
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        return out

# --- CONSTANTES GLOBAIS ---
NOME_FICHEIRO_LOGO = "Logo.png"
CORES_GRAFICOS = ['#3498db', '#e74c3c', '#2ecc71', '#f1c40f', '#9b59b6', '#34495e', '#1abc9c', '#d35400', '#c0392b', '#8e44ad']

UF_VALIDAS = {
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
    'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
}


def formatar_moeda_str(valor):
    return f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

def obter_caminho_logo_absoluto():
    caminho_abs = resource_path(NOME_FICHEIRO_LOGO)
    if os.path.exists(caminho_abs):
        return pathlib.Path(caminho_abs).as_uri()
    return ""

def _asset_path(assets_dir, filename):
    return os.path.join(assets_dir, filename)


def _render_template(template_path, context):
    with open(template_path, 'r', encoding='utf-8') as f:
        template_html = f.read()

    if Environment is None:
        # Fallback: mantém compatibilidade caso Jinja2 não esteja instalado.
        import re

        def _eval_if_expr(expr: str) -> bool:
            """Avalia expressões booleanas simples do Jinja2 no modo fallback.

            Suporta:
            - VAR
            - not VAR
            - VAR and VAR
            - VAR or VAR

            Não suporta parênteses/operadores complexos.
            """
            try:
                s = (expr or '').strip()
                if not s:
                    return False

                # normaliza espaços
                s = re.sub(r"\s+", " ", s)

                def eval_atom(atom: str) -> bool:
                    atom = atom.strip()
                    neg = False
                    if atom.lower().startswith('not '):
                        neg = True
                        atom = atom[4:].strip()
                    val = bool(context.get(atom))
                    return (not val) if neg else val

                # OR tem menor precedência
                or_parts = [p.strip() for p in re.split(r"\s+or\s+", s, flags=re.IGNORECASE)]
                result_or = False
                for part in or_parts:
                    and_parts = [p.strip() for p in re.split(r"\s+and\s+", part, flags=re.IGNORECASE)]
                    result_and = True
                    for atom in and_parts:
                        if atom:
                            result_and = result_and and eval_atom(atom)
                    result_or = result_or or result_and
                return bool(result_or)
            except Exception:
                return False

        html_final = template_html

        # Suporte mínimo a condicionais Jinja2 (inclui expressões simples com or/and/not).
        # Isso evita que as tags apareçam no PDF quando o projeto roda sem Jinja2.
        pattern = re.compile(
            r"\{%\s*if\s+(?P<expr>[^%]+?)\s*%\}(?P<then>.*?)"
            r"(?:\{%\s*else\s*%\}(?P<else>.*?))?\{%\s*endif\s*%\}",
            flags=re.DOTALL,
        )

        for _ in range(50):
            m = pattern.search(html_final)
            if not m:
                break
            expr = (m.group('expr') or '').strip()
            then_block = m.group('then') or ''
            else_block = m.group('else') or ''
            replacement = then_block if _eval_if_expr(expr) else else_block
            html_final = html_final[:m.start()] + replacement + html_final[m.end():]

        for k, v in context.items():
            html_final = html_final.replace('{{' + k + '}}', str(v))
        return html_final

    env = Environment(
        loader=BaseLoader(),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.from_string(template_html)
    return template.render(**context)


def _require_columns(df, required, nome_relatorio):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"O ficheiro selecionado não tem as colunas necessárias para {nome_relatorio}. "
            f"Faltando: {', '.join(missing)}"
        )


def _clean_dataframe_base(df, *, nome_relatorio, log_queue=None):
    required = ['Dt.Pedido', 'Vlr.Total', 'Nro.Pedido']
    _require_columns(df, required, nome_relatorio)

    df = df.copy()
    df['Dt.Pedido'] = pd.to_datetime(df['Dt.Pedido'], errors='coerce')
    df['Vlr.Total'] = pd.to_numeric(df['Vlr.Total'], errors='coerce')
    df['Nro.Pedido'] = df['Nro.Pedido'].astype(str)

    before = len(df)
    df = df.dropna(subset=['Dt.Pedido', 'Vlr.Total'])
    dropped = before - len(df)
    if dropped and log_queue is not None:
        log_queue.put(f"Aviso: {dropped} linhas inválidas foram descartadas (data/valor).")

    # Campos opcionais: se existirem, padroniza
    for col in ['Cliente', 'Vendedor', 'UF', 'Descrição']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if 'Qtde' in df.columns:
        df['Qtde'] = pd.to_numeric(df['Qtde'], errors='coerce').fillna(0)
    if 'Cod.Material' in df.columns:
        df['Cod.Material'] = df['Cod.Material'].astype(str).str.strip()

    return df


def _produto_label_map(*dfs):
    """Constrói mapeamento Cod.Material → Descrição mais frequente.

    Aceita um ou mais DataFrames. Em empate de frequência, usa a descrição
    que aparece por último na ordem de iteração (determinístico).
    """
    counts: dict[str, Counter] = {}
    last_pos: dict[tuple[str, str], int] = {}
    idx = 0
    for df in dfs:
        if 'Cod.Material' not in df.columns or 'Descrição' not in df.columns:
            continue
        for code, desc in zip(df['Cod.Material'], df['Descrição']):
            c = str(code).strip()
            d = str(desc).strip()
            if c and c not in ('nan', 'None', ''):
                counts.setdefault(c, Counter())[d] += 1
                last_pos[(c, d)] = idx
            idx += 1

    def _pick_label(mat: str, ctr: Counter) -> str:
        max_ct = max(ctr.values())
        tied = [d for d, n in ctr.items() if n == max_ct]
        return max(tied, key=lambda d: last_pos.get((mat, d), -1))

    return {mat: _pick_label(mat, ctr) for mat, ctr in counts.items()}


_MAPA_CACHE = None


def _load_mapa_cache():
    global _MAPA_CACHE
    if _MAPA_CACHE is not None:
        return _MAPA_CACHE
    import pickle
    caminho_mapa_pkl = resource_path("mapa_brasil_dados.pkl")
    with open(caminho_mapa_pkl, 'rb') as f:
        _MAPA_CACHE = pickle.load(f)
    return _MAPA_CACHE

def gerar_mapa_calor(df_vendas_por_estado, log_queue, assets_dir, cancel_event=None):
    import geopandas
    
    log_queue.put("Iniciando geração do mapa de calor...")
    try:
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()

        dados_mapa = _load_mapa_cache()
        
        gdf_brasil = geopandas.GeoDataFrame(list(dados_mapa.keys()), geometry=list(dados_mapa.values()), columns=['abbrev_state'])
        
        df_vendas_mapa = df_vendas_por_estado.reset_index()
        df_vendas_mapa.rename(columns={'UF': 'abbrev_state', 'Vlr.Total': 'Valor'}, inplace=True)
        
        mapa_com_dados = gdf_brasil.merge(df_vendas_mapa, on='abbrev_state', how='left')
        mapa_com_dados['Valor'] = mapa_com_dados['Valor'].fillna(0)

        fig, ax = plt.subplots(1, 1, figsize=(10, 10))
        mapa_com_dados.plot(column='Valor', cmap='viridis', linewidth=0.8, ax=ax, edgecolor='0.8', legend=True,
                                legend_kwds={'label': "Valor Total de Vendas (R$)", 'orientation': "horizontal", 'shrink': 0.6})
        
        ax.set_axis_off()
        ax.set_title('Distribuição de Vendas por Estado', fontdict={'fontsize': '16', 'fontweight' : '3'})
        
        nome_grafico = 'grafico_mapa_calor.png'
        plt.savefig(_asset_path(assets_dir, nome_grafico), dpi=150)
        plt.close(fig)
        log_queue.put("Mapa de calor gerado com sucesso.")
        return nome_grafico
    except CancelledError:
        log_queue.put("Geração do mapa cancelada pelo usuário.")
        return None
    except Exception as e:
        log_queue.put(f"ERRO ao gerar mapa de calor: {e}")
        return None

# --- FUNÇÃO 1 (MODIFICADA PARA TAREFA 1) ---
def gerar_relatorio_completo(
    caminho_dados_consolidados,
    log_queue,
    save_path,
    cancel_event=None,
    progress_cb=None,
    output_mode='pdf',
    include_resumo_exec=True,
    include_vendas_tempo=True,
    include_vendedores=True,
    include_clientes=True,
    include_produtos=True,
    include_mapa_calor=True,
    include_geo_detalhada=True,
):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    timings = TimingCollector()
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        if progress_cb:
            progress_cb(2)
        update_status("Iniciando o processo de geração de relatório...")
        update_status(f"Lendo ficheiro de dados consolidados: {os.path.basename(caminho_dados_consolidados)}")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()

        with timings.stage("Leitura Excel"):
            df_final = pd.read_excel(caminho_dados_consolidados)
        with timings.stage("Limpeza/Validação"):
            df_final = _clean_dataframe_base(df_final, nome_relatorio="Relatório de Período", log_queue=log_queue)
            required_extra = ['Vendedor', 'UF', 'Cliente', 'Cod.Material', 'Descrição', 'Qtde']
            _require_columns(df_final, required_extra, "Relatório de Período")

        update_status("Ficheiro de dados lido com sucesso.")

        if progress_cb:
            progress_cb(10)

        periodo_label = df_final['Dt.Pedido'].min().strftime('%B de %Y').capitalize() if not df_final.empty else "Período Geral"
        
        update_status(f"Período identificado: {periodo_label}")
        update_status("Iniciando análises...")

        t_analises = time.perf_counter()
        
        analise_vendedor = None
        if include_vendedores:
            analise_vendedor = df_final.groupby('Vendedor').agg(
                Valor_Total_Vendido=('Vlr.Total', 'sum'),
                Num_Pedidos=('Nro.Pedido', 'nunique')
            ).sort_values(by='Valor_Total_Vendido', ascending=False)
            analise_vendedor['Ticket_Medio'] = analise_vendedor['Valor_Total_Vendido'] / analise_vendedor['Num_Pedidos']
        
        vendas_por_estado = df_final[df_final['UF'] != 'N/D'].groupby('UF')['Vlr.Total'].sum().sort_values(ascending=False)
        vendas_diarias = None
        if include_vendas_tempo:
            vendas_diarias = df_final.groupby(df_final['Dt.Pedido'].dt.date)['Vlr.Total'].sum()

        vendas_por_produto_valor = None
        vendas_por_produto_qtd = None
        if include_produtos:
            _lmap = _produto_label_map(df_final)
            _top_val = df_final.groupby('Cod.Material')['Vlr.Total'].sum().sort_values(ascending=False).head(20)
            _top_qtd = df_final.groupby('Cod.Material')['Qtde'].sum().sort_values(ascending=False).head(20)
            vendas_por_produto_valor = _top_val.rename(index=lambda c: _lmap.get(c, c))
            vendas_por_produto_qtd = _top_qtd.rename(index=lambda c: _lmap.get(c, c))

        vendas_por_cliente_todos = None
        vendas_por_cliente_top20 = None
        if include_clientes:
            vendas_por_cliente_todos = df_final.groupby('Cliente')['Vlr.Total'].sum().sort_values(ascending=False)
            vendas_por_cliente_top20 = vendas_por_cliente_todos.head(20)

        df_pareto = None
        df_pareto_grafico = None
        if include_clientes and vendas_por_cliente_todos is not None and not vendas_por_cliente_todos.empty:
            df_pareto = vendas_por_cliente_todos.to_frame()
            df_pareto.rename(columns={'Vlr.Total': 'Valor_Total'}, inplace=True)
            df_pareto['% Acumulado'] = df_pareto['Valor_Total'].cumsum() / df_pareto['Valor_Total'].sum() * 100
            df_pareto_grafico = df_pareto.head(30)

        try:
            timings.add("Análises", time.perf_counter() - t_analises)
        except Exception:
            pass

        update_status("Análises concluídas. Gerando gráficos...")

        t_graficos = time.perf_counter()

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(20)
        
        nome_grafico_pareto_cliente = ''
        if include_clientes and df_pareto_grafico is not None and not df_pareto_grafico.empty:
            nome_grafico_pareto_cliente = 'grafico_pareto_cliente.png'
            fig_pareto, ax_pareto = plt.subplots(figsize=(12, 6))
            ax_pareto.bar(df_pareto_grafico.index, df_pareto_grafico['Valor_Total'], color="C0")
            ax_pareto.tick_params(axis='x', rotation=90, labelsize=8)
            ax2_pareto = ax_pareto.twinx()
            ax2_pareto.plot(df_pareto_grafico.index, df_pareto_grafico["% Acumulado"], color="C1", marker="o", ms=5)
            ax2_pareto.yaxis.set_major_formatter(_get_mpl_PercentFormatter()())
            ax_pareto.set_title("Análise de Pareto por Cliente (Top 30)", fontsize=16)
            fig_pareto.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_pareto_cliente), dpi=150)
            plt.close(fig_pareto)
            update_status("Gráfico 'Análise de Pareto' criado.")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(30)
        
        nome_grafico_vendas_tempo = ''
        if include_vendas_tempo and vendas_diarias is not None and not vendas_diarias.empty:
            nome_grafico_vendas_tempo = 'grafico_vendas_tempo.png'
            fig_tempo, ax_tempo = plt.subplots(figsize=(12, 6))
            vendas_diarias.plot(ax=ax_tempo, marker='o', linestyle='-')
            ax_tempo.set_title(f'Evolução das Vendas - {periodo_label}', fontsize=16)
            fig_tempo.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_vendas_tempo), dpi=150)
            plt.close(fig_tempo)
            update_status("Gráfico 'Vendas no Tempo' criado.")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(40)

        nome_grafico_vendedor = ''
        if include_vendedores and analise_vendedor is not None and not analise_vendedor.empty:
            nome_grafico_vendedor = 'grafico_vendas_vendedor.png'
            fig_vendedor, ax_vendedor = plt.subplots(figsize=(12, 8))
            sns.barplot(y=analise_vendedor.index, x=analise_vendedor['Valor_Total_Vendido'], ax=ax_vendedor, orient='h')
            ax_vendedor.set_title('Vendas por Vendedor', fontsize=16)
            for patch in ax_vendedor.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_vendedor.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
            ax_vendedor.set_xlim(right=ax_vendedor.get_xlim()[1] * 1.25)
            fig_vendedor.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_vendedor), dpi=150)
            plt.close(fig_vendedor)
            update_status("Gráfico 'Vendas por Vendedor' criado.")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(50)

        nome_grafico_clientes = ''
        if include_clientes and vendas_por_cliente_top20 is not None and not vendas_por_cliente_top20.empty:
            nome_grafico_clientes = 'grafico_top_clientes.png'
            fig_cliente, ax_cliente = plt.subplots(figsize=(12, 8))
            sns.barplot(y=vendas_por_cliente_top20.index, x=vendas_por_cliente_top20.values, ax=ax_cliente, orient='h')
            ax_cliente.set_title('Top 20 Clientes por Valor de Compra', fontsize=16)
            for patch in ax_cliente.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_cliente.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
            ax_cliente.set_xlim(right=ax_cliente.get_xlim()[1] * 1.25)
            fig_cliente.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_clientes), dpi=150)
            plt.close(fig_cliente)
            update_status("Gráfico 'Top 20 Clientes' criado.")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(60)
        
        nome_grafico_produtos_valor = ''
        nome_grafico_produtos_qtd = ''
        if include_produtos and vendas_por_produto_valor is not None and not vendas_por_produto_valor.empty:
            nome_grafico_produtos_valor = 'grafico_top_produtos_valor.png'
            fig_produto_valor, ax_prod_val = plt.subplots(figsize=(12, 10))
            vendas_por_produto_valor.sort_values().plot(kind='barh', ax=ax_prod_val)
            ax_prod_val.set_title('Top 20 Produtos por Valor de Venda (R$)', fontsize=16)
            for patch in ax_prod_val.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_prod_val.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
            ax_prod_val.set_xlim(right=ax_prod_val.get_xlim()[1] * 1.25)
            fig_produto_valor.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_produtos_valor), dpi=150)
            plt.close(fig_produto_valor)
            update_status("Gráfico 'Top 20 Produtos por Valor' criado.")

        if include_produtos and vendas_por_produto_qtd is not None and not vendas_por_produto_qtd.empty:
            nome_grafico_produtos_qtd = 'grafico_top_produtos_qtd.png'
            fig_produto_qtd, ax_prod_qtd = plt.subplots(figsize=(12, 10))
            vendas_por_produto_qtd.sort_values().plot(kind='barh', ax=ax_prod_qtd)
            ax_prod_qtd.set_title('Top 20 Itens por Quantidade Vendida (Unidades)', fontsize=16)
            for patch in ax_prod_qtd.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_prod_qtd.text(width * 1.01, y + height / 2, f'{int(width)} un', va='center')
            ax_prod_qtd.set_xlim(right=ax_prod_qtd.get_xlim()[1] * 1.25)
            fig_produto_qtd.tight_layout()
            plt.savefig(_asset_path(assets_dir, nome_grafico_produtos_qtd), dpi=150)
            plt.close(fig_produto_qtd)
            update_status("Gráfico 'Top 20 Itens por Quantidade' criado.")

        if progress_cb:
            progress_cb(70)
        
        nome_grafico_mapa_calor = None
        if include_mapa_calor:
            nome_grafico_mapa_calor = gerar_mapa_calor(vendas_por_estado, log_queue, assets_dir, cancel_event=cancel_event)
        nome_grafico_pizza_estado = None
        nome_grafico_outros_estados = None
        if include_geo_detalhada and (not vendas_por_estado.empty):
            top_5_estados = vendas_por_estado.head(5)
            soma_outros = vendas_por_estado.iloc[5:].sum()
            dados_pizza = pd.concat([top_5_estados, pd.Series({'Outros': soma_outros})]) if soma_outros > 0 else top_5_estados
            
            fig_pizza, ax_pizza = plt.subplots(figsize=(10, 7))
            total_pizza = float(dados_pizza.sum())

            def autopct_format_pizza(pct):
                valor = (pct / 100.0) * total_pizza
                return f"{pct:.1f}%\n({formatar_moeda_str(valor)})"

            ax_pizza.pie(
                dados_pizza,
                labels=dados_pizza.index,
                autopct=autopct_format_pizza,
                startangle=90,
            )
            ax_pizza.set_title('Distribuição Percentual de Vendas por Estado', fontsize=16)
            ax_pizza.axis('equal')
            nome_grafico_pizza_estado = 'grafico_pizza_estado.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_pizza_estado), dpi=150)
            plt.close(fig_pizza)
            update_status("Gráfico 'Pizza de Estados' criado.")
            
            outros_estados = vendas_por_estado.iloc[5:]
            if not outros_estados.empty:
                fig_outros, ax_outros = plt.subplots(figsize=(12, 6))
                sns.barplot(y=outros_estados.index, x=outros_estados.values, ax=ax_outros, orient='h')
                ax_outros.set_title('Detalhamento de Vendas - Outros Estados', fontsize=16)
                for patch in ax_outros.patches:
                    width = patch.get_width()
                    y = patch.get_y()
                    height = patch.get_height()
                    ax_outros.text(width * 1.01, y + height / 2, f' {formatar_moeda_str(width)}', va='center')
                ax_outros.set_xlim(right=ax_outros.get_xlim()[1] * 1.25)
                fig_outros.tight_layout()
                nome_grafico_outros_estados = 'grafico_outros_estados.png'
                plt.savefig(_asset_path(assets_dir, nome_grafico_outros_estados), dpi=150)
                plt.close(fig_outros)
                update_status("Gráfico 'Outros Estados' criado.")

        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if progress_cb:
            progress_cb(85)

        try:
            timings.add("Gráficos", time.perf_counter() - t_graficos)
        except Exception:
            pass

        update_status("Montando o relatório PDF...")
        t_html = time.perf_counter()
        template_path = resource_path('template.html')
            
        caminho_logo_final = obter_caminho_logo_absoluto()
        
        total_vendas = float(df_final['Vlr.Total'].sum())
        num_pedidos = int(df_final['Nro.Pedido'].nunique())
        ticket_medio = (total_vendas / num_pedidos) if num_pedidos > 0 else 0.0

        # --- RESUMO EXECUTIVO (meta mensal fixa + YoY) ---
        exec_periodo_str = ''
        exec_meta_mensal_str = ''
        exec_ating_meta_str = ''
        exec_clientes_str = ''
        exec_vendas_yoy_str = ''
        exec_vendas_yoy_delta_str = ''
        exec_dias_uteis_periodo_str = ''
        exec_vendas_dia_util_str = ''

        def _fmt_pct_ptbr(x):
            try:
                if x is None:
                    return '-'
                return f"{(float(x) * 100):.1f}%".replace('.', ',')
            except Exception:
                return '-'

        def _fmt_delta_pct_ptbr(x):
            try:
                if x is None:
                    return '-'
                v = float(x) * 100
                s = f"{v:+.1f}%"
                return s.replace('.', ',')
            except Exception:
                return '-'

        if include_resumo_exec:
            try:
                dt_min = pd.to_datetime(df_final['Dt.Pedido'].min()) if not df_final.empty else None
                dt_max = pd.to_datetime(df_final['Dt.Pedido'].max()) if not df_final.empty else None
            except Exception:
                dt_min = None
                dt_max = None

            if dt_min is not None and dt_max is not None:
                try:
                    start = pd.to_datetime(dt_min).normalize()
                    end = pd.to_datetime(dt_max).normalize()
                    exec_periodo_str = f"{start.strftime('%d/%m/%Y')} → {end.strftime('%d/%m/%Y')}"
                except Exception:
                    exec_periodo_str = ''

            # Vendas por dia útil (Mon–Fri)
            try:
                if dt_min is not None and dt_max is not None:
                    start = pd.to_datetime(dt_min).normalize()
                    end = pd.to_datetime(dt_max).normalize()
                    dias_uteis = int(pd.bdate_range(start, end).size)
                    exec_dias_uteis_periodo_str = str(dias_uteis) if dias_uteis >= 0 else '-'
                    if dias_uteis > 0:
                        exec_vendas_dia_util_str = formatar_moeda_str(total_vendas / dias_uteis)
                    else:
                        exec_vendas_dia_util_str = '-'
                else:
                    exec_dias_uteis_periodo_str = '-'
                    exec_vendas_dia_util_str = '-'
            except Exception:
                exec_dias_uteis_periodo_str = '-'
                exec_vendas_dia_util_str = '-'

            try:
                n_clientes = int(df_final['Cliente'].nunique()) if 'Cliente' in df_final.columns else None
                exec_clientes_str = str(n_clientes) if n_clientes is not None else '-'
            except Exception:
                exec_clientes_str = '-'

            # Meta mensal fixa (informada pelo usuário)
            meta_mensal = 2530000.0
            exec_meta_mensal_str = formatar_moeda_str(meta_mensal)
            exec_ating_meta_str = _fmt_pct_ptbr((total_vendas / meta_mensal) if meta_mensal > 0 else None)

            # Comparação YoY (mesmo intervalo do ano anterior)
            try:
                if dt_min is not None and dt_max is not None:
                    start = pd.to_datetime(dt_min).normalize()
                    end = pd.to_datetime(dt_max).normalize()
                    start_yoy = start - pd.DateOffset(years=1)
                    end_yoy = end - pd.DateOffset(years=1)
                    mask_yoy = (df_final['Dt.Pedido'] >= start_yoy) & (df_final['Dt.Pedido'] <= end_yoy)
                    df_yoy = df_final.loc[mask_yoy]
                    yoy_total = float(df_yoy['Vlr.Total'].sum()) if not df_yoy.empty else None
                else:
                    yoy_total = None

                if yoy_total is None or yoy_total == 0:
                    exec_vendas_yoy_str = '-'
                    exec_vendas_yoy_delta_str = 'YoY: sem base'
                else:
                    exec_vendas_yoy_str = formatar_moeda_str(yoy_total)
                    delta_abs = total_vendas - yoy_total
                    delta_pct = (total_vendas / yoy_total) - 1.0
                    exec_vendas_yoy_delta_str = f"YoY: {formatar_moeda_str(delta_abs)} ({_fmt_delta_pct_ptbr(delta_pct)})"
            except Exception:
                exec_vendas_yoy_str = '-'
                exec_vendas_yoy_delta_str = 'YoY: -'


        tabela_html = ''
        if include_vendedores and analise_vendedor is not None:
            for vendedor, row in analise_vendedor.iterrows():
                total_formatado = formatar_moeda_str(row['Valor_Total_Vendido'])
                pedidos_formatado = int(row['Num_Pedidos'])
                ticket_medio_formatado = formatar_moeda_str(row['Ticket_Medio'])
                tabela_html += f"<tr><td>{vendedor}</td><td style='text-align:right;'>{total_formatado}</td><td style='text-align:center;'>{pedidos_formatado}</td><td style='text-align:right;'>{ticket_medio_formatado}</td></tr>"

        # --- LÓGICA PARA O RESUMO DE PARETO ---
        texto_resumo_pareto = ""
        if include_clientes and df_pareto is not None and not df_pareto.empty:
            clientes_pareto_df = df_pareto[df_pareto['% Acumulado'] <= 80]
            num_clientes_pareto = len(clientes_pareto_df) + 1
            texto_resumo_pareto = f"Análise de Pareto: Cerca de 80% do valor total de vendas do período está concentrado em <strong>{num_clientes_pareto} clientes</strong>."
            update_status(f"Resumo de Pareto calculado: {num_clientes_pareto} clientes representam 80% das vendas.")
        
        secao_mapa = ''
        if include_mapa_calor and nome_grafico_mapa_calor:
            secao_mapa = f'<div class="new-page"><h2>Distribuição Geográfica das Vendas</h2><img src="{nome_grafico_mapa_calor}" alt="Mapa de Calor" class="chart"></div>'

        percentual_outros = 0.0
        lista_estados_outros = "Nenhum"
        grafico_outros_html = ''
        grafico_pizza_path = nome_grafico_pizza_estado or ''
        if include_geo_detalhada and nome_grafico_pizza_estado:
            percentual_outros = (soma_outros / vendas_por_estado.sum()) * 100 if vendas_por_estado.sum() > 0 else 0
            lista_estados_outros = ", ".join(outros_estados.index) if 'outros_estados' in locals() and not outros_estados.empty else "Nenhum"
            if nome_grafico_outros_estados:
                grafico_outros_html = f'<img src="{nome_grafico_outros_estados}" alt="Detalhamento de Outros Estados" class="chart">'
        else:
            grafico_pizza_path = ''

        context = {
            'TITULO_RELATORIO': f"Relatório de Vendas - {periodo_label}",
            'DATA_GERACAO': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'CAMINHO_LOGO': caminho_logo_final,
            'VALOR_TOTAL_VENDAS': formatar_moeda_str(total_vendas),
            'NUMERO_PEDIDOS': str(num_pedidos),
            'TICKET_MEDIO': formatar_moeda_str(ticket_medio),
            'EXEC_PERIODO': exec_periodo_str,
            'EXEC_META_MENSAL': exec_meta_mensal_str,
            'EXEC_ATINGIMENTO_META': exec_ating_meta_str,
            'EXEC_CLIENTES': exec_clientes_str,
            'EXEC_VENDAS_YOY': exec_vendas_yoy_str,
            'EXEC_VENDAS_YOY_DELTA': exec_vendas_yoy_delta_str,
            'EXEC_DIAS_UTEIS_PERIODO': exec_dias_uteis_periodo_str,
            'EXEC_VENDAS_DIA_UTIL': exec_vendas_dia_util_str,
            'GRAFICO_VENDAS_TEMPO_PATH': nome_grafico_vendas_tempo,
            'GRAFICO_VENDEDOR_PATH': nome_grafico_vendedor,
            'TABELA_VENDEDORES': Markup(tabela_html) if Markup else tabela_html,
            'GRAFICO_PARETO_CLIENTE_PATH': nome_grafico_pareto_cliente,
            'GRAFICO_CLIENTES_PATH': nome_grafico_clientes,
            'GRAFICO_PRODUTOS_VALOR_PATH': nome_grafico_produtos_valor,
            'GRAFICO_PRODUTOS_QTD_PATH': nome_grafico_produtos_qtd,
            'TEXTO_RESUMO_PARETO': Markup(texto_resumo_pareto) if Markup else texto_resumo_pareto,
            'SECAO_MAPA_CALOR': Markup(secao_mapa) if Markup else secao_mapa,
            'GRAFICO_PIZZA_ESTADO_PATH': grafico_pizza_path,
            'PERCENTUAL_OUTROS': f'{percentual_outros:.1f}',
            'LISTA_ESTADOS_OUTROS': lista_estados_outros,
            'GRAFICO_OUTROS_ESTADOS_HTML': Markup(grafico_outros_html) if Markup else grafico_outros_html,
            'INCLUDE_RESUMO_EXECUTIVO': bool(include_resumo_exec),
            'INCLUDE_VENDAS_TEMPO': bool(include_vendas_tempo),
            'INCLUDE_VENDEDORES': bool(include_vendedores),
            'INCLUDE_CLIENTES': bool(include_clientes),
            'INCLUDE_PRODUTOS': bool(include_produtos),
            'INCLUDE_MAPA_CALOR': bool(include_mapa_calor),
            'INCLUDE_GEO_DETALHADA': bool(include_geo_detalhada),
        }

        html_final = _render_template(template_path, context)
        try:
            timings.add("HTML", time.perf_counter() - t_html)
        except Exception:
            pass

        effective_path = _export_pdf_with_html_fallback(
            html_final=html_final,
            assets_dir=assets_dir,
            save_path=save_path,
            output_mode=output_mode,
            timings=timings,
            update_status=update_status,
            cancel_event=cancel_event,
        )
        pdf_out, html_side = _resolve_output_paths(save_path, output_mode)
        op_mode = (output_mode or 'pdf').strip().lower()
        if op_mode == 'html' and effective_path:
            update_status(f"Relatório HTML '{os.path.basename(effective_path)}' gerado com sucesso!")
        elif effective_path and pdf_out and os.path.abspath(effective_path) == os.path.abspath(pdf_out):
            if html_side and os.path.isfile(html_side):
                update_status(
                    f"PDF e HTML gerados com sucesso: '{os.path.basename(pdf_out)}' e '{os.path.basename(html_side)}'."
                )
            else:
                update_status(f"Relatório PDF '{os.path.basename(pdf_out)}' gerado com sucesso!")
        elif effective_path:
            update_status(f"Relatório salvo em HTML: '{os.path.basename(effective_path)}' (PDF indisponível nesta máquina).")

        if progress_cb:
            progress_cb(100)
        return effective_path
    except CancelledError:
        update_status("Processo cancelado pelo usuário.")
        raise
        
    except Exception as e:
        update_status(f"ERRO CRÍTICO: {e}")
        pass  # mensagem à UI: thread principal / status
        raise e
    finally:
        try:
            timings.log_to(update_status)
        except Exception:
            pass
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)

# --- FUNÇÃO 2 (ORIGINAL E INTACTA) ---
def gerar_relatorio_comparativo(
    caminho_dados_A,
    caminho_dados_B,
    log_queue,
    save_path,
    cancel_event=None,
    progress_cb=None,
    output_mode='pdf',
    include_churn=True,
    include_graf_vendedor=True,
    include_graf_produtos=True,
    include_graf_clientes=True,
):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    timings = TimingCollector()
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("Iniciando relatório comparativo...")
        if progress_cb:
            progress_cb(2)

        _raise_if_cancelled(cancel_event)
        
        with timings.stage("Leitura Excel"):
            df_A_raw = pd.read_excel(caminho_dados_A)
            df_B_raw = pd.read_excel(caminho_dados_B)
        with timings.stage("Limpeza/Validação"):
            df_A = _clean_dataframe_base(df_A_raw, nome_relatorio="Comparativo", log_queue=log_queue)
            df_B = _clean_dataframe_base(df_B_raw, nome_relatorio="Comparativo", log_queue=log_queue)

        _raise_if_cancelled(cancel_event)
        if progress_cb:
            progress_cb(10)
        required_extra = ['Cliente', 'Vendedor', 'Cod.Material', 'Descrição', 'UF']
        _require_columns(df_A, required_extra, "Comparativo")
        _require_columns(df_B, required_extra, "Comparativo")
        update_status("Ficheiros de dados lidos com sucesso.")

        df_A['Dt.Pedido'] = pd.to_datetime(df_A['Dt.Pedido'])
        df_B['Dt.Pedido'] = pd.to_datetime(df_B['Dt.Pedido'])
        
        label_A = df_A['Dt.Pedido'].min().strftime('%B/%Y').capitalize()
        label_B = df_B['Dt.Pedido'].min().strftime('%B/%Y').capitalize()
        update_status(f"Comparando Período A ({label_A}) com Período B ({label_B}).")
        if progress_cb:
            progress_cb(15)

        _raise_if_cancelled(cancel_event)
        
        t_analises = time.perf_counter()

        # --- Análise de KPIs ---
        kpis = {}
        for df, label in [(df_A, 'A'), (df_B, 'B')]:
            vendas = df['Vlr.Total'].sum()
            pedidos = df['Nro.Pedido'].nunique()
            ticket = vendas / pedidos if pedidos > 0 else 0
            kpis[label] = {'vendas': vendas, 'pedidos': pedidos, 'ticket': ticket}

        def calcular_variacao(val_A, val_B):
            if val_A > 0:
                var = ((val_B - val_A) / val_A) * 100
                cor = "positive" if var >= 0 else "negative"
                return f"{var:,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.'), cor
            return "N/A", ""

        var_ven, cor_ven = calcular_variacao(kpis['A']['vendas'], kpis['B']['vendas'])
        var_ped, cor_ped = calcular_variacao(kpis['A']['pedidos'], kpis['B']['pedidos'])
        var_tic, cor_tic = calcular_variacao(kpis['A']['ticket'], kpis['B']['ticket'])
        
        tabela_churn_html = ""
        resumo_churn = ""
        if include_churn:
            # --- Análise de Churn ---
            update_status("Iniciando análise de Clientes Inativos (Churn)...")
            t_churn = time.perf_counter()
            _raise_if_cancelled(cancel_event)
            clientes_A = set(df_A['Cliente'].unique())
            clientes_B = set(df_B['Cliente'].unique())
            clientes_inativos = clientes_A - clientes_B

            if not clientes_inativos:
                tabela_churn_html = "<tr><td colspan='2'>Nenhum cliente importante ficou inativo neste período.</td></tr>"
                resumo_churn = "<p>Nenhum cliente relevante ficou inativo neste período.</p>"
            else:
                df_inativos = df_A[df_A['Cliente'].isin(clientes_inativos)]
                vendas_inativos = df_inativos.groupby('Cliente')['Vlr.Total'].sum().sort_values(ascending=False)
                valor_total_inativo = vendas_inativos.sum()

                resumo_churn = f"<p>No total, <strong>{len(clientes_inativos)} clientes</strong> que representavam <strong>{formatar_moeda_str(valor_total_inativo)}</strong> em vendas em {label_A} não compraram em {label_B}.</p>"

                for cliente, valor in vendas_inativos.head(15).items():
                    tabela_churn_html += f"<tr><td>{cliente}</td><td style='text-align:right;'>{formatar_moeda_str(valor)}</td></tr>"
            update_status(f"Análise de Churn concluída: {len(clientes_inativos)} clientes inativos.")
            try:
                timings.add("Churn", time.perf_counter() - t_churn)
            except Exception:
                pass
        else:
            update_status("Churn desativado (configuração do relatório).")

        _raise_if_cancelled(cancel_event)
        if progress_cb:
            progress_cb(25)

        # --- Gráficos Comparativos ---
        update_status("Gerando gráficos comparativos...")
        t_graficos = time.perf_counter()
        if progress_cb:
            progress_cb(30)

        _raise_if_cancelled(cancel_event)
        
        def formatar_moeda_simples(valor, pos=None):
            return f'R$ {int(valor):,.0f}'.replace(',', '.')
        
        def criar_grafico_comparativo(df1, df2, label1, label2, coluna_grupo, coluna_valor, titulo, nome_ficheiro, label_map=None):
            _raise_if_cancelled(cancel_event)
            dados1 = df1.groupby(coluna_grupo)[coluna_valor].sum().rename(label1)
            dados2 = df2.groupby(coluna_grupo)[coluna_valor].sum().rename(label2)

            if label_map:
                dados1.index = dados1.index.map(lambda c: label_map.get(str(c), str(c)))
                dados2.index = dados2.index.map(lambda c: label_map.get(str(c), str(c)))
                dados1.index.name = 'Descrição'
                dados2.index.name = 'Descrição'
                coluna_grupo = 'Descrição'

            df_comp = pd.merge(dados1, dados2, on=coluna_grupo, how='outer').fillna(0)
            df_comp['Total'] = df_comp[label1] + df_comp[label2]
            df_comp = df_comp.sort_values(by='Total', ascending=False).head(20).drop(columns=['Total'])

            df_grafico = df_comp.reset_index().melt(id_vars=coluna_grupo, var_name='Período', value_name='Vendas')

            fig, ax = plt.subplots(figsize=(12, 8))
            sns.barplot(data=df_grafico, y=coluna_grupo, x='Vendas', hue='Período', ax=ax)
            ax.set_title(titulo, fontsize=16)
            ax.xaxis.set_major_formatter(_get_mpl_FuncFormatter()(formatar_moeda_simples))

            for container in ax.containers:
                ax.bar_label(container, fmt=lambda x: formatar_moeda_simples(x) if x > 0 else '', padding=3, fontsize=8)

            ax.set_xlim(right=ax.get_xlim()[1] * 1.3)
            fig.tight_layout()
            _raise_if_cancelled(cancel_event)
            plt.savefig(_asset_path(assets_dir, nome_ficheiro), dpi=150)
            plt.close(fig)
            update_status(f"Gráfico '{titulo}' criado.")
            return nome_ficheiro

        nome_grafico_vendedor = ''
        nome_grafico_produto = ''
        nome_grafico_cliente = ''
        if include_graf_vendedor:
            _raise_if_cancelled(cancel_event)
            nome_grafico_vendedor = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Vendedor', 'Vlr.Total', 'Comparativo de Vendas por Vendedor', 'grafico_comp_vendedor.png')
            if progress_cb:
                progress_cb(45)
        if include_graf_produtos:
            _raise_if_cancelled(cancel_event)
            _lmap_comp = _produto_label_map(df_A, df_B)
            nome_grafico_produto = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Cod.Material', 'Vlr.Total', 'Comparativo de Top 20 Produtos', 'grafico_comp_produto.png', label_map=_lmap_comp)
            if progress_cb:
                progress_cb(60)
        if include_graf_clientes:
            _raise_if_cancelled(cancel_event)
            nome_grafico_cliente = criar_grafico_comparativo(df_A, df_B, label_A, label_B, 'Cliente', 'Vlr.Total', 'Comparativo de Top 20 Clientes', 'grafico_comp_cliente.png')
            if progress_cb:
                progress_cb(75)

        _raise_if_cancelled(cancel_event)

        try:
            timings.add("Análises", time.perf_counter() - t_analises)
        except Exception:
            pass

        try:
            timings.add("Gráficos", time.perf_counter() - t_graficos)
        except Exception:
            pass

        # Montagem do HTML
        update_status("Montando HTML do relatório...")
        t_html = time.perf_counter()
        template_path = resource_path('template_comparativo.html')
        
        caminho_logo_final = obter_caminho_logo_absoluto()
        
        context = {
            'TITULO_RELATORIO': f'Relatório Comparativo: {label_A} vs {label_B}',
            'DATA_GERACAO': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'CAMINHO_LOGO': caminho_logo_final,
            'LABEL_A': label_A,
            'LABEL_B': label_B,
            'VALOR_TOTAL_A': formatar_moeda_str(kpis['A']['vendas']),
            'VALOR_TOTAL_B': formatar_moeda_str(kpis['B']['vendas']),
            'VARIACAO_VALOR': var_ven,
            'COR_VARIACAO_VALOR': cor_ven,
            'NUM_PEDIDOS_A': str(kpis['A']['pedidos']),
            'NUM_PEDIDOS_B': str(kpis['B']['pedidos']),
            'VARIACAO_PEDIDOS': var_ped,
            'COR_VARIACAO_PEDIDOS': cor_ped,
            'TICKET_MEDIO_A': formatar_moeda_str(kpis['A']['ticket']),
            'TICKET_MEDIO_B': formatar_moeda_str(kpis['B']['ticket']),
            'VARIACAO_TICKET': var_tic,
            'COR_VARIACAO_TICKET': cor_tic,
            'GRAFICO_COMP_VENDEDOR_PATH': nome_grafico_vendedor,
            'GRAFICO_COMP_PRODUTO_PATH': nome_grafico_produto,
            'GRAFICO_COMP_CLIENTE_PATH': nome_grafico_cliente,
            'RESUMO_CHURN': Markup(resumo_churn) if Markup else resumo_churn,
            'TABELA_CHURN': Markup(tabela_churn_html) if Markup else tabela_churn_html,
            'INCLUDE_CHURN': bool(include_churn),
            'INCLUDE_COMP_VENDEDOR': bool(include_graf_vendedor),
            'INCLUDE_COMP_PRODUTOS': bool(include_graf_produtos),
            'INCLUDE_COMP_CLIENTES': bool(include_graf_clientes),
        }
        html_final = _render_template(template_path, context)
        try:
            timings.add("HTML", time.perf_counter() - t_html)
        except Exception:
            pass
        if progress_cb:
            progress_cb(85)

        _raise_if_cancelled(cancel_event)

        effective_path = _export_pdf_with_html_fallback(
            html_final=html_final,
            assets_dir=assets_dir,
            save_path=save_path,
            output_mode=output_mode,
            timings=timings,
            update_status=update_status,
            cancel_event=cancel_event,
        )
        pdf_out, html_side = _resolve_output_paths(save_path, output_mode)
        op_mode = (output_mode or 'pdf').strip().lower()
        if op_mode == 'html' and effective_path:
            update_status(f"Relatório Comparativo (HTML) '{os.path.basename(effective_path)}' gerado com sucesso!")
        elif effective_path and pdf_out and os.path.abspath(effective_path) == os.path.abspath(pdf_out):
            if html_side and os.path.isfile(html_side):
                update_status(
                    f"PDF e HTML gerados com sucesso: '{os.path.basename(pdf_out)}' e '{os.path.basename(html_side)}'."
                )
            else:
                update_status(f"Relatório Comparativo '{os.path.basename(pdf_out)}' gerado com sucesso!")
        elif effective_path:
            update_status(f"Comparativo salvo em HTML: '{os.path.basename(effective_path)}' (PDF indisponível nesta máquina).")

        if progress_cb:
            progress_cb(100)
        return effective_path
    except CancelledError:
        update_status("Processo cancelado pelo usuário.")
        raise
        
    except Exception as e:
        update_status(f"ERRO CRÍTICO: {e}")
        pass  # mensagem à UI: thread principal / status
        raise e
    finally:
        try:
            timings.log_to(update_status)
        except Exception:
            pass
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)

# --- FUNÇÃO 3 (ORIGINAL E INTACTA) ---
def gerar_relatorio_anual(
    caminho_dados_anuais,
    log_queue,
    save_path,
    cancel_event=None,
    progress_cb=None,
    output_mode='pdf',
    include_evolucao=True,
    include_pedidos_ticket=True,
    include_rank_produtos=True,
    include_rank_clientes=True,
    include_vendedores=True,
    include_uf=True,
):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    timings = TimingCollector()
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("Iniciando Análise Anual...")
        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(5)

        with timings.stage("Leitura Excel"):
            df_raw = pd.read_excel(caminho_dados_anuais)
        with timings.stage("Limpeza/Validação"):
            df = _clean_dataframe_base(df_raw, nome_relatorio="Análise Anual", log_queue=log_queue)
            required_extra = ['Cliente', 'Cod.Material', 'Descrição', 'Qtde', 'Nro.Pedido', 'Dt.Pedido', 'Vlr.Total']
            _require_columns(df, required_extra, "Análise Anual")
        update_status("Ficheiro de dados anuais lido com sucesso.")

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(12)

        t_analises = time.perf_counter()
        venda_total = df['Vlr.Total'].sum()
        pedidos_unicos = df['Nro.Pedido'].nunique()
        ticket_medio = venda_total / pedidos_unicos if pedidos_unicos > 0 else 0
        unidades_vendidas = df['Qtde'].sum()
        
        update_status("KPIs anuais calculados. Gerando gráficos...")

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(18)
        
        def formatar_inteiro(valor, pos=None):
            return f'R$ {int(valor):,.0f}'.replace(',', '.')

        t_graficos = time.perf_counter()
        nome_grafico_mensal = ''
        nome_grafico_trimestral = ''
        melhor_mes_label = '-'
        melhor_mes_valor = '-'
        pior_mes_label = '-'
        pior_mes_valor = '-'
        if include_evolucao:
            _raise_if_cancelled(cancel_event)
            # Gráfico de Evolução Mensal
            vendas_mensais = df.set_index('Dt.Pedido')['Vlr.Total'].resample('ME').sum()

            try:
                if not vendas_mensais.empty:
                    ts_best = vendas_mensais.idxmax()
                    best_val = float(vendas_mensais.loc[ts_best])
                    melhor_mes_label = pd.to_datetime(ts_best).strftime('%b')
                    melhor_mes_valor = formatar_moeda_str(best_val)

                    vendas_pos = vendas_mensais[vendas_mensais > 0]
                    if not vendas_pos.empty:
                        ts_worst = vendas_pos.idxmin()
                        worst_val = float(vendas_pos.loc[ts_worst])
                    else:
                        ts_worst = vendas_mensais.idxmin()
                        worst_val = float(vendas_mensais.loc[ts_worst])

                    pior_mes_label = pd.to_datetime(ts_worst).strftime('%b')
                    pior_mes_valor = formatar_moeda_str(worst_val)
            except Exception:
                melhor_mes_label = '-'
                melhor_mes_valor = '-'
                pior_mes_label = '-'
                pior_mes_valor = '-'

            fig_mensal, ax_mensal = plt.subplots(figsize=(12, 6))
            vendas_mensais.index = vendas_mensais.index.strftime('%b')
            sns.barplot(x=vendas_mensais.index, y=vendas_mensais.values, ax=ax_mensal, color='royalblue')
            ax_mensal.set_title('Evolução Mensal de Vendas', fontsize=16)
            for container in ax_mensal.containers:
                ax_mensal.bar_label(container, fmt=lambda x: formatar_inteiro(x) if x > 0 else '', padding=3)
            fig_mensal.tight_layout()
            nome_grafico_mensal = 'grafico_evolucao_mensal.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_mensal), dpi=150)
            plt.close(fig_mensal)
            update_status("Gráfico 'Evolução Mensal' criado.")

            _raise_if_cancelled(cancel_event)

            # Gráfico de Vendas Trimestral
            vendas_trimestrais = df.set_index('Dt.Pedido')['Vlr.Total'].resample('QE').sum()
            fig_trimestral, ax_trimestral = plt.subplots(figsize=(8, 8))

            def autopct_format_anual(pct):
                total = vendas_trimestrais.sum()
                val = int(round(pct * total / 100.0))
                return f'{pct:.1f}%\n({formatar_inteiro(val)})'

            ax_trimestral.pie(vendas_trimestrais, labels=[f"T{i.quarter}" for i in vendas_trimestrais.index], autopct=autopct_format_anual, startangle=90)
            ax_trimestral.set_title('Distribuição das Vendas por Trimestre', fontsize=16)
            ax_trimestral.axis('equal')
            nome_grafico_trimestral = 'grafico_vendas_trimestral.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_trimestral), dpi=150)
            plt.close(fig_trimestral)
            update_status("Gráfico 'Vendas por Trimestre' criado.")

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(40)

        nome_grafico_pedidos_ticket = ''
        if include_pedidos_ticket:
            _raise_if_cancelled(cancel_event)
            # Pedidos por mês + Ticket médio por mês
            df_m = df.set_index('Dt.Pedido').resample('ME').agg(
                vendas=('Vlr.Total', 'sum'),
                pedidos=('Nro.Pedido', pd.Series.nunique),
            )
            # Garante 12 meses (caso falte algum mês no arquivo)
            if not df_m.empty:
                ano_ref = int(df['Dt.Pedido'].dt.year.mode().iloc[0])
                idx = pd.date_range(start=f"{ano_ref}-01-01", end=f"{ano_ref}-12-31", freq='ME')
                df_m = df_m.reindex(idx, fill_value=0)
            df_m['ticket'] = df_m.apply(lambda r: (r['vendas'] / r['pedidos']) if r['pedidos'] else 0, axis=1)
            meses = [d.strftime('%b') for d in df_m.index]

            fig, (ax_ped, ax_tic) = plt.subplots(2, 1, figsize=(12, 7), sharex=True)
            with sns.axes_style("whitegrid"):
                sns.barplot(x=meses, y=df_m['pedidos'].values, ax=ax_ped, color=CORES_GRAFICOS[0])
            ax_ped.set_title('Pedidos por Mês', fontsize=14)
            ax_ped.set_ylabel('Nº de pedidos')
            for container in ax_ped.containers:
                ax_ped.bar_label(container, fmt=lambda x: f"{int(x)}" if x > 0 else '', padding=2)

            ax_tic.plot(meses, df_m['ticket'].values, marker='o', linewidth=2, color=CORES_GRAFICOS[1])
            ax_tic.set_title('Ticket Médio por Mês', fontsize=14)
            ax_tic.set_ylabel('Ticket médio (R$)')
            ax_tic.yaxis.set_major_formatter(_get_mpl_FuncFormatter()(lambda val, pos: formatar_moeda_str(val)))
            ax_tic.grid(axis='y', alpha=0.25)

            fig.tight_layout()
            nome_grafico_pedidos_ticket = 'grafico_pedidos_ticket_mensal.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_pedidos_ticket), dpi=150)
            plt.close(fig)
            update_status("Gráfico 'Pedidos e Ticket por Mês' criado.")

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(55)

        nome_grafico_vendedores = ''
        if include_vendedores:
            _raise_if_cancelled(cancel_event)
            if 'Vendedor' in df.columns:
                vendas_vend = df.groupby('Vendedor')['Vlr.Total'].sum().sort_values(ascending=False)
                top_n = 15
                if len(vendas_vend) > top_n:
                    top = vendas_vend.head(top_n)
                    outros = vendas_vend.iloc[top_n:].sum()
                    vendas_plot = pd.concat([top, pd.Series({'Outros': outros})])
                else:
                    vendas_plot = vendas_vend

                fig, ax = plt.subplots(figsize=(12, 9))
                vendas_plot.sort_values().plot(kind='barh', ax=ax, color='#2980b9')
                ax.set_title('Vendas por Vendedor (Top)', fontsize=16)
                for patch in ax.patches:
                    width = patch.get_width()
                    y = patch.get_y()
                    height = patch.get_height()
                    ax.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
                ax.set_xlim(right=ax.get_xlim()[1] * 1.25)
                fig.tight_layout()
                nome_grafico_vendedores = 'grafico_vendas_por_vendedor.png'
                plt.savefig(_asset_path(assets_dir, nome_grafico_vendedores), dpi=150)
                plt.close(fig)
                update_status("Gráfico 'Vendas por Vendedor' criado.")
            else:
                update_status("AVISO: coluna 'Vendedor' não encontrada; pulando indicador de vendedores.")

        _raise_if_cancelled(cancel_event)

        nome_grafico_uf = ''
        if include_uf:
            _raise_if_cancelled(cancel_event)
            if 'UF' in df.columns:
                vendas_uf = df.groupby('UF')['Vlr.Total'].sum().sort_values(ascending=False)
                top_n = 15
                if len(vendas_uf) > top_n:
                    top = vendas_uf.head(top_n)
                    outros = vendas_uf.iloc[top_n:].sum()
                    vendas_plot = pd.concat([top, pd.Series({'Outros': outros})])
                else:
                    vendas_plot = vendas_uf

                fig, ax = plt.subplots(figsize=(12, 9))
                vendas_plot.sort_values().plot(kind='barh', ax=ax, color='#16a085')
                ax.set_title('Vendas por UF (Top)', fontsize=16)
                for patch in ax.patches:
                    width = patch.get_width()
                    y = patch.get_y()
                    height = patch.get_height()
                    ax.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
                ax.set_xlim(right=ax.get_xlim()[1] * 1.25)
                fig.tight_layout()
                nome_grafico_uf = 'grafico_vendas_por_uf.png'
                plt.savefig(_asset_path(assets_dir, nome_grafico_uf), dpi=150)
                plt.close(fig)
                update_status("Gráfico 'Vendas por UF' criado.")
            else:
                update_status("AVISO: coluna 'UF' não encontrada; pulando indicador por UF.")

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(75)

        nome_grafico_top_produtos = ''
        nome_grafico_top_produtos_qtd = ''
        nome_grafico_top_clientes = ''

        # Gráficos de Rankings Anuais
        top_n_rank = 20

        if include_rank_produtos or include_rank_clientes:
            _raise_if_cancelled(cancel_event)
            _lmap_anual = _produto_label_map(df)
            _top_p_val = df.groupby('Cod.Material')['Vlr.Total'].sum().nlargest(top_n_rank)
            _top_p_qtd = df.groupby('Cod.Material')['Qtde'].sum().nlargest(top_n_rank)
            top_produtos = _top_p_val.rename(index=lambda c: _lmap_anual.get(c, c))
            top_produtos_qtd = _top_p_qtd.rename(index=lambda c: _lmap_anual.get(c, c))
            top_clientes = df.groupby('Cliente')['Vlr.Total'].sum().nlargest(top_n_rank)

        if include_rank_produtos:
            _raise_if_cancelled(cancel_event)
            fig_prods, ax_prods = plt.subplots(figsize=(12, 12))
            top_produtos.sort_values().plot(kind='barh', ax=ax_prods)
            ax_prods.set_title('Top 20 Produtos do Ano (por Valor de Venda)', fontsize=16)
            ax_prods.tick_params(axis='y', labelsize=8)
            for patch in ax_prods.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_prods.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
            ax_prods.set_xlim(right=ax_prods.get_xlim()[1] * 1.25)
            fig_prods.tight_layout()
            nome_grafico_top_produtos = 'grafico_top_produtos.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_top_produtos), dpi=150)
            plt.close(fig_prods)
            update_status("Gráfico 'Top 20 Produtos por Valor' criado.")

            _raise_if_cancelled(cancel_event)

            fig_prods_qtd, ax_prods_qtd = plt.subplots(figsize=(12, 12))
            top_produtos_qtd.sort_values().plot(kind='barh', ax=ax_prods_qtd)
            ax_prods_qtd.set_title('Top 20 Produtos do Ano (por Quantidade Vendida)', fontsize=16)
            ax_prods_qtd.tick_params(axis='y', labelsize=8)
            for patch in ax_prods_qtd.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_prods_qtd.text(width * 1.01, y + height / 2, f'{int(width)} un', va='center')
            ax_prods_qtd.set_xlim(right=ax_prods_qtd.get_xlim()[1] * 1.25)
            fig_prods_qtd.tight_layout()
            nome_grafico_top_produtos_qtd = 'grafico_top_produtos_qtd.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_top_produtos_qtd), dpi=150)
            plt.close(fig_prods_qtd)
            update_status("Gráfico 'Top 20 Produtos por Qtd' criado.")

        _raise_if_cancelled(cancel_event)

        if include_rank_clientes:
            _raise_if_cancelled(cancel_event)
            fig_clientes, ax_clientes = plt.subplots(figsize=(12, 12))
            top_clientes.sort_values().plot(kind='barh', ax=ax_clientes)
            ax_clientes.set_title('Top 20 Clientes do Ano (por Valor de Compra)', fontsize=16)
            ax_clientes.tick_params(axis='y', labelsize=8)
            for patch in ax_clientes.patches:
                width = patch.get_width()
                y = patch.get_y()
                height = patch.get_height()
                ax_clientes.text(width * 1.01, y + height / 2, f'{formatar_moeda_str(width)}', va='center')
            ax_clientes.set_xlim(right=ax_clientes.get_xlim()[1] * 1.25)
            fig_clientes.tight_layout()
            nome_grafico_top_clientes = 'grafico_top_clientes.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_top_clientes), dpi=150)
            plt.close(fig_clientes)
            update_status("Gráfico 'Top 20 Clientes' criado.")

        _raise_if_cancelled(cancel_event)

        try:
            timings.add("Análises", time.perf_counter() - t_analises)
        except Exception:
            pass

        try:
            timings.add("Gráficos", time.perf_counter() - t_graficos)
        except Exception:
            pass

        update_status("Montando HTML do relatório anual...")
        t_html = time.perf_counter()
        caminho_logo_final = obter_caminho_logo_absoluto()

        _raise_if_cancelled(cancel_event)

        if progress_cb:
            progress_cb(85)

        template_path = resource_path('template_anual.html')
        context = {
            'DATA_GERACAO': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'CAMINHO_LOGO': caminho_logo_final,
            'VENDA_TOTAL_ANUAL': formatar_moeda_str(venda_total),
            'NUM_PEDIDOS_ANUAL': str(pedidos_unicos),
            'TICKET_MEDIO_ANUAL': formatar_moeda_str(ticket_medio),
            'UNIDADES_VENDIDAS_ANUAL': str(int(unidades_vendidas)),
            'GRAFICO_EVOLUCAO_MENSAL_PATH': nome_grafico_mensal,
            'GRAFICO_VENDAS_TRIMESTRAL_PATH': nome_grafico_trimestral,
            'GRAFICO_PEDIDOS_TICKET_MENSAL_PATH': nome_grafico_pedidos_ticket,
            'GRAFICO_TOP_PRODUTOS_PATH': nome_grafico_top_produtos,
            'GRAFICO_TOP_PRODUTOS_QTD_ANUAL_PATH': nome_grafico_top_produtos_qtd,
            'GRAFICO_TOP_CLIENTES_PATH': nome_grafico_top_clientes,
            'GRAFICO_VENDAS_POR_VENDEDOR_PATH': nome_grafico_vendedores,
            'GRAFICO_VENDAS_POR_UF_PATH': nome_grafico_uf,
            'MELHOR_MES_LABEL': melhor_mes_label,
            'MELHOR_MES_VALOR': melhor_mes_valor,
            'PIOR_MES_LABEL': pior_mes_label,
            'PIOR_MES_VALOR': pior_mes_valor,
            'INCLUDE_EVOLUCAO': bool(include_evolucao),
            'INCLUDE_PEDIDOS_TICKET': bool(include_pedidos_ticket),
            'INCLUDE_RANK_PRODUTOS': bool(include_rank_produtos),
            'INCLUDE_RANK_CLIENTES': bool(include_rank_clientes),
            'INCLUDE_VENDEDORES': bool(include_vendedores),
            'INCLUDE_UF': bool(include_uf),
        }

        html_final = _render_template(template_path, context)
        try:
            timings.add("HTML", time.perf_counter() - t_html)
        except Exception:
            pass

        pdf_probe, _ = _resolve_output_paths(save_path, output_mode)
        if pdf_probe and progress_cb:
            progress_cb(90)

        effective_path = _export_pdf_with_html_fallback(
            html_final=html_final,
            assets_dir=assets_dir,
            save_path=save_path,
            output_mode=output_mode,
            timings=timings,
            update_status=update_status,
            cancel_event=cancel_event,
        )
        pdf_out, html_side = _resolve_output_paths(save_path, output_mode)
        op_mode = (output_mode or 'pdf').strip().lower()
        if op_mode == 'html' and effective_path:
            update_status(f"Relatório Anual (HTML) '{os.path.basename(effective_path)}' gerado com sucesso!")
        elif effective_path and pdf_out and os.path.abspath(effective_path) == os.path.abspath(pdf_out):
            if html_side and os.path.isfile(html_side):
                update_status(
                    f"PDF e HTML gerados com sucesso: '{os.path.basename(pdf_out)}' e '{os.path.basename(html_side)}'."
                )
            else:
                update_status(f"Relatório Anual '{os.path.basename(pdf_out)}' gerado com sucesso!")
        elif effective_path:
            update_status(f"Análise Anual salva em HTML: '{os.path.basename(effective_path)}' (PDF indisponível nesta máquina).")

        if progress_cb:
            progress_cb(100)
        return effective_path

    except CancelledError:
        update_status("Processo cancelado pelo usuário.")
        raise

    except Exception as e:
        update_status(f"ERRO CRÍTICO na Análise Anual: {e}")
        pass  # mensagem à UI: thread principal / status
        raise e
    finally:
        try:
            timings.log_to(update_status)
        except Exception:
            pass
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)
        
# --- FUNÇÃO 4 (ORIGINAL E INTACTA) ---
def gerar_relatorio_comparativo_anual(
    caminho_ano_A,
    caminho_ano_B,
    log_queue,
    save_path,
    cancel_event=None,
    progress_cb=None,
    output_mode='pdf',
    include_churn=True,
    include_evolucao_mensal=True,
    include_tabela_yoy_mensal=True,
    include_top_produtos=True,
    include_top_clientes=True,
):
    def update_status(message):
        log_queue.put(message)
    assets_dir = None
    timings = TimingCollector()
    try:
        assets_dir = tempfile.mkdtemp(prefix="relatorio_assets_")
        update_status("--- INICIANDO RELATÓRIO COMPARATIVO ANUAL ---")
        if progress_cb:
            progress_cb(2)

        _raise_if_cancelled(cancel_event)
        
        update_status(f"Lendo dados do Ano A: {os.path.basename(caminho_ano_A)}")
        with timings.stage("Leitura Excel"):
            df_A_raw = pd.read_excel(caminho_ano_A)
            df_B_raw = pd.read_excel(caminho_ano_B)
        with timings.stage("Limpeza/Validação"):
            df_A = _clean_dataframe_base(df_A_raw, nome_relatorio="Comparativo Anual", log_queue=log_queue)
            update_status(f"Lendo dados do Ano B: {os.path.basename(caminho_ano_B)}")
            df_B = _clean_dataframe_base(df_B_raw, nome_relatorio="Comparativo Anual", log_queue=log_queue)

        _raise_if_cancelled(cancel_event)
        if progress_cb:
            progress_cb(12)
        required_extra = ['Cliente', 'Cod.Material', 'Descrição']
        _require_columns(df_A, required_extra, "Comparativo Anual")
        _require_columns(df_B, required_extra, "Comparativo Anual")

        ano_A = df_A['Dt.Pedido'].dt.year.iloc[0]
        ano_B = df_B['Dt.Pedido'].dt.year.iloc[0]
        update_status(f"Anos identificados: {ano_A} vs {ano_B}")

        _raise_if_cancelled(cancel_event)

        update_status("Calculando KPIs principais...")
        t_analises = time.perf_counter()
        venda_total_A = df_A['Vlr.Total'].sum()
        pedidos_A = df_A['Nro.Pedido'].nunique()
        ticket_medio_A = venda_total_A / pedidos_A if pedidos_A > 0 else 0
        
        venda_total_B = df_B['Vlr.Total'].sum()
        pedidos_B = df_B['Nro.Pedido'].nunique()
        ticket_medio_B = venda_total_B / pedidos_B if pedidos_B > 0 else 0
        
        def calcular_variacao_anual(val_A, val_B):
            if val_A > 0:
                var = ((val_B - val_A) / val_A) * 100
                cor = "positive" if var >= 0 else "negative"
                sinal = "+" if var >= 0 else ""
                return f"{sinal}{var:.2f}%", cor
            return "N/A", ""

        var_venda, cor_venda = calcular_variacao_anual(venda_total_A, venda_total_B)
        var_pedidos, cor_pedidos = calcular_variacao_anual(pedidos_A, pedidos_B)
        var_ticket, cor_ticket = calcular_variacao_anual(ticket_medio_A, ticket_medio_B)
        update_status("KPIs calculados. Gerando gráficos...")
        t_graficos = time.perf_counter()
        if progress_cb:
            progress_cb(20)

        _raise_if_cancelled(cancel_event)
        
        # Gráfico 1: Evolução Mensal Comparativa (opcional)
        nome_grafico_evolucao = ''
        nomes_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        vendas_mensais_A = None
        vendas_mensais_B = None

        def _calc_vendas_mensais_12m(_df):
            tmp = _df.copy()
            tmp['Mes'] = tmp['Dt.Pedido'].dt.month
            return tmp.groupby('Mes')['Vlr.Total'].sum().reindex(range(1, 13), fill_value=0)

        if include_evolucao_mensal or include_tabela_yoy_mensal:
            try:
                vendas_mensais_A = _calc_vendas_mensais_12m(df_A)
                vendas_mensais_B = _calc_vendas_mensais_12m(df_B)
            except Exception:
                vendas_mensais_A = None
                vendas_mensais_B = None

        if include_evolucao_mensal:
            _raise_if_cancelled(cancel_event)
            if vendas_mensais_A is None or vendas_mensais_B is None:
                vendas_mensais_A = _calc_vendas_mensais_12m(df_A)
                vendas_mensais_B = _calc_vendas_mensais_12m(df_B)

            fig, ax = plt.subplots(figsize=(12, 6))
            bar_width = 0.4
            index = np.arange(len(nomes_meses))
            ax.bar(index - bar_width/2, vendas_mensais_A, bar_width, label=f'{ano_A}', color=CORES_GRAFICOS[0])
            ax.bar(index + bar_width/2, vendas_mensais_B, bar_width, label=f'{ano_B}', color=CORES_GRAFICOS[1])
            ax.set_title(f'Comparativo de Vendas Mensais ({ano_A} vs {ano_B})', fontsize=16)
            ax.set_xticks(index)
            ax.set_xticklabels(nomes_meses)
            ax.yaxis.set_major_formatter(_get_mpl_FuncFormatter()(lambda val, pos: formatar_moeda_str(val)))
            ax.legend()
            fig.tight_layout()
            nome_grafico_evolucao = 'grafico_comp_anual_mensal.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_evolucao), dpi=150)
            plt.close(fig)
            update_status("Gráfico 'Comparativo Mensal' criado.")
            if progress_cb:
                progress_cb(45)

        _raise_if_cancelled(cancel_event)

        # Tabela YoY por mês (opcional)
        tabela_yoy_mensal_html = ''
        if include_tabela_yoy_mensal:
            try:
                if vendas_mensais_A is None or vendas_mensais_B is None:
                    vendas_mensais_A = _calc_vendas_mensais_12m(df_A)
                    vendas_mensais_B = _calc_vendas_mensais_12m(df_B)

                for i, mes_nome in enumerate(nomes_meses, start=1):
                    a = float(vendas_mensais_A.loc[i]) if vendas_mensais_A is not None else 0.0
                    b = float(vendas_mensais_B.loc[i]) if vendas_mensais_B is not None else 0.0
                    delta_abs = b - a
                    delta_pct = ((b / a) - 1.0) if a > 0 else None

                    cls = 'positive' if delta_abs >= 0 else 'negative'
                    sinal_abs = '+' if delta_abs >= 0 else ''
                    if delta_pct is None:
                        delta_pct_str = 'N/A'
                    else:
                        sinal_pct = '+' if delta_pct >= 0 else ''
                        delta_pct_str = f"{sinal_pct}{(delta_pct * 100):.2f}%"

                    tabela_yoy_mensal_html += (
                        "<tr>"
                        f"<td>{mes_nome}</td>"
                        f"<td>{formatar_moeda_str(a)}</td>"
                        f"<td>{formatar_moeda_str(b)}</td>"
                        f"<td class='{cls}'>{sinal_abs}{formatar_moeda_str(delta_abs)}</td>"
                        f"<td class='{cls}'>{delta_pct_str}</td>"
                        "</tr>"
                    )
            except Exception:
                tabela_yoy_mensal_html = "<tr><td colspan='5'>Não foi possível calcular a tabela YoY mensal.</td></tr>"

        # Gráfico 2: Comparativo Top 20 Produtos (opcional)
        nome_grafico_produtos = ''
        if include_top_produtos:
            _raise_if_cancelled(cancel_event)
            _lmap_ca = _produto_label_map(df_A, df_B)
            top_produtos_B = df_B.groupby('Cod.Material')['Vlr.Total'].sum().nlargest(20)
            top_produtos_A = df_A[df_A['Cod.Material'].isin(top_produtos_B.index)].groupby('Cod.Material')['Vlr.Total'].sum().reindex(top_produtos_B.index, fill_value=0)
            df_comp_produtos = pd.DataFrame({f'{ano_A}': top_produtos_A, f'{ano_B}': top_produtos_B}).sort_values(by=f'{ano_B}', ascending=True)
            df_comp_produtos.index = df_comp_produtos.index.map(lambda c: _lmap_ca.get(str(c), str(c)))

            fig, ax = plt.subplots(figsize=(12, 8))
            df_comp_produtos.plot(kind='barh', ax=ax, color=[CORES_GRAFICOS[0], CORES_GRAFICOS[1]])
            ax.set_title(f'Comparativo de Vendas: Top 20 Produtos de {ano_B}', fontsize=16)
            ax.xaxis.set_major_formatter(_get_mpl_FuncFormatter()(lambda val, pos: formatar_moeda_str(val)))
            fig.tight_layout()
            nome_grafico_produtos = 'grafico_comp_anual_produtos.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_produtos), dpi=150)
            plt.close(fig)
            update_status("Gráfico 'Comparativo Top Produtos' criado.")
            if progress_cb:
                progress_cb(60)

        _raise_if_cancelled(cancel_event)
        
        # Gráfico 3: Comparativo Top 20 Clientes (opcional)
        nome_grafico_clientes = ''
        if include_top_clientes:
            _raise_if_cancelled(cancel_event)
            top_clientes_B = df_B.groupby('Cliente')['Vlr.Total'].sum().nlargest(20)
            top_clientes_A = df_A[df_A['Cliente'].isin(top_clientes_B.index)].groupby('Cliente')['Vlr.Total'].sum().reindex(top_clientes_B.index, fill_value=0)
            df_comp_clientes = pd.DataFrame({f'{ano_A}': top_clientes_A, f'{ano_B}': top_clientes_B}).sort_values(by=f'{ano_B}', ascending=True)

            fig, ax = plt.subplots(figsize=(12, 8))
            df_comp_clientes.plot(kind='barh', ax=ax, color=[CORES_GRAFICOS[0], CORES_GRAFICOS[1]])
            ax.set_title(f'Comparativo de Compras: Top 20 Clientes de {ano_B}', fontsize=16)
            ax.xaxis.set_major_formatter(_get_mpl_FuncFormatter()(lambda val, pos: formatar_moeda_str(val)))
            fig.tight_layout()
            nome_grafico_clientes = 'grafico_comp_anual_clientes.png'
            plt.savefig(_asset_path(assets_dir, nome_grafico_clientes), dpi=150)
            plt.close(fig)
            update_status("Gráfico 'Comparativo Top Clientes' criado.")
            if progress_cb:
                progress_cb(75)

        _raise_if_cancelled(cancel_event)

        try:
            timings.add("Análises", time.perf_counter() - t_analises)
        except Exception:
            pass

        try:
            timings.add("Gráficos", time.perf_counter() - t_graficos)
        except Exception:
            pass

        update_status("Montando relatório PDF...")
        t_html = time.perf_counter()
        template_path = resource_path('template_comparativo_anual.html')

        _raise_if_cancelled(cancel_event)

        # --- Churn Anual (opcional) ---
        tabela_churn_html = ""
        resumo_churn = ""
        if include_churn:
            update_status("Iniciando análise de Clientes Inativos (Churn anual)...")
            clientes_A = set(df_A['Cliente'].unique())
            clientes_B = set(df_B['Cliente'].unique())
            clientes_inativos = clientes_A - clientes_B

            if not clientes_inativos:
                tabela_churn_html = "<tr><td colspan='2'>Nenhum cliente importante ficou inativo neste período.</td></tr>"
                resumo_churn = "<p>Nenhum cliente relevante ficou inativo neste período.</p>"
            else:
                df_inativos = df_A[df_A['Cliente'].isin(clientes_inativos)]
                vendas_inativos = df_inativos.groupby('Cliente')['Vlr.Total'].sum().sort_values(ascending=False)
                valor_total_inativo = vendas_inativos.sum()
                resumo_churn = f"<p>No total, <strong>{len(clientes_inativos)} clientes</strong> que representavam <strong>{formatar_moeda_str(valor_total_inativo)}</strong> em vendas em {ano_A} não compraram em {ano_B}.</p>"
                for cliente, valor in vendas_inativos.head(15).items():
                    tabela_churn_html += f"<tr><td>{cliente}</td><td style='text-align:right;'>{formatar_moeda_str(valor)}</td></tr>"
            update_status(f"Análise de Churn anual concluída: {len(clientes_inativos)} clientes inativos.")
        else:
            update_status("Churn anual desativado (configuração do relatório).")

        _raise_if_cancelled(cancel_event)

        # Se Jinja2 estiver disponível, renderiza os blocos condicionais corretamente.
        context_extra = {
            'INCLUDE_EVOLUCAO_MENSAL': bool(include_evolucao_mensal),
            'INCLUDE_TABELA_YOY_MENSAL': bool(include_tabela_yoy_mensal),
            'INCLUDE_TOP_PRODUTOS': bool(include_top_produtos),
            'INCLUDE_TOP_CLIENTES': bool(include_top_clientes),
            'INCLUDE_CHURN': bool(include_churn),
            'TABELA_YOY_MENSAL': Markup(tabela_yoy_mensal_html) if Markup else tabela_yoy_mensal_html,
            'RESUMO_CHURN': Markup(resumo_churn) if Markup else resumo_churn,
            'TABELA_CHURN': Markup(tabela_churn_html) if Markup else tabela_churn_html,
        }

        # Mantém compatibilidade com substituições existentes e adiciona as chaves novas.
        # (No modo Jinja2, isso renderiza os {% if %} do template.)
        update_status("Montando HTML do relatório...")
        html_final = _render_template(template_path, {
            'DATA_GERACAO': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'CAMINHO_LOGO': obter_caminho_logo_absoluto(),
            'ANO_A': str(ano_A),
            'ANO_B': str(ano_B),
            'VENDA_A': formatar_moeda_str(venda_total_A),
            'VENDA_B': formatar_moeda_str(venda_total_B),
            'VARIACAO_VENDA': var_venda,
            'COR_VARIACAO_VENDA': cor_venda,
            'PEDIDOS_A': str(int(pedidos_A)),
            'PEDIDOS_B': str(int(pedidos_B)),
            'VARIACAO_PEDIDOS': var_pedidos,
            'COR_VARIACAO_PEDIDOS': cor_pedidos,
            'TICKET_A': formatar_moeda_str(ticket_medio_A),
            'TICKET_B': formatar_moeda_str(ticket_medio_B),
            'VARIACAO_TICKET': var_ticket,
            'COR_VARIACAO_TICKET': cor_ticket,
            'GRAFICO_EVOLUCAO_MENSAL_COMP_PATH': nome_grafico_evolucao,
            'GRAFICO_TOP_PRODUTOS_COMP_PATH': nome_grafico_produtos,
            'GRAFICO_TOP_CLIENTES_COMP_PATH': nome_grafico_clientes,
            **context_extra,
        })
        try:
            timings.add("HTML", time.perf_counter() - t_html)
        except Exception:
            pass

        if progress_cb:
            progress_cb(90)

        _raise_if_cancelled(cancel_event)

        effective_path = _export_pdf_with_html_fallback(
            html_final=html_final,
            assets_dir=assets_dir,
            save_path=save_path,
            output_mode=output_mode,
            timings=timings,
            update_status=update_status,
            cancel_event=cancel_event,
        )
        pdf_out, html_side = _resolve_output_paths(save_path, output_mode)
        op_mode = (output_mode or 'pdf').strip().lower()
        if op_mode == 'html' and effective_path:
            update_status(f"Relatório Comparativo Anual (HTML) '{os.path.basename(effective_path)}' gerado com sucesso!")
        elif effective_path and pdf_out and os.path.abspath(effective_path) == os.path.abspath(pdf_out):
            if html_side and os.path.isfile(html_side):
                update_status(
                    f"PDF e HTML gerados com sucesso: '{os.path.basename(pdf_out)}' e '{os.path.basename(html_side)}'."
                )
            else:
                update_status(f"Relatório Comparativo Anual '{os.path.basename(pdf_out)}' gerado com sucesso!")
        elif effective_path:
            update_status(
                f"Comparativo Anual salvo em HTML: '{os.path.basename(effective_path)}' (PDF indisponível nesta máquina)."
            )

        if progress_cb:
            progress_cb(100)
        return effective_path

    except CancelledError:
        update_status("Processo cancelado pelo usuário.")
        raise

    except Exception as e:
        update_status(f"ERRO CRÍTICO no Comparativo Anual: {e}")
        pass  # mensagem à UI: thread principal / status
        raise e
    finally:
        try:
            timings.log_to(update_status)
        except Exception:
            pass
        if assets_dir:
            shutil.rmtree(assets_dir, ignore_errors=True)


def gerar_relatorio_qualidade_dados(
    caminho_dados,
    log_queue,
    save_path,
    cancel_event=None,
    progress_cb=None,
):
    def update_status(message):
        log_queue.put(message)

    def safe_head(df, n=50):
        try:
            return df.head(int(n))
        except Exception:
            return df

    def df_table_html(df, max_rows=200):
        try:
            if df is None:
                return ''
            df2 = df.copy()
            if max_rows is not None and len(df2) > int(max_rows):
                df2 = df2.head(int(max_rows))
            return df2.to_html(index=False, escape=True)
        except Exception:
            return ''

    update_status("--- INICIANDO RELATÓRIO DE QUALIDADE DE DADOS ---")
    _raise_if_cancelled(cancel_event)
    if progress_cb:
        progress_cb(2)

    update_status(f"Lendo ficheiro: {os.path.basename(caminho_dados)}")
    df_raw = pd.read_excel(caminho_dados)
    _raise_if_cancelled(cancel_event)
    if progress_cb:
        progress_cb(15)

    before_rows = int(len(df_raw)) if hasattr(df_raw, '__len__') else None
    df = _clean_dataframe_base(df_raw, nome_relatorio="Qualidade de Dados", log_queue=log_queue)
    after_rows = int(len(df)) if hasattr(df, '__len__') else None
    _raise_if_cancelled(cancel_event)
    if progress_cb:
        progress_cb(25)

    columns_present = list(df.columns)
    expected_optional = ['Vendedor', 'UF', 'Cliente', 'Descrição', 'Qtde']
    missing_optional = [c for c in expected_optional if c not in columns_present]

    # Resumo básico
    date_min = None
    date_max = None
    try:
        date_min = df['Dt.Pedido'].min()
        date_max = df['Dt.Pedido'].max()
    except Exception:
        pass

    # Métricas e problemas
    problems = []
    tables = []

    def add_problem(title, count=None, details=None, df_sample=None, severity='Média'):
        problems.append({
            'title': title,
            'count': count,
            'details': details,
            'table': df_sample,
            'severity': severity,
        })

    # Valores inválidos
    _raise_if_cancelled(cancel_event)
    try:
        zeros_total = int((df['Vlr.Total'] == 0).sum())
        neg_total = int((df['Vlr.Total'] < 0).sum())
        if zeros_total:
            add_problem("Linhas com Vlr.Total = 0", zeros_total, severity='Média')
        if neg_total:
            add_problem(
                "Linhas com Vlr.Total negativo",
                neg_total,
                df_sample=safe_head(df[df['Vlr.Total'] < 0][['Nro.Pedido', 'Dt.Pedido', 'Vlr.Total']]),
                severity='Alta',
            )
    except Exception:
        pass

    # Qtde
    _raise_if_cancelled(cancel_event)
    if 'Qtde' in df.columns:
        try:
            qtde_zero = int((df['Qtde'] == 0).sum())
            qtde_neg = int((df['Qtde'] < 0).sum())
            if qtde_zero:
                add_problem("Linhas com Qtde = 0", qtde_zero, severity='Média')
            if qtde_neg:
                add_problem(
                    "Linhas com Qtde negativa",
                    qtde_neg,
                    df_sample=safe_head(df[df['Qtde'] < 0][['Nro.Pedido', 'Descrição', 'Qtde']]),
                    severity='Alta',
                )
        except Exception:
            pass

    # Vendedor não encontrado / vazio
    _raise_if_cancelled(cancel_event)
    if 'Vendedor' in df.columns:
        try:
            s = df['Vendedor'].astype(str).str.strip()
            mask_nf = (s == '') | (s.str.lower() == 'nan') | (s.str.lower() == 'vendedor não encontrado') | (s.str.lower() == 'vendedor nao encontrado')
            n_rows = int(mask_nf.sum())
            if n_rows:
                n_orders = None
                try:
                    n_orders = int(df.loc[mask_nf, 'Nro.Pedido'].nunique())
                except Exception:
                    n_orders = None
                det = f"{n_rows} linhas" + (f" | {n_orders} pedidos" if n_orders is not None else '')
                add_problem(
                    "Vendedor ausente / não encontrado",
                    n_rows,
                    details=det,
                    df_sample=safe_head(df.loc[mask_nf, ['Nro.Pedido', 'Dt.Pedido', 'Vendedor']].drop_duplicates()),
                    severity='Alta',
                )
        except Exception:
            pass

    # UF inválida
    _raise_if_cancelled(cancel_event)
    if 'UF' in df.columns:
        try:
            uf = df['UF'].astype(str).str.strip().str.upper()
            invalid = (uf == '') | (uf == 'N/D') | (uf == 'ND') | (uf == 'N\\D') | (~uf.isin(UF_VALIDAS))
            n_rows = int(invalid.sum())
            if n_rows:
                n_orders = None
                try:
                    n_orders = int(df.loc[invalid, 'Nro.Pedido'].nunique())
                except Exception:
                    n_orders = None
                det = f"{n_rows} linhas" + (f" | {n_orders} pedidos" if n_orders is not None else '')
                add_problem(
                    "UF inválida / ausente",
                    n_rows,
                    details=det,
                    df_sample=safe_head(df.loc[invalid, ['Nro.Pedido', 'UF']].drop_duplicates()),
                    severity='Alta',
                )
        except Exception:
            pass

    # Cliente / Descrição vazios
    _raise_if_cancelled(cancel_event)
    for col in ['Cliente', 'Descrição']:
        if col in df.columns:
            try:
                s = df[col].astype(str).str.strip()
                mask = (s == '') | (s.str.lower() == 'nan')
                n_rows = int(mask.sum())
                if n_rows:
                    add_problem(
                        f"{col} ausente",
                        n_rows,
                        df_sample=safe_head(df.loc[mask, ['Nro.Pedido', col]].drop_duplicates()),
                        severity='Baixa',
                    )
            except Exception:
                pass

    # Consistência por pedido
    _raise_if_cancelled(cancel_event)
    key = 'Nro.Pedido'
    if key in df.columns:
        meta_cols = [c for c in ['Vendedor', 'UF', 'Cliente'] if c in df.columns]
        if meta_cols:
            try:
                agg = {}
                for c in meta_cols:
                    agg[c] = pd.Series.nunique
                g = df.groupby(key)[meta_cols].agg(agg)
                bad_mask = None
                for c in meta_cols:
                    m = g[c] > 1
                    bad_mask = m if bad_mask is None else (bad_mask | m)
                if bad_mask is not None and bool(bad_mask.any()):
                    bad = g[bad_mask].reset_index().rename(columns={c: f"{c}_valores" for c in meta_cols})
                    add_problem(
                        "Inconsistência por pedido (mesmo Nro.Pedido com múltiplos valores)",
                        int(len(bad)),
                        details="Geralmente indica erro de junção/duplicidade de cadastro.",
                        df_sample=safe_head(bad, 80),
                        severity='Alta',
                    )
            except Exception:
                pass

    if progress_cb:
        progress_cb(75)

    # Monta HTML
    _raise_if_cancelled(cancel_event)
    update_status("Montando relatório de qualidade (HTML)...")
    generated_at = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    periodo_str = "-"
    try:
        if date_min is not None and date_max is not None and pd.notna(date_min) and pd.notna(date_max):
            periodo_str = f"{pd.to_datetime(date_min).strftime('%d/%m/%Y')} → {pd.to_datetime(date_max).strftime('%d/%m/%Y')}"
    except Exception:
        pass

    missing_optional_html = ''
    if missing_optional:
        missing_optional_html = "<p><strong>Observação:</strong> colunas opcionais ausentes (algumas checagens foram puladas): " + ", ".join(missing_optional) + "</p>"

    issues_html_parts = []
    summary_rows = []
    severity_rank = {'Alta': 0, 'Média': 1, 'Media': 1, 'Baixa': 2}

    def _sev_badge_html(sev: str) -> str:
        sev2 = (sev or 'Média').strip()
        sev_l = sev2.lower()
        if sev_l.startswith('a'):
            cls = 'sev sev-high'
        elif sev_l.startswith('b'):
            cls = 'sev sev-low'
        else:
            cls = 'sev sev-med'
        return f"<span class='{cls}'>{sev2}</span>"

    if not problems:
        issues_html_parts.append("<div class='ok'>Nenhum problema relevante encontrado nas checagens padrão.</div>")
    else:
        try:
            for p in problems:
                summary_rows.append({
                    'Problema': p.get('title') or '-',
                    'Severidade': p.get('severity') or 'Média',
                    'Ocorrências': int(p.get('count')) if p.get('count') is not None else None,
                })
        except Exception:
            summary_rows = []

        def _sort_key(r):
            sev = (r.get('Severidade') or 'Média').strip()
            sev_rank = severity_rank.get(sev, severity_rank.get(sev.title(), 1))
            occ = r.get('Ocorrências')
            occ2 = int(occ) if occ is not None else -1
            return (sev_rank, -occ2, r.get('Problema') or '')

        try:
            summary_rows = sorted(summary_rows, key=_sort_key)
        except Exception:
            pass

        for p in problems:
            count = p.get('count')
            details = p.get('details')
            sev = p.get('severity') or 'Média'
            issues_html_parts.append("<div class='card'>")
            issues_html_parts.append(f"<h3>{p.get('title')} {_sev_badge_html(sev)}</h3>")
            if count is not None:
                issues_html_parts.append(f"<p><strong>Ocorrências:</strong> {count}</p>")
            if details:
                issues_html_parts.append(f"<p class='muted'>{details}</p>")
            if p.get('table') is not None:
                issues_html_parts.append("<div class='table'>" + df_table_html(p.get('table')) + "</div>")
            issues_html_parts.append("</div>")

    summary_html = ""
    if not summary_rows:
        summary_html = "<div class='muted'>Sem problemas a resumir.</div>"
    else:
        try:
            rows_html = []
            for r in summary_rows:
                occ = r.get('Ocorrências')
                occ_str = str(occ) if occ is not None else "-"
                rows_html.append(
                    "<tr>"
                    f"<td>{str(r.get('Problema') or '-')}</td>"
                    f"<td>{_sev_badge_html(r.get('Severidade') or 'Média')}</td>"
                    f"<td style='text-align:right'>{occ_str}</td>"
                    "</tr>"
                )
            summary_html = (
                "<div class='card'>"
                "<div class='muted' style='margin-bottom:8px'>Ordenado por severidade e volume.</div>"
                "<div class='table'>"
                "<table>"
                "<thead><tr><th>Problema</th><th>Severidade</th><th style='text-align:right'>Ocorrências</th></tr></thead>"
                f"<tbody>{''.join(rows_html)}</tbody>"
                "</table>"
                "</div>"
                "</div>"
            )
        except Exception:
            summary_html = ""

    html_final = f"""<!doctype html>
<html lang='pt-br'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>Relatório de Qualidade de Dados</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }}
    h1 {{ margin: 0 0 6px 0; }}
    .muted {{ color: #6b7280; }}
    .grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; margin: 12px 0 18px; }}
    .kpi {{ border: 1px solid #e5e7eb; border-radius: 10px; padding: 12px; background: #fff; }}
    .kpi .label {{ font-size: 12px; color: #6b7280; }}
    .kpi .value {{ font-size: 16px; font-weight: 600; margin-top: 4px; }}
    .card {{ border: 1px solid #e5e7eb; border-radius: 10px; padding: 12px; margin: 10px 0; background: #fff; }}
    .ok {{ border: 1px solid #bbf7d0; background: #f0fdf4; padding: 12px; border-radius: 10px; }}
        .sev {{ display: inline-block; margin-left: 8px; padding: 2px 8px; border-radius: 999px; font-size: 12px; border: 1px solid #e5e7eb; }}
        .sev-high {{ background: #fef2f2; border-color: #fecaca; color: #991b1b; }}
        .sev-med {{ background: #fffbeb; border-color: #fed7aa; color: #92400e; }}
        .sev-low {{ background: #eff6ff; border-color: #bfdbfe; color: #1e40af; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 12px; }}
    th, td {{ border: 1px solid #e5e7eb; padding: 6px 8px; text-align: left; }}
    th {{ background: #f9fafb; }}
    .table {{ overflow-x: auto; }}
  </style>
</head>
<body>
  <h1>Relatório de Qualidade de Dados</h1>
  <div class='muted'>Gerado em: {generated_at}</div>
  <div class='muted'>Arquivo: {caminho_dados}</div>

  <div class='grid'>
    <div class='kpi'><div class='label'>Período detectado</div><div class='value'>{periodo_str}</div></div>
    <div class='kpi'><div class='label'>Linhas (antes → depois limpeza)</div><div class='value'>{before_rows} → {after_rows}</div></div>
    <div class='kpi'><div class='label'>Pedidos únicos</div><div class='value'>{int(df['Nro.Pedido'].nunique()) if 'Nro.Pedido' in df.columns else '-'}</div></div>
    <div class='kpi'><div class='label'>Colunas</div><div class='value'>{len(columns_present)}</div></div>
  </div>

  {missing_optional_html}

    <h2>Resumo</h2>
    {summary_html}

  <h2>Problemas encontrados</h2>
  {''.join(issues_html_parts)}

  <h2>Notas</h2>
  <ul>
    <li>Este relatório é uma checagem automática (não substitui validação manual).</li>
    <li>Pedidos podem ter várias linhas (itens). A checagem de inconsistência procura pedidos com metadados conflitantes.</li>
  </ul>
</body>
</html>"""

    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(html_final)

    if progress_cb:
        progress_cb(100)
    update_status(f"Relatório de Qualidade salvo: {save_path}")

