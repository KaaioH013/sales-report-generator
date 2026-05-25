# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gerar_relatorio.py'],
    pathex=[],
    binaries=[],
    datas=[('template.html', '.'), ('template_comparativo.html', '.'), ('template_anual.html', '.'), ('template_comparativo_anual.html', '.'), ('Logo.png', '.'), ('mapa_brasil_dados.pkl', '.')],
    hiddenimports=[
        'relatorio_servicos',
        'pandas', 'pandas.io.formats.style',
        'numpy',
        'matplotlib', 'matplotlib.pyplot', 'matplotlib.ticker',
        'matplotlib.backends.backend_agg', 'matplotlib.backends.backend_svg',
        'seaborn', 'seaborn.categorical', 'seaborn.distributions',
        'jinja2', 'markupsafe',
        'weasyprint', 'weasyprint.text.fonts',
        'geopandas', 'geopandas.io',
        'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        'pyogrio', 'pyproj', 'shapely', 'shapely.geometry',
        'cssselect2', 'html5lib', 'tinycss2',
        'customtkinter', 'customtkinter.windows', 'customtkinter.windows.widgets',
        'customtkinter.windows.widgets.theme',
        'darkdetect',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    [],
    exclude_binaries=True,
    name='gerar_relatorio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    splash.binaries,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='gerar_relatorio',
)
