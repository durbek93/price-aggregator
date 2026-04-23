# Сборка одного exe (без консоли):
#   pip install pyinstaller
#   pyinstaller price_aggregator.spec
# Готовый файл: dist\PriceAggregator.exe — его и запускают пользователи (двойной щелчок).
#
# Иконка exe (Проводник, ярлыки): положите assets\app_icon.ico — параметр icon ниже подхватит файл.
# Иконка окна/панели задач в рантайме: app_icon.ico или app_icon.png в папке assets (копируется в сборку).

from pathlib import Path

SPEC_DIR = Path(SPECPATH).resolve()
ASSETS_DIR = SPEC_DIR / "assets"
datas = [(str(ASSETS_DIR), "assets")] if ASSETS_DIR.is_dir() else []
ICO_FILE = ASSETS_DIR / "app_icon.ico"

block_cipher = None

a = Analysis(
    ["app_main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "yaml",
        "openpyxl",
        "xlrd",
        "rapidfuzz",
        "pandas",
        "svodny_price_vertical",
        "assortment_prices",
        "cheapest_svodny_prices",
        "pipeline",
        "app_settings",
        "best_suppliers_analysis",
        "version",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

_exe_kw = dict(
    name="PriceAggregator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if ICO_FILE.is_file():
    _exe_kw["icon"] = str(ICO_FILE)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    **_exe_kw,
)
