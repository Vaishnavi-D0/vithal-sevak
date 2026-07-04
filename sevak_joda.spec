# PyInstaller spec for Sevak Joda - Member Registration App
# Build on a Windows machine with:
#   pyinstaller sevak_joda.spec
#
# Produces a single-file, windowed dist\SevakJoda.exe with all Python
# dependencies packed in. credentials.json is bundled alongside it in
# dist\ (not embedded inside the exe) so it can be swapped without rebuilding.

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

hidden_imports = (
    collect_submodules("win32com")
    + collect_submodules("googleapiclient")
    + collect_submodules("google_auth_httplib2")
    + [
        "win32timezone",
        "gspread",
        "deep_translator",
        "bs4",
        "PIL",
        "PIL._tkinter_finder",
    ]
)

# googleapiclient ships static discovery documents as package data;
# without these, Drive/Sheets API calls fail at runtime in the frozen exe.
extra_datas = collect_data_files("googleapiclient")

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("marathi_keyboard.py", "."),
        ("scanner_capture.py", "."),
        ("translator.py", "."),
        ("drive_helper.py", "."),
        ("marathi_text_render.py", "."),
        ("fonts/NotoSansDevanagari-Regular.ttf", "fonts"),
    ] + extra_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="SevakJoda",
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
