# -*- mode: python ; coding: utf-8 -*-
# main_all_in_without_api.py を onefile exe にビルドするための spec。
#
# ・エントリは main_all_in_without_api.py(APIキー=ダミー / 辞書=ヘッダーのみ /
#   署名=ダミー、プロンプトは _EMBEDDED_PROMPT_TEMPLATES として埋め込み済み)。
# ・datas に追加するのは、各ライブラリ自身のデータと reply_prompt_*.txt のみ。
#   gemini_api_key.txt / pii_dictionary.csv / 署名.txt は絶対に追加しない。
# ・オプションは build_exe.ps1 と同じ(--windowed / --collect-all ... / --hidden-import ...)。
#
# 実行: python build_onefile.py  (ビルド後に混入チェックも行う)
#   出力先: dist/PII_Anonymizer_Tool.exe

import os

from PyInstaller.utils.hooks import collect_all

ENTRY_SCRIPT = "main_all_in_without_api.py"
EXE_NAME = "PII_Anonymizer_Tool"

COLLECT_ALL_PACKAGES = [
    "customtkinter",
    "tksheet",
    "spacy",
    "ginza",
    "ja_ginza",
    "sudachipy",
    "sudachidict_core",
    "llama_cpp",
    "google.genai",
]
HIDDEN_IMPORTS = [
    "win32timezone",
    "win32com.gen_py",
]

# 同梱を許可する外部ファイル(存在するものだけ追加する)。
PROMPT_FILES = [
    "reply_prompt_short.txt",
    "reply_prompt_template.txt",
    "reply_prompt_long_search.txt",
]

datas = [
    (os.path.join(SPECPATH, f), ".")
    for f in PROMPT_FILES
    if os.path.isfile(os.path.join(SPECPATH, f))
]
binaries = []
hiddenimports = list(HIDDEN_IMPORTS)
for pkg in COLLECT_ALL_PACKAGES:
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,  # --windowed
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
