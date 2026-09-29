# -*- coding: utf-8 -*-
"""PII_Anonymizer_Tool.exe を onefile でビルドし、APIキーと辞書が混入していないことを検査する。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PyInstaller.archive.readers import CArchiveReader

ROOT = Path(__file__).resolve().parent
SPEC = ROOT / "PII_Anonymizer_Tool.spec"
DIST_EXE = ROOT / "dist" / "PII_Anonymizer_Tool.exe"
API_KEY_FILE = ROOT / "gemini_api_key.txt"

PROMPT_MARKERS = [
    "あなたは会計事務所に勤務するスタッフとして、クライアントへの短い確認・お礼メールの",
    "Google検索によるグラウンディングを使い、国税庁の公表資料・TKC・税理士が",
]

FORBIDDEN_ARCHIVE_NAMES = (
    "gemini_api_key.txt",
    "pii_dictionary.csv",
    "main_all_in.py",
    "署名.txt",
)
# reply_prompt_*.txt は同梱を許可している(プロンプトは公開してよい内容のため)。
# そのため FORBIDDEN_ARCHIVE_NAMES には含めない。


def _blob_from_extract(extracted) -> bytes:
    if isinstance(extracted, (bytes, bytearray)):
        return bytes(extracted)
    if isinstance(extracted, tuple):
        for item in extracted:
            if isinstance(item, (bytes, bytearray)):
                return bytes(item)
    raise TypeError(f"unexpected extract payload: {type(extracted)!r}")


def _api_key_needles() -> list[bytes]:
    needles: list[bytes] = []
    if not API_KEY_FILE.is_file():
        return needles
    for line in API_KEY_FILE.read_text(encoding="utf-8").splitlines():
        key = line.strip()
        if not key or key.startswith("#") or key == "YOUR_API_KEY_HERE":
            continue
        if len(key) < 12:
            continue
        needles.append(key.encode("utf-8"))
    return needles


def _assert_archive_safe(exe_path: Path) -> None:
    reader = CArchiveReader(str(exe_path))
    names = [str(n) for n in reader.toc]
    lowered = [n.replace("\\", "/").lower() for n in names]
    for forbidden in FORBIDDEN_ARCHIVE_NAMES:
        if any(forbidden.lower() in n for n in lowered):
            raise SystemExit(f"onefileアーカイブに同梱禁止ファイルがあります: {forbidden}")

    script = _blob_from_extract(reader.extract("main_all_in_without_api"))
    missing = [m[:40] for m in PROMPT_MARKERS if m.encode("utf-8") not in script]
    if missing:
        raise SystemExit("エントリスクリプトにプロンプトが埋め込まれていません: " + "; ".join(missing))

    if b"YOUR_API_KEY_HERE" not in script:
        raise SystemExit("埋め込みAPIキーがダミー値ではありません。")

    hits = [i for i, needle in enumerate(_api_key_needles(), start=1) if needle in script]
    if hits:
        raise SystemExit("エントリスクリプトに実APIキーが含まれています。")


def main() -> None:
    skip_build = "--skip-build" in sys.argv
    if not skip_build:
        pyinstaller = Path(sys.executable).with_name("pyinstaller.exe")
        cmd = (
            [str(pyinstaller)]
            if pyinstaller.is_file()
            else [sys.executable, "-m", "PyInstaller"]
        ) + ["--noconfirm", "--clean", str(SPEC)]
        print(" ".join(cmd), flush=True)
        subprocess.check_call(cmd, cwd=ROOT)

    if not DIST_EXE.is_file():
        raise SystemExit(f"exeが見つかりません: {DIST_EXE}")

    _assert_archive_safe(DIST_EXE)
    size_mb = DIST_EXE.stat().st_size / (1024 * 1024)
    print(f"OK: {DIST_EXE} ({size_mb:.1f} MB)  プロンプト同梱 / API・辞書なし", flush=True)


if __name__ == "__main__":
    main()