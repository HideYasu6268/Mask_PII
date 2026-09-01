# -*- coding: utf-8 -*-
"""
pii_core.py

GUIに依存しないPII検出・適用ロジック。
  - detect_pii(text)      : 正規表現+NERでPII候補を検出し、一覧を返す
  - apply_mapping(text, mapping) : 対応表(元の値 -> 匿名化後の値)を原文に適用する

すべてローカル処理のみで完結し、外部通信は行わない。
"""

from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass
from pathlib import Path


REGEX_PATTERNS: dict[str, str] = {
    "EMAIL": r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
    "URL": r"https?://[^\s]+",
    "PHONE": r"(?<!\d)0\d{1,4}-\d{1,4}-\d{4}(?!\d)|(?<!\d)0\d{9,10}(?!\d)",
    "POSTAL_CODE": r"(?<!\d)〒?\d{3}-\d{4}(?!-?\d)",
    "CREDIT_CARD": r"(?<!\d)(?:\d{4}[- ]?){3}\d{4}(?!\d)",
    "MY_NUMBER": r"(?<!\d)\d{4}\s?\d{4}\s?\d{4}(?!\d)",
    "IP_ADDRESS": r"\b(?:\d{1,3}\.){3}\d{1,3}\b",
}
REGEX_ORDER = ["EMAIL", "URL", "PHONE", "POSTAL_CODE", "CREDIT_CARD",
               "MY_NUMBER", "IP_ADDRESS"]

NER_LABEL_MAP = {
    "PERSON": "PERSON",
    "ORG": "ORG",
    "GPE": "LOCATION",
    "LOC": "LOCATION",
}


@dataclass
class PiiItem:
    type: str
    value: str
    start: int
    end: int


_nlp_cache = {}


def _get_ner_model(model_name: str = "ja_ginza"):
    """spaCy/GiNZAモデルを遅延ロードしキャッシュする。失敗時はNoneを返す(NERスキップ)。"""
    if model_name in _nlp_cache:
        return _nlp_cache[model_name]
    try:
        import spacy
    except ImportError:
        print("[info] spaCy未インストールのためNERをスキップします。", file=sys.stderr)
        _nlp_cache[model_name] = None
        return None

    # ja_ginza特有の既知の不具合対策: 新しめのspaCy(3.8.12以降)+ja_ginzaの組み合わせだと
    # compound_splitterコンポーネントのsplit_modeがNoneのままロードされConfigValidationError
    # になることがある。split_modeを明示的に渡すことで回避する。
    load_kwargs = {}
    if "ginza" in model_name:
        load_kwargs["config"] = {"components": {"compound_splitter": {"split_mode": "A"}}}

    nlp = None
    try:
        nlp = spacy.load(model_name, **load_kwargs)
    except OSError:
        try:
            nlp = spacy.load("ja_core_news_lg")
        except OSError:
            print(f"[info] 日本語モデル '{model_name}' が見つからないためNERをスキップします。",
                  file=sys.stderr)
            nlp = None
    except Exception as e:  # noqa: BLE001
        # ConfigValidationErrorなど、OSError以外の予期しないロードエラー。
        # NERはあくまで補助機能なので、ここで落ちてもアプリ全体は止めず、
        # NERだけスキップして正規表現ベースの検出は継続させる。
        print(f"[warn] '{model_name}' のロードに失敗したためNERをスキップします: {e}",
              file=sys.stderr)
        nlp = None
    _nlp_cache[model_name] = nlp
    return nlp


def detect_pii(text: str, use_ner: bool = True) -> list[PiiItem]:
    """テキストからPII候補を検出し、開始位置順のリストで返す(重複区間は除去)。"""
    spans: list[PiiItem] = []
    occupied: list[tuple[int, int]] = []

    def overlaps(s: int, e: int) -> bool:
        return any(not (e <= os or s >= oe) for os, oe in occupied)

    # 1. 正規表現(定型パターンを優先確定)
    for label in REGEX_ORDER:
        for m in re.finditer(REGEX_PATTERNS[label], text):
            s, e = m.start(), m.end()
            if overlaps(s, e):
                continue
            spans.append(PiiItem(type=label, value=m.group(0), start=s, end=e))
            occupied.append((s, e))

    # 2. NER(人名・組織名・地名)。正規表現で確定済みの区間とは重複させない。
    if use_ner:
        nlp = _get_ner_model()
        if nlp is not None:
            doc = nlp(text)
            for ent in doc.ents:
                label = NER_LABEL_MAP.get(ent.label_)
                if label is None:
                    continue
                s, e = ent.start_char, ent.end_char
                if overlaps(s, e):
                    continue
                spans.append(PiiItem(type=label, value=ent.text, start=s, end=e))
                occupied.append((s, e))

    spans.sort(key=lambda item: item.start)
    return spans


def dedupe_values(items: list[PiiItem]) -> list[PiiItem]:
    """同一の値(例: 同じ人名が複数回出現)は対応表上では1行にまとめる。"""
    seen: dict[str, PiiItem] = {}
    for item in items:
        if item.value not in seen:
            seen[item.value] = item
    return list(seen.values())


def merge_llm_items(items: list[PiiItem], llm_found: list[dict], text: str) -> list[PiiItem]:
    """LLMが追加で見つけたPII候補を、既存の検出結果にマージする。

    llm_found: [{"type": "PERSON", "value": "山田太郎"}, ...]
    すでに同じ値が検出済みの場合はスキップ(重複防止)。
    位置(start/end)は原文中で最初に出現する箇所を採用する。
    """
    existing_values = {it.value for it in items}
    merged = list(items)
    for entry in llm_found:
        value = entry.get("value", "")
        type_ = entry.get("type", "OTHER")
        if not value or value in existing_values:
            continue
        idx = text.find(value)
        if idx == -1:
            continue  # 原文に存在しない値は採用しない(LLMの創作を防ぐ)
        merged.append(PiiItem(type=type_, value=value, start=idx, end=idx + len(value)))
        existing_values.add(value)
    merged.sort(key=lambda it: it.start)
    return merged


def build_placeholder_mapping(items: list[PiiItem]) -> dict[str, str]:
    """種別ごとの連番タグ(例: [PERSON_1], [EMAIL_2])を機械的に割り当てる。

    同じ値には同じタグを割り当てて一貫性を保つ。出現順に番号を振る。
    """
    counters: dict[str, int] = {}
    mapping: dict[str, str] = {}
    for item in sorted(items, key=lambda it: it.start):
        if item.value in mapping:
            continue
        counters[item.type] = counters.get(item.type, 0) + 1
        mapping[item.value] = f"[{item.type}_{counters[item.type]}]"
    return mapping


def reverse_mapping(mapping: dict[str, str]) -> dict[str, str]:
    """対応表(元の値 -> 匿名化後の値)を反転し、匿名化解除用の対応表を作る。

    匿名化後の値(タグ)が空欄の行は対象外。
    """
    return {replacement: original for original, replacement in mapping.items() if replacement}


DICTIONARY_HEADERS = ["元の値", "匿名化後", "種別"]
# 非エンジニアがExcelでそのまま開く前提のため、日本語Windows既定のShift-JIS
# (cp932)で読み書きする。BOMは存在しない文字コードなので、追記のたびに
# 開き直してもutf-8-sigのような重複BOM問題は起きない。
DICTIONARY_ENCODING = "cp932"


def load_master_dictionary(path: str | Path) -> dict[str, dict[str, str]]:
    """マスター辞書CSV(蓄積型)を読み込み、「元の値」をキーにした辞書で返す。

    ファイルが存在しない場合は空辞書を返す。
    """
    path = Path(path)
    if not path.exists():
        return {}
    entries: dict[str, dict[str, str]] = {}
    with path.open("r", encoding=DICTIONARY_ENCODING, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            original = (row.get("元の値") or "").strip()
            if not original:
                continue
            entries[original] = {
                "replacement": (row.get("匿名化後") or "").strip(),
                "type": (row.get("種別") or "").strip(),
            }
    return entries


def append_to_master_dictionary(rows: list[dict], path: str | Path) -> tuple[int, list[str]]:
    """現在の対応表(rows)のうち、マスター辞書CSVにまだ無い「元の値」だけを追記する。

    既存の値は上書きしない(タグの連番はセッションごとに振り直されるため、
    マスター辞書側は最初に登録された内容を優先し、値そのものの蓄積に徹する)。
    Shift-JISに存在しない文字(一部の絵文字など)を含む値は1行ずつ捕捉して
    スキップし、他の行の追記は継続する(全体を巻き込んで失敗させない)。

    戻り値: (新規に追記した件数, 文字コードの都合で書き込めずスキップした「元の値」のリスト)
    """
    path = Path(path)
    existing = load_master_dictionary(path)
    new_rows = [r for r in rows if r.get("original") and r["original"] not in existing]
    if not new_rows:
        return 0, []

    file_exists = path.exists() and path.stat().st_size > 0
    mode = "a" if file_exists else "w"
    added = 0
    skipped: list[str] = []
    with path.open(mode, encoding=DICTIONARY_ENCODING, newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(DICTIONARY_HEADERS)
        for r in new_rows:
            try:
                writer.writerow([r["original"], r.get("replacement", ""), r.get("type", "")])
            except UnicodeEncodeError:
                skipped.append(r["original"])
                continue
            added += 1
    return added, skipped


def detect_from_dictionary(text: str, dictionary: dict[str, dict[str, str]]) -> list[PiiItem]:
    """マスター辞書(元の値 -> {type, replacement})のうち、原文に実際に出現する
    値だけを検出結果として返す。値が長い順に走査し、短い値が長い値の一部として
    誤って検出されるのを避ける(apply_mappingと同じ考え方)。

    正規表現/NER検出と同様、毎回の自動検出のたびにこの関数を呼び、過去に
    人手で確認済みの値を優先的に対応表へ差し戻すために使う。
    """
    items: list[PiiItem] = []
    for original in sorted(dictionary.keys(), key=len, reverse=True):
        if not original:
            continue
        idx = text.find(original)
        if idx == -1:
            continue
        type_ = dictionary[original].get("type") or "OTHER"
        items.append(PiiItem(type=type_, value=original, start=idx, end=idx + len(original)))
    items.sort(key=lambda it: it.start)
    return items


def apply_mapping(text: str, mapping: dict[str, str]) -> str:
    """対応表(元の値 -> 匿名化後の値)を原文に適用する。

    値の長い順に置換することで、短い値が長い値の一部として誤って
    先に置換されてしまう事故を防ぐ(例: '田中' が '田中太郎' の一部として
    誤爆するのを防ぐ)。
    """
    result = text
    for original in sorted(mapping.keys(), key=len, reverse=True):
        replacement = mapping[original]
        if not original:
            continue
        result = result.replace(original, replacement)
    return result