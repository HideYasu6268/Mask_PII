# -*- coding: utf-8 -*-
"""
gemini_client.py

Google Gemini APIとの通信を行う(このアプリで唯一、外部にネットワーク通信する箇所)。
送信するのは必ず「匿名化後の文章」であり、原文やPIIそのものは送らない。

- APIキーは同じディレクトリの gemini_api_key.txt から読み込む。1行に1つずつ、
  改行区切りで複数書ける。上から順に使い、あるキーがレート制限(429)に達したら
  次の行のキーに自動で切り替えて再試行する(全キーが429の場合のみエラーになる)。
  このファイルは.gitignore対象で、リポジトリには仮の値しか入っていない。
  実際に使う際は自分のAPIキーに書き換えること。
- 使用するGeminiモデルは、キーごとに models.list() で「実際に使えるflash系モデル」を
  取得し、バージョンが新しいものから順に試す(get_flash_models参照)。
  アカウントごとに使えるモデルが異なるため、以前のような接頭辞(AIzaSy.../AQ.)による
  推奨モデルの決め打ちは廃止した。一覧の取得に失敗した場合のみ、固定の
  DEFAULT_MODEL → FALLBACK_MODEL にフォールバックする。
  ・404(モデルが使えない)/ 503等のサーバー側エラー: 次に新しいモデルで再試行
  ・429(レート制限): 既定では同じキーで古いモデルを先に試し(TRY_OLDER_MODEL_ON_429=True)、
    それも駄目なら次のAPIキーへ(Falseなら429で即・次のキーへ)
- 判断ステップ(decide_reply_style)はflash-lite系(get_lite_models)を新しい順に試し、
  生成ステップ(generate_reply)はflash系(get_flash_models)を使う。クォータを分けるため。
- プロンプトテンプレートは同じディレクトリの3ファイルから読み込む(PROMPT_STYLES参照)。
  テンプレート中の {reply_intent} / {anonymized_text} が、それぞれアプリ上の
  「どういう返信をしたいか」欄の内容・匿名化後の文章に置き換わる。
- 税務・会計の質問に対する根拠を裏付けられるよう、Google検索によるグラウンディング
  (Grounding with Google Search)を有効にして呼び出す。実際に検索したクエリや
  参照したWebページはレスポンスの grounding_metadata から取り出し、
  GeminiReplyResult.sources / search_queries としてGUI側に返す
  (検索が行われなかった場合は両方とも空になる)。

# AIエージェント機能(判断→生成の2段階)について
- 返信プロンプトは short/standard/long_search の3種類を用意している(PROMPT_STYLES)。
  generate_reply_with_agent() を呼ぶと、まず decide_reply_style() が匿名化後の文章と
  返信方針からどのプロンプトを使うべきかをGeminiに判断させ、その結果に基づいて
  generate_reply() が実際の返信文を生成する。
- 判断ステップ(decide_reply_style)・生成ステップ(generate_reply)とも、追加の依存を
  増やさないため google-genai SDK に統一している。判断ステップはGoogle検索による
  グラウンディングが不要な単純な分類呼び出しなので、tools を指定しない
  generate_content 呼び出し1回で済ませる(exe配布時の軽量さを優先。
  READMEの「Ollamaを使わない理由」と同じ考え方)。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Callable

# PyInstallerでexe化した場合、__file__はexeの実体とは別の展開先(onefileなら
# 起動のたびに消える一時フォルダ)を指してしまうため、frozen時はexe自身の
# あるフォルダを基準にする(でないとAPIキー/プロンプトの編集内容がexe再起動の
# たびに消えたり、意図しない場所に書き込まれたりする)。
APP_DIR = (
    os.path.dirname(os.path.abspath(sys.executable))
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
API_KEY_PATH = os.path.join(APP_DIR, "gemini_api_key.txt")

# 返信プロンプトの3種類。キーはアプリ内部・GUI間で使う識別子、値はファイル名。
# "standard" は既存の reply_prompt_template.txt をそのまま流用する(後方互換のため
# ファイル名は変更しない)。
PROMPT_STYLES: dict[str, str] = {
    "short": "reply_prompt_short.txt",
    "standard": "reply_prompt_template.txt",
    "long_search": "reply_prompt_long_search.txt",
}
PROMPT_TEMPLATE_PATHS: dict[str, str] = {
    style: os.path.join(APP_DIR, filename) for style, filename in PROMPT_STYLES.items()
}
# 後方互換用(main.pyの旧コードや外部からの参照向けに残す)。"standard" 用のパスと同じ。
PROMPT_TEMPLATE_PATH = PROMPT_TEMPLATE_PATHS["standard"]

DEFAULT_STYLE = "standard"

# GUI表示用の日本語ラベル。
STYLE_LABELS: dict[str, str] = {
    "short": "短め",
    "standard": "通常",
    "long_search": "検索+丁寧(長め)",
}

# ---------------------------------------------------------------------------
# モデル選択の設定
# ---------------------------------------------------------------------------
# models.list() でモデル一覧を取得できなかった場合(通信エラー・権限不足など)に
# だけ使う固定のフォールバック。通常は get_flash_models() が返す新しい順の
# リストが使われる。
DEFAULT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.6-flash"

# 判断ステップ(decide_reply_style)専用。生成ステップと違い検索が不要な単純な
# 分類なので、flash-lite系を新しい順に試す(get_lite_models参照)。これにより
# 生成ステップ用のflash系モデルの無料枠(RPD/RPM)を消費しない。
# 下は models.list() の取得に失敗した場合だけ使う固定のフォールバック。
DECIDE_FALLBACK_MODELS = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]
DECIDE_MODEL = DECIDE_FALLBACK_MODELS[0]  # 後方互換用

# long_search(Google検索グラウンディングあり)で試すモデル数の上限。None なら
# models.list() の全候補を新しい順に試す。アカウント(無料枠など)によっては新しい
# モデルで検索ツールが使えないことがあるため、MAX_MODELS_TO_TRY の制限は受けない。
# 429/404は0.2秒程度で返るので、全候補を試しても待ち時間は小さい。
GROUNDING_MAX_MODELS: int | None = None

# 「そのキー+そのモデル+検索ツール」の組み合わせで失敗した場合、この秒数だけ
# スキップして、次回以降は通るモデルへ直行する(期限が切れたら再挑戦する)。
GROUNDING_BLOCK_TTL_SEC = 30 * 60

# preview / exp 系のモデルを候補に含めるか。Falseにすると安定版のみを使う。
# 同じバージョンなら安定版を preview より先に試す。
INCLUDE_PREVIEW_MODELS = True

# 1回のAPI呼び出しで、1つのキーにつき最大何個のモデルを新しい順に試すか。
MAX_MODELS_TO_TRY = 4

# 429(レート制限)は通常モデル単位のクォータなので、同じキーで古いモデルなら
# 通る場合がある。Trueにすると、429でも次のキーへ行く前に同じキーで古いモデルを
# 試す(Falseにすると従来通り「429なら次のキーへ」)。
TRY_OLDER_MODEL_ON_429 = True

# モデル一覧のキャッシュ有効期間(秒)。アプリ起動中は使い回し、返信生成のたびに
# 一覧取得を行わない。
MODEL_CACHE_TTL_SEC = 6 * 60 * 60

# gemini_api_key.txt に最初から入っている仮の値。これがそのまま残っている場合は
# 未設定とみなし、実際にはAPIを呼ばずにエラーを返す。
_PLACEHOLDER_KEY = "YOUR_API_KEY_HERE"

# デバッグ出力(不要になったら False に)
DEBUG_LOG = True


def _log(msg: str) -> None:
    if DEBUG_LOG:
        print(f"[gemini_client {time.strftime('%H:%M:%S')}] {msg}", file=sys.stderr, flush=True)


def _notify(on_progress: Callable[[str], None] | None, msg: str) -> None:
    """GUI等へ進捗メッセージを通知する(コールバック未指定・失敗時は何もしない)。
    ワーカースレッドから呼ばれる前提なので、UI更新は呼び出し側でメインスレッドに渡すこと。"""
    if on_progress is None:
        return
    try:
        on_progress(msg)
    except Exception:  # noqa: BLE001
        pass


def _key_label(idx: int, total: int) -> str:
    """複数キーのときだけ「(キー2/3)」のような表示を返す。"""
    return f"(キー{idx}/{total})" if total > 1 else ""


def _key_id(api_key: str) -> str:
    """ログ用にキーを識別する(末尾4文字のみ。キー全体は出さない)。"""
    return f"...{api_key[-4:]}"


# {(キー識別子, モデル名): 失敗した時刻} 検索ツール付きで使えなかった組み合わせの記憶。
_grounding_blocked: dict[tuple[str, str], float] = {}
_grounding_lock = threading.Lock()


def _mark_grounding_blocked(api_key: str, model: str) -> None:
    with _grounding_lock:
        _grounding_blocked[(_key_id(api_key), model)] = time.time()


def _is_grounding_blocked(api_key: str, model: str) -> bool:
    with _grounding_lock:
        t = _grounding_blocked.get((_key_id(api_key), model))
        if t is None:
            return False
        if time.time() - t > GROUNDING_BLOCK_TTL_SEC:
            del _grounding_blocked[(_key_id(api_key), model)]
            return False
        return True


# 判断ステップ(decide_reply_style)でGeminiに渡すシステムプロンプト。
# 出力はJSON形式({"style": ..., "reason": ...})に固定する。
DECIDE_STYLE_SYSTEM_PROMPT = """あなたは、会計事務所のスタッフが送るメール返信の
「作成方針」を判断するアシスタントです。与えられた匿名化済みメール本文(個人情報は
[TYPE_連番] 形式のタグに置き換え済み)と、スタッフが指定した返信方針(空の場合あり)を
踏まえて、返信文をどの方針で作成すべきか、次の3種類から1つだけ選んでください。

- short: 「資料を受け取りました」「了解しました」「承知しました」のような、単純な
  確認・お礼・日程調整などで完結する内容。簡潔な返信で十分な場合。
- standard: 上記ほど単純ではないが、税務・会計について踏み込んだ調査を要するほどでも
  ない、一般的なやり取りの場合。
- long_search: 税務・会計についての具体的な質問・相談が含まれ、根拠(法令・通達・
  情報源など)を調べたうえで、丁寧かつ詳しく回答する必要がある場合。

スタッフが返信方針で「短く」「簡潔に」等を明示している場合はshortを、
「詳しく」「丁寧に」「調べて」等を明示している場合はlong_searchを優先してください。

出力は必ず次のJSON形式のみとし、説明文やコードブロック記号は一切含めないでください。

{"style": "short", "reason": "資料受領の確認のみのため"}

style は short/standard/long_search のいずれか、reason は判断理由を日本語で
15〜30文字程度の一文で書いてください。
"""


# Groq API設定(gskで始まるキーの場合はGroqで通信)
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "openai/gpt-oss-120b"

# gpt-oss系の推論強度("low" / "medium" / "high")。未指定だとGroq側の既定は "medium"。
# 推論トークンは max_tokens にも含まれるため、速度やトークン消費を抑えたい場合は
# "low" に、品質を優先する場合は "medium" / "high" にする。
GROQ_EFFORT_DECIDE = "medium"  # 判断ステップ
GROQ_EFFORT_BY_STYLE: dict[str, str] = {
    "short": "medium",
    "standard": "medium",
    "long_search": "medium",
}
GROQ_EFFORT_DEFAULT = "medium"


def _is_groq_key(api_key: str) -> bool:
    """gskで始まるキーはGroq APIキーと判定する。"""
    return api_key.startswith("gsk")


def _call_groq_chat(
    api_key: str,
    messages: list[dict],
    model: str = GROQ_MODEL,
    temperature: float = 0.3,
    max_tokens: int = 2500,
    reasoning_effort: str | None = GROQ_EFFORT_DEFAULT,
) -> str:
    """Groq API(OpenAI互換)を呼び出す。外部ライブラリ依存を避けるため標準のurllibを使用。

    reasoning_effort: gpt-oss系の推論強度("low" / "medium" / "high")。
    None を渡すとパラメータを送らず、Groq側の既定("medium")になる。
    """
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    if reasoning_effort:
        payload["reasoning_effort"] = reasoning_effort
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        GROQ_API_URL,
        data=data,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": "Mask_PII/1.0",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        res_data = json.loads(resp.read().decode("utf-8"))
        choices = res_data.get("choices", [])
        if not choices:
            raise GeminiError("Groq APIの応答に choices が含まれていませんでした。")
        content = choices[0].get("message", {}).get("content", "")
        return content or ""


class GeminiError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# 利用可能なflashモデルの取得(新しい順)
# ---------------------------------------------------------------------------
# 例: "gemini-2.5-flash" / "gemini-2.5-flash-001" / "gemini-3.0-flash-preview"
# 先頭が "gemini-<数字>-flash" の形のものだけを対象にする。
# ("gemini-live-2.5-flash-..." のように数字が先頭に来ないものは自動的に対象外)
_FLASH_NAME_RE = re.compile(r"^gemini-(?P<ver>\d+(?:\.\d+)*)-flash(?:-(?P<suffix>.+))?$")

# サフィックスにこれらを含むものは、テキスト返信用途ではないので除外する。
_EXCLUDED_SUFFIX_KEYWORDS = (
    "lite", "image", "tts", "audio", "live", "embedding",
    "robotics", "computer-use", "8b", "thinking",
)

# {キーのハッシュ: (取得時刻, [モデル名, ...新しい順])}
_model_cache: dict[str, tuple[float, list[str]]] = {}
_model_cache_lock = threading.Lock()


def _static_fallback_models(lite: bool = False) -> list[str]:
    """モデル一覧を取得できなかった場合に使う固定リスト(重複除去済み)。
    lite=True なら判断ステップ用のlite系、Falseなら生成ステップ用のflash系。"""
    result: list[str] = []
    candidates = DECIDE_FALLBACK_MODELS if lite else (DEFAULT_MODEL, FALLBACK_MODEL)
    for m in candidates:
        if m not in result:
            result.append(m)
    return result


def _parse_model_name(name: str, lite: bool = False) -> tuple[str, tuple[int, ...], int] | None:
    """モデル名を解析し、(短い名前, バージョンのタプル, 品質ランク) を返す。
    対象外のモデルなら None。品質ランクは小さいほど安定版寄り:
      0: 別名(gemini-2.5-flash)  1: 日付/連番固定(gemini-2.0-flash-001)
      2以上: preview / exp / その他のサフィックス付き

    lite=False: 通常のflash系(gemini-3.8-flash など。lite付きは除外)。
    lite=True : flash-lite系(gemini-3.5-flash-lite / gemini-3.1-flash-lite-preview など。
                lite-image / lite-tts のような用途違いは除外)。
    """
    short = name.split("/")[-1]  # "models/gemini-..." の接頭辞を外す
    m = _FLASH_NAME_RE.match(short)
    if not m:
        return None
    suffix = (m.group("suffix") or "").lower()
    if lite:
        # "lite" または "lite-xxx" だけを対象にし、"lite" 以降を品質判定に使う
        if suffix == "lite":
            tail = ""
        elif suffix.startswith("lite-"):
            tail = suffix[len("lite-"):]
        else:
            return None
    else:
        tail = suffix
    if any(k in tail for k in _EXCLUDED_SUFFIX_KEYWORDS):
        return None
    version = tuple(int(x) for x in m.group("ver").split("."))
    if not tail:
        quality = 0
    elif re.fullmatch(r"\d{3}", tail):
        quality = 1
    elif "preview" in tail or "exp" in tail:
        quality = 2
    else:
        quality = 3
    return short, version, quality


def _parse_flash_model(name: str) -> tuple[str, tuple[int, ...], int] | None:
    """後方互換用。通常のflash系の解析(_parse_model_name参照)。"""
    return _parse_model_name(name, lite=False)


def get_flash_models(
    api_key: str, client=None, force_refresh: bool = False, all_models: bool = False
) -> list[str]:
    """生成ステップ用: そのキーで使える通常のflash系モデルを新しい順で返す(_get_models参照)。
    all_models=True なら MAX_MODELS_TO_TRY で切らず、全候補を返す。"""
    return _get_models(api_key, client, force_refresh, lite=False, all_models=all_models)


def get_lite_models(api_key: str, client=None, force_refresh: bool = False) -> list[str]:
    """判断ステップ用: そのキーで使えるflash-lite系モデルを新しい順で返す(_get_models参照)。"""
    return _get_models(api_key, client, force_refresh, lite=True)


def _get_models(
    api_key: str,
    client=None,
    force_refresh: bool = False,
    lite: bool = False,
    all_models: bool = False,
) -> list[str]:
    """そのAPIキーで実際に使えるflash系(lite=Trueならflash-lite系)モデルを、
    新しい順のリストで返す。

    - models.list() の結果から generateContent 対応のものだけを取り出し、
      lite / image / tts / live 等の用途違いは除外する。
    - バージョン(2.5 < 3.0 < 3.8 ...)が新しい順。同じバージョンでは安定版を
      preview より先にする。同じバージョン・同じ種別(安定/preview)で複数ある場合は
      最も安定寄りの1つだけ残す(例: gemini-2.5-flash と gemini-2.5-flash-001 は前者のみ)。
    - 既定では最大 MAX_MODELS_TO_TRY 個まで(all_models=True なら全候補)。
      全候補のリストを MODEL_CACHE_TTL_SEC 秒キャッシュする。
    - 一覧の取得に失敗、または該当が0件の場合は固定の DEFAULT_MODEL → FALLBACK_MODEL
      を返す(この場合はキャッシュしない=次回また取得を試みる)。
    """
    label = "lite" if lite else "flash"
    cache_id = hashlib.sha256(api_key.encode("utf-8")).hexdigest()[:16] + f":{label}"
    now = time.time()
    if not force_refresh:
        with _model_cache_lock:
            cached = _model_cache.get(cache_id)
        if cached and now - cached[0] < MODEL_CACHE_TTL_SEC:
            _log(f"key {_key_id(api_key)}: {label}モデル一覧はキャッシュ利用 -> {cached[1]}")
            return list(cached[1] if all_models else cached[1][:MAX_MODELS_TO_TRY])

    try:
        if client is None:
            from google import genai

            client = genai.Client(api_key=api_key)

        # (バージョン, previewか) ごとに、最も安定寄りの1つを残す
        best: dict[tuple[tuple[int, ...], bool], tuple[int, str]] = {}
        all_names: list[str] = []  # models.list() の生の結果(ログ用)
        for model_info in client.models.list(config={"page_size": 100}):
            name = getattr(model_info, "name", None)
            if not name:
                continue
            all_names.append(name)
            actions = getattr(model_info, "supported_actions", None)
            if actions is not None and "generateContent" not in actions:
                continue
            parsed = _parse_model_name(name, lite)
            if parsed is None:
                continue
            short, version, quality = parsed
            is_preview = quality >= 2
            if is_preview and not INCLUDE_PREVIEW_MODELS:
                continue
            key = (version, is_preview)
            current = best.get(key)
            if current is None or quality < current[0]:
                best[key] = (quality, short)

        _log(f"key {_key_id(api_key)}: models.list() 全{len(all_names)}件")
        _log(f"  うちgemini系: {[n for n in all_names if 'gemini' in n]}")

        # バージョン降順、同バージョンでは安定版(is_preview=False)が先
        ordered = sorted(best.items(), key=lambda kv: (kv[0][0], not kv[0][1]), reverse=True)
        all_candidates = [value[1] for _, value in ordered]
        _log(f"  {label}候補(絞り込み後・新しい順): {all_candidates}")
        candidates = all_candidates
        _log(f"  試行対象(通常は上位{MAX_MODELS_TO_TRY}件 / 検索ありは全件): {candidates[:MAX_MODELS_TO_TRY]}")
    except Exception as e:  # noqa: BLE001
        _log(
            f"key {_key_id(api_key)}: models.list() 失敗 -> 固定フォールバック "
            f"{_static_fallback_models(lite)} ({type(e).__name__}: {e})"
        )
        return _static_fallback_models(lite)

    if not candidates:
        _log(f"key {_key_id(api_key)}: {label}該当モデル0件 -> 固定フォールバック {_static_fallback_models(lite)}")
        return _static_fallback_models(lite)

    with _model_cache_lock:
        _model_cache[cache_id] = (now, candidates)
    return list(candidates if all_models else candidates[:MAX_MODELS_TO_TRY])


@dataclass
class GroundingSource:
    title: str
    uri: str


@dataclass
class GeminiReplyResult:
    text: str
    sources: list[GroundingSource] = field(default_factory=list)
    search_queries: list[str] = field(default_factory=list)
    # generate_reply_with_agent() 経由で呼んだ場合のみ設定される
    # (直接 generate_reply() を呼んだ場合は呼び出し元が指定したstyleがそのまま入る)。
    style: str = DEFAULT_STYLE
    style_reason: str = ""
    # 実際に返信生成に使われたモデル名(GUI側で表示できる)。
    model: str = ""


def _load_api_keys() -> list[str]:
    """gemini_api_key.txtを改行区切りで読み込み、有効なAPIキーのリストを返す
    (空行・仮の値の行は無視する)。複数行ある場合、上から順にレート制限(429)への
    フォールバック先として使う(generate_reply参照)。
    """
    if not os.path.isfile(API_KEY_PATH):
        raise GeminiError(
            f"APIキーファイルが見つかりません: {API_KEY_PATH}"
            " (プロジェクト直下に gemini_api_key.txt を作成し、1行目にAPIキーを書いてください)"
        )
    with open(API_KEY_PATH, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    keys = [line.strip() for line in lines if line.strip() and line.strip() != _PLACEHOLDER_KEY]
    if not keys:
        raise GeminiError(
            "gemini_api_key.txt が未設定です(仮の値のままです)。"
            " 実際のGemini APIキーに書き換えてください"
            "(複数ある場合は1行に1つずつ、改行区切りで書いてください)。"
        )
    return keys


def _load_prompt_template(style: str = DEFAULT_STYLE) -> str:
    path = PROMPT_TEMPLATE_PATHS.get(style, PROMPT_TEMPLATE_PATHS[DEFAULT_STYLE])
    if not os.path.isfile(path):
        raise GeminiError(f"プロンプトテンプレートが見つかりません: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def build_prompt(anonymized_text: str, reply_intent: str, style: str = DEFAULT_STYLE) -> str:
    template = _load_prompt_template(style)
    return template.format(
        reply_intent=reply_intent.strip() or "(特に指定なし。文面から適切に判断してください)",
        anonymized_text=anonymized_text,
    )


def _extract_json_object(raw: str) -> dict | None:
    """LLM出力からJSONオブジェクト部分を頑健に取り出す(decide_reply_style用)。
    抽出・パースに失敗した場合は None を返す(例外は送出しない。呼び出し元で
    既定値へのフォールバックを行うため)。
    """
    cleaned = raw.strip()
    cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    try:
        obj = json.loads(cleaned)
        return obj if isinstance(obj, dict) else None
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if not match:
        return None
    try:
        obj = json.loads(match.group(0))
        return obj if isinstance(obj, dict) else None
    except json.JSONDecodeError:
        return None


def decide_reply_style(
    anonymized_text: str,
    reply_intent: str,
    on_progress: Callable[[str], None] | None = None,
) -> tuple[str, str]:
    """匿名化後の文章と返信方針から、3種類の返信プロンプト(PROMPT_STYLES)のうち
    どれを使うべきかをLLMに判断させる。
    - gskで始まるキー: Groq API (openai/gpt-oss-120b)
    - その他のキー: Gemini API(そのキーで使える最新のflash-lite系モデルから順に試す。
      生成ステップのflash系の無料枠を消費しないため)

    on_progress: 「どのモデルに何を依頼中か」を文字列で受け取るコールバック(省略可)。

    戻り値: (style, reason)。style は PROMPT_STYLES のキーのいずれか。
    """
    try:
        api_keys = _load_api_keys()
    except GeminiError as e:
        return DEFAULT_STYLE, f"判断呼び出し前にエラー({e})のため既定の方針を使用"

    user_content = (
        f"返信の方針(ユーザー指定): {reply_intent.strip() or '(特に指定なし)'}\n\n"
        f"匿名化された原文:\n{anonymized_text}"
    )

    last_error: Exception | None = None
    for key_idx, api_key in enumerate(api_keys, 1):
        content = None
        if _is_groq_key(api_key):
            _notify(on_progress, f"Groq({GROQ_MODEL}) に返信方針の判断を依頼中... {_key_label(key_idx, len(api_keys))}")
            try:
                messages = [
                    {"role": "system", "content": DECIDE_STYLE_SYSTEM_PROMPT},
                    {"role": "user", "content": user_content},
                ]
                content = _call_groq_chat(
                    api_key,
                    messages,
                    model=GROQ_MODEL,
                    temperature=0.1,
                    max_tokens=1000,
                    reasoning_effort=GROQ_EFFORT_DECIDE,
                )
            except urllib.error.HTTPError as e:
                last_error = e
                _log(f"[decide] Groq HTTPError key {_key_id(api_key)}: {e}")
                continue
            except Exception as e:  # noqa: BLE001
                last_error = e
                _log(f"[decide] Groq Error key {_key_id(api_key)}: {type(e).__name__}: {e}")
                continue
        else:
            try:
                from google import genai
                from google.genai import types
                from google.genai import errors as genai_errors
            except ImportError:
                last_error = RuntimeError("google-genaiが未インストールです")
                continue

            # 最新のflash系は内部で「思考」にトークンを使うことがあり、出力上限が
            # 小さいと本文が空になるため、JSONが十分収まる範囲で余裕を持たせる。
            config = types.GenerateContentConfig(
                system_instruction=DECIDE_STYLE_SYSTEM_PROMPT,
                temperature=0.1,
                max_output_tokens=1024,
            )
            client = genai.Client(api_key=api_key)
            for m in get_lite_models(api_key, client):
                _log(f"[decide] key {_key_id(api_key)} -> {m} を試行中...")
                _notify(on_progress, f"{m} に返信方針の判断を依頼中... {_key_label(key_idx, len(api_keys))}")
                try:
                    resp = client.models.generate_content(
                        model=m, contents=user_content, config=config,
                    )
                    content = getattr(resp, "text", None)
                    _log(f"[decide] OK: {m} (text={'あり' if content else '空'})")
                    if content:
                        break
                except genai_errors.ClientError as e:
                    last_error = e
                    _log(f"[decide] ClientError: {m} code={e.code} {e}")
                    if e.code == 429 and not TRY_OLDER_MODEL_ON_429:
                        break  # 次のキーへ
                    continue  # 429(古いモデルを試す設定)/404/400等は次のモデルへ
                except Exception as e:  # noqa: BLE001
                    last_error = e
                    _log(f"[decide] Error: {m} {type(e).__name__}: {e}")
                    continue

        if not content:
            continue

        obj = _extract_json_object(content)
        if obj is None:
            last_error = RuntimeError(f"判断ステップの応答をJSONとして解釈できませんでした: {content[:200]!r}")
            continue

        style = str(obj.get("style", "")).strip()
        reason = str(obj.get("reason", "")).strip()
        if style not in PROMPT_STYLES:
            last_error = RuntimeError(f"判断ステップが未知のstyleを返しました: {style!r}")
            continue

        return style, reason or "(判断理由は返されませんでした)"

    return DEFAULT_STYLE, f"判断呼び出しに失敗したため既定の方針(標準)を使用({last_error})"


def generate_reply(
    anonymized_text: str,
    reply_intent: str,
    style: str = DEFAULT_STYLE,
    model: str | None = None,
    on_progress: Callable[[str], None] | None = None,
) -> GeminiReplyResult:
    """匿名化後の文章と返信方針から、指定されたstyle(PROMPT_STYLES参照)の
    プロンプトを組み立て、LLM APIに送信して返信文の案を取得する。
    - gskで始まるキー: Groq API (openai/gpt-oss-120b)
    - その他のキー: Gemini API (Google検索によるグラウンディングを使用)

    model を省略(None)、または DEFAULT_MODEL を渡した場合は、そのキーで使える
    最新のflash系モデルから順に試す(get_flash_models参照)。それ以外のモデル名を
    渡した場合は、そのモデルだけを使う(フォールバックなし)。

    どのstyleを使うべきかを自動判断させたい場合は、この関数を直接呼ぶのではなく
    generate_reply_with_agent() を使うこと。

    gemini_api_key.txtに複数のAPIキーが書かれている場合、上から順に試し、
    レート制限(429)に達したキーは次のキーに自動で切り替えて再試行する。

    on_progress: 「どのモデルに何を依頼中か」を文字列で受け取るコールバック(省略可)。
    モデルを試すたびに呼ばれる(呼び出しはワーカースレッド上で行われる)。
    """
    if not anonymized_text.strip():
        raise GeminiError("匿名化後の文章が空です。")
    if style not in PROMPT_STYLES:
        style = DEFAULT_STYLE

    api_keys = _load_api_keys()
    prompt = build_prompt(anonymized_text, reply_intent, style=style)

    # DEFAULT_MODELの明示指定は、従来から「自動選択」と同じ扱い(後方互換)。
    pinned_model = model if (model and model != DEFAULT_MODEL) else None

    resp = None
    chosen_model = ""
    last_error: Exception | None = None
    tried_models: list[str] = []

    for key_idx, api_key in enumerate(api_keys, 1):
        if _is_groq_key(api_key):
            _notify(on_progress, f"Groq({GROQ_MODEL}) に文章作成を依頼中... {_key_label(key_idx, len(api_keys))}")
            try:
                messages = [{"role": "user", "content": prompt}]
                effort = GROQ_EFFORT_BY_STYLE.get(style, GROQ_EFFORT_DEFAULT)
                text = _call_groq_chat(
                    api_key,
                    messages,
                    model=GROQ_MODEL,
                    temperature=0.3,
                    max_tokens=2500,
                    reasoning_effort=effort,
                )
                if text:
                    return GeminiReplyResult(
                        text=text,
                        sources=[],
                        search_queries=[],
                        style=style,
                        model=GROQ_MODEL,
                    )
            except urllib.error.HTTPError as e:
                last_error = e
                _log(f"[generate_reply] Groq HTTPError key {_key_id(api_key)}: {e}")
                continue
            except Exception as e:  # noqa: BLE001
                last_error = e
                _log(f"[generate_reply] Groq Error key {_key_id(api_key)}: {type(e).__name__}: {e}")
                continue
        else:
            try:
                from google import genai
                from google.genai import types
                from google.genai import errors as genai_errors
            except ImportError as e:
                last_error = e
                continue

            # 検索グラウンディングは long_search のときだけ付ける。
            # (無料枠では Gemini 3系 + 検索ツールが即429になるため、short/standard は
            #  検索なしで最新モデルをそのまま使う。)
            if style == "long_search":
                config = types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                )
            else:
                config = types.GenerateContentConfig()
            client = genai.Client(api_key=api_key)
            use_search = style == "long_search"
            if pinned_model:
                models_to_try = [pinned_model]
            elif use_search:
                # 検索ツール付きは、新しいモデルほど無料枠で使えないことがあるため、
                # 全候補を新しい順に試す。過去に失敗した組み合わせは(期限内)スキップし、
                # 全部スキップ対象になってしまう場合は念のため全件を再試行する。
                candidates = get_flash_models(api_key, client, all_models=True)
                models_to_try = [m for m in candidates if not _is_grounding_blocked(api_key, m)]
                if not models_to_try:
                    models_to_try = candidates
                if GROUNDING_MAX_MODELS is not None:
                    models_to_try = models_to_try[:GROUNDING_MAX_MODELS]
            else:
                models_to_try = get_flash_models(api_key, client)

            _log(f"[generate_reply] key {_key_id(api_key)} 試行モデル順: {models_to_try}")
            skipped: list[str] = []  # このキーで使えなかったモデル(進捗表示用)
            for m in models_to_try:
                if m not in tried_models:
                    tried_models.append(m)
                _log(f"[generate_reply] -> {m} を試行中...")
                note = f"{m} に文章作成を依頼中..."
                if use_search:
                    note += "(Google検索あり)"
                if skipped:
                    note += f" ※先に試した{len(skipped)}件は使用不可"
                _notify(on_progress, f"{note} {_key_label(key_idx, len(api_keys))}")
                t0 = time.time()
                try:
                    resp = client.models.generate_content(model=m, contents=prompt, config=config)
                    chosen_model = m
                    _log(f"[generate_reply] OK: {m} ({time.time() - t0:.1f}秒)")
                    break
                except genai_errors.ClientError as e:
                    last_error = e
                    skipped.append(m)
                    _log(f"[generate_reply] ClientError: {m} code={e.code} ({time.time() - t0:.1f}秒) {e}")
                    if use_search and e.code in (400, 404, 429):
                        # このキーではこのモデルに検索ツールを付けられない(または枠がない)
                        # ので記憶し、しばらく試さない。
                        _mark_grounding_blocked(api_key, m)
                    if e.code == 429:
                        if TRY_OLDER_MODEL_ON_429 or use_search:
                            continue  # 同じキーで古いモデルを試す
                        break  # 次のキーへ
                    if e.code in (400, 404):
                        # モデルが使えない/そのモデルが検索ツール等に非対応の場合は、
                        # 次に新しいモデルで再試行する。
                        continue
                    raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
                except genai_errors.ServerError as e:
                    # 503(混雑)等。新しいモデルほど起きやすいので次のモデルへ。
                    last_error = e
                    skipped.append(m)
                    _log(
                        f"[generate_reply] ServerError: {m} code={getattr(e, 'code', '?')} "
                        f"({time.time() - t0:.1f}秒) {e}"
                    )
                    continue
                except Exception as e:  # noqa: BLE001
                    _log(f"[generate_reply] 想定外の例外: {m} {type(e).__name__}: {e}")
                    raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
            if resp is not None:
                break

    if resp is None:
        _log(
            f"[generate_reply] 全滅: 試したモデル={tried_models} / "
            f"最後のエラー={type(last_error).__name__}: {last_error}"
        )
        last_code = getattr(last_error, "code", None)
        if tried_models and last_code == 404:
            message = (
                f"指定されたモデル({', '.join(tried_models)})がいずれも利用できませんでした。"
                " gemini_client.pyの設定を見直してください。"
            )
        elif len(api_keys) == 1:
            message = f"API呼び出しでエラーまたはレート制限(429)に達しました({last_error})。しばらく待ってから再試行してください。"
        else:
            message = (
                f"登録されている{len(api_keys)}件のAPIキーすべてで"
                f"エラーまたはレート制限(429)に達しました({last_error})。しばらく待つか、有効なAPIキーを追加してください。"
            )
        raise GeminiError(message) from last_error

    text = getattr(resp, "text", None)
    if not text:
        raise GeminiError("APIの応答が空でした。")

    sources, search_queries = _extract_grounding(resp)
    return GeminiReplyResult(
        text=text,
        sources=sources,
        search_queries=search_queries,
        style=style,
        model=chosen_model,
    )


def generate_reply_with_agent(
    anonymized_text: str,
    reply_intent: str,
    model: str | None = None,
    on_progress: Callable[[str], None] | None = None,
) -> GeminiReplyResult:
    """「🤖 AIによる返答生成」のエージェント版。
    1. decide_reply_style() で、匿名化後の文章と返信方針からどのプロンプト
       (short/standard/long_search)を使うべきかをまずGeminiに判断させる
       (google-genai SDK、tools無しの軽量呼び出し)。
    2. その判断結果のstyleで generate_reply() を呼び、実際の返信文を生成する
       (google-genai SDK、Google検索によるグラウンディングを使用)。

    戻り値の GeminiReplyResult.style / style_reason に、実際に使われたstyleと
    その判断理由が入るので、GUI側で「どの方針を選んだか」を表示できる。
    判断ステップ自体が失敗した場合でも例外にはせず既定値にフォールバックするため、
    この関数が送出する例外は実質的に generate_reply() 側のもの(APIキー未設定・
    通信エラーなど)のみになる。
    """
    if not anonymized_text.strip():
        raise GeminiError("匿名化後の文章が空です。")

    style, reason = decide_reply_style(anonymized_text, reply_intent, on_progress=on_progress)
    result = generate_reply(
        anonymized_text, reply_intent, style=style, model=model, on_progress=on_progress
    )
    result.style_reason = reason
    return result


def _extract_grounding(resp) -> tuple[list[GroundingSource], list[str]]:
    """Google検索によるグラウンディングの結果(検索クエリ・参照したWebページ)を
    レスポンスから取り出す。フィールドの構造はSDKのバージョンに依存しやすく、
    かつ検索が行われなかった応答では存在しないため、全体をベストエフォートで
    取り出し、失敗しても本文の生成自体は成功として扱う(空リストを返すのみ)。
    """
    sources: list[GroundingSource] = []
    search_queries: list[str] = []
    try:
        candidates = getattr(resp, "candidates", None) or []
        if not candidates:
            return sources, search_queries
        metadata = getattr(candidates[0], "grounding_metadata", None)
        if not metadata:
            return sources, search_queries

        search_queries = list(getattr(metadata, "web_search_queries", None) or [])

        seen_uris: set[str] = set()
        for chunk in getattr(metadata, "grounding_chunks", None) or []:
            web = getattr(chunk, "web", None)
            uri = getattr(web, "uri", None) if web else None
            if not uri or uri in seen_uris:
                continue
            seen_uris.add(uri)
            title = getattr(web, "title", None) or uri
            sources.append(GroundingSource(title=title, uri=uri))
    except Exception:  # noqa: BLE001
        return sources, search_queries
    return sources, search_queries