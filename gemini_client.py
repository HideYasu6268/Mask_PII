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
  各キーはアカウントごとに使えるモデルが異なるため、キーの接頭辞から
  推奨モデルを自動選択する(_preferred_model_for_key参照。AIzaSy...形式→
  DEFAULT_MODEL、AQ.形式→FALLBACK_MODEL)。推奨モデルが404(NOT_FOUND)の
  場合は、そのキーのままもう一方のモデルで再試行する。
- プロンプトテンプレートは同じディレクトリの reply_prompt_template.txt から読み込む。
  テンプレート中の {reply_intent} / {anonymized_text} が、それぞれアプリ上の
  「どういう返信をしたいか」欄の内容・匿名化後の文章に置き換わる。
- 税務・会計の質問に対する根拠を裏付けられるよう、Google検索によるグラウンディング
  (Grounding with Google Search)を有効にして呼び出す。実際に検索したクエリや
  参照したWebページはレスポンスの grounding_metadata から取り出し、
  GeminiReplyResult.sources / search_queries としてGUI側に返す
  (検索が行われなかった場合は両方とも空になる)。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field

APP_DIR = os.path.dirname(os.path.abspath(__file__))
API_KEY_PATH = os.path.join(APP_DIR, "gemini_api_key.txt")
PROMPT_TEMPLATE_PATH = os.path.join(APP_DIR, "reply_prompt_template.txt")

DEFAULT_MODEL = "gemini-2.5-flash"

# DEFAULT_MODELが404(NOT_FOUND、新規アカウントで利用不可などモデル自体が
# 使えないケース)になった場合に自動で切り替えて再試行するモデル。
FALLBACK_MODEL = "gemini-3.6-flash"

# gemini_api_key.txt に最初から入っている仮の値。これがそのまま残っている場合は
# 未設定とみなし、実際にはAPIを呼ばずにエラーを返す。
_PLACEHOLDER_KEY = "YOUR_API_KEY_HERE"


def _preferred_model_for_key(api_key: str) -> str:
    """APIキーの接頭辞からアカウント種別を判定し、そのアカウントで通ることが
    分かっているモデルを返す(AIzaSy...形式のキー→DEFAULT_MODEL、
    AQ.形式のキー→FALLBACK_MODEL)。どちらにも一致しないキーはDEFAULT_MODELを使う。
    """
    if api_key.startswith("AQ"):
        return FALLBACK_MODEL
    return DEFAULT_MODEL


class GeminiError(RuntimeError):
    pass


@dataclass
class GroundingSource:
    title: str
    uri: str


@dataclass
class GeminiReplyResult:
    text: str
    sources: list[GroundingSource] = field(default_factory=list)
    search_queries: list[str] = field(default_factory=list)


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


def _load_prompt_template() -> str:
    if not os.path.isfile(PROMPT_TEMPLATE_PATH):
        raise GeminiError(f"プロンプトテンプレートが見つかりません: {PROMPT_TEMPLATE_PATH}")
    with open(PROMPT_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return f.read()


def build_prompt(anonymized_text: str, reply_intent: str) -> str:
    template = _load_prompt_template()
    return template.format(
        reply_intent=reply_intent.strip() or "(特に指定なし。文面から適切に判断してください)",
        anonymized_text=anonymized_text,
    )


def generate_reply(
    anonymized_text: str,
    reply_intent: str,
    model: str = DEFAULT_MODEL,
) -> GeminiReplyResult:
    """匿名化後の文章と返信方針からプロンプトを組み立て、Gemini APIに送信して
    返信文の案を取得する。Google検索によるグラウンディングを有効にしているため、
    税務・会計の質問については実際にWeb検索した上で回答が組み立てられる
    (根拠を示す必要がない単純な返信では検索自体が行われないこともある)。

    gemini_api_key.txtに複数のAPIキーが書かれている場合、上から順に試し、
    レート制限(429)に達したキーは次のキーに自動で切り替えて再試行する
    (キーが1つだけの場合は、これまで通りそのキーのみで呼び出す)。

    ネットワーク呼び出しのためGUIをフリーズさせないよう、呼び出し元(GUI)は
    必ず別スレッドから呼ぶこと(local_llm.pyの各関数と同様の使い方、main.py参照)。
    APIキー/テンプレート未設定、通信/APIエラー時は GeminiError を送出する。
    """
    if not anonymized_text.strip():
        raise GeminiError("匿名化後の文章が空です。")

    api_keys = _load_api_keys()
    prompt = build_prompt(anonymized_text, reply_intent)

    try:
        from google import genai
        from google.genai import types
        from google.genai import errors as genai_errors
    except ImportError as e:
        raise GeminiError(
            "google-genai がインストールされていません。"
            " `pip install google-genai` を実行してください。"
        ) from e

    config = types.GenerateContentConfig(
        tools=[types.Tool(google_search=types.GoogleSearch())],
    )

    # modelが呼び出し元から明示的に指定されていない(=デフォルトのまま)場合のみ、
    # キーの接頭辞ごとの推奨モデル→もう一方のモデルの順で試す。明示的に指定された
    # 場合はそのモデルのみを使う。
    caller_specified_model = model != DEFAULT_MODEL

    resp = None
    last_error: Exception | None = None
    tried_models: list[str] = []
    for api_key in api_keys:
        client = genai.Client(api_key=api_key)
        if caller_specified_model:
            models_to_try = [model]
        else:
            preferred = _preferred_model_for_key(api_key)
            other = FALLBACK_MODEL if preferred == DEFAULT_MODEL else DEFAULT_MODEL
            models_to_try = [preferred, other]
        for m in models_to_try:
            if m not in tried_models:
                tried_models.append(m)
            try:
                resp = client.models.generate_content(model=m, contents=prompt, config=config)
                break
            except genai_errors.ClientError as e:
                if e.code == 404:
                    # このモデルが使えない。もう一方のモデルがあればそちらで再試行する。
                    last_error = e
                    continue
                if e.code == 429:
                    # このキーがレート制限に達した。次のキーがあればそちらで再試行する。
                    last_error = e
                    break
                raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
            except Exception as e:  # noqa: BLE001
                raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
        if resp is not None:
            break

    if resp is None:
        if isinstance(last_error, genai_errors.ClientError) and last_error.code == 404:
            message = (
                f"指定されたモデル({', '.join(tried_models)})がいずれも利用できませんでした。"
                " gemini_client.pyのDEFAULT_MODEL / FALLBACK_MODELを見直してください。"
            )
        elif len(api_keys) == 1:
            message = "Gemini APIのレート制限(429)に達しました。しばらく待ってから再試行してください。"
        else:
            message = (
                f"gemini_api_key.txtに登録されている{len(api_keys)}件のAPIキーすべてで"
                "レート制限(429)に達しました。しばらく待つか、有効なAPIキーを追加してください。"
            )
        raise GeminiError(message) from last_error

    text = getattr(resp, "text", None)
    if not text:
        raise GeminiError("Gemini APIの応答が空でした。")

    sources, search_queries = _extract_grounding(resp)
    return GeminiReplyResult(text=text, sources=sources, search_queries=search_queries)


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
