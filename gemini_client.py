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

import json
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass, field

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

DEFAULT_MODEL = "gemini-3.8-flash"

# DEFAULT_MODELが404(NOT_FOUND、特定アカウントで利用不可などモデル自体が
# 使えないケース)になった場合に自動で切り替えて再試行するモデル。
FALLBACK_MODEL = "gemini-3.6-flash"

# 判断ステップ(decide_reply_style)専用の軽量モデル。生成ステップと違い検索が
# 不要なため、速度重視でflash系固定でよい。
DECIDE_MODEL = "gemini-3.8-flash"

# gemini_api_key.txt に最初から入っている仮の値。これがそのまま残っている場合は
# 未設定とみなし、実際にはAPIを呼ばずにエラーを返す。
_PLACEHOLDER_KEY = "YOUR_API_KEY_HERE"

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
GROQ_MODEL = "openai/gpt-oss-20b"


def _is_groq_key(api_key: str) -> bool:
    """gskで始まるキーはGroq APIキーと判定する。"""
    return api_key.startswith("gsk")


def _call_groq_chat(
    api_key: str,
    messages: list[dict],
    model: str = GROQ_MODEL,
    temperature: float = 0.3,
    max_tokens: int = 2500,
) -> str:
    """Groq API(OpenAI互換)を呼び出す。外部ライブラリ依存を避けるため標準のurllibを使用。"""
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
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


def _preferred_model_for_key(api_key: str) -> str:
    """APIキーの接頭辞から推奨モデルを返す。
    基本は最新の DEFAULT_MODEL を優先する。
    """
    if _is_groq_key(api_key):
        return GROQ_MODEL
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
    # generate_reply_with_agent() 経由で呼んだ場合のみ設定される
    # (直接 generate_reply() を呼んだ場合は呼び出し元が指定したstyleがそのまま入る)。
    style: str = DEFAULT_STYLE
    style_reason: str = ""
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


def decide_reply_style(anonymized_text: str, reply_intent: str) -> tuple[str, str]:
    """匿名化後の文章と返信方針から、3種類の返信プロンプト(PROMPT_STYLES)のうち
    どれを使うべきかをLLMに判断させる。
    - gskで始まるキー: Groq API (openai/gpt-oss-20b)
    - その他のキー: Gemini API (gemini-3.8-flash)

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
    for api_key in api_keys:
        content = None
        if _is_groq_key(api_key):
            try:
                messages = [
                    {"role": "system", "content": DECIDE_STYLE_SYSTEM_PROMPT},
                    {"role": "user", "content": user_content},
                ]
                content = _call_groq_chat(
                    api_key, messages, model=GROQ_MODEL, temperature=0.1, max_tokens=1000
                )
            except urllib.error.HTTPError as e:
                last_error = e
                continue
            except Exception as e:  # noqa: BLE001
                last_error = e
                continue
        else:
            try:
                from google import genai
                from google.genai import types
                from google.genai import errors as genai_errors
            except ImportError:
                last_error = RuntimeError("google-genaiが未インストールです")
                continue

            config = types.GenerateContentConfig(
                system_instruction=DECIDE_STYLE_SYSTEM_PROMPT,
                temperature=0.1,
                max_output_tokens=200,
            )
            client = genai.Client(api_key=api_key)
            models_to_try = [DECIDE_MODEL] if DECIDE_MODEL == FALLBACK_MODEL else [DECIDE_MODEL, FALLBACK_MODEL]
            for m in models_to_try:
                try:
                    resp = client.models.generate_content(
                        model=m, contents=user_content, config=config,
                    )
                    content = getattr(resp, "text", None)
                    break
                except genai_errors.ClientError as e:
                    last_error = e
                    if e.code == 404:
                        continue
                    if e.code == 429:
                        break
                except Exception as e:  # noqa: BLE001
                    last_error = e
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
    model: str = DEFAULT_MODEL,
) -> GeminiReplyResult:
    """匿名化後の文章と返信方針から、指定されたstyle(PROMPT_STYLES参照)の
    プロンプトを組み立て、LLM APIに送信して返信文の案を取得する。
    - gskで始まるキー: Groq API (openai/gpt-oss-20b)
    - その他のキー: Gemini API (Google検索によるグラウンディングを使用)

    どのstyleを使うべきかを自動判断させたい場合は、この関数を直接呼ぶのではなく
    generate_reply_with_agent() を使うこと。

    gemini_api_key.txtに複数のAPIキーが書かれている場合、上から順に試し、
    レート制限(429)に達したキーは次のキーに自動で切り替えて再試行する。
    """
    if not anonymized_text.strip():
        raise GeminiError("匿名化後の文章が空です。")
    if style not in PROMPT_STYLES:
        style = DEFAULT_STYLE

    api_keys = _load_api_keys()
    prompt = build_prompt(anonymized_text, reply_intent, style=style)

    caller_specified_model = model != DEFAULT_MODEL

    resp = None
    chosen_model = ""
    last_error: Exception | None = None
    tried_models: list[str] = []

    for api_key in api_keys:
        if _is_groq_key(api_key):
            try:
                messages = [{"role": "user", "content": prompt}]
                text = _call_groq_chat(
                    api_key, messages, model=GROQ_MODEL, temperature=0.3, max_tokens=2500
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
                continue
            except Exception as e:  # noqa: BLE001
                last_error = e
                continue
        else:
            try:
                from google import genai
                from google.genai import types
                from google.genai import errors as genai_errors
            except ImportError as e:
                last_error = e
                continue

            config = types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
            )
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
                    chosen_model = m
                    break
                except genai_errors.ClientError as e:
                    if e.code == 404:
                        last_error = e
                        continue
                    if e.code == 429:
                        last_error = e
                        break
                    raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
                except Exception as e:  # noqa: BLE001
                    raise GeminiError(f"Gemini APIとの通信中にエラーが発生しました: {e}") from e
            if resp is not None:
                break

    if resp is None:
        if isinstance(last_error, genai_errors.ClientError) if "genai_errors" in locals() else False and getattr(last_error, "code", None) == 404:
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
    model: str = DEFAULT_MODEL,
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

    style, reason = decide_reply_style(anonymized_text, reply_intent)
    result = generate_reply(anonymized_text, reply_intent, style=style, model=model)
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
