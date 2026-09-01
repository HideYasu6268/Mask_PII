# -*- coding: utf-8 -*-
"""
local_llm.py

Ollamaのような常駐サービスを使わず、llama-cpp-python で GGUF モデルを
アプリのプロセス内に直接ロードして推論する。

- モデルは初回のみ HuggingFace からダウンロードし、以降はローカルキャッシュ
  (既定: ~/.cache/huggingface/hub)を使う。初回ダウンロード以降は完全オフライン
- モデルのロードは数秒〜十数秒かかるため、必ず別スレッドから呼び出すこと(main.py参照)
- 置換値そのものは作らせず、「見落としPIIの追加検出」役に徹する設計は
  ollama_client.py と同じ(pii_core.build_placeholder_mapping()がタグを機械的に生成)
"""

from __future__ import annotations

import json
import os
import re
import threading


DEFAULT_REPO_ID = "bartowski/Qwen2.5-7B-Instruct-GGUF"
DEFAULT_FILENAME = "Qwen2.5-7B-Instruct-Q4_K_M.gguf"

# モデルはHuggingFaceの共有キャッシュ(~/.cache/huggingface/hub)ではなく、
# 配布のしやすさを優先してプロジェクト直下の models/ フォルダに配置する。
# フォルダごとコピーすれば、コピー先PCはネットに繋がず動かせる。
APP_DIR = os.path.dirname(os.path.abspath(__file__))
MODELS_DIR = os.path.join(APP_DIR, "models")

ALLOWED_TYPES = ["PERSON", "ORG", "LOCATION", "EMAIL", "PHONE",
                  "POSTAL_CODE", "CREDIT_CARD", "MY_NUMBER", "IP_ADDRESS", "OTHER"]

SYSTEM_PROMPT = f"""あなたは日本語テキストの個人情報(PII)検出を支援するアシスタントです。
与えられた文章から、匿名化すべき情報(個人を特定しうる情報)を漏れなく抜き出してください。
既にマスキング済みの部分(例: [EMAIL_1] のような角括弧のタグ)は対象外です。

対象になりうる種別: {", ".join(ALLOWED_TYPES)}
(該当する種別が上記にない場合は "OTHER" としてください)

出力は必ず次のJSON配列形式のみとし、説明文やコードブロック記号は
一切含めないでください。文章中に実際に出現する文字列をそのまま "value" に入れてください。

[
  {{"type": "PERSON", "value": "山田太郎"}},
  {{"type": "LOCATION", "value": "大阪府"}}
]

該当が無い場合は空配列 [] を返してください。
"""

CLASSIFY_SYSTEM_PROMPT = f"""あなたは日本語テキストの個人情報(PII)分類を支援するアシスタントです。
与えられた文章を文脈として、指定された値それぞれについて、種別を次のいずれかから
1つ選んでください: {", ".join(ALLOWED_TYPES)}
(判断がつかない場合は "OTHER" としてください)

出力は必ず次のJSON配列形式のみとし、説明文やコードブロック記号は
一切含めないでください。与えられた値をそのまま "value" に入れてください。

[
  {{"value": "山田太郎", "type": "PERSON"}},
  {{"value": "大阪府", "type": "LOCATION"}}
]
"""


class LocalLLMError(RuntimeError):
    pass


# llama-cpp-pythonのLlamaインスタンスは同一オブジェクトへの並行呼び出しに対して
# スレッドセーフではない(KVキャッシュ等の内部状態を共有するため)。①②タブ問わず
# _llm_instanceは単一のシングルトンなので、推論呼び出し(create_chat_completion)は
# このロックで直列化し、ボタンの連打や複数箇所からの呼び出しが重なっても
# ネイティブクラッシュに繋がらないようにする。
_inference_lock = threading.Lock()


# モデルは一度ロードしたらプロセス内で使い回す(毎回ロードすると数秒〜十数秒無駄になるため)
_llm_instance = None
_current_config: tuple[str, str] | None = None
_load_lock = threading.Lock()


def is_model_cached(repo_id: str = DEFAULT_REPO_ID, filename: str = DEFAULT_FILENAME) -> bool:
    """指定モデルがプロジェクト内 models/ フォルダに既に存在するか確認する(UI表示用)。
    repo_id は将来的な複数モデル対応のために引数として残しているが、
    現状の判定はファイル名(ローカルパス)の存在のみを見る。
    """
    return os.path.isfile(_local_model_path(filename))


def _local_model_path(filename: str) -> str:
    return os.path.join(MODELS_DIR, filename)


def _ensure_model_file(repo_id: str, filename: str) -> str:
    """models/ フォルダにモデルファイルがあることを保証し、ローカルパスを返す。

    既に存在すればネットワークには一切アクセスしない(完全オフライン)。
    無い場合のみ HuggingFace からダウンロードして models/ 直下に配置する。
    """
    local_path = _local_model_path(filename)
    if os.path.isfile(local_path):
        return local_path

    try:
        from huggingface_hub import hf_hub_download
    except ImportError as e:
        raise LocalLLMError(
            "huggingface_hub がインストールされていません。"
            " `pip install huggingface_hub` を実行してください。"
        ) from e

    os.makedirs(MODELS_DIR, exist_ok=True)
    try:
        # local_dir を指定すると、実体ファイルがそのフォルダ直下に配置される
        # (新しめの huggingface_hub ではこれが既定動作。シンボリックリンクは作られない)
        downloaded_path = hf_hub_download(
            repo_id=repo_id,
            filename=filename,
            local_dir=MODELS_DIR,
        )
    except Exception as e:  # noqa: BLE001
        raise LocalLLMError(
            f"モデルのダウンロードに失敗しました(repo_id={repo_id}, filename={filename}): {e}"
        ) from e
    return downloaded_path


def _load_model(repo_id: str, filename: str):
    """モデルをロードする(ローカルに無ければHuggingFaceからダウンロードされる)。

    プロセス内でシングルトンとして保持し、同じ repo_id/filename なら再ロードしない。
    """
    global _llm_instance, _current_config

    with _load_lock:
        if _llm_instance is not None and _current_config == (repo_id, filename):
            return _llm_instance

        try:
            from llama_cpp import Llama
        except ImportError as e:
            raise LocalLLMError(
                "llama-cpp-python がインストールされていません。"
                " `pip install llama-cpp-python` を実行してください。"
            ) from e

        model_path = _ensure_model_file(repo_id, filename)

        try:
            n_threads = os.cpu_count() or 4
            llm = Llama(
                model_path=model_path,
                n_ctx=4096,
                n_threads=n_threads,
                verbose=False,
            )
        except Exception as e:  # noqa: BLE001
            raise LocalLLMError(
                f"モデルのロードに失敗しました(model_path={model_path}): {e}"
            ) from e

        _llm_instance = llm
        _current_config = (repo_id, filename)
        return llm


def preload_model_async(repo_id: str, filename: str, on_done, on_error):
    """GUIから呼ぶ用のヘルパー。別スレッドでダウンロード+ロードし、完了をコールバックで通知する。

    on_done()      : 成功時に引数なしで呼ばれる
    on_error(msg)   : 失敗時にエラーメッセージ文字列を渡して呼ばれる
    呼び出し元(GUI)は on_done/on_error の中で self.after(0, ...) を使って
    UIスレッドに処理を戻すこと。
    """
    def worker():
        try:
            _load_model(repo_id, filename)
            on_done()
        except LocalLLMError as e:
            on_error(str(e))

    threading.Thread(target=worker, daemon=True).start()


def _extract_json_array(raw: str) -> list[dict]:
    """LLM出力からJSON配列部分を頑健に取り出す。"""
    cleaned = raw.strip()
    cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            for v in parsed.values():
                if isinstance(v, list):
                    return v
            return []
        return parsed
    except json.JSONDecodeError:
        pass
    match = re.search(r"\[.*\]", cleaned, re.DOTALL)
    if match:
        return json.loads(match.group(0))
    raise LocalLLMError(f"LLMの応答からJSON配列を抽出できませんでした: {raw[:200]!r}")


def find_additional_pii(
    text: str,
    repo_id: str = DEFAULT_REPO_ID,
    filename: str = DEFAULT_FILENAME,
) -> list[dict]:
    """テキストからLLMにPII候補を追加検出させる。

    戻り値: [{"type": "PERSON", "value": "山田太郎"}, ...]
    モデルが未ロードの場合はこの呼び出し内でロードされる(初回は時間がかかる)。
    通信/推論エラー時は LocalLLMError を送出する。
    """
    if not text.strip():
        return []

    llm = _load_model(repo_id, filename)

    try:
        # 同じLlamaインスタンスへの同時呼び出しはスレッドセーフではなく、内部状態の
        # 破壊によるネイティブクラッシュ(Pythonの例外として捕捉できず、エラー表示
        # 無しにアプリごと落ちる)につながるため、推論呼び出し全体をロックで直列化する。
        with _inference_lock:
            resp = llm.create_chat_completion(
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": text},
                ],
                temperature=0.1,
                max_tokens=512,
            )
    except Exception as e:  # noqa: BLE001
        raise LocalLLMError(f"推論中にエラーが発生しました: {e}") from e

    content = resp["choices"][0]["message"]["content"]
    if not content:
        raise LocalLLMError("モデルの応答が空でした。")

    items = _extract_json_array(content)

    cleaned_items = []
    for it in items:
        if not isinstance(it, dict):
            continue
        value = str(it.get("value", "")).strip()
        type_ = str(it.get("type", "OTHER")).strip().upper()
        if not value:
            continue
        if type_ not in ALLOWED_TYPES:
            type_ = "OTHER"
        if value not in text:  # 原文に存在しない値(創作)は採用しない
            continue
        cleaned_items.append({"type": type_, "value": value})

    return cleaned_items


def classify_pii_types(
    text: str,
    values: list[str],
    repo_id: str = DEFAULT_REPO_ID,
    filename: str = DEFAULT_FILENAME,
) -> dict[str, str]:
    """手動入力された「元の値」の種別を、原文を文脈としてLLMに判定させる。

    対応表で「元の値」だけ入力され「種別」が未確定の行を、まとめて分類する用途。
    戻り値: {"山田太郎": "PERSON", ...}(valuesに無いキーは含まない)
    モデルが未ロードの場合はこの呼び出し内でロードされる(初回は時間がかかる)。
    通信/推論エラー時は LocalLLMError を送出する。判定不能な値は "OTHER" になる。
    """
    values = [v for v in values if v.strip()]
    if not values:
        return {}

    llm = _load_model(repo_id, filename)

    user_content = (
        f"文章:\n{text}\n\n"
        f"分類対象の値: {json.dumps(values, ensure_ascii=False)}"
    )

    try:
        # 同じLlamaインスタンスへの同時呼び出しはスレッドセーフではなく、内部状態の
        # 破壊によるネイティブクラッシュ(Pythonの例外として捕捉できず、エラー表示
        # 無しにアプリごと落ちる)につながるため、推論呼び出し全体をロックで直列化する。
        with _inference_lock:
            resp = llm.create_chat_completion(
                messages=[
                    {"role": "system", "content": CLASSIFY_SYSTEM_PROMPT},
                    {"role": "user", "content": user_content},
                ],
                temperature=0.1,
                max_tokens=512,
            )
    except Exception as e:  # noqa: BLE001
        raise LocalLLMError(f"推論中にエラーが発生しました: {e}") from e

    content = resp["choices"][0]["message"]["content"]
    if not content:
        raise LocalLLMError("モデルの応答が空でした。")

    items = _extract_json_array(content)

    result: dict[str, str] = {}
    requested = set(values)
    for it in items:
        if not isinstance(it, dict):
            continue
        value = str(it.get("value", "")).strip()
        type_ = str(it.get("type", "OTHER")).strip().upper()
        if value not in requested:
            continue
        if type_ not in ALLOWED_TYPES:
            type_ = "OTHER"
        result[value] = type_

    # LLMが判定を返さなかった値は OTHER として補完する
    for v in values:
        result.setdefault(v, "OTHER")

    return result
