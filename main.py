# -*- coding: utf-8 -*-
"""
main.py

PII匿名化デスクトップアプリ(プロトタイプ)

レイアウト(左から、各ペイン直下に関連ボタンを配置):
  [原文入力 / ①メール取得・匿名化対象を検出 / 返信の方針入力欄]
  [対応表(編集可能) / 行操作ボタン]
  [匿名化後プレビュー / 反映・AI返答ボタン]

流れ:
  1. 「①メール取得」→ Outlookの受信トレイから最新の未読メール本文を原文欄に入れる
     (または左端に原文を手動で入力/貼り付け)。その下の「どういう返信をしたいか」欄は
     「🤖 AIによる返答生成」(未実装)向けに、返信の方針を書いておく任意入力欄
  2. 「匿名化対象を検出」→ 辞書突き合わせ・正規表現/NER・ローカルLLMを一連の流れで
     まとめて実行し、対応表に行を追加する(元の値・匿名化後・種別)。非エンジニアが
     手法を意識しなくて済むよう、内部で複数の検出方式を自動で順に実行する
  3. 対応表は自由に編集可能(「+ 手動で行追加」した行は種別=OTHERでタグが仮に入る)。
     「🏷 タグを割当」→ 種別が未確定(OTHER)の行を、ローカルLLMがPERSON/ORG/
     LOCATION等に判定し、タグも判定後の種別に振り直す
  4. 「プレビュー更新」→ 対応表を原文に適用し、右端にプレビュー表示
     (対応表の新規の値はマスター辞書pii_dictionary.csvにも蓄積される)
  5. 「🤖 AIによる返答生成」は未実装のプレースホルダー(今後実装予定)

すべてローカル完結。ネットワーク通信はモデルの初回ダウンロード時のみ。
"""

from __future__ import annotations

import threading
import time
import tkinter as tk
import tkinter.messagebox as messagebox
from pathlib import Path

import customtkinter as ctk
from tksheet import Sheet

from pii_core import (
    detect_pii, merge_llm_items, build_placeholder_mapping, apply_mapping,
    reverse_mapping, append_to_master_dictionary, load_master_dictionary,
    detect_from_dictionary, PiiItem,
)
import gemini_client
import local_llm
import outlook_client


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

SHEET_HEADERS = ["元の値", "匿名化後", "種別"]

# プレビュー更新のたびに、対応表の内容(元の値・匿名化後・種別)を蓄積していく
# マスター辞書CSV。既知の値は上書きせず追記のみ(pii_core.append_to_master_dictionary参照)。
MASTER_DICTIONARY_PATH = Path(__file__).resolve().parent / "pii_dictionary.csv"

# 匿名化解除後の返信文の末尾に付ける署名。ファイルが無ければ単に付けない
# (必須ではない)。呼び出しのたびに読み直すので、アプリを再起動しなくても
# 内容の変更がすぐ反映される。
SIGNATURE_PATH = Path(__file__).resolve().parent / "署名.txt"


def _load_signature() -> str:
    if not SIGNATURE_PATH.is_file():
        return ""
    try:
        return SIGNATURE_PATH.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


class PiiAnonymizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PII匿名化ツール(ローカルLLM連携プロトタイプ)")
        self.geometry("1400x800")
        # 1400x800は1280x720などの小さい画面だと収まりきらず、右端のプレビュー
        # ペイン(weight=3で一番幅を取る設計)が画面外にはみ出てしまうため、
        # 起動時は画面サイズに関わらず最大化しておく(Windows専用アプリのため
        # state("zoomed")で問題ない)。
        self.after(0, lambda: self.state("zoomed"))

        self._build_top_bar()
        self._build_main_area()
        self._build_status_bar()

        # 現在の対応表の行を保持(sheetのデータと同期させる)
        self.rows: list[dict] = []  # [{"type":..., "original":..., "replacement":...}]

        # ローカルLLM(単一のLlamaインスタンスをシングルトンで使い回す)への
        # 呼び出しが重ならないようにするフラグ。連打や複数ボタンからの同時呼び出しは
        # ネイティブクラッシュにつながるため(_llm_call_begin/_llm_call_end参照)。
        self._llm_busy = False
        # ローカルLLM推論中の経過秒数表示用(_tick_llm_progress参照)。
        # トークン単位の進捗はllama-cpp-python側に無いため、経過時間のみ表示する。
        self._llm_progress_start: float | None = None
        self._llm_progress_after_id: str | None = None

        # Gemini API(「🤖 AIによる返答生成」)呼び出し中の連打防止フラグ。
        # ローカルLLMとは別のシングルトンロックは不要(呼び出しごとにHTTPリクエストが
        # 独立するだけ)だが、同じボタンの連打で問い合わせが重複するのを防ぐ。
        self._gemini_busy = False

        # 「①メール取得」で最後に取得したメール
        # ({"subject","sender","body","received","entry_id","store_id"})。
        # 「返信メール作成」がOutlook側でこのメールを再度特定し、標準の「返信」
        # (引用を自動生成)を作るのに使う。手動入力時など未取得の場合はNoneのままで、
        # その場合「返信メール作成」は使えない旨を案内する。
        self._fetched_mail: dict | None = None

        # 「返信メール作成」(Outlook操作)の連打防止フラグ。
        self._outlook_busy = False

    # ------------------------------------------------------------------
    # UI構築
    # ------------------------------------------------------------------
    def _build_top_bar(self):
        bar = ctk.CTkFrame(self)
        bar.pack(side="top", fill="x", padx=8, pady=(8, 4))

        ctk.CTkLabel(bar, text="Model repo:").pack(side="left", padx=(8, 2))
        self.repo_entry = ctk.CTkEntry(bar, width=280)
        self.repo_entry.insert(0, local_llm.DEFAULT_REPO_ID)
        self.repo_entry.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(bar, text="File:").pack(side="left", padx=(0, 2))
        self.filename_entry = ctk.CTkEntry(bar, width=260)
        self.filename_entry.insert(0, local_llm.DEFAULT_FILENAME)
        self.filename_entry.pack(side="left", padx=(0, 12))

        ctk.CTkButton(bar, text="モデル準備(初回はDL)", width=150,
                      command=self.on_prepare_model).pack(side="left", padx=(0, 8))

        self.use_ner_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(bar, text="NER(人名/組織名)も使う", variable=self.use_ner_var
                         ).pack(side="left", padx=(12, 0))

        # 「匿名化対象を検出」「タグ割当」でのローカルLLM推論中、経過秒数を表示する
        # (_llm_call_begin/_llm_call_end/_tick_llm_progress参照)。トークン単位の
        # 進捗はllama-cpp-python側に無いため、経過時間のみの簡易表示にとどめる。
        self.llm_progress_label = ctk.CTkLabel(bar, text="", text_color="gray")
        self.llm_progress_label.pack(side="left", padx=(12, 0))

    def _build_main_area(self):
        self.tabview = ctk.CTkTabview(self, command=self._on_tab_changed)
        self.tabview.pack(side="top", fill="both", expand=True, padx=8, pady=4)
        anonymize_tab = self.tabview.add("① 匿名化")
        deanonymize_tab = self.tabview.add("② 匿名化解除")

        self._build_anonymize_tab(anonymize_tab)
        self._build_deanonymize_tab(deanonymize_tab)

    def _on_tab_changed(self):
        # ②タブに切り替えた際、①タブの対応表の最新内容を参照用シートに反映する。
        if self.tabview.get() == "② 匿名化解除":
            self._refresh_deanon_ref_sheet()

    def _build_anonymize_tab(self, area):
        # 対応表は3列(元の値/匿名化後/種別)で情報量が少なく済むため幅を絞り、
        # その分をプレビュー(最終確認したい内容)に割り当てる。原文の比率は維持。
        # 各ペインの操作ボタンは、そのペイン自体(原文/対応表/プレビュー)の
        # 直下にそれぞれ配置する(3ペイン共通の1本のボタン行にはしない)。
        # そのためボタン行の横幅はペインの幅に収まり、対応表ペインの列幅比率
        # (weight)がボタン側の要求幅に引っ張られて崩れることもない。
        # weight(2:1:3)はminsizeで手動計算した比率と重ねると衝突して崩れることが
        # あったため、weightは使わずminsizeのみで比率を制御する(_on_anonymize_area_resize)。
        area.grid_columnconfigure(0, weight=0)  # 原文
        area.grid_columnconfigure(1, weight=0)  # 対応表
        area.grid_columnconfigure(2, weight=0)  # プレビュー
        area.grid_rowconfigure(0, weight=1)

        # weightはあくまで「最小幅を満たした後の余り」にしか効かない。原文/プレビュー
        # の各ペインはテキストボックスやラベル自体の要求幅がそれなりに大きく、
        # 小さめの画面では3ペイン合計の最小要求幅が画面幅に迫り、weight比による
        # 分配分がほとんど残らず2:1:3に見えなくなる。そこで実際のペイン全体の幅
        # (<Configure>で分かる)から比率を逆算し、各列のminsizeとして毎回強制する。
        # CTkFrameは.bind()を内部キャンバス(!ctkcanvas)へ委譲するため、
        # event.widgetはareaと一致しない。必ずareaをクロージャで直接参照すること。
        self._anonymize_area = area
        area.bind("<Configure>", self._on_anonymize_area_resize)

        # --- 左: 原文 + 検出ボタン ---
        left = ctk.CTkFrame(area)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        ctk.CTkLabel(left, text="原文(ここに入力・貼り付け)", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.source_box = ctk.CTkTextbox(left, wrap="word")
        self.source_box.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        # CTkButtonは既定でwidth=140(最小幅)を取り、sticky="ew"でも下限としては
        # 効き続ける。2個並ぶと280px超がペインの最小要求幅になり、weight比による
        # 分配(2:1:3)より最小幅の要求が上回って比率が崩れるため、明示的に狭める。
        left_btn_row = ctk.CTkFrame(left, fg_color="transparent")
        left_btn_row.pack(fill="x", padx=8, pady=(0, 8))
        left_btn_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(left_btn_row, text="①メール取得", width=1,
                      command=self.on_fetch_mail).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.btn_detect_all = ctk.CTkButton(left_btn_row, text="匿名化対象を検出", width=1,
                      command=self.on_detect_all)
        self.btn_detect_all.grid(row=0, column=1, sticky="ew", padx=(3, 0))

        # 「🤖 AIによる返答生成」(現状未実装)向けに、どんな返信をしたいかを
        # あらかじめ書いておくスペース。原文欄より小さい固定高さにとどめ、
        # 原文の入力領域を圧迫しないようにする。
        ctk.CTkLabel(left, text="どういう返信をしたいか(任意、AIによる返答生成で使用)",
                     font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(0, 2))
        self.reply_intent_box = ctk.CTkTextbox(left, wrap="word", height=90)
        self.reply_intent_box.pack(fill="x", padx=8, pady=(0, 8))

        # --- 中央: 対応表 + 編集ボタン ---
        # tksheetはセルの実内容(長いメールアドレス等)に合わせて列幅を自動拡張する
        # ため、幅を指定しないとこのペインの要求幅がweightの比率(2:1:3)を超えて
        # 広がり、対応表が一番大きく表示されてしまう。②タブの参照用対応表と同じく
        # 幅を固定してpack_propagate(False)で強制し、はみ出た分はシート内蔵の
        # 横スクロールで見る形にする。
        # CTkFrameのwidth=はCTk内部のDPIスケーリングを受けてしまい、grid側の
        # minsize(スケーリングされない生のTk値)と食い違うため、widthは指定せず
        # pack_propagate(False)のみで「子要素に引っ張られて広がる」ことだけ防ぎ、
        # 実際のサイズは_on_anonymize_area_resizeのminsize(mid_width)に一本化する。
        mid = ctk.CTkFrame(area)
        mid.grid(row=0, column=1, sticky="nsew", padx=4)
        mid.pack_propagate(False)
        ctk.CTkLabel(mid, text="対応表(編集可能)", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))

        sheet_frame = ctk.CTkFrame(mid, fg_color="transparent")
        sheet_frame.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.sheet = Sheet(
            sheet_frame,
            headers=SHEET_HEADERS,
            data=[],
            height=500,
        )
        self.sheet.enable_bindings(
            "single_select", "row_select", "column_select",
            "arrowkeys", "edit_cell", "delete_key", "copy", "paste",
            "right_click_popup_menu",
        )
        self.sheet.pack(fill="both", expand=True)

        # CTkButtonは既定でwidth=140(最小幅)を取り、対応表ペインの幅が
        # weightの比率(1/6)より広がってしまう主因になるため、gridの等分割
        # (sticky="ew"、widthは指定しない)で枠の幅ぴったりに均等配置する。
        mid_btn_row = ctk.CTkFrame(mid, fg_color="transparent")
        mid_btn_row.pack(fill="x", padx=8, pady=(0, 8))
        mid_btn_row.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ctk.CTkButton(mid_btn_row, text="手動追加",
                      command=self.on_add_row).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.btn_llm_classify = ctk.CTkButton(mid_btn_row, text="タグ割当",
                      command=self.on_llm_classify)
        self.btn_llm_classify.grid(row=0, column=1, sticky="ew", padx=3)
        ctk.CTkButton(mid_btn_row, text="手動削除", fg_color="#a33",
                      hover_color="#822", command=self.on_delete_row
                      ).grid(row=0, column=2, sticky="ew", padx=3)
        ctk.CTkButton(mid_btn_row, text="出現順に並替",
                      command=self.on_sort_rows).grid(row=0, column=3, sticky="ew", padx=(3, 0))

        # --- 右: 匿名化後プレビュー + プレビュー/AI返答ボタン ---
        right = ctk.CTkFrame(area)
        right.grid(row=0, column=2, sticky="nsew", padx=(4, 0))
        ctk.CTkLabel(right, text="匿名化後プレビュー", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.preview_box = ctk.CTkTextbox(right, wrap="word")
        self.preview_box.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.preview_box.configure(state="disabled")

        right_btn_row = ctk.CTkFrame(right, fg_color="transparent")
        right_btn_row.pack(fill="x", padx=8, pady=(0, 8))
        right_btn_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(right_btn_row, text="プレビュー更新", width=1,
                      command=self.on_apply_mapping).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.btn_generate_reply = ctk.CTkButton(
            right_btn_row, text="🤖 AIによる返答生成", width=1,
            command=self.on_generate_reply)
        self.btn_generate_reply.grid(row=0, column=1, sticky="ew", padx=(3, 0))

    def _on_anonymize_area_resize(self, event):
        # CTkFrameは.bind()を内部キャンバスへ委譲するためevent.widget/event.width
        # は信用できない。必ずareaそのものから実測する。
        area = self._anonymize_area
        total = area.winfo_width()
        if total <= 1:
            return
        mid_width = 440
        remaining = max(total - mid_width, 0)
        left_width = remaining * 2 // 5
        right_width = remaining - left_width  # 2:3のうち残り側(3)をこちらに寄せる
        area.grid_columnconfigure(0, minsize=left_width)
        area.grid_columnconfigure(1, minsize=mid_width)
        area.grid_columnconfigure(2, minsize=right_width)

    def _build_deanonymize_tab(self, area):
        # レイアウトの向きは①タブと揃える(左=匿名化前/解除後の平文、右=匿名化語)。
        # 対応表は参照用で情報量が少ないため、①タブ以上にテキストエリア側へ幅を寄せる。
        area.grid_columnconfigure(0, weight=3)  # 解除後プレビュー(平文)
        area.grid_columnconfigure(1, weight=1)  # 対応表(参照用)
        area.grid_columnconfigure(2, weight=3)  # AIによる返信案(匿名化)

        ctk.CTkLabel(
            area,
            text="①タブで「🤖 AIによる返答生成」を実行すると、生成された返信案がこのタブの"
                 "右側に自動で反映され、対応表(①タブと連動・参照用)を使って左側に元の値へ"
                 "解除したプレビューが自動表示されます。税務・会計の質問でGoogle検索による"
                 "根拠確認(グラウンディング)が行われた場合、その根拠情報も右下に表示されます。",
            anchor="w", justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="ew", padx=4, pady=(0, 4))

        # --- 左: 解除後プレビュー(平文) ---
        left = ctk.CTkFrame(area)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        ctk.CTkLabel(left, text="匿名化解除後プレビュー", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.deanon_preview_box = ctk.CTkTextbox(left, wrap="word")
        self.deanon_preview_box.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.deanon_preview_box.configure(state="disabled")

        deanon_btn_row = ctk.CTkFrame(left, fg_color="transparent")
        deanon_btn_row.pack(fill="x", padx=8, pady=(0, 8))
        deanon_btn_row.grid_columnconfigure((0, 1), weight=1)
        self.btn_create_reply_mail = ctk.CTkButton(deanon_btn_row, text="返信メール作成", width=1,
                      command=self.on_create_reply_mail)
        self.btn_create_reply_mail.grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(deanon_btn_row, text="削除", fg_color="#a33", hover_color="#822", width=1,
                      command=self.on_clear_all).grid(row=0, column=1, sticky="ew", padx=(3, 0))

        # --- 中央: 対応表(①タブの内容を表示するだけの参照用、念のため) ---
        # 匿名化解除は①タブの「🤖 AIによる返答生成」から自動実行される想定のため、
        # ここに操作ボタンは置かない。
        # 中身(表)の自然な要求幅は無視し、幅300pxで固定する。
        # (grid列のweight比による配分は、中身の要求幅がそれより大きいと
        # 無効化されてしまうため、pack_propagate(False)で強制的に絞る)
        mid = ctk.CTkFrame(area, width=160)
        mid.grid(row=1, column=1, sticky="nsew", padx=4)
        mid.pack_propagate(False)
        ctk.CTkLabel(mid, text="対応表(参照用)", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        ref_sheet_frame = ctk.CTkFrame(mid, fg_color="transparent")
        ref_sheet_frame.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        self.deanon_ref_sheet = Sheet(
            ref_sheet_frame,
            headers=SHEET_HEADERS,
            data=[],
            height=500,
            # 参照専用で行番号は不要。既定列幅も狭める(はみ出た分は
            # シート内蔵の横スクロールで見られる)。
            show_row_index=False,
            default_column_width=90,
        )
        # 編集は①タブの対応表でのみ行う想定のため、閲覧・コピーのみ許可。
        self.deanon_ref_sheet.enable_bindings(
            "single_select", "row_select", "column_select", "arrowkeys", "copy",
        )
        self.deanon_ref_sheet.pack(fill="both", expand=True)

        # --- 右: AIによる返信案(匿名化)。手動で他の匿名化済みテキストを貼り付けてもよい ---
        right = ctk.CTkFrame(area)
        right.grid(row=1, column=2, sticky="nsew", padx=(4, 0))
        ctk.CTkLabel(right, text="AIによる返信案(匿名化)", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.deanon_input_box = ctk.CTkTextbox(right, wrap="word")
        self.deanon_input_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # AIによる返答生成時、Google検索によるグラウンディングで実際に参照した
        # Webページ(根拠)があれば表示する。手動で貼り付けたテキストの匿名化解除
        # では根拠情報が無いため、その旨を表示する(_on_generate_reply_done参照)。
        ctk.CTkLabel(right, text="AIが参照した根拠情報(Google検索によるグラウンディング)",
                     font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(0, 2))
        self.deanon_sources_box = ctk.CTkTextbox(right, wrap="word", height=140)
        self.deanon_sources_box.pack(fill="x", padx=8, pady=(0, 8))
        self._set_deanon_sources_text("(まだAIによる返答生成は実行されていません)")

        area.grid_rowconfigure(1, weight=1)

    def _build_status_bar(self):
        self.status_label = ctk.CTkLabel(self, text="準備完了", anchor="w")
        self.status_label.pack(side="bottom", fill="x", padx=12, pady=(0, 6))

    def set_status(self, text: str):
        self.status_label.configure(text=text)

    # ------------------------------------------------------------------
    # ローカルLLM呼び出しの排他制御
    # ------------------------------------------------------------------
    def _llm_call_begin(self) -> bool:
        """ローカルLLM呼び出し(推論)の開始前に必ず呼ぶ。既に呼び出し中ならFalseを
        返し、案内を出して呼び出し元は処理を中断すること。

        「匿名化対象を検出」と「タグ割当」は内部で同じLlamaシングルトンを共有して
        いるが、ボタンの連打や別々のボタンからの同時クリックで並行に推論を開始
        すると、Pythonの例外として捕捉できないネイティブクラッシュ(アプリが無言で
        落ちる)につながる。ボタンを処理中はdisabledにして視覚的にも防止する。
        """
        if self._llm_busy:
            messagebox.showinfo(
                "確認",
                "既にローカルLLMで処理中です。完了するまでお待ちください。",
            )
            return False
        self._llm_busy = True
        self.btn_detect_all.configure(state="disabled")
        self.btn_llm_classify.configure(state="disabled")
        self._llm_progress_start = time.monotonic()
        self._tick_llm_progress()
        return True

    def _llm_call_end(self):
        self._llm_busy = False
        self.btn_detect_all.configure(state="normal")
        self.btn_llm_classify.configure(state="normal")
        self._llm_progress_start = None
        if self._llm_progress_after_id is not None:
            self.after_cancel(self._llm_progress_after_id)
            self._llm_progress_after_id = None
        self.llm_progress_label.configure(text="")

    def _tick_llm_progress(self):
        """ローカルLLM推論中、1秒おきに経過秒数を表示し続ける。トークン単位の
        進捗はllama-cpp-python側に無いため、あくまで「動いている」ことが
        分かる程度の簡易表示(正確な完了見込み時間ではない)。
        """
        if not self._llm_busy or self._llm_progress_start is None:
            self.llm_progress_label.configure(text="")
            self._llm_progress_after_id = None
            return
        elapsed = int(time.monotonic() - self._llm_progress_start)
        self.llm_progress_label.configure(text=f"🧠 ローカルLLM推論中...({elapsed}秒経過)")
        self._llm_progress_after_id = self.after(1000, self._tick_llm_progress)

    # ------------------------------------------------------------------
    # sheet <-> rows 同期ヘルパー
    # ------------------------------------------------------------------
    def _refresh_sheet_from_rows(self):
        data = [[r["original"], r["replacement"], r["type"]] for r in self.rows]
        self.sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.sheet.set_all_column_widths()

    def _pull_rows_from_sheet(self) -> list[dict]:
        """現在sheetに表示されている内容(手動編集も含む)を rows 形式で取得する。"""
        data = self.sheet.get_sheet_data()
        rows = []
        for r in data:
            if len(r) < 3:
                continue
            original, replacement, type_ = (r + ["", "", ""])[:3]
            original = str(original).strip()
            if not original:
                continue
            rows.append({
                "type": str(type_).strip() or "OTHER",
                "original": original,
                "replacement": str(replacement).strip(),
            })
        return rows

    def _existing_originals(self) -> set[str]:
        return {r["original"] for r in self._pull_rows_from_sheet()}

    def _sort_rows_by_appearance(self, rows: list[dict], text: str) -> list[dict]:
        """対応表の行を、原文中で最初に出現する位置の順に並べ替える。

        原文中に見つからない行(手動追加でまだ値を書きかけ、など)は
        末尾にまとめ、その中では元の順序を保つ(安定ソート)。
        """
        def sort_key(row: dict) -> int:
            pos = text.find(row["original"]) if row["original"] else -1
            return pos if pos != -1 else 10**9

        return sorted(rows, key=sort_key)

    def _type_counters(self, rows: list[dict]) -> dict[str, int]:
        """種別ごとの現在の最大連番を把握する(既存タグ [TYPE_n] から拾う)。"""
        import re as _re
        counters: dict[str, int] = {}
        for r in rows:
            m = _re.match(r"^\[([A-Z_]+)_(\d+)\]$", r["replacement"])
            if m:
                t, n = m.group(1), int(m.group(2))
                counters[t] = max(counters.get(t, 0), n)
        return counters

    def _fill_blank_replacements(self, rows: list[dict]) -> int:
        """「元の値」が入力済みで「匿名化後」が空欄の行に、機械的にタグを振る。
        (手動で元の値だけ入力し、匿名化後を書き忘れているケースの保険)
        戻り値: 実際に補完した行数。
        """
        counters = self._type_counters(rows)
        tag_by_value: dict[str, str] = {
            r["original"]: r["replacement"] for r in rows if r["original"] and r["replacement"]
        }
        filled = 0
        for r in rows:
            if not r["original"] or r["replacement"]:
                continue
            if r["original"] in tag_by_value:
                r["replacement"] = tag_by_value[r["original"]]
                continue
            type_ = r["type"] or "OTHER"
            r["type"] = type_
            counters[type_] = counters.get(type_, 0) + 1
            tag = f"[{type_}_{counters[type_]}]"
            r["replacement"] = tag
            tag_by_value[r["original"]] = tag
            filled += 1
        return filled

    def _append_items_as_rows(self, items: list[PiiItem]) -> int:
        """既存行と重複しないPiiItemだけを、タグを振って対応表に追加する。
        追加後、対応表全体を原文の出現順に並べ替える。
        戻り値: 実際に新規追加された行数(重複除外後)。
        """
        text = self.source_box.get("1.0", "end-1c")
        current_rows = self._pull_rows_from_sheet()
        self._fill_blank_replacements(current_rows)
        existing_values = {r["original"] for r in current_rows}
        counters = self._type_counters(current_rows)

        new_items = [it for it in items if it.value not in existing_values]
        # 同一valueは1行にまとめる
        seen = set()
        deduped = []
        for it in new_items:
            if it.value in seen:
                continue
            seen.add(it.value)
            deduped.append(it)

        for it in deduped:
            counters[it.type] = counters.get(it.type, 0) + 1
            tag = f"[{it.type}_{counters[it.type]}]"
            current_rows.append({"type": it.type, "original": it.value, "replacement": tag})

        self.rows = self._sort_rows_by_appearance(current_rows, text)
        self._refresh_sheet_from_rows()
        return len(deduped)

    # ------------------------------------------------------------------
    # イベントハンドラ
    # ------------------------------------------------------------------
    def on_fetch_mail(self):
        """「①メール取得」: Outlookの受信トレイから最新の未読メール本文を原文欄に入れる。"""
        self.set_status("Outlookから未読メールを取得中...")

        def on_done(mail):
            self.after(0, lambda: self._on_fetch_mail_done(mail))

        def on_error(msg):
            self.after(0, lambda: self._on_fetch_mail_error(msg))

        outlook_client.fetch_latest_unread_email_async(on_done, on_error)

    def _on_fetch_mail_done(self, mail: dict | None):
        if mail is None:
            self.set_status("未読メールが見つかりませんでした。")
            messagebox.showinfo("確認", "未読メールが見つかりませんでした。")
            return
        self.source_box.delete("1.0", "end")
        self.source_box.insert("1.0", mail["body"])
        self._fetched_mail = mail
        self.set_status(f"未読メールを取得しました(件名: {mail['subject']} / 差出人: {mail['sender']})。")

    def _on_fetch_mail_error(self, msg: str):
        self.set_status(f"メール取得エラー: {msg}")
        messagebox.showwarning("メール取得エラー", msg)

    def on_detect_all(self):
        """「匿名化対象を検出」: 辞書突き合わせ・正規表現/NER・ローカルLLMを、非エンジニアが
        使い分けを意識しなくて済むよう一連の流れとしてまとめて実行する。
        辞書/正規表現/NERは同期処理で速いため先に反映し、続けてローカルLLMによる
        追加検出を別スレッドで実行する(モデルロード込みで数十秒かかりうるため)。
        """
        text = self.source_box.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo("確認", "原文が空です。")
            return

        try:
            regex_ner_items = detect_pii(text, use_ner=self.use_ner_var.get())
        except Exception as e:  # noqa: BLE001
            # 正規表現/NER検出は同期処理(GUIスレッド上)なので、ここで想定外の
            # 例外が起きてもアプリごと落とさず、エラー内容を表示して継続する。
            self.set_status(f"検出中にエラーが発生しました: {e}")
            messagebox.showwarning("検出エラー", str(e))
            return

        # 過去に人手で確認済みのマスター辞書も、毎回の検出のたびに必ず突き合わせる
        # (これが辞書を蓄積している目的そのものなので、自動で行う)。
        dict_entries = load_master_dictionary(MASTER_DICTIONARY_PATH)
        dict_items = detect_from_dictionary(text, dict_entries) if dict_entries else []

        added_sync = self._append_items_as_rows(dict_items + regex_ner_items)

        if not self._llm_call_begin():
            self.set_status(
                f"辞書一致{len(dict_items)}件・正規表現/NER {len(regex_ner_items)}件のうち"
                f"{added_sync}件を追加。(ローカルLLMは他の処理が完了するまでスキップしました)"
            )
            return

        repo_id = self.repo_entry.get().strip() or local_llm.DEFAULT_REPO_ID
        filename = self.filename_entry.get().strip() or local_llm.DEFAULT_FILENAME
        self.set_status(
            f"辞書一致{len(dict_items)}件・正規表現/NER {len(regex_ner_items)}件のうち"
            f"{added_sync}件を追加。続けてローカルLLMで見落としを推論中..."
            "(初回はモデルロードも含め数十秒かかります)"
        )

        def worker():
            try:
                found = local_llm.find_additional_pii(text, repo_id=repo_id, filename=filename)
                error = None
            except local_llm.LocalLLMError as e:
                found, error = [], str(e)
            self.after(0, lambda: (self._llm_call_end(), self._on_detect_all_llm_done(found, error, added_sync)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_detect_all_llm_done(self, found: list[dict], error: str | None, added_sync: int):
        if error:
            self.set_status(
                f"辞書/正規表現/NERで{added_sync}件を追加。ローカルLLM検出でエラー: {error}"
            )
            messagebox.showwarning("ローカルLLM検出エラー", error)
            return

        items = [PiiItem(type=f["type"], value=f["value"], start=0, end=0) for f in found]
        added_llm = self._append_items_as_rows(items)
        skipped_llm = len(found) - added_llm

        msg = f"匿名化対象を検出: 辞書/正規表現/NERで{added_sync}件、ローカルLLMで{added_llm}件を対応表に追加"
        if skipped_llm > 0:
            msg += f"(LLM検出{len(found)}件中{skipped_llm}件は対応表に既に存在したためスキップ)"
        self.set_status(msg + "。")

    def on_llm_classify(self):
        """対応表のうち、元の値はあるが種別がまだ既定値"OTHER"のまま(=手動追加直後など
        未確定)の行を対象に、原文を文脈としてLLMに種別を判定させ、タグも判定後の
        種別に合わせて振り直す(対応表ペイン直下の「🏷 種別を判定してタグを割当」用)。
        """
        text = self.source_box.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo("確認", "原文が空です。")
            return
        current_rows = self._pull_rows_from_sheet()
        pending_values = [r["original"] for r in current_rows if r["original"] and r["type"] == "OTHER"]
        if not pending_values:
            messagebox.showinfo("確認", "種別が未確定(OTHER)の行がありません。")
            return

        if not self._llm_call_begin():
            return

        repo_id = self.repo_entry.get().strip() or local_llm.DEFAULT_REPO_ID
        filename = self.filename_entry.get().strip() or local_llm.DEFAULT_FILENAME

        self.set_status("ローカルLLMで種別を判定中... (初回はモデルロードも含め数十秒かかります)")

        def worker():
            try:
                classified = local_llm.classify_pii_types(
                    text, pending_values, repo_id=repo_id, filename=filename
                )
                error = None
            except local_llm.LocalLLMError as e:
                classified, error = {}, str(e)
            self.after(0, lambda: (self._llm_call_end(), self._on_llm_classify_done(classified, error)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_llm_classify_done(self, classified: dict[str, str], error: str | None):
        if error:
            self.set_status(f"種別判定エラー: {error}")
            messagebox.showwarning("種別判定エラー", error)
            return

        current_rows = self._pull_rows_from_sheet()
        counters = self._type_counters(current_rows)
        classified_count = 0
        for r in current_rows:
            if r["original"] not in classified or r["type"] != "OTHER":
                continue
            new_type = classified[r["original"]]
            if new_type != r["type"]:
                # 種別が変わった分、既に振られていた[OTHER_n]タグは
                # 新しい種別のタグに振り直す(古い番号は欠番のままでよい)。
                r["type"] = new_type
                counters[new_type] = counters.get(new_type, 0) + 1
                r["replacement"] = f"[{new_type}_{counters[new_type]}]"
            classified_count += 1
        self.rows = current_rows
        self._refresh_sheet_from_rows()
        self.set_status(f"種別判定: {classified_count}件のタグを更新しました。")

    def on_add_row(self):
        rows = self._pull_rows_from_sheet()
        type_ = "OTHER"
        counters = self._type_counters(rows)
        counters[type_] = counters.get(type_, 0) + 1
        tag = f"[{type_}_{counters[type_]}]"
        rows.append({"type": type_, "original": "", "replacement": tag})
        self.rows = rows
        self._refresh_sheet_from_rows()

    def on_sort_rows(self):
        text = self.source_box.get("1.0", "end-1c")
        rows = self._pull_rows_from_sheet()
        self.rows = self._sort_rows_by_appearance(rows, text)
        self._refresh_sheet_from_rows()
        self.set_status("対応表を原文の出現順に並べ替えました。")

    def on_delete_row(self):
        selected = self.sheet.get_selected_rows()
        if not selected:
            messagebox.showinfo("確認", "削除する行を選択してください。")
            return
        rows = self._pull_rows_from_sheet()
        rows = [r for i, r in enumerate(rows) if i not in selected]
        self.rows = rows
        self._refresh_sheet_from_rows()

    def on_apply_mapping(self):
        text = self.source_box.get("1.0", "end-1c")
        rows = self._pull_rows_from_sheet()
        filled = self._fill_blank_replacements(rows)
        if filled > 0:
            self.rows = self._sort_rows_by_appearance(rows, text)
            self._refresh_sheet_from_rows()

        mapping = {r["original"]: r["replacement"] for r in rows if r["replacement"]}
        result = apply_mapping(text, mapping)

        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", result)
        self.preview_box.configure(state="disabled")

        added_to_dict = 0
        dict_skipped: list[str] = []
        try:
            added_to_dict, dict_skipped = append_to_master_dictionary(rows, MASTER_DICTIONARY_PATH)
        except OSError as e:
            messagebox.showwarning(
                "辞書保存エラー",
                f"マスター辞書CSVへの保存に失敗しました: {e}",
            )

        status = f"対応表({len(mapping)}件)を適用してプレビューを更新しました。"
        if filled > 0:
            status += f" (匿名化後が空欄だった{filled}件にタグを自動補完しました。)"
        if added_to_dict > 0:
            status += f" マスター辞書に{added_to_dict}件を追記しました。"
        if dict_skipped:
            status += f" (文字コードの都合で{len(dict_skipped)}件は辞書に保存できませんでした。)"
        self.set_status(status)

    def on_generate_reply(self):
        """「🤖 AIによる返答生成」: 匿名化後プレビューの文章 + 「どういう返信をしたいか」欄の
        内容 + プロンプトテンプレート(reply_prompt_template.txt)を組み立て、Geminiに
        送信して返信文の案を生成する。外部に送るのはあくまで匿名化後の文章のみ。
        """
        anonymized_text = self.preview_box.get("1.0", "end-1c")
        if not anonymized_text.strip():
            messagebox.showinfo(
                "確認",
                "匿名化後プレビューが空です。先に「プレビュー更新」を押してください。",
            )
            return

        if self._gemini_busy:
            messagebox.showinfo("確認", "既にGemini APIに問い合わせ中です。完了するまでお待ちください。")
            return

        reply_intent = self.reply_intent_box.get("1.0", "end-1c")

        self._gemini_busy = True
        self.btn_generate_reply.configure(state="disabled")
        self.set_status("Gemini APIに問い合わせ中...")

        def worker():
            try:
                result = gemini_client.generate_reply(anonymized_text, reply_intent)
                error = None
            except gemini_client.GeminiError as e:
                result, error = None, str(e)
            self.after(0, lambda: self._on_generate_reply_done(result, error))

        threading.Thread(target=worker, daemon=True).start()

    def _on_generate_reply_done(
        self, result: "gemini_client.GeminiReplyResult | None", error: str | None
    ):
        self._gemini_busy = False
        self.btn_generate_reply.configure(state="normal")

        if error:
            self.set_status(f"AIによる返答生成エラー: {error}")
            messagebox.showwarning("AIによる返答生成エラー", error)
            return

        # 生成結果(タグ付きのまま)を②タブの「AIによる返信案(匿名化)」欄に自動反映し、
        # そのまま匿名化解除まで自動実行する(ユーザーがボタンを押す手間を省く)。
        # on_deanonymize側が解除結果の詳細なステータスを表示するので、ここでは
        # 上書きせずタブ切替・反映のみ行う。
        self.tabview.set("② 匿名化解除")
        self.deanon_input_box.delete("1.0", "end")
        self.deanon_input_box.insert("1.0", result.text)
        self._update_deanon_sources(result.sources, result.search_queries)
        self.on_deanonymize()

    def _set_deanon_sources_text(self, text: str):
        self.deanon_sources_box.configure(state="normal")
        self.deanon_sources_box.delete("1.0", "end")
        self.deanon_sources_box.insert("1.0", text)
        self.deanon_sources_box.configure(state="disabled")

    def _update_deanon_sources(self, sources: list, search_queries: list[str]):
        """AIによる返答生成の結果、Google検索によるグラウンディングで実際に参照した
        Webページ(根拠)があれば一覧表示する。検索が行われなかった場合はその旨を表示する。
        """
        if not sources:
            self._set_deanon_sources_text(
                "(この返信案の作成にあたり、Web検索による根拠確認は行われませんでした)"
            )
            return

        lines = []
        if search_queries:
            lines.append("検索クエリ: " + " / ".join(search_queries))
            lines.append("")
        lines.append("参照した情報源:")
        for i, s in enumerate(sources, start=1):
            lines.append(f"{i}. {s.title}\n   {s.uri}")
        self._set_deanon_sources_text("\n".join(lines))

    def _refresh_deanon_ref_sheet(self):
        rows = self._pull_rows_from_sheet()
        data = [[r["original"], r["replacement"], r["type"]] for r in rows]
        self.deanon_ref_sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.deanon_ref_sheet.set_all_column_widths()

    def on_deanonymize(self) -> bool:
        """戻り値: 匿名化解除後プレビューを実際に更新できたかどうか
        (on_create_reply_mailが、この後に引用を追記してよいか判断するのに使う)。
        """
        self._refresh_deanon_ref_sheet()
        encoded_text = self.deanon_input_box.get("1.0", "end-1c")
        if not encoded_text.strip():
            messagebox.showinfo("確認", "「AIによる返信案(匿名化)」欄が空です。")
            return False

        rows = self._pull_rows_from_sheet()
        mapping = {r["original"]: r["replacement"] for r in rows if r["replacement"]}
        if not mapping:
            messagebox.showinfo("確認", "対応表が空です。①タブで対応表を作成してから実行してください。")
            return False

        rev_mapping = reverse_mapping(mapping)
        result = apply_mapping(encoded_text, rev_mapping)

        signature = _load_signature()
        if signature:
            result = f"{result}\n\n{signature}"

        self.deanon_preview_box.configure(state="normal")
        self.deanon_preview_box.delete("1.0", "end")
        self.deanon_preview_box.insert("1.0", result)
        self.deanon_preview_box.configure(state="disabled")

        # 入力中に見つからなかったタグ(対応表に無い/既に手動編集された等)を検知して警告
        unresolved = [tag for tag in rev_mapping if tag not in encoded_text]
        found_tags = sum(1 for tag in rev_mapping if tag in encoded_text)
        self.set_status(f"対応表({len(rev_mapping)}件)を使って匿名化を解除しました(一致したタグ: {found_tags}件)。")
        if unresolved and found_tags == 0:
            self.set_status(
                f"対応表({len(rev_mapping)}件)のタグが入力テキスト中に見つかりませんでした。"
                "対応表とタグ付きテキストの組み合わせを確認してください。"
            )
        return True

    def on_create_reply_mail(self):
        """「返信メール作成」: 匿名化解除後プレビューを最新の対応表・返信案で作り直した
        うえで、Outlook標準の「返信」(引用を自動生成)を使い、その本文の先頭に
        この返信案を差し込んだ状態でOutlookの作成画面を開く。
        ①メール取得でOutlookから取得したメールに対してのみ実行できる
        (元のメールをOutlook側で再度特定する必要があるため)。
        送信は行わない。内容の確認・編集・送信はOutlook上でユーザー自身が行う。
        """
        if not self.on_deanonymize():
            return

        mail = self._fetched_mail
        if not mail or not mail.get("entry_id"):
            messagebox.showinfo(
                "確認",
                "①タブの「①メール取得」でOutlookから取得したメールに対してのみ、"
                "引用返信を作成できます。",
            )
            return

        if self._outlook_busy:
            messagebox.showinfo("確認", "既にOutlookで返信メールを作成中です。完了するまでお待ちください。")
            return

        reply_body = self.deanon_preview_box.get("1.0", "end-1c")

        self._outlook_busy = True
        self.btn_create_reply_mail.configure(state="disabled")
        self.set_status("Outlookで引用返信を作成中...")

        def on_done():
            self.after(0, self._on_create_reply_mail_done)

        def on_error(msg):
            self.after(0, lambda: self._on_create_reply_mail_error(msg))

        outlook_client.create_quoted_reply_async(
            mail["entry_id"], mail["store_id"], reply_body, on_done, on_error
        )

    def _on_create_reply_mail_done(self):
        self._outlook_busy = False
        self.btn_create_reply_mail.configure(state="normal")
        self.set_status("Outlookで引用返信を作成しました(内容を確認のうえ送信してください)。")

    def _on_create_reply_mail_error(self, msg: str):
        self._outlook_busy = False
        self.btn_create_reply_mail.configure(state="normal")
        self.set_status(f"引用返信の作成エラー: {msg}")
        messagebox.showwarning("引用返信の作成エラー", msg)

    def on_clear_all(self):
        """「削除」: 対応表を含め、①②タブのテキストエリアをすべて空の状態に戻す
        (1件のメール対応が終わり、次のメールに取り掛かる前のリセット用)。
        """
        self.source_box.delete("1.0", "end")
        self.reply_intent_box.delete("1.0", "end")

        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.configure(state="disabled")

        self.rows = []
        self._refresh_sheet_from_rows()
        self._refresh_deanon_ref_sheet()

        self.deanon_input_box.delete("1.0", "end")

        self.deanon_preview_box.configure(state="normal")
        self.deanon_preview_box.delete("1.0", "end")
        self.deanon_preview_box.configure(state="disabled")

        self._set_deanon_sources_text("(まだAIによる返答生成は実行されていません)")

        self._fetched_mail = None
        self.set_status("すべてのテキストエリアと対応表をクリアしました。")

    def on_prepare_model(self):
        repo_id = self.repo_entry.get().strip() or local_llm.DEFAULT_REPO_ID
        filename = self.filename_entry.get().strip() or local_llm.DEFAULT_FILENAME

        if local_llm.is_model_cached(repo_id, filename):
            self.set_status(f"モデルはプロジェクト内 models/ に配置済みです。ロードします... ({filename})")
        else:
            self.set_status(f"モデルをHuggingFaceからダウンロードし、models/ に配置中... ({filename}、数分かかる場合があります)")

        def handle_done():
            self.set_status(f"モデル準備完了: {repo_id} / {filename}")

        def handle_error(msg):
            self.set_status(f"モデル準備エラー: {msg}")
            messagebox.showwarning("モデル準備エラー", msg)

        def on_done():
            self.after(0, handle_done)

        def on_error(msg):
            self.after(0, lambda: handle_error(msg))

        local_llm.preload_model_async(repo_id, filename, on_done, on_error)


if __name__ == "__main__":
    app = PiiAnonymizerApp()
    app.mainloop()