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
import local_llm
import outlook_client


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

SHEET_HEADERS = ["元の値", "匿名化後", "種別"]

# プレビュー更新のたびに、対応表の内容(元の値・匿名化後・種別)を蓄積していく
# マスター辞書CSV。既知の値は上書きせず追記のみ(pii_core.append_to_master_dictionary参照)。
MASTER_DICTIONARY_PATH = Path(__file__).resolve().parent / "pii_dictionary.csv"


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
        ctk.CTkButton(right_btn_row, text="🤖 AIによる返答生成", width=1, state="disabled"
                      ).grid(row=0, column=1, sticky="ew", padx=(3, 0))

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
        area.grid_columnconfigure(2, weight=3)  # 匿名化されたテキスト入力

        ctk.CTkLabel(
            area,
            text="右側に匿名化語(タグ付き)のテキストを貼り付け、対応表(①タブと連動・参照用)を使って"
                 "左側に元の値へ解除したプレビューを表示します。"
                 "(例: 匿名化したテキストを外部LLMに渡し、返ってきた結果をここに貼り付ける)",
            anchor="w", justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="ew", padx=4, pady=(0, 4))

        # --- 左: 解除後プレビュー(平文) ---
        left = ctk.CTkFrame(area)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        ctk.CTkLabel(left, text="匿名化解除後プレビュー", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.deanon_preview_box = ctk.CTkTextbox(left, wrap="word")
        self.deanon_preview_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.deanon_preview_box.configure(state="disabled")

        # --- 中央: 対応表(①タブの内容を表示するだけの参照用、念のため) ---
        # 操作ボタンはここ(対応表の直下)に置く。左右のテキストエリアの列とは
        # 別の列に留めているため、テキストエリアにボタンが被さることはない。
        # 中身(表・ボタン)の自然な要求幅は無視し、幅300pxで固定する。
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
        ctk.CTkButton(mid, text="🔓 匿名化を解除",
                      command=self.on_deanonymize).pack(fill="x", padx=8, pady=(0, 8))

        # --- 右: 匿名化されたテキストの入力 ---
        right = ctk.CTkFrame(area)
        right.grid(row=1, column=2, sticky="nsew", padx=(4, 0))
        ctk.CTkLabel(right, text="匿名化されたテキスト(貼り付け)", font=ctk.CTkFont(weight="bold")
                     ).pack(anchor="w", padx=8, pady=(8, 2))
        self.deanon_input_box = ctk.CTkTextbox(right, wrap="word")
        self.deanon_input_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

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
        return True

    def _llm_call_end(self):
        self._llm_busy = False
        self.btn_detect_all.configure(state="normal")
        self.btn_llm_classify.configure(state="normal")

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

    def _refresh_deanon_ref_sheet(self):
        rows = self._pull_rows_from_sheet()
        data = [[r["original"], r["replacement"], r["type"]] for r in rows]
        self.deanon_ref_sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.deanon_ref_sheet.set_all_column_widths()

    def on_deanonymize(self):
        self._refresh_deanon_ref_sheet()
        encoded_text = self.deanon_input_box.get("1.0", "end-1c")
        if not encoded_text.strip():
            messagebox.showinfo("確認", "匿名化されたテキストが空です。")
            return

        rows = self._pull_rows_from_sheet()
        mapping = {r["original"]: r["replacement"] for r in rows if r["replacement"]}
        if not mapping:
            messagebox.showinfo("確認", "対応表が空です。①タブで対応表を作成してから実行してください。")
            return

        rev_mapping = reverse_mapping(mapping)
        result = apply_mapping(encoded_text, rev_mapping)

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