from __future__ import annotations

import csv
import io
import json
import tempfile
import zipfile
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import numpy as np
from openpyxl import Workbook
from PIL import Image, ImageGrab, ImageOps, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

from image_processing import (
    ProcessingSettings,
    compute_two_threshold_otsu,
    count_labels,
    grayscale_preview_image,
    load_image_arrays,
    process_image,
    result_image,
    save_result_image,
)

APP_TITLE = "Trinary Image App"
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
PREVIEW_MAX_SIDE = 1200
PROCESS_DEBOUNCE_MS = 180
DISPLAY_DEBOUNCE_MS = 120


@dataclass
class ImageItem:
    path: Path
    name: str
    base_name: str
    width: int
    height: int
    grayscale: np.ndarray
    alpha: np.ndarray
    labels: np.ndarray
    settings_used: ProcessingSettings


I18N = {
    "ja": {
        "title": "グレースケール3値化アプリ",
        "lead": "画像をグレースケール輝度に変換し、2つのしきい値で黒・グレー・白の3クラスに分類します。初期値は下限80、上限180、中間値130、表示色は黒・グレー・白です。",
        "images": "画像",
        "add_images": "画像を追加",
        "paste": "貼り付け",
        "delete_images": "削除",
        "drop_hint": "画像ファイルをドラッグ＆ドロップできます。" if DND_AVAILABLE else "画像は「画像を追加」または貼り付けで読み込めます。",
        "processing": "3値化設定",
        "lower": "下限しきい値",
        "upper": "上限しきい値",
        "middle": "中間値の濃さ",
        "otsu": "2しきい値版の大津法",
        "invert": "白黒反転",
        "filters": "フィルタ",
        "denoise": "3値専用ノイズ除去",
        "none": "なし",
        "majority": "多数決フィルタ 3×3",
        "mode": "最頻値フィルタ 3×3",
        "morphology": "3値モルフォロジー",
        "dilate": "膨張",
        "erode": "収縮",
        "opening": "オープニング",
        "closing": "クロージング",
        "iterations": "繰り返し回数",
        "display": "表示",
        "colorized": "クラス表示色を使う",
        "black_color": "黒クラスの表示色",
        "gray_color": "グレークラスの表示色",
        "white_color": "白クラスの表示色",
        "batch": "複数画像",
        "apply_same": "設定変更を全画像に自動適用",
        "reprocess": "現在の設定で全画像を再処理",
        "lock_apply": "1枚目の設定を全画像へ適用してロック",
        "unlock": "設定ロックを解除",
        "export": "保存・出力",
        "format": "画像形式",
        "save_current": "現在の画像を保存",
        "save_zip": "全画像をZIP保存",
        "csv": "集計をCSV出力",
        "excel": "集計をExcel出力",
        "reset": "リセット",
        "processing_defaults": "3値化設定をデフォルトに戻す",
        "processing_defaults_restored": "3値化設定をデフォルトに戻しました。",
        "preview": "現在選択中の画像",
        "original": "元画像",
        "grayscale": "グレースケール",
        "result": "3値化結果",
        "black": "黒クラス",
        "gray": "グレークラス",
        "white": "白クラス",
        "total": "総ピクセル数",
        "selected": "選択中",
        "loaded": "{count}枚の画像を追加しました。合計{total}枚です。",
        "loaded_locked": "{count}枚を追加し、ロック済み設定を適用しました。合計{total}枚です。",
        "load_failed": "画像の読み込みに失敗しました: {name}",
        "processed_all": "{count}枚を現在の設定で処理しました。",
        "processed_selected": "選択中の画像を処理しました。",
        "deleted": "{count}枚の画像を削除しました。",
        "delete_no_selection": "削除する画像を1枚以上選択してください。",
        "display_all": "表示設定を全画像へ反映しました。",
        "display_selected": "表示設定を選択中の画像へ反映しました。",
        "otsu_done": "1枚目から大津法を実行しました。下限 {low} / 上限 {high}",
        "locked": "1枚目の設定を全画像へ適用し、設定をロックしました。",
        "unlocked": "設定ロックを解除しました。",
        "saved": "保存しました: {path}",
        "zip_saved": "ZIPを保存しました: {path}",
        "csv_saved": "CSVを保存しました: {path}",
        "excel_saved": "Excelを保存しました: {path}",
        "reset_confirm": "読み込んだ画像と設定をすべてリセットしますか？",
        "reset_done": "リセットしました。",
        "select_image": "先に画像を読み込んでください。",
        "clipboard_empty": "クリップボードに画像がありません。",
        "save_dialog": "保存先を選択",
        "open_dialog": "画像を選択",
        "drop_ignored": "対応している画像ファイルが見つかりませんでした。",
        "error": "エラー",
        "settings_locked": "設定はロックされています。解除してから変更してください。",
        "apply_all_mode": "以降の変更は全画像へ自動適用します。",
        "apply_selected_mode": "以降の変更は選択中の画像だけに適用します。",
        "no_image_name": "画像が選択されていません",
        "color_dialog": "表示色を選択",
    },
    "en": {
        "title": "Grayscale Trinarization App",
        "lead": "Converts each image to grayscale intensity and classifies pixels into black, gray, and white using two thresholds. Defaults: lower 80, upper 180, middle gray 130, with black / gray / white display colors.",
        "images": "Images",
        "add_images": "Add images",
        "paste": "Paste",
        "delete_images": "Delete",
        "drop_hint": "Drag and drop image files into this window." if DND_AVAILABLE else "Use Add images or Paste to load images.",
        "processing": "Trinarization settings",
        "lower": "Lower threshold",
        "upper": "Upper threshold",
        "middle": "Middle gray level",
        "otsu": "Two-threshold Otsu method",
        "invert": "Invert black / white",
        "filters": "Filters",
        "denoise": "Trinary denoising",
        "none": "None",
        "majority": "Majority filter 3×3",
        "mode": "Mode filter 3×3",
        "morphology": "Trinary morphology",
        "dilate": "Dilation",
        "erode": "Erosion",
        "opening": "Opening",
        "closing": "Closing",
        "iterations": "Iterations",
        "display": "Display",
        "colorized": "Use class display colors",
        "black_color": "Black-class display color",
        "gray_color": "Gray-class display color",
        "white_color": "White-class display color",
        "batch": "Multiple images",
        "apply_same": "Automatically apply setting changes to all images",
        "reprocess": "Reprocess all with current settings",
        "lock_apply": "Apply first-image settings to all and lock",
        "unlock": "Unlock settings",
        "export": "Save / export",
        "format": "Image format",
        "save_current": "Save current image",
        "save_zip": "Save all images as ZIP",
        "csv": "Export statistics as CSV",
        "excel": "Export statistics as Excel",
        "reset": "Reset",
        "processing_defaults": "Restore trinarization defaults",
        "processing_defaults_restored": "Restored the trinarization settings to their defaults.",
        "preview": "Current image",
        "original": "Original",
        "grayscale": "Grayscale",
        "result": "Trinarization result",
        "black": "Black class",
        "gray": "Gray class",
        "white": "White class",
        "total": "Total pixels",
        "selected": "Selected",
        "loaded": "Added {count} image(s). Total: {total}.",
        "loaded_locked": "Added {count} image(s) and applied the locked settings. Total: {total}.",
        "load_failed": "Failed to load image: {name}",
        "processed_all": "Processed {count} image(s) with the current settings.",
        "processed_selected": "Processed the selected image.",
        "deleted": "Deleted {count} image(s).",
        "delete_no_selection": "Select one or more images to delete.",
        "display_all": "Applied display settings to all images.",
        "display_selected": "Applied display settings to the selected image.",
        "otsu_done": "Ran Otsu on the first image. Lower {low} / Upper {high}",
        "locked": "Applied the first-image settings to all images and locked the controls.",
        "unlocked": "Settings unlocked.",
        "saved": "Saved: {path}",
        "zip_saved": "Saved ZIP: {path}",
        "csv_saved": "Saved CSV: {path}",
        "excel_saved": "Saved Excel: {path}",
        "reset_confirm": "Reset all loaded images and settings?",
        "reset_done": "Reset complete.",
        "select_image": "Load an image first.",
        "clipboard_empty": "No image found on the clipboard.",
        "save_dialog": "Choose save location",
        "open_dialog": "Select images",
        "drop_ignored": "No supported image files were found.",
        "error": "Error",
        "settings_locked": "Settings are locked. Unlock them before making changes.",
        "apply_all_mode": "Future changes will be applied to all images automatically.",
        "apply_selected_mode": "Future changes will be applied only to the selected image.",
        "no_image_name": "No image selected",
        "color_dialog": "Choose display color",
    },
}


class TrinaryApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.items: list[ImageItem] = []
        self.locked_settings: ProcessingSettings | None = None
        self.temp_dir = tempfile.TemporaryDirectory(prefix="trinary_tk_")
        self.config_path = Path.home() / ".trinary_image_app.json"
        self.language = self._load_language()
        self._updating_controls = False
        self._process_job = None
        self._display_job = None
        self._status_key: str | None = None
        self._status_vars: dict = {}
        self._preview_refs: list[ImageTk.PhotoImage] = []
        self._colors = {
            "black": (0, 0, 0),
            "gray": (120, 120, 120),
            "white": (255, 255, 255),
        }

        self._build_vars()
        self._build_ui()
        self._bind_events()
        self._set_defaults()
        self._render_language()
        self._refresh_enabled_state()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_language(self) -> str:
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
            return "en" if data.get("language") == "en" else "ja"
        except Exception:
            return "ja"

    def _save_language(self):
        try:
            self.config_path.write_text(json.dumps({"language": self.language}), encoding="utf-8")
        except Exception:
            pass

    def tr(self, key: str, **kwargs) -> str:
        return I18N[self.language].get(key, key).format(**kwargs)

    def _build_vars(self):
        self.low_var = tk.IntVar(value=80)
        self.high_var = tk.IntVar(value=180)
        self.middle_var = tk.IntVar(value=130)
        self.invert_var = tk.BooleanVar(value=False)
        self.denoise_var = tk.StringVar(value="none")
        self.morphology_var = tk.StringVar(value="none")
        self.iterations_var = tk.IntVar(value=1)
        self.colorized_var = tk.BooleanVar(value=True)
        self.apply_same_var = tk.BooleanVar(value=True)
        self.format_var = tk.StringVar(value="png")

    def _build_ui(self):
        self.root.title(APP_TITLE)
        self.root.geometry("1420x880")
        self.root.minsize(1040, 680)

        default_font = ("Meiryo UI", 9)
        self.root.option_add("*Font", default_font)
        self.root.option_add("*TCombobox*Listbox.font", "Meiryo UI 9")
        self.root.option_add("*Menu.font", "Meiryo UI 9")

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure(".", font=default_font)
        style.configure("Title.TLabel", font=("Meiryo UI", 18, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Meiryo UI", 10, "bold"))
        style.configure("Stat.TLabel", font=default_font, anchor="center", padding=8)

        outer = ttk.Frame(self.root, padding=(14, 10, 14, 10))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        title_box = ttk.Frame(header)
        title_box.grid(row=0, column=0, sticky="ew")
        title_box.columnconfigure(0, weight=1)
        self.title_label = ttk.Label(title_box, style="Title.TLabel")
        self.title_label.grid(row=0, column=0, sticky="w")
        self.lead_label = ttk.Label(title_box, wraplength=980, foreground="#666666")
        self.lead_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

        lang_frame = ttk.Frame(header)
        lang_frame.grid(row=0, column=1, sticky="ne", padx=(12, 0))
        self.ja_btn = tk.Button(lang_frame, text="日本語", relief="flat", bd=0, font=("Meiryo UI", 9), command=lambda: self._set_language("ja"))
        self.ja_btn.pack(side="left")
        ttk.Label(lang_frame, text="/").pack(side="left", padx=2)
        self.en_btn = tk.Button(lang_frame, text="English", relief="flat", bd=0, font=("Meiryo UI", 9), command=lambda: self._set_language("en"))
        self.en_btn.pack(side="left")

        self.drop_hint = ttk.Label(outer, foreground="#777777")
        self.drop_hint.pack(fill="x", pady=(0, 8))

        paned = ttk.Panedwindow(outer, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Left scrollable settings pane
        left_shell = ttk.Frame(paned, width=390)
        paned.add(left_shell, weight=0)
        left_canvas = tk.Canvas(left_shell, highlightthickness=0, width=390)
        left_scrollbar = ttk.Scrollbar(left_shell, orient="vertical", command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        self.left_inner = ttk.Frame(left_canvas, padding=(2, 2, 8, 2))
        self.left_window = left_canvas.create_window((0, 0), window=self.left_inner, anchor="nw")
        self.left_inner.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.bind("<Configure>", lambda e: left_canvas.itemconfigure(self.left_window, width=e.width))

        # Images
        self.images_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.images_group.pack(fill="x", pady=(0, 8))
        file_row = ttk.Frame(self.images_group)
        file_row.pack(fill="x", padx=8, pady=(8, 4))
        self.add_btn = ttk.Button(file_row, command=self._open_images)
        self.add_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.paste_btn = ttk.Button(file_row, command=self._paste_clipboard)
        self.paste_btn.pack(side="left", fill="x", expand=True, padx=4)
        self.delete_btn = ttk.Button(file_row, command=self._delete_selected_images)
        self.delete_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))
        list_frame = ttk.Frame(self.images_group)
        list_frame.pack(fill="both", padx=8, pady=(4, 8))
        self.image_list = tk.Listbox(
            list_frame, height=6, exportselection=False, selectmode=tk.EXTENDED,
            font=("Meiryo UI", 9)
        )
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.image_list.yview)
        self.image_list.configure(yscrollcommand=list_scroll.set)
        self.image_list.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        # Processing
        self.processing_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.processing_group.pack(fill="x", pady=(0, 8))
        self.low_label = ttk.Label(self.processing_group)
        self.low_label.pack(anchor="w", padx=8, pady=(8, 2))
        self.low_scale, self.low_spin = self._make_scale_spin(self.processing_group, self.low_var, 0, 255)
        self.high_label = ttk.Label(self.processing_group)
        self.high_label.pack(anchor="w", padx=8, pady=(8, 2))
        self.high_scale, self.high_spin = self._make_scale_spin(self.processing_group, self.high_var, 0, 255)
        self.middle_label = ttk.Label(self.processing_group)
        self.middle_label.pack(anchor="w", padx=8, pady=(8, 2))
        self.middle_scale, self.middle_spin = self._make_scale_spin(self.processing_group, self.middle_var, 32, 223)
        self.otsu_btn = ttk.Button(self.processing_group, command=self._run_otsu)
        self.otsu_btn.pack(fill="x", padx=8, pady=(8, 4))
        self.invert_check = ttk.Checkbutton(self.processing_group, variable=self.invert_var, command=self._schedule_display_change)
        self.invert_check.pack(anchor="w", padx=8, pady=(4, 4))
        self.processing_defaults_btn = ttk.Button(self.processing_group, command=self._restore_processing_defaults)
        self.processing_defaults_btn.pack(fill="x", padx=8, pady=(4, 8))

        # Filters
        self.filters_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.filters_group.pack(fill="x", pady=(0, 8))
        self.denoise_label = ttk.Label(self.filters_group)
        self.denoise_label.pack(anchor="w", padx=8, pady=(8, 2))
        self.denoise_combo = ttk.Combobox(self.filters_group, state="readonly")
        self.denoise_combo.pack(fill="x", padx=8, pady=(0, 6))
        self.morphology_label = ttk.Label(self.filters_group)
        self.morphology_label.pack(anchor="w", padx=8, pady=(4, 2))
        self.morphology_combo = ttk.Combobox(self.filters_group, state="readonly")
        self.morphology_combo.pack(fill="x", padx=8, pady=(0, 6))
        iter_row = ttk.Frame(self.filters_group)
        iter_row.pack(fill="x", padx=8, pady=(4, 8))
        self.iterations_label = ttk.Label(iter_row)
        self.iterations_label.pack(side="left")
        self.iterations_spin = ttk.Spinbox(iter_row, from_=1, to=10, increment=1, textvariable=self.iterations_var, width=6)
        self.iterations_spin.pack(side="right")

        # Display
        self.display_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.display_group.pack(fill="x", pady=(0, 8))
        self.colorized_check = ttk.Checkbutton(self.display_group, variable=self.colorized_var, command=self._schedule_display_change)
        self.colorized_check.pack(anchor="w", padx=8, pady=(8, 6))
        self.black_color_btn = tk.Button(self.display_group, command=lambda: self._choose_color("black"), anchor="w")
        self.gray_color_btn = tk.Button(self.display_group, command=lambda: self._choose_color("gray"), anchor="w")
        self.white_color_btn = tk.Button(self.display_group, command=lambda: self._choose_color("white"), anchor="w")
        for btn in (self.black_color_btn, self.gray_color_btn, self.white_color_btn):
            btn.pack(fill="x", padx=8, pady=3)
        ttk.Frame(self.display_group, height=4).pack()

        # Batch
        self.batch_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.batch_group.pack(fill="x", pady=(0, 8))
        self.apply_same_check = ttk.Checkbutton(self.batch_group, variable=self.apply_same_var, command=self._on_apply_same_changed)
        self.apply_same_check.pack(anchor="w", padx=8, pady=(8, 6))
        self.reprocess_btn = ttk.Button(self.batch_group, command=self._reprocess_all)
        self.reprocess_btn.pack(fill="x", padx=8, pady=3)
        self.lock_btn = ttk.Button(self.batch_group, command=self._toggle_lock)
        self.lock_btn.pack(fill="x", padx=8, pady=(3, 8))

        # Export
        self.export_group = ttk.LabelFrame(self.left_inner, style="Section.TLabelframe")
        self.export_group.pack(fill="x", pady=(0, 8))
        format_row = ttk.Frame(self.export_group)
        format_row.pack(fill="x", padx=8, pady=(8, 4))
        self.format_label = ttk.Label(format_row)
        self.format_label.pack(side="left")
        self.format_combo = ttk.Combobox(format_row, textvariable=self.format_var, state="readonly", width=8, values=("png", "jpeg"))
        self.format_combo.pack(side="right")
        self.save_current_btn = ttk.Button(self.export_group, command=self._save_current)
        self.zip_btn = ttk.Button(self.export_group, command=self._save_zip)
        self.csv_btn = ttk.Button(self.export_group, command=self._export_csv)
        self.excel_btn = ttk.Button(self.export_group, command=self._export_excel)
        self.reset_btn = ttk.Button(self.export_group, command=self._reset)
        for btn in (self.save_current_btn, self.zip_btn, self.csv_btn, self.excel_btn, self.reset_btn):
            btn.pack(fill="x", padx=8, pady=3)
        ttk.Frame(self.export_group, height=4).pack()

        # Right preview pane
        right = ttk.Frame(paned, padding=(10, 2, 2, 2))
        paned.add(right, weight=1)
        self.preview_title = ttk.Label(right, font=("Meiryo UI", 12, "bold"))
        self.preview_title.pack(anchor="w")
        self.current_name_label = ttk.Label(right, foreground="#666666")
        self.current_name_label.pack(anchor="w", pady=(1, 8))

        preview_row = ttk.Frame(right)
        preview_row.pack(fill="both", expand=True)
        for col in range(3):
            preview_row.columnconfigure(col, weight=1, uniform="preview")
        preview_row.rowconfigure(1, weight=1)
        self.original_title = ttk.Label(preview_row, anchor="center", font=("Meiryo UI", 10, "bold"))
        self.gray_title = ttk.Label(preview_row, anchor="center", font=("Meiryo UI", 10, "bold"))
        self.result_title = ttk.Label(preview_row, anchor="center", font=("Meiryo UI", 10, "bold"))
        for i, title in enumerate((self.original_title, self.gray_title, self.result_title)):
            title.grid(row=0, column=i, sticky="ew", padx=4, pady=(0, 4))
        self.original_preview = tk.Label(preview_row, bg="#f4f4f4", relief="solid", bd=1)
        self.gray_preview = tk.Label(preview_row, bg="#f4f4f4", relief="solid", bd=1)
        self.result_preview = tk.Label(preview_row, bg="#f4f4f4", relief="solid", bd=1)
        for i, label in enumerate((self.original_preview, self.gray_preview, self.result_preview)):
            label.grid(row=1, column=i, sticky="nsew", padx=4)
            label.bind("<Configure>", lambda e: self._schedule_preview_refresh())

        stats = ttk.Frame(right)
        stats.pack(fill="x", pady=(10, 0))
        for i in range(4):
            stats.columnconfigure(i, weight=1, uniform="stat")
        self.stat_labels = []
        for i in range(4):
            lbl = ttk.Label(stats, style="Stat.TLabel", relief="solid")
            lbl.grid(row=0, column=i, sticky="ew", padx=3)
            self.stat_labels.append(lbl)

        self.status_var = tk.StringVar(value="")
        status = ttk.Label(outer, textvariable=self.status_var, relief="sunken", anchor="w", padding=(6, 3))
        status.pack(fill="x", pady=(8, 0))

    def _make_scale_spin(self, parent, variable: tk.IntVar, minimum: int, maximum: int):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8)
        scale = ttk.Scale(row, from_=minimum, to=maximum, orient="horizontal")
        scale.pack(side="left", fill="x", expand=True, padx=(0, 8))
        spin = ttk.Spinbox(row, from_=minimum, to=maximum, increment=1, width=6, textvariable=variable)
        spin.pack(side="right")

        def var_to_scale(*_):
            if self._updating_controls:
                return
            try:
                value = int(variable.get())
            except (tk.TclError, ValueError):
                return
            value = max(minimum, min(maximum, value))
            if abs(float(scale.get()) - value) > 0.5:
                scale.set(value)

        def scale_to_var(value):
            if self._updating_controls:
                return
            variable.set(int(round(float(value))))

        variable.trace_add("write", var_to_scale)
        scale.configure(command=scale_to_var)
        scale.set(variable.get())
        return scale, spin

    def _bind_events(self):
        self.image_list.bind("<<ListboxSelect>>", self._on_image_selected)
        self.image_list.bind("<Delete>", lambda e: self._delete_selected_images())
        self.root.bind_all("<Control-v>", lambda e: self._paste_clipboard())
        self.root.bind_all("<Command-v>", lambda e: self._paste_clipboard())

        for var in (self.low_var, self.high_var):
            var.trace_add("write", lambda *_: self._schedule_processing_change())
        self.middle_var.trace_add("write", lambda *_: self._schedule_display_change())
        self.iterations_var.trace_add("write", lambda *_: self._schedule_processing_change())
        self.denoise_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule_processing_change(immediate=True))
        self.morphology_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule_processing_change(immediate=True))

        if DND_AVAILABLE and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _set_defaults(self):
        self._updating_controls = True
        self.low_var.set(80)
        self.high_var.set(180)
        self.middle_var.set(130)
        self.invert_var.set(False)
        self.denoise_var.set("none")
        self.morphology_var.set("none")
        self.iterations_var.set(1)
        self.colorized_var.set(True)
        self.apply_same_var.set(True)
        self.format_var.set("png")
        self._colors = {"black": (0, 0, 0), "gray": (120, 120, 120), "white": (255, 255, 255)}
        self._updating_controls = False
        self._update_color_buttons()

    def _render_language(self):
        self.root.title(self.tr("title"))
        self.title_label.configure(text=self.tr("title"))
        self.lead_label.configure(text=self.tr("lead"))
        self.drop_hint.configure(text=self.tr("drop_hint"))
        self.images_group.configure(text=self.tr("images"))
        self.add_btn.configure(text=self.tr("add_images"))
        self.paste_btn.configure(text=self.tr("paste"))
        self.delete_btn.configure(text=self.tr("delete_images"))
        self.processing_group.configure(text=self.tr("processing"))
        self.low_label.configure(text=self.tr("lower"))
        self.high_label.configure(text=self.tr("upper"))
        self.middle_label.configure(text=self.tr("middle"))
        self.otsu_btn.configure(text=self.tr("otsu"))
        self.invert_check.configure(text=self.tr("invert"))
        self.processing_defaults_btn.configure(text=self.tr("processing_defaults"))
        self.filters_group.configure(text=self.tr("filters"))
        self.denoise_label.configure(text=self.tr("denoise"))
        self.morphology_label.configure(text=self.tr("morphology"))
        self.iterations_label.configure(text=self.tr("iterations"))
        self.display_group.configure(text=self.tr("display"))
        self.colorized_check.configure(text=self.tr("colorized"))
        self.batch_group.configure(text=self.tr("batch"))
        self.apply_same_check.configure(text=self.tr("apply_same"))
        self.reprocess_btn.configure(text=self.tr("reprocess"))
        self.lock_btn.configure(text=self.tr("unlock") if self.locked_settings else self.tr("lock_apply"))
        self.export_group.configure(text=self.tr("export"))
        self.format_label.configure(text=self.tr("format"))
        self.save_current_btn.configure(text=self.tr("save_current"))
        self.zip_btn.configure(text=self.tr("save_zip"))
        self.csv_btn.configure(text=self.tr("csv"))
        self.excel_btn.configure(text=self.tr("excel"))
        self.reset_btn.configure(text=self.tr("reset"))
        self.preview_title.configure(text=self.tr("preview"))
        self.original_title.configure(text=self.tr("original"))
        self.gray_title.configure(text=self.tr("grayscale"))
        self.result_title.configure(text=self.tr("result"))

        denoise_values = {
            "none": self.tr("none"),
            "majority3": self.tr("majority"),
            "mode3": self.tr("mode"),
        }
        morphology_values = {
            "none": self.tr("none"),
            "dilate": self.tr("dilate"),
            "erode": self.tr("erode"),
            "open": self.tr("opening"),
            "close": self.tr("closing"),
        }
        self._set_combo_display(self.denoise_combo, self.denoise_var, denoise_values)
        self._set_combo_display(self.morphology_combo, self.morphology_var, morphology_values)
        self._update_color_buttons()
        self._update_language_button_styles()
        self._refresh_current_view()
        if self._status_key:
            self.status_var.set(self.tr(self._status_key, **self._status_vars))

    @staticmethod
    def _set_combo_display(combo: ttk.Combobox, variable: tk.StringVar, mapping: dict[str, str]):
        current_key = variable.get()
        combo["values"] = list(mapping.values())
        combo._key_to_label = mapping
        combo._label_to_key = {v: k for k, v in mapping.items()}
        if current_key in mapping:
            combo.set(mapping[current_key])

    def _sync_combo_keys(self):
        for combo, var in ((self.denoise_combo, self.denoise_var), (self.morphology_combo, self.morphology_var)):
            mapping = getattr(combo, "_label_to_key", {})
            label = combo.get()
            if label in mapping:
                var.set(mapping[label])

    def _update_language_button_styles(self):
        active = {"font": ("Meiryo UI", 9, "underline"), "fg": "#222222"}
        normal = {"font": ("Meiryo UI", 9), "fg": "#777777"}
        self.ja_btn.configure(**(active if self.language == "ja" else normal))
        self.en_btn.configure(**(active if self.language == "en" else normal))

    def _set_language(self, language: str):
        self.language = "en" if language == "en" else "ja"
        self._save_language()
        self._render_language()

    def _set_status(self, key: str | None = None, **vars_):
        self._status_key = key
        self._status_vars = vars_
        self.status_var.set(self.tr(key, **vars_) if key else "")

    def _current_item(self) -> ImageItem | None:
        selection = self.image_list.curselection()
        if not selection:
            return None
        idx = selection[0]
        return self.items[idx] if 0 <= idx < len(self.items) else None

    def _get_controls_settings(self) -> ProcessingSettings:
        self._sync_combo_keys()
        try:
            low = int(self.low_var.get())
            high = int(self.high_var.get())
        except (tk.TclError, ValueError):
            low, high = 80, 180
        settings = ProcessingSettings(
            low=low,
            high=high,
            invert=bool(self.invert_var.get()),
            denoise=self.denoise_var.get(),
            morphology=self.morphology_var.get(),
            iterations=int(self.iterations_var.get() or 1),
            gray_level=int(self.middle_var.get() or 130),
            colorized=bool(self.colorized_var.get()),
            black_color=self._colors["black"],
            gray_color=self._colors["gray"],
            white_color=self._colors["white"],
        ).normalized()
        if settings.low != low or settings.high != high:
            self._apply_threshold_values(settings.low, settings.high)
        return settings

    def _apply_threshold_values(self, low: int, high: int):
        self._updating_controls = True
        self.low_var.set(low)
        self.high_var.set(high)
        self.low_scale.set(low)
        self.high_scale.set(high)
        self._updating_controls = False

    def _apply_settings_to_controls(self, settings: ProcessingSettings):
        settings = settings.normalized()
        self._updating_controls = True
        self.low_var.set(settings.low)
        self.high_var.set(settings.high)
        self.middle_var.set(settings.gray_level)
        self.low_scale.set(settings.low)
        self.high_scale.set(settings.high)
        self.middle_scale.set(settings.gray_level)
        self.invert_var.set(settings.invert)
        self.denoise_var.set(settings.denoise)
        self.morphology_var.set(settings.morphology)
        self.iterations_var.set(settings.iterations)
        self.colorized_var.set(settings.colorized)
        self._colors = {
            "black": tuple(settings.black_color),
            "gray": tuple(settings.gray_color),
            "white": tuple(settings.white_color),
        }
        self._updating_controls = False
        self._render_language()

    def _schedule_processing_change(self, immediate: bool = False):
        if self._updating_controls:
            return
        if self.locked_settings:
            self._set_status("settings_locked")
            return
        if self._process_job:
            self.root.after_cancel(self._process_job)
        self._process_job = self.root.after(1 if immediate else PROCESS_DEBOUNCE_MS, self._apply_processing_change)

    def _schedule_display_change(self):
        if self._updating_controls:
            return
        if self.locked_settings:
            self._set_status("settings_locked")
            return
        if self._display_job:
            self.root.after_cancel(self._display_job)
        self._display_job = self.root.after(DISPLAY_DEBOUNCE_MS, self._apply_display_change)

    def _apply_processing_change(self):
        self._process_job = None
        if self._updating_controls or self.locked_settings or not self.items:
            return
        settings = self._get_controls_settings()
        if self.apply_same_var.get():
            self._process_items(self.items, settings)
            self._set_status("processed_all", count=len(self.items))
        else:
            item = self._current_item()
            if item:
                self._process_items([item], settings)
                self._set_status("processed_selected")
        self._refresh_current_view()

    def _apply_display_change(self):
        self._display_job = None
        if self._updating_controls or self.locked_settings or not self.items:
            return
        source = self._get_controls_settings()
        targets = self.items if self.apply_same_var.get() else [self._current_item()]
        for item in targets:
            if not item:
                continue
            item.settings_used = replace(
                item.settings_used,
                invert=source.invert,
                gray_level=source.gray_level,
                colorized=source.colorized,
                black_color=source.black_color,
                gray_color=source.gray_color,
                white_color=source.white_color,
            )
        self._set_status("display_all" if self.apply_same_var.get() else "display_selected")
        self._refresh_current_view()

    def _process_items(self, items: list[ImageItem], settings: ProcessingSettings):
        settings = settings.normalized()
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            for item in items:
                item.labels = process_image(item.grayscale, settings)
                item.settings_used = settings
                self.root.update_idletasks()
        finally:
            self.root.configure(cursor="")

    def _reprocess_all(self):
        if not self.items:
            self._set_status("select_image")
            return
        if self.locked_settings:
            self._set_status("settings_locked")
            return
        settings = self._get_controls_settings()
        self._process_items(self.items, settings)
        self._set_status("processed_all", count=len(self.items))
        self._refresh_current_view()

    def _restore_processing_defaults(self):
        if self.locked_settings:
            self._set_status("settings_locked")
            return

        self._updating_controls = True
        try:
            self.low_var.set(80)
            self.high_var.set(180)
            self.middle_var.set(130)
            self.low_scale.set(80)
            self.high_scale.set(180)
            self.middle_scale.set(130)
            self.invert_var.set(False)
        finally:
            self._updating_controls = False

        if self.items:
            settings = self._get_controls_settings()
            if self.apply_same_var.get():
                self._process_items(self.items, settings)
            else:
                item = self._current_item()
                if item:
                    self._process_items([item], settings)
            self._refresh_current_view()

        self._set_status("processing_defaults_restored")

    def _run_otsu(self):
        if not self.items:
            self._set_status("select_image")
            return
        if self.locked_settings:
            self._set_status("settings_locked")
            return
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            low, high = compute_two_threshold_otsu(self.items[0].grayscale)
        finally:
            self.root.configure(cursor="")
        self._apply_threshold_values(low, high)
        self._apply_processing_change()
        self._set_status("otsu_done", low=low, high=high)

    def _toggle_lock(self):
        if not self.items:
            self._set_status("select_image")
            return
        if self.locked_settings:
            self.locked_settings = None
            item = self._current_item()
            if item:
                self._apply_settings_to_controls(item.settings_used)
            self._set_status("unlocked")
        else:
            first_settings = self.items[0].settings_used
            self.locked_settings = first_settings
            self._apply_settings_to_controls(first_settings)
            self._process_items(self.items, first_settings)
            self._set_status("locked")
        self._refresh_enabled_state()
        self._render_language()
        self._refresh_current_view()

    def _on_apply_same_changed(self):
        if self.locked_settings:
            self.apply_same_var.set(True)
            self._set_status("settings_locked")
            return
        self._set_status("apply_all_mode" if self.apply_same_var.get() else "apply_selected_mode")

    def _choose_color(self, which: str):
        if self.locked_settings:
            self._set_status("settings_locked")
            return
        rgb = self._colors[which]
        initial = "#%02x%02x%02x" % rgb
        _, hex_color = colorchooser.askcolor(initialcolor=initial, title=self.tr("color_dialog"), parent=self.root)
        if not hex_color:
            return
        value = hex_color.lstrip("#")
        self._colors[which] = tuple(int(value[i:i+2], 16) for i in (0, 2, 4))
        self._update_color_buttons()
        self._schedule_display_change()

    def _update_color_buttons(self):
        configs = [
            (self.black_color_btn, "black", "black_color"),
            (self.gray_color_btn, "gray", "gray_color"),
            (self.white_color_btn, "white", "white_color"),
        ]
        for btn, key, text_key in configs:
            rgb = self._colors[key]
            hex_color = "#%02x%02x%02x" % rgb
            luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
            fg = "#ffffff" if luminance < 135 else "#000000"
            btn.configure(text=f"{self.tr(text_key)}   RGB {rgb[0]}, {rgb[1]}, {rgb[2]}", bg=hex_color, fg=fg, activebackground=hex_color, activeforeground=fg)

    def _open_images(self):
        files = filedialog.askopenfilenames(
            parent=self.root,
            title=self.tr("open_dialog"),
            filetypes=[
                ("Images", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"),
                ("All files", "*.*"),
            ],
        )
        if files:
            self._load_paths([Path(p) for p in files])

    def _load_paths(self, paths: list[Path]):
        paths = [p for p in paths if p.exists() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
        if not paths:
            self._set_status("drop_ignored")
            return
        added = 0
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            for path in paths:
                try:
                    width, height, gray, alpha = load_image_arrays(path)
                    settings = self.locked_settings or self._get_controls_settings()
                    labels = process_image(gray, settings)
                    item = ImageItem(
                        path=path,
                        name=path.name,
                        base_name=path.stem,
                        width=width,
                        height=height,
                        grayscale=gray,
                        alpha=alpha,
                        labels=labels,
                        settings_used=settings,
                    )
                    self.items.append(item)
                    self.image_list.insert("end", item.name)
                    added += 1
                    self.root.update_idletasks()
                except Exception:
                    self._set_status("load_failed", name=path.name)
        finally:
            self.root.configure(cursor="")

        if added:
            if not self.image_list.curselection():
                self.image_list.selection_set(0)
                self.image_list.activate(0)
            self._refresh_enabled_state()
            self._refresh_current_view()
            if self.locked_settings:
                self._set_status("loaded_locked", count=added, total=len(self.items))
            else:
                self._set_status("loaded", count=added, total=len(self.items))

    def _delete_selected_images(self):
        selection = sorted(self.image_list.curselection())
        if not selection:
            self._set_status("delete_no_selection")
            return

        next_index = selection[0]
        for idx in reversed(selection):
            if 0 <= idx < len(self.items):
                del self.items[idx]
            self.image_list.delete(idx)

        if not self.items:
            self.locked_settings = None
        else:
            next_index = min(next_index, len(self.items) - 1)
            self.image_list.selection_clear(0, "end")
            self.image_list.selection_set(next_index)
            self.image_list.activate(next_index)
            self.image_list.see(next_index)

        self._refresh_enabled_state()
        self._render_language()
        self._refresh_current_view()
        self._set_status("deleted", count=len(selection))

    def _paste_clipboard(self):
        try:
            content = ImageGrab.grabclipboard()
        except Exception:
            content = None
        if isinstance(content, Image.Image):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            path = Path(self.temp_dir.name) / f"clipboard_{stamp}.png"
            content.convert("RGBA").save(path, "PNG")
            self._load_paths([path])
            return
        if isinstance(content, list):
            self._load_paths([Path(p) for p in content])
            return
        self._set_status("clipboard_empty")

    def _on_drop(self, event):
        try:
            paths = [Path(p) for p in self.root.tk.splitlist(event.data)]
        except Exception:
            paths = []
        self._load_paths(paths)

    def _on_image_selected(self, _event=None):
        item = self._current_item()
        if item and not self.locked_settings:
            self._apply_settings_to_controls(item.settings_used)
        self._refresh_enabled_state()
        self._refresh_current_view()

    def _schedule_preview_refresh(self):
        if hasattr(self, "_preview_job") and self._preview_job:
            try:
                self.root.after_cancel(self._preview_job)
            except Exception:
                pass
        self._preview_job = self.root.after(80, self._refresh_current_view)

    def _refresh_current_view(self):
        item = self._current_item()
        if not item:
            self.current_name_label.configure(text=self.tr("no_image_name"))
            for label in (self.original_preview, self.gray_preview, self.result_preview):
                label.configure(image="", text="")
            self._preview_refs = []
            self._update_stats(None)
            return

        self.current_name_label.configure(text=f"{item.name}   {item.width} × {item.height}")
        try:
            with Image.open(item.path) as src:
                original = ImageOps.exif_transpose(src).convert("RGBA")
                original.thumbnail((PREVIEW_MAX_SIDE, PREVIEW_MAX_SIDE), Image.Resampling.LANCZOS)
                original = original.copy()
        except Exception:
            original = Image.new("RGBA", (10, 10), "white")
        gray = grayscale_preview_image(item.grayscale, item.alpha, PREVIEW_MAX_SIDE)
        result = result_image(item.labels, item.alpha, item.settings_used, PREVIEW_MAX_SIDE)

        refs = []
        for pil_img, label in zip((original, gray, result), (self.original_preview, self.gray_preview, self.result_preview)):
            w = max(180, label.winfo_width() - 12)
            h = max(180, label.winfo_height() - 12)
            img = pil_img.copy()
            img.thumbnail((w, h), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            label.configure(image=photo)
            refs.append(photo)
        self._preview_refs = refs
        self._update_stats(item)

    def _update_stats(self, item: ImageItem | None):
        keys = ("black", "gray", "white", "total")
        if item:
            stats = count_labels(item.labels, item.settings_used.invert)
            for lbl, key in zip(self.stat_labels, keys):
                lbl.configure(text=f"{self.tr(key)}\n{stats[key]:,} px")
        else:
            for lbl, key in zip(self.stat_labels, keys):
                lbl.configure(text=f"{self.tr(key)}\n—")

    def _stats_rows(self):
        rows = []
        for idx, item in enumerate(self.items, 1):
            s = item.settings_used.normalized()
            stats = count_labels(item.labels, s.invert)
            rows.append({
                "index": idx,
                "file_name": item.name,
                "width": item.width,
                "height": item.height,
                "total_pixels": stats["total"],
                "black_pixels": stats["black"],
                "gray_pixels": stats["gray"],
                "white_pixels": stats["white"],
                "black_ratio": stats["black"] / stats["total"] if stats["total"] else 0,
                "gray_ratio": stats["gray"] / stats["total"] if stats["total"] else 0,
                "white_ratio": stats["white"] / stats["total"] if stats["total"] else 0,
                "threshold_low": s.low,
                "threshold_high": s.high,
                "inverted": int(s.invert),
                "denoise": s.denoise,
                "morphology": s.morphology,
                "morphology_iterations": s.iterations,
                "gray_level": s.gray_level,
                "colorized": int(s.colorized),
                "black_color": "#%02x%02x%02x" % s.black_color,
                "gray_color": "#%02x%02x%02x" % s.gray_color,
                "white_color": "#%02x%02x%02x" % s.white_color,
            })
        return rows

    @staticmethod
    def _csv_text(rows: list[dict]) -> str:
        if not rows:
            return ""
        output = io.StringIO(newline="")
        writer = csv.DictWriter(
            output, fieldnames=list(rows[0].keys()), lineterminator="\n"
        )
        writer.writeheader()
        writer.writerows(rows)
        return "\ufeff" + output.getvalue()

    def _save_current(self):
        item = self._current_item()
        if not item:
            self._set_status("select_image")
            return
        fmt = self.format_var.get()
        ext = ".jpg" if fmt == "jpeg" else ".png"
        filetypes = [("JPEG", "*.jpg")] if fmt == "jpeg" else [("PNG", "*.png")]
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("save_dialog"),
            defaultextension=ext,
            initialfile=f"{item.base_name}-trinary{ext}",
            filetypes=filetypes,
        )
        if not path:
            return
        try:
            save_result_image(path, item.labels, item.alpha, item.settings_used, fmt)
            self._set_status("saved", path=path)
        except Exception as exc:
            messagebox.showerror(self.tr("error"), str(exc), parent=self.root)

    def _save_zip(self):
        if not self.items:
            self._set_status("select_image")
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("save_dialog"),
            defaultextension=".zip",
            initialfile=f"trinary_images_{datetime.now():%Y%m%d_%H%M%S}.zip",
            filetypes=[("ZIP", "*.zip")],
        )
        if not path:
            return
        fmt = self.format_var.get()
        ext = ".jpg" if fmt == "jpeg" else ".png"
        try:
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in self.items:
                    image = result_image(item.labels, item.alpha, item.settings_used)
                    buffer = io.BytesIO()
                    if fmt == "jpeg":
                        background = Image.new("RGB", image.size, "white")
                        background.paste(image, mask=image.getchannel("A"))
                        background.save(buffer, "JPEG", quality=92)
                    else:
                        image.save(buffer, "PNG")
                    zf.writestr(f"{item.base_name}-trinary{ext}", buffer.getvalue())
                zf.writestr("pixel_stats.csv", self._csv_text(self._stats_rows()).encode("utf-8"))
            self._set_status("zip_saved", path=path)
        except Exception as exc:
            messagebox.showerror(self.tr("error"), str(exc), parent=self.root)

    def _export_csv(self):
        if not self.items:
            self._set_status("select_image")
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("save_dialog"),
            defaultextension=".csv",
            initialfile=f"trinary_pixel_stats_{datetime.now():%Y%m%d_%H%M%S}.csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return
        try:
            Path(path).write_bytes(self._csv_text(self._stats_rows()).encode("utf-8"))
            self._set_status("csv_saved", path=path)
        except Exception as exc:
            messagebox.showerror(self.tr("error"), str(exc), parent=self.root)

    def _export_excel(self):
        if not self.items:
            self._set_status("select_image")
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("save_dialog"),
            defaultextension=".xlsx",
            initialfile=f"trinary_pixel_stats_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            rows = self._stats_rows()
            wb = Workbook(write_only=True)
            ws = wb.create_sheet(title="pixel_stats")
            if rows:
                headers = list(rows[0].keys())
                ws.append(headers)
                for row in rows:
                    ws.append([row[h] for h in headers])
            wb.save(path)
            self._set_status("excel_saved", path=path)
        except Exception as exc:
            messagebox.showerror(self.tr("error"), str(exc), parent=self.root)

    def _refresh_enabled_state(self):
        has = bool(self.items)
        locked = self.locked_settings is not None
        normal_state = "disabled" if locked else "normal"
        readonly_state = "disabled" if locked else "readonly"

        for widget in (self.low_spin, self.high_spin, self.middle_spin, self.iterations_spin):
            widget.configure(state=normal_state)
        for widget in (self.low_scale, self.high_scale, self.middle_scale):
            widget.state(["disabled"] if locked else ["!disabled"])
        self.denoise_combo.configure(state=readonly_state)
        self.morphology_combo.configure(state=readonly_state)
        for widget in (self.invert_check, self.colorized_check, self.apply_same_check, self.black_color_btn, self.gray_color_btn, self.white_color_btn):
            try:
                widget.configure(state="disabled" if locked else "normal")
            except tk.TclError:
                pass

        self.otsu_btn.configure(state="normal" if has and not locked else "disabled")
        self.processing_defaults_btn.configure(state="normal" if not locked else "disabled")
        self.reprocess_btn.configure(state="normal" if has and not locked else "disabled")
        self.lock_btn.configure(state="normal" if has else "disabled")
        self.delete_btn.configure(state="normal" if self.image_list.curselection() else "disabled")
        for widget in (self.save_current_btn, self.zip_btn, self.csv_btn, self.excel_btn, self.reset_btn):
            widget.configure(state="normal" if has else "disabled")

    def _reset(self):
        if not self.items:
            return
        if not messagebox.askyesno(self.tr("reset"), self.tr("reset_confirm"), parent=self.root):
            return
        self.items.clear()
        self.image_list.delete(0, "end")
        self.locked_settings = None
        self._set_defaults()
        self._refresh_enabled_state()
        self._render_language()
        self._set_status("reset_done")

    def _on_close(self):
        try:
            self.temp_dir.cleanup()
        except Exception:
            pass
        self.root.destroy()


def create_root():
    if DND_AVAILABLE:
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
    return tk.Tk()


def main():
    root = create_root()
    app = TrinaryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()