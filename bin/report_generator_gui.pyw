from __future__ import annotations

import ctypes
import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import image_input_pipeline as image_pipeline
import process_excel_to_word as report_pipeline
import type2_report_pipeline as type2_pipeline
from process_excel_to_word import (
    create_gemini_client,
    extract_guarantor_details_from_text,
)
from image_input_pipeline import list_person_folders

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False


IS_FROZEN = bool(getattr(sys, "frozen", False))
if IS_FROZEN:
    executable_dir = Path(sys.executable).resolve().parent
    RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", executable_dir)) / "resources"
    if (
        executable_dir.name.lower() == "dist"
        and (executable_dir.parent / "bin").is_dir()
    ):
        PROJECT_ROOT = executable_dir.parent
    else:
        PROJECT_ROOT = executable_dir
    BIN_DIR = PROJECT_ROOT / "bin"
else:
    BIN_DIR = Path(__file__).resolve().parent
    PROJECT_ROOT = BIN_DIR.parent
    RESOURCE_ROOT = BIN_DIR / "template"
TYPE1_RESOURCE_DIR = RESOURCE_ROOT / "1"
TYPE2_RESOURCE_DIR = RESOURCE_ROOT / "2"
PROCESS_SCRIPT = BIN_DIR / "process_excel_to_word.py"
IMAGE_PIPELINE_SCRIPT = BIN_DIR / "image_input_pipeline.py"
TYPE2_PROCESS_SCRIPT = BIN_DIR / "type2_report_pipeline.py"
DEFAULT_INPUT_DIR = PROJECT_ROOT / "Input"
JSON_DIR = BIN_DIR / "json"
DEFAULT_RECORDS_FILE = JSON_DIR / "图片提取数据.json"
CURRENT_RECORDS_FILE = JSON_DIR / "当前人员数据.json"
TYPE2_RECORDS_FILE = JSON_DIR / "类型2输入数据.json"
DEFAULT_INDIVIDUAL_RECORDS_DIR = JSON_DIR / "人员数据"
SETTINGS_FILE = JSON_DIR / "panel_settings.json"
GUARANTOR_STORE_FILE = JSON_DIR / "保证人资料.json"
DEFAULT_SUPPLEMENTS = (
    BIN_DIR / "template" / "1" / "图片输入补充数据.json"
)
DEFAULT_TEMPLATE = TYPE1_RESOURCE_DIR / "价值分析报告-自动生成基底模板.docx"
DEFAULT_PROMPT = TYPE1_RESOURCE_DIR / "价值分析报告自动生成-Prompt.md"
DEFAULT_RULES = TYPE1_RESOURCE_DIR / "价值分析报告生成规则.json"
TYPE2_REPORT_TEMPLATE = (
    TYPE2_RESOURCE_DIR / "企业价值评估报告-自动生成基底模板.docx"
)
TYPE2_EXPLANATION_TEMPLATE = (
    TYPE2_RESOURCE_DIR / "企业价值评估说明-自动生成基底模板.docx"
)
TYPE2_PROMPT = TYPE2_RESOURCE_DIR / "企业价值评估报告自动生成-Prompt.md"
TYPE2_RULES = TYPE2_RESOURCE_DIR / "企业价值评估报告生成规则.json"
DEFAULT_OUTPUT = PROJECT_ROOT / "word"
API_KEY_FILE = PROJECT_ROOT / "gemini_api.txt"
APP_ICON_FILE = BIN_DIR / "with a pen" / "256x256.ico"

DOCUMENT_TYPE_CONFIG = {
    "1": {
        "name": "当前价值分析报告",
        "template": DEFAULT_TEMPLATE,
        "prompt": DEFAULT_PROMPT,
        "rules": DEFAULT_RULES,
    },
    "2": {
        "name": "企业价值评估报告及评估说明",
        "template": TYPE2_REPORT_TEMPLATE,
        "explanation_template": TYPE2_EXPLANATION_TEMPLATE,
        "prompt": TYPE2_PROMPT,
        "rules": TYPE2_RULES,
    },
}

WINDOW_BG = "#F4F5F2"
PANEL_BG = "#FFFFFF"
INPUT_BG = "#FCFCFA"
TEXT_DARK = "#202821"
TEXT_MID = "#59645B"
BORDER = "#C9CEC8"
GREEN = "#315C45"
GREEN_ACTIVE = "#274B38"
RUST = "#B85B32"
RUST_ACTIVE = "#9E4D2A"
GRAY_BUTTON = "#657069"
GRAY_ACTIVE = "#525B55"
ERROR = "#A33A32"
SERIAL_NEXT_DELAY_MS = 3000


def enable_windows_dpi_awareness() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(-4):
            return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass


def apply_tk_scaling(root: tk.Tk) -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        root.tk.call("tk", "scaling", float(root.winfo_fpixels("1i")) / 72.0)
    except Exception:
        pass


class ReportGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self._settings_ready = False
        self._settings_save_after_id: str | None = None
        self.saved_settings = self._load_panel_settings()
        self.processing = False
        self.person_folders: list[Path] = []
        self.current_person_folder: Path | None = None
        self.input_workbook: Path | None = None
        self.type2_project_folders: list[Path] = []
        self.current_type2_project_folder: Path | None = None
        self.total_type2_project_count = 0
        self.completed_type2_project_count = 0
        self.total_person_count = 0
        self.completed_person_count = 0
        self.batch_records: list[dict] = []
        self.batch_reports: list[dict] = []
        self.batch_errors: list[dict] = []
        self.batch_warnings: list[dict] = []
        self.batch_manifest_base: dict = {}
        self.pending_guarantors: list[str] = []
        self.current_guarantor = ""
        self.guarantor_dialog: tk.Toplevel | None = None
        self.guarantor_source_text: tk.Text | None = None
        self.guarantor_status_var = tk.StringVar(value="")
        self.extraction_in_progress = False
        self._path_entries: list[tk.Entry] = []
        saved_input_dir = str(self.saved_settings.get("input_dir") or "").strip()
        old_default_input = (PROJECT_ROOT / "InputPic").resolve()
        try:
            saved_input_path = Path(saved_input_dir).expanduser().resolve()
        except OSError:
            saved_input_path = DEFAULT_INPUT_DIR.resolve()
        if not saved_input_dir or saved_input_path == old_default_input:
            saved_input_path = DEFAULT_INPUT_DIR.resolve()
        self.input_dir_var = tk.StringVar(value=str(saved_input_path))
        self.output_var = tk.StringVar(
            value=str(self.saved_settings.get("output_dir") or DEFAULT_OUTPUT)
        )
        self.document_type_var = tk.StringVar(
            value=str(self.saved_settings.get("document_type") or "1")
        )
        self.model_var = tk.StringVar(
            value=str(self.saved_settings.get("model") or "gemini-3.5-flash")
        )
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE))
        self.prompt_var = tk.StringVar(value=str(DEFAULT_PROMPT))
        self.api_key_var = tk.StringVar(value=self._read_saved_api_key())
        self.show_key_var = tk.BooleanVar(value=False)
        self.save_key_var = tk.BooleanVar(
            value=bool(self.saved_settings.get("save_api_key", True))
        )
        self.status_var = tk.StringVar(value="就绪")

        self.root.title("价值分析报告生成器")
        self.root.geometry("980x850")
        self.root.minsize(900, 780)
        self.root.configure(bg=WINDOW_BG)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        if APP_ICON_FILE.is_file():
            try:
                self.root.iconbitmap(default=str(APP_ICON_FILE))
            except Exception:
                pass

        self.main = tk.Frame(self.root, bg=WINDOW_BG, padx=28, pady=22)
        self.main.grid(row=0, column=0, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(5, weight=1)

        self._build_ui()
        self._on_document_type_changed()
        self._register_drop_target()
        self._bind_setting_persistence()
        self._settings_ready = True
        self._refresh_state()

    def _read_saved_api_key(self) -> str:
        if not bool(self.saved_settings.get("save_api_key", True)):
            return ""
        try:
            return API_KEY_FILE.read_text(encoding="utf-8-sig").strip()
        except OSError:
            return ""

    def _load_panel_settings(self) -> dict:
        try:
            payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8-sig"))
            return payload if isinstance(payload, dict) else {}
        except (OSError, ValueError):
            return {}

    def _bind_setting_persistence(self) -> None:
        for variable in (
            self.input_dir_var,
            self.output_var,
            self.document_type_var,
            self.model_var,
            self.api_key_var,
            self.save_key_var,
        ):
            variable.trace_add("write", self._schedule_settings_save)

    def _schedule_settings_save(self, *_args: object) -> None:
        if not self._settings_ready:
            return
        if self._settings_save_after_id is not None:
            self.root.after_cancel(self._settings_save_after_id)
        self._settings_save_after_id = self.root.after(
            300,
            self._save_panel_settings,
        )

    def _save_panel_settings(self) -> None:
        self._settings_save_after_id = None
        payload = {
            "schema_version": 1,
            "input_dir": self.input_dir_var.get().strip(),
            "output_dir": self.output_var.get().strip(),
            "document_type": self.document_type_var.get(),
            "model": self.model_var.get(),
            "save_api_key": bool(self.save_key_var.get()),
        }
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            dir=SETTINGS_FILE.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        try:
            temp_path.replace(SETTINGS_FILE)
        finally:
            if temp_path.exists():
                temp_path.unlink()
        if self.save_key_var.get() and self.api_key_var.get().strip():
            API_KEY_FILE.write_text(
                self.api_key_var.get().strip(),
                encoding="utf-8",
            )

    def _apply_document_type(self) -> None:
        config = DOCUMENT_TYPE_CONFIG.get(self.document_type_var.get())
        if config is None:
            self.template_var.set("")
            self.prompt_var.set("")
            return
        self.template_var.set(str(config["template"]))
        self.prompt_var.set(str(config["prompt"]))

    def _build_ui(self) -> None:
        tk.Label(
            self.main,
            text="价值分析报告生成器",
            font=("Microsoft YaHei UI", 21, "bold"),
            bg=WINDOW_BG,
            fg=TEXT_DARK,
        ).grid(row=0, column=0, sticky="w", pady=(0, 14))

        key_panel = tk.Frame(
            self.main,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
            padx=20,
            pady=14,
        )
        key_panel.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        key_panel.columnconfigure(1, weight=1)
        tk.Label(
            key_panel,
            text="Gemini API Key",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        self.api_key_entry = tk.Entry(
            key_panel,
            textvariable=self.api_key_var,
            show="*",
            font=("Consolas", 10),
            bg=INPUT_BG,
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            bd=1,
        )
        self.api_key_entry.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(16, 10),
            ipady=7,
        )
        self.show_key_check = tk.Checkbutton(
            key_panel,
            text="显示",
            variable=self.show_key_var,
            command=self._toggle_key_visibility,
            font=("Microsoft YaHei UI", 9),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            activebackground=PANEL_BG,
            bd=0,
            highlightthickness=0,
        )
        self.show_key_check.grid(row=0, column=2, padx=(0, 8))
        self.save_key_check = tk.Checkbutton(
            key_panel,
            text="保存到本机",
            variable=self.save_key_var,
            font=("Microsoft YaHei UI", 9),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            activebackground=PANEL_BG,
            bd=0,
            highlightthickness=0,
        )
        self.save_key_check.grid(row=0, column=3)

        settings = tk.Frame(
            self.main,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
            padx=20,
            pady=16,
        )
        settings.grid(row=2, column=0, sticky="ew")
        settings.columnconfigure(1, weight=1)

        self.input_dir_label = self._path_row(
            settings,
            0,
            "输入根目录",
            self.input_dir_var,
            self._select_input_dir,
        )
        self.input_dir_entry = self._path_entries[-1]
        self._path_row(
            settings,
            1,
            "输出目录",
            self.output_var,
            self._select_output,
        )

        tk.Label(
            settings,
            text="文档类型",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=(14, 0))
        type_frame = tk.Frame(settings, bg=PANEL_BG)
        type_frame.grid(row=2, column=1, sticky="w", pady=(14, 0))
        for text, value in (("1", "1"), ("2", "2")):
            tk.Radiobutton(
                type_frame,
                text=text,
                value=value,
                variable=self.document_type_var,
                command=self._on_document_type_changed,
                font=("Microsoft YaHei UI", 10),
                bg=PANEL_BG,
                fg=TEXT_DARK,
                activebackground=PANEL_BG,
                selectcolor=INPUT_BG,
                bd=0,
                highlightthickness=0,
                padx=5,
                cursor="hand2",
            ).pack(side="left", padx=(0, 10))
        self.type_note = tk.Label(
            type_frame,
            text="类型 1：当前价值分析报告",
            font=("Microsoft YaHei UI", 9),
            bg=PANEL_BG,
            fg=TEXT_MID,
        )
        self.type_note.pack(side="left", padx=(8, 0))

        tk.Label(
            settings,
            text="Gemini 模型",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.model_combo = ttk.Combobox(
            settings,
            textvariable=self.model_var,
            values=("gemini-3.5-flash", "gemini-3.1-flash-lite"),
            state="readonly",
            font=("Microsoft YaHei UI", 10),
        )
        self.model_combo.grid(row=3, column=1, sticky="ew", pady=(12, 0))

        actions = tk.Frame(self.main, bg=WINDOW_BG)
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 12))
        actions.columnconfigure(0, weight=1)
        left_actions = tk.Frame(actions, bg=WINDOW_BG)
        left_actions.grid(row=0, column=0, sticky="w")
        self.open_output_button = self._button(
            left_actions,
            "打开输出目录",
            self.open_output_folder,
            GRAY_BUTTON,
            GRAY_ACTIVE,
        )
        self.open_output_button.pack(side="left")
        self.clear_log_button = self._button(
            left_actions,
            "清空日志",
            self.clear_log,
            GRAY_BUTTON,
            GRAY_ACTIVE,
        )
        self.clear_log_button.pack(side="left", padx=(10, 0))
        self.run_button = self._button(
            actions,
            "开始生成",
            self.start_generation,
            RUST,
            RUST_ACTIVE,
            width=13,
            pady=10,
        )
        self.run_button.grid(row=0, column=1, sticky="e")

        status_bar = tk.Frame(self.main, bg=WINDOW_BG)
        status_bar.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        status_bar.columnconfigure(0, weight=1)
        tk.Label(
            status_bar,
            text="运行日志",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg=WINDOW_BG,
            fg=TEXT_DARK,
        ).grid(row=0, column=0, sticky="w")
        self.status_label = tk.Label(
            status_bar,
            textvariable=self.status_var,
            font=("Microsoft YaHei UI", 9),
            bg=WINDOW_BG,
            fg=TEXT_MID,
        )
        self.status_label.grid(row=0, column=1, sticky="e")

        log_frame = tk.Frame(
            self.main,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        log_frame.grid(row=5, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            state="disabled",
            font=("Consolas", 10),
            bg="#FAFBF9",
            fg="#263129",
            relief="flat",
            padx=12,
            pady=10,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _path_row(
        self,
        parent: tk.Widget,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Callable[[], None],
    ) -> tk.Label:
        label_widget = tk.Label(
            parent,
            text=label,
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        )
        label_widget.grid(row=row, column=0, sticky="w", pady=5)
        entry = tk.Entry(
            parent,
            textvariable=variable,
            font=("Microsoft YaHei UI", 10),
            bg=INPUT_BG,
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            bd=1,
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(16, 10), pady=5, ipady=6)
        self._path_entries.append(entry)
        self._button(
            parent,
            "选择",
            command,
            GREEN,
            GREEN_ACTIVE,
            padx=14,
            pady=7,
        ).grid(row=row, column=2, pady=5)
        return label_widget

    def _button(
        self,
        parent: tk.Widget,
        text: str,
        command: Callable[[], None],
        color: str,
        active_color: str,
        *,
        width: int | None = None,
        padx: int = 16,
        pady: int = 8,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=color,
            fg="white",
            activebackground=active_color,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=padx,
            pady=pady,
            width=width,
            cursor="hand2",
        )

    def _register_drop_target(self) -> None:
        if not DND_AVAILABLE:
            return
        for widget in (self.root, self.input_dir_entry):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event: object) -> None:
        paths = [Path(item) for item in self.root.tk.splitlist(event.data)]
        directory = next((path for path in paths if path.is_dir()), None)
        if directory:
            self.input_dir_var.set(str(directory.resolve()))

    def _select_input_dir(self) -> None:
        path = filedialog.askdirectory(title="选择 Input 输入根目录")
        if path:
            self.input_dir_var.set(path)

    def _select_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_var.set(path)

    def _toggle_key_visibility(self) -> None:
        self.api_key_entry.configure(show="" if self.show_key_var.get() else "*")

    def _document_input_dir(self, document_type: str | None = None) -> Path:
        selected = Path(self.input_dir_var.get()).expanduser().resolve()
        document_type = document_type or self.document_type_var.get()
        if selected.name in {"1", "2"}:
            return selected.parent / document_type
        return selected / document_type

    def _on_document_type_changed(self) -> None:
        self._apply_document_type()
        if self.document_type_var.get() == "1":
            self.type_note.configure(text="类型 1：当前价值分析报告", fg=TEXT_MID)
        else:
            self.type_note.configure(
                text="类型 2：企业价值评估报告及评估说明",
                fg=TEXT_MID,
            )
        self.input_dir_label.configure(text="输入根目录")
        self._refresh_state()

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, text.rstrip() + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def open_output_folder(self) -> None:
        output = Path(self.output_var.get()).expanduser()
        output.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(output)
        except OSError as exc:
            messagebox.showerror("打开失败", f"无法打开输出目录：\n{output}\n\n{exc}")

    def _validate_inputs(self) -> tuple[bool, str]:
        document_type = self.document_type_var.get()
        config = DOCUMENT_TYPE_CONFIG.get(document_type)
        if config is None:
            return False, f"未知文档类型：{document_type}"
        for label, raw_path in (
            ("Word 模板", self.template_var.get()),
            ("Prompt", self.prompt_var.get()),
            ("规则文件", str(config["rules"])),
        ):
            if not Path(raw_path).expanduser().is_file():
                return False, f"{label}不存在：\n{raw_path}"
        if document_type == "2":
            explanation_template = Path(config["explanation_template"])
            if not explanation_template.is_file():
                return False, f"评估说明模板不存在：\n{explanation_template}"
        input_root = Path(self.input_dir_var.get()).expanduser()
        if not input_root.is_dir():
            return False, f"输入根目录不存在：\n{self.input_dir_var.get()}"
        input_dir = self._document_input_dir(document_type)
        if not input_dir.is_dir():
            return False, (
                f"文档类型 {document_type} 的输入目录不存在：\n"
                f"{input_dir}"
            )
        if document_type == "1":
            try:
                folders = list_person_folders(input_dir.resolve())
                for folder in folders:
                    image_pipeline.find_input_workbook(
                        folder,
                        fallback_dir=input_dir.resolve(),
                    )
            except Exception as exc:
                return False, str(exc)
        else:
            try:
                type2_pipeline.discover_project_folders(input_dir.resolve())
            except Exception as exc:
                return False, str(exc)
        if not self.api_key_var.get().strip():
            return False, "请填写 Gemini API Key。"
        return True, ""

    def _format_selectable_input_item(
        self,
        input_dir: Path,
        folder: Path,
    ) -> str:
        try:
            return str(folder.resolve().relative_to(input_dir.resolve()))
        except ValueError:
            return folder.name

    def _choose_processing_items(
        self,
        document_type: str,
        folders: list[Path],
    ) -> list[Path] | None:
        if not folders:
            return []

        input_dir = self._document_input_dir(document_type)
        dialog = tk.Toplevel(self.root)
        dialog.title(f"选择类型 {document_type} 处理项目")
        dialog.geometry("760x560")
        dialog.minsize(640, 420)
        dialog.configure(bg=WINDOW_BG)
        dialog.transient(self.root)
        dialog.grab_set()

        result: dict[str, list[Path] | None] = {"folders": None}
        selected = set(range(len(folders)))
        row_height = 34
        drag_state: dict[str, object] = {
            "start": None,
            "rect": None,
            "dragging": False,
        }

        container = tk.Frame(dialog, bg=WINDOW_BG, padx=22, pady=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        tk.Label(
            container,
            text="选择本次要处理的数据文件夹",
            font=("Microsoft YaHei UI", 13, "bold"),
            bg=WINDOW_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        note_var = tk.StringVar()
        tk.Label(
            container,
            textvariable=note_var,
            font=("Microsoft YaHei UI", 9),
            bg=WINDOW_BG,
            fg=TEXT_MID,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(4, 12))

        list_frame = tk.Frame(
            container,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        list_frame.grid(row=2, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        canvas = tk.Canvas(
            list_frame,
            bg=INPUT_BG,
            highlightthickness=0,
            bd=0,
            selectbackground=GREEN,
        )
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = tk.Scrollbar(list_frame, command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        def update_note() -> None:
            note_var.set(
                f"已选择 {len(selected)} / {len(folders)} 项。"
                "按住鼠标左键拖动可框选。"
            )

        def redraw() -> None:
            canvas.delete("row")
            width = max(canvas.winfo_width(), 620)
            height = max(len(folders) * row_height, canvas.winfo_height())
            canvas.configure(scrollregion=(0, 0, width, height))
            for index, folder in enumerate(folders):
                y = index * row_height
                is_selected = index in selected
                fill = "#EAF2EC" if is_selected else INPUT_BG
                canvas.create_rectangle(
                    0,
                    y,
                    width,
                    y + row_height,
                    fill=fill,
                    outline=BORDER,
                    tags=("row",),
                )
                box_x = 12
                box_y = y + 9
                canvas.create_rectangle(
                    box_x,
                    box_y,
                    box_x + 15,
                    box_y + 15,
                    outline=GREEN,
                    width=1,
                    fill=PANEL_BG,
                    tags=("row",),
                )
                if is_selected:
                    canvas.create_line(
                        box_x + 3,
                        box_y + 8,
                        box_x + 7,
                        box_y + 12,
                        box_x + 13,
                        box_y + 4,
                        fill=GREEN,
                        width=2,
                        tags=("row",),
                    )
                display = self._format_selectable_input_item(input_dir, folder)
                canvas.create_text(
                    38,
                    y + row_height / 2,
                    text=f"{index + 1:02d}  {display}",
                    anchor="w",
                    fill=TEXT_DARK,
                    font=("Microsoft YaHei UI", 10),
                    tags=("row",),
                )
            update_note()

        def set_all(value: bool) -> None:
            selected.clear()
            if value:
                selected.update(range(len(folders)))
            redraw()

        def row_at(y_value: float) -> int | None:
            index = int(y_value // row_height)
            if 0 <= index < len(folders):
                return index
            return None

        def on_press(event: tk.Event) -> None:
            drag_state["start"] = (
                canvas.canvasx(event.x),
                canvas.canvasy(event.y),
            )
            drag_state["dragging"] = False
            rect = drag_state.get("rect")
            if rect is not None:
                canvas.delete(rect)
                drag_state["rect"] = None

        def on_motion(event: tk.Event) -> None:
            start = drag_state.get("start")
            if not isinstance(start, tuple):
                return
            x0, y0 = start
            x1 = canvas.canvasx(event.x)
            y1 = canvas.canvasy(event.y)
            if abs(x1 - x0) < 4 and abs(y1 - y0) < 4:
                return
            drag_state["dragging"] = True
            rect = drag_state.get("rect")
            if rect is None:
                rect = canvas.create_rectangle(
                    x0,
                    y0,
                    x1,
                    y1,
                    outline=RUST,
                    dash=(4, 3),
                    width=2,
                )
                drag_state["rect"] = rect
            else:
                canvas.coords(rect, x0, y0, x1, y1)

        def on_release(event: tk.Event) -> None:
            start = drag_state.get("start")
            if not isinstance(start, tuple):
                return
            x0, y0 = start
            x1 = canvas.canvasx(event.x)
            y1 = canvas.canvasy(event.y)
            rect = drag_state.get("rect")
            if rect is not None:
                canvas.delete(rect)
                drag_state["rect"] = None
            if drag_state.get("dragging"):
                top, bottom = sorted((y0, y1))
                boxed = {
                    index
                    for index in range(len(folders))
                    if index * row_height <= bottom
                    and (index + 1) * row_height >= top
                }
                if event.state & 0x0004:
                    selected.update(boxed)
                else:
                    selected.clear()
                    selected.update(boxed)
            else:
                index = row_at(canvas.canvasy(event.y))
                if index is not None:
                    if index in selected:
                        selected.remove(index)
                    else:
                        selected.add(index)
            drag_state["start"] = None
            drag_state["dragging"] = False
            redraw()

        def on_mousewheel(event: tk.Event) -> None:
            delta = -1 if event.delta > 0 else 1
            canvas.yview_scroll(delta * 3, "units")

        def confirm() -> None:
            if not selected:
                messagebox.showwarning(
                    "未选择项目",
                    "请至少选择一个要处理的数据文件夹。",
                    parent=dialog,
                )
                return
            result["folders"] = [folders[index] for index in sorted(selected)]
            dialog.destroy()

        def cancel() -> None:
            result["folders"] = None
            dialog.destroy()

        canvas.bind("<Configure>", lambda _event: redraw())
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_motion)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<MouseWheel>", on_mousewheel)
        dialog.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: cancel())
        dialog.protocol("WM_DELETE_WINDOW", cancel)

        actions = tk.Frame(container, bg=WINDOW_BG)
        actions.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        actions.columnconfigure(0, weight=1)
        left_actions = tk.Frame(actions, bg=WINDOW_BG)
        left_actions.grid(row=0, column=0, sticky="w")
        self._button(
            left_actions,
            "全选",
            lambda: set_all(True),
            GRAY_BUTTON,
            GRAY_ACTIVE,
            padx=14,
            pady=7,
        ).pack(side="left")
        self._button(
            left_actions,
            "取消全选",
            lambda: set_all(False),
            GRAY_BUTTON,
            GRAY_ACTIVE,
            padx=14,
            pady=7,
        ).pack(side="left", padx=(10, 0))
        self._button(
            actions,
            "取消",
            cancel,
            GRAY_BUTTON,
            GRAY_ACTIVE,
            padx=14,
            pady=7,
        ).grid(row=0, column=1, padx=(0, 10), sticky="e")
        self._button(
            actions,
            "开始处理",
            confirm,
            RUST,
            RUST_ACTIVE,
            padx=18,
            pady=7,
        ).grid(row=0, column=2, sticky="e")

        redraw()
        dialog.wait_window()
        return result["folders"]

    def _save_api_key(self) -> None:
        if self.save_key_var.get():
            API_KEY_FILE.write_text(self.api_key_var.get().strip(), encoding="utf-8")

    def _load_rules(self) -> dict:
        rules = json.loads(DEFAULT_RULES.read_text(encoding="utf-8-sig"))
        try:
            saved_payload = json.loads(
                GUARANTOR_STORE_FILE.read_text(encoding="utf-8-sig")
            )
        except (OSError, ValueError):
            saved_payload = {}
        saved_guarantors = saved_payload.get("guarantors")
        if isinstance(saved_guarantors, dict):
            rules["guarantors"] = saved_guarantors
        return rules

    def _save_rules(self, rules: dict) -> None:
        payload = {
            "schema_version": 1,
            "guarantors": rules.get("guarantors", {}),
        }
        GUARANTOR_STORE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            dir=GUARANTOR_STORE_FILE.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        try:
            temp_path.replace(GUARANTOR_STORE_FILE)
        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _find_guarantors(
        self,
        records_file: Path = CURRENT_RECORDS_FILE,
    ) -> list[str]:
        payload = json.loads(
            records_file.read_text(encoding="utf-8-sig")
        )
        guarantors: list[str] = []
        for item in payload.get("records", []):
            if not isinstance(item, dict):
                continue
            data = item.get("data") or {}
            guarantor = str(data.get("保证人") or "").strip()
            if guarantor and guarantor not in guarantors:
                guarantors.append(guarantor)
        return guarantors

    def _format_saved_guarantor(self, details: dict) -> str:
        fields = (
            "保证人",
            "统一社会信用代码",
            "类型",
            "法定代表人",
            "成立日期",
            "营业场所",
            "经营范围",
        )
        return "\n".join(
            f"{field}：{details.get(field, '')}"
            for field in fields
            if details.get(field)
        )

    def _has_complete_guarantor(self, details: dict) -> bool:
        return all(
            str(details.get(field, "")).strip()
            for field in (
                "保证人",
                "统一社会信用代码",
                "类型",
                "法定代表人",
                "成立日期",
                "营业场所",
                "经营范围",
            )
        )

    def _copy_guarantor_name(self) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(self.current_guarantor)
        self.root.update()
        self.guarantor_status_var.set("保证人名称已复制")

    def _show_next_guarantor_dialog(self) -> None:
        if not self.pending_guarantors:
            self._launch_generation()
            return
        self.current_guarantor = self.pending_guarantors.pop(0)
        rules = self._load_rules()
        saved_details = rules.get("guarantors", {}).get(self.current_guarantor, {})

        dialog = tk.Toplevel(self.root)
        self.guarantor_dialog = dialog
        dialog.title("保证人资料")
        dialog.geometry("840x650")
        dialog.minsize(720, 560)
        dialog.configure(bg=WINDOW_BG)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(2, weight=1)
        dialog.protocol("WM_DELETE_WINDOW", self._cancel_guarantor_workflow)

        name_panel = tk.Frame(
            dialog,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
            padx=18,
            pady=14,
        )
        name_panel.grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 12))
        name_panel.columnconfigure(1, weight=1)
        tk.Label(
            name_panel,
            text="保证人",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
        ).grid(row=0, column=0, sticky="w")
        name_var = tk.StringVar(value=self.current_guarantor)
        name_entry = tk.Entry(
            name_panel,
            textvariable=name_var,
            state="readonly",
            readonlybackground=INPUT_BG,
            fg=TEXT_DARK,
            font=("Microsoft YaHei UI", 11),
            relief="solid",
            bd=1,
        )
        name_entry.grid(row=0, column=1, sticky="ew", padx=(14, 10), ipady=7)
        self.copy_guarantor_button = self._button(
            name_panel,
            "复制",
            self._copy_guarantor_name,
            GREEN,
            GREEN_ACTIVE,
            padx=14,
            pady=7,
        )
        self.copy_guarantor_button.grid(row=0, column=2)

        tk.Label(
            dialog,
            text="粘贴保证人企业资料",
            font=("Microsoft YaHei UI", 11, "bold"),
            bg=WINDOW_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 8))

        text_panel = tk.Frame(
            dialog,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        text_panel.grid(row=2, column=0, sticky="nsew", padx=22)
        text_panel.columnconfigure(0, weight=1)
        text_panel.rowconfigure(0, weight=1)
        self.guarantor_source_text = tk.Text(
            text_panel,
            wrap="word",
            font=("Microsoft YaHei UI", 10),
            bg=INPUT_BG,
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="flat",
            padx=12,
            pady=10,
            undo=True,
        )
        self.guarantor_source_text.grid(row=0, column=0, sticky="nsew")
        text_scrollbar = tk.Scrollbar(
            text_panel,
            command=self.guarantor_source_text.yview,
        )
        text_scrollbar.grid(row=0, column=1, sticky="ns")
        self.guarantor_source_text.configure(yscrollcommand=text_scrollbar.set)
        if saved_details:
            self.guarantor_source_text.insert(
                "1.0",
                self._format_saved_guarantor(saved_details),
            )

        footer = tk.Frame(dialog, bg=WINDOW_BG)
        footer.grid(row=3, column=0, sticky="ew", padx=22, pady=18)
        footer.columnconfigure(0, weight=1)
        tk.Label(
            footer,
            textvariable=self.guarantor_status_var,
            font=("Microsoft YaHei UI", 9),
            bg=WINDOW_BG,
            fg=TEXT_MID,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        self.cancel_guarantor_button = self._button(
            footer,
            "取消生成",
            self._cancel_guarantor_workflow,
            GRAY_BUTTON,
            GRAY_ACTIVE,
            padx=14,
            pady=8,
        )
        self.cancel_guarantor_button.grid(row=0, column=1, padx=(10, 0))
        self.use_saved_button = self._button(
            footer,
            "使用已保存信息",
            self._use_saved_guarantor,
            GRAY_BUTTON,
            GRAY_ACTIVE,
            padx=14,
            pady=8,
        )
        self.use_saved_button.grid(row=0, column=2, padx=(10, 0))
        if not self._has_complete_guarantor(saved_details):
            self.use_saved_button.configure(state="disabled")
        self.extract_guarantor_button = self._button(
            footer,
            "提取并保存",
            self._start_guarantor_extraction,
            RUST,
            RUST_ACTIVE,
            padx=18,
            pady=8,
        )
        self.extract_guarantor_button.grid(row=0, column=3, padx=(10, 0))

        if self._has_complete_guarantor(saved_details):
            status_text = "已载入完整的保存信息，可直接使用"
        elif saved_details:
            status_text = "保存的信息不完整，请补充资料后重新提取"
        else:
            status_text = "请粘贴查询到的企业资料"
        self.guarantor_status_var.set(status_text)
        self.guarantor_source_text.focus_set()

    def _set_guarantor_dialog_state(self, busy: bool) -> None:
        self.extraction_in_progress = busy
        state = "disabled" if busy else "normal"
        if self.guarantor_source_text is not None:
            self.guarantor_source_text.configure(state=state)
        for button_name in (
            "copy_guarantor_button",
            "cancel_guarantor_button",
            "extract_guarantor_button",
        ):
            button = getattr(self, button_name, None)
            if button is not None:
                button.configure(state=state)
        if getattr(self, "use_saved_button", None) is not None:
            rules = self._load_rules()
            has_saved = self._has_complete_guarantor(
                rules.get("guarantors", {}).get(self.current_guarantor, {})
            )
            self.use_saved_button.configure(
                state="disabled" if busy or not has_saved else "normal"
            )

    def _start_guarantor_extraction(self) -> None:
        if self.extraction_in_progress or self.guarantor_source_text is None:
            return
        source_text = self.guarantor_source_text.get("1.0", tk.END).strip()
        if not source_text:
            messagebox.showerror(
                "缺少资料",
                "请先把查询到的保证人企业资料粘贴到大文字框中。",
                parent=self.guarantor_dialog,
            )
            return
        self._set_guarantor_dialog_state(True)
        self.guarantor_status_var.set("Gemini 正在提取字段...")
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get()
        guarantor = self.current_guarantor
        rules = self._load_rules()
        ai_config = rules.get("ai", {})
        thinking_level = str(ai_config.get("thinking_level") or "medium")
        store_interactions = bool(ai_config.get("store_interactions", False))
        worker = threading.Thread(
            target=self._guarantor_extraction_worker,
            args=(
                api_key,
                model,
                thinking_level,
                guarantor,
                source_text,
                store_interactions,
            ),
            daemon=True,
        )
        worker.start()

    def _guarantor_extraction_worker(
        self,
        api_key: str,
        model: str,
        thinking_level: str,
        guarantor: str,
        source_text: str,
        store_interactions: bool,
    ) -> None:
        try:
            audit = extract_guarantor_details_from_text(
                client=create_gemini_client(api_key),
                model=model,
                thinking_level=thinking_level,
                guarantor_name=guarantor,
                source_text=source_text,
                store_interactions=store_interactions,
            )
            self.root.after(0, self._finish_guarantor_extraction, audit)
        except Exception as exc:
            self.root.after(0, self._fail_guarantor_extraction, str(exc))

    def _finish_guarantor_extraction(self, audit: object) -> None:
        self._set_guarantor_dialog_state(False)
        if audit.warnings:
            self.guarantor_status_var.set(audit.warnings[0])
            messagebox.showwarning(
                "资料不完整",
                audit.warnings[0] + "\n\n请补充原始资料后重新提取。",
                parent=self.guarantor_dialog,
            )
            return
        details = {
            field: audit.guarantor_details.get(field, "")
            for field in (
                "保证人",
                "统一社会信用代码",
                "类型",
                "法定代表人",
                "成立日期",
                "营业场所",
                "经营范围",
            )
        }
        rules = self._load_rules()
        rules.setdefault("guarantors", {})[self.current_guarantor] = details
        try:
            self._save_rules(rules)
        except OSError as exc:
            self._fail_guarantor_extraction(f"写入规则 JSON 失败: {exc}")
            return
        interaction_note = (
            f" | Gemini interaction: {audit.interaction_id}"
            if audit.interaction_id
            else ""
        )
        self._append_log(
            f"已保存保证人资料：{self.current_guarantor}{interaction_note}"
        )
        self._complete_current_guarantor()

    def _fail_guarantor_extraction(self, error: str) -> None:
        self._set_guarantor_dialog_state(False)
        self.guarantor_status_var.set("提取失败")
        messagebox.showerror(
            "提取失败",
            error,
            parent=self.guarantor_dialog,
        )

    def _use_saved_guarantor(self) -> None:
        if self.extraction_in_progress:
            return
        self._append_log(f"使用已保存保证人资料：{self.current_guarantor}")
        self._complete_current_guarantor()

    def _complete_current_guarantor(self) -> None:
        if self.guarantor_dialog is not None:
            self.guarantor_dialog.grab_release()
            self.guarantor_dialog.destroy()
        self.guarantor_dialog = None
        self.guarantor_source_text = None
        self.current_guarantor = ""
        self.root.after(0, self._show_next_guarantor_dialog)

    def _cancel_guarantor_workflow(self) -> None:
        if self.extraction_in_progress:
            return
        if self.guarantor_dialog is not None:
            self.guarantor_dialog.grab_release()
            self.guarantor_dialog.destroy()
        self.guarantor_dialog = None
        self.guarantor_source_text = None
        self.pending_guarantors.clear()
        self.person_folders.clear()
        self.current_person_folder = None
        self.current_guarantor = ""
        self.processing = False
        self.status_var.set("已取消")
        self._refresh_state()
        self._append_log("已取消生成。")

    def _refresh_state(self) -> None:
        state = "disabled" if self.processing else "normal"
        for entry in self._path_entries:
            entry.configure(state=state)
        self.api_key_entry.configure(state=state)
        self.show_key_check.configure(state=state)
        self.save_key_check.configure(state=state)
        self.model_combo.configure(state="disabled" if self.processing else "readonly")
        self.clear_log_button.configure(state=state)
        self.open_output_button.configure(state=state)
        can_run = not self.processing and self.document_type_var.get() in {"1", "2"}
        self.run_button.configure(
            state="normal" if can_run else "disabled",
            text="生成中..." if self.processing else "开始生成",
        )

    def start_generation(self) -> None:
        if self.processing:
            return
        valid, error = self._validate_inputs()
        if not valid:
            messagebox.showerror("无法开始", error)
            return
        try:
            self._save_api_key()
        except OSError as exc:
            messagebox.showerror("保存密钥失败", str(exc))
            return

        if self.document_type_var.get() == "2":
            self._start_type2_generation()
            return

        try:
            input_dir = self._document_input_dir("1")
            self.person_folders = list_person_folders(input_dir)
            self.input_workbook = None
        except Exception as exc:
            messagebox.showerror("读取图片目录失败", str(exc))
            return
        selected_folders = self._choose_processing_items("1", self.person_folders)
        if selected_folders is None:
            self.status_var.set("已取消选择")
            return
        self.person_folders = selected_folders

        self.processing = True
        self.total_person_count = len(self.person_folders)
        self.completed_person_count = 0
        self.current_person_folder = None
        self.pending_guarantors.clear()
        self.batch_records = []
        self.batch_reports = []
        self.batch_errors = []
        self.batch_warnings = []
        self.batch_manifest_base = {}
        self._write_batch_records()
        self.status_var.set("准备串行生成")
        self._refresh_state()
        self._append_log(
            f"发现人员文件夹: {self.total_person_count}。"
            "将按“图片提取 → Excel 匹配覆盖 → Word 完成 → 下一人”"
            "串行处理。"
        )
        self._start_next_person()

    def _start_type2_generation(self) -> None:
        input_root = self._document_input_dir("2")
        try:
            self.type2_project_folders = (
                type2_pipeline.discover_project_folders(input_root)
            )
        except Exception as exc:
            messagebox.showerror("读取类型 2 输入目录失败", str(exc))
            return
        selected_folders = self._choose_processing_items(
            "2",
            self.type2_project_folders,
        )
        if selected_folders is None:
            self.status_var.set("已取消选择")
            return
        self.type2_project_folders = selected_folders
        self.processing = True
        self.total_type2_project_count = len(self.type2_project_folders)
        self.completed_type2_project_count = 0
        self.current_type2_project_folder = None
        self.batch_reports = []
        self.batch_errors = []
        self.batch_warnings = []
        self.batch_manifest_base = {}
        self.status_var.set("准备串行生成类型 2")
        self._refresh_state()
        self._append_log(
            f"发现类型 2 项目文件夹: {self.total_type2_project_count}。"
            "每个项目完成 Excel 提取、Gemini 生成和两份 Word 后，"
            "再处理下一个项目。"
        )
        self._start_next_type2_project()

    def _start_next_type2_project(self) -> None:
        if not self.type2_project_folders:
            self._finish_type2_batch_success()
            return
        self.current_type2_project_folder = self.type2_project_folders.pop(0)
        project_number = self.completed_type2_project_count + 1
        self.status_var.set(
            f"正在处理类型 2 项目 "
            f"{project_number}/{self.total_type2_project_count}"
        )
        self._append_log(
            f"[{project_number}/{self.total_type2_project_count}] "
            f"开始类型 2 项目：{self.current_type2_project_folder.name}"
        )
        arguments = self._build_type2_arguments(
            self.current_type2_project_folder
        )
        command = [
            str(self._console_python()),
            "-X",
            "utf8",
            str(TYPE2_PROCESS_SCRIPT),
            *arguments,
        ]
        worker = threading.Thread(
            target=self._type2_generation_worker,
            args=(command, arguments, self._build_environment()),
            daemon=True,
        )
        worker.start()

    def _build_type2_arguments(self, project_folder: Path) -> list[str]:
        config = DOCUMENT_TYPE_CONFIG["2"]
        records_dir = JSON_DIR / "类型2项目数据"
        records_name = (
            type2_pipeline.sanitize_filename_part(project_folder.name)
            + ".json"
        )
        return [
            "--input-dir",
            str(project_folder.resolve()),
            "--word-dir",
            str(Path(self.output_var.get()).expanduser().resolve()),
            "--records-file",
            str(records_dir / records_name),
            "--report-template",
            str(Path(config["template"]).resolve()),
            "--explanation-template",
            str(Path(config["explanation_template"]).resolve()),
            "--prompt-file",
            str(Path(config["prompt"]).resolve()),
            "--rules-file",
            str(Path(config["rules"]).resolve()),
            "--method-library",
            str((TYPE2_RESOURCE_DIR / "资产基础法评估方法库.json").resolve()),
            "--model",
            self.model_var.get(),
            "--api-key-file",
            str(API_KEY_FILE),
        ]

    def _type2_generation_worker(
        self,
        command: list[str],
        arguments: list[str],
        environment: dict[str, str],
    ) -> None:
        if IS_FROZEN:
            return_code = self._run_embedded_worker(
                type2_pipeline.main,
                arguments,
            )
            self.root.after(0, self._finish_type2_generation, return_code)
            return
        creation_flags = (
            subprocess.CREATE_NO_WINDOW
            if sys.platform.startswith("win")
            else 0
        )
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                creationflags=creation_flags,
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.root.after(0, self._append_log, line)
            return_code = process.wait()
            self.root.after(0, self._finish_type2_generation, return_code)
        except Exception as exc:
            self.root.after(0, self._finish_generation_error, str(exc))

    def _finish_type2_generation(self, return_code: int) -> None:
        if return_code == 0:
            try:
                self._capture_type2_results()
            except Exception as exc:
                self.status_var.set("汇总类型 2 项目结果失败")
                self._stop_batch()
                messagebox.showerror("汇总失败", str(exc))
                return
            self.completed_type2_project_count += 1
            current_name = (
                self.current_type2_project_folder.name
                if self.current_type2_project_folder is not None
                else "当前项目"
            )
            self._append_log(
                f"类型 2 项目已完成：{current_name} "
                f"({self.completed_type2_project_count}/"
                f"{self.total_type2_project_count})"
            )
            self.current_type2_project_folder = None
            self.status_var.set("等待下一个类型 2 项目")
            self.root.after(
                SERIAL_NEXT_DELAY_MS,
                self._start_next_type2_project,
            )
            return
        self.status_var.set("类型 2 生成失败")
        self._append_log(f"类型 2 生成失败，退出代码：{return_code}")
        self._stop_batch()
        messagebox.showerror("生成失败", "请查看面板中的运行日志。")

    def _capture_type2_results(self) -> None:
        output_dir = Path(self.output_var.get()).expanduser().resolve()
        manifest_file = output_dir / "generation_manifest.json"
        manifest = json.loads(
            manifest_file.read_text(encoding="utf-8-sig")
        )
        project_name = (
            self.current_type2_project_folder.name
            if self.current_type2_project_folder is not None
            else ""
        )
        for report in manifest.get("reports") or []:
            report_item = dict(report)
            report_item["project_folder"] = project_name
            self.batch_reports.append(report_item)
        for warning in manifest.get("warnings") or []:
            self.batch_warnings.append(
                {
                    "project_folder": project_name,
                    "output_status": "warning",
                    "warning": str(warning),
                }
            )
        for error in manifest.get("errors") or []:
            error_item = dict(error)
            error_item["project_folder"] = project_name
            self.batch_errors.append(error_item)
        self._write_type2_batch_manifest()

    def _write_type2_batch_manifest(self) -> None:
        output_dir = Path(self.output_var.get()).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        payload = {
            "document_type": "2",
            "source_type": "report_project_folders",
            "source_root": str(
                self._document_input_dir("2")
            ),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "reports": self.batch_reports,
            "warnings": self.batch_warnings,
            "errors": self.batch_errors,
        }
        (output_dir / "generation_manifest.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def _finish_type2_batch_success(self) -> None:
        self._write_type2_batch_manifest()
        self.processing = False
        self.current_type2_project_folder = None
        self.status_var.set("类型 2 全部生成完成")
        self._refresh_state()
        document_count = len(self.batch_reports)
        self._append_log(
            f"类型 2 全部完成，共处理 "
            f"{self.completed_type2_project_count} 个项目，"
            f"生成 {document_count} 份 Word。"
        )
        messagebox.showinfo(
            "生成完成",
            f"已按顺序处理 {self.completed_type2_project_count} 个项目，"
            f"生成 {document_count} 份文档。\n"
            f"输出目录：\n{self.output_var.get()}",
        )

    def _start_next_person(self) -> None:
        if not self.person_folders:
            self._finish_batch_success()
            return
        self.current_person_folder = self.person_folders.pop(0)
        person_number = self.completed_person_count + 1
        self.status_var.set(
            f"正在处理 {person_number}/{self.total_person_count}"
        )
        self._append_log(
            f"[{person_number}/{self.total_person_count}] "
            f"开始图片提取：{self.current_person_folder.name}"
        )
        command, environment = self._build_extraction_command(
            self.current_person_folder
        )
        worker = threading.Thread(
            target=self._image_extraction_worker,
            args=(command, environment),
            daemon=True,
        )
        worker.start()

    def _console_python(self) -> Path:
        python_executable = Path(sys.executable)
        if python_executable.name.lower() == "pythonw.exe":
            console_python = python_executable.with_name("python.exe")
            if console_python.is_file():
                return console_python
        return python_executable

    def _build_environment(self) -> dict[str, str]:
        environment = os.environ.copy()
        environment["PYTHONUTF8"] = "1"
        environment["PYTHONIOENCODING"] = "utf-8"
        if not self.save_key_var.get():
            environment["GEMINI_API_KEY"] = self.api_key_var.get().strip()
        return environment

    def _build_extraction_command(
        self,
        folder: Path | None = None,
    ) -> tuple[list[str], dict[str, str]]:
        folder = folder or self.current_person_folder
        arguments = self._build_extraction_arguments(folder)
        command = [
            str(self._console_python()),
            "-X",
            "utf8",
            str(IMAGE_PIPELINE_SCRIPT),
            *arguments,
        ]
        return command, self._build_environment()

    def _build_extraction_arguments(
        self,
        folder: Path | None,
    ) -> list[str]:
        arguments = [
            "--input-dir",
            str(self._document_input_dir("1")),
            "--output-file",
            str(CURRENT_RECORDS_FILE),
            "--individual-dir",
            str(DEFAULT_INDIVIDUAL_RECORDS_DIR),
            "--supplements-file",
            str(DEFAULT_SUPPLEMENTS),
            "--rules-file",
            str(DEFAULT_RULES),
            "--model",
            self.model_var.get(),
            "--api-key-file",
            str(API_KEY_FILE),
        ]
        if folder is not None:
            arguments.extend(["--folder", str(folder.resolve())])
        return arguments

    def _image_extraction_worker(
        self,
        command: list[str],
        environment: dict[str, str],
    ) -> None:
        if IS_FROZEN:
            return_code = self._run_embedded_worker(
                image_pipeline.main,
                self._build_extraction_arguments(self.current_person_folder),
            )
            self.root.after(0, self._finish_image_extraction, return_code)
            return
        creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                creationflags=creation_flags,
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.root.after(0, self._append_log, line)
            return_code = process.wait()
            self.root.after(0, self._finish_image_extraction, return_code)
        except Exception as exc:
            self.root.after(0, self._finish_generation_error, str(exc))

    def _finish_image_extraction(self, return_code: int) -> None:
        if return_code != 0:
            self.status_var.set("当前人员图片提取失败")
            error_message = self._get_image_extraction_error_message(
                CURRENT_RECORDS_FILE
            )
            self._append_log(error_message)
            self._stop_batch()
            messagebox.showerror(
                "图片提取未完成",
                error_message,
            )
            return
        try:
            payload = json.loads(
                CURRENT_RECORDS_FILE.read_text(encoding="utf-8-sig")
            )
            records = payload.get("records") or []
            if len(records) != 1:
                raise ValueError(
                    f"当前人员 JSON 应包含 1 条记录，实际为 {len(records)} 条"
                )
            current_record = records[0]
            if current_record.get("status") != "ready":
                missing = current_record.get("missing_fields") or []
                raise ValueError(
                    "当前人员数据不完整: "
                    + ", ".join(str(field) for field in missing)
                )
            self.pending_guarantors = self._find_guarantors(
                CURRENT_RECORDS_FILE
            )
        except Exception as exc:
            self.status_var.set("读取图片 JSON 失败")
            self._stop_batch()
            messagebox.showerror("读取图片 JSON 失败", str(exc))
            return
        self._append_log(
            f"图片 JSON 已生成：{CURRENT_RECORDS_FILE}"
        )
        if self.pending_guarantors:
            self._append_log(
                f"识别到保证人：{self.pending_guarantors[0]}，"
                "请确认资料。"
            )
            self.status_var.set("等待保证人资料")
            self._show_next_guarantor_dialog()
            return
        self._launch_generation()

    def _get_image_extraction_error_message(
        self,
        records_file: Path = CURRENT_RECORDS_FILE,
    ) -> str:
        fallback = (
            "图片提取出现错误或存在缺失字段，请查看运行日志和：\n"
            f"{records_file}"
        )
        try:
            payload = json.loads(
                records_file.read_text(encoding="utf-8-sig")
            )
            errors = payload.get("errors") or []
            first_error = str((errors[0] or {}).get("error") or "").strip()
        except Exception:
            return fallback
        lowered = first_error.lower()
        if (
            "prepayment credits are depleted" in lowered
            or "预付费额度已耗尽" in first_error
        ):
            return (
                "Gemini API 返回 429（too_many_requests / "
                "RESOURCE_EXHAUSTED）。\n\n"
                "API 的详细原因是：预付费额度已耗尽。\n"
                "Google AI Studio 将 RPM、TPM、RPD 和消费额度限制统一"
                "归类为 429，因此控制台显示 too many requests 并不代表"
                "一定是并发请求过多。\n\n"
                "请为当前项目充值，或在面板顶部更换有可用额度的 API Key。"
            )
        if "api key" in lowered and any(
            marker in lowered
            for marker in ("invalid", "expired", "not valid")
        ):
            return "Gemini API Key 无效或已失效，请在面板顶部更换后重试。"
        if first_error:
            return f"图片提取失败：\n{first_error}\n\n详情：{records_file}"
        return fallback

    def _launch_generation(self) -> None:
        current_name = (
            self.current_person_folder.name
            if self.current_person_folder is not None
            else "当前人员"
        )
        self.status_var.set("正在生成当前人员 Word")
        self._append_log(f"开始生成 Word：{current_name}")
        command, environment = self._build_process_command()
        worker = threading.Thread(
            target=self._generation_worker,
            args=(command, environment),
            daemon=True,
        )
        worker.start()

    def _build_process_command(self) -> tuple[list[str], dict[str, str]]:
        arguments = self._build_process_arguments()
        command = [
            str(self._console_python()),
            "-X",
            "utf8",
            str(PROCESS_SCRIPT),
            *arguments,
        ]
        return command, self._build_environment()

    def _build_process_arguments(self) -> list[str]:
        return [
            "--records-file",
            str(CURRENT_RECORDS_FILE),
            "--template-file",
            str(Path(self.template_var.get()).expanduser().resolve()),
            "--prompt-file",
            str(Path(self.prompt_var.get()).expanduser().resolve()),
            "--rules-file",
            str(DEFAULT_RULES),
            "--guarantors-file",
            str(GUARANTOR_STORE_FILE),
            "--word-dir",
            str(Path(self.output_var.get()).expanduser().resolve()),
            "--document-type",
            self.document_type_var.get(),
            "--model",
            self.model_var.get(),
            "--skip-gemini",
            "--api-key-file",
            str(API_KEY_FILE),
        ]

    def _generation_worker(
        self,
        command: list[str],
        environment: dict[str, str],
    ) -> None:
        if IS_FROZEN:
            return_code = self._run_embedded_worker(
                report_pipeline.main,
                self._build_process_arguments(),
            )
            self.root.after(0, self._finish_generation, return_code)
            return
        creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                creationflags=creation_flags,
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.root.after(0, self._append_log, line)
            return_code = process.wait()
            self.root.after(0, self._finish_generation, return_code)
        except Exception as exc:
            self.root.after(0, self._finish_generation_error, str(exc))

    def _run_embedded_worker(
        self,
        worker: Callable[[list[str]], int],
        arguments: list[str],
    ) -> int:
        buffer = io.StringIO()
        previous_api_key = os.environ.get("GEMINI_API_KEY")
        if not self.save_key_var.get():
            os.environ["GEMINI_API_KEY"] = self.api_key_var.get().strip()
        try:
            with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
                return_code = worker(arguments)
        except Exception:
            traceback.print_exc(file=buffer)
            return_code = 1
        finally:
            if previous_api_key is None:
                os.environ.pop("GEMINI_API_KEY", None)
            else:
                os.environ["GEMINI_API_KEY"] = previous_api_key
        output = buffer.getvalue()
        if output:
            for line in output.splitlines():
                self.root.after(0, self._append_log, line)
        return return_code

    def _finish_generation(self, return_code: int) -> None:
        if return_code == 0:
            current_name = (
                self.current_person_folder.name
                if self.current_person_folder is not None
                else "当前人员"
            )
            try:
                self._capture_current_results()
            except Exception as exc:
                self.status_var.set("汇总当前人员结果失败")
                self._stop_batch()
                messagebox.showerror("汇总失败", str(exc))
                return
            self.completed_person_count += 1
            self._append_log(
                f"Word 已完成：{current_name} "
                f"({self.completed_person_count}/{self.total_person_count})"
            )
            self.current_person_folder = None
            self.pending_guarantors.clear()
            self.status_var.set("等待下一人")
            self.root.after(SERIAL_NEXT_DELAY_MS, self._start_next_person)
        else:
            self.status_var.set("生成失败")
            self._append_log(f"生成失败，退出代码：{return_code}")
            self._stop_batch()
            messagebox.showerror("生成失败", "请查看面板中的运行日志。")

    def _capture_current_results(self) -> None:
        records_payload = json.loads(
            CURRENT_RECORDS_FILE.read_text(encoding="utf-8-sig")
        )
        current_records = records_payload.get("records") or []
        if len(current_records) != 1:
            raise ValueError("当前人员提取结果数量不正确")
        self.batch_records.append(current_records[0])

        manifest_file = Path(self.output_var.get()).expanduser().resolve() / "generation_manifest.json"
        manifest = json.loads(manifest_file.read_text(encoding="utf-8-sig"))
        if not self.batch_manifest_base:
            self.batch_manifest_base = {
                key: value
                for key, value in manifest.items()
                if key not in {"reports", "errors"}
            }
        self.batch_reports.extend(manifest.get("reports") or [])
        self.batch_errors.extend(manifest.get("errors") or [])
        self._write_batch_records()
        self._write_batch_manifest()

    def _write_batch_records(self) -> None:
        excel_files = sorted(
            {
                str((record.get("excel") or {}).get("file") or "")
                for record in self.batch_records
                if str((record.get("excel") or {}).get("file") or "")
            }
        )
        payload = {
            "schema_version": "2.0",
            "source_type": "report_folders_with_images_and_excel",
            "source_root": str(
                self._document_input_dir("1")
            ),
            "excel_files": excel_files,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "model": self.model_var.get(),
            "records": self.batch_records,
            "errors": self.batch_errors,
        }
        DEFAULT_RECORDS_FILE.parent.mkdir(parents=True, exist_ok=True)
        DEFAULT_RECORDS_FILE.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def _write_batch_manifest(self) -> None:
        output_dir = Path(self.output_var.get()).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        manifest = dict(self.batch_manifest_base)
        manifest["reports"] = self.batch_reports
        manifest["errors"] = self.batch_errors
        (output_dir / "generation_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def _finish_batch_success(self) -> None:
        self._write_batch_records()
        self._write_batch_manifest()
        self.processing = False
        self.current_person_folder = None
        self.status_var.set("全部生成完成")
        self._refresh_state()
        self._append_log(
            f"全部完成，共生成 {self.completed_person_count} 份 Word。"
        )
        messagebox.showinfo(
            "生成完成",
            f"已按顺序生成 {self.completed_person_count} 份文档。\n"
            f"输出目录：\n{self.output_var.get()}",
        )

    def _stop_batch(self) -> None:
        self.processing = False
        self.person_folders.clear()
        self.type2_project_folders.clear()
        self.current_person_folder = None
        self.current_type2_project_folder = None
        self.pending_guarantors.clear()
        self._refresh_state()

    def _finish_generation_error(self, error: str) -> None:
        self.status_var.set("启动失败")
        self._stop_batch()
        self._append_log(error)
        messagebox.showerror("启动失败", error)


def create_root() -> tk.Tk:
    if DND_AVAILABLE:
        return TkinterDnD.Tk()
    return tk.Tk()


def show_fatal_error(exc: BaseException) -> None:
    error_log = PROJECT_ROOT / "error.log"
    error_log.write_text(traceback.format_exc(), encoding="utf-8")
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "程序启动失败",
        f"{exc}\n\n详细信息已写入：\n{error_log}",
    )
    root.destroy()


def main() -> None:
    enable_windows_dpi_awareness()
    root = create_root()
    apply_tk_scaling(root)
    ReportGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        show_fatal_error(exc)
