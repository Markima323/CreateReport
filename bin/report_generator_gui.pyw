from __future__ import annotations

import ctypes
import json
import os
import subprocess
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from process_excel_to_word import (
    create_gemini_client,
    extract_guarantor_details_from_text,
)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False


BIN_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BIN_DIR.parent
PROCESS_SCRIPT = BIN_DIR / "process_excel_to_word.py"
IMAGE_PIPELINE_SCRIPT = BIN_DIR / "image_input_pipeline.py"
DEFAULT_INPUT_DIR = PROJECT_ROOT / "InputPic"
DEFAULT_RECORDS_FILE = PROJECT_ROOT / "json" / "图片提取数据.json"
DEFAULT_INDIVIDUAL_RECORDS_DIR = PROJECT_ROOT / "json" / "人员数据"
DEFAULT_SUPPLEMENTS = BIN_DIR / "template" / "图片输入补充数据.json"
DEFAULT_TEMPLATE = BIN_DIR / "template" / "价值分析报告-自动生成基底模板.docx"
DEFAULT_PROMPT = BIN_DIR / "template" / "价值分析报告自动生成-Prompt.md"
DEFAULT_RULES = BIN_DIR / "template" / "价值分析报告生成规则.json"
DEFAULT_OUTPUT = PROJECT_ROOT / "word"
API_KEY_FILE = PROJECT_ROOT / "gemini_api.txt"

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
        self.processing = False
        self.pending_guarantors: list[str] = []
        self.current_guarantor = ""
        self.guarantor_dialog: tk.Toplevel | None = None
        self.guarantor_source_text: tk.Text | None = None
        self.guarantor_status_var = tk.StringVar(value="")
        self.extraction_in_progress = False
        self._path_entries: list[tk.Entry] = []
        self.input_dir_var = tk.StringVar(value=str(DEFAULT_INPUT_DIR))
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE))
        self.prompt_var = tk.StringVar(value=str(DEFAULT_PROMPT))
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.document_type_var = tk.StringVar(value="1")
        self.model_var = tk.StringVar(value="gemini-3.5-flash")
        self.api_key_var = tk.StringVar(value=self._read_saved_api_key())
        self.show_key_var = tk.BooleanVar(value=False)
        self.save_key_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就绪")

        self.root.title("价值分析报告生成器")
        self.root.geometry("980x850")
        self.root.minsize(900, 780)
        self.root.configure(bg=WINDOW_BG)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.main = tk.Frame(self.root, bg=WINDOW_BG, padx=28, pady=22)
        self.main.grid(row=0, column=0, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(5, weight=1)

        self._build_ui()
        self._register_drop_target()
        self._refresh_state()

    def _read_saved_api_key(self) -> str:
        try:
            return API_KEY_FILE.read_text(encoding="utf-8-sig").strip()
        except OSError:
            return ""

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

        self._path_row(
            settings,
            0,
            "图片数据目录",
            self.input_dir_var,
            self._select_input_dir,
        )
        self.input_dir_entry = self._path_entries[-1]
        self._path_row(
            settings,
            1,
            "Word 模板",
            self.template_var,
            self._select_template,
        )
        self._path_row(
            settings,
            2,
            "Prompt",
            self.prompt_var,
            self._select_prompt,
        )
        self._path_row(
            settings,
            3,
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
        ).grid(row=4, column=0, sticky="w", pady=(14, 0))
        type_frame = tk.Frame(settings, bg=PANEL_BG)
        type_frame.grid(row=4, column=1, sticky="w", pady=(14, 0))
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
        ).grid(row=5, column=0, sticky="w", pady=(12, 0))
        self.model_combo = ttk.Combobox(
            settings,
            textvariable=self.model_var,
            values=("gemini-3.5-flash", "gemini-3.1-flash-lite"),
            state="readonly",
            font=("Microsoft YaHei UI", 10),
        )
        self.model_combo.grid(row=5, column=1, sticky="ew", pady=(12, 0))

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
    ) -> None:
        tk.Label(
            parent,
            text=label,
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=5)
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
        path = filedialog.askdirectory(title="选择图片数据根目录")
        if path:
            self.input_dir_var.set(path)

    def _select_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Word 模板",
            filetypes=[("Word 模板", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.template_var.set(path)

    def _select_prompt(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Prompt",
            filetypes=[("Markdown", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if path:
            self.prompt_var.set(path)

    def _select_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_var.set(path)

    def _toggle_key_visibility(self) -> None:
        self.api_key_entry.configure(show="" if self.show_key_var.get() else "*")

    def _on_document_type_changed(self) -> None:
        if self.document_type_var.get() == "1":
            self.type_note.configure(text="类型 1：当前价值分析报告", fg=TEXT_MID)
        else:
            self.type_note.configure(text="类型 2：尚未配置", fg=ERROR)
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
        if self.document_type_var.get() != "1":
            return False, "文档类型 2 尚未配置模板和生成规则。"
        for label, raw_path in (
            ("Word 模板", self.template_var.get()),
            ("Prompt", self.prompt_var.get()),
        ):
            if not Path(raw_path).expanduser().is_file():
                return False, f"{label}不存在：\n{raw_path}"
        if not Path(self.input_dir_var.get()).expanduser().is_dir():
            return False, f"图片数据目录不存在：\n{self.input_dir_var.get()}"
        if not DEFAULT_RULES.is_file():
            return False, f"规则文件不存在：\n{DEFAULT_RULES}"
        if not self.api_key_var.get().strip():
            return False, "请填写 Gemini API Key。"
        return True, ""

    def _save_api_key(self) -> None:
        if self.save_key_var.get():
            API_KEY_FILE.write_text(self.api_key_var.get().strip(), encoding="utf-8")

    def _load_rules(self) -> dict:
        return json.loads(DEFAULT_RULES.read_text(encoding="utf-8-sig"))

    def _save_rules(self, rules: dict) -> None:
        DEFAULT_RULES.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            dir=DEFAULT_RULES.parent,
            delete=False,
        ) as handle:
            temp_path = Path(handle.name)
            json.dump(rules, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        try:
            temp_path.replace(DEFAULT_RULES)
        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _find_guarantors(self) -> list[str]:
        payload = json.loads(
            DEFAULT_RECORDS_FILE.read_text(encoding="utf-8-sig")
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
        can_run = not self.processing and self.document_type_var.get() == "1"
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

        self.processing = True
        self.status_var.set("正在从图片提取 JSON")
        self._refresh_state()
        self._append_log("开始遍历图片文件夹并调用 Gemini 提取 JSON...")
        command, environment = self._build_extraction_command()
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
        if not self.save_key_var.get():
            environment["GEMINI_API_KEY"] = self.api_key_var.get().strip()
        return environment

    def _build_extraction_command(self) -> tuple[list[str], dict[str, str]]:
        command = [
            str(self._console_python()),
            str(IMAGE_PIPELINE_SCRIPT),
            "--input-dir",
            str(Path(self.input_dir_var.get()).expanduser().resolve()),
            "--output-file",
            str(DEFAULT_RECORDS_FILE),
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
        return command, self._build_environment()

    def _image_extraction_worker(
        self,
        command: list[str],
        environment: dict[str, str],
    ) -> None:
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
            self.processing = False
            self.status_var.set("图片提取不完整")
            self._refresh_state()
            messagebox.showerror(
                "图片提取未完成",
                "图片提取出现错误或存在缺失字段，请查看运行日志和 "
                f"{DEFAULT_RECORDS_FILE}。",
            )
            return
        try:
            self.pending_guarantors = self._find_guarantors()
        except Exception as exc:
            self.processing = False
            self.status_var.set("读取图片 JSON 失败")
            self._refresh_state()
            messagebox.showerror("读取图片 JSON 失败", str(exc))
            return
        self._append_log(f"图片 JSON 已生成：{DEFAULT_RECORDS_FILE}")
        if self.pending_guarantors:
            self._append_log(
                f"识别到 {len(self.pending_guarantors)} 个唯一保证人，"
                "请逐一确认资料。"
            )
            self.status_var.set("等待保证人资料")
            self._show_next_guarantor_dialog()
            return
        self._launch_generation()

    def _launch_generation(self) -> None:
        self.status_var.set("正在生成")
        self._append_log("开始生成文档...")
        command, environment = self._build_process_command()
        worker = threading.Thread(
            target=self._generation_worker,
            args=(command, environment),
            daemon=True,
        )
        worker.start()

    def _build_process_command(self) -> tuple[list[str], dict[str, str]]:
        command = [
            str(self._console_python()),
            str(PROCESS_SCRIPT),
            "--records-file",
            str(DEFAULT_RECORDS_FILE),
            "--template-file",
            str(Path(self.template_var.get()).expanduser().resolve()),
            "--prompt-file",
            str(Path(self.prompt_var.get()).expanduser().resolve()),
            "--rules-file",
            str(DEFAULT_RULES),
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
        return command, self._build_environment()

    def _generation_worker(
        self,
        command: list[str],
        environment: dict[str, str],
    ) -> None:
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

    def _finish_generation(self, return_code: int) -> None:
        self.processing = False
        self._refresh_state()
        if return_code == 0:
            self.status_var.set("生成完成")
            self._append_log("生成完成。")
            messagebox.showinfo("生成完成", f"文档已输出到：\n{self.output_var.get()}")
        else:
            self.status_var.set("生成失败")
            self._append_log(f"生成失败，退出代码：{return_code}")
            messagebox.showerror("生成失败", "请查看面板中的运行日志。")

    def _finish_generation_error(self, error: str) -> None:
        self.processing = False
        self.status_var.set("启动失败")
        self._refresh_state()
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
