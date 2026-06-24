from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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
DEFAULT_WORKBOOK = PROJECT_ROOT / "excel" / "数据表.xlsx"
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
        self._path_entries: list[tk.Entry] = []
        self.workbook_var = tk.StringVar(value=str(DEFAULT_WORKBOOK))
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
        self.root.geometry("980x790")
        self.root.minsize(900, 720)
        self.root.configure(bg=WINDOW_BG)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.main = tk.Frame(self.root, bg=WINDOW_BG, padx=28, pady=22)
        self.main.grid(row=0, column=0, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(4, weight=1)

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

        settings = tk.Frame(
            self.main,
            bg=PANEL_BG,
            highlightbackground=BORDER,
            highlightthickness=1,
            padx=20,
            pady=16,
        )
        settings.grid(row=1, column=0, sticky="ew")
        settings.columnconfigure(1, weight=1)

        self._path_row(
            settings,
            0,
            "Excel 数据",
            self.workbook_var,
            self._select_workbook,
        )
        self.workbook_entry = self._path_entries[-1]
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

        tk.Label(
            settings,
            text="API Key",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            anchor="w",
        ).grid(row=6, column=0, sticky="w", pady=(12, 0))
        key_frame = tk.Frame(settings, bg=PANEL_BG)
        key_frame.grid(row=6, column=1, sticky="ew", pady=(12, 0))
        key_frame.columnconfigure(0, weight=1)
        self.api_key_entry = tk.Entry(
            key_frame,
            textvariable=self.api_key_var,
            show="*",
            font=("Consolas", 10),
            bg=INPUT_BG,
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            bd=1,
        )
        self.api_key_entry.grid(row=0, column=0, sticky="ew", ipady=6)
        self.show_key_check = tk.Checkbutton(
            key_frame,
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
        self.show_key_check.grid(row=0, column=1, padx=(10, 0))
        self.save_key_check = tk.Checkbutton(
            key_frame,
            text="保存到本机",
            variable=self.save_key_var,
            font=("Microsoft YaHei UI", 9),
            bg=PANEL_BG,
            fg=TEXT_DARK,
            activebackground=PANEL_BG,
            bd=0,
            highlightthickness=0,
        )
        self.save_key_check.grid(row=0, column=2, padx=(8, 0))

        actions = tk.Frame(self.main, bg=WINDOW_BG)
        actions.grid(row=2, column=0, sticky="ew", pady=(16, 12))
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
        status_bar.grid(row=3, column=0, sticky="ew", pady=(0, 8))
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
        log_frame.grid(row=4, column=0, sticky="nsew")
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
        for widget in (self.root, self.workbook_entry):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event: object) -> None:
        paths = [Path(item) for item in self.root.tk.splitlist(event.data)]
        workbook = next((path for path in paths if path.suffix.lower() in {".xlsx", ".xlsm"}), None)
        if workbook:
            self.workbook_var.set(str(workbook.resolve()))

    def _select_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 数据表",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.workbook_var.set(path)

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
            ("Excel 数据", self.workbook_var.get()),
            ("Word 模板", self.template_var.get()),
            ("Prompt", self.prompt_var.get()),
        ):
            if not Path(raw_path).expanduser().is_file():
                return False, f"{label}不存在：\n{raw_path}"
        if not DEFAULT_RULES.is_file():
            return False, f"规则文件不存在：\n{DEFAULT_RULES}"
        if not self.api_key_var.get().strip():
            return False, "请填写 Gemini API Key。"
        return True, ""

    def _save_api_key(self) -> None:
        if self.save_key_var.get():
            API_KEY_FILE.write_text(self.api_key_var.get().strip(), encoding="utf-8")

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
        self.status_var.set("正在生成")
        self._refresh_state()
        self._append_log("开始生成文档...")
        command, environment = self._build_process_command()
        worker = threading.Thread(
            target=self._generation_worker,
            args=(command, environment),
            daemon=True,
        )
        worker.start()

    def _build_process_command(self) -> tuple[list[str], dict[str, str]]:
        python_executable = Path(sys.executable)
        if python_executable.name.lower() == "pythonw.exe":
            console_python = python_executable.with_name("python.exe")
            if console_python.is_file():
                python_executable = console_python
        command = [
            str(python_executable),
            str(PROCESS_SCRIPT),
            "--workbook",
            str(Path(self.workbook_var.get()).expanduser().resolve()),
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
            "--api-key-file",
            str(API_KEY_FILE),
        ]
        environment = os.environ.copy()
        if not self.save_key_var.get():
            environment["GEMINI_API_KEY"] = self.api_key_var.get().strip()
        return command, environment

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
