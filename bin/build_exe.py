from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


def main() -> int:
    project_root = Path(__file__).resolve().parent.parent
    source_file = project_root / "bin" / "report_generator_gui.pyw"
    icon_file = project_root / "bin" / "with a pen" / "256x256.ico"
    template_file = (
        project_root
        / "bin"
        / "template"
        / "价值分析报告-自动生成基底模板.docx"
    )
    prompt_file = (
        project_root
        / "bin"
        / "template"
        / "价值分析报告自动生成-Prompt.md"
    )
    rules_file = (
        project_root
        / "bin"
        / "template"
        / "价值分析报告生成规则.json"
    )
    output_file = project_root / "dist" / "CreateReport.exe"

    for label, path in (
        ("GUI source", source_file),
        ("ICO file", icon_file),
        ("Word template", template_file),
        ("Prompt file", prompt_file),
        ("Rules file", rules_file),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label} not found: {path}")

    embedded_dir = project_root / "build" / "embedded_resources"
    embedded_dir.mkdir(parents=True, exist_ok=True)
    embedded_template = embedded_dir / "report_template.docx"
    embedded_prompt = embedded_dir / "report_prompt.md"
    embedded_rules = embedded_dir / "report_rules.json"
    shutil.copy2(template_file, embedded_template)
    shutil.copy2(prompt_file, embedded_prompt)
    shutil.copy2(rules_file, embedded_rules)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "CreateReport",
        "--icon",
        str(icon_file),
        "--distpath",
        str(project_root / "dist"),
        "--workpath",
        str(project_root / "build" / "CreateReport"),
        "--specpath",
        str(project_root / "build"),
        "--paths",
        str(project_root / "bin"),
        "--collect-all",
        "tkinterdnd2",
        "--collect-all",
        "openpyxl",
        "--collect-submodules",
        "win32com",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "win32timezone",
        "--add-data",
        f"{embedded_template}{os.pathsep}resources",
        "--add-data",
        f"{embedded_prompt}{os.pathsep}resources",
        "--add-data",
        f"{embedded_rules}{os.pathsep}resources",
        str(source_file),
    ]
    print("Building CreateReport.exe...")
    subprocess.run(command, cwd=project_root, check=True)
    if not output_file.is_file():
        raise FileNotFoundError(f"PyInstaller did not create: {output_file}")
    print(f"EXE created: {output_file}")
    print(f"Embedded ICO: {icon_file}")
    print(f"Embedded Word template: {template_file}")
    print(f"Embedded Prompt: {prompt_file}")
    print(f"Embedded Rules: {rules_file}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
