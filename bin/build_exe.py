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
    template_root = project_root / "bin" / "template"
    resource_files = {
        "1": [
            template_root / "1" / "价值分析报告-自动生成基底模板.docx",
            template_root / "1" / "价值分析报告自动生成-Prompt.md",
            template_root / "1" / "价值分析报告生成规则.json",
        ],
        "2": [
            template_root / "2" / "企业价值评估报告-自动生成基底模板.docx",
            template_root / "2" / "企业价值评估说明-自动生成基底模板.docx",
            template_root / "2" / "企业价值评估报告自动生成-Prompt.md",
            template_root / "2" / "企业价值评估报告生成规则.json",
            template_root / "2" / "企业价值评估输入数据格式.json",
            template_root / "2" / "资产基础法评估方法库.json",
            template_root / "2" / "类型2文档生成规律说明.md",
        ],
    }
    output_file = project_root / "dist" / "CreateReport.exe"

    required_files = [
        ("GUI source", source_file),
        ("ICO file", icon_file),
    ]
    required_files.extend(
        (f"Type {document_type} resource", path)
        for document_type, paths in resource_files.items()
        for path in paths
    )
    for label, path in required_files:
        if not path.is_file():
            raise FileNotFoundError(f"{label} not found: {path}")

    embedded_dir = project_root / "build" / "embedded_resources"
    if embedded_dir.exists():
        shutil.rmtree(embedded_dir)
    embedded_dir.mkdir(parents=True)
    for document_type, paths in resource_files.items():
        target_dir = embedded_dir / document_type
        target_dir.mkdir()
        for path in paths:
            shutil.copy2(path, target_dir / path.name)

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
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "win32timezone",
        "--exclude-module",
        "pandas",
        "--exclude-module",
        "numpy",
        "--exclude-module",
        "scipy",
        "--exclude-module",
        "torch",
        "--exclude-module",
        "tensorflow",
        "--exclude-module",
        "matplotlib",
        "--exclude-module",
        "PySide6",
        "--exclude-module",
        "IPython",
        "--exclude-module",
        "sphinx",
        "--exclude-module",
        "pyarrow",
        "--exclude-module",
        "tables",
        "--exclude-module",
        "numba",
        "--exclude-module",
        "dask",
        "--add-data",
        f"{embedded_dir / '1'}{os.pathsep}resources/1",
        "--add-data",
        f"{embedded_dir / '2'}{os.pathsep}resources/2",
        str(source_file),
    ]
    print("Building CreateReport.exe...")
    subprocess.run(command, cwd=project_root, check=True)
    if not output_file.is_file():
        raise FileNotFoundError(f"PyInstaller did not create: {output_file}")
    print(f"EXE created: {output_file}")
    print(f"Embedded ICO: {icon_file}")
    for document_type, paths in resource_files.items():
        for path in paths:
            print(f"Embedded type {document_type} resource: {path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
