from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    project_root = Path(__file__).resolve().parent.parent
    source_file = project_root / "bin" / "report_generator_gui.pyw"
    icon_file = project_root / "bin" / "with a pen" / "256x256.ico"
    output_file = project_root / "dist" / "CreateReport.exe"

    for label, path in (
        ("GUI source", source_file),
        ("ICO file", icon_file),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label} not found: {path}")

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
        "--collect-submodules",
        "win32com",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "win32timezone",
        str(source_file),
    ]
    print("Building CreateReport.exe...")
    subprocess.run(command, cwd=project_root, check=True)
    if not output_file.is_file():
        raise FileNotFoundError(f"PyInstaller did not create: {output_file}")
    print(f"EXE created: {output_file}")
    print(f"Embedded ICO: {icon_file}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
