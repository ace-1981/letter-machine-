"""Build portable LetterGenerator_release (debug console first, then windowed release)."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RELEASE_DIR_NAME = os.environ.get("LG_RELEASE_DIR", "LetterGenerator_release")
RELEASE_DIR = ROOT.parent / RELEASE_DIR_NAME
TEMPLATE_FILES = (
    "תחשיב זכויות אישי.json",
    "תחשיב זכויות אישי.docx",
)
SAMPLE_EXCEL = ROOT / "samples" / "sample_data.xlsx"


def run_pyinstaller(*, console: bool, name: str, clean: bool) -> Path:
    dist_exe = ROOT / "dist" / f"{name}.exe"
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm"]
    if clean:
        cmd.append("--clean")
    cmd.extend(
        [
            "--onefile",
            "--name",
            name,
            "--distpath",
            str(ROOT / "dist"),
            "--workpath",
            str(ROOT / "build"),
            "--specpath",
            str(ROOT),
            "--collect-all",
            "PySide6",
            "--hidden-import",
            "fitz",
            "--hidden-import",
            "docxtpl",
            "--hidden-import",
            "jinja2",
            "--hidden-import",
            "lxml",
            "--hidden-import",
            "pandas",
            "--hidden-import",
            "openpyxl",
            "--hidden-import",
            "xlrd",
            "--hidden-import",
            "pypdf",
            "--hidden-import",
            "win32com",
            "--hidden-import",
            "win32com.client",
            "--hidden-import",
            "docxcompose",
            "--collect-submodules",
            "docxcompose",
        ]
    )
    if console:
        cmd.append("--console")
    else:
        cmd.append("--windowed")
    cmd.append(str(ROOT / "app.py"))

    label = "debug (console)" if console else "release (windowed)"
    print(f"Running PyInstaller — {label}: {name}.exe")
    subprocess.run(cmd, check=True, cwd=ROOT)
    if not dist_exe.is_file():
        raise FileNotFoundError(f"Expected exe not found: {dist_exe}")
    return dist_exe


def _stage_portable(exe_path: Path, target: Path) -> None:
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)
    shutil.copy2(exe_path, target / "app.exe")
    templates_dir = target / "templates"
    templates_dir.mkdir()
    for name in TEMPLATE_FILES:
        src = ROOT / "templates" / name
        if not src.is_file():
            raise FileNotFoundError(f"Missing template: {src}")
        shutil.copy2(src, templates_dir / name)
    (target / "output").mkdir()
    output_readme = ROOT / "release" / "OUTPUT_README.txt"
    if output_readme.is_file():
        shutil.copy2(output_readme, target / "output" / "README.txt")
    readme = ROOT / "release" / "README.txt"
    if "V1.2" in RELEASE_DIR_NAME:
        readme = ROOT / "release" / "README_V1.2.txt"
    elif "V1.1" in RELEASE_DIR_NAME:
        readme = ROOT / "release" / "README_V1.1.txt"
    shutil.copy2(readme, target / "README.txt")


def smoke_test_debug(exe_path: Path) -> None:
    """Verify debug exe reads external templates before building windowed release."""
    if not SAMPLE_EXCEL.is_file():
        raise FileNotFoundError(f"Sample Excel missing: {SAMPLE_EXCEL}")

    with tempfile.TemporaryDirectory(prefix="lg_debug_smoke_") as tmp:
        stage = Path(tmp) / "portable"
        _stage_portable(exe_path, stage)
        shutil.copy2(SAMPLE_EXCEL, stage / "sample_data.xlsx")

        check = subprocess.run(
            [str(stage / "app.exe"), "check"],
            capture_output=True,
            cwd=stage,
            env={**os.environ, "PYTHONUTF8": "1"},
        )
        if check.returncode != 0:
            raise RuntimeError(
                f"Debug smoke test (check) failed:\n"
                f"{check.stdout.decode('utf-8', errors='replace')}\n"
                f"{check.stderr.decode('utf-8', errors='replace')}"
            )

        preview = subprocess.run(
            [
                str(stage / "app.exe"),
                "preview",
                str(stage / "sample_data.xlsx"),
                "--row",
                "1",
            ],
            capture_output=True,
            cwd=stage,
            timeout=180,
            env={**os.environ, "PYTHONUTF8": "1"},
        )
        if preview.returncode != 0:
            raise RuntimeError(
                f"Debug smoke test (preview) failed:\n"
                f"{preview.stdout.decode('utf-8', errors='replace')}\n"
                f"{preview.stderr.decode('utf-8', errors='replace')}"
            )
        pdf = Path(preview.stdout.decode("utf-8", errors="replace").strip())
        if not pdf.is_file():
            raise RuntimeError(f"Debug smoke test: PDF not created: {pdf}")
        print(f"Debug smoke test OK — PDF: {pdf.name}")


def assemble_release(exe_path: Path) -> Path:
    _stage_portable(exe_path, RELEASE_DIR)
    return RELEASE_DIR


def main() -> int:
    verify = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "verify_portable_design.py")],
        cwd=ROOT,
    )
    if verify.returncode != 0:
        print("Aborting build — portable design check failed.")
        return 1

    debug_exe = run_pyinstaller(console=True, name="app_debug", clean=True)
    smoke_test_debug(debug_exe)
    release_exe = run_pyinstaller(console=False, name="app", clean=False)
    release_path = assemble_release(release_exe)
    print(f"Portable release ready: {release_path}")
    print(f"Debug exe (for troubleshooting): {debug_exe}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
