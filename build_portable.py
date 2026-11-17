"""Builds dist/PaperDesk (and dist/PaperDesk.zip): a folder that runs on any Windows PC that has
Microsoft Word, with nothing else to install.  It carries its own Python and Node.js.

    python build_portable.py
"""
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "dist" / "PaperDesk"

SKIP_DIRS = {"__pycache__", ".git", "data", "tests", "sources", "tools"}
SKIP_NAMES = {"plan.md", "start.bat", "server.js", "CCE_original_backup.docx", "IARJSET_published_production.docx"}


def ignore(_dir, names):
    return [n for n in names if n in SKIP_DIRS or n in SKIP_NAMES or n.endswith((".pyc", ".log"))
            or n.endswith("-Paper-Format.docx")]


def copy_tree(src, dst):
    shutil.copytree(ROOT / src, OUT / dst, ignore=ignore)


def main():
    if OUT.parent.exists():
        shutil.rmtree(OUT.parent)
    OUT.mkdir(parents=True)

    # --- the app
    for d in ("webapp", "formatter", "templates", "certificate-generator"):
        copy_tree(d, d)
    (OUT / "webapp" / "data").mkdir()
    for f in ("requirements.txt",):
        shutil.copy2(ROOT / f, OUT / f)

    # --- Python: a copy of this PC's installation, without its libraries, then only what the app needs
    base = Path(sys.base_prefix)
    py = OUT / "python"
    py.mkdir()
    for f in base.iterdir():
        if f.is_file() and f.suffix.lower() in (".exe", ".dll") and "uninst" not in f.name.lower():
            shutil.copy2(f, py / f.name)
    shutil.copytree(base / "DLLs", py / "DLLs")
    shutil.copytree(base / "Lib", py / "Lib",
                    ignore=shutil.ignore_patterns("site-packages", "__pycache__", "test", "tests", "idlelib",
                                                  "turtledemo", "lib2to3", "tkinter", "ensurepip", "*.pyc"))
    site = py / "Lib" / "site-packages"
    site.mkdir()
    subprocess.run([str(py / "python.exe"), "-m", "pip", "--version"], capture_output=True)   # (pip is not shipped)
    r = subprocess.run([sys.executable, "-m", "pip", "install", "--target", str(site), "--no-compile",
                        "--only-binary", ":all:", "-r", str(ROOT / "requirements.txt")], text=True)
    if r.returncode:
        sys.exit("Installing the Python packages failed.")

    # --- Node.js: the certificates are made with it
    node = shutil.which("node")
    if not node:
        sys.exit("Node.js was not found on this PC.")
    (OUT / "node").mkdir()
    shutil.copy2(node, OUT / "node" / "node.exe")

    # --- start files
    (OUT / "Start Paper Desk.bat").write_text(START_BAT, encoding="ascii", newline="\r\n")
    (OUT / "READ ME.txt").write_text(README, encoding="ascii", newline="\r\n")

    # --- zip
    zp = OUT.parent / "PaperDesk.zip"
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        for f in OUT.rglob("*"):
            if f.is_file():
                z.write(f, Path("PaperDesk") / f.relative_to(OUT))
    print("Built", OUT, "and", zp, "(%.0f MB)" % (zp.stat().st_size / 1e6))


START_BAT = r"""@echo off
cd /d "%~dp0"
title Paper Desk
reg query "HKCR\Word.Application" >nul 2>&1
if errorlevel 1 (
  echo Microsoft Word was not found on this computer.
  echo Paper Desk needs Word to count pages and make PDFs.
  pause
  exit /b 1
)
echo.
echo   Paper Desk is starting. Keep this window open while you work.
echo   Close this window to stop it.
echo.
start "" cmd /c "ping 127.0.0.1 -n 4 >nul & start http://127.0.0.1:5000"
python\python.exe webapp\app.py
echo.
echo Paper Desk stopped. If it did not start, another copy may already be open.
pause
"""

README = """PAPER DESK
==========
1. Double-click "Start Paper Desk.bat".  A black window opens and then the browser opens the app.
2. Work in the browser.  Keep the black window open.
3. To stop, close the black window.

Needs Microsoft Word on this computer.  Nothing else has to be installed.
Leave this folder where it is; you can put a shortcut to "Start Paper Desk.bat" on the desktop.
"""

if __name__ == "__main__":
    main()
