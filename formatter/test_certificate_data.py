"""Reads the authors and designations of several real student papers and prints them for checking.
Run:  python formatter/test_certificate_data.py"""
import io
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import ijarcce_format as F  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DESK = Path.home() / "Desktop"
CASES = [
    ROOT / "student_originals" / "1.docx", ROOT / "student_originals" / "2.docx",
    ROOT / "student_originals" / "3.docx", ROOT / "student_originals" / "4.docx",
    DESK / "gokul rajendran.docx", DESK / "te 3 Dhanashri Deshmukh.docx",
    DESK / "te 1 mk7605 srmist.edu.in.docx", DESK / "te 2 ritheeshchittimalla39.docx",
]
for src in CASES:
    if not src.exists():
        print("(missing)", src)
        continue
    with tempfile.TemporaryDirectory() as d:
        meta = F.make_meta("15", "11", "November", "2026", 1, 1)
        fm, _ = F.format_paper(src, Path(d) / "o.docx", meta)
        data = F.certificate_data(fm)
    print("=" * 90)
    print(src.name, "|", data["title"][:80])
    for a in data["authors"]:
        print("   -", a["name"], "||", a["designation"][:100])
    for w in data["warnings"]:
        print("   !", w)
