"""Format every paper in inbox/ for one issue.

1. Edit issue.ini (volume, issue, month, year, first paper number, first page).
2. Put the students' .docx files in inbox/.
3. Run:  python run_issue.py
Formatted papers, one report per paper and summary.txt appear in output/.

Papers are taken in file-name order (1.docx, 2.docx, 10.docx ...).  Each gets
the next paper number (for its DOI) and starts on the page after the previous
paper ends.  Page counts come from Microsoft Word, so Word must be installed.
"""
import configparser
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "formatter"))
from ijarcce_format import format_paper, make_meta  # noqa: E402

INBOX, OUTPUT = ROOT / "inbox", ROOT / "output"
PAGES_SCRIPT = ROOT / "formatter" / "word_pages.ps1"


def natural(path):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", path.name)]


def word_page_counts(paths):
    """Return {path: pages} using Word, or {} if Word could not be used."""
    quoted = ",".join("'%s'" % str(p).replace("'", "''") for p in paths)
    script = str(PAGES_SCRIPT).replace("'", "''")
    cmd = "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '%s'))) -Paths %s" % (script, quoted)
    try:
        res = subprocess.run(["powershell", "-NoProfile", "-Command", cmd],
                             capture_output=True, text=True, timeout=900)
    except (OSError, subprocess.TimeoutExpired) as exc:
        print("Could not run Word to count pages:", exc)
        return {}
    counts = {}
    for line in res.stdout.splitlines():
        bits = line.split("\t")
        if len(bits) >= 2 and bits[0].isdigit():
            counts[Path(bits[1]).resolve()] = int(bits[0])
        elif bits[0] == "ERROR":
            print("Word could not open", bits[1], "-", bits[-1])
    return counts


def main():
    cfg = configparser.ConfigParser()
    if not cfg.read(ROOT / "issue.ini", encoding="utf-8"):
        sys.exit("issue.ini not found next to run_issue.py")
    c = cfg["issue"]
    vol, issue, month, year = c["volume"], c["issue"], c["month"], c["year"]
    paper_no = c.getint("first_paper_number")
    first_page = c.getint("first_page")

    INBOX.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)
    papers = sorted((p for p in INBOX.glob("*.docx") if not p.name.startswith("~$")), key=natural)
    if not papers:
        sys.exit("No .docx files in %s" % INBOX)

    jobs = []
    for i, src in enumerate(papers):
        jobs.append({"src": src, "out": OUTPUT / (src.stem + "_formatted.docx"),
                     "paper": paper_no + i, "start": first_page})

    def run_all():
        for j in jobs:
            meta = make_meta(vol, issue, month, year, j["paper"], j["start"])
            try:
                f, rep = format_paper(j["src"], j["out"], meta)
                j.update(doi=meta["DOI"], flags=len(f.flags), title=f.title, error=None)
            except Exception as exc:  # one bad file must not stop the issue
                j.update(doi=meta["DOI"], flags=0, title="", error=str(exc))

    run_all()
    done = [j for j in jobs if not j["error"]]
    counts = word_page_counts([j["out"] for j in done])
    chained = bool(done) and all(j["out"].resolve() in counts for j in done)
    if chained:
        page = first_page
        for j in done:
            j["start"] = page
            j["pages"] = counts[j["out"].resolve()]
            page += j["pages"]
        run_all()   # second pass writes the real first-page numbers

    lines = ["IJARCCE Vol. %s, Issue %s, %s %s" % (vol, issue, month, year), ""]
    if not chained:
        lines += ["WARNING: page counts could not be read from Word, so every paper starts at page "
                  "%d. Set the start page by hand (Layout > Page Number > Format)." % first_page, ""]
    for j in jobs:
        if j["error"]:
            lines.append("%-28s FAILED: %s" % (j["src"].name, j["error"]))
            continue
        pages = "pages %d-%d" % (j["start"], j["start"] + j["pages"] - 1) if chained else "pages ?"
        lines.append("%-28s paper %-4d %-16s flagged %-3d DOI %s" % (
            j["src"].name, j["paper"], pages, j["flags"], j["doi"]))
        lines.append("    %s" % j["title"][:110])
    text = "\n".join(lines) + "\n"
    (OUTPUT / "summary.txt").write_text(text, encoding="utf-8")
    print(text)
    print("Output folder:", OUTPUT)


if __name__ == "__main__":
    main()
