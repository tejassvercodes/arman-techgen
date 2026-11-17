"""IJARCCE / IARJSET paper formatter - issue workspace (runs on this computer only).

Start with run_website.bat, then open http://127.0.0.1:5000

An "issue" holds the journal details and a list of papers.  Papers are formatted one
after another in the background; DOIs and first-page numbers are chained from paper
to paper automatically, and can be overridden per paper.
"""
import csv
import logging
import logging.handlers
import os
import io
import json
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "formatter"))
from ijarcce_format import (format_paper, make_meta, ptext, certificate_data,  # noqa: E402
                            DOI_PREFIXES, JOURNAL_NAMES)
from core import (JOURNALS, MONTHS, settings_ok, clean_doi, parse_doi, clean_settings,  # noqa: E402
                  auto_doi, describe)

APP_VERSION = "2.0.0"
HERE = Path(__file__).resolve().parent

# ---- configuration (environment variables; the defaults suit one PC used by one person) ----------
HOST = os.environ.get("PAPERDESK_HOST", "127.0.0.1")        # 0.0.0.0 to accept other computers
PORT = int(os.environ.get("PAPERDESK_PORT") or os.environ.get("PORT") or "5000")   # hosts such as Render set PORT
DATA = Path(os.environ.get("PAPERDESK_DATA", HERE / "data"))   # issues, papers, certificates, logs
API_TOKEN = os.environ.get("PAPERDESK_TOKEN", "")            # if set, every request must carry it
ALLOWED_ORIGINS = {o.strip().rstrip("/") for o in os.environ.get("PAPERDESK_ORIGINS", "").split(",") if o.strip()}
MAX_UPLOAD_MB = int(os.environ.get("PAPERDESK_MAX_UPLOAD_MB", "500"))
BATCHES = DATA / "batches"
BATCHES.mkdir(parents=True, exist_ok=True)
DEFAULTS_FILE = DATA / "defaults.json"
PS_PAGES = ROOT / "formatter" / "word_pages.ps1"
PS_CONVERT = ROOT / "formatter" / "word_convert.ps1"

# website journal -> certificate-generator template key
CERT_KEYS = {"ijarcce": "CCE", "iarjset": "SET", "ijireeice": "ICE", "imrjr": "RJR"}
CERT_DIR = ROOT / "certificate-generator"
CERT_CLI = CERT_DIR / "make-certificates.js"
_BUNDLED_NODE = ROOT / "node" / "node.exe"            # the portable folder carries its own Node.js
NODE = str(_BUNDLED_NODE) if _BUNDLED_NODE.exists() else shutil.which("node")
ALLOWED = {".docx", ".doc", ".rtf", ".odt"}

app = Flask(__name__, static_folder=str(HERE / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.json.sort_keys = False            # journals keep their order: IJARCCE, IJIREEICE, IARJSET, IMRJR

(DATA / "logs").mkdir(parents=True, exist_ok=True)
log = logging.getLogger("paperdesk")
log.setLevel(logging.INFO)
if not log.handlers:
    _fh = logging.handlers.RotatingFileHandler(DATA / "logs" / "app.log", maxBytes=2_000_000, backupCount=5,
                                               encoding="utf-8")
    _fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(_fh)
    log.addHandler(logging.StreamHandler())

LOCK = threading.RLock()          # protects the JSON files
WORD_LOCK = threading.Lock()      # Word is used by one job at a time
RUN_Q = queue.Queue()
PENDING = set()


# ----------------------------------------------------------------------------
# storage
# ----------------------------------------------------------------------------

def valid_id(s):
    return bool(re.fullmatch(r"[0-9a-f]{8,32}", s or ""))


def bdir(bid):
    if not valid_id(bid):
        abort(404)
    return BATCHES / bid


def load(bid):
    f = bdir(bid) / "batch.json"
    if not f.exists():
        abort(404)
    with LOCK:
        return json.loads(f.read_text(encoding="utf-8"))


def save(b):
    f = bdir(b["id"]) / "batch.json"
    tmp = f.with_suffix(".tmp")
    with LOCK:
        tmp.write_text(json.dumps(b, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(f)


def mutate(bid, fn):
    with LOCK:
        b = load(bid)
        fn(b)
        save(b)
        return b


def get_paper(b, pid):
    for p in b["papers"]:
        if p["id"] == pid:
            return p
    abort(404)


def load_defaults():
    now = datetime.now()
    d = {"journal": "ijarcce", "mode": "with", "first_doi": "", "volume": "", "issue": "", "month": "",
         "year": str(now.year), "first_paper": "1", "first_page": "1"}
    try:
        pass                              # nothing is remembered between sessions
    except (OSError, ValueError):
        pass
    return d


def label_of(s):
    if not (s.get("volume") and s.get("issue")):
        return "New issue (details not filled in)"
    return "%s Vol. %s, Issue %s, %s %s" % (JOURNALS.get(s["journal"], s["journal"]), s["volume"],
                                            s["issue"], s["month"], s["year"])


def new_batch(settings=None):
    s = load_defaults()
    if settings:
        s.update(settings)
    bid = uuid.uuid4().hex[:16]
    (BATCHES / bid).mkdir()
    b = {"id": bid, "created": time.time(), "settings": s, "papers": []}
    save(b)
    return b


def summary_of(b):
    return {"id": b["id"], "label": label_of(b["settings"]), "count": len(b["papers"]),
            "created": b["created"]}


def list_batches():
    out = []
    for d in BATCHES.iterdir():
        f = d / "batch.json"
        if f.exists():
            try:
                out.append(summary_of(json.loads(f.read_text(encoding="utf-8"))))
            except ValueError:
                pass
    return sorted(out, key=lambda x: -x["created"])


# ----------------------------------------------------------------------------
# Word helpers
# ----------------------------------------------------------------------------

def _ps(script, extra):
    cmd = "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '%s'))) %s" % (
        str(script).replace("'", "''"), extra)
    try:
        res = subprocess.run(["powershell", "-NoProfile", "-Command", cmd],
                             capture_output=True, text=True, timeout=900)
    except (OSError, subprocess.SubprocessError) as exc:
        log.error("Word script %s failed: %s", Path(script).name, exc)
        return ""
    if res.returncode != 0 and res.stderr.strip():
        log.warning("Word script %s: %s", Path(script).name, res.stderr.strip()[:300])
    return res.stdout


def _q(path):
    return "'%s'" % str(path).replace("'", "''")


# ---- document engine -------------------------------------------------------------------------------
# Page counts, PDFs and .doc conversion come from Microsoft Word when it is installed (Windows).  On a
# server without Word (a Linux container in the cloud) LibreOffice does the same jobs.
# PAPERDESK_ENGINE=word|libreoffice forces one; PAPERDESK_SOFFICE points at soffice.
_SOFFICE_PATHS = [r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                  "/usr/bin/soffice", "/usr/bin/libreoffice", "/opt/libreoffice/program/soffice"]
SOFFICE = os.environ.get("PAPERDESK_SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice") \
    or next((p for p in _SOFFICE_PATHS if Path(p).exists()), "")


def engine():
    """'word', 'libreoffice' or '' (nothing to count pages and make PDFs with)."""
    pref = os.environ.get("PAPERDESK_ENGINE", "").lower()
    word = PS_PAGES.exists() and shutil.which("powershell") is not None and _word_installed()
    if pref == "libreoffice":
        return "libreoffice" if SOFFICE else ""
    if word and pref != "libreoffice":
        return "word"
    return "libreoffice" if SOFFICE else ""


def _lo_convert(src, fmt, outdir):
    """Convert one file with LibreOffice; returns the output path or None.  Uses its own profile folder
    so it never fights a LibreOffice the user has open."""
    profile = (DATA / "lo_profile").resolve()
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [SOFFICE, "-env:UserInstallation=" + profile.as_uri(), "--headless", "--norestore", "--nolockcheck",
           "--convert-to", fmt, "--outdir", str(outdir), str(src)]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    except (OSError, subprocess.SubprocessError) as exc:
        log.error("LibreOffice failed on %s: %s", Path(src).name, exc)
        return None
    out = Path(outdir) / (Path(src).stem + "." + fmt.split(":")[0])
    if not out.exists():
        log.warning("LibreOffice made nothing from %s: %s", Path(src).name, (res.stderr or res.stdout).strip()[:300])
        return None
    return out


def _lo_pages(paths, pdf):
    import tempfile
    import pymupdf
    counts = {}
    for p in paths:
        p = Path(p)
        with tempfile.TemporaryDirectory() as d:
            tmp = Path(d) / ("doc" + p.suffix)                # a short plain name: safe for every tool
            shutil.copyfile(p, tmp)
            out = _lo_convert(tmp, "pdf", d)
            if not out:
                continue
            with pymupdf.open(out) as doc:
                counts[p.resolve()] = doc.page_count
            if pdf:
                shutil.copyfile(out, p.with_suffix(".pdf"))
    return counts


def word_pages(paths, pdf=False):
    """{path: pages}; also writes a .pdf next to each file when pdf=True."""
    if not paths:
        return {}
    if engine() == "libreoffice":
        with WORD_LOCK:
            return _lo_pages(paths, pdf)
    extra = "-Paths %s%s" % (",".join(_q(p) for p in paths), " -Pdf" if pdf else "")
    with WORD_LOCK:
        out = _ps(PS_PAGES, extra)
    counts = {}
    for line in out.splitlines():
        bits = line.split("\t")
        if len(bits) >= 2 and bits[0].isdigit():
            counts[Path(bits[1]).resolve()] = int(bits[0])
    return counts


def _pdf_for(path):
    """A PDF of a student's file for the formatter (it takes pictures of tables and equations a PDF
    conversion broke).  Uses the same one-at-a-time Word lock as everything else."""
    import tempfile
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d) / "source.docx"
        shutil.copyfile(path, tmp)
        word_pages([tmp], pdf=True)
        pdf = tmp.with_suffix(".pdf")
        return pdf.read_bytes() if pdf.exists() else None


import ijarcce_format  # noqa: E402
ijarcce_format.PDF_PROVIDER = _pdf_for


def word_convert(src, dst):
    if engine() == "libreoffice":
        import tempfile
        with WORD_LOCK, tempfile.TemporaryDirectory() as d:
            tmp = Path(d) / ("in" + Path(src).suffix)
            shutil.copyfile(src, tmp)
            out = _lo_convert(tmp, "docx:MS Word 2007 XML", d)
            if out:
                shutil.copyfile(out, dst)
            return bool(out) and Path(dst).exists()
    with WORD_LOCK:
        out = _ps(PS_CONVERT, "-InFile %s -OutFile %s" % (_q(src), _q(dst)))
    return "OK" in out and Path(dst).exists()


# ----------------------------------------------------------------------------
# the pipeline: format -> count pages -> chain DOIs / page numbers -> re-format
# ----------------------------------------------------------------------------

def targets(b):
    s = b["settings"]
    running = int(s["first_page"])
    first = int(s["first_paper"])
    out = {}
    for i, p in enumerate(b["papers"]):
        no = first + i
        # "Without DOI": no DOI anywhere, neither in the paper nor on the certificates
        doi = "" if s.get("mode") == "without" else (p.get("doi_manual") or auto_doi(s, no))
        start = p.get("start_manual") or running
        out[p["id"]] = {"paper_no": no, "doi": doi, "start": start}
        if p["status"] == "done" and p.get("pages"):
            running = start + p["pages"]
    return out


def snapshot(b, p, t):
    s = b["settings"]
    return [s["journal"], s["volume"], s["issue"], s["month"], s["year"],
            t["doi"], t["start"], p.get("rev", 0)]


def src_path(b, p):
    d = bdir(b["id"]) / "papers" / p["id"]
    return d / "source.docx"


def out_path(b, p):
    return bdir(b["id"]) / "papers" / p["id"] / "formatted.docx"


def set_paper(bid, pid, **fields):
    def fn(b):
        p = get_paper(b, pid)
        p.update(fields)
    mutate(bid, fn)


def format_one(bid, p, t):
    b = load(bid)
    s = b["settings"]
    pid = p["id"]
    set_paper(bid, pid, status="working", stage="Formatting", error=None)
    try:
        d = bdir(bid) / "papers" / pid
        src = d / "source.docx"
        if not src.exists():
            orig = next(d.glob("original.*"))
            if orig.suffix.lower() == ".docx":
                shutil.copyfile(orig, src)
            else:
                set_paper(bid, pid, stage="Converting to .docx with Word")
                if not word_convert(orig, src):
                    raise RuntimeError("Word could not open this file. Open it in Word and save it as .docx.")
        meta = make_meta(s["volume"], s["issue"], s["month"], s["year"], doi=t["doi"],
                         start_page=t["start"], journal=s["journal"], no_doi=not t["doi"])
        out = d / "formatted.docx"
        fm, _ = format_paper(src, out, meta, journal=s["journal"], report_path=d / "report.txt")
        info = describe(fm)
        flags, found = info["flags"], info["found"]
        for stale in (d / "formatted.pdf", d / "preview"):
            if stale.is_dir():
                shutil.rmtree(stale, ignore_errors=True)
            elif stale.exists():
                stale.unlink()
        cd = certificate_data(fm)
        fields = dict(status="done", stage="", title=fm.title, flags=flags, fixes=fm.fixes,
                      found=found, snap=snapshot(b, p, t), error=None, cert_data_auto=cd)
        if not get_paper(load(bid), pid).get("cert_manual"):
            fields["cert_data"] = cd
        if p.get("reset_pages", True):
            fields["pages"] = None
        set_paper(bid, pid, reset_pages=False, **fields)
    except Exception as exc:
        set_paper(bid, pid, status="error", stage="", error=str(exc)[:300], flags=[], fixes=[],
                  found={}, pages=None, snap=None, reset_pages=True)


def cert_dir(bid, pid):
    return BATCHES / bid / "papers" / pid / "certs"


def cert_snapshot(b, p, t):
    s = b["settings"]
    cd = p.get("cert_data") or {}
    return json.dumps([s["journal"], s["volume"], s["issue"], s["year"], t["doi"], t["paper_no"],
                       cd.get("title"), cd.get("authors")], sort_keys=True, ensure_ascii=False)


def strip_fixed_doi(path):
    """Remove a DOI printed as fixed text on a certificate (IMRJR's template carries the journal's own
    "DOI: 10.17148/IMRJR"), for issues made without DOIs."""
    import zipfile
    pat = re.compile(r"(<w:t(?: [^>]*)?>)\s*DOI\s*:?\s*10\.\d{4,9}/[^<]*(</w:t>)")
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    changed = False
    for n, (info, data) in enumerate(items):
        if info.filename.startswith("word/") and info.filename.endswith(".xml") and b"DOI" in data:
            new = pat.sub(r"\1\2", data.decode("utf8")).encode("utf8")
            if new != data:
                items[n] = (info, new)
                changed = True
    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for info, data in items:
                z.writestr(info, data)


def certificate_fit_levels(key, cd, doi, s):
    """[{level, fits}] for every author's certificate and, last, the group certificate."""
    import cert_fit

    def dedupe_join(values):
        seen = []
        for v in values:
            v = (v or "").strip()
            if v and v not in seen:
                seen.append(v)
        return ", ".join(seen)

    base = {"DOI": doi, "Volume": str(int(s["volume"])), "Issue": str(int(s["issue"])), "Year": s["year"],
            "Month": s["month"]}
    title = cd.get("title", "")
    records = [dict(base, NAME=a["name"], Designation=a.get("designation", ""), PaperTitle=title)
               for a in cd["authors"]]
    group = dict(base, NAME=dedupe_join(r["NAME"] for r in records),
                 Designation=dedupe_join(r["Designation"] for r in records), PaperTitle=title)
    levels, too_long = [], []
    for i, rec in enumerate(records + [group]):
        level, fits = cert_fit.fit_level(key, rec)
        levels.append({"level": level, "fits": fits})
        if not fits:
            who = "The group certificate" if i == len(records) else "Certificate %d (%s)" % (i + 1, rec["NAME"])
            too_long.append("%s: the text is too long for the box even at the smallest size, so the DOI line "
                            "may be cut off. Shorten the designation in the table below." % who)
    return levels, too_long


def make_certs_one(bid, pid, t):
    """Fill the journal's certificate template for every author, then let Word make the PDFs."""
    b = load(bid)
    p = get_paper(b, pid)
    s = b["settings"]
    snap = cert_snapshot(b, p, t)
    set_paper(bid, pid, certs={"status": "working", "files": [], "warnings": [], "error": None})
    d = cert_dir(bid, pid)
    try:
        if not NODE:
            raise RuntimeError("Node.js was not found, so certificates cannot be made. Install Node.js from nodejs.org.")
        cd = p.get("cert_data") or {}
        if not cd.get("authors"):
            raise RuntimeError("No authors were found in the paper, so there is nothing to put on the certificates. "
                               "Add them in the table below.")
        if d.exists():
            shutil.rmtree(d)                                  # our own folder: certificates of an earlier run
        d.mkdir(parents=True)
        req = {"journal": CERT_KEYS[s["journal"]], "title": cd.get("title", ""), "authors": cd["authors"],
               "doi": t["doi"], "outDir": str(d),
               # used by the certificate maker when there is no DOI
               "volume": str(int(s["volume"])), "issue": str(int(s["issue"])), "year": str(s["year"]),
               "fileNumber": "%02d" % int(t["paper_no"])}
        # measure each certificate's text with the real fonts so the Volume and DOI lines always stay
        # inside the box: long titles and designations are set smaller, only as much as needed
        too_long = []
        try:
            req["fitLevels"], too_long = certificate_fit_levels(CERT_KEYS[s["journal"]], cd, t["doi"], s)
        except Exception:                              # fall back to the certificate maker's own estimate
            log.exception("certificate fit measurement failed")
        reqf = d.parent / "certs_request.json"
        reqf.write_text(json.dumps(req, ensure_ascii=False), encoding="utf-8")
        res = subprocess.run([NODE, str(CERT_CLI), str(reqf)], capture_output=True, text=True, encoding="utf-8",
                             timeout=180, cwd=str(CERT_DIR))
        lines = res.stdout.strip().splitlines()
        out = json.loads(lines[-1]) if lines else {}
        if res.returncode != 0 or out.get("error"):
            raise RuntimeError((out.get("error") or res.stderr[-300:] or "The certificate maker failed.")[:400])
        if not t["doi"]:
            for f in out["files"]:
                strip_fixed_doi(d / f["file"])               # "Without DOI": not even the journal's own DOI
        word_pages([d / f["file"] for f in out["files"]], pdf=True)
        entries = [{"stem": f["file"][:-5], "kind": f["kind"], "index": f["index"],
                    "pdf": (d / (f["file"][:-5] + ".pdf")).exists()} for f in out["files"]]
        pub = out.get("publication", {})
        decoded = (str(pub.get("volume")), str(pub.get("issue")), str(pub.get("year")))
        expected = (str(int(s["volume"])), str(int(s["issue"])), str(s["year"]))
        gen = [w for w in out.get("warnings", [])
               if not (w.startswith("Check the DOI:") and decoded == expected)]   # we know the issue: no doubt
        warnings = list(cd.get("warnings") or []) + gen + too_long
        if t["doi"] and decoded != expected:
            warnings.append("The DOI on the certificates reads as volume %s, issue %s, %s, but this issue is "
                            "volume %s, issue %s, %s. Check the DOI." % (decoded + expected))
        if any(not e["pdf"] for e in entries):
            warnings.append("Some certificate PDFs could not be made; the Word files are fine.")
        set_paper(bid, pid, certs={"status": "ok", "files": entries, "warnings": warnings, "error": None},
                  cert_snap=snap)
    except Exception as exc:
        set_paper(bid, pid, certs={"status": "error", "files": [], "warnings": [], "error": str(exc)[:400]},
                  cert_snap=snap)


def pipeline(bid):
    for _ in range(6):
        try:
            b = load(bid)
        except Exception:
            return
        if not settings_ok(b["settings"]):
            return
        t = targets(b)
        todo = []
        for p in b["papers"]:
            if p["status"] == "error":
                continue
            if p["status"] == "queued" or p.get("snap") != snapshot(b, p, t[p["id"]]):
                todo.append((p, t[p["id"]]))
        for p, tt in todo:
            format_one(bid, p, tt)
        b = load(bid)
        need = [p for p in b["papers"] if p["status"] == "done" and p.get("pages") is None]
        if need:
            for p in need:
                set_paper(bid, p["id"], stage="Counting pages")
            counts = word_pages([out_path(b, p) for p in need])
            for p in need:
                n = counts.get(out_path(b, p).resolve())
                set_paper(bid, p["id"], pages=n, stage="",
                          **({} if n else {"status": "error", "error": "Word could not count the pages."}))
        b = load(bid)
        t = targets(b)
        cert_todo = [p["id"] for p in b["papers"]
                     if p["status"] == "done" and p.get("pages")
                     and (p.get("certs") is None
                          or (p["certs"].get("status") != "working"
                              and p.get("cert_snap") != cert_snapshot(b, p, t[p["id"]])))]
        for pid in cert_todo:
            make_certs_one(bid, pid, t[pid])
        if not todo and not need and not cert_todo:
            return


def worker():
    while True:
        bid = RUN_Q.get()
        with LOCK:
            PENDING.discard(bid)
        try:
            pipeline(bid)
        except (FileNotFoundError, HTTPException):
            pass                                              # the issue or paper was deleted meanwhile
        except Exception:  # keep the worker alive whatever happens
            log.exception("pipeline failed for issue %s", bid)


def submit(bid):
    with LOCK:
        if bid in PENDING:
            return
        PENDING.add(bid)
    RUN_Q.put(bid)


# ----------------------------------------------------------------------------
# views
# ----------------------------------------------------------------------------

def view(b):
    t = targets(b) if settings_ok(b["settings"]) else {}
    papers = []
    for p in b["papers"]:
        tt = t.get(p["id"], {})
        pages = p.get("pages")
        papers.append({
            "id": p["id"], "name": p["name"], "status": p["status"], "stage": p.get("stage", ""),
            "error": p.get("error"), "title": p.get("title", ""), "flags": p.get("flags", []),
            "fixes": p.get("fixes", []), "found": p.get("found", {}), "pages": pages,
            "paper_no": tt.get("paper_no"), "doi": tt.get("doi", ""),
            "doi_manual": bool(p.get("doi_manual")), "start": tt.get("start"),
            "start_manual": bool(p.get("start_manual")),
            "end": (tt["start"] + pages - 1) if (pages and tt) else None,
            "certs": p.get("certs"), "cert_data": p.get("cert_data"), "cert_manual": bool(p.get("cert_manual")),
        })
    busy = any(p["status"] in ("queued", "working") or (p["status"] == "done" and p.get("pages") is None)
               or (p["status"] == "done" and (p.get("certs") is None or p["certs"].get("status") == "working"))
               for p in b["papers"]) and settings_ok(b["settings"])
    return {"id": b["id"], "label": label_of(b["settings"]), "settings": b["settings"],
            "ready": settings_ok(b["settings"]), "busy": busy, "papers": papers}


def health():
    return {"word": bool(engine()), "engine": engine() or "none",
            "node": bool(NODE) and CERT_CLI.exists(),
            "templates": {k: Path(v).exists() for k, v in ijarcce_format.JOURNAL_TEMPLATES.items()},
            "preview": _has_pymupdf()}


def _has_pymupdf():
    try:
        import pymupdf  # noqa: F401
        return True
    except ImportError:
        return False


_WORD = None


def _word_installed():
    global _WORD
    if _WORD is None:
        try:
            import winreg
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
            _WORD = True
        except Exception:
            _WORD = False
    return _WORD


@app.before_request
def guard():
    """Who may call the server.  Locally: this page only.  When PAPERDESK_ORIGINS / PAPERDESK_TOKEN are
    set (the page is hosted elsewhere), the listed sites, and every request must carry the token."""
    origin = (request.headers.get("Origin") or "").rstrip("/")
    if request.method == "OPTIONS":
        return ("", 204)
    protected = request.path.startswith(("/api/", "/files/", "/preview/", "/certpreview/"))
    if origin and protected and request.method != "GET":
        same = origin.split("://", 1)[-1] == request.host
        if not same and origin not in ALLOWED_ORIGINS:
            return jsonify(error="This site is not allowed to use the server."), 403
    if API_TOKEN and protected and request.path != "/api/health":
        given = request.headers.get("X-API-Token") or request.args.get("token") or ""
        if given != API_TOKEN:
            return jsonify(error="Access token missing or wrong."), 401


@app.after_request
def headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    origin = (request.headers.get("Origin") or "").rstrip("/")
    if origin and origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-API-Token"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, PATCH, DELETE, OPTIONS"
        resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return resp


@app.errorhandler(HTTPException)
def http_error(exc):
    if request.path.startswith(("/api/", "/files/", "/preview/")):
        return jsonify(error=exc.description or exc.name), exc.code
    return exc


@app.errorhandler(Exception)
def server_error(exc):
    log.exception("request failed: %s %s", request.method, request.path)
    return jsonify(error="Something went wrong on the server. The details are in the log file."), 500


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html", max_age=0)


@app.get("/config.js")
def config_js():
    return send_from_directory(app.static_folder, "config.js", max_age=0)


@app.get("/logos/<key>.jpg")
def logo(key):
    if key not in CERT_KEYS.values():
        abort(404)
    return send_from_directory(CERT_DIR / "logos", key + ".jpg", max_age=86400)


@app.get("/api/health")
def api_health():
    h = health()
    ok = h["word"] and h["node"] and all(h["templates"].values())
    return jsonify(status="ok" if ok else "degraded", version=APP_VERSION, **h), (200 if ok else 503)


@app.get("/api/meta")
def api_meta():
    return jsonify(version=APP_VERSION, months=MONTHS, journals=JOURNALS, prefixes=DOI_PREFIXES,
                   logos={k: "/logos/%s.jpg" % v for k, v in CERT_KEYS.items()}, health=health())


@app.get("/api/parse-doi")
def api_parse_doi():
    """Live preview for the 'With DOI' box: what the DOI of the first paper says."""
    return jsonify(parse_doi(request.args.get("doi", "")))


@app.get("/api/init")
def api_init():
    # Every session starts empty: whatever an earlier session left is deleted.  `keep` is the issue of
    # this browser tab, so reloading the page does not lose the work in progress.
    keep = request.args.get("keep", "")
    for old in list_batches():
        if old["id"] != keep:
            shutil.rmtree(BATCHES / old["id"], ignore_errors=True)
    batches = list_batches()
    if not batches:
        batches = [summary_of(new_batch())]
    return jsonify(batches=batches, current=batches[0]["id"])


@app.post("/api/batches")
def api_new_batch():
    data = request.get_json(silent=True) or {}
    try:
        s = clean_settings(data) if data else None
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    b = new_batch(s)
    resp = view(b)
    resp["batches"] = list_batches()
    return jsonify(resp)


@app.get("/api/batches/<bid>")
def api_batch(bid):
    return jsonify(view(load(bid)))


@app.get("/api/batches")
def api_batches():
    return jsonify(batches=list_batches())


@app.delete("/api/batches/<bid>")
def api_delete_batch(bid):
    shutil.rmtree(bdir(bid), ignore_errors=True)
    return jsonify(batches=list_batches())


@app.put("/api/batches/<bid>/settings")
def api_settings(bid):
    try:
        s = clean_settings(request.get_json(force=True))
    except ValueError as exc:
        return jsonify(error=str(exc)), 400

    def fn(b):
        old = b["settings"]
        b["settings"] = s
        if old.get("journal") != s["journal"]:
            for p in b["papers"]:
                p["doi_manual"] = ""                          # its prefix names the old journal
        if any(old.get(k) != s[k] for k in ("journal", "volume", "issue", "month", "year")):
            for p in b["papers"]:
                if p["status"] != "error":
                    p["status"] = "queued"
                    p["reset_pages"] = True
                    p["rev"] = p.get("rev", 0) + 1

    b = mutate(bid, fn)
    submit(bid)
    resp = view(b)
    resp["batches"] = list_batches()
    return jsonify(resp)


@app.post("/api/batches/<bid>/papers")
def api_upload(bid):
    files = request.files.getlist("papers")
    if not files:
        return jsonify(error="Choose at least one Word file."), 400
    b = load(bid)
    if not settings_ok(b["settings"]):
        return jsonify(error="Fill in the issue details first (volume, issue, month, year, first paper number, first page)."), 400
    added, rejected = [], []
    for f in files:
        name = secure_filename(f.filename or "") or "paper.docx"
        ext = Path(name).suffix.lower()
        if ext not in ALLOWED:
            rejected.append(f.filename)
            continue
        pid = uuid.uuid4().hex[:12]
        d = BATCHES / bid / "papers" / pid
        d.mkdir(parents=True)
        f.save(d / ("original" + ext))
        added.append({"id": pid, "name": f.filename or name, "status": "queued", "stage": "",
                      "reset_pages": True, "rev": 0, "pages": None, "flags": [], "fixes": [],
                      "found": {}, "title": ""})
    if added:
        mutate(bid, lambda b: b["papers"].extend(added))
        submit(bid)
    resp = view(load(bid))
    if rejected:
        resp["warning"] = "Skipped (not a Word file): " + ", ".join(rejected)
    return jsonify(resp)


@app.patch("/api/batches/<bid>/papers/<pid>")
def api_patch_paper(bid, pid):
    data = request.get_json(force=True)

    def fn(b):
        p = get_paper(b, pid)
        if "doi" in data:
            doi = (data["doi"] or "").strip()
            prefix = DOI_PREFIXES[b["settings"]["journal"]]
            if doi and not doi.lower().startswith("10."):
                doi = prefix + doi
            p["doi_manual"] = doi if doi and doi != prefix else ""
        if "start" in data:
            st = str(data["start"] or "").strip()
            if st and not st.isdigit():
                raise ValueError("Start page must be a number.")
            p["start_manual"] = int(st) if st else 0

    try:
        b = mutate(bid, fn)
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    submit(bid)
    return jsonify(view(b))


@app.post("/api/batches/<bid>/papers/<pid>/move")
def api_move(bid, pid):
    d = -1 if (request.get_json(force=True) or {}).get("dir") == "up" else 1

    def fn(b):
        ps = b["papers"]
        i = next((k for k, p in enumerate(ps) if p["id"] == pid), None)
        if i is None:
            abort(404)
        j = i + d
        if 0 <= j < len(ps):
            ps[i], ps[j] = ps[j], ps[i]

    b = mutate(bid, fn)
    submit(bid)
    return jsonify(view(b))


@app.post("/api/batches/<bid>/papers/<pid>/redo")
def api_redo(bid, pid):
    def fn(b):
        p = get_paper(b, pid)
        p["status"] = "queued"
        p["reset_pages"] = True
        p["rev"] = p.get("rev", 0) + 1

    b = mutate(bid, fn)
    submit(bid)
    return jsonify(view(b))


@app.delete("/api/batches/<bid>/papers/<pid>")
def api_delete_paper(bid, pid):
    mutate(bid, lambda b: b.__setitem__("papers", [p for p in b["papers"] if p["id"] != pid]))
    shutil.rmtree(BATCHES / bid / "papers" / pid, ignore_errors=True)
    submit(bid)
    return jsonify(view(load(bid)))


# ----------------------------------------------------------------------------
# files: downloads, PDF, page previews, zip
# ----------------------------------------------------------------------------

def nice_name(p, suffix):
    stem = Path(p["name"]).stem
    return secure_filename(stem) + suffix


def ensure_pdf(b, p):
    d = bdir(b["id"]) / "papers" / p["id"]
    docx, pdf = d / "formatted.docx", d / "formatted.pdf"
    if not docx.exists():
        abort(404)
    if not pdf.exists() or pdf.stat().st_mtime < docx.stat().st_mtime:
        word_pages([docx], pdf=True)
    if not pdf.exists():
        abort(500, "Word could not make a PDF of this paper.")
    return pdf


@app.get("/files/<bid>/<pid>/<kind>")
def api_file(bid, pid, kind):
    b = load(bid)
    p = get_paper(b, pid)
    d = bdir(bid) / "papers" / pid
    if p["status"] != "done":
        abort(404)
    if kind == "docx":
        return send_file(d / "formatted.docx", as_attachment=True,
                         download_name=nice_name(p, "_formatted.docx"))
    if kind == "report":
        return send_file(d / "report.txt", as_attachment=True,
                         download_name=nice_name(p, "_report.txt"))
    if kind == "pdf":
        return send_file(ensure_pdf(b, p), as_attachment=True, download_name=nice_name(p, "_formatted.pdf"))
    abort(404)


@app.get("/files/<bid>/<pid>/cert")
def api_cert_file(bid, pid):
    name, fmt = request.args.get("name", ""), request.args.get("fmt", "docx")
    p = get_paper(load(bid), pid)
    entry = next((e for e in (p.get("certs") or {}).get("files", []) if e["stem"] == name), None)
    if entry is None or fmt not in ("docx", "pdf"):
        abort(404)
    f = cert_dir(bid, pid) / (entry["stem"] + "." + fmt)
    if not f.exists():
        abort(404)
    return send_file(f, as_attachment=True, download_name=f.name)


@app.put("/api/batches/<bid>/papers/<pid>/certdata")
def api_certdata(bid, pid):
    data = request.get_json(force=True) or {}
    authors = [{"name": str(a.get("name", "")).strip(), "designation": str(a.get("designation", "")).strip()}
               for a in data.get("authors", []) if str(a.get("name", "")).strip()]
    if not authors:
        return jsonify(error="Keep at least one author with a name."), 400

    def fn(b):
        q = get_paper(b, pid)
        q["cert_data"] = {"title": str(data.get("title", "")).strip(), "authors": authors, "warnings": []}
        q["cert_manual"] = True
        q["certs"] = None

    b = mutate(bid, fn)
    submit(bid)
    return jsonify(view(b))


@app.post("/api/batches/<bid>/papers/<pid>/certs/redo")
def api_certs_redo(bid, pid):
    reset = bool((request.get_json(silent=True) or {}).get("reset"))

    def fn(b):
        q = get_paper(b, pid)
        if reset and q.get("cert_data_auto"):
            q["cert_data"] = q["cert_data_auto"]
            q["cert_manual"] = False
        q["certs"] = None

    b = mutate(bid, fn)
    submit(bid)
    return jsonify(view(b))


@app.post("/api/batches/<bid>/papers/<pid>/certpreview")
def api_cert_preview(bid, pid):
    """Page pictures of the 'all certificates in one file' PDF."""
    try:
        import pymupdf
    except ImportError:
        return jsonify(error="Previews need the pymupdf package (pip install pymupdf)."), 501
    p = get_paper(load(bid), pid)
    entry = next((e for e in (p.get("certs") or {}).get("files", []) if e["kind"] == "all"), None)
    if entry is None:
        return jsonify(error="The certificates are not ready."), 400
    d = cert_dir(bid, pid)
    pdf = d / (entry["stem"] + ".pdf")
    docx = d / (entry["stem"] + ".docx")
    source = pdf if pdf.exists() else docx
    if not source.exists():
        return jsonify(error="The certificates are not ready."), 400
    version = int(source.stat().st_mtime)
    pdir = d.parent / "certpreview"
    stamp = pdir / ("v%d" % version)
    if not stamp.exists():
        data = pdf.read_bytes() if pdf.exists() else _pdf_for(docx)   # short temporary path: long names are safe
        if not data:
            return jsonify(error="Word could not make a PDF of the certificates."), 500
        shutil.rmtree(pdir, ignore_errors=True)
        pdir.mkdir(parents=True)
        with pymupdf.open(stream=data, filetype="pdf") as doc:
            for i, page in enumerate(doc, 1):
                page.get_pixmap(dpi=90).save(str(pdir / ("%d.png" % i)))
        stamp.touch()
    return jsonify(pages=len(list(pdir.glob("*.png"))), version=version)


@app.get("/certpreview/<bid>/<pid>/<int:n>.png")
def api_cert_preview_page(bid, pid, n):
    f = bdir(bid) / "papers" / pid / "certpreview" / ("%d.png" % n)
    if not valid_id(pid) or not f.exists():
        abort(404)
    return send_file(f, mimetype="image/png", max_age=3600)


@app.get("/files/<bid>/<pid>/bundle")
def api_bundle(bid, pid):
    """One ZIP for one paper: the formatted paper (Word and PDF), its report and all its certificates."""
    b = load(bid)
    p = get_paper(b, pid)
    if p["status"] != "done":
        abort(404)
    d = bdir(bid) / "papers" / pid
    base = nice_name(p, "")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(d / "formatted.docx", base + "_formatted.docx")
        z.write(d / "report.txt", base + "_report.txt")
        try:
            z.write(ensure_pdf(b, p), base + "_formatted.pdf")
        except Exception:
            pass
        for e in (p.get("certs") or {}).get("files", []):
            for ext in ("docx", "pdf"):
                f = cert_dir(bid, pid) / (e["stem"] + "." + ext)
                if f.exists():
                    z.write(f, "certificates/" + f.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=base + "_paper_and_certificates.zip",
                     mimetype="application/zip")


@app.post("/api/batches/<bid>/papers/<pid>/preview")
def api_preview(bid, pid):
    """Make page pictures of the formatted paper; return how many pages there are."""
    try:
        import pymupdf
    except ImportError:
        return jsonify(error="Page previews need the pymupdf package (pip install pymupdf)."), 501
    b = load(bid)
    p = get_paper(b, pid)
    if p["status"] != "done":
        return jsonify(error="This paper is not ready."), 400
    pdf = ensure_pdf(b, p)
    pdir = bdir(bid) / "papers" / pid / "preview"
    stamp = pdir / ("v%d" % int(pdf.stat().st_mtime))
    if not stamp.exists():
        shutil.rmtree(pdir, ignore_errors=True)
        pdir.mkdir(parents=True)
        with pymupdf.open(pdf) as doc:
            for i, page in enumerate(doc, 1):
                page.get_pixmap(dpi=90).save(str(pdir / ("%d.png" % i)))
        stamp.touch()
    n = len(list(pdir.glob("*.png")))
    return jsonify(pages=n, version=int(pdf.stat().st_mtime))


@app.get("/preview/<bid>/<pid>/<int:n>.png")
def api_preview_page(bid, pid, n):
    f = bdir(bid) / "papers" / pid / "preview" / ("%d.png" % n)
    if not valid_id(pid) or not f.exists():
        abort(404)
    return send_file(f, mimetype="image/png", max_age=3600)


@app.get("/api/batches/<bid>/zip")
def api_zip(bid):
    b = load(bid)
    v = view(b)
    buf = io.BytesIO()
    rows = [["Paper no.", "File", "Title", "DOI", "First page", "Last page", "Pages", "Items to check", "Certificates"]]
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        used = set()
        for p, pv in zip(b["papers"], v["papers"]):
            if p["status"] != "done":
                continue
            d = bdir(bid) / "papers" / p["id"]
            base = nice_name(p, "")
            while base in used:
                base += "_"
            used.add(base)
            z.write(d / "formatted.docx", base + "_formatted.docx")
            z.write(d / "report.txt", "reports/" + base + "_report.txt")
            ncert = 0
            for e in (p.get("certs") or {}).get("files", []):
                if e["kind"] == "single":
                    ncert += 1
                for ext in ("docx", "pdf"):
                    f = cert_dir(bid, p["id"]) / (e["stem"] + "." + ext)
                    if f.exists():
                        z.write(f, "certificates/%s/%s" % (base, f.name))
            rows.append([pv["paper_no"], p["name"], pv["title"], pv["doi"],
                         pv["start"], pv["end"], pv["pages"], len(pv["flags"]), ncert])
        sio = io.StringIO()
        csv.writer(sio).writerows(rows)
        z.writestr("contents.csv", "﻿" + sio.getvalue())
    buf.seek(0)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", v["label"]) + ".zip"
    return send_file(buf, as_attachment=True, download_name=name, mimetype="application/zip")


def requeue_unfinished():
    """After a restart, papers that were mid-way go back in the queue."""
    for d in BATCHES.iterdir():
        f = d / "batch.json"
        if not f.exists():
            continue
        try:
            b = json.loads(f.read_text(encoding="utf-8"))
        except ValueError:
            continue
        changed = False
        for p in b["papers"]:
            if p["status"] == "working":
                p["status"] = "queued"
                p["reset_pages"] = True
                changed = True
            if (p.get("certs") or {}).get("status") == "working":
                p["certs"] = None
                changed = True
        if changed:
            save(b)
        if any(p["status"] in ("queued",) or (p["status"] == "done" and (p.get("pages") is None or p.get("certs") is None))
               for p in b["papers"]):
            submit(b["id"])


threading.Thread(target=worker, daemon=True).start()
requeue_unfinished()

if __name__ == "__main__":
    log.info("Paper Desk %s on http://%s:%d  (data: %s)", APP_VERSION, HOST, PORT, DATA)
    try:
        from waitress import serve                           # a production server; Flask's own is the fallback
        serve(app, host=HOST, port=PORT, threads=8, channel_timeout=1200)
    except ImportError:
        app.run(host=HOST, port=PORT, debug=False, threaded=True)
