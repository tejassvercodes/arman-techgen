"""IJARCCE / IARJSET paper formatter – Vercel serverless edition.

Runs as a Vercel Python serverless function.  All file I/O goes to /tmp/paperdesk/
(ephemeral, lost on cold start — that's fine: upload → format → download in one session).

Differences from the local edition:
  • No Word COM (no page counting, no PDF, no .doc/.rtf/.odt conversion)
  • No Node.js (certificates are generated in pure Python)
  • No background worker thread (processing is synchronous)
  • No persistent data directory
  • No waitress
"""
import csv
import io
import json
import logging
import os
import re
import shutil
import sys
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from werkzeug.exceptions import HTTPException
from werkzeug.utils import secure_filename

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "formatter"))
try:
    from ijarcce_format import (format_paper, make_meta, ptext, certificate_data,  # noqa: E402
                                DOI_PREFIXES, JOURNAL_NAMES, auto_doi as journal_auto_doi)
except ImportError:
    from formatter.ijarcce_format import (format_paper, make_meta, ptext, certificate_data,  # noqa: E402
                                          DOI_PREFIXES, JOURNAL_NAMES, auto_doi as journal_auto_doi)

APP_VERSION = "2.0.0-cloud"
HERE = Path(__file__).resolve().parent

# ---- configuration -----------------------------------------------------------------------
HOST = os.environ.get("PAPERDESK_HOST", "127.0.0.1")
PORT = int(os.environ.get("PAPERDESK_PORT", "5000"))
DATA = Path(os.environ.get("PAPERDESK_DATA", "/tmp/paperdesk"))
API_TOKEN = os.environ.get("PAPERDESK_TOKEN", "")
ALLOWED_ORIGINS = {o.strip().rstrip("/") for o in os.environ.get("PAPERDESK_ORIGINS", "").split(",") if o.strip()}
MAX_UPLOAD_MB = int(os.environ.get("PAPERDESK_MAX_UPLOAD_MB", "4"))
BATCHES = DATA / "batches"

JOURNALS = dict(JOURNAL_NAMES)
CERT_KEYS = {"ijarcce": "CCE", "iarjset": "SET", "ijireeice": "ICE", "imrjr": "RJR"}
CERT_DIR = ROOT / "certificate-generator"
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]
ALLOWED = {".docx"}

app = Flask(__name__, static_folder=str(HERE / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.json.sort_keys = False

# ---- logging (stdout only, Vercel captures it) -------------------------------------------
log = logging.getLogger("paperdesk")
log.setLevel(logging.INFO)
if not log.handlers:
    log.addHandler(logging.StreamHandler())

# ---- ensure /tmp dirs lazily ---------------------------------------------------------------

def _ensure_dirs():
    BATCHES.mkdir(parents=True, exist_ok=True)

# ---- PDF provider stub (no Word COM on Vercel) --------------------------------------------

import ijarcce_format  # noqa: E402
ijarcce_format.PDF_PROVIDER = lambda path: None   # always return None → formatter skips PDF


# ============================================================================
# storage
# ============================================================================

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
    return json.loads(f.read_text(encoding="utf-8"))


def save(b):
    f = bdir(b["id"]) / "batch.json"
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(b, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(f)


def mutate(bid, fn):
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
    return {"journal": "ijarcce", "mode": "with", "first_doi": "", "volume": "", "issue": "", "month": "",
            "year": str(now.year), "first_paper": "1", "first_page": "1"}


def label_of(s):
    if not (s.get("volume") and s.get("issue")):
        return "New issue (details not filled in)"
    return "%s Vol. %s, Issue %s, %s %s" % (JOURNALS.get(s["journal"], s["journal"]), s["volume"],
                                            s["issue"], s["month"], s["year"])


def new_batch(settings=None):
    _ensure_dirs()
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
    _ensure_dirs()
    out = []
    for d in BATCHES.iterdir():
        f = d / "batch.json"
        if f.exists():
            try:
                out.append(summary_of(json.loads(f.read_text(encoding="utf-8"))))
            except ValueError:
                pass
    return sorted(out, key=lambda x: -x["created"])


def settings_ok(s):
    return all(str(s.get(k, "")).strip() for k in ("volume", "issue", "month", "year", "first_page", "first_paper"))


# How the issue details work: With DOI / Without DOI.
FIRST_JOURNAL_YEAR = 2011
DOI_LAYOUTS = [(2, 2, 2), (2, 1, 2), (2, 1, 3), (2, 2, 3), (1, 2, 2), (1, 1, 2)]


def clean_doi(raw):
    raw = (raw or "").strip()
    for _ in range(2):
        raw = re.sub(r"^doi\s*:?\s*", "", raw, flags=re.I)
        raw = re.sub(r"^(?:https?://)?(?:dx\.)?doi\.org/", "", raw, flags=re.I)
    return re.sub(r"\s+", "", raw)


def parse_doi(doi):
    out = {"journal": "", "volume": "", "issue": "", "year": "", "month": "", "file_number": "",
           "error": "", "note": ""}
    doi = clean_doi(doi)
    if not doi:
        return out
    m = re.match(r"10\.17148/([A-Za-z]+)\.", doi)
    if m:
        out["journal"] = next((k for k, v in JOURNALS.items() if v.lower() == m.group(1).lower()), "")
        if not out["journal"]:
            out["error"] = "The DOI names the journal %s, which this tool doesn't know." % m.group(1)
            return out
    runs = re.findall(r"\d+", doi)
    if len(runs) < 2:
        out["error"] = "Couldn't find volume, issue and year numbers in the DOI."
        return out
    last, year = runs[-1], runs[-2]
    if len(year) != 4:
        out["error"] = "Couldn't find a 4-digit year in the DOI (found \"%s\")." % year
        return out
    year_n = int(year)
    if year_n < FIRST_JOURNAL_YEAR or year_n > datetime.now().year + 1:
        out["error"] = "The year %s in the DOI isn't possible for these journals." % year
        return out
    cands = []
    for v, i, f in DOI_LAYOUTS:
        if v + i + f != len(last):
            continue
        vol, issue, file = int(last[:v]), int(last[v:v + i]), last[v + i:]
        if 1 <= vol <= year_n - FIRST_JOURNAL_YEAR + 1 and 1 <= issue <= 12:
            cands.append((vol, issue, file))
    if not cands:
        out["error"] = "Couldn't read a valid volume and issue from \"%s\" in the DOI. Use Without DOI and type them." % last
        return out
    vol, issue, file = cands[0]
    alt = next((c for c in cands if c[1] != issue and not c[2].startswith("0")), None)
    if alt:
        out["note"] = ("Check the DOI: \"%s\" was read as issue %d (%s), file %s, but could also be "
                       "issue %d (%s), file %s." % (last, issue, MONTHS[issue - 1], file,
                                                    alt[1], MONTHS[alt[1] - 1], alt[2]))
    out.update(volume=str(vol), issue=str(issue), year=year, month=MONTHS[issue - 1], file_number=file)
    return out


def clean_settings(src):
    s = {"mode": "with" if str(src.get("mode", "without")) == "with" else "without",
         "first_doi": clean_doi(src.get("first_doi", ""))}
    for k in ("journal", "volume", "issue", "year", "first_paper", "first_page"):
        s[k] = str(src.get(k, "")).strip()
    if s["journal"] not in JOURNALS:
        s["journal"] = "ijarcce"
    if s["mode"] == "without":
        s["first_doi"] = ""
    if s["mode"] == "with":
        if s["first_doi"]:
            d = parse_doi(s["first_doi"])
            if d["error"]:
                raise ValueError(d["error"])
            if d["journal"]:
                s["journal"] = d["journal"]
            s.update(volume=d["volume"], issue=d["issue"], year=d["year"],
                     first_paper=str(int(d["file_number"])))
        else:
            s.update(volume="", issue="", year="", first_paper="")
    for k in ("volume", "issue", "year", "first_paper", "first_page"):
        if s[k] and not s[k].isdigit():
            raise ValueError("%s must be a number." % k.replace("_", " ").capitalize())
    if s["issue"] and not 1 <= int(s["issue"]) <= 12:
        raise ValueError("Issue must be a number from 1 to 12 (it sets the month).")
    if s["year"] and len(s["year"]) != 4:
        raise ValueError("Year must be a 4-digit number.")
    s["month"] = MONTHS[int(s["issue"]) - 1] if s["issue"] else ""
    return s


# ============================================================================
# pipeline: format → chain DOIs / page numbers
# ============================================================================

def auto_doi(s, paper_no):
    return journal_auto_doi(s["journal"], s["year"], s["volume"], s["issue"], paper_no)


def targets(b):
    s = b["settings"]
    running = int(s["first_page"])
    first = int(s["first_paper"])
    out = {}
    for i, p in enumerate(b["papers"]):
        no = first + i
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
    return bdir(b["id"]) / "papers" / p["id"] / "source.docx"


def out_path(b, p):
    return bdir(b["id"]) / "papers" / p["id"] / "formatted.docx"


def set_paper(bid, pid, **fields):
    def fn(b):
        p = get_paper(b, pid)
        p.update(fields)
    mutate(bid, fn)


def estimate_pages(docx_path):
    """Estimate page count of a formatted paper without Word COM."""
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            if "docProps/app.xml" in z.namelist():
                app_xml = z.read("docProps/app.xml").decode("utf-8", errors="ignore")
                m = re.search(r"<Pages>(\d+)</Pages>", app_xml)
                if m and int(m.group(1)) > 0:
                    return int(m.group(1))
            if "word/document.xml" in z.namelist():
                doc_xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
                chars = sum(len(t) for t in re.findall(r"<w:t[^>]*>([^<]*)</w:t>", doc_xml))
                imgs = doc_xml.count("<w:drawing>") + doc_xml.count("<w:pict>")
                tables = doc_xml.count("</w:tbl>")
                est = round((chars / 3000.0) + (imgs + tables) * 0.3)
                return max(1, est)
    except Exception:
        pass
    return 1


def format_one(bid, p, t):
    b = load(bid)
    s = b["settings"]
    pid = p["id"]
    set_paper(bid, pid, status="working", stage="Formatting", error=None)
    try:
        d = bdir(bid) / "papers" / pid
        src = d / "source.docx"
        if not src.exists():
            orig = next(d.glob("original.*"), None)
            if orig is None:
                raise RuntimeError("Source file not found.")
            if orig.suffix.lower() == ".docx":
                shutil.copyfile(orig, src)
            else:
                raise RuntimeError("Only .docx files are supported on the cloud deployment. "
                                   "Open the file in Word and save it as .docx first.")
        meta = make_meta(s["volume"], s["issue"], s["month"], s["year"], doi=t["doi"],
                         start_page=t["start"], journal=s["journal"], no_doi=not t["doi"])
        out = d / "formatted.docx"
        fm, _ = format_paper(src, out, meta, journal=s["journal"], report_path=d / "report.txt")
        flags, shown = [], {}
        for el, msg in fm.flags:
            snippet = " ".join(ptext(el).split())[:70] if el is not None else ""
            shown[msg] = shown.get(msg, 0) + 1
            total = fm.repeat_counts.get(msg, 1)
            flags.append({"message": msg, "near": snippet,
                          "more": total - 3 if total > 3 and shown[msg] == 3 else 0})
        c = fm.counts
        found = {"Sections": c.get("h1", 0), "Sub-headings": c.get("h2", 0) + c.get("h3", 0),
                 "Figures": c.get("figure", 0), "Tables": c.get("table", 0),
                 "Equations": c.get("equation", 0), "References": getattr(fm, "ref_count", 0)}
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
        # Estimate page count (no Word COM on cloud). User can override manually if needed.
        fields["pages"] = p.get("pages") or estimate_pages(out)
        set_paper(bid, pid, reset_pages=False, **fields)
    except Exception as exc:
        log.exception("format_one failed for paper %s in batch %s", pid, bid)
        set_paper(bid, pid, status="error", stage="", error=str(exc)[:300], flags=[], fixes=[],
                  found={}, pages=None, snap=None, reset_pages=True)


# ============================================================================
# certificate generation – pure Python (no Node.js)
# ============================================================================

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


def _render_cert_docx(template_path, record):
    """Render a single certificate from a .docx template using pure-Python ZIP + XML substitution.

    This replicates what docxtemplater does in the Node.js code:
      1. Open the .docx (which is a ZIP) and read all XML parts
      2. Replace {placeholder} tags with actual values
      3. Handle {#hasDoi}...{/hasDoi} conditional blocks
      4. Return the rendered ZIP as bytes
    """
    values = {
        "NAME": (record.get("NAME") or "").upper(),
        "name": (record.get("NAME") or "").upper(),
        "NAMe": (record.get("NAME") or "").upper(),  # RJR template uses this spelling
        "Designation": record.get("Designation") or "",
        "designation": record.get("Designation") or "",
        "Paper Title": record.get("PaperTitle") or "",
        "paper title": record.get("PaperTitle") or "",
        "PaperTitle": record.get("PaperTitle") or "",
        "papertitle": record.get("PaperTitle") or "",
        "DOI": record.get("DOI") or "",
        "doi": record.get("DOI") or "",
        "vol": record.get("Volume") or "",
        "Vol": record.get("Volume") or "",
        "Volume": record.get("Volume") or "",
        "volume": record.get("Volume") or "",
        "issue": record.get("Issue") or "",
        "Issue": record.get("Issue") or "",
        "year": record.get("Year") or "",
        "Year": record.get("Year") or "",
        "month": record.get("Month") or "",
        "Month": record.get("Month") or "",
    }
    has_doi = bool(record.get("DOI"))

    with zipfile.ZipFile(template_path, "r") as zin:
        infos_and_data = [(info, zin.read(info.filename)) for info in zin.infolist()]

    result = []
    for info, data in infos_and_data:
        # Only process XML files inside word/
        if info.filename.endswith(".xml") or info.filename.endswith(".xml.rels"):
            text = data.decode("utf-8")

            # Handle {#hasDoi}...{/hasDoi} conditional blocks.
            # In the OOXML, the tags may be split across runs, e.g.:
            #   <w:t>{#has</w:t></w:r><w:r><w:t>Doi}</w:t>
            # So we work on the raw XML string.
            # First, strip XML tags to find the conditional in the logical text, but we need
            # to handle it in the XML. We'll use a regex that works across run boundaries.

            # Pattern for {#hasDoi}...{/hasDoi} that may span XML tags
            # We match the opening and closing tags allowing XML markup between characters
            def _make_xml_aware_pattern(tag_text):
                """Build a regex that matches tag_text even when split across <w:t> runs."""
                parts = []
                for ch in tag_text:
                    if ch in r"\.^$*+?{}[]|()":
                        parts.append("\\" + ch)
                    else:
                        parts.append(ch)
                    # Allow optional XML tags between characters
                    parts.append(r"(?:</w:t>(?:\s*</w:r>\s*<w:r(?:\s[^>]*)?>)?\s*<w:t(?:\s[^>]*)?>)?")
                # Remove the trailing optional part after the last character
                parts.pop()
                return "".join(parts)

            open_pat = _make_xml_aware_pattern("{#hasDoi}")
            close_pat = _make_xml_aware_pattern("{/hasDoi}")

            # Find and handle the conditional
            open_m = re.search(open_pat, text, re.DOTALL)
            close_m = re.search(close_pat, text, re.DOTALL)
            if open_m and close_m:
                if has_doi:
                    # Keep the content, remove just the tags
                    text = text[:open_m.start()] + text[open_m.end():close_m.start()] + text[close_m.end():]
                else:
                    # Remove everything from opening tag to closing tag (inclusive)
                    # But we need to handle paragraph-level removal.
                    # Remove the entire block between {#hasDoi} and {/hasDoi}
                    # Find the enclosing <w:p> elements if the block spans full paragraphs
                    before = text[:open_m.start()]
                    after = text[close_m.end():]

                    # Check if the opening tag starts at the beginning of a paragraph's text content
                    # and the closing tag ends at the end — if so, remove the whole paragraph(s)
                    p_start = before.rfind("<w:p ")
                    if p_start == -1:
                        p_start = before.rfind("<w:p>")
                    p_end = after.find("</w:p>")

                    if p_start >= 0 and p_end >= 0:
                        # Check that the content between p_start and open_m.start() is only XML markup
                        between_before = before[p_start:]
                        # Strip XML tags and check if remaining text is only whitespace
                        text_only = re.sub(r"<[^>]+>", "", between_before)
                        if not text_only.strip():
                            text = before[:p_start] + after[p_end + len("</w:p>"):]
                        else:
                            text = before + after
                    else:
                        text = before + after

            # Replace {placeholder} tags (may be split across runs in OOXML)
            # Strategy: for each known placeholder, build a regex that matches it across runs
            for tag, value in values.items():
                tag_with_braces = "{" + tag + "}"
                pat = _make_xml_aware_pattern(tag_with_braces)
                # Replace — the value itself must not contain XML special chars
                safe_value = (value.replace("&", "&amp;").replace("<", "&lt;")
                              .replace(">", "&gt;").replace('"', "&quot;"))
                text = re.sub(pat, safe_value, text, flags=re.DOTALL)

            data = text.encode("utf-8")

        result.append((info, data))

    return result


def _write_cert_zip(infos_and_data, out_path):
    """Write a list of (ZipInfo, bytes) tuples as a .docx file."""
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in infos_and_data:
            z.writestr(info, data)


def _extract_body_content(xml_text):
    """Extract paragraph content between <w:body> and the trailing <w:sectPr>."""
    body_open = xml_text.find("<w:body>")
    if body_open == -1:
        return xml_text
    body_open += len("<w:body>")
    sect_pr = xml_text.rfind("<w:sectPr")
    if sect_pr == -1:
        return xml_text[body_open:]
    return xml_text[body_open:sect_pr]


def _make_ids_unique(xml, counters):
    """Remap OOXML element IDs so they're unique across concatenated certificate pages."""
    bookmark_map = {}

    def docpr_repl(m):
        counters["docPr"] += 1
        return m.group(1) + str(counters["docPr"]) + m.group(2)
    xml = re.sub(r'(<wp:docPr\s+id=")\d+(")', docpr_repl, xml)

    def bookmark_repl(m):
        old_id = m.group(2)
        if old_id not in bookmark_map:
            counters["bookmark"] += 1
            bookmark_map[old_id] = str(counters["bookmark"])
        return m.group(1) + bookmark_map[old_id] + m.group(3)
    xml = re.sub(r'(<w:bookmark(?:Start|End)\s+w:id=")([\d]+)(")', bookmark_repl, xml)

    def bookmarkname_repl(m):
        counters["bookmarkName"] += 1
        return m.group(1) + m.group(2) + "_" + str(counters["bookmarkName"]) + m.group(3)
    xml = re.sub(r'(<w:bookmarkStart\s+w:id="\d+"\s+w:name=")([^"]*)(")', bookmarkname_repl, xml)

    def hex_repl(m):
        counters["hex"] += 1
        return m.group(1) + format(counters["hex"], "08X") + m.group(2)
    xml = re.sub(r'(w14:(?:para|text)Id=")[0-9A-Fa-f]{1,8}(")', hex_repl, xml)

    def vml_repl(m):
        counters["vml"] += 1
        return m.group(1) + str(counters["vml"]) + m.group(2)
    xml = re.sub(r'(o:spid="_x0000_s)\d+(")', vml_repl, xml)

    return xml


def _build_combined_cert(template_path, records):
    """Build a multi-page certificate .docx: one page per record, combined into one file."""
    if not records:
        return None

    # Render the first record to get the base ZIP structure
    first_parts = _render_cert_docx(template_path, records[0])

    # Find document.xml in the parts
    doc_idx = None
    for i, (info, data) in enumerate(first_parts):
        if info.filename == "word/document.xml":
            doc_idx = i
            break

    if doc_idx is None:
        # Fallback: just return the first cert
        return first_parts

    xml0 = first_parts[doc_idx][1].decode("utf-8")

    # Extract structure
    body_open_idx = xml0.find("<w:body>") + len("<w:body>")
    sect_pr_start = xml0.rfind("<w:sectPr")
    sect_pr_end = xml0.find("</w:sectPr>", sect_pr_start) + len("</w:sectPr>")
    sect_pr_xml = xml0[sect_pr_start:sect_pr_end]
    prefix = xml0[:body_open_idx]
    suffix = xml0[sect_pr_end:]  # "</w:body></w:document>"

    counters = {"docPr": 0, "bookmark": 0, "bookmarkName": 0, "hex": 0, "vml": 0}
    bodies = [_make_ids_unique(_extract_body_content(xml0), counters)]

    for rec in records[1:]:
        rec_parts = _render_cert_docx(template_path, rec)
        for info, data in rec_parts:
            if info.filename == "word/document.xml":
                rec_xml = data.decode("utf-8")
                bodies.append(_make_ids_unique(_extract_body_content(rec_xml), counters))
                break

    section_break = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>'
                     '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
                     + sect_pr_xml + '</w:pPr></w:p>')

    combined_xml = prefix + section_break.join(bodies) + sect_pr_xml + suffix
    first_parts[doc_idx] = (first_parts[doc_idx][0], combined_xml.encode("utf-8"))
    return first_parts


def _sanitize_filename(name, max_len=80):
    """Produce a safe filename from certificate data."""
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:max_len] if name else "certificate"


def make_certs_one(bid, pid, t):
    """Fill the journal's certificate template for every author — pure Python, no Node.js."""
    b = load(bid)
    p = get_paper(b, pid)
    s = b["settings"]
    snap = cert_snapshot(b, p, t)
    set_paper(bid, pid, certs={"status": "working", "files": [], "warnings": [], "error": None})
    d = cert_dir(bid, pid)
    try:
        cd = p.get("cert_data") or {}
        if not cd.get("authors"):
            raise RuntimeError("No authors were found in the paper, so there is nothing to put on the certificates. "
                               "Add them in the table below.")
        if d.exists():
            shutil.rmtree(d)
        d.mkdir(parents=True)

        key = CERT_KEYS[s["journal"]]
        template_path = CERT_DIR / "templates" / (key + ".docx")
        if not template_path.exists():
            raise RuntimeError("Certificate template %s.docx not found." % key)

        title = cd.get("title", "")
        authors = cd["authors"]
        doi = t["doi"]
        file_number = "%02d" % int(t["paper_no"])

        # Build author list line for filenames
        author_list = ", ".join(a["name"] for a in authors)

        def dedupe_join(values):
            seen = []
            for v in values:
                v = (v or "").strip()
                if v and v not in seen:
                    seen.append(v)
            return ", ".join(seen)

        # Base publication info for all records
        base = {
            "DOI": doi,
            "Volume": str(int(s["volume"])),
            "Issue": str(int(s["issue"])),
            "Year": s["year"],
            "Month": s["month"],
        }

        # Individual records
        records = []
        for a in authors:
            rec = dict(base, NAME=a["name"], Designation=a.get("designation", ""), PaperTitle=title)
            records.append(rec)

        # Group record
        group_rec = dict(base,
                         NAME=dedupe_join(r["NAME"] for r in records),
                         Designation=dedupe_join(r["Designation"] for r in records),
                         PaperTitle=title)

        # Build filenames
        base_stem = _sanitize_filename(("%s %s" % (file_number, author_list)).strip())

        warnings = list(cd.get("warnings") or [])
        entries = []

        # 1. All-in-one document (all individual certs combined)
        combined_parts = _build_combined_cert(template_path, records)
        all_stem = base_stem
        all_path = d / (all_stem + ".docx")
        _write_cert_zip(combined_parts, all_path)
        if not doi:
            strip_fixed_doi(all_path)
        entries.append({"stem": all_stem, "kind": "all", "index": 0, "pdf": False})

        # 2. Individual certificates (one per author)
        for i, rec in enumerate(records):
            indiv_parts = _render_cert_docx(template_path, rec)
            indiv_stem = "%s %d" % (base_stem, i + 1)
            indiv_path = d / (indiv_stem + ".docx")
            _write_cert_zip(indiv_parts, indiv_path)
            if not doi:
                strip_fixed_doi(indiv_path)
            entries.append({"stem": indiv_stem, "kind": "single", "index": i + 1, "pdf": False})

        # 3. Group certificate
        group_parts = _build_combined_cert(template_path, [group_rec])
        group_stem = base_stem + " (group)"
        group_path = d / (group_stem + ".docx")
        _write_cert_zip(group_parts, group_path)
        if not doi:
            strip_fixed_doi(group_path)
        entries.append({"stem": group_stem, "kind": "group", "index": 0, "pdf": False})

        # Warnings about missing info
        for i, a in enumerate(authors):
            if not a.get("designation"):
                warnings.append("No designation for certificate #%d." % (i + 1))
        if not title:
            warnings.append("No paper title.")

        set_paper(bid, pid, certs={"status": "ok", "files": entries, "warnings": warnings, "error": None},
                  cert_snap=snap)
    except Exception as exc:
        log.exception("certificate generation failed for paper %s in batch %s", pid, bid)
        set_paper(bid, pid, certs={"status": "error", "files": [], "warnings": [], "error": str(exc)[:400]},
                  cert_snap=snap)


def pipeline(bid):
    """Synchronous pipeline: format all papers, then generate certificates."""
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
        # No page counting step (no Word COM)
        b = load(bid)
        t = targets(b)
        cert_todo = [p["id"] for p in b["papers"]
                     if p["status"] == "done"
                     and (p.get("certs") is None
                          or (p["certs"].get("status") != "working"
                              and p.get("cert_snap") != cert_snapshot(b, p, t[p["id"]])))]
        for pid in cert_todo:
            make_certs_one(bid, pid, t[pid])
        if not todo and not cert_todo:
            return


# ============================================================================
# views
# ============================================================================

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
    return {"word": True,           # cloud mode: pure-python lxml docx pipeline ready
            "node": True,           # pure Python certificate generation ready
            "templates": {k: Path(v).exists() for k, v in ijarcce_format.JOURNAL_TEMPLATES.items()},
            "preview": False,
            "cloud": True}


@app.before_request
def guard():
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
    return jsonify(error="Something went wrong on the server."), 500


@app.get("/")
def index():
    for candidate in [
        Path(app.static_folder) / "index.html",
        HERE / "static" / "index.html",
        ROOT / "webapp" / "static" / "index.html",
        Path("webapp/static/index.html")
    ]:
        if candidate.exists():
            return send_file(candidate, max_age=0)
    return "<h1>Arman Paper Desk</h1><p>index.html not found</p>", 404


@app.get("/config.js")
def config_js():
    for candidate in [
        Path(app.static_folder) / "config.js",
        HERE / "static" / "config.js",
        ROOT / "webapp" / "static" / "config.js",
        Path("webapp/static/config.js")
    ]:
        if candidate.exists():
            return send_file(candidate, max_age=0)
    return "// config.js", 200, {"Content-Type": "application/javascript"}


@app.get("/logos/<key>.jpg")
def logo(key):
    if key not in CERT_KEYS.values():
        abort(404)
    for candidate_dir in [
        CERT_DIR / "logos",
        ROOT / "certificate-generator" / "logos",
        Path("certificate-generator/logos")
    ]:
        p = candidate_dir / (key + ".jpg")
        if p.exists():
            return send_file(p, max_age=86400)
    abort(404)


@app.get("/api/health")
def api_health():
    h = health()
    ok = all(h["templates"].values())
    return jsonify(status="ok" if ok else "degraded", version=APP_VERSION, **h), (200 if ok else 503)


@app.get("/api/meta")
def api_meta():
    return jsonify(version=APP_VERSION, months=MONTHS, journals=JOURNALS, prefixes=DOI_PREFIXES,
                   logos={k: "/logos/%s.jpg" % v for k, v in CERT_KEYS.items()}, health=health())


@app.get("/api/parse-doi")
def api_parse_doi():
    return jsonify(parse_doi(request.args.get("doi", "")))


@app.get("/api/init")
def api_init():
    _ensure_dirs()
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
                p["doi_manual"] = ""
        if any(old.get(k) != s[k] for k in ("journal", "volume", "issue", "month", "year")):
            for p in b["papers"]:
                if p["status"] != "error":
                    p["status"] = "queued"
                    p["reset_pages"] = True
                    p["rev"] = p.get("rev", 0) + 1

    b = mutate(bid, fn)
    # Synchronous processing
    pipeline(bid)
    b = load(bid)
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
        # Format synchronously
        pipeline(bid)
    resp = view(load(bid))
    if rejected:
        resp["warning"] = "Skipped (not a .docx file): " + ", ".join(rejected)
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
    pipeline(bid)
    return jsonify(view(load(bid)))


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
    pipeline(bid)
    return jsonify(view(load(bid)))


@app.post("/api/batches/<bid>/papers/<pid>/redo")
def api_redo(bid, pid):
    def fn(b):
        p = get_paper(b, pid)
        p["status"] = "queued"
        p["reset_pages"] = True
        p["rev"] = p.get("rev", 0) + 1

    b = mutate(bid, fn)
    pipeline(bid)
    return jsonify(view(load(bid)))


@app.delete("/api/batches/<bid>/papers/<pid>")
def api_delete_paper(bid, pid):
    mutate(bid, lambda b: b.__setitem__("papers", [p for p in b["papers"] if p["id"] != pid]))
    shutil.rmtree(BATCHES / bid / "papers" / pid, ignore_errors=True)
    pipeline(bid)
    return jsonify(view(load(bid)))


# ============================================================================
# files: downloads, PDF, page previews, zip
# ============================================================================

def nice_name(p, suffix):
    stem = Path(p["name"]).stem
    return secure_filename(stem) + suffix


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
        return jsonify(error="PDF preview not available on cloud deployment. Download the .docx file."), 404
    abort(404)


@app.get("/files/<bid>/<pid>/cert")
def api_cert_file(bid, pid):
    name, fmt = request.args.get("name", ""), request.args.get("fmt", "docx")
    p = get_paper(load(bid), pid)
    entry = next((e for e in (p.get("certs") or {}).get("files", []) if e["stem"] == name), None)
    if entry is None or fmt not in ("docx",):
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
    pipeline(bid)
    return jsonify(view(load(bid)))


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
    pipeline(bid)
    return jsonify(view(load(bid)))


@app.post("/api/batches/<bid>/papers/<pid>/certpreview")
def api_cert_preview(bid, pid):
    return jsonify(error="Certificate preview not available on cloud deployment. Download the .docx file."), 501


@app.get("/certpreview/<bid>/<pid>/<int:n>.png")
def api_cert_preview_page(bid, pid, n):
    abort(404)


@app.get("/files/<bid>/<pid>/bundle")
def api_bundle(bid, pid):
    """One ZIP for one paper: the formatted paper (Word), its report and all its certificates."""
    b = load(bid)
    p = get_paper(b, pid)
    if p["status"] != "done":
        abort(404)
    d = bdir(bid) / "papers" / pid
    base = nice_name(p, "")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        docx_path = d / "formatted.docx"
        if docx_path.exists():
            z.write(docx_path, base + "_formatted.docx")
        report_path = d / "report.txt"
        if report_path.exists():
            z.write(report_path, base + "_report.txt")
        for e in (p.get("certs") or {}).get("files", []):
            for ext in ("docx",):
                f = cert_dir(bid, pid) / (e["stem"] + "." + ext)
                if f.exists():
                    z.write(f, "certificates/" + f.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=base + "_paper_and_certificates.zip",
                     mimetype="application/zip")


@app.post("/api/batches/<bid>/papers/<pid>/preview")
def api_preview(bid, pid):
    return jsonify(error="PDF preview not available on cloud deployment. Download the .docx file."), 501


@app.get("/preview/<bid>/<pid>/<int:n>.png")
def api_preview_page(bid, pid, n):
    abort(404)


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
            docx_path = d / "formatted.docx"
            if docx_path.exists():
                z.write(docx_path, base + "_formatted.docx")
            report_path = d / "report.txt"
            if report_path.exists():
                z.write(report_path, "reports/" + base + "_report.txt")
            ncert = 0
            for e in (p.get("certs") or {}).get("files", []):
                if e["kind"] == "single":
                    ncert += 1
                for ext in ("docx",):
                    f = cert_dir(bid, p["id"]) / (e["stem"] + "." + ext)
                    if f.exists():
                        z.write(f, "certificates/%s/%s" % (base, f.name))
            rows.append([pv["paper_no"], p["name"], pv["title"], pv["doi"],
                         pv["start"], pv["end"], pv["pages"], len(pv["flags"]), ncert])
        sio = io.StringIO()
        csv.writer(sio).writerows(rows)
        z.writestr("contents.csv", "\ufeff" + sio.getvalue())
    buf.seek(0)
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", v["label"]) + ".zip"
    return send_file(buf, as_attachment=True, download_name=name, mimetype="application/zip")


# ============================================================================
# manual page-count setter (since Word COM isn't available to count them)
# ============================================================================

@app.patch("/api/batches/<bid>/papers/<pid>/pages")
def api_set_pages(bid, pid):
    data = request.get_json(force=True) or {}
    pages = data.get("pages")
    if pages is not None:
        try:
            pages = int(pages)
            if pages < 1:
                raise ValueError
        except (ValueError, TypeError):
            return jsonify(error="Pages must be a positive number."), 400
    set_paper(bid, pid, pages=pages)
    pipeline(bid)
    return jsonify(view(load(bid)))
