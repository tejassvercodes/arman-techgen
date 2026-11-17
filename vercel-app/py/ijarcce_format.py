"""Format a student paper (.docx) to the IJARCCE layout.

The student's file is used as the package (so images, equations, charts,
lists and footnotes survive untouched) and is then given the journal's
styles, header, footer and page set-up from templates/CCE_production.docx.
Every paragraph is classified (title, author, heading, caption, reference...)
and restyled.  Anything the tool is unsure about is left in place and marked
with a Word comment, and listed in a text report.

Usage:
    python ijarcce_format.py paper.docx out.docx --vol 15 --issue 11 \
        --month November --year 2026 --paper 136 --start-page 1034
"""
import argparse
import copy
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
V = "urn:schemas-microsoft-com:vml"
XML = "http://www.w3.org/XML/1998/namespace"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "CCE_production.docx"
JOURNAL_TEMPLATES = {"ijarcce": DEFAULT_TEMPLATE, "iarjset": TEMPLATE_DIR / "IARJSET_production.docx",
                     "ijireeice": TEMPLATE_DIR / "IJIREEICE_production.docx",
                     "imrjr": TEMPLATE_DIR / "IMRJR_production.docx"}
JOURNAL_NAMES = {"ijarcce": "IJARCCE", "ijireeice": "IJIREEICE", "iarjset": "IARJSET", "imrjr": "IMRJR"}
DOI_PREFIXES = {k: "10.17148/%s." % v for k, v in JOURNAL_NAMES.items()}


def auto_doi(journal, year, vol, issue, paper):
    """10.17148/<JOURNAL>.<year>.<volume, 2 digits><issue><paper number, 2 digits> (the layout the
    certificate generator reads back)."""
    return "%s%s.%02d%s%02d" % (DOI_PREFIXES[journal], year, int(vol), int(issue), int(paper))

PAGE_W, PAGE_H, MARGIN = 11909, 16834, 1152
TEXT_WIDTH = PAGE_W - 2 * MARGIN                 # twips
TEXT_EMU = TEXT_WIDTH * 635                      # same, in EMU
MAX_IMG_H_EMU = int(8.3 * 914400)

SYMBOL_FONTS = ("symbol", "wingdings", "webdings", "cambria math", "mt extra",
                "segoe ui symbol", "lucida sans unicode", "arial unicode ms", "ms gothic",
                "ms mincho", "math")
MONO_FONTS = ("courier", "consolas", "lucida console", "monospace", "menlo")


def w(tag):
    return "{%s}%s" % (W, tag)


PDF_PROVIDER = None     # callable(path) -> bytes of a PDF of that Word file; the website replaces it


def default_pdf_provider(path):
    """A PDF of the student's file, made by Microsoft Word (None if Word is not available)."""
    import shutil
    import subprocess
    import tempfile
    ps1 = Path(__file__).with_name("word_pages.ps1")
    if not ps1.exists():
        return None
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d) / "source.docx"
        shutil.copyfile(path, tmp)
        cmd = "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '%s'))) -Paths '%s' -Pdf" % (
            str(ps1).replace("'", "''"), str(tmp).replace("'", "''"))
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", cmd], capture_output=True,
                           text=True, timeout=300)
        except (OSError, subprocess.SubprocessError):
            return None
        pdf = tmp.with_suffix(".pdf")
        return pdf.read_bytes() if pdf.exists() else None


def m(tag):
    return "{%s}%s" % (M, tag)


# Sub-trees whose text is not part of the paragraph's own running text.
OPAQUE = {w("drawing"), w("pict"), w("object"), "{%s}AlternateContent" % MC,
          m("oMath"), m("oMathPara"), w("txbxContent"), w("del"), w("instrText"),
          w("footnoteReference"), w("endnoteReference"), w("commentReference")}

KNOWN_H1 = {
    "introduction", "related work", "related works", "literature review", "literature survey",
    "background", "methodology", "methods", "materials and methods", "proposed system",
    "proposed methodology", "proposed method", "proposed work", "system architecture",
    "system design", "system overview", "existing system", "problem statement", "objectives",
    "implementation", "system implementation", "experimental setup", "experimental results",
    "results", "result", "results and discussion", "result and discussion", "discussion",
    "limitations", "future work", "future scope", "conclusion", "conclusions",
    "conclusion and future work", "conclusion and future scope",
}
UNNUMBERED_H1 = {
    "acknowledgment", "acknowledgments", "acknowledgement", "acknowledgements", "references",
    "reference", "biography", "biographies", "appendix", "conflict of interest",
    "conflicts of interest", "funding", "author contributions", "data availability",
    "declaration", "declarations",
}
AFFIL_RE = re.compile(
    r"universit|institut|college|department|dept\b|school|facult|laborator|engineering|"
    r"technology|student|professor|scholar|lecturer|india\b|@|corresponding|\bhod\b|campus|"
    r"academy|polytechnic|pvt|\bltd\b|centre|center\b|\bguide\b", re.I)
NAME_RE = re.compile(r"(?:[A-Z][\w.'’-]*\s*){1,6}")
REGNO_RE = re.compile(r"[\s,;(]*\bReg(?:\.|istration)?\s*(?:No\.?|Number)\s*[:.]?\s*[\w/-]+\)?", re.I)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
CAPTION_VERB_RE = re.compile(
    r"^(shows?|illustrates?|depicts?|presents?|displays?|demonstrates?|is|are|represents?|and|"
    r"gives?|provides?|compares?|summari[sz]es?|lists?|below|above|,)\b", re.I)
FIG_RE = re.compile(r"^\s*(fig(?:ure)?s?\.?)\s*(\d+)\s*[.:\-–—]?\s*(.*)$", re.I | re.S)
TAB_RE = re.compile(r"^\s*(table)\s*([IVXLC]+|\d+)\b\s*[.:\-–—]?\s*(.*)$", re.I | re.S)
ROMAN_RE = re.compile(r"^\s*([IVXLC]{1,6})\s*[.)]\s*(\S.*)$", re.S)
ALPHA_RE = re.compile(r"^\s*([A-Z])\s*[.)]\s*(\S.*)$", re.S)
NUM_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,2})\s*[.)]?\s+(\S.*)$", re.S)
LABEL_RE = re.compile(
    r"^\s*(?P<label>authors?(?:\s*names?)?|names?|guide(?:\s*names?)?|supervisor|mentor|submitted\s+by|by)"
    r"\s*[:\-\u2013\u2014]\s*(?=\S)", re.I)
REFNUM_RE = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\d+\)|\d{1,3}[.)])\s*")
EQNUM_RE = re.compile(r"\(\s*\d+[a-z]?\s*\)\s*$")


def to_roman(n):
    out = ""
    for val, sym in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                     (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= val:
            out += sym
            n -= val
    return out


def from_roman(s):
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}
    total = 0
    for i, ch in enumerate(s):
        v = vals[ch]
        total += -v if i + 1 < len(s) and vals[s[i + 1]] > v else v
    return total if to_roman(total) == s else None


# --------------------------------------------------------------------------
# paragraph helpers
# --------------------------------------------------------------------------

def atoms(p):
    """Text-bearing nodes of a paragraph in document order: (element, kind)."""
    out = []

    def walk(el):
        for ch in el:
            if ch.tag in OPAQUE:
                continue
            if ch.tag == w("t"):
                out.append((ch, "t"))
            elif ch.tag == w("tab") and ch.getparent().tag == w("r"):
                out.append((ch, "tab"))
            elif ch.tag in (w("br"), w("cr")):
                out.append((ch, "br"))
            elif ch.tag == w("noBreakHyphen"):
                out.append((ch, "hy"))
            else:
                walk(ch)
    walk(p)
    return out


def ptext(p):
    parts = []
    for el, kind in atoms(p):
        parts.append((el.text or "") if kind == "t" else {"tab": "\t", "br": "\n", "hy": "-"}[kind])
    return "".join(parts)


def math_count(p):
    return sum(1 for _ in p.iter(m("oMath")))


def has_graphic(p):
    """A picture or drawing that carries no words of its own.  A text box that holds text
    (a drop cap, an equation fragment) is not a 'graphic'."""
    for tag in (w("drawing"), w("pict"), w("object")):
        for el in p.iter(tag):
            if any((t.text or "").strip() for t in el.iter(w("t"))):
                continue
            return True
    return False


def has_shape(p):
    return any(next(p.iter(tag), None) is not None for tag in (w("drawing"), w("pict"), w("object")))


def token_text(n):
    return "(%d)" % n


def split_lines(p):
    """Split a paragraph at its line breaks into one paragraph per line."""
    ppr = p.find(w("pPr"))
    out = []

    def new():
        q = etree.Element(w("p"))
        if ppr is not None:
            q.append(copy.deepcopy(ppr))
        out.append(q)
        return q

    def is_line_break(e):
        return e.tag == w("br") and e.get(w("type")) in (None, "textWrapping")

    cur = new()
    for ch in list(p):
        if ch.tag == w("pPr"):
            continue
        if ch.tag == w("r") and any(is_line_break(e) for e in ch):
            rpr = ch.find(w("rPr"))
            seg = None
            for sub in list(ch):
                if sub.tag == w("rPr"):
                    continue
                if is_line_break(sub):
                    cur = new()
                    seg = None
                    continue
                if seg is None:
                    seg = etree.SubElement(cur, w("r"))
                    if rpr is not None:
                        seg.append(copy.deepcopy(rpr))
                seg.append(sub)
        else:
            cur.append(ch)
    return out


def split_paragraph_at(p, offset):
    """Cut a paragraph after `offset` characters of its text; return the new paragraph
    that holds the rest (p keeps the beginning)."""
    newp = etree.Element(w("p"))
    ppr = p.find(w("pPr"))
    if ppr is not None:
        newp.append(copy.deepcopy(ppr))
    count = 0
    started = False
    for ch in list(p):
        if ch.tag == w("pPr"):
            continue
        if started:
            newp.append(ch)
            continue
        at = atoms(ch) if ch.tag == w("r") else []
        ln = sum(len(e.text or "") if k == "t" else 1 for e, k in at)
        if count + ln <= offset:
            count += ln
            continue
        pos = offset - count
        r2 = copy.deepcopy(ch)
        mine, theirs = list(atoms(ch)), list(atoms(r2))
        acc = 0
        for idx, ((e, k), (e2, _)) in enumerate(zip(mine, theirs)):
            n = len(e.text or "") if k == "t" else 1
            if acc + n <= pos:
                e2.getparent().remove(e2)
                acc += n
                continue
            if k == "t" and acc < pos:
                cut = pos - acc
                full = e.text or ""
                e.text, e2.text = full[:cut], full[cut:]
                e.set("{%s}space" % XML, "preserve")
                e2.set("{%s}space" % XML, "preserve")
                for e3, _k in mine[idx + 1:]:
                    e3.getparent().remove(e3)
            else:
                for e3, _k in mine[idx:]:
                    e3.getparent().remove(e3)
            break
        newp.append(r2)
        started = True
    return newp


def rstrip_chars(p, chars):
    for el, kind in reversed(atoms(p)):
        if kind != "t":
            continue
        t = (el.text or "").rstrip(chars)
        el.text = t
        if t:
            break


def delete_prefix(p, n):
    """Remove the first n characters of the paragraph's running text."""
    for el, kind in atoms(p):
        if n <= 0:
            break
        if kind == "t":
            t = el.text or ""
            take = min(n, len(t))
            el.text = t[take:]
            el.set("{%s}space" % XML, "preserve")
            n -= take
        else:
            el.getparent().remove(el)
            n -= 1


def rstrip_text(p):
    for el, kind in reversed(atoms(p)):
        if kind != "t":
            el.getparent().remove(el)
            continue
        t = (el.text or "").rstrip()
        el.text = t
        if t:
            break


def map_text(p, fn):
    for el, kind in atoms(p):
        if kind == "t" and el.text:
            el.text = fn(el.text)


def make_run(text, bold=False, tab_after=False):
    r = etree.Element(w("r"))
    if bold:
        rpr = etree.SubElement(r, w("rPr"))
        etree.SubElement(rpr, w("b"))
        etree.SubElement(rpr, w("bCs"))
    t = etree.SubElement(r, w("t"))
    t.text = text
    t.set("{%s}space" % XML, "preserve")
    if tab_after:
        etree.SubElement(r, w("tab"))
    return r


def prepend_run(p, run):
    ppr = p.find(w("pPr"))
    p.insert(0 if ppr is None else list(p).index(ppr) + 1, run)


def append_paragraph(dst, src, sep=" "):
    """Move the content of paragraph src onto the end of dst."""
    if sep:
        dst.append(make_run(sep))
    for ch in list(src):
        if ch.tag != w("pPr"):
            dst.append(ch)


def flatten_fields(p):
    """Replace fields (SEQ, REF, HYPERLINK...) with their displayed result."""
    for fs in list(p.iter(w("fldSimple"))):
        parent = fs.getparent()
        idx = list(parent).index(fs)
        for ch in list(fs):
            parent.insert(idx, ch)
            idx += 1
        parent.remove(fs)
    state = 0  # 0 outside, 1 in instruction, 2 in result
    for r in list(p.iter(w("r"))):
        fc = r.find(w("fldChar"))
        if fc is not None:
            kind = fc.get(w("fldCharType"))
            state = {"begin": 1, "separate": 2, "end": 0}.get(kind, state)
            r.getparent().remove(r)
        elif state == 1:
            r.getparent().remove(r)


def unwrap(el):
    parent = el.getparent()
    idx = list(parent).index(el)
    for ch in list(el):
        parent.insert(idx, ch)
        idx += 1
    parent.remove(el)


def onoff(el):
    return el is not None and el.get(w("val")) not in ("0", "false", "none", "off")


class Styles:
    """Read-only view of the student's styles, used only for classification."""

    def __init__(self, xml):
        self.by_id = {}
        if xml:
            root = etree.fromstring(xml)
            for st in root.findall(w("style")):
                self.by_id[st.get(w("styleId"))] = st

    def name(self, sid):
        st = self.by_id.get(sid)
        if st is None:
            return ""
        n = st.find(w("name"))
        return (n.get(w("val")) if n is not None else "").lower()

    def chain(self, sid):
        seen = set()
        while sid and sid not in seen and sid in self.by_id:
            seen.add(sid)
            st = self.by_id[sid]
            yield st
            b = st.find(w("basedOn"))
            sid = b.get(w("val")) if b is not None else None

    def run_prop(self, sid, tag):
        for st in self.chain(sid):
            el = st.find(w("rPr") + "/" + w(tag))
            if el is not None:
                return el
        return None

    def para_prop(self, sid, tag):
        for st in self.chain(sid):
            el = st.find(w("pPr") + "/" + w(tag))
            if el is not None:
                return el
        return None


class Para:
    """Facts about a paragraph gathered before it is restyled."""

    def __init__(self, p, styles):
        self.el = p
        self.text = ptext(p)
        self.stripped = self.text.strip()
        ppr = p.find(w("pPr"))
        sid = None
        if ppr is not None and ppr.find(w("pStyle")) is not None:
            sid = ppr.find(w("pStyle")).get(w("val"))
        self.style_name = styles.name(sid) if sid else ""
        jc = ppr.find(w("jc")) if ppr is not None else None
        if jc is None and sid:
            jc = styles.para_prop(sid, "jc")
        self.jc = jc.get(w("val")) if jc is not None else None
        self.centered = self.jc == "center"
        self.numbered = ppr is not None and ppr.find(w("numPr")) is not None and \
            (ppr.find(w("numPr") + "/" + w("numId")) is None or
             ppr.find(w("numPr") + "/" + w("numId")).get(w("val")) != "0")
        sp = ppr.find(w("spacing")) if ppr is not None else None
        self.exact_line = sp is not None and sp.get(w("lineRule")) == "exact"
        self.graphic = has_graphic(p)
        self.shape = has_shape(p)
        self.math = math_count(p)
        self.math_para = next(p.iter(m("oMathPara")), None) is not None
        self.textbox = next(p.iter(w("txbxContent")), None) is not None
        self.anchored = next(p.iter("{%s}anchor" % WP), None) is not None

        runs = [r for r in p.iter(w("r"))
                if any((t.text or "").strip() for t in r.findall(w("t")))]
        runs = [r for r in runs if not any(a.tag in OPAQUE for a in r.iterancestors())]

        def flag(r, tag):
            el = r.find(w("rPr") + "/" + w(tag))
            if el is None and sid:
                el = styles.run_prop(sid, tag)
            return onoff(el)

        def size(r):
            el = r.find(w("rPr") + "/" + w("sz"))
            if el is None and sid:
                el = styles.run_prop(sid, "sz")
            try:
                return int(el.get(w("val"))) / 2 if el is not None else None
            except (TypeError, ValueError):
                return None

        def font(r):
            el = r.find(w("rPr") + "/" + w("rFonts"))
            return ((el.get(w("ascii")) or "") if el is not None else "").lower()

        self.all_bold = bool(runs) and all(flag(r, "b") for r in runs)
        self.all_italic = bool(runs) and all(flag(r, "i") for r in runs)
        self.caps = bool(runs) and all(flag(r, "caps") or flag(r, "smallCaps") for r in runs)
        letters = [c for c in self.stripped if c.isalpha()]
        self.upper = len(letters) >= 3 and all(c.isupper() for c in letters)
        sizes = [s for s in (size(r) for r in runs) if s]
        self.max_size = max(sizes) if sizes else None
        self.mono = bool(runs) and all(any(f in font(r) for f in MONO_FONTS) for r in runs)

    @property
    def empty(self):
        return not self.stripped and not self.shape and not self.math


PICTURE_PARAGRAPH = """<w:p xmlns:w="%(w)s" xmlns:r="%(r)s"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:r><w:drawing>
 <wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%(cx)d" cy="%(cy)d"/>
 <wp:docPr id="%(n)d" name="%(name)s"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
 <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
 <pic:nvPicPr><pic:cNvPr id="0" name="%(name)s"/><pic:cNvPicPr/></pic:nvPicPr>
 <pic:blipFill><a:blip r:embed="%(rid)s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
 <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%(cx)d" cy="%(cy)d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
 </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"""


class Block:
    def __init__(self, role, el, info=None):
        self.role = role
        self.el = el
        self.info = info
        self.extra = {}


# --------------------------------------------------------------------------
# the formatter
# --------------------------------------------------------------------------

class Formatter:
    def __init__(self, src, template=DEFAULT_TEMPLATE):
        self.src = Path(src)
        with zipfile.ZipFile(self.src) as z:
            self.parts = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
        with zipfile.ZipFile(template) as z:
            self.tpl = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
        if "word/document.xml" not in self.parts:
            raise ValueError("%s is not a Word .docx document" % self.src)
        self.doc = etree.fromstring(self.parts["word/document.xml"])
        self.body = self.doc.find(w("body"))
        self.styles = Styles(self.parts.get("word/styles.xml"))
        self.fixes = []      # things done automatically
        self.flags = []      # (paragraph element or None, message)
        self.counts = {}
        self.title = ""

    # -- reporting ---------------------------------------------------------
    def fixed(self, msg):
        if msg not in self.fixes:
            self.fixes.append(msg)

    def flag(self, el, msg):
        self.flags.append((el, msg))

    # -- step 1: tidy the raw body ----------------------------------------
    def preprocess(self):
        body = self.body
        for tag in ("ins", "moveTo", "smartTag", "customXml"):
            for el in list(body.iter(w(tag))):
                unwrap(el)
        for tag in ("del", "moveFrom", "proofErr", "lastRenderedPageBreak"):
            for el in list(body.iter(w(tag))):
                el.getparent().remove(el)
        for sdt in list(body.iter(w("sdt"))):
            content = sdt.find(w("sdtContent"))
            parent = sdt.getparent()
            idx = list(parent).index(sdt)
            if content is not None:
                for ch in list(content):
                    parent.insert(idx, ch)
                    idx += 1
            parent.remove(sdt)
        for br in list(body.iter(w("br"))):
            if br.get(w("type")) in ("page", "column"):
                br.getparent().remove(br)
        cols = 1
        for sect in list(body.iter(w("sectPr"))):
            c = sect.find(w("cols"))
            if c is not None and c.get(w("num")):
                cols = max(cols, int(c.get(w("num"))))
            if sect.getparent() is not body:
                sect.getparent().remove(sect)
        if cols > 1:
            self.fixed("Converted a %d-column layout to single column" % cols)
        for hl in list(body.iter(w("hyperlink"))):
            unwrap(hl)
        self.unwrap_textbox_tables()
        self.fix_drop_caps()

    def fix_drop_caps(self):
        """A big first letter kept in a floating text box: put the letter back in front of the
        next paragraph (which then starts with the rest of the word) and drop the box."""
        fixed = 0
        paras = [q for q in self.body.iter(w("p"))
                 if not any(a.tag == w("txbxContent") for a in q.iterancestors())]
        for idx, p in enumerate(paras):
            if p.getparent() is None:
                continue
            boxes = [t for t in p.iter(w("txbxContent"))
                     if not any(a.tag == w("txbxContent") for a in t.iterancestors())]
            if not boxes:
                continue
            letters = "".join(x.text or "" for x in boxes[0].iter(w("t"))).strip()
            if not (1 <= len(letters) <= 2 and letters.isalpha() and letters.isupper()):
                continue
            sizes = [int(e.get(w("val"))) for e in boxes[0].iter(w("sz")) if (e.get(w("val")) or "").isdigit()]
            if not sizes or max(sizes) < 30:       # a drop cap is set in a much larger size
                continue
            nxt = next((q for q in paras[idx + 1: idx + 6] if ptext(q).strip()), None)
            if nxt is None:
                continue
            word = re.match(r"\s*([A-Za-z]+)", ptext(nxt))
            if not word or len(word.group(1)) < 2 or not (word.group(1).isupper() or word.group(1).islower()):
                continue
            for r in list(p.iter(w("r"))):
                if next(r.iter(w("txbxContent")), None) is not None and r.getparent() is not None:
                    r.getparent().remove(r)
            prepend_run(nxt, make_run(letters))
            fixed += 1
        if fixed:
            self.fixed("Put %d drop-cap letter(s) back into the text" % fixed)

    def unwrap_textbox_tables(self):
        """A PDF-to-Word conversion often puts a table inside a floating text box, where it
        overlaps the text.  Move such tables into the normal text flow."""
        moved = 0
        for p in list(self.body.iter(w("p"))):
            if p.getparent() is None or any(a.tag == w("txbxContent") for a in p.iterancestors()):
                continue
            boxes = [t for t in p.iter(w("txbxContent"))
                     if not any(a.tag == w("txbxContent") for a in t.iterancestors())
                     and t.find(w("tbl")) is not None]
            if not boxes:
                continue
            # the modern and the fallback copy of the box hold the same table: use the first
            tables = [copy.deepcopy(t) for t in boxes[0].findall(w("tbl"))]
            for r in list(p.iter(w("r"))):
                if next(r.iter(w("txbxContent")), None) is not None and r.getparent() is not None:
                    r.getparent().remove(r)
            parent = p.getparent()
            idx = list(parent).index(p) + 1
            for t in tables:
                parent.insert(idx, t)
                idx += 1
            moved += len(tables)
        if moved:
            self.fixed("Moved %d table(s) out of floating text boxes into the text" % moved)

    # -- step 2: classify --------------------------------------------------
    def heading_level(self, info, state, in_refs=False):
        """Return (level, numbered?, clean_text, uncertain) or None."""
        t = info.stripped
        if not t or info.graphic or info.math or len(t) > 150:
            return None
        style = info.style_name
        style_level = None
        ms = re.search(r"heading (\d)", style)
        if ms:
            style_level = min(int(ms.group(1)), 3)
        low = t.lower().rstrip(" :.")
        core_low = re.sub(r"^\s*(?:[ivxlc]+|[a-z]|\d+(?:\.\d+)*)\s*[.)]\s*", "", low).strip()

        def clean(pattern_match):
            return pattern_match.group(2).strip() if pattern_match else t

        mr = ROMAN_RE.match(t)
        if mr and from_roman(mr.group(1)) is None:
            mr = None
        ma = ALPHA_RE.match(t)
        mn = NUM_RE.match(t)
        strip_any = clean(mr or ma or mn)

        if low in UNNUMBERED_H1 or core_low in UNNUMBERED_H1:
            return 1, False, strip_any, False
        if in_refs:
            return None
        ends_sentence = t.endswith((".", ";", ",")) and not (mr or ma)

        letter_heading = ma is not None and ma.group(1) == state["next_letter"] and not (
            mr and from_roman(mr.group(1)) == state["next_roman"])      # "C. Objectives" after "B." is a sub-heading
        known = core_low in KNOWN_H1 and len(t) < 60 and (mr or not ma) and not letter_heading             and not (mn and "." in mn.group(1)) and not getattr(self, "has_h1_style", False)
        if known:                      # 'Introduction', 'Methodology :', 'II. Results' ... are always main sections
            return 1, True, strip_any, False
        if style_level:
            if style_level >= 2 and mr and info.upper and not (ma and ma.group(1) == state["next_letter"]):
                return 1, True, clean(mr), False
            return style_level, True, strip_any, False
        if mr and len(t) < 100:
            ambiguous = ma is not None and ma.group(1) == state["next_letter"] and \
                from_roman(mr.group(1)) != state["next_roman"]
            if not ambiguous and (info.upper or info.caps or info.centered or info.all_bold
                                  or core_low in KNOWN_H1):
                return 1, True, clean(mr), False
        if ma and len(t) < 100 and not t.endswith("."):
            if info.all_italic or info.all_bold or \
                    (ma.group(1) == state["next_letter"] and len(t) < 70 and state["h1_seen"]):
                return 2, True, clean(ma), not (info.all_italic or info.all_bold)
        if low in KNOWN_H1 or (core_low in KNOWN_H1 and (mn or mr) and len(t) < 60):
            if mn and "." in mn.group(1):
                return 2, True, clean(mn), False
            return 1, True, strip_any, False
        if mn and len(t) < 100 and not ends_sentence and (info.all_bold or info.upper or info.caps):
            depth = mn.group(1).count(".") + 1
            return min(depth, 3), True, clean(mn), False
        if info.numbered and len(t) < 100 and not t.endswith((".", ";", ":", ",")):
            if info.caps or info.upper:
                return 1, True, t, False
            if info.all_italic:
                return 2, True, t, False
        if ends_sentence or t.endswith(":"):
            return None
        if info.all_italic and not info.numbered and not info.centered and 4 <= len(t) < 70 \
                and len(t.split()) <= 9 and re.search(r"[A-Za-z]{3}", t):
            return 2, True, t, False   # a short italic line on its own is a sub-heading
        if info.centered and (info.all_bold or info.upper) and len(t) < 80 and len(t.split()) <= 10:
            return 1, True, t, True
        if info.all_bold and not info.centered and len(t) < 80 and len(t.split()) <= 10 \
                and not info.numbered:
            return 2, True, t, True
        return None

    def classify(self):
        raw = []
        for el in list(self.body):
            if el.tag == w("p"):
                info = Para(el, self.styles)
                if not info.empty:
                    raw.append(("p", el, info))
            elif el.tag == w("tbl"):
                raw.append(("tbl", el, None))
        self.raw = raw
        # a paper that uses 'Heading 1' for its sections already says what is a section: trust its styles
        self.has_h1_style = sum(1 for k, el, info in raw if k == "p" and re.search(r"heading 1", info.style_name)) >= 2

        def neighbour_is(i, kind, step):
            j = i + step
            if 0 <= j < len(raw):
                k, el, info = raw[j]
                if kind == "tbl":
                    return k == "tbl"
                return k == "p" and (info.graphic or (info.shape and not info.stripped)) \
                    and len(info.stripped) <= 3
            return False

        state = {"next_roman": 1, "next_letter": "A", "h1_seen": False}
        blocks = []

        # ---- front matter ----
        abstract_at = None
        first_h1 = None
        for i, (k, el, info) in enumerate(raw[:60]):
            if k != "p":
                continue
            if re.match(r"^\s*abstract\b", info.stripped, re.I) and abstract_at is None:
                abstract_at = i
                break
            if first_h1 is None:
                h = self.heading_level(info, state)
                if h and h[0] == 1 and not h[3] and i > 0:
                    first_h1 = i
        front_end = abstract_at if abstract_at is not None else (first_h1 if first_h1 is not None else 0)

        # lines of an author block that were separated by line breaks become paragraphs of their own
        front = []
        for k, el, info in raw[:front_end]:
            if k == "p" and info.stripped and not info.graphic and next(el.iter(w("br")), None) is not None \
                    and len(front) > 0:
                for part in split_lines(el):
                    pinfo = Para(part, self.styles)
                    if not pinfo.empty:
                        front.append(("p", part, pinfo))
            else:
                front.append((k, el, info))

        seen_title = False
        seen_affil = False
        for k, el, info in front:
            if k != "p":
                blocks.append(Block("table", el))
                continue
            if (info.graphic or (info.shape and not info.stripped)) and len(info.stripped) <= 3:
                blocks.append(Block("figure", el, info))
                self.flag(el, "Image found before the abstract; check it belongs here.")
                continue
            txt = info.stripped
            label = LABEL_RE.match(info.text)
            if not seen_title:
                seen_title = True
                self.title = txt
                blocks.append(Block("title", el, info))
            elif blocks and blocks[-1].role == "title" and "title" in info.style_name \
                    and not info.all_bold and info.max_size == blocks[-1].info.max_size:
                append_paragraph(blocks[-1].el, el)
                self.title += " " + txt
            elif EMAIL_RE.search(txt) and len(re.sub(r"(?i)e-?mail|[^A-Za-z0-9]", "",
                                                     EMAIL_RE.sub("", txt))) < 4:
                blocks.append(Block("dropped", el, info))          # a line that is only an e-mail address
                self.fixed("Removed e-mail address(es) from the author block")
            elif label:
                delete_prefix(el, label.end())
                info = Para(el, self.styles)
                b = Block("author", el, info)
                b.extra["guide"] = bool(re.search(r"guide|supervisor|mentor", label.group("label"), re.I))
                blocks.append(b)
            elif AFFIL_RE.search(txt) or (
                    seen_affil and not info.all_bold and not NAME_RE.fullmatch(txt)):
                seen_affil = True
                blocks.append(Block("affiliation", el, info))
            else:
                blocks.append(Block("author", el, info))

        i = front_end
        if abstract_at is not None:
            k, el, info = raw[i]
            rest = re.sub(r"^\s*abstract\s*[:.\-–—]*\s*", "", info.stripped, flags=re.I)
            b = Block("abstract", el, info)
            if not rest and i + 1 < len(raw) and raw[i + 1][0] == "p":
                # "Abstract" was a heading on its own line: use the next paragraph.
                i += 1
                b = Block("abstract", raw[i][1], raw[i][2])
                b.extra["label_missing"] = True
            blocks.append(b)
            i += 1
            while i < len(raw):
                k, el, info = raw[i]
                if k != "p":
                    break
                if re.match(r"^\s*(key\s*-?\s*words?|index\s+terms?)\b", info.stripped, re.I):
                    kb = Block("keywords", el, info)
                    rest = re.sub(r"^\s*(key\s*-?\s*words?|index\s+terms?)\s*[:.\-–—]*\s*", "",
                                  info.stripped, flags=re.I)
                    if not rest and i + 1 < len(raw) and raw[i + 1][0] == "p":
                        i += 1
                        kb = Block("keywords", raw[i][1], raw[i][2])
                        kb.extra["label_missing"] = True
                    blocks.append(kb)
                    i += 1
                    break
                if self.heading_level(info, state):
                    break
                blocks.append(Block("abstract_more", el, info))
                i += 1
            # the abstract is one paragraph, however many pieces the student's file cut it into
            abs_block = next(b for b in blocks if b.role == "abstract")
            for extra in [b for b in blocks if b.role == "abstract_more"]:
                append_paragraph(abs_block.el, extra.el)
                blocks.remove(extra)
            if not any(b.role == "keywords" for b in blocks):
                text = ptext(abs_block.el)
                hits = list(re.finditer(r"\bkey\s*-?\s*words?\b|\bindex\s+terms?\b", text, re.I))
                if hits and hits[-1].start() > 0.5 * len(text) and len(text) - hits[-1].start() < 400:
                    rest = split_paragraph_at(abs_block.el, hits[-1].start())
                    blocks.insert(blocks.index(abs_block) + 1, Block("keywords", rest, Para(rest, self.styles)))
                    self.fixed("Moved the keywords out of the abstract paragraph into their own line")
                else:
                    self.flag(abs_block.el, "No 'Keywords' line was found after the abstract.")
        else:
            self.flag(None, "No abstract was found at the top of the paper.")

        # ---- main body ----
        in_refs = False
        ref_explicit = 0
        while i < len(raw):
            k, el, info = raw[i]
            if k == "tbl":
                blocks.append(Block("table", el))
                i += 1
                continue
            t = info.stripped
            h = self.heading_level(info, state, in_refs)
            fig_adjacent = neighbour_is(i, "fig", -1) or neighbour_is(i, "fig", 1)
            tab_adjacent = neighbour_is(i, "tbl", -1) or neighbour_is(i, "tbl", 1)

            if (info.graphic or (info.shape and not t)) and len(t) <= 3:
                blocks.append(Block("figure", el, info))
                i += 1
                continue

            mf = FIG_RE.match(t)
            mt = TAB_RE.match(t)
            if mt and mt.group(2).isalpha() and from_roman(mt.group(2).upper()) is None:
                mt = None
            upper_cap = bool(mt) and bool(re.match(r"^\s*TABLE\s", info.text)) and len(t) < 250
            relaxed_fig = bool(mf) and len(t) < 250 and \
                bool(re.match(r"^\s*Fig(?:ure)?s?\.?\s*\d+\s*[.:]\s+\S", t))
            if mf and len(t) < 400 and not CAPTION_VERB_RE.match(mf.group(3)) and not in_refs and \
                    ("caption" in info.style_name or (info.centered and len(t) < 300) or relaxed_fig
                     or (fig_adjacent and len(t) < 250)):
                if info.graphic:
                    gp = self.split_graphic(el)
                    if gp is not None:
                        blocks.append(Block("figure", gp, Para(gp, self.styles)))
                        info = Para(el, self.styles)
                b = Block("fig_caption", el, info)
                b.extra = {"num": int(mf.group(2)), "prefix": len(t) - len(mf.group(3))}
                blocks.append(b)
                i += 1
                continue
            if mt and len(t) < 400 and not CAPTION_VERB_RE.match(mt.group(3)) and not in_refs and \
                    ("caption" in info.style_name or (info.centered and len(t) < 300) or upper_cap
                     or (tab_adjacent and len(t) < 250)
                     or (not mt.group(3) and i + 2 < len(raw) and raw[i + 2][0] == "tbl")):
                num = mt.group(2)
                n = int(num) if num.isdigit() else from_roman(num.upper())
                b = Block("table_caption", el, info)
                b.extra = {"num": n, "prefix": len(t) - len(mt.group(3))}
                # Title on the following line ("TABLE I" / "TITLE") -> join them.
                if not mt.group(3).strip() and i + 1 < len(raw) and raw[i + 1][0] == "p" \
                        and len(raw[i + 1][2].stripped) < 300 and not raw[i + 1][2].graphic:
                    append_paragraph(el, raw[i + 1][1])
                    i += 1
                blocks.append(b)
                i += 1
                continue

            if "caption" in info.style_name and not mf and not mt and len(t) < 300 and not in_refs:
                # a caption the student styled but never numbered
                tbl_side = neighbour_is(i, "tbl", 1) or (neighbour_is(i, "tbl", -1) and not fig_adjacent)
                b = Block("table_caption" if tbl_side else "fig_caption", el, info)
                b.extra = {"num": None, "prefix": 0}
                blocks.append(b)
                i += 1
                continue

            if h:
                level, numbered, text, uncertain = h
                b = Block("h%d" % level, el, info)
                b.extra = {"numbered": numbered, "text": text, "uncertain": uncertain}
                blocks.append(b)
                low = text.lower().rstrip(" :.")
                if level == 1:
                    in_refs = low in ("references", "reference")
                    state["h1_seen"] = True
                    if numbered and low not in UNNUMBERED_H1:
                        state["next_roman"] += 1
                    state["next_letter"] = "A"
                elif level == 2:
                    state["next_letter"] = chr(min(ord(state["next_letter"]) + 1, ord("Z")))
                i += 1
                continue

            if in_refs:
                explicit = bool(REFNUM_RE.match(t)) or info.numbered
                refs_so_far = [b for b in blocks if b.role == "reference"]
                if not explicit and refs_so_far and ref_explicit >= max(2, len(refs_so_far) - 1):
                    append_paragraph(refs_so_far[-1].el, el)
                    self.fixed("Joined a reference that was split over two paragraphs")
                else:
                    ref_explicit += explicit
                    blocks.append(Block("reference", el, info))
                i += 1
                continue

            if info.textbox and len(t) <= 3:
                blocks.append(Block("figure", el, info))
                i += 1
                continue

            # pieces of an equation left behind by a PDF-to-Word conversion
            words = re.findall(r"[A-Za-z]{3,}", t)
            if len(t) <= 30 and len(words) <= 2 and not info.numbered and not info.math and \
                    (info.exact_line or (info.max_size and info.max_size <= 7.5)
                     or re.search(r"[∑√⎷︂˜⌈-⌏⃓]", t)):
                blocks.append(Block("fragment", el, info))
                i += 1
                continue

            non_math = EQNUM_RE.sub("", t).strip(" \t.,;")
            if info.math and len(non_math) <= 15:
                blocks.append(Block("equation", el, info))
            elif not info.math and "=" in t and len(t) < 160 and not info.numbered and \
                    (info.centered or t != info.text.lstrip("\t ") and info.text.startswith("\t")
                     or (EQNUM_RE.search(t) and len(words) < 12)) and not t.endswith((".", ":")) or \
                    (not info.math and info.centered and "=" in t and len(t) < 160):
                blocks.append(Block("equation", el, info))
            elif info.mono:
                blocks.append(Block("code", el, info))
            elif info.numbered:
                blocks.append(Block("list", el, info))
            else:
                blocks.append(Block("body", el, info))
            i += 1
        self.blocks = blocks

    @staticmethod
    def tabs_to_spaces(p):
        """Tabs inside running text (left by column joins) become single spaces, and
        underlined blanks stop being underlined."""
        for el, kind in atoms(p):
            if kind == "tab":
                el.tag = w("t")
                el.text = " "
                el.set("{%s}space" % XML, "preserve")
        for r in p.iter(w("r")):
            if any(a.tag in OPAQUE for a in r.iterancestors()):
                continue
            text = "".join(t.text or "" for t in r.findall(w("t")))
            if text and not text.strip():
                u = r.find(w("rPr") + "/" + w("u"))
                if u is not None:
                    u.getparent().remove(u)
        for t in p.iter(w("t")):
            if t.text and "  " in t.text:
                t.text = re.sub(r" {2,}", " ", t.text)
                t.set("{%s}space" % XML, "preserve")

    @staticmethod
    def split_graphic(p):
        """Move the picture(s) of a paragraph into a new paragraph; return it (or None)."""
        runs = [r for r in p.findall(w("r"))
                if next(r.iter(w("drawing")), None) is not None
                or next(r.iter(w("pict")), None) is not None
                or next(r.iter("{%s}AlternateContent" % MC), None) is not None]
        if not runs:
            return None
        newp = etree.Element(w("p"))
        for r in runs:
            newp.append(r)
        return newp

    @staticmethod
    def anchors_to_inline(p):
        """Turn floating pictures into in-line ones so that they stay where the text puts them."""
        for anc in list(p.iter("{%s}anchor" % WP)):
            graphic = anc.find("{%s}graphic" % A)
            if graphic is None:
                continue
            data = graphic.find("{%s}graphicData" % A)
            if data is None or "picture" not in (data.get("uri") or ""):
                continue
            inline = etree.Element("{%s}inline" % WP, nsmap=anc.nsmap)
            for k in ("distT", "distB", "distL", "distR"):
                inline.set(k, "0")
            for tag in ("extent", "effectExtent", "docPr", "cNvGraphicFramePr"):
                e = anc.find("{%s}%s" % (WP, tag))
                if e is not None:
                    inline.append(copy.deepcopy(e))
            inline.append(copy.deepcopy(graphic))
            anc.getparent().replace(anc, inline)

    # -- tables the student's file only holds as loose text (PDF conversions) --------------
    def get_pdf(self):
        if not hasattr(self, "_pdf"):
            self._pdf = None
            try:
                import pymupdf
                provider = PDF_PROVIDER or default_pdf_provider
                data = provider(self.src)
                if data:
                    self._pdf = pymupdf.open(stream=data, filetype="pdf")
            except Exception:
                self._pdf = None
        return self._pdf

    def pair_captions(self):
        """Pair every real table with its caption; return the table captions left over."""
        blocks = self.blocks
        caps = [i for i, b in enumerate(blocks) if b.role == "table_caption"]
        used = set()
        for ti, b in enumerate(blocks):
            if b.role != "table":
                continue
            before = [c for c in caps if c < ti and c not in used and ti - c <= 14]
            after = [c for c in caps if c > ti and c not in used and c - ti <= 3]
            pick = before[-1] if before else (after[0] if after else None)
            if pick is not None:
                used.add(pick)
        return [c for c in caps if c not in used]

    def rebuild_loose_tables(self):
        orphans = self.pair_captions()
        if not orphans:
            return
        doc = self.get_pdf()
        for cap in [self.blocks[ci] for ci in orphans]:
            if doc is None:
                self.flag(cap.el, "This table is not a real table in the student's file (it is loose text "
                                  "from a PDF conversion) and could not be rebuilt here. Please insert it again.")
                continue
            self.table_from_pdf(cap, doc)

    def table_from_pdf(self, cap, doc):
        import pymupdf
        text = " ".join(ptext(cap.el).split())
        m = re.match(r"(?i)\s*table\s+([IVXLC]+|\d+)\b[.:\s-]*(.*)", text)
        if not m:
            return
        label = "TABLE " + m.group(1).upper()
        title_words = [re.sub(r"[^a-z0-9]", "", x.lower()) for x in m.group(2).split()]
        title_words = [x for x in title_words if x]

        # 1. find the caption line in the PDF (a short line that is just the label, or starts with it)
        found = None
        for page in doc:
            for r in page.search_for(label):
                line = " ".join(wd[4] for wd in page.get_text("words")
                                if abs((wd[1] + wd[3]) / 2 - (r.y0 + r.y1) / 2) < 3 and wd[0] >= r.x0 - 1)
                if line.upper().strip().startswith(label) and \
                        (len(line.strip()) <= len(label) + 2 or len(line) < 90) and \
                        not re.match(r"(?i)table\s+[ivxlc\d]+\s+(is|are|shows?|lists?|summari[sz]es|presents?|compares?|contrasts?)\b", line):
                    found = (page, r)
                    break
            if found:
                break
        if not found:
            self.flag(cap.el, "Could not find this table in the student's pages, so it was left as it is.")
            return
        page, cr = found
        words = page.get_text("words")
        # the table's rules (thin horizontal lines) below the caption
        rules = []
        for d in page.get_drawings():
            rc = d["rect"]
            if rc.height <= 2.5 and rc.width >= 40 and rc.y0 >= cr.y1 - 1:
                rules.append(rc)
        rules.sort(key=lambda r: r.y0)
        cx = (cr.x0 + cr.x1) / 2           # the caption is centred over its table (tables can sit side by side)
        first = next((r for r in rules if r.y0 - cr.y1 < 100 and r.x0 - 8 <= cx <= r.x1 + 8), None)
        if first is None:
            self.flag(cap.el, "Could not tell where this table ends in the student's page, so it was left as it is.")
            return
        table_rules = [first]
        for r in rules:
            if r is first:
                continue
            if r.x0 < first.x0 - 20 or r.x1 > first.x1 + 20 or r.x1 < first.x0 + 20:
                continue
            if r.y0 - table_rules[-1].y1 <= 130 and r.y0 > table_rules[-1].y0:
                table_rules.append(r)
        x0 = min(r.x0 for r in table_rules) - 3
        x1 = max(r.x1 for r in table_rules) + 3
        top = first.y0 - 1.5
        bottom = table_rules[-1].y1 + 3
        clip = pymupdf.Rect(x0, top, x1, bottom)
        self.__dict__.setdefault("_table_regions", []).append((page.number, clip))
        caption_zone = pymupdf.Rect(x0, cr.y1 - 1, x1, top)

        # 3. which loose paragraphs of the student's file are the table's cells
        inside = [wd for wd in words if clip.x0 - 1 <= wd[0] and wd[2] <= clip.x1 + 1
                  and clip.y0 - 1 <= wd[1] and wd[3] <= clip.y1 + 1]
        blob = re.sub(r"[^a-z0-9]", "", "".join(wd[4] for wd in inside).lower())
        if len(blob) < 10:
            self.flag(cap.el, "The student's page shows no table text here, so it was left as it is.")
            return
        region_words = {re.sub(r"[^a-z0-9]", "", wd[4].lower()) for wd in inside}
        zone_words = [wd for wd in words if caption_zone.x0 - 1 <= wd[0] and wd[2] <= caption_zone.x1 + 1
                      and caption_zone.y0 - 1 <= wd[1] and wd[3] <= caption_zone.y1 + 1]
        zone_blob = re.sub(r"[^a-z0-9]", "", "".join(wd[4] for wd in zone_words).lower())
        blocks = self.blocks
        ci = blocks.index(cap)
        gone, last = [], ci
        for j, b in enumerate(blocks):
            if j == ci or b.el.tag != w("p") or b.role in ("table_caption", "fig_caption", "title", "author",
                                                           "affiliation", "abstract", "keywords", "reference"):
                continue
            t = re.sub(r"[^a-z0-9]", "", ptext(b.el).lower())
            if t and len(t) < 4:
                if t in region_words:
                    gone.append(b)
                    last = max(last, j)
            elif t and len(t) <= 200 and t in blob:
                gone.append(b)
                last = max(last, j)
            elif t and len(t) <= 80 and t in zone_blob:
                append_paragraph(cap.el, b.el)          # the rest of the caption
                gone.append(b)
                last = max(last, j)
        shapes = [b for j, b in enumerate(blocks) if ci < j < last and b.role == "figure"
                  and not b.info.graphic and not b.info.stripped and next(b.el.iter(w("txbxContent")), None) is None]
        for b in gone + shapes:
            if b in blocks:
                blocks.remove(b)

        # 4. take the picture and put it where the table belongs
        scale = min(480.0 / clip.width, 1.25)
        png = page.get_pixmap(clip=clip, dpi=250).tobytes("png")
        n = len(getattr(self, "new_images", [])) + 1
        name = "ijarcce_table%d.png" % n
        rid = "rIdJImg%d" % n
        self.new_images = getattr(self, "new_images", []) + [(rid, name, png)]
        cx, cy = int(clip.width * scale * 12700), int(clip.height * scale * 12700)
        para = etree.fromstring(PICTURE_PARAGRAPH % {"w": W, "r": R, "cx": cx, "cy": cy, "rid": rid, "n": 9000 + n,
                                                    "name": name})
        newb = Block("tableimg", para, Para(para, self.styles))
        blocks.insert(blocks.index(cap) + 1, newb)
        self.flag(para, "%s is a picture of the student's table: in their file it was loose text from a PDF "
                        "conversion that could not be rebuilt as a table. Retype it as a Word table if it "
                        "must be editable." % label.title())
        self.fixed("Showed the tables that were only loose text in the student's file as pictures of their page")

    @staticmethod
    def is_line_shape(p):
        """True if every shape in the paragraph is a thin line (and there is no picture or text)."""
        shapes = 0
        for tag in (w("drawing"), w("pict")):
            for el in p.iter(tag):
                if el.getparent() is not None and any(a.tag == "{%s}Fallback" % MC for a in el.iterancestors()):
                    continue                    # the fallback copy of a shape already counted
                if next(el.iter("{%s}blip" % A), None) is not None or next(el.iter("{%s}imagedata" % V), None) is not None:
                    return False
                if any((t.text or "").strip() for t in el.iter(w("t"))):
                    return False
                thin = False
                for e in el.iter("{%s}extent" % WP):
                    if (e.get("cx") or "").isdigit() and (e.get("cy") or "").isdigit():
                        thin = thin or min(int(e.get("cx")), int(e.get("cy"))) <= 38100
                for e in el.iter("{%s}prstGeom" % A):
                    thin = thin or (e.get("prst") or "") in ("line", "straightConnector1")
                for e in el.iter("{%s}line" % V):
                    thin = True
                for e in el.iter():
                    st = e.get("style") or ""
                    for key in ("width", "height"):
                        m = re.search(r"(?:^|;)%s:\s*([\d.]+)pt" % key, st)
                        if m and float(m.group(1)) <= 3:
                            thin = True
                if not thin:
                    return False
                shapes += 1
        return shapes > 0

    def drop_rule_shapes(self):
        """Paragraphs that hold nothing but a thin drawn line (table rules left by a PDF conversion)."""
        gone = 0
        for b in list(self.blocks):
            if b.role != "figure" or b.info.stripped or not self.is_line_shape(b.el):
                continue
            self.blocks.remove(b)
            gone += 1
        if gone:
            self.fixed("Removed %d stray drawn line(s) left over from a PDF conversion" % gone)

    # -- equations a PDF conversion broke into pieces -----------------------------------
    def rebuild_loose_equations(self):
        if not any(b.role == "fragment" for b in self.blocks):
            return
        nums = {}
        for b in self.blocks:
            if b.el.tag != w("p") or b.role not in ("fragment", "equation", "body") or b.info.math:
                continue
            t = ptext(b.el).strip()
            m = EQNUM_RE.search(t) or re.search(r"\t\(\s*(\d+)\s*\)", t)
            if m and len(t) <= 140:
                nums.setdefault(int(re.sub(r"\D", "", m.group(0))), b)
        if not nums:
            return
        doc = self.get_pdf()
        if doc is None:
            return
        self._eqtext = [
            re.sub(r"[^a-z0-9]", "", ptext(b.el).lower()) for b in self.blocks
            if b.el.tag == w("p") and not b.info.math and (
                b.role in ("fragment", "equation") or (b.role == "body" and len(ptext(b.el)) <= 60))]
        for n in sorted(nums):
            self.equation_from_pdf(n, doc)
        slots = [i for i, b in enumerate(self.blocks) if b.role == "eqimg"]
        ordered = sorted((self.blocks[i] for i in slots), key=lambda b: b.extra["num"])
        for i, b in zip(slots, ordered):
            self.blocks[i] = b

    def equation_from_pdf(self, n, doc):
        token = "(%d)" % n
        for page in doc:
            words = page.get_text("words")
            for wd in words:
                if wd[4] == token and self.apply_equation(n, token, page, words, wd):
                    return True
        return False

    def apply_equation(self, n, token, page, words, wd):
        """Try one place where '(n)' stands in the student's page; True if it was an equation
        label whose pieces were found and replaced by a picture."""
        import pymupdf
        yc = (wd[1] + wd[3]) / 2
        label = pymupdf.Rect(wd[:4])
        if any(pn == page.number and label.intersects(r) for pn, r in getattr(self, "_table_regions", [])):
            return False                                    # a number inside a table, not an equation
        # justified text lines of the same column end at the same right edge as the number
        band = [x for x in words if 3 < abs((x[1] + x[3]) / 2 - yc) < 300 and abs(x[2] - wd[2]) <= 12]
        if len(band) < 3:
            return False
        lefts = []
        for e in band:                                      # walk left along each such line to find the column's left edge
            eyc = (e[1] + e[3]) / 2
            row = sorted([x for x in words if abs((x[1] + x[3]) / 2 - eyc) < 3 and x[0] < e[0]], key=lambda q: -q[0])
            left, count, prev = e[0], 1, e[0]
            for x in row:
                if prev - x[2] > 14:
                    break
                left, prev, count = x[0], x[0], count + 1
            if count >= 6:
                lefts.append(left)
        if not lefts:
            return False
        colx0 = min(lefts)
        colx1 = wd[2]
        colw = colx1 - colx0
        line = sorted([x for x in words if abs((x[1] + x[3]) / 2 - yc) < 3 and colx0 - 2 <= x[0] and x[2] <= wd[2] + 2],
                      key=lambda x: x[0])
        if not line or line[-1] is not wd or len(line) - 1 > 12:
            return False
        if any(0 < x[0] - wd[2] <= 6 and abs((x[1] + x[3]) / 2 - yc) < 3 for x in words):
            return False                                    # more words follow it on the line: a reference, not a label

        inc = [x for x in words if x[0] >= colx0 - 2 and x[2] <= colx1 + 2]
        rows = []
        for x in sorted(inc, key=lambda q: (q[1] + q[3]) / 2):
            yy = (x[1] + x[3]) / 2
            if rows and abs(rows[-1]["yc"] - yy) < 3.2:
                rows[-1]["w"].append(x)
            else:
                rows.append({"yc": yy, "w": [x]})
        taken = self.__dict__.setdefault("_eq_rows", set())
        for r in rows:
            r["x0"] = min(q[0] for q in r["w"])
            r["x1"] = max(q[2] for q in r["w"])
            r["y0"] = min(q[1] for q in r["w"])
            r["y1"] = max(q[3] for q in r["w"])
            r["key"] = (page.number, round(r["yc"] / 2), round(colx0 / 20))
        k0 = next(i for i, r in enumerate(rows) if wd in r["w"])
        eqtext = self._eqtext

        def equationish(r):
            if r["key"] in taken:
                return False
            if any(re.fullmatch(r"\(\d+\)", q[4]) and q is not wd for q in r["w"]):
                return False                                # the next equation's number: stop there
            if r["x0"] - colx0 < 10 and r["x1"] - r["x0"] >= 0.85 * colw:
                return False                                # a full text line
            rt = re.sub(r"[^a-z0-9]", "", "".join(q[4] for q in r["w"] if q is not wd).lower())
            if r["x0"] - colx0 < 8 and len(rt) >= 3:
                return False                                # a line of text starting at the margin
            return rt == "" or any(rt in piece for piece in eqtext)

        lo = hi = k0
        while lo > 0 and k0 - lo < 8 and rows[lo]["y0"] - rows[lo - 1]["y1"] <= 18 and equationish(rows[lo - 1]):
            lo -= 1
        while hi + 1 < len(rows) and hi - k0 < 8 and rows[hi + 1]["y0"] - rows[hi]["y1"] <= 18 and equationish(rows[hi + 1]):
            hi += 1
        used = rows[lo:hi + 1]
        pieces = [q for r in used for q in r["w"] if q is not wd]
        if not pieces:
            return False
        rect = pymupdf.Rect(min(q[0] for q in pieces), min(q[1] for q in pieces),
                            max(q[2] for q in pieces), max(q[3] for q in pieces))
        shapes = 0
        for d in page.get_drawings():                       # fraction bars, braces, roots
            rc = d["rect"]
            if rc.x0 >= colx0 - 2 and rc.x1 <= colx1 + 2 and rc.y1 >= rect.y0 - 6 and rc.y0 <= rect.y1 + 6 \
                    and rc.width < colw * 0.9 and rc.height < 80:
                rect |= rc
                shapes += 1
        rect = pymupdf.Rect(rect.x0 - 3, rect.y0 - 2, min(rect.x1 + 3, wd[0] - 2), rect.y1 + 2)
        inside = [q for q in words if rect.x0 - 2 <= q[0] and q[2] <= rect.x1 + 2 and rect.y0 - 2 <= q[1] and q[3] <= rect.y1 + 2]
        blob = re.sub(r"[^a-z0-9]", "", "".join(q[4] for q in inside).lower()) + str(n)
        raw = "".join(q[4] for q in inside)

        blocks = self.blocks
        gone = []
        for b in blocks:
            if b.el.tag != w("p") or b.role not in ("fragment", "equation", "body") or b.info.math:
                continue
            t = re.sub(r"[^a-z0-9]", "", ptext(b.el).lower())
            if t.isdigit() and t != str(n):
                continue                                    # another equation's number
            if t == str(n):
                gone.append(b)                              # the block that holds this equation's number
                continue
            if t and len(t) <= 160 and (t in blob if len(t) >= 2 else t in {re.sub(r"[^a-z0-9]", "", q[4].lower()) for q in inside}):
                gone.append(b)
        broken = any(b.role == "fragment" for b in gone) or shapes > 0 or len(used) >= 2 or any(
            re.search(r"[\u2308-\u23ff\ufe00-\ufe0f\u02dc\u20d3]", ptext(b.el)) for b in gone)
        if not gone or not broken:
            return False

        pos = min(blocks.index(b) for b in gone)
        for b in gone:
            blocks.remove(b)
        for b in blocks:                                    # one-glyph text boxes (subscripts, limits) of this equation
            if b.el.tag != w("p"):
                continue
            for r in list(b.el.iter(w("r"))):
                if r.getparent() is None:
                    continue
                box = next(r.iter(w("txbxContent")), None)
                if box is not None:
                    txt = "".join("".join(x.text or "" for x in box.iter(w("t"))).split())
                    if txt and len(txt) <= 6 and txt in raw:
                        r.getparent().remove(r)
        png = page.get_pixmap(clip=rect, dpi=250).tobytes("png")
        k = len(getattr(self, "new_images", [])) + 1
        name, rid = "ijarcce_equation%d.png" % k, "rIdJImg%d" % k
        self.new_images = getattr(self, "new_images", []) + [(rid, name, png)]
        cx, cy = int(rect.width * 12700), int(rect.height * 12700)
        para = etree.fromstring(PICTURE_PARAGRAPH % {"w": W, "r": R, "cx": cx, "cy": cy, "rid": rid, "n": 9000 + k,
                                                    "name": name})
        run = para.find(w("r"))
        tab1 = etree.Element(w("r"))
        etree.SubElement(tab1, w("tab"))
        run.addprevious(tab1)
        tab2 = etree.Element(w("r"))
        etree.SubElement(tab2, w("tab"))
        para.append(tab2)
        para.append(make_run(token))
        taken.update(r["key"] for r in used)
        eb = Block("eqimg", para, Para(para, self.styles))
        eb.extra["num"] = n
        blocks.insert(pos, eb)
        self.flag(para, "Equation (%d) is a picture of the student's equation: in their file it was broken into "
                        "pieces by a PDF conversion. Retype it with Word's equation editor if it must be "
                        "editable." % n)
        self.fixed("Showed the broken equations from the student's page as pictures")
        return True

    def promote_headings(self):
        """If every heading was typed at the same (lower) level, make them main sections."""
        blocks = self.blocks
        main_styles = {b.info.style_name for b in blocks if b.role == "h1" and "heading" in b.info.style_name
                       and b.extra["text"].lower().rstrip(" :.") in KNOWN_H1}
        peers = 0
        for b in blocks:
            if not self.has_h1_style and b.role == "h2" and b.info.style_name in main_styles and not ALPHA_RE.match(ptext(b.el)):
                b.role = "h1"
                peers += 1
        if peers:
            self.fixed("Treated %d heading(s) typed like the other main sections as main sections" % peers)
        numbered_h1 = any(b.role == "h1" and b.extra.get("numbered")
                          and b.extra["text"].lower().rstrip(" :.") not in UNNUMBERED_H1 for b in blocks)
        if numbered_h1 or not any(b.role in ("h2", "h3") for b in blocks):
            return
        for b in blocks:
            if b.role == "h2":
                b.role = "h1"
            elif b.role == "h3":
                b.role = "h2"
        self.fixed("All headings were at the same level; made them main sections")

    def fix_hyphenation(self):
        """'map-ping', 'dis-ease': a word cut at a line end by the student's source.  Rejoined
        only when the whole word appears elsewhere in the paper."""
        roles = ("body", "list", "abstract", "reference")
        ps = [b.el for b in self.blocks if b.role in roles and b.el.tag == w("p")]
        vocab = set()
        for el in ps:
            for tok in re.findall(r"(?<![\w-])[A-Za-z]{4,}(?![\w-])", ptext(el)):
                vocab.add(tok.lower())
        prefixes = {"pre", "post", "non", "multi", "anti", "semi", "self", "cross", "inter", "intra", "sub",
                    "over", "under", "bi", "uni", "pseudo", "meta", "micro", "macro", "mid", "co", "out", "de"}
        pat = re.compile(r"(?<![\w-])([A-Za-z]{2,})-[ \t]*([a-z]{2,})(?![\w-])")
        done = 0

        def fix(m):
            nonlocal done
            joined = (m.group(1) + m.group(2)).lower()
            left, right = m.group(1).lower(), m.group(2).lower()
            if left in prefixes or joined not in vocab:
                return m.group(0)
            if len(left) >= 4 and len(right) >= 4 and left in vocab and right in vocab:
                return m.group(0)       # two real words: a genuine compound such as 'risk-category'
            done += 1
            return m.group(1) + m.group(2)

        for el in ps:
            for t in el.iter(w("t")):
                if t.text and "-" in t.text:
                    t.text = pat.sub(fix, t.text)
        if done:
            self.fixed("Rejoined %d word(s) that had been hyphenated at a line end" % done)

    def merge_split_paragraphs(self):
        """A paragraph that stops mid-sentence and is continued (in lower case) further on was
        cut by a column or page break in the student's file.  Join the pieces."""
        terminal = (".", "?", "!", ":", ";", ")", "]", '"', "\u201d", "\u2019")
        skip = {"table_caption", "fig_caption", "figure", "table", "fragment", "equation"}
        blocks = self.blocks
        longs = [ptext(b.el).rstrip() for b in blocks if b.role == "body" and len(ptext(b.el)) > 60]
        if longs and sum(t.endswith(terminal) for t in longs) / len(longs) < 0.6:
            return          # this paper hardly uses punctuation: joining by punctuation would be guesswork
        joined = 0
        i = 0
        while i < len(blocks):
            a = blocks[i]
            text = ptext(a.el).rstrip() if a.role == "body" else ""
            if a.role == "body" and not a.info.graphic and len(text) > 25 and not text.endswith(terminal):
                j = i + 1
                while j < len(blocks) and j <= i + 5 and blocks[j].role in skip:
                    j += 1
                if j < len(blocks) and blocks[j].role == "body" and not blocks[j].info.graphic:
                    c = ptext(blocks[j].el).lstrip()[:1]
                    nxt = ptext(blocks[j].el).lstrip()
                    if c and not c.isupper() and c not in "[\u2022\u2013-" and \
                            not re.match(r"^[a-z]\.\s", nxt) and \
                            not re.match(r"^\d+\s*[:.)]\s", nxt) and \
                            not re.match(r"^(\d+\s*[:.)]|require|ensure|input|output)\b", text.lstrip(), re.I):
                        append_paragraph(a.el, blocks[j].el)
                        self.flag(a.el, "Two pieces of text were joined here because a column or page "
                                        "break had cut a paragraph. Please check that it reads correctly.")
                        del blocks[j]
                        joined += 1
                        continue            # the joined paragraph may continue again
            i += 1
        if joined:
            self.fixed("Joined %d paragraph(s) that had been cut in two by a column or page break" % joined)

    def group_authors(self):
        """Authors and affiliations as the journal wants them: 'A1, B2, C3' on one line and the
        affiliations numbered below.  Handles one block per author, several authors that share
        one affiliation, and several names typed on one line.  Identical affiliations are written
        once."""
        blocks = self.blocks
        idx = [i for i, b in enumerate(blocks) if b.role in ("author", "affiliation")]
        if not idx or blocks[idx[0]].role != "author":
            return
        seq = [blocks[i] for i in idx]

        def text_of(b):
            return " ".join(ptext(b.el).replace("[", "").replace("]", "").split()).strip(" ,;")

        groups = []                     # [author blocks, affiliation blocks]
        for b in seq:
            if b.role == "author":
                if groups and not groups[-1][1]:
                    groups[-1][0].append(b)
                else:
                    groups.append([[b], []])
            else:
                groups[-1][1].append(b)
        if any(not aff for _, aff in groups):
            return
        marks = re.compile(r"[\d*\u2020\u2021\u00b9\u00b2\u00b3\u2070-\u2079]")
        names = []                      # (name, group number)
        for g, (auths, _) in enumerate(groups):
            for b in auths:
                t = text_of(b)
                if marks.search(t) or next(b.el.iter(w("vertAlign")), None) is not None:
                    return              # already numbered by the student
                parts = [x.strip(" ,;") for x in re.split(r",\s*(?:and\s+)?|\s+and\s+|\s*&\s*", t) if x.strip(" ,;")]
                if not parts or any(len(x.split()) > 6 for x in parts):
                    return
                names += [(x, g) for x in parts]
        if len(names) < 2:
            return

        def norm(t):
            return re.sub(r"[^a-z0-9]", "", t.lower())

        affs = []                       # affiliation text per group
        for _, aff in groups:
            parts = [" ".join(ptext(x.el).split()).strip(" ,;") for x in aff]
            t = ", ".join(x for x in parts if x)
            affs.append(REGNO_RE.sub("", t).strip(" ,;"))
        unique = []                     # (key, text)
        for a in affs:
            if norm(a) not in [k for k, _ in unique]:
                unique.append((norm(a), a))
        refs = {k: [n + 1 for n, (_, g) in enumerate(names) if norm(affs[g]) == k] for k, _ in unique}

        def ref_text(nums):
            out, k = [], 0
            while k < len(nums):
                e = k
                while e + 1 < len(nums) and nums[e + 1] == nums[e] + 1:
                    e += 1
                out.append("%d-%d" % (nums[k], nums[e]) if e > k else str(nums[k]))
                k = e + 1
            return ",".join(out)

        def sup(text):
            r = make_run(text)
            rpr = etree.Element(w("rPr"))
            etree.SubElement(rpr, w("vertAlign")).set(w("val"), "superscript")
            r.insert(0, rpr)
            return r

        first_author, first_aff = groups[0][0][0], groups[0][1][0]
        new_author = etree.Element(w("p"))
        for k, (name, _) in enumerate(names):
            new_author.append(make_run(name))
            new_author.append(sup(str(k + 1)))
            if k + 1 < len(names):
                new_author.append(make_run(", "))
        new_blocks = [Block("author", new_author, first_author.info)]
        for key, text in unique:
            para = etree.Element(w("p"))
            para.append(make_run(text))
            para.append(sup(ref_text(refs[key])))
            new_blocks.append(Block("affiliation", para, first_aff.info))
        at = idx[0]
        drop = set(idx)
        self.blocks = blocks[:at] + new_blocks + [b for i, b in enumerate(blocks) if i > at and i not in drop]
        self.fixed("Put the authors on one line with numbered affiliations (%d author(s), %d affiliation(s))"
                   % (len(names), len(unique)))

    # -- step 3: restyle ---------------------------------------------------
    def set_ppr(self, p, style, keep=(), jc=None, ind=None):
        old = p.find(w("pPr"))
        new = etree.Element(w("pPr"))
        etree.SubElement(new, w("pStyle")).set(w("val"), style)
        if old is not None:
            for tag in keep:
                el = old.find(w(tag))
                if el is not None and tag != "jc":
                    new.append(el)
        if ind:
            e = etree.SubElement(new, w("ind"))
            for k_, v_ in ind.items():
                e.set(w(k_), str(v_))
        if jc is None and "jc" in keep and old is not None and old.find(w("jc")) is not None:
            jc = old.find(w("jc")).get(w("val"))
        if jc:
            etree.SubElement(new, w("jc")).set(w("val"), jc)
        if old is not None:
            p.remove(old)
        p.insert(0, new)

    def clean_runs(self, p, keep_emphasis=True, keep_color=False):
        keep = {w(x) for x in ("vertAlign", "strike", "dstrike", "rtl", "cs", "lang", "noProof",
                               "vanish", "em")}
        if keep_emphasis:
            keep |= {w(x) for x in ("b", "bCs", "i", "iCs", "u")}
        if keep_color:
            keep.add(w("color"))
        for r in p.iter(w("r")):
            if any(a.tag in OPAQUE for a in r.iterancestors() if a is not p):
                continue
            rpr = r.find(w("rPr"))
            if rpr is None:
                continue
            pos = rpr.find(w("position"))
            sz = rpr.find(w("sz"))
            for ch in list(rpr):
                if ch.tag in keep:
                    continue
                if ch.tag == w("rFonts"):
                    names = " ".join(v for v in ch.attrib.values()).lower()
                    if any(f in names for f in SYMBOL_FONTS) or r.find(w("sym")) is not None:
                        continue
                rpr.remove(ch)
            # PDF conversions fake super/subscripts with raised small text
            if pos is not None and rpr.find(w("vertAlign")) is None:
                try:
                    v = int(pos.get(w("val")))
                    small = sz is not None and int(sz.get(w("val"))) <= 16
                except (TypeError, ValueError):
                    v, small = 0, False
                if v and small:
                    etree.SubElement(rpr, w("vertAlign")).set(
                        w("val"), "superscript" if v > 0 else "subscript")
            if len(rpr) == 0:
                r.remove(rpr)

    def strip_lead(self, p):
        t = ptext(p)
        n = len(t) - len(t.lstrip(" \t \n"))
        if n:
            delete_prefix(p, n)
        rstrip_text(p)

    def label_paragraph(self, b, label, pattern):
        p = b.el
        self.strip_lead(p)
        if not b.extra.get("label_missing"):
            mm = re.match(pattern, ptext(p), re.I)
            if mm:
                delete_prefix(p, mm.end())
        self.clean_runs(p, keep_emphasis=False)
        prepend_run(p, make_run(label + ": ", bold=True))
        self.set_ppr(p, "IEEEAbtract")

    def restyle(self):
        h1 = h2 = h3 = 0
        ref_no = 0
        fig_nums, tab_nums = [], []
        for b in self.blocks:
            p = b.el
            self.counts[b.role] = self.counts.get(b.role, 0) + 1
            if b.role == "table":
                self.do_table(p)
                continue
            if b.role in ("fragment", "dropped"):
                continue
            if b.role == "figure":
                self.anchors_to_inline(p)
            if next(p.iter("{%s}anchor" % WP), None) is not None:
                self.flag(p, "Floating image or shape: its position may have moved. "
                             "Consider setting it 'In line with text'.")

            if b.role == "title":
                for br in list(p.iter(w("br"))):
                    if br.getparent() is not None and br.getparent().tag == w("r"):
                        br.tag = w("t")
                        br.attrib.clear()
                        br.text = " "
                self.strip_lead(p)
                self.clean_runs(p, keep_emphasis=False)
                self.set_ppr(p, "IEEETitle")
            elif b.role == "author":
                self.strip_lead(p)
                if "[" in b.info.text or "]" in b.info.text:
                    map_text(p, lambda s: s.replace("[", "").replace("]", ""))
                    self.fixed("Removed square brackets from author names")
                self.clean_runs(p, keep_emphasis=False)
                self.set_ppr(p, "IEEEAuthorName")
            elif b.role == "affiliation":
                self.strip_lead(p)
                txt = ptext(p)
                if EMAIL_RE.search(txt):
                    if re.search(r"corresponding", txt, re.I):
                        delete_prefix(p, len(txt))
                        prepend_run(p, make_run("*Corresponding Author"))
                    else:
                        map_text(p, lambda s: EMAIL_RE.sub("", s))
                    self.fixed("Removed e-mail address(es) from the author block")
                    if not ptext(p).strip(" ,;:|"):
                        b.role = "dropped"
                        continue
                if REGNO_RE.search(ptext(p)):
                    map_text(p, lambda s: REGNO_RE.sub("", s))
                    self.fixed("Removed registration numbers from the author block")
                self.clean_runs(p, keep_emphasis=False)
                self.set_ppr(p, "IEEEAuthorAffiliation")
            elif b.role == "abstract":
                self.label_paragraph(b, "Abstract", r"^\s*abstract\s*[:.\-–—]*\s*")
            elif b.role == "abstract_more":
                self.strip_lead(p)
                self.clean_runs(p, keep_emphasis=False)
                self.set_ppr(p, "IEEEAbtract")
            elif b.role == "keywords":
                self.label_paragraph(
                    b, "Keywords", r"^\s*(key\s*-?\s*words?|index\s+terms?)\s*[:.\-–—]*\s*")
            elif b.role in ("h1", "h2", "h3"):
                self.strip_lead(p)
                cur = ptext(p)
                want = b.extra["text"]
                cut = len(cur) - len(want) if cur.endswith(want) else 0
                if cut:
                    delete_prefix(p, cut)
                if b.role != "h3":
                    rstrip_chars(p, " :\u00a0\t")
                self.clean_runs(p, keep_emphasis=False)
                low = want.lower().rstrip(" :.")
                if b.role == "h1":
                    h2 = h3 = 0
                    map_text(p, lambda s: s.upper())
                    if low not in UNNUMBERED_H1 and b.extra["numbered"]:
                        h1 += 1
                        prepend_run(p, make_run(to_roman(h1) + ". "))
                    self.set_ppr(p, "IEEEHeading1")
                elif b.role == "h2":
                    h2 += 1
                    h3 = 0
                    prepend_run(p, make_run((chr(64 + h2) if h2 <= 26 else str(h2)) + ". "))
                    self.set_ppr(p, "IEEEHeading2")
                else:
                    h3 += 1
                    prepend_run(p, make_run("%d) " % h3))
                    self.set_ppr(p, "IEEEHeading3")
                if b.extra["uncertain"]:
                    self.flag(p, "Treated as a %s because of its formatting. If it is ordinary "
                                 "text, apply the 'IEEE Paragraph' style."
                              % ("section heading" if b.role == "h1" else "sub-heading"))
            elif b.role == "figure":
                self.resize_images(p)
                self.set_ppr(p, "IEEEFigure")
                for r in p.iter(w("r")):
                    rpr = r.find(w("rPr"))
                    if rpr is not None and not any(a.tag in OPAQUE for a in r.iterancestors()):
                        r.remove(rpr)
                if b.info.textbox:
                    self.flag(p, "Text box or drawing shape: not reformatted, please check it.")
            elif b.role == "fig_caption":
                flatten_fields(p)
                self.strip_lead(p)
                mm = FIG_RE.match(ptext(p))
                if mm:
                    delete_prefix(p, mm.start(3))
                self.clean_runs(p, keep_emphasis=False)
                num = b.extra["num"]
                if num is None:
                    num = fig_nums[-1][0] + 1 if fig_nums else 1
                    self.flag(p, "This caption had no figure number, so it was numbered %d automatically. "
                                 "Check that the text refers to it as Fig. %d." % (num, num))
                prepend_run(p, make_run("Fig. %d " % num))
                self.set_ppr(p, "IEEEFigureCaptionMulti-Lines")
                fig_nums.append((num, p))
            elif b.role == "table_caption":
                flatten_fields(p)
                self.strip_lead(p)
                mm = TAB_RE.match(ptext(p))
                if mm:
                    delete_prefix(p, mm.start(3))
                self.strip_lead(p)
                self.clean_runs(p, keep_emphasis=False)
                map_text(p, lambda s: s.upper())
                n = b.extra["num"]
                if n is None:
                    n = tab_nums[-1][0] + 1 if tab_nums else 1
                    self.flag(p, "This caption had no table number, so it was numbered %s automatically. "
                                 "Check that the text refers to it as Table %s." % (to_roman(n), to_roman(n)))
                n = n or 0
                prepend_run(p, make_run("TABLE %s " % (to_roman(n) if n else "?")))
                self.set_ppr(p, "IEEETableCaption")
                tab_nums.append((n, p))
            elif b.role == "reference":
                self.strip_lead(p)
                mm = REFNUM_RE.match(ptext(p))
                if mm:
                    delete_prefix(p, mm.end())
                self.strip_lead(p)
                self.clean_runs(p, keep_emphasis=True)
                for r in p.iter(w("r")):
                    rpr = r.find(w("rPr"))
                    if rpr is not None:
                        for tag in ("b", "bCs", "u"):
                            e = rpr.find(w(tag))
                            if e is not None:
                                rpr.remove(e)
                ref_no += 1
                prepend_run(p, make_run("[%d]" % ref_no, tab_after=True))
                self.set_ppr(p, "IEEEReferenceItem")
            elif b.role == "tableimg":
                self.set_ppr(p, "IEEEFigure")
            elif b.role == "eqimg":
                self.set_ppr(p, "IEEEEquation")
            elif b.role == "equation":
                self.do_equation(b)
            elif b.role == "code":
                self.clean_runs(p)
                self.set_ppr(p, "IEEECode")
            elif b.role == "list":
                self.strip_lead(p)
                self.clean_runs(p)
                self.set_ppr(p, "IEEEListItem", keep=("numPr",), ind={"left": 540, "hanging": 360})
            else:
                self.strip_lead(p)
                self.tabs_to_spaces(p)
                self.clean_runs(p)
                if b.info.graphic:
                    self.resize_images(p)
                    self.flag(p, "Image sits inside a line of text; check its position.")
                self.set_ppr(p, "IEEEParagraph")

        for label, nums in (("Figure", fig_nums), ("Table", tab_nums)):
            expect = 1
            for n, p in nums:
                if n != expect:
                    self.flag(p, "%s numbering is out of sequence here (found %s, expected %s). "
                                 "Captions keep the student's numbers; please check the text too."
                              % (label, n, expect))
                    expect = n
                expect += 1
        self.blocks = [b for b in self.blocks if b.role != "dropped"]
        self.ref_count = ref_no

    def do_equation(self, b):
        """Centre the equation; if it has a number like (3), push that to the right margin."""
        p = b.el
        self.clean_runs(p)
        if b.info.math_para:
            self.set_ppr(p, "IEEEParagraph", jc="center")
            return
        # Walk the paragraph in order (maths included): drop tabs and spaces in
        # front of the equation, and collapse runs of tabs after it.
        seen = False
        prev_tab = False
        last_tab = None
        for el in list(p.iter()):
            if any(a.tag in OPAQUE for a in el.iterancestors() if a is not p):
                continue
            if el.tag in (m("oMath"), w("drawing"), w("object"), w("pict")):
                seen, prev_tab = True, False
            elif el.tag == w("tab") and el.getparent().tag == w("r"):
                if not seen or prev_tab:
                    el.getparent().remove(el)
                else:
                    prev_tab, last_tab = True, el
            elif el.tag == w("t"):
                if not seen:
                    el.text = (el.text or "").lstrip(" 	 ")
                if (el.text or "").strip():
                    seen, prev_tab = True, False
        rstrip_text(p)
        text = ptext(p)
        mm = EQNUM_RE.search(text)
        numbered = False
        if mm:
            if text[:mm.start()].endswith("	"):
                numbered = True
            else:
                # number typed straight after the equation: put a tab before it
                for el, kind in reversed(atoms(p)):
                    if kind == "t" and (el.text or "").strip():
                        if EQNUM_RE.fullmatch((el.text or "").strip()):
                            el.text = el.text.strip()
                            el.addprevious(etree.Element(w("tab")))
                            numbered = True
                        break
        if numbered:
            r = etree.Element(w("r"))
            etree.SubElement(r, w("tab"))
            prepend_run(p, r)
            self.set_ppr(p, "IEEEEquation")
        else:
            self.set_ppr(p, "IEEEParagraph", jc="center")

    def resize_images(self, p):
        inlines = list(p.iter("{%s}inline" % WP))
        if not inlines:
            return
        total = 0
        tallest = 0
        for il in inlines:
            ext = il.find("{%s}extent" % WP)
            if ext is not None:
                total += int(ext.get("cx"))
                tallest = max(tallest, int(ext.get("cy")))
        factor = 1.0
        if total > TEXT_EMU:
            factor = TEXT_EMU / total
        if tallest * factor > MAX_IMG_H_EMU:
            factor = MAX_IMG_H_EMU / tallest
        if factor >= 0.999:
            return
        for il in inlines:
            for ext in [il.find("{%s}extent" % WP)] + list(il.iter("{%s}ext" % A)):
                if ext is not None and ext.get("cx") and ext.get("cy"):
                    ext.set("cx", str(int(int(ext.get("cx")) * factor)))
                    ext.set("cy", str(int(int(ext.get("cy")) * factor)))
        self.fixed("Shrank image(s) that were wider or taller than the page area")

    def do_table(self, tbl):
        tblpr = tbl.find(w("tblPr"))
        if tblpr is None:
            tblpr = etree.Element(w("tblPr"))
            tbl.insert(0, tblpr)
        for tag in ("tblpPr", "tblInd", "jc", "tblOverlap"):
            e = tblpr.find(w(tag))
            if e is not None:
                tblpr.remove(e)
        tw = tblpr.find(w("tblW"))
        grid_total = 0
        grid = tbl.find(w("tblGrid"))
        if grid is not None:
            for gc in grid.findall(w("gridCol")):
                try:
                    grid_total += int(gc.get(w("w")) or 0)
                except ValueError:
                    pass
        width = 0
        if tw is not None and tw.get(w("type")) == "dxa":
            try:
                width = int(tw.get(w("w")))
            except (TypeError, ValueError):
                width = 0
        if max(width, grid_total) > TEXT_WIDTH + 20:
            if tw is None:
                tw = etree.SubElement(tblpr, w("tblW"))
            tw.set(w("w"), "5000")
            tw.set(w("type"), "pct")
            lay = tblpr.find(w("tblLayout"))
            if lay is not None:
                tblpr.remove(lay)
            self.fixed("Narrowed table(s) that were wider than the page area")
        # children of tblPr must stay in schema order
        order = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
                 "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
                 "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
        jc = etree.Element(w("jc"))
        jc.set(w("val"), "center")
        tblpr.append(jc)
        old = tblpr.find(w("tblBorders"))
        full_grid = tblpr.find(w("tblStyle")) is not None or (
            old is not None and old.find(w("insideH")) is not None and old.find(w("insideV")) is not None
            and all(e.get(w("val")) not in ("nil", "none") for e in old))
        if not full_grid:
            if old is not None:
                tblpr.remove(old)
            for tc_b in list(tbl.iter(w("tcBorders"))):
                tc_b.getparent().remove(tc_b)
            borders = etree.SubElement(tblpr, w("tblBorders"))
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                e = etree.SubElement(borders, w(side))
                e.set(w("val"), "single")
                e.set(w("sz"), "4")
                e.set(w("space"), "0")
                e.set(w("color"), "auto")
            self.fixed("Gave table(s) a full grid of borders")
        kids = sorted(tblpr, key=lambda e: order.index(etree.QName(e).localname)
                      if etree.QName(e).localname in order else 99)
        for e in kids:
            tblpr.append(e)
        for trpr in tbl.iter(w("trPr")):
            for e in trpr.findall(w("jc")):
                trpr.remove(e)
        for p in tbl.iter(w("p")):
            if any(a.tag in OPAQUE for a in p.iterancestors()):
                continue
            shaded = False
            for a in p.iterancestors(w("tc")):
                shd = a.find(w("tcPr") + "/" + w("shd"))
                fill = (shd.get(w("fill")) or "").upper() if shd is not None else ""
                shaded = fill not in ("", "AUTO", "FFFFFF")
                break
            self.clean_runs(p, keep_color=shaded)
            self.set_ppr(p, "IEEETableCell", keep=("jc", "numPr"))
        first_row = tbl.find(w("tr"))
        if first_row is not None:
            for r in first_row.iter(w("r")):
                if any(a.tag in OPAQUE for a in r.iterancestors()):
                    continue
                rpr = r.find(w("rPr"))
                if rpr is None:
                    rpr = etree.Element(w("rPr"))
                    r.insert(0, rpr)
                if rpr.find(w("b")) is None:
                    pos = sum(1 for ch in rpr if ch.tag in (w("rStyle"), w("rFonts")))
                    rpr.insert(pos, etree.Element(w("b")))
        if next(tbl.iter(w("drawing")), None) is not None or next(tbl.iter(w("pict")), None) is not None:
            first = next(tbl.iter(w("p")), None)
            self.flag(first, "This table contains image(s); check the layout.")
            for p in tbl.iter(w("p")):
                self.resize_images(p)

    # -- step 4: spacing and order ----------------------------------------
    @staticmethod
    def gap(prev, cur, prev_b, cur_b):
        if prev is None:
            return False
        same_group = {("author", "author"), ("author", "affiliation"), ("affiliation", "affiliation"),
                      ("list", "list"), ("reference", "reference"), ("code", "code"),
                      ("fragment", "fragment"), ("figure", "fig_caption"), ("figure", "figure"),
                      ("abstract", "abstract_more"), ("abstract_more", "abstract_more"),
                      ("h2", "body"), ("h3", "body"), ("h2", "list"), ("h3", "list"),
                      ("body", "list"), ("h2", "h3")}
        if (prev, cur) in same_group:
            return False
        if prev == "body" and cur == "body":
            # runs of short lines (algorithm steps, addresses) stay together
            a, c = prev_b.info.stripped, cur_b.info.stripped
            step = re.compile(r"^(\d+\s*[:.)]|(require|ensure|input|output|do|step)\b)", re.I)
            return not (len(a) < 90 and len(c) < 90
                        and (not a.endswith(".") or (step.match(a) and step.match(c))))
        return True

    def layout(self):
        blocks = self.blocks
        # caption on the wrong side of its table / figure -> move it
        i = 0
        while i < len(blocks) - 1:
            a, b = blocks[i], blocks[i + 1]
            before = blocks[i - 1].role if i else None
            after = blocks[i + 2].role if i + 2 < len(blocks) else None
            if a.role == "table" and b.role == "table_caption" and before != "table_caption" \
                    and after != "table":
                blocks[i], blocks[i + 1] = b, a
                self.fixed("Moved table caption(s) above the table")
                i += 2
                continue
            if a.role == "fig_caption" and b.role == "figure" and before != "figure" \
                    and after != "fig_caption":
                blocks[i], blocks[i + 1] = b, a
                self.fixed("Moved figure caption(s) below the figure")
                i += 2
                continue
            i += 1
        k = 0
        while k < len(blocks):
            if blocks[k].role == "table" and not (k and blocks[k - 1].role == "table_caption"):
                for back in range(2, 5):
                    c = k - back
                    if c >= 0 and blocks[c].role == "table_caption" and \
                            not (c + 1 < len(blocks) and blocks[c + 1].role == "table"):
                        cap = blocks.pop(c)
                        blocks.insert(k - 1, cap)
                        break
            k += 1
        for idx, b in enumerate(blocks):
            if b.role == "table":
                if not (idx and blocks[idx - 1].role == "table_caption"):
                    first = next(b.el.iter(w("p")), None)
                    self.flag(first, "No caption was found for this table.")
            if b.role == "figure" and not b.info.textbox:
                nxt = blocks[idx + 1].role if idx + 1 < len(blocks) else None
                if nxt not in ("fig_caption", "figure"):
                    self.flag(b.el, "No caption was found under this figure.")

        sect = self.body.find(w("sectPr"))
        for el in list(self.body):
            self.body.remove(el)
        prev = None
        for b in blocks:
            if self.gap(prev.role if prev else None, b.role, prev, b):
                blank = etree.SubElement(self.body, w("p"))
                ppr = etree.SubElement(blank, w("pPr"))
                style = "IEEEParagraph"      # the blank line after the title is 10 pt, not 24 pt
                etree.SubElement(ppr, w("pStyle")).set(w("val"), style)
                if prev.role == "table_caption":        # keep the caption, the blank line and the table together
                    etree.SubElement(ppr, w("keepNext"))
            self.body.append(b.el)
            prev = b
        self.old_sect = sect

    # -- step 5: package-level changes ------------------------------------
    def merge_styles(self):
        tpl = etree.fromstring(self.tpl["word/styles.xml"])
        have = {s.get(w("styleId")) for s in tpl.findall(w("style"))}
        names = {s.find(w("name")).get(w("val")).lower() for s in tpl.findall(w("style"))
                 if s.find(w("name")) is not None}
        if "word/styles.xml" in self.parts:
            mine = etree.fromstring(self.parts["word/styles.xml"])
            for st in mine.findall(w("style")):
                nm = st.find(w("name"))
                if st.get(w("styleId")) in have or st.get(w("default")) in ("1", "true"):
                    continue
                if nm is not None and nm.get(w("val")).lower() in names:
                    continue
                tpl.append(copy.deepcopy(st))
        self.parts["word/styles.xml"] = etree.tostring(
            tpl, xml_declaration=True, encoding="UTF-8", standalone=True)

    def rels_path(self, part):
        folder, name = part.rsplit("/", 1)
        return "%s/_rels/%s.rels" % (folder, name)

    def install_header_footer(self, meta):
        rels_name = "word/_rels/document.xml.rels"
        rels = etree.fromstring(self.parts[rels_name])
        for rel in list(rels):
            if rel.get("Type", "").endswith(("/header", "/footer")):
                part = "word/" + rel.get("Target").lstrip("/").replace("word/", "")
                self.parts.pop(part, None)
                self.parts.pop(self.rels_path(part), None)
                rels.remove(rel)
        used = {r.get("Id") for r in rels}

        def new_id(base):
            n = 1
            while "%s%d" % (base, n) in used:
                n += 1
            used.add("%s%d" % (base, n))
            return "%s%d" % (base, n)

        # which header/footer does the template's first section use?
        tdoc = self.tpl["word/document.xml"].decode("utf8")
        trels = self.tpl["word/_rels/document.xml.rels"].decode("utf8")

        def tpl_part(kind):
            rid = re.search(r'<w:%sReference w:type="default" r:id="(\w+)"' % kind, tdoc).group(1)
            rel = re.search(r'<Relationship [^>]*Id="%s"[^>]*>' % rid, trels).group(0)
            return "word/" + re.search(r'Target="([^"]+)"', rel).group(1)

        ids = {}
        for kind in ("header", "footer"):
            src = tpl_part(kind)
            dst = "word/ijarcce_%s.xml" % kind
            data = self.tpl[src].decode("utf8")
            if not meta.get("DOI"):
                data = strip_doi(data)                        # "Without DOI": no DOI line at all
            for key, val in meta.items():
                data = data.replace("{{%s}}" % key, str(val))
            self.parts[dst] = data.encode("utf8")
            src_rels = self.rels_path(src)
            if src_rels in self.tpl:
                self.parts[self.rels_path(dst)] = self.tpl[src_rels]
                for target in re.findall(r'Target="(media/[^"]+)"', self.tpl[src_rels].decode("utf8")):
                    self.parts["word/" + target] = self.tpl["word/" + target]
            rid = new_id("rIdJ")
            ids[kind] = rid
            rel = etree.SubElement(rels, "{%s}Relationship" % REL_NS)
            rel.set("Id", rid)
            rel.set("Type", "%s/%s" % (R, kind))
            rel.set("Target", "ijarcce_%s.xml" % kind)
        self.parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

        sect = etree.SubElement(self.body, w("sectPr"))
        for kind in ("header", "footer"):
            ref = etree.SubElement(sect, w(kind + "Reference"))
            ref.set(w("type"), "default")
            ref.set("{%s}id" % R, ids[kind])
        e = etree.SubElement(sect, w("pgSz"))
        e.set(w("w"), str(PAGE_W)); e.set(w("h"), str(PAGE_H)); e.set(w("code"), "9")
        e = etree.SubElement(sect, w("pgMar"))
        for k_ in ("top", "right", "bottom", "left"):
            e.set(w(k_), str(MARGIN))
        e.set(w("header"), "720"); e.set(w("footer"), "720"); e.set(w("gutter"), "0")
        if meta.get("START_PAGE"):
            etree.SubElement(sect, w("pgNumType")).set(w("start"), str(meta["START_PAGE"]))
        etree.SubElement(sect, w("cols")).set(w("space"), "720")
        etree.SubElement(sect, w("docGrid")).set(w("linePitch"), "360")

        if "word/settings.xml" in self.parts:
            s = self.parts["word/settings.xml"].decode("utf8")
            s = re.sub(r"<w:(evenAndOddHeaders|mirrorMargins|gutterAtTop|trackRevisions)\b[^>]*/>", "", s)
            s = re.sub(r"<w:documentProtection\b[^>]*/>", "", s)
            self.parts["word/settings.xml"] = s.encode("utf8")

    def add_new_images(self):
        imgs = getattr(self, "new_images", [])
        if not imgs:
            return
        rels_name = "word/_rels/document.xml.rels"
        rels = etree.fromstring(self.parts[rels_name])
        for rid, name, png in imgs:
            self.parts["word/media/" + name] = png
            rel = etree.SubElement(rels, "{%s}Relationship" % REL_NS)
            rel.set("Id", rid)
            rel.set("Type", R + "/image")
            rel.set("Target", "media/" + name)
        self.parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    def add_comments(self):
        notes = [(el, msg) for el, msg in self.flags if el is not None and el.getparent() is not None]
        if not notes:
            return
        name = "word/comments.xml"
        if name in self.parts:
            root = etree.fromstring(self.parts[name])
        else:
            root = etree.Element(w("comments"), nsmap={"w": W})
            rels_name = "word/_rels/document.xml.rels"
            rels = etree.fromstring(self.parts[rels_name])
            rel = etree.SubElement(rels, "{%s}Relationship" % REL_NS)
            rel.set("Id", "rIdJComments")
            rel.set("Type", R + "/comments")
            rel.set("Target", "comments.xml")
            self.parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8",
                                                   standalone=True)
        ids = [int(c.get(w("id"))) for c in root.findall(w("comment")) if (c.get(w("id")) or "").isdigit()]
        ids += [int(e.get(w("id"))) for e in self.doc.iter(w("commentRangeStart"))
                if (e.get(w("id")) or "").isdigit()]
        nxt = max(ids) + 1 if ids else 0
        for el, msg in notes:
            cid = str(nxt)
            nxt += 1
            c = etree.SubElement(root, w("comment"))
            c.set(w("id"), cid)
            c.set(w("author"), "Formatter")
            c.set(w("initials"), "F")
            cp = etree.SubElement(c, w("p"))
            cr = etree.SubElement(cp, w("r"))
            etree.SubElement(cr, w("t")).text = msg
            start = etree.Element(w("commentRangeStart"))
            start.set(w("id"), cid)
            ppr = el.find(w("pPr"))
            el.insert(0 if ppr is None else list(el).index(ppr) + 1, start)
            end = etree.SubElement(el, w("commentRangeEnd"))
            end.set(w("id"), cid)
            rr = etree.SubElement(el, w("r"))
            etree.SubElement(rr, w("commentReference")).set(w("id"), cid)
        self.parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def fix_content_types(self):
        ct = etree.fromstring(self.parts["[Content_Types].xml"])
        ns = ct.nsmap[None]
        for ov in list(ct.findall("{%s}Override" % ns)):
            if ov.get("PartName").lstrip("/") not in self.parts:
                ct.remove(ov)
        have = {ov.get("PartName") for ov in ct.findall("{%s}Override" % ns)}
        base = "application/vnd.openxmlformats-officedocument.wordprocessingml."
        for part, kind in (("/word/ijarcce_header.xml", "header+xml"),
                           ("/word/ijarcce_footer.xml", "footer+xml"),
                           ("/word/comments.xml", "comments+xml"),
                           ("/word/styles.xml", "styles+xml")):
            if part.lstrip("/") in self.parts and part not in have:
                ov = etree.SubElement(ct, "{%s}Override" % ns)
                ov.set("PartName", part)
                ov.set("ContentType", base + kind)
        exts = {d.get("Extension").lower() for d in ct.findall("{%s}Default" % ns)}
        if "png" not in exts:
            d = etree.Element("{%s}Default" % ns)
            d.set("Extension", "png")
            d.set("ContentType", "image/png")
            ct.insert(0, d)
        self.parts["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        if "word/styles.xml" in self.parts:
            rels_name = "word/_rels/document.xml.rels"
            rels = etree.fromstring(self.parts[rels_name])
            if not any(r.get("Type", "").endswith("/styles") for r in rels):
                rel = etree.SubElement(rels, "{%s}Relationship" % REL_NS)
                rel.set("Id", "rIdJStyles")
                rel.set("Type", R + "/styles")
                rel.set("Target", "styles.xml")
                self.parts[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8",
                                                       standalone=True)

    # -- driver -------------------------------------------------------------
    def run(self, out, meta):
        self.preprocess()
        self.classify()
        self.promote_headings()
        self.rebuild_loose_tables()
        self.drop_rule_shapes()
        self.rebuild_loose_equations()
        self.merge_split_paragraphs()
        self.fix_hyphenation()
        self.group_authors()
        self.restyle()
        self.layout()
        self.document_checks()
        self.merge_styles()
        self.install_header_footer(meta)
        self.add_comments()
        self.add_new_images()
        self.parts["word/document.xml"] = etree.tostring(
            self.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.fix_content_types()
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            order = ["[Content_Types].xml"] + [n for n in self.parts if n != "[Content_Types].xml"]
            for n in order:
                z.writestr(n, self.parts[n])
        return out

    def document_checks(self):
        c = self.counts
        frag = c.get("fragment", 0)
        if frag:
            first = next(b.el for b in self.blocks if b.role == "fragment")
            self.flag(first, "This paper looks like it was converted from a PDF: %d small text "
                             "fragments appear to be pieces of broken equations. They were left "
                             "exactly as submitted; the equations need retyping with Word's "
                             "equation editor." % frag)
            prev = None
            for b in self.blocks:
                if b.role == "fragment" and prev != "fragment" and b.el is not first:
                    self.flag(b.el, "Broken equation fragment(s) start here.")
                prev = b.role
        if not c.get("title"):
            self.flag(None, "No title was identified.")
        if not c.get("author"):
            self.flag(None, "No author-name line was identified; check the block under the title.")
        if not c.get("reference"):
            self.flag(None, "No reference list was found (looked for a 'References' heading).")
        if not c.get("h1"):
            self.flag(None, "No section headings were recognised.")
        for b in self.blocks:
            if b.role == "equation" and not b.info.math:
                self.flag(b.el, "Equation typed as plain text; consider retyping it with the "
                                "equation editor.")
                break
        # the same message on many paragraphs is noise: keep the first three
        seen = {}
        kept = []
        for el, msg in self.flags:
            seen[msg] = seen.get(msg, 0) + 1
            if seen[msg] <= 3 or el is None:
                kept.append((el, msg))
        self.repeat_counts = seen
        self.flags = kept

    def report(self, out_path, meta):
        c = self.counts
        lines = ["IJARCCE formatter report", "=" * 24, "",
                 "Input : %s" % self.src.name, "Output: %s" % Path(out_path).name,
                 "Title : %s" % (self.title or "(not found)"),
                 "Header: Vol. %s, Issue %s, %s %s   DOI %s   first page %s" % (
                     meta["VOL"], meta["ISSUE"], meta["MONTH"], meta["YEAR"],
                     meta["DOI"] if any(b"{{DOI}}" in v for v in self.tpl.values()) else "(not used by this journal)",
                     meta.get("START_PAGE") or "(not set)"),
                 "",
                 "What was found",
                 "  sections %d, sub-headings %d, figures %d (captions %d), tables %d (captions %d),"
                 % (c.get("h1", 0), c.get("h2", 0) + c.get("h3", 0), c.get("figure", 0),
                    c.get("fig_caption", 0), c.get("table", 0), c.get("table_caption", 0)),
                 "  equations %d, references %d, author lines %d, affiliation lines %d"
                 % (c.get("equation", 0), getattr(self, "ref_count", 0), c.get("author", 0),
                    c.get("affiliation", 0)),
                 "",
                 "Done automatically",
                 "  - Page set to A4 with 0.8 inch margins; journal header and footer added",
                 "  - Title, authors, abstract, headings, captions, tables and references restyled",
                 "  - Headings renumbered (I, II, ... / A, B, ...); references renumbered [1]..[n]",
                 "  - Spacing replaced by single blank lines"]
        lines += ["  - " + f for f in self.fixes]
        lines += ["", "Needs your eye (%d)" % len(self.flags)]
        if not self.flags:
            lines.append("  Nothing was flagged.")
        shown = {}
        for el, msg in self.flags:
            where = ""
            if el is not None:
                snippet = " ".join(ptext(el).split())[:60]
                where = ' [near: "%s"]' % snippet if snippet else ""
            shown[msg] = shown.get(msg, 0) + 1
            total = self.repeat_counts.get(msg, 1)
            more = " (and %d more like this, not marked)" % (total - 3)                 if total > 3 and shown[msg] == 3 else ""
            lines.append("  - %s%s%s" % (msg, where, more))
        lines += ["", "Each flagged place also carries a comment in the Word file.",
                  "Remove them with Review > Delete > Delete All Comments in Document.", ""]
        return "\n".join(lines)


def _flat_with_markers(p):
    """Text of a paragraph with every superscript (real superscript runs, unicode superscripts and
    numbers glued to the end of a word) written as ⟦...⟧."""
    sup_map = dict(zip("\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079\u207b\u207a", "0123456789-+"))
    parts = []
    for r in p.iter(w("r")):
        if any(a.tag in OPAQUE for a in r.iterancestors()):
            continue
        text = "".join((e.text or "") if k == "t" else " " for e, k in atoms(r))
        if not text:
            continue
        va = r.find(w("rPr") + "/" + w("vertAlign"))
        if va is not None and va.get(w("val")) == "superscript" and text.strip():
            parts.append("\u27e6" + text.strip() + "\u27e7")
        else:
            parts.append(text)
    flat = "".join(parts)
    flat = re.sub("[\u2070\u00b9\u00b2\u00b3\u2074-\u2079\u207b\u207a]+",
                  lambda m: "\u27e6" + "".join(sup_map.get(c, c) for c in m.group(0)) + "\u27e7", flat)
    flat = re.sub(r"\u27e7\s*\u27e6", "", flat)
    return " ".join(flat.split())


def _marker_numbers(marker):
    nums = set()
    for part in re.split(r"[,;]", marker or ""):
        part = part.strip().strip("*\u2020\u2021")
        m = re.fullmatch(r"(\d+)\s*[-\u2013]\s*(\d+)", part)
        if m and int(m.group(2)) - int(m.group(1)) < 30:
            nums.update(range(int(m.group(1)), int(m.group(2)) + 1))
        elif part.isdigit():
            nums.add(int(part))
    return nums


def certificate_data(fm):
    """What the certificates need, read from a formatted paper: its title and, for every author,
    the name and the designation (the author's own affiliation).  Works whether the authors were
    numbered by the formatter or by the student."""
    warnings = []
    title = ""
    for b in fm.blocks:
        if b.role == "title":
            title = " ".join(ptext(b.el).split())
            break
    names = []                                           # [name, marker text]
    for b in fm.blocks:
        if b.role != "author":
            continue
        flat = _flat_with_markers(b.el)
        # a number glued to the end of a name ("Rajendra1") is a marker too
        flat = re.sub(r"(?<=[A-Za-z.\]])(\d{1,2}(?:[,\-\u2013]\d{1,2})*)(?=\s*(?:,|&|\band\b|$|\u27e6))",
                      lambda m: "\u27e6" + m.group(1) + "\u27e7", flat)
        for tok in re.findall(r"\u27e6[^\u27e7]*\u27e7|[^\u27e6\u27e7]+", flat):
            if tok.startswith("\u27e6"):
                if names:
                    names[-1][1] = tok[1:-1]
            else:
                for piece in re.split(r",|\band\b|&", tok):
                    piece = piece.strip(" ;*\u2020\u2021")
                    if piece:
                        names.append([piece, ""])
    affs = []                                            # [text, marker numbers]
    marked_from = 0                                      # affs[marked_from:] are the marked ones seen so far
    for b in fm.blocks:
        if b.role != "affiliation":
            continue
        flat = _flat_with_markers(b.el)
        flat = re.sub(r"(?<=[A-Za-z])(\d{1,2}(?:[,\-–]\d{1,2})*)$", lambda m: "⟦" + m.group(1) + "⟧", flat)
        if re.fullmatch(r"\*?\s*corresponding authors?.*", flat.replace("⟦", "").replace("⟧", "").strip(), re.I):
            continue
        if flat.startswith("⟦"):                      # "1Assistant Professor; 2-5UG Students, ..."
            pieces = [(t, m) for m, t in re.findall(r"⟦([^⟧]*)⟧([^⟦]*)", flat)]
        elif re.search(r"⟦[^⟧]*⟧\s*$", flat):   # "..., India1-3"
            pieces = [(t, m) for t, m in re.findall(r"([^⟦]*)⟦([^⟧]*)⟧", flat)]
        else:
            pieces = [(flat, "")]
        for text, marker in pieces:
            text = text.replace("⟦", "").replace("⟧", "").strip(" ,;")
            if not text:
                continue
            nums = _marker_numbers(marker)
            if nums or not affs:
                affs.append([text, nums])
            else:
                # an unmarked line after marked ones is the rest of their address (institute, city ...)
                for entry in affs:
                    if entry[1] and text not in entry[0]:
                        entry[0] = entry[0] + ", " + text

    authors = []
    for i, (name, marker) in enumerate(names):
        nums = _marker_numbers(marker)
        if nums:
            mine = [a for a in affs if a[1] & nums]
        else:
            mine = list(affs) if len(affs) == 1 else []
        if not mine and nums and len(affs) == 1:
            mine = list(affs)
        if not mine:
            warnings.append("No affiliation could be matched to %s: the designation is empty." % name)
        authors.append({"name": name, "designation": "; ".join(a[0] for a in mine)})
    if not authors:
        warnings.append("No author names were found.")
    if not title:
        warnings.append("No paper title was found.")
    return {"title": title, "authors": authors, "warnings": warnings}


def strip_doi(xml):
    """Header/footer XML without its DOI: a paragraph that is only "DOI: {{DOI}}" is removed, and a DOI
    that shares a line ("... 2026  |  DOI {{DOI}}") is cut out of that line."""
    if "{{DOI}}" not in xml:
        return xml
    root = etree.fromstring(xml.encode("utf8"))
    for para in list(root.iter(w("p"))):
        texts = [t for t in para.iter(w("t"))]
        joined = "".join(t.text or "" for t in texts)
        if "{{DOI}}" not in joined:
            continue
        plain = not any(True for _ in para.iter(w("drawing"))) and not any(True for _ in para.iter(w("pict")))
        if plain and re.fullmatch(r"\s*DOI\s*:?\s*\{\{DOI\}\}\s*", joined, flags=re.I) and para.getparent() is not None \
                and len(para.getparent().findall(w("p"))) > 1:
            para.getparent().remove(para)
            continue
        m = re.search(r"\s*\|?\s*(?:DOI\s*:?\s*)?\{\{DOI\}\}", joined, flags=re.I)
        pos = 0
        for t in texts:                                       # delete the matched characters across the runs
            txt = t.text or ""
            a, b = max(m.start() - pos, 0), min(m.end() - pos, len(txt))
            pos += len(txt)
            if a < b:
                t.text = txt[:a] + txt[b:]
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf8")


def make_meta(vol, issue, month, year, paper=None, start_page=None, doi=None, journal="ijarcce", no_doi=False):
    if no_doi:
        doi = ""
    elif not doi:
        doi = auto_doi(journal, year, vol, issue, paper) if paper else \
            "%s%s.%02d%sxx" % (DOI_PREFIXES[journal], year, int(vol), issue)
    return {"VOL": vol, "ISSUE": issue, "MONTH": month, "YEAR": year, "DOI": doi,
            "START_PAGE": start_page}


def format_paper(src, out, meta, template=None, report_path=None, journal="ijarcce"):
    f = Formatter(src, template or JOURNAL_TEMPLATES[journal])
    f.run(out, meta)
    text = f.report(out, meta)
    report_path = Path(report_path) if report_path else Path(out).with_name(Path(out).stem + "_report.txt")
    report_path.write_text(text, encoding="utf-8")
    return f, report_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="Format a paper to the IJARCCE layout")
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--vol", required=True)
    ap.add_argument("--issue", required=True)
    ap.add_argument("--month", required=True)
    ap.add_argument("--year", required=True)
    ap.add_argument("--paper", help="paper number within the issue (used for the DOI)")
    ap.add_argument("--start-page", type=int)
    ap.add_argument("--doi", help="full DOI, if it should not be built from vol/issue/paper")
    ap.add_argument("--journal", choices=sorted(JOURNAL_TEMPLATES), default="ijarcce",
                    help="which journal's header and footer to use")
    ap.add_argument("--template", help="use this template file instead of the journal's own")
    a = ap.parse_args(argv)
    meta = make_meta(a.vol, a.issue, a.month, a.year, a.paper, a.start_page, a.doi, a.journal)
    f, rep = format_paper(a.input, a.output, meta, a.template, journal=a.journal)
    print("wrote %s  (%d item(s) flagged, see %s)" % (a.output, len(f.flags), rep.name))


if __name__ == "__main__":
    sys.exit(main())
