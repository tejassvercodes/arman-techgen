"""Estimates how many pages Word will make of a formatted paper, without Word.

The papers come out of ijarcce_format.py in a very regular layout (one column, Times New Roman, a few
paragraph styles), so the page count can be worked out by laying the document out the way Word does:
real character widths -> lines -> paragraph heights -> pages, with the keep-with-next and widow/orphan
rules.  It reads only the .docx (XML) and a table of Times New Roman character widths, so it runs anywhere
Python runs, including the browser.

    estimate(docx_bytes) -> {"pages": n, "unsure": [...], "paragraphs": [(text, page, y), ...]}
"""
import io
import json
import math
import re
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

TW = 1 / 20.0                    # twips -> points
EMU = 1 / 12700.0                # EMU -> points

_METRICS = None


def metrics():
    global _METRICS
    if _METRICS is None:
        p = Path(__file__).with_name("tnr_metrics.json")
        if not p.exists():
            p = Path(__file__).resolve().parent.parent / "vercel-app" / "py" / "tnr_metrics.json"
        _METRICS = json.loads(p.read_text(encoding="utf-8"))
    return _METRICS


def q(tag, ns=W):
    return "{%s}%s" % (ns, tag)


def val(el, default=None):
    return el.get(q("val")) if el is not None else default


def on(el):
    """A true/false property element: present and not switched off."""
    if el is None:
        return False
    return val(el, "true") not in ("0", "false", "off")


# --------------------------------------------------------------------------------------- tuning
LINE_FACTOR = 1.15               # Times New Roman: one line of single spacing = 1.15 x the font size
SMALLCAPS_SCALE = 0.8
JUSTIFY_SLACK = 0.0159           # a justified line may run over by this fraction of its width: Word squeezes it in (measured in Word)
INLINE_PICTURE_EXTRA = 0.0       # points added below an inline picture (calibrated)
TABLE_BORDER = 0.5               # points per horizontal cell border
TABLE_ROW_EXTRA = 0.0            # calibrated
BOTTOM_EXTRA = 54.3              # the footer pushes the bottom of the text up (measured in Word); see BOTTOM_BY_JOURNAL
# Each journal's footer is a different height, so the lowest line that fits differs.  Measured against Word.
BOTTOM_BY_JOURNAL = {"IJARCCE": 54.3, "IARJSET": 54.3, "IJIREEICE": 51.0, "IMRJR": 37.5}


class Style:
    def __init__(self):
        self.p = {}              # paragraph properties (spacing, ind, jc, keepNext ...)
        self.r = {}              # run properties (sz, b, i, caps, smallCaps)


def _read_ppr(ppr, into):
    if ppr is None:
        return
    for key in ("keepNext", "keepLines", "pageBreakBefore", "widowControl"):
        e = ppr.find(q(key))
        if e is not None:
            into[key] = on(e)
    sp = ppr.find(q("spacing"))
    if sp is not None:
        for a in ("before", "after", "line", "lineRule", "beforeAutospacing", "afterAutospacing"):
            if sp.get(q(a)) is not None:
                into["sp_" + a] = sp.get(q(a))
    ind = ppr.find(q("ind"))
    if ind is not None:
        for a in ("left", "right", "firstLine", "hanging", "start", "end"):
            if ind.get(q(a)) is not None:
                into["ind_" + a] = int(float(ind.get(q(a))))
    jc = ppr.find(q("jc"))
    if jc is not None:
        into["jc"] = val(jc)
    tabs = ppr.find(q("tabs"))
    if tabs is not None:
        into["tabs"] = [(val(t) if False else t.get(q("val")), int(t.get(q("pos")))) for t in tabs.findall(q("tab"))]


def _read_rpr(rpr, into):
    if rpr is None:
        return
    sz = rpr.find(q("sz"))
    if sz is not None:
        into["sz"] = int(sz.get(q("val"))) / 2.0
    for key, tag in (("b", "b"), ("i", "i"), ("caps", "caps"), ("smallCaps", "smallCaps"), ("vanish", "vanish")):
        e = rpr.find(q(tag))
        if e is not None:
            into[key] = on(e)
    va = rpr.find(q("vertAlign"))
    if va is not None:
        into["vertAlign"] = val(va)
    wsc = rpr.find(q("w"))
    if wsc is not None and wsc.get(q("val")):
        into["scale"] = float(wsc.get(q("val"))) / 100.0            # character width scaling, percent
    csp = rpr.find(q("spacing"))
    if csp is not None and csp.get(q("val")):
        into["cspace"] = float(csp.get(q("val"))) * TW              # extra space after every character, points


class Styles:
    def __init__(self, xml):
        root = etree.fromstring(xml)
        self.by_id = {}
        self.default_para = None
        dd = root.find(q("docDefaults"))
        self.dp, self.dr = {}, {}
        if dd is not None:
            _read_ppr(dd.find(q("pPrDefault") + "/" + q("pPr")), self.dp)
            _read_rpr(dd.find(q("rPrDefault") + "/" + q("rPr")), self.dr)
        for s in root.findall(q("style")):
            sid = s.get(q("styleId"))
            node = {"based": val(s.find(q("basedOn"))), "p": {}, "r": {}, "type": s.get(q("type"))}
            _read_ppr(s.find(q("pPr")), node["p"])
            _read_rpr(s.find(q("rPr")), node["r"])
            self.by_id[sid] = node
            if s.get(q("default")) == "1" and s.get(q("type")) == "paragraph":
                self.default_para = sid
        self._cache = {}

    def resolve(self, sid):
        sid = sid or self.default_para
        if sid in self._cache:
            return self._cache[sid]
        chain, cur, seen = [], sid, set()
        while cur and cur in self.by_id and cur not in seen:
            seen.add(cur)
            chain.append(self.by_id[cur])
            cur = self.by_id[cur]["based"]
        st = Style()
        st.p = dict(self.dp)
        st.r = dict(self.dr)
        for node in reversed(chain):
            st.p.update(node["p"])
            st.r.update(node["r"])
        self._cache[sid] = st
        return st


# --------------------------------------------------------------------------------------- measuring
def char_w(ch, bold, italic, size, small_caps_lower=False):
    m = metrics()["bi" if bold and italic else "b" if bold else "i" if italic else "r"]
    c = ch.upper() if small_caps_lower else ch
    w = m.get(c)
    if w is None:
        w = m["_"] if ord(ch) < 0x2E80 else 1000.0           # East Asian characters are full width
    return w * size / 1000.0 * (SMALLCAPS_SCALE if small_caps_lower else 1.0)


def line_height(size):
    return size * LINE_FACTOR


class Atom:
    """One breakable piece of a line: a word (with its trailing space) or a picture."""
    __slots__ = ("w", "space", "h", "size", "pic", "tab")

    def __init__(self, w, space, size, h=0.0, pic=False, tab=False):
        self.w, self.space, self.size, self.h, self.pic, self.tab = w, space, size, h, pic, tab


class Para:
    def __init__(self, el, style, runs_atoms, text):
        self.el, self.style, self.atoms, self.text = el, style, runs_atoms, text
        self.lines = []          # heights of each line, in points
        self.before = self.after = 0.0


class Layout:
    def __init__(self, docx_bytes):
        z = zipfile.ZipFile(io.BytesIO(docx_bytes))
        self.doc = etree.fromstring(z.read("word/document.xml"))
        self.styles = Styles(z.read("word/styles.xml"))
        self.body = self.doc.find(q("body"))
        self.unsure = []
        self.line_info = []        # (text, number of lines, height of one line), top-level paragraphs only (for tests)
        sect = self.body.find(q("sectPr"))
        pg = sect.find(q("pgSz"))
        mar = sect.find(q("pgMar"))
        self.page_w = int(pg.get(q("w"))) * TW
        self.page_h = int(pg.get(q("h"))) * TW
        self.m_top = int(mar.get(q("top"))) * TW
        self.m_bottom = int(mar.get(q("bottom"))) * TW
        self.m_left = int(mar.get(q("left"))) * TW
        self.m_right = int(mar.get(q("right"))) * TW
        cols = sect.find(q("cols"))
        self.cols = int(cols.get(q("num"), "1")) if cols is not None else 1
        if self.cols > 1:
            self.unsure.append("The document has %d columns; the page count is a rough guess." % self.cols)
        self.text_w = self.page_w - self.m_left - self.m_right
        extra = BOTTOM_EXTRA
        jname = self.journal_of(z)
        if jname in BOTTOM_BY_JOURNAL:
            extra = BOTTOM_BY_JOURNAL[jname]
        else:
            self.unsure.append("The header and footer do not match a known journal; the bottom of the page is a guess.")
        self.journal = jname
        # things that float on the page push text around in ways this estimate cannot follow
        n_anchor = len(self.doc.findall(".//" + q("anchor", WP)))
        n_box = len(self.doc.findall(".//" + q("txbxContent")))
        n_pict = len(self.doc.findall(".//" + q("pict"))) + len(self.doc.findall(".//" + q("object")))
        n_frame = len(self.doc.findall(".//" + q("framePr")))
        if n_anchor or n_box or n_pict or n_frame:
            self.unsure.append("The paper has floating pictures, text boxes or frames (%d); text may wrap around them, so the page count may be off." % (n_anchor + n_box + n_pict + n_frame))
        self.page_text_h = self.page_h - self.m_top - self.m_bottom - extra

    @staticmethod
    def journal_of(z):
        """Which journal's header/footer the paper carries (read from the header text)."""
        for name in z.namelist():
            if re.match(r"word/(header|footer)\d*\.xml$|word/ijarcce_(header|footer)\.xml$", name):
                txt = re.sub(r"<[^>]+>", "", z.read(name).decode("utf8", "ignore")).upper()
                for j in ("IJIREEICE", "IARJSET", "IMRJR", "IJARCCE"):
                    if j in txt:
                        return j
        return None

    # ---- a paragraph's atoms (words, pictures) with their widths
    def para_atoms(self, p, style):
        atoms = []
        text_parts = []
        base_r = style.r

        def run_props(r):
            rp = dict(base_r)
            rst = r.find(q("rPr") + "/" + q("rStyle"))
            if rst is not None and val(rst) in self.styles.by_id:
                rp.update(self.styles.by_id[val(rst)]["r"])
            _read_rpr(r.find(q("rPr")), rp)
            return rp

        cur = {"w": 0.0, "n": 0, "size": 0.0}

        def flush(space_w, size):
            if cur["n"]:
                atoms.append(Atom(cur["w"], space_w, cur["size"] or size, tab=False))
            cur["w"], cur["n"], cur["size"] = 0.0, 0, 0.0

        for r in p.iter(q("r")):
            if self._in_textbox(r, p):
                continue
            rp = run_props(r)
            if rp.get("vanish"):
                continue
            size = rp.get("sz", 12.0)
            if rp.get("vertAlign") in ("superscript", "subscript"):
                size = size * 0.65
            bold, ital = rp.get("b", False), rp.get("i", False)
            caps, sc = rp.get("caps", False), rp.get("smallCaps", False)
            scale, cspace = rp.get("scale", 1.0), rp.get("cspace", 0.0)
            for ch_el in r:
                tag = etree.QName(ch_el).localname
                if tag == "t" or tag == "delText":
                    s = ch_el.text or ""
                    text_parts.append(s)
                    for ch in s:
                        if ch in (" ", " ") and tag == "t":
                            if ch == " ":
                                cur["w"] += char_w(ch, bold, ital, size) * scale + cspace
                                cur["n"] += 1
                                cur["size"] = max(cur["size"], size)
                            else:
                                flush(char_w(" ", bold, ital, size) * scale + cspace, size)
                                if not atoms or atoms[-1].space == 0:
                                    pass
                        else:
                            lower = sc and ch.islower() and not caps
                            cur["w"] += char_w(ch.upper() if caps else ch, bold, ital, size, small_caps_lower=lower) * scale + cspace
                            cur["n"] += 1
                            cur["size"] = max(cur["size"], size)
                            if ch in "-–—/" and cur["n"] > 1:
                                flush(0.0, size)             # Word may break after a hyphen
                elif tag == "tab":
                    flush(0.0, size)
                    atoms.append(Atom(0.0, 0.0, size, tab=True))
                    text_parts.append("\t")
                elif tag == "br":
                    flush(0.0, size)
                    atoms.append(Atom(0.0, 0.0, size, h=-1.0))       # a line break
                elif tag == "drawing":
                    flush(0.0, size)
                    inline = ch_el.find(q("inline", WP))
                    if inline is not None:
                        ext = inline.find(q("extent", WP))
                        atoms.append(Atom(int(ext.get("cx")) * EMU, 0.0, size, h=int(ext.get("cy")) * EMU, pic=True))
                    else:
                        anc = ch_el.find(q("anchor", WP))
                        if anc is not None and anc.find(q("wrapTopAndBottom", WP)) is not None:
                            ext = anc.find(q("extent", WP))
                            atoms.append(Atom(int(ext.get("cx")) * EMU, 0.0, size, h=int(ext.get("cy")) * EMU, pic=True))
                            self.unsure.append("A floating picture was counted as if it sat in the text.")
                elif tag == "pict" or tag == "object":
                    flush(0.0, size)
                    self.unsure.append("An older-style picture or object (%s) could not be measured." % tag)
            flush(0.0, size)
        for om in p.iter(q("oMath", M)):
            atoms.append(Atom(0.0, 0.0, 10.0, h=16.0, pic=True))       # a display formula: about 1.4 lines tall
        return atoms, "".join(text_parts)

    def _in_textbox(self, el, p):
        a = el.getparent()
        while a is not None and a is not p:
            if a.tag == q("txbxContent"):
                return True
            a = a.getparent()
        return False

    # ---- lay out one paragraph into lines for a given available width
    def lay_para(self, p, avail_w, in_cell=False):
        sid = val(p.find(q("pPr") + "/" + q("pStyle")))
        st = self.styles.resolve(sid)
        pp = dict(st.p)
        _read_ppr(p.find(q("pPr")), pp)
        atoms, text = self.para_atoms(p, st)
        left = pp.get("ind_left", pp.get("ind_start", 0)) * TW
        right = pp.get("ind_right", pp.get("ind_end", 0)) * TW
        first = pp.get("ind_firstLine", 0) * TW - pp.get("ind_hanging", 0) * TW
        width = avail_w - left - right
        lw0 = width - first if first >= 0 else width - first      # (a hanging indent starts further left)
        lw0 = width - (left + first - left) if False else width - first
        # the paragraph mark's own size sets the height of an empty line
        mark = st.r.get("sz", 12.0)
        rpr_mark = p.find(q("pPr") + "/" + q("rPr"))
        if rpr_mark is not None:
            mrp = {}
            _read_rpr(rpr_mark, mrp)
            mark = mrp.get("sz", mark)
        line_rule = pp.get("sp_lineRule", "auto")
        line_val = float(pp.get("sp_line", 240))

        def height_of(size, pic_h):
            single = line_height(size)
            if line_rule == "exact":
                return line_val * TW
            if line_rule == "atLeast":
                return max(single, line_val * TW, pic_h)
            h = single * line_val / 240.0
            if pic_h:
                h = max(h, pic_h + INLINE_PICTURE_EXTRA) if line_val <= 240 else pic_h + INLINE_PICTURE_EXTRA + (h - single)
            return h

        lines = []
        cur_w, cur_size, cur_pic, started = 0.0, 0.0, 0.0, False
        limit = lw0
        pending_space = 0.0
        cur_sp = 0.0                                         # width of the spaces inside the current line
        slack = JUSTIFY_SLACK * width if pp.get("jc") == "both" else 0.0

        def close(size_default):
            nonlocal cur_w, cur_size, cur_pic, started, limit, pending_space, cur_sp
            lines.append(height_of(cur_size or size_default, cur_pic))
            cur_w, cur_size, cur_pic, started, pending_space, cur_sp = 0.0, 0.0, 0.0, False, 0.0, 0.0
            limit = width

        for a in atoms:
            if a.h == -1.0:                                  # explicit line break
                close(a.size)
                continue
            if a.tab:
                stops = [pos * TW for (kind, pos) in pp.get("tabs", []) if pos * TW > cur_w + left]
                nxt = (min(stops) - left) if stops else (math.floor((cur_w + left) / 36.0) + 1) * 36.0 - left
                cur_w = min(max(nxt, cur_w), limit)
                started = True
                cur_size = max(cur_size, a.size)
                continue
            if a.pic:
                if started and cur_w + pending_space + a.w > limit + 0.01:
                    close(a.size)
                cur_w += pending_space + a.w
                pending_space = 0.0
                cur_pic = max(cur_pic, a.h)
                cur_size = max(cur_size, a.size)
                started = True
                continue
            need = cur_w + pending_space + a.w
            if started and need > limit + slack + 0.01:
                close(a.size)
                need = a.w
            # a word longer than a whole line is broken wherever it falls
            while a.w > limit + 0.01 and not started:
                lines.append(height_of(a.size, 0))
                a = Atom(a.w - limit, a.space, a.size)
                limit = width
                if a.w <= limit:
                    break
            if started:
                cur_sp += pending_space
                cur_w = need
            else:
                cur_w = a.w
            started = True
            cur_size = max(cur_size, a.size)
            pending_space = a.space
        if started or not lines:
            lines.append(height_of(cur_size or mark, cur_pic))
        before = float(pp.get("sp_before", 0)) * TW
        after = float(pp.get("sp_after", 0)) * TW
        para = Para(p, pp, atoms, text)
        para.lines, para.before, para.after = lines, before, after
        para.keep_next = pp.get("keepNext", False)
        para.keep_lines = pp.get("keepLines", False)
        para.page_break_before = pp.get("pageBreakBefore", False)
        para.widow = pp.get("widowControl", True)
        return para

    # ---- tables
    @staticmethod
    def _border_pt(el):
        """Width in points of a border element, 0 when there is none."""
        if el is None or el.get(q("val")) in (None, "none", "nil"):
            return 0.0
        sz = el.get(q("sz"))
        return (int(sz) / 8.0) if sz else 0.5

    def lay_table(self, tbl):
        grid = [int(g.get(q("w"))) * TW for g in tbl.findall(q("tblGrid") + "/" + q("gridCol"))]
        tpr = tbl.find(q("tblPr"))
        dmar = {"left": 5.4, "right": 5.4, "top": 0.0, "bottom": 0.0}
        bottom_tbl = inside_h = 0.0
        if tpr is not None:
            cm = tpr.find(q("tblCellMar"))
            if cm is not None:
                for side in dmar:
                    e = cm.find(q(side))
                    if e is not None and e.get(q("w")) is not None:
                        dmar[side] = int(e.get(q("w"))) * TW
            tb = tpr.find(q("tblBorders"))
            if tb is not None:
                bottom_tbl = self._border_pt(tb.find(q("bottom")))
                inside_h = self._border_pt(tb.find(q("insideH")))
        trs = tbl.findall(q("tr"))
        rows = []
        for ri, tr in enumerate(trs):
            ci, cells_h, texts = 0, [], []
            trh = tr.find(q("trPr") + "/" + q("trHeight"))
            tr_h = int(trh.get(q("val"))) * TW if trh is not None and trh.get(q("val")) else 0.0
            ncells = len(tr.findall(q("tc")))
            for tc in tr.findall(q("tc")):
                tcpr = tc.find(q("tcPr"))
                span, w, vmerge_cont = 1, None, False
                mar = dict(dmar)
                if tcpr is not None:
                    gs = tcpr.find(q("gridSpan"))
                    span = int(val(gs, "1")) if gs is not None else 1
                    tcw = tcpr.find(q("tcW"))
                    if tcw is not None and tcw.get(q("type")) == "dxa":
                        w = int(tcw.get(q("w"))) * TW
                    vm = tcpr.find(q("vMerge"))
                    if vm is not None and val(vm, "continue") == "continue":
                        vmerge_cont = True
                    cm = tcpr.find(q("tcMar"))
                    if cm is not None:
                        for side in mar:
                            e = cm.find(q(side))
                            if e is not None and e.get(q("w")) is not None:
                                mar[side] = int(e.get(q("w"))) * TW
                if w is None:
                    w = sum(grid[ci:ci + span]) if grid else self.text_w / max(1, ncells)
                ci += span
                if vmerge_cont:
                    continue
                inner = w - mar["left"] - mar["right"]
                h = mar["top"] + mar["bottom"]
                for el in tc:
                    if el.tag == q("p"):
                        para = self.lay_para(el, inner, in_cell=True)
                        h += para.before + sum(para.lines) + para.after
                        texts.append(para.text)
                    elif el.tag == q("tbl"):
                        h += sum(r["h"] for r in self.lay_table(el)["rows"])
                        self.unsure.append("A table inside a table was found.")
                cells_h.append(h)
            rh = max(max(cells_h) if cells_h else 0.0, tr_h) + TABLE_ROW_EXTRA
            trp = tr.find(q("trPr"))
            rows.append({"h": rh, "cant_split": trp is not None and trp.find(q("cantSplit")) is not None,
                         "header": trp is not None and trp.find(q("tblHeader")) is not None,
                         "text": next((t for t in texts if t.strip()), "")})
            # Word adds the thickness of the row's bottom border to the row (a cell's own border wins over
            # the table's; the next row's top border counts when it is heavier)
            bots, tops_next = [], []
            for tc in tr.findall(q("tc")):
                tcb = tc.find(q("tcPr") + "/" + q("tcBorders"))
                bot = tcb.find(q("bottom")) if tcb is not None else None
                bots.append(self._border_pt(bot) if bot is not None else (bottom_tbl if ri == len(trs) - 1 else inside_h))
            if ri + 1 < len(trs):
                for tc in trs[ri + 1].findall(q("tc")):
                    tcb = tc.find(q("tcPr") + "/" + q("tcBorders"))
                    top = tcb.find(q("top")) if tcb is not None else None
                    if top is not None:
                        tops_next.append(self._border_pt(top))
            rows[-1]["h"] += max(bots + tops_next + [0.0])
        return {"rows": rows}

    # ---- the whole document
    def run(self):
        blocks = []
        for el in self.body:
            if el.tag == q("p"):
                para = self.lay_para(el, self.text_w)
                self.line_info.append((para.text, len(para.lines), para.lines[0] if para.lines else 0.0))
                blocks.append(("p", para))
            elif el.tag == q("tbl"):
                blocks.append(("t", self.lay_table(el)))
        return self.paginate(blocks)

    def paginate(self, blocks):
        H = self.page_text_h
        page, y = 1, 0.0
        marks = []                 # (text, page, y of its first line)
        i = 0
        EPS = 0.05

        def first_need(b):
            kind, obj = b
            if kind == "t":
                return obj["rows"][0]["h"] if obj["rows"] else 0.0
            n = len(obj.lines)
            take = 1 if n < 2 or not obj.widow else min(2, n)
            if obj.keep_lines or n <= 3 and obj.widow:
                take = n
            return obj.before + sum(obj.lines[:take])

        while i < len(blocks):
            kind, obj = blocks[i]
            # a chain of keep-with-next paragraphs must sit on one page together with the start of what follows
            if kind == "p" and obj.keep_next:
                j, need = i, 0.0
                while j < len(blocks) and blocks[j][0] == "p" and blocks[j][1].keep_next:
                    pj = blocks[j][1]
                    need += pj.before + sum(pj.lines) + pj.after
                    j += 1
                if j < len(blocks):
                    need += first_need(blocks[j])
                if y > 0 and y + need > H + EPS:
                    page, y = page + 1, 0.0
            if kind == "p":
                if obj.page_break_before and y > 0:
                    page, y = page + 1, 0.0
                before = obj.before if y > 0 else 0.0
                lines = obj.lines
                n = len(lines)
                total = before + sum(lines)
                if y + total <= H + EPS:
                    marks.append((obj.text, page, y + before))
                    y += total + obj.after
                else:
                    # split across pages, honouring keep-lines and widow/orphan control
                    placed, k = y + before, 0
                    first_page_mark = None
                    fits = 0
                    acc = placed
                    for h in lines:
                        if acc + h <= H + EPS:
                            acc += h
                            fits += 1
                        else:
                            break
                    if obj.keep_lines or fits == 0:
                        fits_ok = 0
                    else:
                        fits_ok = fits
                        if obj.widow and n >= 4:
                            if fits_ok < 2:
                                fits_ok = 0
                            elif n - fits_ok < 2:
                                fits_ok = n - 2
                        elif obj.widow and n < 4:
                            fits_ok = 0
                    if fits_ok == 0:
                        page, y = page + 1, 0.0
                        marks.append((obj.text, page, y))
                        rest = list(lines)
                    else:
                        marks.append((obj.text, page, y + before))
                        rest = lines[fits_ok:]
                        page, y = page + 1, 0.0
                        # spill the remainder (it may itself need several pages)
                    while True:
                        acc = 0.0
                        fit = 0
                        for h in rest:
                            if acc + h <= H + EPS:
                                acc += h
                                fit += 1
                            else:
                                break
                        if fit == len(rest):
                            y = acc + obj.after
                            break
                        if fit == 0:
                            fit = 1
                            acc = rest[0]
                        rest = rest[fit:]
                        page, y = page + 1, 0.0
                i += 1
            else:
                rows = obj["rows"]
                headers = []
                for r in rows:
                    if r["header"]:
                        headers.append(r)
                    else:
                        break
                first = True
                for ri, r in enumerate(rows):
                    if y + r["h"] > H + EPS and y > 0:
                        page, y = page + 1, 0.0
                        if not r["header"]:
                            y += sum(h["h"] for h in headers if h is not r)      # header rows repeat on the new page
                    if first:
                        marks.append((r["text"], page, y))
                        first = False
                    y += r["h"]
                i += 1

        return page, marks


def estimate(docx_bytes):
    lay = Layout(docx_bytes)
    pages, marks = lay.run()
    marks = [(t, pg, y + lay.m_top) for (t, pg, y) in marks]      # y from the top edge of the page, as Word reports it
    return {"pages": pages, "unsure": sorted(set(lay.unsure)), "paragraphs": marks}
