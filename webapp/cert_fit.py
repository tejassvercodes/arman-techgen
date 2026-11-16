"""How small the certificate text must be so that the Volume and DOI lines stay inside the box.

This is the certificate generator's own PDF layout (certificate-generator/pdf-export.js: word wrap, line
heights, the box size from templates/pdf/<KEY>.json) done in Python with the real Windows fonts, so the
paper formatter can tell the certificate maker exactly which fit level to use.  The levels themselves are
read from certificate-generator/app.js (FIT_LEVELS), so the two never drift apart.
"""
import json
import math
import re
from functools import lru_cache
from pathlib import Path

import pymupdf

CERT_DIR = Path(__file__).resolve().parent.parent / "certificate-generator"
FONT_DIR = Path(r"C:\Windows\Fonts")
FIT_MARGIN_PT = 2            # same safety margin as pdf-export.js

# family -> (regular, bold, italic, bold italic) font files
FONT_FILES = {
    "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "algerian": ("ALGER.TTF",) * 4,
    "times new roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "calisto mt": ("CALIST.TTF", "CALISTB.TTF", "CALISTI.TTF", "CALISTBI.TTF"),
    "nirmala ui": ("Nirmala.ttc",) * 4,
    "microsoft phagspa": ("phagspa.ttf", "phagspab.ttf", "phagspa.ttf", "phagspab.ttf"),
}


@lru_cache(maxsize=None)
def fit_levels():
    text = (CERT_DIR / "app.js").read_text(encoding="utf-8")
    block = re.search(r"const FIT_LEVELS = \[(.*?)\];", text, re.S).group(1)
    return [(float(a), float(b)) for a, b in re.findall(r"line:\s*([\d.]+),\s*font:\s*([\d.]+)", block)]


@lru_cache(maxsize=None)
def spec_for(key):
    return json.loads((CERT_DIR / "templates" / "pdf" / (key + ".json")).read_text(encoding="utf-8"))


def family_of(name):
    return re.sub(r"\s+(Bold|Italic|Regular)$", "", name, flags=re.I)


@lru_cache(maxsize=None)
def _font(family, bold, italic):
    """(font, width factor).  When the bold face can't be loaded on its own (Nirmala.ttc keeps all its
    faces in one file, Algerian has no bold), the regular face is used and widened by 6 % so the text is
    never measured narrower than Word sets it."""
    files = FONT_FILES.get(family.lower())
    if not files:
        return pymupdf.Font("helv"), 1.06
    want = files[(1 if bold else 0) + (2 if italic else 0)]
    path = FONT_DIR / want
    widen = 1.0
    if not path.exists() or (bold and want == files[0]):
        path, widen = FONT_DIR / files[0], (1.06 if bold else 1.0)
    try:
        return pymupdf.Font(fontfile=str(path)), widen
    except Exception:
        return pymupdf.Font("helv"), 1.06


def font_for(family, bold, italic):
    return _font(family, bold, italic)[0]


def _width(text, run, size):
    if not text:
        return 0.0
    if run.get("caps"):
        text = text.upper()
    f, widen = _font(family_of(run["font"]), bool(run.get("bold")), bool(run.get("italic")))
    return f.text_length(text, fontsize=size) * widen + (run.get("spacing") or 0) * len(text)


def _measure(chars, sizes):
    total, i = 0.0, 0
    while i < len(chars):
        j = i
        while j < len(chars) and chars[j][1] is chars[i][1]:
            j += 1
        run = chars[i][1]
        total += _width("".join(c[0] for c in chars[i:j]), run, sizes[id(run)])
        i = j
    return total


def _fill(paragraph, values):
    chars = [(ch, run) for run in paragraph["runs"] for ch in run["text"]]
    text = "".join(c[0] for c in chars)
    m = re.search(r"\{#hasDoi\}([\s\S]*?)\{/hasDoi\}", text)
    if m:
        if not values["doi"]:
            return None
        a, b = m.start(), m.end()
        chars = chars[:a] + chars[a + len("{#hasDoi}"):b - len("{/hasDoi}")] + chars[b:]
        text = "".join(c[0] for c in chars)
    out, last = [], 0
    for t in re.finditer(r"\{([^{}]+)\}", text):
        out += chars[last:t.start()]
        run = chars[t.start()][1]
        out += [(ch, run) for ch in str(values.get(re.sub(r"[\s_]+", "", t.group(1).lower()), ""))]
        last = t.end()
    return out + chars[last:]


def _wrap(chars, available, sizes):
    lines, line, word = [], [], []

    def flush():
        nonlocal line, word
        if not word:
            return
        if line and _measure(line + word, sizes) > available:
            while line and line[-1][0] == " ":
                line.pop()
            lines.append(line)
            line = []
        if not line and _measure(word, sizes) > available:
            piece = []
            for c in word:
                if piece and _measure(piece + [c], sizes) > available:
                    lines.append(piece)
                    piece = []
                piece.append(c)
            line = piece
        else:
            line = line + word
        word = []

    for i, c in enumerate(chars):
        if c[0] == "\n":
            flush()
            lines.append(line)
            line = []
            continue
        if c[0] == " ":
            flush()
            line.append(c)
            continue
        word.append(c)
        nxt = chars[i + 1] if i + 1 < len(chars) else None
        if re.match(r"[-\u2010\u2013\u2014]", c[0]) and nxt and nxt[0] not in " \n" and len(word) > 1:
            flush()
    flush()
    while line and line[-1][0] == " ":
        line.pop()
    lines.append(line)
    return lines


def fits(key, record, level):
    spec = spec_for(key)
    line_f, font_f = fit_levels()[level]
    box = spec["box"]
    values = {"name": (record.get("NAME") or "").upper(), "designation": record.get("Designation") or "",
              "papertitle": record.get("PaperTitle") or "", "doi": record.get("DOI") or "",
              "vol": record.get("Volume") or "", "volume": record.get("Volume") or "",
              "issue": record.get("Issue") or "", "year": record.get("Year") or "", "month": record.get("Month") or ""}
    content_w = box["width"] - box["insetLeft"] - box["insetRight"]
    y = box["insetTop"]
    limit = box["height"] - FIT_MARGIN_PT
    for p in spec["paragraphs"]:
        sizes = {id(r): (r["size"] if font_f == 1 else max(1, math.floor(r["size"] * font_f * 2 + 0.5) / 2)) for r in p["runs"]}
        chars = _fill(p, values)
        if chars is None:
            continue
        avail = content_w - (p["leftIndent"] + p["rightIndent"]) - max(0, p["firstLineIndent"])
        y += p["spaceBefore"]
        rule, value = p["line"]["rule"], p["line"]["value"]
        if rule == "multiple":
            value = math.floor(value * 240 * line_f + 0.5) / 240     # rounded like Math.round in the JS
        for line in _wrap(chars, avail, sizes):
            runs = list({id(c[1]): c[1] for c in line}.values()) if line else [p["runs"][0]]
            asc = desc = 0.0
            for r in runs:
                m = spec.get("fontMetrics", {}).get("%s|%d|%d" % (family_of(r["font"]), 1 if r.get("bold") else 0,
                                                                   1 if r.get("italic") else 0))
                sz = sizes[id(r)]
                if m:
                    a, d = m["ascent"] * sz, m["descent"] * sz
                else:
                    f = font_for(family_of(r["font"]), bool(r.get("bold")), bool(r.get("italic")))
                    a, d = f.ascender * sz, -f.descender * sz
                asc, desc = max(asc, a), max(desc, d)
            natural = asc + desc
            height = natural * value if rule == "multiple" else (max(natural, value) if rule == "atLeast" else value)
            if line and y + natural > limit:
                return False
            y += height
        y += p["spaceAfter"]
    return True


def fit_level(key, record):
    """(level, fits) — the first level at which every line stays inside the box."""
    levels = fit_levels()
    for i in range(len(levels)):
        if fits(key, record, i):
            return i, True
    return len(levels) - 1, False
