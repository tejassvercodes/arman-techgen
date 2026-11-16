"""One-off builder for templates/CCE_production.docx.

Reads   templates/CCE.docx          (the journal's own template, never modified)
        dad_finals/2.docx           (source of the header/footer as currently published)
Writes  templates/CCE_original_backup.docx   (byte-for-byte copy of CCE.docx)
        templates/CCE_production.docx        (what the formatter uses)

Changes made in the production copy (and why) are listed in README.md.
Re-run only if the journal's header/footer or style rules change.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "templates" / "CCE.docx"
BACKUP = ROOT / "templates" / "CCE_original_backup.docx"
OUT = ROOT / "templates" / "CCE_production.docx"
HEADER_SOURCE = ROOT / "templates" / "sources" / "IJARCCE_example.docx"
IARJSET_SOURCE = ROOT / "templates" / "sources" / "IARJSET_example.docx"
IARJSET_OUT = ROOT / "templates" / "IARJSET_published_production.docx"   # the old header (8.311, no DOI line), kept for reference

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def w(tag):
    return "{%s}%s" % (W, tag)


# Text width between the 0.8" margins on A4, in twips.
TEXT_WIDTH = 11909 - 2 * 1152

# style id -> (name, basedOn, paragraph props, run props)
# These values are what the published papers actually use.
STYLES = {
    "IEEETitle": ("IEEE Title", "Normal",
                  {"jc": "center", "spacing": {"before": "0", "after": "0"}},
                  {"sz": "48"}),
    "IEEEAuthorName": ("IEEE Author Name", "Normal",
                       {"jc": "center", "spacing": {"before": "0", "after": "0", "line": "360", "lineRule": "auto"}},
                       {"b": True, "sz": "24"}),
    "IEEEAuthorAffiliation": ("IEEE Author Affiliation", "Normal",
                              {"jc": "center", "spacing": {"before": "0", "after": "0", "line": "360", "lineRule": "auto"}},
                              {"i": True, "sz": "20"}),
    "IEEEAbtract": ("IEEE Abtract", "Normal",
                    {"jc": "both", "spacing": {"before": "0", "after": "0"}},
                    {"sz": "20"}),
    "IEEEParagraph": ("IEEE Paragraph", "Normal",
                      {"jc": "both", "spacing": {"before": "0", "after": "0"}, "ind": {"firstLine": "0"}},
                      {"sz": "20"}),
    "IEEEHeading1": ("IEEE Heading 1", "Normal",
                     {"jc": "center", "spacing": {"before": "0", "after": "0"}, "keepNext": True},
                     {"b": True, "smallCaps": True, "sz": "20"}),
    "IEEEHeading2": ("IEEE Heading 2", "Normal",
                     {"jc": "left", "spacing": {"before": "0", "after": "0"}, "keepNext": True},
                     {"i": True, "sz": "20"}),
    "IEEEHeading3": ("IEEE Heading 3", "Normal",
                     {"jc": "left", "spacing": {"before": "0", "after": "0"}, "ind": {"left": "216"}, "keepNext": True},
                     {"i": True, "sz": "20"}),
    "IEEETableCaption": ("IEEE Table Caption", "Normal",
                         {"jc": "center", "spacing": {"before": "0", "after": "0"}, "keepNext": True},
                         {"smallCaps": True, "sz": "20"}),
    "IEEEFigure": ("IEEE Figure", "Normal",
                   {"jc": "center", "spacing": {"before": "0", "after": "0"}, "keepNext": True},
                   {"sz": "20"}),
    "IEEEFigureCaptionMulti-Lines": ("IEEE Figure Caption Multi-Lines", "Normal",
                                     {"jc": "center", "spacing": {"before": "0", "after": "0"}},
                                     {"sz": "20"}),
    "IEEETableCell": ("IEEE Table Cell", "Normal",
                      {"jc": "left", "spacing": {"before": "0", "after": "0"}},
                      {"sz": "20"}),
    "IEEETableHeaderCentered": ("IEEE Table Header Centered", "IEEETableCell",
                                {"jc": "center"}, {"b": True}),
    "IEEETableHeaderLeft-Justified": ("IEEE Table Header Left-Justified", "IEEETableCell",
                                      {}, {"b": True}),
    "IEEEReferenceItem": ("IEEE Reference Item", "Normal",
                          {"jc": "both", "spacing": {"before": "0", "after": "0"},
                           "ind": {"left": "540", "hanging": "540"}},
                          {"sz": "20"}),
    "IEEEListItem": ("IEEE List Item", "Normal",
                     {"jc": "both", "spacing": {"before": "0", "after": "0"},
                      "ind": {"left": "540", "hanging": "360"}},
                     {"sz": "20"}),
    "IEEEEquation": ("IEEE Equation", "Normal",
                     {"jc": "left", "spacing": {"before": "0", "after": "0"},
                      "tabs": [("center", str(TEXT_WIDTH // 2)), ("right", str(TEXT_WIDTH))]},
                     {"sz": "20"}),
    "IEEECode": ("IEEE Code", "Normal",
                 {"jc": "left", "spacing": {"before": "0", "after": "0"}, "ind": {"left": "216"}},
                 {"font": "Courier New", "sz": "18"}),
}


def make_style(style_id, spec):
    name, based, ppr_spec, rpr_spec = spec
    st = etree.Element(w("style"), nsmap={"w": W})
    st.set(w("type"), "paragraph")
    st.set(w("customStyle"), "1")
    st.set(w("styleId"), style_id)
    etree.SubElement(st, w("name")).set(w("val"), name)
    etree.SubElement(st, w("basedOn")).set(w("val"), based)
    etree.SubElement(st, w("qFormat"))
    ppr = etree.SubElement(st, w("pPr"))
    if ppr_spec.get("keepNext"):
        etree.SubElement(ppr, w("keepNext"))
    if "tabs" in ppr_spec:
        tabs = etree.SubElement(ppr, w("tabs"))
        for kind, pos in ppr_spec["tabs"]:
            t = etree.SubElement(tabs, w("tab"))
            t.set(w("val"), kind)
            t.set(w("pos"), pos)
    if "spacing" in ppr_spec:
        sp = etree.SubElement(ppr, w("spacing"))
        for k, v in ppr_spec["spacing"].items():
            sp.set(w(k), v)
    if "ind" in ppr_spec:
        ind = etree.SubElement(ppr, w("ind"))
        for k, v in ppr_spec["ind"].items():
            ind.set(w(k), v)
    if "jc" in ppr_spec:
        etree.SubElement(ppr, w("jc")).set(w("val"), ppr_spec["jc"])
    rpr = etree.SubElement(st, w("rPr"))
    if "font" in rpr_spec:
        f = etree.SubElement(rpr, w("rFonts"))
        for a in ("ascii", "hAnsi", "cs"):
            f.set(w(a), rpr_spec["font"])
    for flag in ("b", "i", "smallCaps"):
        if rpr_spec.get(flag):
            etree.SubElement(rpr, w(flag))
    if "sz" in rpr_spec:
        etree.SubElement(rpr, w("sz")).set(w("val"), rpr_spec["sz"])
        etree.SubElement(rpr, w("szCs")).set(w("val"), rpr_spec["sz"])
    return st


def build_styles(xml):
    root = etree.fromstring(xml)
    existing = {s.get(w("styleId")): s for s in root.findall(w("style"))}
    for sid, spec in STYLES.items():
        new = make_style(sid, spec)
        if sid in existing:
            root.replace(existing[sid], new)
        else:
            root.append(new)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_header(xml, with_doi=True):
    s = xml.decode("utf8")
    s, n1 = re.subn(r"Vol\. \d+, Issue \d+, [A-Za-z]+ \d{4}",
                    "Vol. {{VOL}}, Issue {{ISSUE}}, {{MONTH}} {{YEAR}}", s)
    # The DOI number is split over two runs in the source; put a token in the
    # first and empty the second.
    s, n2 = re.subn(r"(<w:t>)10\.17148/IJARCCE\.[\d.]+(</w:t></w:r><w:r>.*?<w:t>)\d+(</w:t>)",
                    r"\g<1>{{DOI}}\g<2>\g<3>", s, flags=re.S)
    if n1 != 1 or (with_doi and n2 != 1):
        sys.exit("Header layout in %s is not what this script expects (%d, %d)" % ("header source", n1, n2))
    return s.encode("utf8")


VOL_RULE = (r"Vol\.\s*\d+,\s*Issue\s*\w+,\s*\w+\s*\d{4}", "Vol. {{VOL}}, Issue {{ISSUE}}, {{MONTH}} {{YEAR}}")
VOLUME_RULE = (r"Volume\s*\d+,\s*Issue\s*\w+,\s*\w+\s*\d{4}", "Volume {{VOL}}, Issue {{ISSUE}}, {{MONTH}} {{YEAR}}")
DOI_RULE = (r"10\.17148/\w+\.\d{4}\.\w+", "{{DOI}}")

# journals whose header comes from a journal-supplied "...-Paper-Format.docx" (text split over many runs)
RUN_AWARE = {
    "ijireeice": {"source": "IJIREEICE-Paper-Format.docx", "out": "IJIREEICE_production.docx",
                  "rules": [VOL_RULE, DOI_RULE]},
    "imrjr": {"source": "IMRJR-Paper-Format.docx", "out": "IMRJR_production.docx",
              "rules": [VOLUME_RULE, DOI_RULE]},
    # IARJSET now uses the journal's current template (impact factor 8.066, with a DOI line)
    "iarjset": {"source": "IARJSET-Paper-Format.docx", "out": "IARJSET_production.docx",
                    "rules": [VOL_RULE, DOI_RULE]},
}


def replace_span(p, pattern, repl):
    """Replace the first match of `pattern` in the paragraph's text, even when the text is split
    over several runs: the replacement goes in the first run the match touches, the other runs
    lose their part of the match."""
    ts = list(p.iter(w("t")))
    text = "".join(t.text or "" for t in ts)
    m = re.search(pattern, text)
    if not m:
        return False
    pos, first = 0, None
    for t in ts:
        txt = t.text or ""
        a, b = pos, pos + len(txt)
        pos = b
        if b <= m.start() or a >= m.end():
            continue
        head = txt[:max(0, m.start() - a)]
        tail = txt[m.end() - a:] if b > m.end() else ""
        if first is None:
            first = t
            t.text = head + repl + tail
        else:
            t.text = tail
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return True


def tokenize_header(xml, rules):
    root = etree.fromstring(xml)
    for pattern, repl in rules:
        hits = 0
        for p in root.iter(w("p")):
            in_fallback = any(a.tag.endswith("}Fallback") for a in p.iterancestors())
            if replace_span(p, pattern, repl) and not in_fallback:
                hits += 1
        if hits != 1:
            sys.exit("Header pattern %r matched %d paragraph(s), expected 1" % (pattern, hits))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build(header_source, out_path, with_doi, rules=None, key="ijarcce"):
    if not BACKUP.exists():
        shutil.copyfile(SRC, BACKUP)
    src = zipfile.ZipFile(SRC)
    hs = zipfile.ZipFile(header_source)

    hs_rels = hs.read("word/_rels/document.xml.rels").decode("utf8")
    hs_doc = hs.read("word/document.xml").decode("utf8")
    first_sect = re.search(r"<w:sectPr.*?</w:sectPr>", hs_doc, re.S).group(0)
    hid = re.search(r'<w:headerReference w:type="default" r:id="(\w+)"', first_sect).group(1)
    fid = re.search(r'<w:footerReference w:type="default" r:id="(\w+)"', first_sect).group(1)

    def target(rid):
        m = re.search(r'<Relationship [^>]*Id="%s"[^>]*>' % rid, hs_rels).group(0)
        return re.search(r'Target="([^"]+)"', m).group(1)

    h_part, f_part = target(hid), target(fid)
    if rules:
        header = tokenize_header(hs.read("word/" + h_part), rules)
    else:
        header = build_header(hs.read("word/" + h_part), with_doi)
    footer = hs.read("word/" + f_part)
    h_rels = hs.read("word/_rels/%s.rels" % h_part).decode("utf8")
    f_rels = hs.read("word/_rels/%s.rels" % f_part).decode("utf8")
    media = {}                       # new package name -> bytes (every picture the header or footer uses)
    if rules:
        for which, rels_text in (("h", h_rels), ("f", f_rels)):
            for tgt in sorted(set(re.findall(r'Target="(media/[^"]+)"', rels_text))):
                new = "media/%s_%s" % (key, tgt.split("/")[-1])
                media["word/" + new] = hs.read("word/" + tgt)
                if which == "h":
                    h_rels = h_rels.replace(tgt, new)
                else:
                    f_rels = f_rels.replace(tgt, new)
    else:
        logo_name = re.search(r'Target="(media/[^"]+)"', h_rels).group(1)
        media["word/media/ijarcce_logo.png"] = hs.read("word/" + logo_name)
        h_rels = h_rels.replace(logo_name, "media/ijarcce_logo.png")

    drop = re.compile(r"word/(_rels/)?(header|footer)\d+\.xml(\.rels)?$")
    doc_rels = src.read("word/_rels/document.xml.rels").decode("utf8")
    doc_rels = re.sub(r'<Relationship [^>]*Target="(header|footer)\d+\.xml"[^>]*/>', "", doc_rels)
    doc_rels = doc_rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdJH" Type="%s/header" Target="header1.xml"/>'
        '<Relationship Id="rIdJF" Type="%s/footer" Target="footer1.xml"/></Relationships>' % (R, R))

    doc = src.read("word/document.xml").decode("utf8")
    doc = re.sub(r"<w:(header|footer)Reference [^>]*/>", "", doc)
    doc = re.sub(r'(<w:pgMar [^>]*?)w:header="\d+" w:footer="\d+"', r'\1w:header="720" w:footer="720"', doc)
    refs = ('<w:headerReference w:type="default" r:id="rIdJH"/>'
            '<w:footerReference w:type="default" r:id="rIdJF"/>')
    doc = re.sub(r"(<w:sectPr[^>]*>)", r"\1" + refs, doc, count=1)

    ct = src.read("[Content_Types].xml").decode("utf8")
    ct = re.sub(r'<Override PartName="/word/(header|footer)\d+\.xml"[^>]*/>', "", ct)
    ct = ct.replace(
        "</Types>",
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            name = item.filename
            if drop.match(name):
                continue
            data = src.read(name)
            if name == "word/styles.xml":
                data = build_styles(data)
            elif name == "word/document.xml":
                data = doc.encode("utf8")
            elif name == "word/_rels/document.xml.rels":
                data = doc_rels.encode("utf8")
            elif name == "[Content_Types].xml":
                data = ct.encode("utf8")
            out.writestr(name, data)
        out.writestr("word/header1.xml", header)
        out.writestr("word/footer1.xml", footer)
        out.writestr("word/_rels/header1.xml.rels", h_rels)
        out.writestr("word/_rels/footer1.xml.rels", f_rels)
        for name, data in media.items():
            out.writestr(name, data)
    print("wrote", out_path)


def main():
    build(HEADER_SOURCE, OUT, True)
    build(IARJSET_SOURCE, IARJSET_OUT, False)
    for key, cfg in RUN_AWARE.items():
        src = ROOT / "templates" / cfg["source"]
        if src.exists():
            build(src, ROOT / "templates" / cfg["out"], True, rules=cfg["rules"], key=key)


if __name__ == "__main__":
    main()
