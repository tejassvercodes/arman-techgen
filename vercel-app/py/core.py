"""Logic shared by the desktop server (webapp/app.py) and the in-browser edition (vercel-app):
issue settings, DOI reading, the page/DOI chain, and the description of a formatted paper.
Pure Python, no Flask, no Word: it runs in Pyodide too."""
import re
from datetime import datetime

from ijarcce_format import JOURNAL_NAMES, DOI_PREFIXES, auto_doi as journal_auto_doi, ptext

JOURNALS = dict(JOURNAL_NAMES)
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


def settings_ok(s):
    return all(str(s.get(k, "")).strip() for k in ("volume", "issue", "month", "year", "first_page", "first_paper"))


# How the issue details work, the same idea as the certificate generator's "With DOI / Without DOI":
#   With DOI     the DOI of the issue's FIRST paper gives volume, issue, year and the first paper number.
#   Without DOI  volume, issue, year and the first paper number are typed.
# The issue number IS the month (issue 8 = August), so the month is never typed.  The first page is
# always typed.  The DOI also names the journal (10.17148/IJARCCE...), so a DOI from another journal is
# refused, or switches the journal if the choice is left to it.
FIRST_JOURNAL_YEAR = 2011
DOI_LAYOUTS = [(2, 2, 2), (2, 1, 2), (2, 1, 3), (2, 2, 3), (1, 2, 2), (1, 1, 2)]   # volume, issue, file digits


def clean_doi(raw):
    raw = (raw or "").strip()
    for _ in range(2):                                   # "DOI: https://doi.org/10...." in either order
        raw = re.sub(r"^doi\s*:?\s*", "", raw, flags=re.I)
        raw = re.sub(r"^(?:https?://)?(?:dx\.)?doi\.org/", "", raw, flags=re.I)
    return re.sub(r"\s+", "", raw)


def parse_doi(doi):
    """{'journal','volume','issue','year','month','file_number','error','note'} read from a DOI such as
    10.17148/IJARCCE.2026.15817 (last number = volume 2 digits + issue + file number 2 digits)."""
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
        out["note"] = "Check the DOI: \"%s\" was read as issue %d (%s), file %s, but could also be issue %d (%s), file %s." % (
            last, issue, MONTHS[issue - 1], file, alt[1], MONTHS[alt[1] - 1], alt[2])
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
        s["first_doi"] = ""                                  # typed details only: no DOI left over from before
    if s["mode"] == "with":
        if s["first_doi"]:
            d = parse_doi(s["first_doi"])
            if d["error"]:
                raise ValueError(d["error"])
            if d["journal"]:
                s["journal"] = d["journal"]                  # the DOI names the journal
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
    s["month"] = MONTHS[int(s["issue"]) - 1] if s["issue"] else ""     # the issue number sets the month
    return s


def auto_doi(s, paper_no):
    return journal_auto_doi(s["journal"], s["year"], s["volume"], s["issue"], paper_no)


def chain(settings, papers):
    """The paper number, DOI and first page of each paper, in order.  A paper is a dict with 'id',
    optional 'doi_manual' / 'start_manual', and, once its length is known, 'pages'.
    'Without DOI' means no DOI at all."""
    running = int(settings["first_page"])
    first = int(settings["first_paper"])
    out = {}
    for i, p in enumerate(papers):
        no = first + i
        doi = "" if settings.get("mode") == "without" else (p.get("doi_manual") or auto_doi(settings, no))
        start = p.get("start_manual") or running
        out[p["id"]] = {"paper_no": no, "doi": doi, "start": start}
        if p.get("pages"):
            running = start + p["pages"]
    return out


def describe(fm):
    """What the page shows about a formatted paper: flags, counts, fixes (from a finished Formatter)."""
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
    return {"flags": flags, "found": found, "fixes": list(fm.fixes), "title": fm.title}
