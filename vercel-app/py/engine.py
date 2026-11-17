"""Entry points the browser calls (through Pyodide).  Everything here takes and returns JSON text."""
import json
import sys

sys.path.insert(0, "/app/formatter")

import ijarcce_format as F          # noqa: E402
import core                         # noqa: E402
import pagecount                    # noqa: E402

F.PDF_PROVIDER = lambda path: None   # no PDF in the browser: tables/equations a PDF conversion broke are flagged


def meta():
    return json.dumps({"months": core.MONTHS, "journals": core.JOURNALS, "prefixes": F.DOI_PREFIXES})


def parse_doi(doi):
    return json.dumps(core.parse_doi(doi))


def clean_settings(src_json):
    try:
        return json.dumps({"ok": core.clean_settings(json.loads(src_json))})
    except ValueError as exc:
        return json.dumps({"error": str(exc)})


def chain(settings_json, papers_json):
    return json.dumps(core.chain(json.loads(settings_json), json.loads(papers_json)))


def format_paper(job_json):
    """Format /work/in.docx into /work/out.docx; returns what the page shows about it."""
    job = json.loads(job_json)
    doi = job.get("doi") or ""
    m = F.make_meta(job["volume"], job["issue"], job["month"], job["year"], doi=doi or None,
                    start_page=job["start"], journal=job["journal"], no_doi=not doi)
    fm, _ = F.format_paper("/work/in.docx", "/work/out.docx", m, journal=job["journal"],
                           report_path="/work/report.txt")
    info = core.describe(fm)
    info["cert_data"] = F.certificate_data(fm)
    est = pagecount.estimate(open("/work/out.docx", "rb").read())
    info["pages"] = est["pages"]
    info["unsure"] = est["unsure"]
    info["report"] = open("/work/report.txt", encoding="utf-8").read()
    return json.dumps(info)


def count_pages():
    """Estimated pages of /work/out.docx (after the paper changed, e.g. a header edit)."""
    est = pagecount.estimate(open("/work/out.docx", "rb").read())
    return json.dumps({"pages": est["pages"], "unsure": est["unsure"]})
