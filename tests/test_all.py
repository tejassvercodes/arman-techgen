"""Checks for the whole product.  Run:  python -m unittest discover -s tests -v

Needs Python with lxml, flask and pymupdf.  The formatter checks use the sample papers in
student_originals/ and dad_finals/.  Nothing here needs Microsoft Word.
"""
import difflib
import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.environ["PAPERDESK_DATA"] = tempfile.mkdtemp(prefix="paperdesk_test_")
sys.path.insert(0, str(ROOT / "webapp"))
sys.path.insert(0, str(ROOT / "formatter"))

import app as web            # noqa: E402
import cert_fit              # noqa: E402
import compare               # noqa: E402
import ijarcce_format as F   # noqa: E402


class DoiRules(unittest.TestCase):
    def test_reads_volume_issue_file(self):
        d = web.parse_doi("10.17148/IJARCCE.2026.15817")
        self.assertEqual((d["journal"], d["volume"], d["issue"], d["month"], d["year"], d["file_number"]),
                         ("ijarcce", "15", "8", "August", "2026", "17"))

    def test_prefixes_are_removed(self):
        d = web.parse_doi("DOI: https://doi.org/10.17148/IMRJR.2026.031105")
        self.assertEqual((d["journal"], d["volume"], d["issue"]), ("imrjr", "3", "11"))
        self.assertTrue(d["note"])                      # issue 11 or issue 1: flagged, not guessed silently

    def test_bad_doi_is_an_error(self):
        self.assertTrue(web.parse_doi("nonsense")["error"])
        self.assertTrue(web.parse_doi("10.17148/UNKNOWN.2026.15817")["error"])

    def test_auto_doi_round_trip(self):
        for journal, vol, issue, paper in (("ijarcce", 15, 11, 136), ("imrjr", 3, 8, 5), ("iarjset", 13, 1, 46)):
            doi = F.auto_doi(journal, 2026, vol, issue, paper)
            back = web.parse_doi(doi)
            self.assertEqual((back["journal"], int(back["volume"]), int(back["issue"])), (journal, vol, issue), doi)

    def test_settings_rules(self):
        s = web.clean_settings({"journal": "ijarcce", "mode": "with", "first_doi": "10.17148/IARJSET.2026.13817",
                                "first_page": "5"})
        self.assertEqual((s["journal"], s["volume"], s["issue"], s["month"], s["first_paper"]),
                         ("iarjset", "13", "8", "August", "17"))
        with self.assertRaises(ValueError):
            web.clean_settings({"mode": "without", "volume": "1", "issue": "13", "year": "2026"})


class Formatter(unittest.TestCase):
    BASELINE = {1: 56, 2: 154, 3: 429}        # paragraphs identical to the hand-finished papers

    def test_matches_hand_finished_papers(self):
        with tempfile.TemporaryDirectory() as d:
            for n, least in self.BASELINE.items():
                out = Path(d) / ("%d.docx" % n)
                F.format_paper(ROOT / "student_originals" / ("%d.docx" % n), out,
                               F.make_meta("15", "11", "November", "2026", 130 + n, 1000))
                a, b = compare.lines(out), compare.lines(ROOT / "dad_finals" / ("%d.docx" % n))
                same = sum(x.size for x in difflib.SequenceMatcher(None, a, b).get_matching_blocks())
                self.assertGreaterEqual(same, least, "paper %d" % n)

    def test_every_journal_formats(self):
        with tempfile.TemporaryDirectory() as d:
            for journal in F.JOURNAL_TEMPLATES:
                out = Path(d) / (journal + ".docx")
                fm, _ = F.format_paper(ROOT / "student_originals" / "1.docx", out,
                                       F.make_meta("3", "8", "August", "2026", 7, 1, journal=journal), journal=journal)
                self.assertTrue(out.exists() and out.stat().st_size > 20000)
                self.assertTrue(fm.title.startswith("Hint-Ladder"))

    def test_authors_for_certificates(self):
        with tempfile.TemporaryDirectory() as d:
            fm, _ = F.format_paper(ROOT / "student_originals" / "1.docx", Path(d) / "o.docx",
                                   F.make_meta("15", "11", "November", "2026", 1, 1))
            data = F.certificate_data(fm)
        self.assertEqual([a["name"] for a in data["authors"]],
                         ["Mali Dnyaneshwari Rajendra", "Prof. Manoj V. Nikum", "Prof. Shivam B. Limbhare"])
        self.assertTrue(all(a["designation"].endswith("India") for a in data["authors"]))
        self.assertEqual(data["warnings"], [])


class CertificateFit(unittest.TestCase):
    BASE = {"DOI": "10.17148/IJARCCE.2026.15817", "Volume": "15", "Issue": "8", "Year": "2026", "Month": "August"}

    def test_short_text_is_untouched(self):
        rec = dict(self.BASE, NAME="Jane Doe", Designation="Student, ABC College, India", PaperTitle="A Short Title")
        for key in ("CCE", "SET", "ICE", "RJR"):
            self.assertEqual(cert_fit.fit_level(key, rec), (0, True), key)

    def test_long_text_is_made_to_fit(self):
        long = "Department of Computer Science and Engineering, A Very Long Institute Name of Technology, City, India"
        rec = dict(self.BASE, NAME="First Author Name, Second Author Name, Third Author Name",
                   Designation=", ".join([long] * 4), PaperTitle="A Long Paper Title " * 9)
        for key in ("CCE", "SET", "ICE", "RJR"):
            level, fits = cert_fit.fit_level(key, rec)
            self.assertTrue(fits and level > 0, key)


class Api(unittest.TestCase):
    def setUp(self):
        self.c = web.app.test_client()

    def test_meta_and_health(self):
        m = self.c.get("/api/meta").get_json()
        self.assertEqual(set(m["journals"]), {"ijarcce", "iarjset", "ijireeice", "imrjr"})
        self.assertIn(self.c.get("/api/health").status_code, (200, 503))
        self.assertEqual(self.c.get("/").status_code, 200)
        self.assertEqual(self.c.get("/config.js").status_code, 200)
        self.assertEqual(self.c.get("/logos/CCE.jpg").status_code, 200)
        self.assertEqual(self.c.get("/logos/nope.jpg").status_code, 404)

    def test_issue_lifecycle(self):
        b = self.c.post("/api/batches", json={}).get_json()
        url = "/api/batches/%s" % b["id"]
        r = self.c.put(url + "/settings",
                       json={"mode": "with", "first_doi": "10.17148/IJARCCE.2026.15817", "first_page": "1"}).get_json()
        self.assertTrue(r["ready"])
        self.assertEqual(r["label"], "IJARCCE Vol. 15, Issue 8, August 2026")
        self.assertEqual(self.c.put(url + "/settings", json={"mode": "with", "first_doi": "junk"}).status_code, 400)
        self.assertEqual(self.c.post(url + "/papers", data={}).status_code, 400)
        self.assertEqual(self.c.get("/api/batches/zzzz").status_code, 404)
        self.assertEqual(self.c.delete(url).status_code, 200)

    def test_token_and_origin_guards(self):
        web.API_TOKEN = "secret"
        try:
            self.assertEqual(self.c.get("/api/meta").status_code, 401)
            self.assertEqual(self.c.get("/api/meta", headers={"X-API-Token": "secret"}).status_code, 200)
            self.assertEqual(self.c.get("/api/meta?token=secret").status_code, 200)
        finally:
            web.API_TOKEN = ""
        r = self.c.post("/api/batches", json={}, headers={"Origin": "https://evil.example"})
        self.assertEqual(r.status_code, 403)


if __name__ == "__main__":
    unittest.main()
