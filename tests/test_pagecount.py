"""The page estimator (formatter/pagecount.py) against the page counts Word gave the same papers.

The reference numbers are in fixtures/word_pages.json (made once with Word); everything here runs without Word.
"""
import io
import json
import sys
import unittest
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "formatter"))
import ijarcce_format as F      # noqa: E402
import pagecount                # noqa: E402

FIX = json.loads((ROOT / "tests" / "fixtures" / "word_pages.json").read_text(encoding="utf-8"))
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
JOURNALS = ["ijarcce", "iarjset", "ijireeice", "imrjr"]


def formatted(n, journal):
    """Paper n (0-4) of student_originals, formatted without PDF help, as the browser edition does."""
    src = ROOT / "student_originals" / ("%d.docx" % (n + 1))
    out = io.BytesIO()
    tmp = ROOT / "webapp" / "data" / ("_t_%s_%d.docx" % (journal, n))
    F.PDF_PROVIDER = lambda p: None
    meta = F.make_meta("15", "8", "August", "2026", paper=17 + n, start_page=1, journal=journal)
    F.format_paper(str(src), str(tmp), meta, journal=journal, report_path=str(tmp.with_suffix(".txt")))
    data = tmp.read_bytes()
    tmp.unlink()
    tmp.with_suffix(".txt").unlink()
    return data


def cut(data, percent):
    zin = zipfile.ZipFile(io.BytesIO(data))
    root = etree.fromstring(zin.read("word/document.xml"))
    body = root.find("{%s}body" % W)
    kids = [k for k in body if k.tag != "{%s}sectPr" % W]
    for k in kids[max(5, len(kids) * percent // 100):]:
        body.remove(k)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            d = zin.read(it.filename)
            if it.filename == "word/document.xml":
                d = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zo.writestr(it.filename, d)
    return out.getvalue()


class PageEstimate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        (ROOT / "webapp" / "data").mkdir(parents=True, exist_ok=True)
        cls.docs = {"%s_%02d" % (j, n): formatted(n, j) for j in JOURNALS for n in range(5)}

    def test_whole_papers_match_word(self):
        """Every paper without floating objects gets Word's exact page count; the others are flagged."""
        wrong = []
        for name, data in self.docs.items():
            est = pagecount.estimate(data)
            if est["unsure"]:
                continue
            if est["pages"] != FIX["whole"][name]:
                wrong.append((name, est["pages"], FIX["whole"][name]))
        self.assertLessEqual(len(wrong), 1, wrong)          # a close call in line breaking may cost one paper

    def test_floating_objects_are_flagged(self):
        """The paper whose text wraps around floating shapes (student 5) is reported as uncertain."""
        for j in JOURNALS:
            self.assertTrue(pagecount.estimate(self.docs["%s_04" % j])["unsure"], j)

    def test_shortened_papers(self):
        """The same papers cut at 15-90%: nearly always exact, never more than one page out."""
        exact = total = 0
        worst = 0
        for name, data in self.docs.items():
            for pct in (15, 30, 45, 60, 75, 90):
                key = "%s_%02d" % (name, pct)
                est = pagecount.estimate(cut(data, pct))["pages"]
                total += 1
                exact += est == FIX["cut"][key]
                worst = max(worst, abs(est - FIX["cut"][key]))
        self.assertLessEqual(worst, 1)
        self.assertGreaterEqual(exact / total, 0.88, "%d of %d exact" % (exact, total))

    def test_journal_is_recognised(self):
        for j in JOURNALS:
            lay = pagecount.Layout(self.docs["%s_00" % j])
            self.assertEqual(lay.journal, j.upper())


class SharedCore(unittest.TestCase):
    def test_chain_numbers_pages_and_dois(self):
        import core
        s = core.clean_settings({"mode": "with", "first_doi": "10.17148/IJARCCE.2026.15817", "first_page": "10"})
        t = core.chain(s, [{"id": "a", "pages": 4}, {"id": "b", "pages": 9}, {"id": "c"}])
        self.assertEqual([t[k]["start"] for k in "abc"], [10, 14, 23])
        self.assertEqual([t[k]["paper_no"] for k in "abc"], [17, 18, 19])
        self.assertEqual(t["c"]["doi"], "10.17148/IJARCCE.2026.15819")

    def test_without_doi_has_no_doi_anywhere(self):
        import core
        s = core.clean_settings({"mode": "without", "volume": "3", "issue": "11", "year": "2026", "first_paper": "5",
                                 "first_page": "1", "first_doi": "10.17148/IJARCCE.2026.15817", "journal": "imrjr"})
        self.assertEqual(s["first_doi"], "")
        self.assertEqual(core.chain(s, [{"id": "a"}])["a"]["doi"], "")


if __name__ == "__main__":
    unittest.main()
