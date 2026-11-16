"""Show text-level differences between a formatter output and a hand-finished paper."""
import difflib, io, sys, zipfile
from lxml import etree
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
from ijarcce_format import ptext, w

def lines(path):
    doc = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    out = []
    for p in doc.iter(w("p")):
        t = " ".join(ptext(p).split())
        if t:
            out.append(t)
    return out

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    a, b = lines(sys.argv[1]), lines(sys.argv[2])
    same = sum(1 for x in difflib.SequenceMatcher(None, a, b).get_matching_blocks() for _ in range(x.size))
    print("%s vs %s: %d of %d paragraphs identical" % (sys.argv[1], sys.argv[2], same, len(b)))
    for l in difflib.unified_diff(a, b, "formatter", "final", lineterm="", n=0):
        if not l.startswith(("---", "+++", "@@")):
            print("  " + l[:150])
