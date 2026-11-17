"""Builds vercel-app/ (the browser edition) from the desktop edition's files.

    python build_web.py

Copies the Python (formatter, shared core, page estimator), the four paper templates and the logos, and turns
webapp/static/index.html into vercel-app/index.html.  The certificate maker (vercel-app/cert/) is a copy of
certificate-generator/ made once with tools in that folder (see vercel-app/README.md); it is not touched here.
"""
import re
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WEB = ROOT / "vercel-app"

# ---- plain copies
for f in ("ijarcce_format.py", "core.py", "pagecount.py"):
    shutil.copyfile(ROOT / "formatter" / f, WEB / "py" / f)
(WEB / "templates").mkdir(exist_ok=True)
for t in ("CCE", "IARJSET", "IJIREEICE", "IMRJR"):
    shutil.copyfile(ROOT / "templates" / (t + "_production.docx"), WEB / "templates" / (t + "_production.docx"))
(WEB / "logos").mkdir(exist_ok=True)
for f in (ROOT / "certificate-generator" / "logos").glob("*"):
    shutil.copyfile(f, WEB / "logos" / f.name)

# ---- the page
s = (ROOT / "webapp" / "static" / "index.html").read_text(encoding="utf-8")


def rep(old, new, count=1):
    global s
    assert s.count(old) == count, (s.count(old), old[:90])
    s = s.replace(old, new)


def cut(start, end, new):
    """Replace everything from `start` up to (not including) `end`."""
    global s
    assert s.count(start) == 1 and s.count(end) == 1, (s.count(start), s.count(end), start[:60])
    i, j = s.index(start), s.index(end)
    assert i < j
    s = s[:i] + new + s[j:]


# scripts: the libraries (same versions and hashes as the certificate maker), then our two files
rep('<script src="config.js"></script>', '''<script src="https://unpkg.com/pizzip@3.2.0/dist/pizzip.js" integrity="sha384-EKMdeygnsgFUPocowqWjGXhnzFAiDXk3KuOx8g8lsB5/JNEQYTWaADHxmioDiFU8" crossorigin="anonymous"></script>
<script src="https://unpkg.com/jszip@3.10.1/dist/jszip.min.js" integrity="sha384-+mbV2IY1Zk/X1p/nWllGySJSUN8uMs+gUAN10Or95UBH0fpj6GfKgPmgC5EXieXG" crossorigin="anonymous"></script>
<script src="https://unpkg.com/docx-preview@0.4.0/dist/docx-preview.min.js" integrity="sha384-UkwbeBm1NknJfLd5UU4RI1j7PidBeZE3tIiKY1k+n7RQ+kEJ0dkTC0u53xTtRC9r" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js"></script>
<script>if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";</script>
<script src="app/certs.js"></script>
<script src="app/localapi.js"></script>''')

# a private tool: no search engines, a clear message for browsers that cannot run it, and a warning before work is lost
rep('<meta name="viewport" content="width=device-width, initial-scale=1.0">', '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<meta name="robots" content="noindex, nofollow">')
rep('<div class="container">', '''<noscript><p style="padding:1rem;color:#fcd34d">This page needs JavaScript. Please turn it on and reload.</p></noscript>
<p id="compat" role="alert" hidden style="margin:0;padding:.7rem 1rem;background:#7f1d1d;color:#fff;text-align:center">This browser is too old for this page. Please open it in a current Chrome, Edge, Firefox or Safari.</p>
<div class="container">''')
# the "server" is in this browser
cut('const API_BASE = ', 'const jsonOpts = ', '''const link = (path) => LocalApi.link(path);
const api = (path, opts) => LocalApi.api(path, opts);
''')
rep('a.href = link(path); a.download = ""; a.style.display = "none";', 'a.href = link(path); a.download = LocalApi.filename(path); a.style.display = "none";')

# word of a paper: Word only (no PDF in this edition); one ZIP built here
rep('''      done ? h("button", { class: "linkbtn" + on("paper"), type: "button", title: "Download the paper as Word and PDF", "aria-label": "Download the paper as Word and PDF",
        onclick: e => { pick(e, "paper"); downloadMany(["/files/" + cur + "/" + p.id + "/docx", "/files/" + cur + "/" + p.id + "/pdf"]); } }, h("i", { class: "fa-solid fa-download" }), " Word + PDF") : null,''',
    '''      done ? h("button", { class: "linkbtn" + on("paper"), type: "button", title: "Download the formatted paper (Word)", "aria-label": "Download the formatted paper (Word)",
        onclick: e => { pick(e, "paper"); downloadMany(["/files/" + cur + "/" + p.id + "/docx"]); } }, h("i", { class: "fa-solid fa-download" }), " Word") : null,''')
rep('''      done ? h("a", { class: "linkbtn" + on("all"), href: link("/files/" + cur + "/" + p.id + "/bundle"), title: "Paper (Word and PDF), report and all certificates in one ZIP", onclick: e => pick(e, "all") }, "Download everything") : null,''',
    '''      done ? h("button", { class: "linkbtn" + on("all"), type: "button", title: "The paper, its report and all certificates (Word and PDF) in one ZIP",
        onclick: async e => { pick(e, "all"); try { const z = await LocalApi.bundle(p.id); const a = document.createElement("a"); a.href = z.url; a.download = z.name; a.style.display = "none"; document.body.append(a); a.click(); setTimeout(() => { a.remove(); URL.revokeObjectURL(z.url); }, 4000); } catch (ex) { toast(ex.message, true); } } }, "Download everything") : null,''')
rep('''    paths.push("/files/" + cur + "/" + p.id + "/docx", "/files/" + cur + "/" + p.id + "/pdf");''', '''    paths.push("/files/" + cur + "/" + p.id + "/docx");''')
rep('data.papers.filter(p => p.status === "done" && p.pages != null)', 'data.papers.filter(p => p.status === "done" && p.pages != null)', count=s.count('data.papers.filter(p => p.status === "done" && p.pages != null)'))
rep('title: "Every paper (Word and PDF) and every certificate (Word and PDF, one per author)"', 'title: "Every paper (Word) and every certificate (Word and PDF, one per author)"', count=s.count('title: "Every paper (Word and PDF) and every certificate (Word and PDF, one per author)"')) if 'title: "Every paper (Word and PDF) and every certificate' in s else None
s = s.replace('Every paper (Word and PDF) and every certificate (Word and PDF, one per author)', 'Every paper (Word) and every certificate (Word and PDF, one per author)')

# the page count is an estimate in this edition: shown, and editable
rep('''    done && p.doi ? h("div", { class: "prow" },
      h("div", { class: "fld doi" },
        h("div", { class: "doi-wrap" }, h("span", {}, prefixNow()), doi)),
      ) : null,''', '''    done ? h("div", { class: "prow" },
      p.doi ? h("div", { class: "fld doi" }, h("div", { class: "doi-wrap" }, h("span", {}, prefixNow()), doi)) : null,
      h("div", { class: "fld pgs" + (p.unsure && p.unsure.length && !p.pages_manual ? " unsure" : "") },
        h("label", {}, p.pages_manual ? "Pages (yours)" : "Pages (estimate)"),
        h("input", { type: "number", min: "1", max: "999", value: p.pages != null ? p.pages : "", "aria-label": "Number of pages of this paper in Word",
          title: "The number of pages Word shows for this paper. Change it if Word shows a different number: the next papers are renumbered.",
          onchange: e => patch(p.id, { pages: e.target.value }) }))) : null,
    done && p.unsure && p.unsure.length && !p.pages_manual ? h("p", { class: "pgnote" }, "Check the page count in Word: " + p.unsure[0]) : null,''')

# previews are drawn here from the Word files (docx-preview), not from PDFs
cut('async function previewCerts(p) {', '''/* live preview of the "all certificates in one file" PDF of the selected paper */''', '''// Word files drawn in the page (docx-preview): `fit` is "width" (pages stacked) or "page" (one page fills the box).
async function renderDocx(box, source, fit) {
  box.replaceChildren();
  const host = document.createElement("div");
  box.append(host);
  await docx.renderAsync(source, host, null, { className: "docx", inWrapper: false, ignoreWidth: false, ignoreHeight: false,
    breakPages: true, useBase64URL: true, renderHeaders: true, renderFooters: true });
  const secs = [...host.querySelectorAll("section.docx")];
  const bw = box.clientWidth - (fit === "page" ? 0 : 8), bh = box.clientHeight;
  secs.forEach(sec => {
    const sw = sec.offsetWidth, sh = sec.offsetHeight;
    const k = fit === "page" ? Math.min(bw / sw, bh / sh) : Math.min(1, bw / sw);
    const pg = document.createElement("div");
    pg.className = "pgbox";
    pg.style.cssText = fit === "page" ? "height:100%;display:flex;align-items:center;justify-content:center" : "margin:0 auto 0.9rem";
    const inner = document.createElement("div");
    inner.style.cssText = "width:" + Math.round(sw * k) + "px;height:" + Math.round(sh * k) + "px;overflow:hidden;flex:none;box-shadow:0 2px 12px rgba(0,0,0,.45)";
    sec.style.transformOrigin = "top left";
    sec.style.transform = "scale(" + k + ")";
    sec.style.margin = "0";
    sec.replaceWith(pg);
    pg.append(inner);
    inner.append(sec);
  });
  host.replaceChildren(...[...host.querySelectorAll(".pgbox")]);
  return secs.length;
}
// The certificate PDF drawn page by page (pdf.js): what is shown is the very file that gets downloaded.
async function renderPdf(box, blob, fit) {
  const token = box._tok = (box._tok || 0) + 1;            // only the newest drawing may fill the box
  const pdf = await pdfjsLib.getDocument({ data: await blob.arrayBuffer() }).promise;
  const bw = box.clientWidth - (fit === "page" ? 0 : 8), bh = box.clientHeight, dpr = window.devicePixelRatio || 1;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    if (box._tok !== token) return 0;
    const page = await pdf.getPage(i);
    const v1 = page.getViewport({ scale: 1 });
    const k = fit === "page" ? Math.min(bw / v1.width, bh / v1.height) : Math.min(1.6, bw / v1.width);
    const v = page.getViewport({ scale: k * dpr });
    const cv = document.createElement("canvas");
    cv.width = Math.floor(v.width); cv.height = Math.floor(v.height);
    cv.style.cssText = "width:" + Math.floor(v.width / dpr) + "px;height:" + Math.floor(v.height / dpr) + "px;display:block;background:#fff;box-shadow:0 2px 12px rgba(0,0,0,.45)";
    cv.setAttribute("role", "img"); cv.setAttribute("aria-label", "Certificate " + i + " of " + pdf.numPages);
    await page.render({ canvasContext: cv.getContext("2d"), viewport: v }).promise;
    const pg = document.createElement("div");
    pg.className = "pgbox";
    pg.style.cssText = fit === "page" ? "height:100%;display:flex;align-items:center;justify-content:center" : "margin:0 auto 0.9rem;width:fit-content";
    pg.append(cv);
    pages.push(pg);
  }
  if (box._tok !== token) return 0;
  box.replaceChildren(...pages);
  return pdf.numPages;
}
async function renderCerts(box, p, fit) {
  const all = p.certs.files.find(f => f.kind === "all");
  const pdf = LocalApi.certPdf(p.id, all.stem);
  if (pdf && window.pdfjsLib) return renderPdf(box, pdf, fit);
  const docxSrc = LocalApi.certDocx(p.id, all.stem);
  if (!docxSrc) throw new Error("The certificates are not ready.");
  return renderDocx(box, docxSrc, fit);
}
async function previewCerts(p) {
  const all = p.certs && p.certs.files ? p.certs.files.find(f => f.kind === "all") : null;
  const src = all ? LocalApi.certDocx(p.id, all.stem) : null;
  $("mTitle").textContent = "Certificates — " + (p.title || p.name);
  $("mFlags").replaceChildren(h("h3", {}, "Certificates"),
    ...((p.cert_data && p.cert_data.authors) || []).map((a, i) => h("div", { class: "good", style: "margin-bottom:.4rem" }, (i + 1) + ". " + a.name, h("small", { style: "display:block;opacity:.8" }, a.designation))));
  modal.showModal();
  if (!src) { $("mPages").replaceChildren(h("div", { class: "loading" }, "The certificates are not ready.")); return; }
  $("mPages").replaceChildren(h("div", { class: "loading" }, "Preparing the preview…"));
  try { await renderCerts($("mPages"), p, "width"); } catch (ex) { $("mPages").replaceChildren(h("div", { class: "loading" }, ex.message)); }
}
''')
cut('''  if (!$("cpPages").querySelector("img")) cpNote("Preparing the preview…");''', '''function showTab(name) {''', '''  if (!$("cpPages").querySelector(".pgbox")) cpNote("Preparing the preview…");
  try {
    await renderCerts($("cpPages"), p, "page");
    if (run !== cpRun) return;                           // a newer request has taken over
    $("cpPages").scrollTop = 0;
  } catch (ex) { if (run === cpRun) { cpKey = ""; cpNote(ex.message); } }
}
''')
cut('async function preview(p) {', '''/* ---------- start ---------- */''', '''async function preview(p) {
  $("mTitle").textContent = (p.title || p.name);
  $("mPages").replaceChildren(h("div", { class: "loading" }, "Preparing the preview…"));
  $("mFlags").replaceChildren(
    h("h3", {}, p.flags.length ? "Needs your eye" : "All clear"),
    ...(p.flags.length ? p.flags.map(f => h("div", { class: "flag" }, f.message + (f.more ? " (and " + f.more + " more)" : ""), f.near ? h("small", {}, "near: “" + f.near + "”") : null))
                   : [h("div", { class: "good" }, "Nothing was flagged.")]),
    h("p", { class: "note" }, "This preview is drawn by the browser and can differ slightly from Word (fonts, spacing). Open the Word file to check the final look."));
  modal.showModal();
  const src = LocalApi.docxBytes(p.id);
  if (!src) { $("mPages").replaceChildren(h("div", { class: "loading" }, "The paper is not ready.")); return; }
  try { await renderDocx($("mPages"), src, "width"); } catch (ex) { $("mPages").replaceChildren(h("div", { class: "loading" }, ex.message)); }
}

''')

rep('Word files (.docx, .doc). You can add several at once.', 'Word files (.docx). You can add several at once.', count=s.count('Word files (.docx, .doc). You can add several at once.'))
rep('accept=".docx,.doc,.rtf,.odt"', 'accept=".docx"')
# start the engines while the user fills in the details
rep('''    const meta = await api("/api/meta");''', '''    log("Starting the formatter in this browser (the first visit loads about 15 MB)…");
    window.Certs.warm();
    const meta = await api("/api/meta");
    log("Formatter ready.", "success");''')
rep('Workspace started. Choose a journal, fill in the issue details and upload the manuscripts.', 'Workspace started. Everything stays in this browser tab; nothing is uploaded.') if 'Workspace started. Choose a journal' in s else None

# styles for the new bits
rep('</style>', '''  .pgs { max-width: 9rem; }
  .pgs input { width: 5rem; }
  .pgs.unsure input { border-color: #f59e0b; box-shadow: 0 0 0 2px rgba(245,158,11,.2); }
  .pgnote { margin: 0.35rem 0 0; font-size: 0.78rem; color: #fcd34d; }
  .cprev .cppages { position: relative; }
  .cprev .cppages .pgbox { scroll-snap-align: start; scroll-snap-stop: always; }
  .mpages .pgbox { background: transparent; }
</style>''')
rep('</script>\n</body>', '''</script>
<script>
// everything lives in this tab: warn before it is closed with papers in it, and say so if the browser cannot run the page
(function () {
  if (!window.Worker || !window.WebAssembly || !window.HTMLDialogElement) document.getElementById("compat").hidden = false;
  window.addEventListener("beforeunload", function (e) {
    if (typeof data !== "undefined" && data && data.papers && data.papers.length) { e.preventDefault(); e.returnValue = ""; }
  });
})();
</script>
</body>''')
(WEB / "index.html").write_text(s, encoding="utf-8")
print("built", WEB / "index.html")
