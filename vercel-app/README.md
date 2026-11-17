# Arman - Workspace, browser edition (Vercel)

Everything runs in the visitor's browser: the formatter (Python through Pyodide, in a background worker),
the page estimate and the certificates. There is no server code, nothing is uploaded and nothing is kept
after the tab is closed. Vercel only serves these files.

## Deploy
```
cd vercel-app
vercel --prod
```
(or import the repository in Vercel and set the **Root Directory** to `vercel-app`; no build step, no framework.)

## What is different from the desktop edition
- Page counts are **estimated** (`py/pagecount.py`) instead of counted by Word. Each paper shows "Pages (estimate)";
  type Word's number there if it differs and every following paper is renumbered. Papers with floating pictures or
  text boxes are marked "check the page count".
- No PDF of the paper and no PDF preview (save as PDF in Word). Paper preview is drawn by the browser.
- `.docx` only (`.doc`, `.rtf`, `.odt` must be saved as .docx in Word first).
- Broken tables/equations from PDF-converted papers are flagged, not rebuilt.
- Certificate PDFs use the fonts installed on the visitor's computer (Windows with Office fonts: exact).

## Keeping it in step with the desktop edition
```
python build_web.py      # from the project root: copies the Python, templates, logos and rebuilds index.html
```
`cert/` is a copy of `certificate-generator/` plus `templates/RJR0.docx` and `templates/pdf/RJR0.*` (IMRJR without
the journal's printed DOI); rebuild those with the generator's `tools/build-pdf-assets.js` (needs Word) if a
certificate template changes.

## Tests
`python -m unittest discover -s tests` checks the page estimator against page counts Word gave
(`tests/fixtures/word_pages.json`).
