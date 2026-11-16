# IJARCCE paper formatter

Turns a student's Word file into the IJARCCE layout and lists what still needs a human look.

## Using it

1. Open `issue.ini` and set the volume, issue, month, year, the number of the first paper and the first page number.
2. Put the students' `.docx` files in `inbox/`. They are processed in file-name order.
3. Run `python run_issue.py`.
4. Open `output/`:
   - `<name>_formatted.docx` is the formatted paper. Places that need checking carry a Word comment.
   - `<name>_report.txt` lists what was fixed and what was flagged.
   - `summary.txt` lists each paper's DOI and page range.

When a paper is reviewed, remove the comments with Review > Delete > Delete All Comments in Document.

A single paper can also be formatted directly:

```
python formatter/ijarcce_format.py paper.docx out.docx --vol 15 --issue 11 --month November --year 2026 --paper 136 --start-page 1034
```

Requirements: Python 3.10+, `pip install lxml`, and Microsoft Word (used only to count pages so page numbers run on from paper to paper).

## What it does

- A4 page, 0.8 inch margins, journal header and footer, DOI and starting page number.
- Title 24 pt; author names 12 pt bold; affiliations 10 pt italic; e-mail addresses removed from the author block.
- "Abstract:" and "Keywords:" labels in bold, text 10 pt.
- Section headings in capitals with Roman numerals; sub-headings in italic with A., B.; both renumbered.
- Body text 10 pt justified, student spacing and indents removed, one blank line between blocks.
- Figures centred and shrunk if larger than the page area; captions become "Fig. 1 ..." below.
- Tables centred, 10 pt, narrowed if too wide; captions become "TABLE I ..." in capitals above.
- Equations centred with the number at the right margin. Word equations are never altered.
- References 10 pt with hanging indent, renumbered [1]..[n].
- Student headers, footers, columns, page breaks and hyperlinks removed.

## What it flags instead of fixing

- Papers converted from PDF, where equations are broken into fragments (left exactly as submitted).
- Equations typed as plain text.
- Missing abstract, keywords, references, or figure and table captions.
- Figure or table numbers out of sequence.
- Floating images, text boxes, images inside tables or inside a line of text.
- Headings it recognised only from their formatting.

It does not change wording, title capitalisation, or the order of author and affiliation lines, and it does not make by-eye decisions about image size beyond fitting the page.

## Template files

- `templates/CCE.docx` is the journal's template, untouched.
- `templates/CCE_original_backup.docx` is a byte-for-byte backup of it.
- `templates/CCE_production.docx` is the copy the formatter uses. It is built by `formatter/build_template.py` and differs from the original as follows:

| Change | Why |
|---|---|
| Header replaced with the one in the published papers (Impact Factor 8.471, "Peer-reviewed & Refereed journal", 9 pt) | The original still says 8.102 and uses older wording |
| Volume, issue, month, year and DOI in the header are fill-in tokens (`{{VOL}}`, `{{ISSUE}}`, `{{MONTH}}`, `{{YEAR}}`, `{{DOI}}`) | So the formatter can fill them per paper |
| One header and one footer instead of four of each (empty "even page" and "first page" ones removed) | The extra ones were blank and unused |
| Author names: new style, 12 pt bold | Matches the published papers (rule text says 11 pt regular) |
| Abstract 10 pt regular; captions, table text and references 10 pt | Matches the published papers (styles were 9 pt bold, 8 pt, 9 pt, 8 pt) |
| Paragraphs: no first-line indent, no space before or after | Published papers use blank lines instead |
| Heading and reference styles no longer auto-number | The formatter types the numbers, so they cannot drift |
| New styles: IEEE Heading 3, IEEE List Item, IEEE Equation, IEEE Code | Needed for content the original had no style for |

The instruction text in the body of the template was not edited, so a few of its written rules (for example "references must be 8 pt") no longer match the styles. To change the impact factor or header wording later, edit the header in `CCE_production.docx` directly in Word and keep the `{{...}}` tokens.

## DOI format

`10.17148/IJARCCE.<year>.<volume><issue><paper number>`, for example volume 15, issue 11, paper 136 gives `10.17148/IJARCCE.2026.1511136`. Paper numbers below 10 are written with two digits.

## Checking against hand-finished papers

`python formatter/compare.py output/2_formatted.docx dad_finals/2.docx` lists the text differences between a formatted paper and a hand-finished one.

## Website

Double-click `run_website.bat` (or run `python webapp/app.py`) and open http://127.0.0.1:5000. It runs only on this computer. Needs `pip install flask lxml pymupdf` and Microsoft Word.

- **Issues:** each issue (journal, volume, issue, month, year, first paper number, first page) is saved with its papers. Switch between issues from the top bar; start a new one with New issue.
- **Upload:** drop several student files at once (.docx, and .doc/.rtf/.odt, which Word converts). Papers are formatted one after another in the background.
- **Automatic numbering:** DOIs and first-page numbers run on from paper to paper using the real page counts from Word. Change the DOI or first page of any one paper and the following papers update; clear the box to go back to automatic.
- **Per paper:** Preview (page images, with the list of things to check), download the Word file, PDF or report, see the details, move up or down, format again, remove.
- **Whole issue:** Download all gives a ZIP with every formatted paper, the reports and a `contents.csv` (paper number, title, DOI, pages). Changing the journal, volume, issue, month or year formats all papers again.
- Files are kept in `webapp/data/`. Delete an issue from the top bar when it is finished.

## Second journal (IARJSET) and PDF-converted papers

- `--journal iarjset` (or the Journal box on the website) uses the IARJSET header and footer from `templates/IARJSET_production.docx` (ISSN, impact factor 8.311, no DOI line). The header and footer come from `templates/sources/`; `python formatter/build_template.py` rebuilds both production templates.
- Tables that a PDF conversion left inside floating text boxes are moved into the text. Paragraphs cut in two by a column or page break are joined and marked with a comment, because a few of those joins may be wrong when the conversion scrambled the order of the text.
- Authors listed one per block, each with an affiliation, become one line "A1, B2, C3" with numbered affiliations; identical affiliations are written once. Registration numbers are removed.
- Tables get a full grid of borders and a bold first row; floating figures are made in-line; "TABLE II." and "Fig. 6." lines are recognised as captions wherever they sit.

## Badly converted student files (PDF to Word)

Some submissions are PDFs that were converted to Word: text in columns, loose text instead of tables, equations in pieces. The formatter now:

- cuts apart author blocks, reads "Author Name:" / "Guide Name:" labels, drops e-mail lines, numbers the authors and writes identical affiliations once;
- keeps the abstract as one paragraph and moves "Keywords" that were typed inside it onto their own line;
- recognises section headings typed with a colon or all in one style ("Introduction :", "Research Gap :"), and puts back drop-cap letters;
- rejoins words hyphenated at line ends (only when the whole word appears elsewhere in the paper) and, if the paper uses punctuation, paragraphs cut by a column break;
- numbers captions that the student styled but never numbered;
- **takes a picture of a table or equation from the student's own page when the file only holds it as loose pieces.** This needs Microsoft Word (it makes a PDF of the student's file) and `pip install pymupdf`. Each picture is flagged in the report and with a comment in the Word file so it can be retyped if needed.

The blank line after the title is 10 pt (the title itself is 24 pt).

## Paper + certificates in one click

Pick the journal (IJARCCE, IARJSET, IJIREEICE or IMRJR), fill in the issue details once and drop in the students' files. For every paper the website:

1. formats it (header, footer, DOI and page numbers for that journal);
2. reads the title, the authors and each author's designation (their affiliation) from the formatted paper;
3. fills that journal's certificate template with the certificate generator's own code (`certificate-generator/make-certificates.js`): one file with all certificates, one per author and one group certificate;
4. lets Word make PDFs of all of them.

Each paper card shows the downloads (Word and PDF) and **Check / edit names**: a table of the names, designations and title that went on the certificates. Fix anything that was read wrongly, press *Save and re-make certificates* (only the certificates are redone), or go back to what was read from the paper. **Download everything** gives one ZIP per paper (paper, report, all certificates); the whole-issue ZIP has a `certificates/` folder per paper and a Certificates column in `contents.csv`.

Needs Node.js (for the certificates) as well as Python and Word. The DOI is `10.17148/<JOURNAL>.<year>.<volume as 2 digits><issue><paper no. as 2 digits>`, which is exactly what the certificate generator reads back (volume, issue, month, file number). The paper templates for the four journals are `templates/*_production.docx`, built by `formatter/build_template.py` from the journals' own `*-Paper-Format.docx` files. IARJSET uses the journal's current template (impact factor 8.066, with a DOI line); the older header from your father's published final (8.311, no DOI line) is kept as `templates/IARJSET_published_production.docx` and is not used.

## Issue details work like the certificate generator's "With DOI / Without DOI"

- **With DOI** (default): paste the DOI of the issue's first paper, for example `10.17148/IJARCCE.2026.15817`. Volume (15), issue (8), month (August), year (2026) and the first paper number (17) are read from it, and the journal too (a pasted `DOI:` or `https://doi.org/` prefix is fine). A line under the box shows what was read, with a warning if the DOI could be read two ways.
- **Without DOI**: type volume, issue, year and the first paper number.
- The **month is never typed**: the issue number is the month (issue 8 = August).
- The **first page** is always typed. The DOIs and first pages of the following papers run on automatically.

## Certificates: long text no longer pushes the DOI off the certificate

The certificate text box has a fixed size, so a long title, a long designation or a group certificate used to lose its Volume and DOI lines. Each certificate is now measured (`webapp/cert_fit.py`, the same layout rules as the certificate generator, with the real Windows fonts) and its text set just small enough to fit; the certificate maker writes that into the Word file and Word makes the PDF. A certificate that already fits is unchanged. If even the smallest size is not enough the card says which certificate and asks for a shorter designation.
