# Putting Paper Desk in the cloud (no PC needed)

The app runs in a container on Linux and uses LibreOffice for page counts, PDFs and previews
(on your own Windows PC it still uses Microsoft Word automatically).

## 1. Fonts (do this once, before building)
The certificates use Calibri, Algerian, Calisto MT, Times New Roman, Arial and others. Copy these files from
`C:\Windows\Fonts` into the `fonts` folder next to the Dockerfile, keeping the file names:
`calibri.ttf calibrib.ttf calibrii.ttf calibriz.ttf ALGER.TTF times.ttf timesbd.ttf timesi.ttf timesbi.ttf
arial.ttf arialbd.ttf ariali.ttf arialbi.ttf CALIST.TTF CALISTB.TTF CALISTI.TTF CALISTBI.TTF`
Without them the app still works, but the certificates use look-alike fonts and may differ slightly.
Keep the repository private if fonts are included.

## 2. Host (example: Render, a Docker web service)
- New Web Service from the repository, runtime **Docker**, an instance with at least 1 GB RAM.
- Environment variables:
  - `PAPERDESK_TOKEN` = a long secret (everyone must enter it once in the browser)
  - `PAPERDESK_MAX_UPLOAD_MB` = 100 (optional)
- Open the service address. The page asks for the token the first time.

## 3. Check before relying on it
Page numbers come from LibreOffice here, not Word. Format a few finished papers and compare the page counts
with Word's. If they differ, the first-page numbers of the following papers will shift; use "Edit page#".
