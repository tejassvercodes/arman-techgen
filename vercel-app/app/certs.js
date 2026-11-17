// Certificates in the browser: the certificate generator (./cert/) runs in a hidden frame and this file
// drives it exactly the way the desktop edition drives it from Node, then also makes each PDF.
(function () {
    'use strict';
    const KEYS = { ijarcce: 'CCE', iarjset: 'SET', ijireeice: 'ICE', imrjr: 'RJR' };
    let framePromise = null;
    let queue = Promise.resolve();

    function frame() {
        if (framePromise) return framePromise;
        framePromise = new Promise((resolve, reject) => {
            const f = document.createElement('iframe');
            f.src = 'cert/';
            f.setAttribute('aria-hidden', 'true');
            f.tabIndex = -1;
            f.style.cssText = 'position:fixed;left:-99999px;top:0;width:1200px;height:900px;border:0;visibility:hidden';
            f.onload = () => {
                const w = f.contentWindow;
                const t0 = Date.now();
                (function wait() {                                  // the generator is ready once its pieces exist
                    let haveState = false;
                    try { haveState = w.eval('typeof state') === 'object'; } catch (e) { /* not parsed yet */ }
                    if (haveState && w.PdfExport && w.PizZip && w.docxtemplater && w.buildCombinedDocxBlob) return resolve(w);
                    if (Date.now() - t0 > 30000) return reject(new Error('The certificate maker did not load. Check the internet connection and reload.'));
                    setTimeout(wait, 100);
                })();
            };
            f.onerror = () => reject(new Error('The certificate maker could not be loaded.'));
            (document.body || document.documentElement).appendChild(f);
        });
        framePromise.catch(() => { framePromise = null; });
        return framePromise;
    }

    const PREPARE = `(function () {
        const r = __req;
        state.templateKey = r.key;
        state.docxBuffer = __template;
        if (r.doi) { state.doiMode = 'with'; el.doiInput.value = r.doi; }
        else { state.doiMode = 'without'; el.volInput.value = r.volume; el.issueInput.value = r.issue; el.yearInput.value = r.year; el.fileNumInput.value = r.fileNumber; }
        const pub = getPublicationInfo();
        const warnings = [];
        if (pub.error) return { error: pub.error };
        if (pub.note) warnings.push(pub.note);
        const missing = describePublicationIssue(pub);
        if (missing) return { error: missing };
        const records = r.authors.map(a => ({
            NAME: a.name, Designation: a.designation, PaperTitle: r.title,
            DOI: pub.doi, Volume: pub.vol, Issue: pub.issue, Year: pub.year, Month: pub.month, __minFitLevel: 0,
        }));
        state.records = records;
        state.authorListLine = r.authors.map((a, i) => a.name + (i + 1)).join(', ');
        records.forEach((rec, i) => { if (!textFits(rec)) warnings.push(tooLongMessage('Certificate #' + (i + 1))); });
        records.forEach((rec, i) => { if (!rec.Designation) warnings.push('No designation for certificate #' + (i + 1) + '.'); });
        if (!r.title) warnings.push('No paper title.');
        return { warnings, pub };
    })()`;

    async function make(req) {
        const w = await frame();
        const authors = (req.authors || []).filter(a => a && String(a.name || '').trim())
            .map(a => ({ name: String(a.name).trim(), designation: String(a.designation || '').trim() }));
        if (!authors.length) throw new Error('No authors were given.');
        const journalKey = KEYS[req.journal];
        if (!journalKey) throw new Error('Unknown journal ' + req.journal);
        // IMRJR's template has the journal's own DOI printed on it: without a DOI it uses a copy without that line
        const key = (!req.doi && journalKey === 'RJR') ? 'RJR0' : journalKey;
        const tplRes = await fetch('cert/templates/' + key + '.docx');
        if (!tplRes.ok) throw new Error('The certificate template could not be loaded.');
        w.__template = await tplRes.arrayBuffer();
        w.__req = { key, title: String(req.title || ''), authors, doi: String(req.doi || ''),
                    volume: String(req.volume || ''), issue: String(req.issue || ''), year: String(req.year || ''),
                    fileNumber: String(req.fileNumber || '') };
        await w.PdfExport.preload(key);
        const prepared = w.eval(PREPARE);
        if (prepared.error) throw new Error(prepared.error);
        const warnings = prepared.warnings.slice();
        try {
            const missing = await w.PdfExport.missingFontsFor(key);
            if (missing && missing.length) warnings.push('This computer is missing the font' + (missing.length > 1 ? 's ' : ' ') + missing.join(', ') +
                ': the PDFs use a look-alike, so the text may look slightly different. The Word files are fine.');
        } catch (e) { /* the check is a courtesy */ }

        const files = [];
        const add = async (stem, docx, records, kind, index) => {
            let pdf = null;
            try { pdf = await w.PdfExport.build(key, records); } catch (e) { warnings.push('The PDF for "' + stem + '" could not be made (' + e.message + '). The Word file is fine.'); }
            files.push({ stem, kind, index, docx, pdf });
        };
        const stem = w.eval('buildCombinedFilename()');
        const records = w.eval('state.records');
        await add(stem, w.eval('buildCombinedDocxBlob(state.records)'), records, 'all', 0);
        for (let i = 0; i < records.length; i++) {
            w.__i = i;
            await add(w.eval('buildIndividualFilename(__i)'), w.eval('renderDocxBlobForRecord(state.records[__i])'), [records[i]], 'single', i + 1);
        }
        const group = w.eval('buildGroupRecord()');
        group.__minFitLevel = 0;
        w.__group = group;
        if (!w.eval('textFits(__group)')) warnings.push(w.eval("tooLongMessage('Group certificate')"));
        await add(stem + ' (group)', w.eval('buildCombinedDocxBlob([__group])'), [group], 'group', 0);
        const pub = prepared.pub;
        return { files, warnings, key, publication: { doi: pub.doi, volume: pub.vol, issue: pub.issue, month: pub.month, year: pub.year } };
    }

    window.Certs = {
        // one job at a time: the generator keeps its state in the frame
        make: (req) => (queue = queue.then(() => make(req), () => make(req))),
        warm: () => frame().catch(() => {}),
    };
})();
