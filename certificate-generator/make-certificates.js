#!/usr/bin/env node
// Command-line certificate maker.
//
// It runs the live site's own app.js (unchanged) inside a sandbox with a stand-in for the page, so the
// certificates it writes are produced by exactly the same code as the website's buttons: same template
// filling, same DOI reading, same combined / group documents, same file names.
//
//   node make-certificates.js request.json
//
// request.json:
//   { "journal": "CCE" | "SET" | "ICE" | "RJR",
//     "title": "Paper title",
//     "authors": [ { "name": "Jane Doe", "designation": "Student, ABC College, India" }, ... ],
//     "doi": "10.17148/IJARCCE.2026.15817"      // or "" for no DOI
//     "volume": "15", "issue": "8", "year": "2026", "fileNumber": "17",   // used when there is no DOI
//     "outDir": "C:\\path\\to\\folder" }
//
// Writes (Word only) into outDir:
//   "<file number> <authors>.docx"            all certificates in one file
//   "<file number> <authors> <n>.docx"        one file per author
//   "<file number> <authors> (group).docx"    one certificate for everybody
// and prints a JSON summary { files: [...], warnings: [...] } on stdout.
//
// No npm packages: the two libraries the page loads from a CDN are kept in vendor/ (same versions,
// same hashes as the <script> tags in index.html).
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = __dirname;
const GROUP_SUFFIX = ' (group)';

function fail(message) {
    process.stdout.write(JSON.stringify({ error: message }) + '\n');
    process.exit(1);
}

// A stand-in for any DOM element: absorbs every property read, call and write.
function makeStub(name) {
    const store = { value: '', textContent: '', innerHTML: '', hidden: false, disabled: false, checked: false };
    const fn = function () {};
    return new Proxy(fn, {
        get(_t, prop) {
            if (prop in store) return store[prop];
            if (prop === Symbol.toPrimitive) return () => '';
            if (prop === 'classList') return { add() {}, remove() {}, toggle() {}, contains() { return false; } };
            if (prop === 'style' || prop === 'dataset') return {};
            if (prop === 'children' || prop === 'childNodes') return [];
            return makeStub(String(prop));
        },
        set(_t, prop, value) { store[prop] = value; return true; },
        apply() { return makeStub('result'); },
    });
}

function loadSite() {
    const elements = new Map();
    const document = {
        readyState: 'loading',                       // keeps the page's init() from running
        getElementById(id) { if (!elements.has(id)) elements.set(id, makeStub(id)); return elements.get(id); },
        querySelector() { return makeStub('q'); },
        querySelectorAll() { return []; },
        createElement() { return makeStub('e'); },
        addEventListener() {},
        body: makeStub('body'),
    };
    const sandbox = {
        document, Blob, URL, console, Buffer, setTimeout, clearTimeout, TextEncoder, TextDecoder,
        Uint8Array, ArrayBuffer, Promise, Map, Set, Date, JSON, Math, parseInt, parseFloat, isNaN,
        navigator: { userAgent: 'node' }, location: { href: 'http://localhost/' },
        fetch() { throw new Error('fetch is not available here'); },
        matchMedia() { return { matches: false, addEventListener() {} }; },
        addEventListener() {},
    };
    const xmldom = require('./vendor/xmldom');          // docxtemplater needs an XML parser; browsers have one built in
    sandbox.DOMParser = xmldom.DOMParser;
    sandbox.XMLSerializer = xmldom.XMLSerializer;
    sandbox.window = sandbox;
    sandbox.self = sandbox;
    sandbox.globalThis = sandbox;
    const context = vm.createContext(sandbox);
    for (const file of ['vendor/pizzip-3.2.0.js', 'vendor/docxtemplater-3.37.11.js', 'app.js']) {
        vm.runInContext(fs.readFileSync(path.join(ROOT, file), 'utf8'), context, { filename: file });
    }
    return context;
}

async function main() {
    const requestPath = process.argv[2];
    if (!requestPath) fail('Usage: node make-certificates.js request.json');
    let req;
    try { req = JSON.parse(fs.readFileSync(requestPath, 'utf8')); } catch (e) { fail('Could not read the request: ' + e.message); }

    const key = String(req.journal || '').toUpperCase();
    if (!['CCE', 'SET', 'ICE', 'RJR'].includes(key)) fail('Unknown journal "' + req.journal + '" (use CCE, SET, ICE or RJR).');
    const authors = (req.authors || []).filter(a => a && String(a.name || '').trim());
    if (authors.length === 0) fail('No authors were given.');
    if (!req.outDir) fail('No outDir was given.');

    const ctx = loadSite();
    const templatePath = path.join(ROOT, 'templates', key + '.docx');
    const tpl = fs.readFileSync(templatePath);

    ctx.__req = {
        key,
        title: String(req.title || ''),
        authors: authors.map(a => ({ name: String(a.name).trim(), designation: String(a.designation || '').trim() })),
        doi: String(req.doi || ''),
        volume: String(req.volume || ''), issue: String(req.issue || ''),
        year: String(req.year || ''), fileNumber: String(req.fileNumber || ''),
        minFit: Math.max(0, parseInt(req.minFitLevel, 10) || 0),   // ask for smaller text (after a check showed a cut-off)
        // exact fit levels measured by the caller with the real fonts: one per author, then the group
        fitLevels: Array.isArray(req.fitLevels) ? req.fitLevels : null,
    };
    ctx.__template = tpl.buffer.slice(tpl.byteOffset, tpl.byteOffset + tpl.byteLength);

    // Everything below runs inside the sandbox, against the website's own functions.
    const prepare = `
        (function () {
            const r = __req;
            state.templateKey = r.key;
            state.docxBuffer = __template;
            if (r.doi) {
                state.doiMode = 'with';
                el.doiInput.value = r.doi;
            } else {
                state.doiMode = 'without';
                el.volInput.value = r.volume;
                el.issueInput.value = r.issue;
                el.yearInput.value = r.year;
                el.fileNumInput.value = r.fileNumber;
            }
            const pub = getPublicationInfo();
            const warnings = [];
            if (pub.error) return { error: pub.error };
            if (pub.note) warnings.push(pub.note);
            const missing = describePublicationIssue(pub);
            if (missing) return { error: missing };
            const records = r.authors.map(a => ({
                NAME: a.name, Designation: a.designation, PaperTitle: r.title,
                DOI: pub.doi, Volume: pub.vol, Issue: pub.issue, Year: pub.year, Month: pub.month,
                __minFitLevel: r.minFit,
            }));
            if (r.fitLevels) records.forEach((rec, i) => {
                const f = r.fitLevels[i];
                if (f && Number.isInteger(f.level)) { rec.__fitLevel = f.level; rec.__fitFits = f.fits !== false; }
            });
            // numbered the way pasted data is: the author list line keeps each name's number
            state.records = records;
            state.authorListLine = r.authors.map((a, i) => a.name + (i + 1)).join(', ');
            records.forEach((rec, i) => { if (!textFits(rec)) warnings.push(tooLongMessage('Certificate #' + (i + 1))); });
            records.forEach((rec, i) => {
                if (!rec.Designation) warnings.push('No designation for certificate #' + (i + 1) + '.');
            });
            if (!r.title) warnings.push('No paper title.');
            return { warnings, pub };
        })()
    `;
    const prepared = vm.runInContext(prepare, vm.createContext ? ctx : ctx);
    if (prepared.error) fail(prepared.error);

    const toBuffer = async blob => Buffer.from(await blob.arrayBuffer());
    const outDir = path.resolve(req.outDir);
    fs.mkdirSync(outDir, { recursive: true });
    const written = [];
    const save = async (stem, blob, kind, index) => {
        const file = stem + '.docx';
        fs.writeFileSync(path.join(outDir, file), await toBuffer(blob));
        written.push({ file, kind, index });
    };

    const stem = vm.runInContext('buildCombinedFilename()', ctx);
    const records = vm.runInContext('state.records', ctx);

    // 1. all in one
    await save(stem, vm.runInContext('buildCombinedDocxBlob(state.records)', ctx), 'all', 0);
    // 2. one per author
    for (let i = 0; i < records.length; i++) {
        ctx.__i = i;
        const blob = vm.runInContext('renderDocxBlobForRecord(state.records[__i])', ctx);
        await save(vm.runInContext('buildIndividualFilename(__i)', ctx), blob, 'single', i + 1);
    }
    // 3. group (everybody on one certificate)
    const group = vm.runInContext('buildGroupRecord()', ctx);
    group.__minFitLevel = ctx.__req.minFit;
    const gf = ctx.__req.fitLevels && ctx.__req.fitLevels[records.length];
    if (gf && Number.isInteger(gf.level)) { group.__fitLevel = gf.level; group.__fitFits = gf.fits !== false; }
    ctx.__group = group;
    const warnings = prepared.warnings.slice();
    if (!vm.runInContext('textFits(__group)', ctx)) warnings.push(vm.runInContext("tooLongMessage('Group certificate')", ctx));
    await save(stem + GROUP_SUFFIX, vm.runInContext('buildCombinedDocxBlob([__group])', ctx), 'group', 0);

    const fits = records.map(r => { ctx.__r = r; return vm.runInContext('certificateFit(__r)', ctx).level; });
    fits.push(vm.runInContext('certificateFit(__group)', ctx).level);
    process.stdout.write(JSON.stringify({
        files: written,
        fitLevels: fits,
        warnings,
        publication: { doi: prepared.pub.doi, volume: prepared.pub.vol, issue: prepared.pub.issue,
                       month: prepared.pub.month, year: prepared.pub.year, fileNumber: prepared.pub.fileNumber },
    }) + '\n');
}

main().catch(err => fail(String(err && err.stack || err)));
