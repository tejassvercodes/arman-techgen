// The formatter, running in the browser: Python (Pyodide) in a background thread, so the page never freezes.
// Messages in:  { id, op, ... }   Messages out: { id, ok, result | error }
'use strict';
importScripts('https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js');

const PY_FILES = ['ijarcce_format.py', 'core.py', 'pagecount.py', 'engine.py'];
const TEMPLATES = ['CCE', 'IARJSET', 'IJIREEICE', 'IMRJR'];
let py = null, ready = null;

async function bytes(url) {
    const r = await fetch(url, { cache: 'force-cache' });
    if (!r.ok) throw new Error('Could not load ' + url);
    return new Uint8Array(await r.arrayBuffer());
}

function start(base) {
    if (ready) return ready;
    ready = (async () => {
        py = await loadPyodide();
        await py.loadPackage('lxml');
        py.FS.mkdirTree('/app/formatter');
        py.FS.mkdirTree('/app/templates');
        py.FS.mkdirTree('/work');
        for (const f of PY_FILES) py.FS.writeFile('/app/formatter/' + f, await bytes(base + 'py/' + f));
        py.FS.writeFile('/app/formatter/tnr_metrics.json', await bytes(base + 'py/tnr_metrics.json'));
        for (const t of TEMPLATES) py.FS.writeFile('/app/templates/' + t + '_production.docx', await bytes(base + 'templates/' + t + '_production.docx'));
        py.runPython('import sys; sys.path.insert(0, "/app/formatter"); import engine');
        return true;
    })();
    return ready;
}

const call = (fn, ...args) => {
    const pf = py.globals.get('engine')[fn] || py.runPython('engine.' + fn);
    return pf(...args);
};

self.onmessage = async (ev) => {
    const m = ev.data;
    try {
        if (m.op === 'start') { await start(m.base); self.postMessage({ id: m.id, ok: true, result: true }); return; }
        await ready;
        let result;
        if (m.op === 'meta') result = JSON.parse(py.runPython('engine.meta()'));
        else if (m.op === 'parse_doi') { py.globals.set('_a', m.doi); result = JSON.parse(py.runPython('engine.parse_doi(_a)')); }
        else if (m.op === 'clean_settings') { py.globals.set('_a', JSON.stringify(m.settings)); result = JSON.parse(py.runPython('engine.clean_settings(_a)')); }
        else if (m.op === 'chain') { py.globals.set('_a', JSON.stringify(m.settings)); py.globals.set('_b', JSON.stringify(m.papers)); result = JSON.parse(py.runPython('engine.chain(_a, _b)')); }
        else if (m.op === 'format') {
            py.FS.writeFile('/work/in.docx', new Uint8Array(m.bytes));
            py.globals.set('_a', JSON.stringify(m.job));
            const info = JSON.parse(py.runPython('engine.format_paper(_a)'));
            const out = py.FS.readFile('/work/out.docx');
            info.docx = out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
            self.postMessage({ id: m.id, ok: true, result: info }, [info.docx]);
            return;
        } else throw new Error('Unknown operation ' + m.op);
        self.postMessage({ id: m.id, ok: true, result });
    } catch (e) {
        self.postMessage({ id: m.id, ok: false, error: String(e && e.message || e).split('\n').slice(-3).join(' ') });
    }
};
