// The "server" of the browser edition.  The page talks to it exactly as it talks to the desktop server
// (api(path, options) and link(path)), but everything happens here, in this browser tab: the formatter runs
// in a worker (Pyodide), the certificates in the hidden certificate maker.  Nothing is uploaded anywhere and
// nothing is kept once the tab is closed.
(function () {
    'use strict';
    const ROOT = new URL('../', document.currentScript.src).href;     // the site's root folder
    const VERSION = '1.0';
    const CERT_KEYS = { ijarcce: 'CCE', iarjset: 'SET', ijireeice: 'ICE', imrjr: 'RJR' };

    /* ------------------------------------------------------------------ the formatter (worker) */
    let worker = null, seq = 0, ready = null;
    const pending = {};
    function startWorker() {
        if (ready) return ready;
        worker = new Worker(ROOT + 'app/engine.worker.js');
        worker.onmessage = (e) => {
            const p = pending[e.data.id];
            if (!p) return;
            delete pending[e.data.id];
            e.data.ok ? p.res(e.data.result) : p.rej(new Error(e.data.error));
        };
        worker.onerror = (e) => { const msg = 'The formatter stopped: ' + (e.message || 'unknown error'); Object.keys(pending).forEach(id => { const p = pending[id]; delete pending[id]; p.rej(new Error(msg)); }); worker = null; ready = null; };
        ready = call('start', { base: ROOT });
        ready.catch(() => { ready = null; });
        return ready;
    }
    // Every call has a time limit; a stuck formatter is stopped and started again so one bad file cannot freeze the page.
    const LIMITS = { start: 180000, format: 150000 };
    function restartWorker() {
        try { worker && worker.terminate(); } catch (e) { /* already gone */ }
        worker = null; ready = null;
        Object.keys(pending).forEach(id => { const p = pending[id]; delete pending[id]; p.rej(new Error('The formatter was restarted. Please try again.')); });
    }
    function call(op, extra, transfer) {
        return new Promise((res, rej) => {
            const id = ++seq;
            const timer = setTimeout(() => {
                if (!pending[id]) return;
                delete pending[id];
                restartWorker();
                rej(new Error(op === 'format' ? 'This paper took too long to format and was stopped.' : 'The formatter did not answer in time. Reload the page.'));
            }, LIMITS[op] || 60000);
            pending[id] = { res: (v) => { clearTimeout(timer); res(v); }, rej: (e) => { clearTimeout(timer); rej(e); } };
            worker.postMessage(Object.assign({ id, op }, extra), transfer || []);
        });
    }
    const engine = async (op, extra, transfer) => { await startWorker(); return call(op, extra, transfer); };

    /* ------------------------------------------------------------------ state */
    let B = null;                                   // the one issue of this session
    const uid = () => Math.random().toString(16).slice(2, 14).padEnd(12, '0');
    const blobs = {};                               // paper id -> { sources, docx: ArrayBuffer, certs: { stem: { docx, pdf } } }
    const urlCache = new Map();

    function defaults() {
        return { journal: 'ijarcce', mode: 'with', first_doi: '', volume: '', issue: '', month: '',
                 year: String(new Date().getFullYear()), first_paper: '1', first_page: '1' };
    }
    function newBatch() { B = { id: uid(), created: Date.now(), settings: defaults(), papers: [] }; return B; }
    const settingsOk = (s) => ['volume', 'issue', 'month', 'year', 'first_page', 'first_paper'].every(k => String(s[k] || '').trim());
    const effPages = (p) => p.pages_manual || p.pages_est || null;
    const secure = (s) => s.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^[._]+|[._]+$/g, '') || 'paper';
    const niceName = (p) => secure(p.name.replace(/\.[^.]+$/, ''));

    function label(s) {
        const names = window.__journals || {};
        if (!(s.volume && s.issue)) return 'New issue (details not filled in)';
        return (names[s.journal] || s.journal) + ' Vol. ' + s.volume + ', Issue ' + s.issue + ', ' + s.month + ' ' + s.year;
    }
    function summary(b) { return { id: b.id, label: label(b.settings), count: b.papers.length, created: b.created }; }

    async function chainFor(b) {
        const papers = b.papers.map(p => ({ id: p.id, doi_manual: p.doi_manual || '', start_manual: p.start_manual || 0, pages: p.status === 'done' ? effPages(p) : null }));
        return engine('chain', { settings: b.settings, papers });
    }
    const snapOf = (b, p, t) => JSON.stringify([b.settings.journal, b.settings.volume, b.settings.issue, b.settings.month, b.settings.year, t.doi, t.start, p.rev || 0]);
    const certSnapOf = (b, p, t) => { const cd = p.cert_data || {}; return JSON.stringify([b.settings.journal, b.settings.volume, b.settings.issue, b.settings.year, t.doi, t.paper_no, cd.title, cd.authors]); };

    let lastTargets = {};
    function view(b) {
        const t = settingsOk(b.settings) ? lastTargets : {};
        const papers = b.papers.map(p => {
            const tt = t[p.id] || {};
            const pages = p.status === 'done' ? effPages(p) : null;
            return {
                id: p.id, name: p.name, status: p.status, stage: p.stage || '', error: p.error || null, title: p.title || '',
                flags: p.flags || [], fixes: p.fixes || [], found: p.found || {}, pages,
                pages_est: p.pages_est || null, pages_manual: p.pages_manual || 0, unsure: p.unsure || [],
                paper_no: tt.paper_no, doi: tt.doi || '', doi_manual: !!p.doi_manual, start: tt.start,
                start_manual: !!p.start_manual, end: (pages && tt.start) ? tt.start + pages - 1 : null,
                certs: p.certs || null, cert_data: p.cert_data || null, cert_manual: !!p.cert_manual,
            };
        });
        const busy = settingsOk(b.settings) && b.papers.some(p => p.status === 'queued' || p.status === 'working' ||
            (p.status === 'done' && (p.certs == null || p.certs.status === 'working')));
        return { id: b.id, label: label(b.settings), settings: b.settings, ready: settingsOk(b.settings), busy, papers };
    }

    /* ------------------------------------------------------------------ the pipeline: format -> chain pages -> certificates */
    let running = false, again = false;
    function pump() {
        if (running) { again = true; return; }
        running = true;
        (async () => {
            try { do { again = false; await pipeline(); } while (again); }
            catch (e) { console.error(e); }
            finally { running = false; }
        })();
    }
    const alive = (b, p) => B === b && b.papers.includes(p);

    async function pipeline() {
        for (let round = 0; round < 6; round++) {
            const b = B;
            if (!b || !settingsOk(b.settings)) return;
            let t = lastTargets = await chainFor(b);
            const todo = b.papers.filter(p => p.status !== 'error' && (p.status === 'queued' || p.snap !== snapOf(b, p, t[p.id])));
            for (const p of todo) { if (alive(b, p)) await formatOne(b, p, t[p.id]); }
            t = lastTargets = await chainFor(b);
            const certTodo = b.papers.filter(p => p.status === 'done' && effPages(p) &&
                (p.certs == null || (p.certs.status !== 'working' && p.cert_snap !== certSnapOf(b, p, t[p.id]))));
            for (const p of certTodo) { if (alive(b, p)) await makeCerts(b, p, t[p.id]); }
            if (!todo.length && !certTodo.length) return;
        }
    }

    async function formatOne(b, p, t) {
        p.status = 'working'; p.stage = 'Formatting'; p.error = null;
        try {
            const src = blobs[p.id].source.slice(0);
            const s = b.settings;
            const job = { volume: s.volume, issue: s.issue, month: s.month, year: s.year, doi: t.doi, start: t.start, journal: s.journal };
            const r = await engine('format', { bytes: src, job }, [src]);
            if (!alive(b, p)) return;
            blobs[p.id].docx = r.docx;
            blobs[p.id].report = r.report;
            Object.assign(p, { status: 'done', stage: '', title: r.title, flags: r.flags, fixes: r.fixes, found: r.found,
                               pages_est: r.pages, unsure: r.unsure, snap: snapOf(b, p, t), error: null, cert_data_auto: r.cert_data });
            if (!p.cert_manual) p.cert_data = r.cert_data;
        } catch (e) {
            if (!alive(b, p)) return;
            Object.assign(p, { status: 'error', stage: '', error: String(e.message || e).slice(0, 300), flags: [], fixes: [], found: {}, pages_est: null, snap: null });
        }
    }

    async function makeCerts(b, p, t) {
        const snap = certSnapOf(b, p, t);
        p.certs = { status: 'working', files: [], warnings: [], error: null };
        try {
            const cd = p.cert_data || {};
            if (!(cd.authors || []).length) throw new Error('No authors were found in the paper, so there is nothing to put on the certificates. Add them with Check / edit names.');
            const s = b.settings;
            const req = { journal: s.journal, title: cd.title || '', authors: cd.authors, doi: t.doi,
                          volume: s.volume, issue: s.issue, year: s.year, fileNumber: String(t.paper_no).padStart(2, '0') };
            let r;
            try { r = await window.Certs.make(req); }
            catch (first) { await new Promise(res => setTimeout(res, 800)); r = await window.Certs.make(req); }   // one quiet retry
            if (!alive(b, p)) return;
            const store = blobs[p.id].certs = {};
            r.files.forEach(f => { store[f.stem] = { docx: f.docx, pdf: f.pdf }; });
            p.certs = { status: 'ok', files: r.files.map(f => ({ stem: f.stem, kind: f.kind, index: f.index, pdf: !!f.pdf })),
                        warnings: [...(cd.warnings || []), ...r.warnings], error: null };
            p.cert_snap = snap;
        } catch (e) {
            if (!alive(b, p)) return;
            p.certs = { status: 'error', files: [], warnings: [], error: String(e.message || e).slice(0, 400) };
            p.cert_snap = snap;
        }
    }

    /* ------------------------------------------------------------------ the routes the page uses */
    const bad = (msg) => { throw new Error(msg); };
    const getPaper = (pid) => B.papers.find(p => p.id === pid) || bad('That paper is no longer here.');
    const markAll = (fn) => B.papers.forEach(p => { if (p.status !== 'error') fn(p); });

    async function route(method, path, query, body, isForm) {
        let m;
        if (method === 'GET' && path === '/api/meta') {
            const meta = await engine('meta');
            window.__journals = meta.journals;
            const logos = {}; Object.keys(CERT_KEYS).forEach(k => { logos[k] = '/logos/' + CERT_KEYS[k] + '.jpg'; });
            return Object.assign(meta, { version: VERSION, logos, health: { word: true, node: true, engine: 'browser' } });
        }
        if (method === 'GET' && path === '/api/parse-doi') return engine('parse_doi', { doi: query.get('doi') || '' });
        if (method === 'GET' && path === '/api/init') { if (!B) newBatch(); return { batches: [summary(B)], current: B.id }; }
        if (method === 'POST' && path === '/api/batches') { newBatch(); return Object.assign(view(B), { batches: [summary(B)] }); }
        if ((m = path.match(/^\/api\/batches\/([^/]+)(?:\/(.*))?$/))) {
            const rest = m[2] || '';
            if (!B || B.id !== m[1]) bad('This issue is no longer here. Reload the page.');
            if (method === 'GET' && rest === '') return view(B);
            if (method === 'DELETE' && rest === '') { Object.keys(blobs).forEach(k => delete blobs[k]); B = null; lastTargets = {}; return { batches: [] }; }
            if (method === 'PUT' && rest === 'settings') {
                const r = await engine('clean_settings', { settings: body });
                if (r.error) bad(r.error);
                const s = r.ok, old = B.settings;
                B.settings = s;
                if (old.journal !== s.journal) B.papers.forEach(p => { p.doi_manual = ''; });
                if (['journal', 'volume', 'issue', 'month', 'year'].some(k => old[k] !== s[k])) markAll(p => { p.status = 'queued'; p.rev = (p.rev || 0) + 1; });
                if (settingsOk(s)) lastTargets = await chainFor(B);
                pump();
                return Object.assign(view(B), { batches: [summary(B)] });
            }
            if (method === 'POST' && rest === 'papers') {
                if (!settingsOk(B.settings)) bad('Fill in the issue details first (volume, issue, month, year, first paper number, first page).');
                const files = body.getAll('papers');
                if (!files.length) bad('Choose at least one Word file.');
                const rejected = [], old = [], broken = [];
                if (files.length + B.papers.length > 80) bad('That is more than 80 papers in one issue. Do the rest in the next issue.');
                for (const f of files) {
                    const ext = (f.name.match(/\.[^.]+$/) || [''])[0].toLowerCase();
                    if (ext !== '.docx') { (ext === '.doc' || ext === '.rtf' || ext === '.odt' ? old : rejected).push(f.name); continue; }
                    if (f.size > 60 * 1024 * 1024) { broken.push(f.name + ' (over 60 MB)'); continue; }
                    const head = new Uint8Array(await f.slice(0, 4).arrayBuffer());
                    if (f.size < 1000 || head[0] !== 0x50 || head[1] !== 0x4b) { broken.push(f.name + ' (not a valid Word file)'); continue; }
                    const p = { id: uid(), name: f.name, status: 'queued', stage: '', rev: 0, flags: [], fixes: [], found: {}, title: '', pages_est: null };
                    blobs[p.id] = { source: await f.arrayBuffer(), certs: {} };
                    B.papers.push(p);
                }
                if (settingsOk(B.settings)) lastTargets = await chainFor(B);
                pump();
                const resp = view(B);
                const warn = [];
                if (old.length) warn.push('Open ' + old.join(', ') + ' in Word and save as .docx first: this version reads .docx files only.');
                if (rejected.length) warn.push('Skipped (not a Word file): ' + rejected.join(', '));
                if (broken.length) warn.push('Skipped (cannot be read): ' + broken.join(', ') + '. Open it in Word and save it again.');
                if (warn.length) resp.warning = warn.join(' ');
                return resp;
            }
            if ((m = rest.match(/^papers\/([^/]+)(?:\/(.*))?$/))) {
                const pid = m[1], sub = m[2] || '';
                if (method === 'PATCH' && sub === '') {
                    const p = getPaper(pid);
                    if ('doi' in body) {
                        let doi = String(body.doi || '').trim();
                        const prefix = (await engine('meta')).prefixes[B.settings.journal];
                        if (doi && !doi.toLowerCase().startsWith('10.')) doi = prefix + doi;
                        p.doi_manual = doi && doi !== prefix ? doi : '';
                    }
                    if ('start' in body) {
                        const st = String(body.start || '').trim();
                        if (st && !/^\d+$/.test(st)) bad('Start page must be a number.');
                        p.start_manual = st ? parseInt(st, 10) : 0;
                    }
                    if ('pages' in body) {
                        const n = String(body.pages || '').trim();
                        if (n && !/^[1-9]\d{0,2}$/.test(n)) bad('The page count must be a whole number.');
                        p.pages_manual = n ? parseInt(n, 10) : 0;
                    }
                    lastTargets = await chainFor(B); pump();
                    return view(B);
                }
                if (method === 'DELETE' && sub === '') {
                    B.papers = B.papers.filter(p => p.id !== pid); delete blobs[pid];
                    lastTargets = settingsOk(B.settings) ? await chainFor(B) : {}; pump();
                    return view(B);
                }
                if (method === 'POST' && sub === 'move') {
                    const i = B.papers.findIndex(p => p.id === pid), j = i + (body && body.dir === 'up' ? -1 : 1);
                    if (i < 0) bad('That paper is no longer here.');
                    if (j >= 0 && j < B.papers.length) { const t = B.papers[i]; B.papers[i] = B.papers[j]; B.papers[j] = t; }
                    lastTargets = await chainFor(B); pump();
                    return view(B);
                }
                if (method === 'POST' && sub === 'redo') { const p = getPaper(pid); p.status = 'queued'; p.rev = (p.rev || 0) + 1; p.pages_manual = 0; pump(); return view(B); }
                if (method === 'PUT' && sub === 'certdata') {
                    const p = getPaper(pid);
                    const authors = (body.authors || []).map(a => ({ name: String(a.name || '').trim(), designation: String(a.designation || '').trim() })).filter(a => a.name);
                    if (!authors.length) bad('Keep at least one author with a name.');
                    p.cert_data = { title: String(body.title || '').trim(), authors, warnings: [] };
                    p.cert_manual = true; p.certs = null; pump();
                    return view(B);
                }
                if (method === 'POST' && sub === 'certs/redo') {
                    const p = getPaper(pid);
                    if (body && body.reset && p.cert_data_auto) { p.cert_data = p.cert_data_auto; p.cert_manual = false; }
                    p.certs = null; pump();
                    return view(B);
                }
            }
        }
        bad('Not available in this version: ' + method + ' ' + path);
    }

    async function api(path, opts) {
        opts = opts || {};
        const method = (opts.method || 'GET').toUpperCase();
        const u = new URL(path, 'http://local');
        let body = opts.body;
        const isForm = typeof FormData !== 'undefined' && body instanceof FormData;
        if (!isForm && typeof body === 'string' && body) { try { body = JSON.parse(body); } catch (e) { /* leave as text */ } }
        return route(method, u.pathname, u.searchParams, isForm ? body : (body || {}), isForm);
    }

    /* ------------------------------------------------------------------ downloads */
    function blobURL(key, make) {
        if (!urlCache.has(key)) urlCache.set(key, URL.createObjectURL(make()));
        return urlCache.get(key);
    }
    const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', pdf: 'application/pdf', zip: 'application/zip' };
    const toBlob = (v, type) => (v instanceof Blob ? v : new Blob([v], { type }));

    function parseFile(path) {
        const u = new URL(path, 'http://local');
        const m = u.pathname.match(/^\/files\/[^/]+\/([^/]+)\/([^/]+)$/);
        return m ? { pid: m[1], kind: m[2], q: u.searchParams } : null;
    }
    function link(path) {
        if (path.startsWith('/logos/') || path.startsWith('/static/')) return ROOT + path.slice(1);
        const f = parseFile(path);
        if (!f) return path;
        const p = B && B.papers.find(x => x.id === f.pid), st = blobs[f.pid];
        if (!p || !st) return '#';
        const rev = p.snap + '|' + (p.cert_snap || '');
        if (f.kind === 'docx' && st.docx) return blobURL('d|' + f.pid + '|' + p.snap, () => toBlob(st.docx, MIME.docx));
        if (f.kind === 'cert') {
            const e = st.certs[f.q.get('name')], fmt = f.q.get('fmt') === 'pdf' ? 'pdf' : 'docx';
            return e && e[fmt] ? blobURL('c|' + f.pid + '|' + f.q.get('name') + '|' + fmt + '|' + p.cert_snap, () => toBlob(e[fmt], MIME[fmt])) : '#';
        }
        return '#';
    }
    function filename(path) {
        const f = parseFile(path);
        if (!f || !B) return '';
        const p = B.papers.find(x => x.id === f.pid);
        if (!p) return '';
        if (f.kind === 'docx') return niceName(p) + '.docx';
        if (f.kind === 'cert') return f.q.get('name') + '.' + (f.q.get('fmt') === 'pdf' ? 'pdf' : 'docx');
        return '';
    }
    // One ZIP for one paper: the formatted paper, its report and every certificate (Word and PDF).
    async function bundle(pid) {
        const p = B && B.papers.find(x => x.id === pid), st = blobs[pid];
        if (!p || !st || !st.docx) throw new Error('This paper is not ready.');
        const z = new window.PizZip(), base = niceName(p);
        z.file(base + '_formatted.docx', st.docx);
        z.file(base + '_report.txt', st.report || '');
        for (const [stem, e] of Object.entries(st.certs)) {
            if (e.docx) z.file('certificates/' + stem + '.docx', await e.docx.arrayBuffer());
            if (e.pdf) z.file('certificates/' + stem + '.pdf', await e.pdf.arrayBuffer());
        }
        return { url: URL.createObjectURL(z.generate({ type: 'blob', mimeType: MIME.zip })), name: base + '_paper_and_certificates.zip' };
    }

    window.LocalApi = {
        api, link, filename, bundle, start: startWorker,
        docxBytes: (pid) => (blobs[pid] && blobs[pid].docx) || null,
        certPdf: (pid, stem) => (blobs[pid] && blobs[pid].certs[stem] && blobs[pid].certs[stem].pdf) || null,
        certDocx: (pid, stem) => (blobs[pid] && blobs[pid].certs[stem] && blobs[pid].certs[stem].docx) || null,
    };
})();
