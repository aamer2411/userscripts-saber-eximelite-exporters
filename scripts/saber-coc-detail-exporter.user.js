// ==UserScript==
// @name         SABER Conformity Certificates Detail Exporter
// @namespace    saber-exporter
// @version      0.5
// @description  Export every Conformity Certificate with its full detail-page fields (Certification Body + 14 fields) to Excel. Unlocks locked certificates first via the review page's Skip (default) or a 5-star Send Review.
// @author       Mohammed Aamer
// @license      MIT
// @homepageURL  https://github.com/aamer2411/userscripts-saber-eximelite-exporters
// @supportURL   https://github.com/aamer2411/userscripts-saber-eximelite-exporters/issues
// @updateURL    https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/saber-coc-detail-exporter.user.js
// @downloadURL  https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/saber-coc-detail-exporter.user.js
// @match        https://saber.sa/ContactOfficer/*
// @match        https://saber.sa/contactofficer/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

//
// HOW IT WORKS (quick map for future edits)
// -----------------------------------------
// The Conformity Certificates list (/ContactOfficer/HRRequest/RegulatedCertificatesSearch)
// shows 7 rows per page. Each row's Options cell is one of:
//   * UNLOCKED: a link to .../ViewHRCoCApprovedProductCertificateReport?query=<token>
//     (the eye icon) - the detail page is readable straight away.
//   * LOCKED: a link to /ContactOfficer/Reviews?query=<token> (the star icon) - the
//     detail page's own token does not exist yet; the site only creates it after the
//     mandatory Certification-Body review is submitted OR skipped.
//
// So to read every certificate the script:
//   1. Pages the list, collecting each cert's number + its detail URL (if unlocked)
//      or its review URL (if locked).
//   2. For each locked cert: GETs the review page, re-submits its exact form (with the
//      page's own antiforgery token) to /Reviews/SkipEvaluation (UNLOCK_MODE='skip',
//      posts no rating) or /Reviews/SubmitEvaluation with both answers = 5
//      (UNLOCK_MODE='submit5'). This is the same action a human does by hand.
//   3. Re-pages the list once so every row now carries a detail URL.
//   4. Fetches each detail page and parses the 15 columns.
//   5. Writes one sheet, 15 columns, one row per certificate.
//
// READ/WRITE: step 2 is the ONLY write. Everything else is GET. It never touches any
// other action link (Cancel, Delete, ...). If a page redirects to /Account/Login the
// run stops and asks you to log in again.
//
// Common changes:
//   * skip vs 5-star rating       -> UNLOCK_MODE.
//   * columns / order / headers   -> COLUMNS.
//   * parallel detail fetches     -> DETAIL_CONCURRENCY (site is slow; keep low).
//

(function () {
    'use strict';

    // 'skip'    -> press the review page's Skip button (no rating posted). Default.
    // 'submit5' -> send the review with the top score (5) on both questions.
    const UNLOCK_MODE = 'skip';

    const LIST_PATH = '/ContactOfficer/HRRequest/RegulatedCertificatesSearch';
    const DETAIL_PATH = '/ContactOfficer/HRRequest/ViewHRCoCApprovedProductCertificateReport';
    const MAX_PAGES = 100;           // safety cap for list paging
    const DETAIL_CONCURRENCY = 3;    // detail pages fetched at once (site is slow)

    // Output columns, in order. Each label is matched against the detail page.
    // 'certBody' is the "We ( ... )" office name; the rest are label/value pairs.
    const COLUMNS = [
        { header: 'Certification Body', kind: 'certBody' },
        { header: 'Certificate Number', label: 'Certificate Number' },
        { header: 'Issue Date', label: 'Issue Date' },
        { header: 'Expire Date', label: 'Expire Date' },
        { header: 'Certificate Type', label: 'Certificate Type' },
        { header: 'Commercial Registration No', label: 'Commercial Registration No' },
        { header: 'Establishment Address', label: 'Establishment Address' },
        { header: 'Model Name', label: 'Model Name' },
        { header: 'Trade Mark', label: 'Trade Mark' },
        { header: 'Product Name', label: 'Product Name' },
        { header: 'Product Description', label: 'Product Description' },
        { header: 'Country of origin', label: 'Country of origin' },
        { header: 'HS Code', label: 'HS Code' },
        { header: 'Technical Regulation', label: 'Technical Regulation' },
        { header: 'Manufacturer name', label: 'Manufacturer name' },
    ];

    // ---------- fetching ----------

    async function fetchDoc(url) {
        const res = await fetch(url, { credentials: 'include' });
        if (/\/Account\/Login/i.test(res.url)) throw new Error('Logged out. Please log in again and retry.');
        if (!res.ok) throw new Error('HTTP ' + res.status + ' for ' + url);
        return new DOMParser().parseFromString(await res.text(), 'text/html');
    }

    // ---------- list paging ----------

    // Returns [{ certNo, detailUrl|null, reviewUrl|null }] for every row across all pages.
    async function collectListRows(setStatus) {
        const rows = [];
        const seen = new Set();
        for (let page = 1; page <= MAX_PAGES; page++) {
            const url = new URL(LIST_PATH, location.origin);
            url.searchParams.set('PageNumber', page);
            const doc = await fetchDoc(url.href);
            const trs = [...doc.querySelectorAll('table tbody tr')].filter(tr => tr.querySelector('td'));
            if (!trs.length) break;
            let newOnPage = 0;
            for (const tr of trs) {
                const certNo = (tr.querySelector('td') || {}).textContent;
                const no = (certNo || '').trim();
                if (!no || seen.has(no)) continue;
                seen.add(no);
                newOnPage++;
                const view = tr.querySelector('a[href*="ViewHRCoC"]');
                const review = tr.querySelector('a[href*="Reviews"]');
                rows.push({
                    certNo: no,
                    detailUrl: view ? view.href : null,
                    reviewUrl: review ? review.href : null,
                });
            }
            setStatus('Listing ' + rows.length);
            if (!newOnPage) break; // guard against a repeating last page
        }
        return rows;
    }

    // ---------- unlocking ----------

    // GET the review page and re-submit its own form (with its antiforgery token) to
    // Skip or Submit. Returns nothing; throws on HTTP error or logout.
    async function unlock(reviewUrl) {
        const doc = await fetchDoc(reviewUrl);
        const form = doc.querySelector('form[action*="SubmitEvaluation"]');
        if (!form) throw new Error('review form not found on ' + reviewUrl);
        const body = new FormData();
        for (const el of form.querySelectorAll('input, select, textarea')) {
            if (!el.name || el.type === 'file') continue;
            if (el.type === 'radio') {
                // only the chosen score is sent; 5 = top rating, for submit5 mode only
                if (UNLOCK_MODE === 'submit5' && el.value === '5') body.append(el.name, '5');
                continue;
            }
            body.append(el.name, el.value);
        }
        const action = UNLOCK_MODE === 'submit5'
            ? '/ContactOfficer/Reviews/SubmitEvaluation'
            : '/ContactOfficer/Reviews/SkipEvaluation';
        const res = await fetch(new URL(action, location.origin).href, {
            method: 'POST', credentials: 'include', body,
        });
        if (/\/Account\/Login/i.test(res.url)) throw new Error('Logged out. Please log in again and retry.');
        if (!res.ok) throw new Error('unlock HTTP ' + res.status);
    }

    // ---------- detail parsing ----------

    function clean(s) {
        return (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
    }

    // Value for a label: the first non-empty element sibling after the leaf element
    // whose text is exactly the label, walking up to 3 ancestors to find it.
    function valueForLabel(doc, label) {
        const want = label.toLowerCase();
        const el = [...doc.querySelectorAll('*')].find(
            e => e.children.length === 0 && e.textContent.trim().toLowerCase() === want);
        if (!el) return '';
        let n = el;
        for (let up = 0; up < 3; up++) {
            let s = n.nextElementSibling;
            while (s && !s.textContent.trim()) s = s.nextElementSibling;
            if (s) return clean(s.textContent);
            n = n.parentElement;
            if (!n) break;
        }
        return '';
    }

    // "We ( ....OFFICE NAME.... ) office number ..." -> OFFICE NAME (dots are padding).
    function certificationBody(doc) {
        const m = doc.body.textContent.match(/We\s*\(([\s\S]*?)\)/);
        if (!m) return '';
        return m[1].replace(/^[\s.]+|[\s.]+$/g, '').replace(/\s+/g, ' ').trim();
    }

    function parseDetail(doc) {
        return COLUMNS.map(c => c.kind === 'certBody'
            ? certificationBody(doc)
            : valueForLabel(doc, c.label));
    }

    // Fetch detail pages in small waves.
    async function fetchDetails(rows, setStatus) {
        const out = new Array(rows.length);
        let done = 0;
        for (let i = 0; i < rows.length; i += DETAIL_CONCURRENCY) {
            const slice = rows.slice(i, i + DETAIL_CONCURRENCY);
            await Promise.all(slice.map(async (row, k) => {
                const url = new URL(DETAIL_PATH, location.origin);
                url.search = new URL(row.detailUrl).search; // reuse the query token as-is
                const doc = await fetchDoc(url.href);
                out[i + k] = parseDetail(doc);
                setStatus('Details (' + (++done) + '/' + rows.length + ')');
            }));
        }
        return out;
    }

    // ---------- export ----------

    // SheetJS community can't write Excel Tables or frozen panes, so write the workbook,
    // then splice the table part, its relationship, exact column widths and a frozen
    // header row into the .xlsx zip. Ported from eximelite-exporter.user.js so the two
    // stay mergeable.
    function xmlEsc(s) {
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
    }

    function tableXml(id, name, ref, headers, stripes) {
        const cols = headers.map((h, i) => '<tableColumn id="' + (i + 1) + '" name="' + xmlEsc(h) + '"/>').join('');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="' + id +
            '" name="' + name + '" displayName="' + name + '" ref="' + ref + '" totalsRowShown="0">' +
            '<autoFilter ref="' + ref + '"/>' +
            '<tableColumns count="' + headers.length + '">' + cols + '</tableColumns>' +
            '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="' + (stripes ? 1 : 0) + '" showColumnStripes="0"/>' +
            '</table>';
    }

    // Freeze the header row by injecting a frozen pane into the worksheet's <sheetView>.
    function freezeTopRow(xml) {
        const pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
        if (/<sheetView\b[^>]*\/>/.test(xml)) {
            return xml.replace(/<sheetView\b([^>]*)\/>/, '<sheetView$1>' + pane + '</sheetView>');
        }
        return xml.replace(/<sheetView\b([^>]*)>[\s\S]*?<\/sheetView>/, '<sheetView$1>' + pane + '</sheetView>');
    }

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = filename;
        (document.body || document.documentElement).appendChild(a);
        a.click();
        setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1500);
    }

    // Write wb, then splice a real Table (with autofilter), a frozen header row and exact
    // column widths into its single worksheet (sheet1), and download it.
    async function writeWorkbookWithTables(wb, spec, filename) {
        const b64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
        const zip = await JSZip.loadAsync(b64, { base64: true });

        const ct = await zip.file('[Content_Types].xml').async('string');
        zip.file('[Content_Types].xml', ct.replace('</Types>',
            '<Override PartName="/xl/tables/table' + spec.tableId +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));

        zip.file('xl/tables/table' + spec.tableId + '.xml', tableXml(spec.tableId, spec.name, spec.ref, spec.headers, spec.stripes));
        zip.file('xl/worksheets/_rels/sheet' + spec.sheetIndex + '.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + spec.tableId + '.xml"/>' +
            '</Relationships>');

        const sn = 'xl/worksheets/sheet' + spec.sheetIndex + '.xml';
        let xml = await zip.file(sn).async('string');
        xml = freezeTopRow(xml);
        const cols = '<cols>' + spec.widths.map((w, i) =>
            '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>';
        xml = xml.replace('<sheetData', cols + '<sheetData');
        xml = xml.replace('</worksheet>', '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>');
        zip.file(sn, xml);

        const blob = await zip.generateAsync({
            type: 'blob', compression: 'DEFLATE',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        triggerDownload(blob, filename);
    }

    async function runExport(setStatus) {
        let rows = await collectListRows(setStatus);
        if (!rows.length) throw new Error('No certificates found.');

        const locked = rows.filter(r => !r.detailUrl && r.reviewUrl);
        for (let i = 0; i < locked.length; i++) {
            setStatus('Unlocking (' + (i + 1) + '/' + locked.length + ')');
            await unlock(locked[i].reviewUrl);
        }
        if (locked.length) {
            rows = await collectListRows(setStatus); // refresh: every row should now be unlocked
        }

        const ready = rows.filter(r => r.detailUrl);
        const data = await fetchDetails(ready, setStatus);

        // Sort by Expire Date ascending (dd/mm/yyyy); blanks/bad dates sink to the end.
        const expIdx = COLUMNS.findIndex(c => c.header === 'Expire Date');
        const asDate = s => {
            const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec((s || '').trim());
            return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : Infinity;
        };
        data.sort((a, b) => asDate(a[expIdx]) - asDate(b[expIdx]));

        // Prepend a 1-based Sl No column.
        const headers = ['Sl No', ...COLUMNS.map(c => c.header)];
        const body = data.map((r, i) => [i + 1, ...r]);
        const ws = XLSX.utils.aoa_to_sheet([headers, ...body]);
        const widths = headers.map((h, i) =>
            Math.min(60, Math.max(h.length, ...body.map(r => String(r[i] || '').length)) + 2));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Conformity Certificates');

        const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        await writeWorkbookWithTables(wb, {
            tableId: 1, sheetIndex: 1, name: 'Conformity_Certificates',
            ref: ws['!ref'], headers, widths, stripes: true,
        }, 'SABER_Conformity_Certificates_' + stamp + '.xlsx');

        const skipped = rows.length - ready.length;
        return { total: ready.length, skipped };
    }

    // ---------- button ----------
    // Same look, size and placement as the EximElite FAB (circular, top-right, 44px,
    // #217346) so the two can be merged later. Shows on every ContactOfficer page; the
    // export always runs over the whole certificates list regardless of the page it is
    // clicked from. The tooltip (title) carries progress while it runs.

    const SVG_DOWNLOAD = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M5 21h14"/></svg>';
    const SVG_CHECK = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>';

    function injectStyle() {
        if (document.getElementById('saber-coc-style')) return;
        const s = document.createElement('style');
        s.id = 'saber-coc-style';
        s.textContent =
            '#saber-coc-fab{position:fixed;top:16px;right:16px;width:44px;height:44px;border-radius:50%;' +
            'background:#217346;border:none;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.35);' +
            'z-index:2147483647;display:flex;align-items:center;justify-content:center;padding:0;transition:background .2s;}' +
            '#saber-coc-fab:hover{background:#1a5c38;}' +
            '#saber-coc-fab:disabled{cursor:default;}' +
            // progress pill: hides behind the button, slides out to the left while running
            '#saber-coc-pill{position:fixed;top:16px;right:16px;width:44px;height:44px;border-radius:22px;' +
            'background:#217346;color:#fff;font:600 13px/44px system-ui,Segoe UI,Arial,sans-serif;' +
            'white-space:nowrap;overflow:hidden;box-sizing:border-box;z-index:2147483646;pointer-events:none;' +
            'transition:width .35s ease,background .3s;}' +
            '#saber-coc-pill span{display:inline-block;padding:0 10px 0 18px;opacity:0;transition:opacity .2s;}' +
            '#saber-coc-pill.open{box-shadow:0 2px 8px rgba(0,0,0,.35);}' +
            '#saber-coc-pill.open span{opacity:1;}' +
            '#saber-coc-pill.err{background:#c0392b;}' +
            '.saber-coc-spin{width:20px;height:20px;border:3px solid rgba(255,255,255,.35);' +
            'border-top-color:#fff;border-radius:50%;animation:saber-coc-rot .8s linear infinite;}' +
            '@keyframes saber-coc-rot{to{transform:rotate(360deg);}}';
        (document.head || document.documentElement).appendChild(s);
    }

    function makeFab(onClick) {
        injectStyle();
        let btn = document.getElementById('saber-coc-fab');
        if (!btn) {
            btn = document.createElement('button');
            btn.id = 'saber-coc-fab';
            document.body.appendChild(btn);
        }
        btn.onclick = onClick;
        let pill = document.getElementById('saber-coc-pill');
        if (!pill) {
            pill = document.createElement('div');
            pill.id = 'saber-coc-pill';
            pill.innerHTML = '<span></span>';
            document.body.appendChild(pill);
        }
        const text = pill.firstChild;
        // Show text in the pill; its width follows the text (44px is the button under its right end).
        const show = t => { text.textContent = t; pill.classList.add('open'); pill.style.width = (44 + text.scrollWidth) + 'px'; };
        const hide = () => { pill.classList.remove('open'); pill.style.width = '44px'; setTimeout(() => pill.classList.remove('err'), 350); };
        let timer = null;
        const later = (fn, ms) => { clearTimeout(timer); timer = setTimeout(fn, ms); };
        const api = {
            el: btn,
            idle() { btn.disabled = false; btn.title = 'Export certificates to Excel'; btn.innerHTML = SVG_DOWNLOAD; },
            busy(n) {
                clearTimeout(timer);
                btn.disabled = true; btn.removeAttribute('title');
                if (!btn.querySelector('.saber-coc-spin')) btn.innerHTML = '<span class="saber-coc-spin"></span>';
                pill.classList.remove('err');
                show(n || 'Loading...');
            },
            done(n) {
                btn.disabled = false; btn.removeAttribute('title'); btn.innerHTML = SVG_CHECK;
                show('Done: ' + n + ' rows');
                later(() => { hide(); later(() => api.idle(), 400); }, 2000);
            },
            error(msg) {
                api.idle();
                pill.classList.add('err');
                show('Export failed');
                later(() => { if (msg) alert(msg); later(hide, 1500); }, 450); // let the red pill paint before alert blocks
            },
        };
        api.idle();
        return api;
    }

    let fab = null;
    async function onClick() {
        if (!fab) return;
        fab.busy('Starting...');
        try {
            const { total } = await runExport(s => fab.busy(s));
            fab.done(total);
        } catch (e) {
            console.error('[SABER CoC exporter]', e);
            fab.error('SABER export failed: ' + e.message);
        }
    }

    // Re-ensure the button every 2 s so it survives the site's in-page navigation.
    function bootFab() {
        const ensure = () => { if (document.body && !document.getElementById('saber-coc-fab')) fab = makeFab(onClick); };
        if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ensure);
        else ensure();
        setInterval(ensure, 2000);
    }

    // Console hook for testing without the button.
    window.saberCoc = { collectListRows, unlock, fetchDetails, parseDetail, runExport, COLUMNS };

    bootFab();
})();
