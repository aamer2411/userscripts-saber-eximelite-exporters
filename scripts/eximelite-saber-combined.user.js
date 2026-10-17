// ==UserScript==
// @name         EximElite + SABER Combined Exporter
// @namespace    eximelite-exporter
// @version      1.5
// @description  One script, two tools by site: EximElite EDA-EX bulk exporter on powerbi/eximelite, and the SABER Conformity Certificates detail exporter on saber.sa.
// @author       Mohammed Aamer
// @license      MIT
// @homepageURL  https://github.com/aamer2411/userscripts-saber-eximelite-exporters
// @supportURL   https://github.com/aamer2411/userscripts-saber-eximelite-exporters/issues
// @updateURL    https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/eximelite-saber-combined.user.js
// @downloadURL  https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/eximelite-saber-combined.user.js
// @match        https://app.powerbi.com/view*
// @match        https://eximelite.com/*
// @match        https://saber.sa/ContactOfficer/*
// @match        https://saber.sa/contactofficer/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js
// @grant        none
// @run-at       document-start
// ==/UserScript==

//
// GENERATED FILE: do not edit by hand. Built by tools/build-combined.js from
// scripts/eximelite-exporter.user.js and scripts/saber-coc-detail-exporter.user.js.
// Two independent tools, selected by hostname:
//   * saber.sa           -> runSaber()  (Conformity Certificates detail exporter)
//   * powerbi / eximelite -> runExim()  (EDA-EX bulk exporter, runs in both frames)
// Each tool is the original script wrapped verbatim in its own function, so nothing is
// shared except the @require libraries (xlsx/jszip in both; pdf.js used by SABER only). Edit
// a tool by editing its original file and re-running the build.
//

(function () {
    'use strict';
    if (/(^|\.)saber\.sa$/i.test(location.hostname)) runSaber();
    else runExim();

    //
    // HOW IT WORKS (quick map for future edits)
    // -----------------------------------------
    // The eximelite page embeds a Publish-to-Web Power BI report in an app.powerbi.com
    // iframe. This script runs in BOTH frames (see the two @match lines):
    //   * app.powerbi.com frame  -> initEngine(): hooks fetch/XHR to capture the
    //     report's own querydata POST (auth key + current slicer filters), then on
    //     demand replays it to pull every row-level column, pages the result, decodes
    //     Power BI's DSR payload, and writes the .xlsx.
    //   * eximelite.com frame    -> initHostButton(): draws the sticky green button on
    //     the outer page and relays a start message to the engine (so the button stays
    //     put while the outer page scrolls). Standalone PBI gets a local button too.
    //
    // The single fact table EE_EDA_EX_Monthly_Data holds all detail columns at
    // shipment grain, so one query gets everything - no per-row drill-through needed.
    //
    // Common changes:
    //   * add/remove/reorder export columns -> edit COLUMNS (Shipments/KPI sheets) or
    //                                           RAW_EXTRA_COLUMNS (Raw Data sheet only).
    //   * filename slots                     -> buildFilename() / extractFilters().
    //   * output sheets                      -> exportWorkbook().
    //   * duplicate handling / row count     -> COUNT_PROP / COUNT_FN + exportWorkbook().
    //
    
    function runExim() {
        'use strict';
    
        const FACT = 'EE_EDA_EX_Monthly_Data';
        const QUERY_RE = /analysis\.windows\.net\/public\/reports\/querydata/;
        const PAGE_ROWS = 30000;   // rows requested per query window
        const MAX_PAGES = 500;     // safety cap (=> up to 15M rows)
    
        // Columns to export, in output order. 'col' is a plain column; buildSelect also
        // supports date-hierarchy kinds 'year'/'month'/'day'. Year/Month come from the
        // Date Hierarchy on '01 Shipping Bill Date'. The six address/IEC/contact fields
        // are the exporter's, keyed to the exporter name.
        const COLUMNS = [
            { header: 'Year', kind: 'year' },
            { header: 'Month', kind: 'month' },
            { header: 'Day', kind: 'day' },
            { header: 'Exporter Name', kind: 'col', prop: '07 Expoerter Names' },
            { header: 'HS Code', kind: 'col', prop: '04 HS Code' },
            { header: 'Consignee Name', kind: 'col', prop: '18 Consignee Names' },
            { header: 'IEC No', kind: 'col', prop: '06 IEC No' },
            { header: 'Exporter Pin Code', kind: 'col', prop: '09 Expoerter Pin Code' },
            { header: 'Exporter Address', kind: 'col', prop: '08 Expoerter Address' },
            { header: 'Exporter District', kind: 'col', prop: '10 Expoerter District' },
            { header: 'Exporter State', kind: 'col', prop: '11 Expoerter State' },
            { header: 'Exporter Phone No', kind: 'col', prop: '12 Expoerter Phone No' },
            { header: 'Exporter Email ID', kind: 'col', prop: '13 Expoerter Email ID' },
            { header: 'Product Description', kind: 'col', prop: '05 Product Description' },
        ];
    
        // Extra columns, queried alongside COLUMNS but only surfaced on the "Raw Data"
        // sheet (Shipments/KPI sheets stay on the original 13). Property strings pulled
        // from the report's own conceptualschema (exact spellings/double-spaces kept
        // as-is, e.g. '17 Indian  Port Name', '25.1  Currency').
        const RAW_EXTRA_COLUMNS = [
            { header: 'Indian Port Code', kind: 'col', prop: '14 Indian Port Code' },
            { header: 'Indian Port Name', kind: 'col', prop: '17 Indian  Port Name' },
            { header: 'Consignee Address', kind: 'col', prop: '19 Consignee Address' },
            { header: 'Foreign Port', kind: 'col', prop: '20 Foreign Port' },
            { header: 'Foreign Country', kind: 'col', prop: '21 Foreign Country' },
            { header: 'Mode', kind: 'col', prop: '35 Mode' },
            { header: 'Trade', kind: 'col', prop: '32 Trade' },
            { header: 'Quantity', kind: 'col', prop: '22 Quantity' },
            { header: 'Unit Quantity', kind: 'col', prop: '23 Unit Quantity' },
            { header: 'Unit Value (Invoice)', kind: 'col', prop: '24 Unit Value as Per Invoice' },
            { header: 'Currency', kind: 'col', prop: '25.1  Currency' },
            { header: 'Total Value (FC)', kind: 'col', prop: '26 Total Value in FC' },
            { header: 'Unit Rate (USD)', kind: 'col', prop: '27 Unit Rate USD' },
            { header: 'Exchange Rate', kind: 'col', prop: '28 Exchange Rate' },
            { header: 'Total Value (USD)', kind: 'col', prop: '29 Total Value in USD' },
            { header: 'Unit Rate (INR)', kind: 'col', prop: '30 Unit Rate in INR' },
            { header: 'FOB Value (INR)', kind: 'col', prop: '31.1 FOB Value in INR' },
            { header: 'Drawback', kind: 'col', prop: '34 Drawback' },
        ];
        const ALL_COLUMNS = COLUMNS.concat(RAW_EXTRA_COLUMNS);
    
        // The "N shipment" card counts physical rows, including records that are
        // byte-identical across every column. A grouped query merges those, so the
        // export used to fall short. Fix: append a CountNonNull aggregation over an
        // always-present column; each distinct row then carries how many physical
        // copies it stands for, and we repeat it that many times on export.
        const COUNT_PROP = '04 HS Code'; // always populated for a shipment
        const COUNT_FN = 5;              // Power BI QueryAggregateFunction.CountNonNull
    
        // ---- Near-duplicate name matching (KPI sheets only) ----
        // The same company/consignee is registered under slightly different spellings
        // (spacing, plurals, punctuation, abbreviations). Conservative signature: lowercase,
        // drop punctuation, expand a few known abbreviations, singularise simple plurals,
        // then remove all spaces so "Hyper Markets" == "Hypermarket". Genuinely different
        // wordings (e.g. two distinct trade names sharing one word) are NOT merged - by
        // design; that keeps false merges out at the cost of missing the loosest matches.
        const NAME_ABBR = { trdg: 'trading', pvt: 'private', ltd: 'limited', co: 'company', est: 'establishment', intl: 'international', corp: 'corporation', mfg: 'manufacturing' };
        function nameSig(s) {
            let t = String(s == null ? '' : s).toLowerCase().replace(/&/g, ' and ');
            t = t.replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
            if (!t) return '';
            return t.split(' ')
                .map(w => NAME_ABBR[w] || w)
                .map(w => (w.length > 3 && w.endsWith('s') && !w.endsWith('ss')) ? w.slice(0, -1) : w)
                .join('');
        }
    
        const IN_PBI = /(^|\.)app\.powerbi\.com$/.test(location.hostname);
    
        // =====================================================================
        //  ENGINE (runs inside the app.powerbi.com frame)
        // =====================================================================
    
        // Latest richest fact query seen; carries current auth + slicer filters.
        let template = null; // {url, key, parsed, selectLen}
        // Most recent fact query seen, regardless of size - lets the field-dump debug
        // helper (eximeliteListFields) reflect whatever visual you're currently looking
        // at, even when its Select is smaller than the running "richest" template.
        let lastQuery = null; // {url, parsed, selectLen}
    
        function considerCapture(url, headers, body) {
            if (!QUERY_RE.test(url) || !body) return;
            let parsed;
            try { parsed = JSON.parse(typeof body === 'string' ? body : ''); } catch (e) { return; }
            const cmd = parsed && parsed.queries && parsed.queries[0] &&
                parsed.queries[0].Query && parsed.queries[0].Query.Commands &&
                parsed.queries[0].Query.Commands[0] &&
                parsed.queries[0].Query.Commands[0].SemanticQueryDataShapeCommand;
            if (!cmd) return;
            const from = (cmd.Query && cmd.Query.From) || [];
            if (!from.some(f => f.Entity === FACT)) return;
            const selectLen = (cmd.Query.Select || []).length;
            lastQuery = { url, parsed, selectLen };
            // Keep the richest Select (that is the main detail table, whose Where is
            // the full slicer filter set). ">=" so the LATEST such query wins and the
            // template tracks the current filter instead of freezing on the first one.
            if (!template || selectLen >= template.selectLen) {
                const key = getHeader(headers, 'x-powerbi-resourcekey') ||
                    (template && template.key);
                template = { url, key, parsed, selectLen };
            }
        }
    
        function getHeader(headers, name) {
            if (!headers) return null;
            name = name.toLowerCase();
            if (typeof headers.forEach === 'function' && !Array.isArray(headers)) {
                let found = null;
                headers.forEach((v, k) => { if (String(k).toLowerCase() === name) found = v; });
                return found;
            }
            for (const k in headers) if (k.toLowerCase() === name) return headers[k];
            return null;
        }
    
        const origFetch = window.fetch;
    
        function installHooks() {
            window.fetch = function (input, init) {
                try {
                    const url = (typeof input === 'string') ? input : (input && input.url);
                    const headers = (init && init.headers) || (input && input.headers);
                    const body = init && init.body;
                    if (url) considerCapture(url, headers, body);
                } catch (e) { /* ignore */ }
                return origFetch.apply(this, arguments);
            };
            const XP = XMLHttpRequest.prototype;
            const origOpen = XP.open, origSend = XP.send, origSet = XP.setRequestHeader;
            XP.open = function (method, url) { this.__u = url; this.__h = {}; return origOpen.apply(this, arguments); };
            XP.setRequestHeader = function (k, v) { if (this.__h) this.__h[k] = v; return origSet.apply(this, arguments); };
            XP.send = function (body) {
                try { if (this.__u) considerCapture(this.__u, this.__h, body); } catch (e) { /* ignore */ }
                return origSend.apply(this, arguments);
            };
        }
    
        function guid() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
                const r = Math.random() * 16 | 0;
                return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
            });
        }
    
        function factAliasFrom(cmd) {
            const f = (cmd.Query.From || []).find(x => x.Entity === FACT);
            return f ? f.Name : 'e';
        }
    
        function buildSelect(alias) {
            return ALL_COLUMNS.map(c => {
                if (c.kind === 'col') {
                    return {
                        Column: { Expression: { SourceRef: { Source: alias } }, Property: c.prop },
                        Name: FACT + '.' + c.prop,
                        NativeReferenceName: c.header,
                    };
                }
                const level = c.kind.charAt(0).toUpperCase() + c.kind.slice(1); // Year/Month/Day
                return {
                    HierarchyLevel: {
                        Expression: {
                            Hierarchy: {
                                Expression: {
                                    PropertyVariationSource: {
                                        Expression: { SourceRef: { Source: alias } },
                                        Name: 'Variation',
                                        Property: '01 Shipping Bill Date',
                                    },
                                },
                                Hierarchy: 'Date Hierarchy',
                            },
                        },
                        Level: level,
                    },
                    Name: FACT + '.01 Shipping Bill Date.Variation.Date Hierarchy.' + level,
                    NativeReferenceName: c.header,
                };
            });
        }
    
        function buildQueryBody(restartTokens) {
            const body = JSON.parse(JSON.stringify(template.parsed)); // deep clone
            const cmd = body.queries[0].Query.Commands[0].SemanticQueryDataShapeCommand;
            const alias = factAliasFrom(cmd);
    
            const sel = buildSelect(alias);
            // Extra projection: physical-row count per distinct group (see COUNT_PROP).
            sel.push({
                Aggregation: {
                    Expression: { Column: { Expression: { SourceRef: { Source: alias } }, Property: COUNT_PROP } },
                    Function: COUNT_FN,
                },
                Name: 'EE_RowCount',
                NativeReferenceName: '__count',
            });
            cmd.Query.Select = sel;
            // Order by a selected column for a deterministic sequence (needed for paging).
            cmd.Query.OrderBy = [{
                Direction: 1,
                Expression: { Column: { Expression: { SourceRef: { Source: alias } }, Property: COUNT_PROP } },
            }];
    
            const win = { Count: PAGE_ROWS };
            if (restartTokens) win.RestartTokens = restartTokens;
            cmd.Binding = {
                Primary: { Groupings: [{ Projections: sel.map((_, i) => i) }] },
                DataReduction: { DataVolume: 3, Primary: { Window: win } },
                Version: 1,
            };
            // cmd.Query.Where is left untouched: it carries the slicer filters.
            if ('ExecutionMetricsKind' in cmd) delete cmd.ExecutionMetricsKind;
            return body;
        }
    
        // ---- DSR decode (validated offline + live against real captures) ----
        // Power BI's response is dictionary + delta encoded: DS[0].ValueDicts holds
        // string dictionaries; rows live in PH[0].DM0 where S=schema, C=values,
        // R=repeat-from-previous bitmask, Ø=null bitmask, DN=ValueDict index per column;
        // RT=restart tokens, IC:false means more pages remain.
        function bitSet(mask, i) {
            // avoid 32-bit bitwise overflow for >31 columns
            return Math.floor(mask / Math.pow(2, i)) % 2 === 1;
        }
    
        function decodeDSR(obj) {
            // A malformed query (e.g. a wrong column/hierarchy name) still returns HTTP 200
            // but with an error payload instead of results. Surface Power BI's actual
            // message - it names the offending field - rather than crashing later reading a
            // missing property.
            const res = obj && obj.results && obj.results[0] && obj.results[0].result;
            if (!res || !res.data || !res.data.dsr || !res.data.dsr.DS) {
                const err = (obj && obj.error) ||
                    (obj && obj.results && obj.results[0] && obj.results[0].error);
                const msg = err ? (err.message || err.code || JSON.stringify(err))
                    : JSON.stringify(obj || {}).slice(0, 500);
                try { console.error('[eximelite] query rejected:', obj); } catch (e) { /* ignore */ }
                throw new Error('Power BI rejected the query: ' + msg);
            }
            const ds = res.data.dsr.DS[0];
            const vd = ds.ValueDicts || {};
            const dm = (ds.PH && ds.PH[0] && ds.PH[0].DM0) || [];
            let schema = null, prev = null;
            const rows = [];
            for (const r of dm) {
                if (r.S) schema = r.S;
                const ncol = schema.length;
                const rbits = r.R || 0;
                const nbits = r['Ø'] || 0;
                const cvals = r.C || [];
                let cur = 0;
                const out = new Array(ncol);
                for (let i = 0; i < ncol; i++) {
                    if (rbits && bitSet(rbits, i)) out[i] = prev[i];
                    else if (nbits && bitSet(nbits, i)) out[i] = null;
                    else out[i] = cvals[cur++];
                }
                prev = out;
                const resolved = new Array(ncol);
                for (let i = 0; i < ncol; i++) {
                    const v = out[i];
                    const dn = schema[i].DN;
                    resolved[i] = (dn != null && typeof v === 'number') ? vd[dn][v] : v;
                }
                rows.push(resolved);
            }
            return { rows, rt: ds.RT || null, ic: ds.IC };
        }
    
        async function runQuery(body) {
            const resp = await origFetch(template.url, {
                method: 'POST',
                headers: {
                    'content-type': 'application/json;charset=UTF-8',
                    'x-powerbi-resourcekey': template.key,
                    'activityid': guid(),
                    'requestid': guid(),
                },
                body: JSON.stringify(body),
            });
            if (!resp.ok) throw new Error('Query HTTP ' + resp.status + ': ' + (await resp.text()).slice(0, 300));
            return resp.json();
        }
    
        async function collectAllRows(onCount) {
            const all = [];
            let restart = null;
            let prevKey = null;
            for (let page = 0; page < MAX_PAGES; page++) {
                const obj = await runQuery(buildQueryBody(restart));
                const { rows, rt, ic } = decodeDSR(obj);
                for (const r of rows) all.push(r);
                if (onCount) onCount(all.length);
                if (ic !== false || !rt) break; // ic:false + rt => more pages
                const key = JSON.stringify(rt);
                if (!rows.length || key === prevKey) break; // no progress: stop
                prevKey = key;
                restart = rt;
            }
            return all;
        }
    
        // ---- filename from the current filter (kept in template.Where) ----
        function cleanLiteral(v) {
            if (typeof v !== 'string') return String(v);
            let s = v;
            if (s.length >= 2 && s[0] === "'" && s[s.length - 1] === "'") s = s.slice(1, -1);
            // Power BI numeric literals carry a type suffix, e.g. 2024L or 12.5D. Strip it.
            else if (/^-?\d+(\.\d+)?[LD]$/.test(s)) s = s.slice(0, -1);
            return s;
        }
    
        // Identify what a filter expression targets: a plain column property, or a
        // date-hierarchy level (Year/Month/Day). Year is normally a hierarchy slicer,
        // which is why a plain-column-only reader used to miss it and show All-Years.
        function exprKey(expr) {
            if (!expr) return null;
            if (expr.Column && expr.Column.Property) return expr.Column.Property;
            if (expr.HierarchyLevel && expr.HierarchyLevel.Level) return expr.HierarchyLevel.Level;
            return null;
        }
    
        function extractFilters() {
            const cmd = template.parsed.queries[0].Query.Commands[0].SemanticQueryDataShapeCommand;
            const where = (cmd.Query && cmd.Query.Where) || [];
            const byProp = {};
            for (const cond of where) {
                const inCond = cond.Condition && cond.Condition.In; // only positive slicer selections
                if (!inCond) continue;
                const key = exprKey(inCond.Expressions && inCond.Expressions[0]);
                if (!key) continue;
                const vals = (inCond.Values || []).map(v => cleanLiteral(v[0] && v[0].Literal && v[0].Literal.Value));
                byProp[key] = (byProp[key] || []).concat(vals);
            }
            return byProp;
        }
    
        // The year is not a slicer: each report is a per-year dataset whose URL looks
        // like ".../eda-ex-2025/". Pull it from the host page URL (passed in from the
        // outer page) or, failing that, the iframe referrer.
        function reportYearFrom(url) {
            const m = url && /\b(19|20)\d{2}\b/.exec(url);
            return m ? m[0] : null;
        }
    
        function buildFilename(pageUrl) {
            const f = extractFilters();
            const part = (arr, multiLabel, allLabel) =>
                (!arr || !arr.length) ? allLabel : (arr.length > 1 ? multiLabel : arr[0]);
            const year = reportYearFrom(pageUrl) || reportYearFrom(document.referrer) ||
                part(f['Year'], 'Multiple-Years', 'All-Years');
            const month = part(f['Month'], 'Multiple-Months', 'All-Months');
            const chapter = part(f['Chapter 1'] || f['02 Chapter'], 'Multiple-Chapters', 'All-Chapters');
            const hsn = part(f['HS Code 2'] || f['04 HS Code'], 'Multiple-HSN', 'All-HSN');
            const country = part(f['21 Foreign Country'], 'Multiple-Countries', 'All-Countries');
            let name = [year, month, chapter, hsn, country].join('_');
            name = name.replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
            return name + '.xlsx';
        }
    
        // ---- Excel Table + width/freeze helpers (used by writeWorkbookWithTables) ----
        function colLetter(n) { // 1-based column index -> A, B, ... Z, AA, ...
            let s = '';
            while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
            return s;
        }
    
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
    
        // Freeze the header row: inject a frozen pane into the worksheet's <sheetView>
        // (SheetJS community doesn't write panes). Handles both the open/close and the
        // self-closing sheetView forms SheetJS emits.
        function freezeTopRow(xml) {
            const pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
            if (/<sheetView\b[^>]*\/>/.test(xml)) {
                return xml.replace(/<sheetView\b([^>]*)\/>/, '<sheetView$1>' + pane + '</sheetView>');
            }
            return xml.replace(/<sheetView\b([^>]*)>[\s\S]*?<\/sheetView>/, '<sheetView$1>' + pane + '</sheetView>');
        }
    
        // ---- Pivot table helpers (hand-authored OOXML: pivot cache + pivot table
        // definition). SheetJS community can't write native PivotTables at all, so these
        // work around it the same way tableXml() above works around missing Table support:
        // write the base workbook, then splice these parts into the zip. Structure verified
        // by unzipping a real Excel-built pivot workbook (pivot.xlsx) and reading its raw
        // XML - not guessed. Only two fixed shapes are needed here: a 2-visible-level pivot
        // (outer field + one nested level, KPI_1) and a 1-visible-level pivot (outer field
        // only, KPI_2), both with everything below the visible depth collapsed by default
        // (Excel's "Collapse Entire Field"). The collapse rule that matches a real
        // Excel-built file: the DEEPEST VISIBLE field gets sd="0" on every item (that is
        // what hides its children); every field above it is left expanded (no sd); fields
        // already hidden below it are inert (never reached by rowItems either way).
        function buildCacheDict(rows, colIdx) {
            const order = [];
            const indexOf = new Map();
            for (const r of rows) {
                const v = String(r[colIdx]);
                if (!indexOf.has(v)) { indexOf.set(v, order.length); order.push(v); }
            }
            return { order, indexOf };
        }
    
        function cacheFieldXml(name, dict) {
            const items = dict.order.map(v => '<s v="' + xmlEsc(v) + '"/>').join('');
            return '<cacheField name="' + xmlEsc(name) + '" numFmtId="0"><sharedItems count="' +
                dict.order.length + '">' + items + '</sharedItems></cacheField>';
        }
    
        function numericCacheFieldXml(name, values) {
            let min = Infinity, max = -Infinity;
            for (const v of values) { const n = Number(v) || 0; if (n < min) min = n; if (n > max) max = n; }
            if (!isFinite(min)) { min = 0; max = 0; }
            return '<cacheField name="' + xmlEsc(name) + '" numFmtId="0"><sharedItems containsSemiMixedTypes="0" ' +
                'containsString="0" containsNumber="1" containsInteger="1" minValue="' + min + '" maxValue="' + max + '"/></cacheField>';
        }
    
        function pivotItemsXml(dict, collapse) {
            const items = dict.order.map((_, i) => '<item x="' + i + '"' + (collapse ? ' sd="0"' : '') + '/>').join('');
            return '<items count="' + (dict.order.length + 1) + '">' + items +
                '<item t="default"' + (collapse ? ' sd="0"' : '') + '/></items>';
        }
    
        // rowItems for a pivot collapsed to exactly ONE visible level (outer field only).
        function rowItems1Level(names, dict) {
            const parts = names.map(name => {
                const i = dict.indexOf.get(name);
                return '<i><x' + (i ? ' v="' + i + '"' : '') + '/></i>';
            });
            parts.push('<i t="grand"><x/></i>');
            return '<rowItems count="' + parts.length + '">' + parts.join('') + '</rowItems>';
        }
    
        // rowItems for a pivot collapsed to exactly TWO visible levels (outer + one nested
        // level). groups: [{ outer: name, inner: [name, ...] }, ...] in display order.
        function rowItems2Level(groups, outerDict, innerDict) {
            const parts = [];
            for (const g of groups) {
                const oi = outerDict.indexOf.get(g.outer);
                parts.push('<i><x' + (oi ? ' v="' + oi + '"' : '') + '/></i>');
                for (const name of g.inner) {
                    const ii = innerDict.indexOf.get(name);
                    parts.push('<i r="1"><x' + (ii ? ' v="' + ii + '"' : '') + '/></i>');
                }
            }
            parts.push('<i t="grand"><x/></i>');
            return '<rowItems count="' + parts.length + '">' + parts.join('') + '</rowItems>';
        }
    
        // Assembles the pivotCacheDefinition/Records XML and the pivotTableDefinition XML
        // for one pivot table. `rows`/`headers` = the full underlying record set (this is
        // the cache - lets Excel show correct data if the user later expands a collapsed
        // level); `rowFieldIdxs` = which header columns form the row hierarchy, in nesting
        // order; `dataFieldIdx` = the Values column; `visibleDepth` = how many leading row
        // fields are shown before the default-collapsed state hides the rest.
        // opts: { dicts (pre-built, shared with the caller's rowItemsXml so indices line
        // up), sourceTableName, cacheId, name, location, rowItemsXml, dataFieldName }.
        function buildPivotParts(rows, headers, rowFieldIdxs, dataFieldIdx, visibleDepth, opts) {
            const dicts = opts.dicts;
            const subtotal = opts.subtotal || 'sum';
            const dataIsNumeric = subtotal === 'sum';
            const cacheFieldsXml = headers.map((h, i) => i === dataFieldIdx && dataIsNumeric
                ? numericCacheFieldXml(h, rows.map(r => r[i]))
                : cacheFieldXml(h, dicts[i])).join('');
            const cacheDefXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" ' +
                'createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="' + rows.length + '">' +
                '<cacheSource type="worksheet"><worksheetSource name="' + opts.sourceTableName + '"/></cacheSource>' +
                '<cacheFields count="' + headers.length + '">' + cacheFieldsXml + '</cacheFields>' +
                '</pivotCacheDefinition>';
    
            const recXml = rows.map(r => {
                const cells = headers.map((h, i) => i === dataFieldIdx && dataIsNumeric
                    ? '<n v="' + (Number(r[i]) || 0) + '"/>'
                    : '<x v="' + dicts[i].indexOf.get(String(r[i])) + '"/>').join('');
                return '<r>' + cells + '</r>';
            }).join('');
            const cacheRecordsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="' +
                rows.length + '">' + recXml + '</pivotCacheRecords>';
    
            const pivotFieldsXml = headers.map((h, i) => {
                if (i === dataFieldIdx) return '<pivotField dataField="1" showAll="0"/>';
                const pos = rowFieldIdxs.indexOf(i);
                if (pos === -1) return '<pivotField showAll="0"/>';
                return '<pivotField axis="axisRow" showAll="0">' + pivotItemsXml(dicts[i], pos >= visibleDepth - 1) + '</pivotField>';
            }).join('');
            const rowFieldsXml = rowFieldIdxs.map(i => '<field x="' + i + '"/>').join('');
            const pivotTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                'name="' + opts.name + '" cacheId="' + opts.cacheId + '" applyNumberFormats="0" applyBorderFormats="0" ' +
                'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0" ' +
                'dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" itemPrintTitles="1" createdVersion="8" ' +
                'indent="0" outline="1" outlineData="1" multipleFieldFilters="0">' +
                '<location ref="' + opts.location + '" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
                '<pivotFields count="' + headers.length + '">' + pivotFieldsXml + '</pivotFields>' +
                '<rowFields count="' + rowFieldIdxs.length + '">' + rowFieldsXml + '</rowFields>' +
                opts.rowItemsXml +
                '<colItems count="1"><i/></colItems>' +
                '<dataFields count="1"><dataField name="' + xmlEsc(opts.dataFieldName) + '" fld="' + dataFieldIdx +
                '" baseField="0" baseItem="0"' + (subtotal !== 'sum' ? ' subtotal="' + subtotal + '"' : '') + '/></dataFields>' +
                '<pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
                '</pivotTableDefinition>';
    
            return { cacheDefXml, cacheRecordsXml, pivotTableXml };
        }
    
        // Splices the pivot cache/table parts for one or more pivots (all living on one
        // sheet) into the zip: parts + their .rels, Content-Types overrides, the sheet's
        // own rels (-> both pivotTables), fixed column widths (autofit-on-update stays off
        // via applyWidthHeightFormats="0" above), and the workbook-level <pivotCaches>
        // wiring (+ workbook.xml.rels) that ties each pivotTable's cacheId to its actual
        // pivotCacheDefinition part.
        async function injectPivotTables(zip, pivotSpec) {
            const ct = await zip.file('[Content_Types].xml').async('string');
            let overrides = '';
            for (const p of pivotSpec.pivots) {
                overrides += '<Override PartName="/xl/pivotCache/pivotCacheDefinition' + p.id +
                    '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>';
                overrides += '<Override PartName="/xl/pivotCache/pivotCacheRecords' + p.id +
                    '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>';
                overrides += '<Override PartName="/xl/pivotTables/pivotTable' + p.id +
                    '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>';
            }
            zip.file('[Content_Types].xml', ct.replace('</Types>', overrides + '</Types>'));
    
            for (const p of pivotSpec.pivots) {
                zip.file('xl/pivotCache/pivotCacheDefinition' + p.id + '.xml', p.cacheDefXml);
                zip.file('xl/pivotCache/pivotCacheRecords' + p.id + '.xml', p.cacheRecordsXml);
                zip.file('xl/pivotCache/_rels/pivotCacheDefinition' + p.id + '.xml.rels',
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords' + p.id + '.xml"/>' +
                    '</Relationships>');
                zip.file('xl/pivotTables/pivotTable' + p.id + '.xml', p.pivotTableXml);
                zip.file('xl/pivotTables/_rels/pivotTable' + p.id + '.xml.rels',
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition' + p.id + '.xml"/>' +
                    '</Relationships>');
            }
    
            const sheetRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                pivotSpec.pivots.map((p, i) =>
                    '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable' + p.id + '.xml"/>'
                ).join('') + '</Relationships>';
            zip.file('xl/worksheets/_rels/sheet' + pivotSpec.sheetIndex + '.xml.rels', sheetRelsXml);
    
            const sn = 'xl/worksheets/sheet' + pivotSpec.sheetIndex + '.xml';
            let xml = await zip.file(sn).async('string');
            const cols = '<cols>' + pivotSpec.widths.map(w =>
                '<col min="' + w.col + '" max="' + w.col + '" width="' + w.width + '" customWidth="1"/>').join('') + '</cols>';
            xml = xml.replace('<sheetData', cols + '<sheetData');
            zip.file(sn, xml);
    
            const relsPath = 'xl/_rels/workbook.xml.rels';
            let relsXml = await zip.file(relsPath).async('string');
            const newRelIds = pivotSpec.pivots.map((p, i) => 'rId9' + (i + 10)); // rId910, rId911 - unused range
            let relAdds = '';
            pivotSpec.pivots.forEach((p, i) => {
                relAdds += '<Relationship Id="' + newRelIds[i] + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition' + p.id + '.xml"/>';
            });
            relsXml = relsXml.replace('</Relationships>', relAdds + '</Relationships>');
            zip.file(relsPath, relsXml);
    
            const wbPath = 'xl/workbook.xml';
            let wbXml = await zip.file(wbPath).async('string');
            const pivotCachesXml = '<pivotCaches>' + pivotSpec.pivots.map((p, i) =>
                '<pivotCache cacheId="' + p.id + '" r:id="' + newRelIds[i] + '"/>'
            ).join('') + '</pivotCaches>';
            if (wbXml.indexOf('<calcPr') !== -1) {
                wbXml = wbXml.replace(/(<calcPr\b[^>]*\/>)/, '$1' + pivotCachesXml);
            } else {
                wbXml = wbXml.replace('</sheets>', '</sheets>' + pivotCachesXml);
            }
            zip.file(wbPath, wbXml);
    
            if (pivotSpec.titleCells || pivotSpec.headerCells) {
                await injectPivotSheetStyling(zip, pivotSpec.sheetIndex, pivotSpec.titleCells || [], pivotSpec.headerCells || []);
            }
        }
    
        // Appends one entry to a styles.xml section (<fonts>/<fills>/<cellXfs>/...) and
        // returns the new zero-based index for it, bumping the section's count="N" to
        // match. SheetJS community can't write cell styles either (same limitation as
        // Tables/pivots), so this is the same splice-after-the-fact workaround.
        function appendStyleEntry(xml, tag, entryXml) {
            const re = new RegExp('<' + tag + ' count="(\\d+)">([\\s\\S]*?)</' + tag + '>');
            const m = re.exec(xml);
            const count = parseInt(m[1], 10);
            const newSection = '<' + tag + ' count="' + (count + 1) + '">' + m[2] + entryXml + '</' + tag + '>';
            return { xml: xml.slice(0, m.index) + newSection + xml.slice(m.index + m[0].length), index: count };
        }
    
        // Sets (or replaces) the style index on one specific cell, identified by its A1
        // reference, wherever it already appears in a worksheet's XML.
        function setCellStyle(xml, cellRef, styleIdx) {
            const re = new RegExp('(<c r="' + cellRef + '"[^>]*)(/?>)');
            return xml.replace(re, (full, pre, close) => pre.replace(/\s+s="\d+"/, '') + ' s="' + styleIdx + '"' + close);
        }
    
        // Dark blue fill + white bold centered text for the pivot sheet's merged title
        // cells, and a plain bold (default color) style for its "Row Labels"/"Sum of ..."
        // header cells. Mirrors the default font (size 12 Calibri) SheetJS itself emits,
        // just adding <b/> (and white color for the title font).
        async function injectPivotSheetStyling(zip, sheetIndex, titleCells, headerCells) {
            let stylesXml = await zip.file('xl/styles.xml').async('string');
            let r = appendStyleEntry(stylesXml, 'fonts',
                '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
            stylesXml = r.xml;
            const titleFontIdx = r.index;
            r = appendStyleEntry(stylesXml, 'fonts',
                '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
            stylesXml = r.xml;
            const boldFontIdx = r.index;
            r = appendStyleEntry(stylesXml, 'fills',
                '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>');
            stylesXml = r.xml;
            const fillIdx = r.index;
            r = appendStyleEntry(stylesXml, 'cellXfs',
                '<xf numFmtId="0" fontId="' + titleFontIdx + '" fillId="' + fillIdx + '" borderId="0" xfId="0" ' +
                'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>');
            stylesXml = r.xml;
            const titleXfIdx = r.index;
            r = appendStyleEntry(stylesXml, 'cellXfs',
                '<xf numFmtId="0" fontId="' + boldFontIdx + '" fillId="0" borderId="0" xfId="0" applyFont="1"/>');
            stylesXml = r.xml;
            const headerXfIdx = r.index;
            zip.file('xl/styles.xml', stylesXml);
    
            const sn = 'xl/worksheets/sheet' + sheetIndex + '.xml';
            let sheetXml = await zip.file(sn).async('string');
            for (const ref of titleCells) sheetXml = setCellStyle(sheetXml, ref, titleXfIdx);
            for (const ref of headerCells) sheetXml = setCellStyle(sheetXml, ref, headerXfIdx);
            zip.file(sn, sheetXml);
        }
    
        function triggerDownload(blob, filename) {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = filename;
            (document.body || document.documentElement).appendChild(a);
            a.click();
            setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1500);
        }
    
        // SheetJS community can't emit table parts or column widths, so write the workbook,
        // then splice into the .xlsx zip: exact column widths, a table XML per sheet, a
        // sheet->table relationship, a [Content_Types] override, a frozen header row, and a
        // <tableParts> ref per worksheet. sheetIndex maps to sheetN.xml (append order).
        // Downloads the result.
        async function writeWorkbookWithTables(wb, specs, filename, pivotSpec) {
            const b64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
            const zip = await JSZip.loadAsync(b64, { base64: true });
    
            const ct = await zip.file('[Content_Types].xml').async('string');
            let overrides = '';
            for (const s of specs) {
                overrides += '<Override PartName="/xl/tables/table' + s.tableId +
                    '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>';
            }
            zip.file('[Content_Types].xml', ct.replace('</Types>', overrides + '</Types>'));
    
            for (const s of specs) {
                zip.file('xl/tables/table' + s.tableId + '.xml', tableXml(s.tableId, s.name, s.ref, s.headers, s.stripes));
                zip.file('xl/worksheets/_rels/sheet' + s.sheetIndex + '.xml.rels',
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + s.tableId + '.xml"/>' +
                    '</Relationships>');
                const sn = 'xl/worksheets/sheet' + s.sheetIndex + '.xml';
                let xml = await zip.file(sn).async('string');
                xml = freezeTopRow(xml);
                // Exact column widths (SheetJS's own auto-fit runs too wide). <cols> must
                // sit immediately before <sheetData>.
                const cols = '<cols>' + s.widths.map((w, i) =>
                    '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>';
                xml = xml.replace('<sheetData', cols + '<sheetData');
                xml = xml.replace('</worksheet>',
                    '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>');
                zip.file(sn, xml);
            }
    
            if (pivotSpec) await injectPivotTables(zip, pivotSpec);
    
            const blob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            });
            triggerDownload(blob, filename);
        }
    
        // Build the workbook from raw grouped rows (each ending with its physical-copy
        // count). Eight sheets, headers on row 1:
        //  - "Shipments": every physical row (each distinct row repeated by its count),
        //    ordered Year(desc) -> Month(desc) -> Day(desc) -> Exporter(volume desc) ->
        //    HS Code(count desc) -> Consignee(count desc). Raw names, unmerged; Day is
        //    Shipments-only.
        //  - "KPI_1_Pivot" (HIDDEN): the 5-level hierarchy (Year -> Month -> Exporter ->
        //    HS Code -> Consignee) laid out as columns, on the MERGED+disambiguated names
        //    (see nameSig/buildNameGroups - an Exporter also needs a matching Exporter Pin
        //    Code, a Consignee also needs a matching Foreign Port, to count as the same
        //    entity). Every column is filled on every row - no print-once blanks - so the
        //    sheet is directly pivotable. Export Count/HS Count are that row's Exporter's/
        //    HS Code's own group total; Consignee Count is the leaf total. Exists only to
        //    feed the "KPI Pivot Tables" left PivotTable's source data.
        //  - "KPI_1" (the visible, user-facing sheet): the same rows as KPI_1_Pivot, but
        //    Year/Month are tied to the Exporter level (both reprint alongside every
        //    distinct Exporter+Pin+Export Count, not just when they themselves change),
        //    then HS Code+HS Count, then Consignee+Foreign Port each print once per group
        //    and blank on repeat rows - a classic report layout.
        //  - "KPI_2_Pivot" (HIDDEN): group Exporter -> Year -> Month -> Day -> Consignee on
        //    the merged+disambiguated names. Exporter ranks by total volume (desc); Year/
        //    Month/Day descending. Each leaf row is one consignee (+ Foreign Port) on one
        //    day, with a comma-separated list of the distinct HS Codes it shipped that day.
        //    Monthly Shipment Count is the row count of consignee entries for that
        //    Exporter/Year/Month (not a sum), filled on every row of that group. Exists
        //    only to feed the "KPI Pivot Tables" right PivotTable's source data.
        //  - "KPI_2" (the visible, user-facing sheet): the same rows as KPI_2_Pivot, but
        //    Exporter+Pin/Year are tied to the Month level (all reprint alongside every
        //    distinct Month+Monthly Shipment Count), then Day, then Consignee+Foreign Port
        //    each print once per group and blank on repeat rows.
        //  - "KPI Pivot Tables": two real Excel PivotTables side by side, hand-authored
        //    OOXML (see buildPivotParts/injectPivotTables) since SheetJS community can't
        //    write native pivots. Left (B3, source KPI_1_Pivot): rows Year -> Month ->
        //    Exporter Name -> Exporter Pin Code -> HS Code -> Consignee Name, collapsed
        //    below Month, Sum of Consignee Count. Right (E3, source KPI_2_Pivot): rows
        //    Exporter Name -> Year -> Month -> Day -> Consignee Name -> Foreign Port,
        //    collapsed below Exporter Name, a Count (not a Sum, so it stays correct at
        //    every drill level) labelled "Shipment Count". Both PivotStyleMedium2,
        //    autofit-on-update off.
        //  - "Name Mapping": one row per name that a KPI sheet collapsed from >1 spelling -
        //    Type (Exporter/Consignee), the shared Exporter Pin Code/Foreign Port that
        //    disambiguates it, the canonical name shown there, and every original spelling
        //    it merged.
        //  - "Raw Data": the same rows and order as Shipments, but with all 18 extra detail
        //    columns appended (port/consignee-address/foreign-port/trade/mode/quantity/
        //    value/currency/rate/drawback fields, see RAW_EXTRA_COLUMNS) - the fullest
        //    single-row-per-shipment dump, deduped across all 31 columns.
        // Near-duplicate Exporter/Consignee spellings are merged conservatively (case,
        // punctuation, spacing, plurals, common abbreviations) for the KPI sheets only,
        // and only within the same Exporter Pin Code / Foreign Port (points 2/3). KPI
        // sheets show ONE canonical name (the most-frequent spelling); the full spelling
        // list lives on the Name Mapping sheet so the KPI cells stay uncluttered.
        // Shipments/KPI_1_Pivot/KPI_1/KPI_2_Pivot/KPI_2/Name Mapping/Raw Data are real
        // Excel Tables (filterable) with exact column widths and per-row stripes; KPI
        // Pivot Tables carries the two PivotTables described above instead.
        async function exportWorkbook(raw, filename) {
            const header = COLUMNS.map(c => c.header);
            const idx = h => COLUMNS.findIndex(c => c.header === h);
            const iYear = idx('Year'), iMonth = idx('Month'), iDay = idx('Day');
            const iExporter = idx('Exporter Name'), iHS = idx('HS Code'), iConsignee = idx('Consignee Name');
            const iExporterPin = idx('Exporter Pin Code');
            const iForeignPort = ALL_COLUMNS.findIndex(c => c.header === 'Foreign Port');
            const dataOf = r => r.slice(0, r.length - 1);
            const countOf = r => Math.max(1, Math.round(Number(r[r.length - 1]) || 1));
            const str = v => (v == null ? '' : String(v));
            const cmpStr = (a, b) => a < b ? -1 : (a > b ? 1 : 0);
    
            // Year sorts by number; Month by calendar position whether Power BI hands it
            // back as a name ("January") or a number ("1").
            const yearNum = v => { const n = parseInt(str(v).replace(/[^\d]/g, ''), 10); return isFinite(n) ? n : 0; };
            const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
            const monthIndex = v => {
                const s = str(v).trim().toLowerCase();
                if (/^\d+$/.test(s)) return parseInt(s, 10);
                return MONTHS[s.slice(0, 3)] || 99;
            };
            const dayNum = v => { const n = parseInt(str(v).replace(/[^\d]/g, ''), 10); return isFinite(n) ? n : 0; };
    
            // Indian ITC-HS codes are 8 digits; Power BI hands them back as numbers, so
            // chapters 01-09 lose their leading zero. Re-pad and keep them as text.
            const HS_WIDTH = 8;
            const fmtHS = v => {
                let s = str(v).trim();
                if (/^\d+$/.test(s) && s.length < HS_WIDTH) s = s.padStart(HS_WIDTH, '0');
                return s;
            };
    
            // Distinct rows with physical counts; normalise HS code up front so both
            // sheets and every aggregation key agree.
            const groups = raw.map(r => {
                const d = dataOf(r);
                d[iHS] = fmtHS(d[iHS]);
                return { d, n: countOf(r) };
            });
    
            // For the KPI sheets: map each raw Exporter/Consignee name + its disambiguator
            // (Exporter Pin Code / Foreign Port) to a single canonical display name = the
            // most-frequent original spelling sharing both the same nameSig AND the same
            // disambiguator value - two otherwise-identical names with different pin codes
            // / foreign ports are kept as distinct entities (points 2/3). `merges` lists
            // the groups that actually collapsed >1 spelling, for the Name Mapping sheet.
            // Shipments is untouched and keeps the raw per-row names.
            function buildNameGroups(field, extraField) {
                const rawCount = new Map(); // sig -> Map(name -> count)
                for (const g of groups) {
                    const nm = str(g.d[field]); if (!nm) continue;
                    const extra = str(g.d[extraField]).trim();
                    const sig = nameSig(nm) + '' + extra;
                    let m = rawCount.get(sig); if (!m) { m = new Map(); rawCount.set(sig, m); }
                    m.set(nm, (m.get(nm) || 0) + g.n);
                }
                const canon = new Map(); // sig -> { name, extra }
                const merges = [];       // { canon, extra, variants: [names] } for sigs with >1 spelling
                for (const [sig, nameCounts] of rawCount) {
                    const arr = [...nameCounts.entries()].map(([nm, c]) => ({ nm, c }));
                    arr.sort((a, b) => (b.c - a.c) || cmpStr(a.nm, b.nm));
                    const extra = sig.slice(sig.indexOf('') + 1);
                    canon.set(sig, { name: arr[0].nm, extra });
                    if (arr.length > 1) merges.push({ canon: arr[0].nm, extra, variants: arr.map(x => x.nm) });
                }
                merges.sort((a, b) => (b.variants.length - a.variants.length) || cmpStr(a.canon, b.canon));
                const sigOf = (nm, extraVal) => nameSig(nm) + '' + str(extraVal).trim();
                return {
                    keyOf: sigOf,
                    dispOf: (nm, extraVal) => { const c = canon.get(sigOf(nm, extraVal)); return c ? c.name : str(nm); },
                    merges,
                };
            }
            const exG = buildNameGroups(iExporter, iExporterPin), coG = buildNameGroups(iConsignee, iForeignPort);
    
            // ---- Aggregate Year -> Month -> Exporter -> HS Code -> Consignee counts
            //      (Shipments sort order; raw names, no Pin/Foreign-Port disambiguation). ----
            const years = new Map();
            for (const g of groups) {
                const y = str(g.d[iYear]), mo = str(g.d[iMonth]), ex = str(g.d[iExporter]), hs = g.d[iHS], co = str(g.d[iConsignee]);
                let Y = years.get(y); if (!Y) { Y = { total: 0, months: new Map() }; years.set(y, Y); }
                Y.total += g.n;
                let M = Y.months.get(mo); if (!M) { M = { total: 0, exps: new Map() }; Y.months.set(mo, M); }
                M.total += g.n;
                let E = M.exps.get(ex); if (!E) { E = { total: 0, hs: new Map() }; M.exps.set(ex, E); }
                E.total += g.n;
                let H = E.hs.get(hs); if (!H) { H = { total: 0, cons: new Map() }; E.hs.set(hs, H); }
                H.total += g.n;
                H.cons.set(co, (H.cons.get(co) || 0) + g.n);
            }
    
            // ---- Same aggregation on the MERGED+disambiguated names, for KPI_1/KPI_1a.
            //      (Shipments above stays on raw names.) An Exporter also needs a matching
            //      Exporter Pin Code, a Consignee also needs a matching Foreign Port, to
            //      count as the same entity (points 2/3). ----
            const cyears = new Map();
            for (const g of groups) {
                const y = str(g.d[iYear]), mo = str(g.d[iMonth]), hs = g.d[iHS];
                const exPin = str(g.d[iExporterPin]).trim();
                const exKey = exG.keyOf(g.d[iExporter], exPin), exName = exG.dispOf(g.d[iExporter], exPin);
                const fPort = str(g.d[iForeignPort]).trim();
                const coKey = coG.keyOf(g.d[iConsignee], fPort), coName = coG.dispOf(g.d[iConsignee], fPort);
                let Y = cyears.get(y); if (!Y) { Y = { total: 0, months: new Map() }; cyears.set(y, Y); }
                Y.total += g.n;
                let M = Y.months.get(mo); if (!M) { M = { total: 0, exps: new Map() }; Y.months.set(mo, M); }
                M.total += g.n;
                let E = M.exps.get(exKey); if (!E) { E = { name: exName, pin: exPin, total: 0, hs: new Map() }; M.exps.set(exKey, E); }
                E.total += g.n;
                let H = E.hs.get(hs); if (!H) { H = { total: 0, cons: new Map() }; E.hs.set(hs, H); }
                H.total += g.n;
                let C = H.cons.get(coKey); if (!C) { C = { name: coName, fPort, count: 0 }; H.cons.set(coKey, C); }
                C.count += g.n;
            }
    
            // ---- Shipments sheet: order by the 5 sort keys, then expand each distinct
            //      row by its physical count. ----
            for (const g of groups) {
                const y = str(g.d[iYear]), mo = str(g.d[iMonth]), ex = str(g.d[iExporter]), hs = g.d[iHS], co = str(g.d[iConsignee]);
                const E = years.get(y).months.get(mo).exps.get(ex), H = E.hs.get(hs);
                g.k = { yi: yearNum(y), mi: monthIndex(mo), di: dayNum(g.d[iDay]), exTot: E.total, ex, hsTot: H.total, hs, coCnt: H.cons.get(co), co };
            }
            const sorted = groups.slice().sort((a, b) =>
                (b.k.yi - a.k.yi) || (b.k.mi - a.k.mi) || (b.k.di - a.k.di) ||
                (b.k.exTot - a.k.exTot) || cmpStr(a.k.ex, b.k.ex) ||
                (b.k.hsTot - a.k.hsTot) || cmpStr(a.k.hs, b.k.hs) ||
                (b.k.coCnt - a.k.coCnt) || cmpStr(a.k.co, b.k.co));
            const data = [];
            const rawData = [];
            for (const g of sorted) {
                for (let k = 0; k < g.n; k++) {
                    data.push(g.d.slice(0, COLUMNS.length));
                    rawData.push(g.d);
                }
            }
    
            // ---- KPI_1 sheet: 5-level hierarchy on the MERGED+disambiguated names. Every
            //      column is filled on every row (no print-once blanks), so the sheet pivots
            //      directly. Export Count/HS Count are that row's Exporter's/HS Code's own
            //      group total (point 5). Striped like Shipments. ----
            const metrics = [];
            // Year->Month subtotals for the "KPI Pivot Tables" sheet's left (KPI_1) pivot,
            // which is collapsed below Month - captured here instead of re-walking cyears.
            const kpi1Grid = [];
            const yearsRanked = [...cyears.entries()].sort((a, b) => yearNum(b[0]) - yearNum(a[0]));
            for (const [y, Y] of yearsRanked) {
                const monthsRanked = [...Y.months.entries()].sort((a, b) => monthIndex(b[0]) - monthIndex(a[0]));
                kpi1Grid.push({ year: y, yearTotal: Y.total, months: monthsRanked.map(([mo, M]) => ({ month: mo, total: M.total })) });
                for (const [mo, M] of monthsRanked) {
                    const expsRanked = [...M.exps.entries()].sort((a, b) => (b[1].total - a[1].total) || cmpStr(a[1].name, b[1].name));
                    for (const [, E] of expsRanked) {
                        const hsRanked = [...E.hs.entries()].sort((a, b) => (b[1].total - a[1].total) || cmpStr(a[0], b[0]));
                        for (const [hs, H] of hsRanked) {
                            const consRanked = [...H.cons.entries()].sort((a, b) => (b[1].count - a[1].count) || cmpStr(a[1].name, b[1].name));
                            for (const [, C] of consRanked) {
                                metrics.push([y, mo, E.name, E.pin, E.total, hs, H.total, C.name, C.fPort, C.count]);
                            }
                        }
                    }
                }
            }
            const metricsHeader = ['Year', 'Month', 'Exporter Name', 'Exporter Pin Code', 'Export Count', 'HS Code', 'HS Count', 'Consignee Name', 'Foreign Port', 'Consignee Count'];
    
            // ---- KPI_2 sheet: group Exporter -> Year -> Month -> Day -> Consignee on the
            //      MERGED+disambiguated names. Exporter ranks by total volume; Year/Month/
            //      Day descending; Consignee-day leaf by shipment count. Each leaf row is one
            //      consignee (+ Foreign Port) on one day, with a comma-separated list of its
            //      distinct HS Codes that day. Monthly Shipment Count is the row count of
            //      consignee entries for that Exporter/Year/Month (not a sum), filled on
            //      every row of that group. Exporter/Year/Month/Day all repeat on every
            //      row, so the sheet pivots directly (including by day). ----
            const k2 = new Map(); // exKey -> { name, pin, total, years: Map }
            for (const g of groups) {
                const exPin = str(g.d[iExporterPin]).trim();
                const exKey = exG.keyOf(g.d[iExporter], exPin), exName = exG.dispOf(g.d[iExporter], exPin);
                const y = str(g.d[iYear]), mo = str(g.d[iMonth]), day = str(g.d[iDay]), hs = g.d[iHS];
                const fPort = str(g.d[iForeignPort]).trim();
                const coKey = coG.keyOf(g.d[iConsignee], fPort), coName = coG.dispOf(g.d[iConsignee], fPort);
                let X = k2.get(exKey); if (!X) { X = { name: exName, pin: exPin, total: 0, years: new Map() }; k2.set(exKey, X); }
                X.total += g.n;
                let Y = X.years.get(y); if (!Y) { Y = { months: new Map() }; X.years.set(y, Y); }
                let M = Y.months.get(mo); if (!M) { M = { days: new Map() }; Y.months.set(mo, M); }
                let D = M.days.get(day); if (!D) { D = { cons: new Map() }; M.days.set(day, D); }
                let C = D.cons.get(coKey); if (!C) { C = { name: coName, fPort, count: 0, hs: new Set() }; D.cons.set(coKey, C); }
                C.count += g.n;
                if (hs) C.hs.add(hs);
            }
            // Row count per Exporter/Year/Month (point 6) - a second pass over the already-
            // built tree (rather than accumulated per shipment above) so repeat HS codes for
            // the same day+consignee, which collapse into one row, are only counted once.
            for (const X of k2.values()) {
                X.rowCount = 0;
                for (const Y of X.years.values()) {
                    for (const M of Y.months.values()) {
                        let rc = 0;
                        for (const D of M.days.values()) rc += D.cons.size;
                        M.rowCount = rc;
                        X.rowCount += rc;
                    }
                }
            }
            const kpi2 = [];
            const k2ExRanked = [...k2.entries()].sort((a, b) => (b[1].total - a[1].total) || cmpStr(a[1].name, b[1].name));
            for (const [, X] of k2ExRanked) {
                const yRanked = [...X.years.entries()].sort((a, b) => yearNum(b[0]) - yearNum(a[0]));
                for (const [y, Y] of yRanked) {
                    const mRanked = [...Y.months.entries()].sort((a, b) => monthIndex(b[0]) - monthIndex(a[0]));
                    for (const [mo, M] of mRanked) {
                        const dRanked = [...M.days.entries()].sort((a, b) => dayNum(b[0]) - dayNum(a[0]));
                        for (const [day, D] of dRanked) {
                            const cRanked = [...D.cons.entries()].sort((a, b) => (b[1].count - a[1].count) || cmpStr(a[1].name, b[1].name));
                            for (const [, C] of cRanked) {
                                const hsList = [...C.hs].sort(cmpStr).join(', ');
                                kpi2.push([X.name, X.pin, y, mo, M.rowCount, day, C.name, C.fPort, hsList]);
                            }
                        }
                    }
                }
            }
            const kpi2Header = ['Exporter Name', 'Exporter Pin Code', 'Year', 'Month', 'Monthly Shipment Count', 'Day', 'Consignee Name', 'Foreign Port', 'HS Codes'];
            // Exporter subtotals for the "KPI Pivot Tables" sheet's right (KPI_2) pivot,
            // which is collapsed below Exporter Name. k2ExRanked is already total-desc sorted.
            const kpi2Grid = k2ExRanked.map(([, X]) => ({ exporter: X.name, total: X.rowCount }));
    
            // ---- Print-once copies (KPI_1a/KPI_2a): same rows as KPI_1/KPI_2, but each
            //      hierarchy level's cells are blanked on rows where every level from the
            //      top down through that one is unchanged from the row above - i.e. each
            //      distinct group is printed once, like a classic report layout. Year/Month
            //      are tied to the Exporter level (so they reprint alongside every distinct
            //      Exporter, not just when they themselves change); similarly Exporter/Pin/
            //      Year are tied to the Month level in KPI_2a. Leaf metrics that aren't a
            //      sorting level (Consignee Count, HS Codes) are always shown. ----
            function printOnceRows(rows, levels) {
                const out = [];
                for (let i = 0; i < rows.length; i++) {
                    const r = rows[i].slice();
                    if (i > 0) {
                        const prev = rows[i - 1];
                        let same = true;
                        for (const cols of levels) {
                            if (same && cols.every(c => String(r[c]) === String(prev[c]))) {
                                for (const c of cols) r[c] = '';
                            } else {
                                same = false;
                            }
                        }
                    }
                    out.push(r);
                }
                return out;
            }
            const metrics1a = printOnceRows(metrics, [[0, 1, 2, 3, 4], [5, 6], [7, 8]]);
            const kpi2a = printOnceRows(kpi2, [[0, 1, 2, 3, 4], [5], [6, 7]]);
    
            // ---- Name Mapping sheet: for each name that a KPI sheet collapsed from >1
            //      spelling, the shared Exporter Pin Code/Foreign Port that disambiguates it
            //      (point 4), the canonical (most-frequent) name shown there, and every
            //      original spelling it stands for (" / "-joined, most-frequent first).
            //      Exporters first, then consignees; biggest merges on top. ----
            const nameMap = [];
            for (const m of exG.merges) nameMap.push(['Exporter', m.canon, m.extra, m.variants.join(' / ')]);
            for (const m of coG.merges) nameMap.push(['Consignee', m.canon, m.extra, m.variants.join(' / ')]);
            const nameMapHeader = ['Type', 'Name (most frequent)', 'Exporter Pin Code/ Foreign Port', 'Merged Spellings'];
    
            // Exact widths taken from width.xlsx; columns not in that file get fixed widths.
            const SHIP_W = { 'Year': 8, 'Month': 11, 'Day': 7, 'Exporter Name': 35.125, 'HS Code': 13.57, 'Consignee Name': 30.25, 'IEC No': 15, 'Exporter Pin Code': 20.71, 'Exporter Address': 62.125, 'Exporter District': 18, 'Exporter State': 17.86, 'Exporter Phone No': 22.86, 'Exporter Email ID': 20.71, 'Product Description': 140 };
            const MET_W = { 'Year': 8, 'Month': 11, 'Exporter Name': 41.25, 'Exporter Pin Code': 20.71, 'Export Count': 13, 'HS Code': 9.875, 'HS Count': 12, 'Consignee Name': 36, 'Foreign Port': 18, 'Consignee Count': 17.25 };
            const K2_W = { 'Exporter Name': 41.25, 'Exporter Pin Code': 20.71, 'Year': 8, 'Month': 11, 'Monthly Shipment Count': 17, 'Day': 7, 'Consignee Name': 36, 'Foreign Port': 18, 'HS Codes': 135.5 };
            const MAP_W = { 'Type': 11, 'Exporter Pin Code/ Foreign Port': 24, 'Name (most frequent)': 41.25, 'Merged Spellings': 242.14 };
            // No width.xlsx reference for these 18 - reasonable fixed widths instead.
            const RAW_EXTRA_W = { 'Indian Port Code': 12, 'Indian Port Name': 20, 'Consignee Address': 45, 'Foreign Port': 18, 'Foreign Country': 16, 'Mode': 8, 'Trade': 10, 'Quantity': 10, 'Unit Quantity': 12, 'Unit Value (Invoice)': 14, 'Currency': 9, 'Total Value (FC)': 14, 'Unit Rate (USD)': 12, 'Exchange Rate': 12, 'Total Value (USD)': 14, 'Unit Rate (INR)': 12, 'FOB Value (INR)': 14, 'Drawback': 10 };
            const RAW_W = Object.assign({}, SHIP_W, RAW_EXTRA_W);
    
            const rawHeader = ALL_COLUMNS.map(c => c.header);
    
            const shipAoa = [header, ...data];
            const kpi1Aoa = [metricsHeader, ...metrics];
            const kpi1aAoa = [metricsHeader, ...metrics1a];
            const kpi2Aoa = [kpi2Header, ...kpi2];
            const kpi2aAoa = [kpi2Header, ...kpi2a];
            const nameMapAoa = [nameMapHeader, ...nameMap];
            const rawAoa = [rawHeader, ...rawData];
            const wsShip = XLSX.utils.aoa_to_sheet(shipAoa);
            const wsMet = XLSX.utils.aoa_to_sheet(kpi1Aoa);
            const wsMet1a = XLSX.utils.aoa_to_sheet(kpi1aAoa);
            const wsK2 = XLSX.utils.aoa_to_sheet(kpi2Aoa);
            const wsK2a = XLSX.utils.aoa_to_sheet(kpi2aAoa);
            const wsMap = XLSX.utils.aoa_to_sheet(nameMapAoa);
            const wsRaw = XLSX.utils.aoa_to_sheet(rawAoa);
    
            // ---- "KPI Pivot Tables" sheet: rendered grid for both pivots, side by side
            //      (B:C = KPI_1 collapsed to Year>Month, E:F = KPI_2 collapsed to Exporter).
            //      Excel re-applies PivotStyleMedium2 live from the style name once the file
            //      opens, so the data rows themselves need no manual styling; only the merged
            //      title cells and the "Row Labels"/data-caption header cells get one (see
            //      injectPivotSheetStyling, called from injectPivotTables). ----
            const kpi1GrandTotal = kpi1Grid.reduce((s, y) => s + y.yearTotal, 0);
            const kpi2GrandTotal = kpi2Grid.reduce((s, e) => s + e.total, 0);
            const kpi1DataRows = kpi1Grid.reduce((s, y) => s + 1 + y.months.length, 0) + 1; // + grand total
            const kpi2DataRows = kpi2Grid.length + 1; // + grand total
            const pivotTotalRows = 3 + Math.max(kpi1DataRows, kpi2DataRows); // blank + title + header + data
            const pivotAoa = [];
            for (let r = 0; r < pivotTotalRows; r++) pivotAoa.push([]);
            pivotAoa[1][1] = 'KPI_1_Pivot'; pivotAoa[1][4] = 'KPI_2_Pivot';
            pivotAoa[2][1] = 'Row Labels'; pivotAoa[2][2] = 'Sum of Consignee Count';
            pivotAoa[2][4] = 'Row Labels'; pivotAoa[2][5] = 'Shipment Count';
            let pr = 3;
            for (const y of kpi1Grid) {
                pivotAoa[pr][1] = y.year; pivotAoa[pr][2] = y.yearTotal; pr++;
                for (const m of y.months) { pivotAoa[pr][1] = m.month; pivotAoa[pr][2] = m.total; pr++; }
            }
            pivotAoa[pr][1] = 'Grand Total'; pivotAoa[pr][2] = kpi1GrandTotal;
            pr = 3;
            for (const e of kpi2Grid) { pivotAoa[pr][4] = e.exporter; pivotAoa[pr][5] = e.total; pr++; }
            pivotAoa[pr][4] = 'Grand Total'; pivotAoa[pr][5] = kpi2GrandTotal;
            const wsPivot = XLSX.utils.aoa_to_sheet(pivotAoa);
            wsPivot['!merges'] = [XLSX.utils.decode_range('B2:C2'), XLSX.utils.decode_range('E2:F2')];
    
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, wsShip, 'Shipments');
            XLSX.utils.book_append_sheet(wb, wsMet, 'KPI_1_Pivot');
            XLSX.utils.book_append_sheet(wb, wsMet1a, 'KPI_1');
            XLSX.utils.book_append_sheet(wb, wsK2, 'KPI_2_Pivot');
            XLSX.utils.book_append_sheet(wb, wsK2a, 'KPI_2');
            XLSX.utils.book_append_sheet(wb, wsPivot, 'KPI Pivot Tables');
            XLSX.utils.book_append_sheet(wb, wsMap, 'Name Mapping');
            XLSX.utils.book_append_sheet(wb, wsRaw, 'Raw Data');
            // KPI_1_Pivot/KPI_2_Pivot (sheet indices 1 and 3, 0-based) only exist to feed the
            // "KPI Pivot Tables" PivotTables' source data; hide them from normal view.
            wb.Workbook = wb.Workbook || {};
            wb.Workbook.Sheets = [{}, { Hidden: 1 }, {}, { Hidden: 1 }, {}, {}, {}, {}];
    
            const specs = [
                { sheetIndex: 1, tableId: 1, name: 'Table1', headers: header, stripes: 1,
                  widths: header.map(h => SHIP_W[h]),
                  ref: 'A1:' + colLetter(header.length) + shipAoa.length },
                { sheetIndex: 2, tableId: 2, name: 'Table2', headers: metricsHeader, stripes: 1,
                  widths: metricsHeader.map(h => MET_W[h]),
                  ref: 'A1:' + colLetter(metricsHeader.length) + kpi1Aoa.length },
                { sheetIndex: 3, tableId: 3, name: 'Table3', headers: metricsHeader, stripes: 1,
                  widths: metricsHeader.map(h => MET_W[h]),
                  ref: 'A1:' + colLetter(metricsHeader.length) + kpi1aAoa.length },
                { sheetIndex: 4, tableId: 4, name: 'Table4', headers: kpi2Header, stripes: 1,
                  widths: kpi2Header.map(h => K2_W[h]),
                  ref: 'A1:' + colLetter(kpi2Header.length) + kpi2Aoa.length },
                { sheetIndex: 5, tableId: 5, name: 'Table5', headers: kpi2Header, stripes: 1,
                  widths: kpi2Header.map(h => K2_W[h]),
                  ref: 'A1:' + colLetter(kpi2Header.length) + kpi2aAoa.length },
                { sheetIndex: 7, tableId: 6, name: 'Table6', headers: nameMapHeader, stripes: 1,
                  widths: nameMapHeader.map(h => MAP_W[h]),
                  ref: 'A1:' + colLetter(nameMapHeader.length) + nameMapAoa.length },
                { sheetIndex: 8, tableId: 7, name: 'Table7', headers: rawHeader, stripes: 1,
                  widths: rawHeader.map(h => RAW_W[h]),
                  ref: 'A1:' + colLetter(rawHeader.length) + rawAoa.length },
            ];
    
            // KPI_1 pivot (left, source KPI_1_Pivot sheet/Table2): rows = Year/Month/
            // Exporter Name/Exporter Pin Code/HS Code/Consignee Name, data = Consignee
            // Count, collapsed below Month (visibleDepth=2). Export Count/HS Count/Foreign
            // Port are defined but left off both axes (available in the field list, unused
            // by default).
            const iMetYear = metricsHeader.indexOf('Year'), iMetMonth = metricsHeader.indexOf('Month'),
                iMetEx = metricsHeader.indexOf('Exporter Name'), iMetPin = metricsHeader.indexOf('Exporter Pin Code'),
                iMetHS = metricsHeader.indexOf('HS Code'), iMetCo = metricsHeader.indexOf('Consignee Name'),
                iMetCoCount = metricsHeader.indexOf('Consignee Count');
            const kpi1Dicts = metricsHeader.map((h, i) => i === iMetCoCount ? null : buildCacheDict(metrics, i));
            const kpi1RowItemsXml = rowItems2Level(
                kpi1Grid.map(y => ({ outer: y.year, inner: y.months.map(m => m.month) })),
                kpi1Dicts[iMetYear], kpi1Dicts[iMetMonth]);
            const kpi1Pivot = buildPivotParts(metrics, metricsHeader, [iMetYear, iMetMonth, iMetEx, iMetPin, iMetHS, iMetCo], iMetCoCount, 2, {
                dicts: kpi1Dicts, sourceTableName: 'Table2', cacheId: 1, name: 'PivotTable1',
                location: 'B3:C' + (3 + kpi1DataRows), rowItemsXml: kpi1RowItemsXml,
                dataFieldName: 'Sum of Consignee Count',
            });
    
            // KPI_2 pivot (right, source KPI_2_Pivot sheet/Table4): rows = Exporter Name/
            // Year/Month/Day/Consignee Name/Foreign Port, collapsed below Exporter Name
            // (visibleDepth=1). Data is a Count on Monthly Shipment Count, not a Sum: that
            // column is a per-month total repeated on every row of that month, so summing
            // it would overcount; a Count of rows stays correct at every drill level and
            // matches it exactly when collapsed to Month.
            const iK2Ex = kpi2Header.indexOf('Exporter Name'), iK2Year = kpi2Header.indexOf('Year'),
                iK2Month = kpi2Header.indexOf('Month'), iK2Day = kpi2Header.indexOf('Day'),
                iK2Co = kpi2Header.indexOf('Consignee Name'), iK2FPort = kpi2Header.indexOf('Foreign Port'),
                iK2ShipCount = kpi2Header.indexOf('Monthly Shipment Count');
            const kpi2Dicts = kpi2Header.map((h, i) => buildCacheDict(kpi2, i));
            const kpi2RowItemsXml = rowItems1Level(kpi2Grid.map(e => e.exporter), kpi2Dicts[iK2Ex]);
            const kpi2Pivot = buildPivotParts(kpi2, kpi2Header, [iK2Ex, iK2Year, iK2Month, iK2Day, iK2Co, iK2FPort], iK2ShipCount, 1, {
                dicts: kpi2Dicts, sourceTableName: 'Table4', cacheId: 2, name: 'PivotTable2',
                location: 'E3:F' + (3 + kpi2DataRows), rowItemsXml: kpi2RowItemsXml,
                dataFieldName: 'Shipment Count', subtotal: 'count',
            });
    
            const pivotSpec = {
                sheetIndex: 6,
                widths: [{ col: 2, width: 45 }, { col: 3, width: 30 }, { col: 5, width: 45 }, { col: 6, width: 30 }],
                titleCells: ['B2', 'E2'],
                headerCells: ['B3', 'C3', 'E3', 'F3'],
                pivots: [
                    { id: 1, cacheDefXml: kpi1Pivot.cacheDefXml, cacheRecordsXml: kpi1Pivot.cacheRecordsXml, pivotTableXml: kpi1Pivot.pivotTableXml },
                    { id: 2, cacheDefXml: kpi2Pivot.cacheDefXml, cacheRecordsXml: kpi2Pivot.cacheRecordsXml, pivotTableXml: kpi2Pivot.pivotTableXml },
                ],
            };
    
            await writeWorkbookWithTables(wb, specs, filename, pivotSpec);
            return { all: data.length };
        }
    
        async function runExport(onCount, pageUrl) {
            if (!template || !template.key) throw new Error('not-ready');
            const raw = await collectAllRows(onCount);
            if (!raw.length) throw new Error('No rows returned for the current filter.');
            const filename = buildFilename(pageUrl);
            const stats = await exportWorkbook(raw, filename);
            return { count: stats.all, filename };
        }
    
        // =====================================================================
        //  BUTTON (round FAB with download -> spinner -> check states)
        // =====================================================================
        const SVG_DOWNLOAD = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M5 21h14"/></svg>';
        const SVG_CHECK = '<svg class="chk" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12l5 5L20 6"/></svg>';
    
        function injectStyle() {
            if (document.getElementById('eximelite-style')) return;
            const s = document.createElement('style');
            s.id = 'eximelite-style';
            s.textContent =
                '#eximelite-fab{position:fixed;top:16px;right:16px;width:44px;height:44px;border-radius:50%;' +
                'background:#217346;border:none;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.35);' +
                'z-index:2147483647;display:flex;align-items:center;justify-content:center;padding:0;transition:background .2s;}' +
                '#eximelite-fab:hover{background:#1a5c38;}' +
                '#eximelite-fab:disabled{cursor:default;}' +
                // progress pill: hides behind the button, slides out to the left while running
                '#eximelite-pill{position:fixed;top:16px;right:16px;width:44px;height:44px;border-radius:22px;' +
                'background:#217346;color:#fff;font:600 13px/44px system-ui,Segoe UI,Arial,sans-serif;' +
                'white-space:nowrap;overflow:hidden;box-sizing:border-box;z-index:2147483646;pointer-events:none;' +
                'transition:width .35s ease,background .3s;}' +
                '#eximelite-pill span{display:inline-block;font:500 12.5px/44px ui-monospace,Consolas,monospace;padding:0 10px 0 18px;opacity:0;transition:opacity .2s;}' +
                '#eximelite-pill.open{box-shadow:0 2px 8px rgba(0,0,0,.35);}' +
                '#eximelite-pill.open span{opacity:1;}' +
                '#eximelite-pill.err{background:#c0392b;}' +
                '.eximelite-spin{display:grid;gap:3px;width:20px;}' +
                '.eximelite-spin i{display:block;height:3px;border-radius:2px;background:#fff;transform-origin:left;animation:eximelite-fill 1.2s ease-in-out infinite;}' +
                '.eximelite-spin i:nth-child(2){animation-delay:.2s;}.eximelite-spin i:nth-child(3){animation-delay:.4s;}' +
                '@keyframes eximelite-fill{0%,100%{transform:scaleX(.25);opacity:.5;}50%{transform:scaleX(1);opacity:1;}}' +
                // done: the check mark draws itself like a pen stroke
                '#eximelite-fab .chk path{stroke-dasharray:24;stroke-dashoffset:24;animation:eximelite-draw .45s ease-out forwards;}' +
                '@keyframes eximelite-draw{to{stroke-dashoffset:0;}}';
            (document.head || document.documentElement).appendChild(s);
        }
    
        function makeFab(onClick) {
            injectStyle();
            let btn = document.getElementById('eximelite-fab');
            if (!btn) {
                btn = document.createElement('button');
                btn.id = 'eximelite-fab';
                document.body.appendChild(btn);
            }
            btn.onclick = onClick;
            let pill = document.getElementById('eximelite-pill');
            if (!pill) {
                pill = document.createElement('div');
                pill.id = 'eximelite-pill';
                pill.innerHTML = '<span></span>';
                document.body.appendChild(pill);
            }
            const text = pill.firstChild;
            // Show text in the pill; its width follows the text (44px is the button under its right end).
            const show = t => { text.textContent = t; pill.classList.add('open'); pill.style.width = (44 + text.scrollWidth) + 'px'; };
            const hide = () => { pill.classList.remove('open'); pill.style.width = '44px'; setTimeout(() => pill.classList.remove('err'), 350); };
            let timer = null;
            const later = (fn, ms) => { clearTimeout(timer); timer = setTimeout(fn, ms); };
            // Hover opens the pill with a hint; `active` keeps it off while progress/result shows.
            let active = false;
            btn.onmouseenter = () => { if (!active) show('Click to export all to Excel'); };
            btn.onmouseleave = () => { if (!active) hide(); };
            const api = {
                el: btn,
                idle() { active = false; btn.disabled = false; btn.innerHTML = SVG_DOWNLOAD; },
                busy(n) {
                    clearTimeout(timer);
                    active = true; btn.disabled = true;
                    if (!btn.querySelector('.eximelite-spin')) btn.innerHTML = '<span class="eximelite-spin"><i></i><i></i><i></i></span>';
                    pill.classList.remove('err');
                    show(n ? ('Loading ' + n + ' rows...') : 'Loading...');
                },
                done(n) {
                    btn.disabled = false; btn.innerHTML = SVG_CHECK;
                    show('Done: ' + n + ' rows');
                    later(() => { hide(); later(() => api.idle(), 400); }, 2000);
                },
                error(msg) {
                    api.idle(); active = true;
                    pill.classList.add('err');
                    show('Export failed');
                    later(() => { if (msg) alert(msg); later(() => { hide(); active = false; }, 1500); }, 450); // let the red pill paint before alert blocks
                },
            };
            api.idle();
            return api;
        }
    
        function bootFab(onClick, shouldShow) {
            let api = null;
            const ensure = () => {
                if (!document.body) return;
                if (shouldShow && !shouldShow()) {
                    const ex = document.getElementById('eximelite-fab');
                    if (ex) { ex.remove(); api = null; }
                    const pill = document.getElementById('eximelite-pill');
                    if (pill) pill.remove();
                    return;
                }
                if (!document.getElementById('eximelite-fab')) api = makeFab(onClick);
                else if (!api) api = makeFab(onClick);
            };
            const start = () => { ensure(); setInterval(ensure, 2000); };
            if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
            else start();
            return { get: () => api };
        }
    
        // =====================================================================
        //  WIRING
        // =====================================================================
        function post(action, extra) {
            try { window.parent.postMessage(Object.assign({ source: 'eximelite-engine', action: action }, extra || {}), '*'); } catch (e) { /* ignore */ }
        }
    
        // Debug helper (not part of normal operation): dump the Property/Name strings
        // of every column in the MOST RECENT fact query seen, so a new field can be
        // located by its exact Power BI name before adding it to COLUMNS. Click into /
        // scroll / expand the visual that shows the field you're after (so it fires its
        // own query), THEN run eximeliteListFields() in the devtools console, scoped to
        // the app.powerbi.com frame.
        function initFieldDump() {
            window.eximeliteListFields = function () {
                const q = lastQuery || template;
                if (!q) { console.log('[eximelite] no query captured yet - interact with the report first.'); return; }
                const cmd = q.parsed.queries[0].Query.Commands[0].SemanticQueryDataShapeCommand;
                const sel = cmd.Query.Select || [];
                console.table(sel.map(s => ({
                    Name: s.Name,
                    NativeReferenceName: s.NativeReferenceName,
                    Property: s.Column && s.Column.Property,
                    HierarchyLevel: s.HierarchyLevel && s.HierarchyLevel.Level,
                })));
            };
        }
    
        function initEngine() {
            installHooks();
            initFieldDump();
    
            // Triggered by the button on the host page.
            window.addEventListener('message', (e) => {
                const d = e.data;
                if (!d || d.source !== 'eximelite-btn' || d.action !== 'start') return;
                if (!template || !template.key) return; // stay silent (not the right frame / not ready)
                post('ack');
                runExport(n => post('progress', { count: n }), d.pageUrl)
                    .then(r => post('done', { count: r.count, filename: r.filename }))
                    .catch(err => post('error', { message: (err && err.message) || String(err) }));
            });
    
            // Standalone: report opened directly on app.powerbi.com -> local button.
            if (window.top === window.self) {
                const holder = bootFab(() => {
                    const api = holder.get(); if (!api) return;
                    if (!template || !template.key) { api.error('Not ready: interact with the report (apply a filter) first, then click again.'); return; }
                    api.busy(0);
                    runExport(n => api.busy(n), location.href).then(r => api.done(r.count)).catch(err => api.error(err.message));
                });
            }
        }
    
        function initHostButton() {
            let ackReceived = false, ackTimer = null, holder = null;
    
            const start = () => {
                const api = holder.get(); if (!api) return;
                const iframe = document.querySelector('iframe[src*="app.powerbi.com"]');
                if (!iframe || !iframe.contentWindow) { api.error('Report not loaded yet. Wait for the report to appear, then try again.'); return; }
                api.busy(0);
                ackReceived = false;
                iframe.contentWindow.postMessage({ source: 'eximelite-btn', action: 'start', pageUrl: location.href }, '*');
                clearTimeout(ackTimer);
                ackTimer = setTimeout(() => { if (!ackReceived) api.error('Could not reach the report. Apply a filter so the table loads, then try again.'); }, 6000);
            };
    
            window.addEventListener('message', (e) => {
                const d = e.data, api = holder && holder.get();
                if (!d || d.source !== 'eximelite-engine' || !api) return;
                if (d.action === 'ack') { ackReceived = true; clearTimeout(ackTimer); }
                else if (d.action === 'progress') { api.busy(d.count); }
                else if (d.action === 'done') { ackReceived = true; clearTimeout(ackTimer); api.done(d.count); }
                else if (d.action === 'error') { ackReceived = true; clearTimeout(ackTimer); api.error(d.message); }
            });
    
            // Only show the button on pages that actually embed the report.
            holder = bootFab(start, () => !!document.querySelector('iframe[src*="app.powerbi.com"]'));
        }
    
        if (IN_PBI) initEngine();
        else initHostButton();
    }

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
    //   4. Fetches each detail page and parses the 16 columns (Certificate Status comes from
    //      the list). Rows with no detail page (download-only) take number, dates, product,
    //      model and status from the list; for a certificate PDF (/reports/HRCoCProductCertificate,
    //      or /reports/QMCertificate for IECEE) the rest is read from the PDF in the browser
    //      (pdf.js). If the file isn't a PDF in SABER's template, the rest stays blank.
    //   5. Writes one sheet, 16 columns, one row per certificate.
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
    
    function runSaber() {
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
        // `list` = the column's index in the list row (Number, Issue, Expiry, Product, Model,
        // Status), used for Certificate Status and for rows that have no detail page.
        // `pdf` = the label in the certificate PDF, where it differs from the detail page's.
        const COLUMNS = [
            { header: 'Certification Body', kind: 'certBody' },
            { header: 'Certificate Number', label: 'Certificate Number', list: 0 },
            { header: 'Issue Date', label: 'Issue Date', list: 1, date: true },
            { header: 'Expire Date', label: 'Expire Date', list: 2, date: true },
            { header: 'Certificate Status', list: 5 },
            { header: 'Certificate Type', label: 'Certificate Type' },
            { header: 'Commercial Registration No', label: 'Commercial Registration No', pdf: 'Registration No.' },
            { header: 'Establishment Address', label: 'Establishment Address' },
            { header: 'Model Name', label: 'Model Name', list: 4 },
            { header: 'Trade Mark', label: 'Trade Mark', pdf: 'Trademark' },
            { header: 'Product Name', label: 'Product Name', list: 3 },
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
    
        // Returns [{ certNo, detailUrl|null, reviewUrl|null, pdfUrl|null, cells }] for every row across all pages.
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
                    const pdf = tr.querySelector('a[href*="/reports/HRCoCProductCertificate"], a[href*="/reports/QMCertificate"]');
                    rows.push({
                        certNo: no,
                        detailUrl: view ? view.href : null,
                        reviewUrl: review ? review.href : null,
                        pdfUrl: pdf ? pdf.href : null,
                        cells: [...tr.querySelectorAll('td')].map(td => clean(td.textContent)),
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
        // Ends at ") office", so a name with its own brackets ("TUV Austria (SHANGHAI) Co.")
        // is kept whole; falls back to the first ")" if that wording is missing.
        function certificationBody(text) {
            const m = text.match(/We\s*\(([\s\S]*?)\)\s*office/i) || text.match(/We\s*\(([\s\S]*?)\)/);
            if (!m) return '';
            return m[1].replace(/^[\s.]+|[\s.]+$/g, '').replace(/\s+/g, ' ').trim();
        }
    
        // Value of column `c` from the list row; list dates are yyyy-mm-dd, output is dd/mm/yyyy.
        function listValue(row, c) {
            const v = row.cells[c.list] || '';
            const m = c.date && /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
            return m ? m[3] + '/' + m[2] + '/' + m[1] : v;
        }
    
        function parseDetail(doc, row) {
            return COLUMNS.map(c => c.kind === 'certBody' ? certificationBody(doc.body.textContent)
                : c.label ? valueForLabel(doc, c.label)
                : listValue(row, c));
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
                    out[i + k] = parseDetail(doc, row);
                    setStatus('Details (' + (++done) + '/' + rows.length + ')');
                }));
            }
            return out;
        }
    
        // ---------- certificate PDFs ----------
        // Download-only certificates have no detail page, only a PDF in SABER's own template:
        // on page 1 each label is 8pt text and its value 10pt text to its right on the same
        // row. pdf.worker is @required too, so pdf.js runs in-page without loading a worker.
    
        const isValue = it => it.h >= 9.5;
    
        // Page 1's text pieces with their position (presentation-form Arabic normalised).
        async function pdfItems(buf) {
            const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
            const { items } = await (await pdf.getPage(1)).getTextContent();
            return items.filter(it => it.str.trim()).map(it => ({
                s: clean(it.str.normalize('NFKC')), x: it.transform[4], y: it.transform[5], h: it.height,
            }));
        }
    
        // Value pieces right of the label, within ~one row of it, up to the next label on that row.
        function pdfValue(items, label) {
            const want = label.toLowerCase();
            const lab = items.find(it => !isValue(it) && it.s.toLowerCase() === want);
            if (!lab) return '';
            const near = it => Math.abs(it.y - lab.y) <= 9 && it.x > lab.x;
            const stop = Math.min(...items.filter(it => !isValue(it) && near(it)).map(it => it.x));
            return items.filter(it => isValue(it) && near(it) && it.x < stop)
                .sort((a, b) => b.y - a.y || a.x - b.x).map(it => it.s).join(' ');
        }
    
        // The "We (... NAME ...) Office Number" line, pieces put back in left-to-right order.
        function pdfCertBody(items) {
            const we = items.find(it => it.s === 'We');
            if (!we) return '';
            return certificationBody(items.filter(it => Math.abs(it.y - we.y) <= 2)
                .sort((a, b) => a.x - b.x).map(it => it.s).join(' '));
        }
    
        // List values first (number, dates, status, product, model); the rest from the PDF.
        function parsePdf(items, row) {
            return COLUMNS.map(c => c.list != null ? listValue(row, c)
                : c.kind === 'certBody' ? pdfCertBody(items)
                : pdfValue(items, c.pdf || c.label));
        }
    
        // Fetch certificate PDFs in small waves. A PDF that can't be read leaves the row with
        // its list values only, so one bad file never stops the export.
        async function fetchPdfDetails(rows, setStatus) {
            const out = new Array(rows.length);
            let done = 0;
            for (let i = 0; i < rows.length; i += DETAIL_CONCURRENCY) {
                const slice = rows.slice(i, i + DETAIL_CONCURRENCY);
                await Promise.all(slice.map(async (row, k) => {
                    const res = await fetch(row.pdfUrl, { credentials: 'include' });
                    if (/\/Account\/Login/i.test(res.url)) throw new Error('Logged out. Please log in again and retry.');
                    let items = [];
                    try {
                        if (!res.ok) throw new Error('HTTP ' + res.status);
                        items = await pdfItems(await res.arrayBuffer());
                    } catch (e) {
                        console.warn('[SABER CoC exporter] PDF not read for ' + row.certNo, e);
                    }
                    out[i + k] = parsePdf(items, row);
                    setStatus('PDFs (' + (++done) + '/' + rows.length + ')');
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
            // No view or review button (download-only): certificate PDFs are read, anything
            // else gets list values only with the other columns blank.
            const listOnly = rows.filter(r => !r.detailUrl && !r.reviewUrl);
            const pdfRows = listOnly.filter(r => r.pdfUrl);
            data.push(...await fetchPdfDetails(pdfRows, setStatus));
            data.push(...listOnly.filter(r => !r.pdfUrl).map(r => COLUMNS.map(c => listValue(r, c))));
    
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
    
            const skipped = rows.length - data.length;
            return { total: data.length, skipped };
        }
    
        // ---------- button ----------
        // Same look, size and placement as the EximElite FAB (circular, top-right, 44px,
        // #217346) so the two can be merged later. Shows on every ContactOfficer page; the
        // export always runs over the whole certificates list regardless of the page it is
        // clicked from. Hovering opens the pill with a hint; it carries progress while it runs.
    
        const SVG_DOWNLOAD = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M5 21h14"/></svg>';
        const SVG_CHECK = '<svg class="chk" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12l5 5L20 6"/></svg>';
    
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
                '#saber-coc-pill span{display:inline-block;font:500 12.5px/44px ui-monospace,Consolas,monospace;padding:0 10px 0 18px;opacity:0;transition:opacity .2s;}' +
                '#saber-coc-pill.open{box-shadow:0 2px 8px rgba(0,0,0,.35);}' +
                '#saber-coc-pill.open span{opacity:1;}' +
                '#saber-coc-pill.err{background:#c0392b;}' +
                '.saber-coc-spin{display:grid;gap:3px;width:20px;}' +
                '.saber-coc-spin i{display:block;height:3px;border-radius:2px;background:#fff;transform-origin:left;animation:saber-coc-fill 1.2s ease-in-out infinite;}' +
                '.saber-coc-spin i:nth-child(2){animation-delay:.2s;}.saber-coc-spin i:nth-child(3){animation-delay:.4s;}' +
                '@keyframes saber-coc-fill{0%,100%{transform:scaleX(.25);opacity:.5;}50%{transform:scaleX(1);opacity:1;}}' +
                // done: the check mark draws itself like a pen stroke
                '#saber-coc-fab .chk path{stroke-dasharray:24;stroke-dashoffset:24;animation:saber-coc-draw .45s ease-out forwards;}' +
                '@keyframes saber-coc-draw{to{stroke-dashoffset:0;}}';
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
            // Hover opens the pill with a hint; `active` keeps it off while progress/result shows.
            let active = false;
            btn.onmouseenter = () => { if (!active) show('Click to export certificates to Excel'); };
            btn.onmouseleave = () => { if (!active) hide(); };
            const api = {
                el: btn,
                idle() { active = false; btn.disabled = false; btn.innerHTML = SVG_DOWNLOAD; },
                busy(n) {
                    clearTimeout(timer);
                    active = true; btn.disabled = true;
                    if (!btn.querySelector('.saber-coc-spin')) btn.innerHTML = '<span class="saber-coc-spin"><i></i><i></i><i></i></span>';
                    pill.classList.remove('err');
                    show(n || 'Loading...');
                },
                done(n) {
                    btn.disabled = false; btn.innerHTML = SVG_CHECK;
                    show('Done: ' + n + ' rows');
                    later(() => { hide(); later(() => api.idle(), 400); }, 2000);
                },
                error(msg) {
                    api.idle(); active = true;
                    pill.classList.add('err');
                    show('Export failed');
                    later(() => { if (msg) alert(msg); later(() => { hide(); active = false; }, 1500); }, 450); // let the red pill paint before alert blocks
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
        window.saberCoc = { collectListRows, unlock, fetchDetails, fetchPdfDetails, parseDetail, runExport, COLUMNS };
    
        bootFab();
    }
})();
