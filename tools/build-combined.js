// Build a single userscript that runs EximElite on powerbi/eximelite and the SABER
// conformity-cert exporter on saber.sa, branching by hostname. Each original is wrapped
// verbatim in its own function so their internal helper names never collide.
//
// Usage: node tools/build-combined.js   (or: npm run build)
const fs = require('fs');
const path = require('path');

const SCRIPTS = path.resolve(__dirname, '..', 'scripts');
const exim = fs.readFileSync(path.join(SCRIPTS, 'eximelite-exporter.user.js'), 'utf8');
const saber = fs.readFileSync(path.join(SCRIPTS, 'saber-coc-detail-exporter.user.js'), 'utf8');

// Strip the ==UserScript== metadata block and re-wrap the body's outer IIFE as a named
// function: `(function () {` -> `function NAME() {`, and drop the final `})();`.
function toFunction(src, name) {
    let body = src.replace(/\/\/ ==UserScript==[\s\S]*?\/\/ ==\/UserScript==\r?\n?/, '').trim();
    const open = body.indexOf('(function () {');
    if (open < 0) throw new Error('opening IIFE not found in ' + name);
    body = body.slice(0, open) + 'function ' + name + '() {' + body.slice(open + '(function () {'.length);
    const close = body.lastIndexOf('})();');
    if (close < 0) throw new Error('closing IIFE not found in ' + name);
    body = body.slice(0, close) + '}' + body.slice(close + '})();'.length);
    return body.trim();
}

const RAW = 'https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/eximelite-saber-combined.user.js';
const header = `// ==UserScript==
// @name         EximElite + SABER Combined Exporter
// @namespace    eximelite-exporter
// @version      1.6
// @description  One script, two tools by site: EximElite EDA-EX bulk exporter on powerbi/eximelite, and the SABER Conformity Certificates detail exporter on saber.sa.
// @author       Mohammed Aamer
// @license      MIT
// @homepageURL  https://github.com/aamer2411/userscripts-saber-eximelite-exporters
// @supportURL   https://github.com/aamer2411/userscripts-saber-eximelite-exporters/issues
// @updateURL    ${RAW}
// @downloadURL  ${RAW}
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
`;

const out = header + `
(function () {
    'use strict';
    if (/(^|\\.)saber\\.sa$/i.test(location.hostname)) runSaber();
    else runExim();

    ${toFunction(exim, 'runExim').replace(/\n/g, '\n    ')}

    ${toFunction(saber, 'runSaber').replace(/\n/g, '\n    ')}
})();
`;

const dest = path.join(SCRIPTS, 'eximelite-saber-combined.user.js');
fs.writeFileSync(dest, out);
console.log('wrote ' + path.relative(process.cwd(), dest) + ' (' + out.split('\n').length + ' lines)');
