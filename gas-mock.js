// In-memory stand-in for the Apps Script services used by Code.gs.
// The real Code.gs runs unchanged inside a Node vm context. Sheets conversions that matter are emulated:
// "2026-11-02" becomes a Date, "09:00" becomes a time, digit strings become numbers, a leading
// apostrophe forces text, and a leading "=" would become a formula (recorded so tests can fail on it).
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function createBackend(codePath) {
  const sheets = new Map();
  const props = new Map();
  const outbox = [];
  const triggers = [];
  const logs = [];
  const formulaCells = [];

  const host = {
    sha256Bytes: function (text) { return Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()); },
    b64FromBytes: function (arr) { return Buffer.from(arr).toString('base64'); },
    b64FromText: function (text) { return Buffer.from(String(text), 'utf8').toString('base64'); },
    b64ToBytes: function (s) { return Array.from(Buffer.from(String(s), 'base64')); },
    uuid: function () { return crypto.randomUUID(); },
    sheets: sheets, props: props, outbox: outbox, triggers: triggers, logs: logs, formulaCells: formulaCells,
    cache: new Map()
  };

  const ctx = vm.createContext({ __host: host, Logger: { log: function (m) { logs.push(String(m)); } } });
  const prelude = `
    function toCell(v) {
      if (typeof v !== 'string') return v;
      if (v.charAt(0) === "'") return v.slice(1);
      if (v.charAt(0) === '=') { __host.formulaCells.push(v); return v; }
      let m = /^(\\d{4})-(\\d{2})-(\\d{2})$/.exec(v);
      if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      m = /^(\\d{1,2}):(\\d{2})$/.exec(v);
      if (m) return new Date(1899, 11, 30, Number(m[1]), Number(m[2]));
      if (/^-?\\d+(\\.\\d+)?$/.test(v)) return Number(v);
      return v;
    }
    class Sheet {
      constructor(name) { this.name = name; this.rows = []; }
      getLastRow() { return this.rows.length; }
      getLastColumn() { return this.rows.length ? this.rows[0].length : 0; }
      getRange(row, col, numRows, numCols) {
        const self = this;
        const nr = numRows || 1, nc = numCols || 1;
        return {
          getValues() {
            const out = [];
            for (let r = row - 1; r < row - 1 + nr; r++) {
              const src = self.rows[r] || [];
              const line = [];
              for (let c = col - 1; c < col - 1 + nc; c++) line.push(src[c] === undefined ? '' : src[c]);
              out.push(line);
            }
            return out;
          },
          setValues(vals) {
            vals.forEach(function (line, i) {
              while (self.rows.length < row + i) self.rows.push([]);
              line.forEach(function (v, j) {
                const target = self.rows[row - 1 + i];
                while (target.length < col - 1 + j) target.push('');
                target[col - 1 + j] = toCell(v);
              });
            });
            return this;
          },
          setValue(v) { return this.setValues([[v]]); },
          setFontWeight() { return this; }
        };
      }
      getDataRange() {
        const self = this;
        return { getValues() { return self.rows.map(function (r) { return r.slice(); }); } };
      }
      appendRow(values) {
        this.rows.push(values.map(toCell));
      }
      setFrozenRows() {}
      deleteRow(n) { this.rows.splice(n - 1, 1); }
    }
    const SpreadsheetApp = {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) { return __host.sheets.get(name) || null; },
          insertSheet(name) { const s = new Sheet(name); __host.sheets.set(name, s); return s; }
        };
      },
      openById() { return SpreadsheetApp.getActiveSpreadsheet(); }
    };
    const Utilities = {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest(alg, value) { return __host.sha256Bytes(value); },
      base64Encode(v) { return Array.isArray(v) ? __host.b64FromBytes(v) : __host.b64FromText(v); },
      base64Decode(s) { return __host.b64ToBytes(s); },
      getUuid() { return __host.uuid(); },
      newBlob(bytes, mime, name) { return { getName() { return name; }, getContentType() { return mime; }, getBytes() { return bytes; } }; },
      formatDate(d, tz, fmt) {
        const p = function (n) { return (n < 10 ? '0' : '') + n; };
        return fmt
          .replace('yyyy', String(d.getFullYear()))
          .replace('MM', p(d.getMonth() + 1))
          .replace('dd', p(d.getDate()))
          .replace('HH', p(d.getHours()))
          .replace('mm', p(d.getMinutes()))
          .replace('ss', p(d.getSeconds()));
      }
    };
    const CacheService = {
      getScriptCache() {
        return {
          put(k, v) { __host.cache.set(k, String(v)); },
          get(k) { return __host.cache.has(k) ? __host.cache.get(k) : null; },
          remove(k) { __host.cache.delete(k); }
        };
      }
    };
    const Session = { getScriptTimeZone() { return 'Asia/Kolkata'; } };
    const LockService = { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } };
    const PropertiesService = {
      getScriptProperties() {
        return {
          getProperty(k) { return __host.props.has(k) ? __host.props.get(k) : null; },
          setProperty(k, v) { __host.props.set(k, String(v)); },
          deleteProperty(k) { __host.props.delete(k); }
        };
      }
    };
    const ContentService = {
      MimeType: { JSON: 'application/json', JAVASCRIPT: 'application/javascript' },
      createTextOutput(text) {
        const o = { text: String(text), mime: '' };
        o.setMimeType = function (m) { o.mime = m; return o; };
        return o;
      }
    };
    const MailApp = { sendEmail(opts) { __host.outbox.push({ to: opts.to, subject: opts.subject, body: opts.body }); } };
    const ScriptApp = {
      getProjectTriggers() { return __host.triggers.map(function (t) { return { getHandlerFunction() { return t.fn; }, id: t.id }; }); },
      deleteTrigger(t) { const i = __host.triggers.findIndex(function (x) { return x.id === t.id; }); if (i >= 0) __host.triggers.splice(i, 1); },
      newTrigger(fn) {
        const b = { fn: fn };
        b.timeBased = function () { return b; };
        b.everyHours = function () { return b; };
        b.create = function () { __host.triggers.push({ fn: fn, id: __host.triggers.length + 1 }); return b; };
        return b;
      }
    };
    const DriveApp = { getFoldersByName() { return { hasNext() { return false; } }; }, createFolder() { return {}; } };
    const Drive = {};
  `;
  vm.runInContext(prelude, ctx);
  vm.runInContext(fs.readFileSync(codePath, 'utf8'), ctx, { filename: 'Code.gs' });

  // Calls doGet with a parameter map (no callback: returns plain JSON text).
  function get(params) {
    const out = vm.runInContext('doGet(__e)', Object.assign(ctx, { __e: { parameter: params || {} } }));
    return JSON.parse(out.text);
  }
  // Sends a POST body the way the page does (JSON text in the body).
  function post(body) {
    const out = vm.runInContext('doPost(__e)', Object.assign(ctx, { __e: { postData: { contents: JSON.stringify(body) } } }));
    return out.text;
  }
  // Login/register/forgot flow: POST, then fetchResult (what the page does).
  function postThenFetch(body) {
    const requestId = 'req' + Math.random().toString(36).slice(2, 12);
    post(Object.assign({}, body, { requestId: requestId }));
    return get({ action: 'fetchResult', requestId: requestId });
  }
  // Calls an action with a token; returns the parsed result.
  // Same route the page uses: the token travels in the POST body, the result is collected with fetchResult.
  function api(action, token, params) {
    return postThenFetch(Object.assign({}, params || {}, { action: action, token: token || '' }));
  }
  function rows(name) {
    const s = sheets.get(name);
    if (!s) return [];
    const headers = s.rows[0] || [];
    return s.rows.slice(1).map(function (r) {
      const o = {};
      headers.forEach(function (h, i) { o[h] = r[i] === undefined ? '' : r[i]; });
      return o;
    });
  }
  function runTrigger(fnName) {
    return vm.runInContext(fnName + '()', ctx);
  }
  // Raw text as Apps Script would send it (JSONP wrapper included when a callback is given).
  function getRaw(params) {
    return vm.runInContext("doGet(__e)", Object.assign(ctx, { __e: { parameter: params || {} } })).text;
  }
  // Fixes the backend clock to a local ISO time (same method as the unit tests).
  function setClock(localIso) {
    const Real = vm.runInContext("Date", ctx);
    const fixed = new Real(localIso).getTime();
    ctx.__NewDate = class extends Real {
      constructor(...args) { if (args.length === 0) super(fixed); else super(...args); }
      static now() { return fixed; }
    };
    vm.runInContext("Date = __NewDate", ctx);
  }
  function setGlobal(name, value) { ctx[name] = value; }
  return { get, getRaw, post, postThenFetch, api, rows, sheets, props, outbox, triggers, logs, formulaCells, runTrigger, setGlobal, setClock, ctx, host };
}

module.exports = { createBackend };
