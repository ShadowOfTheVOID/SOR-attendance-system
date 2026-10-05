// Minimal in-memory stand-in for the Apps Script services used by Code.js.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

class FakeRange {
  constructor(sheet, row, col, nr, nc) { Object.assign(this, { sheet, row, col, nr, nc }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const line = [];
      for (let c = 0; c < this.nc; c++) line.push(this.sheet.get(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  setValues(v) {
    if (v.length !== this.nr || v[0].length !== this.nc) throw new Error('dimension mismatch');
    v.forEach((line, r) => line.forEach((x, c) => this.sheet.set(this.row + r, this.col + c, x)));
    return this;
  }
  getDisplayValues() {
    return this.getValues().map((line) => line.map((v) => {
      if (Object.prototype.toString.call(v) === '[object Date]') {
        return `${v.getUTCMonth() + 1}/${v.getUTCDate()}/${v.getUTCFullYear()}`;
      }
      return String(v ?? '');
    }));
  }
  getValue() { return this.sheet.get(this.row, this.col); }
  setValue(x) { this.sheet.set(this.row, this.col, x); return this; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  insertCheckboxes() { return this; }
}

class FakeSheet {
  constructor(name) { this.name = name; this.cells = []; this.maxRows = 1000; this.maxCols = 26; }
  get(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, x) {
    if (r > this.maxRows || c > this.maxCols) throw new Error(`out of bounds R${r}C${c}`);
    (this.cells[r - 1] ||= [])[c - 1] = x;
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const col = a.charCodeAt(0) - 64;
      const m = /^[A-Z](\d*)/.exec(a);
      const row = m[1] ? Number(m[1]) : 1;
      return new FakeRange(this, row, col, this.maxRows - row + 1, 1);
    }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  getLastRow() {
    for (let r = this.cells.length; r > 0; r--) {
      if ((this.cells[r - 1] || []).some((x) => x !== '' && x != null)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.cells.forEach((row) => (row || []).forEach((x, i) => {
      if (x !== '' && x != null) max = Math.max(max, i + 1);
    }));
    return max;
  }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertColumnBefore(col) {
    this.maxCols++;
    this.cells.forEach((row) => row && row.splice(col - 1, 0, ''));
  }
  insertColumnAfter() { this.maxCols++; }
  appendRow(values) { const r = this.getLastRow() + 1; values.forEach((x, i) => this.set(r, i + 1, x)); }
  setFrozenRows() {}
  setFrozenColumns() {}
  setConditionalFormatRules() {}
  getName() { return this.name; }
}

const toSigned = (buf) => Array.from(buf).map((b) => (b > 127 ? b - 256 : b));

// tokens: { idToken: { email, emailVerified, displayName } } accepted by the fake lookup API.
function load({ now = new Date('2026-10-05T09:00:00Z'), tokens = {}, apiKey = 'test-key', keyInProps = true } = {}) {
  const makeSpreadsheet = (store) => ({
    getSheetByName: (n) => store[n] || null,
    insertSheet: (n) => (store[n] = new FakeSheet(n)),
    getSpreadsheetTimeZone: () => 'Etc/UTC',
    toast() {}
  });
  const sheets = {};
  const others = {}; // other spreadsheets by ID
  const ss = {
    ...makeSpreadsheet(sheets),
  };
  const props = apiKey && keyInProps ? { FIREBASE_API_KEY: apiKey } : {};
  const fetches = [];
  const openedIds = [];
  const cache = {};
  const rule = () => {
    const b = { whenTextEqualTo: () => b, setBackground: () => b, setRanges: () => b, build: () => ({}) };
    return b;
  };
  const RealDate = Date;
  class FixedDate extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(clock.now); }
    static now() { return clock.now; }
  }
  const clock = { now: now.getTime() };

  const context = {
    Date: FixedDate,
    SpreadsheetApp: {
      getActive: () => ss,
      openById: (id) => { openedIds.push(id); return others[id] || ss; },
      flush() {},
      getUi: () => ({ alert() {}, createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) }),
      newConditionalFormatRule: rule
    },
    Utilities: {
      formatDate(d, tz, fmt) {
        const iso = new RealDate(d.getTime()).toISOString();
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        if (fmt === 'HH:mm') return iso.slice(11, 16);
        throw new Error('unsupported format ' + fmt);
      },
      DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (alg, v) => toSigned(crypto.createHash(alg).update(v).digest()),
      computeHmacSha256Signature: (v, key) => toSigned(crypto.createHmac('sha256', key).update(v).digest()),
      base64EncodeWebSafe: (bytes) => Buffer.from(bytes.map((b) => b & 255)).toString('base64url'),
      getUuid: () => crypto.randomUUID()
    },
    Logger: { log() {} },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
    UrlFetchApp: {
      fetch(url, opts) {
        fetches.push(url);
        const key = new URL(url).searchParams.get('key');
        const account = tokens[JSON.parse(opts.payload).idToken];
        const ok = key === apiKey && account;
        const body = ok ? { users: [account] } : { error: { message: 'INVALID_ID_TOKEN' } };
        return { getResponseCode: () => (ok ? 200 : 400), getContentText: () => JSON.stringify(body) };
      }
    },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: (text) => ({ setMimeType() { return this; }, getContent: () => text })
    },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) }
  };
  vm.createContext(context);
  for (const f of ['Logic.js', 'Code.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), context, { filename: f });
  }
    // Calls the API the way the website does and unwraps the JSON response.
  const call = (action, idToken, params = {}) => {
    const out = context.doPost({ postData: { contents: JSON.stringify({ action, idToken, ...params }) } });
    return JSON.parse(out.getContent());
  };
  return {
    app: context, sheets, props, fetches, openedIds, call, tokens,
    /** Adds another spreadsheet that openById(id) returns; gives back its sheets store. */
    addSpreadsheet: (id) => { const store = {}; others[id] = makeSpreadsheet(store); return others[id]; },
    setTime: (iso) => { clock.now = new RealDate(iso).getTime(); }
  };
}

module.exports = { load };
