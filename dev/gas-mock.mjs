// Google Apps Script จำลอง (เฉพาะส่วนที่ Code.gs ใช้) สำหรับทดสอบในเครื่อง
import fs from 'node:fs';
import vm from 'node:vm';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export function loadGas({ persist } = {}) {
  let db = { props: {}, cache: {}, sheets: [], files: {}, folders: {}, n: 0 };
  if (persist && fs.existsSync(persist)) db = JSON.parse(fs.readFileSync(persist, 'utf8'));
  const save = () => { if (persist) fs.writeFileSync(persist, JSON.stringify(db)); };
  const nid = p => p + (++db.n);
  const colIdx = s => s.charCodeAt(0) - 64;

  function sheetApi(sh) {
    const lastRow = () => { for (let i = sh.rows.length - 1; i >= 0; i--) if (sh.rows[i].some(v => v !== '' && v != null)) return i + 1; return 0; };
    const range = (r, c, nr = 1, nc = 1) => ({
      setValues(vals) { for (let i = 0; i < nr; i++) { while (sh.rows.length < r + i) sh.rows.push([]); const row = sh.rows[r + i - 1]; for (let j = 0; j < nc; j++) row[c + j - 1] = vals[i][j]; } save(); return this; },
      getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = sh.rows[r + i - 1] || []; out.push(Array.from({ length: nc }, (_, j) => row[c + j - 1] ?? '')); } return out; },
      clearContent() { for (let i = 0; i < nr; i++) { const row = sh.rows[r + i - 1]; if (row) for (let j = 0; j < nc; j++) row[c + j - 1] = ''; } save(); return this; },
      setNumberFormat() { return this; }, setFontWeight() { return this; },
    });
    return {
      getName: () => sh.name, setName(n) { sh.name = n; save(); return this; },
      getLastRow: lastRow, setFrozenRows() {},
      getRange(a, b, c, d) { if (typeof a === 'string') { const [x, y] = a.split(':'); return range(1, colIdx(x), 1, colIdx(y) - colIdx(x) + 1); } return range(a, b, c, d); },
      deleteRows(start, n) { sh.rows.splice(start - 1, n); save(); },
      _rows: () => sh.rows,
    };
  }
  const SSID = 'ss1';
  const ss = {
    getId: () => SSID, getUrl: () => 'https://docs.google.com/spreadsheets/d/' + SSID,
    getSheets: () => { if (!db.sheets.length) { db.sheets.push({ name: 'ชีต1', rows: [] }); save(); } return db.sheets.map(sheetApi); },
    getSheetByName: n => { const s = db.sheets.find(x => x.name === n); return s ? sheetApi(s) : null; },
    insertSheet: n => { const s = { name: n, rows: [] }; db.sheets.push(s); save(); return sheetApi(s); },
  };
  const iter = arr => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
  function folderApi(id) {
    const f = db.folders[id];
    return {
      getId: () => id, getUrl: () => 'https://drive.google.com/drive/folders/' + id,
      createFolder(name) { const k = nid('fo'); db.folders[k] = { name, parent: id }; save(); return folderApi(k); },
      getFoldersByName(name) { return iter(Object.keys(db.folders).filter(k => db.folders[k].parent === id && db.folders[k].name === name).map(folderApi)); },
      createFile(a, content, mime) {
        const k = nid('fi');
        if (typeof a === 'string') db.files[k] = { name: a, data: Buffer.from(content).toString('base64'), parent: id, type: mime };
        else db.files[k] = { name: a.name, data: Buffer.from(a.bytes).toString('base64'), parent: id, type: a.type };
        save(); return fileApi(k);
      },
    };
  }
  function fileApi(id) {
    if (id === SSID) return { getParents: () => iter([folderApi('root')]) };
    const f = db.files[id]; if (!f) throw new Error('file not found ' + id);
    return {
      getId: () => id,
      getBlob: () => ({ getDataAsString: () => Buffer.from(f.data, 'base64').toString('utf8') }),
      setContent(t) { f.data = Buffer.from(t).toString('base64'); save(); return this; },
      setSharing(a, p) { f.sharing = a + '/' + p; save(); return this; },
    };
  }
  db.folders.root = db.folders.root || { name: 'root' };
  const bytesOf = v => Buffer.isBuffer(v) ? v : Buffer.from(typeof v === 'string' ? v : v);
  const ctx = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ({ alert() {}, ButtonSet: {}, createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }) },
    DriveApp: { getFileById: fileApi, getFolderById: id => { if (!db.folders[id]) throw new Error('no folder'); return folderApi(id); }, getRootFolder: () => folderApi('root'), Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => db.props[k] ?? null, setProperty: (k, v) => { db.props[k] = String(v); save(); } }) },
    CacheService: { getScriptCache: () => ({ get: k => { const c = db.cache[k]; return c && c.exp > Date.now() ? c.v : null; }, put: (k, v, s) => { db.cache[k] = { v, exp: Date.now() + s * 1000 }; save(); } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: {},
    Utilities: {
      getUuid: () => crypto.randomUUID(), sleep() {}, Charset: { UTF_8: 'utf8' }, DigestAlgorithm: { SHA_256: 'sha256' }, computeDigest: (alg, v) => [...crypto.createHash(alg).update(String(v), 'utf8').digest()].map(b => b > 127 ? b - 256 : b),
      base64EncodeWebSafe: v => bytesOf(v).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      base64DecodeWebSafe: s => Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64'),
      base64Decode: s => Buffer.from(s, 'base64'),
      computeHmacSha256Signature: (v, k) => crypto.createHmac('sha256', k).update(v).digest(),
      newBlob: (bytes, type, name) => ({ bytes: bytesOf(bytes), type, name, getDataAsString: () => bytesOf(bytes).toString('utf8') }),
      formatDate: (d, tz, f) => { const x = new Date(d.getTime() + 7 * 3600e3); return x.getUTCFullYear() + '-' + String(x.getUTCMonth() + 1).padStart(2, '0'); },
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return { content: s }; } }) },
  };
  vm.createContext(ctx);
  const here = path.dirname(fileURLToPath(import.meta.url));
  vm.runInContext(fs.readFileSync(path.join(here, '..', 'google', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  return { ctx, db, sheet: n => ss.getSheetByName(n), files: () => db.files };
}
