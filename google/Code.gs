/* ===== แผนที่มิเตอร์ไฟฟ้า–น้ำประปา ม.๑ พัน.๓ รอ. — ตัวกลางบน Google Apps Script =====
 วิธีติดตั้ง: ดู "คู่มือตั้งค่าระบบออนไลน์.md"
 ข้อมูล: ชีต "ตัวเลขรายเดือน" (แก้ตัวเลขในชีตได้ตรง ๆ) · ชีต "มิเตอร์" (สำเนาให้อ่าน แก้ในเว็บเท่านั้น) · ชีต "ประวัติ"
         โฟลเดอร์ Drive "meter-map ข้อมูลระบบ" เก็บ config.json (อาคาร แนวท่อ กองร้อย รหัส) และรูปหน้าปัด
 ห้ามแชร์ชีตหรือโฟลเดอร์นี้ให้คนอื่นแก้ไข — config.json มีรหัสของทุกกองร้อย */

var SH_RD = 'ตัวเลขรายเดือน', SH_M = 'มิเตอร์', SH_LOG = 'ประวัติ', SH_LOG_OLD = 'ประวัติเก่า';
var RD_HEAD = ['รหัสมิเตอร์ (id)', 'เดือน (ค.ศ. ปี-เดือน)', 'เลขมิเตอร์', 'รูปหน้าปัด', 'ผู้บันทึก', 'เวลาบันทึก'];
var TOKEN_H = 12, LOG_MAX = 5000;
var PHOTO_RE = /^https:\/\/drive\.google\.com\/thumbnail\?id=[A-Za-z0-9_-]{10,100}&sz=w\d{2,4}$/;
var YM = /^\d{4}-(0[1-9]|1[0-2])$/, SAFE_ID = /^[A-Za-z0-9_-]{1,40}$/;

function props_() { return PropertiesService.getScriptProperties(); }
// รหัสชีต (ส่วนยาวใน URL ระหว่าง /d/ กับ /edit) — ใช้เมื่อสร้างโครงการจากหน้า script.google.com แทนเมนูในชีต
var SHEET_ID = '';
function ss_() { var a = SpreadsheetApp.getActiveSpreadsheet(); return a || SpreadsheetApp.openById(SHEET_ID); }
function err_(status, msg) { var e = new Error(msg); e.status = status; return e; }

/* ---------- เมนูในชีต ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('แผนที่มิเตอร์').addItem('ตั้งค่าเริ่มต้น / ดูรหัสชั่วคราว', 'setup').addToUi();
}
function setup() {
  ensure_();
  var p = props_(), code = p.getProperty('ADMIN_CODE');
  if (!code) { code = String(100000 + Math.floor(Math.random() * 900000)); p.setProperty('ADMIN_CODE', code); }
  var has = !!readConfig_().data;
  var msg = (has ? 'ระบบมีข้อมูลแล้ว ใช้รหัสผู้ดูแลที่ระบบสร้างตอนอัปโหลดข้อมูลตั้งต้น'
         : 'รหัสผู้ดูแลชั่วคราว: ' + code + '\nใช้เข้าสู่ระบบบนเว็บครั้งแรก แล้วกด "อัปโหลดข้อมูลตั้งต้นขึ้นระบบ"') +
    '\n\nขั้นต่อไป: เมนู ทำให้ใช้งานได้ (Deploy) › การทำให้ใช้งานได้รายการใหม่ › เว็บแอป';
  try { SpreadsheetApp.getUi().alert('ตั้งค่าเรียบร้อย', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { console.log('ตั้งค่าเรียบร้อย — ' + msg); } // รันจากหน้าแก้โค้ด: ดูข้อความใน "บันทึกการดำเนินการ"
}

/* ---------- ที่เก็บ ---------- */
function ensure_() {
  var ss = ss_();
  var rd = ss.getSheetByName(SH_RD);
  if (!rd) {
    rd = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(SH_RD) : ss.insertSheet(SH_RD);
    rd.getRange(1, 1, 1, RD_HEAD.length).setValues([RD_HEAD]).setFontWeight('bold');
    rd.getRange('A:B').setNumberFormat('@'); rd.getRange('F:F').setNumberFormat('@'); rd.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SH_M)) { var m = ss.insertSheet(SH_M); m.getRange(1, 1, 1, 7).setValues([['รหัส', 'ชื่อ/ที่ตั้ง', 'ประเภท', 'กองร้อย', 'เลขผู้ใช้ไฟ/น้ำ', 'อาคาร', 'id']]).setFontWeight('bold'); m.setFrozenRows(1); }
  if (!ss.getSheetByName(SH_LOG)) { var l = ss.insertSheet(SH_LOG); l.getRange(1, 1, 1, 3).setValues([['เวลา', 'ผู้ทำ', 'รายการ']]).setFontWeight('bold'); l.getRange('A:A').setNumberFormat('@'); l.setFrozenRows(1); }
  var p = props_(), fid = p.getProperty('FOLDER_ID'), folder = null;
  if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; } }
  if (!folder) {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    folder = (parents.hasNext() ? parents.next() : DriveApp.getRootFolder()).createFolder('meter-map ข้อมูลระบบ');
    p.setProperty('FOLDER_ID', folder.getId());
  }
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  return folder;
}
function folder_() { return ensure_(); }
function photoFolder_() {
  var f = folder_(), it = f.getFoldersByName('รูปหน้าปัด');
  return it.hasNext() ? it.next() : f.createFolder('รูปหน้าปัด');
}
function readConfig_() {
  var p = props_(), id = p.getProperty('CFG_ID');
  if (!id) return { data: null, etag: null };
  var txt = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
  return { data: JSON.parse(txt), etag: p.getProperty('CFG_VER') || '1' };
}
function writeConfig_(cfg) {
  var p = props_(), id = p.getProperty('CFG_ID'), txt = JSON.stringify(cfg);
  if (id) DriveApp.getFileById(id).setContent(txt);
  else p.setProperty('CFG_ID', folder_().createFile('config.json', txt, 'application/json').getId());
  var v = String(Number(p.getProperty('CFG_VER') || '0') + 1); p.setProperty('CFG_VER', v);
  mirrorMeters_(cfg);
  return v;
}
function mirrorMeters_(cfg) {
  var sh = ss_().getSheetByName(SH_M), co = {}, bl = {};
  (cfg.companies || []).forEach(function (c) { co[c.id] = c.name; });
  (cfg.blds || []).forEach(function (b) { bl[b.id] = b.name; });
  var rows = (cfg.meters || []).map(function (m) { return [m.code, (m.name || '') + (m.room ? ' [ห้อง ' + m.room + (m.shop ? ' ' + m.shop : '') + ']' : ''), m.type === 'elec' ? 'ไฟฟ้า' : 'ประปา', co[m.companyId] || '', m.account || '', bl[m.bldId] || '', m.id]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  rows = rows.map(function (r) { return r.map(cell_); });
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
}
function isDate_(v) { return v != null && typeof v === "object" && typeof v.getTime === "function"; }
function ymText_(v) { return isDate_(v) ? Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM') : String(v).trim(); }
function readReadings_() {
  var sh = ss_().getSheetByName(SH_RD), n = sh.getLastRow(), out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 6).getValues().forEach(function (r) {
    var id = String(r[0]).trim(), ym = ymText_(r[1]), v = Number(String(r[2]).replace(/,/g, ''));
    if (!id || !YM.test(ym) || r[2] === '' || !isFinite(v)) return;
    (out[id] = out[id] || {})[ym] = { reading: v, photo: r[3] ? String(r[3]) : null, by: String(r[4] || ''), at: isDate_(r[5]) ? r[5].toISOString() : String(r[5] || '') };
  });
  return out;
}
function writeAllReadings_(map) {
  var sh = ss_().getSheetByName(SH_RD), rows = [];
  Object.keys(map).sort().forEach(function (id) { Object.keys(map[id]).sort().forEach(function (ym) { var r = map[id][ym]; rows.push([id, ym, r.reading, r.photo || '', cell_(r.by || ''), r.at || '']); }); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) { sh.getRange(2, 1, rows.length, 2).setNumberFormat('@'); sh.getRange(2, 6, rows.length, 1).setNumberFormat('@'); sh.getRange(2, 1, rows.length, 6).setValues(rows); }
}
function upsertReading_(id, ym, rec) {
  var sh = ss_().getSheetByName(SH_RD), n = sh.getLastRow(), row = -1;
  if (n > 1) { var keys = sh.getRange(2, 1, n - 1, 2).getValues(); for (var i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === id && ymText_(keys[i][1]) === ym) { row = i + 2; break; } }
  if (row < 0) row = n + 1;
  sh.getRange(row, 1, 1, 2).setNumberFormat('@'); sh.getRange(row, 6).setNumberFormat('@');
  sh.getRange(row, 1, 1, 6).setValues([[id, ym, rec.reading, rec.photo || '', cell_(rec.by || ''), rec.at || '']]);
}
// ข้อความจากผู้ใช้ที่ขึ้นต้นด้วย = + - @ จะถูกชีตตีเป็นสูตร → ใส่ ' นำหน้าให้เป็นข้อความ
function cell_(v) { return typeof v === 'string' && /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }
function addLog_(who, texts) {
  texts = (Array.isArray(texts) ? texts : [texts]).filter(function (t) { return typeof t === 'string' && t.trim(); }).slice(0, 50);
  if (!texts.length) return;
  var sh = ss_().getSheetByName(SH_LOG), at = new Date().toISOString();
  var rows = texts.map(function (t) { return [at, cell_(who), cell_(t.slice(0, 300))]; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  if (sh.getLastRow() > LOG_MAX + 500) { // ย้ายแถวเก่าไปชีต "ประวัติเก่า" แทนการลบทิ้ง (เก็บร่องรอยตรวจสอบ)
    var ss = ss_(), ar = ss.getSheetByName(SH_LOG_OLD);
    if (!ar) { ar = ss.insertSheet(SH_LOG_OLD); ar.getRange(1, 1, 1, 3).setValues([['เวลา', 'ผู้ทำ', 'รายการ']]); }
    ar.getRange(ar.getLastRow() + 1, 1, 500, 3).setValues(sh.getRange(2, 1, 500, 3).getValues());
    sh.deleteRows(2, 500);
  }
}
function readLog_() {
  var sh = ss_().getSheetByName(SH_LOG), n = sh.getLastRow();
  if (n < 2) return [];
  var k = Math.min(500, n - 1);
  return sh.getRange(n - k + 1, 1, k, 3).getValues().reverse().map(function (r) { return { at: String(r[0]), who: String(r[1]), text: String(r[2]) }; });
}

/* ---------- บัญชีผู้ใช้ (สมัครเอง ผู้ดูแลอนุมัติ) ----------
 เก็บใน users.json ในโฟลเดอร์ข้อมูลระบบ (ไม่อยู่ใน config จึงไม่ถูกเขียนทับตอนผู้ดูแลบันทึกแผนที่)
 รหัสผ่านเก็บเป็นค่าแฮช SHA-256 แบบใส่เกลือ วนซ้ำ ไม่มีใครอ่านรหัสจริงได้ */
var USER_RE = /^[a-z0-9._-]{3,30}$/, PW_ROUNDS = 300, REG_MAX = 300;
function readUsers_() {
  var id = props_().getProperty('USERS_ID');
  if (!id) return [];
  return JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')) || [];
}
function writeUsers_(list) {
  var p = props_(), id = p.getProperty('USERS_ID'), txt = JSON.stringify(list);
  if (id) DriveApp.getFileById(id).setContent(txt);
  else p.setProperty('USERS_ID', folder_().createFile('users.json', txt, 'application/json').getId());
}
function hex_(bytes) { var o = ''; for (var i = 0; i < bytes.length; i++) o += ('0' + (bytes[i] & 255).toString(16)).slice(-2); return o; }
function hashPw_(salt, pw) {
  var h = salt + '|' + pw;
  for (var i = 0; i < PW_ROUNDS; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  return h;
}
function checkPw_(pw) { pw = String(pw || ''); if (pw.length < 6 || pw.length > 100) throw err_(400, 'รหัสผ่านต้องยาว 6 ตัวขึ้นไป'); return pw; }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function coName_(c, id) { return (((c && c.companies) || []).filter(function (x) { return x.id === id; })[0] || {}).name || ''; }
function pubUser_(u) { return { id: u.id, user: u.user, name: u.name, pos: u.pos, phone: u.phone, reqCo: u.reqCo, note: u.note || '', status: u.status, role: u.role || null, co: u.co || null, at: u.at, lastLogin: u.lastLogin || null }; }
function byName_(s, c) {
  if (s.n) return s.r === 'admin' ? s.n + ' (ผู้ดูแล)' : s.n + (coName_(c, s.c) ? ' (' + coName_(c, s.c) + ')' : '');
  return s.r === 'admin' ? 'ผู้ดูแลระบบ' : coName_(c, s.c) || 'กองร้อย';
}

/* ---------- โทเค็น ---------- */
function sig_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, props_().getProperty('SECRET'))); }
function sign_(o) { var p = Utilities.base64EncodeWebSafe(JSON.stringify(o), Utilities.Charset.UTF_8); return p + '.' + sig_(p); }
function same_(x, y) { x = String(x); y = String(y); if (x.length !== y.length) return false; var d = 0; for (var i = 0; i < x.length; i++) d |= x.charCodeAt(i) ^ y.charCodeAt(i); return d === 0; }
function verify_(t) {
  if (!t || typeof t !== 'string' || t.indexOf('.') < 0) return null;
  var a = t.split('.'); if (a.length !== 2 || !same_(sig_(a[0]), a[1])) return null;
  try { var o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(a[0])).getDataAsString('UTF-8')); return o.exp > Date.now() ? o : null; } catch (e) { return null; }
}
/* ---------- รหัสเข้าดูเว็บไซต์ ----------
 ผู้ดูแลตั้ง settings.viewCode ในหน้าเว็บ (ไม่ถูกส่งออกไปกับข้อมูลสาธารณะ) เมื่อตั้งแล้ว GET state ต้องมีโทเค็น
 โทเค็นดูอย่างเดียว {r:'view', k} ผูกกับรหัสปัจจุบัน — เปลี่ยนรหัสเมื่อไร โทเค็นเก่าใช้ไม่ได้ทันที */
var VIEW_DAYS = 30;
function viewKey_(c) { var v = c && c.settings && c.settings.viewCode; return v ? sig_('view|' + v).slice(0, 16) : null; }
function viewTok_(c) { var k = viewKey_(c); return k ? sign_({ r: 'view', k: k, exp: Date.now() + VIEW_DAYS * 864e5 }) : null; }
// โทเค็นจากรหัส 6 หลัก (ผู้ดูแล/กองร้อย) ผูกกับรหัสปัจจุบัน — ผู้ดูแลเปลี่ยนรหัสเมื่อไร โทเค็นเก่าใช้ไม่ได้ทันที
function codeVer_(s, c) {
  var code = s.r === 'admin' ? ((c && c.settings && c.settings.adminCode) || props_().getProperty('ADMIN_CODE'))
    : (((c && c.companies) || []).filter(function (x) { return x.id === s.c; })[0] || {}).code;
  return code ? sig_('cv|' + s.r + '|' + code).slice(0, 12) : null;
}
function who_(t, need) {
  var s = verify_(t);
  if (!s) throw err_(401, 'ต้องเข้าสู่ระบบใหม่');
  if (!s.u && s.r !== 'view' && (!s.cv || s.cv !== codeVer_(s, readConfig_().data))) throw err_(401, 'รหัสถูกเปลี่ยนแล้ว กรุณาเข้าสู่ระบบใหม่');
  if (s.r === 'view') { // โทเค็นดูอย่างเดียว ใช้ได้แค่ดูหน้าแสดงผล
    if (need !== 'view') throw err_(401, 'ต้องเข้าสู่ระบบ');
    if (s.k !== viewKey_(readConfig_().data)) throw err_(401, 'รหัสเข้าดูถูกเปลี่ยนแล้ว ใส่รหัสใหม่');
    return s;
  }
  if (s.u) { // บัญชีผู้ใช้: ใช้สิทธิ์ล่าสุดเสมอ ผู้ดูแลระงับหรือเปลี่ยนสิทธิ์แล้วมีผลทันที
    var u = readUsers_().filter(function (x) { return x.id === s.u; })[0];
    if (!u || u.status !== 'active' || u.v !== s.v) throw err_(401, 'สิทธิ์ของบัญชีนี้ถูกเปลี่ยน กรุณาเข้าสู่ระบบใหม่');
    s = { r: u.role, c: u.co, u: u.id, n: u.name, exp: s.exp };
  }
  if (need === 'admin' && s.r !== 'admin') throw err_(403, 'เฉพาะผู้ดูแลระบบ');
  return s;
}

/* ---------- ตัวช่วย ---------- */
function sanitize_(c) {
  if (!c) return null;
  var x = JSON.parse(JSON.stringify(c));
  x.companies = (x.companies || []).map(function (co) { delete co.code; return co; });
  if (x.settings) { delete x.settings.adminCode; delete x.settings.viewCode; }
  return x;
}
function prevYm_(ym) { var a = ym.split('-').map(Number), y = a[0], m = a[1] - 1; if (!m) { m = 12; y--; } return y + '-' + ('0' + m).slice(-2); }
function openYm_(c) { return (c && c.settings && c.settings.openYm) || prevYm_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM')); }
function savePhoto_(meterId, ym, dataUrl) {
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/.exec(dataUrl || '');
  if (!m) throw err_(400, 'ไฟล์รูปไม่ถูกต้อง');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 3e6) throw err_(413, 'รูปใหญ่เกินไป');
  var f = photoFolder_().createFile(Utilities.newBlob(bytes, 'image/' + m[1], meterId + '_' + ym + '_' + Date.now().toString(36) + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1])));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600';
}
function cleanRec_(meterId, ym, rec, byName) {
  if (rec == null) return null;
  var n = Number(rec.reading);
  if (!isFinite(n) || n < 0 || n > 1e12) throw err_(400, 'เลขมิเตอร์ไม่ถูกต้อง');
  var photo = rec.photo || null;
  if (photo && photo.indexOf('data:') === 0) photo = savePhoto_(meterId, ym, photo);
  else if (photo && !PHOTO_RE.test(photo)) photo = null;
  return { reading: n, photo: photo, by: byName != null ? byName : String(rec.by || '').slice(0, 60), at: rec.at && byName == null ? String(rec.at).slice(0, 30) : new Date().toISOString() };
}
function locked_(fn) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
// กันเดารหัส: นับแยกตามเป้าหมาย (บัญชีผู้ใช้ / รหัส 6 หลัก / รหัสเข้าดู) + เพดานรวมทั้งระบบ · พัก 10 นาที
var FAIL_MAX = { user: 10, code: 30, view: 30, all: 100 };
function throttle_(ok, key) {
  var c = CacheService.getScriptCache(), kind = String(key || 'code').split(':')[0];
  var nAll = Number(c.get('fails') || 0), nKey = Number(c.get('fails:' + key) || 0);
  if (ok === undefined) {
    if (nAll >= FAIL_MAX.all || nKey >= (FAIL_MAX[kind] || FAIL_MAX.code)) throw err_(429, 'มีการใส่รหัสผิดจำนวนมาก ระบบพักการเข้าสู่ระบบ 10 นาที');
    return;
  }
  if (!ok) { c.put('fails', String(nAll + 1), 600); c.put('fails:' + key, String(nKey + 1), 600); }
}

/* ---------- ตลาด: มิเตอร์รายห้อง ----------
 มิเตอร์ที่ใส่ "ห้อง" (m.room) คือมิเตอร์ย่อยของร้านค้าหรือผู้ประกอบการ
 คนทั่วไปได้แค่ยอดรวมต่ออาคาร (ไม่เห็นเลขห้อง ชื่อร้าน หรือยอดรายห้อง)
 ผู้กรอกเห็นเฉพาะห้องของกองร้อยตัวเอง ผู้ดูแลเห็นทั้งหมด */
function isRoom_(m) { return !!(m && m.room); }
// มิเตอร์วนรอบ (เลขลดลงและครั้งก่อนใกล้เพดาน 999…9) คิดแบบหน่วย: (999…9 − ครั้งก่อน) + ครั้งนี้ — ต้องตรงกับ rollU() ในหน้าเว็บ
function rollU_(p, c) { if (c >= p) return c - p; var max = Math.pow(10, String(Math.floor(p)).length) - 1; return p >= max * 0.9 ? max - p + c : c - p; }
function units_(r, ym) { var a = r && r[ym], b = r && r[prevYm_(ym)]; return a && b ? rollU_(b.reading, a.reading) : null; }
function peek_(t) { if (!t) return null; try { return who_(t, 'view'); } catch (e) { return null; } }
function canRoom_(s, m) { return !!s && (s.r === 'admin' || s.c === m.companyId); }
function publicView_(cfg, rd, s) {
  var x = sanitize_(cfg), keep = [], groups = {}, order = [], out = {};
  (x.meters || []).forEach(function (m) {
    if (!isRoom_(m) || canRoom_(s, m)) { keep.push(m); if (rd[m.id]) out[m.id] = rd[m.id]; return; }
    var k = (m.bldId || '') + '|' + m.type + '|' + m.companyId, g = groups[k];
    if (!g) { g = groups[k] = { id: ('agg-' + (m.bldId || 'x') + '-' + m.type + '-' + m.companyId).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40), code: 'รวม', name: '', type: m.type, companyId: m.companyId, bldId: m.bldId || '', lng: 0, lat: 0, agg: true, n: 0, u: {} }; order.push(k); }
    g.n++; g.lng += Number(m.lng) || 0; g.lat += Number(m.lat) || 0;
    var r = rd[m.id] || {};
    Object.keys(r).forEach(function (ym) { var u = units_(r, ym); if (u != null) g.u[ym] = Math.round(((g.u[ym] || 0) + u) * 1000) / 1000; });
  });
  order.forEach(function (k) { var g = groups[k]; g.lng /= g.n; g.lat /= g.n; g.name = 'รวม ' + g.n + ' ห้อง'; keep.push(g); });
  x.meters = keep;
  return { config: x, readings: out };
}
// รูปตำแหน่งที่ติดมิเตอร์ (จุดละ 1 รูป) เก็บแยกใน locphotos.json ไม่ปนกับ config จึงไม่ชนกับการบันทึกแผนที่ของผู้ดูแล
function readLoc_() {
  var id = props_().getProperty('LOC_ID');
  if (!id) return {};
  return JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')) || {};
}
function writeLoc_(o) {
  var p = props_(), id = p.getProperty('LOC_ID'), txt = JSON.stringify(o);
  if (id) DriveApp.getFileById(id).setContent(txt);
  else p.setProperty('LOC_ID', folder_().createFile('locphotos.json', txt, 'application/json').getId());
}
// ผลตรวจการรั่วไหลรายเดือน (เดือนที่ระบบเตือนว่าใช้มากผิดปกติ) เก็บแยกใน checks.json (พร็อพ CHK_ID)
// คีย์ = รหัสมิเตอร์ หรือ "b:<อาคาร>:<elec|water>" (ยอดรวมห้องในอาคาร) → { ym: { s: 'ok'|'leak', note, by, at } }
var CHK_KEY = /^(?:[A-Za-z0-9_-]{1,40}|b:[A-Za-z0-9_-]{1,40}:(?:elec|water))$/;
function readChecks_() {
  var id = props_().getProperty('CHK_ID');
  if (!id) return {};
  return JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')) || {};
}
function writeChecks_(o) {
  var p = props_(), id = p.getProperty('CHK_ID'), txt = JSON.stringify(o);
  if (id) DriveApp.getFileById(id).setContent(txt);
  else p.setProperty('CHK_ID', folder_().createFile('checks.json', txt, 'application/json').getId());
}
function pickChecks_(meters) {
  var all = readChecks_(), ok = {}, o = {};
  (meters || []).forEach(function (m) { if (!m.agg) ok[m.id] = 1; });
  Object.keys(all).forEach(function (k) { if (k.indexOf('b:') === 0 || ok[k]) o[k] = all[k]; });
  return o;
}
function pickLoc_(meters) { var all = readLoc_(), o = {}; (meters || []).forEach(function (m) { if (all[m.id]) o[m.id] = all[m.id]; }); return o; }

/* ---------- เส้นทาง ---------- */
var ROUTES = {
  'GET ping': function () { return { app: 'meter-map', storage: 'google', connected: true, account: null, gate: !!viewKey_(readConfig_().data) }; },
  'GET state': function (b, t) {
    var c = readConfig_().data;
    if (!c) return { config: null, readings: {}, locPhotos: {} };
    var who = peek_(t);
    if (viewKey_(c) && !who) throw err_(401, 'ใส่รหัสเข้าดูเว็บไซต์');
    var v = publicView_(c, readReadings_(), who);
    v.locPhotos = pickLoc_(v.config.meters);
    v.checks = pickChecks_(v.config.meters);
    return v;
  },
  'POST register': function (b) {
    b = b || {};
    var c = CacheService.getScriptCache(), n = Number(c.get('regs') || 0);
    if (n >= 20) throw err_(429, 'มีการสมัครจำนวนมาก ลองใหม่ในอีก 10 นาที');
    var user = str_(b.user, 30).toLowerCase(), name = str_(b.name, 80), pos = str_(b.pos, 80), phone = str_(b.phone, 20), note = str_(b.note, 200), reqCo = str_(b.co, 40);
    if (!USER_RE.test(user)) throw err_(400, 'ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 . _ - ยาว 3–30 ตัว');
    if (name.length < 2) throw err_(400, 'ใส่ยศ ชื่อ สกุล');
    var pw = checkPw_(b.pass), cfg = readConfig_().data;
    if (reqCo && reqCo !== 'admin' && !coName_(cfg, reqCo)) throw err_(400, 'ไม่พบกองร้อยที่เลือก');
    locked_(function () {
      var list = readUsers_();
      if (list.some(function (x) { return x.user === user; })) throw err_(409, 'ชื่อผู้ใช้นี้มีคนใช้แล้ว');
      if (list.filter(function (x) { return x.status === 'pending'; }).length >= 50 || list.length >= REG_MAX) throw err_(429, 'คำขอสมัครเต็ม ติดต่อผู้ดูแลระบบ');
      var salt = Utilities.getUuid();
      list.push({ id: 'u' + Date.now().toString(36) + Math.floor(Math.random() * 1e4), user: user, name: name, pos: pos, phone: phone, note: note, reqCo: reqCo, status: 'pending', role: null, co: null, salt: salt, hash: hashPw_(salt, pw), v: 1, at: new Date().toISOString() });
      writeUsers_(list);
    });
    c.put('regs', String(n + 1), 600);
    addLog_(name, 'สมัครใช้งาน (ชื่อผู้ใช้ ' + user + ') ขอสิทธิ์ ' + (reqCo === 'admin' ? 'ผู้ดูแลระบบ' : coName_(cfg, reqCo) || '-'));
    return { ok: true };
  },
  'POST login': function (b) {
    if (b && b.user != null) {
      var un = str_(b.user, 30).toLowerCase(), pw = String(b.pass || '');
      throttle_(undefined, 'user:' + un);
      var u = readUsers_().filter(function (x) { return x.user === un; })[0];
      var good = !!u && same_(u.hash, hashPw_(u.salt, pw));
      throttle_(good, 'user:' + un);
      if (!good) { Utilities.sleep(400); throw err_(401, 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'); }
      if (u.status === 'pending') throw err_(403, 'ส่งคำขอแล้ว รอผู้ดูแลระบบอนุมัติ');
      if (u.status !== 'active') throw err_(403, 'บัญชีนี้ถูกระงับ ติดต่อผู้ดูแลระบบ');
      if (u.role === 'company' && !coName_(readConfig_().data, u.co)) throw err_(403, 'ยังไม่ได้กำหนดกองร้อยให้บัญชีนี้ ติดต่อผู้ดูแลระบบ');
      locked_(function () { var l2 = readUsers_(); l2.forEach(function (x) { if (x.id === u.id) x.lastLogin = new Date().toISOString(); }); writeUsers_(l2); });
      return { token: sign_({ u: u.id, v: u.v, exp: Date.now() + TOKEN_H * 3600e3 }), role: u.role, co: u.co || null, name: u.name, user: u.user, view: viewTok_(readConfig_().data) };
    }
    if (b && b.view != null) {
      throttle_(undefined, 'view');
      var cv = readConfig_().data, want = cv && cv.settings && cv.settings.viewCode, gv = !!want && same_(String(b.view).trim(), String(want));
      throttle_(gv, 'view');
      if (!gv) { Utilities.sleep(400); throw err_(401, want ? 'รหัสเข้าดูไม่ถูกต้อง' : 'ระบบนี้ไม่ได้ตั้งรหัสเข้าดู'); }
      return { token: viewTok_(cv), role: 'view' };
    }
    var code = String((b && b.code) || '').trim();
    if (!code) throw err_(400, 'ใส่รหัส');
    throttle_(undefined, 'code');
    var c = readConfig_().data, adminCode = (c && c.settings && c.settings.adminCode) || props_().getProperty('ADMIN_CODE'), s = null;
    if (adminCode && same_(code, String(adminCode))) s = { r: 'admin' };
    else { var co = ((c && c.companies) || []).filter(function (x) { return x.code && same_(String(x.code), code); })[0]; if (co) s = { r: 'company', c: co.id }; }
    throttle_(!!s, 'code');
    if (!s) { Utilities.sleep(400); throw err_(401, 'รหัสไม่ถูกต้อง'); }
    s.cv = codeVer_(s, c);
    s.exp = Date.now() + TOKEN_H * 3600e3;
    return { token: sign_(s), role: s.r === 'admin' ? 'admin' : 'company', co: s.c || null, view: viewTok_(c) };
  },
  'GET admin/state': function (b, t) {
    who_(t, 'admin');
    var c = readConfig_();
    return { config: c.data, etag: c.etag, readings: readReadings_(), log: readLog_(), locPhotos: readLoc_(), checks: readChecks_(), sheetUrl: ss_().getUrl(), folderUrl: folder_().getUrl() };
  },
  'GET admin/users': function (b, t) {
    who_(t, 'admin');
    return { users: readUsers_().map(pubUser_) };
  },
  'POST admin/user': function (b, t) {
    var me = who_(t, 'admin'); b = b || {};
    var op = String(b.op || ''), cfg = readConfig_().data, out = null;
    locked_(function () {
      var list = readUsers_(), u = list.filter(function (x) { return x.id === b.id; })[0];
      if (!u) throw err_(404, 'ไม่พบบัญชี');
      if (me.u === u.id) throw err_(400, 'แก้สิทธิ์ของบัญชีตัวเองไม่ได้ ให้ผู้ดูแลคนอื่นทำ หรือใช้เมนูเปลี่ยนรหัสผ่านของฉัน');
      if (op === 'approve' || op === 'save') {
        var role = b.role === 'admin' ? 'admin' : 'company', co = role === 'company' ? String(b.co || '') : null;
        if (role === 'company' && !coName_(cfg, co)) throw err_(400, 'เลือกกองร้อยให้บัญชีนี้');
        if (u.role !== role || u.co !== co || u.status !== 'active') u.v++;
        u.role = role; u.co = co; u.status = 'active';
        out = 'ให้สิทธิ์ ' + u.name + ' (' + u.user + ') เป็น ' + (role === 'admin' ? 'ผู้ดูแลระบบ' : 'ผู้กรอกข้อมูล ' + coName_(cfg, co));
      } else if (op === 'disable') { u.status = 'disabled'; u.v++; out = 'ระงับบัญชี ' + u.name + ' (' + u.user + ')'; }
      else if (op === 'reset') { var pw = checkPw_(b.pass); u.salt = Utilities.getUuid(); u.hash = hashPw_(u.salt, pw); u.v++; out = 'ตั้งรหัสผ่านใหม่ให้ ' + u.name + ' (' + u.user + ')'; }
      else if (op === 'delete') { list = list.filter(function (x) { return x.id !== u.id; }); out = 'ลบบัญชี ' + u.name + ' (' + u.user + ')'; }
      else throw err_(400, 'คำสั่งไม่ถูกต้อง');
      writeUsers_(list);
    });
    addLog_(byName_(me, cfg), out);
    return { users: readUsers_().map(pubUser_) };
  },
  'POST me/password': function (b, t) {
    var s = who_(t); b = b || {};
    if (!s.u) throw err_(400, 'เปลี่ยนได้เฉพาะบัญชีผู้ใช้');
    var pw = checkPw_(b.pass), nv = null;
    locked_(function () {
      var list = readUsers_(), u = list.filter(function (x) { return x.id === s.u; })[0];
      if (!u || u.hash !== hashPw_(u.salt, String(b.old || ''))) throw err_(403, 'รหัสผ่านเดิมไม่ถูกต้อง');
      u.salt = Utilities.getUuid(); u.hash = hashPw_(u.salt, pw); u.v++; nv = u.v; writeUsers_(list);
    });
    addLog_(byName_(s, readConfig_().data), 'เปลี่ยนรหัสผ่านของตัวเอง');
    return { token: sign_({ u: s.u, v: nv, exp: Date.now() + TOKEN_H * 3600e3 }) };
  },
  'PUT admin/config': function (b, t) {
    who_(t, 'admin');
    var config = b && b.config;
    if (!config || !Array.isArray(config.companies) || !Array.isArray(config.meters) || !config.settings) throw err_(400, 'ข้อมูลไม่ครบ');
    var vc = config.settings.viewCode;
    if (vc != null && (typeof vc !== 'string' || vc.trim().length < 6 || vc.length > 40)) throw err_(400, 'รหัสเข้าดูต้องยาว 6–40 ตัว');
    // รหัสผู้ดูแล/กองร้อย: อย่างน้อย 6 ตัว ไม่เกิน 40 และห้ามซ้ำกัน · สีกองร้อยต้องเป็นรหัสสี #xxxxxx
    var seen = {}, ac = config.settings.adminCode;
    if (ac != null && ac !== '') { ac = String(ac).trim(); if (ac.length < 6 || ac.length > 40) throw err_(400, 'รหัสผู้ดูแลต้องยาว 6–40 ตัว'); seen[ac] = 1; }
    config.companies.forEach(function (co) {
      if (co && co.code != null && co.code !== '') {
        var k = String(co.code).trim();
        if (k.length < 6 || k.length > 40) throw err_(400, 'รหัสของ ' + str_(co.name, 40) + ' ต้องยาว 6–40 ตัว');
        if (seen[k]) throw err_(400, 'รหัสของ ' + str_(co.name, 40) + ' ซ้ำกับรหัสอื่น');
        seen[k] = 1;
      }
      if (co && !/^#[0-9a-fA-F]{3,8}$/.test(String(co.color || ''))) co.color = '#888888';
    });
    return locked_(function () {
      var cur = readConfig_();
      if (cur.data && !b.etag) throw err_(409, 'มีข้อมูลอยู่แล้ว');
      if (cur.data && String(b.etag) !== String(cur.etag)) throw err_(409, 'มีการแก้ไขจากที่อื่นระหว่างนี้');
      return { etag: writeConfig_(config) };
    });
  },
  'POST admin/readings': function (b, t) {
    who_(t, 'admin');
    var ch = b && b.changes;
    if (!Array.isArray(ch) || ch.length > 20000) throw err_(400, 'ข้อมูลไม่ถูกต้อง');
    var clean = ch.map(function (x) {
      if (!SAFE_ID.test((x && x.meterId) || '') || !YM.test((x && x.ym) || '')) throw err_(400, 'รหัสมิเตอร์หรือเดือนไม่ถูกต้อง');
      return { meterId: x.meterId, ym: x.ym, rec: cleanRec_(x.meterId, x.ym, x.rec, null) };
    });
    locked_(function () {
      var rd = readReadings_();
      clean.forEach(function (x) { if (x.rec) (rd[x.meterId] = rd[x.meterId] || {})[x.ym] = x.rec; else if (rd[x.meterId]) { delete rd[x.meterId][x.ym]; if (!Object.keys(rd[x.meterId]).length) delete rd[x.meterId]; } });
      writeAllReadings_(rd);
    });
    var photos = {}; clean.forEach(function (x) { if (x.rec && x.rec.photo) (photos[x.meterId] = photos[x.meterId] || {})[x.ym] = x.rec.photo; });
    return { ok: true, n: clean.length, photos: photos };
  },
  'POST reading': function (b, t) {
    var s = who_(t);
    b = b || {};
    if (!SAFE_ID.test(b.meterId || '') || !YM.test(b.ym || '')) throw err_(400, 'ข้อมูลไม่ถูกต้อง');
    var c = readConfig_().data;
    var m = ((c && c.meters) || []).filter(function (x) { return x.id === b.meterId; })[0];
    if (!m) throw err_(404, 'ไม่พบมิเตอร์');
    var byName = byName_(s, c);
    if (s.r !== 'admin') {
      if (m.companyId !== s.c) throw err_(403, 'ไม่มีสิทธิ์กรอกมิเตอร์นี้');
      if (c.settings && c.settings.closed) throw err_(403, 'ปิดรับข้อมูลแล้ว');
      if (b.ym !== openYm_(c)) throw err_(403, 'เดือนนี้ไม่ได้เปิดให้กรอก');
    }
    var rec0 = cleanRec_(b.meterId, b.ym, { reading: b.reading, photo: b.photo || null }, byName), prev = null;
    var rec = locked_(function () {
      var rd = readReadings_(); prev = (rd[b.meterId] || {})[b.ym] || null;
      var r = { reading: rec0.reading, photo: rec0.photo || (prev && prev.photo) || null, by: rec0.by, at: rec0.at };
      upsertReading_(b.meterId, b.ym, r); return r;
    });
    addLog_(byName, (prev ? 'แก้' : 'กรอก') + 'เลข ' + m.code + ' เดือน ' + b.ym + ' = ' + rec.reading);
    return { rec: rec };
  },
  'POST meter/locphoto': function (b, t) {
    var s = who_(t); b = b || {};
    if (!SAFE_ID.test(b.meterId || '')) throw err_(400, 'ข้อมูลไม่ถูกต้อง');
    var c = readConfig_().data, m = ((c && c.meters) || []).filter(function (x) { return x.id === b.meterId; })[0];
    if (!m) throw err_(404, 'ไม่พบมิเตอร์');
    if (s.r !== 'admin' && m.companyId !== s.c) throw err_(403, 'ไม่มีสิทธิ์แก้รูปของมิเตอร์นี้');
    var url = b.photo ? savePhoto_(b.meterId, 'loc', b.photo) : null;
    locked_(function () { var o = readLoc_(); if (url) o[b.meterId] = url; else delete o[b.meterId]; writeLoc_(o); });
    addLog_(byName_(s, c), (url ? 'ใส่รูปตำแหน่ง ' : 'ลบรูปตำแหน่ง ') + m.code);
    return { url: url };
  },
  'POST check': function (b, t) {
    var s = who_(t); b = b || {};
    var key = String(b.key || ''), ym = String(b.ym || '');
    if (!CHK_KEY.test(key) || !YM.test(ym)) throw err_(400, 'ข้อมูลไม่ถูกต้อง');
    var c = readConfig_().data, owner = null, label = '';
    if (key.indexOf('b:') === 0) {
      var bid = key.split(':')[1], bl = ((c && c.blds) || []).filter(function (x) { return x.id === bid; })[0];
      if (!bl) throw err_(404, 'ไม่พบอาคาร');
      owner = bl.companyId; label = bl.name + ' (รวมห้อง ' + (key.split(':')[2] === 'elec' ? 'ไฟฟ้า' : 'ประปา') + ')';
    } else {
      var m = ((c && c.meters) || []).filter(function (x) { return x.id === key; })[0];
      if (!m) throw err_(404, 'ไม่พบมิเตอร์');
      owner = m.companyId; label = m.code;
    }
    if (s.r !== 'admin' && owner !== s.c) throw err_(403, 'ไม่มีสิทธิ์บันทึกผลตรวจของจุดนี้');
    var st = b.s === 'ok' || b.s === 'leak' ? b.s : null, rec = null, who = byName_(s, c);
    if (st) rec = { s: st, note: str_(b.note, 200), by: who, at: new Date().toISOString() };
    locked_(function () { var o = readChecks_(); o[key] = o[key] || {}; if (rec) o[key][ym] = rec; else { delete o[key][ym]; if (!Object.keys(o[key]).length) delete o[key]; } writeChecks_(o); });
    addLog_(who, rec ? 'บันทึกผลตรวจรั่ว ' + label + ' เดือน ' + ym + ': ' + (st === 'ok' ? 'ไม่พบรั่ว' : 'พบรั่ว/แจ้งซ่อม') + (rec.note ? ' — ' + rec.note : '') : 'ลบผลตรวจรั่ว ' + label + ' เดือน ' + ym);
    return { rec: rec };
  },
  'POST log': function (b, t) {
    var s = who_(t), c = readConfig_().data;
    addLog_(byName_(s, c), b && b.entries);
    return { ok: true };
  },
};

function handle_(method, route, body, token) {
  try {
    var fn = ROUTES[method + ' ' + route];
    if (!fn) throw err_(404, 'ไม่พบ');
    ensure_();
    return { ok: true, data: fn(body, token) };
  } catch (e) {
    var st = e.status || 500;
    if (st === 500) console.error(e && e.stack || e);
    return { ok: false, status: st, error: st === 500 ? 'ระบบขัดข้อง กรุณาลองใหม่ (เวลา ' + new Date().toISOString() + ')' : e.message };
  }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// เว็บส่งทุกคำขอเป็น POST แบบ text/plain: {r: เส้นทาง, m: GET|POST|PUT, t: โทเค็น, b: ข้อมูล}
/* ด่านตัวส่งต่อ: เว็บจริงอยู่หลัง Cloudflare Access แล้วส่งคำขอต่อมาที่นี่พร้อมกุญแจ px
 ตั้งกุญแจที่ ตั้งค่าโครงการ › พร็อพเพอร์ตีของสคริปต์ › PROXY_KEY (ค่าเดียวกับ PROXY_KEY ใน Cloudflare)
 เมื่อตั้งแล้ว ใครเรียกลิงก์ /exec ตรง ๆ โดยไม่ผ่านด่านจะได้ 403 ทุกเส้นทาง · ลบพร็อพเพอร์ตีนี้ = กลับไปเปิดตรงเหมือนเดิม */
function proxyOk_(px) { var k = props_().getProperty('PROXY_KEY'); return !k || (typeof px === 'string' && same_(px, k)); }
var DENY_ = { ok: false, status: 403, error: 'ต้องเข้าผ่านเว็บของหน่วยเท่านั้น' };
function doPost(e) {
  var q = {}; try { q = JSON.parse(e.postData.contents); } catch (x) {}
  if (!proxyOk_(q.px)) return json_(DENY_);
  return json_(handle_(q.m || 'POST', String(q.r || ''), q.b || null, q.t || ''));
}
function doGet(e) {
  var p = e.parameter || {};
  if (!proxyOk_(p.px)) return json_(DENY_);
  return json_(handle_('GET', String(p.r || 'ping'), null, ''));
}
