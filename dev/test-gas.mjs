// ทดสอบ google/Code.gs ด้วย Google จำลอง: node dev/test-gas.mjs
import assert from 'node:assert/strict';
import { loadGas } from './gas-mock.mjs';

const g = loadGas();
const call = (r, b, t, m = 'POST') => JSON.parse(g.ctx.doPost({ postData: { contents: JSON.stringify({ r, m, t, b }) } }).content);
const get = r => JSON.parse(g.ctx.doGet({ parameter: { r } }).content);
const pad = x => String(x).padStart(2, '0');
const d = new Date(Date.now() + 7 * 3600e3); let y = d.getUTCFullYear(), mo = d.getUTCMonth(); if (!mo) { mo = 12; y--; } const OPEN = y + '-' + pad(mo);
let ok = 0; const t = async (name, fn) => { await fn(); ok++; console.log('✓', name); };

g.ctx.setup();
const TEMP = g.db.props.ADMIN_CODE;
let A, C1, ET;
await t('setup creates sheets + temp code', () => { assert.ok(g.sheet('ตัวเลขรายเดือน') && g.sheet('มิเตอร์') && g.sheet('ประวัติ')); assert.match(TEMP, /^\d{6}$/); });
await t('ping', () => { const r = get('ping'); assert.equal(r.data.app, 'meter-map'); assert.equal(r.data.storage, 'google'); assert.equal(r.data.account, null); });
await t('empty state', () => { assert.equal(get('state').data.config, null); });
await t('temp admin login', () => { const r = call('login', { code: TEMP }); assert.ok(r.ok, JSON.stringify(r)); A = r.data.token; });
const cfg = { v: 3, companies: [{ id: 'c1', name: 'ร้อย.ก', color: '#f00', code: '111111' }, { id: 'c2', name: 'ร้อย.ข', color: '#0f0', code: '222222' }],
  meters: [{ id: 'E01', code: 'E01', type: 'elec', companyId: 'c1', name: 'ไฟ ก' }, { id: 'W02', code: 'W02', type: 'water', companyId: 'c2' }], blds: [], shapes: [], settings: { adminCode: '888888', openYm: null, closed: false } };
await t('first config upload + meters mirrored to sheet', () => { const r = call('admin/config', { config: cfg, etag: null }, A, 'PUT'); assert.ok(r.ok, JSON.stringify(r)); ET = r.data.etag; assert.equal(g.sheet('มิเตอร์')._rows()[1][0], 'E01'); A = call('login', { code: '888888' }).data.token; /* รหัสผู้ดูแลเปลี่ยนจากรหัสชั่วคราว → โทเค็นเดิมหมดอายุ */ assert.equal(g.sheet('มิเตอร์')._rows()[1][3], 'ร้อย.ก'); });
await t('blind second upload rejected', () => { assert.equal(call('admin/config', { config: cfg, etag: null }, A, 'PUT').status, 409); });
await t('public state hides codes', () => { const s = JSON.stringify(get('state')); assert.ok(!s.includes('111111') && !s.includes('888888')); });
await t('temp code dead once config has admin code', () => { assert.equal(call('login', { code: TEMP }).status, 401); const r = call('login', { code: '888888' }); assert.equal(r.data.role, 'admin'); A = r.data.token; });
await t('company login', () => { const r = call('login', { code: '111111' }); assert.equal(r.data.co, 'c1'); C1 = r.data.token; });
await t('company blocked from admin', () => { assert.equal(call('admin/state', null, C1, 'GET').status, 403); assert.equal(call('admin/config', { config: cfg, etag: ET }, C1, 'PUT').status, 403); });
await t('company saves own meter (row in sheet)', () => { const r = call('reading', { meterId: 'E01', ym: OPEN, reading: 1234 }, C1); assert.ok(r.ok, JSON.stringify(r)); const row = g.sheet('ตัวเลขรายเดือน')._rows()[1]; assert.deepEqual(row.slice(0, 3), ['E01', OPEN, 1234]); assert.equal(row[4], 'ร้อย.ก'); });
await t('re-save updates same row', () => { call('reading', { meterId: 'E01', ym: OPEN, reading: 1240 }, C1); assert.equal(g.sheet('ตัวเลขรายเดือน').getLastRow(), 2); });
await t('blocked: other company / month / negative', () => {
  assert.equal(call('reading', { meterId: 'W02', ym: OPEN, reading: 5 }, C1).status, 403);
  assert.equal(call('reading', { meterId: 'E01', ym: '2020-01', reading: 5 }, C1).status, 403);
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: -1 }, C1).status, 400);
});
await t('photo saved to Drive as public-link thumbnail, kept on re-save', () => {
  const px = 'data:image/jpeg;base64,' + Buffer.from('jpeg').toString('base64');
  const r = call('reading', { meterId: 'E01', ym: OPEN, reading: 1250, photo: px }, C1);
  assert.match(r.data.rec.photo, /^https:\/\/drive\.google\.com\/thumbnail\?id=fi\d+&sz=w1600$/);
  const f = Object.values(g.files()).find(x => x.type === 'image/jpeg'); assert.equal(f.sharing, 'ANYONE_WITH_LINK/VIEW');
  const r2 = call('reading', { meterId: 'E01', ym: OPEN, reading: 1251 }, C1); assert.equal(r2.data.rec.photo, r.data.rec.photo);
});
await t('foreign photo URL dropped', () => { const r = call('admin/readings', { changes: [{ meterId: 'W02', ym: '2026-01', rec: { reading: 1, photo: 'https://evil.example/x.jpg' } }] }, A); assert.equal(r.data.photos.W02, undefined); });
await t('closed month blocks company', () => {
  const s = call('admin/state', null, A, 'GET').data; s.config.settings.closed = true;
  const r = call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT'); assert.ok(r.ok); ET = r.data.etag;
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: 1300 }, C1).status, 403);
  s.config.settings.closed = false; ET = call('admin/config', { config: s.config, etag: ET }, A, 'PUT').data.etag;
});
await t('stale etag -> 409', () => { assert.equal(call('admin/config', { config: cfg, etag: '1' }, A, 'PUT').status, 409); });
await t('admin batch merges & keeps company entry; delete works', () => {
  call('admin/readings', { changes: [{ meterId: 'W02', ym: '2026-02', rec: { reading: 20, by: 'นำเข้า', at: '2026-03-01' } }, { meterId: 'W02', ym: '2026-03', rec: { reading: 30 } }] }, A);
  let s = call('admin/state', null, A, 'GET').data; assert.equal(s.readings.W02['2026-02'].by, 'นำเข้า'); assert.ok(s.readings.E01[OPEN]);
  call('admin/readings', { changes: [{ meterId: 'W02', ym: '2026-03', rec: null }] }, A);
  s = call('admin/state', null, A, 'GET').data; assert.equal(s.readings.W02['2026-03'], undefined);
});
await t('numbers typed directly in sheet are read back', () => {
  const sh = g.sheet('ตัวเลขรายเดือน'); const n = sh.getLastRow(); sh.getRange(n + 1, 1, 1, 6).setValues([['W02', '2026-04', '1,234', '', 'แก้ในชีต', '']]);
  assert.equal(get('state').data.readings.W02['2026-04'].reading, 1234);
});
await t('Date-typed month cell tolerated', () => { const sh = g.sheet('ตัวเลขรายเดือน'); sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([['W02', new Date('2026-05-01T00:00:00+07:00'), 50, '', '', '']]); assert.equal(get('state').data.readings.W02['2026-05'].reading, 50); });
await t('log', () => { call('log', { entries: ['ทดสอบ'] }, C1); const l = call('admin/state', null, A, 'GET').data.log; assert.equal(l[0].who, 'ร้อย.ก'); assert.ok(l.some(x => x.text.startsWith('กรอกเลข E01'))); });
// ---------- บัญชีผู้ใช้ ----------
let U1, U2, uid1, uid2;
await t('register validates input', () => {
  assert.equal(call('register', { user: 'ab', pass: '123456', name: 'ส.อ.ทดสอบ' }).status, 400);
  assert.equal(call('register', { user: 'somchai', pass: '123', name: 'ส.อ.ทดสอบ' }).status, 400);
  assert.equal(call('register', { user: 'somchai', pass: '123456', name: '' }).status, 400);
  assert.equal(call('register', { user: 'somchai', pass: '123456', name: 'ส.อ.ทดสอบ', co: 'zz' }).status, 400);
});
await t('register ok, duplicate rejected, hash not plain', () => {
  assert.ok(call('register', { user: 'Somchai', pass: 'pass1234', name: 'ส.อ.สมชาย ใจดี', pos: 'เสมียน', phone: '081', co: 'c1' }).ok);
  assert.equal(call('register', { user: 'somchai', pass: 'pass1234', name: 'x y' }).status, 409);
  assert.ok(call('register', { user: 'boss', pass: 'boss5678', name: 'ร.ต.หัวหน้า', co: 'admin' }).ok);
  const raw = Object.values(g.files()).map(f => Buffer.from(f.data, 'base64').toString()).find(s => s.includes('somchai'));
  assert.ok(raw && !raw.includes('pass1234'));
});
await t('pending cannot log in; wrong password 401', () => {
  assert.equal(call('login', { user: 'somchai', pass: 'pass1234' }).status, 403);
  assert.equal(call('login', { user: 'somchai', pass: 'wrong!!' }).status, 401);
});
await t('users list admin-only, no hashes', () => {
  assert.equal(call('admin/users', null, C1, 'GET').status, 403);
  const l = call('admin/users', null, A, 'GET').data.users; assert.equal(l.length, 2); assert.ok(!JSON.stringify(l).includes('hash') && !JSON.stringify(l).includes('salt'));
  uid1 = l.find(x => x.user === 'somchai').id; uid2 = l.find(x => x.user === 'boss').id; assert.equal(l[0].status, 'pending');
});
await t('approve needs a company', () => { assert.equal(call('admin/user', { id: uid1, op: 'approve', role: 'company', co: 'nope' }, A).status, 400); });
await t('approved company user enters only own company, named in sheet', () => {
  assert.ok(call('admin/user', { id: uid1, op: 'approve', role: 'company', co: 'c2' }, A).ok);
  const r = call('login', { user: 'SOMCHAI', pass: 'pass1234' }); assert.ok(r.ok, JSON.stringify(r)); assert.equal(r.data.co, 'c2'); U1 = r.data.token;
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: 9 }, U1).status, 403);
  assert.ok(call('reading', { meterId: 'W02', ym: OPEN, reading: 77 }, U1).ok);
  assert.equal(get('state').data.readings.W02[OPEN].by, 'ส.อ.สมชาย ใจดี (ร้อย.ข)');
  assert.equal(call('admin/state', null, U1, 'GET').status, 403);
});
await t('role change takes effect immediately (old token dead)', () => {
  call('admin/user', { id: uid1, op: 'save', role: 'company', co: 'c1' }, A);
  assert.equal(call('reading', { meterId: 'W02', ym: OPEN, reading: 78 }, U1).status, 401);
  U1 = call('login', { user: 'somchai', pass: 'pass1234' }).data.token;
  assert.ok(call('reading', { meterId: 'E01', ym: OPEN, reading: 1260 }, U1).ok);
});
await t('user admin can manage, but not itself', () => {
  call('admin/user', { id: uid2, op: 'approve', role: 'admin' }, A);
  U2 = call('login', { user: 'boss', pass: 'boss5678' }).data.token;
  assert.ok(call('admin/state', null, U2, 'GET').ok);
  assert.equal(call('admin/user', { id: uid2, op: 'disable' }, U2).status, 400);
  assert.ok(call('admin/user', { id: uid1, op: 'disable' }, U2).ok);
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: 1261 }, U1).status, 401);
  assert.equal(call('login', { user: 'somchai', pass: 'pass1234' }).status, 403);
  assert.ok(call('admin/state', null, U2, 'GET').data.log.some(x => x.who === 'ร.ต.หัวหน้า (ผู้ดูแล)'));
});
await t('reset password + own password change', () => {
  assert.ok(call('admin/user', { id: uid1, op: 'reset', pass: 'newpass1' }, A).ok);
  call('admin/user', { id: uid1, op: 'approve', role: 'company', co: 'c1' }, A);
  assert.equal(call('login', { user: 'somchai', pass: 'pass1234' }).status, 401);
  U1 = call('login', { user: 'somchai', pass: 'newpass1' }).data.token;
  assert.equal(call('me/password', { old: 'bad', pass: 'mine9999' }, U1).status, 403);
  const r = call('me/password', { old: 'newpass1', pass: 'mine9999' }, U1); assert.ok(r.ok);
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: 1262 }, U1).status, 401); // โทเค็นเก่าใช้ไม่ได้
  assert.ok(call('reading', { meterId: 'E01', ym: OPEN, reading: 1262 }, r.data.token).ok);
});
await t('delete user; admin config save keeps users', () => {
  const s = call('admin/state', null, A, 'GET').data; assert.ok(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').ok);
  assert.equal(call('admin/users', null, A, 'GET').data.users.length, 2);
  call('admin/user', { id: uid1, op: 'delete' }, A);
  assert.equal(call('admin/users', null, A, 'GET').data.users.length, 1);
  assert.equal(call('login', { user: 'somchai', pass: 'mine9999' }).status, 401);
});
// ---------- ตลาด: มิเตอร์รายห้อง + รูปตำแหน่ง ----------
let C1b;
await t('market rooms: setup', () => {
  const s = call('admin/state', null, A, 'GET').data;
  s.config.meters.push({ id: 'R01', code: 'E50', type: 'elec', companyId: 'c1', bldId: 'b9', room: 'ก-1', shop: 'ร้านลับมาก', name: 'ห้อง ก-1', lng: 100, lat: 13 },
    { id: 'R02', code: 'E51', type: 'elec', companyId: 'c1', bldId: 'b9', room: 'ก-2', shop: 'ร้านสอง', name: 'ห้อง ก-2', lng: 102, lat: 15 },
    { id: 'R03', code: 'W50', type: 'water', companyId: 'c1', bldId: 'b9', room: 'ก-1', shop: 'ร้านลับมาก', name: 'น้ำ ก-1', lng: 100, lat: 13 });
  ET = call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').data.etag;
  call('admin/readings', { changes: [
    { meterId: 'R01', ym: '2026-01', rec: { reading: 100 } }, { meterId: 'R01', ym: '2026-02', rec: { reading: 130.5 } },
    { meterId: 'R02', ym: '2026-01', rec: { reading: 10 } }, { meterId: 'R02', ym: '2026-02', rec: { reading: 20 } }, { meterId: 'R02', ym: '2026-03', rec: { reading: 25 } },
    { meterId: 'R03', ym: '2026-01', rec: { reading: 5 } }, { meterId: 'R03', ym: '2026-02', rec: { reading: 7 } }] }, A);
  C1b = call('login', { code: '111111' }).data.token;
});
await t('market rooms: public sees only building totals', () => {
  const st = get('state').data, j = JSON.stringify(st);
  assert.ok(!j.includes('ร้านลับมาก') && !j.includes('ก-1') && !j.includes('R01') && !j.includes('130.5'));
  const ae = st.config.meters.find(m => m.agg && m.type === 'elec'), aw = st.config.meters.find(m => m.agg && m.type === 'water');
  assert.equal(ae.n, 2); assert.equal(ae.bldId, 'b9'); assert.deepEqual(ae.u, { '2026-02': 40.5, '2026-03': 5 }); assert.equal(ae.lng, 101);
  assert.deepEqual(aw.u, { '2026-02': 2 });
  assert.ok(st.config.meters.some(m => m.id === 'E01'));
});
await t('market rooms: other company token = public view; own company sees rooms', () => {
  const C2 = call('login', { code: '222222' }).data.token;
  assert.ok(!JSON.stringify(call('state', null, C2, 'GET').data).includes('ร้านลับมาก'));
  const own = call('state', null, C1b, 'GET').data;
  assert.equal(own.config.meters.find(m => m.id === 'R01').shop, 'ร้านลับมาก');
  assert.equal(own.readings.R01['2026-02'].reading, 130.5);
  assert.ok(!own.config.meters.some(m => m.agg));
  assert.ok(!JSON.stringify(call('state', null, 'bad.token', 'GET').data).includes('ร้านลับมาก'));
});
await t('market rooms: company fills own room; sheet shows room', () => {
  assert.ok(call('reading', { meterId: 'R01', ym: OPEN, reading: 140 }, C1b).ok);
  assert.ok(g.sheet('มิเตอร์')._rows().some(r => r[0] === 'E50' && String(r[1]).includes('[ห้อง ก-1 ร้านลับมาก]')));
});
await t('market rooms: meter rollover counted like the unit sheet (9663 → 385 = 721)', () => {
  call('admin/readings', { changes: [{ meterId: 'R02', ym: '2026-04', rec: { reading: 9663 } }, { meterId: 'R02', ym: '2026-05', rec: { reading: 385 } }, { meterId: 'R01', ym: '2026-04', rec: { reading: 500 } }, { meterId: 'R01', ym: '2026-05', rec: { reading: 400 } }] }, A);
  const ae = get('state').data.config.meters.find(m => m.agg && m.type === 'elec');
  assert.equal(ae.u['2026-05'], 721 - 100); // R02 วนรอบ = 721, R01 เลขลดแต่ไม่ใกล้เพดาน = -100 (ให้ผู้ดูแลเห็นว่าผิด)
});
await t('location photo: owner + admin set, others blocked, public filtered', () => {
  const px = 'data:image/jpeg;base64,' + Buffer.from('loc').toString('base64');
  const C2 = call('login', { code: '222222' }).data.token;
  assert.equal(call('meter/locphoto', { meterId: 'E01', photo: px }, C2).status, 403);
  assert.equal(call('meter/locphoto', { meterId: 'E01', photo: px }).status, 401);
  assert.equal(call('meter/locphoto', { meterId: 'E01', photo: 'data:text/html;base64,AAAA' }, C1b).status, 400);
  const r = call('meter/locphoto', { meterId: 'E01', photo: px }, C1b); assert.match(r.data.url, /^https:\/\/drive\.google\.com\/thumbnail/);
  assert.ok(call('meter/locphoto', { meterId: 'R01', photo: px }, A).ok);
  const pub = get('state').data.locPhotos; assert.equal(pub.E01, r.data.url); assert.equal(pub.R01, undefined);
  assert.ok(call('state', null, C1b, 'GET').data.locPhotos.R01);
  assert.ok(call('admin/state', null, A, 'GET').data.locPhotos.R01);
  const s = call('admin/state', null, A, 'GET').data; assert.ok(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').ok);
  assert.equal(get('state').data.locPhotos.E01, r.data.url); // บันทึกแผนที่ไม่ลบรูปตำแหน่ง
  assert.ok(call('meter/locphoto', { meterId: 'E01', photo: null }, C1b).ok);
  assert.equal(get('state').data.locPhotos.E01, undefined);
});
await t('leak checks: owner/admin record, others blocked, public sees only visible keys', () => {
  const C2 = call('login', { code: '222222' }).data.token;
  assert.equal(call('check', { key: 'E01', ym: '2026-05', s: 'ok' }).status, 401);
  assert.equal(call('check', { key: 'E01', ym: '2026-05', s: 'ok' }, C2).status, 403);
  assert.equal(call('check', { key: 'bad key!', ym: '2026-05', s: 'ok' }, C1b).status, 400);
  const r = call('check', { key: 'E01', ym: '2026-05', s: 'leak', note: 'ก๊อกห้องน้ำรั่ว แจ้งซ่อมแล้ว' }, C1b);
  assert.ok(r.ok, JSON.stringify(r)); assert.equal(r.data.rec.s, 'leak'); assert.equal(r.data.rec.by, 'ร้อย.ก');
  assert.ok(call('check', { key: 'R01', ym: '2026-05', s: 'ok' }, C1b).ok);
  assert.equal(call('check', { key: 'b:b9:elec', ym: '2026-05', s: 'ok' }, A).status, 404);
  { const st = call('admin/state', null, A, 'GET').data; st.config.blds.push({ id: 'b9', name: 'ตลาดทดสอบ', companyId: 'c1', h: 3, pts: [[100, 13], [100.1, 13], [100.1, 13.1]] }); assert.ok(call('admin/config', { config: st.config, etag: st.etag }, A, 'PUT').ok); }
  assert.equal(call('check', { key: 'b:b9:elec', ym: '2026-05', s: 'ok' }, C2).status, 403);
  assert.ok(call('check', { key: 'b:b9:elec', ym: '2026-05', s: 'ok' }, C1b).ok);
  const pub = get('state').data.checks;
  assert.equal(pub.E01['2026-05'].note, 'ก๊อกห้องน้ำรั่ว แจ้งซ่อมแล้ว'); assert.equal(pub.R01, undefined); assert.equal(pub['b:b9:elec']['2026-05'].s, 'ok');
  assert.ok(call('state', null, C1b, 'GET').data.checks.R01);
  assert.ok(call('admin/state', null, A, 'GET').data.checks.R01);
  assert.ok(call('check', { key: 'E01', ym: '2026-05', s: null }, C1b).ok);
  assert.equal(get('state').data.checks.E01, undefined);
});
await t('forged/expired token rejected', () => { assert.equal(call('admin/state', null, A.slice(0, -2) + 'xx', 'GET').status, 401); assert.equal(call('admin/state', null, 'a.b', 'GET').status, 401); });
await t('view code: off by default, state open, ping gate=false', () => { assert.equal(get('ping').data.gate, false); assert.ok(get('state').ok); assert.equal(call('login', { view: 'anything' }).status, 401); });
let V1;
await t('view code: admin sets it, public state blocked, code hidden', () => {
  const s = call('admin/state', null, A, 'GET').data; s.config.settings.viewCode = 'short';
  assert.equal(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').status, 400);
  s.config.settings.viewCode = 'ดูได้2569'; assert.ok(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').ok);
  assert.equal(get('ping').data.gate, true);
  const r = get('state'); assert.equal(r.status, 401); assert.ok(!JSON.stringify(r).includes('readings'));
  assert.equal(call('state', null, 'bad.token', 'GET').status, 401);
  assert.equal(call('login', { view: 'ผิด' }).status, 401);
});
await t('view code: right code -> view token sees public data only', () => {
  const r = call('login', { view: ' ดูได้2569 ' }); assert.ok(r.ok, JSON.stringify(r)); assert.equal(r.data.role, 'view'); V1 = r.data.token;
  const st = call('state', null, V1, 'GET'); assert.ok(st.ok); const j = JSON.stringify(st.data);
  assert.ok(!j.includes('ดูได้2569') && !j.includes('ร้านลับมาก')); assert.ok(st.data.config.blds);
});
await t('view token cannot write, log or reach admin', () => {
  assert.equal(call('reading', { meterId: 'E01', ym: OPEN, reading: 99999 }, V1).status, 401);
  assert.equal(call('log', { entries: ['x'] }, V1).status, 401);
  assert.equal(call('check', { key: 'E01', ym: '2026-05', s: 'ok' }, V1).status, 401);
  assert.equal(call('meter/locphoto', { meterId: 'E01', photo: null }, V1).status, 401);
  assert.equal(call('admin/state', null, V1, 'GET').status, 401);
  assert.equal(call('me/password', { old: 'x', pass: 'yyyyyy' }, V1).status, 401);
});
await t('company/admin tokens still see state; login hands out view token', () => {
  assert.ok(call('state', null, C1b, 'GET').ok); assert.ok(call('state', null, A, 'GET').ok);
  const r = call('login', { code: '888888' }); assert.ok(r.data.view); assert.ok(call('state', null, r.data.view, 'GET').ok);
});
await t('changing view code kills old view tokens; removing it reopens', () => {
  let s = call('admin/state', null, A, 'GET').data; s.config.settings.viewCode = 'รหัสใหม่9999'; assert.ok(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').ok);
  assert.equal(call('state', null, V1, 'GET').status, 401);
  assert.ok(call('state', null, call('login', { view: 'รหัสใหม่9999' }).data.token, 'GET').ok);
  s = call('admin/state', null, A, 'GET').data; delete s.config.settings.viewCode; assert.ok(call('admin/config', { config: s.config, etag: s.etag }, A, 'PUT').ok);
  assert.ok(get('state').ok); assert.equal(get('ping').data.gate, false);
});
await t('code token dies when its code changes; short/duplicate codes rejected; bad colour reset', () => {
  let s0 = call('admin/state', null, A, 'GET').data, cfg0 = s0.config;
  const co = cfg0.companies[0], oldCode = String(co.code);
  const tk = call('login', { code: oldCode }).data.token; assert.ok(call('state', null, tk, 'GET').ok);
  let c2 = JSON.parse(JSON.stringify(cfg0)); c2.companies[0].code = '12345';
  assert.equal(call('admin/config', { config: c2, etag: s0.etag }, A, 'PUT').status, 400);
  c2.companies[0].code = String(c2.companies[1].code); assert.equal(call('admin/config', { config: c2, etag: s0.etag }, A, 'PUT').status, 400);
  c2.companies[0].code = '7654321'; c2.companies[0].color = 'red;background:url(x)';
  assert.ok(call('admin/config', { config: c2, etag: s0.etag }, A, 'PUT').ok);
  const st = call('admin/state', null, A, 'GET').data; assert.equal(st.config.companies[0].color, '#888888');
  assert.equal(call('me/password', { pass: 'xxxxxxxx', old: 'y' }, tk).status, 401); /* state เปิดสาธารณะเมื่อไม่ตั้งรหัสเข้าดู จึงตรวจด้วยเส้นทางที่ต้องเข้าสู่ระบบ */
  assert.ok(call('login', { code: '7654321' }).ok);
  const legacy = g.ctx.sign_({ r: 'admin', exp: Date.now() + 3600e3 }); assert.equal(call('admin/state', null, legacy, 'GET').status, 401);
  st.config.companies[0].code = oldCode; st.config.companies[0].color = co.color; assert.ok(call('admin/config', { config: st.config, etag: st.etag }, A, 'PUT').ok);
});
await t('history overflow is archived, not deleted', () => {
  const sh = g.sheet('ประวัติ'), before = sh.getLastRow(); g.ctx.LOG_MAX = 10;
  for (let i = 0; i < 12; i++) call('log', { entries: Array.from({ length: 50 }, (_, j) => 'x' + i + '-' + j) }, A);
  const ar = g.sheet('ประวัติเก่า'); assert.ok(ar, 'archive sheet'); assert.ok(ar.getLastRow() >= 501);
  assert.ok(sh.getLastRow() <= 10 + 500 + 50); g.ctx.LOG_MAX = 5000;
});
await t('throttle per target: 30 bad codes lock code login, user login still works', () => {
  let r; for (let i = 0; i < 31; i++) r = call('login', { code: String(900000 + i) }); assert.equal(r.status, 429);
  assert.notEqual(call('login', { user: 'nobody', pass: 'x' }).status, 429);
  for (let i = 0; i < 11; i++) r = call('login', { user: 'victim', pass: 'bad' + i }); assert.equal(r.status, 429);
  assert.notEqual(call('login', { user: 'other', pass: 'x' }).status, 429);
});
await t('sheet formula injection is neutralised', () => {
  assert.equal(g.ctx.cell_('=HYPERLINK("x")'), "'=HYPERLINK(\"x\")"); assert.equal(g.ctx.cell_('+1'), "'+1"); assert.equal(g.ctx.cell_('@a'), "'@a");
  assert.equal(g.ctx.cell_('ร้อย.ก'), 'ร้อย.ก'); assert.equal(g.ctx.cell_(12), 12);
  call('log', { entries: ['=IMPORTXML("http://x")'] }, A);
  const sh = g.sheet('ประวัติ'), rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues(); assert.ok(rows.some(r => String(r[2]).startsWith("'=IMPORTXML")));
});
await t('photo URL must match drive thumbnail exactly', () => {
  assert.ok(g.ctx.PHOTO_RE.test('https://drive.google.com/thumbnail?id=abcDEF123_-xyz&sz=w1600'));
  assert.ok(!g.ctx.PHOTO_RE.test('https://drive.google.com/thumbnail?id=abcDEF123_-xyz&sz=w1600" onerror="alert(1)'));
});
await t('proxy key gate: /exec direct = 403 once PROXY_KEY is set', () => {
  g.db.props.PROXY_KEY = 'k'.repeat(40);
  assert.equal(get('ping').status, 403); assert.equal(call('login', { code: TEMP }).status, 403);
  const viaProxy = JSON.parse(g.ctx.doPost({ postData: { contents: JSON.stringify({ r: 'ping', m: 'GET', px: 'k'.repeat(40) }) } }).content);
  assert.ok(viaProxy.ok); assert.equal(viaProxy.data.app, 'meter-map');
  const wrong = JSON.parse(g.ctx.doPost({ postData: { contents: JSON.stringify({ r: 'ping', m: 'GET', px: 'k'.repeat(39) + 'x' }) } }).content);
  assert.equal(wrong.status, 403);
  assert.ok(JSON.parse(g.ctx.doGet({ parameter: { r: 'ping', px: 'k'.repeat(40) } }).content).ok);
  delete g.db.props.PROXY_KEY; assert.ok(get('ping').ok);
});
await t('unknown route', () => { assert.equal(call('nope', null, A).status, 404); });
console.log(`\nผ่าน ${ok} ข้อ`);
