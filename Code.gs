/**
 * Residency CVT — Google Apps Script backend (Phase 1)
 * ฐานข้อมูล: Google Sheets  ·  ไฟล์: Google Drive  ·  หน้าเว็บ: HtmlService
 *
 * ขั้นตอนติดตั้งดูใน deploy/คู่มือการ deploy.md
 */

// ---------- ตั้งค่า ----------
const CONFIG = {
  SHEET_ID: '',            // ใส่ ID ของ Google Sheet (ว่างไว้ = ใช้ Sheet ที่ผูกกับสคริปต์)
  DRIVE_FOLDER_ID: '',     // ใส่ ID โฟลเดอร์ Drive สำหรับเก็บไฟล์ (ว่างไว้ = สร้างให้อัตโนมัติ)
  SESSION_HOURS: 8,
  MAX_UPLOAD_MB: 45,       // ข้อจำกัดของ google.script.run ~50 MB ต่อครั้ง
  APP_TITLE: 'Residency CVT · หลักสูตรการพยาบาลเฉพาะทาง',
  SHARE: 'DOMAIN',         // 'DOMAIN' = เปิดดูได้เฉพาะในองค์กร · 'ANYONE' = ทุกคนที่มีลิงก์
};

const ROLE_EN = { 'ผู้ดูแลระบบ':'admin', 'Preceptor':'preceptor', 'พี่เลี้ยง':'preceptor', 'ผู้เรียน':'learner',
  admin:'admin', preceptor:'preceptor', learner:'learner' };
const ROLE_TH = { admin:'ผู้ดูแลระบบ', preceptor:'Preceptor', learner:'ผู้เรียน' };
const CHUNK = 45000;   // ขนาดสูงสุดต่อเซลล์ของ Google Sheets คือ 50,000 ตัวอักษร

const TABLES = {
  Users:        ['sap','name','role','codeHash','active','ward','cohortId','updatedAt'],
  Cohorts:      ['id','name','year','yr','period','status'],
  Students:     ['code','sap','name','nameEn','license','member','ward','preceptor','cohortId','status','updatedAt'],
  Applications: ['no','at','sap','name','nameEn','code','license','member','ward','phone','email','cohort','photoFileId','status','reviewedBy','reviewedAt'],
  Progress:     ['sap','key','value','updatedAt'],         // pre/post-test, สื่อที่เปิดแล้ว, tick ประเด็น
  Logbook:      ['id','sap','book','unit','item','date','supervisor','status','signedBy','signatureFileId','signedAt','note','updatedAt'],
  Works:        ['sap','workId','status','topic','sentAt','fileIds','updatedAt'],
  Evaluations:  ['id','sap','workId','form','total','max','pct','verdict','note','by','at'],
  Attendance:   ['sessionId','sap','status','checkedAt','by'],
  Sessions:     ['id','cohortId','topic','dateFrom','dateTo','time','place','kind','fileIds','createdBy','createdAt'],
  Announcements:['id','title','body','target','by','at'],
  Files:        ['id','driveId','name','mime','size','owner','scope','refKey','at'],
  Shared:       ['key','value','updatedBy','updatedAt'],   // ค่าที่ผู้ดูแลระบบตั้ง (ชุดการเรียนรู้, ข้อสอบ, ตั้งค่า)
  Audit:        ['at','sap','action','detail'],
};

// ---------- Web App ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// ---------- ติดตั้งครั้งแรก (รันด้วยมือ 1 ครั้ง) ----------
function setup() {
  const ss = db_();
  Object.keys(TABLES).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(TABLES[name]);
      sh.setFrozenRows(1);
    }
  });
  folder_();
  folder_('assets');
  db_().getSheetByName('Shared').getRange('B:B').setNumberFormat('@');
  db_().getSheetByName('Progress').getRange('C:C').setNumberFormat('@');
  // บัญชีผู้ดูแลระบบตั้งต้น — เปลี่ยนรหัสทันทีหลังเข้าใช้ครั้งแรก
  if (!findRow_('Users', 'sap', '10006735')) {
    insert_('Users', { sap:'10006735', name:'พว.วิชชุดา ดอกผึ้ง', role:'admin',
      codeHash: hash_('Sap10006735'), active:true, updatedAt: now_() });
  }
  return 'ติดตั้งเรียบร้อย';
}

// ---------- Auth ----------
function login(sap, code) {
  sap = String(sap || '').trim();
  const u = findRow_('Users', 'sap', sap);
  if (!u || u.codeHash !== hash_(String(code || ''))) throw new Error('รหัส SAP หรือรหัสผ่านไม่ถูกต้อง');
  if (String(u.active) === 'false') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token,
    JSON.stringify({ sap:u.sap, name:u.name, role:u.role, ward:u.ward, cohortId:u.cohortId }),
    CONFIG.SESSION_HOURS * 3600);
  audit_(u.sap, 'login', '');
  return { token, user:{ sap:u.sap, name:u.name, role:u.role, ward:u.ward, cohortId:u.cohortId } };
}

function logout(token) {
  CacheService.getScriptCache().remove('s:' + token);
  return true;
}

function session_(token, roles) {
  const raw = CacheService.getScriptCache().get('s:' + token);
  if (!raw) throw new Error('หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  const s = JSON.parse(raw);
  if (roles && roles.indexOf(s.role) < 0) throw new Error('ไม่มีสิทธิ์ใช้งานส่วนนี้');
  return s;
}

// ---------- โหลดข้อมูลเริ่มต้นตามบทบาท ----------
function bootstrap(token) {
  const s = session_(token);
  const shared = {};
  rows_('Shared').forEach(r => { try { shared[r.key] = JSON.parse(r.value); } catch (e) {} });
  const out = { user:s, shared, cohorts: rows_('Cohorts'), announcements: rows_('Announcements') };

  if (s.role === 'learner') {
    out.progress = rows_('Progress').filter(r => r.sap === s.sap);
    out.logbook  = rows_('Logbook').filter(r => r.sap === s.sap);
    out.works    = rows_('Works').filter(r => r.sap === s.sap);
    out.evals    = rows_('Evaluations').filter(r => r.sap === s.sap);
  } else {
    const students = rows_('Students');
    const mine = s.role === 'admin' ? students : students.filter(x => x.ward === s.ward);
    const saps = mine.map(x => String(x.sap));
    const pick = t => rows_(t).filter(r => saps.indexOf(String(r.sap)) >= 0);
    out.students = mine;
    out.logbook = pick('Logbook');
    out.works = pick('Works');
    out.evals = pick('Evaluations');
    if (s.role === 'admin') {
      out.users = rows_('Users').map(u => ({ sap:u.sap, name:u.name, role:u.role, active:u.active, ward:u.ward }));
      out.applications = rows_('Applications');
      out.progress = rows_('Progress');
      out.attendance = rows_('Attendance');
      out.sessions = rows_('Sessions');
    }
  }
  return out;
}

// ---------- ใช้กับหน้าเว็บ (bridge) ----------
function resume(token) { return session_(token); }

function loadState(token) {
  const s = session_(token);
  const shared = readJson_(rows_('Shared'));
  const personal = readJson_(rows_('Progress').filter(r => String(r.sap) === String(s.sap)));
  const accounts = s.role === 'admin'
    ? rows_('Users').map(u => ({ name:u.name, sap:String(u.sap), role: ROLE_TH[u.role] || u.role,
        code:'', active: String(u.active) !== 'false' }))
    : null;
  return { user:s, shared, personal, accounts };
}

// ข้อมูลส่วนตัวของผู้ใช้ที่ล็อกอิน (ผลสอบ, Logbook, งาน ฯลฯ)
function saveProgress(token, key, value) {
  const s = session_(token);
  withLock_(() => putJson_('Progress', { sap:String(s.sap), updatedAt: now_() }, key, JSON.stringify(value)));
  return true;
}

// ข้อมูลกลางที่ผู้ดูแลระบบตั้ง (ชุดการเรียนรู้, ตารางเรียน, ทะเบียน, ตั้งค่า ฯลฯ)
function saveShared(token, key, value) {
  const s = session_(token, ['admin']);
  if (key === 'accounts') { withLock_(() => syncUsers_(value || [])); return true; }
  withLock_(() => putJson_('Shared', { updatedBy:String(s.sap), updatedAt: now_() }, key, JSON.stringify(value)));
  return true;
}

// Preceptor ส่งผลประเมิน / แนบเอกสารงาน → รวมเข้ากับข้อมูลกลาง
function mergeShared(token, key, patch) {
  const s = session_(token, ['preceptor','admin']);
  if (['evRes','aWork'].indexOf(key) < 0) throw new Error('ไม่อนุญาตให้แก้ไข ' + key);
  withLock_(() => {
    const cur = readJson_(rows_('Shared'))[key] || {};
    putJson_('Shared', { updatedBy:String(s.sap), updatedAt: now_() }, key, JSON.stringify(Object.assign(cur, patch || {})));
  });
  return true;
}

// ผู้เรียนยื่นใบสมัคร → ต่อท้ายรายการใบสมัคร
function appendShared(token, key, items) {
  const s = session_(token);
  if (key !== 'apps') throw new Error('ไม่อนุญาตให้แก้ไข ' + key);
  withLock_(() => {
    const cur = readJson_(rows_('Shared'))[key] || [];
    const add = (items || []).map(x => Object.assign({}, x, { sap:String(s.sap) }));
    putJson_('Shared', { updatedBy:String(s.sap), updatedAt: now_() }, key, JSON.stringify(cur.concat(add)));
  });
  return true;
}

// รายการไฟล์ในโฟลเดอร์ assets บน Drive → { ชื่อไฟล์: fileId }
function assetMap() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('assetMap');
  if (hit) return JSON.parse(hit);
  const map = {};
  const it = folder_('assets').getFiles();
  while (it.hasNext()) { const f = it.next(); map[f.getName()] = f.getId(); }
  cache.put('assetMap', JSON.stringify(map), 600);
  return map;
}

function syncUsers_(list) {
  const keep = {};
  list.forEach(u => {
    const sap = String(u.sap || '').trim();
    if (!sap) return;
    keep[sap] = true;
    const patch = { sap, name:u.name, role: ROLE_EN[u.role] || 'learner',
      active: u.active !== false, updatedAt: now_() };
    if (u.code) patch.codeHash = hash_(String(u.code));
    upsert_('Users', r => String(r.sap) === sap, patch, false);
  });
  rows_('Users').forEach(r => {
    if (!keep[String(r.sap)] && String(r.active) !== 'false')
      upsert_('Users', x => String(x.sap) === String(r.sap), { active:false, updatedAt: now_() }, true);
  });
}

function putJson_(table, base, key, json) {
  const sh = db_().getSheetByName(table);
  const head = TABLES[table];
  const kc = head.indexOf('key'), sc = head.indexOf('sap');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const vals = sh.getRange(2, 1, n, head.length).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const k = String(vals[i][kc]);
      const sameKey = k === key || k.indexOf(key + '#') === 0;
      const sameUser = sc < 0 || String(vals[i][sc]) === String(base.sap || '');
      if (sameKey && sameUser) sh.deleteRow(i + 2);
    }
  }
  const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  if (!parts.length) parts.push('null');
  const rows = parts.map((p, i) => head.map(h =>
    h === 'key' ? (i ? key + '#' + i : key) : h === 'value' ? "'" + p : (base[h] !== undefined ? base[h] : '')));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
}

function readJson_(rows) {
  const groups = {};
  rows.forEach(r => {
    const m = String(r.key).match(/^(.*?)(?:#(\d+))?$/);
    (groups[m[1]] = groups[m[1]] || [])[Number(m[2] || 0)] = String(r.value);
  });
  const out = {};
  Object.keys(groups).forEach(k => { try { out[k] = JSON.parse(groups[k].join('')); } catch (e) {} });
  return out;
}

// ---------- ผู้เรียน ----------

function addLogEntry(token, e) {
  const s = session_(token, ['learner']);
  const row = { id: Utilities.getUuid(), sap:s.sap, book:e.book || 'main', unit:e.unit, item:e.item,
    date:e.date, supervisor:e.supervisor, status:'wait', note:e.note || '', updatedAt: now_() };
  withLock_(() => insert_('Logbook', row));
  return row;
}

function editLogEntry(token, id, patch) {
  const s = session_(token, ['learner']);
  return withLock_(() => upsert_('Logbook', r => r.id === id && r.sap === s.sap,
    { date:patch.date, supervisor:patch.supervisor, status:'wait', updatedAt: now_() }, true));
}

function saveWork(token, workId, data) {
  const s = session_(token, ['learner','preceptor','admin']);
  const sap = s.role === 'learner' ? s.sap : data.sap;
  if (s.role === 'learner' && data.status === 'ผ่านแล้ว') throw new Error('ผู้เรียนไม่สามารถตั้งสถานะผ่านแล้ว');
  if (s.role === 'preceptor') { delete data.status; }   // Preceptor แนบเอกสารได้อย่างเดียว
  return withLock_(() => upsert_('Works', r => r.sap === sap && r.workId === workId,
    Object.assign({ sap, workId, updatedAt: now_() }, data), true));
}

function submitApplication(token, form) {
  const s = session_(token);
  const list = rows_('Applications');
  const no = 'AP-' + String(form.cohort || '').replace(/\D/g, '') + '-' + String(list.length + 1).padStart(3, '0');
  withLock_(() => insert_('Applications', Object.assign({}, form,
    { no, at: now_(), sap:s.sap, status:'pending' })));
  return no;
}

// ---------- Preceptor ----------
function signLogEntry(token, id, signatureDataUrl) {
  const s = session_(token, ['preceptor','admin']);
  const entry = findRow_('Logbook', 'id', id);
  if (!entry) throw new Error('ไม่พบรายการ');
  if (s.role === 'preceptor') {
    const stu = findRow_('Students', 'sap', entry.sap);
    if (!stu || stu.ward !== s.ward) throw new Error('ผู้เรียนนี้ไม่อยู่ในความดูแล');
  }
  let sigId = '';
  if (signatureDataUrl) {
    const blob = Utilities.newBlob(Utilities.base64Decode(signatureDataUrl.split(',')[1]), 'image/png', 'sig-' + id + '.png');
    sigId = folder_('signatures').createFile(blob).getId();
  }
  withLock_(() => upsert_('Logbook', r => r.id === id,
    { status:'ok', signedBy:s.name, signatureFileId:sigId, signedAt: now_(), updatedAt: now_() }, true));
  audit_(s.sap, 'sign', id);
  return true;
}

function requestFix(token, id, note) {
  const s = session_(token, ['preceptor','admin']);
  return withLock_(() => upsert_('Logbook', r => r.id === id,
    { status:'fix', note: note || '', signedBy:s.name, updatedAt: now_() }, true));
}

function saveEvaluation(token, ev) {
  const s = session_(token, ['preceptor','admin']);
  const row = Object.assign({ id: Utilities.getUuid(), by:s.name, at: now_() }, ev);
  withLock_(() => {
    insert_('Evaluations', row);
    upsert_('Works', r => r.sap === ev.sap && r.workId === ev.workId,
      { sap:ev.sap, workId:ev.workId, status: ev.pct >= 60 ? 'ผ่านแล้ว' : 'ส่งแล้ว', updatedAt: now_() }, false);
  });
  return row;
}

// ---------- ผู้ดูแลระบบ ----------

function saveUser(token, u) {
  session_(token, ['admin']);
  const patch = { sap:String(u.sap), name:u.name, role:u.role, active:u.active !== false,
    ward:u.ward || '', cohortId:u.cohortId || '', updatedAt: now_() };
  if (u.code) patch.codeHash = hash_(u.code);
  return withLock_(() => upsert_('Users', r => String(r.sap) === String(u.sap), patch, false));
}

function saveStudent(token, st) {
  session_(token, ['admin']);
  return withLock_(() => upsert_('Students', r => String(r.code) === String(st.code),
    Object.assign({}, st, { updatedAt: now_() }), false));
}

function reviewApplication(token, no, decision, cohortId) {
  const s = session_(token, ['admin']);
  const app = findRow_('Applications', 'no', no);
  if (!app) throw new Error('ไม่พบใบสมัคร');
  withLock_(() => {
    upsert_('Applications', r => r.no === no,
      { status: decision, reviewedBy:s.name, reviewedAt: now_() }, true);
    if (decision === 'accepted') {
      upsert_('Students', r => String(r.code) === String(app.code), {
        code:app.code, sap:app.sap, name:app.name, nameEn:app.nameEn, license:app.license,
        member:app.member, ward:app.ward, cohortId, status:'active', updatedAt: now_() }, false);
    }
  });
  return true;
}

function saveSession(token, sess) {
  const s = session_(token, ['admin']);
  const row = Object.assign({ id: sess.id || Utilities.getUuid(), createdBy:s.name, createdAt: now_() }, sess);
  withLock_(() => upsert_('Sessions', r => r.id === row.id, row, false));
  return row;
}

function markAttendance(token, sessionId, sap, status) {
  const s = session_(token, ['admin']);
  return withLock_(() => upsert_('Attendance', r => r.sessionId === sessionId && String(r.sap) === String(sap),
    { sessionId, sap, status, checkedAt: now_(), by:s.name }, false));
}

// ผู้เรียนสแกน QR → เปิด URL ?checkin=<sessionId>
function selfCheckIn(token, sessionId) {
  const s = session_(token, ['learner']);
  return withLock_(() => upsert_('Attendance', r => r.sessionId === sessionId && String(r.sap) === String(s.sap),
    { sessionId, sap:s.sap, status:'มา', checkedAt: now_(), by:'QR' }, false));
}

function postAnnouncement(token, a) {
  const s = session_(token, ['admin']);
  const row = { id: Utilities.getUuid(), title:a.title, body:a.body, target:a.target, by:s.name, at: now_() };
  withLock_(() => insert_('Announcements', row));
  return row;
}

// ---------- ไฟล์ ----------
// ไฟล์ ≤ 45 MB ส่งผ่านสคริปต์ได้ · ไฟล์ใหญ่ (วิดีโอ 900 MB) ให้อัปโหลดเข้า Drive/YouTube แล้วบันทึกเป็นลิงก์
function uploadFile(token, meta, base64) {
  const s = session_(token);
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > CONFIG.MAX_UPLOAD_MB * 1024 * 1024)
    throw new Error('ไฟล์เกิน ' + CONFIG.MAX_UPLOAD_MB + ' MB — กรุณาอัปโหลดเข้า Google Drive แล้ววางลิงก์แทน');
  const file = folder_(meta.scope || 'misc')
    .createFile(Utilities.newBlob(bytes, meta.mime || 'application/octet-stream', meta.name));
  try {
    file.setSharing(CONFIG.SHARE === 'ANYONE' ? DriveApp.Access.ANYONE_WITH_LINK : DriveApp.Access.DOMAIN_WITH_LINK,
      DriveApp.Permission.VIEW);
  } catch (e) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // บัญชี Gmail ทั่วไปไม่มีโดเมน
  }
  const row = { id: Utilities.getUuid(), driveId: file.getId(), name: meta.name, mime: meta.mime,
    size: bytes.length, owner: s.sap, scope: meta.scope || 'misc', refKey: meta.refKey || '', at: now_() };
  withLock_(() => insert_('Files', row));
  const isImg = /^image\//.test(meta.mime || '');
  return { id: row.id, driveId: row.driveId,
    url: isImg ? 'https://lh3.googleusercontent.com/d/' + row.driveId
               : 'https://drive.google.com/file/d/' + row.driveId + '/preview' };
}

function listFiles(token, refKey) {
  session_(token);
  return rows_('Files').filter(f => f.refKey === refKey)
    .map(f => ({ id:f.id, name:f.name, mime:f.mime, url:'https://drive.google.com/file/d/' + f.driveId + '/preview' }));
}

// ---------- ส่งออก ----------
function exportSheetUrl(token, table) {
  session_(token, ['admin']);
  const sh = db_().getSheetByName(table);
  if (!sh) throw new Error('ไม่พบตาราง ' + table);
  const id = db_().getId(), gid = sh.getSheetId();
  return {
    xlsx: 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx&gid=' + gid,
    pdf:  'https://docs.google.com/spreadsheets/d/' + id + '/export?format=pdf&gid=' + gid,
  };
}

// ---------- helpers ----------
function db_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function folder_(sub) {
  const props = PropertiesService.getScriptProperties();
  let rootId = CONFIG.DRIVE_FOLDER_ID || props.getProperty('ROOT_FOLDER');
  let root;
  if (rootId) root = DriveApp.getFolderById(rootId);
  else { root = DriveApp.createFolder('Residency CVT — Files'); props.setProperty('ROOT_FOLDER', root.getId()); }
  if (!sub) return root;
  const it = root.getFoldersByName(sub);
  return it.hasNext() ? it.next() : root.createFolder(sub);
}

function rows_(table) {
  const sh = db_().getSheetByName(table);
  if (!sh || sh.getLastRow() < 2) return [];
  const [head, ...data] = sh.getDataRange().getValues();
  return data.map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])));
}

function findRow_(table, col, val) {
  return rows_(table).find(r => String(r[col]) === String(val)) || null;
}

function insert_(table, obj) {
  const sh = db_().getSheetByName(table);
  sh.appendRow(TABLES[table].map(h => obj[h] !== undefined ? obj[h] : ''));
  return obj;
}

function upsert_(table, match, patch, mustExist) {
  const sh = db_().getSheetByName(table);
  const head = TABLES[table];
  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  for (let i = 0; i < data.length; i++) {
    const row = Object.fromEntries(head.map((h, j) => [h, data[i][j]]));
    if (match(row)) {
      const next = Object.assign(row, patch);
      sh.getRange(i + 2, 1, 1, head.length).setValues([head.map(h => next[h] !== undefined ? next[h] : '')]);
      return next;
    }
  }
  if (mustExist) throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');
  return insert_(table, patch);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function hash_(s) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || (() => {
    const v = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('SALT', v);
    return v;
  })();
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function now_() { return new Date().toISOString(); }

function audit_(sap, action, detail) {
  try { insert_('Audit', { at: now_(), sap, action, detail }); } catch (e) {}
}
