/**
 * ระบบรายงานอุบัติการณ์ (Incident Report) — Apps Script backend
 * โรงพยาบาลสัตว์เมืองเอก
 *
 * ติดตั้ง:
 *   1) เปิด Google Sheets ที่ใช้อยู่ > ส่วนขยาย > Apps Script แล้ววางไฟล์นี้ทั้งไฟล์
 *   2) เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" (อนุญาตสิทธิ์ครั้งแรก)
 *   3) การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *        เรียกใช้ในฐานะ: ฉัน · ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 *   4) นำ URL ที่ลงท้าย /exec ไปใส่ใน config.js ของหน้าเว็บ
 *   รายละเอียดเพิ่มเติมดู README.md
 */

var CFG = {
  // ไอดีไฟล์ Google Sheets (ปล่อยว่างได้ถ้าสคริปต์ผูกกับชีตอยู่แล้ว)
  SPREADSHEET_ID: '1QmfuUYwjiDnyie-iXHryYsB-ioWB8xqQ5V1U_kbWNMY',
  TZ: 'Asia/Bangkok',
  SHEET_INC: 'IR_Incidents',
  SHEET_USERS: 'IR_Users',
  SHEET_SETTINGS: 'IR_Settings',
  LEGACY_SHEET: 'จาก ภายใน',       // ชีตคำตอบจาก Google Form เดิม
  SESSION_SECONDS: 21600           // 6 ชม. ต่ออายุทุกครั้งที่ใช้งาน
};

var INC_COLS = ['id', 'createdAt', 'reporter', 'branch', 'type', 'category', 'severity', 'location', 'departments',
  'incidentDate', 'incidentTime', 'hour', 'petRef', 'detail', 'claim', 'confidential', 'status', 'owner',
  'actions', 'attachments', 'closedAt', 'closedBy', 'source', 'legacyKey', 'updatedAt', 'deleted'];
var INC_HEAD = ['เลขที่', 'วันที่แจ้ง', 'ผู้แจ้ง', 'สาขา', 'ประเภท', 'หมวดปัญหา', 'ระดับความรุนแรง', 'สถานที่เกิดเหตุ', 'แผนกที่เกี่ยวข้อง',
  'วันที่เกิดเหตุ', 'เวลาเกิดเหตุ', 'ชั่วโมง', 'HN / ชื่อสัตว์', 'รายละเอียด', 'มูลค่าเคลม (บาท)', 'ความลับ', 'สถานะ', 'ผู้รับผิดชอบ',
  'บันทึกการดำเนินการ (JSON)', 'ไฟล์แนบ (JSON)', 'วันที่ปิดเรื่อง', 'ผู้ปิดเรื่อง', 'ที่มา', 'อ้างอิงฟอร์มเดิม', 'แก้ไขล่าสุด', 'ลบ'];
var USER_COLS = ['username', 'name', 'role', 'branch', 'passHash', 'salt', 'active', 'createdAt'];

var STATUS = { NEW: 'รอรับเรื่อง', WORK: 'กำลังดำเนินการ', DONE: 'ปิดเรื่อง' };
var ROLES = { admin: 'ผู้ดูแลระบบ', manager: 'ผู้บริหาร / QA', branch: 'หัวหน้าสาขา' };
var SEVERITY = [
  { v: 1, name: 'ไม่มีผลกระทบ / เกือบเกิดเหตุ' },
  { v: 2, name: 'เล็กน้อย' },
  { v: 3, name: 'ปานกลาง' },
  { v: 4, name: 'รุนแรง' }
];
var DEFAULTS = {
  orgName: 'โรงพยาบาลสัตว์เมืองเอก',
  branches: [{ code: 'MA', name: 'เมืองเอก' }, { code: 'K3', name: 'คลองสาม' }, { code: 'K4', name: 'คลองสี่' }],
  types: ['ลูกค้าร้องเรียน', 'ความผิดพลาด (Error)', 'อุบัติเหตุ', 'พนักงานร้องเรียน', 'เหตุผิดปกติ / ไม่คาดคิด'],
  categories: ['การใช้ยา / จัดยา', 'การรอคอย / คิว', 'อาบน้ำตัดขน', 'พฤติกรรมบริการ', 'การประสานงานภายใน',
    'นัดหมาย / แจ้งผล', 'ค่าใช้จ่าย / การเงิน', 'การดูแลสัตว์ป่วย / ฝากเลี้ยง', 'การตรวจรักษา', 'รับ-ส่งสัตว์',
    'อุบัติเหตุ / บาดเจ็บ', 'อาคารสถานที่', 'เอกสาร / จัดซื้อ / คลัง', 'อัตรากำลัง / ภาระงาน', 'อื่น ๆ'],
  departments: ['เวชระเบียน / Front', 'พยาบาล', 'สัตวแพทย์', 'ห้องยา', 'อาบน้ำตัดขน', 'รับ-ส่งสัตว์', 'ธุรการ', 'จัดซื้อ / คลัง', 'การเงิน', 'บริหาร', 'อาคาร / แม่บ้าน'],
  locations: ['OPD', 'IPD / Ward', 'ห้องตรวจ', 'ห้องยา', 'ห้องอาบน้ำตัดขน', 'เคาน์เตอร์ / เวชระเบียน', 'โซนลูกค้านั่งรอ', 'ห้องผ่าตัด', 'บ้านลูกค้า / ระหว่างรับส่ง', 'ทางโทรศัพท์ / LINE', 'หน้าโรงพยาบาล'],
  reportPin: '',
  notifyMinSeverity: 1,
  lineTo: '',
  appUrl: '',
  photoFolderId: '',
  legacySheet: CFG.LEGACY_SHEET
};

/* ======================================================================
 *  HTTP
 * ==================================================================== */
function doGet(e) {
  if (e && e.parameter && e.parameter.fn === 'ping') return json_({ ok: { pong: true, at: now_() } });
  try {
    return HtmlService.createHtmlOutputFromFile('index').setTitle('รายงานอุบัติการณ์')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return json_({ ok: { service: 'Incident Report API', at: now_() } });
  }
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'คำขอไม่ถูกต้อง' }); }
  if (body.fn === 'batch') {
    return json_({ ok: (body.args[0] || []).map(function (c) { return run_(c[0], c[1]); }) });
  }
  return json_(run_(body.fn, body.args || []));
}

/** สำหรับหน้าเว็บที่เปิดผ่าน Apps Script โดยตรง (google.script.run) */
function api(fn, args) { var r = run_(fn, args || []); if (r.error) throw new Error(r.error); return r.ok; }
function apiBatch(calls) { return calls.map(function (c) { return run_(c[0], c[1]); }); }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

var PUBLIC = { ping: 1, config: 1, submitIncident: 1, login: 1, logout: 1 };
function run_(fn, args) {
  try {
    if (!Object.prototype.hasOwnProperty.call(API, fn)) throw new Error('ไม่พบคำสั่ง ' + fn);
    if (!PUBLIC[fn]) {
      var user = auth_(args[0]);
      return { ok: API[fn].apply(null, [user].concat(args.slice(1))) };
    }
    return { ok: API[fn].apply(null, args) };
  } catch (err) {
    return { error: String((err && err.message) || err).replace(/^(Exception|Error):\s*/, '') };
  }
}

/* ======================================================================
 *  API
 * ==================================================================== */
var API = {
  ping: function () { return { pong: true, at: now_() }; },

  /** รายการตัวเลือกสำหรับฟอร์มแจ้งเหตุ (ไม่ต้องเข้าสู่ระบบ) */
  config: function () { return publicConfig_(); },

  /** แจ้งเหตุใหม่ (ไม่ต้องเข้าสู่ระบบ; ถ้าตั้งรหัสแจ้งเหตุไว้ต้องส่งรหัสมาด้วย) */
  submitIncident: function (pin, d) {
    var s = settings_();
    if (s.reportPin && String(pin || '') !== String(s.reportPin)) throw new Error('รหัสแจ้งเหตุไม่ถูกต้อง');
    d = d || {};
    var reporter = clean_(d.reporter, 120), detail = clean_(d.detail, 20000);
    if (!reporter) throw new Error('กรุณากรอกชื่อผู้แจ้ง');
    if (!d.branch) throw new Error('กรุณาเลือกสาขา');
    if (!d.type) throw new Error('กรุณาเลือกประเภทเหตุการณ์');
    if (detail.length < 5) throw new Error('กรุณาเล่ารายละเอียดเหตุการณ์');
    var sev = Math.max(1, Math.min(4, parseInt(d.severity, 10) || 2));
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var created = now_();
      var rec = {
        id: nextId_(created), createdAt: created, reporter: reporter, branch: clean_(d.branch, 20),
        type: clean_(d.type, 60), category: clean_(d.category, 60) || guessCategory_(detail), severity: sev,
        location: clean_(d.location, 200), departments: (d.departments || []).map(function (x) { return clean_(x, 60); }).join(', '),
        incidentDate: /^\d{4}-\d{2}-\d{2}$/.test(d.incidentDate || '') ? d.incidentDate : created.slice(0, 10),
        incidentTime: clean_(d.incidentTime, 40), hour: hourOf_(d.incidentTime), petRef: clean_(d.petRef, 120),
        detail: detail, claim: num_(d.claim), confidential: d.confidential ? 'TRUE' : 'FALSE',
        status: STATUS.NEW, owner: '', actions: '[]', attachments: '[]', closedAt: '', closedBy: '',
        source: 'app', legacyKey: '', updatedAt: created, deleted: ''
      };
      var files = (d.photos || []).slice(0, 6).map(function (p, i) { return savePhoto_(rec.id, i, p); }).filter(Boolean);
      rec.attachments = JSON.stringify(files);
      appendInc_(rec);
    } finally { lock.releaseLock(); }
    try { notifyNew_(rec); } catch (err) { console.warn('notify', err); }
    return { id: rec.id, createdAt: rec.createdAt };
  },

  login: function (username, password) {
    username = String(username || '').trim().toLowerCase();
    var u = findUser_(username);
    if (!u || String(u.active) === 'FALSE' || hash_(password, u.salt) !== u.passHash) {
      Utilities.sleep(600);
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }
    var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    CacheService.getScriptCache().put('s_' + token, username, CFG.SESSION_SECONDS);
    return { token: token, user: publicUser_(u) };
  },

  logout: function (token) { if (token) CacheService.getScriptCache().remove('s_' + token); return true; },

  bootstrap: function (user) {
    return { user: publicUser_(user), config: publicConfig_(), roles: ROLES, status: STATUS };
  },

  /** อุบัติการณ์ทั้งหมดที่ผู้ใช้มีสิทธิ์เห็น */
  listIncidents: function (user) {
    return readInc_().filter(function (r) { return !r.deleted && canSee_(user, r); }).map(function (r) { return view_(user, r); });
  },

  updateIncident: function (user, id, patch) {
    patch = patch || {};
    return withRow_(id, user, function (r) {
      var ok = ['type', 'category', 'location', 'incidentDate', 'incidentTime', 'petRef', 'owner', 'branch'];
      ok.forEach(function (k) { if (patch[k] != null) r[k] = clean_(patch[k], 300); });
      if (patch.detail != null && user.role !== 'branch') r.detail = clean_(patch.detail, 20000);
      if (patch.incidentTime != null) r.hour = hourOf_(patch.incidentTime);
      if (patch.severity != null) r.severity = Math.max(1, Math.min(4, parseInt(patch.severity, 10) || 2));
      if (patch.claim != null) r.claim = num_(patch.claim);
      if (patch.departments != null) r.departments = [].concat(patch.departments).join(', ');
      if (patch.confidential != null) r.confidential = patch.confidential ? 'TRUE' : 'FALSE';
      return r;
    });
  },

  /** บันทึกการดำเนินการ และเปลี่ยนสถานะ */
  addAction: function (user, id, text, status) {
    text = clean_(text, 8000);
    if (!text && !status) throw new Error('กรุณากรอกการดำเนินการ');
    return withRow_(id, user, function (r) {
      var acts = parseJson_(r.actions, []);
      var at = now_();
      if (status && [STATUS.NEW, STATUS.WORK, STATUS.DONE].indexOf(status) < 0) throw new Error('สถานะไม่ถูกต้อง');
      acts.push({ at: at, by: user.name, text: text, status: status && status !== r.status ? status : '' });
      r.actions = JSON.stringify(acts);
      if (status) {
        r.status = status;
        if (status === STATUS.DONE) { r.closedAt = at; r.closedBy = user.name; } else { r.closedAt = ''; r.closedBy = ''; }
      } else if (r.status === STATUS.NEW) r.status = STATUS.WORK;
      if (!r.owner) r.owner = user.name;
      return r;
    });
  },

  deleteIncident: function (user, id) {
    needRole_(user, 'admin');
    return withRow_(id, user, function (r) { r.deleted = 'TRUE'; return r; }), true;
  },

  changePassword: function (user, oldPw, newPw) {
    if (String(newPw || '').length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
    if (hash_(oldPw, user.salt) !== user.passHash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    setPassword_(user.username, newPw);
    return true;
  },

  listUsers: function (user) { needRole_(user, 'admin'); return readUsers_().map(publicUser_); },

  saveUser: function (user, u) {
    needRole_(user, 'admin');
    var username = String(u.username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{2,30}$/.test(username)) throw new Error('ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 . _ - (2-30 ตัว)');
    if (!ROLES[u.role]) throw new Error('บทบาทไม่ถูกต้อง');
    if (u.role === 'branch' && !u.branch) throw new Error('หัวหน้าสาขาต้องระบุสาขา');
    var sh = sheet_(CFG.SHEET_USERS), rows = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === username) {
        sh.getRange(i + 1, 2, 1, 3).setValues([[clean_(u.name, 80), u.role, u.role === 'branch' ? u.branch : '']]);
        sh.getRange(i + 1, 7).setValue(u.active === false ? 'FALSE' : 'TRUE');
        if (u.password) setPassword_(username, u.password);
        return true;
      }
    }
    if (String(u.password || '').length < 6) throw new Error('ผู้ใช้ใหม่ต้องตั้งรหัสผ่านอย่างน้อย 6 ตัวอักษร');
    var salt = Utilities.getUuid();
    sh.appendRow([username, clean_(u.name, 80), u.role, u.role === 'branch' ? u.branch : '', hash_(u.password, salt), salt, 'TRUE', now_()]);
    return true;
  },

  getSettings: function (user) {
    needRole_(user, 'admin');
    var s = settings_();
    s.lineTokenSet = !!PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
    s.webAppUrl = ScriptApp.getService().getUrl() || '';
    return s;
  },

  saveSettings: function (user, s) {
    needRole_(user, 'admin');
    var keys = ['orgName', 'branches', 'types', 'categories', 'departments', 'locations', 'reportPin', 'notifyMinSeverity', 'lineTo', 'appUrl', 'legacySheet'];
    var cur = settings_();
    keys.forEach(function (k) { if (s[k] != null) cur[k] = s[k]; });
    writeSettings_(cur);
    if (s.lineToken) PropertiesService.getScriptProperties().setProperty('LINE_TOKEN', String(s.lineToken).trim());
    if (s.lineToken === '') PropertiesService.getScriptProperties().deleteProperty('LINE_TOKEN');
    return true;
  },

  testNotify: function (user) {
    needRole_(user, 'admin');
    var r = pushLine_('ทดสอบการแจ้งเตือนจากระบบรายงานอุบัติการณ์ ✅ (' + user.name + ')');
    if (r !== true) throw new Error(r);
    return true;
  },

  /** ดึงข้อมูลจาก Google Form เดิมเข้าระบบ (ข้ามแถวที่เคยดึงแล้ว) */
  importLegacy: function (user) { needRole_(user, 'admin'); return importLegacy_(); }
};

/* ======================================================================
 *  การติดตั้ง
 * ==================================================================== */
function setup() {
  var ss = ss_();
  ensureSheet_(ss, CFG.SHEET_INC, INC_HEAD);
  ensureSheet_(ss, CFG.SHEET_USERS, ['ชื่อผู้ใช้', 'ชื่อ', 'บทบาท', 'สาขา', 'passHash', 'salt', 'ใช้งาน', 'สร้างเมื่อ']);
  ensureSheet_(ss, CFG.SHEET_SETTINGS, ['key', 'value']);
  var s = settings_();
  if (!s.photoFolderId) {
    var f = DriveApp.createFolder('Incident Report - รูปภาพ');
    s.photoFolderId = f.getId();
  }
  writeSettings_(s);
  if (!findUser_('admin')) {
    var salt = Utilities.getUuid();
    sheet_(CFG.SHEET_USERS).appendRow(['admin', 'ผู้ดูแลระบบ', 'admin', '', hash_('admin1234', salt), salt, 'TRUE', now_()]);
  }
  var res = importLegacy_();
  console.log('ติดตั้งเรียบร้อย · ผู้ใช้เริ่มต้น admin / admin1234 (เปลี่ยนรหัสผ่านทันที) · ดึงข้อมูลเดิม ' + res.imported + ' เรื่อง');
}

/** ดึงคำตอบใหม่จาก Google Form อัตโนมัติ (ใช้ระหว่างช่วงเปลี่ยนผ่าน) */
function installFormSync() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncLegacy') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncLegacy').forSpreadsheet(ss_()).onFormSubmit().create();
  console.log('เปิดการซิงก์จาก Google Form แล้ว');
}
function syncLegacy() {
  var r = importLegacy_();
  if (r.imported) readInc_().slice(-r.imported).forEach(function (rec) { try { notifyNew_(rec); } catch (e) { } });
}

/* ======================================================================
 *  ข้อมูล
 * ==================================================================== */
var _ss;
function ss_() { return _ss || (_ss = CFG.SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet()); }
function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้ติดตั้งระบบ (ไม่พบชีต ' + name + ') กรุณาเรียกใช้ฟังก์ชัน setup ใน Apps Script');
  return sh;
}
function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e3f1ef');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@');
  }
  return sh;
}
function readInc_() {
  var sh = sheet_(CFG.SHEET_INC), v = sh.getDataRange().getDisplayValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var r = {}; INC_COLS.forEach(function (k, j) { r[k] = v[i][j] == null ? '' : v[i][j]; });
    r._row = i + 1;
    out.push(r);
  }
  return out;
}
function rowOf_(r) { return INC_COLS.map(function (k) { var x = r[k] == null ? '' : String(r[k]); return /^[=+\-@]/.test(x) ? "'" + x : x; }); }
function appendInc_(r) { sheet_(CFG.SHEET_INC).appendRow(rowOf_(r)); }
function withRow_(id, user, fn) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var r = readInc_().filter(function (x) { return x.id === id && !x.deleted; })[0];
    if (!r) throw new Error('ไม่พบเรื่อง ' + id);
    if (!canSee_(user, r)) throw new Error('ไม่มีสิทธิ์แก้ไขเรื่องนี้');
    r = fn(r);
    r.updatedAt = now_();
    sheet_(CFG.SHEET_INC).getRange(r._row, 1, 1, INC_COLS.length).setValues([rowOf_(r)]);
    return view_(user, r);
  } finally { lock.releaseLock(); }
}
function view_(user, r) {
  var o = {};
  INC_COLS.forEach(function (k) { o[k] = r[k]; });
  o.severity = parseInt(r.severity, 10) || 2;
  o.claim = num_(r.claim);
  o.hour = r.hour === '' ? null : parseInt(r.hour, 10);
  o.confidential = String(r.confidential) === 'TRUE';
  o.departments = String(r.departments || '').split(/\s*,\s*/).filter(String);
  o.actions = parseJson_(r.actions, []);
  o.attachments = parseJson_(r.attachments, []);
  delete o.deleted; delete o.legacyKey;
  if (o.confidential && user.role === 'branch') { o.reporter = '(ปกปิด)'; }
  return o;
}
function nextId_(created) {
  var y = (parseInt(created.slice(0, 4), 10) + 543) % 100, m = created.slice(5, 7);
  var prefix = 'IR' + (y < 10 ? '0' : '') + y + m + '-';
  var max = 0;
  readInc_().forEach(function (r) { if (r.id.indexOf(prefix) === 0) max = Math.max(max, parseInt(r.id.slice(prefix.length), 10) || 0); });
  var n = max + 1;
  return prefix + (n < 10 ? '00' : n < 100 ? '0' : '') + n;
}

function settings_() {
  var s = JSON.parse(JSON.stringify(DEFAULTS));
  var sh = ss_().getSheetByName(CFG.SHEET_SETTINGS);
  if (!sh) return s;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (!r[0]) return;
    try { s[r[0]] = JSON.parse(r[1]); } catch (e) { s[r[0]] = r[1]; }
  });
  return s;
}
function writeSettings_(s) {
  var sh = sheet_(CFG.SHEET_SETTINGS);
  var rows = Object.keys(s).filter(function (k) { return k !== 'lineTokenSet' && k !== 'webAppUrl'; })
    .map(function (k) { return [k, JSON.stringify(s[k])]; });
  sh.getRange(2, 1, Math.max(sh.getLastRow(), 2), 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}
function publicConfig_() {
  var s = settings_();
  return {
    orgName: s.orgName, branches: s.branches, types: s.types, categories: s.categories,
    departments: s.departments, locations: s.locations, severity: SEVERITY, needPin: !!s.reportPin
  };
}

/* ---------- ผู้ใช้ ---------- */
function readUsers_() {
  var sh = ss_().getSheetByName(CFG.SHEET_USERS);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    var u = {}; USER_COLS.forEach(function (k, j) { u[k] = r[j]; }); return u;
  });
}
function findUser_(username) { return readUsers_().filter(function (u) { return String(u.username).toLowerCase() === username; })[0] || null; }
function publicUser_(u) { return { username: u.username, name: u.name, role: u.role, branch: u.branch, active: String(u.active) !== 'FALSE' }; }
function setPassword_(username, pw) {
  var sh = sheet_(CFG.SHEET_USERS), rows = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).toLowerCase() === username) {
    var salt = Utilities.getUuid();
    sh.getRange(i + 1, 5, 1, 2).setValues([[hash_(pw, salt), salt]]);
    return;
  }
}
function hash_(pw, salt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(pw), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ((x + 256) % 256).toString(16).padStart(2, '0'); }).join('');
}
function auth_(token) {
  if (!token) throw new Error('SESSION_EXPIRED');
  var c = CacheService.getScriptCache(), username = c.get('s_' + token);
  if (!username) throw new Error('SESSION_EXPIRED');
  var u = findUser_(username);
  if (!u || String(u.active) === 'FALSE') throw new Error('SESSION_EXPIRED');
  c.put('s_' + token, username, CFG.SESSION_SECONDS);
  return u;
}
function needRole_(user, role) { if (user.role !== role) throw new Error('เฉพาะ' + ROLES[role] + 'เท่านั้น'); }
function canSee_(user, r) { return user.role !== 'branch' || r.branch === user.branch; }

/* ---------- รูปภาพ ---------- */
function savePhoto_(id, i, p) {
  try {
    if (!p || !p.b64) return null;
    var s = settings_();
    var folder = s.photoFolderId ? DriveApp.getFolderById(s.photoFolderId) : DriveApp.getRootFolder();
    var blob = Utilities.newBlob(Utilities.base64Decode(p.b64), p.mime || 'image/jpeg', id + '-' + (i + 1) + '.jpg');
    var f = folder.createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
    return { id: f.getId(), name: f.getName(), url: 'https://drive.google.com/file/d/' + f.getId() + '/view', thumb: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w800' };
  } catch (err) { console.warn('photo', err); return null; }
}

/* ---------- LINE ---------- */
function notifyNew_(r) {
  var s = settings_();
  if (!s.lineTo || (parseInt(r.severity, 10) || 2) < (parseInt(s.notifyMinSeverity, 10) || 1)) return;
  var br = (s.branches.filter(function (b) { return b.code === r.branch; })[0] || {}).name || r.branch;
  var sev = SEVERITY[(parseInt(r.severity, 10) || 2) - 1];
  var detail = String(r.confidential) === 'TRUE' ? '(เรื่องความลับ ดูรายละเอียดในระบบ)' : String(r.detail).slice(0, 350);
  var msg = (r.severity >= 3 ? '🚨' : '🔔') + ' แจ้งอุบัติการณ์ ' + r.id + '\n' +
    'สาขา' + br + ' · ระดับ ' + sev.v + ' ' + sev.name + '\n' +
    r.type + (r.category ? ' · ' + r.category : '') + '\n\n' + detail + '\n\nผู้แจ้ง: ' + r.reporter +
    (s.appUrl ? '\n' + s.appUrl + '#' + r.id : '');
  pushLine_(msg);
}
function pushLine_(text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN'), to = settings_().lineTo;
  if (!token || !to) return 'ยังไม่ได้ตั้งค่า Channel access token หรือ ID กลุ่มปลายทาง';
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text.slice(0, 4900) }] })
  });
  return res.getResponseCode() === 200 ? true : 'LINE ตอบกลับ ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200);
}

/* ======================================================================
 *  ดึงข้อมูลจาก Google Form เดิม
 * ==================================================================== */
function importLegacy_() {
  var s = settings_();
  var sh = ss_().getSheetByName(s.legacySheet || CFG.LEGACY_SHEET);
  if (!sh) return { imported: 0, skipped: 0, note: 'ไม่พบชีต ' + s.legacySheet };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var have = {};
    readInc_().forEach(function (r) { if (r.legacyKey) have[r.legacyKey] = 1; });
    var rows = sh.getDataRange().getValues().slice(1), imported = 0, skipped = 0;
    var recs = rows.map(function (row) { return legacyRowToIncident(row); }).filter(function (r) { return r && r.createdAt; })
      .sort(function (a, b) { return a.createdAt < b.createdAt ? -1 : 1; });
    recs.forEach(function (r) {
      if (have[r.legacyKey]) { skipped++; return; }
      r.id = nextId_(r.createdAt);
      r.updatedAt = now_();
      appendInc_(r); have[r.legacyKey] = 1; imported++;
    });
    return { imported: imported, skipped: skipped };
  } finally { lock.releaseLock(); }
}

/**
 * แปลง 1 แถวจากชีตคำตอบ Google Form เดิมเป็นข้อมูลรูปแบบใหม่
 * คอลัมน์: ประทับเวลา, ชื่อ, สาขา, ประเภท, สถานที่, แจ้งเรื่อง, แผนก, ช่วงเวลา, วันที่, มูลค่าเคลม, คอลัมน์10, การดำเนินการ, ลงชื่อ, ลิงก์
 */
function legacyRowToIncident(row) {
  var c = function (i) { return row[i] == null ? '' : row[i]; };
  var createdAt = toDateTime_(c(0));
  if (!createdAt) return null;
  var detail = String(c(5)).trim();
  var claimText = String(c(9)).trim();
  var claim = num_(claimText);
  var actionText = String(c(11)).trim();
  if (claimText.replace(/[\d,.\s]|บาท|ไม่มี|ไม่ทราบ/g, '').length > 25) actionText = actionText ? claimText + '\n' + actionText : claimText;
  var sign = String(c(12)).trim(), link = String(c(13)).trim();
  var actions = [];
  if (actionText) {
    var m = /(.*?)\s*(\d{1,2}\/\d{1,2}\/\d{4})\s*$/.exec(sign);
    var by = m ? m[1] : sign, at = m ? toDateTime_(m[2]) : createdAt;
    actions.push({ at: at || createdAt, by: by || 'ฟอร์มเดิม', text: actionText, status: '' });
  }
  var status = STATUS.NEW;
  if (actions.length) status = /รอติดต่อ|รอดำเนิน|อยู่ระหว่าง|รอสรุป|รอผล|นัดคุย/.test(actionText) ? STATUS.WORK : STATUS.DONE;
  if (status === STATUS.DONE) { actions[0].status = STATUS.DONE; }
  var type = normType_(c(3));
  var depts = normDepts_(c(6), detail + ' ' + c(4));
  var t = parseTime_(c(7));
  var attachments = /^https?:\/\//.test(link) ? [{ name: 'ลิงก์ข้อมูล', url: link }] : [];
  return {
    id: '', createdAt: createdAt, reporter: String(c(1)).trim(), branch: normBranch_(c(2)), type: type,
    category: guessCategory_(detail + ' ' + c(4)), severity: guessSeverity_(detail, claim),
    location: String(c(4)).trim(), departments: depts.join(', '),
    incidentDate: toDateTime_(c(8)).slice(0, 10) || createdAt.slice(0, 10),
    incidentTime: t.text, hour: t.hour == null ? '' : t.hour,
    petRef: extractPet_(detail) || (attachments.length ? '' : link),
    detail: detail, claim: claim, confidential: /รักษาความลับ|ไม่ต้องบอก/.test(detail) ? 'TRUE' : 'FALSE',
    status: status, owner: actions.length ? actions[0].by : '', actions: JSON.stringify(actions), attachments: JSON.stringify(attachments),
    closedAt: status === STATUS.DONE ? actions[0].at : '', closedBy: status === STATUS.DONE ? actions[0].by : '',
    source: 'form', legacyKey: createdAt + '|' + String(c(1)).trim(), updatedAt: '', deleted: ''
  };
}

function normBranch_(b) { b = String(b).trim().toUpperCase().replace(/[.\s]/g, ''); return b; }
function normType_(t) {
  t = String(t).trim();
  if (/^error$/i.test(t) || /ผิดพลาด/.test(t)) return 'ความผิดพลาด (Error)';
  if (/อุบัติ/.test(t)) return 'อุบัติเหตุ';
  if (/พนักงาน/.test(t)) return 'พนักงานร้องเรียน';
  if (/ลูกค้า/.test(t)) return 'ลูกค้าร้องเรียน';
  return t || 'เหตุผิดปกติ / ไม่คาดคิด';
}
function normDepts_(raw, context) {
  var s = String(raw || '').toLowerCase(), out = [];
  var add = function (d) { if (out.indexOf(d) < 0) out.push(d); };
  if (/front|เวช|เคาน์เตอร์|เคาท์เตอร์/.test(s)) add('เวชระเบียน / Front');
  if (/พยาบาล|ward|ผู้ช่วย/.test(s)) add('พยาบาล');
  if (/สัตวแพทย์|หมอ|แพทย์/.test(s)) add('สัตวแพทย์');
  if (/ห้องยา|เภสัช/.test(s)) add('ห้องยา');
  if (/groom|gm|อาบน้ำ|ตัดขน/.test(s)) add('อาบน้ำตัดขน');
  if (/รับส่ง|รับ-ส่ง/.test(s)) add('รับ-ส่งสัตว์');
  if (/ธุรการ/.test(s)) add('ธุรการ');
  if (/จัดซื้อ|คลัง/.test(s)) add('จัดซื้อ / คลัง');
  if (/การเงิน/.test(s)) add('การเงิน');
  if (/บริหาร/.test(s)) add('บริหาร');
  return out;
}
/** "17.36น." -> 17:36 · "18.4" -> 18:40 · "14" -> 14:00 · "18.41-19.20น." -> 18:41–19:20 */
function parseTime_(v) {
  if (v instanceof Date) { var hh = v.getHours(), mm = v.getMinutes(); return { text: pad2_(hh) + ':' + pad2_(mm), hour: hh }; }
  var s = String(v || '').trim();
  var re = /(\d{1,2})(?:\s*[.:]\s*(\d{1,2}))?/g, m, parts = [];
  while ((m = re.exec(s)) && parts.length < 2) {
    var h = parseInt(m[1], 10); if (h > 24) continue;
    var mi = m[2] == null ? 0 : (m[2].length === 1 ? parseInt(m[2], 10) * 10 : parseInt(m[2], 10));
    if (mi > 59) continue;
    parts.push(pad2_(h % 24) + ':' + pad2_(mi));
  }
  if (!parts.length) return { text: s.slice(0, 40), hour: null };
  return { text: parts.join('–'), hour: parseInt(parts[0], 10) };
}
function hourOf_(t) { var h = parseTime_(t).hour; return h == null ? '' : h; }
function guessCategory_(t) {
  t = String(t || '');
  var rules = [
    ['อาบน้ำตัดขน', /อาบน้ำ|ตัดขน|กรูม|groom|\bgm\b|ช่าง/i],
    ['การใช้ยา / จัดยา', /จัดยา|จ่ายยา|ยาผิด|ป้อนยา|\d\s*mg|rate|limit|ยาเก่า|ยาเดิม/i],
    ['อุบัติเหตุ / บาดเจ็บ', /โดนกัด|ถูกกัด|ตกจาก|หกล้ม|ลื่น|บาดเจ็บ|ข่วน/],
    ['การรอคอย / คิว', /รอตรวจ|รอนาน|ยังไม่ได้เข้าตรวจ|รอคิว|รอพบ|ไม่รอ/],
    ['ค่าใช้จ่าย / การเงิน', /ค่าใช้จ่าย|ราคา|คิดเงิน|คืนเงิน|แพง|เก็บตัง/],
    ['นัดหมาย / แจ้งผล', /นัด|ผลเลือด|แจ้งผล/],
    ['เอกสาร / จัดซื้อ / คลัง', /บิล|สินค้า|จัดซื้อ|คลัง|ส่งของ|เอกสาร/],
    ['รับ-ส่งสัตว์', /รับส่ง|รับ-ส่ง|คนรับ/],
    ['พฤติกรรมบริการ', /พูดจา|หน้าบึ้ง|หน้าบิ้ง|ไม่สุภาพ|มองหน้า|ทำแรง/],
    ['อัตรากำลัง / ภาระงาน', /เคสเยอะ|คนน้อย|หมอน้อย|พยาบาลน้อย/],
    ['อาคารสถานที่', /แอร์|ห้องน้ำ|อึหมา|พื้น|ไฟดับ|สกปรก/],
    ['การดูแลสัตว์ป่วย / ฝากเลี้ยง', /ฝากเลี้ยง|ฝากรักษา|เลอะ|ward/i]
  ];
  for (var i = 0; i < rules.length; i++) if (rules[i][1].test(t)) return rules[i][0];
  return 'อื่น ๆ';
}
function guessSeverity_(t, claim) {
  t = String(t || '');
  if (/เสียชีวิต|กู้ชีพ|น็อค|no HR|ฟ้อง/i.test(t) || claim >= 10000) return 4;
  if (/บาดเจ็บ|โดนกัด|ตกจาก|x-ray|ยาผิด|จ่ายยาผิด|ผิดขนาด|ติดเชื้อ/i.test(t) || claim >= 1000) return 3;
  if (/ยังไม่ได้ป้อน|ก่อนป้อน|ตรวจพบก่อน|เกือบ/.test(t)) return 1;
  return 2;
}
function extractPet_(t) {
  t = String(t || '');
  var code = /(HN\.?\s*[\d-]+|(?:MA|K)\s?\d{6,9}|\b\d{2}-\d{3,4}\b)/i.exec(t);
  var name = /(?:น้อง(?![ๆ\s]|ผู้|คน)|F\.\s?)\s*([^\s(]{1,20})/.exec(t);
  var out = [name ? 'น้อง' + name[1] : '', code ? code[1].replace(/\s+/g, ' ') : ''].filter(String).join(' ');
  return out;
}

/* ======================================================================
 *  helpers
 * ==================================================================== */
function now_() { return fmt_(new Date()); }
function fmt_(d) {
  var t = new Date(d.getTime() + 7 * 3600000);
  return t.getUTCFullYear() + '-' + pad2_(t.getUTCMonth() + 1) + '-' + pad2_(t.getUTCDate()) + ' ' +
    pad2_(t.getUTCHours()) + ':' + pad2_(t.getUTCMinutes()) + ':' + pad2_(t.getUTCSeconds());
}
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
/** Date หรือ "d/m/yyyy[, h:mm:ss]" (ค.ศ. หรือ พ.ศ.) -> "yyyy-MM-dd HH:mm:ss" */
function toDateTime_(v) {
  if (v instanceof Date && !isNaN(v)) return fmt_(v);
  var m = /(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(String(v || ''));
  if (!m) return '';
  var y = parseInt(m[3], 10); if (y < 100) y += 2500; if (y > 2400) y -= 543;
  return y + '-' + pad2_(+m[2]) + '-' + pad2_(+m[1]) + ' ' + pad2_(+(m[4] || 0)) + ':' + (m[5] || '00') + ':' + (m[6] || '00');
}
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var m = /(\d[\d,]*(?:\.\d+)?)/.exec(String(v || ''));
  return m ? parseFloat(m[1].replace(/,/g, '')) : 0;
}
function clean_(v, max) { return String(v == null ? '' : v).replace(/\u0000/g, '').trim().slice(0, max || 500); }
function parseJson_(s, d) { try { var v = JSON.parse(s); return v == null ? d : v; } catch (e) { return d; } }
