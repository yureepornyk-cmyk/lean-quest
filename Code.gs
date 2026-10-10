// Lean Quest — กระดานคะแนนกลางบน Google Sheets
// วางไฟล์นี้ใน Extensions → Apps Script ของชีตเปล่า แล้ว Deploy → New deployment → Web app
//   Execute as: Me · Who has access: Anyone  → คัดลอก URL ที่ลงท้ายด้วย /exec ไปใส่ในเกม (แท็บ "ฉัน" → กระดานคะแนนกลาง)
// ชีตที่สร้างให้อัตโนมัติ: scores (ทุกผลการเล่น), season (สรุปอันดับซีซันปัจจุบัน), players (รายชื่อ)

var SEASON_EPOCH = new Date(2026, 9, 12);   // วันเริ่มซีซัน 1 (12 ต.ค. 2026) — แก้ได้
var SEASON_DAYS  = 14;                      // ความยาวซีซัน (วัน)
var STREAK_PTS   = 50, STREAK_MAX_DAYS = 10;
var LV_NAMES     = ['👶 ทารก CPS', '🎒 นักเรียน CPS', '🎓 มหาลัย CPS', '🥋 เซียน CPS', '⚡ เทพ CPS', '👑 ตำนาน CPS'];
var ADMIN_PIN    = '2468';                  // รหัสเข้าหลังบ้านในเกม (แท็บ ฉัน → หลังบ้านวิทยากร) — เปลี่ยนได้

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.ping) return out_({ ok: true, v: 1 });
  var sn = p.season === 'current' || !p.season ? currentSeason_() : Number(p.season);
  if (p.admin != null) {
    if (String(p.admin) !== String(ADMIN_PIN)) return out_({ ok: false, err: 'pin' });
    if (p.list === 'scores') return out_({ ok: true, rows: recentScores_(Number(p.n) || 2000), seasons: seasonsList_(), current: currentSeason_() });
    if (p.list === 'players') return out_({ ok: true, online: onlineList_(), players: playersList_(), now: Date.now() });
  }
  var rows = leaderboard_(sn);
  return out_({ ok: true, season: sn, label: seasonLabel_(sn), rows: rows, current: currentSeason_(), updated: new Date().toISOString() });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.hb) { online_(body.hb); return out_({ ok: true }); }
    var rows = Array.isArray(body.rows) ? body.rows : [];
    if (!rows.length) return out_({ ok: false, err: 'no rows' });
    var sh = sheet_('scores', ['เวลา', 'ซีซัน', 'รหัสพนักงาน', 'ชื่อ', 'แผนก', 'เกม', 'ด่าน/ระดับ', 'คะแนน', 'กำไร', 'LBE', 'ส่งมอบ%', 'ผ่าน', 'วันที่เล่น', 'ชุดคำถาม', 'เครื่อง', 'gid', 'ระดับ CPS', 'ทักษะ CPS %']);
    var out = [];
    rows.slice(0, 200).forEach(function (r) {
      var t = r.ts ? new Date(Number(r.ts)) : new Date();
      out.push([t, seasonOf_(t), s_(r.emp), s_(r.name), s_(r.team), s_(r.game), r.stage == null ? '' : r.stage, n_(r.score), n_(r.profit), n_(r.lbe), n_(r.deliv), r.pass == null ? '' : (r.pass ? 'Y' : 'N'), n_(r.days), s_(r.title), s_(r.dev), s_(r.gid), r.lvl == null ? '' : LV_NAMES[Number(r.lvl)] || '', n_(r.skill)]);
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    updatePlayers_(rows);
    rebuildSeason_();
    return out_({ ok: true, n: out.length });
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  }
}

// ---------- leaderboard: season score per player ----------
// คะแนนซีซัน = ผลรวมคะแนนด่านโรงงานของฉันในซีซัน + คะแนนห้องเรียนที่ดีที่สุด + ผลรวม Live Quiz + โบนัสเข้าเล่น (50/วัน สูงสุด 10 วัน)
function leaderboard_(sn) {
  var sh = sheet_('scores');
  var last = sh.getLastRow(); if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 18).getValues(), P = {};
  v.forEach(function (r) {
    if (Number(r[1]) !== sn) return;
    var key = s_(r[2]) || ('n:' + s_(r[3]) + '|' + s_(r[14]));
    var p = P[key] || (P[key] = { emp: s_(r[2]), name: s_(r[3]), team: s_(r[4]), career: 0, stage: 0, cls: 0, live: 0, days: {}, n: 0, lvl: '', skill: '' });
    var game = s_(r[5]), sc = Number(r[7]) || 0, d = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    p.name = s_(r[3]) || p.name; p.team = s_(r[4]) || p.team; p.days[d] = 1; p.n++;
    if (game === 'career') { p.career += sc; p.stage = Math.max(p.stage, Number(r[6]) || 0); if (r[16]) { p.lvl = s_(r[16]); p.skill = r[17]; } }
    else if (game === 'factory') { p.cls = Math.max(p.cls, sc); }
    else if (game === 'live') { p.live += sc; }
  });
  var rows = Object.keys(P).map(function (k) {
    var p = P[k], days = Object.keys(p.days).length;
    var bonus = Math.min(STREAK_MAX_DAYS, days) * STREAK_PTS;
    return { emp: p.emp, name: p.name, team: p.team, score: Math.round(p.career + p.cls + p.live + bonus), career: Math.round(p.career), stage: p.stage, cls: p.cls, live: p.live, days: days, plays: p.n, lvl: p.lvl, skill: p.skill };
  });
  rows.sort(function (a, b) { return b.score - a.score; });
  return rows;
}

function rebuildSeason_() {
  var sn = currentSeason_(), rows = leaderboard_(sn);
  var sh = sheet_('season', ['อันดับ', 'รหัสพนักงาน', 'ชื่อ', 'แผนก', 'คะแนนซีซัน', 'คะแนนโรงงาน', 'ด่านสูงสุด', 'ห้องเรียน (ดีที่สุด)', 'Live Quiz', 'วันที่เล่น', 'จำนวนครั้ง', 'ระดับ CPS', 'ทักษะ CPS %']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 13).clearContent();
  sh.getRange(1, 15).setValue('ซีซัน ' + sn + ' · ' + seasonLabel_(sn) + ' · อัปเดต ' + new Date());
  if (!rows.length) return;
  var out = rows.map(function (r, i) { return [i + 1, r.emp, r.name, r.team, r.score, r.career, r.stage, r.cls, r.live, r.days, r.plays, r.lvl, r.skill]; });
  sh.getRange(2, 1, out.length, 13).setValues(out);
}

function updatePlayers_(rows) {
  var sh = sheet_('players', ['รหัสพนักงาน', 'ชื่อ', 'แผนก', 'เล่นล่าสุด']);
  var last = sh.getLastRow(), idx = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { idx[s_(r[0])] = i + 2; });
  rows.forEach(function (r) {
    var emp = s_(r.emp); if (!emp) return;
    var row = idx[emp];
    if (!row) { row = sh.getLastRow() + 1; idx[emp] = row; }
    sh.getRange(row, 1, 1, 4).setValues([[emp, s_(r.name), s_(r.team), new Date()]]);
  });
}

function recentScores_(n) {
  var sh = sheet_('scores'), last = sh.getLastRow(); if (last < 2) return [];
  var from = Math.max(2, last - n + 1), v = sh.getRange(from, 1, last - from + 1, 18).getValues();
  return v.map(function (r) { return { ts: new Date(r[0]).getTime(), season: r[1], emp: r[2], name: r[3], team: r[4], game: r[5], stage: r[6], score: r[7], profit: r[8], lbe: r[9], deliv: r[10], pass: r[11], days: r[12], title: r[13], lvl: r[16], skill: r[17] }; }).reverse();
}
function seasonsList_() {
  var sh = sheet_('scores'), last = sh.getLastRow(); if (last < 2) return [currentSeason_()];
  var seen = {}; sh.getRange(2, 2, last - 1, 1).getValues().forEach(function (r) { seen[Number(r[0])] = 1; }); seen[currentSeason_()] = 1;
  return Object.keys(seen).map(Number).sort(function (a, b) { return a - b; });
}

// ---------- presence: who is online now (heartbeat every ~90 s from each open game) ----------
var ONLINE_TTL_MS = 3 * 60 * 1000;
function online_(hb) {
  var lock = LockService.getScriptLock(); try { lock.waitLock(3000); } catch (e) { return; }
  try {
    var c = CacheService.getScriptCache(), m = JSON.parse(c.get('online') || '{}'), now = Date.now();
    var id = s_(hb.emp) || ('dev:' + s_(hb.dev));
    m[id] = { emp: s_(hb.emp), name: s_(hb.name), team: s_(hb.team), screen: s_(hb.screen), stage: hb.stage == null ? '' : hb.stage, lvl: hb.lvl == null ? '' : hb.lvl, ts: now };
    Object.keys(m).forEach(function (k) { if (now - m[k].ts > ONLINE_TTL_MS) delete m[k]; });
    c.put('online', JSON.stringify(m), 21600);
  } finally { lock.releaseLock(); }
}
function onlineList_() {
  var m = JSON.parse(CacheService.getScriptCache().get('online') || '{}'), now = Date.now();
  return Object.keys(m).map(function (k) { return m[k]; }).filter(function (r) { return now - r.ts <= ONLINE_TTL_MS; }).sort(function (a, b) { return b.ts - a.ts; });
}
function playersList_() {
  var sh = sheet_('scores'), last = sh.getLastRow(), P = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 18).getValues().forEach(function (r) {
    var key = s_(r[2]) || ('n:' + s_(r[3]) + '|' + s_(r[14])); if (s_(r[5]) === 'live' && !s_(r[2])) key = 'n:' + s_(r[3]);
    var p = P[key] || (P[key] = { emp: s_(r[2]), name: s_(r[3]), team: s_(r[4]), plays: 0, first: r[0], last: r[0], stage: 0, lvl: '', skill: '', total: 0 });
    p.plays++; p.total += Number(r[7]) || 0; if (r[0] > p.last) p.last = r[0]; if (r[0] < p.first) p.first = r[0];
    p.name = s_(r[3]) || p.name; p.team = s_(r[4]) || p.team;
    if (s_(r[5]) === 'career') { p.stage = Math.max(p.stage, Number(r[6]) || 0); if (r[16]) { p.lvl = s_(r[16]); p.skill = r[17]; } }
  });
  return Object.keys(P).map(function (k) { var p = P[k]; p.first = new Date(p.first).getTime(); p.last = new Date(p.last).getTime(); return p; }).sort(function (a, b) { return b.last - a.last; });
}

// ---------- helpers ----------
function currentSeason_() { return seasonOf_(new Date()); }
function seasonOf_(d) { var days = Math.floor((d - SEASON_EPOCH) / 864e5); return Math.max(1, Math.floor(days / SEASON_DAYS) + 1); }
function seasonLabel_(n) { var a = new Date(SEASON_EPOCH.getTime() + (n - 1) * SEASON_DAYS * 864e5), b = new Date(a.getTime() + (SEASON_DAYS - 1) * 864e5); var f = function (d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'd MMM yyyy'); }; return f(a) + ' – ' + f(b); }
function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (header) { sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1); } }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function s_(v) { return v == null ? '' : String(v).slice(0, 60); }
function n_(v) { return v == null || v === '' ? '' : Number(v); }
