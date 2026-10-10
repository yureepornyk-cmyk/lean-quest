// Lean Quest — กระดานคะแนนกลางบน Google Sheets
// วางไฟล์นี้ใน Extensions → Apps Script ของชีตเปล่า แล้ว Deploy → New deployment → Web app
//   Execute as: Me · Who has access: Anyone  → คัดลอก URL ที่ลงท้ายด้วย /exec ไปใส่ในเกม (แท็บ "ฉัน" → กระดานคะแนนกลาง)
// ชีตที่สร้างให้อัตโนมัติ: scores (ทุกผลการเล่น), season (สรุปอันดับซีซันปัจจุบัน), players (รายชื่อ)

var SEASON_EPOCH = new Date(2026, 9, 12);   // วันเริ่มซีซัน 1 (12 ต.ค. 2026) — แก้ได้
var SEASON_DAYS  = 14;                      // ความยาวซีซัน (วัน)
var STREAK_PTS   = 50, STREAK_MAX_DAYS = 10;

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.ping) return out_({ ok: true, v: 1 });
  var sn = p.season === 'current' || !p.season ? currentSeason_() : Number(p.season);
  var rows = leaderboard_(sn);
  return out_({ ok: true, season: sn, label: seasonLabel_(sn), rows: rows, updated: new Date().toISOString() });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var rows = Array.isArray(body.rows) ? body.rows : [];
    if (!rows.length) return out_({ ok: false, err: 'no rows' });
    var sh = sheet_('scores', ['เวลา', 'ซีซัน', 'รหัสพนักงาน', 'ชื่อ', 'แผนก', 'เกม', 'ด่าน/ระดับ', 'คะแนน', 'กำไร', 'LBE', 'ส่งมอบ%', 'ผ่าน', 'วันที่เล่น', 'ชุดคำถาม', 'เครื่อง', 'gid']);
    var out = [];
    rows.slice(0, 200).forEach(function (r) {
      var t = r.ts ? new Date(Number(r.ts)) : new Date();
      out.push([t, seasonOf_(t), s_(r.emp), s_(r.name), s_(r.team), s_(r.game), r.stage == null ? '' : r.stage, n_(r.score), n_(r.profit), n_(r.lbe), n_(r.deliv), r.pass == null ? '' : (r.pass ? 'Y' : 'N'), n_(r.days), s_(r.title), s_(r.dev), s_(r.gid)]);
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
  var v = sh.getRange(2, 1, last - 1, 16).getValues(), P = {};
  v.forEach(function (r) {
    if (Number(r[1]) !== sn) return;
    var key = s_(r[2]) || ('n:' + s_(r[3]) + '|' + s_(r[14]));
    var p = P[key] || (P[key] = { emp: s_(r[2]), name: s_(r[3]), team: s_(r[4]), career: 0, stage: 0, cls: 0, live: 0, days: {}, n: 0 });
    var game = s_(r[5]), sc = Number(r[7]) || 0, d = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    p.name = s_(r[3]) || p.name; p.team = s_(r[4]) || p.team; p.days[d] = 1; p.n++;
    if (game === 'career') { p.career += sc; p.stage = Math.max(p.stage, Number(r[6]) || 0); }
    else if (game === 'factory') { p.cls = Math.max(p.cls, sc); }
    else if (game === 'live') { p.live += sc; }
  });
  var rows = Object.keys(P).map(function (k) {
    var p = P[k], days = Object.keys(p.days).length;
    var bonus = Math.min(STREAK_MAX_DAYS, days) * STREAK_PTS;
    return { emp: p.emp, name: p.name, team: p.team, score: Math.round(p.career + p.cls + p.live + bonus), career: Math.round(p.career), stage: p.stage, cls: p.cls, live: p.live, days: days, plays: p.n };
  });
  rows.sort(function (a, b) { return b.score - a.score; });
  return rows;
}

function rebuildSeason_() {
  var sn = currentSeason_(), rows = leaderboard_(sn);
  var sh = sheet_('season', ['อันดับ', 'รหัสพนักงาน', 'ชื่อ', 'แผนก', 'คะแนนซีซัน', 'คะแนนโรงงาน', 'ด่านสูงสุด', 'ห้องเรียน (ดีที่สุด)', 'Live Quiz', 'วันที่เล่น', 'จำนวนครั้ง']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 11).clearContent();
  sh.getRange(1, 13).setValue('ซีซัน ' + sn + ' · ' + seasonLabel_(sn) + ' · อัปเดต ' + new Date());
  if (!rows.length) return;
  var out = rows.map(function (r, i) { return [i + 1, r.emp, r.name, r.team, r.score, r.career, r.stage, r.cls, r.live, r.days, r.plays]; });
  sh.getRange(2, 1, out.length, 11).setValues(out);
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
