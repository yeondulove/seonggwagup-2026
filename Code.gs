/***** ================================================================
 *  성과급 배분 · 자동집계 웹앱  (Google Apps Script)
 *  - 팀장   : 비밀번호 입력 → 본인 팀 등급 입력 → 확정 → 시트에 자동 기록
 *  - 인사팀 : 관리자 비밀번호 → 전체 집계 대시보드(+ 구글시트 자동 집계)
 *  설치/배포 방법은 설치가이드.md 참고.
 *  ================================================================ *****/

/* ============ 설정 (필요 시 수정) ============ */

// ★ 관리자(여규동 부대표) 비밀번호 — 운영 시 더 강력한 값으로 바꾸길 권장합니다.
var ADMIN_PASSWORD = '871114';

// 팀 데이터(재원·비밀번호 포함)는 서버에만 있으며 팀장 화면으로 내려가지 않습니다.
// (팀장은 인증 후 '본인 팀'만 받습니다.)
// 데이터 기준: 성과급_배분표_260713.xlsx (재원 원 단위 반올림, 등급 대문자 통일)
var TEAMS = [
  { id:'commerce', name:'커머스팀', fund:3907277, leader:'이현하', password:'6uN2-YiWY-Lhch',
    members:[{name:'민지원',role:'인턴',grade:'C'},{name:'백소정',role:'매니저',grade:'A'},{name:'한상현',role:'매니저',grade:'B'}] },

  { id:'agency', name:'대행팀(마케팅파트)', fund:16162436, leader:'김환운', password:'K3oE-Ypbq-CL7N',
    members:[{name:'김민석',role:'매니저',grade:'A'},{name:'김유라',role:'매니저',grade:'B'},{name:'박지원',role:'매니저',grade:'C'},
             {name:'성창훈',role:'매니저',grade:'B'},{name:'윤선영',role:'매니저',grade:'A'},{name:'이명림',role:'인턴',grade:'C'},
             {name:'정성은',role:'매니저',grade:'B'},{name:'정은채',role:'매니저',grade:'C'},{name:'박혜원',role:'파트장',grade:'B'}] },

  { id:'youtube', name:'유튜브파트', fund:4848731, leader:'박혜원', password:'SVvr-4rrd-gPND',
    members:[{name:'김진규',role:'매니저',grade:'B'},{name:'류현',role:'매니저',grade:'B'},{name:'이다연',role:'매니저',grade:'A'}] },

  { id:'academy', name:'아카데미팀', fund:7561619, leader:'서연교', password:'ANnd-5SxV-bhra',
    members:[{name:'김다빈',role:'매니저',grade:'B'},{name:'신효림',role:'인턴',grade:'C'},{name:'이수정',role:'매니저',grade:'B'},
             {name:'최은영',role:'매니저',grade:'C'},{name:'현승우',role:'파트장',grade:'B'}] },

  { id:'prdp', name:'프드프파트', fund:21961936, leader:'성나현', password:'BMJv-wsej-sBJM',
    members:[{name:'김나영',role:'매니저',grade:'B'},{name:'박수민',role:'매니저',grade:'B'},{name:'성나현',role:'파트장',grade:'B'}] },

  { id:'publish', name:'출판기획파트', fund:14941794, leader:'경정은', password:'kMFs-jof9-sAKA',
    members:[{name:'경정은',role:'파트장',grade:'A'}] },

  { id:'content', name:'콘텐츠기획팀', fund:4214249, leader:'김지우', password:'3SMC-ZFXc-jNQZ',
    members:[{name:'신민석',role:'매니저',grade:'A'},{name:'전유성',role:'매니저',grade:'B'},{name:'하민지',role:'매니저',grade:'A'}] },

  { id:'brand', name:'브랜드미디어팀', fund:8087775, leader:'전유진', password:'Sc4y-nusX-pkRB',
    members:[{name:'김은솔',role:'인턴',grade:'C'},{name:'문재성',role:'매니저',grade:'A'},{name:'성낙훈',role:'인턴',grade:'C'},
             {name:'유주상',role:'매니저',grade:'B'},{name:'최주영',role:'매니저',grade:'B'},{name:'황재현',role:'매니저',grade:'B'}] },

  { id:'hr_main', name:'인사팀(메인)', fund:3001022, leader:'강지혜', password:'GaCn-bpEV-uP9r',
    members:[{name:'박현종',role:'파트장',grade:'A'},{name:'박수연',role:'매니저',grade:'S'}] },

  { id:'hr_dev', name:'인재개발파트', fund:3001022, leader:'박현종', password:'HS4L-m9b3-ckjH',
    members:[{name:'이지현',role:'매니저',grade:'C'},{name:'오경민',role:'매니저',grade:'A'}] }
];

/* ============ 상수 ============ */
var WEIGHT = {S:4, A:3, B:2, C:1, D:0};
var FIXED_INTERN = 500000; // 인턴 고정 지급액
var RESULT_SHEET = '결과';
var STATUS_SHEET = '집계현황';

/* ============ 웹앱 진입점 ============ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('성과급 배분 · 자동집계')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ============ 최초 1회 실행: 시트 준비 ============ */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var r = ss.getSheetByName(RESULT_SHEET) || ss.insertSheet(RESULT_SHEET);
  r.clear();
  r.getRange(1, 1, 1, 8).setValues([['확정시각','팀','팀장','이름','직책','등급','가중치','예상성과급']]);
  r.setFrozenRows(1);

  var s = ss.getSheetByName(STATUS_SHEET) || ss.insertSheet(STATUS_SHEET);
  s.clear();
  s.getRange(1, 1, 1, 5).setValues([['팀','팀장','재원','확정여부','확정시각']]);
  for (var i = 0; i < TEAMS.length; i++) {
    s.getRange(i + 2, 1, 1, 5).setValues([[TEAMS[i].name, TEAMS[i].leader, TEAMS[i].fund, '미확정', '']]);
  }
  s.setFrozenRows(1);
  return '초기화 완료 — 결과/집계현황 시트가 준비되었습니다.';
}

/* ============ 성과급 계산: 인턴 50만원 고정, 나머지 재원을 비인턴이 가중치 비례(최대잉여법) ============ */
function computePay(members, fund) {
  var n = members.length;
  var pay = members.map(function(){ return 0; });
  var internTotal = 0;
  for (var i = 0; i < n; i++) { if (members[i].role === '인턴') { pay[i] = FIXED_INTERN; internTotal += FIXED_INTERN; } }
  var remaining = fund - internTotal;

  var idx = [], w = [];
  for (var j = 0; j < n; j++) { if (members[j].role !== '인턴') { idx.push(j); w.push(WEIGHT[members[j].grade]); } }
  var sumW = w.reduce(function(a,b){ return a+b; }, 0);

  if (sumW > 0 && remaining > 0) {
    var raw = w.map(function(x){ return x / sumW * remaining; });
    var fl = raw.map(Math.floor);
    var rem = remaining - fl.reduce(function(a,b){ return a+b; }, 0);
    var ord = raw.map(function(rr,k){ return { k:k, f: rr - Math.floor(rr) }; })
                 .sort(function(a,b){ return b.f - a.f; });
    for (var t = 0; t < ord.length && rem > 0; t++) { fl[ord[t].k] += 1; rem--; }
    for (var q = 0; q < idx.length; q++) { pay[idx[q]] = fl[q]; }
  }
  return { pay: pay, sumW: sumW, remaining: remaining, internTotal: internTotal };
}

/* ============ 팀장 인증 → 본인 팀 반환 ============ */
function authTeam(pw) {
  pw = (pw || '').trim();
  var t = null;
  for (var i = 0; i < TEAMS.length; i++) { if (TEAMS[i].password === pw) { t = TEAMS[i]; break; } }
  if (!t) return { ok: false };

  var members = t.members.map(function(m){ return { name:m.name, role:m.role, grade:m.grade }; });
  var saved = readConfirmed_(t.name);
  if (saved && saved.grades.length === members.length) {
    for (var j = 0; j < members.length; j++) { if (saved.grades[j]) members[j].grade = saved.grades[j]; }
  }
  return {
    ok: true,
    team: { id: t.id, name: t.name, leader: t.leader, fund: t.fund, members: members },
    confirmed: !!saved,
    confirmedAt: saved ? saved.time : ''
  };
}

/* ============ 확정 제출 → 시트 기록 ============ */
function submitResult(pw, grades) {
  pw = (pw || '').trim();
  var t = null;
  for (var i = 0; i < TEAMS.length; i++) { if (TEAMS[i].password === pw) { t = TEAMS[i]; break; } }
  if (!t) return { ok: false, msg: '인증에 실패했습니다.' };
  if (!grades || grades.length !== t.members.length) return { ok: false, msg: '등급 데이터가 올바르지 않습니다.' };
  for (var g = 0; g < grades.length; g++) { if (!(grades[g] in WEIGHT)) return { ok: false, msg: '등급 값 오류: ' + grades[g] }; }

  var members = t.members.map(function(m, i){ return { name:m.name, role:m.role, grade: grades[i] }; });
  var res = computePay(members, t.fund);
  if (res.remaining < 0) return { ok: false, msg: '인턴 고정액이 재원을 초과합니다.' };
  var pay = res.pay;
  var sumPay0 = pay.reduce(function(a,b){ return a+b; }, 0);
  if (sumPay0 === 0) return { ok: false, msg: '배분액이 0원입니다. 등급을 확인하세요.' };

  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var r = ss.getSheetByName(RESULT_SHEET);
    // 같은 팀 기존 행 삭제(재확정 시 덮어쓰기)
    var data = r.getDataRange().getValues();
    for (var d = data.length - 1; d >= 1; d--) { if (data[d][1] === t.name) r.deleteRow(d + 1); }
    // 새 행 기록
    var rows = members.map(function(m, i){ return [ts, t.name, t.leader, m.name, m.role, m.grade, WEIGHT[m.grade], pay[i]]; });
    r.getRange(r.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
    // 현황 갱신
    var s = ss.getSheetByName(STATUS_SHEET);
    var sd = s.getDataRange().getValues();
    for (var k = 1; k < sd.length; k++) { if (sd[k][0] === t.name) { s.getRange(k + 1, 4, 1, 2).setValues([['확정', ts]]); break; } }
  } finally {
    lock.releaseLock();
  }

  return {
    ok: true,
    ts: ts,
    members: members.map(function(m, i){ return { name:m.name, role:m.role, grade:m.grade, weight:WEIGHT[m.grade], pay:pay[i] }; })
  };
}

/* ============ 확정 내역 읽기 ============ */
function readConfirmed_(teamName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = ss.getSheetByName(RESULT_SHEET);
  if (!r) return null;
  var data = r.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) { if (data[i][1] === teamName) rows.push(data[i]); }
  if (!rows.length) return null;
  return { time: rows[0][0], grades: rows.map(function(x){ return x[5]; }) };
}

/* ============ 관리자 대시보드 ============ */
function getDashboard(pw) {
  if ((pw || '').trim() !== ADMIN_PASSWORD) return { ok: false };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(STATUS_SHEET);
  var sd = s ? s.getDataRange().getValues() : [];
  var rows = [], confirmedCount = 0;
  for (var i = 1; i < sd.length; i++) {
    var conf = sd[i][3] === '확정';
    if (conf) confirmedCount++;
    rows.push({ team: sd[i][0], leader: sd[i][1], fund: sd[i][2], confirmed: conf, time: sd[i][4] });
  }
  var totalFund = TEAMS.reduce(function(a, t){ return a + t.fund; }, 0);
  return { ok: true, rows: rows, confirmedCount: confirmedCount, totalTeams: TEAMS.length, totalFund: totalFund };
}
