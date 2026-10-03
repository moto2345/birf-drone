/**
 * 2026 BIRF 드론 비행 현황판 — 구글 시트 백엔드 (Google Apps Script)
 *
 * 설치
 *  1) 새 구글 시트 만들기 → 메뉴 [확장 프로그램] → [Apps Script]
 *  2) 기본 코드를 모두 지우고 이 파일 내용을 붙여넣기 → 저장
 *  3) [배포] → [새 배포] → 유형: 웹 앱
 *       - 다음 사용자 인증정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *  4) 나온 웹 앱 URL(…/exec)을 현황판 첫 화면에 붙여넣기
 *
 * 시트 탭 '현황'(드론별 현재 상태)과 '비행기록'(이착륙 로그)은 첫 요청 때 자동으로 만들어집니다.
 * 하루 마감 때는 시트 메뉴 [드론 현황판] → [전체 착륙 처리]를 누르면 됩니다.
 */

var SHEET_STATE = '현황';
var SHEET_LOG = '비행기록';
var TZ = 'Asia/Seoul';
var STATE_HEADERS = ['드론ID', '상태', '구역', 'x', 'y', '고도', '메모', '이륙시각(ms)', '갱신(ms)', '갱신시각'];
var LOG_HEADERS = ['시각', '드론ID', '동작', '구역', '고도', '메모', '비행시간(분)'];
var LOG_RETURN = 30;      // 화면에 돌려줄 최근 기록 수
var CACHE_SEC = 300;      // 캐시 유지 시간(초). 쓰기 때마다 즉시 갱신됨

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    out = (p.a === 'set') ? setState_(p) : getState_(p);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  return doGet({ parameter: p });
}

/* ---------- 조회: 캐시 우선 (바뀐 게 없으면 아주 짧은 응답) ---------- */
function getState_(p) {
  var cache = CacheService.getScriptCache();
  var got = cache.getAll(['rev', 'snap']);
  var rev = got.rev, snap = got.snap;
  if (rev && snap && p.rev && String(p.rev) === rev) {
    return { ok: true, same: true, rev: rev, now: Date.now() };
  }
  var o;
  if (rev && snap) {
    o = JSON.parse(snap);
  } else {
    o = buildSnap_(SpreadsheetApp.getActiveSpreadsheet());
    putCache_(o);
  }
  o.ok = true;
  o.now = Date.now();
  return o;
}

/* ---------- 저장: fly(이륙) / land(착륙) / move(비행 중 위치·정보 변경) ---------- */
function setState_(p) {
  var id = String(p.u || '').trim();
  if (!/^[A-Za-z0-9_]{1,24}$/.test(id)) throw new Error('bad unit');
  var action = String(p.s || '');
  if (['fly', 'land', 'move'].indexOf(action) < 0) throw new Error('bad state');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureSheet_(ss, SHEET_STATE, STATE_HEADERS);
    var lg = ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
    var now = Date.now();

    var data = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) { row = i; break; }
    }
    var prev = row > 0 ? rowToUnit_(data[row]) : null;
    var prevState = prev ? prev.s : 'land';

    var newState = action === 'move' ? prevState : action;
    var since = 0;
    if (newState === 'fly') {
      since = (prevState === 'fly' && prev && prev.since) ? prev.since : now;
    }
    var z = clip_(p.z, 12), al = clip_(p.al, 10), n = clip_(p.n, 40);
    var x = num01_(p.x), y = num01_(p.y);

    var vals = [id, newState, z, x, y, al, safeText_(n), since, now, fmt_(now)];
    if (row > 0) sh.getRange(row + 1, 1, 1, vals.length).setValues([vals]);
    else sh.appendRow(vals);

    // 비행기록
    var act = '', mins = '';
    if (action === 'fly') act = prevState === 'fly' ? '정보변경' : '이륙';
    else if (action === 'land') {
      act = prevState === 'fly' ? '착륙' : '착륙(재확인)';
      if (prevState === 'fly' && prev && prev.since) mins = Math.max(0, Math.round((now - prev.since) / 60000));
    } else if (action === 'move' && (!prev || prev.z !== z)) act = '구역변경';
    if (act) lg.appendRow([fmt_(now), id, act, z, al, safeText_(n), mins]);

    var o = buildSnap_(ss);
    putCache_(o);
    o.ok = true;
    o.now = now;
    return o;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 → 화면용 데이터 ---------- */
function buildSnap_(ss) {
  var sh = ensureSheet_(ss, SHEET_STATE, STATE_HEADERS);
  var lg = ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
  var units = {};
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    units[String(data[i][0])] = rowToUnit_(data[i]);
  }
  var log = [];
  var last = lg.getLastRow();
  if (last > 1) {
    var start = Math.max(2, last - LOG_RETURN + 1);
    var rows = lg.getRange(start, 1, last - start + 1, LOG_HEADERS.length).getValues();
    for (var j = rows.length - 1; j >= 0; j--) {
      var r = rows[j];
      log.push([String(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), r[6] === '' ? '' : Number(r[6])]);
    }
  }
  return { rev: String(Date.now()) + Math.floor(Math.random() * 1000), units: units, log: log };
}

function rowToUnit_(r) {
  return {
    s: String(r[1] || 'land'),
    z: String(r[2] || ''),
    x: r[3] === '' ? null : Number(r[3]),
    y: r[4] === '' ? null : Number(r[4]),
    al: String(r[5] || ''),
    n: String(r[6] || ''),
    since: Number(r[7]) || 0,
    t: Number(r[8]) || 0
  };
}

function putCache_(o) {
  CacheService.getScriptCache().putAll({ rev: o.rev, snap: JSON.stringify(o) }, CACHE_SEC);
}

function clearCache_() {
  CacheService.getScriptCache().removeAll(['rev', 'snap']);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clip_(v, n) { return String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, n); }
function num01_(v) {
  var f = parseFloat(v);
  if (!isFinite(f)) return '';
  return Math.round(Math.min(1, Math.max(0, f)) * 10000) / 10000;
}
function safeText_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; } // 수식 주입 방지
function fmt_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd HH:mm:ss'); }

/* ---------- 시트 메뉴 / 직접 수정 반영 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('드론 현황판')
    .addItem('전체 착륙 처리', 'landAll')
    .addItem('화면 캐시 새로고침', 'clearCache_')
    .addToUi();
}

// 시트를 손으로 고치면 현황판에도 바로 반영되도록 캐시를 비움
function onEdit() { clearCache_(); }

function landAll() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureSheet_(ss, SHEET_STATE, STATE_HEADERS);
    var lg = ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
    var data = sh.getDataRange().getValues();
    var now = Date.now();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === 'fly') {
        sh.getRange(i + 1, 2).setValue('land');
        sh.getRange(i + 1, 8, 1, 3).setValues([[0, now, fmt_(now)]]);
        lg.appendRow([fmt_(now), String(data[i][0]), '전체착륙 처리', String(data[i][2]), '', '', '']);
      }
    }
    clearCache_();
  } finally {
    lock.releaseLock();
  }
}
