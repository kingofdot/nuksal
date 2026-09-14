// ============================================================
// 넉살 FC 통장관리 — Google Apps Script 웹앱
// ============================================================
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = e.parameter || {};
  let body = {};
  if (params.payload) {
    try { body = JSON.parse(params.payload); } catch(_) {}
  } else if (e.postData) {
    try { body = JSON.parse(e.postData.contents || '{}'); } catch(_) {}
  }
  const action = params.action || body.action;
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    let result;
    switch (action) {
      case 'load':         result = loadAll();                   break;
      case 'saveBank':     result = saveBank(body.rows);         break;
      case 'saveMembers':  result = saveMembers(body.rows);      break;
      case 'saveBackNumber': result = saveBackNumber(body.rows); break;
      case 'saveDonations':  result = saveDonations(body.rows);  break;
      case 'checkAuth':    result = checkAuth(body);             break;
      case 'requestAuth':  result = requestAuth(body);           break;
      case 'log':          result = writeLog(body);              break;
      case 'ping':         result = { msg: 'pong' };             break;
      default:             result = { ok: false, msg: 'unknown action: ' + action };
    }
    out.setContent(JSON.stringify({ ok: true, ...result }));
  } catch (err) {
    out.setContent(JSON.stringify({ ok: false, msg: err.message }));
  }
  return out;
}

// ── 기기 권한 확인
function checkAuth(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('auth');
  if (!ws) {
    // auth 시트 없으면 생성
    ws = ss.insertSheet('auth');
    ws.getRange(1,1,1,6).setValues([['기기ID','별명','권한','기기종류','브라우저','요청일']]);
    ws.getRange(1,1,1,6).setFontWeight('bold');
  }
  const last = ws.getLastRow();
  if (last < 2) return { authorized: false };
  const rows = ws.getRange(2,1,last-1,3).getValues();
  for (const row of rows) {
    if (row[0] === body.deviceId && String(row[2]).trim() === 'O') {
      return { authorized: true, nickname: row[1] };
    }
  }
  return { authorized: false };
}

// ── 권한 요청
function requestAuth(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('auth');
  if (!ws) {
    ws = ss.insertSheet('auth');
    ws.getRange(1,1,1,6).setValues([['기기ID','별명','권한','기기종류','브라우저','요청일']]);
    ws.getRange(1,1,1,6).setFontWeight('bold');
  }
  // 이미 있는 기기인지 확인
  const last = ws.getLastRow();
  if (last >= 2) {
    const ids = ws.getRange(2,1,last-1,1).getValues().flat();
    if (ids.includes(body.deviceId)) {
      // 이미 요청됨 — 업데이트만
      const idx = ids.indexOf(body.deviceId) + 2;
      ws.getRange(idx,2).setValue(body.nickname||'');
      return { msg: '이미 요청된 기기입니다. 승인을 기다려주세요.' };
    }
  }
  ws.appendRow([
    body.deviceId,
    body.nickname || '',
    '',
    body.deviceType || '',
    body.browser || '',
    body.requestTime || new Date().toLocaleString('ko-KR')
  ]);
  return { msg: '권한 요청이 접수됐어요!' };
}

// ── 행동 로그
function writeLog(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName('log');
  if (!ws) {
    ws = ss.insertSheet('log');
    ws.getRange(1,1,1,6).setValues([['시간','기기ID','기기종류','브라우저','행동','상세']]);
    ws.getRange(1,1,1,6).setFontWeight('bold');
  }
  ws.appendRow([
    body.time || new Date().toLocaleString('ko-KR'),
    body.deviceId || '',
    body.deviceType || '',
    body.browser || '',
    body.action || '',
    body.detail || ''
  ]);
  return { msg: 'logged' };
}

function loadAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    bankbook:         readSheet(ss, 'bankbook'),
    memberManagement: readSheet(ss, 'memberManagement'),
    donation:         readSheet(ss, 'DonationOfGoods'),
  };
}

function saveBank(rows)       { writeSheet('bankbook', rows);        return { msg: '통장내역 저장 완료', count:(rows||[]).length }; }
// 등번호는 회원 시트 E열에서만 관리한다. 예전 화면의 등번호 목록 저장은 받지 않는다(BackNumber 시트는 건드리지 않음).
function saveBackNumber(rows) { throw new Error('등번호는 이제 회원 시트(E열)에서만 관리해요. 새로고침(Ctrl+F5) 후 저장해주세요'); }
function saveDonations(rows)  { writeSheet('DonationOfGoods', rows);  return { msg: '물품찬조 저장 완료', count:(rows||[]).length }; }

// ── 회원 시트
// 열: 순번|구분|직책|이름|등번호|Year|나이|연락처|상세상태|m1~m12|회원ID|탈퇴일
//     A    B    C    D    E     F    G    H     I       J~U     V      W      (V·W 는 2026-09 추가)
// 등번호는 이 시트 E열 하나가 원본이다. BackNumber 시트는 읽지도 쓰지도 않는다(코드가 지우지는 않는다).
const MEMBER_COLS = 23;
const MEMBER_EXTRA_HEADERS = { 22: '회원ID', 23: '탈퇴일' };

function saveMembers(rows) {
  const width = (rows && rows.length) ? rows[0].length : 0;
  // 회원ID 열이 생긴 뒤에 예전 화면(21열)이 저장하면 거절한다.
  // 저장은 행을 이름순으로 다시 쓰는데 예전 화면은 V·W열을 모른다 — 받아 주면 회원ID 가 다른 사람 행에 붙는다.
  if (width && width < MEMBER_COLS && hasMemberIdColumn()) {
    throw new Error('예전 화면에서 보낸 저장이라 막았어요. 새로고침(Ctrl+F5) 후 다시 저장해주세요');
  }
  if (width === MEMBER_COLS) ensureHeaders('memberManagement', MEMBER_EXTRA_HEADERS);
  writeSheet('memberManagement', rows);
  return { msg: '회원 정보 저장 완료', count:(rows||[]).length };
}

function hasMemberIdColumn() {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('memberManagement');
  return !!ws && ws.getLastColumn() >= 22 && String(ws.getRange(1, 22).getValue()).trim() === '회원ID';
}

// 제목 칸이 비어 있을 때만 채운다. 이미 뭔가 적혀 있으면 건드리지 않는다.
function ensureHeaders(name, headers) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!ws) return;
  Object.keys(headers).forEach(col => {
    const cell = ws.getRange(1, Number(col));
    if (String(cell.getValue()).trim() === '') cell.setValue(headers[col]);
  });
}

// ── 저장 직전 자동 백업
// 시트를 통째로 복사해 숨겨 둔다. 이름: bak_<시트이름>_yyyyMMdd_HHmmss_SSS
// 시트마다 최근 BACKUP_KEEP 개만 남기고 오래된 것부터 지운다.
// 되돌리기: 스프레드시트 아래 시트 탭 목록 → 숨긴 시트 보기 → 해당 bak_ 시트 내용을 복사해 붙여넣기
const BACKUP_KEEP = 20;

function backupSheet(ss, ws, name) {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss_SSS');
  const copy = ws.copyTo(ss);
  copy.setName('bak_' + name + '_' + stamp);
  copy.hideSheet();
  const prefix = 'bak_' + name + '_';
  const baks = ss.getSheets()
    .filter(s => s.getName().indexOf(prefix) === 0)
    .sort((a, b) => (a.getName() < b.getName() ? 1 : -1));      // 새것부터
  baks.slice(BACKUP_KEEP).forEach(s => ss.deleteSheet(s));
}

function readSheet(ss, name) {
  const ws = ss.getSheetByName(name);
  if (!ws) return [];
  const last = ws.getLastRow();
  if (last < 2) return [];
  return ws.getRange(2, 1, last-1, ws.getLastColumn()).getValues().map(row => row.map(cell => {
    if (cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Seoul', 'yyyy-MM-dd');
    if (typeof cell === 'number') return cell;
    if (typeof cell === 'boolean') return cell;
    return String(cell);
  }));
}

function writeSheet(name, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(name);
  if (!ws) throw new Error('시트 없음: ' + name);
  backupSheet(ss, ws, name);   // 지우기 전에 시트를 통째로 복사해 둔다
  // 지우는 범위는 예전 그대로(마지막 열까지). 저장마다 행 순서가 바뀌므로,
  // 일부 열만 지우고 쓰면 남은 열 값이 다른 사람 행에 붙는다.
  const last = ws.getLastRow();
  if (last >= 2) ws.getRange(2, 1, last-1, ws.getLastColumn()).clearContent();
  if (rows && rows.length > 0) ws.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
