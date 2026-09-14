// 넉살 Apps Script(Code.gs) — 가짜 스프레드시트로 저장·백업·옛 형식 거절 시험 (실제 시트를 부르지 않는다)
const fs = require('fs'), vm = require('vm');
const code = fs.readFileSync('D:/nuksal/gas/Code.gs', 'utf8');

class Sheet {
  constructor(ss, name, data) { this.ss = ss; this.name = name; this.data = data.map(r => r.slice()); this.hidden = false; }
  getName() { return this.name; }
  setName(n) {
    if (this.ss.sheets.some(s => s !== this && s.name === n)) throw new Error('같은 이름의 시트가 있음: ' + n);
    this.name = n; return this;
  }
  getLastRow() { for (let r = this.data.length - 1; r >= 0; r--) if (this.data[r].some(v => v !== '' && v != null)) return r + 1; return 0; }
  getLastColumn() { let c = 0; this.data.forEach(r => r.forEach((v, i) => { if (v !== '' && v != null) c = Math.max(c, i + 1); })); return c; }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  copyTo(ss) { const s = new Sheet(ss, this.name + ' 사본', this.data); ss.sheets.push(s); return s; }
  hideSheet() { this.hidden = true; return this; }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  _cell(i, j) { const row = this.sh.data[this.r - 1 + i] || []; const v = row[this.c - 1 + j]; return v == null ? '' : v; }
  getValues() { return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this._cell(i, j))); }
  getValue() { return this._cell(0, 0); }
  setValues(v) {
    if (v.length !== this.nr || v.some(row => row.length !== this.nc)) throw new Error('범위 크기와 값 크기가 다름');
    v.forEach((row, i) => row.forEach((x, j) => {
      const ri = this.r - 1 + i, ci = this.c - 1 + j;
      while (this.sh.data.length <= ri) this.sh.data.push([]);
      while (this.sh.data[ri].length <= ci) this.sh.data[ri].push('');
      this.sh.data[ri][ci] = x;
    }));
    return this;
  }
  setValue(x) { return this.setValues([[x]]); }
  clearContent() { this.setValues(Array.from({ length: this.nr }, () => Array(this.nc).fill(''))); return this; }
  setFontWeight() { return this; }
}
function makeSS() {
  const ss = { sheets: [] };
  ss.getSheetByName = n => ss.sheets.find(s => s.name === n) || null;
  ss.getSheets = () => ss.sheets.slice();
  ss.insertSheet = n => { const s = new Sheet(ss, n, []); ss.sheets.push(s); return s; };
  ss.deleteSheet = s => { ss.sheets = ss.sheets.filter(x => x !== s); };
  ss.add = (n, data) => { const s = new Sheet(ss, n, data); ss.sheets.push(s); return s; };
  return ss;
}

let SS, tick = 0;
const sandbox = {
  console, JSON, Math, Date, Object, Array, String, Number,
  SpreadsheetApp: { getActiveSpreadsheet: () => SS },
  Utilities: { formatDate: () => '20260914_120000_' + String(tick++).padStart(3, '0') },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: () => ({ setMimeType() {}, setContent(c) { this.c = c; } }) },
};
vm.createContext(sandbox);
vm.runInContext(code, sandbox, { filename: 'Code.gs' });

const call = body => JSON.parse(sandbox.handleRequest({ parameter: {}, postData: { contents: JSON.stringify(body) } }).c);
const H21 = ['순번', '구분', '직책', '이름', '등번호', 'Year', '나이', '연락처', '상세상태', 'm1', 'm2', 'm3', 'm4', 'm5', 'm6', 'm7', 'm8', 'm9', 'm10', 'm11', 'm12'];
const r21 = (i, name, j) => [i, 6, '회원', name, j, 1990, 37, '010-0000-000' + i, '정상', ...Array(12).fill('')];
const r23 = (i, name, j, id) => [...r21(i, name, j), id, ''];
let n = 0;
const ok = (c, msg) => { if (!c) { console.error('  ✗ ' + msg); process.exitCode = 1; throw new Error('실패: ' + msg); } n++; console.log('  ✓ ' + msg); };
const baks = name => SS.sheets.filter(s => s.name.startsWith('bak_' + name + '_'));

console.log('① 합치기 저장(23열) — 백업 후 쓰고, 새 열 제목을 채운다');
SS = makeSS();
const ws = SS.add('memberManagement', [H21, r21(1, '가상A', 7), r21(2, '가상B', 10), r21(3, '가상C', 0)]);
let res = call({ action: 'saveMembers', rows: [r23(1, '가상A', 7, 'M0001'), r23(2, '가상B', 10, 'M0002'), r23(3, '가상C', 5, 'M0003')] });
ok(res.ok === true && res.count === 3, '저장 성공 3건');
ok(baks('memberManagement').length === 1 && baks('memberManagement')[0].hidden, '저장 직전 숨긴 백업 시트 1개');
ok(baks('memberManagement')[0].data[3][4] === 0, '백업에는 저장 전 값(가상C 등번호 0)이 남아 있음');
ok(ws.data[0][21] === '회원ID' && ws.data[0][22] === '탈퇴일', 'V1·W1 제목 = 회원ID·탈퇴일');
ok(ws.data[3][21] === 'M0003' && ws.data[3][4] === 5, '가상C 행: 회원ID M0003, 등번호 5');

console.log('② 회원ID 열이 생긴 뒤 옛 화면(21열) 저장은 거절');
res = call({ action: 'saveMembers', rows: [r21(1, '가상A', 99), r21(2, '가상B', 99)] });
ok(res.ok === false && /예전 화면/.test(res.msg), '거절 + 안내 문구');
ok(ws.data[1][4] === 7 && ws.data[3][21] === 'M0003' && ws.getLastRow() === 4, '시트는 그대로(값·행 수·회원ID)');
ok(baks('memberManagement').length === 1, '거절된 저장은 백업도 만들지 않음');

console.log('③ 행이 줄면 남은 줄은 비운다 · 백업은 최근 20개만');
res = call({ action: 'saveMembers', rows: [r23(1, '가상A', 7, 'M0001'), r23(2, '가상B', 10, 'M0002')] });
ok(res.ok && ws.getLastRow() === 3 && ws.data[3].every(v => v === ''), '3명→2명 저장 시 넷째 줄이 비워짐');
for (let i = 0; i < 25; i++) call({ action: 'saveMembers', rows: [r23(1, '가상A', 7, 'M0001')] });
ok(baks('memberManagement').length === 20, '백업 시트는 20개로 유지');
const names = baks('memberManagement').map(s => s.name).sort();
ok(names[0] > 'bak_memberManagement_20260914_120000_006', '오래된 백업부터 지워짐');

console.log('④ 합치기 전 시트(회원ID 열 없음)에는 21열 저장이 지금처럼 된다');
SS = makeSS();
const ws2 = SS.add('memberManagement', [H21, r21(1, '가상A', 7)]);
res = call({ action: 'saveMembers', rows: [r21(1, '가상A', 8)] });
ok(res.ok && ws2.data[1][4] === 8 && ws2.data[0].length === 21, '21열 저장 성공, 새 열 제목은 안 만듦');

console.log('⑤ V1 에 이미 다른 제목이 있으면 덮어쓰지 않는다');
SS = makeSS();
const H = H21.concat(['메모']);
const ws3 = SS.add('memberManagement', [H, r21(1, '가상A', 7).concat(['x'])]);
res = call({ action: 'saveMembers', rows: [r23(1, '가상A', 7, 'M0001')] });
ok(res.ok && ws3.data[0][21] === '메모' && ws3.data[0][22] === '탈퇴일', "V1 '메모' 는 그대로, 빈 W1 만 채움");

console.log('⑥ 통장 저장도 백업을 남긴다 · 다른 동작은 그대로');
SS = makeSS();
const wb = SS.add('bankbook', [['날짜', '금액', '이름', '구분', '내용', '비고', '잔액'], ['2026.01.05', 20000, '가상A', '회비', '1월', '', 20000]]);
res = call({ action: 'saveBank', rows: [['2026.01.05', 20000, '가상A', '회비', '1월', '', 20000], ['2026.01.06', -5000, '물', '물', '', '', 15000]] });
ok(res.ok && baks('bankbook').length === 1 && wb.getLastRow() === 3, '통장 2건 저장 + 백업 1개');
ok(call({ action: 'ping' }).msg === 'pong', 'ping 그대로');
SS.add('memberManagement', [H21]); SS.add('BackNumber', [['이름', '등번호'], ['가상A', 7]]); SS.add('DonationOfGoods', [['날짜', '이름', '내용']]);
const load = call({ action: 'load' });
ok(load.ok && load.backNumber === undefined && load.bankbook.length === 2, 'load 는 BackNumber 를 내려주지 않음(등번호 원본 = 회원 시트 E열)');

console.log('⑦ 예전 등번호 목록 저장(saveBackNumber)은 거절 — BackNumber 시트는 그대로');
const bnBefore = JSON.stringify(SS.getSheetByName('BackNumber').data);
res = call({ action: 'saveBackNumber', rows: [['가상B', 99]] });
ok(res.ok === false && /회원 시트/.test(res.msg), '거절 + 안내 문구');
ok(JSON.stringify(SS.getSheetByName('BackNumber').data) === bnBefore && baks('BackNumber').length === 0, 'BackNumber 시트 값 그대로, 백업도 안 만듦');
console.log('\n통과 ' + n + '개');
