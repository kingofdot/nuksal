// 넉살 index.html 회원 단일 원본(등번호 = 회원 시트 E열) — 가짜 데이터로 로직 시험 (실제 시트·주소를 부르지 않는다)
const fs = require('fs'), vm = require('vm');
const html = fs.readFileSync('D:/nuksal/index.html', 'utf8');
const a = html.lastIndexOf('<script>'), b = html.lastIndexOf('</script>');
const code = html.slice(a + '<script>'.length, b);

function el() {
  return { innerHTML: '', textContent: '', style: {}, value: '', dataset: {}, disabled: false,
    remove() {}, appendChild() {}, querySelector() { return null }, querySelectorAll() { return [] },
    classList: { add() {}, remove() {}, toggle() {} }, addEventListener() {}, focus() {}, select() {} };
}
const elements = {};
const calls = [];
let mode = 'normal';
let LOAD = null;

const sandbox = {
  console, URL, JSON, Math, Date, Promise, Array, Object, String, Number, Set, Map, parseInt, isNaN,
  setTimeout: () => 0, clearTimeout: () => {},
  alert: (m) => { sandbox.__alerts.push(m); }, __alerts: [],
  confirm: () => sandbox.__confirm, __confirm: true,
  navigator: { userAgent: 'node-test', clipboard: {} },
  localStorage: { _s: {}, getItem(k) { return this._s[k] || null }, setItem(k, v) { this._s[k] = v } },
  sessionStorage: { _s: {}, getItem(k) { return this._s[k] || null }, setItem(k, v) { this._s[k] = v } },
  document: {
    getElementById: id => elements[id] || (elements[id] = el()),
    createElement: () => el(), querySelector: () => null, querySelectorAll: () => [],
    body: el(), head: el(),
  },
  async fetch(url, opt = {}) {
    calls.push({ url, opt });
    if (opt.mode === 'no-cors') return {};
    if (opt.method === 'POST') {
      if (mode === 'throwjson') return { json: async () => { throw new Error('CORS 흉내') } };
      if (mode === 'servererr') return { json: async () => ({ ok: false, msg: '서버 오류 흉내' }) };
      const body = JSON.parse(opt.body);
      return { json: async () => ({ ok: true, msg: body.action + ' 완료', count: (body.rows || []).length }) };
    }
    const action = new URL(url).searchParams.get('action');
    if (action === 'load') return { json: async () => ({ ok: true, ...LOAD }) };
    return { json: async () => ({ ok: true, msg: 'ok' }) };
  },
};
sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(code, sandbox, { filename: 'index.html<script>' });

// 21열(예전 형식) 회원 행
const mrow = (seq, group, pos, name, jersey, year, contact, status) =>
  [seq, group, pos, name, jersey, year, '', contact, status, ...Array(12).fill('')];
const OLD_LOAD = {
  bankbook: [['2026.01.05', 20000, '가상A', '회비', '1월', '', 20000]],
  memberManagement: [
    mrow(1, 6, '회원', '가상A', 7, 1990, '010-0000-0001', '정상'),
    mrow(2, 6, '회원', '가상B', 10, 1991, '010-0000-0002', '정상'),
    mrow(3, 6, '회원', '가상C', 0, 1992, '010-0000-0003', '정상'),
    mrow(4, 6, '회원', '가상D', 3, 1993, '010-0000-0004', '정상'),
    mrow(5, 6, '회원', '가상E', 0, 1994, '010-0000-0005', '부상'),
  ],
  backNumber: [['가상A', 7], ['가상B', 11], ['가상C', 5]],   // 예전 GAS 가 보내더라도 쓰지 않아야 한다
  donation: [],
};
LOAD = OLD_LOAD;
sandbox.__setMode = m => { mode = m; };
sandbox.__calls = calls;
sandbox.__elements = elements;
sandbox.__setLoad = d => { LOAD = d; };
sandbox.__OLD_LOAD = OLD_LOAD;
sandbox.__mrow = mrow;

const test = `
(async () => {
  let n = 0;
  const ok = (c, msg) => { if (!c) throw new Error('실패: ' + msg); n++; console.log('  ✓ ' + msg); };
  const posts = () => __calls.filter(c => c.opt.method === 'POST');
  const byName = nm => MEMBER_DATA.find(m => m.name === nm);

  console.log('① 불러오기 — 등번호는 E열만, BackNumber 는 무시');
  await loadAllFromSheets();
  ok(MEMBER_DATA.length === 5, '회원 5명');
  ok(byName('가상B').jersey === 10 && byName('가상C').jersey === 0, 'E열 값 그대로(가상B 10, 가상C 없음) — BackNumber 11·5 는 안 씀');
  ok(state.backNumbers === undefined, 'BackNumber 목록을 화면 상태에 두지 않음');
  ok(MEMBER_DATA.map(m => m.id).join() === 'M0001,M0002,M0003,M0004,M0005', '회원ID 가 없던 5명에게 불러온 순서대로 M0001~M0005');

  console.log('② 저장 전에 새로고침해도 회원ID 는 같다');
  await loadAllFromSheets();
  ok(MEMBER_DATA.map(m => m.id).join() === 'M0001,M0002,M0003,M0004,M0005', '같은 순서 → 같은 회원ID');

  console.log('③ 처음 저장 — 23열로 회원ID 가 같이 적힌다 (막는 단계 없음)');
  const ok1 = await saveMembersToSheet();
  const p = posts().at(-1); const body = JSON.parse(p.opt.body);
  ok(ok1 === true && body.action === 'saveMembers', '저장 요청 성공');
  ok(body.rows.length === 5 && body.rows.every(r => r.length === 23), '5행 × 23열(회원ID·탈퇴일 포함)');
  const rB = body.rows.find(r => r[3] === '가상B');
  ok(rB[4] === 10 && rB[21] === 'M0002' && rB[22] === '', '가상B 행: 등번호 10 · 회원ID M0002 · 탈퇴일 빈칸');
  ok(body.rows.find(r => r[3] === '가상E')[8] === '부상', '상세상태 부상 유지');
  ok(!__calls.some(c => /saveBackNumber/.test(c.url + (c.opt.body || ''))), 'saveBackNumber 는 한 번도 안 부름');

  console.log('④ 탈퇴 = 상태값, 행은 남는다');
  const idxA = MEMBER_DATA.indexOf(byName('가상A'));
  withdrawMember(idxA);
  ok(MEMBER_DATA.length === 5 && byName('가상A').status === '탈퇴' && /^\\d{4}-\\d{2}-\\d{2}$/.test(byName('가상A').withdrawDate), '가상A 상태=탈퇴, 탈퇴일 기록, 명단에서 안 빠짐');
  ok(!getCurrentMembers().includes(byName('가상A')) && getWithdrawnMembers().length === 1, '재적 목록에서 빠지고 탈퇴 목록에 들어감');
  ok(!(jerseyHolders()[7] || []).length, '탈퇴 회원 번호(7)는 비어 있는 번호로 보임');
  await saveMembersToSheet();
  const rowA = JSON.parse(posts().at(-1).opt.body).rows.find(r => r[3] === '가상A');
  ok(rowA && rowA[8] === '탈퇴' && rowA[22] === byName('가상A').withdrawDate, '저장해도 가상A 행이 남고 상세상태=탈퇴·탈퇴일이 들어감');

  console.log('⑤ 등번호 배정 — 한 번호 한 사람 (등번호 탭도 같은 E열 데이터)');
  const idx = nm => String(MEMBER_DATA.indexOf(byName(nm)));
  assignJersey(7, idx('가상B'));
  ok(byName('가상B').jersey === 7, '가상B 에 7번 → 옛 번호 10 은 저절로 풀림');
  assignJersey(5, idx('가상C'));
  assignJersey(5, idx('가상E'));
  ok(byName('가상E').jersey === 5 && byName('가상C').jersey === 0, '가상C 가 쓰던 5번을 가상E 에게 → 가상C 번호 비움');
  __confirm = false;
  assignJersey(3, idx('가상E'));
  ok(byName('가상E').jersey === 5 && byName('가상D').jersey === 3, '확인 창에서 취소하면 아무것도 안 바뀜');
  __confirm = true;
  assignJersey(3, null);
  ok(byName('가상D').jersey === 0, '번호 비우기');

  console.log('⑥ 복귀 — 그 사이 번호를 남이 받았으면 비운다');
  restoreMember(idxA);
  ok(byName('가상A').status === '정상' && byName('가상A').withdrawDate === '' && byName('가상A').jersey === 0, '가상A 복귀, 7번은 가상B 가 써서 비움');

  console.log('⑦ 회원 수정 — 등번호 입력이 바로 되고, 회원ID·상세상태가 사라지지 않는다');
  const set = (id, v) => { __elements[id] = Object.assign(__elements[id] || {}, { value: v, remove() {} }); };
  const idxE = MEMBER_DATA.indexOf(byName('가상E'));
  set('mm_name', '가상E'); set('mm_pos', '회원'); set('mm_year', '1994'); set('mm_contact', '010-0000-0005'); set('mm_jersey', '7');
  for (let i = 1; i <= 12; i++) set('mm_m' + i, '');
  set('memberModalOverlay', '');
  saveMember(idxE);
  ok(byName('가상E').id === 'M0005' && byName('가상E').status === '부상', '회원ID M0005 · 상세상태 부상 그대로');
  ok(byName('가상E').jersey === 7 && byName('가상B').jersey === 0, '7번으로 바꾸면 가상B 번호를 비움(확인 후)');
  set('mm_name', '가상F'); set('mm_jersey', '');
  saveMember(null);
  ok(byName('가상F') && byName('가상F').id === 'M0006' && byName('가상F').status === '정상', '새 회원은 M0006 · 정상');

  console.log('⑧ 저장 실패를 성공으로 띄우지 않는다');
  __setMode('servererr');
  const nPost = posts().length;
  ok(await saveMembersToSheet() === false && posts().length === nPost + 1, '서버 오류면 실패로 끝남 — 다시 보내지 않음');
  __setMode('throwjson');
  const r2 = await saveMembersToSheet();
  const last2 = __calls.slice(-2);
  ok(r2 === false && last2[1].opt.mode === 'no-cors', '응답을 못 읽으면 한 번 더 보내고 "확인 못함"으로 끝남(false)');
  __setMode('normal');
  MEMBER_DATA.length = 0;
  ok(await saveMembersToSheet() === false, '회원이 비어 있으면 저장 안 함');

  console.log('⑨ 저장된 시트 다시 불러오기 + 시트에 직접 넣은 회원ID 없는 행');
  __setLoad({ bankbook: [], memberManagement: body.rows.concat([__mrow(6, 6, '회원', '가상G', 22, 1995, '', '정상')]), backNumber: [['가상B', 11]], donation: [] });
  await loadAllFromSheets();
  ok(byName('가상B').jersey === 10 && byName('가상B').id === 'M0002', '가상B = 시트 값 10 · 회원ID 그대로');
  ok(byName('가상G').id === 'M0006' && byName('가상G').jersey === 22, '회원ID 없는 새 행은 다음 번호 M0006');

  console.log('⑩ 모든 탭 화면이 오류 없이 그려진다');
  for (const t of ['bank', 'dues', 'report', 'members', 'jersey']) { state.tab = t; render(); }
  state.memberSubTab = 'withdrawn'; state.tab = 'members'; render();
  openMemberModal(0); openJerseyPicker(10);
  ok(true, '통장·회비·보고서·회원(재적/탈퇴)·등번호 탭 + 회원 수정·번호 배정 창');
  console.log('\\n통과 ' + n + '개');
})().catch(e => { console.error(e.stack || e); process.exitCode = 1; });
`;
sandbox.process = process;
vm.runInContext(test, sandbox, { filename: 'test' });
