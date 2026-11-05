// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 작업일보는 아직 받지 않았습니다(기획서 10장 ①). 여기서는 제출 글의 항목대로 만든 작은 가상 표와 가상 샘플 12개월로 검사합니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const E = require('../js/ai-endpoint.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const near = (a, b, eps = 1e-9) => assert.ok(Math.abs(a - b) < eps, a + ' ≠ ' + b);

// 제출 글의 작업일보 항목 그대로인 작은 표 (제목 · 빈 줄 · 머리 · 자료 · 합계)
const HEAD = ['작업일자', '작업자', 'Hydro press Main 압력(bar)', '쿠션 압력(bar)', '금형', '작업모델', '소재 메이커', '소재규격', '작업수량', '불량수량'];
const TINY = [
  ['■ 작업일보'], [], HEAD,
  ['2026-01-05', '작업자A', 200, 50, 'D-1', 'M-1', '메이커가', 'STS 0.8t', 1000, 50],
  ['2026-01-05', '작업자B', '210 bar', 52, 'D-2', 'M-2', '메이커나', 'STS 1.0t', '1,000', 10],
  ['2026-01-06', '작업자A', 205, 70, 'D-1', 'M-1', '메이커가', 'STS 0.8t', 500, null],      // 불량 빈칸 → 0
  ['2026-01-06', '작업자A', 205, 70, 'D-1', 'M-1', '메이커가', 'STS 0.8t', 500, null],      // 같은 날 같은 조건 두 번 → 따로 셈
  ['2026-01-07', '작업자B', 200, 50, 'D-2', 'M-2', '메이커나', 'STS 1.0t', 100, 120],       // 불량 > 작업 → 건너뜀
  ['2026-01-07', '작업자B', 200, 50, 'D-2', 'M-2', '메이커나', 'STS 1.0t', 0, 0],           // 작업 0 → 건너뜀
  ['날짜아님', '작업자B', 200, 50, 'D-2', 'M-2', '메이커나', 'STS 1.0t', 100, 1],           // 날짜 못 읽음 → 건너뜀
  [],
  ['합계', null, null, null, null, null, null, null, 3100, 60]
];

console.log('열 찾기');
test('제목·빈 줄 아래 머리행을 찾고 제출 글의 10개 항목을 모두 연결', () => {
  const d = L.detectColumns(TINY);
  assert.equal(d.headerRow, 2);
  ['date', 'worker', 'main', 'cushion', 'die', 'model', 'maker', 'spec', 'qty', 'defect'].forEach((f, i) => assert.equal(d.map[f], i, f));
  assert.deepEqual(d.missing, []);
});
test('열 이름이 달라도 찾음 — 일자·Main 압력·쿠션압력·금형번호·모델·메이커·규격, 단위·괄호 무시', () => {
  const d = L.detectColumns([Sample.HEAD_B, ['2026-04-01', '작업자A', 200, 50, 'D', 'M', 'K', 'K-2604', 'S', 100, 1, 1, 0, 0, '']]);
  assert.equal(d.headerRow, 0);
  assert.equal(d.map.date, 0); assert.equal(d.map.main, 2); assert.equal(d.map.cushion, 3); assert.equal(d.map.die, 4);
  assert.equal(d.map.model, 5); assert.equal(d.map.maker, 6); assert.equal(d.map.lot, 7, 'LOT No → 소재 로트'); assert.equal(d.map.spec, 8);
  assert.deepEqual(d.typeCols.map(x => x.type), ['Crack', '주름', '기타']);
  assert.equal(L.matchField('Main 압력 (MPa)'), 'main');
  assert.equal(L.matchField('쿠션 설정압'), 'cushion');
  assert.equal(L.matchField('소재 MAKER'), 'maker');
  assert.equal(L.matchField('작업수량(EA)'), 'qty');
  assert.equal(L.matchType('CRACK 수량'), 'Crack');
});
test('두 줄 머리(윗줄 「압력」, 아랫줄 「Main」「쿠션」)도 윗줄을 붙여 찾음', () => {
  const aoa = [[null, null, '압력', null, null, null, null], ['작업일자', '금형', 'Main', '쿠션', '작업수량', '불량수량', '불량유형'], ['2026-01-05', 'D-1', 200, 50, 100, 3, 'Crack']];
  const d = L.detectColumns(aoa);
  assert.equal(d.headerRow, 1); assert.equal(d.groupRow, 0);
  assert.equal(d.map.main, 2); assert.equal(d.map.cushion, 3);
  const p = L.parseSheet(aoa);
  assert.deepEqual(p.records[0].types, { Crack: 3 }, '불량유형 글자 칸 → 그 행 불량 전부를 그 유형으로');
});
test('필수 열이 없으면 알리고 열 연결(덮어쓰기)로 고칠 수 있음', () => {
  const aoa = [['작업일자', '금형', '생산 실적', '불량수량'], ['2026-01-05', 'D-1', 100, 2]];
  const d = L.detectColumns(aoa);
  assert.equal(d.headerRow, -1, '작업수량이 없으면 머리행으로 보지 않음');
  const aoa2 = [['작업일자', '금형', '작업수량', 'NG 개수'], ['2026-01-05', 'D-1', 100, 2]];
  const d2 = L.detectColumns(aoa2);
  assert.deepEqual(d2.missing, ['defect']);
  const d3 = L.detectColumns(aoa2, { defect: 3 });
  assert.deepEqual(d3.missing, []);
  assert.equal(L.parseSheet(aoa2, { overrides: { defect: 3 } }).records[0].defect, 2);
});
test('여러 시트 중 작업일보 시트를 고름', () => {
  const s = L.pickSheet([{ name: '요약', aoa: [['월', '불량률'], ['1월', 0.05]] }, { name: '일보', aoa: TINY }]);
  assert.equal(s.name, '일보');
});

console.log('행 읽기');
test('행 점검 — 단위 붙은 숫자·쉼표 읽기, 빈 불량 = 0, 불량>작업·작업 0·날짜 오류는 건너뜀, 합계 줄 제외', () => {
  const p = L.parseSheet(TINY, { fileName: 't.xlsx' });
  assert.equal(p.records.length, 4);
  assert.equal(p.records[1].main, 210); assert.equal(p.records[1].qty, 1000);
  assert.equal(p.records[2].defect, 0);
  assert.equal(p.skipped, 3);
  assert.equal(p.summary, 1);
  assert.ok(p.warnings.some(w => w.msg.includes('작업수량(100)보다 커서')));
  assert.ok(p.det.warnings.some(w => w.includes('빈 행 2개는 불량 0')));
});
test('같은 날 같은 조건 두 행은 따로 세고(열쇠 #1·#2), 같은 파일을 다시 넣으면 두 번 세지 않음', () => {
  const p = L.parseSheet(TINY);
  assert.notEqual(p.records[2].key, p.records[3].key);
  const m1 = L.mergeRecords([], p.records);
  const m2 = L.mergeRecords(m1.records, L.parseSheet(TINY).records);
  assert.equal(m2.records.length, 4);
  assert.deepEqual(m2.stat, { added: 0, updated: 4 });
});
test('날짜 — 엑셀 일련번호·Date·「2026.01.05(월)」·「1/5」(연도 추정)', () => {
  assert.equal(L.toISO(46027), '2026-01-05');
  assert.equal(L.toISO(new Date('2026-01-05T00:00:00')), '2026-01-05');
  assert.equal(L.toISO('2026.01.05(월)'), '2026-01-05');
  assert.equal(L.toISO('1/5', 2026), '2026-01-05');
  assert.equal(L.toISO('2026-02-30'), '');
  assert.equal(L.toNum('1,200개'), 1200); assert.equal(L.toNum('12.5 MPa'), 12.5); assert.ok(isNaN(L.toNum('없음')));
});

console.log('집계');
test('불량률은 작업량 가중(합계 ÷ 합계) — 행 불량률의 평균이 아님', () => {
  const recs = L.parseSheet(TINY).records;
  const a = L.aggregate(recs);
  assert.equal(a.qty, 3000); assert.equal(a.cnt, 60);
  near(a.rate, 0.02);
  const meanOfRates = recs.reduce((s, r) => s + r.defect / r.qty, 0) / recs.length;
  assert.ok(Math.abs(meanOfRates - a.rate) > 0.001, '평균과 달라야 이 검사가 뜻이 있음');
  const g = L.groupBy(L.buildContext(recs), 'die', 'all', 'rate');
  assert.equal(g.rows[0].key, 'D-1'); near(g.rows[0].rate, 50 / 2000); near(g.rows[1].rate, 10 / 1000);
  near(g.rows[0].excess, 50 - 2000 * 0.02);
});
test('계절·월 — 12~2월 겨울, 월별 추이와 목표 초과 표시', () => {
  assert.deepEqual(['2026-01-05', '2026-03-01', '2026-06-30', '2026-09-01', '2025-12-31'].map(L.seasonOf), ['겨울', '봄', '여름', '가을', '겨울']);
  const ctx = L.buildContext(L.parseSheet(TINY).records, { settings: { target: 0.015 } });
  const t = L.trend(ctx, 'month');
  assert.equal(t.length, 1); assert.equal(t[0].key, '2026-01'); assert.equal(t[0].overTarget, true);
  const d = L.trend(ctx, 'day', 'all', '2026-01');
  assert.deepEqual(d.map(x => x.key), ['2026-01-05', '2026-01-06']);
  near(d[1].rate, 0);
});
test('압력 구간 — 같은 폭 구간, 정상 범위(설정) 아래·안·위', () => {
  const b = L.makeBins([41, 49, 55, 66, 74], 10);
  assert.equal(b.step, 5); assert.equal(b.start, 40);
  assert.equal(L.binOf(b, 66), 65); assert.equal(L.binOf(b, 65), 65); assert.equal(L.binOf(b, 64.9), 60);
  const ctx = L.buildContext(L.parseSheet(TINY).records, { settings: { cushionMin: 45, cushionMax: 60 } });
  const g = L.groupBy(ctx, 'cushionRange');
  assert.deepEqual(g.rows.map(r => r.key), ['in', 'above']);
  assert.equal(g.rows[1].qty, 1000);
  assert.equal(L.dimAvailable(L.buildContext(L.parseSheet(TINY).records), 'cushionRange'), false, '범위를 안 넣으면 그 기준은 없음');
});
test('표본 적음 — 작업량·행 기준', () => {
  const ctx = L.buildContext(L.parseSheet(TINY).records, { settings: { minQty: 1500, minRows: 2 } });
  const g = L.groupBy(ctx, 'die');
  assert.deepEqual(g.rows.map(r => [r.key, r.small]), [['D-1', false], ['D-2', true]]);
});
test('윌슨 구간 — 작업량이 적으면 넓고, 많으면 좁음', () => {
  const a = L.wilson(5, 100), b = L.wilson(500, 10000);
  assert.ok(a.lo < 0.05 && a.hi > 0.05 && b.lo < 0.05 && b.hi > 0.05);
  assert.ok((a.hi - a.lo) > 5 * (b.hi - b.lo));
});

console.log('금형 보수 이력');
test('붙여 넣기 글 읽기 — 머리 줄 건너뛰기, 탭·쉼표, 잘못된 줄 번호, 같은 이력 하나로', () => {
  const r = L.parseMaintText('금형\t일자\t구분\t내용\nD-1\t2026-01-06\t보수\tR부 연마\nD-1,2026-01-06,보수\nD-2\t어제\t교체\n\nD-2\t2026/02/01\t교체');
  assert.equal(r.items.length, 2);
  assert.deepEqual(r.bad, [4]);
  assert.equal(r.items[1].kind, '교체'); assert.equal(r.items[1].date, '2026-02-01');
});
test('누적 작업량 — 보수일 작업부터 0, 이력보다 앞선 행은 시작점 모름', () => {
  const recs = L.parseSheet(TINY).records; // D-1: 01-05 1000, 01-06 500, 01-06 500
  const w = L.computeWear(recs, [{ die: 'D-1', date: '2026-01-06', kind: '보수' }]);
  const d1 = recs.filter(r => r.die === 'D-1');
  assert.equal(w[d1[0].key].known, false);
  assert.deepEqual([w[d1[1].key].cum, w[d1[2].key].cum], [0, 500]);
  assert.equal(w[d1[1].key].known, true);
  assert.equal(w[recs.find(r => r.die === 'D-2').key].known, false, '이력 없는 금형');
});

console.log('가상 샘플 12개월 — 심어 둔 패턴을 찾는가');
const SB = Sample.build();
let SRECS = [];
for (const f of SB.files) {
  // 앱과 같은 길: 엑셀 통합문서로 만들었다가 다시 읽기
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(f.aoa), '작업일보');
  const back = XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellDates: true });
  const p = L.parseSheet(XLSX.utils.sheet_to_json(back.Sheets['작업일보'], { header: 1, raw: true, defval: null }), { fileName: f.fileName });
  assert.equal(p.warnings.length, 0, f.fileName + ' 경고');
  SRECS = L.mergeRecords(SRECS, p.records).records;
}
const CTX = L.buildContext(SRECS, { maint: SB.maint, stds: SB.stds, repairs: SB.repairs });
const AN = L.analyze(CTX, 'all');
test('샘플 — 12개 파일 1,130행 모두, 빠진 달 없음, 불량 합계가 원본과 같음', () => {
  assert.equal(SRECS.length, SB.rows.length);
  assert.equal(L.monthCoverage(SRECS).missing.length, 0);
  assert.equal(L.monthsOf(SRECS).length, 12);
  assert.equal(L.aggregate(SRECS).defect, SB.rows.reduce((s, r) => s + r.defect, 0));
  assert.ok(AN.overall.rate > 0.04 && AN.overall.rate < 0.065, '제출 글의 「약 5%」 수준: ' + AN.overall.rate);
});
const hotText = AN.hot.list.map(h => L.condText(h.parts));
test('심어 둔 ① 겨울 × 소재사B 를 조합으로 찾음(둘 각각보다 높음)', () => {
  assert.ok(hotText.includes('계절 겨울 × 소재 메이커 소재사B'), hotText.join(' / '));
  const winterB = AN.hot.list.find(h => L.condText(h.parts) === '계절 겨울 × 소재 메이커 소재사B');
  assert.ok(winterB.agg.rate > AN.dims.season.rows.find(r => r.key === '겨울').rate * 1.2);
});
test('심어 둔 ② 쿠션 압력 65 이상 구간을 찾음 — Crack 으로 보면 더 뚜렷', () => {
  assert.ok(AN.hot.list.some(h => h.parts.length === 1 && h.parts[0].dim === 'cushion' && h.parts[0].key >= 65));
  const crack = L.analyze(CTX, 'Crack');
  const c = crack.dims.cushion.rows.find(r => +r.key === 70), n = crack.dims.cushion.rows.find(r => +r.key === 50);
  assert.ok(c.rate > n.rate * 2, 'Crack: 70~75 ' + c.rate + ' vs 50~55 ' + n.rate);
});
test('심어 둔 ③ 금형 DIE-B01 마모 — 누적 작업량이 늘수록 오르고, 보수하면 떨어짐', () => {
  const b01 = AN.dies.find(d => d.die === 'DIE-B01');
  const rates = b01.curve.filter(c => !c.small).map(c => c.rate);
  for (let i = 1; i < rates.length; i++) assert.ok(rates[i] > rates[i - 1], '누적 구간마다 오름: ' + rates.join(', '));
  assert.ok(b01.threshold != null, '보수 직후보다 1.3배 넘는 구간');
  const eff = AN.maint.filter(m => m.die === 'DIE-B01' && m.enough);
  assert.ok(eff.filter(m => m.change < 0).length >= 3, '보수 뒤 내려감: ' + eff.map(m => m.change).join(', '));
  const others = AN.dies.filter(d => d.die !== 'DIE-B01' && d.curve.length);
  others.forEach(d => assert.equal(d.threshold, null, d.die + ' 은 마모 패턴 없음'));
});
test('대조군 — Main 압력은 홀로 걸리면 반드시 「겹침」(금형)이 붙음, 작업자는 걸리지 않음', () => {
  AN.hot.list.filter(h => h.parts.some(p => p.dim === 'main')).forEach(h => {
    assert.ok(h.overlap.some(o => o.dim === 'die' && o.label === 'DIE-B01'), L.condText(h.parts) + ' 겹침 없음');
  });
  assert.ok(!AN.hot.list.some(h => h.parts.length === 1 && h.parts[0].dim === 'worker'));
});
test('같은 행 묶음(금형=모델=규격 1:1)은 한 번만 — 나머지는 「같은 조건」으로', () => {
  const sigs = AN.hot.list.map(h => h.agg.rows + '|' + h.agg.qty + '|' + h.agg.cnt);
  assert.equal(new Set(sigs).size, sigs.length);
  assert.ok(!hotText.some(t => t.includes('작업모델 SMP-H210') && t.includes('겨울')), '겨울 × 모델은 겨울 × 금형과 같은 행이라 빠짐');
});
test('불량이 몰린 조건 거름망 — 기준을 올리면 줄고, 근거 숫자가 맞음', () => {
  const strict = L.findHotspots(L.buildContext(SRECS, { maint: SB.maint, settings: { minLift: 1.8 } }), 'all');
  assert.ok(strict.total < AN.hot.total && strict.total > 0);
  strict.list.forEach(h => assert.ok(h.agg.lift >= 1.8 && !h.agg.small));
  const h = AN.hot.list[0];
  near(h.agg.excess, h.agg.cnt - h.agg.qty * AN.overall.rate, 1e-6);
  assert.match(L.hotEvidence(h, AN.hot.overall), /작업 [\d,]+개 중 불량 [\d,]+개\([\d.]+%\) — 전체 [\d.]+%의 [\d.]+배/);
});
test('두 조건 조합은 「조합만의 효과」만 — 두 조건 각각보다 pairGain 배 이상, 하한도 더 높음', () => {
  const S0 = CTX.S, single = {};
  Object.keys(L.DIMS).forEach(d => { if (L.dimAvailable(CTX, d)) L.groupBy(CTX, d).rows.forEach(r => { single[d + '=' + r.key] = r.rate; }); });
  const all = L.findHotspots(CTX, 'all', 999).list.filter(h => h.parts.length === 2);
  assert.ok(all.length > 0);
  all.forEach(h => {
    const parent = Math.max(...h.parts.map(p => single[p.dim + '=' + p.key]));
    assert.ok(h.agg.rate >= parent * S0.pairGain - 1e-12 && h.agg.lo > parent, L.condText(h.parts) + ' ' + h.agg.rate + ' vs ' + parent);
  });
});
test('교차표 — 계절 × 소재 메이커의 칸 합계 = 전체', () => {
  const ct = L.crossTab(CTX, 'season', 'maker');
  let q = 0; ct.rows.forEach(r => ct.cols.forEach(c => { const a = ct.cell(r.key, c.key); if (a) q += a.qty; }));
  assert.equal(q, AN.overall.qty);
  assert.deepEqual(ct.rows.map(r => r.key), ['봄', '여름', '가을', '겨울']);
});
test('기간 거르기 — 누적 작업량은 기간 앞 작업까지 세어 같은 값', () => {
  const part = L.buildContext(SRECS, { maint: SB.maint, from: '2026-03-01', to: '2026-03-31' });
  const r = part.recs.find(x => x.die === 'DIE-B01' && x.date === '2026-03-10');
  const full = CTX.recs.find(x => x.key === r.key);
  assert.equal(r.wear, full.wear);
  assert.ok(part.recs.every(x => x.date >= '2026-03-01' && x.date <= '2026-03-31'));
});

console.log('AI 프롬프트 · 가리기');
const AL = L.assignAllAliases({}, SRECS);
const INFO = L.buildAiPromptInfo(AN, AL);
test('가리기 — 금형·모델·메이커·작업자 이름이 프롬프트에 없고, 숫자는 그대로', () => {
  ['DIE-', 'SMP-', '소재사', '작업자A', '작업자B'].forEach(x => assert.ok(!INFO.prompt.includes(x), x + ' 가 남음'));
  assert.deepEqual(INFO.leaks, []);
  assert.ok(INFO.prompt.includes(AL.maker['소재사B']), '메이커 별칭');
  const plain = L.buildAiPrompt(AN);
  assert.ok(plain.includes('DIE-B01'));
  const nums = s => s.replace(/(금형|모델|메이커|작업자|로트)\d+/g, 'X').replace(/\n금형·모델[^\n]*/, '').match(/\d[\d,.]*/g).join(' ');
  assert.equal(nums(INFO.prompt), nums(plain.replace(/DIE-[A-Z]\d+|SMP-[A-Z]\d+|소재사[A-C]|작업자[A-E]|[A-C]-\d{4}/g, 'X')));
  assert.ok(!/[A-C]-\d{4}/.test(INFO.prompt), '소재 로트 이름도 가림');
});
test('프롬프트 — 사실만·추정 표시·목표·4M 요청이 들어 있음', () => {
  assert.match(INFO.prompt, /숫자를 새로 만들거나 바꾸지 말아줘/);
  assert.match(INFO.prompt, /「추정」/);
  assert.match(INFO.prompt, /목표 불량률 2%/);
  assert.match(INFO.prompt, /4M/);
  assert.match(INFO.prompt, /겹침: 이 조건 작업의 \d+%가 금형\d/);
  assert.ok(!/비고|첫 작업/.test(INFO.prompt.split('[사실]')[1]), '비고(자유 글)는 넣지 않음');
});
test('별칭 — 같은 값은 늘 같은 별칭, 새 값만 다음 번호', () => {
  const a1 = L.assignAliases({}, 'die', ['B', 'A']);
  assert.deepEqual(a1, { A: '금형1', B: '금형2' });
  const a2 = L.assignAliases({ die: a1 }, 'die', ['C', 'A']);
  assert.deepEqual(a2, { A: '금형1', B: '금형2', C: '금형3' });
});
test('AI 답 되돌리기 — 별칭 → 원래 이름, 「금형 3개」「금형3개」는 그대로', () => {
  const ans = AL.die['DIE-B01'] + ' 보수 주기 확인, ' + AL.maker['소재사B'] + ' 겨울 로트 확인. 금형 3개, 금형3개는 그대로.';
  const back = L.unmaskText(ans, AL);
  assert.ok(back.startsWith('DIE-B01 보수 주기 확인, 소재사B 겨울 로트 확인.'), back);
  assert.ok(back.endsWith('금형 3개, 금형3개는 그대로.'));
});
test('안전망 — 남은 이름을 찾음, 더 긴 다른 이름 속은 아님', () => {
  assert.deepEqual(L.findLeaks('DIE-B01 확인, DIE-B012 는 다름', { die: { 'DIE-B01': '금형1', 'DIE-A01': '금형2' } }), ['DIE-B01']);
  assert.deepEqual(L.findLeaks('DIE-B012 만', { die: { 'DIE-B01': '금형1' } }), []);
  assert.deepEqual(L.findLeaks('수량 12', { die: { '12': '금형1' } }), [], '짧은 숫자 이름은 숫자와 헷갈려 검사하지 않음');
});

console.log('3차 반영 — 소재 메이커·규격 · 소재 lot · 압력 작업표준 · 압력 저하 전파 · 설비 수선');
test('두 줄 머리 — 윗줄 「소재」(가로 병합) 아래 「메이커」「규격」이면 메이커·규격을 제자리로 (전에는 규격 칸에 메이커 값)', () => {
  // 작업일자·금형처럼 세로로 병합된 칸은 아랫줄이 비고, 「압력」「소재」는 가로로 병합된 모양
  const aoa = [
    ['■ 작업일보'],
    ['작업일자', '작업자', '압력', null, '금형', '작업모델', '소재', null, null, '작업수량', '불량수량'],
    [null, null, 'Main', '쿠션', null, null, '메이커', 'lot', '규격', null, null],
    ['2026-01-05', '작업자A', 200, 50, 'D-1', 'M-1', '소재사A', 'A-2601', 'STS304 0.8t', 1000, 20],
    ['2026-01-05', '작업자B', 210, 52, 'D-2', 'M-2', '소재사B', 'B-2601', 'STS430 1.0t', 500, 5]
  ];
  const p = L.parseSheet(aoa);
  assert.equal(p.det.headerRow, 2); assert.equal(p.det.groupRow, 1);
  assert.deepEqual([p.det.map.main, p.det.map.cushion, p.det.map.maker, p.det.map.lot, p.det.map.spec], [2, 3, 6, 7, 8]);
  assert.equal(p.records.length, 2);
  assert.deepEqual(p.records.map(r => [r.maker, r.lot, r.spec, r.main, r.cushion]), [['소재사A', 'A-2601', 'STS304 0.8t', 200, 50], ['소재사B', 'B-2601', 'STS430 1.0t', 210, 52]]);
  const ctx = L.buildContext(p.records);
  assert.deepEqual(L.groupBy(ctx, 'maker', 'all', 'order').rows.map(r => r.label), ['소재사A', '소재사B']);
  assert.deepEqual(L.groupBy(ctx, 'spec', 'all', 'order').rows.map(r => r.label), ['STS304 0.8t', 'STS430 1.0t']);
  // 자료 줄을 아랫머리로 잘못 붙이지 않음(한 줄 머리 작업일보는 그대로)
  assert.equal(L.detectColumns(TINY).headerRow, 2);
});
test('압력 칸 이름 — 「기준·표준」은 작업표준 칸, 「설정압」은 지금처럼 압력 값, 설정·실제가 둘 다 있으면 실제를 압력으로', () => {
  assert.equal(L.matchField('Main 압력 기준'), 'mainSpec'); assert.equal(L.matchField('쿠션 압력(표준)'), 'cushionSpec');
  assert.equal(L.matchField('쿠션 설정압'), 'cushion'); assert.equal(L.matchField('호기'), 'press');
  const aoa = [['작업일자', '금형', 'Main 압력', null, '작업수량', '불량수량'], [null, null, '설정', '실제', null, null], ['2026-01-05', 'D-1', 235, 214, 100, 2]];
  const p = L.parseSheet(aoa);
  assert.equal(p.records[0].main, 214); assert.equal(p.records[0].mainSpec, 235);
  assert.ok(p.det.warnings.some(w => w.includes('실제 값')));
});
const FILE_READ = (() => {
  // 리포의 samples/*.xlsx 를 앱과 같은 길로 다시 읽기 — 내보낸 파일의 칸 순서가 맞는지까지 봄
  const fs = require('fs'), path = require('path');
  const dir = path.join(path.dirname(require('url').fileURLToPath(import.meta.url)), '..', 'samples');
  let recs = [];
  fs.readdirSync(dir).filter(n => n.startsWith('작업일보_샘플_')).sort().forEach(n => {
    const wb = XLSX.read(fs.readFileSync(path.join(dir, n)), { type: 'buffer', cellDates: true });
    recs = L.mergeRecords(recs, L.parseSheet(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null }), { fileName: n }).records).records;
  });
  return recs;
})();
test('② 항목별 — 소재 메이커·소재 lot·소재규격 표가 원본 값으로 따로 센 것과 같음(서로 바뀌지 않음, 샘플 xlsx 파일도)', () => {
  const by = (rows, k) => { const m = {}; rows.forEach(r => { const x = m[r[k]] || (m[r[k]] = { qty: 0, d: 0 }); x.qty += r.qty; x.d += r.defect; }); return m; };
  [['maker', SB.rows], ['spec', SB.rows], ['lot', SB.rows]].forEach(([k, rows]) => {
    const want = by(rows, k), g = L.groupBy(CTX, k, 'all', 'order');
    assert.deepEqual(g.rows.map(r => r.label).sort(), Object.keys(want).sort(), k + ' 이름');
    g.rows.forEach(r => { assert.equal(r.qty, want[r.label].qty, k + ' ' + r.label); assert.equal(r.cnt, want[r.label].d); });
  });
  assert.ok(L.groupBy(CTX, 'maker').rows.every(r => /^소재사[A-C]$/.test(r.label)));
  assert.ok(L.groupBy(CTX, 'spec').rows.every(r => /^STS\d{3} \d\.\dt$/.test(r.label)));
  assert.equal(FILE_READ.length, SB.rows.length);
  const key = r => r.date + '|' + r.die + '|' + r.qty + '|' + r.defect;
  const orig = {}; SB.rows.forEach(r => { orig[key(r)] = r; });
  FILE_READ.forEach(r => { const o = orig[key(r)]; assert.ok(o, key(r)); assert.deepEqual([r.maker, r.lot, r.spec], [o.maker, o.lot, o.spec], key(r)); });
});
test('같은 행 묶음 알림 — 샘플은 금형마다 소재규격이 하나라 소재규격 표 = 금형 표, 메이커는 아님', () => {
  assert.equal(L.sameGrouping(CTX, 'spec', 'die'), true);
  assert.equal(L.sameGrouping(CTX, 'spec', 'model'), true);
  assert.equal(L.sameGrouping(CTX, 'maker', 'die'), false);
  assert.equal(L.sameGrouping(CTX, 'lot', 'maker'), false, '로트는 메이커보다 잘게 나뉨');
  const d = L.groupBy(CTX, 'die', 'all', 'order').rows.map(r => r.qty).sort(), sp = L.groupBy(CTX, 'spec', 'all', 'order').rows.map(r => r.qty).sort();
  assert.deepEqual(d, sp);
});
test('소재 lot — 메이커·달마다 로트, 심어 둔 로트 A-2607 이 「불량이 몰린 조건」에 걸리고 메이커 겹침이 붙음', () => {
  const g = L.groupBy(CTX, 'lot', 'all', 'rate');
  assert.equal(g.rows.length, 36);
  const h = AN.hot.list.find(x => x.parts.length === 1 && x.parts[0].dim === 'lot' && x.parts[0].key === Sample.BAD_LOT);
  assert.ok(h, AN.hot.list.map(x => L.condText(x.parts)).join(' / '));
  assert.ok(h.overlap.some(o => o.dim === 'maker' && o.label === '소재사A'));
  const july = L.groupBy(L.buildContext(SRECS, { from: '2026-07-01', to: '2026-07-31' }), 'lot', 'Crack', 'rate').rows;
  assert.equal(july[0].key, Sample.BAD_LOT, '7월 Crack 1위 로트');
});
test('작업표준 표 — 붙여 넣기(머리 줄·탭·쉼표·못 읽은 줄), 허용 비우면 0, 작업일보의 기준 칸이 표보다 먼저', () => {
  const r = L.parseStdText('작업모델\tMain 기준\tMain 허용\t쿠션 기준\t쿠션 허용\nM-1\t200\t10\t50\t5\nM-2,210,,52,\n\nM-3\t없음\t\t\t\nM-1\t999\t1\t1\t1');
  assert.deepEqual(r.bad, [5]);
  assert.deepEqual(r.stds.rows.map(x => [x.key, x.main, x.mainTol, x.cushion, x.cushionTol]), [['M-1', 200, 10, 50, 5], ['M-2', 210, null, 52, null]]);
  const st = L.stdOf(r.stds, { model: 'M-2', main: 205 });
  assert.equal(st.main, 210); assert.equal(st.mainTol, 0); assert.equal(st.mainFrom, '표');
  assert.equal(L.stdOf(r.stds, { model: 'M-2', mainSpec: 230 }).main, 230, '칸 값이 먼저');
  assert.equal(L.stdOf(L.cleanStds({ by: 'die', rows: [{ key: 'D-1', main: 100 }] }), { die: 'D-1', model: 'M-1' }).main, 100, '금형으로 찾기');
  const dv = L.devOf(r.stds, { model: 'M-1', main: 190, cushion: 44 });
  assert.equal(dv.mainDev, -10); assert.equal(dv.mainLow, false, '허용 폭과 같으면 저하 아님'); assert.equal(dv.cushionLow, true);
});
test('기준 대비 — 샘플: 쿠션 기준보다 높게 세팅한 작업의 불량률이 높고, 모델별 표가 기준·실제를 따로 보여 줌', () => {
  const g = AN.dims.cushionStd;
  const hi = g.rows.find(r => r.key === 'high'), ok = g.rows.find(r => r.key === 'ok');
  assert.ok(hi.rate > ok.rate * 1.5, hi.rate + ' vs ' + ok.rate);
  assert.equal(AN.std.cushion.length, 5);
  const b01 = AN.std.main.find(x => x.key === 'SMP-H210');
  assert.equal(b01.std, 235); assert.equal(b01.tol, 15);
  assert.ok(Math.abs(b01.devMean) < 3, '평소에는 기준 둘레: ' + b01.devMean);
  assert.ok(AN.dims.mainDev && AN.dims.cushionDev);
  // Main 은 대조군 — 기준보다 낮아도 불량률이 뚜렷이 오르지 않음
  const mlow = AN.dims.mainStd.rows.find(r => r.key === 'low');
  assert.ok(mlow.rate < AN.overall.rate * 1.3);
});
test('작업일보 압력 값이 기준 값 그대로여도 돎 — 차이 0 · 전파 0 · 「기준 값을 적은 것」 표시', () => {
  const same = SRECS.map(r => Object.assign({}, r, { main: SB.stds.rows.find(x => x.key === r.model).main, cushion: SB.stds.rows.find(x => x.key === r.model).cushion }));
  const sug = L.suggestStds(same, 'model');
  assert.deepEqual(sug.rows.map(x => [x.key, x.main, x.cushion]), SB.stds.rows.map(x => [x.key, x.main, x.cushion]).sort());
  const an = L.analyze(L.buildContext(same, { stds: SB.stds }), 'all');
  assert.equal(an.alerts.length, 0); assert.equal(an.std.zeroShare, 1);
  assert.match(L.buildAiPrompt(an), /기준\(설정\) 값을 적은 것으로 보임/);
  assert.match(L.buildReport(an), /실제 압력을 적기 시작하면/);
  assert.equal(L.analyze(L.buildContext(SRECS), 'all').std.has, false, '작업표준이 없으면 기준 대비는 비어 있음');
});
test('압력 저하 전파 목록 — 기준 − 허용보다 낮은 행만, 심어 둔 두 번의 저하가 한 묶음으로, 글자·엑셀', () => {
  const al = AN.alerts;
  al.forEach(a => assert.ok(a.dev < -a.tol, a.date + ' ' + a.dev));
  const cu = al.filter(a => a.kind === 'cushion'), inDrop = cu.filter(a => a.date >= Sample.CUSHION_DROP.from && a.date <= Sample.CUSHION_DROP.to);
  assert.ok(inDrop.length >= 25 && inDrop.length / cu.length > 0.6, inDrop.length + ' / ' + cu.length);
  const mn = al.filter(a => a.kind === 'main' && a.date >= Sample.MAIN_DROP.from && a.date <= Sample.MAIN_DROP.to);
  assert.ok(mn.length >= 10);
  const t = L.alertText(al, { from: AN.from, to: AN.to });
  assert.match(t, /^\[설비보전·공무팀 전파\] Hydro press 압력 저하 \d+건/);
  assert.match(t, /■ 쿠션 압력 — \d+건, [^\n]*2026-05-11~2026-05-21/);
  assert.match(t, /■ Main 압력 — \d+건, [^\n]*2026-08-03~2026-08-07/);
  assert.equal(L.alertsToAoa(al).length, al.length + 1);
  assert.match(L.alertText([], { from: 'a', to: 'b' }), /낮은 압력이 없습니다/);
});
test('설비(Press) 수선 이력 — 붙여 넣기, 수선 뒤 압력 저하·주름이 줄어듦, 설비 칸이 있으면 같은 설비만', () => {
  const r = L.parseRepairText('설비\t일자\t구분\t내용\nHydro press\t2026-05-22\t부품교체\t씰 교체\n\t2026/06/01\t점검\n2호기\t어제\t수선');
  assert.equal(r.items.length, 2); assert.deepEqual(r.bad, [4]);
  assert.equal(r.items[1].equip, '프레스', '설비 칸 비면 「프레스」'); assert.equal(r.items[0].kind, '부품교체');
  const e = AN.repairs.find(x => x.date === '2026-05-22');
  assert.ok(e.before.lowRows > 10 * Math.max(1, e.after.lowRows) / 2 && e.after.lowRows < 5, e.before.lowRows + ' → ' + e.after.lowRows);
  assert.ok(e.before.types['주름'] / e.before.qty > e.after.types['주름'] / e.after.qty * 1.3, '주름 줄어듦');
  assert.ok(e.before.cushionDev < e.after.cushionDev - 3, '쿠션 기준 대비 회복: ' + e.before.cushionDev + ' → ' + e.after.cushionDev);
  const m = AN.repairs.find(x => x.date === '2026-08-10');
  assert.ok(m.before.mainDev < m.after.mainDev - 2, 'Main 기준 대비 회복: ' + m.before.mainDev + ' → ' + m.after.mainDev);
  // 설비 칸이 있는 작업일보: 1호기 수선은 1호기 작업만
  const two = [{ date: '2026-01-05', die: 'D', qty: 100, defect: 9, press: '1호기', key: 'a' }, { date: '2026-01-06', die: 'D', qty: 100, defect: 1, press: '1호기', key: 'b' }, { date: '2026-01-05', die: 'D', qty: 100, defect: 0, press: '2호기', key: 'c' }];
  const ef = L.repairEffects(L.buildContext(two, { repairs: [{ equip: '1호기', date: '2026-01-06', kind: '수선' }, { equip: '9호기', date: '2026-01-06' }] }), 'all');
  near(ef[0].before.rate, 0.09); near(ef[0].after.rate, 0.01);
  assert.equal(ef[1].matched, false, '작업일보에 없는 설비');
  assert.equal(L.aoaToText([['설비', '일자'], ['P', new Date('2026-05-22T00:00:00')]]), '설비\t일자\nP\t2026-05-22');
});
test('프롬프트·보고서·저장 — 작업표준·전파·수선이 들어가고, 수선 내용(자유 글)은 프롬프트에 넣지 않음', () => {
  assert.match(INFO.prompt, /압력 작업표준\(작업모델별 기준 ± 허용\)보다 낮은 작업: Main \d+건, 쿠션 \d+건/);
  assert.match(INFO.prompt, /설비\(Press\) 수선 전후 30일 불량률: [^\n]*Hydro press 2026-05-22 부품교체 전 [\d.]+%\(압력 저하 \d+행\)/);
  assert.ok(!/누유|씰 교체/.test(INFO.prompt));
  const db = L.restoreDb({ records: [], stds: { by: 'die', rows: [{ key: 'D', main: '200', mainTol: 'x' }, { key: '' }] }, repairs: [{ date: '2026-01-01', kind: '교체' }, { date: 'x' }] });
  assert.deepEqual(db.stds, { by: 'die', rows: [{ key: 'D', main: 200, mainTol: null, cushion: null, cushionTol: null }] });
  assert.deepEqual(db.repairs, [{ equip: '프레스', date: '2026-01-01', kind: '부품교체', note: '' }]);
  assert.deepEqual(L.emptyDb().stds, { by: 'model', rows: [] });
});

console.log('보고서 · 내보내기 · 저장');
test('보고서 — 총괄·월별·금형·소재·조건·보수 절, AI 의견은 넣을 때만', () => {
  const t = L.buildReport(AN);
  ['1. 총괄', '2. 월별 불량률', '3. 금형·소재별', '4. 불량이 몰린 조건', '5. 금형 보수·교체', '6. 압력 작업표준 대비 · 설비(Press) 수선'].forEach(s => assert.ok(t.includes(s), s));
  assert.ok(!t.includes('7. AI'));
  assert.ok(L.buildReport(AN, '요약: 확인 필요').includes('7. AI 분석 의견'));
  assert.match(t, /쿠션 실린더 씰 교체/);
  assert.ok(t.includes('전 기록 없음(데이터 시작 전)'));
});
test('엑셀 표 — 묶음표 행 수·전체 줄, 행 목록 유형 칸', () => {
  const g = AN.dims.die;
  const aoa = L.groupToAoa(g);
  assert.equal(aoa.length, 2 + g.rows.length + 1);
  assert.equal(aoa[aoa.length - 1][3], AN.overall.qty);
  const ra = L.recordsToAoa(SRECS.slice(0, 3));
  assert.deepEqual(ra[0].slice(12, 15), ['Crack', '주름', '기타']);
  assert.equal(L.trendToAoa(AN.months, 't').length, 14);
});
test('설정 정리 — 잘못된 값은 기본값, 범위 뒤바뀌면 바로잡음', () => {
  const s = L.cleanSettings({ target: 'x', mainMin: 250, mainMax: 200, minQty: -5, cushionMin: '' });
  assert.equal(s.target, 0.02); assert.equal(s.mainMin, 200); assert.equal(s.mainMax, 250);
  assert.equal(s.minQty, L.DEFAULT_SETTINGS.minQty); assert.equal(s.cushionMin, null);
});
test('저장 형식 — 되살리기, 잘못된 행·별칭 버림, 가리기 기본 켬', () => {
  const db = L.restoreDb({ records: [{ date: '2026-01-05', die: 'D', qty: 10, defect: 1 }, { date: '', die: 'X', qty: 1 }], maint: [{ die: 'D', date: '2026-01-01' }, { die: '', date: 'x' }], aliases: { die: { D: '금형1', E: 7 } } });
  assert.equal(db.records.length, 1); assert.ok(db.records[0].key);
  assert.equal(db.maint.length, 1); assert.deepEqual(db.aliases.die, { D: '금형1' });
  assert.equal(db.mask, true); assert.equal(L.restoreDb({ mask: false }).mask, false);
});

console.log('버튼·결과 불일치 점검 (2026-10-27)');
test('「최근 N개월」 시작 = 끝 날짜가 든 달까지 N개 달의 첫날 (전에는 07-02 · 04-02)', () => {
  assert.equal(L.monthsBackStart('2026-09-30', 3), '2026-07-01');
  assert.equal(L.monthsBackStart('2026-09-30', 6), '2026-04-01');
  assert.equal(L.monthsBackStart('2026-09-30', 1), '2026-09-01');
  assert.equal(L.monthsBackStart('2026-02-15', 6), '2025-09-01', '해 넘김');
  assert.equal(L.monthsBackStart('2026-03-31', 12), '2025-04-01');
  // 샘플: 최근 3개월 = 7·8·9월 세 달이 통째로
  const ctx = L.buildContext(Sample.build().rows, { from: L.monthsBackStart('2026-09-30', 3), to: '2026-09-30' });
  assert.deepEqual(L.trend(ctx, 'month').map(m => m.key), ['2026-07', '2026-08', '2026-09']);
  assert.equal(L.trend(ctx, 'day', 'all', '2026-07')[0].key.slice(0, 7), '2026-07');
  const july1 = Sample.build().rows.filter(r => r.date === '2026-07-01').length;
  if (july1) assert.ok(ctx.recs.some(r => r.date === '2026-07-01'), '7월 1일 작업이 빠지지 않음');
});
// 1단계 판(열 읽기 판 기록 없음)으로 가져와 저장된 모양 — 두 줄 머리 결함으로 메이커 값이 규격 칸에 들어간 행
const OLD_PAYLOAD = { schemaVersion: 1, records: [
  { date: '2026-01-05', die: 'D-1', model: 'M-1', maker: '', spec: '소재사A', qty: 1000, defect: 20, types: {}, key: 'a#1' },
  { date: '2026-01-05', die: 'D-2', model: 'M-2', maker: '', spec: '소재사B', qty: 500, defect: 5, types: {}, key: 'b#1' }],
  files: [{ name: '작업일보_2026-01.xlsx', sheet: '작업일보', rows: 2, at: '2026-10-20 12:00' }] };
const TWO_ROW = [
  ['■ 작업일보'],
  ['작업일자', '작업자', '압력', null, '금형', '작업모델', '소재', null, '소재 lot', '작업수량', '불량수량'],
  [null, null, 'Main', '쿠션', null, null, '메이커', '규격', null, null, null],
  ['2026-01-05', '작업자A', 200, 50, 'D-1', 'M-1', '소재사A', 'STS304 0.8t', 'A-2601', 1000, 20],
  ['2026-01-05', '작업자B', 210, 52, 'D-2', 'M-2', '소재사B', 'STS430 1.0t', 'B-2601', 500, 5]
];
test('옛 판에서 가져와 저장된 자료 → importStatus.stale (다시 가져오기 알림), 지금 판으로 읽은 자료는 아님', () => {
  const old = L.restoreDb(JSON.parse(JSON.stringify(OLD_PAYLOAD)));
  const st = L.importStatus(old);
  assert.equal(st.stale, true); assert.equal(st.oldRows, 2); assert.deepEqual(st.oldFiles, ['작업일보_2026-01.xlsx']);
  const p = L.parseSheet(TWO_ROW);
  assert.ok(p.records.every(r => r.pv === L.PARSER_VERSION), '읽은 행마다 열 읽기 판');
  const fresh = L.restoreDb(JSON.parse(JSON.stringify({ records: p.records, files: [{ name: 'x.xlsx', parser: L.PARSER_VERSION, cols: L.fileColumns(p.det) }] })));
  assert.equal(L.importStatus(fresh).stale, false, '저장 → 되살리기 뒤에도 판 기록 유지');
  // 지우지 않고 같은 파일을 다시 넣으면 메이커·규격이 바뀐 옛 행은 키가 달라 남음 → 여전히 알림
  const mixed = L.restoreDb(JSON.parse(JSON.stringify(OLD_PAYLOAD)));
  mixed.records = L.mergeRecords(mixed.records, p.records).records;
  mixed.files = [{ name: '작업일보_2026-01.xlsx', parser: L.PARSER_VERSION, cols: L.fileColumns(p.det) }];
  assert.equal(L.importStatus(mixed).stale, true); assert.equal(L.importStatus(mixed).oldRows, 2);
  assert.equal(L.importStatus(L.emptyDb()).stale, false);
});
test('탭의 칸 출처 — 두 줄 머리 「소재」 아래 메이커·규격 + 「소재 lot」: 메이커 ← G열, 규격 ← H열, 로트 ← I열', () => {
  const p = L.parseSheet(TWO_ROW);
  assert.deepEqual(p.records.map(r => [r.maker, r.spec, r.lot]), [['소재사A', 'STS304 0.8t', 'A-2601'], ['소재사B', 'STS430 1.0t', 'B-2601']]);
  const files = [{ name: 'a', cols: L.fileColumns(p.det) }, { name: 'b', cols: L.fileColumns(p.det) }, { name: 'old' }];
  const mk = L.columnSources(files, 'maker'), sp = L.columnSources(files, 'spec'), lt = L.columnSources(files, 'lot');
  assert.deepEqual(mk.list, [{ text: 'G열 「소재 메이커」', files: 2 }]); assert.equal(mk.unknown, 1);
  assert.deepEqual(sp.list, [{ text: 'H열 「소재 규격」', files: 2 }]);
  assert.deepEqual(lt.list, [{ text: 'I열 「소재 lot」', files: 2 }]);
  assert.equal(L.columnSources(files, 'worker').list[0].text, 'B열 「작업자」');
});
console.log('실제 작업일보 머리 이름 점검 (2026-10-27) — 소재 메이커 · 소재규격');
const MK = ['소재사A', '소재사B', '소재사C', '소재사A', '소재사B', '소재사C'];
const SP = ['STS304 0.8t', 'SPCC 1.2T', 'STS430 1.0t', 'AL5052 1.5mm', 'SPHC 2.0t', 'STS304 0.8t'];
// 작업일자 · 작업자 · 금형 · (소재 칸들) · 작업수량 · 불량수량
function sheetWith(heads, cols) {
  const aoa = [['작업일자', '작업자', '금형', ...heads, '작업수량', '불량수량']];
  for (let i = 0; i < 6; i++) aoa.push(['2026-01-0' + (i + 1), '작업자A', 'D-' + i, ...cols.map(c => c[i]), 1000, 10 + i]);
  return aoa;
}
const uniq = a => [...new Set(a)];
function rightPlace(aoa, msg) {
  const p = L.parseSheet(aoa);
  assert.deepEqual(uniq(p.records.map(r => r.maker)), uniq(MK), msg + ' — 소재 메이커 값');
  assert.deepEqual(uniq(p.records.map(r => r.spec)), uniq(SP), msg + ' — 소재규격 값');
  assert.ok(!p.det.warnings.some(w => /한 번 더/.test(w)), msg + ' — 소재 칸을 「한 번 더」로 버리지 않음: ' + p.det.warnings.join(' / '));
  assert.ok(!(p.det.material && p.det.material.warn), msg + ' — 어긋남 알림 없음');
  return p;
}
test('머리 이름 — 「소재 | 규격」「소재 | 소재규격」: 「소재」 = 메이커, 「규격」 = 규격 (전에는 「소재」 가 규격을 차지하고 「규격」 은 버려짐)', () => {
  rightPlace(sheetWith(['소재', '규격'], [MK, SP]), '소재 | 규격');
  rightPlace(sheetWith(['소재', '소재규격'], [MK, SP]), '소재 | 소재규격');
  rightPlace(sheetWith(['소재', '소재 lot', '규격'], [MK, MK.map((m, i) => 'L-' + i), SP]), '소재 | 소재 lot | 규격 (작업일보 순서)');
});
test('머리 이름 — 순서를 바꿔도 「규격 | 소재」, 「소재 사양 | 소재명」', () => {
  rightPlace(sheetWith(['규격', '소재'], [SP, MK]), '규격 | 소재');
  rightPlace(sheetWith(['소재명', '소재 사양'], [MK, SP]), '소재명 | 소재 사양');
  rightPlace(sheetWith(['소재 사양', '소재명'], [SP, MK]), '소재 사양 | 소재명');
});
test('머리 이름 — 「원소재 | 제조사」「자재 | 공급업체」: 회사 쪽이 메이커, 남은 소재 칸이 규격', () => {
  rightPlace(sheetWith(['원소재', '제조사'], [SP, MK]), '원소재 | 제조사');
  rightPlace(sheetWith(['자재', '공급업체'], [SP, MK]), '자재 | 공급업체');
  rightPlace(sheetWith(['공급사', '재질'], [MK, SP]), '공급사 | 재질');
});
test('머리 이름 — 영어: Material | Spec, Supplier | Material, Maker | Grade', () => {
  rightPlace(sheetWith(['Material', 'Spec'], [MK, SP]), 'Material | Spec');
  rightPlace(sheetWith(['Supplier', 'Material'], [MK, SP]), 'Supplier | Material');
  rightPlace(sheetWith(['Maker', 'Grade'], [MK, SP]), 'Maker | Grade');
});
test('머리 이름 — 두 줄 병합 「소재」 아래 「메이커」「규격」(lot 없음) · 「소재 | 소재 메이커 | 소재규격」 은 구체 이름이 이김', () => {
  const aoa = [['작업일자', '작업자', '금형', '소재', null, '작업수량', '불량수량'], [null, null, null, '메이커', '규격', null, null]];
  for (let i = 0; i < 6; i++) aoa.push(['2026-01-0' + (i + 1), '작업자A', 'D-' + i, MK[i], SP[i], 1000, 10 + i]);
  rightPlace(aoa, '두 줄 병합');
  const p = L.parseSheet(sheetWith(['소재', '소재 메이커', '소재규격'], [MK.map(() => '냉연'), MK, SP]));
  assert.deepEqual([p.det.map.maker, p.det.map.spec], [4, 5], '구체 이름(소재 메이커 · 소재규격)이 연결됨');
  assert.deepEqual(uniq(p.records.map(r => r.maker)), uniq(MK));
});
test('값 보기 — 머리가 둘 다 애매하면(「재질 | 소재」) 아래 값으로: 회사 이름 = 메이커, STS304 0.8t = 규격', () => {
  rightPlace(sheetWith(['재질', '소재'], [MK, SP]), '재질(회사 이름) | 소재(규격 값)');
  rightPlace(sheetWith(['소재', '재질'], [SP, MK]), '소재(규격 값) | 재질(회사 이름)');
  // 소재 칸 하나뿐: 값이 회사 이름이면 메이커, 규격 값이면 규격
  let p = L.parseSheet(sheetWith(['소재'], [MK]));
  assert.equal(p.det.map.maker, 3); assert.equal(p.det.map.spec, undefined);
  p = L.parseSheet(sheetWith(['소재'], [SP]));
  assert.equal(p.det.map.spec, 3); assert.equal(p.det.map.maker, undefined);
  assert.equal(L.valueKind(['STS304 0.8t', 'SPCC 1.2T', 'AL5052 1.5mm', '0.8t', 'SUS430 1.0×1219']), 'spec');
  assert.equal(L.valueKind(['소재사A', '소재사B', '㈜가나금속', 'ABC Steel', '소재사A']), 'maker');
  assert.equal(L.valueKind(['A-2601', 'B-2602', 'C-2603']), null, '로트 번호는 어느 쪽도 아님');
});
test('값 보기 — 머리는 「메이커 | 규격」 인데 값이 뒤바뀌어 있으면 연결은 머리대로 두고 「두 칸 바꾸기」 알림, 바꾸면 알림 없음', () => {
  const aoa = sheetWith(['메이커', '규격'], [SP, MK]);
  const p = L.parseSheet(aoa);
  assert.deepEqual([p.det.map.maker, p.det.map.spec], [3, 4], '머리 이름대로 연결');
  assert.ok(p.det.material.swap, '바꾸기 제안'); assert.ok(/소재 메이커 칸에 규격처럼 보이는 값/.test(p.det.material.warn));
  assert.ok(p.det.warnings.some(w => /열 연결을 확인해 주세요/.test(w)));
  assert.deepEqual(p.det.material.maker.samples.slice(0, 3), SP.slice(0, 3), '예시 값 3개');
  const q = L.parseSheet(aoa, { overrides: { maker: 4, spec: 3 } });
  assert.deepEqual(uniq(q.records.map(r => r.maker)), uniq(MK)); assert.ok(!q.det.material.warn, '바꾼 뒤 알림 없음');
  // 메이커 칸이 없고 규격 칸에 회사 이름만 있으면 — 바꿀 칸은 없지만 알림
  const r = L.parseSheet(sheetWith(['규격'], [MK]));
  assert.ok(/소재규격 칸에 메이커/.test(r.det.material.warn || '')); assert.ok(!r.det.material.swap);
});
test('쌓인 자료 점검 — 메이커·규격 값이 뒤바뀐 행을 알아보고 「두 칸 바꾸기」로 고침(열쇠도 새로)', () => {
  const p = L.parseSheet(sheetWith(['메이커', '규격'], [SP, MK]));
  const mc = L.materialCheck(p.records);
  assert.ok(mc.swap && /소재 메이커 칸에 규격처럼/.test(mc.warn));
  const fixed = L.swapMaterial(p.records);
  assert.deepEqual(uniq(fixed.map(r => r.maker)), uniq(MK)); assert.ok(!L.materialCheck(fixed).warn);
  assert.ok(fixed.every(r => r.key.startsWith([r.date, r.die, r.model, r.worker, r.maker, r.spec].join('|'))), '열쇠도 바뀐 값으로');
  assert.equal(new Set(fixed.map(r => r.key)).size, fixed.length);
  assert.ok(!L.materialCheck(L.parseSheet(TINY).records).warn, '바른 자료는 알림 없음');
  // 옛 결함 모양(메이커 비고 규격에 회사 이름) — 바꿀 수 없으니 다시 가져오기 안내
  const old = L.materialCheck(OLD_PAYLOAD.records.concat(OLD_PAYLOAD.records, OLD_PAYLOAD.records));
  assert.ok(old.warn && !old.swap);
});
test('열 읽기 판 3 — 오늘 오전 판(2)으로 가져온 행도 다시 가져오기 알림, 고친 열 연결은 머리행마다 기억', () => {
  assert.ok(L.PARSER_VERSION >= 3);
  const db = L.restoreDb({ records: [{ date: '2026-01-05', die: 'D', qty: 10, defect: 1, maker: '', spec: '소재사A', pv: 2, key: 'k#1' }], files: [{ name: 'a.xlsx', parser: 2 }], colOverrides: { 'x|y': { maker: 4, spec: 3 } } });
  assert.equal(L.importStatus(db).stale, true);
  assert.deepEqual(db.colOverrides, { 'x|y': { maker: 4, spec: 3 } }, '되살리기 뒤에도 열 연결 기억');
  assert.deepEqual(L.emptyDb().colOverrides, {});
});
const { versions } = await import('../scripts/bump-version.mjs');
test('판 맞추기 — index.html ?v= · logic.js · app.js · sample-data.js 가 같은 판 (scripts/bump-version.mjs)', () => {
  const v = versions();
  assert.equal(v.logic, L.VERSION); assert.equal(v.sample, Sample.VERSION);
  assert.equal(v.app, L.VERSION, 'app.js APP_VERSION'); assert.ok(v.html.length >= 9 && v.html.every(x => x === L.VERSION), 'index.html 의 모든 js·css 에 ?v=' + L.VERSION);
});

console.log('AI 연결 설정 (선택 기능)');
test('주소 정리·점검·요청 — 키는 넣을 때만 머리글, https 페이지에서 http 주소 막음', () => {
  assert.equal(E.normalizeBaseUrl('https://x.example/v1/chat/completions/'), 'https://x.example/v1');
  assert.ok(E.validateConfig({ baseUrl: 'http://10.0.0.5/v1', model: 'm' }, 'https:').errors.length);
  assert.equal(E.validateConfig({ baseUrl: 'http://10.0.0.5/v1', model: 'm' }, 'file:').errors.length, 0);
  const req = E.buildChatRequest({ baseUrl: 'https://x/v1', model: 'm' }, E.promptMessages('hi'));
  assert.equal(req.init.headers.Authorization, undefined);
  assert.equal(E.buildChatRequest({ baseUrl: 'https://x/v1', model: 'm', apiKey: 'k' }, []).init.headers.Authorization, 'Bearer k');
  assert.equal(E.toStored({ baseUrl: 'https://x/v1', model: 'm', apiKey: 'k', remember: false }).apiKey, '', '기억 끄면 키 저장 안 함');
  assert.equal(E.parseChatResponse({ choices: [{ message: { content: '답' } }] }), '답');
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
