// 헤드리스 브라우저로 「버튼 · 탭 · 고르기 상자와 그 아래 결과가 같은가」를 잽니다 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어
//  ① 샘플 → ② 항목별: 빠른 기간 3개 × 보는 불량 전부 × 탭 전부 × 정렬 3개마다
//     눌린 탭 = 그래프 제목 = 막대 이름·값 = 표 줄 = 원본 행을 따로 더한 값(logic.js 를 쓰지 않고 셈) · 탭의 칸 출처
//  ② 빠른 기간: 최근 3·6개월 = 7월 1일 · 4월 1일부터, 눌린 단추 = 지금 기간
//  ③ 기간 → 탭 → 보는 불량을 바꾼 뒤 다른 화면에 갔다 와도 그대로 · 다른 화면의 제목·보고서 머리도 같은 상태 · 일별 달 고르기
//  ④ 새로고침하면 기본값(금형 · 전체 불량 · 전체 기간)으로, 단추와 결과가 함께
//  ⑤ 옛 판에서 가져온 자료(열 읽기 판 기록 없음)를 넣고 열면 「다시 가져오기」 알림 → 작업일보만 지우기 →
//     두 줄 병합 머리(윗줄 「소재」 아래 「메이커」「규격」 + 「소재 lot」) xlsx 를 파일로 넣기 → 메이커·로트·규격 탭 = 원본 칸 값
//  ⑥ 작업자·lot 칸이 없는 파일로 바꾸면(작업자 탭을 보던 중) 눌린 탭과 내용이 같은 기준
//  ⑦ 옛 logic.js 가 섞여 들어오면(판 다름) 계산하지 않고 새로고침 안내
//  ⑧ 실제 작업일보 머리 이름: 「소재 | 소재 lot | 규격」 xlsx → 소재 칸 확인 카드 · ② 메이커/규격 탭 = 원본 칸, 값이 뒤바뀐 「메이커 | 규격」 → 알림 · 두 칸 바꾸기 · 기억, ② 쌓인 행 바꾸기, 판 2 자료 알림
// BREAK=1 이면 app.js 가 탭과 다른 기준으로 표를 그리게 바꿔(브라우저 안에서만) 검사가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const url = 'file://' + path.join(root, 'index.html');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'data10-11-check-'));

const errors = [], problems = [];
let checks = 0;
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
// fresh-guard 는 file:// 에서 HEAD 를 못 보내 콘솔에 한 줄 남깁니다(잡아서 무시함 — 배포 주소에서는 안 남)
page.on('console', m => { if (m.type() === 'error' && !/Fetch API cannot load file:.*_fg=/.test(m.text())) errors.push('console: ' + m.text()); });
page.on('dialog', d => d.accept());
function ok(cond, msg, quiet) { checks++; if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else if (!quiet) console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(120); }

const LAB = { die: '금형', maker: '소재 메이커', lot: '소재 로트', spec: '소재규격', model: '작업모델', worker: '작업자' };
const FIELD = Object.fromEntries(Object.entries(LAB).map(([k, v]) => [v, k]));
async function snap() {
  return page.evaluate(() => {
    const q = s => document.querySelector(s);
    const db = JSON.parse(localStorage.getItem('data10-11.db'));
    const tabsCard = q('.tabs') && q('.tabs').closest('section');
    const svg = tabsCard && tabsCard.querySelector('svg.hbar');
    const tbl = tabsCard && tabsCard.querySelector('table.list');
    return {
      records: db.records,
      dates: [...document.querySelectorAll('section.filters input[type=date]')].map(i => i.value),
      metric: q('select[aria-label="보는 불량"]') && q('select[aria-label="보는 불량"]').value,
      pressed: [...document.querySelectorAll('.tabs button[aria-pressed=true]')].map(b => b.textContent),
      tabs: [...document.querySelectorAll('.tabs button')].map(b => b.textContent),
      sort: q('select[aria-label="정렬"]') && q('select[aria-label="정렬"]').value,
      cap: tabsCard && tabsCard.querySelector('figcaption') && tabsCard.querySelector('figcaption').textContent,
      src: q('.col-source') && q('.col-source').textContent,
      bars: svg ? [...svg.querySelectorAll('title')].map(t => t.textContent) : [],
      rows: tbl ? [...tbl.querySelectorAll('tbody tr')].map(tr => [...tr.cells].map(td => td.textContent)) : []
    };
  });
}
// logic.js 를 쓰지 않고 원본 행을 그대로 더한 값
function recompute(recs, from, to, dim, metric) {
  const m = {};
  let Q = 0, C = 0;
  recs.filter(r => r.date >= from && r.date <= to).forEach(r => {
    const c = metric === 'all' ? r.defect : (r.types && r.types[metric]) || 0;
    Q += r.qty; C += c;
    const k = r[dim]; if (!k) return;
    (m[k] ||= { q: 0, c: 0 }); m[k].q += r.qty; m[k].c += c;
  });
  return { m, rate: Q ? C / Q : null };
}
const pctNum = s => parseFloat(String(s).replace(/[^0-9.\-]/g, ''));
const r2 = v => Math.round(v * 10000) / 100;
async function checkTable(tag, quiet = true) {
  const s = await snap();
  ok(s.pressed.length === 1, tag + ' — 눌린 탭 하나 (' + s.pressed.join(',') + ')', quiet);
  const dim = FIELD[s.pressed[0]];
  const ml = s.metric === 'all' ? '전체 불량' : s.metric + ' 불량';
  ok(s.cap === s.pressed[0] + '별 ' + ml + '률', tag + ' — 그래프 제목 = 눌린 탭 · 보는 불량 (' + s.cap + ')', quiet);
  ok(s.src && s.src.startsWith(s.pressed[0] + ' ← '), tag + ' — 칸 출처 줄 = 눌린 탭 (' + s.src + ')', quiet);
  const rc = recompute(s.records, s.dates[0], s.dates[1], dim, s.metric);
  const body = s.rows.filter(r => r[0] !== '전체');
  ok(body.length === Object.keys(rc.m).length, tag + ' — 표 줄 수 = 원본 값 가짓수 (' + body.length + '/' + Object.keys(rc.m).length + ')', quiet);
  const bad = body.filter(r => { const x = rc.m[r[0]]; return !x || r[3] !== x.q.toLocaleString('en-US') || pctNum(r[5]) !== r2(x.c / x.q); });
  ok(!bad.length, tag + ' — 표 값 = 원본을 따로 더한 값' + (bad.length ? ' (' + bad[0].join('|') + ')' : ''), quiet);
  ok(s.bars.length === body.length && s.bars.every((t, i) => t.startsWith(body[i][0] + ' · ') && pctNum(t.split(' ').pop()) === pctNum(body[i][5])), tag + ' — 막대 = 표 순서·값', quiet);
  const tot = s.rows.find(r => r[0] === '전체');
  ok(tot && pctNum(tot[5]) === r2(rc.rate), tag + ' — 전체 줄 = 원본', quiet);
  // 정렬
  const key = { excess: i => -pctNum(body[i][7]), rate: i => -pctNum(body[i][5]), order: null }[s.sort];
  if (key) ok(body.every((r, i) => i === 0 || key(i - 1) <= key(i) + 1e-9), tag + ' — 정렬 「' + s.sort + '」 대로', quiet);
  else ok(body.every((r, i) => i === 0 || body[i - 1][0] <= r[0]), tag + ' — 이름 순', quiet);
  return s;
}
async function quickDates(label) {
  await page.locator('.quick-period button', { hasText: label }).click(); await page.waitForTimeout(80);
  return page.evaluate(() => ({ d: [...document.querySelectorAll('section.filters input[type=date]')].map(i => i.value), on: [...document.querySelectorAll('.quick-period button[aria-pressed=true]')].map(b => b.textContent) }));
}

await page.goto(url);
await page.evaluate(() => localStorage.clear());
await page.reload();
if (process.env.BREAK) {
  // 탭 단추를 그리는 순서와 표를 그리는 기준을 한 칸 어긋나게: 누른 탭의 다음 탭 기준으로 표를 그리도록 groupBy 를 감쌈
  await page.evaluate(() => { const L = window.DDLogic, g = L.groupBy, ds = ['die', 'maker', 'lot', 'spec', 'model', 'worker']; L.groupBy = function (c, d, m, s) { const i = ds.indexOf(d); return g(c, i >= 0 && i < 2 ? ds[i + 1] : d, m, s); }; });
}
await page.locator('button', { hasText: '샘플 12개월로 시작' }).first().click();
await page.waitForTimeout(800);

console.log('① ② 항목별 — 빠른 기간 × 보는 불량 × 탭 × 정렬');
await go('#/table');
const metrics = await page.$$eval('select[aria-label="보는 불량"] option', os => os.map(o => o.value));
let combos = 0;
for (const qb of ['전체 기간', '최근 6개월', '최근 3개월']) {
  await quickDates(qb);
  for (const mt of metrics) {
    await page.selectOption('select[aria-label="보는 불량"]', mt); await page.waitForTimeout(60);
    for (const t of await page.$$eval('.tabs button', a => a.map(b => b.textContent))) {
      await page.locator('.tabs button', { hasText: t }).first().click(); await page.waitForTimeout(40);
      for (const so of ['excess', 'rate', 'order']) {
        await page.selectOption('select[aria-label="정렬"]', so); await page.waitForTimeout(30);
        await checkTable([qb, mt, t, so].join(' / ')); combos++;
      }
    }
  }
}
ok(problems.length === 0, combos + '가지 상태 모두 탭 · 그래프 · 표 · 원본 일치');

console.log('② 빠른 기간');
let q3 = await quickDates('최근 3개월'), q6 = await quickDates('최근 6개월'), qa = await quickDates('전체 기간');
ok(q3.d.join('~') === '2026-07-01~2026-09-30' && q3.on.join() === '최근 3개월', '최근 3개월 = 2026-07-01 ~ 09-30, 그 단추만 눌림 (' + q3.d.join('~') + ' · ' + q3.on + ')');
ok(q6.d.join('~') === '2026-04-01~2026-09-30' && q6.on.join() === '최근 6개월', '최근 6개월 = 2026-04-01 ~ 09-30 (' + q6.d.join('~') + ')');
ok(qa.d.join('~') === '2025-10-01~2026-09-30' && qa.on.join() === '전체 기간', '전체 기간 (' + qa.d.join('~') + ')');
await page.fill('section.filters input[type=date] >> nth=0', '2026-05-10'); await page.dispatchEvent('section.filters input[type=date] >> nth=0', 'change'); await page.waitForTimeout(100);
ok((await page.$$eval('.quick-period button[aria-pressed=true]', a => a.length)) === 0, '날짜를 직접 고치면 빠른 기간 단추는 눌리지 않음');
await checkTable('직접 고친 기간 2026-05-10 ~', false);

console.log('③ 다른 화면에 갔다 오기');
await quickDates('최근 3개월');
await page.locator('.tabs button', { hasText: '소재 로트' }).click(); await page.waitForTimeout(60);
await page.selectOption('select[aria-label="보는 불량"]', 'Crack'); await page.waitForTimeout(60);
await page.selectOption('select[aria-label="정렬"]', 'rate'); await page.waitForTimeout(60);
for (const p of ['#/trend', '#/cond', '#/ai', '#/report', '#/press', '#/settings', '#/import']) {
  await go(p);
  if (p !== '#/settings' && p !== '#/import') {
    const r = await page.evaluate(() => ({ d: [...document.querySelectorAll('section.filters input[type=date]')].map(i => i.value), m: document.querySelector('select[aria-label="보는 불량"]').value, on: [...document.querySelectorAll('.quick-period button[aria-pressed=true]')].map(b => b.textContent), t: document.querySelector('main').textContent }));
    ok(r.d.join('~') === '2026-07-01~2026-09-30' && r.m === 'Crack' && r.on.join() === '최근 3개월', p + ' — 기간 · 보는 불량 · 눌린 단추가 같음', true);
    ok(r.t.includes('Crack 불량'), p + ' — 내용이 Crack 불량 기준', true);
    if (p === '#/report') ok(r.t.includes('2026-07-01 ~ 2026-09-30 · Crack 불량'), '보고서 머리 = 기간 · 보는 불량', true);
    if (p === '#/ai') ok(r.t.includes('AI 답 (2026-07-01 ~ 2026-09-30, Crack 불량)'), 'AI 답 칸 = 기간 · 보는 불량', true);
  }
  await go('#/table');
  const s = await checkTable('복귀 ' + p);
  ok(s.pressed[0] === '소재 로트' && s.metric === 'Crack' && s.sort === 'rate', '복귀 ' + p + ' — 소재 로트 · Crack · 불량률 순 그대로', true);
}
ok(true, '7개 화면에 갔다 와도 ② 의 탭 · 보는 불량 · 정렬 · 기간과 결과가 같음');
await go('#/trend');
for (const m of await page.$$eval('select[aria-label="일별로 볼 달"] option', a => a.map(o => o.value))) {
  await page.selectOption('select[aria-label="일별로 볼 달"]', m); await page.waitForTimeout(50);
  const r = await page.evaluate(() => ({ v: document.querySelector('select[aria-label="일별로 볼 달"]').value, h: [...document.querySelectorAll('h2')].map(x => x.textContent).find(x => x.startsWith('일별')), days: [...document.querySelectorAll('details table tbody tr')].map(tr => tr.cells[0].textContent) }));
  ok(r.v === m && r.days.length && r.days.every(d => d.startsWith(m)) && r.h.endsWith(+m.slice(5) + '월'), '① 일별 달 ' + m + ' — 고른 달 = 제목 = 표 날짜', true);
}
ok(true, '① 일별 달 고르기 3개월치 모두 일치');

console.log('④ 새로고침');
await go('#/table'); await page.reload(); await page.waitForTimeout(400);
let s = await checkTable('새로고침 뒤', false);
ok(s.pressed[0] === '금형' && s.metric === 'all' && s.dates.join('~') === '2025-10-01~2026-09-30', '새로고침하면 금형 · 전체 불량 · 전체 기간 (단추와 결과 함께)');

console.log('⑤ 옛 판에서 가져온 자료 → 다시 가져오기 → 두 줄 병합 머리 xlsx');
await page.evaluate(() => {
  const db = JSON.parse(localStorage.getItem('data10-11.db'));
  db.sample = false;
  db.records = db.records.slice(0, 300).map(r => { const o = Object.assign({}, r, { maker: '', spec: r.maker, lot: '' }); delete o.pv; return o; });   // 1단계 결함 모양: 메이커 값이 규격 칸
  db.files = [{ name: '작업일보_2026-01.xlsx', sheet: '작업일보', rows: 300, at: '2026-10-20 12:00' }];
  localStorage.setItem('data10-11.db', JSON.stringify(db));
});
await page.reload(); await page.waitForTimeout(400);
ok(await page.locator('#staleImport').isVisible(), '옛 판 자료 → 「이전 판에서 가져온 자료입니다」 알림');
await go('#/table');
s = await snap();
ok(s.tabs.indexOf('소재 메이커') < 0 && /칸 기록 없음|칸 기록이 없는/.test(s.src || ''), '② 에 옛 자료의 칸 출처가 없음을 보여 줌 (' + s.src + ')');
if (await page.locator('#staleImport a').count()) await page.locator('#staleImport a', { hasText: '다시 가져오기 안내' }).click(); else await go('#/import');
await page.waitForTimeout(200);
ok(await page.locator('#reimportGuide').isVisible(), '「다시 가져오기 안내」 → 가져오기 화면의 안내 카드');
if (await page.locator('#reimportGuide').count()) await page.locator('#reimportGuide button', { hasText: '작업일보만 지우기' }).click();
else await page.evaluate(() => { const db = JSON.parse(localStorage.getItem('data10-11.db')); db.records = []; db.files = []; localStorage.setItem('data10-11.db', JSON.stringify(db)); });
await page.waitForTimeout(200);
const head1 = ['작업일자', '작업자', '압력', null, '금형', '작업모델', '소재', null, '소재 lot', '작업수량', '불량수량', 'Crack'];
const head2 = [null, null, 'Main', '쿠션', null, null, '메이커', '규격', null, null, null, null];
const raw = [];
for (let d = 1; d <= 20; d++) for (let k = 0; k < 6; k++) {
  const mk = ['소재사A', '소재사B', '소재사C'][k % 3], sp = ['STS304 0.8t', 'STS430 1.2t'][k % 2], lot = mk.slice(-1) + '-26' + String(d % 4);
  raw.push(['2026-01-' + String(d).padStart(2, '0'), '작업자' + 'ABC'[k % 3], 200 + k, 50 + k, 'D-' + (k % 4), 'M-' + (k % 2), mk, sp, lot, 1000 + 37 * k + d, (k * 7 + d) % 40, (k * 3 + d) % 10]);
}
const ws = XLSX.utils.aoa_to_sheet([['■ 작업일보 (2026년 1월)'], head1, head2, ...raw]);
ws['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } }, { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } }, { s: { r: 1, c: 4 }, e: { r: 2, c: 4 } }, { s: { r: 1, c: 5 }, e: { r: 2, c: 5 } },
  { s: { r: 1, c: 6 }, e: { r: 1, c: 7 } }, { s: { r: 1, c: 8 }, e: { r: 2, c: 8 } }, { s: { r: 1, c: 9 }, e: { r: 2, c: 9 } }, { s: { r: 1, c: 10 }, e: { r: 2, c: 10 } }, { s: { r: 1, c: 11 }, e: { r: 2, c: 11 } }];
const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '작업일보');
const twoRow = path.join(tmp, '작업일보_두줄머리_2026-01.xlsx'); fs.writeFileSync(twoRow, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
await page.setInputFiles('input[type=file][multiple]:not([webkitdirectory])', twoRow); await page.waitForTimeout(800);
await page.locator('button', { hasText: /확인한 1개 파일 반영/ }).click(); await page.waitForTimeout(300);
ok(!(await page.locator('#staleImport').count()), '다시 가져오면 옛 판 알림이 사라짐');
await go('#/table');
for (const [t, ci] of [['소재 메이커', 6], ['소재 로트', 8], ['소재규격', 7]]) {
  await page.locator('.tabs button', { hasText: t }).click(); await page.waitForTimeout(60);
  s = await checkTable('두 줄 머리 ' + t, false);
  const want = {}; raw.forEach(r => { (want[r[ci]] ||= { q: 0 }).q += r[9]; });
  const got = Object.fromEntries(s.rows.filter(r => r[0] !== '전체').map(r => [r[0], +r[3].replace(/,/g, '')]));
  ok(JSON.stringify(Object.keys(got).sort()) === JSON.stringify(Object.keys(want).sort()) && Object.keys(want).every(k => got[k] === want[k].q), t + ' 탭 = 엑셀 ' + XLSX.utils.encode_col(ci) + '열 값 그대로 (' + Object.keys(got).slice(0, 3).join(', ') + ' …)');
  ok(s.src.includes(XLSX.utils.encode_col(ci) + '열'), t + ' 칸 출처 = ' + XLSX.utils.encode_col(ci) + '열 (' + s.src + ')');
}

await go('#/import');
await page.setInputFiles('input[type=file][multiple]:not([webkitdirectory])', twoRow); await page.waitForTimeout(800);
await page.locator('button', { hasText: /확인한 1개 파일 반영/ }).click(); await page.waitForTimeout(300);
ok((await page.textContent('main')).includes('가져온 파일 1개 보기'), '같은 파일을 한 번 더 넣어도 가져온 파일 목록은 1개(옛 기록을 새 판 기록으로 바꿈)');
await go('#/table');

console.log('⑥ 작업자 · lot 칸이 없는 파일 (작업자 탭을 보던 중)');
await page.locator('.tabs button', { hasText: '작업자' }).click(); await page.waitForTimeout(60);
await go('#/import');
await page.locator('button', { hasText: '모든 데이터 지우기' }).click(); await page.waitForTimeout(200);
const noW = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(noW, XLSX.utils.aoa_to_sheet([['작업일자', 'Main 압력', '쿠션 압력', '금형', '작업모델', '소재 메이커', '소재규격', '작업수량', '불량수량'], ...raw.map(r => [r[0], r[2], r[3], r[4], r[5], r[6], r[7], r[9], r[10]])]), '작업일보');
const noWf = path.join(tmp, '작업일보_작업자없음.xlsx'); fs.writeFileSync(noWf, XLSX.write(noW, { type: 'buffer', bookType: 'xlsx' }));
await page.setInputFiles('input[type=file][multiple]:not([webkitdirectory])', noWf); await page.waitForTimeout(800);
await page.locator('button', { hasText: /확인한 1개 파일 반영/ }).click(); await page.waitForTimeout(300);
await go('#/table');
s = await checkTable('작업자 · lot 없는 파일', false);
ok(s.tabs.join() === '금형,소재 메이커,소재규격,작업모델' && s.pressed[0] === '금형', '작업자 · 소재 로트 탭이 빠지고 눌린 탭 = 금형 = 내용 (' + s.tabs.join() + ' / ' + s.pressed + ')');

console.log('⑦ 옛 logic.js 가 섞여 들어오면');
const p2 = await ctx.newPage();
await p2.addInitScript(() => { let v; Object.defineProperty(window, 'DDLogic', { configurable: true, get: () => v, set: x => { x.VERSION = '2026-10-20.1'; v = x; } }); });
await p2.goto(url); await p2.waitForTimeout(300);
ok(await p2.locator('#versionMismatch').isVisible() && (await p2.textContent('#versionMismatch')).includes('Ctrl+F5'), '판이 다르면 「새 판이 나왔습니다 — 새로고침(Ctrl+F5)」 안내, 계산 화면을 그리지 않음');
ok(!(await p2.locator('.tabs, table.list').count()), '판이 다를 때 표 · 탭 없음');
await p2.close();
await page.reload(); await page.waitForTimeout(200);
ok((await page.textContent('.site-footer')).includes('판 2026-'), '바닥글에 판 표시');


console.log('⑧ 실제 작업일보 머리 이름 — 「소재 | 소재 lot | 규격」 · 값이 뒤바뀐 「메이커 | 규격」 · 쌓인 행 바꾸기 · 판 2 자료');
async function importOne(file, before) {
  await go('#/import');
  await page.locator('button', { hasText: '모든 데이터 지우기' }).click(); await page.waitForTimeout(200);
  await page.setInputFiles('input[type=file][multiple]:not([webkitdirectory])', file); await page.waitForTimeout(800);
  if (before) await before();
  await page.locator('button', { hasText: /확인한 1개 파일 반영/ }).click(); await page.waitForTimeout(300);
}
async function tabValues(t) {
  await go('#/table');
  await page.locator('.tabs button', { hasText: new RegExp('^' + t + '$') }).click(); await page.waitForTimeout(60);
  const s = await checkTable('머리 이름 ' + t);
  return Object.fromEntries(s.rows.filter(r => r[0] !== '전체').map(r => [r[0], +r[3].replace(/,/g, '')]));
}
const want = (rows, ci) => { const m = {}; rows.forEach(r => { m[r[ci]] = (m[r[ci]] || 0) + r[9]; }); return m; };
const same = (a, b) => JSON.stringify(Object.keys(a).sort().map(k => [k, a[k]])) === JSON.stringify(Object.keys(b).sort().map(k => [k, b[k]]));
const HEAD_S = ['작업일자', '작업자', 'Main 압력', '쿠션 압력', '금형', '작업모델', '소재', '소재 lot', '규격', '작업수량', '불량수량'];
const rawS = raw.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[8], r[7], r[9], r[10]]);   // G=메이커 값 · H=lot · I=규격 값
const wbS = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wbS, XLSX.utils.aoa_to_sheet([HEAD_S, ...rawS]), '작업일보');
const fS = path.join(tmp, '작업일보_소재_규격.xlsx'); fs.writeFileSync(fS, XLSX.write(wbS, { type: 'buffer', bookType: 'xlsx' }));
await importOne(fS, async () => {
  const card = await page.evaluate(() => { const c = document.querySelector('.mat-check'); return c && { t: c.textContent, bad: c.classList.contains('bad'), rows: [...c.querySelectorAll('tbody tr')].map(tr => [...tr.cells].map(td => td.textContent)) }; });
  ok(card && card.rows[0][1].startsWith('G열') && card.rows[2][1].startsWith('I열') && !card.bad, '가져오기 「소재 칸 확인」 — 소재 메이커 ← G열 「소재」, 소재규격 ← I열 「규격」, 알림 없음 (' + (card && card.rows.map(r => r[1]).join(' / ')) + ')');
  ok(card && card.rows[0][2].split(' · ').length === 3 && card.rows[0][2].includes('소재사A') && card.rows[2][2].includes('STS'), '예시 값 3개씩 (' + (card && card.rows[0][2] + ' | ' + card.rows[2][2]) + ')');
});
ok(same(await tabValues('소재 메이커'), want(rawS, 6)), '② 소재 메이커 탭 = 엑셀 G열 「소재」(회사 이름) 값');
ok(same(await tabValues('소재규격'), want(rawS, 8)), '② 소재규격 탭 = 엑셀 I열 「규격」(STS…) 값');
ok(same(await tabValues('소재 로트'), want(rawS, 7)), '② 소재 로트 탭 = 엑셀 H열 값');
ok(!(await page.locator('#materialAlert').count()), '바르게 읽은 자료에는 ② 소재 알림 없음');

// 머리는 「메이커 | 규격」인데 값이 뒤바뀐 파일 → 알림 + 「두 칸 바꾸기」
const HEAD_W = ['작업일자', '작업자', 'Main 압력', '쿠션 압력', '금형', '작업모델', '메이커', '소재 lot', '규격', '작업수량', '불량수량'];
const rawW = rawS.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[8], r[7], r[6], r[9], r[10]]);   // G=규격 값 · I=회사 이름
const wbW = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wbW, XLSX.utils.aoa_to_sheet([HEAD_W, ...rawW]), '작업일보');
const fW = path.join(tmp, '작업일보_값뒤바뀜.xlsx'); fs.writeFileSync(fW, XLSX.write(wbW, { type: 'buffer', bookType: 'xlsx' }));
await importOne(fW, async () => {
  const t = await page.textContent('.mat-check');
  ok(t.includes('소재 메이커 칸에 규격처럼 보이는 값') && t.includes('열 연결을 확인해 주세요'), '값이 뒤바뀐 파일 → 「소재 메이커 칸에 규격처럼 보이는 값이 있습니다 — 열 연결을 확인해 주세요」');
  const sb = page.locator('.mat-check button', { hasText: '두 칸 바꾸기' });
  ok(await sb.count() === 1, '가져오기 「소재 칸 확인」에 「두 칸 바꾸기」 단추');
  if (await sb.count()) { await sb.click(); await page.waitForTimeout(200); }
  const c2 = await page.evaluate(() => { const c = document.querySelector('.mat-check'); return { bad: c.classList.contains('bad'), m: c.querySelector('tbody tr').cells[1].textContent }; });
  ok(!c2.bad && c2.m.startsWith('I열'), '「두 칸 바꾸기」 → 파일을 다시 읽어 메이커 ← I열, 알림 사라짐 (' + c2.m + ')');
});
ok(same(await tabValues('소재 메이커'), want(rawW, 8)) && same(await tabValues('소재규격'), want(rawW, 6)), '바꾼 뒤 ② 메이커 탭 = 회사 이름, 규격 탭 = 규격 값');
const memo = await page.evaluate(() => JSON.parse(localStorage.getItem('data10-11.db')).colOverrides || {});
ok(Object.values(memo).some(o => o.maker === 8 && o.spec === 6), '고친 열 연결을 머리행 모양마다 기억 (' + JSON.stringify(memo) + ')');
await go('#/import');
await page.setInputFiles('input[type=file][multiple]:not([webkitdirectory])', fW); await page.waitForTimeout(800);
const again = await page.evaluate(() => document.querySelector('.mat-check').textContent);
ok(again.includes('전에 고친 열 연결을 그대로') && !again.includes('규격처럼 보이는'), '같은 모양 파일을 다시 넣으면 고친 연결을 그대로 씀');
await page.locator('button', { hasText: '취소' }).click(); await page.waitForTimeout(100);

// 쌓인 행이 뒤바뀌어 있으면(② 화면) 알림 + 두 칸 바꾸기
await page.evaluate(() => { const db = JSON.parse(localStorage.getItem('data10-11.db')); db.records = db.records.map(r => Object.assign({}, r, { maker: r.spec, spec: r.maker })); db.colOverrides = {}; localStorage.setItem('data10-11.db', JSON.stringify(db)); });
await page.reload(); await page.waitForTimeout(300); await go('#/table');
ok(await page.locator('#materialAlert').isVisible() && (await page.textContent('#materialAlert')).includes('소재 메이커 칸에 규격처럼'), '② 쌓인 행의 메이커 · 규격이 뒤바뀌면 알림');
if (await page.locator('#materialAlert button').count()) { await page.locator('#materialAlert button', { hasText: '두 칸 바꾸기' }).click(); await page.waitForTimeout(300); }
ok(!(await page.locator('#materialAlert').count()) && same(await tabValues('소재 메이커'), want(rawW, 8)), '② 「두 칸 바꾸기」 → 메이커 탭 = 회사 이름, 알림 사라짐');

// 오늘 오전 판(열 읽기 판 2)으로 가져온 행 → 다시 가져오기 알림
await page.evaluate(() => { const db = JSON.parse(localStorage.getItem('data10-11.db')); db.records.forEach(r => { r.pv = 2; }); db.files.forEach(f => { f.parser = 2; }); localStorage.setItem('data10-11.db', JSON.stringify(db)); });
await page.reload(); await page.waitForTimeout(300);
ok(await page.locator('#staleImport').isVisible(), '열 읽기 판 2 로 가져온 자료 → 「이전 판에서 가져온 자료입니다」 알림');

await page.setViewportSize({ width: 390, height: 900 });
await go('#/table');
const over = await page.evaluate(() => { const W = document.documentElement.clientWidth, bad = []; document.querySelectorAll('main *').forEach(el => { const r = el.getBoundingClientRect(); if (!r.width) return; let p = el.parentElement; while (p) { const o = getComputedStyle(p).overflowX; if (o === 'auto' || o === 'scroll') return; p = p.parentElement; } if (r.right > W + 1) bad.push(el.tagName + '.' + el.className); }); return bad.slice(0, 3); });
ok(!over.length, '390px ② 가로 넘침 없음 ' + over.join(', '));

ok(!errors.length, 'JS 오류 없음' + (errors.length ? ' — ' + errors.slice(0, 3).join(' | ') : ''));
await browser.close();
console.log('\n검사 ' + checks + '개 · ' + (problems.length ? '문제 ' + problems.length + '건' : '모두 통과'));
process.exitCode = problems.length ? 1 : 0;
