// 일부러 깨뜨려 검사가 잡는지(전역 규칙 §5.5): node scripts/break-check.mjs  [--browser]
// 원본은 이 스크립트 변수에만 두고 finally 로 되돌립니다(공용 폴더에 사본을 두지 않음 — 1단계 사고 교훈).
// --browser 를 주면 app.js 를 깨뜨린 것은 scripts/check-browser.mjs 로 잽니다(PLAYWRIGHT_CORE 필요).
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath } from 'node:url'; import { execSync } from 'node:child_process';
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const breaks = [
  // [이름, 파일, 찾을 글, 바꿀 글, 잴 것(node|browser)]
  ['「최근 N개월」을 예전처럼 N×30.4일 빼기', 'js/logic.js', "var y = +to.slice(0, 4), m = +to.slice(5, 7) - (n - 1);\n    while (m < 1) { m += 12; y--; }\n    return y + '-' + pad(m) + '-01';", "return addDays(to, -Math.round(n * 30.4) + 1);", 'node'],
  ['옛 판 자료를 알아보지 못함(행의 판 안 봄)', 'js/logic.js', "if (!(Number(r.pv) >= PARSER_VERSION)) out.oldRows++;", "", 'node'],
  ['파일 목록만 보고 판단(다시 넣어 남은 옛 행 놓침)', 'js/logic.js', "out.stale = out.oldRows > 0;", "out.stale = out.oldFiles.length > 0;", 'node'],
  ['행에 열 읽기 판을 안 적음', 'js/logic.js', "res.records.forEach(function (r) { r.pv = PARSER_VERSION; });", "", 'node'],
  ['칸 출처를 머리 글 대신 첫 머리로', 'js/logic.js', "i: c, head: (det.headers && det.headers[c]) || ''", "i: c, head: (det.headers && det.headers[0]) || ''", 'node'],
  ['app.js 판만 올리고 logic.js 는 그대로', 'js/app.js', "var APP_VERSION = '", "var APP_VERSION = 'x", 'node'],
  ['빠른 기간 단추를 예전 계산으로(app.js)', 'js/app.js', "var f0 = months ? L.monthsBackStart(r.to, months) : r.from;", "var f0 = months ? L.addDays(r.to, -Math.round(months * 30.4) + 1) : r.from;", 'browser'],
  ['판이 섞여도 그대로 계산(판 맞추기 빼기)', 'js/app.js', "if (mixed.length || !S || !C || !AIP) {", "if (false) {", 'browser'],
  ['옛 판 알림 빼기', 'js/app.js', "add(main, staleBanner());", "", 'browser'],
  // 2026-10-27 실제 작업일보 머리 이름 점검
  ['「소재」를 예전처럼 규격 이름으로(= 고치기 전)', 'js/logic.js', "'강종', 'spec',", "'강종', '소재', '재질', 'spec',", 'node'],
  ['「원소재」「자재」를 소재 칸으로 알아보지 않음(= 고치기 전)', 'js/logic.js', "'원소재', '원자재', '원자재명', '자재',", "", 'node'],
  ['애매한 소재 칸에서 값 보기 빼기', 'js/logic.js', "kind: valueKind(columnValues(aoa, hr, h.col, 200)) }; });", "kind: null }; });", 'node'],
  ['메이커·규격 값 어긋남 알림 빼기', 'js/logic.js', "out.material = materialCheckCols(aoa, out);", "out.material = {};", 'node'],
  ['열 읽기 판을 2 그대로(오전 판 자료 알림 없음)', 'js/logic.js', "var PARSER_VERSION = 3;", "var PARSER_VERSION = 2;", 'node'],
  ['쌓인 행 바꾸기에서 열쇠를 그대로', 'js/logic.js', "o.key = baseKey(o) + '#' +", "o.key = r.key || baseKey(o) + '#' +", 'node'],
  ['가져오기 「두 칸 바꾸기」가 아무것도 안 함', 'js/app.js', "setOverrides(it, { maker: det.map.spec, spec: det.map.maker });", "", 'browser'],
  ['고친 열 연결을 기억하지 않음', 'js/app.js', "var saved = db.colOverrides && db.colOverrides[L.headerSignature(item.result.det)];", "var saved = null;", 'browser'],
  ['② 쌓인 행 소재 알림 빼기', 'js/app.js', "        storedMaterialAlert(),\n", "", 'browser'],
  ['같은 이름 파일을 다시 넣어도 옛 파일 기록을 남김', 'js/app.js', "db.files = db.files.filter(function (f) { return f.name !== it.fileName; });", "", 'browser']
];
const browser = process.argv.includes('--browser');
const res = [];
for (const [name, file, a, b, how] of breaks) {
  if (how === 'browser' && !browser) { res.push([name, '(--browser 때만)']); continue; }
  const F = path.join(ROOT, file), orig = fs.readFileSync(F, 'utf8');
  if (!orig.includes(a)) { res.push([name, '바꿀 곳을 못 찾음']); continue; }
  try {
    fs.writeFileSync(F, orig.replace(a, b));
    let out = '';
    const cmd = how === 'node' ? 'node test/logic.test.mjs' : 'node scripts/check-browser.mjs';
    try { out = execSync(cmd, { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 64 << 20 }); } catch (e) { out = (e.stdout || '') + (e.stderr || ''); }
    const fails = out.split('\n').filter(l => /FAIL/.test(l));
    if (!/개 통과|검사 \d+개/.test(out)) { res.push([name, '검사가 돌지 않음 — ' + out.trim().split('\n').pop().slice(0, 80)]); continue; }
    res.push([name, fails.length ? '잡힘 — 실패 ' + fails.length + '개: ' + fails[0].trim().slice(0, 90) : '못 잡음!']);
  } finally { fs.writeFileSync(F, orig); }
}
res.forEach(r => console.log(r[0].padEnd(34), r[1]));
console.log('원복 확인: git diff 로 이 스크립트 실행 전과 같은지 보세요');
