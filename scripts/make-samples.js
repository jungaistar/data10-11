// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 12개월치를 작업일보 모양의 .xlsx 12개(월별)와 금형 보수이력 · 설비(Press) 수선이력 · 압력 작업표준 .xlsx 를 씁니다(전부 지어낸 값).
// 날짜 칸은 엑셀 날짜로 넣고, 방금 쓴 파일을 앱과 같은 방식으로 다시 읽어 행 수·합계가 맞는지 확인합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });
const b = Sample.build();
let recs = [];
for (const f of b.files) {
  const aoa = f.aoa.map((row, i) => i < 3 ? row : row.map((v, c) => (c === 0 && typeof v === 'string' && /^\d{4}-\d\d-\d\d$/.test(v)) ? new Date(v + 'T00:00:00') : v));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'yyyy-mm-dd' });
  ws['!cols'] = [{ wch: 11 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 9 }, { wch: 12 }, { wch: 9 }, { wch: 9 }, { wch: 7 }, { wch: 7 }, { wch: 7 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws, '작업일보');
  const file = path.join(out, f.fileName);
  fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellDates: true }));
  const back = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
  const p = L.parseSheet(XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1, raw: true, defval: null }), { fileName: f.fileName });
  const want = b.rows.filter(r => r.date.slice(0, 7) === f.month);
  if (p.det.missing.length || p.warnings.length || p.records.length !== want.length || p.summary !== 1)
    throw new Error(f.fileName + ' 다시 읽기 실패: ' + JSON.stringify(p.det.warnings.concat(p.warnings).slice(0, 3)) + ' 행 ' + p.records.length + '/' + want.length);
  if (p.records.some((r, i) => r.date !== want[i].date || r.qty !== want[i].qty || r.defect !== want[i].defect)) throw new Error(f.fileName + ' 값이 다름');
  // 소재 메이커 · 소재 lot · 소재규격이 제 칸으로 읽히는지(3차 — 「메이커와 규격이 바뀐 것 같다」 확인)
  if (p.records.some((r, i) => r.maker !== want[i].maker || r.lot !== want[i].lot || r.spec !== want[i].spec)) throw new Error(f.fileName + ' 메이커·로트·규격 칸이 다름');
  recs = L.mergeRecords(recs, p.records).records;
  console.log(f.fileName, '행', p.records.length);
}
// 금형 보수이력 — 설정 화면의 「엑셀에서 붙여 넣기」에 그대로 복사해 넣는 모양
const maint = [['금형', '일자', '구분', '내용']].concat(b.maint.map(m => [m.die, new Date(m.date + 'T00:00:00'), m.kind, m.note]));
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet(maint, { cellDates: true, dateNF: 'yyyy-mm-dd' });
ws2['!cols'] = [{ wch: 10 }, { wch: 11 }, { wch: 6 }, { wch: 26 }];
XLSX.utils.book_append_sheet(wb2, ws2, '금형 보수이력');
fs.writeFileSync(path.join(out, '금형_보수이력_샘플.xlsx'), XLSX.write(wb2, { bookType: 'xlsx', type: 'buffer', cellDates: true }));
console.log('금형_보수이력_샘플.xlsx', '건', b.maint.length);
// 설비(Press) 수선이력 — 설정 화면의 「엑셀 파일로 넣기」나 붙여 넣기에 그대로
function writeSheet(name, sheet, aoa, cols) {
  const wb = XLSX.utils.book_new(), ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'yyyy-mm-dd' });
  ws['!cols'] = cols.map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, sheet);
  fs.writeFileSync(path.join(out, name), XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellDates: true }));
}
writeSheet('설비_수선이력_샘플.xlsx', '설비 수선이력', [['설비', '일자', '구분', '내용']].concat(b.repairs.map(m => [m.equip, new Date(m.date + 'T00:00:00'), m.kind, m.note])), [12, 11, 8, 36]);
const back2 = XLSX.read(fs.readFileSync(path.join(out, '설비_수선이력_샘플.xlsx')), { type: 'buffer', cellDates: true });
const rp = L.parseRepairText(L.aoaToText(XLSX.utils.sheet_to_json(back2.Sheets[back2.SheetNames[0]], { header: 1, raw: true, defval: null })));
if (rp.items.length !== b.repairs.length || rp.bad.length) throw new Error('설비 수선이력 다시 읽기 실패');
console.log('설비_수선이력_샘플.xlsx', '건', b.repairs.length);
// 압력 작업표준 — 작업모델 · Main 기준 · Main 허용(±) · 쿠션 기준 · 쿠션 허용(±)
writeSheet('압력_작업표준_샘플.xlsx', '압력 작업표준', [['작업모델', 'Main 기준', 'Main 허용(±)', '쿠션 기준', '쿠션 허용(±)']].concat(b.stds.rows.map(r => [r.key, r.main, r.mainTol, r.cushion, r.cushionTol])), [11, 10, 12, 10, 12]);
const back3 = XLSX.read(fs.readFileSync(path.join(out, '압력_작업표준_샘플.xlsx')), { type: 'buffer' });
const sp = L.parseStdText(L.aoaToText(XLSX.utils.sheet_to_json(back3.Sheets[back3.SheetNames[0]], { header: 1, raw: true, defval: null })), 'model');
if (sp.stds.rows.length !== b.stds.rows.length || sp.bad.length) throw new Error('압력 작업표준 다시 읽기 실패');
console.log('압력_작업표준_샘플.xlsx', '모델', sp.stds.rows.length);
const cov = L.monthCoverage(recs);
console.log('누적 행', recs.length, '·', cov.first, '~', cov.last, '· 빠진 달', cov.missing.length);
