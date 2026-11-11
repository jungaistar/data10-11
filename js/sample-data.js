/* 가상 샘플 — 12개월치 Deep drawing 작업일보 (월별 파일 12개, 전부 지어낸 값)
   · 금형 DIE-A01 … , 작업모델 SMP-…, 소재 메이커 소재사A~C, 작업자A~E — 실제 회사·제품과 관계없는 가짜 이름입니다.
   · 열은 제출 글의 작업일보 항목 그대로 + 불량 유형(Crack·주름·기타) 칸. 유형 칸은 기획서 10장 ② 확인 전이라 가정입니다.
   · 앞 6개월 파일과 뒤 6개월 파일은 머리 글자를 일부러 다르게 적었습니다(「작업일자」↔「일자」, 「금형」↔「금형번호」 …).
     열 이름이 조금 달라도 찾는지 보기 위해서입니다. 파일 끝에는 「합계」 줄도 넣었습니다(가져올 때 빠져야 함).
   · 일부러 심어 둔 패턴 세 가지 — ③ 화면에서 「불량이 몰린 조건」으로 바로 보이게:
       ① 금형 DIE-B01 — 보수 뒤 누적 작업량이 늘수록 불량률이 오르고(마모), 보수하면 떨어진다. 보수 이력도 함께 둡니다.
       ② 소재사B — 겨울(12~2월)에만 Crack 이 크게 는다.
       ③ 쿠션 압력이 높게(65 bar 이상) 세팅된 날 — Crack 이 는다.
     Main 압력은 불량과 관계없게 만들었습니다(아무것도 안 걸려야 맞는 「대조군」).
   · 3차(2026-10-27, 기획서 11장)에 더한 것:
       - 「소재 lot」 칸 — 소재 메이커와 소재규격 사이. 메이커마다 달마다 새 로트(예: A-2607). 소재사A 의 2026년 7월 로트(A-2607)만 Crack 이 늘게 심었습니다.
       - 모델별 압력 작업표준(STDS) — 모델마다 Main·쿠션 기준과 허용 폭이 다릅니다. 작업일보의 압력 값은 기준 둘레의 「실제 값」으로 지어냈습니다.
       - 설비(Press) 수선 이력(REPAIRS)과 그에 맞춘 압력 저하 두 번:
           쿠션 압력 저하 2026-05-11~05-21(쿠션 −12, 주름 증가) → 05-22 쿠션 실린더 씰 교체로 회복
           Main 압력 저하 2026-08-03~08-07(Main −22, 불량 변화 없음 — Main 은 계속 대조군) → 08-10 유압 펌프 수선으로 회복
     난수 순서를 1·2차와 같게 두려고 로트·저하는 난수를 더 쓰지 않고 날짜로 정합니다(그래서 다른 행의 값은 그대로입니다).
   브라우저에서는 window.DDSample, node 에서는 require('./sample-data.js'). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DDSample = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var START = '2025-10-01', END = '2026-09-30';
  var HOLIDAYS = ['2025-10-03', '2025-10-06', '2025-10-07', '2025-10-08', '2025-10-09', '2025-12-25', '2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18',
    '2026-03-02', '2026-05-05', '2026-05-25', '2026-08-17', '2026-09-24', '2026-09-25'];
  var DIES = [
    { die: 'DIE-A01', model: 'SMP-C110', spec: 'STS304 0.8t', main: 200, cushion: 50, run: 0.85, wear: 0.1 },
    { die: 'DIE-A02', model: 'SMP-C120', spec: 'STS304 1.0t', main: 210, cushion: 52, run: 0.7, wear: 0.1 },
    { die: 'DIE-B01', model: 'SMP-H210', spec: 'STS430 1.2t', main: 235, cushion: 55, run: 0.8, wear: 1.3 },
    { die: 'DIE-B02', model: 'SMP-H220', spec: 'STS430 1.0t', main: 225, cushion: 53, run: 0.6, wear: 0.1 },
    { die: 'DIE-C01', model: 'SMP-S310', spec: 'STS304 0.6t', main: 185, cushion: 46, run: 0.65, wear: 0.1 }
  ];
  /* 보수·교체 이력 (가상) — DIE-B01 은 마모 패턴을 보이도록 여러 번, DIE-A01 은 두 번, 나머지는 이력 없음 */
  var MAINT = [
    { die: 'DIE-B01', date: '2025-09-26', kind: '보수', note: '펀치 R부 연마(가상)' },
    { die: 'DIE-B01', date: '2025-12-12', kind: '보수', note: '다이 R부 연마·랩핑(가상)' },
    { die: 'DIE-B01', date: '2026-03-13', kind: '보수', note: '다이 R부 연마(가상)' },
    { die: 'DIE-B01', date: '2026-06-12', kind: '교체', note: '다이 인서트 교체(가상)' },
    { die: 'DIE-B01', date: '2026-09-04', kind: '보수', note: '펀치·다이 연마(가상)' },
    { die: 'DIE-A01', date: '2025-09-29', kind: '보수', note: '정기 점검(가상)' },
    { die: 'DIE-A01', date: '2026-04-03', kind: '보수', note: '정기 점검(가상)' }
  ];
  var MAKERS = [{ name: '소재사A', p: 0.55 }, { name: '소재사B', p: 0.3 }, { name: '소재사C', p: 0.15 }];
  /* 모델별 압력 작업표준 (가상) — 기준은 금형마다 세팅하는 값, 허용 폭은 ± */
  var STDS = { by: 'model', rows: DIES.map(function (d) { return { key: d.model, main: d.main, mainTol: 15, cushion: d.cushion, cushionTol: 8 }; }) };
  /* 설비(Press) 수선 이력 (가상) — 프레스 한 대(작업일보에 설비 칸이 없어 모든 작업에 적용) */
  var REPAIRS = [
    { equip: 'Hydro press', date: '2026-01-09', kind: '점검', note: '정기 점검 — 유압유 교환(가상)' },
    { equip: 'Hydro press', date: '2026-05-22', kind: '부품교체', note: '쿠션 실린더 씰 교체 — 누유(가상)' },
    { equip: 'Hydro press', date: '2026-08-10', kind: '수선', note: '메인 유압 펌프 수선·릴리프 밸브 조정(가상)' }
  ];
  var CUSHION_DROP = { from: '2026-05-11', to: '2026-05-21', by: 12 }, MAIN_DROP = { from: '2026-08-03', to: '2026-08-07', by: 22 };
  var BAD_LOT = 'A-2607';
  function lotOf(maker, date) { return maker.slice(-1) + '-' + date.slice(2, 4) + date.slice(5, 7); }
  var WORKERS = ['작업자A', '작업자B', '작업자C', '작업자D', '작업자E'];
  var HEAD_A = ['작업일자', '작업자', 'Hydro press Main 압력(bar)', '쿠션 압력(bar)', '금형', '작업모델', '소재 메이커', '소재 lot', '소재규격', '작업수량', '불량수량', 'Crack', '주름', '기타', '비고'];
  var HEAD_B = ['일자', '작업자', 'Main 압력', '쿠션압력', '금형번호', '모델', '메이커', 'LOT No', '규격', '작업수량', '불량수량', 'Crack', '주름', '기타', '비고'];

  function rng(seed) { // mulberry32 — 같은 씨앗이면 늘 같은 값
    return function () {
      seed |= 0; seed = seed + 0x6D2B79F5 | 0;
      var t = Math.imul(seed ^ seed >>> 15, 1 | seed);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(t) { var d = new Date(t); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function utc(s) { var p = s.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); }
  function seasonOf(s) { var m = +s.slice(5, 7); return m === 12 || m <= 2 ? '겨울' : m <= 5 ? '봄' : m <= 8 ? '여름' : '가을'; }

  /* → { rows:[기록], maint, files:[{fileName, month, aoa}] } */
  function build() {
    var R = rng(20261020), gauss = function () { var u = 1 - R(), v = R(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); };
    function binom(n, p) { if (p <= 0) return 0; var k = 0; for (var i = 0; i < n; i++) if (R() < p) k++; return k; }
    function pick(list) { var x = R(), s = 0; for (var i = 0; i < list.length; i++) { s += list[i].p; if (x < s) return list[i].name; } return list[list.length - 1].name; }
    var cum = {}; DIES.forEach(function (d) { cum[d.die] = 0; });
    var rows = [];
    for (var t = utc(START); t <= utc(END); t += 86400000) {
      var date = iso(t), dow = new Date(t).getUTCDay();
      if (dow === 0 || dow === 6 || HOLIDAYS.indexOf(date) >= 0) continue;
      MAINT.forEach(function (m) { if (m.date === date) cum[m.die] = 0; });
      var season = seasonOf(date);
      DIES.forEach(function (d, di) {
        if (R() > d.run) return;
        var shifts = R() < 0.25 ? 2 : 1;
        for (var s = 0; s < shifts; s++) {
          var qty = Math.max(150, Math.min(1000, Math.round(560 + gauss() * 150)));
          var maker = pick(MAKERS);
          var main = Math.round(d.main + gauss() * 6);
          var highCushion = R() < 0.12;
          var cushion = Math.round(highCushion ? 66 + R() * 9 : d.cushion + gauss() * 3.5);
          var cDrop = date >= CUSHION_DROP.from && date <= CUSHION_DROP.to, mDrop = date >= MAIN_DROP.from && date <= MAIN_DROP.to;
          if (cDrop) cushion -= CUSHION_DROP.by;
          if (mDrop) main -= MAIN_DROP.by;
          var lot = lotOf(maker, date);
          var worker = WORKERS[(di + s + (R() < 0.3 ? 1 : 0)) % WORKERS.length];
          // 불량률 = 바탕 + 심어 둔 패턴 (유형별)
          var pCrack = 0.015, pWr = 0.011, pEtc = 0.006;
          if (season === '겨울') pCrack += 0.003;
          if (maker === '소재사B') pCrack += season === '겨울' ? 0.045 : 0.003;
          if (maker === '소재사C') pWr += 0.003;
          if (cushion >= 65) pCrack += 0.05;
          if (lot === BAD_LOT) pCrack += 0.035;   // 소재 로트 하나만 나쁨
          if (cDrop) pWr += 0.025;                // 쿠션(주름 누름) 압력이 빠지면 주름이 는다
          var w = d.wear * cum[d.die] / 10000 / 100; // 1만 개당 d.wear %p
          pCrack += w * 0.6; pWr += w * 0.4;
          var crack = binom(qty, pCrack), wr = binom(qty, pWr), etc = binom(qty, pEtc);
          var note = '';
          if (MAINT.some(function (m) { return m.die === d.die && m.date === date; }) && s === 0) note = '금형 ' + (MAINT.filter(function (m) { return m.die === d.die && m.date === date; })[0].kind) + ' 후 첫 작업';
          rows.push({ date: date, worker: worker, main: main, cushion: cushion, die: d.die, model: d.model, maker: maker, lot: lot, spec: d.spec,
            qty: qty, defect: crack + wr + etc, crack: crack, wr: wr, etc: etc, note: note });
          cum[d.die] += qty;
        }
      });
    }
    // 월별 파일
    var byMonth = {};
    rows.forEach(function (r) { (byMonth[r.date.slice(0, 7)] || (byMonth[r.date.slice(0, 7)] = [])).push(r); });
    var files = Object.keys(byMonth).sort().map(function (ym, i) {
      var head = ym < '2026-04' ? HEAD_A : HEAD_B;
      var aoa = [['■ Deep drawing 작업일보 — ' + ym.slice(0, 4) + '년 ' + (+ym.slice(5)) + '월 (가상 샘플)'], [], head];
      var tq = 0, td = 0;
      byMonth[ym].forEach(function (r) {
        aoa.push([r.date, r.worker, r.main, r.cushion, r.die, r.model, r.maker, r.lot, r.spec, r.qty, r.defect, r.crack, r.wr, r.etc, r.note || null]);
        tq += r.qty; td += r.defect;
      });
      aoa.push([]);
      aoa.push(['합계', null, null, null, null, null, null, null, null, tq, td, null, null, null, null]);
      return { fileName: '작업일보_샘플_' + ym + '.xlsx', month: ym, aoa: aoa };
    });
    return { rows: rows, maint: MAINT.slice(), stds: JSON.parse(JSON.stringify(STDS)), repairs: REPAIRS.slice(), files: files };
  }

  return { VERSION: '2026-11-11.1', build: build, MAINT: MAINT, STDS: STDS, REPAIRS: REPAIRS, CUSHION_DROP: CUSHION_DROP, MAIN_DROP: MAIN_DROP, BAD_LOT: BAD_LOT, HEAD_A: HEAD_A, HEAD_B: HEAD_B, START: START, END: END };
});
