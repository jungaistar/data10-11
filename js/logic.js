/* Deep drawing 불량률 분석 — 순수 로직 (화면·저장소와 무관)
   브라우저에서는 window.DDLogic, node 에서는 require('./logic.js') 로 씁니다.
   기준 문서: docs/01_프로젝트_기획서.md (3장 데이터, 5장 기능, 8장 1단계)

   흐름: 작업일보 엑셀 → detectColumns(머리행·열 찾기) → parseRows(행 → 작업 기록)
        → mergeRecords(여러 파일 누적·중복 제거) → buildContext(기간·계절·압력 구간·금형 누적 작업량)
        → trend(일별·월별) · groupBy(금형·소재·모델·작업자별) · findHotspots(불량이 몰린 조건)
        · maintEffects(보수 전후) · dieStatus · analyze → buildAiPromptInfo · buildReport */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DDLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SCHEMA_VERSION = 1;
  /* 판(版) — index.html 의 ?v= · app.js 의 APP_VERSION · sample-data.js 와 같아야 합니다(scripts/bump-version.mjs 로 함께 올림).
     브라우저가 옛 파일과 새 파일을 섞어 쓰면(GitHub Pages 캐시 10분) 버튼과 결과가 어긋날 수 있어, app.js 가 시작할 때 맞춰 봅니다. */
  var VERSION = '2026-11-11.1';
  /* 열 읽기 판 — 가져온 파일마다 db.files[].parser 에 적습니다. 2 = 3차(두 줄 머리 「소재 ― 메이커·규격」 수정 · 소재 lot).
     이보다 낮거나 없는 파일로 쌓인 작업 행은 메이커·규격이 뒤바뀌어 있을 수 있어 다시 가져오라고 알립니다(기록에는 원본 칸이 없어 다시 읽을 수 없음). */
  var PARSER_VERSION = 3;
  /* 3 = 2026-10-27 실제 작업일보 머리 이름 점검 — 「소재 | 규격」처럼 애매한 「소재」 머리가 소재규격을 먼저 차지하고 진짜 「규격」 칸을
     「한 번 더 보입니다」로 버리던 결함(메이커 탭이 사라지고 규격 탭에 회사 이름). 판 2 로 읽은 행도 다시 가져오라고 알립니다. */

  /* ── 1. 열 이름 ────────────────────────────────────────────────
     제출 글의 작업일보 항목: 작업일자 · 작업자 · Hydro press main 압력 · 쿠션 압력 · 금형 · 작업모델 · 소재 메이커 · 소재규격 · 작업수량 · 불량수량.
     머리행 글자는 공백·괄호·가운데점·빗금을 빼고 소문자로 맞추고, 끝의 단위(bar·MPa·kgf/cm2·개 …)를 뗀 뒤 별칭과 비교합니다.
     소재 로트·불량유형은 제출 글에 없지만(기획서 10장 ②④ 확인 중) 있으면 함께 읽습니다.
     3차(11장): 압력 기준 칸(「Main 압력 기준」 등)·설비(호기) 칸이 있으면 함께 읽습니다. 없으면 설정의 모델별 작업표준을 씁니다. */
  var FIELDS = [
    { id: 'date', label: '작업일자', req: true, aliases: ['작업일자', '작업일', '일자', '날짜', '생산일자', '작업날짜', 'date'] },
    { id: 'worker', label: '작업자', req: false, aliases: ['작업자', '작업자명', '성명', '작업자이름', '담당자'] },
    { id: 'main', label: 'Hydro press Main 압력', req: false, aliases: ['hydropressmain압력', 'hydropressmain', 'hydropress압력', 'main압력', '메인압력', '주압력', '프레스압력', 'press압력', 'main', 'mainpressure', 'hydro압력', '유압압력'] },
    { id: 'cushion', label: '쿠션 압력', req: false, aliases: ['쿠션압력', '쿠션', '쿠션압', 'cushion압력', 'cushion', 'cushionpressure'] },
    { id: 'mainSpec', label: 'Main 압력 기준(작업표준)', req: false, aliases: ['main압력기준', 'main기준', 'main기준압력', '기준main압력', 'main압력표준', 'main표준', 'mainstd', 'mainspec', '메인압력기준', 'hydropressmain압력기준'] },
    { id: 'cushionSpec', label: '쿠션 압력 기준(작업표준)', req: false, aliases: ['쿠션압력기준', '쿠션기준', '쿠션기준압력', '기준쿠션압력', '쿠션압력표준', '쿠션표준', 'cushionstd', 'cushionspec'] },
    { id: 'press', label: '설비(프레스 호기)', req: false, aliases: ['설비', '설비명', '설비번호', '설비no', '프레스', '프레스호기', '호기', 'press', 'pressno', '장비', '장비명'] },
    { id: 'die', label: '금형', req: true, aliases: ['금형', '금형번호', '금형no', '금형명', '금형코드', 'die', 'dieno', '다이'] },
    { id: 'model', label: '작업모델', req: false, aliases: ['작업모델', '모델', '모델명', '제품', '제품명', '품번', '품명', '작업품번'] },
    // 소재 메이커 · 소재규격은 「그 칸이라고 분명한 이름」만 여기 둡니다. 「소재」「재질」「제조사」처럼 어느 쪽인지 애매한 이름은
    // 아래 MATERIAL_HINTS 로 따로 모아, 분명한 칸을 먼저 연결한 뒤 남은 쪽 · 아래 값 모양으로 정합니다(resolveMaterial).
    { id: 'maker', label: '소재 메이커', req: false, aliases: ['소재메이커', '메이커', '메이커명', '소재업체', '소재제조사', '소재공급사', '소재공급업체', '공급업체', '소재사', 'maker', 'makername', '소재maker', 'supplier', 'vendor', 'manufacturer'] },
    { id: 'spec', label: '소재규격', req: false, aliases: ['소재규격', '규격', '소재사양', '사양', '소재스펙', '스펙', '재질규격', '강종', 'spec', 'specification', 'materialspec', 'grade'] },
    { id: 'lot', label: '소재 로트', req: false, aliases: ['소재로트', '로트', 'lot', 'lotno', '로트번호', '소재lot', 'heatno', '히트번호', 'coilno', '코일번호'] },
    { id: 'qty', label: '작업수량', req: true, aliases: ['작업수량', '생산수량', '투입수량', '작업량', '수량', '총수량', '생산량'] },
    { id: 'defect', label: '불량수량', req: true, aliases: ['불량수량', '불량', '불량수', 'ng수량', 'ng', '불량합계', '총불량', '불량계'] },
    { id: 'defectType', label: '불량유형(글자 한 칸)', req: false, aliases: ['불량유형', '불량구분', '불량내용', '불량현상', '불량명'] },
    { id: 'note', label: '비고', req: false, aliases: ['비고', '특이사항', '메모'] }
  ];
  /* 애매한 소재 머리 — amb: 메이커일 수도 규격일 수도(「소재」「재질」), maker: 메이커 쪽으로 기우는 일반 이름(「제조사」 — 금형·설비 제조사일 수도 있어 분명하지는 않음) */
  var MATERIAL_HINTS = {
    amb: ['소재', '소재명', '원소재', '원자재', '원자재명', '자재', '자재명', '자재코드', '소재코드', '재료', '재질', '소재정보', 'material', 'materialname', 'mat'],
    maker: ['제조사', '공급사', '제조업체', '제조사명', '공급사명']
  };
  var FIELD_BY_ID = {};
  FIELDS.forEach(function (f) { FIELD_BY_ID[f.id] = f; });

  /* 불량 유형별 수량 열 — 「Crack」「주름」 같은 머리의 열이 있으면 유형별로 읽습니다 */
  var DEFECT_TYPES = [
    { id: 'Crack', aliases: ['crack', '크랙', '균열', '터짐', 'crack수량', '크랙수량'] },
    { id: '주름', aliases: ['주름', 'wrinkle', '주름수량'] },
    { id: '찢어짐', aliases: ['찢어짐', '찢김', '파단', 'tear', '찢어짐수량'] },
    { id: '스크래치', aliases: ['스크래치', '긁힘', '찍힘', 'scratch'] },
    { id: '치수', aliases: ['치수', '치수불량'] },
    { id: '기타', aliases: ['기타', '기타불량'] }
  ];

  var UNIT_RE = /(bar|mpa|kgfcm2|kgf㎠|kg㎠|kgcm2|㎏㎠|㎏fcm2|ton|톤|kn|ea|pcs|개)$/;
  function norm(s) {
    var n = String(s == null ? '' : s).replace(/[\s()（）\[\]·.:_\/\-~]/g, '').toLowerCase();
    var m = n.replace(UNIT_RE, '');
    return m || n;
  }
  function str(v) { return v == null ? '' : String(v).trim(); }

  /* ── 2. 숫자·날짜 ─────────────────────────────────────────── */
  /* 빈칸·공백·「-」는 null, 숫자로 못 읽으면 NaN(경고용). 「120 bar」「1,200개」처럼 단위가 붙어도 읽습니다. */
  function toNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (typeof v === 'boolean') return NaN;
    var s = String(v).trim();
    if (!s || s === '-') return null;
    if (s.charAt(0) === '=') return null; // 계산값이 저장되지 않은 수식
    s = s.replace(/,/g, '').replace(/\s*(bar|mpa|kgf\/cm2|kgf\/㎠|㎏\/㎠|kg\/cm2|ton|톤|개|ea|pcs)$/i, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
    return Number(s);
  }
  function round(n, d) { var p = Math.pow(10, d == null ? 2 : d); return Math.round(n * p) / p; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function isoFromUTC(ms) { var d = new Date(ms); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function isoOk(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
    var t = Date.UTC(y, m - 1, d); var x = new Date(t);
    return x.getUTCMonth() === m - 1 ? isoFromUTC(t) : '';
  }
  /* 엑셀 날짜(일련번호 · Date · 글자) → 'YYYY-MM-DD'. 못 읽으면 ''.
     Date 는 시간대 때문에 전날 23:59 로 들어오는 일이 있어 12시간을 더해 날짜만 씁니다. */
  function toISO(v, defaultYear) {
    if (v == null || v === '') return '';
    if (v instanceof Date || Object.prototype.toString.call(v) === '[object Date]') {
      if (isNaN(v.getTime())) return '';
      var d = new Date(v.getTime() + 12 * 3600 * 1000);
      return isoOk(d.getFullYear(), d.getMonth() + 1, d.getDate());
    }
    if (typeof v === 'number') {
      if (v < 20000 || v > 80000) return '';
      return isoFromUTC(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    }
    var s = String(v).trim();
    var m = /^(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/.exec(s);
    if (m) return isoOk(+m[1], +m[2], +m[3]);
    m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$/.exec(s);
    if (m) return isoOk(2000 + +m[1], +m[2], +m[3]);
    m = /^(\d{1,2})[-./월\s]+(\d{1,2})일?(\s*\(.\))?$/.exec(s);
    if (m && defaultYear) return isoOk(defaultYear, +m[1], +m[2]);
    if (/^\d{5}(\.\d+)?$/.test(s)) return toISO(Number(s));
    return '';
  }
  function isoToUTC(iso) { var p = iso.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); }
  function addDays(iso, n) { return isoFromUTC(isoToUTC(iso) + n * 86400000); }
  function daysBetween(a, b) { return Math.round((isoToUTC(b) - isoToUTC(a)) / 86400000); }

  /* 계절 — 봄 3~5월 · 여름 6~8월 · 가을 9~11월 · 겨울 12~2월 */
  var SEASONS = ['봄', '여름', '가을', '겨울'];
  function seasonOf(iso) {
    var m = +String(iso).slice(5, 7);
    return m >= 3 && m <= 5 ? '봄' : m >= 6 && m <= 8 ? '여름' : m >= 9 && m <= 11 ? '가을' : '겨울';
  }
  function monthOf(iso) { return String(iso).slice(0, 7); }
  function monthLabel(ym) { return ym ? ym.slice(0, 4) + '년 ' + (+ym.slice(5, 7)) + '월' : ''; }
  function colName(i) {
    var s = ''; i = Number(i);
    do { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; } while (i >= 0);
    return s;
  }

  /* ── 3. 머리행 찾기 ────────────────────────────────────────── */
  function matchType(text) {
    var n = norm(text);
    if (!n) return null;
    for (var i = 0; i < DEFECT_TYPES.length; i++) if (DEFECT_TYPES[i].aliases.indexOf(n) >= 0) return DEFECT_TYPES[i].id;
    if (/crack|크랙/.test(n)) return 'Crack';
    return null;
  }
  function matchField(text) {
    var n = norm(text);
    if (!n) return null;
    for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i].aliases.indexOf(n) >= 0) return FIELDS[i].id;
    // 비슷한 머리(예: 「Hydro press 압력」「쿠션 설정압」「소재 MAKER」)
    // 「기준·표준·std」가 붙으면 작업표준 칸입니다(「설정압」은 지금처럼 압력 값으로 읽음 — 작업일보에 설정값을 적는 곳이 많아서)
    var isStd = /(기준|표준|std|spec)/.test(n);
    if (/쿠션|cushion/.test(n)) return isStd ? 'cushionSpec' : 'cushion';
    if (/(main|메인|hydro|주압)/.test(n) && /(압|press|pressure)/.test(n)) return isStd ? 'mainSpec' : 'main';
    if (/메이커|maker|supplier|vendor/.test(n)) return 'maker';
    if (/로트|lot/.test(n)) return 'lot';
    if (/(소재|재질|자재|원자재|material).*(규격|사양|스펙|spec|grade)/.test(n)) return 'spec';
    if (/소재.*(업체|제조|공급)/.test(n)) return 'maker';
    return null;
  }
  function materialHint(text) {
    var n = norm(text);
    if (!n) return null;
    if (MATERIAL_HINTS.amb.indexOf(n) >= 0) return 'amb';
    if (MATERIAL_HINTS.maker.indexOf(n) >= 0) return 'maker';
    return null;
  }

  /* 값 모양 — 소재규격처럼 보이는 값(STS304 0.8t · SPCC 1.2T · AL5052 1.5mm · 1.0×1219) / 메이커(회사) 이름처럼 보이는 값 */
  var GRADE_RE = /(^|[^a-z])(sts|sus|spcc|spcd|spce|sphc|sphd|sphe|scp|saph|sgcc|secc|sghc|sgcd|swch|ss\d|sm\d|al\d|a\d{4}|c\d{4}|aisi|gi|ga|egi|hgi|cgi|dp\d|hsla)(?![a-z])/i;
  var THICK_RE = /\d(\.\d+)?\s*(t|mm|㎜)(?![a-z])/i;
  var DIM_RE = /\d\s*[x×*]\s*\d/i;
  var MAT_WORD_RE = /(스테인리스|스텐|알루미늄|냉연|열연|아연|도금|갈바|강판|코일|판재|후판|박판)/;
  var COMPANY_RE = /(㈜|\(주\)|주식회사|제철|철강|스틸|금속|메탈|상사|산업|공업|소재사|steel|metal|corp|inc\b|co\.|ltd)/i;
  function looksSpec(v) { var s = str(v); return !!s && (GRADE_RE.test(s) || THICK_RE.test(s) || DIM_RE.test(s) || MAT_WORD_RE.test(s)); }
  function looksMaker(v) {
    var s = str(v);
    if (!s || looksSpec(s)) return false;
    if (COMPANY_RE.test(s)) return true;
    return !/\d/.test(s) && /[a-z가-힣]/i.test(s);   // 숫자 없는 이름(로트 번호 · 수량은 아님)
  }
  /* 값 목록이 어느 쪽인지 — 'spec' | 'maker' | null(모름). 비지 않은 값이 3개 미만이면 모름 */
  function valueKind(values) {
    var n = 0, sp = 0, mk = 0;
    (values || []).forEach(function (v) { if (str(v) === '') return; n++; if (looksSpec(v)) sp++; else if (looksMaker(v)) mk++; });
    if (n < 3) return null;
    if (sp / n >= 0.6) return 'spec';
    if (mk / n >= 0.6 && sp / n <= 0.1) return 'maker';
    return null;
  }

  /* aoa: 시트를 행 배열로 읽은 것(SheetJS sheet_to_json header:1).
     반환: { headerRow, groupRow, map:{field:col}, typeCols:[{type,col}], headers:[글자], missing:[field], warnings:[] }
     머리가 두 줄(윗줄 「압력」, 아랫줄 「Main」「쿠션」)이어도 윗줄 글자를 붙여 다시 맞춰 봅니다. */
  function detectColumns(aoa, overrides) {
    aoa = aoa || [];
    var limit = Math.min(aoa.length, 30);
    function groupsAbove(r, width) {
      var g = [], cur = '', any = false, row = r > 0 ? (aoa[r - 1] || []) : [];
      for (var c = 0; c < width; c++) { if (str(row[c])) { cur = str(row[c]); any = true; } g.push(cur); }
      return any ? g : [];
    }
    function readRow(r) {
      var row = aoa[r] || [], width = row.length;
      for (var k = r; k < Math.min(aoa.length, r + 3); k++) width = Math.max(width, (aoa[k] || []).length);
      var groups = groupsAbove(r, width);
      var map = {}, typeCols = [], headers = [], dup = [], hints = [];
      for (var c = 0; c < width; c++) {
        var t = str(row[c]);
        var fid = matchField(t), ty = fid ? null : matchType(t);
        var label = t;
        if (t && !fid && !ty && groups[c] && groups[c] !== t) {
          fid = matchField(groups[c] + t); ty = fid ? null : matchType(groups[c] + t);
          if (fid || ty) label = groups[c] + ' ' + t;
        }
        headers.push(label);
        if (fid) { if (map[fid] == null) map[fid] = c; else dup.push({ label: label, fid: fid }); }
        else if (ty && !typeCols.some(function (x) { return x.type === ty; })) typeCols.push({ type: ty, col: c });
        else if (!ty && materialHint(t)) hints.push({ col: c, hint: materialHint(t), label: label });
      }
      return { map: map, typeCols: typeCols, headers: headers, dup: dup, groups: groups, hints: hints };
    }
    /* 두 줄 머리 — 윗줄이 머리행이고 아랫줄에 작은 머리가 있는 모양.
       예) 윗줄: 작업일자 | 압력(가로 병합) | | 소재(가로 병합) | | 작업수량 …   아랫줄: | Main | 쿠션 | 메이커 | 규격 | …
       「작업일자」처럼 세로로 병합된 칸은 아랫줄이 비어 있어 윗줄만 보면 머리행이 되는데, 그러면 「소재」 칸 하나가
       소재규격으로 읽혀 메이커 값이 소재규격으로 들어갔습니다(3차 수정 — 기획서 11.1). 아랫줄에 글자가 있으면 「윗줄 + 아랫줄」로 다시 맞춥니다. */
    function isSubHeader(r) {
      var row = aoa[r + 1] || [], txt = 0;
      for (var c = 0; c < row.length; c++) {
        var v = row[c];
        if (v == null || str(v) === '') continue;
        if (typeof v === 'number' || v instanceof Date || Object.prototype.toString.call(v) === '[object Date]' || toISO(v)) return false; // 자료 줄
        if (!isNaN(toNum(v)) && toNum(v) != null) return false;
        txt++;
      }
      return txt > 0;
    }
    function readTwo(r) {
      var top = aoa[r] || [], sub = aoa[r + 1] || [], width = Math.max(top.length, sub.length);
      var map = {}, typeCols = [], headers = [], dup = [], hints = [], cur = '';
      for (var c = 0; c < width; c++) {
        var t = str(top[c]), b = str(sub[c]);
        if (t) cur = t; else if (!b) cur = '';
        var g = t || cur, fid = null, ty = null, label = t, hint = null;
        if (b) {
          fid = g ? matchField(g + b) : null; ty = fid || !g ? null : matchType(g + b);
          if (!fid && !ty) { fid = matchField(b); ty = fid ? null : matchType(b); }
          label = g && g !== b ? g + ' ' + b : b;
          if (!fid && !ty) hint = materialHint(b) || (g ? materialHint(g) : null);
        } else if (t) { fid = matchField(t); ty = fid ? null : matchType(t); if (!fid && !ty) hint = materialHint(t); }
        headers.push(label);
        if (fid) { if (map[fid] == null) map[fid] = c; else dup.push({ label: label, fid: fid }); }
        else if (ty && !typeCols.some(function (x) { return x.type === ty; })) typeCols.push({ type: ty, col: c });
        else if (hint) hints.push({ col: c, hint: hint, label: label });
      }
      return { map: map, typeCols: typeCols, headers: headers, dup: dup, groups: [], two: true, hints: hints };
    }
    function scoreOf(rd) { return Object.keys(rd.map).length + (rd.typeCols.length ? 1 : 0); }
    function usable(m) { return m.date != null && m.qty != null && (m.defect != null || m.die != null); }
    var best = -1, bestScore = 0, bestRead = null;
    for (var r = 0; r < limit; r++) {
      var rd = readRow(r);
      var score = scoreOf(rd);
      if (usable(rd.map) && score > bestScore) { best = r; bestScore = score; bestRead = rd; }
      if (r + 1 < aoa.length && isSubHeader(r)) {
        var r2 = readTwo(r), s2 = scoreOf(r2) + r2.dup.filter(function (d) { return d.fid === 'main' || d.fid === 'cushion'; }).length; // 「설정·실제」 두 칸도 셈
        if (usable(r2.map) && s2 > bestScore) { best = r + 1; bestScore = s2; bestRead = r2; }
      }
    }
    var out = { headerRow: best, groupRow: -1, map: {}, typeCols: [], headers: [], missing: [], warnings: [] };
    if (best < 0 || bestScore < 3) {
      out.headerRow = -1;
      out.warnings.push('머리행(작업일자·금형·작업수량·불량수량 …)을 찾지 못했습니다. 작업일보 시트인지 확인해 주세요.');
      FIELDS.forEach(function (f) { if (f.req) out.missing.push(f.id); });
      return out;
    }
    out.map = bestRead.map; out.typeCols = bestRead.typeCols; out.headers = bestRead.headers;
    if (bestRead.groups.length || bestRead.two) out.groupRow = best - 1;
    resolveMaterial(aoa, best, out, bestRead.hints || []);
    // 같은 압력 칸이 두 번(예: 「Main 압력 설정」「Main 압력 실제」) — 「실제·측정」 쪽을 압력 값으로, 다른 쪽을 기준으로
    bestRead.dup = bestRead.dup.filter(function (d) {
      if (d.fid !== 'main' && d.fid !== 'cushion') return true;
      var sp = d.fid + 'Spec';
      if (out.map[sp] != null) return true;
      var col = out.headers.indexOf(d.label), first = out.map[d.fid];
      if (/(실제|측정|실측|act)/i.test(d.label)) { out.map[sp] = first; out.map[d.fid] = col; }
      else if (/(실제|측정|실측|act)/i.test(out.headers[first])) out.map[sp] = col;
      else return true;
      out.warnings.push(FIELD_BY_ID[d.fid].label + ' 칸이 두 개라 「' + out.headers[out.map[d.fid]] + '」 을 실제 값, 「' + out.headers[out.map[sp]] + '」 을 기준 값으로 읽습니다.');
      return false;
    });
    bestRead.dup.forEach(function (d) { out.warnings.push('「' + d.label + '」 열이 ' + FIELD_BY_ID[d.fid].label + '(으)로 한 번 더 보입니다. 앞의 것(' + colName(out.map[d.fid]) + ')을 씁니다.'); });
    // 사용자가 고른 연결이 있으면 덮어씁니다 (값 -1 = 이 파일에는 없음)
    if (overrides) Object.keys(overrides).forEach(function (fid) {
      if (!FIELD_BY_ID[fid]) return;
      var v = overrides[fid];
      if (v === -1 || v === '' || v == null) delete out.map[fid]; else out.map[fid] = Number(v);
    });
    out.material = materialCheckCols(aoa, out);
    if (out.material.warn) out.warnings.unshift(out.material.warn);
    FIELDS.forEach(function (f) { if (f.req && out.map[f.id] == null) out.missing.push(f.id); });
    if (out.missing.length) out.warnings.push('꼭 필요한 열이 없습니다: ' + out.missing.map(function (id) { return FIELD_BY_ID[id].label; }).join(', ') + ' — 열 연결에서 골라 주세요.');
    var soft = ['main', 'cushion', 'maker', 'worker'].filter(function (id) { return out.map[id] == null; });
    if (soft.length) out.warnings.push('이 열이 없어 해당 분석이 비어 보입니다: ' + soft.map(function (id) { return FIELD_BY_ID[id].label; }).join(', ') + '. 있는데 못 찾았다면 열 연결에서 골라 주세요.');
    return out;
  }

  /* 머리행 아래 자료 줄에서 한 열의 값(비지 않은 것, 최대 n개, 합계 줄 제외) */
  function columnValues(aoa, headerRow, col, n) {
    var out = [];
    if (col == null || headerRow < 0) return out;
    for (var r = headerRow + 1; r < aoa.length && out.length < (n || 200); r++) {
      var row = aoa[r] || [];
      if (isSummaryRow(row)) continue;
      var v = row[col];
      if (v instanceof Date || Object.prototype.toString.call(v) === '[object Date]') v = toISO(v);
      if (str(v) !== '') out.push(str(v));
    }
    return out;
  }
  /* 미리 보기용 — 서로 다른 값 앞에서 n개 */
  function columnSamples(aoa, headerRow, col, n) {
    var seen = {}, out = [];
    columnValues(aoa, headerRow, col, 200).forEach(function (v) { if (out.length < (n || 3) && !seen[v]) { seen[v] = 1; out.push(v); } });
    return out;
  }
  /* 애매한 소재 머리(「소재」「재질」「제조사」 …)를 메이커 · 규격 중 비어 있는 쪽에 — 분명한 이름이 먼저, 그다음 아래 값 모양, 그다음 이름이 기우는 쪽,
     마지막으로 둘 다 비면 작업일보 순서(메이커 → 규격, 칸이 하나면 규격 — 예전 동작). 쓴 결정은 알림에 적어 열 연결 카드에서 보이게 합니다. */
  function resolveMaterial(aoa, hr, out, hints) {
    var used = {}; Object.keys(out.map).forEach(function (f) { used[out.map[f]] = 1; });
    var cands = hints.filter(function (h) { return !used[h.col]; }).map(function (h) { return { col: h.col, hint: h.hint, label: h.label, kind: valueKind(columnValues(aoa, hr, h.col, 200)) }; });
    if (!cands.length) return;
    function free(f) { return out.map[f] == null; }
    function take(c, f, why) {
      out.map[f] = c.col; c.done = true;
      out.notes = out.notes || [];
      out.notes.push('「' + c.label + '」 열(' + colName(c.col) + ')을 ' + FIELD_BY_ID[f].label + '(으)로 읽었습니다 — ' + why + '.');
    }
    cands.forEach(function (c) { if (c.kind && free(c.kind)) take(c, c.kind, c.kind === 'spec' ? '아래 값이 규격 모양(STS304 0.8t 등)' : '아래 값이 회사 이름 모양'); });
    cands.forEach(function (c) { if (!c.done && !c.kind && c.hint === 'maker' && free('maker')) take(c, 'maker', '이름이 메이커 쪽'); });
    var rest = cands.filter(function (c) { return !c.done && !c.kind; });
    var bothFree = free('maker') && free('spec');
    rest.forEach(function (c, i) {
      if (free('maker') && !free('spec')) take(c, 'maker', '소재규격 칸이 따로 있어서');
      else if (free('spec') && !free('maker')) take(c, 'spec', '소재 메이커 칸이 따로 있어서');
      else if (bothFree && free('maker') && free('spec')) take(c, rest.length > 1 && i === 0 ? 'maker' : 'spec', rest.length > 1 ? '작업일보 순서(메이커 → 규격)' : '소재 칸이 하나라 규격으로(값으로는 가르지 못함)');
    });
    // 값이 반대 모양인데 그쪽 칸이 이미 찼으면 연결하지 않고 알림만
    cands.forEach(function (c) { if (!c.done) (out.notes = out.notes || []).push('「' + c.label + '」 열(' + colName(c.col) + ')은 소재 메이커 · 소재규격 칸이 이미 있어 연결하지 않았습니다. 이 칸이 맞다면 열 연결에서 골라 주세요.'); });
  }
  /* 연결된 메이커 · 규격 칸의 값이 반대 모양인지 — { maker:{col,head,kind,samples}, spec:{…}, swap, warn } */
  function materialCheckCols(aoa, det) {
    var out = { maker: null, spec: null, swap: false, warn: '' };
    ['maker', 'spec'].forEach(function (f) {
      var c = det.map[f];
      if (c == null) return;
      out[f] = { col: c, head: (det.headers && det.headers[c]) || '', kind: valueKind(columnValues(aoa, det.headerRow, c, 200)), samples: columnSamples(aoa, det.headerRow, c, 3) };
    });
    return materialVerdict(out);
  }
  function materialVerdict(out) {
    var mk = out.maker && out.maker.kind, sp = out.spec && out.spec.kind;
    if (mk === 'spec' && sp === 'maker') { out.swap = true; out.warn = '소재 메이커 칸에 규격처럼 보이는 값이 있고, 소재규격 칸에는 회사 이름이 있습니다 — 두 칸이 서로 바뀐 것 같습니다. 열 연결을 확인해 주세요(「두 칸 바꾸기」).'; }
    else if (mk === 'spec') { out.swap = !!out.spec; out.warn = '소재 메이커 칸에 규격처럼 보이는 값이 있습니다 — 열 연결을 확인해 주세요.'; }
    else if (sp === 'maker' && mk !== 'maker') { out.swap = !!out.maker; out.warn = '소재규격 칸에 메이커(회사) 이름처럼 보이는 값이 있습니다 — 열 연결을 확인해 주세요.'; }
    return out;
  }
  /* 쌓인 작업 행 점검(② 화면) — 행의 메이커 · 규격 값 모양이 반대인지 */
  function materialCheck(records) {
    var mk = [], sp = [];
    (records || []).forEach(function (r) { if (mk.length < 3000) { mk.push(r.maker); sp.push(r.spec); } });
    var anyMk = mk.some(function (v) { return str(v) !== ''; }), anySp = sp.some(function (v) { return str(v) !== ''; });
    return materialVerdict({ maker: anyMk ? { kind: valueKind(mk) } : null, spec: anySp ? { kind: valueKind(sp) } : null, swap: false, warn: '' });
  }
  /* 쌓인 작업 행의 메이커 ↔ 규격 맞바꾸기 — 열쇠도 바뀐 값으로(같은 파일을 다시 넣어도 두 번 세지 않게) */
  function swapMaterial(records) {
    return (records || []).map(function (r) {
      var o = {}; Object.keys(r).forEach(function (k) { o[k] = r[k]; });
      o.maker = r.spec || ''; o.spec = r.maker || '';
      var n = String(r.key || '').split('#').pop();
      o.key = baseKey(o) + '#' + (/^\d+$/.test(n) ? n : 1);
      return o;
    });
  }

  /* ── 4. 행 → 작업 기록 ─────────────────────────────────────── */
  function isSummaryRow(row) {
    for (var c = 0; c < Math.min(row.length, 8); c++) {
      if (/^(합계|소계|총계|계|누계|평균|total|sum|subtotal)$/i.test(str(row[c]).replace(/\s/g, ''))) return true;
    }
    return false;
  }
  // 설비(호기) 칸은 3차에 더했습니다 — 칸이 없는 작업일보는 열쇠가 1단계와 같게(다시 넣어도 두 번 세지 않게) 값이 있을 때만 붙입니다
  function baseKey(r) { return [r.date, r.die, r.model, r.worker, r.maker, r.spec, r.lot, r.main, r.cushion].join('|') + (r.press ? '|' + r.press : ''); }

  /* opts: { fileName, defaultYear } → { records, warnings:[{row,msg}], skipped, summary, blankDefect } */
  function parseRows(aoa, det, opts) {
    opts = opts || {};
    var out = { records: [], warnings: [], skipped: 0, summary: 0, blankDefect: 0 };
    if (!det || det.headerRow < 0 || det.missing.length) return out;
    var m = det.map;
    var year = opts.defaultYear || guessYear(aoa, det) || new Date().getFullYear();
    function cell(row, fid) { return m[fid] == null ? null : row[m[fid]]; }
    function num(row, fid, r) {
      var v = toNum(cell(row, fid));
      if (v != null && isNaN(v)) { out.warnings.push({ row: r + 1, msg: FIELD_BY_ID[fid].label + ' 칸 「' + str(cell(row, fid)) + '」 을 숫자로 읽지 못해 비워 둡니다.' }); return null; }
      return v;
    }
    var occ = {};
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var rawDate = cell(row, 'date');
      var empty = ['date', 'die', 'qty', 'defect'].every(function (f) { var v = cell(row, f); return v == null || str(v) === ''; });
      if (empty) continue; // 빈 줄
      if (isSummaryRow(row)) { out.summary++; continue; } // 합계·소계 줄
      var date = toISO(rawDate, year);
      if (!date) { out.skipped++; out.warnings.push({ row: r + 1, msg: '작업일자 「' + str(rawDate) + '」 를 날짜로 읽지 못해 건너뜁니다.' }); continue; }
      var qty = num(row, 'qty', r);
      if (qty == null || qty <= 0) { out.skipped++; out.warnings.push({ row: r + 1, msg: '작업수량이 비어 있거나 0 이하라 건너뜁니다.' }); continue; }
      var dRaw = cell(row, 'defect');
      var defect = num(row, 'defect', r);
      if (defect == null) { if (dRaw == null || str(dRaw) === '' || str(dRaw) === '-') { defect = 0; out.blankDefect++; } else { out.skipped++; continue; } }
      if (defect < 0) { out.skipped++; out.warnings.push({ row: r + 1, msg: '불량수량이 음수(' + defect + ')라 건너뜁니다.' }); continue; }
      if (defect > qty) { out.skipped++; out.warnings.push({ row: r + 1, msg: '불량수량(' + defect + ')이 작업수량(' + qty + ')보다 커서 건너뜁니다. 칸이 바뀌지 않았는지 확인해 주세요.' }); continue; }
      var die = str(cell(row, 'die'));
      if (!die) { die = '(금형 없음)'; out.warnings.push({ row: r + 1, msg: '금형 칸이 비어 「(금형 없음)」으로 묶습니다.' }); }
      var types = {}, tsum = 0;
      det.typeCols.forEach(function (tc) {
        var v = toNum(row[tc.col]);
        if (v != null && isNaN(v)) { out.warnings.push({ row: r + 1, msg: tc.type + ' 칸 「' + str(row[tc.col]) + '」 을 숫자로 읽지 못했습니다.' }); return; }
        if (v > 0) { types[tc.type] = v; tsum += v; }
      });
      var dtype = str(cell(row, 'defectType'));
      if (!det.typeCols.length && dtype && defect > 0) { types[dtype] = defect; tsum = defect; }
      if (tsum > defect) out.warnings.push({ row: r + 1, msg: '불량 유형별 합계(' + tsum + ')가 불량수량(' + defect + ')보다 큽니다.' });
      var main = num(row, 'main', r), cushion = num(row, 'cushion', r);
      if (main != null && main <= 0) { out.warnings.push({ row: r + 1, msg: 'Main 압력 ' + main + ' 은 0 이하라 비워 둡니다.' }); main = null; }
      if (cushion != null && cushion < 0) { out.warnings.push({ row: r + 1, msg: '쿠션 압력 ' + cushion + ' 은 음수라 비워 둡니다.' }); cushion = null; }
      var mainSpec = num(row, 'mainSpec', r), cushionSpec = num(row, 'cushionSpec', r);
      if (mainSpec != null && mainSpec <= 0) mainSpec = null;
      if (cushionSpec != null && cushionSpec < 0) cushionSpec = null;
      var rec = {
        date: date,
        worker: str(cell(row, 'worker')),
        main: main,
        cushion: cushion,
        die: die,
        model: str(cell(row, 'model')),
        maker: str(cell(row, 'maker')),
        spec: str(cell(row, 'spec')),
        lot: str(cell(row, 'lot')),
        press: str(cell(row, 'press')),
        mainSpec: mainSpec,
        cushionSpec: cushionSpec,
        qty: qty,
        defect: defect,
        types: types,
        note: str(cell(row, 'note')),
        file: opts.fileName || ''
      };
      // 같은 파일 안에서 모든 칸이 같은 행은 「몇 번째」를 붙여 따로 셉니다(같은 날 같은 조건으로 두 번 작업한 경우).
      // 같은 파일을 다시 넣으면 같은 열쇠가 나오므로 두 번 세지 않습니다.
      var bk = baseKey(rec);
      occ[bk] = (occ[bk] || 0) + 1;
      rec.key = bk + '#' + occ[bk];
      out.records.push(rec);
    }
    return out;
  }
  function guessYear(aoa, det) {
    var m = det.map;
    for (var r = det.headerRow + 1; r < Math.min(aoa.length, det.headerRow + 80); r++) {
      var d = toISO(m.date != null ? (aoa[r] || [])[m.date] : null);
      if (d) return +d.slice(0, 4);
    }
    return 0;
  }

  /* 시트 하나를 한 번에 — 앱과 테스트가 같은 길을 씁니다 */
  function parseSheet(aoa, opts) {
    opts = opts || {};
    var det = detectColumns(aoa, opts.overrides);
    var res = parseRows(aoa, det, opts);
    res.records.forEach(function (r) { r.pv = PARSER_VERSION; });   // 어느 열 읽기 판으로 읽은 행인지(importStatus)
    if (res.blankDefect) det.warnings.push('불량수량이 빈 행 ' + res.blankDefect + '개는 불량 0 으로 셉니다.');
    return { det: det, records: res.records, warnings: res.warnings, skipped: res.skipped, summary: res.summary };
  }
  /* 여러 시트 중 작업일보 시트 고르기 — 머리행 점수가 가장 높은 시트 */
  function pickSheet(sheets) {
    var best = null, bestN = -1;
    (sheets || []).forEach(function (s) {
      var det = detectColumns(s.aoa);
      var n = det.headerRow < 0 ? -1 : Object.keys(det.map).length + det.typeCols.length;
      if (n > bestN) { bestN = n; best = s; }
    });
    return bestN < 0 ? null : best;
  }
  function headerSignature(det) { return det && det.headers ? det.headers.map(norm).join('|') : ''; }

  /* ── 5. 누적 · 중복 제거 ─────────────────────────────────────── */
  function cmp(a, b) { a = a == null ? '' : String(a); b = b == null ? '' : String(b); return a < b ? -1 : a > b ? 1 : 0; }
  function sortRecords(list) {
    return list.slice().sort(function (a, b) { return cmp(a.date, b.date) || cmp(a.die, b.die) || cmp(a.model, b.model) || cmp(a.key, b.key); });
  }
  function mergeRecords(existing, incoming) {
    var byKey = {}, stat = { added: 0, updated: 0 };
    (existing || []).forEach(function (r) { byKey[r.key || baseKey(r) + '#1'] = r; });
    (incoming || []).forEach(function (r) {
      var k = r.key || baseKey(r) + '#1';
      if (byKey[k]) stat.updated++; else stat.added++;
      byKey[k] = r;
    });
    return { records: sortRecords(Object.keys(byKey).map(function (k) { return byKey[k]; })), stat: stat };
  }
  function monthsOf(records) {
    var s = {}; (records || []).forEach(function (r) { s[monthOf(r.date)] = 1; });
    return Object.keys(s).sort();
  }
  /* 첫 달 ~ 마지막 달 사이에 작업 기록이 없는 달 */
  function monthCoverage(records) {
    var ms = monthsOf(records), out = { first: ms[0] || '', last: ms[ms.length - 1] || '', months: ms, missing: [] };
    if (!ms.length) return out;
    var have = {}; ms.forEach(function (m) { have[m] = 1; });
    var y = +out.first.slice(0, 4), mo = +out.first.slice(5, 7);
    while (true) {
      var k = y + '-' + pad(mo);
      if (k > out.last) break;
      if (!have[k]) out.missing.push(k);
      mo++; if (mo > 12) { mo = 1; y++; }
    }
    return out;
  }
  function dateRange(records) {
    var a = '', b = '';
    (records || []).forEach(function (r) { if (!a || r.date < a) a = r.date; if (!b || r.date > b) b = r.date; });
    return { from: a, to: b };
  }

  /* ── 6. 설정 ──────────────────────────────────────────────── */
  var DEFAULT_SETTINGS = {
    target: 0.02,       // 목표 불량률 2% (제출 글)
    current: 0.05,      // 현재 불량률 약 5% (제출 글 — 비교용 표시)
    mainMin: null, mainMax: null,       // Hydro press Main 압력 정상 범위 (기획서 10장 ⑤ 확인 전이라 비워 둠)
    cushionMin: null, cushionMax: null, // 쿠션 압력 정상 범위
    minQty: 3000,       // 이보다 작업량이 적은 묶음은 「표본 적음」
    minRows: 5,         // 이보다 행이 적은 묶음은 「표본 적음」
    minLift: 1.3,       // 「불량이 몰린 조건」: 전체 불량률의 몇 배 이상
    pairGain: 1.2,      // 두 조건 조합은 각 조건 하나일 때보다 이만큼 더 높아야 따로 보여 줌
    windowDays: 30      // 금형 보수 전후 비교 기간(일)
  };
  function numOr(v, d, lo, hi) {
    if (v === '' || v == null) return d;
    var n = Number(v);
    if (!isFinite(n) || n < lo || n > hi) return d;
    return n;
  }
  function cleanSettings(p) {
    p = p || {};
    var s = {};
    s.target = numOr(p.target, DEFAULT_SETTINGS.target, 0.0001, 0.5);
    s.current = numOr(p.current, DEFAULT_SETTINGS.current, 0, 1);
    ['mainMin', 'mainMax', 'cushionMin', 'cushionMax'].forEach(function (k) { s[k] = numOr(p[k], null, 0, 1e6); });
    if (s.mainMin != null && s.mainMax != null && s.mainMin > s.mainMax) { var t = s.mainMin; s.mainMin = s.mainMax; s.mainMax = t; }
    if (s.cushionMin != null && s.cushionMax != null && s.cushionMin > s.cushionMax) { var t2 = s.cushionMin; s.cushionMin = s.cushionMax; s.cushionMax = t2; }
    s.minQty = Math.round(numOr(p.minQty, DEFAULT_SETTINGS.minQty, 0, 1e9));
    s.minRows = Math.round(numOr(p.minRows, DEFAULT_SETTINGS.minRows, 1, 1e6));
    s.minLift = numOr(p.minLift, DEFAULT_SETTINGS.minLift, 1, 10);
    s.pairGain = numOr(p.pairGain, DEFAULT_SETTINGS.pairGain, 1, 5);
    s.windowDays = Math.round(numOr(p.windowDays, DEFAULT_SETTINGS.windowDays, 3, 365));
    return s;
  }

  /* ── 7. 금형 보수·교체 이력 ─────────────────────────────────── */
  function cleanMaint(list) {
    var out = [], seen = {};
    (list || []).forEach(function (m) {
      if (!m) return;
      var die = str(m.die), date = toISO(m.date);
      if (!die || !date) return;
      var kind = /교체/.test(str(m.kind)) ? '교체' : '보수';
      var k = die + '|' + date + '|' + kind;
      if (seen[k]) return; seen[k] = 1;
      out.push({ die: die, date: date, kind: kind, note: str(m.note) });
    });
    return out.sort(function (a, b) { return cmp(a.die, b.die) || cmp(a.date, b.date); });
  }
  /* 엑셀에서 복사한 글(탭·쉼표로 나뉜 줄: 금형 · 날짜 · 구분 · 내용) → 이력. 머리 줄은 건너뜁니다 */
  function parseMaintText(text) {
    var items = [], bad = [];
    String(text || '').split(/\r?\n/).forEach(function (line, i) {
      if (!line.trim()) return;
      var p = line.indexOf('\t') >= 0 ? line.split('\t') : line.split(',');
      p = p.map(str);
      if (i === 0 && !toISO(p[1]) && /금형|die/i.test(p[0])) return;
      var date = toISO(p[1]);
      if (!p[0] || !date) { bad.push(i + 1); return; }
      items.push({ die: p[0], date: date, kind: p[2] || '보수', note: p.slice(3).join(' ') });
    });
    return { items: cleanMaint(items), bad: bad };
  }
  /* 행마다 「마지막 보수·교체 뒤 그 금형으로 앞서 작업한 수량」을 셉니다(그 행 작업량은 넣지 않음).
     보수일에 한 작업은 보수 뒤 작업으로 봅니다. 그 금형의 보수 이력보다 앞선 행은 시작점을 몰라 known=false. */
  function computeWear(records, maint) {
    var byDie = {}, ev = {}, out = {};
    sortRecords(records || []).forEach(function (r) { (byDie[r.die] || (byDie[r.die] = [])).push(r); });
    cleanMaint(maint).forEach(function (m) { (ev[m.die] || (ev[m.die] = [])).push(m); });
    Object.keys(byDie).forEach(function (die) {
      var es = ev[die] || [], i = 0, cum = 0, known = false, last = null;
      byDie[die].forEach(function (r) {
        while (i < es.length && es[i].date <= r.date) { cum = 0; known = true; last = es[i]; i++; }
        out[r.key] = { cum: cum, known: known, since: last ? last.date : '', kind: last ? last.kind : '' };
        cum += r.qty;
      });
    });
    return out;
  }

  /* ── 7-2. 압력 작업표준(모델·금형별 기준 값과 허용 폭) — 3차(기획서 11.3) ──
     모델(또는 금형)마다 Main·쿠션 압력의 기준과 허용 폭(±)이 다르다는 답을 받아, 전체 하나의 정상 범위 대신 행마다 「실제 − 기준」을 봅니다.
     기준을 찾는 순서: ① 작업일보의 기준 칸(「Main 압력 기준」 등)이 있으면 그 값 ② 설정의 작업표준 표(작업모델 또는 금형으로 찾기).
     허용 폭을 비워 두면 0 — 기준보다 조금만 낮아도 「저하」로 셉니다. */
  function cleanStds(o) {
    o = o || {};
    var by = o.by === 'die' ? 'die' : 'model', rows = [], seen = {};
    function n(v, neg) { var x = toNum(v); return x == null || isNaN(x) || (!neg && x < 0) ? null : x; }
    (Array.isArray(o.rows) ? o.rows : []).forEach(function (r) {
      if (!r) return;
      var key = str(r.key);
      if (!key || seen[key]) return; seen[key] = 1;
      var x = { key: key, main: n(r.main), mainTol: n(r.mainTol), cushion: n(r.cushion), cushionTol: n(r.cushionTol) };
      if (x.main == null && x.cushion == null) return;
      rows.push(x);
    });
    rows.sort(function (a, b) { return cmp(a.key, b.key); });
    return { by: by, rows: rows };
  }
  /* 엑셀에서 복사한 글: 모델(또는 금형) · Main 기준 · Main 허용(±) · 쿠션 기준 · 쿠션 허용(±). 머리 줄은 건너뜁니다 */
  function parseStdText(text, by) {
    var rows = [], bad = [];
    String(text || '').split(/\r?\n/).forEach(function (line, i) {
      if (!line.trim()) return;
      var p = (line.indexOf('\t') >= 0 ? line.split('\t') : line.split(',')).map(str);
      if (i === 0 && toNum(p[1]) != null && isNaN(toNum(p[1]))) return; // 머리 줄
      var main = toNum(p[1]), cushion = toNum(p[3]);
      if (!p[0] || ((main == null || isNaN(main)) && (cushion == null || isNaN(cushion)))) { bad.push(i + 1); return; }
      rows.push({ key: p[0], main: p[1], mainTol: p[2], cushion: p[3], cushionTol: p[4] });
    });
    return { stds: cleanStds({ by: by, rows: rows }), bad: bad };
  }
  /* 작업일보에서 기준 초안 만들기 — 모델(금형)마다 압력 값의 가운데값(중앙값). 지금 작업일보에 기준(설정) 값을 적고 있다면 그 값이 그대로 나옵니다 */
  function median(a) { var v = a.slice().sort(function (x, y) { return x - y; }), m = v.length >> 1; return v.length ? (v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2) : null; }
  function suggestStds(records, by, keep) {
    by = by === 'die' ? 'die' : 'model';
    var g = {}, old = {};
    (keep && keep.rows || []).forEach(function (r) { old[r.key] = r; });
    (records || []).forEach(function (r) {
      var k = r[by]; if (!k) return;
      var x = g[k] || (g[k] = { main: [], cushion: [] });
      if (r.main != null) x.main.push(r.main);
      if (r.cushion != null) x.cushion.push(r.cushion);
    });
    return cleanStds({ by: by, rows: Object.keys(g).map(function (k) {
      var o = old[k] || {};
      return { key: k, main: g[k].main.length ? round(median(g[k].main), 1) : null, mainTol: o.mainTol, cushion: g[k].cushion.length ? round(median(g[k].cushion), 1) : null, cushionTol: o.cushionTol };
    }) });
  }
  /* 행의 기준 → { main, mainTol, cushion, cushionTol, from:'칸'|'표' } (없는 값은 null) */
  function stdOf(stds, r) {
    var t = null;
    if (stds && stds.rows) for (var i = 0; i < stds.rows.length; i++) if (stds.rows[i].key === r[stds.by || 'model']) { t = stds.rows[i]; break; }
    var o = { main: null, mainTol: 0, cushion: null, cushionTol: 0, mainFrom: '', cushionFrom: '' };
    ['main', 'cushion'].forEach(function (k) {
      if (r[k + 'Spec'] != null) { o[k] = r[k + 'Spec']; o[k + 'From'] = '칸'; }
      else if (t && t[k] != null) { o[k] = t[k]; o[k + 'From'] = '표'; }
      if (t && t[k + 'Tol'] != null) o[k + 'Tol'] = t[k + 'Tol'];
    });
    return o;
  }
  /* 행마다 실제 − 기준. 압력 값이나 기준이 없으면 null */
  function devOf(stds, r) {
    var s = stdOf(stds, r), o = { std: s };
    ['main', 'cushion'].forEach(function (k) {
      o[k + 'Dev'] = r[k] != null && s[k] != null ? round(r[k] - s[k], 6) : null;
      o[k + 'Low'] = o[k + 'Dev'] != null && o[k + 'Dev'] < -s[k + 'Tol'];
      o[k + 'High'] = o[k + 'Dev'] != null && o[k + 'Dev'] > s[k + 'Tol'];
    });
    return o;
  }

  /* ── 7-3. 설비(Press) 수선 이력 — 금형 보수 이력과 같은 모양(3차, 기획서 11.4) ──
     설비 · 일자 · 구분(수선·점검·부품교체) · 내용. 작업일보에 설비(호기) 칸이 있으면 같은 설비의 작업만, 없으면(프레스 한 대) 모든 작업으로 전후를 봅니다. */
  var REPAIR_KINDS = ['수선', '부품교체', '점검'];
  function cleanRepairs(list) {
    var out = [], seen = {};
    (list || []).forEach(function (m) {
      if (!m) return;
      var equip = str(m.equip) || '프레스', date = toISO(m.date);
      if (!date) return;
      var k0 = str(m.kind), kind = /교체/.test(k0) ? '부품교체' : /점검/.test(k0) ? '점검' : '수선';
      var k = equip + '|' + date + '|' + kind;
      if (seen[k]) return; seen[k] = 1;
      out.push({ equip: equip, date: date, kind: kind, note: str(m.note) });
    });
    return out.sort(function (a, b) { return cmp(a.date, b.date) || cmp(a.equip, b.equip); });
  }
  function parseRepairText(text) {
    var items = [], bad = [];
    String(text || '').split(/\r?\n/).forEach(function (line, i) {
      if (!line.trim()) return;
      var p = (line.indexOf('\t') >= 0 ? line.split('\t') : line.split(',')).map(str);
      if (i === 0 && !toISO(p[1]) && /설비|프레스|press|장비/i.test(p[0])) return;
      var date = toISO(p[1]);
      if (!date) { bad.push(i + 1); return; }
      items.push({ equip: p[0], date: date, kind: p[2] || '수선', note: p.slice(3).join(' ') });
    });
    return { items: cleanRepairs(items), bad: bad };
  }
  /* 엑셀 시트(행 배열) → 붙여 넣기 글과 같은 줄들. 날짜 칸이 Date 여도 읽게 글자로 바꿉니다 */
  function aoaToText(aoa) {
    return (aoa || []).map(function (row) {
      return (row || []).map(function (v) { return v instanceof Date || Object.prototype.toString.call(v) === '[object Date]' ? toISO(v) : str(v); }).join('\t');
    }).join('\n');
  }

  /* ── 8. 합계와 비율 ────────────────────────────────────────── */
  /* 불량률은 늘 「불량 합계 ÷ 작업수량 합계」(작업량 가중). 행마다 낸 불량률의 평균이 아닙니다. */
  function countOf(r, metric) { return !metric || metric === 'all' ? r.defect : (r.types && r.types[metric]) || 0; }
  function emptyAgg() { return { rows: 0, qty: 0, cnt: 0, defect: 0, types: {}, days: 0 }; }
  function aggregate(list, metric) {
    var a = emptyAgg(), days = {};
    (list || []).forEach(function (r) {
      a.rows++; a.qty += r.qty; a.defect += r.defect; a.cnt += countOf(r, metric); days[r.date] = 1;
      Object.keys(r.types || {}).forEach(function (t) { a.types[t] = (a.types[t] || 0) + r.types[t]; });
    });
    a.days = Object.keys(days).length;
    a.rate = a.qty > 0 ? a.cnt / a.qty : null;
    a.defectRate = a.qty > 0 ? a.defect / a.qty : null;
    var typed = 0; Object.keys(a.types).forEach(function (t) { typed += a.types[t]; });
    a.untyped = Math.max(0, a.defect - typed);
    return a;
  }
  function typeList(records) {
    var s = {}; (records || []).forEach(function (r) { Object.keys(r.types || {}).forEach(function (t) { s[t] = 1; }); });
    var order = DEFECT_TYPES.map(function (d) { return d.id; });
    return Object.keys(s).sort(function (a, b) {
      var ia = order.indexOf(a), ib = order.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || cmp(a, b);
    });
  }
  function metricLabel(metric) { return !metric || metric === 'all' ? '전체 불량' : metric + ' 불량'; }
  /* 윌슨 신뢰구간(95%) — 작업량이 적은 묶음의 불량률이 우연히 높게 나온 것을 거르는 데 씁니다 */
  function wilson(k, n, z) {
    if (!(n > 0)) return { lo: null, hi: null };
    z = z || 1.96;
    var p = k / n, d = 1 + z * z / n, c = p + z * z / (2 * n), w = z * Math.sqrt(p * (1 - p) / n + z * z / (4 * n * n));
    return { lo: Math.max(0, (c - w) / d), hi: Math.min(1, (c + w) / d) };
  }

  /* ── 9. 기간·파생값을 붙인 분석 묶음 ────────────────────────── */
  function niceStep(x) {
    if (!(x > 0)) return 1;
    var p = Math.pow(10, Math.floor(Math.log10(x))), n = x / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
  }
  /* 숫자 값을 같은 폭 구간으로 나눕니다(목표 구간 수 target, 기본 5). 값이 하나뿐이면 구간 하나 */
  function makeBins(values, target) {
    var v = (values || []).filter(function (x) { return x != null && isFinite(x); });
    if (!v.length) return null;
    var min = Math.min.apply(null, v), max = Math.max.apply(null, v);
    var step = niceStep((max - min) / (target || 5) || Math.abs(max) / 10 || 1);
    var start = Math.floor(min / step) * step;
    return { start: round(start, 6), step: step, min: min, max: max };
  }
  function binOf(bins, v) {
    if (!bins || v == null || !isFinite(v)) return null;
    var i = Math.floor((v - bins.start) / bins.step + 1e-9);
    return round(bins.start + i * bins.step, 6);
  }
  function fmtNum(n) { return n == null ? '-' : Number(round(n, 2)).toLocaleString('ko-KR'); }
  function binLabel(bins, lo, unit) { return fmtNum(lo) + ' ~ ' + fmtNum(round(lo + bins.step, 6)) + (unit ? ' ' + unit : ''); }

  /* 「실제 − 기준」 구간 — 같은 폭(시작이 폭의 배수라 0 이 늘 구간 경계: 「0 ~ 폭」 = 기준과 같거나 조금 높음) */
  function devBins(values) { return makeBins(values, 8); }
  /* 분석에 쓰는 묶음. opts: { from, to, settings, maint, stds, repairs } */
  function buildContext(records, opts) {
    opts = opts || {};
    var S = cleanSettings(opts.settings);
    var maint = cleanMaint(opts.maint);
    var stds = cleanStds(opts.stds), repairs = cleanRepairs(opts.repairs);
    var all = records || [];
    var wear = computeWear(all, maint);   // 누적 작업량은 기간 밖 작업까지 세야 맞습니다
    var list = all.filter(function (r) { return (!opts.from || r.date >= opts.from) && (!opts.to || r.date <= opts.to); });
    var recs = list.map(function (r) {
      var w = wear[r.key] || { cum: 0, known: false, since: '' };
      var o = {}; Object.keys(r).forEach(function (k) { o[k] = r[k]; });
      o.season = seasonOf(r.date); o.month = monthOf(r.date);
      o.wear = w.known ? w.cum : null; o.wearSince = w.since;
      var dv = devOf(stds, r);
      o.mainStd = dv.std.main; o.mainTol = dv.std.mainTol; o.mainDev = dv.mainDev; o.mainLow = dv.mainLow;
      o.cushionStd = dv.std.cushion; o.cushionTol = dv.std.cushionTol; o.cushionDev = dv.cushionDev; o.cushionLow = dv.cushionLow;
      return o;
    });
    var bins = {
      main: makeBins(recs.map(function (r) { return r.main; }), 10),
      cushion: makeBins(recs.map(function (r) { return r.cushion; }), 10),
      wear: makeBins(recs.map(function (r) { return r.wear; })),
      mainDev: devBins(recs.map(function (r) { return r.mainDev; })),
      cushionDev: devBins(recs.map(function (r) { return r.cushionDev; }))
    };
    var rng = dateRange(recs);
    return { recs: recs, all: all, S: S, maint: maint, stds: stds, repairs: repairs, bins: bins, from: opts.from || rng.from, to: opts.to || rng.to, types: typeList(recs), wearMap: wear };
  }

  /* ── 10. 보는 기준(차원) — 4M + 계절 ─────────────────────────
     4M 묶음은 보기 위한 분류입니다(회사 기준이 있으면 바꿀 수 있음).
     Man = 작업자 · Machine = 금형·금형 누적 작업량·Main 압력·쿠션 압력 · Material = 소재 메이커·규격·로트 · Method = 작업모델 */
  var M4 = [
    { id: 'Man', label: 'Man (사람)', dims: ['worker'] },
    { id: 'Machine', label: 'Machine (설비·금형)', dims: ['die', 'wear', 'main', 'mainRange', 'mainStd', 'mainDev', 'cushion', 'cushionRange', 'cushionStd', 'cushionDev'] },
    { id: 'Material', label: 'Material (소재)', dims: ['maker', 'spec', 'lot'] },
    { id: 'Method', label: 'Method (작업 방법·모델)', dims: ['model'] },
    { id: 'Env', label: '계절 (시기)', dims: ['season', 'month'] }
  ];
  var M4_SHORT = { Man: '사람', Machine: '설비·금형', Material: '소재', Method: '방법', Env: '계절' };
  function m4Text(ids) { return (ids || []).map(function (i) { return M4_SHORT[i] || i; }).join('·'); }
  var DIMS = {
    season: { label: '계절', cat: false },
    month: { label: '월', cat: false },
    worker: { label: '작업자', cat: true, mask: 'worker' },
    die: { label: '금형', cat: true, mask: 'die' },
    wear: { label: '금형 누적 작업량(보수 후)', num: 'wear', unit: '개' },
    main: { label: 'Main 압력 구간', num: 'main' },
    mainRange: { label: 'Main 압력 정상 범위', range: 'main' },
    cushion: { label: '쿠션 압력 구간', num: 'cushion' },
    cushionRange: { label: '쿠션 압력 정상 범위', range: 'cushion' },
    mainStd: { label: 'Main 압력 작업표준 대비', std: 'main' },
    cushionStd: { label: '쿠션 압력 작업표준 대비', std: 'cushion' },
    mainDev: { label: 'Main 압력 − 기준', num: 'mainDev' },
    cushionDev: { label: '쿠션 압력 − 기준', num: 'cushionDev' },
    model: { label: '작업모델', cat: true, mask: 'model' },
    maker: { label: '소재 메이커', cat: true, mask: 'maker' },
    spec: { label: '소재규격', cat: true },
    lot: { label: '소재 로트', cat: true, mask: 'lot' }
  };
  /* 행의 그 기준 값 → { key, label, order } 또는 null(값 없음 — 그 기준에서 빠짐) */
  function dimValue(ctx, dim, r) {
    var D = DIMS[dim];
    if (dim === 'season') return { key: r.season, label: r.season, order: SEASONS.indexOf(r.season) };
    if (dim === 'month') return { key: r.month, label: monthLabel(r.month), order: r.month };
    if (D.cat) { var v = r[dim]; return v ? { key: v, label: v, order: v } : null; }
    if (D.num) {
      var x = r[D.num], b = ctx.bins[D.num];
      if (x == null || !b) return null;
      var lo = binOf(b, x);
      return { key: String(lo), label: binLabel(b, lo, D.unit), order: lo, lo: lo, hi: round(lo + b.step, 6) };
    }
    if (D.std) {
      var dv = r[D.std + 'Dev'], tol = r[D.std + 'Tol'] || 0;
      if (dv == null) return null;
      if (dv < -tol) return { key: 'low', label: '기준보다 허용 넘게 낮음 (저하)', order: 0 };
      if (dv > tol) return { key: 'high', label: '기준보다 허용 넘게 높음', order: 2 };
      return { key: 'ok', label: '허용 안 (기준 ± 허용)', order: 1 };
    }
    if (D.range) {
      var val = r[D.range], mn = ctx.S[D.range + 'Min'], mx = ctx.S[D.range + 'Max'];
      if (val == null || (mn == null && mx == null)) return null;
      if (mn != null && val < mn) return { key: 'below', label: '정상 범위 아래 (' + fmtNum(mn) + ' 미만)', order: 0 };
      if (mx != null && val > mx) return { key: 'above', label: '정상 범위 위 (' + fmtNum(mx) + ' 초과)', order: 2 };
      return { key: 'in', label: '정상 범위 안', order: 1 };
    }
    return null;
  }
  function dimAvailable(ctx, dim) {
    if (DIMS[dim].range && ctx.S[DIMS[dim].range + 'Min'] == null && ctx.S[DIMS[dim].range + 'Max'] == null) return false;
    for (var i = 0; i < ctx.recs.length; i++) if (dimValue(ctx, dim, ctx.recs[i])) return true;
    return false;
  }
  function decorate(a, overall, S) {
    a.lift = overall.rate > 0 && a.rate != null ? a.rate / overall.rate : null;
    a.expected = overall.rate != null ? a.qty * overall.rate : null;
    a.excess = a.expected != null ? a.cnt - a.expected : null;
    a.share = overall.cnt > 0 ? a.cnt / overall.cnt : null;
    var w = wilson(a.cnt, a.qty); a.lo = w.lo; a.hi = w.hi;
    a.small = a.qty < S.minQty || a.rows < S.minRows;
    return a;
  }
  /* 한 기준으로 묶은 표. sort: 'order'(기준 순) · 'rate' · 'excess' */
  function groupBy(ctx, dim, metric, sort) {
    var m = {}, info = {}, missing = [];
    ctx.recs.forEach(function (r) {
      var v = dimValue(ctx, dim, r);
      if (!v) { missing.push(r); return; }
      (m[v.key] || (m[v.key] = [])).push(r); info[v.key] = v;
    });
    var overall = aggregate(ctx.recs, metric);
    var rows = Object.keys(m).map(function (k) {
      var a = decorate(aggregate(m[k], metric), overall, ctx.S);
      a.key = k; a.label = info[k].label; a.order = info[k].order; a.lo_ = info[k].lo; a.hi_ = info[k].hi;
      return a;
    });
    rows.sort(function (a, b) {
      if (sort === 'rate') return (b.rate - a.rate) || cmp(a.label, b.label);
      if (sort === 'excess') return (b.excess - a.excess) || cmp(a.label, b.label);
      return typeof a.order === 'number' && typeof b.order === 'number' ? a.order - b.order : cmp(a.order, b.order);
    });
    return { dim: dim, label: DIMS[dim].label, metric: metric || 'all', overall: overall, rows: rows, missing: aggregate(missing, metric) };
  }
  /* 두 기준 교차표 (예: 계절 × 소재 메이커) */
  function crossTab(ctx, dimA, dimB, metric) {
    var cells = {}, aInfo = {}, bInfo = {};
    ctx.recs.forEach(function (r) {
      var a = dimValue(ctx, dimA, r), b = dimValue(ctx, dimB, r);
      if (!a || !b) return;
      aInfo[a.key] = a; bInfo[b.key] = b;
      (cells[a.key + '\u0001' + b.key] || (cells[a.key + '\u0001' + b.key] = [])).push(r);
    });
    var overall = aggregate(ctx.recs, metric);
    function ord(info) { return Object.keys(info).sort(function (x, y) { var p = info[x].order, q = info[y].order; return typeof p === 'number' ? p - q : cmp(p, q); }); }
    var A = ord(aInfo), B = ord(bInfo);
    return {
      dimA: dimA, dimB: dimB, overall: overall,
      rows: A.map(function (ak) { return { key: ak, label: aInfo[ak].label }; }),
      cols: B.map(function (bk) { return { key: bk, label: bInfo[bk].label }; }),
      cell: function (ak, bk) { var l = cells[ak + '\u0001' + bk]; return l ? decorate(aggregate(l, metric), overall, ctx.S) : null; }
    };
  }

  /* ── 11. 일별·월별 추이 ─────────────────────────────────────── */
  function trend(ctx, unit, metric, month) {
    var m = {};
    ctx.recs.forEach(function (r) {
      if (unit === 'day' && month && r.month !== month) return;
      var k = unit === 'day' ? r.date : r.month;
      (m[k] || (m[k] = [])).push(r);
    });
    return Object.keys(m).sort().map(function (k) {
      var a = aggregate(m[k], metric);
      a.key = k; a.label = unit === 'day' ? k.slice(5).replace('-', '/') : k.slice(2).replace('-', '.');
      a.overTarget = a.rate != null && a.rate > ctx.S.target;
      return a;
    });
  }

  /* ── 12. 불량이 몰린 조건 ─────────────────────────────────────
     한 기준(예: 쿠션 압력 65~70)과 두 기준 조합(예: 겨울 × 소재 메이커 B)마다 불량률을 내서,
       ① 표본이 충분하고(작업량 minQty·행 minRows 이상)
       ② 불량률이 전체의 minLift 배 이상이며
       ③ 작업량을 생각해도(윌슨 95% 하한) 전체 불량률보다 높은 것만 고릅니다.
     조합은 두 조건 각각보다 pairGain 배 이상 높고 하한도 둘보다 높을 때만 따로 보여 줍니다(한 조건으로 설명되는 조합은 빼기).
     「기대보다 많은 불량」= 그 조건의 불량 − 작업량 × 전체 불량률. 이 수가 큰 순서로 늘어놓습니다.
     통계적 원인 판정이 아니라 「먼저 볼 곳」을 고르는 거름망입니다. */
  var HOT_SINGLE = ['season', 'worker', 'die', 'wear', 'main', 'mainRange', 'mainStd', 'cushion', 'cushionRange', 'cushionStd', 'model', 'maker', 'spec', 'lot'];
  var HOT_PAIR = ['season', 'worker', 'die', 'wear', 'main', 'cushion', 'model', 'maker', 'spec'];
  function m4Of(dim) { for (var i = 0; i < M4.length; i++) if (M4[i].dims.indexOf(dim) >= 0) return M4[i].id; return ''; }
  function findHotspots(ctx, metric, limit) {
    var S = ctx.S, overall = aggregate(ctx.recs, metric), out = [];
    if (!(overall.rate > 0)) return { overall: overall, list: [], total: 0 };
    var dims = HOT_SINGLE.filter(function (d) { return dimAvailable(ctx, d); });
    var singleRate = {};
    function hot(a) { return !a.small && a.lift != null && a.lift >= S.minLift && a.lo != null && a.lo > overall.rate; }
    dims.forEach(function (d) {
      groupBy(ctx, d, metric).rows.forEach(function (a) {
        singleRate[d + '=' + a.key] = a.rate;
        if (hot(a)) out.push(hotItem([{ dim: d, key: a.key, label: a.label }], a));
      });
    });
    var pd = HOT_PAIR.filter(function (d) { return dims.indexOf(d) >= 0; });
    for (var i = 0; i < pd.length; i++) for (var j = i + 1; j < pd.length; j++) {
      var ct = crossTab(ctx, pd[i], pd[j], metric);
      ct.rows.forEach(function (ra) {
        ct.cols.forEach(function (cb) {
          var a = ct.cell(ra.key, cb.key);
          if (!a || !hot(a)) return;
          var p1 = singleRate[pd[i] + '=' + ra.key], p2 = singleRate[pd[j] + '=' + cb.key];
          var parent = Math.max(p1 || 0, p2 || 0);
          // 조합만의 효과여야 함: 두 조건 각각보다 pairGain 배 이상 높고, 95% 하한도 둘 중 높은 쪽보다 높아야
          if (!(a.rate >= parent * S.pairGain && a.lo > parent)) return;
          out.push(hotItem([{ dim: pd[i], key: ra.key, label: ra.label }, { dim: pd[j], key: cb.key, label: cb.label }], a));
        });
      });
    }
    out.sort(function (a, b) { return (b.agg.excess - a.agg.excess) || (b.agg.rate - a.agg.rate) || (a.parts.length - b.parts.length); });
    // 금형·모델·규격이 1:1 이면 같은 행 묶음이 이름만 바꿔 여러 번 나옵니다 → 행 묶음이 같으면 처음 것만
    var seen = {}, list = [];
    out.forEach(function (h) {
      var sig = h.agg.rows + '|' + h.agg.qty + '|' + h.agg.cnt + '|' + h.parts.length;
      if (seen[sig]) { seen[sig].same.push(condText(h.parts)); return; }
      h.same = []; seen[sig] = h; list.push(h);
    });
    var total = list.length;
    list = list.slice(0, limit || 12);
    list.forEach(function (h) { h.overlap = overlapOf(ctx, h); });
    return { overall: overall, list: list, total: total };
  }
  function hotItem(parts, a) {
    return { parts: parts, agg: a, m4: uniq(parts.map(function (p) { return m4Of(p.dim); })) };
  }
  /* 겹침 — 이 조건의 작업이 한 금형·메이커·작업자·계절에 몰려 있으면(작업량 70% 이상) 알려 줍니다.
     예: 「Main 압력 240~250」 작업이 거의 다 한 금형이면, 압력 때문인지 금형 때문인지 이 표만으로는 가를 수 없습니다. */
  function overlapOf(ctx, h) {
    var rows = ctx.recs.filter(function (r) {
      return h.parts.every(function (p) { var v = dimValue(ctx, p.dim, r); return v && v.key === p.key; });
    });
    var total = rows.reduce(function (s, r) { return s + r.qty; }, 0), out = [];
    var allQty = ctx.recs.reduce(function (s, r) { return s + r.qty; }, 0);
    ['die', 'maker', 'worker', 'season'].forEach(function (d) {
      if (h.parts.some(function (p) { return p.dim === d; })) return;
      var m = {};
      rows.forEach(function (r) { var v = dimValue(ctx, d, r); if (v) m[v.label] = (m[v.label] || 0) + r.qty; });
      Object.keys(m).forEach(function (k) {
        var share = total ? m[k] / total : 0;
        // 전체에서도 원래 그만큼 큰 값(예: 전체의 80%가 소재사A)이면 겹침이 아닙니다
        var base = ctx.recs.filter(function (r) { var v = dimValue(ctx, d, r); return v && v.label === k; }).reduce(function (s, r) { return s + r.qty; }, 0) / (allQty || 1);
        if (share >= 0.7 && share - base >= 0.3) out.push({ dim: d, label: k, share: share });
      });
    });
    return out;
  }
  function overlapText(h, A) {
    return (h.overlap || []).map(function (o) {
      var D = DIMS[o.dim];
      return '이 조건 작업의 ' + round(o.share * 100, 0) + '%가 ' + (D.mask && A ? A(D.mask, o.label) : D.label + ' ' + o.label);
    }).join(', ');
  }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  /* 조건 이름 — A(dimMask, value) 로 값을 가릴 수 있습니다(AI 프롬프트) */
  function condText(parts, A) {
    return parts.map(function (p) {
      var D = DIMS[p.dim];
      if (D.mask && A) return A(D.mask, p.label); // 별칭(「금형3」「메이커2」)은 그 자체로 무엇인지 드러남
      return D.label + ' ' + p.label;
    }).join(' × ');
  }
  function pctText(v, d) { return v == null ? '-' : round(v * 100, d == null ? 1 : d) + '%'; }
  function hotEvidence(h, overall) {
    var a = h.agg;
    return '작업 ' + fmtNum(a.qty) + '개 중 불량 ' + fmtNum(a.cnt) + '개(' + pctText(a.rate, 2) + ') — 전체 ' + pctText(overall.rate, 2) + '의 ' + round(a.lift, 2) +
      '배, 전체 비율이면 ' + fmtNum(round(a.expected, 0)) + '개였을 것 → ' + fmtNum(round(a.excess, 0)) + '개 더 많음 (' + a.rows + '행, ' + a.days + '일)';
  }

  /* ── 13. 금형 보수 전후 · 금형 상태 ─────────────────────────── */
  function maintEffects(ctx, metric) {
    var W = ctx.S.windowDays, out = [];
    ctx.maint.forEach(function (m) {
      var rows = ctx.all.filter(function (r) { return r.die === m.die; });
      if (!rows.length) return;
      var from = addDays(m.date, -W), to = addDays(m.date, W);
      var before = aggregate(rows.filter(function (r) { return r.date >= from && r.date < m.date; }), metric);
      var after = aggregate(rows.filter(function (r) { return r.date >= m.date && r.date < to; }), metric);
      if (!before.rows && !after.rows) return;
      out.push({
        die: m.die, date: m.date, kind: m.kind, note: m.note, before: before, after: after,
        change: before.rate != null && after.rate != null ? after.rate - before.rate : null,
        enough: before.qty >= ctx.S.minQty && after.qty >= ctx.S.minQty
      });
    });
    return out.sort(function (a, b) { return cmp(a.date, b.date) || cmp(a.die, b.die); });
  }
  /* 금형마다 — 마지막 보수·교체 뒤 누적 작업량(데이터 끝 기준), 최근 30일 불량률,
     그리고 과거 이 금형에서 보수 직후보다 불량률이 minLift 배 이상 높아진 누적 작업량 구간(참고). 예측이 아닙니다. */
  function dieStatus(ctx, metric) {
    var S = ctx.S, end = dateRange(ctx.all).to, out = [];
    var byDie = {};
    ctx.all.forEach(function (r) { (byDie[r.die] || (byDie[r.die] = [])).push(r); });
    Object.keys(byDie).sort(cmp).forEach(function (die) {
      var rows = byDie[die];
      var wearBins = makeBins(rows.map(function (r) { var w = ctx.wearMap[r.key]; return w && w.known ? w.cum : null; }), 6);
      var ev = ctx.maint.filter(function (m) { return m.die === die; });
      var last = ev.length ? ev[ev.length - 1] : null;
      var cumNow = last ? rows.filter(function (r) { return r.date >= last.date; }).reduce(function (s, r) { return s + r.qty; }, 0) : null;
      var recent = aggregate(rows.filter(function (r) { return r.date > addDays(end, -30); }), metric);
      var threshold = null, curve = [], fresh = null;
      if (wearBins && ev.length) {
        var m = {};
        rows.forEach(function (r) { var w = ctx.wearMap[r.key]; if (w && w.known) { var lo = binOf(wearBins, w.cum); (m[lo] || (m[lo] = [])).push(r); } });
        curve = Object.keys(m).map(Number).sort(function (a, b) { return a - b; }).map(function (lo) {
          var a = aggregate(m[lo], metric); a.lo = lo; a.hi = round(lo + wearBins.step, 6); a.label = binLabel(wearBins, lo, '개');
          a.small = a.qty < S.minQty || a.rows < S.minRows; return a;
        });
        // 보수 직후(첫 구간) 불량률보다 minLift 배 이상 높아지고, 그 뒤 구간도 계속 그만큼 높은 첫 구간
        fresh = curve.filter(function (c) { return !c.small; })[0] || null;
        var bar = fresh && fresh.rate > 0 ? fresh.rate * S.minLift : null;
        for (var i = 0; bar != null && i < curve.length; i++) {
          var rest = curve.slice(i).filter(function (c) { return !c.small; });
          if (!curve[i].small && curve[i] !== fresh && curve[i].rate >= bar && rest.every(function (c) { return c.rate >= bar; })) { threshold = curve[i].lo; break; }
        }
      }
      out.push({ die: die, lastMaint: last, events: ev.length, cumNow: cumNow, recent: recent, threshold: threshold, fresh: fresh ? fresh.rate : null, curve: curve, total: aggregate(rows, metric) });
    });
    return out;
  }

  /* ── 14. 한 번에 — 화면 ③·AI·보고서가 같은 결과를 씁니다 ──── */
  function analyze(ctx, metric) {
    metric = metric || 'all';
    var S = ctx.S, overall = aggregate(ctx.recs, metric), all = aggregate(ctx.recs, 'all');
    var gap = overall.qty > 0 ? overall.cnt - overall.qty * S.target : null;
    var dims = {};
    ['die', 'maker', 'spec', 'model', 'worker', 'lot', 'season', 'main', 'mainRange', 'mainStd', 'mainDev', 'cushion', 'cushionRange', 'cushionStd', 'cushionDev', 'wear'].forEach(function (d) {
      if (dimAvailable(ctx, d)) dims[d] = groupBy(ctx, d, metric, DIMS[d].cat ? 'excess' : 'order');
    });
    return {
      from: ctx.from, to: ctx.to, metric: metric, S: S, overall: overall, all: all, gap: gap,
      months: trend(ctx, 'month', metric), dims: dims, types: ctx.types,
      hot: findHotspots(ctx, metric), maint: maintEffects(ctx, metric), dies: dieStatus(ctx, metric),
      hasMaint: ctx.maint.length > 0,
      std: stdSummary(ctx), alerts: pressureAlerts(ctx), repairs: repairEffects(ctx, metric)
    };
  }

  /* ── 14-2. 작업표준 대비 · 압력 저하 전파 · 설비 수선 전후 (3차) ── */
  /* 모델(금형)마다 기준과 실제 — 평균·가장 낮은/높은 값·평균 차이·허용 밖 행 수.
     zeroShare: 실제 − 기준이 0 인 행의 비율. 0.9 이상이면 작업일보 압력 칸에 기준(설정) 값을 적고 있는 것으로 봅니다. */
  function stdSummary(ctx) {
    var by = ctx.stds.by, g = {}, out = { by: by, label: by === 'die' ? '금형' : '작업모델', main: [], cushion: [], rows: 0, zero: 0, noStd: {} };
    ctx.recs.forEach(function (r) {
      var k = r[by] || '(비어 있음)';
      ['main', 'cushion'].forEach(function (p) {
        if (r[p] == null) return;
        if (r[p + 'Std'] == null) { out.noStd[k] = 1; return; }
        var x = (g[p] || (g[p] = {}))[k] || (g[p][k] = { key: k, std: r[p + 'Std'], tol: r[p + 'Tol'], n: 0, qty: 0, sum: 0, min: Infinity, max: -Infinity, devSum: 0, low: 0, high: 0, lowQty: 0, lowCnt: 0 });
        x.n++; x.qty += r.qty; x.sum += r[p]; x.min = Math.min(x.min, r[p]); x.max = Math.max(x.max, r[p]); x.devSum += r[p + 'Dev'];
        if (r[p + 'Dev'] < -(r[p + 'Tol'] || 0)) { x.low++; x.lowQty += r.qty; x.lowCnt += r.defect; }
        else if (r[p + 'Dev'] > (r[p + 'Tol'] || 0)) x.high++;
        out.rows++; if (r[p + 'Dev'] === 0) out.zero++;
      });
    });
    ['main', 'cushion'].forEach(function (p) {
      out[p] = Object.keys(g[p] || {}).sort(cmp).map(function (k) {
        var x = g[p][k]; x.mean = x.sum / x.n; x.devMean = x.devSum / x.n; return x;
      });
    });
    out.noStd = Object.keys(out.noStd).sort(cmp);
    out.zeroShare = out.rows ? out.zero / out.rows : null;
    out.has = out.rows > 0;
    return out;
  }
  /* 압력 저하 전파 목록 — 실제가 「기준 − 허용」보다 낮은 행. 같은 행에서 Main·쿠션 둘 다 낮으면 두 줄 */
  function pressureAlerts(ctx) {
    var out = [];
    ctx.recs.forEach(function (r) {
      [['main', 'Main 압력'], ['cushion', '쿠션 압력']].forEach(function (p) {
        var k = p[0];
        if (r[k + 'Dev'] == null || !(r[k + 'Dev'] < -(r[k + 'Tol'] || 0))) return;
        out.push({ date: r.date, item: p[1], kind: k, die: r.die, model: r.model, press: r.press || '', worker: r.worker, std: r[k + 'Std'], tol: r[k + 'Tol'] || 0, act: r[k],
          dev: r[k + 'Dev'], over: round(-r[k + 'Dev'] - (r[k + 'Tol'] || 0), 6), qty: r.qty, defect: r.defect, rate: r.defect / r.qty, types: r.types || {} });
      });
    });
    return out.sort(function (a, b) { return cmp(a.date, b.date) || cmp(a.item, b.item) || cmp(a.model, b.model); });
  }
  /* 전파 글 — 설비보전·공무팀에 보낼 글자(메신저·메일에 붙여 넣기). 날짜가 이어진 저하는 한 묶음으로 */
  function alertText(alerts, opts) {
    opts = opts || {};
    var L = [];
    L.push('[설비보전·공무팀 전파] Hydro press 압력 저하 ' + (alerts.length ? alerts.length + '건' : '없음') + ' (' + (opts.from || '') + ' ~ ' + (opts.to || '') + ')');
    if (!alerts.length) { L.push('- 작업표준(기준 − 허용)보다 낮은 압력이 없습니다.'); return L.join('\n'); }
    var byItem = {};
    alerts.forEach(function (a) { var k = a.item + '|' + (a.press || ''); (byItem[k] || (byItem[k] = [])).push(a); });
    Object.keys(byItem).sort(cmp).forEach(function (k) {
      var list = byItem[k], dates = uniq(list.map(function (a) { return a.date; })), runs = [], cur = null;
      dates.forEach(function (d) { if (cur && daysBetween(cur.to, d) <= 3) cur.to = d; else { cur = { from: d, to: d }; runs.push(cur); } });
      var it = k.split('|');
      L.push('■ ' + it[0] + (it[1] ? ' (' + it[1] + ')' : '') + ' — ' + list.length + '건, ' + runs.map(function (x) { return x.from === x.to ? x.from : x.from + '~' + x.to; }).join(', '));
      list.forEach(function (a) {
        L.push('  - ' + a.date + ' ' + (a.model || '') + (a.die ? '(' + a.die + ')' : '') + ': 실제 ' + fmtNum(a.act) + ' / 기준 ' + fmtNum(a.std) + (a.tol ? ' ±' + fmtNum(a.tol) : '') +
          ' → ' + fmtNum(-a.dev) + ' 낮음 · 작업 ' + fmtNum(a.qty) + ' · 불량 ' + fmtNum(a.defect) + '(' + pctText(a.rate, 1) + ')');
      });
    });
    L.push('※ 점검 요청: 유압·쿠션 실린더 누유, 압력 센서·게이지, 설정값 입력 확인. 조치 후 설비 수선 이력에 적어 주세요.');
    return L.join('\n');
  }
  function alertsToAoa(alerts) {
    var aoa = [['작업일자', '항목', '설비', '작업모델', '금형', '작업자', '기준', '허용(±)', '실제', '실제 − 기준', '허용 넘은 폭', '작업수량', '불량수량', '불량률(%)']];
    alerts.forEach(function (a) { aoa.push([a.date, a.item, a.press, a.model, a.die, a.worker, a.std, a.tol, a.act, round(a.dev, 2), round(a.over, 2), a.qty, a.defect, round(a.rate * 100, 2)]); });
    return aoa;
  }
  function stdToAoa(st) {
    var aoa = [['구분', st.label, '기준', '허용(±)', '행', '실제 평균', '가장 낮음', '가장 높음', '평균 차이(실제 − 기준)', '허용 넘게 낮음(행)', '허용 넘게 높음(행)']];
    [['main', 'Main 압력'], ['cushion', '쿠션 압력']].forEach(function (p) {
      st[p[0]].forEach(function (x) { aoa.push([p[1], x.key, x.std, x.tol || 0, x.n, round(x.mean, 1), x.min, x.max, round(x.devMean, 2), x.low, x.high]); });
    });
    return aoa;
  }
  /* 설비 수선 전후 — 전후 N일 불량률(유형별 포함), 압력 저하 행 수, 평균 「실제 − 기준」 */
  function repairEffects(ctx, metric) {
    var W = ctx.S.windowDays, out = [], hasPress = ctx.all.some(function (r) { return r.press; });
    ctx.repairs.forEach(function (m) {
      var rows = hasPress ? ctx.all.filter(function (r) { return r.press === m.equip; }) : ctx.all;
      var matched = rows.length > 0;
      if (!matched) rows = hasPress ? [] : ctx.all;
      var from = addDays(m.date, -W), to = addDays(m.date, W);
      function side(list) {
        var a = aggregate(list, metric), low = 0, dsum = { main: 0, cushion: 0 }, dn = { main: 0, cushion: 0 };
        list.forEach(function (r) {
          var d = devOf(ctx.stds, r);
          if (d.mainLow || d.cushionLow) low++;
          ['main', 'cushion'].forEach(function (k) { if (d[k + 'Dev'] != null) { dsum[k] += d[k + 'Dev']; dn[k]++; } });
        });
        a.lowRows = low;
        a.mainDev = dn.main ? dsum.main / dn.main : null; a.cushionDev = dn.cushion ? dsum.cushion / dn.cushion : null;
        return a;
      }
      var before = side(rows.filter(function (r) { return r.date >= from && r.date < m.date; }));
      var after = side(rows.filter(function (r) { return r.date >= m.date && r.date < to; }));
      out.push({ equip: m.equip, date: m.date, kind: m.kind, note: m.note, before: before, after: after, matched: matched || !hasPress,
        change: before.rate != null && after.rate != null ? after.rate - before.rate : null,
        enough: before.qty >= ctx.S.minQty && after.qty >= ctx.S.minQty });
    });
    return out;
  }
  /* 같은 행 묶음 — 두 기준이 1:1 로 짝지어지면(예: 샘플은 금형마다 소재규격이 하나) 두 표의 숫자가 이름만 바꿔 똑같이 나옵니다.
     ② 화면에서 「이 표는 금형 표와 같은 묶음」이라고 알려 주는 데 씁니다. */
  function sameGrouping(ctx, a, b) {
    var ab = {}, ba = {}, n = 0;
    for (var i = 0; i < ctx.recs.length; i++) {
      var r = ctx.recs[i], x = dimValue(ctx, a, r), y = dimValue(ctx, b, r);
      if (!x && !y) continue;
      if (!x || !y) return false;
      if (ab[x.key] != null && ab[x.key] !== y.key) return false;
      if (ba[y.key] != null && ba[y.key] !== x.key) return false;
      ab[x.key] = y.key; ba[y.key] = x.key; n++;
    }
    return n > 0 && Object.keys(ab).length > 1;
  }

  /* ── 15. 값 가리기 (외부 AI 에 금형·모델·메이커·작업자·로트 이름을 보내지 않음) ──
     「금형1」「메이커2」 같은 별칭으로 바꾸고 대응표는 이 브라우저에만 둡니다. 같은 값은 늘 같은 별칭. */
  var MASK_CATS = { die: '금형', model: '모델', maker: '메이커', worker: '작업자', lot: '로트' };
  function assignAliases(map, cat, values) {
    var pre = MASK_CATS[cat], out = {}, used = 0;
    var cur = (map && map[cat]) || {};
    Object.keys(cur).forEach(function (v) {
      var m = new RegExp('^' + pre + '(\\d+)$').exec(cur[v]);
      if (m) { out[v] = cur[v]; used = Math.max(used, +m[1]); }
    });
    uniq((values || []).map(str).filter(Boolean)).sort(cmp).forEach(function (v) { if (!out[v]) out[v] = pre + (++used); });
    return out;
  }
  function assignAllAliases(map, records) {
    var o = {};
    Object.keys(MASK_CATS).forEach(function (cat) { o[cat] = assignAliases(map, cat, (records || []).map(function (r) { return r[cat]; })); });
    return o;
  }
  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&'); }
  function maskable(v) { return /[^0-9]/.test(v) ? v.length >= 2 : v.length >= 6; }
  /* AI 답 → 원래 이름. 「금형3」만 바꾸고 「금형 3개」「금형3개」처럼 수를 세는 말은 그대로 둡니다 */
  function unmaskText(text, aliases) {
    var back = {};
    Object.keys(aliases || {}).forEach(function (cat) { Object.keys(aliases[cat] || {}).forEach(function (v) { back[aliases[cat][v]] = v; }); });
    var pres = Object.keys(MASK_CATS).map(function (c) { return MASK_CATS[c]; }).join('|');
    return String(text == null ? '' : text).replace(new RegExp('(' + pres + ')(\\d+)(?![0-9개건종가명대번곳])', 'g'), function (all) { return back[all] || all; });
  }
  /* 안전망 — 프롬프트에 원래 이름이 남았는지 */
  function findLeaks(text, aliases) {
    var s = String(text || ''), out = [];
    Object.keys(aliases || {}).forEach(function (cat) {
      Object.keys(aliases[cat] || {}).forEach(function (v) {
        if (!maskable(v)) return;
        var re = new RegExp('(^|[^0-9A-Za-z])' + escRe(v) + '(?![0-9A-Za-z])');
        if (re.test(s)) out.push(v);
      });
    });
    return uniq(out);
  }

  /* ── 16. AI 프롬프트 — 계산된 사실만, 원인은 「추정」으로 ─────── */
  function buildAiPrompt(an, aliases) { return buildAiPromptInfo(an, aliases).prompt; }
  function buildAiPromptInfo(an, aliases) {
    var used = {};
    function A(cat, v) {
      if (!aliases || !aliases[cat] || !aliases[cat][v]) return v;
      (used[cat] || (used[cat] = {}))[v] = aliases[cat][v];
      return aliases[cat][v];
    }
    var S = an.S, o = an.overall, L = [];
    L.push('너는 Deep drawing(Hydro press) 공정의 품질 분석 보조야. 아래 [사실]은 작업일보 엑셀에서 도구가 계산한 값이야.');
    L.push('[사실]에 있는 숫자만 근거로 삼아줘. 숫자를 새로 만들거나 바꾸지 말아줘.');
    L.push('「겹침」이 적힌 조건은 다른 조건(금형·메이커 등)과 작업이 겹쳐 있어서, 어느 쪽 때문인지 이 숫자만으로는 가를 수 없다는 뜻이야.');
    L.push('작업일보에는 원인이 적혀 있지 않아. 그래서 원인은 모두 「추정」이라고 표시하고, 왜 그렇게 추정하는지 [사실]의 숫자를 함께 적어줘.');
    L.push('불량률은 모두 「불량 합계 ÷ 작업수량 합계」(작업량 가중)이고, 「표본 적음」 표시가 붙은 묶음은 결론에 쓰지 말아줘.');
    if (aliases) L.push('금형·모델·소재 메이커·작업자 이름은 보안상 「금형1」「메이커2」 같은 별칭으로 바꿔 두었어. 답에서도 별칭을 띄어 쓰지 말고 그대로 써줘(실제 이름을 짐작하지 말아줘).');
    L.push('');
    L.push('[사실]');
    L.push('- 기간: ' + an.from + ' ~ ' + an.to + ' · 보는 불량: ' + metricLabel(an.metric));
    L.push('- 전체: 작업수량 ' + fmtNum(o.qty) + '개, ' + metricLabel(an.metric) + ' ' + fmtNum(o.cnt) + '개, 불량률 ' + pctText(o.rate, 2) + ' (작업 ' + o.rows + '행, ' + o.days + '일)');
    L.push('- 목표 불량률 ' + pctText(S.target, 1) + ' (지금 수준이라고 적은 값 약 ' + pctText(S.current, 0) + ') → 목표에 맞추려면 이 기간에 불량을 약 ' + fmtNum(round(Math.max(0, an.gap || 0), 0)) + '개 줄여야 함');
    if (an.metric === 'all' && an.types.length) {
      var a = an.all;
      L.push('- 불량 유형별: ' + an.types.map(function (t) { return t + ' ' + fmtNum(a.types[t] || 0) + '개(' + pctText(a.qty ? (a.types[t] || 0) / a.qty : null, 2) + ')'; }).join(', ') + (a.untyped ? ', 유형 미기록 ' + fmtNum(a.untyped) + '개' : ''));
    }
    L.push('- 월별 불량률: ' + an.months.map(function (m) { return m.key + ' ' + pctText(m.rate, 2) + '(작업 ' + fmtNum(m.qty) + ')'; }).join(', '));
    function dimLine(d, n) {
      var g = an.dims[d]; if (!g) return;
      var D = DIMS[d];
      var rows = D.cat ? g.rows.slice().sort(function (x, y) { return y.excess - x.excess; }).slice(0, n || 6) : g.rows;
      L.push('- ' + D.label + '별: ' + rows.map(function (r) {
        return (D.mask ? A(D.mask, r.label) : r.label) + ' ' + pctText(r.rate, 2) + '(작업 ' + fmtNum(r.qty) + (r.small ? ', 표본 적음' : '') + ')';
      }).join(', ') + (D.cat && g.rows.length > rows.length ? ' 외 ' + (g.rows.length - rows.length) + '개' : ''));
    }
    ['season', 'die', 'maker', 'spec', 'model', 'worker', 'lot', 'main', 'mainRange', 'mainStd', 'mainDev', 'cushion', 'cushionRange', 'cushionStd', 'cushionDev', 'wear'].forEach(function (d) { dimLine(d); });
    if (an.std && an.std.has) {
      L.push('- 압력 작업표준(' + an.std.label + '별 기준 ± 허용)보다 낮은 작업: Main ' + an.alerts.filter(function (a) { return a.kind === 'main'; }).length + '건, 쿠션 ' + an.alerts.filter(function (a) { return a.kind === 'cushion'; }).length + '건' +
        (an.std.zeroShare >= 0.9 ? ' (압력 칸 값이 거의 모두 기준과 같음 — 작업일보에 실제 값이 아니라 기준(설정) 값을 적은 것으로 보임)' : ''));
    }
    if (an.repairs && an.repairs.length) {
      L.push('- 설비(Press) 수선 전후 ' + S.windowDays + '일 불량률: ' + an.repairs.map(function (m) {
        return m.equip + ' ' + m.date + ' ' + m.kind + ' 전 ' + pctText(m.before.rate, 2) + '(압력 저하 ' + m.before.lowRows + '행) → 후 ' + pctText(m.after.rate, 2) + '(압력 저하 ' + m.after.lowRows + '행)' + (m.enough ? '' : '(표본 적음)');
      }).join('; '));
    }
    if (!an.dims.wear) L.push('- 금형 누적 작업량: 금형 보수·교체 이력이 없어 계산하지 않음');
    var hot = an.hot.list;
    L.push('- 규칙으로 찾은 「불량이 몰린 조건」 ' + hot.length + '건 (기준: 표본 작업 ' + fmtNum(S.minQty) + '개·' + S.minRows + '행 이상, 전체의 ' + S.minLift + '배 이상, 95% 하한도 전체보다 높음)');
    hot.forEach(function (h, i) {
      var ov = overlapText(h, A);
      L.push('  ' + (i + 1) + '. [' + m4Text(h.m4) + '] ' + condText(h.parts, A) + ' — ' + hotEvidence(h, an.hot.overall) + (ov ? ' / 겹침: ' + ov : ''));
    });
    if (!hot.length) L.push('  (기준을 넘은 조건 없음)');
    if (an.maint.length) {
      L.push('- 금형 보수·교체 전후 ' + S.windowDays + '일 불량률: ' + an.maint.filter(function (m) { return m.before.rows && m.after.rows; }).map(function (m) {
        return A('die', m.die) + ' ' + m.date + ' ' + m.kind + ' 전 ' + pctText(m.before.rate, 2) + ' → 후 ' + pctText(m.after.rate, 2) + (m.enough ? '' : '(표본 적음)');
      }).join('; '));
      var th = an.dies.filter(function (d) { return d.threshold != null; });
      if (th.length) L.push('- 보수 직후보다 불량률이 ' + S.minLift + '배 이상 높아진 누적 작업량 구간(과거 기록, 참고): ' + th.map(function (d) {
        return A('die', d.die) + ' 보수 후 ' + fmtNum(d.threshold) + '개~ (보수 직후 ' + pctText(d.fresh, 2) + ', 지금 누적 ' + fmtNum(d.cumNow) + '개)';
      }).join(', '));
    }
    L.push('');
    L.push('[요청] 다음 순서로 한국어로 짧게 정리해줘.');
    L.push('1. 한 줄 요약 (지금 불량률과 목표의 차이)');
    L.push('2. 4M(사람·설비/금형·소재·방법)과 계절로 나눠, 불량이 몰린 조건과 근거 숫자 (조건마다 [사실]의 숫자를 그대로 인용)');
    L.push('3. 원인 추정 — 소재 문제 / 프레스 압력 문제 / 금형 보수 시점 가운데 어디에 가까운지, 각각 「추정」으로');
    L.push('4. 현장에서 확인할 것 (조건마다 2~3개: 소재 성적서·로트, 압력 설정값과 실제값, 금형 마모·보수 기록, 설비 수선 기록 등)');
    L.push('5. 불량률을 목표 쪽으로 낮추기 위해 다음 달에 해 볼 것 (데이터로 효과를 다시 잴 수 있는 것만)');
    L.push('확실하지 않은 것은 「확인 필요」라고 적어줘.');
    var text = L.join('\n');
    return { prompt: text, map: used, leaks: aliases ? findLeaks(text, aliases) : [] };
  }

  /* ── 17. 보고서 (글자) — 화면·Word·인쇄·엑셀이 같은 내용을 씁니다 ── */
  function buildReport(an, aiText) {
    var S = an.S, o = an.overall, L = [];
    L.push('Deep drawing 불량률 분석 보고 — ' + an.from + ' ~ ' + an.to);
    L.push('');
    L.push('1. 총괄 (' + metricLabel(an.metric) + ')');
    L.push('  - 작업수량 ' + fmtNum(o.qty) + '개 · 불량 ' + fmtNum(o.cnt) + '개 · 불량률 ' + pctText(o.rate, 2) + ' (목표 ' + pctText(S.target, 1) + ')');
    L.push('  - 목표까지 줄일 불량 약 ' + fmtNum(round(Math.max(0, an.gap || 0), 0)) + '개 · 목표를 넘은 달 ' + an.months.filter(function (m) { return m.overTarget; }).length + ' / ' + an.months.length + '개월');
    if (an.metric === 'all' && an.types.length) L.push('  - 유형별: ' + an.types.map(function (t) { return t + ' ' + fmtNum(an.all.types[t] || 0) + '개'; }).join(', ') + (an.all.untyped ? ', 유형 미기록 ' + fmtNum(an.all.untyped) + '개' : ''));
    L.push('');
    L.push('2. 월별 불량률');
    an.months.forEach(function (m) { L.push('  - ' + monthLabel(m.key) + ': 작업 ' + fmtNum(m.qty) + ' · 불량 ' + fmtNum(m.cnt) + ' · ' + pctText(m.rate, 2) + (m.overTarget ? ' (목표 초과)' : '')); });
    L.push('');
    L.push('3. 금형·소재별 불량률 (작업량 가중, 기대보다 많은 불량 순)');
    ['die', 'maker', 'lot', 'spec', 'model'].forEach(function (d) {
      var g = an.dims[d]; if (!g) return;
      L.push('  [' + DIMS[d].label + ']');
      g.rows.slice(0, 8).forEach(function (r) { L.push('  - ' + r.label + ': 작업 ' + fmtNum(r.qty) + ' · 불량 ' + fmtNum(r.cnt) + ' · ' + pctText(r.rate, 2) + (r.small ? ' (표본 적음)' : '')); });
    });
    L.push('');
    L.push('4. 불량이 몰린 조건 (4M·계절)');
    if (!an.hot.list.length) L.push('  - 기준을 넘은 조건이 없습니다.');
    an.hot.list.forEach(function (h) { var ov = overlapText(h); L.push('  - [' + m4Text(h.m4) + '] ' + condText(h.parts) + ': ' + hotEvidence(h, an.hot.overall) + (ov ? ' (겹침: ' + ov + ')' : '')); });
    L.push('  ※ 상관을 보여 줄 뿐 원인을 확정하지 않습니다. 현장 확인이 필요합니다.');
    L.push('');
    L.push('5. 금형 보수·교체');
    if (!an.maint.length) L.push('  - 보수·교체 이력이 없어 전후 비교를 하지 않았습니다(설정에서 넣을 수 있습니다).');
    an.maint.forEach(function (m) {
      L.push('  - ' + m.die + ' ' + m.date + ' ' + m.kind + ': ' + (m.before.rows ? '전 ' + S.windowDays + '일 ' + pctText(m.before.rate, 2) : '전 기록 없음(데이터 시작 전)') + ' → 후 ' + pctText(m.after.rate, 2) + (m.enough ? '' : ' (표본 적음)'));
    });
    an.dies.filter(function (d) { return d.threshold != null; }).forEach(function (d) {
      L.push('  - 참고: ' + d.die + ' 은 과거 보수 후 누적 ' + fmtNum(d.threshold) + '개 구간부터 불량률이 보수 직후(' + pctText(d.fresh, 2) + ')의 ' + S.minLift + '배 이상이었습니다. 지금 누적 ' + fmtNum(d.cumNow) + '개.');
    });
    L.push('');
    L.push('6. 압력 작업표준 대비 · 설비(Press) 수선');
    if (!an.std || !an.std.has) L.push('  - 모델(금형)별 압력 작업표준을 넣지 않아 기준 대비를 보지 않았습니다(설정에서 넣을 수 있습니다).');
    else {
      L.push('  - ' + an.std.label + '별 기준 ± 허용보다 낮은 작업(설비보전·공무팀 전파 대상): Main ' + an.alerts.filter(function (a) { return a.kind === 'main'; }).length + '건 · 쿠션 ' + an.alerts.filter(function (a) { return a.kind === 'cushion'; }).length + '건');
      ['mainStd', 'cushionStd'].forEach(function (d) {
        var g = an.dims[d]; if (!g) return;
        L.push('  - ' + DIMS[d].label + ': ' + g.rows.map(function (r) { return r.label + ' ' + pctText(r.rate, 2) + '(작업 ' + fmtNum(r.qty) + (r.small ? ', 표본 적음' : '') + ')'; }).join(', '));
      });
      if (an.std.zeroShare >= 0.9) L.push('  - 압력 칸 값이 거의 모두 기준과 같습니다. 작업일보에 실제 압력을 적기 시작하면 기준 대비 저하를 볼 수 있습니다.');
    }
    if (an.repairs && an.repairs.length) an.repairs.forEach(function (m) {
      L.push('  - ' + m.equip + ' ' + m.date + ' ' + m.kind + (m.note ? '(' + m.note + ')' : '') + ': 전 ' + S.windowDays + '일 ' + pctText(m.before.rate, 2) + ' · 압력 저하 ' + m.before.lowRows + '행 → 후 ' + pctText(m.after.rate, 2) + ' · 압력 저하 ' + m.after.lowRows + '행' + (m.enough ? '' : ' (표본 적음)'));
    });
    else L.push('  - 설비 수선 이력이 없습니다(설정에서 넣을 수 있습니다).');
    if (aiText && str(aiText)) {
      L.push('');
      L.push('7. AI 분석 의견 (검토 후 사용)');
      str(aiText).split('\n').forEach(function (s) { L.push('  ' + s); });
    }
    return L.join('\n');
  }

  /* 엑셀 내보내기용 표 */
  function groupToAoa(g) {
    var aoa = [[g.label + '별 불량률 (' + metricLabel(g.metric) + ', 작업량 가중)'], [g.label, '행', '작업일', '작업수량', '불량수량', '불량률(%)', '전체 대비(배)', '기대보다 많은 불량', '불량 비중(%)', '표본']];
    g.rows.forEach(function (r) {
      aoa.push([r.label, r.rows, r.days, r.qty, r.cnt, r.rate == null ? '' : round(r.rate * 100, 2), r.lift == null ? '' : round(r.lift, 2), r.excess == null ? '' : round(r.excess, 0), r.share == null ? '' : round(r.share * 100, 1), r.small ? '표본 적음' : '']);
    });
    aoa.push(['전체', g.overall.rows, g.overall.days, g.overall.qty, g.overall.cnt, g.overall.rate == null ? '' : round(g.overall.rate * 100, 2), 1, 0, 100, '']);
    return aoa;
  }
  function trendToAoa(list, title) {
    var aoa = [[title], ['기간', '행', '작업수량', '불량수량', '불량률(%)', '목표 초과']];
    list.forEach(function (m) { aoa.push([m.key, m.rows, m.qty, m.cnt, m.rate == null ? '' : round(m.rate * 100, 2), m.overTarget ? '초과' : '']); });
    return aoa;
  }
  function recordsToAoa(records, types) {
    types = types || typeList(records);
    var aoa = [['작업일자', '작업자', 'Main 압력', '쿠션 압력', '금형', '작업모델', '소재 메이커', '소재 로트', '소재규격', '작업수량', '불량수량', '불량률(%)'].concat(types).concat(['Main 기준', 'Main − 기준', '쿠션 기준', '쿠션 − 기준', '설비', '비고', '원본 파일'])];
    (records || []).forEach(function (r) {
      aoa.push([r.date, r.worker, r.main, r.cushion, r.die, r.model, r.maker, r.lot, r.spec, r.qty, r.defect, round(r.defect / r.qty * 100, 2)]
        .concat(types.map(function (t) { return (r.types && r.types[t]) || 0; }))
        .concat([r.mainStd == null ? '' : r.mainStd, r.mainDev == null ? '' : round(r.mainDev, 2), r.cushionStd == null ? '' : r.cushionStd, r.cushionDev == null ? '' : round(r.cushionDev, 2), r.press || '', r.note, r.file]));
    });
    return aoa;
  }

  /* ── 18. 저장 형식 ─────────────────────────────────────────── */
  function emptyDb() {
    return { schemaVersion: SCHEMA_VERSION, records: [], files: [], settings: cleanSettings({}), maint: [], stds: cleanStds({}), repairs: [], aliases: {}, mask: true, aiNotes: {}, sample: false, colOverrides: {} };
  }
  /* 쌓인 자료가 지금 열 읽기 판으로 가져온 것인지 — { stale, oldRows, oldFiles }
     행마다 pv(열 읽기 판)를 적어 두므로, 같은 파일을 지우지 않고 다시 넣어 옛 행이 남아 있어도(메이커·규격이 바뀐 행은 키가 달라 따로 남음) 알아봅니다. */
  function importStatus(db) {
    var out = { stale: false, oldRows: 0, oldFiles: [] };
    if (!db || !db.records || !db.records.length) return out;
    db.records.forEach(function (r) { if (!(Number(r.pv) >= PARSER_VERSION)) out.oldRows++; });
    (db.files || []).forEach(function (f) { if (!(Number(f.parser) >= PARSER_VERSION)) out.oldFiles.push(f.name || '(이름 없음)'); });
    out.stale = out.oldRows > 0;
    return out;
  }
  /* 열 연결 기록 — 파일마다 어느 칸을 어느 항목으로 읽었는지(② 화면에 「소재 메이커 ← G열 '메이커'」로 보여 줌) */
  function fileColumns(det) {
    var m = {};
    Object.keys((det && det.map) || {}).forEach(function (f) { var c = det.map[f]; m[f] = { col: colName(c), i: c, head: (det.headers && det.headers[c]) || '' }; });
    return m;
  }
  /* 그 항목이 쌓인 파일들에서 어느 칸이었는지 — [{ text: "G열 「메이커」", files: 12 }] (같은 것끼리 묶음), 기록 없는 파일 수 */
  function columnSources(files, field) {
    var by = {}, order = [], unknown = 0, none = 0;
    (files || []).forEach(function (f) {
      if (!f.cols) { unknown++; return; }
      var c = f.cols[field];
      if (!c) { none++; return; }
      var t = c.col + '열 「' + c.head + '」';
      if (!by[t]) { by[t] = 0; order.push(t); }
      by[t]++;
    });
    return { list: order.map(function (t) { return { text: t, files: by[t] }; }), unknown: unknown, none: none };
  }
  /* 「최근 N개월」 시작일 — 끝 날짜가 든 달을 포함해 N개 달의 첫날 (2026-09-30, 3 → 2026-07-01).
     예전에는 N × 30.4일을 빼서 07-02 · 04-02 처럼 하루씩 어긋났습니다. */
  function monthsBackStart(to, n) {
    var y = +to.slice(0, 4), m = +to.slice(5, 7) - (n - 1);
    while (m < 1) { m += 12; y--; }
    return y + '-' + pad(m) + '-01';
  }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (Array.isArray(o.records)) db.records = sortRecords(o.records.filter(function (r) { return r && r.date && r.die && r.qty > 0; }).map(function (r) {
      r.types = r.types && typeof r.types === 'object' ? r.types : {};
      if (!r.key) r.key = baseKey(r) + '#1';
      return r;
    }));
    if (Array.isArray(o.files)) db.files = o.files;
    db.settings = cleanSettings(o.settings || {});
    db.maint = cleanMaint(o.maint);
    db.stds = cleanStds(o.stds);
    db.repairs = cleanRepairs(o.repairs);
    if (o.aliases && typeof o.aliases === 'object') Object.keys(MASK_CATS).forEach(function (c) { var a = assignAliases(o.aliases, c, []); if (Object.keys(a).length) db.aliases[c] = a; });
    if (o.mask === false) db.mask = false;
    if (o.aiNotes && typeof o.aiNotes === 'object') db.aiNotes = o.aiNotes;
    db.sample = !!o.sample;
    // 고친 열 연결 — 머리행 모양(headerSignature)마다 기억해 같은 모양 파일을 다시 넣을 때 그대로 씁니다
    if (o.colOverrides && typeof o.colOverrides === 'object') Object.keys(o.colOverrides).forEach(function (k) {
      var v = o.colOverrides[k]; if (v && typeof v === 'object') db.colOverrides[k] = v;
    });
    return db;
  }

  return {
    SCHEMA_VERSION: SCHEMA_VERSION, VERSION: VERSION, PARSER_VERSION: PARSER_VERSION, importStatus: importStatus, fileColumns: fileColumns, valueKind: valueKind, materialHint: materialHint, materialCheck: materialCheck, swapMaterial: swapMaterial, columnSamples: columnSamples, columnValues: columnValues, columnSources: columnSources, monthsBackStart: monthsBackStart, FIELDS: FIELDS, DEFECT_TYPES: DEFECT_TYPES, DEFAULT_SETTINGS: DEFAULT_SETTINGS, SEASONS: SEASONS, M4: M4, M4_SHORT: M4_SHORT, DIMS: DIMS, MASK_CATS: MASK_CATS,
    norm: norm, toNum: toNum, toISO: toISO, addDays: addDays, daysBetween: daysBetween, seasonOf: seasonOf, monthOf: monthOf, monthLabel: monthLabel, colName: colName, round: round,
    matchField: matchField, matchType: matchType, detectColumns: detectColumns, parseRows: parseRows, parseSheet: parseSheet, pickSheet: pickSheet, headerSignature: headerSignature,
    mergeRecords: mergeRecords, sortRecords: sortRecords, monthsOf: monthsOf, monthCoverage: monthCoverage, dateRange: dateRange,
    cleanSettings: cleanSettings, cleanMaint: cleanMaint, parseMaintText: parseMaintText, computeWear: computeWear,
    cleanStds: cleanStds, parseStdText: parseStdText, suggestStds: suggestStds, stdOf: stdOf, devOf: devOf, REPAIR_KINDS: REPAIR_KINDS, cleanRepairs: cleanRepairs, parseRepairText: parseRepairText, aoaToText: aoaToText,
    stdSummary: stdSummary, pressureAlerts: pressureAlerts, alertText: alertText, alertsToAoa: alertsToAoa, stdToAoa: stdToAoa, repairEffects: repairEffects, sameGrouping: sameGrouping,
    aggregate: aggregate, typeList: typeList, metricLabel: metricLabel, wilson: wilson, niceStep: niceStep, makeBins: makeBins, binOf: binOf,
    buildContext: buildContext, dimValue: dimValue, dimAvailable: dimAvailable, groupBy: groupBy, crossTab: crossTab, trend: trend,
    findHotspots: findHotspots, m4Text: m4Text, condText: condText, overlapText: overlapText, hotEvidence: hotEvidence, maintEffects: maintEffects, dieStatus: dieStatus, analyze: analyze,
    assignAliases: assignAliases, assignAllAliases: assignAllAliases, unmaskText: unmaskText, findLeaks: findLeaks,
    buildAiPrompt: buildAiPrompt, buildAiPromptInfo: buildAiPromptInfo, buildReport: buildReport,
    groupToAoa: groupToAoa, trendToAoa: trendToAoa, recordsToAoa: recordsToAoa, fmtNum: fmtNum, pctText: pctText,
    emptyDb: emptyDb, restoreDb: restoreDb
  };
});
