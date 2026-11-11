/* Deep drawing 불량률 분석 — 화면
   계산은 js/logic.js(DDLogic), 그래프는 js/charts.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js·ai-panel.js.
   뼈대(가져오기·열 연결·AI 패널·보고서 내보내기)는 data09-25 에서 가져왔습니다. */
(function () {
  'use strict';
  var L = window.DDLogic, S = window.DDStore, C = window.DDCharts, AIP = window.AIPanel;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 버튼과 결과가 어긋날 수 있어, 섞였으면 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js · sample-data.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.1';
  var mixed = [['js/logic.js', L], ['js/sample-data.js', window.DDSample]].filter(function (x) { return !x[1] || x[1].VERSION !== APP_VERSION; });
  if (mixed.length || !S || !C || !AIP) {
    var m0 = document.getElementById('main');
    m0.innerHTML = '';
    var box = document.createElement('section'); box.className = 'card'; box.id = 'versionMismatch';
    box.innerHTML = '<h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(' +
      mixed.map(function (x) { return x[0] + ' ' + (x[1] && x[1].VERSION ? x[1].VERSION : '옛 판'); }).join(', ') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p>';
    m0.appendChild(box);
    return;
  }
  var db = S.loadDb();
  var ui = { from: '', to: '', metric: 'all', month: '', dim: 'die', sort: 'excess', pending: [], errors: [], progress: '' };
  var main = document.getElementById('main');

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint) {
    return h('label', { class: 'field' }, h('span', null, label), control, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 3500);
  }
  function save() {
    var ok = S.saveDb(db);
    document.getElementById('storeNote').textContent = ok ? '이 브라우저에 저장됨 · ' + S.sizeKb() + 'KB' : '저장 안 됨(브라우저 저장소가 막혔거나 가득 참) — JSON 백업을 받아 두세요';
    return ok;
  }
  var pct = function (v) { return L.pctText(v, 2); };
  var N = L.fmtNum;
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  /* 엑셀 파일(첫 시트) → 붙여 넣기 글과 같은 줄들 — 이력·작업표준 「엑셀 파일로 넣기」 */
  function sheetFileInput(onText) {
    var i = h('input', { type: 'file', accept: '.xlsx,.xlsm,.xls,.csv' });
    i.addEventListener('change', function () {
      var f = i.files[0]; if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true });
          onText(L.aoaToText(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null })));
        } catch (e) { toast(f.name + ': 엑셀을 읽지 못했습니다.', true); }
        i.value = '';
      };
      fr.readAsArrayBuffer(f);
    });
    return i;
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  /* 이 PC 시각(예전에는 toISOString 이라 한국 시각보다 9시간 이른 UTC 가 「가져온 때」에 찍혔습니다) */
  function localStamp(dateOnly) { var d = new Date(); var s = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); return dateOnly ? s : s + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
  function copyText(t) {
    function fallback() { var ta = h('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast('복사했습니다.'); } catch (e) { toast('복사하지 못했습니다. 직접 선택해 복사해 주세요.', true); } ta.remove(); }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(function () { toast('복사했습니다.'); }, fallback); else fallback();
  }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function emptyNotice() {
    return h('section', { class: 'card empty' },
      h('p', null, '아직 가져온 작업일보가 없습니다.'),
      h('p', null, '작업일보 엑셀을 가져오거나, 가상 샘플 12개월로 먼저 둘러봐 주세요.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        h('a', { class: 'btn btn-primary', href: '#/import' }, '데이터 가져오기로'),
        h('button', { type: 'button', class: 'btn', onclick: loadSample }, '샘플 12개월로 시작')));
  }
  function tile(k, v, sub, cls) { return h('div', { class: 'tile' }, h('div', { class: 'k' }, k), h('div', { class: 'v' + (cls ? ' ' + cls : '') }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: ((opts.num && opts.num[i] ? 'num' : '') + (r.tdCls && r.tdCls[i] ? ' ' + r.tdCls[i] : '')) || null }, x)); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }

  /* ── 기간 · 보는 불량 ─────────────────────── */
  function settings() { return L.cleanSettings(db.settings); }
  function fullRange() { return L.dateRange(db.records); }
  function period() {
    var r = fullRange();
    var f = ui.from && ui.from >= r.from && ui.from <= r.to ? ui.from : r.from;
    var t = ui.to && ui.to <= r.to && ui.to >= r.from ? ui.to : r.to;
    if (f > t) { var x = f; f = t; t = x; }
    return { from: f, to: t };
  }
  function context() {
    var p = period();
    return L.buildContext(db.records, { from: p.from, to: p.to, settings: db.settings, maint: db.maint, stds: db.stds, repairs: db.repairs });
  }
  function metric(ctx) { return ui.metric !== 'all' && ctx.types.indexOf(ui.metric) < 0 ? 'all' : ui.metric; }
  function scopeCard(ctx, extra) {
    var p = period(), r = fullRange();
    var f = h('input', { type: 'date', value: p.from, min: r.from, max: r.to }), t = h('input', { type: 'date', value: p.to, min: r.from, max: r.to });
    f.addEventListener('change', function () { ui.from = f.value; render(); });
    t.addEventListener('change', function () { ui.to = t.value; render(); });
    var m = h('select', { 'aria-label': '보는 불량' });
    ['all'].concat(ctx.types).forEach(function (k) { var o = h('option', { value: k }, L.metricLabel(k)); if (k === metric(ctx)) o.selected = true; m.appendChild(o); });
    m.addEventListener('change', function () { ui.metric = m.value; render(); });
    // 「최근 N개월」 = 끝 날짜가 든 달까지 N개 달(달 첫날부터). 지금 기간과 같은 단추는 눌린 모양으로 보여 줍니다.
    function quickFrom(months) { var f0 = months ? L.monthsBackStart(r.to, months) : r.from; return f0 < r.from ? r.from : f0; }
    function quick(label, months) {
      var on = p.to === r.to && p.from === quickFrom(months);
      return h('button', { type: 'button', class: 'btn btn-sm', 'aria-pressed': on ? 'true' : 'false', onclick: function () {
        ui.to = r.to; ui.from = quickFrom(months); render();
      } }, label);
    }
    return h('section', { class: 'card filters no-print' },
      fld('시작일', f), fld('종료일', t),
      fld('보는 불량', m, ctx.types.length ? '유형 칸(' + ctx.types.join('·') + ')이 있어 유형별로도 볼 수 있습니다.' : '불량 유형 칸이 없어 전체 불량만 봅니다.'),
      h('div', { class: 'btn-row span-all quick-period', role: 'group', 'aria-label': '기간 빠르게 고르기' }, quick('전체 기간', 0), quick('최근 6개월', 6), quick('최근 3개월', 3), extra || null));
  }
  function barColors(rows, overall, S0) {
    return rows.map(function (r) { return r.small ? 'var(--series-plan)' : (r.lift != null && r.lift >= S0.minLift ? 'var(--series-2)' : 'var(--series-1)'); });
  }
  function refs(ctx, overallRate) {
    var a = [{ value: ctx.S.target, label: '목표 ' + L.pctText(ctx.S.target, 1) }];
    if (overallRate != null) a.push({ value: overallRate, label: '이 기간 전체 ' + pct(overallRate), cls: 'now' });
    return a;
  }

  /* ── 가져오기 ───────────────────────────── */
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true });
          resolve(prepare(file.name, wb));
        } catch (e) { reject(new Error(file.name + ': 엑셀로 읽지 못했습니다(' + e.message + ')')); }
      };
      fr.onerror = function () { reject(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function prepare(name, wb) {
    var sheets = wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; });
    var pick = L.pickSheet(sheets) || sheets[0];
    var item = { fileName: name, sheets: sheets, sheet: pick ? pick.name : '', overrides: {} };
    reparse(item);
    // 같은 머리행 모양의 파일에서 전에 고친 열 연결이 있으면 그대로 씁니다(머리행마다 기억 — db.colOverrides)
    var saved = db.colOverrides && db.colOverrides[L.headerSignature(item.result.det)];
    if (saved && Object.keys(saved).length) { item.overrides = JSON.parse(JSON.stringify(saved)); item.savedOverrides = true; reparse(item); }
    return item;
  }
  function reparse(item) {
    var sh = item.sheets.filter(function (s) { return s.name === item.sheet; })[0];
    item.aoa = sh ? sh.aoa : [];
    item.result = L.parseSheet(item.aoa, { fileName: item.fileName, overrides: item.overrides });
  }
  function fileId(f) { return f.name + '|' + f.size + '|' + (f.lastModified || 0); }
  /* 여러 달 치를 한 번에: 하나씩 차례로 읽고(화면 멈춤 방지) 진행을 보여 주며, 엑셀 잠금 파일(~$…)·엑셀이 아닌 파일·이미 올려 둔 파일은 건너뜁니다 */
  function handleFiles(list) {
    var all = Array.prototype.slice.call(list || []);
    var skipped = [], batch = {};
    var files = all.filter(function (f) {
      if (batch[fileId(f)]) { skipped.push(f.name + ' (같은 파일을 두 번 고름)'); return false; }
      batch[fileId(f)] = 1;
      if (/^~\$/.test(f.name)) { skipped.push(f.name + ' (엑셀 잠금 파일)'); return false; }
      if (!/\.(xlsx|xlsm|xls|csv)$/i.test(f.name)) { if (!/^\./.test(f.name)) skipped.push(f.name + ' (엑셀 아님)'); return false; }
      if (ui.pending.some(function (p) { return p.fid === fileId(f); })) { skipped.push(f.name + ' (이미 올려 둠)'); return false; }
      return true;
    });
    if (!files.length) { toast(skipped.length ? '가져올 엑셀이 없습니다(' + skipped.length + '개 건너뜀).' : '엑셀 파일(.xlsx·.xls·.csv)을 골라 주세요.', true); return; }
    files.sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });
    ui.errors = skipped.map(function (x) { return '건너뜀: ' + x; });
    var i = 0;
    (function next() {
      if (i >= files.length) { ui.progress = ''; render(); toast(files.length + '개 파일을 읽었습니다. 아래에서 확인한 뒤 반영해 주세요.'); return; }
      var f = files[i++];
      ui.progress = '읽는 중 ' + i + ' / ' + files.length + ' — ' + f.name;
      var pr = document.getElementById('importProgress'); if (pr) pr.textContent = ui.progress;
      readFile(f).then(function (it) {
        it.fid = fileId(f);
        it.already = db.files.some(function (x) { return x.name === f.name; });
        ui.pending.push(it);
        applySameHeader(it);
      }, function (e) { ui.errors.push(e.message); }).then(function () { setTimeout(next, 0); });
    })();
  }
  /* 머리행이 같은 파일에 이미 고친 열 연결이 있으면 그대로 씁니다(첫 파일만 고치면 나머지도 따라감) */
  function applySameHeader(it) {
    var sig = L.headerSignature(it.result.det);
    var src = ui.pending.filter(function (p) { return p !== it && Object.keys(p.overrides).length && L.headerSignature(p.result.det) === sig; })[0];
    if (src) { it.overrides = JSON.parse(JSON.stringify(src.overrides)); reparse(it); }
  }
  function commit(items) {
    var total = { added: 0, updated: 0 }, ok = 0;
    items.forEach(function (it) {
      var r = it.result;
      if (r.det.missing.length) return;
      var m = L.mergeRecords(db.records, r.records);
      db.records = m.records;
      total.added += m.stat.added; total.updated += m.stat.updated;
      // parser·cols: 어느 판으로 · 어느 칸을 읽었는지(다음 판에서 옛 자료를 알아보고, ② 화면에 칸 출처를 보여 줌)
      db.files = db.files.filter(function (f) { return f.name !== it.fileName; });
      var sig = L.headerSignature(r.det);
      db.files.push({ name: it.fileName, sheet: it.sheet, rows: r.records.length, at: localStamp(), parser: L.PARSER_VERSION, cols: L.fileColumns(r.det), sig: sig });
      if (Object.keys(it.overrides).length && !it.sample) { db.colOverrides = db.colOverrides || {}; db.colOverrides[sig] = JSON.parse(JSON.stringify(it.overrides)); }
      ok++;
    });
    db.sample = !!db.sample && items.every(function (it) { return it.sample; });
    ui.pending = ui.pending.filter(function (p) { return items.indexOf(p) < 0; });
    ui.from = ''; ui.to = '';
    save();
    ui.errors = [];
    toast(ok + '개 파일 반영 — 새 행 ' + total.added + ' · 다시 가져와 바뀐 행 ' + total.updated);
    render();
  }
  function loadSample() {
    var b = window.DDSample.build();
    db = L.emptyDb(); db.sample = true;
    db.maint = L.cleanMaint(b.maint);
    db.stds = L.cleanStds(b.stds); db.repairs = L.cleanRepairs(b.repairs);
    // 정상 범위는 실제 값을 받기 전이라(기획서 10장 ⑤) 샘플에만 가상 값을 넣어 둡니다
    db.settings = L.cleanSettings({ mainMin: 180, mainMax: 250, cushionMin: 40, cushionMax: 60 });
    var items = b.files.map(function (f) {
      // 앱이 실제 파일을 읽을 때와 같은 길로 — 엑셀 통합문서로 만들었다가 다시 읽습니다
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(f.aoa), '작업일보');
      var back = XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellDates: true });
      var it = prepare(f.fileName, back); it.sample = true; return it;
    });
    ui.pending = []; ui.metric = 'all';
    commit(items);
    location.hash = '#/cond';
  }

  function pageImport() {
    var input = h('input', { type: 'file', multiple: true, accept: '.xlsx,.xlsm,.xls,.csv' });
    input.addEventListener('change', function () { handleFiles(input.files); input.value = ''; });
    var folder = h('input', { type: 'file', multiple: true, webkitdirectory: true });
    folder.addEventListener('change', function () { handleFiles(folder.files); folder.value = ''; });
    var drop = h('div', { class: 'drop' },
      h('label', { class: 'btn btn-primary' }, input, '작업일보 엑셀 고르기(여러 개 가능)'),
      h('label', { class: 'btn' }, folder, '폴더째 고르기'),
      h('span', { class: 'note' }, '또는 파일을 여기로 끌어다 놓아 주세요. 6개월~1년 치 파일을 한꺼번에 넣어도 됩니다.'));
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); handleFiles(e.dataTransfer.files); });

    var reimport = L.importStatus(db).stale && !db.sample ? h('section', { class: 'card', id: 'reimportGuide' }, h('h2', null, '다시 가져오기 (이전 판 자료)'),
      h('ol', { class: 'steps' },
        h('li', null, '먼저 설정 화면에서 「JSON 백업 내려받기」로 지금 상태를 받아 두면 안전합니다(금형 보수 · 작업표준 · 수선 이력은 지우지 않습니다).'),
        h('li', null, '아래 「작업일보만 지우기」를 누릅니다.'),
        h('li', null, '「작업일보 엑셀 고르기」로 같은 작업일보 파일들을 다시 넣고, 열 연결에서 소재 메이커 · 소재 lot · 소재규격이 맞는 칸인지 확인한 뒤 반영합니다.')),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
        if (!confirm('쌓인 작업일보 행과 가져온 파일 목록만 지울까요? (설정 · 금형 보수 · 작업표준 · 수선 이력 · AI 연결은 그대로 둡니다)')) return;
        db.records = []; db.files = []; db.aiNotes = {}; ui.from = ''; ui.to = ''; save(); toast('작업일보를 지웠습니다. 파일을 다시 넣어 주세요.'); render();
      } }, '작업일보만 지우기'))) : null;
    var out = [reimport, pageHead('데이터', '작업일보 가져오기', 'Deep drawing 공정 작업일보 엑셀을 그대로 넣어 주세요. 제목 줄 아래의 머리행(작업일자·작업자·Main 압력·쿠션 압력·금형·작업모델·소재 메이커·소재규격·작업수량·불량수량)을 알아서 찾습니다.'),
      h('section', { class: 'card' }, h('h2', null, '1. 파일 넣기'), drop,
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { if (!db.records.length || confirm('지금 데이터를 지우고 가상 샘플 12개월을 불러올까요?')) loadSample(); } }, '샘플 12개월로 시작'),
          h('a', { class: 'btn', href: 'samples/' + encodeURIComponent('작업일보_샘플_2026-01.xlsx'), download: '' }, '샘플 파일 내려받기(2026년 1월)')),
        h('ul', { class: 'note' },
          h('li', null, '열 이름이 조금 달라도(「일자」「금형번호」「Main 압력(bar)」 등) 찾습니다. 못 찾은 열은 아래 「열 연결」에서 골라 주세요. 한 파일에서 고치면 머리행이 같은 나머지 파일에도 똑같이 적용됩니다.'),
          h('li', null, '불량률은 늘 「불량수량 합계 ÷ 작업수량 합계」(작업량 가중)로 셉니다. 「합계」「소계」 줄은 빼고, 불량수량이 빈 행은 불량 0 으로 셉니다.'),
          h('li', null, '「Crack」「주름」 같은 불량 유형 칸이 있으면 유형별로도 봅니다.'),
          h('li', null, '같은 파일을 두 번 넣어도 두 번 세지 않습니다(같은 행은 새 값으로 바뀜). 같은 날 같은 조건으로 두 번 작업한 행은 따로 셉니다.'),
          h('li', null, '파일은 서버로 보내지 않고 이 브라우저 안에서만 읽습니다.')),
        h('p', { class: 'note', id: 'importProgress', 'aria-live': 'polite' }, ui.progress || ''),
        ui.errors.length ? h('details', { class: 'warn-list', open: true }, h('summary', null, '읽지 못했거나 건너뛴 파일 ' + ui.errors.length + '개'),
          h('ul', null, ui.errors.map(function (e) { return h('li', null, e); }))) : null)];

    if (ui.pending.length) {
      var box = h('section', { class: 'card' }, h('h2', null, '2. 열 연결 확인 후 반영 (' + ui.pending.length + '개 파일)'));
      var ready = ui.pending.filter(function (it) { return !it.result.det.missing.length; });
      box.appendChild(table(['파일', '기간', '작업 행', '건너뛴 행', '알림', '상태'], ui.pending.map(function (it) {
        var r = it.result, rg = L.dateRange(r.records);
        return { cls: r.det.missing.length ? 'excluded' : null, cells: [it.fileName, rg.from ? rg.from + ' ~ ' + rg.to : '없음', r.records.length, r.skipped, r.warnings.length + r.det.warnings.length,
          r.det.missing.length ? '열 연결 필요' : it.already ? '이전에 가져온 파일(새 값으로 바뀜)' : '반영 준비됨'] };
      }), { num: [0, 0, 1, 1, 1, 0] }));
      box.appendChild(h('h3', null, '파일별 자세히'));
      ui.pending.forEach(function (it) { box.appendChild(pendingCard(it)); });
      box.appendChild(h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', disabled: !ready.length, onclick: function () { commit(ready); } }, '확인한 ' + ready.length + '개 파일 반영'),
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.pending = []; render(); } }, '취소')));
      out.push(box);
    }

    var rg = fullRange(), cov = L.monthCoverage(db.records), a = L.aggregate(db.records);
    out.push(h('section', { class: 'card' }, h('h2', null, '쌓인 데이터'),
      db.records.length ? h('div', null,
        h('div', { class: 'tiles' },
          tile('기간', cov.months.length + '개월', rg.from + ' ~ ' + rg.to),
          tile('작업 행', N(db.records.length), '작업일 ' + N(a.days) + '일'),
          tile('작업수량', N(a.qty) + '개', '불량 ' + N(a.defect) + '개'),
          tile('불량률', pct(a.defectRate), '목표 ' + L.pctText(settings().target, 1))),
        cov.missing.length ? h('p', { class: 'alert warn' }, '작업 기록이 없는 달이 있습니다: ' + cov.missing.map(L.monthLabel).join(', ') + '. 휴무가 아니라면 그 달 파일을 더 넣어 주세요.')
          : h('p', { class: 'note' }, L.monthLabel(cov.first) + ' ~ ' + L.monthLabel(cov.last) + ' — 빠진 달 없이 이어져 있습니다.'),
        h('details', null, h('summary', null, '가져온 파일 ' + db.files.length + '개 보기'),
          table(['가져온 파일', '시트', '행', '가져온 때'], db.files.slice().reverse().map(function (f) { return [f.name, f.sheet, f.rows, f.at]; }), { num: [0, 0, 1, 0] })),
        h('div', { class: 'btn-row', style: 'margin-top:12px' },
          h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
            if (!confirm('이 브라우저에 쌓인 작업일보·설정·보수 이력을 모두 지울까요? (되돌릴 수 없습니다. 먼저 설정 화면에서 JSON 백업을 받아 두면 안전합니다)')) return;
            db = L.emptyDb(); S.clearDb(); ui.from = ''; ui.to = ''; save(); render();
          } }, '모든 데이터 지우기')))
        : h('p', { class: 'note' }, '아직 없습니다.')));
    return out;
  }
  function pendingCard(it) {
    var r = it.result, det = r.det;
    var sheetSel = h('select', { 'aria-label': '시트' });
    it.sheets.forEach(function (s) { var o = h('option', { value: s.name }, s.name); if (s.name === it.sheet) o.selected = true; sheetSel.appendChild(o); });
    sheetSel.addEventListener('change', function () { it.sheet = sheetSel.value; it.overrides = {}; reparse(it); render(); });
    var hdr = det.headers || [];
    var ex = it.aoa[det.headerRow + 1] || [];
    var mapGrid = h('div', { class: 'map-grid' });
    L.FIELDS.forEach(function (f) {
      var sel = h('select', { 'aria-label': f.label + ' 열' });
      sel.appendChild(h('option', { value: '-1' }, '(이 파일에는 없음)'));
      hdr.forEach(function (t, i) {
        var o = h('option', { value: String(i) }, L.colName(i) + ' · ' + (t || '(빈 머리)'));
        if (det.map[f.id] === i) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', function () { var o = {}; o[f.id] = Number(sel.value); setOverrides(it, o); });
      var col = det.map[f.id];
      var exs = col == null ? [] : L.columnSamples(it.aoa, det.headerRow, col, 3);
      mapGrid.appendChild(h('div', { class: 'map-row' + (f.req && col == null ? ' missing' : '') },
        h('span', { class: 'col' }, f.label, f.req ? h('span', { class: 'req' }, ' 필수') : null), sel,
        h('span', { class: 'ex' }, col == null ? '연결 안 됨' : '예: ' + (exs.length ? exs.join(' · ') : '(빈칸)'))));
    });
    var mat = det.material || {};
    var rg = L.dateRange(r.records);
    var needs = det.missing.length > 0;
    return h('details', { class: 'block', open: needs || mat.warn || ui.pending.length <= 3 ? true : null }, h('summary', null, it.fileName + (det.missing.length ? ' — 열 연결 필요' : mat.warn ? ' — 소재 칸 확인 필요' : '')),
      h('div', { class: 'form-grid' },
        fld('시트', sheetSel, it.sheets.length > 1 ? '작업일보 시트를 자동으로 골랐습니다.' : null),
        h('div', { class: 'field' }, h('span', null, '읽은 결과'),
          h('div', null, det.headerRow < 0 ? '머리행을 못 찾음' : '머리행 ' + (det.headerRow + 1) + '행' + (det.groupRow >= 0 ? '(윗줄 묶음 머리 ' + (det.groupRow + 1) + '행)' : '') +
            ' · 작업 행 ' + r.records.length + ' · 건너뛴 행 ' + r.skipped + (r.summary ? ' · 합계 줄 ' + r.summary : '')),
          h('small', { class: 'hint' }, '기간: ' + (rg.from ? rg.from + ' ~ ' + rg.to : '없음') +
            ' · 불량 유형 칸: ' + (det.typeCols.map(function (x) { return x.type + '(' + L.colName(x.col) + ')'; }).join(', ') || '없음')))),
      materialCard(it),
      det.warnings.filter(function (w) { return w !== mat.warn; }).map(function (w) { return h('p', { class: 'alert warn' }, w); }),
      r.warnings.length ? h('details', { class: 'warn-list' }, h('summary', null, '행 점검 알림 ' + r.warnings.length + '건'),
        h('ul', null, r.warnings.slice(0, 80).map(function (w) { return h('li', null, w.row + '행: ' + w.msg); }))) : null,
      h('details', { open: det.missing.length || mat.warn ? true : null }, h('summary', null, '열 연결 보기·고치기 (' + Object.keys(det.map).length + '개 연결됨)'), mapGrid));
  }
  /* 열 연결 바꾸기 — 이 파일과 머리행이 같은 올려 둔 파일 모두. 원본 칸(시트)을 들고 있으니 바로 다시 읽습니다 */
  function setOverrides(it, o) {
    var sig = L.headerSignature(it.result.det), n = 0;
    ui.pending.forEach(function (p) {
      if (p !== it && L.headerSignature(p.result.det) !== sig) return;
      Object.keys(o).forEach(function (k) { p.overrides[k] = o[k]; }); reparse(p); if (p !== it) n++;
    });
    if (n) toast('머리행이 같은 파일 ' + n + '개에도 같은 연결을 적용했습니다.');
    render();
  }
  /* 소재 칸 확인 — 어느 엑셀 열이 소재 메이커 · 소재 lot · 소재규격으로 들어가는지, 값 3개씩. 값이 반대 모양이면 「두 칸 바꾸기」 */
  function materialCard(it) {
    var det = it.result.det, mat = det.material || {};
    if (det.headerRow < 0) return null;
    var rows = ['maker', 'lot', 'spec'].map(function (f) {
      var c = det.map[f];
      return [L.FIELDS.filter(function (x) { return x.id === f; })[0].label, c == null ? '연결 안 됨' : L.colName(c) + '열 「' + (det.headers[c] || '') + '」',
        c == null ? '-' : (L.columnSamples(it.aoa, det.headerRow, c, 3).join(' · ') || '(빈칸)')];
    });
    var swapBtn = mat.swap && det.map.maker != null && det.map.spec != null ? h('button', { type: 'button', class: 'btn btn-primary btn-sm', onclick: function () {
      setOverrides(it, { maker: det.map.spec, spec: det.map.maker });
    } }, '두 칸 바꾸기 (메이커 ↔ 규격)') : null;
    return h('div', { class: 'mat-check' + (mat.warn ? ' bad' : ''), 'data-file': it.fileName },
      h('h4', null, '소재 칸 확인 — 엑셀 열 → 항목'),
      mat.warn ? h('p', { class: 'alert warn' }, h('strong', null, mat.warn), ' ', swapBtn) : null,
      table(['항목', '엑셀 열', '예시 값'], rows, { num: [0, 0, 0] }),
      (det.notes || []).map(function (n) { return h('p', { class: 'note' }, n); }),
      it.savedOverrides ? h('p', { class: 'note' }, '이 머리행 모양에서 전에 고친 열 연결을 그대로 썼습니다.') : null,
      h('p', { class: 'note' }, '값이 맞지 않으면 아래 「열 연결 보기·고치기」에서 바로 고칠 수 있습니다(바꾸면 파일을 다시 읽고, 다음에 같은 모양의 파일을 넣을 때도 기억합니다).'));
  }

  /* ── ① 불량률 추이 ───────────────────────── */
  function pageTrend() {
    var head = pageHead('①', '불량률 추이 (일별 · 월별)', '작업량 대비 불량률을 월별·일별로 보고, 목표 불량률 선과 비교합니다.');
    if (!db.records.length) return [head, emptyNotice()];
    var ctx = context(), mt = metric(ctx), S0 = ctx.S;
    var o = L.aggregate(ctx.recs, mt);
    var months = L.trend(ctx, 'month', mt);
    var gap = o.cnt - o.qty * S0.target;
    var monthsIn = months.map(function (m) { return m.key; });
    if (!ui.month || monthsIn.indexOf(ui.month) < 0) ui.month = monthsIn[monthsIn.length - 1] || '';
    var days = L.trend(ctx, 'day', mt, ui.month);
    var mSel = h('select', { 'aria-label': '일별로 볼 달' });
    monthsIn.slice().reverse().forEach(function (k) { var op = h('option', { value: k }, L.monthLabel(k)); if (k === ui.month) op.selected = true; mSel.appendChild(op); });
    mSel.addEventListener('change', function () { ui.month = mSel.value; render(); });
    var R = refs(ctx, null);
    if (S0.current) R.push({ value: S0.current, label: '지금 수준(제출 글) ' + L.pctText(S0.current, 0), cls: 'now' });
    var typeTable = null;
    if (mt === 'all' && ctx.types.length) {
      typeTable = h('section', { class: 'card' }, h('h2', null, '불량 유형별'),
        table(['유형', '불량수량', '작업량 대비', '불량 중 비중'], ctx.types.map(function (t) {
          return [t, N(o.types[t] || 0), pct(o.qty ? (o.types[t] || 0) / o.qty : null), L.pctText(o.defect ? (o.types[t] || 0) / o.defect : null, 1)];
        }).concat(o.untyped ? [['유형 미기록', N(o.untyped), pct(o.untyped / o.qty), L.pctText(o.untyped / o.defect, 1)]] : []), { num: [0, 1, 1, 1] }),
        h('p', { class: 'note' }, '유형 칸은 기획서 10장 ② 확인 전이라, 실제 작업일보에 유형별 기록이 있는지에 따라 달라집니다.'));
    }
    return [head, scopeCard(ctx),
      h('div', { class: 'tiles' },
        tile(L.metricLabel(mt) + '률', pct(o.rate), '작업 ' + N(o.qty) + '개 중 ' + N(o.cnt) + '개', o.rate > S0.target ? 'over' : 'ok'),
        tile('목표', L.pctText(S0.target, 1), '설정에서 바꿀 수 있습니다'),
        tile('목표까지 줄일 불량', N(Math.max(0, Math.round(gap))) + '개', '이 기간 작업량 기준'),
        tile('목표를 넘은 달', months.filter(function (m) { return m.overTarget; }).length + ' / ' + months.length, period().from + ' ~ ' + period().to)),
      h('section', { class: 'card' }, h('h2', null, '월별'),
        h('div', { class: 'charts' },
          C.figure('월별 ' + L.metricLabel(mt) + '률', { type: 'line', labels: months.map(function (m) { return m.label; }), unit: '%2', refs: R,
            series: [{ name: '불량률', values: months.map(function (m) { return m.rate; }), color: 'var(--series-1)' }] }),
          C.figure('월별 작업수량(개)', { type: 'bar', labels: months.map(function (m) { return m.label; }), unit: '개',
            series: [{ name: '작업수량', values: months.map(function (m) { return m.qty; }), color: 'var(--series-plan)' }] })),
        table(['월', '작업 행', '작업일', '작업수량', '불량수량', '불량률', '목표 대비'], months.map(function (m) {
          return { cells: [L.monthLabel(m.key), m.rows, m.days, N(m.qty), N(m.cnt), pct(m.rate), m.overTarget ? '+' + L.round((m.rate - S0.target) * 100, 2) + '%p' : '목표 안'], tdCls: { 6: m.overTarget ? 'over' : '' } };
        }), { num: [0, 1, 1, 1, 1, 1, 1] })),
      h('section', { class: 'card' }, h('div', { class: 'page-head' }, h('h2', null, '일별 — ' + L.monthLabel(ui.month)), h('div', { class: 'no-print' }, fld('달 고르기', mSel))),
        C.figure('일별 ' + L.metricLabel(mt) + '률 (' + L.monthLabel(ui.month) + ')', { type: 'line', labels: days.map(function (d) { return d.label.slice(3); }), unit: '%2', refs: refs(ctx, null),
          series: [{ name: '불량률', values: days.map(function (d) { return d.rate; }), color: 'var(--series-1)' }] }),
        h('details', null, h('summary', null, '일별 표 (' + days.length + '일)'),
          table(['일자', '작업 행', '작업수량', '불량수량', '불량률'], days.map(function (d) {
            return { cells: [d.key, d.rows, N(d.qty), N(d.cnt), pct(d.rate)], tdCls: { 4: d.overTarget ? 'over' : '' } };
          }), { num: [0, 1, 1, 1, 1] })),
        h('p', { class: 'note' }, '하루 작업량이 적은 날은 불량 몇 개로도 불량률이 크게 흔들립니다. 판단은 월별·조건별 합계로 해 주세요.')),
      typeTable];
  }

  /* ── ② 항목별 불량률 표 ─────────────────── */
  var TABLE_DIMS = ['die', 'maker', 'lot', 'spec', 'model', 'worker'];   // 소재 칸은 작업일보 순서(메이커 · lot · 규격)대로
  function rowCells(r) {
    return { cls: r.small ? 'small' : null, cells: [r.label, r.rows, r.days, N(r.qty), N(r.cnt), pct(r.rate), r.lift == null ? '-' : L.round(r.lift, 2) + '배',
      r.excess == null ? '-' : (r.excess > 0 ? '+' : '') + N(Math.round(r.excess)), L.pctText(r.share, 1), r.small ? h('span', { class: 'small-mark' }, '표본 적음') : ''] };
  }
  var ROW_HEAD = ['이름', '행', '작업일', '작업수량', '불량수량', '불량률', '전체 대비', '기대보다 많은 불량', '불량 비중', '표본'];
  var ROW_NUM = [0, 1, 1, 1, 1, 1, 1, 1, 1, 0];
  function dimFigure(ctx, g, title) {
    return C.figure(title || g.label + '별 ' + L.metricLabel(g.metric) + '률', { type: 'hbar', labels: g.rows.map(function (r) { return r.label; }), unit: '%2', refs: refs(ctx, g.overall.rate),
      series: [{ name: '불량률', values: g.rows.map(function (r) { return r.rate; }), colors: barColors(g.rows, g.overall, ctx.S), color: 'var(--series-1)' }] });
  }
  function pageTable() {
    var head = pageHead('②', '금형 · 소재 · 모델 · 작업자별 불량률', '작업량 대비 불량률(작업량 가중)을 묶음별로 비교합니다. 작업량이 적은 묶음은 「표본 적음」으로 흐리게 표시합니다.',
      db.records.length ? h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { exportTables(); } }, '엑셀로 내려받기(모든 표)')) : null);
    if (!db.records.length) return [head, emptyNotice()];
    var ctx = context(), mt = metric(ctx);
    var dims = TABLE_DIMS.filter(function (d) { return L.dimAvailable(ctx, d); });
    if (dims.indexOf(ui.dim) < 0) ui.dim = dims[0];
    var g = L.groupBy(ctx, ui.dim, mt, ui.sort);
    var sortSel = h('select', { 'aria-label': '정렬' }, [['excess', '기대보다 많은 불량 순'], ['rate', '불량률 높은 순'], ['order', '이름 순']].map(function (x) {
      var o = h('option', { value: x[0] }, x[1]); if (x[0] === ui.sort) o.selected = true; return o;
    }));
    sortSel.addEventListener('change', function () { ui.sort = sortSel.value; render(); });
    var missing = TABLE_DIMS.filter(function (d) { return dims.indexOf(d) < 0; }).map(function (d) { return L.DIMS[d].label; });
    // 두 기준이 1:1 이면(예: 금형마다 소재규격이 하나) 두 표가 이름만 다르고 숫자가 똑같습니다 — 「바뀐 것 같다」로 보이지 않게 알려 줍니다(3차)
    var same = dims.filter(function (d) { return d !== ui.dim && L.DIMS[d].cat && L.sameGrouping(ctx, ui.dim, d); }).map(function (d) { return L.DIMS[d].label; });
    var cross = null;
    if (dims.indexOf('die') >= 0 && dims.indexOf('maker') >= 0) cross = crossCard(ctx, 'die', 'maker', mt, '금형 × 소재 메이커');
    return [head, scopeCard(ctx),
      h('section', { class: 'card' },
        h('div', { class: 'tabs', role: 'group', 'aria-label': '묶을 기준' }, dims.map(function (d) {
          return h('button', { type: 'button', 'aria-pressed': d === ui.dim ? 'true' : 'false', onclick: function () { ui.dim = d; render(); } }, L.DIMS[d].label);
        })),
        sourceNote(ui.dim),
        storedMaterialAlert(),
        h('div', { class: 'filters', style: 'margin-bottom:12px' }, fld('정렬', sortSel),
          h('p', { class: 'note span-all' }, '전체 ' + L.metricLabel(mt) + '률 ' + pct(g.overall.rate) + ' · 「기대보다 많은 불량」 = 그 묶음의 불량 − 작업수량 × 전체 불량률. 양수가 크면 전체 불량을 끌어올리는 묶음입니다.')),
        same.length ? h('p', { class: 'alert info' }, '이 표는 ' + same.join(' · ') + ' 표와 숫자가 똑같습니다. 이 기간에는 ' + g.label + ' 하나마다 짝이 되는 ' + same.join(' · ') + '도 하나뿐이라(같은 작업 묶음) 이름만 다르게 나옵니다. ' +
          (ui.dim === 'spec' || same.indexOf(L.DIMS.spec.label) >= 0 ? '소재규격이 금형·모델마다 정해져 있으면 규격별 차이와 금형별 차이를 이 표만으로는 가를 수 없습니다. 소재 쪽 차이는 「소재 메이커」「소재 로트」 표에서 봐 주세요.' : '')) : null,
        dimFigure(ctx, g),
        table(ROW_HEAD, g.rows.map(rowCells).concat([{ cls: 'sum total', cells: ['전체', g.overall.rows, g.overall.days, N(g.overall.qty), N(g.overall.cnt), pct(g.overall.rate), '1배', '0', '100%', ''] }]), { num: ROW_NUM }),
        g.missing.rows ? h('p', { class: 'note' }, g.label + ' 칸이 빈 행 ' + g.missing.rows + '개(작업 ' + N(g.missing.qty) + '개)는 이 표에서 빠졌습니다.') : null,
        h('ul', { class: 'note' },
          h('li', null, '표본 적음 = 작업수량 ' + N(ctx.S.minQty) + '개 미만 또는 ' + ctx.S.minRows + '행 미만(설정에서 바꿀 수 있음). 불량 몇 개로도 비율이 크게 흔들리니 결론에 쓰지 말아 주세요.'),
          h('li', null, '막대 색: 주황 = 전체의 ' + ctx.S.minLift + '배 이상, 파랑 = 그 밖, 회색 = 표본 적음. 초록 점선 = 목표, 회색 점선 = 이 기간 전체.'),
          missing.length ? h('li', null, '작업일보에 없는 칸: ' + missing.join(', ') + ' (있다면 가져오기의 열 연결에서 골라 주세요)') : null)),
      cross];
  }
  /* 지금 탭이 작업일보의 어느 칸에서 왔는지 — 「소재 메이커 ← G열 「메이커」 (파일 12개)」. 탭과 칸이 어긋나면 여기서 바로 보입니다 */
  function sourceNote(dim) {
    var cs = L.columnSources(db.files, dim);
    var parts = cs.list.map(function (x) { return x.text + (db.files.length > 1 ? ' (파일 ' + x.files + '개)' : ''); });
    var tail = [];
    if (cs.none) tail.push('이 칸이 없는 파일 ' + cs.none + '개');
    if (cs.unknown) tail.push('이전 판에서 가져와 칸 기록이 없는 파일 ' + cs.unknown + '개');
    return h('p', { class: 'note col-source' }, h('strong', null, L.DIMS[dim].label + ' ← '), parts.length ? parts.join(' · ') : '칸 기록 없음', tail.length ? ' — ' + tail.join(' · ') : '');
  }
  /* 쌓인 작업 행의 메이커 · 규격 값 모양이 반대면 ② 에 알림 — 두 칸이 다 차 있으면 「두 칸 바꾸기」(행 값을 맞바꾸고 그 머리행의 열 연결도 기억) */
  function storedMaterialAlert() {
    if (db.sample) return null;
    var mc = L.materialCheck(db.records);
    if (!mc.warn) return null;
    return h('div', { class: 'alert warn', id: 'materialAlert' }, h('strong', null, mc.warn), ' ',
      mc.swap ? h('button', { type: 'button', class: 'btn btn-sm btn-primary', onclick: function () {
        if (!confirm('쌓인 작업 행 ' + db.records.length + '개의 소재 메이커와 소재규격 값을 서로 바꿀까요? 같은 모양의 파일을 다음에 넣을 때도 바꾼 연결로 읽습니다.')) return;
        db.records = L.swapMaterial(db.records);
        db.colOverrides = db.colOverrides || {};
        (db.files || []).forEach(function (f) {
          if (!f.cols) return;
          var m = f.cols.maker, s2 = f.cols.spec;
          f.cols.maker = s2; f.cols.spec = m;
          if (!f.cols.maker) delete f.cols.maker; if (!f.cols.spec) delete f.cols.spec;
          if (f.sig && m && s2 && m.i != null && s2.i != null) db.colOverrides[f.sig] = Object.assign({}, db.colOverrides[f.sig], { maker: s2.i, spec: m.i });
        });
        save(); toast('소재 메이커와 소재규격을 바꿨습니다.'); render();
      } }, '두 칸 바꾸기') : h('a', { class: 'btn btn-sm', href: '#/import?reimport' }, '다시 가져오기 안내'));
  }
  function staleBanner() {
    var st = L.importStatus(db);
    if (!st.stale) return null;
    return h('section', { class: 'alert warn stale-import no-print', id: 'staleImport' },
      h('strong', null, '이전 판에서 가져온 자료입니다 — 작업일보를 다시 가져와 주세요. '),
      db.sample ? '쌓인 것은 예전 판의 샘플입니다(소재 로트 · 3차 항목이 없음). ' :
        '2026-10-27 에 소재 칸 머리(두 줄 머리 · 「소재 | 규격」처럼 「소재」 만 적힌 칸)를 읽는 방법을 고쳤는데, 그 전에 가져온 작업 행 ' + N(st.oldRows) + '개' + (st.oldFiles.length ? '(파일 ' + st.oldFiles.length + '개)' : '') + '는 소재 메이커 · 소재규격 값이 서로 바뀌어 들어가 있을 수 있습니다. 이미 쌓인 행은 원본 칸이 남아 있지 않아 다시 읽을 수 없습니다. ',
      db.sample ? h('button', { type: 'button', class: 'btn btn-sm', onclick: loadSample }, '샘플 다시 불러오기')
        : h('a', { class: 'btn btn-sm', href: '#/import?reimport' }, '다시 가져오기 안내'));
  }
  function crossCard(ctx, a, b, mt, title) {
    var ct = L.crossTab(ctx, a, b, mt), o = ct.overall;
    var rows = ct.rows.map(function (r) {
      var cls = {}, cells = [r.label];
      ct.cols.forEach(function (c, i) {
        var x = ct.cell(r.key, c.key);
        if (!x) { cells.push('-'); return; }
        cells.push(h('span', { title: '작업 ' + N(x.qty) + ' · 불량 ' + N(x.cnt) }, pct(x.rate), x.small ? h('span', { class: 'small-mark' }, ' 표본 적음') : null));
        if (!x.small && x.lift != null) cls[i + 1] = x.lift >= 1.6 ? 'heat-3' : x.lift >= ctx.S.minLift ? 'heat-2' : x.lift >= 1.1 ? 'heat-1' : x.lift < 0.9 ? 'low' : '';
      });
      return { cells: cells, tdCls: cls };
    });
    return h('section', { class: 'card' }, h('h2', null, title + ' 교차표'),
      table([L.DIMS[a].label + ' \\ ' + L.DIMS[b].label].concat(ct.cols.map(function (c) { return c.label; })), rows, { num: [0].concat(ct.cols.map(function () { return 1; })), cls: 'cross' }),
      h('p', { class: 'note' }, '칸 색이 진할수록 전체(' + pct(o.rate) + ')보다 높습니다. 초록 = 전체보다 10% 이상 낮음. 칸에 마우스를 올리면 작업·불량 수가 보입니다.'));
  }
  function exportTables() {
    var ctx = context(), mt = metric(ctx), wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.trendToAoa(L.trend(ctx, 'month', mt), '월별 ' + L.metricLabel(mt) + '률')), '월별');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.trendToAoa(L.trend(ctx, 'day', mt), '일별 ' + L.metricLabel(mt) + '률')), '일별');
    TABLE_DIMS.concat(['season', 'main', 'cushion', 'wear']).forEach(function (d) {
      if (!L.dimAvailable(ctx, d)) return;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.groupToAoa(L.groupBy(ctx, d, mt, L.DIMS[d].cat ? 'excess' : 'order'))), L.DIMS[d].label.replace(/[\\\/?*\[\]:]/g, '').slice(0, 28));
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.recordsToAoa(ctx.recs, ctx.types)), '행 목록');
    var p = period();
    XLSX.writeFile(wb, '불량률_분석표_' + p.from + '_' + p.to + '.xlsx');
  }

  /* ── ③ 4M · 계절 · 불량이 몰린 조건 ─────── */
  function hotCard(h0, overall, i) {
    var ov = L.overlapText(h0);
    return h('article', { class: 'hot' },
      h('div', { class: 'hot-top' }, h('span', { class: 'tag' }, (i + 1) + '위'), h0.m4.map(function (m) { return h('span', { class: 'tag m4' }, L.M4_SHORT[m]); }),
        h('strong', null, L.condText(h0.parts))),
      h('p', { class: 'ev' }, L.hotEvidence(h0, overall)),
      ov ? h('p', { class: 'note' }, '겹침: ' + ov + ' — 어느 쪽 때문인지 이 숫자만으로는 가를 수 없습니다.') : null,
      h0.same && h0.same.length ? h('p', { class: 'note' }, '같은 작업 묶음: ' + h0.same.join(', ')) : null);
  }
  function dimBox(ctx, d, mt) {
    var g = L.groupBy(ctx, d, mt, 'order');
    return h('div', { class: 'dim-box' }, h('h3', null, g.label),
      dimFigure(ctx, g, g.label + '별 ' + L.metricLabel(mt) + '률'),
      h('details', null, h('summary', null, '표 보기'),
        table(ROW_HEAD, g.rows.map(rowCells), { num: ROW_NUM })),
      g.missing.rows ? h('p', { class: 'note' }, '값이 없는 행 ' + g.missing.rows + '개는 빠졌습니다.') : null);
  }
  function pageCond() {
    var head = pageHead('③', '4M · 계절 분석 — 불량이 몰린 조건', '사람·설비/금형·소재·방법(4M)과 계절로 나눠 불량률을 보고, 불량이 평소보다 몰린 조건을 근거 숫자와 함께 찾습니다.');
    if (!db.records.length) return [head, emptyNotice()];
    var ctx = context(), mt = metric(ctx), an = L.analyze(ctx, mt), S0 = ctx.S;
    var out = [head, scopeCard(ctx),
      h('div', { class: 'tiles' },
        tile(L.metricLabel(mt) + '률', pct(an.overall.rate), '목표 ' + L.pctText(S0.target, 1)),
        tile('불량이 몰린 조건', an.hot.total + '건', an.hot.total > an.hot.list.length ? '위 ' + an.hot.list.length + '건 표시' : '기준을 넘은 것'),
        tile('금형 보수 이력', db.maint.length + '건', db.maint.length ? '보수 전후·누적 작업량 분석' : '설정에서 넣어 주세요'),
        tile('계절 차이', (function () { var r = an.dims.season ? an.dims.season.rows.filter(function (x) { return !x.small; }) : []; if (!r.length) return '-'; var hi = r.slice().sort(function (a, b) { return b.rate - a.rate; })[0]; return hi.label + ' ' + pct(hi.rate); })(), '가장 높은 계절')),
      h('section', { class: 'card' }, h('h2', null, '불량이 몰린 조건 (기대보다 많은 불량 순)'),
        an.hot.list.length ? h('div', { class: 'hot-list' }, an.hot.list.map(function (x, i) { return hotCard(x, an.hot.overall, i); }))
          : h('p', { class: 'note' }, '기준을 넘은 조건이 없습니다. 기간을 넓히거나 설정의 기준을 바꿔 보세요.'),
        h('ul', { class: 'note' },
          h('li', null, '기준: 작업 ' + N(S0.minQty) + '개·' + S0.minRows + '행 이상이면서, 불량률이 전체의 ' + S0.minLift + '배 이상이고, 작업량을 생각해도(95% 신뢰구간 하한) 전체보다 높은 조건. 두 조건 조합은 각 조건 하나일 때보다 ' + S0.pairGain + '배 이상 높을 때만 따로 보여 줍니다.'),
          h('li', null, '「기대보다 많은 불량」 = 그 조건의 불량 − 작업수량 × 전체 불량률. 이 수가 크면 그 조건을 고칠 때 줄어들 불량이 많습니다.'),
          h('li', null, '숫자는 「함께 나타난다」는 것을 보여 줄 뿐 원인을 확정하지 않습니다. 소재 성적서·압력 설정값·금형 상태를 현장에서 확인해 주세요.')))];
    if (L.dimAvailable(ctx, 'maker')) out.push(crossCard(ctx, 'season', 'maker', mt, '계절 × 소재 메이커'));
    L.M4.forEach(function (grp) {
      var dims = grp.dims.filter(function (d) { return d !== 'month' && L.dimAvailable(ctx, d); });
      var sec = h('section', { class: 'card m4-sec' }, h('h2', null, grp.label));
      if (!dims.length) sec.appendChild(h('p', { class: 'note' }, '이 묶음에 해당하는 칸이 작업일보에 없습니다.'));
      else sec.appendChild(h('div', { class: 'dim-grid' }, dims.map(function (d) { return dimBox(ctx, d, mt); })));
      if (grp.id === 'Machine') {
        if (!db.maint.length) sec.appendChild(h('p', { class: 'alert info' }, '금형 보수·교체 이력을 설정 화면에 넣으면 「금형 누적 작업량(보수 후)」 구간과 보수 전후 비교가 나타납니다(기획서 10장 ③).'));
        if (S0.mainMin == null && S0.mainMax == null && S0.cushionMin == null && S0.cushionMax == null) sec.appendChild(h('p', { class: 'note' }, '압력 정상 범위를 설정에 넣으면 「정상 범위 아래·안·위」로도 나눠 봅니다(기획서 10장 ⑤).'));
      }
      out.push(sec);
    });
    out.push(maintCard(ctx, an));
    return out;
  }
  function maintCard(ctx, an) {
    if (!db.maint.length) return null;
    var S0 = ctx.S;
    var dies = an.dies.filter(function (d) { return d.curve.length; });
    return h('section', { class: 'card' }, h('h2', null, '금형 보수 · 교체 — 전후 비교와 누적 작업량'),
      table(['금형', '일자', '구분', '전 ' + S0.windowDays + '일 불량률', '후 ' + S0.windowDays + '일 불량률', '변화', '비고'], an.maint.map(function (m) {
        return { cls: m.enough ? null : 'small', cells: [m.die, m.date, m.kind, m.before.rows ? pct(m.before.rate) + ' (작업 ' + N(m.before.qty) + ')' : '기록 없음', m.after.rows ? pct(m.after.rate) + ' (작업 ' + N(m.after.qty) + ')' : '기록 없음',
          m.change == null ? '-' : (m.change > 0 ? '+' : '') + L.round(m.change * 100, 2) + '%p', (m.enough ? '' : '표본 적음 · ') + (m.note || '')], tdCls: { 5: m.change < 0 ? '' : m.change > 0 ? 'over' : '' } };
      }), { num: [0, 0, 0, 1, 1, 1, 0] }),
      dies.map(function (d) {
        return h('div', { class: 'dim-box', style: 'margin-top:14px' }, h('h3', null, d.die + ' — 보수 후 누적 작업량 구간별 불량률'),
          C.figure(d.die + ' 누적 작업량별 불량률', { type: 'hbar', labels: d.curve.map(function (c) { return c.label; }), unit: '%2', refs: refs(ctx, null),
            series: [{ name: '불량률', values: d.curve.map(function (c) { return c.rate; }), colors: d.curve.map(function (c) { return c.small ? 'var(--series-plan)' : d.threshold != null && c.lo >= d.threshold ? 'var(--series-2)' : 'var(--series-1)'; }), color: 'var(--series-1)' }] }),
          h('p', { class: 'note' }, d.threshold != null
            ? '참고: 과거 기록에서 보수 후 누적 ' + N(d.threshold) + '개 구간부터 불량률이 보수 직후(' + pct(d.fresh) + ')의 ' + S0.minLift + '배 이상이었습니다. 지금은 마지막 ' + d.lastMaint.kind + '(' + d.lastMaint.date + ') 뒤 누적 ' + N(d.cumNow) + '개입니다.'
            : '보수 직후보다 뚜렷하게 오르는 구간은 보이지 않습니다. 지금 누적 ' + N(d.cumNow) + '개(마지막 ' + d.lastMaint.kind + ' ' + d.lastMaint.date + ').'));
      }),
      h('p', { class: 'note' }, '보수 시점 「예측」은 이력이 더 쌓인 뒤 2단계에서 다룹니다. 지금은 과거에 어디서부터 올랐는지를 보여 주는 참고값입니다.'));
  }

  /* ── ④ AI 프롬프트 ───────────────────────── */
  function aiKey(ctx, mt) { return ctx.from + '~' + ctx.to + '|' + mt; }
  function pageAi() {
    var head = pageHead('④', 'AI 분석 프롬프트', '계산된 사실만 담은 프롬프트를 만들어 회사가 허용한 AI 에 붙여 넣고, 받은 답을 보고서에 씁니다.');
    if (!db.records.length) return [head, emptyNotice()];
    var ctx = context(), mt = metric(ctx), an = L.analyze(ctx, mt);
    var mask = db.mask !== false;
    if (mask) db.aliases = L.assignAllAliases(db.aliases, db.records);
    var info = L.buildAiPromptInfo(an, mask ? db.aliases : null);
    var prompt = info.prompt, key = aiKey(ctx, mt);
    function back(t) { return mask ? L.unmaskText(t, db.aliases) : t; }
    var answer = h('textarea', { rows: 10, placeholder: 'AI 답을 여기에 붙여 넣어 주세요. 별칭(금형1·메이커2 등)은 원래 이름으로 되돌려 보고서의 「AI 분석 의견」에 넣습니다.' });
    answer.value = (db.aiNotes || {})[key] || '';
    answer.addEventListener('input', function () { db.aiNotes[key] = back(answer.value); save(); });
    answer.addEventListener('change', function () { var b = back(answer.value); if (b !== answer.value) { answer.value = b; toast('답의 별칭을 원래 이름으로 되돌렸습니다.'); } });
    var promptBox = h('textarea', { rows: 14, readonly: true, class: 'mono' }); promptBox.value = prompt;
    var maskBox = h('input', { type: 'checkbox' }); maskBox.checked = mask;
    maskBox.addEventListener('change', function () { db.mask = maskBox.checked; save(); render(); });
    var mapRows = [];
    Object.keys(info.map).forEach(function (cat) {
      Object.keys(info.map[cat]).forEach(function (v) { mapRows.push([info.map[cat][v], v, L.MASK_CATS[cat]]); });
    });
    mapRows.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    return [head, scopeCard(ctx),
      h('section', { class: 'card', id: 'ai' }, h('h2', null, 'AI 분석 (반자동 기본)'),
        h('ol', { class: 'steps' },
          h('li', null, '아래 프롬프트를 복사해 회사가 허용한 AI(ChatGPT 등)에 붙여 넣어 주세요. 도구가 계산한 숫자(불량률·조건·보수 전후)만 들어 있고, 비고 같은 자유 글은 넣지 않았습니다.'),
          h('li', null, '받은 답을 「AI 답」 칸에 붙여 넣으면 별칭이 원래 이름으로 돌아가 ⑤ 보고서에 들어갑니다. 답은 꼭 검토해 주세요 — 원인은 모두 「추정」입니다.')),
        h('div', { class: 'mask-box' },
          h('label', { class: 'opt' }, maskBox, h('span', null, h('span', { class: 't' }, '이름 가리기(외부 AI 에 금형·모델·소재 메이커·작업자·로트 이름을 보내지 않음)'),
            h('span', { class: 's' }, '「금형1」「메이커2」「작업자3」 같은 별칭으로 바꿉니다. 대응표는 이 브라우저에만 있고, 같은 이름은 늘 같은 별칭입니다.'))),
          mask && info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름이 프롬프트에 남아 있습니다: ' + info.leaks.join(', ') + ' — 보내기 전에 지워 주세요.') : null,
          mask ? h('p', { class: 'note' }, info.leaks.length ? '' : '확인: 프롬프트에 실제 금형·모델·메이커·작업자 이름이 없습니다. 소재규격·압력 값·수량은 그대로 들어갑니다.')
            : h('p', { class: 'alert info' }, '지금은 이름이 그대로 들어갑니다. 회사 보안 기준을 확인해 주세요.'),
          mask && mapRows.length ? h('details', null, h('summary', null, '별칭 대응표 — 이 프롬프트에 쓴 ' + mapRows.length + '개 (이 PC 에만 있음)'),
            table(['별칭', '원래 이름', '구분'], mapRows, { cls: 'alias' }),
            h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: function () { copyText(mapRows.map(function (r) { return r.join('\t'); }).join('\n')); } }, '대응표 복사(엑셀 붙여넣기용)'))) : null),
        promptBox,
        h('div', { class: 'btn-row', style: 'margin:10px 0' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(prompt); } }, '프롬프트 복사'),
          AIP.sendButton('설정한 AI 서버로 보내기(선택)', function () { return prompt; }, function (text) { text = back(text); answer.value = text; db.aiNotes[key] = text; save(); toast('AI 답을 받았습니다(별칭은 원래 이름으로 되돌림). 검토 후 보고서에 씁니다.'); },
            { toast: toast, settingsHref: '#/settings', system: '너는 Deep drawing 공정 품질 분석 보조야. 주어진 사실 밖의 숫자를 만들지 마.' })),
        fld('AI 답 (' + ctx.from + ' ~ ' + ctx.to + ', ' + L.metricLabel(mt) + ')', answer, '기간·보는 불량마다 따로 저장됩니다(이 브라우저).'))];
  }

  /* ── ⑤ 보고서 ───────────────────────────── */
  function pageReport() {
    var head0 = pageHead('⑤', '불량률 분석 보고서');
    if (!db.records.length) return [head0, emptyNotice()];
    var ctx = context(), mt = metric(ctx), an = L.analyze(ctx, mt);
    var ai = (db.aiNotes || {})[aiKey(ctx, mt)];
    var text = L.buildReport(an, ai);
    var paper = reportPaper(an, ai);
    return [
      pageHead('⑤', '불량률 분석 보고서', '분석 결과로 보고서 초안을 만듭니다. 고쳐 쓸 곳을 확인한 뒤 인쇄하거나 Word·엑셀로 내려받아 주세요.',
        h('div', { class: 'btn-row no-print' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { window.print(); } }, '인쇄 / PDF'),
          h('button', { type: 'button', class: 'btn', onclick: function () { exportWord(an, paper); } }, 'Word 로'),
          h('button', { type: 'button', class: 'btn', onclick: function () { exportReportXlsx(an, text, ctx); } }, '엑셀로'),
          h('button', { type: 'button', class: 'btn', onclick: function () { copyText(text); } }, '글자로 복사'))),
      scopeCard(ctx),
      h('p', { class: 'note no-print' }, ai ? 'AI 분석 의견은 ④ 화면에서 붙여 넣은 답입니다.' : 'AI 분석 의견을 넣으려면 ④ 화면에서 AI 답을 붙여 넣어 주세요(선택).'),
      h('section', { class: 'card paper' }, paper)];
  }
  function reportPaper(an, ai) {
    var S0 = an.S, o = an.overall;
    function dimTable(d) {
      var g = an.dims[d]; if (!g) return null;
      return [h('h3', null, g.label + '별'), table(['이름', '작업수량', '불량수량', '불량률', '전체 대비', '표본'], g.rows.slice(0, 10).map(function (r) {
        return { cls: r.small ? 'small' : null, cells: [r.label, N(r.qty), N(r.cnt), pct(r.rate), L.round(r.lift, 2) + '배', r.small ? '표본 적음' : ''] };
      }), { num: [0, 1, 1, 1, 1, 0] })];
    }
    return h('div', { class: 'rp' },
      h('h1', null, 'Deep drawing 불량률 분석 보고'),
      h('p', { class: 'note' }, an.from + ' ~ ' + an.to + ' · ' + L.metricLabel(an.metric) + ' · 작성: 구매팀 · 자동 생성 초안(검토 후 사용)'),
      h('h2', null, '1. 총괄'),
      table(['작업수량', '불량수량', '불량률', '목표', '목표까지 줄일 불량', '목표를 넘은 달'],
        [[N(o.qty), N(o.cnt), pct(o.rate), L.pctText(S0.target, 1), N(Math.max(0, Math.round(an.gap))), an.months.filter(function (m) { return m.overTarget; }).length + ' / ' + an.months.length]], { num: [1, 1, 1, 1, 1, 1] }),
      an.metric === 'all' && an.types.length ? h('p', null, '유형별: ' + an.types.map(function (t) { return t + ' ' + N(an.all.types[t] || 0) + '개'; }).join(', ') + (an.all.untyped ? ', 유형 미기록 ' + N(an.all.untyped) + '개' : '')) : null,
      h('h2', null, '2. 월별 불량률'),
      table(['월', '작업수량', '불량수량', '불량률', '목표 대비'], an.months.map(function (m) {
        return [L.monthLabel(m.key), N(m.qty), N(m.cnt), pct(m.rate), m.overTarget ? '초과' : '목표 안'];
      }), { num: [0, 1, 1, 1, 0] }),
      h('h2', null, '3. 금형 · 소재별 불량률 (작업량 가중)'),
      ['die', 'maker', 'lot', 'spec', 'model'].map(dimTable),
      h('h2', null, '4. 불량이 몰린 조건 (4M · 계절)'),
      an.hot.list.length ? h('ol', null, an.hot.list.map(function (x) {
        var ov = L.overlapText(x);
        return h('li', null, h('strong', null, '[' + L.m4Text(x.m4) + '] ' + L.condText(x.parts)), ': ' + L.hotEvidence(x, an.hot.overall) + (ov ? ' (겹침: ' + ov + ')' : ''));
      })) : h('p', null, '기준을 넘은 조건이 없습니다.'),
      h('p', { class: 'note' }, '상관을 보여 줄 뿐 원인을 확정하지 않습니다. 현장 확인이 필요합니다.'),
      h('h2', null, '5. 금형 보수 · 교체'),
      an.maint.length ? table(['금형', '일자', '구분', '전 ' + S0.windowDays + '일', '후 ' + S0.windowDays + '일', '변화'], an.maint.map(function (m) {
        return [m.die, m.date, m.kind, m.before.rows ? pct(m.before.rate) : '기록 없음', m.after.rows ? pct(m.after.rate) : '기록 없음', m.change == null ? '-' : (m.change > 0 ? '+' : '') + L.round(m.change * 100, 2) + '%p'];
      }), { num: [0, 0, 0, 1, 1, 1] }) : h('p', null, '보수·교체 이력이 없어 전후 비교를 하지 않았습니다.'),
      an.dies.filter(function (d) { return d.threshold != null; }).map(function (d) {
        return h('p', null, '참고: ' + d.die + ' 은 과거 보수 후 누적 ' + N(d.threshold) + '개 구간부터 불량률이 보수 직후의 ' + S0.minLift + '배 이상이었습니다. 지금 누적 ' + N(d.cumNow) + '개.');
      }),
      h('h2', null, '6. 압력 작업표준 대비 · 설비(Press) 수선'),
      an.std.has ? [
        h('p', null, an.std.label + '별 기준 ± 허용보다 낮은 작업(설비보전·공무팀 전파 대상): Main ' + an.alerts.filter(function (a) { return a.kind === 'main'; }).length + '건 · 쿠션 ' + an.alerts.filter(function (a) { return a.kind === 'cushion'; }).length + '건' + (an.std.zeroShare >= 0.9 ? ' — 압력 칸 값이 거의 모두 기준과 같아, 작업일보에 기준(설정) 값을 적은 것으로 보입니다.' : '')),
        ['cushionStd', 'mainStd'].map(dimTable)] : h('p', null, '모델(금형)별 압력 작업표준을 넣지 않아 기준 대비를 보지 않았습니다.'),
      an.repairs.length ? table(['설비', '일자', '구분', '전 ' + S0.windowDays + '일', '후 ' + S0.windowDays + '일', '압력 저하 행 전 → 후', '내용'], an.repairs.map(function (m) {
        return [m.equip, m.date, m.kind, m.before.rows ? pct(m.before.rate) : '기록 없음', m.after.rows ? pct(m.after.rate) : '기록 없음', m.before.lowRows + ' → ' + m.after.lowRows, m.note];
      }), { num: [0, 0, 0, 1, 1, 1, 0] }) : h('p', null, '설비 수선 이력이 없습니다.'),
      ai ? [h('h2', null, '7. AI 분석 의견 (검토 후 사용)'), h('div', { class: 'pre' }, ai)] : null);
  }
  function exportWord(an, paper) {
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>불량률 분석 보고</title>' +
      '<style>body{font-family:"맑은 고딕",sans-serif;font-size:10.5pt} table{border-collapse:collapse;margin:6px 0} th,td{border:1px solid #999;padding:3px 6px;font-size:9.5pt} th{background:#eef2f7} h1{font-size:16pt} h2{font-size:12.5pt;margin-top:14pt} h3{font-size:11pt} .pre{white-space:pre-wrap}</style></head><body>' +
      paper.innerHTML + '</body></html>';
    download('불량률_분석보고_' + an.from + '_' + an.to + '.doc', new Blob(['﻿' + html], { type: 'application/msword' }));
  }
  function exportReportXlsx(an, text, ctx) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(text.split('\n').map(function (s) { return [s]; })), '보고서');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.trendToAoa(an.months, '월별 ' + L.metricLabel(an.metric) + '률')), '월별');
    ['die', 'maker', 'lot', 'spec', 'model', 'worker', 'season', 'cushion', 'main', 'cushionStd', 'mainStd', 'wear'].forEach(function (d) {
      if (an.dims[d]) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.groupToAoa(an.dims[d])), L.DIMS[d].label.replace(/[\\\/?*\[\]:]/g, '').slice(0, 28));
    });
    var hot = [['순위', '4M', '조건', '작업수량', '불량수량', '불량률(%)', '전체 대비(배)', '기대보다 많은 불량', '겹침']];
    an.hot.list.forEach(function (x, i) { hot.push([i + 1, L.m4Text(x.m4), L.condText(x.parts), x.agg.qty, x.agg.cnt, L.round(x.agg.rate * 100, 2), L.round(x.agg.lift, 2), Math.round(x.agg.excess), L.overlapText(x)]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hot), '불량이 몰린 조건');
    var mt = [['금형', '일자', '구분', '전 작업수량', '전 불량률(%)', '후 작업수량', '후 불량률(%)', '변화(%p)', '표본']];
    an.maint.forEach(function (m) { mt.push([m.die, m.date, m.kind, m.before.qty, m.before.rate == null ? '' : L.round(m.before.rate * 100, 2), m.after.qty, m.after.rate == null ? '' : L.round(m.after.rate * 100, 2), m.change == null ? '' : L.round(m.change * 100, 2), m.enough ? '' : '표본 적음']); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mt), '보수 전후');
    if (an.std.has) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.stdToAoa(an.std)), '압력 기준 대비');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.alertsToAoa(an.alerts)), '압력 저하 전파');
    }
    if (an.repairs.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(repairAoa(an)), '설비 수선 전후');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.recordsToAoa(ctx.recs, ctx.types)), '행 목록');
    XLSX.writeFile(wb, '불량률_분석보고_' + an.from + '_' + an.to + '.xlsx');
  }

  /* ── ⑥ 압력 작업표준 · 압력 저하 전파 · 설비 수선 (3차, 기획서 11장) ── */
  function stdTable(st, p, title) {
    var list = st[p];
    if (!list.length) return null;
    return [h('h3', null, title), table([st.label, '기준', '허용(±)', '행', '실제 평균', '가장 낮음 ~ 높음', '평균 차이(실제 − 기준)', '허용 넘게 낮음', '허용 넘게 높음'], list.map(function (x) {
      return { cells: [x.key, L.fmtNum(x.std), x.tol ? '±' + L.fmtNum(x.tol) : '0 (비어 있음)', x.n, L.fmtNum(L.round(x.mean, 1)), L.fmtNum(x.min) + ' ~ ' + L.fmtNum(x.max), (x.devMean > 0 ? '+' : '') + L.fmtNum(L.round(x.devMean, 1)),
        x.low ? x.low + '행 (불량률 ' + pct(x.lowQty ? x.lowCnt / x.lowQty : null) + ')' : '0', x.high || '0'], tdCls: { 7: x.low ? 'over' : '' } };
    }), { num: [0, 1, 1, 1, 1, 1, 1, 1, 1] })];
  }
  function devDistCard(ctx, by, dev, title) {
    if (!L.dimAvailable(ctx, dev)) return null;
    var ct = L.crossTab(ctx, by, dev, 'all');
    return h('div', { class: 'dim-box' }, h('h3', null, title),
      table([L.DIMS[by].label + ' \\ ' + L.DIMS[dev].label].concat(ct.cols.map(function (c) { return c.label; })), ct.rows.map(function (r) {
        var cls = {};
        var cells = [r.label].concat(ct.cols.map(function (c, i) {
          var x = ct.cell(r.key, c.key);
          if (!x) return '-';
          if (+c.key < 0) cls[i + 1] = 'heat-1';
          return h('span', { title: '작업 ' + N(x.qty) + ' · 불량률 ' + pct(x.rate) }, String(x.rows));
        }));
        return { cells: cells, tdCls: cls };
      }), { num: [0].concat(ct.cols.map(function () { return 1; })), cls: 'cross' }),
      h('p', { class: 'note' }, '칸 숫자 = 작업 행 수(마우스를 올리면 작업수량·불량률). 0 왼쪽(음수)이 기준보다 낮게 친 작업입니다.'));
  }
  function suggestStdsNow() {
    db.stds = L.suggestStds(db.records, db.stds.by, db.stds); save();
    toast('작업일보 압력 값의 가운데값으로 ' + db.stds.rows.length + '개 기준 초안을 만들었습니다. 허용 폭(±)은 직접 넣어 주세요.'); render();
  }
  function pagePress() {
    var head = pageHead('⑥', '압력 작업표준 · 설비 수선', '모델(금형)마다 다른 Main·쿠션 압력 기준과 실제 값을 비교하고, 기준보다 떨어진 작업을 설비보전·공무팀 전파 목록으로 만듭니다. 설비(Press) 수선 전후도 비교합니다.',
      db.records.length ? h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { exportPress(); } }, '엑셀로 내려받기(기준 대비 · 전파 목록 · 수선 전후)')) : null);
    if (!db.records.length) return [head, emptyNotice()];
    var ctx = context(), mt = metric(ctx), an = L.analyze(ctx, mt), st = an.std, S0 = ctx.S;
    var out = [head, scopeCard(ctx)];
    var mainLow = an.alerts.filter(function (a) { return a.kind === 'main'; }).length, cuLow = an.alerts.length - mainLow;
    out.push(h('div', { class: 'tiles' },
      tile('Main 압력 저하', st.has ? mainLow + '건' : '-', '기준 − 허용보다 낮은 작업'),
      tile('쿠션 압력 저하', st.has ? cuLow + '건' : '-', '기준 − 허용보다 낮은 작업'),
      tile('작업표준', db.stds.rows.length + '개', (db.stds.by === 'die' ? '금형' : '작업모델') + '별 기준' + (st.noStd.length ? ' · 기준 없음 ' + st.noStd.length + '개' : '')),
      tile('설비 수선 이력', db.repairs.length + '건', db.repairs.length ? '수선 전후 비교' : '설정에서 넣어 주세요')));
    if (!st.has) {
      out.push(h('section', { class: 'card' }, h('h2', null, '작업표준을 먼저 넣어 주세요'),
        h('p', { class: 'alert info' }, '모델(또는 금형)마다 Main·쿠션 압력의 기준과 허용 폭(±)을 넣으면, 행마다 「실제 − 기준」을 계산해 저하를 찾습니다. 작업일보에 「Main 압력 기준」 같은 기준 칸이 있으면 그 칸을 먼저 씁니다.'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: suggestStdsNow }, '작업일보에서 기준 초안 만들기(모델별 가운데값)'),
          h('a', { class: 'btn', href: '#/settings' }, '설정에서 직접 넣기 · 엑셀 붙여 넣기'))));
    } else {
      if (st.zeroShare >= 0.9) out.push(h('p', { class: 'alert info' }, '압력 칸 값의 ' + L.pctText(st.zeroShare, 0) + '가 기준과 똑같습니다. 지금 작업일보에 실제 압력이 아니라 기준(설정) 값을 적고 있는 것으로 보입니다. 실제 값(act)을 적기 시작하면 같은 화면에서 저하·차이가 바로 보입니다 — 실제 값을 따로 적는 칸(예: 「Main 압력 실제」)을 더해도 됩니다.'));
      if (st.noStd.length) out.push(h('p', { class: 'alert warn' }, '작업표준이 없는 ' + st.label + ': ' + st.noStd.join(', ') + ' — 이 작업은 기준 대비에서 빠졌습니다. 설정에서 넣어 주세요.'));
      out.push(h('section', { class: 'card' }, h('h2', null, st.label + '별 기준 대비 실제'),
        stdTable(st, 'main', 'Hydro press Main 압력'), stdTable(st, 'cushion', '쿠션 압력'),
        h('p', { class: 'note' }, '기준·허용 폭은 설정의 「압력 작업표준」 표(작업일보에 기준 칸이 있으면 그 값)입니다. 허용 폭이 비어 있으면 0 으로 보아 기준보다 조금만 낮아도 「허용 넘게 낮음」으로 셉니다.')));
      out.push(h('section', { class: 'card' }, h('h2', null, '기준 대비 차이 구간별 ' + L.metricLabel(mt) + '률'),
        h('div', { class: 'dim-grid' }, ['cushionStd', 'mainStd', 'cushionDev', 'mainDev'].filter(function (d) { return L.dimAvailable(ctx, d); }).map(function (d) { return dimBox(ctx, d, mt); })),
        h('p', { class: 'note' }, '「실제 − 기준」이 0 이면 기준 그대로 친 작업입니다. 모델마다 기준이 달라도 차이로 바꾸면 한 그래프에서 비교할 수 있습니다. 상관만 보여 줄 뿐 원인을 확정하지 않습니다.')));
      out.push(h('section', { class: 'card' }, h('h2', null, st.label + '별 「실제 − 기준」 분포'),
        h('div', { class: 'dim-grid' }, devDistCard(ctx, st.by, 'cushionDev', '쿠션 압력 − 기준'), devDistCard(ctx, st.by, 'mainDev', 'Main 압력 − 기준'))));
      if (L.dimAvailable(ctx, 'cushionStd')) out.push(crossCard(ctx, st.by, 'cushionStd', mt, st.label + ' × 쿠션 압력 작업표준 대비'));
      out.push(alertCard(ctx, an));
    }
    out.push(repairCard(ctx, an));
    return out;
  }
  function alertCard(ctx, an) {
    var text = L.alertText(an.alerts, { from: ctx.from, to: ctx.to });
    var box = h('textarea', { rows: 10, readonly: true, class: 'mono' }); box.value = text;
    return h('section', { class: 'card', id: 'alerts' }, h('h2', null, '설비보전 · 공무팀 전파 목록 — 압력 저하 ' + an.alerts.length + '건'),
      h('p', { class: 'note' }, 'Main 또는 쿠션 압력이 그 ' + an.std.label + '의 「기준 − 허용」보다 낮았던 작업입니다. 아래 글을 복사해 메신저·메일로 보내 주세요(이 도구는 직접 보내지 않습니다).'),
      box,
      h('div', { class: 'btn-row', style: 'margin:10px 0' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(text); } }, '전파 글 복사'),
        h('button', { type: 'button', class: 'btn', onclick: function () {
          var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.alertsToAoa(an.alerts)), '압력 저하 전파');
          XLSX.writeFile(wb, '압력저하_전파목록_' + ctx.from + '_' + ctx.to + '.xlsx');
        } }, '전파 목록 엑셀')),
      an.alerts.length ? h('details', null, h('summary', null, '목록 표로 보기 (' + an.alerts.length + '건)'),
        table(['일자', '항목', '작업모델', '금형', '기준', '실제', '실제 − 기준', '작업수량', '불량률'], an.alerts.map(function (a) {
          return [a.date, a.item, a.model, a.die, L.fmtNum(a.std) + (a.tol ? ' ±' + L.fmtNum(a.tol) : ''), L.fmtNum(a.act), L.fmtNum(L.round(a.dev, 1)), N(a.qty), pct(a.rate)];
        }), { num: [0, 0, 0, 0, 1, 1, 1, 1, 1] })) : null);
  }
  function repairCard(ctx, an) {
    var S0 = ctx.S, types = ctx.types.slice(0, 3);
    if (!db.repairs.length) return h('section', { class: 'card' }, h('h2', null, '설비(Press) 수선 전후'),
      h('p', { class: 'alert info' }, '설정 화면에 설비 수선·점검·부품교체 이력을 넣으면 수선 전후 ' + S0.windowDays + '일의 불량률·압력 저하 행 수·「실제 − 기준」을 비교합니다. 예: samples/설비_수선이력_샘플.xlsx'));
    function dv(v) { return v == null ? '-' : (v > 0 ? '+' : '') + L.fmtNum(L.round(v, 1)); }
    return h('section', { class: 'card' }, h('h2', null, '설비(Press) 수선 전후 ' + S0.windowDays + '일'),
      table(['설비', '일자', '구분', '전 불량률', '후 불량률', '변화'].concat(types.map(function (t) { return t + ' 전 → 후'; })).concat(['압력 저하 행 전 → 후', '쿠션 − 기준 전 → 후', 'Main − 기준 전 → 후', '내용']), an.repairs.map(function (m) {
        return { cls: m.enough ? null : 'small', cells: [m.equip, m.date, m.kind, m.before.rows ? pct(m.before.rate) + ' (작업 ' + N(m.before.qty) + ')' : '기록 없음', m.after.rows ? pct(m.after.rate) + ' (작업 ' + N(m.after.qty) + ')' : '기록 없음',
          m.change == null ? '-' : (m.change > 0 ? '+' : '') + L.round(m.change * 100, 2) + '%p'].concat(types.map(function (t) {
            return L.pctText(m.before.qty ? (m.before.types[t] || 0) / m.before.qty : null, 2) + ' → ' + L.pctText(m.after.qty ? (m.after.types[t] || 0) / m.after.qty : null, 2);
          })).concat([m.before.lowRows + ' → ' + m.after.lowRows, dv(m.before.cushionDev) + ' → ' + dv(m.after.cushionDev), dv(m.before.mainDev) + ' → ' + dv(m.after.mainDev), (m.matched ? '' : '작업일보에 없는 설비 · ') + (m.enough ? '' : '표본 적음 · ') + (m.note || '')]),
          tdCls: { 5: m.change < 0 ? '' : m.change > 0 ? 'over' : '' } };
      }), { num: [0, 0, 0, 1, 1, 1].concat(types.map(function () { return 1; })).concat([1, 1, 1, 0]) }),
      h('p', { class: 'note' }, '작업일보에 설비(호기) 칸이 있으면 같은 설비의 작업만, 없으면(프레스 한 대) 모든 작업으로 전후를 봅니다. 압력 저하 행·「실제 − 기준」은 작업표준이 있을 때만 셉니다.'));
  }
  function exportPress() {
    var ctx = context(), mt = metric(ctx), an = L.analyze(ctx, mt), wb = XLSX.utils.book_new();
    if (an.std.has) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.stdToAoa(an.std)), '기준 대비 실제');
      ['mainStd', 'cushionStd', 'mainDev', 'cushionDev'].forEach(function (d) { if (an.dims[d]) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.groupToAoa(an.dims[d])), L.DIMS[d].label.replace(/[\\\/?*\[\]:]/g, '').slice(0, 28)); });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.alertsToAoa(an.alerts)), '압력 저하 전파');
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(repairAoa(an)), '설비 수선 전후');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['작업모델/금형', 'Main 기준', 'Main 허용(±)', '쿠션 기준', '쿠션 허용(±)']].concat(db.stds.rows.map(function (r) { return [r.key, r.main, r.mainTol, r.cushion, r.cushionTol]; }))), '작업표준');
    XLSX.writeFile(wb, '압력기준_설비수선_' + ctx.from + '_' + ctx.to + '.xlsx');
  }
  function repairAoa(an) {
    var aoa = [['설비', '일자', '구분', '내용', '전 작업수량', '전 불량률(%)', '후 작업수량', '후 불량률(%)', '변화(%p)', '전 압력 저하 행', '후 압력 저하 행', '표본']];
    an.repairs.forEach(function (m) { aoa.push([m.equip, m.date, m.kind, m.note, m.before.qty, m.before.rate == null ? '' : L.round(m.before.rate * 100, 2), m.after.qty, m.after.rate == null ? '' : L.round(m.after.rate * 100, 2), m.change == null ? '' : L.round(m.change * 100, 2), m.before.lowRows, m.after.lowRows, m.enough ? '' : '표본 적음']); });
    return aoa;
  }

  /* ── 설정 ───────────────────────────────── */
  function pageSettings() {
    var s = settings();
    function num(v, mul, step) { var i = h('input', { type: 'number', min: 0, step: step || 'any', inputmode: 'decimal', value: v == null ? '' : L.round(v * (mul || 1), 4) }); return i; }
    var inp = {
      target: num(s.target, 100, 0.1), current: num(s.current, 100, 0.1),
      mainMin: num(s.mainMin), mainMax: num(s.mainMax), cushionMin: num(s.cushionMin), cushionMax: num(s.cushionMax),
      minQty: num(s.minQty, 1, 100), minRows: num(s.minRows, 1, 1), minLift: num(s.minLift, 1, 0.05), windowDays: num(s.windowDays, 1, 1)
    };
    function saveSettings() {
      var p = {}; Object.keys(inp).forEach(function (k) { p[k] = inp[k].value === '' ? null : Number(inp[k].value); });
      p.target = p.target == null ? null : p.target / 100; p.current = p.current == null ? null : p.current / 100;
      p.pairGain = s.pairGain;
      db.settings = L.cleanSettings(p); save(); toast('설정을 저장했습니다.'); render();
    }
    // 보수 이력
    var dies = []; db.records.forEach(function (r) { if (dies.indexOf(r.die) < 0) dies.push(r.die); }); dies.sort();
    var dl = h('datalist', { id: 'dieList' }, dies.map(function (d) { return h('option', { value: d }); }));
    var mDie = h('input', { list: 'dieList', placeholder: '예: ' + (dies[0] || 'DIE-01'), autocomplete: 'off' }), mDate = h('input', { type: 'date' });
    var mKind = h('select', null, h('option', { value: '보수' }, '보수'), h('option', { value: '교체' }, '교체')), mNote = h('input', { placeholder: '예: R부 연마' });
    function addMaint() {
      if (!mDie.value.trim() || !mDate.value) { toast('금형과 일자를 넣어 주세요.', true); return; }
      db.maint = L.cleanMaint(db.maint.concat([{ die: mDie.value, date: mDate.value, kind: mKind.value, note: mNote.value }])); save(); toast('보수 이력을 넣었습니다.'); render();
    }
    var paste = h('textarea', { rows: 5, placeholder: '금형\t일자\t구분\t내용\nDIE-01\t2026-03-13\t보수\tR부 연마' });
    function pasteMaint() {
      var r = L.parseMaintText(paste.value);
      if (!r.items.length) { toast('읽을 줄이 없습니다. 「금형 · 일자 · 구분 · 내용」 순서인지 확인해 주세요.', true); return; }
      var before = db.maint.length;
      db.maint = L.cleanMaint(db.maint.concat(r.items)); save();
      toast((db.maint.length - before) + '건 넣었습니다' + (r.bad.length ? ' · 못 읽은 줄 ' + r.bad.join(', ') : '') + '.'); render();
    }
    var maintFile = sheetFileInput(function (text) { paste.value = text; pasteMaint(); });
    // 압력 작업표준 (모델·금형별)
    var stdBy = h('select', null, h('option', { value: 'model' }, '작업모델'), h('option', { value: 'die' }, '금형'));
    stdBy.value = db.stds.by;
    stdBy.addEventListener('change', function () { db.stds = L.cleanStds({ by: stdBy.value, rows: db.stds.rows }); save(); render(); });
    var keyList = []; db.records.forEach(function (r) { var k = r[db.stds.by]; if (k && keyList.indexOf(k) < 0) keyList.push(k); }); keyList.sort();
    var sKey = h('input', { list: 'stdKeyList', placeholder: '예: ' + (keyList[0] || 'SMP-01'), autocomplete: 'off' });
    var sMain = num(null), sMainTol = num(null), sCu = num(null), sCuTol = num(null);
    function addStd() {
      if (!sKey.value.trim() || (sMain.value === '' && sCu.value === '')) { toast((db.stds.by === 'die' ? '금형' : '작업모델') + '과 기준 값을 넣어 주세요.', true); return; }
      var rows = db.stds.rows.filter(function (r) { return r.key !== sKey.value.trim(); });
      rows.push({ key: sKey.value.trim(), main: sMain.value, mainTol: sMainTol.value, cushion: sCu.value, cushionTol: sCuTol.value });
      db.stds = L.cleanStds({ by: db.stds.by, rows: rows }); save(); toast('작업표준을 넣었습니다.'); render();
    }
    var stdPaste = h('textarea', { rows: 5, placeholder: '작업모델\tMain 기준\tMain 허용\t쿠션 기준\t쿠션 허용\nSMP-01\t235\t15\t55\t8' });
    function pasteStds() {
      var r = L.parseStdText(stdPaste.value, db.stds.by);
      if (!r.stds.rows.length) { toast('읽을 줄이 없습니다. 「모델 · Main 기준 · Main 허용 · 쿠션 기준 · 쿠션 허용」 순서인지 확인해 주세요.', true); return; }
      var keys = r.stds.rows.map(function (x) { return x.key; });
      db.stds = L.cleanStds({ by: db.stds.by, rows: db.stds.rows.filter(function (x) { return keys.indexOf(x.key) < 0; }).concat(r.stds.rows) }); save();
      toast(r.stds.rows.length + '개 넣었습니다(같은 이름은 새 값으로)' + (r.bad.length ? ' · 못 읽은 줄 ' + r.bad.join(', ') : '') + '.'); render();
    }
    var stdFile = sheetFileInput(function (text) { stdPaste.value = text; pasteStds(); });
    function stdCell(row, k) {
      var i = num(row[k]);
      i.addEventListener('change', function () {
        var t = db.stds.rows.filter(function (x) { return x.key === row.key; })[0]; if (!t) return;
        t[k] = i.value === '' ? null : Number(i.value); db.stds = L.cleanStds(db.stds); save(); toast('작업표준을 고쳤습니다.');
      });
      i.setAttribute('aria-label', row.key + ' ' + k);
      return i;
    }
    var stdMissing = keyList.filter(function (k) { return !db.stds.rows.some(function (r) { return r.key === k; }); });
    // 설비(Press) 수선 이력
    var equips = []; db.records.forEach(function (r) { if (r.press && equips.indexOf(r.press) < 0) equips.push(r.press); });
    db.repairs.forEach(function (m) { if (equips.indexOf(m.equip) < 0) equips.push(m.equip); });
    var rEquip = h('input', { list: 'equipList', placeholder: '예: ' + (equips[0] || 'Hydro press'), autocomplete: 'off' }), rDate = h('input', { type: 'date' });
    var rKind = h('select', null, L.REPAIR_KINDS.map(function (k) { return h('option', { value: k }, k); })), rNote = h('input', { placeholder: '예: 쿠션 실린더 씰 교체' });
    function addRepair() {
      if (!rDate.value) { toast('일자를 넣어 주세요.', true); return; }
      db.repairs = L.cleanRepairs(db.repairs.concat([{ equip: rEquip.value, date: rDate.value, kind: rKind.value, note: rNote.value }])); save(); toast('설비 수선 이력을 넣었습니다.'); render();
    }
    var rPaste = h('textarea', { rows: 5, placeholder: '설비\t일자\t구분\t내용\nHydro press\t2026-05-22\t부품교체\t쿠션 실린더 씰 교체' });
    function pasteRepairs() {
      var r = L.parseRepairText(rPaste.value);
      if (!r.items.length) { toast('읽을 줄이 없습니다. 「설비 · 일자 · 구분 · 내용」 순서인지 확인해 주세요.', true); return; }
      var before = db.repairs.length;
      db.repairs = L.cleanRepairs(db.repairs.concat(r.items)); save();
      toast((db.repairs.length - before) + '건 넣었습니다' + (r.bad.length ? ' · 못 읽은 줄 ' + r.bad.join(', ') : '') + '.'); render();
    }
    var repairFile = sheetFileInput(function (text) { rPaste.value = text; pasteRepairs(); });
    var ctx = db.records.length ? context() : null;
    var rEff = ctx ? L.repairEffects(L.buildContext(db.records, { settings: db.settings, stds: db.stds, repairs: db.repairs }), 'all') : [];
    var eff = ctx ? L.maintEffects(L.buildContext(db.records, { settings: db.settings, maint: db.maint }), 'all') : [];
    function effOf(m) { return eff.filter(function (e) { return e.die === m.die && e.date === m.date; })[0]; }
    var unknownDies = db.maint.map(function (m) { return m.die; }).filter(function (d, i, a) { return a.indexOf(d) === i && dies.indexOf(d) < 0; });
    var backup = h('input', { type: 'file', accept: '.json' });
    backup.addEventListener('change', function () {
      var f = backup.files[0]; if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t); if (!o || !Array.isArray(o.records)) throw new Error('이 도구의 백업 파일이 아닙니다.');
        db = L.restoreDb(o); save(); toast('백업을 불러왔습니다(' + db.records.length + '행).'); render();
      }).catch(function (e) { toast('불러오지 못했습니다: ' + e.message, true); });
    });
    return [
      pageHead('설정', '목표 · 기준 · 금형 보수 · 압력 작업표준 · 설비 수선 · AI 연결 · 백업'),
      h('section', { class: 'card' }, h('h2', null, '목표와 분석 기준'),
        h('div', { class: 'form-grid' },
          fld('목표 불량률 (%)', inp.target, '제출 글: 2% 수준으로 낮추기'),
          fld('지금 불량률 (%, 비교용)', inp.current, '제출 글: 약 5%'),
          fld('「표본 적음」 작업수량 기준(개)', inp.minQty, '이보다 적으면 결론에서 뺍니다'),
          fld('Main 압력 정상 범위 — 아래', inp.mainMin, '비워 두면 구간으로만 봅니다(기획서 10장 ⑤)'),
          fld('Main 압력 정상 범위 — 위', inp.mainMax),
          fld('「표본 적음」 행 수 기준', inp.minRows),
          fld('쿠션 압력 정상 범위 — 아래', inp.cushionMin, '작업일보와 같은 단위로'),
          fld('쿠션 압력 정상 범위 — 위', inp.cushionMax),
          fld('「불량이 몰린 조건」 — 전체의 몇 배 이상', inp.minLift),
          fld('보수 전후 비교 기간(일)', inp.windowDays)),
        h('div', { class: 'btn-row', style: 'margin-top:12px' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: saveSettings }, '저장'),
          h('button', { type: 'button', class: 'btn', onclick: function () { db.settings = L.cleanSettings({}); save(); render(); } }, '기본값으로'))),
      h('section', { class: 'card' }, h('h2', null, '금형 보수 · 교체 이력 (' + db.maint.length + '건)'),
        h('p', { class: 'note' }, '보수·교체한 날을 넣으면 그날 작업부터 「보수 후 누적 작업량」을 0 부터 다시 세고, 보수 전후 ' + s.windowDays + '일 불량률을 비교합니다. 금형 이름은 작업일보의 금형 칸과 똑같이 적어 주세요.'),
        dl,
        h('div', { class: 'inline-form' }, fld('금형', mDie), fld('일자', mDate), fld('구분', mKind), fld('내용(선택)', mNote),
          h('button', { type: 'button', class: 'btn btn-primary', onclick: addMaint }, '넣기')),
        h('details', { style: 'margin-top:10px' }, h('summary', null, '엑셀에서 여러 줄 붙여 넣기'),
          h('p', { class: 'note' }, '엑셀에서 「금형 · 일자 · 구분 · 내용」 네 칸을 복사해 붙여 넣어 주세요(머리 줄은 있어도 됩니다). 예: samples/금형_보수이력_샘플.xlsx'),
          paste, h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn', onclick: pasteMaint }, '붙여 넣은 줄 넣기'), h('label', { class: 'btn' }, maintFile, '엑셀 파일로 넣기'))),
        unknownDies.length ? h('p', { class: 'alert warn' }, '작업일보에 없는 금형 이름: ' + unknownDies.join(', ') + ' — 이름이 작업일보와 같은지 확인해 주세요.') : null,
        db.maint.length ? table(['금형', '일자', '구분', '내용', '전 ' + s.windowDays + '일', '후 ' + s.windowDays + '일', ''], db.maint.map(function (m, i) {
          var e = effOf(m);
          return [m.die, m.date, m.kind, m.note, e && e.before.rows ? pct(e.before.rate) : '-', e && e.after.rows ? pct(e.after.rate) : '-',
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { db.maint.splice(i, 1); save(); render(); } }, '지우기')];
        }), { num: [0, 0, 0, 0, 1, 1, 0] }) : null,
        db.maint.length ? h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (confirm('보수 이력을 모두 지울까요?')) { db.maint = []; save(); render(); } } }, '이력 모두 지우기')) : null),
      h('section', { class: 'card', id: 'stds' }, h('h2', null, '압력 작업표준 — 모델 · 금형별 기준과 허용 폭 (' + db.stds.rows.length + '개)'),
        h('p', { class: 'note' }, '모델(또는 금형)마다 Main·쿠션 압력의 기준과 허용 폭(±)을 넣으면 ⑥ 화면에서 행마다 「실제 − 기준」을 보고, 「기준 − 허용」보다 낮은 작업을 설비보전·공무팀 전파 목록에 올립니다. 허용을 비우면 0 입니다. 작업일보에 기준 칸(「Main 압력 기준」 등)이 있으면 그 칸이 먼저입니다.'),
        h('div', { class: 'inline-form' }, fld('찾는 칸', stdBy, '작업일보의 이 칸 값으로 기준을 찾습니다')),
        h('datalist', { id: 'stdKeyList' }, keyList.map(function (k) { return h('option', { value: k }); })),
        h('div', { class: 'inline-form' }, fld(db.stds.by === 'die' ? '금형' : '작업모델', sKey), fld('Main 기준', sMain), fld('Main 허용(±)', sMainTol), fld('쿠션 기준', sCu), fld('쿠션 허용(±)', sCuTol),
          h('button', { type: 'button', class: 'btn btn-primary', onclick: addStd }, '넣기')),
        h('div', { class: 'btn-row', style: 'margin-top:8px' },
          db.records.length ? h('button', { type: 'button', class: 'btn', onclick: function () { if (!db.stds.rows.length || confirm('작업일보 압력 값의 가운데값으로 기준을 다시 채울까요? (허용 폭은 그대로 둡니다)')) suggestStdsNow(); } }, '작업일보에서 기준 초안 채우기(가운데값)') : null),
        h('details', { style: 'margin-top:10px' }, h('summary', null, '엑셀에서 여러 줄 붙여 넣기 · 엑셀 파일로 넣기'),
          h('p', { class: 'note' }, '「모델(또는 금형) · Main 기준 · Main 허용 · 쿠션 기준 · 쿠션 허용」 다섯 칸(머리 줄은 있어도 됩니다). 같은 이름은 새 값으로 바뀝니다. 예: samples/압력_작업표준_샘플.xlsx'),
          stdPaste, h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn', onclick: pasteStds }, '붙여 넣은 줄 넣기'), h('label', { class: 'btn' }, stdFile, '엑셀 파일로 넣기'))),
        stdMissing.length && db.stds.rows.length ? h('p', { class: 'alert warn' }, '작업일보에는 있는데 작업표준이 없는 ' + (db.stds.by === 'die' ? '금형' : '작업모델') + ': ' + stdMissing.join(', ')) : null,
        db.stds.rows.length ? table([db.stds.by === 'die' ? '금형' : '작업모델', 'Main 기준', 'Main 허용(±)', '쿠션 기준', '쿠션 허용(±)', ''], db.stds.rows.map(function (r, i) {
          return [r.key, stdCell(r, 'main'), stdCell(r, 'mainTol'), stdCell(r, 'cushion'), stdCell(r, 'cushionTol'),
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { db.stds.rows.splice(i, 1); save(); render(); } }, '지우기')];
        }), { cls: 'std-edit' }) : null,
        db.stds.rows.length ? h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (confirm('작업표준을 모두 지울까요?')) { db.stds = L.cleanStds({ by: db.stds.by }); save(); render(); } } }, '작업표준 모두 지우기')) : null),
      h('section', { class: 'card', id: 'repairs' }, h('h2', null, '설비(Press) 수선 이력 (' + db.repairs.length + '건)'),
        h('p', { class: 'note' }, '프레스 수선·점검·부품교체한 날을 넣으면 ⑥ 화면에서 전후 ' + s.windowDays + '일의 불량률(유형별)·압력 저하 행 수·「실제 − 기준」을 비교합니다. 작업일보에 설비(호기) 칸이 있으면 설비 이름을 그 칸과 똑같이 적어 주세요(없으면 모든 작업에 적용).'),
        h('datalist', { id: 'equipList' }, equips.map(function (d) { return h('option', { value: d }); })),
        h('div', { class: 'inline-form' }, fld('설비', rEquip), fld('일자', rDate), fld('구분', rKind), fld('내용(선택)', rNote),
          h('button', { type: 'button', class: 'btn btn-primary', onclick: addRepair }, '넣기')),
        h('details', { style: 'margin-top:10px' }, h('summary', null, '엑셀에서 여러 줄 붙여 넣기 · 엑셀 파일로 넣기'),
          h('p', { class: 'note' }, '「설비 · 일자 · 구분 · 내용」 네 칸(머리 줄은 있어도 됩니다). 예: samples/설비_수선이력_샘플.xlsx'),
          rPaste, h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn', onclick: pasteRepairs }, '붙여 넣은 줄 넣기'), h('label', { class: 'btn' }, repairFile, '엑셀 파일로 넣기'))),
        db.repairs.length ? table(['설비', '일자', '구분', '내용', '전 ' + s.windowDays + '일', '후 ' + s.windowDays + '일', '압력 저하 행 전 → 후', ''], db.repairs.map(function (m, i) {
          var e = rEff.filter(function (x) { return x.equip === m.equip && x.date === m.date && x.kind === m.kind; })[0];
          return [m.equip, m.date, m.kind, m.note, e && e.before.rows ? pct(e.before.rate) : '-', e && e.after.rows ? pct(e.after.rate) : '-', e ? e.before.lowRows + ' → ' + e.after.lowRows : '-',
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { db.repairs.splice(i, 1); save(); render(); } }, '지우기')];
        }), { num: [0, 0, 0, 0, 1, 1, 1, 0] }) : null,
        db.repairs.length ? h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (confirm('설비 수선 이력을 모두 지울까요?')) { db.repairs = []; save(); render(); } } }, '이력 모두 지우기')) : null),
      AIP.settingsCard({ toast: toast, onChange: render }),
      h('section', { class: 'card' }, h('h2', null, '백업'),
        h('p', { class: 'note' }, '쌓인 작업일보는 이 브라우저에만 있습니다. 다른 PC 로 옮기거나 브라우저 정리 전에 JSON 으로 받아 두세요. (파일에는 작업자 칸도 들어 있으니 보관에 주의해 주세요)'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('불량률분석_백업_' + localStamp(true) + '.json', new Blob([JSON.stringify(db)], { type: 'application/json' })); } }, 'JSON 백업 내려받기'),
          h('label', { class: 'btn' }, backup, 'JSON 백업 불러오기')))
    ];
  }

  /* ── 라우터 ─────────────────────────────── */
  var PAGES = [
    { id: 'import', no: '0', t: '데이터 가져오기', s: '작업일보 엑셀 여러 개', fn: pageImport },
    { id: 'trend', no: '①', t: '불량률 추이', s: '일별·월별·목표선', fn: pageTrend },
    { id: 'table', no: '②', t: '항목별 불량률', s: '금형·소재·모델·작업자', fn: pageTable },
    { id: 'cond', no: '③', t: '4M · 계절 분석', s: '불량이 몰린 조건', fn: pageCond },
    { id: 'ai', no: '④', t: 'AI 프롬프트', s: '사실만 · 이름 가림', fn: pageAi },
    { id: 'report', no: '⑤', t: '보고서', s: '인쇄·Word·엑셀', fn: pageReport },
    { id: 'press', no: '⑥', t: '압력 기준 · 설비', s: '작업표준·저하 전파·수선', fn: pagePress },
    { id: 'settings', no: '설', t: '설정', s: '목표·보수 이력·AI', fn: pageSettings }
  ];
  function route() { var id = (location.hash || '').replace(/^#\//, '').split('?')[0]; return PAGES.filter(function (p) { return p.id === id; })[0] || (db.records.length ? PAGES[3] : PAGES[0]); }
  function render() {
    var p = route();
    var nav = document.getElementById('nav'); nav.innerHTML = '';
    PAGES.forEach(function (x) {
      nav.appendChild(h('a', { href: '#/' + x.id, 'aria-current': x === p ? 'page' : null }, h('span', { class: 'no' }, x.no), h('span', { class: 't' }, x.t), h('span', { class: 's' }, x.s)));
    });
    var rg = fullRange();
    document.getElementById('dataChip').textContent = db.records.length ? '작업일보 ' + L.monthsOf(db.records).length + '개월 · ' + L.fmtNum(db.records.length) + '행 · ' + rg.from + ' ~ ' + rg.to : '데이터 없음';
    var b = document.getElementById('sampleBanner');
    b.hidden = !db.sample;
    b.textContent = '지금 보시는 것은 가상 샘플 12개월입니다(가짜 금형 DIE-…, 모델 SMP-…, 소재사A~C, 작업자A~E). 금형 DIE-B01 마모, 겨울 소재사B Crack, 쿠션 압력 65 이상 Crack 증가, 소재 로트 A-2607 Crack, 5월 쿠션 압력 저하(주름)·8월 Main 압력 저하를 일부러 심어 두었고, 압력 정상 범위·모델별 압력 작업표준·금형 보수·설비 수선 이력도 가상 값입니다. 실제 파일을 넣기 전에 「모든 데이터 지우기」를 해 주세요.';
    main.innerHTML = '';
    add(main, staleBanner());
    add(main, p.fn());
    document.title = p.t + ' — Deep drawing 불량률 분석';
    var ft = document.querySelector('.site-footer');
    if (ft && !ft.querySelector('.ver')) ft.appendChild(h('span', { class: 'ver' }, ' · 판 ' + APP_VERSION));
    save();
  }
  window.addEventListener('hashchange', function () { render(); main.focus(); window.scrollTo(0, 0); });
  render();
})();
