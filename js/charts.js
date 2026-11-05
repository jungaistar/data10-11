/* 작은 SVG 그래프 — 외부 라이브러리 없이 막대·꺾은선·가로 막대 세 가지만 (data09-25 charts.js 를 넓힘).
   한 그래프에 y축은 하나. 목표 불량률 같은 기준선(refs)을 점선으로 그립니다. 막대·점에 마우스를 올리면 값이 보입니다(title).
   색은 CSS 변수(--series-1 …)를 씁니다. */
(function (root) {
  'use strict';
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, attrs, text) {
    var e = document.createElementNS(NS, tag);
    Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, attrs[k]); });
    if (text != null) e.textContent = text;
    return e;
  }
  function niceMax(v) {
    if (!(v > 0)) return 1;
    var p = Math.pow(10, Math.floor(Math.log10(v))), n = v / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
  }
  function fmt(v, unit) {
    if (v == null) return '-';
    if (unit === '%') return (Math.round(v * 1000) / 10) + '%';
    if (unit === '%2') return (Math.round(v * 10000) / 100) + '%';
    return (Math.round(v * 10) / 10).toLocaleString('ko-KR') + (unit ? unit : '');
  }
  function allValues(opts) {
    var all = [];
    opts.series.forEach(function (s) { s.values.forEach(function (v) { if (v != null && isFinite(v)) all.push(v); }); });
    return all;
  }
  function refLines(svg, opts, x1, x2, y) {
    (opts.refs || []).forEach(function (r) {
      var yy = y(r.value);
      svg.appendChild(el('line', { x1: x1, x2: x2, y1: yy, y2: yy, class: 'ref ' + (r.cls || '') }));
      svg.appendChild(el('text', { x: x2 - 2, y: yy - 4, 'text-anchor': 'end', class: 'ref-label ' + (r.cls || '') }, r.label));
    });
  }

  /* opts: { labels:[], series:[{name, values:[], color}], unit, type:'bar'|'line', height, title, refs:[{value,label,cls}] } */
  function chart(opts) {
    if (opts.type === 'hbar') return hbar(opts);
    var W = 640, H = opts.height || 220, padL = 52, padR = 12, padT = 14, padB = 30;
    var n = opts.labels.length, unit = opts.unit || '';
    var all = allValues(opts);
    var max = niceMax(Math.max.apply(null, all.concat([opts.minMax || 0]).concat((opts.refs || []).map(function (r) { return r.value * 1.1; }))));
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, class: 'chart', role: 'img', 'aria-label': opts.title || '' });
    var plotW = W - padL - padR, plotH = H - padT - padB;
    function y(v) { return padT + plotH - (v / max) * plotH; }
    for (var g = 0; g <= 4; g++) {
      var gv = max * g / 4, gy = y(gv);
      svg.appendChild(el('line', { x1: padL, x2: W - padR, y1: gy, y2: gy, class: g === 0 ? 'axis' : 'grid' }));
      svg.appendChild(el('text', { x: padL - 6, y: gy + 4, 'text-anchor': 'end', class: 'tick' }, fmt(gv, unit === '%' || unit === '%2' ? '%' : '')));
    }
    var step = plotW / Math.max(n, 1);
    var every = Math.max(1, Math.ceil(n / 16)); // 칸이 많으면 눈금 글자를 건너뜀
    opts.labels.forEach(function (lb, i) {
      if (i % every) return;
      svg.appendChild(el('text', { x: padL + step * (i + 0.5), y: H - 10, 'text-anchor': 'middle', class: 'tick' }, lb));
    });
    if (!all.length) {
      svg.appendChild(el('text', { x: W / 2, y: H / 2, 'text-anchor': 'middle', class: 'tick' }, '표시할 값이 없습니다'));
      return svg;
    }
    if (opts.type === 'line') {
      opts.series.forEach(function (s) {
        var d = '', pen = false;
        s.values.forEach(function (v, i) {
          if (v == null || !isFinite(v)) { pen = false; return; }
          d += (pen ? 'L' : 'M') + (padL + step * (i + 0.5)).toFixed(1) + ' ' + y(v).toFixed(1) + ' ';
          pen = true;
        });
        svg.appendChild(el('path', { d: d, class: 'ln', style: 'stroke:' + s.color }));
        s.values.forEach(function (v, i) {
          if (v == null || !isFinite(v)) return;
          var c = el('circle', { cx: padL + step * (i + 0.5), cy: y(v), r: n > 24 ? 3 : 4.5, class: 'dot', style: 'fill:' + s.color });
          c.appendChild(el('title', null, opts.labels[i] + ' · ' + s.name + ' ' + fmt(v, unit)));
          svg.appendChild(c);
        });
      });
    } else {
      var k = opts.series.length, gap = 2, bw = Math.min(28, (step * 0.7 - gap * (k - 1)) / k);
      opts.labels.forEach(function (lb, i) {
        var x0 = padL + step * (i + 0.5) - (bw * k + gap * (k - 1)) / 2;
        opts.series.forEach(function (s, j) {
          var v = s.values[i];
          if (v == null || !isFinite(v)) return;
          var top = y(v), h = Math.max(0, padT + plotH - top), r = Math.min(4, bw / 2, h);
          var x = x0 + j * (bw + gap);
          var p = 'M' + x + ' ' + (top + h) + 'V' + (top + r) + 'Q' + x + ' ' + top + ' ' + (x + r) + ' ' + top +
            'H' + (x + bw - r) + 'Q' + (x + bw) + ' ' + top + ' ' + (x + bw) + ' ' + (top + r) + 'V' + (top + h) + 'Z';
          var b = el('path', { d: p, class: 'bar', style: 'fill:' + ((s.colors && s.colors[i]) || s.color) });
          b.appendChild(el('title', null, lb + ' · ' + s.name + ' ' + fmt(v, unit)));
          svg.appendChild(b);
        });
      });
    }
    refLines(svg, opts, padL, W - padR, y);
    return svg;
  }

  /* 가로 막대 — 이름이 긴 항목(금형·소재규격·조건)을 비교할 때. 시리즈는 하나, colors 로 막대마다 색 */
  function hbar(opts) {
    var s = opts.series[0], n = opts.labels.length, unit = opts.unit || '';
    var W = 640, rowH = 30, padT = 10, padB = 28, padL = Math.min(260, 16 + Math.max.apply(null, opts.labels.map(function (l) { return String(l).length; }).concat([4])) * 13), padR = 78;
    var H = padT + padB + rowH * Math.max(n, 1);
    var all = allValues(opts);
    var max = niceMax(Math.max.apply(null, all.concat([0]).concat((opts.refs || []).map(function (r) { return r.value * 1.1; }))));
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, class: 'chart hbar', role: 'img', 'aria-label': opts.title || '' });
    var plotW = W - padL - padR;
    function x(v) { return padL + (v / max) * plotW; }
    for (var g = 0; g <= 4; g++) {
      var gx = x(max * g / 4);
      svg.appendChild(el('line', { x1: gx, x2: gx, y1: padT, y2: H - padB, class: g === 0 ? 'axis' : 'grid' }));
      svg.appendChild(el('text', { x: gx, y: H - 8, 'text-anchor': 'middle', class: 'tick' }, fmt(max * g / 4, unit === '%2' ? '%' : unit)));
    }
    opts.labels.forEach(function (lb, i) {
      var cy = padT + rowH * i + rowH / 2;
      var label = String(lb).length > 20 ? String(lb).slice(0, 19) + '…' : String(lb);
      svg.appendChild(el('text', { x: padL - 8, y: cy + 5, 'text-anchor': 'end', class: 'tick lbl' }, label));
      var v = s.values[i];
      if (v == null || !isFinite(v)) return;
      var b = el('rect', { x: padL, y: cy - 9, width: Math.max(1, x(v) - padL), height: 18, rx: 3, class: 'bar', style: 'fill:' + ((s.colors && s.colors[i]) || s.color) });
      b.appendChild(el('title', null, lb + ' · ' + s.name + ' ' + fmt(v, unit)));
      svg.appendChild(b);
      svg.appendChild(el('text', { x: x(v) + 6, y: cy + 5, class: 'tick val' }, fmt(v, unit)));
    });
    (opts.refs || []).forEach(function (r) {
      var xx = x(r.value);
      svg.appendChild(el('line', { x1: xx, x2: xx, y1: padT - 4, y2: H - padB, class: 'ref ' + (r.cls || '') }));
    });
    return svg;
  }
  /* 범례 — 계열이 둘 이상이거나 기준선이 있을 때 */
  function legend(opts) {
    var box = document.createElement('div'); box.className = 'legend';
    var items = opts.series.length > 1 ? opts.series.map(function (s) { return { name: s.name, color: s.color }; }) : [];
    (opts.refs || []).forEach(function (r) { items.push({ name: r.label, ref: r.cls || '' }); });
    items.forEach(function (s) {
      var i = document.createElement('span'); i.className = 'lg';
      var sw = document.createElement('i');
      if (s.ref != null) sw.className = 'ref-sw ' + s.ref; else sw.style.background = s.color;
      i.appendChild(sw); i.appendChild(document.createTextNode(s.name)); box.appendChild(i);
    });
    return box;
  }
  function figure(title, opts) {
    var f = document.createElement('figure'); f.className = 'fig';
    var cap = document.createElement('figcaption'); cap.textContent = title; f.appendChild(cap);
    opts.title = title;
    f.appendChild(legend(opts));
    f.appendChild(chart(opts));
    return f;
  }
  root.DDCharts = { chart: chart, figure: figure };
})(window);
