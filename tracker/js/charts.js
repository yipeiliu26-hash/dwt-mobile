/*
 * charts.js — 趨勢圖
 *   drawSvgChart(svg, rows)        畫面上的互動圖(SVG,跟著深/淺色主題)
 *   renderChartImage(rowsByWh)     匯出 Excel 用的 PNG(Canvas,固定淺色配色 = 畫面淺色主題的顏色)
 * 兩者的比例尺、格線、面積漸層、標籤位置要保持一致,改一邊就要改另一邊。
 */
(function (root) {
  'use strict';

  var fmtUSD = function (v) {
    if (v == null) return '—';
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(v);
  };
  var parseDate = function (s) { return new Date(s + 'T00:00:00Z'); };
  var SVGNS = 'http://www.w3.org/2000/svg';

  // 淺色主題的顏色(與 css/styles.css 的 :root 一致)
  var LIGHT = { fg: '#26241f', muted: '#78746a', border: '#e3e0d7', accent: '#2f6f63' };

  function drawSvgChart(svg, rows) {
    var box = svg.parentElement;
    var width = Math.max(box.clientWidth, 200);
    var height = 200;
    var margin = { top: 16, right: 16, bottom: 26, left: 10 };
    while (svg.firstChild) svg.removeChild(svg.firstChild);
    svg.setAttribute('viewBox', '0 0 ' + width + ' ' + height);

    if (!rows.length) {
      var t = document.createElementNS(SVGNS, 'text');
      t.setAttribute('x', width / 2); t.setAttribute('y', height / 2);
      t.setAttribute('text-anchor', 'middle'); t.setAttribute('fill', 'var(--fg-muted)'); t.setAttribute('font-size', '12');
      t.textContent = '所選區間內沒有資料';
      svg.appendChild(t);
      return;
    }

    var sorted = rows.slice().sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    var xs = sorted.map(function (d) { return parseDate(d.date).getTime(); });
    var minX = Math.min.apply(null, xs), maxX = Math.max.apply(null, xs);
    var spanX = Math.max(maxX - minX, 1);
    var values = sorted.map(function (d) { return d.usd_amount; });
    var maxY = Math.max.apply(null, values) * 1.15 || 1;

    var xScale = function (t) { return margin.left + (sorted.length === 1 ? (width - margin.left - margin.right) / 2 : (t - minX) / spanX * (width - margin.left - margin.right)); };
    var yScale = function (v) { return (height - margin.bottom) - (v / maxY) * (height - margin.top - margin.bottom); };

    [0.33, 0.66, 1].forEach(function (f) {
      var y = yScale(maxY * f);
      var line = document.createElementNS(SVGNS, 'line');
      line.setAttribute('x1', margin.left); line.setAttribute('x2', width - margin.right);
      line.setAttribute('y1', y); line.setAttribute('y2', y);
      line.setAttribute('stroke', 'var(--border)'); line.setAttribute('stroke-width', '1');
      svg.appendChild(line);
    });

    var gradId = 'dwt-grad-' + Math.random().toString(36).slice(2);
    var defs = document.createElementNS(SVGNS, 'defs');
    var grad = document.createElementNS(SVGNS, 'linearGradient');
    grad.setAttribute('id', gradId); grad.setAttribute('x1', '0'); grad.setAttribute('x2', '0'); grad.setAttribute('y1', '0'); grad.setAttribute('y2', '1');
    [[0, .26], [.5, .08], [1, 0]].forEach(function (pair) {
      var stop = document.createElementNS(SVGNS, 'stop');
      stop.setAttribute('offset', String(pair[0]));
      stop.setAttribute('style', 'stop-color:var(--accent)');
      stop.setAttribute('stop-opacity', String(pair[1]));
      grad.appendChild(stop);
    });
    defs.appendChild(grad);
    svg.appendChild(defs);

    var pts = sorted.map(function (d) { return xScale(parseDate(d.date).getTime()) + ',' + yScale(d.usd_amount); }).join(' ');
    var baseline = height - margin.bottom;
    var area = document.createElementNS(SVGNS, 'polygon');
    area.setAttribute('points', (xScale(minX) + ',' + baseline) + ' ' + pts + ' ' + (xScale(maxX) + ',' + baseline));
    area.setAttribute('fill', 'url(#' + gradId + ')');
    svg.appendChild(area);

    var line = document.createElementNS(SVGNS, 'polyline');
    line.setAttribute('points', pts);
    line.setAttribute('fill', 'none');
    line.setAttribute('stroke', 'var(--accent)');
    line.setAttribute('stroke-width', '2');
    line.setAttribute('stroke-linecap', 'round');
    line.setAttribute('stroke-linejoin', 'round');
    svg.appendChild(line);

    var tooltip = box.querySelector('.tooltip') || (function () {
      var el = document.createElement('div'); el.className = 'tooltip'; box.appendChild(el); return el;
    })();

    sorted.forEach(function (d) {
      var cx = xScale(parseDate(d.date).getTime());
      var cy = yScale(d.usd_amount);
      // 加大的透明熱區,滑鼠比較好對準
      var hit = document.createElementNS(SVGNS, 'circle');
      hit.setAttribute('cx', cx); hit.setAttribute('cy', cy); hit.setAttribute('r', '10'); hit.setAttribute('fill', 'transparent');
      var dot = document.createElementNS(SVGNS, 'circle');
      dot.setAttribute('cx', cx); dot.setAttribute('cy', cy); dot.setAttribute('r', '3.5');
      dot.setAttribute('fill', 'var(--accent)');
      [hit, dot].forEach(function (el) {
        el.addEventListener('mouseenter', function () {
          tooltip.textContent = d.week + ' · ' + fmtUSD(d.usd_amount);
          tooltip.style.left = Math.min(Math.max(cx - 40, 0), width - 90) + 'px';
          tooltip.style.top = Math.max(cy - 34, 0) + 'px';
          tooltip.classList.add('show');
        });
        el.addEventListener('mouseleave', function () { tooltip.classList.remove('show'); });
      });
      svg.appendChild(dot);
      svg.appendChild(hit);

      var label = document.createElementNS(SVGNS, 'text');
      label.setAttribute('x', cx); label.setAttribute('y', height - margin.bottom + 16);
      label.setAttribute('text-anchor', 'middle'); label.setAttribute('font-size', '10'); label.setAttribute('fill', 'var(--fg-muted)');
      label.textContent = d.week;
      svg.appendChild(label);
    });

    var last = sorted[sorted.length - 1];
    var endLabel = document.createElementNS(SVGNS, 'text');
    endLabel.setAttribute('x', Math.min(xScale(parseDate(last.date).getTime()) + 6, width - margin.right - 2));
    endLabel.setAttribute('y', yScale(last.usd_amount) - 8);
    endLabel.setAttribute('text-anchor', 'end'); endLabel.setAttribute('font-size', '11'); endLabel.setAttribute('fill', 'var(--fg)');
    endLabel.textContent = fmtUSD(last.usd_amount);
    svg.appendChild(endLabel);
  }

  function drawCanvasChart(ctx, x0, y0, w, h, rows, title) {
    ctx.save();
    ctx.translate(x0, y0);
    ctx.fillStyle = LIGHT.fg;
    ctx.font = '13px "Microsoft JhengHei", sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(title, 0, 14);

    var margin = { top: 28, right: 14, bottom: 28, left: 10 };
    var plotW = w - margin.left - margin.right;
    var plotH = h - margin.top - margin.bottom;

    if (!rows.length) {
      ctx.fillStyle = LIGHT.muted;
      ctx.font = '12px "Microsoft JhengHei", sans-serif';
      ctx.fillText('所選區間內沒有資料', margin.left, margin.top + plotH / 2);
      ctx.restore();
      return;
    }

    var sorted = rows.slice().sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    var toTime = function (d) { return parseDate(d.date).getTime(); };
    var xs = sorted.map(toTime);
    var minX = Math.min.apply(null, xs), maxX = Math.max.apply(null, xs);
    var spanX = Math.max(maxX - minX, 1);
    var values = sorted.map(function (d) { return d.usd_amount; });
    var maxY = Math.max.apply(null, values) * 1.15 || 1;

    var xScale = function (t) { return margin.left + (sorted.length === 1 ? plotW / 2 : (t - minX) / spanX * plotW); };
    var yScale = function (v) { return margin.top + plotH - (v / maxY) * plotH; };

    ctx.strokeStyle = LIGHT.border;
    ctx.lineWidth = 1;
    [0.33, 0.66, 1].forEach(function (f) {
      var y = yScale(maxY * f);
      ctx.beginPath(); ctx.moveTo(margin.left, y); ctx.lineTo(margin.left + plotW, y); ctx.stroke();
    });

    // 面積漸層(與 SVG 版相同的 0.26 → 0.08 → 0)
    var g = ctx.createLinearGradient(0, margin.top, 0, margin.top + plotH);
    g.addColorStop(0, 'rgba(47,111,99,0.26)');
    g.addColorStop(0.5, 'rgba(47,111,99,0.08)');
    g.addColorStop(1, 'rgba(47,111,99,0)');
    ctx.beginPath();
    ctx.moveTo(xScale(minX), margin.top + plotH);
    sorted.forEach(function (d) { ctx.lineTo(xScale(toTime(d)), yScale(d.usd_amount)); });
    ctx.lineTo(xScale(maxX), margin.top + plotH);
    ctx.closePath();
    ctx.fillStyle = g;
    ctx.fill();

    ctx.beginPath();
    sorted.forEach(function (d, i) {
      var px = xScale(toTime(d)), py = yScale(d.usd_amount);
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    });
    ctx.strokeStyle = LIGHT.accent;
    ctx.lineWidth = 2;
    ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    ctx.stroke();

    ctx.font = '10px "Microsoft JhengHei", sans-serif';
    sorted.forEach(function (d) {
      var px = xScale(toTime(d)), py = yScale(d.usd_amount);
      ctx.beginPath(); ctx.arc(px, py, 3.5, 0, 2 * Math.PI); ctx.fillStyle = LIGHT.accent; ctx.fill();
      ctx.fillStyle = LIGHT.muted;
      ctx.textAlign = 'center';
      ctx.fillText(d.week, px, margin.top + plotH + 16);
    });

    var last = sorted[sorted.length - 1];
    ctx.fillStyle = LIGHT.fg;
    ctx.textAlign = 'right';
    ctx.font = '11px "Microsoft JhengHei", sans-serif';
    ctx.fillText(fmtUSD(last.usd_amount), Math.min(xScale(toTime(last)) + 6, margin.left + plotW - 2), yScale(last.usd_amount) - 8);

    ctx.restore();
  }

  /** rowsByWh = {'大園倉': rows, '新竹倉': rows} → {base64, width, height} */
  function renderChartImage(rowsByWh) {
    var canvas = document.createElement('canvas');
    var W = 980, H = 300, scale = 2;
    canvas.width = W * scale; canvas.height = H * scale;
    var ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    drawCanvasChart(ctx, 10, 10, W / 2 - 20, H - 20, rowsByWh['大園倉'] || [], '大園倉 不良品倉金額趨勢(美金)');
    drawCanvasChart(ctx, W / 2 + 10, 10, W / 2 - 20, H - 20, rowsByWh['新竹倉'] || [], '新竹倉 不良品倉金額趨勢(美金)');
    return { base64: canvas.toDataURL('image/png'), width: W, height: H };
  }

  root.DWTCharts = { fmtUSD: fmtUSD, drawSvgChart: drawSvgChart, renderChartImage: renderChartImage };
})(window);
