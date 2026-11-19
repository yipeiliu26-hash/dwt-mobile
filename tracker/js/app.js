/*
 * app.js — 畫面與操作
 *
 * 資料流:
 *   WMS .xlsx ──(excel-io.readSheetRows)──> rows ──(core.parseSnapshot)──> snapshot
 *   snapshot ──(core.applySnapshot)──> state ──(store: tracker-data.json + localStorage)
 *   state ──(core.weeklyRows / itemRows / rangeDiff / declineSources)──> 畫面、匯出Excel
 *
 * 全域只有 state(持久化的資料)和 view(畫面選擇,不存檔)兩個物件。
 */
(function () {
  'use strict';
  var Core = window.DWTCore, Excel = window.DWTExcel, Store = window.DWTStore, Charts = window.DWTCharts;
  var fmtUSD = Charts.fmtUSD;
  var fmtPct = function (v) { return v == null ? '' : (v >= 0 ? '+' : '') + (v * 100).toFixed(1) + '%'; };
  var fmtQty = function (v) { return v == null ? '—' : Number(v).toLocaleString('en-US'); };
  var fmtTime = function (iso) {
    if (!iso) return '';
    var d = new Date(iso);
    if (isNaN(d)) return '';
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  };
  var $ = function (id) { return document.getElementById(id); };

  // ---------- 狀態 ----------
  var state = loadInitialState();
  var view = { rangeStart: null, rangeEnd: null, knownFirst: null, knownLast: null, whFilter: 'all', diffWhFilter: 'all', searchText: '', tab: 'items' };
  var folder = { handle: null, granted: false, scanning: false, lastScanAt: null, lastScan: null, lastScanStarted: 0 };
  var saveStatus = { text: '', at: null };

  function loadInitialState() {
    var local = Store.loadLocal();
    if (local) return Core.normalizeState(local);
    var seed = window.DWT_SEED_DATA ? Core.normalizeState(window.DWT_SEED_DATA) : Core.emptyState();
    Store.saveLocal(seed);
    return seed;
  }

  // ---------- 儲存 ----------
  var writeChain = Promise.resolve();
  var writeTimer = null;
  function persist(immediate) {
    Store.saveLocal(state);
    if (!(folder.handle && folder.granted)) { setSaveStatus('已儲存在這個瀏覽器'); return Promise.resolve(); }
    clearTimeout(writeTimer);
    return new Promise(function (resolve) {
      writeTimer = setTimeout(function () {
        var snapshot = Core.clone(state);
        writeChain = writeChain.then(function () { return Store.writeDataFile(folder.handle, snapshot); })
          .then(function () { setSaveStatus('已儲存到資料夾 ' + Store.DATA_FILE); })
          .catch(function (e) {
            setSaveStatus('寫入資料夾失敗(已暫存在瀏覽器):' + (e && e.message ? e.message : e));
            if (e && (e.name === 'NotAllowedError' || e.name === 'SecurityError')) { folder.granted = false; renderConn(); }
          })
          .then(resolve);
      }, immediate ? 0 : 400);
    });
  }
  function setSaveStatus(text) {
    saveStatus.text = text; saveStatus.at = new Date();
    var el = $('save-hint');
    if (el) el.textContent = ' (' + text + ',' + fmtTime(saveStatus.at.toISOString()).slice(11) + ')';
  }

  var toastTimer = null;
  function toast(msg, kind, sticky) {
    var el = $('toast');
    el.textContent = msg;
    el.className = 'toast ' + (kind || '');
    el.hidden = false;
    clearTimeout(toastTimer);
    if (!sticky) toastTimer = setTimeout(function () { el.hidden = true; }, 8000);
  }

  // ---------- 區間 ----------
  /** 資料變動後調整區間:沒動過、或原本就貼齊最早/最新週的那一端,跟著延伸 */
  function syncRange() {
    var dates = Object.keys(state.weeks).sort();
    var first = dates[0] || null, last = dates[dates.length - 1] || null;
    if (!view.rangeStart || view.rangeStart === view.knownFirst) view.rangeStart = first;
    if (!view.rangeEnd || view.rangeEnd === view.knownLast) view.rangeEnd = last;
    view.knownFirst = first; view.knownLast = last;
    $('start-input').value = view.rangeStart || '';
    $('end-input').value = view.rangeEnd || '';
  }

  // ---------- 分頁 ----------
  var panels = { items: $('panel-items'), trend: $('panel-trend'), data: $('panel-data'), quality: $('panel-quality') };
  document.querySelectorAll('.tab-btn').forEach(function (btn) {
    btn.addEventListener('click', function () { showTab(btn.dataset.tab); });
  });
  function showTab(tab) {
    view.tab = tab;
    document.querySelectorAll('.tab-btn').forEach(function (b) { b.setAttribute('aria-selected', String(b.dataset.tab === tab)); });
    Object.keys(panels).forEach(function (k) { panels[k].hidden = (k !== tab); });
    // SVG 依容器寬度繪製,隱藏時寬度是 0,顯示時要重畫
    if (tab === 'trend') renderTrend();
    if (tab === 'data') renderData();
    if (tab === 'quality' && window.DWTQuality) window.DWTQuality.render();
  }

  // ---------- 趨勢總覽 ----------
  $('start-input').addEventListener('change', function (e) { if (e.target.value) { view.rangeStart = e.target.value; renderTrend(); } });
  $('end-input').addEventListener('change', function (e) { if (e.target.value) { view.rangeEnd = e.target.value; renderTrend(); } });
  $('reset-range').addEventListener('click', function () { view.rangeStart = null; view.rangeEnd = null; syncRange(); renderTrend(); });

  function updateCardSelection() {
    Core.WAREHOUSES.forEach(function (wh) {
      var card = $('card-' + wh);
      if (card) card.classList.toggle('active', view.diffWhFilter === wh);
    });
  }
  Core.WAREHOUSES.forEach(function (wh) {
    $('card-' + wh).addEventListener('click', function () {
      view.diffWhFilter = (view.diffWhFilter === wh) ? 'all' : wh;
      updateCardSelection();
      renderRangeDiff();
      renderDeclineSources();
    });
  });

  function inRange(d) { return (!view.rangeStart || d.date >= view.rangeStart) && (!view.rangeEnd || d.date <= view.rangeEnd); }

  function renderTrend() {
    var rows = Core.weeklyRows(state);
    Core.WAREHOUSES.forEach(function (wh) {
      var card = $('card-' + wh);
      var valueEl = card.querySelector('.value');
      var deltaEl = card.querySelector('.delta');
      var s = Core.summarize(rows, wh);
      valueEl.classList.remove('skeleton');
      if (!s) {
        valueEl.textContent = '—';
        deltaEl.textContent = '';
      } else {
        valueEl.textContent = fmtUSD(s.latest.usd_amount) + ' ';
        var small = document.createElement('small');
        small.textContent = s.latest.week + ' 當週';
        valueEl.appendChild(small);
        if (s.delta == null) {
          deltaEl.textContent = '尚無上一週可比較';
          deltaEl.className = 'delta flat';
        } else {
          var dir = s.delta > 0 ? 'up' : (s.delta < 0 ? 'down' : 'flat');
          deltaEl.textContent = '較 ' + s.prev.week + '：' + (s.delta > 0 ? '+' : '') + fmtUSD(s.delta) + '（' + fmtPct(s.deltaPct) + '）';
          deltaEl.className = 'delta ' + dir;
        }
      }
      if (!panels.trend.hidden) {
        Charts.drawSvgChart(card.querySelector('svg'), rows.filter(function (d) { return d.warehouse === wh && inRange(d); }));
      }
    });

    var fn = $('trend-footnote');
    if (!rows.length) {
      fn.textContent = '還沒有任何週次的資料。請到「資料來源」連結資料夾或匯入 WMS 檔案。';
    } else {
      var parts = Core.WAREHOUSES.map(function (wh) {
        var s = Core.summarize(rows, wh);
        if (!s) return wh + ' 尚無資料';
        if (s.delta == null) return wh + ' ' + s.latest.week + ' 為 ' + fmtUSD(s.latest.usd_amount);
        var verb = s.delta > 0 ? '上升' : (s.delta < 0 ? '下降' : '持平');
        return wh + '較上週' + verb + ' ' + fmtUSD(Math.abs(s.delta));
      });
      fn.textContent = parts.join('，') + '。資料來自每週 WMS 不良品倉匯出檔,每週一筆快照,不會覆蓋先前週次。';
    }
    renderRangeDiff();
    renderDeclineSources();
  }

  function renderScopeLabel(elId) {
    var scope = $(elId);
    if (!scope) return;
    if (view.diffWhFilter === 'all') {
      scope.textContent = '（點上方倉別卡片可只看該倉）';
      return;
    }
    scope.innerHTML = '';
    var label = document.createElement('span');
    label.textContent = '（只顯示' + view.diffWhFilter + '）';
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = '顯示全部';
    btn.addEventListener('click', function () {
      view.diffWhFilter = 'all';
      updateCardSelection();
      renderRangeDiff();
      renderDeclineSources();
    });
    scope.appendChild(label);
    scope.appendChild(btn);
  }

  function makeCell(text, cls) {
    var s = document.createElement('span');
    if (cls) s.className = cls;
    s.textContent = text;
    s.title = text;
    return s;
  }

  // ---------- 本區間項目變動(單一表格:每個有變動的料號一列,合計 = 圖表的變動) ----------
  var CHG_TYPES = ['新增', '增加', '減少', '轉出'];
  function renderRangeDiff() {
    var listEl = $('chg-list'), totalEl = $('chg-total'), summary = $('chg-summary'), footnote = $('diff-footnote');
    renderScopeLabel('diff-scope');
    var r = Core.rangeChanges(state, view.rangeStart, view.rangeEnd, view.diffWhFilter);
    summary.textContent = ''; totalEl.textContent = '';
    if (r.insufficient) {
      listEl.innerHTML = '<div class="diff-empty-row">所選區間內週次不足,無法比較</div>';
      footnote.textContent = '';
      return;
    }
    var sign = function (v) { return (v > 0 ? '+' : (v < 0 ? '−' : '')) + fmtUSD(Math.abs(v)); };
    var qsign = function (v) { return (v > 0 ? '+' : (v < 0 ? '−' : '')) + fmtQty(Math.abs(v)); };
    if (CHG_TYPES.indexOf(view.chgType) < 0) view.chgType = 'all';
    // 類型篩選(附筆數與金額小計)
    [['all', '全部', r.items.length, r.total]].concat(CHG_TYPES.map(function (k) { return [k, k, r.counts[k], r.sums[k]]; })).forEach(function (p) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'chg-chip chg-' + p[0];
      b.setAttribute('aria-pressed', String(view.chgType === p[0]));
      var name = document.createElement('span'); name.textContent = p[1] + ' ' + p[2];
      var amt = document.createElement('strong'); amt.textContent = sign(p[3]);
      b.appendChild(name); b.appendChild(amt);
      b.addEventListener('click', function () { view.chgType = p[0]; renderRangeDiff(); });
      summary.appendChild(b);
    });
    var items = r.items.filter(function (it) { return view.chgType === 'all' || it.type === view.chgType; });
    if (!items.length) listEl.innerHTML = '<div class="diff-empty-row">沒有符合的項目</div>';
    else {
      listEl.innerHTML = '';
      items.forEach(function (it) {
        var row = document.createElement('div'); row.className = 'chg-row';
        row.appendChild(makeCell(it.part_no, 'part-no'));
        row.appendChild(makeCell(it.warehouse));
        row.appendChild(makeCell(it.name || '', 'name'));
        row.appendChild(makeCell(it.type, 'chg-type t-' + it.type));
        row.appendChild(makeCell(fmtQty(it.startQty) + ' → ' + fmtQty(it.endQty) + '(' + qsign(it.qtyChange) + ')', 'num'));
        row.appendChild(makeCell(fmtUSD(it.startAmt) + ' → ' + fmtUSD(it.endAmt), 'num'));
        row.appendChild(makeCell(sign(it.amtChange), 'num chg-amt ' + (it.amtChange > 0 ? 'up' : 'down')));
        listEl.appendChild(row);
      });
    }
    // 合計列:全部時 = 圖表的變動
    var shownTotal = Core.round2(items.reduce(function (s, it) { return s + it.amtChange; }, 0));
    var who = view.diffWhFilter === 'all' ? '兩倉' : view.diffWhFilter;
    [view.chgType === 'all' ? '合計(' + items.length + ' 項)' : view.chgType + '小計(' + items.length + ' 項)', '', '', '',
      '', fmtUSD(r.startTotal) + ' → ' + fmtUSD(r.endTotal), sign(shownTotal)].forEach(function (txt, i) {
      totalEl.appendChild(makeCell(i === 5 && view.chgType !== 'all' ? '' : txt, i >= 4 ? 'num' : ''));
    });
    footnote.textContent = who + ' ' + r.first.week + ' → ' + r.last.week + ':期初 ' + fmtUSD(r.startTotal) + ',期末 ' + fmtUSD(r.endTotal) +
      ',變動 ' + sign(r.change) + '(= 上方圖表)。列出所有有變動的料號:新增 = 期初沒有;轉出 = 期末已不在不良品倉;增加 / 減少 = 兩週都在、金額變多 / 變少。' +
      '沒出現的一方以 0 計算,所以「金額變化」加總等於總變動。依金額變化大小排序。';
  }
  function renderDeclineSources() {
    var listEl = $('decline-list'), footnote = $('decline-footnote');
    renderScopeLabel('decline-scope');
    var diff = Core.declineSources(state, view.rangeStart, view.rangeEnd);
    if (diff.insufficient) {
      listEl.innerHTML = '<div class="diff-empty-row">所選區間內週次不足,無法計算去化明細</div>';
      footnote.textContent = '';
      return;
    }
    var items = diff.items.filter(function (it) { return view.diffWhFilter === 'all' || it.warehouse === view.diffWhFilter; });
    if (!items.length) {
      listEl.innerHTML = '<div class="diff-empty-row">無金額降低項目</div>';
    } else {
      listEl.innerHTML = '';
      items.forEach(function (it) {
        var row = document.createElement('div'); row.className = 'decline-row';
        row.appendChild(makeCell(it.part_no, 'part-no'));
        row.appendChild(makeCell(it.warehouse));
        row.appendChild(makeCell(it.name || '', 'name'));
        var qtyCell = makeCell(fmtQty(it.startQty) + ' → ' + fmtQty(it.endQty), 'num');
        if (it.removed) {
          var tag = document.createElement('span'); tag.className = 'gone-tag'; tag.textContent = '已消失';
          qtyCell.appendChild(tag); qtyCell.title += '(已消失)';
        }
        row.appendChild(qtyCell);
        row.appendChild(makeCell(fmtUSD(it.startAmt) + ' → ' + fmtUSD(it.endAmt), 'num'));
        row.appendChild(makeCell('−' + fmtUSD(it.decline), 'num decline-amt'));
        listEl.appendChild(row);
      });
    }
    var total = Core.round2(items.reduce(function (s, it) { return s + it.decline; }, 0));
    footnote.textContent = '比較所選區間內最早（' + diff.first.week + '）與最新（' + diff.last.week + '）兩週同一項目的金額,僅列出金額降低的項目（含已消失的料號,視為降至0),依降低金額由大到小排序;合計降低 ' + fmtUSD(total) + '（共' + items.length + '筆）。';
  }

  // ---------- 項目明細 ----------
  document.querySelectorAll('.wh-filter button').forEach(function (btn) {
    btn.addEventListener('click', function () {
      document.querySelectorAll('.wh-filter button').forEach(function (b) { b.setAttribute('aria-selected', String(b === btn)); });
      view.whFilter = btn.dataset.wh;
      renderItems(true);
    });
  });
  $('item-search').addEventListener('input', function (e) { view.searchText = e.target.value.trim().toLowerCase(); renderItems(true); });

  var pendingItemsRender = false;
  $('items-tbody').addEventListener('focusout', function () {
    setTimeout(function () { if (pendingItemsRender && !isEditingCell()) { pendingItemsRender = false; renderItems(); } }, 0);
  });
  function isEditingCell() {
    var a = document.activeElement;
    return !!(a && a.classList && a.classList.contains('cell-input') && $('items-tbody').contains(a));
  }

  /** force=true:使用者自己改了篩選;否則(自動更新)若正在輸入格子就延後重畫,避免打到一半被清掉 */
  function renderItems(force) {
    if (!force && isEditingCell()) { pendingItemsRender = true; return; }
    var tbody = $('items-tbody');
    var match = function (it) {
      if (view.whFilter !== 'all' && it.warehouse !== view.whFilter) return false;
      if (view.searchText) {
        var hay = (it.part_no + ' ' + (it.name || '') + ' ' + (it.supplier || '')).toLowerCase();
        if (hay.indexOf(view.searchText) === -1) return false;
      }
      return true;
    };
    // 主表只有最新一週還在不良品倉的料號;已轉出的放下方「已轉出不良品倉」
    var rows = Core.sortItems(Core.currentItemRows(state).filter(match));
    var latest = Object.keys(state.weeks).sort().pop();
    $('items-week').textContent = latest ? Core.weekLabel(latest) : '—';
    renderTransferred(match);
    if (!rows.length) {
      var empty = !Object.keys(state.weeks).length ? '還沒有資料,請到「資料來源」連結資料夾或匯入 WMS 檔案'
        : (view.whFilter !== 'all' && !view.searchText ? view.whFilter + '本週沒有不良品庫存' : '沒有符合條件的項目');
      tbody.innerHTML = '<tr><td colspan="10" class="empty-state">' + empty + '</td></tr>';
      return;
    }
    tbody.innerHTML = '';
    rows.forEach(function (it) {
      var tr = document.createElement('tr');
      var td = function (text, cls) { var c = document.createElement('td'); if (cls) c.className = cls; c.textContent = text; return c; };
      tr.appendChild(td(it.part_no, 'part-no'));
      tr.appendChild(td(it.name || '', 'name'));
      tr.appendChild(td(it.warehouse, 'wh'));
      function inputCell(spec) {
        var cell = document.createElement('td');
        var input = document.createElement('input');
        input.className = 'cell-input'; input.type = spec[1];
        if (spec[2]) input.placeholder = spec[2];
        input.value = it[spec[0]] || '';
        input.addEventListener('change', function () { saveField(it.id, spec[0], input.value, input); });
        cell.appendChild(input);
        tr.appendChild(cell);
      }
      inputCell(['supplier', 'text', '輸入供應商']);
      if (it.supplier_source === 'wms') {
        // 從「條碼庫存明細」帶入的供應商:灰色顯示;直接改掉就會變成你自己填的
        var supInput = tr.lastChild.firstChild;
        supInput.classList.add('from-wms');
        supInput.title = '來自 WMS 條碼庫存明細;直接輸入可以改成你自己填的';
      }
      tr.appendChild(td(fmtQty(it.qty), 'num'));
      tr.appendChild(td(fmtUSD(it.usd_amount), 'num'));
      [['status', 'text', '輸入處理狀態'], ['owner', 'text', '輸入Owner'], ['eta', 'date', '']].forEach(inputCell);
      // 照片(在手機或這裡加的料號照片;點開看大圖 / 加照片)
      var pc = document.createElement('td');
      var pb = document.createElement('button');
      pb.type = 'button'; pb.className = 'btn small' + (it.photos.length ? '' : ' ghost');
      pb.textContent = it.photos.length ? '照片 ' + it.photos.length : '加照片';
      pb.addEventListener('click', function () { if (window.DWTQuality) window.DWTQuality.openItemPhotos(it.id, it.part_no + ' ' + it.warehouse); });
      pc.appendChild(pb);
      tr.appendChild(pc);
      tbody.appendChild(tr);
    });
  }

  /** 已轉出不良品倉(最近 8 週內出現過、最新一週已經沒有),唯讀 */
  function renderTransferred(match) {
    var list = Core.transferredItemRows(state).filter(match);
    $('transferred-count').textContent = '(' + list.length + ')';
    var tb = $('transferred-tbody');
    tb.textContent = '';
    if (!list.length) {
      var tr0 = document.createElement('tr'), td0 = document.createElement('td');
      td0.colSpan = 8; td0.className = 'empty-state'; td0.textContent = '沒有已轉出的料號';
      tr0.appendChild(td0); tb.appendChild(tr0);
      return;
    }
    list.forEach(function (it) {
      var tr = document.createElement('tr');
      [[it.part_no, 'part-no'], [it.name || '', 'name'], [it.warehouse, 'wh'], [it.updated_week, 'num'],
        [fmtQty(it.qty), 'num'], [fmtUSD(it.usd_amount), 'num'], [it.status || '', ''], [it.owner || '', '']].forEach(function (c) {
        var td = document.createElement('td'); td.className = c[1]; td.textContent = c[0]; tr.appendChild(td);
      });
      tb.appendChild(tr);
    });
  }

  function saveField(id, field, value, inputEl) {
    Core.setAnnotation(state, id, field, value);
    persist().then(function () {
      if (window.DWTQuality) window.DWTQuality.snapshotSoon();   // 手機清單也要更新
      inputEl.classList.add('saved');
      setTimeout(function () { inputEl.classList.remove('saved'); }, 900);
    });
  }

  // ---------- 匯入 / 讀取資料夾 ----------
  async function parseFile(file, fileName) {
    try {
      var buf = await file.arrayBuffer();
      var sheet = await Excel.readSheetRows(window.ExcelJS, buf);
      var res = Core.parseSnapshot(sheet.rows, { fileName: fileName, defaultYear: new Date(file.lastModified || Date.now()).getFullYear() });
      // 不是每週不良品倉匯出 → 看看是不是有供應商欄位的「條碼庫存明細」
      if (!res.ok && res.reason === 'not_wms_export') {
        var sup = Core.parseSupplierReport(sheet.rows);
        if (sup.ok) return sup;
      }
      return res;
    } catch (e) {
      return { ok: false, reason: 'read_error', message: '無法讀取這份 Excel:' + (e && e.message ? e.message : e), warnings: [] };
    }
  }

  /** 把一份檔案的解析結果記錄進 state.sources,回傳 {changedWeek, rec} */
  function recordResult(path, file, res) {
    var rec = { size: file.size, lastModified: file.lastModified, checked_at: new Date().toISOString() };
    var changedWeek = false;
    if (res.ok && res.kind === 'suppliers') {
      var n = Core.applySupplierReport(state, res, { path: path, lastModified: file.lastModified });
      rec.status = 'imported';
      rec.kind = 'suppliers';
      rec.message = '供應商對照表:' + res.parts + ' 個料號' + (n < res.parts ? '(其中 ' + (res.parts - n) + ' 個已有較新的資料)' : '');
      state.sources[path] = rec;
      return { changedWeek: false, suppliers: n, rec: rec };
    }
    if (res.ok) {
      var r = Core.applySnapshot(state, res.snapshot, { path: path, size: file.size, lastModified: file.lastModified });
      rec.date = res.snapshot.date;
      if (r.action === 'kept_existing') {
        rec.status = 'superseded';
        rec.message = '同一天(' + res.snapshot.week + ')已有修改時間較新的檔案「' + r.previous.source_file + '」,這份未採用';
      } else {
        rec.status = 'imported';
        rec.message = (r.action === 'replaced' ? '更新 ' : '新增 ') + res.snapshot.week + ' 週' +
          (r.action === 'replaced' && r.previous.source_file && r.previous.source_file !== path ? '(取代「' + r.previous.source_file + '」)' : '') +
          (res.warnings.length ? ';' + res.warnings.join(';') : '');
        changedWeek = true;
      }
    } else {
      rec.status = res.reason === 'not_wms_export' ? 'skipped' : 'error';
      rec.message = res.message;
    }
    state.sources[path] = rec;
    return { changedWeek: changedWeek, rec: rec };
  }

  async function scanFolder(trigger) {
    if (!folder.handle || !folder.granted || folder.scanning) return;
    folder.scanning = true; folder.lastScanStarted = Date.now();
    renderConn();
    var summary = { added: [], errors: 0 };
    try {
      // 1. 先讀資料夾裡的 tracker-data.json(可能在別台電腦/別的分頁更新過),合併
      var fileState = await Store.readDataFile(folder.handle);
      if (fileState) state = Core.mergeStates(state, fileState);

      // 2. 讀所有 Excel,只解析新的或有變動的
      var files = await Store.listExcelFiles(folder.handle);
      var results = [];
      var seen = {};
      for (var i = 0; i < files.length; i++) {
        var f = files[i];
        seen[f.path] = true;
        var file = await f.handle.getFile();
        var prev = state.sources[f.path];
        if (prev && prev.size === file.size && prev.lastModified === file.lastModified) {
          results.push({ path: f.path, rec: prev, unchanged: true });
          continue;
        }
        var res = await parseFile(file, f.name);
        var out = recordResult(f.path, file, res);
        if (out.changedWeek) summary.added.push(out.rec.date);
        if (out.suppliers) summary.suppliers = (summary.suppliers || 0) + out.suppliers;
        if (out.rec.status === 'error') summary.errors++;
        results.push({ path: f.path, rec: out.rec, unchanged: false });
      }
      // 資料夾裡已經不存在的檔案:保留它匯入的週資料,只在清單上標示
      Object.keys(state.sources).forEach(function (p) {
        if (!seen[p] && p.indexOf('匯入:') !== 0) results.push({ path: p, rec: state.sources[p], missing: true });
      });
      folder.lastScan = results;
      folder.lastScanAt = new Date();
      await persist(true);
      afterDataChange();
      if (summary.added.length || summary.suppliers) {
        var what = [];
        if (summary.added.length) what.push(summary.added.map(Core.weekLabel).join('、') + ' 週的資料');
        if (summary.suppliers) what.push(summary.suppliers + ' 個料號的供應商');
        toast('已從資料夾讀入 ' + what.join(',以及 ') + (summary.errors ? ';另有 ' + summary.errors + ' 個檔案讀取失敗,詳見「資料來源」' : ''), 'ok');
      } else if (summary.errors) {
        toast(summary.errors + ' 個檔案讀取失敗,詳見「資料來源」分頁', 'warn');
      } else if (trigger === 'manual') {
        toast('資料夾沒有新的或變動的 WMS 檔案(共 ' + files.length + ' 個 Excel 檔)', '');
      }
    } catch (e) {
      if (e && (e.name === 'NotAllowedError' || e.name === 'SecurityError')) folder.granted = false;
      toast('讀取資料夾失敗:' + (e && e.message ? e.message : e), 'err', true);
    } finally {
      folder.scanning = false;
      renderConn();
    }
    // 品質異常紀錄:順便處理手機收件匣、更新手機清單
    if (folder.granted && window.DWTQuality) window.DWTQuality.sync(trigger);
  }

  async function importFiles(fileList) {
    var added = [], problems = [], suppliersMsg = [];
    for (var i = 0; i < fileList.length; i++) {
      var file = fileList[i];
      var res = await parseFile(file, file.name);
      var out = recordResult('匯入:' + file.name, file, res);
      if (out.changedWeek) added.push(Core.weekLabel(out.rec.date));
      else if (out.rec.kind === 'suppliers') suppliersMsg.push(res.parts + ' 個料號的供應商');
      else problems.push(file.name + ':' + out.rec.message);
    }
    await persist(true);
    afterDataChange();
    var msg = added.length ? '已匯入 ' + added.join('、') + ' 週的資料' : (suppliersMsg.length ? '' : '沒有匯入任何週次');
    if (suppliersMsg.length) msg += (msg ? ';' : '') + '已匯入 ' + suppliersMsg.join('、');
    if (problems.length) msg += '\n' + problems.join('\n');
    toast(msg, problems.length ? 'warn' : 'ok', !!problems.length);
  }

  // ---------- 資料夾連結 ----------
  $('btn-connect').addEventListener('click', async function () {
    try {
      var h = await Store.pickFolder();
      folder.handle = h; folder.granted = true;
      await scanFolder('connect');
      toast('已連結資料夾「' + h.name + '」。之後把新的 WMS 匯出檔放進這個資料夾,回到這個頁面就會自動讀取。', 'ok');
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      toast('無法連結資料夾:' + (e && e.message ? e.message : e), 'err', true);
    }
  });
  $('btn-reconnect').addEventListener('click', async function () {
    if (!folder.handle) return;
    folder.granted = await Store.ensurePermission(folder.handle, true);
    // 桌面版:資料夾不見了(改名、搬走、隨身碟拔掉)→ 直接請使用者重新選
    if (!folder.granted && Store.isDesktop) {
      toast('找不到資料夾「' + folder.handle.name + '」,請重新選擇。', 'warn');
      $('btn-connect').click();
      return;
    }
    renderConn();
    if (folder.granted) await scanFolder('reconnect');
  });
  $('btn-rescan').addEventListener('click', function () { scanFolder('manual'); });
  $('btn-disconnect').addEventListener('click', async function () {
    if (!confirm('中斷後看板不會再自動讀取這個資料夾(資料夾裡的檔案不會被刪除)。確定嗎?')) return;
    await Store.clearHandle();
    folder.handle = null; folder.granted = false; folder.lastScan = null;
    renderConn(); renderData();
    if (window.DWTQuality) window.DWTQuality.render();
  });
  $('btn-import').addEventListener('click', function () { $('file-import').click(); });
  $('file-import').addEventListener('change', function (e) {
    var files = Array.prototype.slice.call(e.target.files || []);
    e.target.value = '';
    if (files.length) importFiles(files);
  });

  // 回到這個分頁/視窗時自動重新讀取資料夾(節流 5 秒);開著的時候每 60 秒檢查一次
  function autoScan() {
    // 瀏覽器版:分頁在背景時不讀(省資源);桌面版:縮小到工作列也要照常讀,手機傳來的資料才會自動進來
    if (document.visibilityState !== 'visible' && !Store.isDesktop) return;
    if (Date.now() - folder.lastScanStarted < 5000) return;
    scanFolder('auto');
  }
  document.addEventListener('visibilitychange', autoScan);
  window.addEventListener('focus', autoScan);
  setInterval(autoScan, 60000);

  // ---------- 備份 ----------
  $('btn-download-backup').addEventListener('click', function () {
    var stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    var json = JSON.stringify(state, null, 1);
    Store.saveBlob('tracker-data_' + stamp + '.json', function () { return Promise.resolve(new Blob([json], { type: 'application/json' })); }, 'application/json', '.json', '看板資料備份')
      .catch(function (e) { toast('下載備份失敗:' + e.message, 'err', true); });
  });
  $('btn-load-backup').addEventListener('click', function () { $('file-backup').click(); });
  $('file-backup').addEventListener('change', async function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    try {
      var data = JSON.parse(await f.text());
      if (!data || typeof data.weeks !== 'object') throw new Error('不是看板的資料備份檔');
      state = Core.mergeStates(state, data);
      await persist(true);
      afterDataChange();
      toast('已載入備份並合併(共 ' + Object.keys(state.weeks).length + ' 週)', 'ok');
    } catch (err) {
      toast('載入備份失敗:' + err.message, 'err', true);
    }
  });

  // ---------- 資料來源分頁 ----------
  function pill(text, kind) { var s = document.createElement('span'); s.className = 'status-pill ' + (kind || ''); s.textContent = text; return s; }

  function renderData() {
    var tbody = $('weeks-tbody');
    var weeks = Core.sortedWeeks(state).slice().reverse();
    tbody.innerHTML = '';
    if (!weeks.length) tbody.innerHTML = '<tr><td colspan="8" class="empty-state">還沒有任何週次</td></tr>';
    weeks.forEach(function (w) {
      var tr = document.createElement('tr');
      var td = function (content, cls) {
        var c = document.createElement('td'); if (cls) c.className = cls;
        if (content instanceof Node) c.appendChild(content); else c.textContent = content;
        tr.appendChild(c); return c;
      };
      td(w.week + '(' + w.date + ')', 'part-no');
      td(fmtUSD(w.totals['大園倉'].usd_amount), 'num');
      td(fmtUSD(w.totals['新竹倉'].usd_amount), 'num');
      td(String((w.items || []).length), 'num');
      td((w.seeded ? '(初始資料)' : '') + (w.source_file || '—'));
      td(w.seeded ? '' : fmtTime(w.imported_at));
      var notes = [];
      if (w.unassigned && w.unassigned.length) notes.push(w.unassigned.length + ' 筆只標「不良品倉」未計入:' + w.unassigned.map(function (u) { return u.part_no + '(' + (u.storage_location || '?') + ',' + fmtUSD(u.usd_amount) + ')'; }).join('、'));
      if (w.date_source === 'filename') notes.push('日期由檔名推得');
      td(notes.join(';'), 'name');
      var btn = document.createElement('button');
      btn.className = 'btn small danger'; btn.type = 'button'; btn.textContent = '移除';
      btn.addEventListener('click', function () {
        if (!confirm('確定移除 ' + w.week + ' 這一週的資料嗎?\n(來源檔案若沒有更新,之後不會再自動匯入;處理狀態等註記不受影響)')) return;
        Core.removeWeek(state, w.date);
        persist(true);
        afterDataChange();
      });
      td(btn);
      tbody.appendChild(tr);
    });

    var ftbody = $('files-tbody');
    ftbody.innerHTML = '';
    var list = folder.lastScan;
    if (!list) {
      // 沒有掃描過(瀏覽器模式):列出曾經匯入過的檔案紀錄
      list = Object.keys(state.sources).sort().map(function (p) { return { path: p, rec: state.sources[p] }; });
    }
    if (!list.length) { ftbody.innerHTML = '<tr><td colspan="3" class="empty-state">' + (folder.handle ? '資料夾裡沒有 Excel 檔' : '還沒有讀取過資料夾或匯入檔案') + '</td></tr>'; }
    list.forEach(function (item) {
      var rec = item.rec || {};
      var tr = document.createElement('tr');
      var c1 = document.createElement('td'); c1.textContent = item.path; tr.appendChild(c1);
      var c2 = document.createElement('td');
      var removedNow = rec.status === 'imported' && rec.date && !state.weeks[rec.date];
      var replacedNow = rec.status === 'imported' && rec.date && state.weeks[rec.date] && state.weeks[rec.date].source_file !== item.path;
      if (item.missing) c2.appendChild(pill('檔案已不在資料夾', 'muted'));
      else if (removedNow) c2.appendChild(pill('該週已移除', 'muted'));
      else if (replacedNow) c2.appendChild(pill('已被較新檔案取代', 'muted'));
      else if (rec.status === 'imported') c2.appendChild(pill(item.unchanged ? '已匯入(未變動)' : '已匯入', 'ok'));
      else if (rec.status === 'superseded') c2.appendChild(pill('未採用', 'warn'));
      else if (rec.status === 'skipped') c2.appendChild(pill('略過', 'muted'));
      else if (rec.status === 'error') c2.appendChild(pill('讀取失敗', 'err'));
      tr.appendChild(c2);
      var c3 = document.createElement('td'); c3.className = 'name'; c3.textContent = rec.message || '';
      if (removedNow && folder.handle && !item.missing) {
        var again = document.createElement('button');
        again.className = 'btn small'; again.type = 'button'; again.textContent = '重新匯入';
        again.style.marginLeft = '0.4rem';
        again.addEventListener('click', function () { delete state.sources[item.path]; scanFolder('manual'); });
        c3.appendChild(again);
      }
      tr.appendChild(c3);
      ftbody.appendChild(tr);
    });

    $('howto-browser').hidden = Store.supportsFolder;
    $('btn-disconnect').hidden = !folder.handle;
  }

  // ---------- 連線狀態列 ----------
  function renderConn() {
    var dot = $('conn-dot'), text = $('conn-text');
    var bConnect = $('btn-connect'), bRe = $('btn-reconnect'), bScan = $('btn-rescan');
    var weeks = Object.keys(state.weeks).sort();
    var dataInfo = weeks.length ? '目前收錄 ' + weeks.length + ' 週(' + Core.weekLabel(weeks[0]) + ' ~ ' + Core.weekLabel(weeks[weeks.length - 1]) + ')' : '目前沒有資料';
    bConnect.hidden = true; bRe.hidden = true; bScan.hidden = true;
    text.innerHTML = '';
    var strong = document.createElement('strong');
    if (!Store.supportsFolder) {
      dot.className = 'conn-dot warn';
      strong.textContent = '瀏覽器模式';
      text.appendChild(strong);
      text.appendChild(document.createTextNode(' — 這個瀏覽器不能直接讀資料夾,請用「匯入WMS檔案」(建議改用 Chrome 或 Edge)。' + dataInfo));
    } else if (!folder.handle) {
      dot.className = 'conn-dot';
      strong.textContent = '尚未連結資料夾';
      text.appendChild(strong);
      text.appendChild(document.createTextNode(' — 連結放 WMS 匯出檔的資料夾後,新檔案會自動讀取。' + dataInfo));
      bConnect.hidden = false;
    } else if (!folder.granted) {
      dot.className = 'conn-dot warn';
      strong.textContent = Store.isDesktop ? '找不到資料夾「' + folder.handle.name + '」' : '資料夾「' + folder.handle.name + '」需要重新授權';
      text.appendChild(strong);
      text.appendChild(document.createTextNode(Store.isDesktop
        ? ' — 資料夾可能被改名、搬走或在隨身碟上,請按「重新連結」再選一次。' + dataInfo
        : ' — 瀏覽器重新開啟後需要按一下允許讀取。' + dataInfo));
      bRe.hidden = false;
    } else {
      dot.className = 'conn-dot on';
      strong.textContent = '已連結資料夾「' + folder.handle.name + '」';
      text.appendChild(strong);
      text.appendChild(document.createTextNode(' — ' + (folder.scanning ? '讀取中…' : (folder.lastScanAt ? '上次讀取 ' + fmtTime(folder.lastScanAt.toISOString()).slice(11) : '')) + '。' + dataInfo));
      bScan.hidden = false;
      bScan.disabled = folder.scanning;
    }
  }

  // ---------- 匯出 Excel ----------
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  $('export-all').addEventListener('click', async function () {
    if (typeof window.ExcelJS === 'undefined') { toast('匯出功能載入失敗(找不到 vendor/exceljs.min.js)', 'err', true); return; }
    // 先把要匯出的內容固定下來(= 按下按鈕當下畫面上的資料與區間)
    var rows = Core.weeklyRows(state);
    var items = Core.currentItemRows(state);   // 和畫面一致:只匯出本週還在不良品倉的料號
    var rs = view.rangeStart, re = view.rangeEnd;
    var byWh = {};
    Core.WAREHOUSES.forEach(function (wh) { byWh[wh] = rows.filter(function (d) { return d.warehouse === wh && inRange(d); }); });
    try {
      await Store.saveBlob('不良品倉追蹤看板.xlsx', async function () {
        var wb = Excel.buildExportWorkbook(window.ExcelJS, {
          rangeStart: rs, rangeEnd: re, weeklyRows: rows, itemRows: items,
          chartImage: Charts.renderChartImage(byWh)
        });
        var buf = await wb.xlsx.writeBuffer();
        return new Blob([buf], { type: XLSX_MIME });
      }, XLSX_MIME, '.xlsx', 'Excel 活頁簿');
    } catch (e) {
      toast('匯出失敗:' + (e && e.message ? e.message : e), 'err', true);
    }
  });

  // ---------- 整體重畫 ----------
  function afterDataChange() {
    syncRange();
    renderItems();
    renderTrend();
    renderData();
    renderConn();
  }

  window.addEventListener('resize', function () { if (!panels.trend.hidden) renderTrend(); });

  // 同一台電腦開了兩個分頁時,另一個分頁存檔後同步過來
  window.addEventListener('storage', function (e) {
    if (e.key !== 'dwt-tracker-data-v1' || !e.newValue) return;
    try { state = Core.mergeStates(state, JSON.parse(e.newValue)); afterDataChange(); } catch (err) { /* ignore */ }
  });

  // ---------- 啟動 ----------
  async function boot() {
    afterDataChange();
    if (Store.supportsFolder) {
      var h = await Store.getSavedHandle();
      if (h) {
        folder.handle = h;
        folder.granted = await Store.ensurePermission(h, false);
        renderConn();
        if (folder.granted) await scanFolder('startup');
      }
    }
    renderConn();
    renderData();
  }

  // 測試/除錯用(瀏覽器 console 可看 DWTApp.state())
  // ---------- 手機版用(js/quality-app.js 呼叫) ----------
  /** 給手機清單的項目明細 */
  function trackerSnapshot() {
    var dates = Object.keys(state.weeks).sort();
    var last = dates[dates.length - 1] || '';
    return { rows: Core.itemRows(state), annotations: state.annotations, latest_date: last, latest_week: last ? Core.weekLabel(last) : '' };
  }
  /** 手機傳來的註記:較新的才採用。回傳實際更新的筆數 */
  function applyAnnotations(map) {
    var n = 0;
    Object.keys(map || {}).forEach(function (id) {
      var cur = state.annotations[id], y = map[id];
      if (!cur || String(y.updated_at || '') > String(cur.updated_at || '')) {
        state.annotations[id] = { status: y.status || '', owner: y.owner || '', eta: y.eta || '', supplier: y.supplier || '', updated_at: y.updated_at };
        if (String(y.updated_at) > String(state.updated_at || '')) state.updated_at = y.updated_at;
        n++;
      }
    });
    if (!n) return Promise.resolve(0);
    return persist(true).then(function () { renderItems(); return n; });
  }

  /** 料號照片紀錄(手機傳來或這台電腦新增):較新的才採用。回傳實際更新的筆數 */
  function applyItemPhotos(records) {
    var n = 0;
    (records || []).forEach(function (rec) { if (Core.putItemPhoto(state, rec)) n++; });
    if (!n) return Promise.resolve(0);
    return persist(true).then(function () { renderItems(); return n; });
  }

  window.DWTApp = {
    state: function () { return state; }, view: view, folder: folder, scan: scanFolder, importFiles: importFiles, toast: toast,
    trackerSnapshot: trackerSnapshot, applyAnnotations: applyAnnotations, applyItemPhotos: applyItemPhotos
  };

  boot();
})();
