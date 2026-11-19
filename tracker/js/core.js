/*
 * core.js — 純邏輯層(不碰 DOM、不碰檔案系統)
 *
 * 負責:
 *   1. 把 WMS 匯出檔的「列」(array of arrays) 解析成一週的快照 (snapshot)
 *   2. 把快照合併進 tracker 狀態 (state)
 *   3. 從 state 推導畫面要用的資料:週度趨勢、項目明細、本區間新增/減少、金額降低來源
 *
 * 瀏覽器:以 <script> 載入後成為 window.DWTCore
 * Node   :require('./js/core.js')  (tests/ 使用)
 *
 * 注意:index.html 是用 file:// 直接開啟,Chrome 在 file:// 下不能載入 ES module,
 * 所以這裡刻意用 UMD/IIFE 寫法,不要改成 import/export。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DWTCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // v2(2026-10-22):加入 item_photos(料號照片);註記多了 supplier。v1 資料讀進來時自動補上,原本的註記不變。
  var SCHEMA_VERSION = 2;
  var WAREHOUSES = ['大園倉', '新竹倉'];
  var WH_CODE = { '大園倉': 'DY', '新竹倉': 'HC' };

  // WMS 匯出檔(工作表 "Data")的欄位。一律用「欄位名稱」找欄位,不要用欄位位置:
  // 10/07 的檔案「美金金額」在第 21 欄,10/14 之後多了「生管或採購姓名」變成第 22 欄。
  var HEADERS = {
    part_no: '物料號碼',
    qty: '數量',
    date: '年月日',
    name: '內文',
    warehouse: '倉別',
    usd: '美金金額',
    storage: '儲存地點' // 選用;檔案裡有兩個同名欄位,取第一個
  };
  var REQUIRED = ['part_no', 'qty', 'date', 'name', 'warehouse', 'usd'];
  var DEFECTIVE_KEYWORD = '不良品倉';

  // 另一種 WMS 報表「條碼庫存明細」(例如 條碼庫存明細.xlsx):每個條碼一列,有供應商代碼/名稱。
  // 只拿來對照「料號 → 供應商」,不影響每週金額。
  var SUPPLIER_HEADERS = { part_no: '料號', code: '供應商代碼', name: '供應商名稱' };

  // ---------- 小工具 ----------
  function round2(v) { return Math.round((Number(v) + Number.EPSILON) * 100) / 100; }
  function roundQty(v) { return Math.round(Number(v) * 1e6) / 1e6; }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function clone(obj) { return obj == null ? obj : JSON.parse(JSON.stringify(obj)); }
  function nowIso() { return new Date().toISOString(); }

  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = Number(String(v).replace(/,/g, '').trim());
    return isFinite(n) ? n : 0;
  }

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return toDateKey(v) || '';
    return String(v).trim();
  }

  /** 把各種日期表示轉成 'YYYY-MM-DD'。失敗回傳 null。 */
  function toDateKey(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      // ExcelJS 讀到的日期是 UTC 午夜,用 UTC 取值才不會因時區差一天
      return v.getUTCFullYear() + '-' + pad2(v.getUTCMonth() + 1) + '-' + pad2(v.getUTCDate());
    }
    if (typeof v === 'number' && isFinite(v) && v > 20000 && v < 80000) {
      // Excel 序列日期
      return toDateKey(new Date(Math.round((v - 25569) * 86400000)));
    }
    var s = String(v).trim();
    var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    return null;
  }

  /** 檔名推日期(只在檔案內「年月日」欄位全空時才用)。例:record_1019.xlsx、WMS_20261019.xlsx */
  function dateFromFileName(fileName, defaultYear) {
    if (!fileName) return null;
    var base = String(fileName).replace(/\.[^.]+$/, '');
    var m = base.match(/(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/);
    if (m) return validDate(+m[1], +m[2], +m[3]);
    m = base.match(/(?:^|\D)(\d{2})(\d{2})(?:\D|$)/);
    if (m && defaultYear) return validDate(defaultYear, +m[1], +m[2]);
    return null;
  }
  function validDate(y, mo, d) {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    var dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCMonth() !== mo - 1) return null;
    return y + '-' + pad2(mo) + '-' + pad2(d);
  }

  function weekLabel(dateKey) { return dateKey ? dateKey.slice(5, 7) + '/' + dateKey.slice(8, 10) : ''; }
  function dateKeyCompact(dateKey) { return dateKey.replace(/-/g, ''); }
  function itemId(partNo, warehouse) { return partNo + '-' + WH_CODE[warehouse]; }

  /** 倉別字串 → '大園倉' | '新竹倉' | null(是不良品倉但沒寫地點) | undefined(不是不良品倉) */
  function resolveWarehouse(raw) {
    var s = cellText(raw);
    if (s.indexOf(DEFECTIVE_KEYWORD) === -1) return undefined;
    if (s.indexOf('大園') !== -1) return '大園倉';
    if (s.indexOf('新竹') !== -1) return '新竹倉';
    return null;
  }

  // ---------- 解析 ----------
  /** 在前 10 列裡找表頭列。回傳 { rowIndex, index: {key: colIdx} } 或 null */
  function findHeader(rows) {
    var limit = Math.min(rows.length, 10);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var index = {};
      for (var c = 0; c < row.length; c++) {
        var h = cellText(row[c]);
        if (!h) continue;
        Object.keys(HEADERS).forEach(function (k) {
          if (index[k] == null && h === HEADERS[k]) index[k] = c;
        });
      }
      var ok = REQUIRED.every(function (k) { return index[k] != null; });
      if (ok) return { rowIndex: r, index: index };
    }
    return null;
  }

  function missingHeaders(rows) {
    var best = [];
    var limit = Math.min(rows.length, 10);
    var bestCount = -1;
    for (var r = 0; r < limit; r++) {
      var texts = (rows[r] || []).map(cellText);
      var missing = REQUIRED.filter(function (k) { return texts.indexOf(HEADERS[k]) === -1; });
      if (REQUIRED.length - missing.length > bestCount) { bestCount = REQUIRED.length - missing.length; best = missing; }
    }
    return best.map(function (k) { return HEADERS[k]; });
  }

  /**
   * 「條碼庫存明細」→ 料號 → 供應商。同一料號出現多個供應商時取筆數最多的,其餘放在 alt。
   * @returns {{ok:true, kind:'suppliers', suppliers:{料號:{code,name,count,alt:[]}}, parts:number} | {ok:false}}
   */
  function parseSupplierReport(rows) {
    rows = rows || [];
    var limit = Math.min(rows.length, 10), hdr = null;
    for (var r = 0; r < limit && !hdr; r++) {
      var texts = (rows[r] || []).map(cellText);
      var pi = texts.indexOf(SUPPLIER_HEADERS.part_no), ni = texts.indexOf(SUPPLIER_HEADERS.name);
      if (pi !== -1 && ni !== -1) hdr = { rowIndex: r, part: pi, name: ni, code: texts.indexOf(SUPPLIER_HEADERS.code) };
    }
    if (!hdr) return { ok: false };
    var counts = {};
    for (var i = hdr.rowIndex + 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var pn = cellText(row[hdr.part]), name = cellText(row[hdr.name]);
      if (!pn || !name) continue;
      var code = hdr.code !== -1 ? cellText(row[hdr.code]) : '';
      var key = code + '\u0000' + name;
      var m = counts[pn] = counts[pn] || {};
      m[key] = (m[key] || 0) + 1;
    }
    var suppliers = {};
    Object.keys(counts).forEach(function (pn) {
      var list = Object.keys(counts[pn]).map(function (k) {
        var p = k.split('\u0000');
        return { code: p[0], name: p[1], count: counts[pn][k] };
      }).sort(function (a, b) { return b.count - a.count; });
      suppliers[pn] = { code: list[0].code, name: list[0].name, count: list[0].count, alt: list.slice(1) };
    });
    var parts = Object.keys(suppliers).length;
    if (!parts) return { ok: false };
    return { ok: true, kind: 'suppliers', suppliers: suppliers, parts: parts };
  }

  /** 把供應商對照放進 state(就地修改)。同一料號以檔案修改時間較新的為準。回傳更新的料號數 */
  function applySupplierReport(state, res, source) {
    var src = source || {}, at = nowIso(), n = 0;
    Object.keys(res.suppliers).forEach(function (pn) {
      var cur = state.wms_suppliers[pn];
      if (cur && Number(cur.source_mtime || 0) > Number(src.lastModified || 0)) return;
      var s = res.suppliers[pn];
      state.wms_suppliers[pn] = { code: s.code, name: s.name, count: s.count, alt: s.alt, source_file: src.path || '', source_mtime: src.lastModified || 0, imported_at: at };
      n++;
    });
    if (n) state.updated_at = at;
    return n;
  }

  /**
   * 把一個工作表的列解析成週快照。
   * @param {Array<Array>} rows  第一個維度是列,第二個是欄(0-based)
   * @param {{fileName?:string, defaultYear?:number}} opts
   * @returns {{ok:boolean, reason?:string, message?:string, snapshot?:object, warnings:string[]}}
   */
  function parseSnapshot(rows, opts) {
    opts = opts || {};
    var warnings = [];
    var hdr = findHeader(rows || []);
    if (!hdr) {
      return { ok: false, reason: 'not_wms_export', message: '不是 WMS 不良品倉匯出格式(缺少欄位:' + missingHeaders(rows || []).join('、') + ')', warnings: warnings };
    }
    var ix = hdr.index;
    var totals = {};
    WAREHOUSES.forEach(function (w) { totals[w] = { usd_amount: 0, quantity: 0, line_count: 0 }; });
    var itemsMap = {};
    var unassigned = [];
    var dateCount = {};
    var defectiveRows = 0;

    for (var r = hdr.rowIndex + 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var wh = resolveWarehouse(row[ix.warehouse]);
      if (wh === undefined) continue; // 不是不良品倉
      defectiveRows++;
      var dk = toDateKey(row[ix.date]);
      if (dk) dateCount[dk] = (dateCount[dk] || 0) + 1;
      var partNo = cellText(row[ix.part_no]);
      var qty = toNumber(row[ix.qty]);
      var usd = toNumber(row[ix.usd]);
      var name = cellText(row[ix.name]);
      if (!partNo) { warnings.push('第 ' + (r + 1) + ' 列沒有物料號碼,已略過'); continue; }
      if (wh === null) {
        unassigned.push({
          part_no: partNo, name: name, qty: roundQty(qty), usd_amount: round2(usd),
          warehouse_raw: cellText(row[ix.warehouse]),
          storage_location: ix.storage != null ? cellText(row[ix.storage]) : ''
        });
        continue;
      }
      var t = totals[wh];
      t.usd_amount += usd; t.quantity += qty; t.line_count += 1;
      var id = itemId(partNo, wh);
      var it = itemsMap[id];
      if (!it) it = itemsMap[id] = { id: id, part_no: partNo, name: name, warehouse: wh, qty: 0, usd_amount: 0 };
      if (!it.name && name) it.name = name;
      it.qty += qty; it.usd_amount += usd;
    }

    // 快照日期 = 不良品倉列裡出現最多次的「年月日」
    var dates = Object.keys(dateCount).sort(function (a, b) { return dateCount[b] - dateCount[a] || (a < b ? 1 : -1); });
    var date = dates[0] || null;
    var dateSource = 'column';
    if (dates.length > 1) warnings.push('檔案內「年月日」有多個日期(' + dates.join('、') + '),採用最多列的 ' + date);
    if (!date) {
      date = dateFromFileName(opts.fileName, opts.defaultYear || new Date().getFullYear());
      dateSource = 'filename';
      if (date) warnings.push('「年月日」欄位為空,改用檔名推得日期 ' + date);
    }
    if (!date) {
      return { ok: false, reason: 'no_date', message: '無法判斷這份檔案的快照日期(「年月日」為空,檔名也看不出日期)', warnings: warnings };
    }
    if (defectiveRows === 0) warnings.push('檔案裡沒有任何「不良品倉」的列,這週金額會記為 0');
    if (unassigned.length) {
      warnings.push(unassigned.length + ' 筆只標示「不良品倉」沒有寫大園/新竹,未計入任何倉別(' +
        unassigned.map(function (u) { return u.part_no + ' US$' + u.usd_amount; }).join('、') + ')');
    }

    WAREHOUSES.forEach(function (w) {
      totals[w].usd_amount = round2(totals[w].usd_amount);
      totals[w].quantity = roundQty(totals[w].quantity);
    });
    var items = Object.keys(itemsMap).sort().map(function (id) {
      var it = itemsMap[id];
      return { id: it.id, part_no: it.part_no, name: it.name, warehouse: it.warehouse, qty: roundQty(it.qty), usd_amount: round2(it.usd_amount) };
    });

    return {
      ok: true,
      warnings: warnings,
      snapshot: {
        date: date,
        week: weekLabel(date),
        date_source: dateSource,
        totals: totals,
        items: items,
        unassigned: unassigned
      }
    };
  }

  // ---------- 狀態 ----------
  function emptyState() {
    return { schema_version: SCHEMA_VERSION, updated_at: nowIso(), weeks: {}, annotations: {}, sources: {}, removed_weeks: {}, item_photos: {}, wms_suppliers: {} };
  }

  function normalizeState(s) {
    var st = s && typeof s === 'object' ? clone(s) : emptyState();
    if (!st.weeks || typeof st.weeks !== 'object') st.weeks = {};
    if (!st.annotations || typeof st.annotations !== 'object') st.annotations = {};
    // v1 → v2:補上料號照片(照片檔在資料夾的「不良品倉照片」;這裡只記錄 {id, item_id, name, taken_at, updated_at, deleted, source})
    if (!st.item_photos || typeof st.item_photos !== 'object') st.item_photos = {};
    // 料號 → 供應商(從「條碼庫存明細」報表讀來;手動填的供應商在 annotations,優先)
    if (!st.wms_suppliers || typeof st.wms_suppliers !== 'object') st.wms_suppliers = {};
    if (!st.sources || typeof st.sources !== 'object') st.sources = {};
    if (!st.removed_weeks || typeof st.removed_weeks !== 'object') st.removed_weeks = {};
    st.schema_version = SCHEMA_VERSION;
    return st;
  }

  /**
   * 把一份快照放進 state(就地修改 state)。
   * 同一天已經有資料時:來源是同一個檔案、或新檔案的修改時間 >= 舊的,就取代;否則保留舊的。
   * @param source {path, size, lastModified}
   * @returns {{action:'added'|'replaced'|'kept_existing', previous?:object}}
   */
  function applySnapshot(state, snapshot, source) {
    var existing = state.weeks[snapshot.date];
    var src = source || {};
    var week = clone(snapshot);
    week.source_file = src.path || '';
    week.source_size = src.size || 0;
    week.source_mtime = src.lastModified || 0;
    week.imported_at = nowIso();
    if (state.removed_weeks && state.removed_weeks[snapshot.date]) delete state.removed_weeks[snapshot.date];
    if (!existing) {
      state.weeks[snapshot.date] = week;
      state.updated_at = nowIso();
      return { action: 'added' };
    }
    var sameFile = existing.source_file && existing.source_file === week.source_file;
    if (sameFile || (week.source_mtime || 0) >= (existing.source_mtime || 0)) {
      state.weeks[snapshot.date] = week;
      state.updated_at = nowIso();
      return { action: 'replaced', previous: { source_file: existing.source_file } };
    }
    return { action: 'kept_existing', previous: { source_file: existing.source_file } };
  }

  /** 移除一週。留下「墓碑」(removed_weeks),合併其他較舊的 state 時才不會又被加回來。 */
  function removeWeek(state, dateKey) {
    if (!state.weeks[dateKey]) return false;
    delete state.weeks[dateKey];
    state.removed_weeks = state.removed_weeks || {};
    state.removed_weeks[dateKey] = nowIso();
    state.updated_at = state.removed_weeks[dateKey];
    return true;
  }

  var ANNOTATION_FIELDS = ['status', 'owner', 'eta', 'supplier'];

  /** 料號照片:加入或更新一筆紀錄(就地修改 state)。照片檔本身由呼叫端寫進資料夾。 */
  function putItemPhoto(state, rec) {
    var cur = state.item_photos[rec.id];
    if (cur && String(cur.updated_at || '') >= String(rec.updated_at || '')) return false;
    state.item_photos[rec.id] = clone(rec);
    if (String(rec.updated_at || '') > String(state.updated_at || '')) state.updated_at = rec.updated_at;
    return true;
  }
  /** 某料號目前的照片(新到舊) */
  function itemPhotos(state, itemId) {
    return Object.keys(state.item_photos || {}).map(function (k) { return state.item_photos[k]; })
      .filter(function (p) { return p.item_id === itemId && !p.deleted; })
      .sort(function (a, b) { return String(a.taken_at || '') < String(b.taken_at || '') ? 1 : -1; });
  }

  function hasAnnotation(a) { return !!(a && (a.status || a.owner || a.eta || a.supplier)); }
  function latestItemIds(state) {
    var d = Object.keys(state.weeks).sort().pop(), out = {};
    if (d) (state.weeks[d].items || []).forEach(function (it) { out[it.id] = true; });
    return out;
  }
  /**
   * 料號換倉(例如新竹調撥回大園)時,項目 id 會從 -HC 變成 -DY。
   * 新 id 還沒有自己的註記、而同料號另一倉的舊 id 已經不在最新週 → 沿用舊 id 的註記與照片,使用者不用重新輸入。
   * @returns {{ annotation: object|null, from: string|null }}
   */
  function effectiveAnnotation(state, id, latestIds) {
    var own = state.annotations[id];
    if (hasAnnotation(own)) return { annotation: own, from: null };
    latestIds = latestIds || latestItemIds(state);
    var m = /^(.*)-(DY|HC)$/.exec(id);
    if (m) {
      var codes = Object.keys(WH_CODE).map(function (w) { return WH_CODE[w]; });
      for (var i = 0; i < codes.length; i++) {
        var sid = m[1] + '-' + codes[i];
        if (sid !== id && !latestIds[sid] && hasAnnotation(state.annotations[sid])) return { annotation: state.annotations[sid], from: sid };
      }
    }
    return { annotation: own || null, from: null };
  }

  function setAnnotation(state, id, field, value) {
    if (ANNOTATION_FIELDS.indexOf(field) === -1) throw new Error('unknown field ' + field);
    // 以畫面上看到的註記為基礎(含換倉沿用的),只改這一欄
    var base = effectiveAnnotation(state, id).annotation;
    var a = base ? clone(base) : { status: '', owner: '', eta: '', supplier: '' };
    a[field] = value == null ? '' : String(value);
    a.updated_at = nowIso();
    state.annotations[id] = a;
    state.updated_at = a.updated_at;
    return a;
  }

  /**
   * 合併兩份 state(例如:資料夾裡的 tracker-data.json + 瀏覽器 localStorage)。
   * 週資料:同一天取 imported_at 較新者。註記:同一料號取 updated_at 較新者。sources:取較新的 checked_at。
   */
  function mergeStates(a, b) {
    var A = normalizeState(a), B = normalizeState(b);
    var out = normalizeState(A);
    Object.keys(B.weeks).forEach(function (d) {
      var x = out.weeks[d], y = B.weeks[d];
      if (!x || String(y.imported_at || '') > String(x.imported_at || '')) out.weeks[d] = clone(y);
    });
    Object.keys(B.annotations).forEach(function (id) {
      var x = out.annotations[id], y = B.annotations[id];
      if (!x || String(y.updated_at || '') > String(x.updated_at || '')) out.annotations[id] = clone(y);
    });
    Object.keys(B.wms_suppliers).forEach(function (pn) {
      var x = out.wms_suppliers[pn], y = B.wms_suppliers[pn];
      var newer = !x || Number(y.source_mtime || 0) > Number(x.source_mtime || 0) ||
        (Number(y.source_mtime || 0) === Number(x.source_mtime || 0) && String(y.imported_at || '') > String(x.imported_at || ''));
      if (newer) out.wms_suppliers[pn] = clone(y);
    });
    Object.keys(B.item_photos).forEach(function (id) {
      var x = out.item_photos[id], y = B.item_photos[id];
      if (!x || String(y.updated_at || '') > String(x.updated_at || '')) out.item_photos[id] = clone(y);
    });
    Object.keys(B.sources).forEach(function (p) {
      var x = out.sources[p], y = B.sources[p];
      if (!x || String(y.checked_at || '') > String(x.checked_at || '')) out.sources[p] = clone(y);
    });
    // 墓碑:取較晚的移除時間;移除時間比該週匯入時間晚 → 該週維持移除
    Object.keys(B.removed_weeks).forEach(function (d) {
      if (String(B.removed_weeks[d]) > String(out.removed_weeks[d] || '')) out.removed_weeks[d] = B.removed_weeks[d];
    });
    Object.keys(out.removed_weeks).forEach(function (d) {
      var w = out.weeks[d];
      if (!w) return;
      // >=:移除一定發生在該週匯入之後(同一毫秒也算移除)
      if (String(out.removed_weeks[d]) >= String(w.imported_at || '')) delete out.weeks[d];
      else delete out.removed_weeks[d];
    });
    out.updated_at = String(A.updated_at || '') > String(B.updated_at || '') ? A.updated_at : B.updated_at;
    return out;
  }

  // ---------- 推導 ----------
  function sortedWeeks(state) {
    return Object.keys(state.weeks).sort().map(function (d) { return state.weeks[d]; });
  }

  function weeksInRange(state, start, end) {
    return sortedWeeks(state).filter(function (w) {
      return (!start || w.date >= start) && (!end || w.date <= end);
    });
  }

  /** 對應雲端版的 weekly_amounts collection:每週 × 每倉一列 */
  function weeklyRows(state) {
    var out = [];
    sortedWeeks(state).forEach(function (w) {
      WAREHOUSES.forEach(function (wh) {
        var t = (w.totals && w.totals[wh]) || { usd_amount: 0, quantity: 0, line_count: 0 };
        out.push({ date: w.date, week: w.week, warehouse: wh, usd_amount: t.usd_amount, quantity: t.quantity, line_count: t.line_count });
      });
    });
    return out;
  }

  /** 最新一週 vs 前一週 */
  function summarize(rows, warehouse) {
    var r = rows.filter(function (d) { return d.warehouse === warehouse; }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    if (!r.length) return null;
    var latest = r[r.length - 1];
    var prev = r.length > 1 ? r[r.length - 2] : null;
    var delta = prev ? round2(latest.usd_amount - prev.usd_amount) : null;
    var deltaPct = prev && prev.usd_amount ? delta / prev.usd_amount : null;
    return { latest: latest, prev: prev, delta: delta, deltaPct: deltaPct };
  }

  /**
   * 對應雲端版的 items collection:每個料號×倉別一列,數量/金額取「最後一次出現那週」的值,
   * 再合併處理狀態/Owner/預計完成時間註記。已經消失的料號仍會保留(顯示最後一次出現的數字)。
   */
  function itemRows(state) {
    var map = {};
    sortedWeeks(state).forEach(function (w) {
      (w.items || []).forEach(function (it) {
        var cur = map[it.id];
        if (!cur) cur = map[it.id] = { id: it.id, first_seen_week: w.week, first_seen_date: w.date };
        cur.part_no = it.part_no;
        cur.name = it.name || cur.name || '';
        cur.warehouse = it.warehouse;
        cur.qty = it.qty;
        cur.usd_amount = it.usd_amount;
        cur.updated_week = w.week;
        cur.last_seen_date = w.date;
      });
    });
    var latestDate = Object.keys(state.weeks).sort().pop() || null;
    var latestIds = latestItemIds(state);
    return Object.keys(map).map(function (id) {
      var r = map[id];
      r.in_latest_week = r.last_seen_date === latestDate;
      // 註記跟著料號走:換倉時沿用另一倉(已轉出)的註記
      var eff = r.in_latest_week ? effectiveAnnotation(state, id, latestIds) : { annotation: state.annotations[id], from: null };
      var a = eff.annotation || {};
      r.ann_from = eff.from;
      r.ann_updated_at = a.updated_at || '';
      r.status = a.status || '';
      r.owner = a.owner || '';
      r.eta = a.eta || '';
      var ws = state.wms_suppliers[r.part_no];
      r.wms_supplier = ws ? ws.name : '';
      r.supplier = a.supplier || r.wms_supplier;            // 手動填的優先
      r.supplier_source = a.supplier ? 'manual' : (ws ? 'wms' : '');
      var photos = itemPhotos(state, id);
      if (eff.from) photos = photos.concat(itemPhotos(state, eff.from));
      r.photos = photos.map(function (p) { return p.name; });
      return r;
    });
  }
  /** 項目明細主表:只有最新一週還在不良品倉的料號(數量、金額都是最新一週的) */
  function currentItemRows(state) { return itemRows(state).filter(function (r) { return r.in_latest_week; }); }
  /** 已轉出不良品倉:最新一週已經沒有、但最近 keepDays 天內還出現過的料號(備查,不算庫存) */
  function transferredItemRows(state, keepDays) {
    var latestDate = Object.keys(state.weeks).sort().pop();
    if (!latestDate) return [];
    var limit = keepDays == null ? 56 : keepDays;
    var t0 = Date.parse(latestDate);
    return itemRows(state).filter(function (r) {
      return !r.in_latest_week && (t0 - Date.parse(r.last_seen_date)) / 86400000 <= limit;
    }).sort(function (a, b) { return a.last_seen_date < b.last_seen_date ? 1 : (a.last_seen_date > b.last_seen_date ? -1 : (itemSortKey(a) < itemSortKey(b) ? -1 : 1)); });
  }

  function itemSortKey(it) { return it.warehouse + it.part_no; }
  function sortItems(list) { return list.slice().sort(function (a, b) { var x = itemSortKey(a), y = itemSortKey(b); return x < y ? -1 : (x > y ? 1 : 0); }); }

  /**
   * 本區間項目變動:比較區間內最早與最新兩週。
   * 新增 = 最新週有、最早週沒有(數字取最新週);減少 = 最早週有、最新週沒有(數字取區間內最後一次出現那週)。
   */
  function rangeDiff(state, start, end) {
    var weeks = weeksInRange(state, start, end);
    if (weeks.length < 2) return { insufficient: true };
    var first = weeks[0], last = weeks[weeks.length - 1];
    var firstMap = {}, lastMap = {};
    (first.items || []).forEach(function (it) { firstMap[it.id] = it; });
    (last.items || []).forEach(function (it) { lastMap[it.id] = it; });
    var added = (last.items || []).filter(function (it) { return !firstMap[it.id]; }).map(clone);
    var removed = (first.items || []).filter(function (it) { return !lastMap[it.id]; }).map(function (it) {
      var latest = it;
      for (var i = weeks.length - 1; i >= 0; i--) {
        var hit = (weeks[i].items || []).filter(function (x) { return x.id === it.id; })[0];
        if (hit) { latest = hit; break; }
      }
      return clone(latest);
    });
    var byId = function (a, b) { return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0); };
    return { insufficient: false, first: first, last: last, added: added.sort(byId), removed: removed.sort(byId) };
  }

  /**
   * 金額降低來源(去化明細):區間內最早週 → 最新週,同一料號金額變少的都列出;
   * 消失的料號視為降到 0。依降低金額由大到小。
   */
  function declineSources(state, start, end) {
    var weeks = weeksInRange(state, start, end);
    if (weeks.length < 2) return { insufficient: true };
    var first = weeks[0], last = weeks[weeks.length - 1];
    var lastMap = {};
    (last.items || []).forEach(function (it) { lastMap[it.id] = it; });
    var out = [];
    (first.items || []).forEach(function (fi) {
      var li = lastMap[fi.id];
      var startAmt = fi.usd_amount || 0;
      var endAmt = li ? (li.usd_amount || 0) : 0;
      var decline = round2(startAmt - endAmt);
      if (decline > 0.005) {
        out.push({
          id: fi.id, part_no: fi.part_no, warehouse: fi.warehouse, name: fi.name,
          startQty: fi.qty, endQty: li ? li.qty : 0,
          startAmt: startAmt, endAmt: endAmt, decline: decline, removed: !li
        });
      }
    });
    out.sort(function (a, b) { return b.decline - a.decline || (a.id < b.id ? -1 : 1); });
    return { insufficient: false, first: first, last: last, items: out };
  }

  /**
   * 本區間項目變動(單一表格):區間內最早週 → 最新週,每個有變動的料號一列。
   * type:新增(期初沒有)/ 轉出(期末沒有)/ 增加 / 減少(兩週都有,金額變多或變少)。數量和金額都沒變的不列。
   * 期初沒有 → 期初 0;期末沒有 → 期末 0。所以 Σ amtChange = 期末總額 − 期初總額 = 圖表的變動。
   * @param warehouse 'all' 或倉別
   */
  function rangeChanges(state, start, end, warehouse) {
    var weeks = weeksInRange(state, start, end);
    if (weeks.length < 2) return { insufficient: true };
    var first = weeks[0], last = weeks[weeks.length - 1];
    var pick = function (it) { return !warehouse || warehouse === 'all' || it.warehouse === warehouse; };
    var fm = {}, lm = {};
    (first.items || []).filter(pick).forEach(function (it) { fm[it.id] = it; });
    (last.items || []).filter(pick).forEach(function (it) { lm[it.id] = it; });
    var ids = Object.keys(fm).concat(Object.keys(lm).filter(function (id) { return !fm[id]; }));
    var items = [], sums = { '新增': 0, '增加': 0, '減少': 0, '轉出': 0 }, counts = { '新增': 0, '增加': 0, '減少': 0, '轉出': 0 };
    var startTotal = 0, endTotal = 0;
    ids.forEach(function (id) {
      var f = fm[id], l = lm[id], ref = l || f;
      var sq = f ? (f.qty || 0) : 0, eq = l ? (l.qty || 0) : 0;
      var sa = f ? (f.usd_amount || 0) : 0, ea = l ? (l.usd_amount || 0) : 0;
      startTotal += sa; endTotal += ea;
      var dq = roundQty(eq - sq), da = round2(ea - sa);
      if (!dq && !da && f && l) return;
      var type = !f ? '新增' : (!l ? '轉出' : (da > 0 || (da === 0 && dq > 0) ? '增加' : '減少'));
      sums[type] = round2(sums[type] + da); counts[type]++;
      items.push({ id: id, part_no: ref.part_no, warehouse: ref.warehouse, name: ref.name || (f && f.name) || '',
        type: type, startQty: sq, endQty: eq, qtyChange: dq, startAmt: sa, endAmt: ea, amtChange: da });
    });
    items.sort(function (a, b) { return Math.abs(b.amtChange) - Math.abs(a.amtChange) || (a.id < b.id ? -1 : 1); });
    var total = round2(items.reduce(function (s, it) { return s + it.amtChange; }, 0));
    return {
      insufficient: false, first: first, last: last, items: items, sums: sums, counts: counts,
      startTotal: round2(startTotal), endTotal: round2(endTotal), change: round2(endTotal - startTotal), total: total
    };
  }
  return {
    SCHEMA_VERSION: SCHEMA_VERSION,
    WAREHOUSES: WAREHOUSES,
    WH_CODE: WH_CODE,
    HEADERS: HEADERS,
    round2: round2,
    toDateKey: toDateKey,
    dateFromFileName: dateFromFileName,
    weekLabel: weekLabel,
    dateKeyCompact: dateKeyCompact,
    itemId: itemId,
    resolveWarehouse: resolveWarehouse,
    findHeader: findHeader,
    parseSnapshot: parseSnapshot,
    emptyState: emptyState,
    normalizeState: normalizeState,
    applySnapshot: applySnapshot,
    removeWeek: removeWeek,
    setAnnotation: setAnnotation,
    ANNOTATION_FIELDS: ANNOTATION_FIELDS,
    putItemPhoto: putItemPhoto,
    parseSupplierReport: parseSupplierReport,
    applySupplierReport: applySupplierReport,
    itemPhotos: itemPhotos,
    mergeStates: mergeStates,
    sortedWeeks: sortedWeeks,
    weeksInRange: weeksInRange,
    weeklyRows: weeklyRows,
    summarize: summarize,
    itemRows: itemRows,
    sortItems: sortItems,
    rangeDiff: rangeDiff,
    declineSources: declineSources,
    rangeChanges: rangeChanges,
    currentItemRows: currentItemRows,
    transferredItemRows: transferredItemRows,
    effectiveAnnotation: effectiveAnnotation,
    clone: clone
  };
});
