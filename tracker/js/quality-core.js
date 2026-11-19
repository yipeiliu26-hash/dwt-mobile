/*
 * quality-core.js — 品質異常紀錄的純邏輯(無 DOM)
 *
 * 桌面版(js/quality-app.js)和手機版(mobile/quality-mobile.src.html,建置時內嵌)共用這一份。
 * Node 可直接 require(tests/quality.test.js)。
 *
 * 資料模型
 *   issue(一筆異常):{ id, supplier, item, part_no, defect, defect_qty, total_qty, scar, due,
 *                      status, containment, created_at, updated_at, field_times{欄位: ISO}, deleted }
 *   log(現場紀錄):  { id, issue_id, date, type, note, photos[檔名], status_to, source,
 *                      created_at, updated_at, deleted }
 *
 * 同步方式(公司 OneDrive,不需要伺服器):
 *   手機 → 電腦:手機產生「傳送包」(makePackage)存到 OneDrive 品質異常/收件匣/,桌面版 applyPackage 合併。
 *   電腦 → 手機:桌面版寫 品質異常/手機清單.json(makeSnapshot),手機載入後 rebasePending 去掉已確認的變更。
 *
 * 合併規則:異常是「逐欄位」比較 field_times,較新的勝出(手機改狀態、電腦改對策不會互相蓋掉);
 *           紀錄以 updated_at 較新者為準;刪除是軟刪除(deleted: true),合併時不會被舊資料加回。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DWTQualityCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // v2(2026-10-22):issue 加 origin(異常來源)、origin_other(選「其他」時自填,最多 40 字)。
  //   v1 資料讀進來時補空值(顯示「未指定」);舊資料沒有這兩欄的 field_times,合併時不會蓋掉新資料(見 fieldTime)。
  //   注意:log 的 source(phone/desktop)是另一回事,不要混用。
  var SCHEMA_VERSION = 2;
  var STATUSES = ['調查中', '待回覆', '結案'];
  var SCAR_STATES = ['未開', '已開', '已回覆', '已結案'];
  var LOG_TYPES = ['現場觀察', '量測', '會議'];
  var ORIGINS = ['IQC', 'IPQC', '產線', '客戶', '出貨檢驗', '其他'];
  var ORIGIN_OTHER = '其他';
  var ORIGIN_OTHER_MAX = 40;
  var ORIGIN_NONE = '__none__';   // 篩選用:未指定
  var ISSUE_FIELDS = ['supplier', 'item', 'part_no', 'defect', 'defect_qty', 'total_qty', 'scar', 'due', 'status', 'containment', 'origin', 'origin_other', 'deleted'];
  var LOG_FIELDS = ['issue_id', 'date', 'type', 'note', 'photos', 'status_to', 'source', 'created_at', 'deleted'];
  var ID_RE = /^[A-Za-z0-9_-]{4,64}$/;
  var PHOTO_RE = /^[A-Za-z0-9_-]{1,80}\.jpg$/;

  // ---------- 小工具 ----------
  function nowIso() { return new Date().toISOString(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  /** 本地日期 YYYY-MM-DD */
  function localDate(d) { d = d || new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function shortDate(s) { var m = /^\d{4}-(\d{2})-(\d{2})$/.exec(s || ''); return m ? (+m[1]) + '/' + (+m[2]) : (s || ''); }
  function newId(prefix) {
    var rnd = '';
    for (var i = 0; i < 6; i++) rnd += Math.floor(Math.random() * 36).toString(36);
    return (prefix || 'X') + Date.now().toString(36) + rnd;
  }
  function str(v, max) { return v == null ? '' : String(v).slice(0, max || 2000); }
  function numOrEmpty(v) {
    if (v === '' || v == null) return '';
    var n = Number(v);
    return isFinite(n) && n >= 0 ? n : '';
  }
  function dateOrEmpty(v) { return /^\d{4}-\d{2}-\d{2}$/.test(v || '') ? v : ''; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function emptyState() {
    return { schema_version: SCHEMA_VERSION, updated_at: '', issues: {}, logs: {}, processed_packages: {} };
  }
  function normalizeState(s) {
    var out = emptyState();
    if (!s || typeof s !== 'object') return out;
    out.updated_at = str(s.updated_at, 40);
    Object.keys(s.issues || {}).forEach(function (id) { var it = cleanIssue(s.issues[id]); if (it) out.issues[it.id] = it; });
    Object.keys(s.logs || {}).forEach(function (id) { var lg = cleanLog(s.logs[id]); if (lg) out.logs[lg.id] = lg; });
    Object.keys(s.processed_packages || {}).forEach(function (id) { out.processed_packages[id] = s.processed_packages[id]; });
    return out;
  }

  // ---------- 清理外來資料(手機傳送包、檔案)----------
  function cleanIssue(x) {
    if (!x || !ID_RE.test(x.id || '')) return null;
    var ft = {};
    Object.keys(x.field_times || {}).forEach(function (k) { if (ISSUE_FIELDS.indexOf(k) >= 0) ft[k] = str(x.field_times[k], 40); });
    var origin = ORIGINS.indexOf(x.origin) >= 0 ? x.origin : '';
    return {
      id: x.id,
      supplier: str(x.supplier, 100),
      item: str(x.item, 100),
      part_no: str(x.part_no, 60),
      defect: str(x.defect, 300),
      defect_qty: numOrEmpty(x.defect_qty),
      total_qty: numOrEmpty(x.total_qty),
      scar: SCAR_STATES.indexOf(x.scar) >= 0 ? x.scar : '未開',
      due: dateOrEmpty(x.due),
      status: STATUSES.indexOf(x.status) >= 0 ? x.status : '調查中',
      containment: str(x.containment, 300),
      origin: origin,
      // 只有來源是「其他」時才保留自填文字
      origin_other: origin === ORIGIN_OTHER ? str(x.origin_other, ORIGIN_OTHER_MAX).trim() : '',
      created_at: str(x.created_at, 40),
      updated_at: str(x.updated_at, 40),
      field_times: ft,
      deleted: !!x.deleted
    };
  }
  function cleanLog(x) {
    if (!x || !ID_RE.test(x.id || '') || !ID_RE.test(x.issue_id || '')) return null;
    return {
      id: x.id,
      issue_id: x.issue_id,
      date: dateOrEmpty(x.date),
      type: LOG_TYPES.indexOf(x.type) >= 0 ? x.type : '現場觀察',
      note: str(x.note, 4000),
      photos: (Array.isArray(x.photos) ? x.photos : []).filter(function (p) { return PHOTO_RE.test(p); }).slice(0, 20),
      status_to: STATUSES.indexOf(x.status_to) >= 0 ? x.status_to : '',
      source: x.source === 'phone' ? 'phone' : 'desktop',
      created_at: str(x.created_at, 40),
      updated_at: str(x.updated_at, 40),
      deleted: !!x.deleted
    };
  }

  // ---------- 建立 / 修改 ----------
  function newIssue(fields, at) {
    at = at || nowIso();
    var it = cleanIssue(Object.assign({ id: newId('Q'), created_at: at, updated_at: at }, fields || {}));
    ISSUE_FIELDS.forEach(function (f) { it.field_times[f] = at; });
    return it;
  }
  /** 改一個欄位(會記錄該欄位的修改時間)。回傳新物件,不改原本的。 */
  function setIssueField(issue, field, value, at) {
    var ch = {}; ch[field] = value;
    return setIssueFields(issue, ch, at);
  }
  /**
   * 一次改多個欄位(同一個修改時間)。改來源時一定要和 origin_other 一起改,
   * 例如 { origin: 'IQC', origin_other: '' } 或 { origin: '其他', origin_other: '供應商自檢' }。
   */
  function setIssueFields(issue, changes, at) {
    at = at || nowIso();
    var raw = clone(issue);
    Object.keys(changes).forEach(function (f) {
      if (ISSUE_FIELDS.indexOf(f) < 0) throw new Error('未知欄位:' + f);
      raw[f] = changes[f];
    });
    // 來源改成「其他」以外的選項 → 自填文字一併清空(也記錄時間,合併時才會一起生效)
    if ('origin' in changes && changes.origin !== ORIGIN_OTHER) raw.origin_other = '';
    var it = cleanIssue(raw);
    Object.keys(changes).forEach(function (f) { it.field_times[f] = at; });
    if ('origin' in changes) it.field_times.origin_other = at;
    it.updated_at = at;
    return it;
  }

  // ---------- 異常來源 ----------
  /** 顯示文字:未指定 / IQC … / 其他:自填內容 */
  function originLabel(issue) {
    if (!issue || !issue.origin) return '未指定';
    if (issue.origin === ORIGIN_OTHER) return issue.origin_other ? ORIGIN_OTHER + ':' + issue.origin_other : ORIGIN_OTHER;
    return issue.origin;
  }
  /**
   * 表單檢查。isNew = 新增異常(來源必填);編輯舊資料可以維持未指定。
   * @returns {string} 錯誤訊息,沒問題回 ''
   */
  function validateOrigin(origin, other, isNew) {
    if (!origin) return isNew ? '請選擇異常來源' : '';
    if (ORIGINS.indexOf(origin) < 0) return '異常來源不在選項內';
    if (origin === ORIGIN_OTHER) {
      var t = String(other || '').trim();
      if (!t) return '選「其他」時請填寫來源說明';
      if (t.length > ORIGIN_OTHER_MAX) return '來源說明最多 ' + ORIGIN_OTHER_MAX + ' 個字';
    }
    return '';
  }
  function newLog(fields, at) {
    at = at || nowIso();
    return cleanLog(Object.assign({ id: newId('L'), date: localDate(), created_at: at, updated_at: at }, fields || {}));
  }

  // ---------- 合併 ----------
  /**
   * 欄位的修改時間。有 field_times 的資料(v1 起都有)只看該欄位自己的時間,沒記錄就當作最舊;
   * 這樣 v1 資料(沒有 origin 的時間)或舊版手機傳來的包,不會用 updated_at 蓋掉別處填好的來源。
   * 完全沒有 field_times 的極舊資料才退回 updated_at。
   */
  function fieldTime(it, f) {
    var ft = it.field_times || {};
    if (Object.keys(ft).length) return ft[f] || '';
    return it.updated_at || '';
  }
  function mergeIssue(a, b) {
    if (!a) return clone(b);
    if (!b) return clone(a);
    var out = clone(a);
    ISSUE_FIELDS.forEach(function (f) {
      if (fieldTime(b, f) > fieldTime(a, f)) { out[f] = b[f]; out.field_times[f] = fieldTime(b, f); }
    });
    if ((b.updated_at || '') > (out.updated_at || '')) out.updated_at = b.updated_at;
    if (!out.created_at || (b.created_at && b.created_at < out.created_at)) out.created_at = b.created_at;
    return out;
  }
  function mergeStates(a, b) {
    var out = normalizeState(a);
    b = normalizeState(b);
    Object.keys(b.issues).forEach(function (id) { out.issues[id] = mergeIssue(out.issues[id], b.issues[id]); });
    Object.keys(b.logs).forEach(function (id) {
      var cur = out.logs[id];
      if (!cur || (b.logs[id].updated_at || '') > (cur.updated_at || '')) out.logs[id] = clone(b.logs[id]);
    });
    Object.keys(b.processed_packages).forEach(function (id) { if (!out.processed_packages[id]) out.processed_packages[id] = b.processed_packages[id]; });
    if ((b.updated_at || '') > (out.updated_at || '')) out.updated_at = b.updated_at;
    return out;
  }

  // ---------- 推導 ----------
  function daysBetween(fromYmd, toYmd) {
    var a = Date.UTC(+fromYmd.slice(0, 4), +fromYmd.slice(5, 7) - 1, +fromYmd.slice(8, 10));
    var b = Date.UTC(+toYmd.slice(0, 4), +toYmd.slice(5, 7) - 1, +toYmd.slice(8, 10));
    return Math.round((b - a) / 86400000);
  }
  /** 逾期天數;沒逾期回 0。期限當天不算逾期。 */
  function overdueDays(issue, today) {
    if (!issue.due || issue.status === '結案') return 0;
    var d = daysBetween(issue.due, today || localDate());
    return d > 0 ? d : 0;
  }
  /** 清單上的狀態標籤 { label, tone: bad|warn|info|done } */
  function badge(issue, today) {
    var od = overdueDays(issue, today);
    if (od) return { label: '逾期', tone: 'bad', days: od };
    if (issue.status === '待回覆') return { label: '待回覆', tone: 'warn' };
    if (issue.status === '結案') return { label: '結案', tone: 'done' };
    return { label: '調查中', tone: 'info' };
  }
  function qtyText(issue) {
    if (issue.defect_qty === '' && issue.total_qty === '') return '';
    return (issue.defect_qty === '' ? '?' : issue.defect_qty) + '/' + (issue.total_qty === '' ? '?' : issue.total_qty);
  }
  function title(issue) { return [issue.supplier, issue.item].filter(Boolean).join(' ') || '(未命名異常)'; }

  function liveIssues(state) {
    return Object.keys(state.issues).map(function (k) { return state.issues[k]; }).filter(function (it) { return !it.deleted; });
  }
  function counts(state, today) {
    var list = liveIssues(state);
    return {
      open: list.filter(function (it) { return it.status !== '結案'; }).length,
      overdue: list.filter(function (it) { return overdueDays(it, today) > 0; }).length,
      all: list.length
    };
  }
  /**
   * filter: open | overdue | all;q 搜尋供應商、品名、料號、不良描述。
   * 排序:逾期(天數多的在前)→ 未結案依期限近到遠 → 結案依更新時間新到舊。
   */
  /** 來源篩選:'' = 全部;ORIGIN_NONE = 未指定;其餘 = ORIGINS 之一 */
  function matchOrigin(it, origin) {
    if (!origin) return true;
    if (origin === ORIGIN_NONE) return !it.origin;
    return it.origin === origin;
  }
  function filterIssues(state, filter, q, today, origin) {
    today = today || localDate();
    q = (q || '').trim().toLowerCase();
    var list = liveIssues(state).filter(function (it) {
      if (filter === 'open' && it.status === '結案') return false;
      if (filter === 'overdue' && !overdueDays(it, today)) return false;
      if (!matchOrigin(it, origin)) return false;
      if (!q) return true;
      return [it.supplier, it.item, it.part_no, it.defect, originLabel(it)].join(' ').toLowerCase().indexOf(q) >= 0;
    });
    function rank(it) { return it.status === '結案' ? 2 : (overdueDays(it, today) ? 0 : 1); }
    list.sort(function (a, b) {
      var ra = rank(a), rb = rank(b);
      if (ra !== rb) return ra - rb;
      if (ra === 0) return overdueDays(b, today) - overdueDays(a, today);
      if (ra === 1) return (a.due || '9999') < (b.due || '9999') ? -1 : (a.due || '9999') > (b.due || '9999') ? 1 : 0;
      return (b.updated_at || '') < (a.updated_at || '') ? -1 : 1;
    });
    return list;
  }
  /** 各來源的筆數(套用狀態篩選後),給篩選 chips 顯示:{ '': 全部, IQC: n, …, __none__: n } */
  function originCounts(state, filter, today) {
    var list = filterIssues(state, filter, '', today);
    var out = { '': list.length };
    ORIGINS.forEach(function (o) { out[o] = 0; });
    out[ORIGIN_NONE] = 0;
    list.forEach(function (it) { out[it.origin || ORIGIN_NONE]++; });
    return out;
  }

  /** 某筆異常的現場紀錄,新到舊 */
  function issueLogs(state, issueId) {
    return Object.keys(state.logs).map(function (k) { return state.logs[k]; })
      .filter(function (l) { return l.issue_id === issueId && !l.deleted; })
      .sort(function (a, b) {
        if (a.date !== b.date) return a.date < b.date ? 1 : -1;
        return (a.created_at || '') < (b.created_at || '') ? 1 : -1;
      });
  }

  // ---------- 不良品倉項目明細(手機版主畫面)----------
  // 註記 = 看板的 state.annotations[項目id] = { status, owner, eta, updated_at },整筆以 updated_at 較新者為準(與看板 Core.mergeStates 相同)
  var ITEM_ID_RE = /^[A-Za-z0-9._-]{1,80}$/;
  function cleanAnnotation(x) {
    if (!x || typeof x !== 'object') return null;
    return { status: str(x.status, 200), owner: str(x.owner, 100), eta: dateOrEmpty(x.eta), supplier: str(x.supplier, 100), updated_at: str(x.updated_at, 40) };
  }
  /** 料號照片紀錄;照片檔名固定是 id + '.jpg' */
  function cleanItemPhoto(x) {
    if (!x || !ID_RE.test(x.id || '') || !ITEM_ID_RE.test(x.item_id || '') || x.name !== x.id + '.jpg') return null;
    return {
      id: x.id, item_id: x.item_id, name: x.name, taken_at: str(x.taken_at, 40), updated_at: str(x.updated_at, 40),
      deleted: !!x.deleted, source: x.source === 'phone' ? 'phone' : 'desktop'
    };
  }
  function newItemPhoto(itemId, source, at) {
    at = at || nowIso();
    var id = newId('I');
    return { id: id, item_id: itemId, name: id + '.jpg', taken_at: at, updated_at: at, deleted: false, source: source || 'desktop' };
  }
  function cleanAnnotations(map) {
    var out = {};
    Object.keys(map || {}).forEach(function (id) {
      if (!ITEM_ID_RE.test(id)) return;
      var a = cleanAnnotation(map[id]);
      if (a && a.updated_at) out[id] = a;
    });
    return out;
  }
  function num(v) { var n = Number(v); return isFinite(n) ? n : 0; }
  /** 看板 Core.itemRows 的結果 → 手機清單用的精簡版 */
  function trackerForSnapshot(rows, annotations, weeks) {
    return {
      latest_week: str(weeks && weeks.latest_week, 10),
      latest_date: dateOrEmpty(weeks && weeks.latest_date),
      items: (rows || []).map(function (r) {
        // 看板 itemRows 已算好「實際顯示的註記」(含料號換倉沿用),有就用它;舊版看板才退回 annotations
        var a = 'status' in r ? { status: r.status, owner: r.owner, eta: r.eta, supplier: r.supplier === r.wms_supplier && r.supplier_source !== 'manual' ? '' : r.supplier, updated_at: r.ann_updated_at }
          : ((annotations || {})[r.id] || {});
        return {
          id: r.id, part_no: r.part_no, name: r.name || '', warehouse: r.warehouse,
          qty: r.qty, usd_amount: r.usd_amount, last_seen_week: r.updated_week || '', in_latest_week: !!r.in_latest_week,
          status: a.status || '', owner: a.owner || '', eta: a.eta || '', ann_updated_at: a.updated_at || '',
          // supplier = 實際顯示的(手動填的優先,否則用 WMS 條碼庫存明細的);wms_supplier 讓手機知道清空手動值後要顯示什麼
          supplier: a.supplier || r.wms_supplier || '', wms_supplier: r.wms_supplier || '',
          photos: r.photos || []
        };
      })
    };
  }
  function cleanTracker(t) {
    if (!t || !Array.isArray(t.items)) return null;
    return {
      latest_week: str(t.latest_week, 10), latest_date: dateOrEmpty(t.latest_date),
      items: t.items.filter(function (r) { return r && ITEM_ID_RE.test(r.id || ''); }).map(function (r) {
        return {
          id: r.id, part_no: str(r.part_no, 60), name: str(r.name, 200), warehouse: str(r.warehouse, 20),
          qty: num(r.qty), usd_amount: num(r.usd_amount), last_seen_week: str(r.last_seen_week, 10), in_latest_week: !!r.in_latest_week,
          status: str(r.status, 200), owner: str(r.owner, 100), eta: dateOrEmpty(r.eta), supplier: str(r.supplier, 100), wms_supplier: str(r.wms_supplier, 100),
          ann_updated_at: str(r.ann_updated_at, 40),
          photos: (Array.isArray(r.photos) ? r.photos : []).filter(function (p) { return PHOTO_RE.test(p); })
        };
      })
    };
  }
  /** 手機顯示用:電腦清單的項目 + 手機上還沒傳送的註記與照片 */
  function trackerItems(tracker, pendingAnnotations, pendingItemPhotos) {
    if (!tracker) return [];
    var addBy = {}, removed = {};
    Object.keys(pendingItemPhotos || {}).forEach(function (k) {
      var p = pendingItemPhotos[k];
      if (p.deleted) removed[p.name] = true;
      else (addBy[p.item_id] = addBy[p.item_id] || []).push(p.name);
    });
    return tracker.items.map(function (r) {
      var out = r;
      var p = (pendingAnnotations || {})[r.id];
      if (p && (p.updated_at || '') > (r.ann_updated_at || '')) {
        out = Object.assign({}, out, { status: p.status, owner: p.owner, eta: p.eta, supplier: p.supplier || r.wms_supplier || '', ann_updated_at: p.updated_at });
      }
      if (addBy[r.id] || r.photos.some(function (n) { return removed[n]; })) {
        var names = (addBy[r.id] || []).slice().reverse().concat(r.photos).filter(function (n, i, arr) { return !removed[n] && arr.indexOf(n) === i; });
        out = Object.assign({}, out, { photos: names });
      }
      return out;
    }).sort(function (a, b) {
      var x = a.warehouse + a.part_no, y = b.warehouse + b.part_no;
      return x < y ? -1 : x > y ? 1 : 0;
    });
  }

  // ---------- 手機 → 電腦:傳送包 ----------
  /**
   * @param {object} p { issues[], logs[], photos{檔名: base64 JPEG}, thumbs{檔名: base64 JPEG}, annotations{項目id: 註記}, device }
   */
  function makePackage(p) {
    return {
      type: 'dwt-quality-package', version: 1, id: newId('P'),
      created_at: nowIso(), device: str(p.device, 60) || '手機',
      issues: p.issues || [], logs: p.logs || [], photos: p.photos || {}, thumbs: p.thumbs || {},
      annotations: p.annotations || {}, item_photos: p.item_photos || []
    };
  }
  /**
   * 把傳送包合併進 state(會改 state)。
   * @returns {{ ok, duplicate?, message?, issues, logs, photos: {檔名: base64}, thumbs: {檔名: base64} }}
   */
  function applyPackage(state, pkg) {
    if (!pkg || pkg.type !== 'dwt-quality-package' || pkg.version !== 1 || !ID_RE.test(pkg.id || '')) {
      return { ok: false, message: '不是品質異常傳送包' };
    }
    // 註記一律交回看板合併(以 updated_at 判斷,重複處理也不會出錯)
    var annotations = cleanAnnotations(pkg.annotations);
    // 料號照片同理交回看板(以 updated_at 判斷);照片檔另外回傳 item_files / item_thumbs
    var itemPhotos = (Array.isArray(pkg.item_photos) ? pkg.item_photos : []).map(cleanItemPhoto).filter(Boolean);
    var itemNames = {};
    itemPhotos.forEach(function (p) { if (!p.deleted) itemNames[p.name] = true; });
    function pickNames(map, names) {
      var out = {};
      Object.keys(map || {}).forEach(function (name) {
        if (PHOTO_RE.test(name) && names[name] && typeof map[name] === 'string') out[name] = map[name];
      });
      return out;
    }
    var itemFiles = pickNames(pkg.photos, itemNames), itemThumbs = pickNames(pkg.thumbs, itemNames);
    if (state.processed_packages[pkg.id]) {
      return { ok: true, duplicate: true, issues: 0, logs: 0, photos: {}, thumbs: {}, annotations: annotations, item_photos: itemPhotos, item_files: itemFiles, item_thumbs: itemThumbs };
    }
    var incoming = emptyState();
    (pkg.issues || []).forEach(function (x) { var it = cleanIssue(x); if (it) incoming.issues[it.id] = it; });
    (pkg.logs || []).forEach(function (x) { var lg = cleanLog(x); if (lg) { lg.source = 'phone'; incoming.logs[lg.id] = lg; } });
    var referenced = {};
    Object.keys(incoming.logs).forEach(function (id) { incoming.logs[id].photos.forEach(function (p) { referenced[p] = true; }); });
    function pick(map) {
      var out = {};
      Object.keys(map || {}).forEach(function (name) {
        if (PHOTO_RE.test(name) && referenced[name] && typeof map[name] === 'string') out[name] = map[name];
      });
      return out;
    }
    var merged = mergeStates(state, incoming);
    state.issues = merged.issues; state.logs = merged.logs;
    state.processed_packages[pkg.id] = {
      at: nowIso(), device: str(pkg.device, 60), created_at: str(pkg.created_at, 40),
      issues: Object.keys(incoming.issues).length, logs: Object.keys(incoming.logs).length,
      annotations: Object.keys(annotations).length, item_photos: itemPhotos.length
    };
    state.updated_at = nowIso();
    return {
      ok: true, issues: Object.keys(incoming.issues).length, logs: Object.keys(incoming.logs).length,
      photos: pick(pkg.photos), thumbs: pick(pkg.thumbs), annotations: annotations,
      item_photos: itemPhotos, item_files: itemFiles, item_thumbs: itemThumbs
    };
  }

  // ---------- 電腦 → 手機:清單快照 ----------
  /** tracker = trackerForSnapshot(...)(不良品倉項目明細);沒有就省略 */
  function makeSnapshot(state, thumbs, tracker) {
    var s = normalizeState(state);
    var out = {
      type: 'dwt-quality-snapshot', version: 1, generated_at: nowIso(),
      issues: s.issues, logs: s.logs, thumbs: thumbs || {}
    };
    if (tracker) out.tracker = tracker;
    return out;
  }
  function readSnapshot(obj) {
    if (!obj || obj.type !== 'dwt-quality-snapshot' || obj.version !== 1) return null;
    var s = normalizeState({ issues: obj.issues, logs: obj.logs });
    var thumbs = {};
    Object.keys(obj.thumbs || {}).forEach(function (n) { if (PHOTO_RE.test(n) && typeof obj.thumbs[n] === 'string') thumbs[n] = obj.thumbs[n]; });
    return { state: s, thumbs: thumbs, tracker: cleanTracker(obj.tracker), generated_at: str(obj.generated_at, 40) };
  }
  /**
   * 手機載入新的電腦清單後:電腦已經收到的變更就從「待傳送」拿掉。
   * pending = { issues{id: issue}, logs{id: log}, annotations{項目id: 註記} };回傳仍未被電腦確認的部分。
   * tracker = 新清單裡的項目明細(沒有就保留所有待傳送的註記)。
   */
  function rebasePending(base, pending, tracker) {
    var out = { issues: {}, logs: {}, annotations: {}, item_photos: {} };
    var anns = {}, photoNames = {};
    ((tracker && tracker.items) || []).forEach(function (r) {
      anns[r.id] = r.ann_updated_at || '';
      (r.photos || []).forEach(function (n) { photoNames[n] = true; });
    });
    Object.keys(pending.item_photos || {}).forEach(function (id) {
      var p = pending.item_photos[id];
      var confirmed = tracker && (p.deleted ? !photoNames[p.name] : !!photoNames[p.name]);
      if (!confirmed) out.item_photos[id] = p;
    });
    Object.keys(pending.annotations || {}).forEach(function (id) {
      var p = pending.annotations[id];
      if (!(anns[id] != null && anns[id] >= (p.updated_at || ''))) out.annotations[id] = p;
    });
    Object.keys(pending.issues || {}).forEach(function (id) {
      var p = pending.issues[id], b = base.issues[id];
      var confirmed = !!b && Object.keys(p.field_times || {}).every(function (f) { return fieldTime(b, f) >= p.field_times[f]; });
      if (!confirmed) out.issues[id] = p;
    });
    Object.keys(pending.logs || {}).forEach(function (id) {
      var p = pending.logs[id], b = base.logs[id];
      if (!(b && (b.updated_at || '') >= (p.updated_at || ''))) out.logs[id] = p;
    });
    return out;
  }

  return {
    SCHEMA_VERSION: SCHEMA_VERSION, STATUSES: STATUSES, SCAR_STATES: SCAR_STATES, LOG_TYPES: LOG_TYPES,
    ORIGINS: ORIGINS, ORIGIN_OTHER: ORIGIN_OTHER, ORIGIN_OTHER_MAX: ORIGIN_OTHER_MAX, ORIGIN_NONE: ORIGIN_NONE,
    ISSUE_FIELDS: ISSUE_FIELDS, PHOTO_RE: PHOTO_RE,
    setIssueFields: setIssueFields, originLabel: originLabel, validateOrigin: validateOrigin,
    originCounts: originCounts, matchOrigin: matchOrigin,
    nowIso: nowIso, localDate: localDate, shortDate: shortDate, newId: newId,
    emptyState: emptyState, normalizeState: normalizeState, cleanIssue: cleanIssue, cleanLog: cleanLog,
    newIssue: newIssue, setIssueField: setIssueField, newLog: newLog,
    mergeIssue: mergeIssue, mergeStates: mergeStates,
    overdueDays: overdueDays, badge: badge, qtyText: qtyText, title: title,
    liveIssues: liveIssues, counts: counts, filterIssues: filterIssues, issueLogs: issueLogs,
    cleanAnnotations: cleanAnnotations, cleanItemPhoto: cleanItemPhoto, newItemPhoto: newItemPhoto, trackerForSnapshot: trackerForSnapshot, cleanTracker: cleanTracker, trackerItems: trackerItems,
    makePackage: makePackage, applyPackage: applyPackage,
    makeSnapshot: makeSnapshot, readSnapshot: readSnapshot, rebasePending: rebasePending
  };
});
