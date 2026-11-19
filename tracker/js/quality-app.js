/*
 * quality-app.js — 桌面版「品質異常」分頁 + 手機同步 + 料號照片視窗
 *
 * 連結資料夾(放在公司 OneDrive,手機才能同步)底下:
 *   手機同步/收件匣/*.json        手機傳來的傳送包 → 處理後搬到 收件匣/已處理/(舊版的 品質異常/收件匣 也會讀)
 *   手機同步/手機清單.json        給手機載入的最新清單:不良品倉項目明細 + 品質異常 + 縮圖
 *   不良品倉照片/*.jpg            料號照片(I 開頭);縮圖/ 底下是縮圖。紀錄存在看板 tracker-data.json 的 item_photos
 *   品質異常/quality-data.json   品質異常紀錄(與不良品倉是不同的事,資料分開)
 *   品質異常/照片/*.jpg           現場紀錄照片(L 開頭);縮圖/ 底下是縮圖
 *
 * 同步時機:看板每次讀取資料夾後(啟動、回到視窗、每 60 秒)、以及這裡每次修改後。
 * 外來資料(手機傳送包)一律用 textContent 顯示,不用 innerHTML。
 */
(function () {
  'use strict';

  var Q = window.DWTQualityCore, P = window.DWTQualityPhoto, Store = window.DWTStore;
  var DIR = '品質異常';
  var DATA = DIR + '/quality-data.json';
  var PHOTOS = DIR + '/照片';
  var THUMBS = DIR + '/照片/縮圖';
  // 手機同步(不良品倉項目明細 + 品質異常共用)
  var SYNC = '手機同步';
  var SNAP = SYNC + '/手機清單.json';
  var INBOX = SYNC + '/收件匣';
  var OLD_INBOX = DIR + '/收件匣';          // v1.2 的位置,仍會讀取
  var OLD_SNAP = DIR + '/手機清單.json';
  // 不良品倉料號照片(紀錄存在看板的 tracker-data.json,檔案放這裡)
  var ITEM_PHOTOS = '不良品倉照片';
  var ITEM_THUMBS = '不良品倉照片/縮圖';
  var LS_KEY = 'dwt-quality-v1';
  var SETTLE_MS = 30000;          // OneDrive 剛同步進來的檔案先等 30 秒再讀,避免讀到一半
  var BROKEN_AFTER_MS = 600000;   // 10 分鐘後還是讀不了才當成壞檔
  // 手機版網址(claude.ai 私人頁面)
  var MOBILE_URL = 'https://yipeiliu26-hash.github.io/dwt-mobile/';

  var $ = function (id) { return document.getElementById(id); };
  var state = loadLocal();
  // origin:'' = 全部來源,Q.ORIGIN_NONE = 未指定;creating:正在填「新增異常」表單
  var view = { filter: 'open', q: '', origin: '', selected: null, creating: false };
  var thumbCache = {};
  var sync = { running: false, again: false, lastAt: null, packages: 0, snapAt: null, error: '' };
  var writeChain = Promise.resolve();
  var pendingDetail = false;

  // ---------- 小工具 ----------
  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'text') e.textContent = v;
      else if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : v);
    });
    (children || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return e;
  }
  function today() { return Q.localDate(); }
  function hhmm(d) { return d ? String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') : ''; }
  function localStamp(iso) { if (!iso) return ''; var d = new Date(iso); return isNaN(d) ? '' : Q.localDate(d) + ' ' + hhmm(d); }
  function app() { return window.DWTApp; }
  function toast(msg, kind, sticky) { if (app() && app().toast) app().toast(msg, kind, sticky); }
  function folder() { var f = app() && app().folder; return f && f.handle && f.granted ? f.handle : null; }
  function isVisible() { return !$('panel-quality').hidden; }
  function editingDetail() { var a = document.activeElement; return !!(a && $('q-detail').contains(a) && /INPUT|TEXTAREA|SELECT/.test(a.tagName)); }

  // ---------- 儲存 ----------
  function loadLocal() {
    try { var raw = localStorage.getItem(LS_KEY); return raw ? Q.normalizeState(JSON.parse(raw)) : Q.emptyState(); }
    catch (e) { return Q.emptyState(); }
  }
  function saveLocal() { try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch (e) { /* 空間不足時只靠資料夾 */ } }

  /** 讀資料夾裡的 quality-data.json 合併 → 寫回(有變才寫)→ 更新手機清單 */
  function persist() {
    saveLocal();
    var h = folder();
    if (!h) return Promise.resolve();
    writeChain = writeChain.then(function () { return writeFolder(h); }).catch(function (e) {
      toast('品質異常資料存檔失敗:' + (e && e.message ? e.message : e), 'err', true);
    });
    return writeChain;
  }
  async function writeFolder(h) {
    var txt = await Store.readTextAt(h, DATA);
    if (txt) { try { state = Q.mergeStates(state, JSON.parse(txt)); } catch (e) { /* 壞檔 → 用本機的覆蓋 */ } }
    var out = JSON.stringify(state, null, 1);
    if (out !== txt) await Store.writeAt(h, DATA, out);
    saveLocal();
    await writeSnapshot(h);
  }
  // 照片檔名開頭:L = 品質異常現場紀錄,I = 不良品倉料號照片
  function isItemPhoto(name) { return name.charAt(0) === 'I'; }
  function thumbPath(name) { return (isItemPhoto(name) ? ITEM_THUMBS : THUMBS) + '/' + name; }
  function fullPath(name) { return (isItemPhoto(name) ? ITEM_PHOTOS : PHOTOS) + '/' + name; }
  async function loadThumb(h, n) {
    if (thumbCache[n] == null) {
      var b = await Store.readBytesAt(h, thumbPath(n));
      thumbCache[n] = b ? P.bytesToBase64(b) : '';
    }
    return thumbCache[n];
  }

  async function writeSnapshot(h) {
    var names = [];
    Q.liveIssues(state).forEach(function (it) {
      Q.issueLogs(state, it.id).forEach(function (l) { names = names.concat(l.photos); });
    });
    var tracker = null;
    if (app() && app().trackerSnapshot) {
      var t = app().trackerSnapshot();
      tracker = Q.trackerForSnapshot(t.rows, t.annotations, t);
      tracker.items.forEach(function (r) { names = names.concat(r.photos); });
    }
    var thumbs = {};
    for (var i = 0; i < names.length; i++) {
      var tb = await loadThumb(h, names[i]);
      if (tb) thumbs[names[i]] = tb;
    }
    var snap = Q.makeSnapshot(state, thumbs, tracker);
    var old = await Store.readTextAt(h, SNAP);
    var same = false;
    if (old) {
      try {
        var o = JSON.parse(old);
        same = JSON.stringify([o.issues, o.logs, o.thumbs, o.tracker]) === JSON.stringify([snap.issues, snap.logs, snap.thumbs, snap.tracker]);
      } catch (e) { same = false; }
    }
    if (!same) await Store.writeAt(h, SNAP, JSON.stringify(snap));
    sync.snapAt = new Date();
    // 先建好收件匣,手機第一次存檔時才找得到這個資料夾
    if (!sync.inboxReady) {
      var note = INBOX + '/請把手機傳來的檔案存在這裡.txt';
      if ((await Store.readTextAt(h, note)) == null) {
        await Store.writeAt(h, note, '手機版按「傳送到電腦」後,把產生的 手機傳送_日期_時間.json 存到這個資料夾。\r\n電腦上的看板開著時會自動讀取,讀完的檔案會移到「已處理」。\r\n');
      }
      // v1.2 放在「品質異常」裡的手機清單已經不更新,刪掉避免選錯(那是看板自己產生的檔案)
      if ((await Store.readTextAt(h, OLD_SNAP)) != null) { try { await Store.removeAt(h, OLD_SNAP); } catch (e) { /* 留著也無妨 */ } }
      sync.inboxReady = true;
    }
  }

  // ---------- 收件匣(手機傳送包) ----------
  async function writePhoto(h, name, fullB64, thumbB64) {
    await Store.writeAt(h, fullPath(name), P.base64ToBytes(fullB64));
    if (thumbB64) { await Store.writeAt(h, thumbPath(name), P.base64ToBytes(thumbB64)); thumbCache[name] = thumbB64; }
  }
  async function processInbox(h) {
    var got = { packages: 0, issues: 0, logs: 0, photos: 0, broken: [], annotations: {}, item_photos: {} };
    var dirs = [INBOX, OLD_INBOX];
    var now = Date.now();
    for (var d = 0; d < dirs.length; d++) {
      var dir = dirs[d];
      var entries = await Store.listDirAt(h, dir);
      for (var i = 0; i < entries.length; i++) {
        var f = entries[i];
        if (f.kind !== 'file' || !/\.json$/i.test(f.name) || /\.tmp$/i.test(f.name)) continue;
        var age = now - (f.lastModified || 0);
        if (age < SETTLE_MS) { sync.again = true; continue; }
        var path = dir + '/' + f.name;
        var pkg = null;
        try { pkg = JSON.parse(await Store.readTextAt(h, path)); } catch (e) { pkg = null; }
        if (!pkg) {
          if (age > BROKEN_AFTER_MS) { await Store.moveAt(h, path, INBOX + '/無法讀取/' + f.name); got.broken.push(f.name); }
          continue;
        }
        var r = Q.applyPackage(state, pkg);
        if (!r.ok) { await Store.moveAt(h, path, INBOX + '/無法讀取/' + f.name); got.broken.push(f.name); continue; }
        Object.keys(r.annotations || {}).forEach(function (id) {
          var a = r.annotations[id], cur = got.annotations[id];
          if (!cur || a.updated_at > cur.updated_at) got.annotations[id] = a;
        });
        (r.item_photos || []).forEach(function (p) {
          var cur = got.item_photos[p.id];
          if (!cur || p.updated_at > cur.updated_at) got.item_photos[p.id] = p;
        });
        var names = Object.keys(r.photos);
        for (var j = 0; j < names.length; j++) await writePhoto(h, names[j], r.photos[names[j]], r.thumbs[names[j]]);
        var inames = Object.keys(r.item_files || {});
        for (var k = 0; k < inames.length; k++) await writePhoto(h, inames[k], r.item_files[inames[k]], r.item_thumbs[inames[k]]);
        await Store.moveAt(h, path, INBOX + '/已處理/' + f.name);
        if (!r.duplicate) { got.packages++; got.issues += r.issues; got.logs += r.logs; got.photos += names.length + inames.length; }
      }
    }
    return got;
  }

  /** 看板讀取資料夾後呼叫;也可手動按「立即同步」 */
  async function syncNow(trigger) {
    var h = folder();
    if (!h) { render(); return; }
    if (sync.running) return;
    sync.running = true; sync.again = false; sync.error = '';
    renderSync();
    try {
      var txt = await Store.readTextAt(h, DATA);
      if (txt) state = Q.mergeStates(state, JSON.parse(txt));
      var got = await processInbox(h);
      sync.packages += got.packages;
      // 手機改的項目明細註記 → 寫回看板(先寫回,手機清單才會帶到最新的註記)
      var annCount = Object.keys(got.annotations).length && app() && app().applyAnnotations ? await app().applyAnnotations(got.annotations) : 0;
      var ipList = Object.keys(got.item_photos).map(function (k) { return got.item_photos[k]; });
      var ipCount = ipList.length && app() && app().applyItemPhotos ? await app().applyItemPhotos(ipList) : 0;
      await writeFolder(h);
      sync.lastAt = new Date();
      if (got.packages) {
        var parts = [];
        if (annCount) parts.push(annCount + ' 筆項目明細');
        if (ipCount) parts.push(ipCount + ' 筆料號照片變更');
        if (got.issues) parts.push(got.issues + ' 筆異常');
        if (got.logs) parts.push(got.logs + ' 筆現場紀錄');
        if (got.photos) parts.push(got.photos + ' 張照片');
        toast('已收到手機傳來的資料' + (parts.length ? ':' + parts.join('、') : '(沒有新的變更)'), 'ok');
      }
      if (got.broken.length) toast('收件匣有 ' + got.broken.length + ' 個檔案無法讀取,已移到「手機同步/收件匣/無法讀取」:' + got.broken.join('、'), 'warn', true);
      else if (trigger === 'manual' && !got.packages) toast('收件匣沒有新的手機資料' + (sync.again ? '(有檔案剛同步進來,30 秒後會自動再讀)' : ''), '');
    } catch (e) {
      sync.error = e && e.message ? e.message : String(e);
    } finally {
      sync.running = false;
      render();
      if (sync.again) setTimeout(function () { syncNow('retry'); }, SETTLE_MS + 2000);
    }
  }

  // ---------- 修改 ----------
  function updateIssue(id, field, value) {
    var it = state.issues[id];
    if (!it) return;
    state.issues[id] = Q.setIssueField(it, field, value);
    state.updated_at = Q.nowIso();
    persist();
    renderList(); renderDetailHead();
  }
  /** 打開「新增異常」表單(異常來源必填,建立後才會出現在清單) */
  function addIssue() {
    view.creating = true;
    view.selected = null;
    renderList();
    renderDetail();
    var first = $('q-n-supplier'); if (first) first.focus();
  }
  function createIssue(fields) {
    var it = Q.newIssue(Object.assign({ status: '調查中', scar: '未開' }, fields));
    state.issues[it.id] = it;
    state.updated_at = Q.nowIso();
    view.creating = false;
    view.selected = it.id;
    if (view.filter === 'overdue') view.filter = 'open';
    if (view.origin && !Q.matchOrigin(it, view.origin)) { view.origin = ''; $('q-origin-filter').value = ''; }
    view.q = ''; $('q-search').value = '';
    persist();
    render();
    toast('已新增異常「' + Q.title(it) + '」', 'ok');
  }
  /** 改來源:一定和 origin_other 一起存(同一個修改時間) */
  function updateOrigin(id, origin, other) {
    var it = state.issues[id];
    if (!it) return;
    state.issues[id] = Q.setIssueFields(it, { origin: origin, origin_other: origin === Q.ORIGIN_OTHER ? other : '' });
    state.updated_at = Q.nowIso();
    persist();
    renderList(); renderDetailHead();
  }
  function deleteIssue(id) {
    var it = state.issues[id];
    if (!it || !confirm('確定刪除「' + Q.title(it) + '」?(手機下次載入清單後也會消失)')) return;
    updateIssue(id, 'deleted', true);
    view.selected = null;
    render();
  }
  async function addLog(issueId, form) {
    var files = Array.prototype.slice.call(form.photos.files || []);
    var h = folder();
    if (files.length && !h) { toast('要先連結資料夾才能加照片(照片存在資料夾的「品質異常/照片」)', 'warn', true); return; }
    var note = form.note.value.trim();
    if (!note && !files.length) { toast('請輸入註解或選擇照片', 'warn'); return; }
    var btn = form.querySelector('button[type="submit"]');
    btn.disabled = true; btn.textContent = '儲存中…';
    try {
      var log = Q.newLog({ issue_id: issueId, date: form.date.value || today(), type: form.type.value, note: note, status_to: form.status.value, source: 'desktop' });
      var names = [];
      for (var i = 0; i < files.length; i++) {
        var pic = await P.prepare(files[i]);
        var name = log.id + '-' + (i + 1) + '.jpg';
        await Store.writeAt(h, PHOTOS + '/' + name, new Uint8Array(await pic.full.arrayBuffer()));
        var tb = new Uint8Array(await pic.thumb.arrayBuffer());
        await Store.writeAt(h, THUMBS + '/' + name, tb);
        thumbCache[name] = P.bytesToBase64(tb);
        names.push(name);
      }
      log.photos = names;
      state.logs[log.id] = log;
      if (log.status_to && state.issues[issueId].status !== log.status_to) {
        state.issues[issueId] = Q.setIssueField(state.issues[issueId], 'status', log.status_to);
      }
      state.updated_at = Q.nowIso();
      await persist();
      toast('已新增現場紀錄' + (names.length ? '(' + names.length + ' 張照片)' : ''), 'ok');
      render();
    } catch (e) {
      toast('新增紀錄失敗:' + (e && e.message ? e.message : e), 'err', true);
      btn.disabled = false; btn.textContent = '新增紀錄';
    }
  }
  function deleteLog(id) {
    var lg = state.logs[id];
    if (!lg || !confirm('確定刪除這筆 ' + Q.shortDate(lg.date) + ' 的現場紀錄?')) return;
    var copy = JSON.parse(JSON.stringify(lg));
    copy.deleted = true; copy.updated_at = Q.nowIso();
    state.logs[id] = copy;
    state.updated_at = Q.nowIso();
    persist();
    render();
  }

  // ---------- 照片 ----------
  function thumbImg(name) {
    var img = el('img', { alt: '現場照片', loading: 'lazy' });
    function set(b64) { if (b64) img.src = P.dataUrl(b64); else img.classList.add('missing'); }
    if (thumbCache[name] != null) set(thumbCache[name]);
    else if (folder()) loadThumb(folder(), name).then(set, function () { set(''); });
    else set('');
    return el('button', { type: 'button', class: 'q-thumb', title: '看大圖', onclick: function () { openPhoto(name); } }, [img]);
  }
  async function openPhoto(name) {
    var h = folder();
    if (!h) return;
    var bytes = await Store.readBytesAt(h, fullPath(name));
    if (!bytes) { toast('找不到原圖:' + name, 'warn'); return; }
    var url = URL.createObjectURL(new Blob([bytes], { type: 'image/jpeg' }));
    var box = el('div', { class: 'q-lightbox', role: 'dialog', 'aria-label': '照片', onclick: close }, [
      el('img', { src: url, alt: '現場照片' }),
      el('button', { type: 'button', class: 'btn', text: '關閉', onclick: close })
    ]);
    function close() { URL.revokeObjectURL(url); box.remove(); document.removeEventListener('keydown', onKey); }
    function onKey(e) { if (e.key === 'Escape') close(); }
    document.addEventListener('keydown', onKey);
    document.body.appendChild(box);
  }

  // ---------- 不良品倉料號照片(項目明細的「照片」按鈕) ----------
  function openItemPhotos(itemId, label) {
    var Core = window.DWTCore;
    var box = el('div', { class: 'q-lightbox q-gallery', role: 'dialog', 'aria-label': label + ' 的照片' });
    function close() { box.remove(); document.removeEventListener('keydown', onKey); }
    function onKey(e) { if (e.key === 'Escape' && !document.querySelector('.q-lightbox:not(.q-gallery)')) close(); }
    function draw() {
      box.textContent = '';
      var list = Core.itemPhotos(app().state(), itemId);
      var input = el('input', { type: 'file', accept: 'image/*', multiple: true, id: 'q-item-photo-input' });
      input.addEventListener('change', function () { addItemPhotos(itemId, input.files).then(draw); });
      box.appendChild(el('div', { class: 'q-gallery-panel' }, [
        el('div', { class: 'q-gallery-head' }, [
          el('h3', { text: label }),
          el('button', { type: 'button', class: 'btn', text: '關閉', onclick: close })
        ]),
        list.length ? el('div', { class: 'q-thumbs' }, list.map(function (p) {
          return el('div', { class: 'q-gallery-item' }, [
            thumbImg(p.name),
            el('span', { class: 'footnote', text: localStamp(p.taken_at) + (p.source === 'phone' ? ' · 手機' : '') }),
            el('button', { type: 'button', class: 'q-link', text: '刪除', onclick: function () {
              if (!confirm('確定刪除這張照片?')) return;
              var rec = Object.assign({}, p, { deleted: true, updated_at: Q.nowIso() });
              app().applyItemPhotos([rec]).then(function () { snapshotSoon(); draw(); });
            } })
          ]);
        })) : el('p', { class: 'empty-state', text: '還沒有照片。可以在手機版的項目明細拍照,或在這裡加入。' }),
        folder() ? el('label', { class: 'btn primary', for: 'q-item-photo-input', text: '加照片' }) : el('p', { class: 'footnote', text: '要先連結資料夾才能加照片。' }),
        input
      ]));
      input.hidden = true;
    }
    box.addEventListener('click', function (e) { if (e.target === box) close(); });
    document.addEventListener('keydown', onKey);
    document.body.appendChild(box);
    draw();
  }
  async function addItemPhotos(itemId, files) {
    var h = folder();
    files = Array.prototype.slice.call(files || []);
    if (!h || !files.length) return;
    var recs = [];
    try {
      for (var i = 0; i < files.length; i++) {
        var pic = await P.prepare(files[i]);
        var rec = Q.newItemPhoto(itemId, 'desktop');
        await Store.writeAt(h, fullPath(rec.name), new Uint8Array(await pic.full.arrayBuffer()));
        var tb = new Uint8Array(await pic.thumb.arrayBuffer());
        await Store.writeAt(h, thumbPath(rec.name), tb);
        thumbCache[rec.name] = P.bytesToBase64(tb);
        recs.push(rec);
      }
    } catch (e) {
      toast('加照片失敗:' + (e && e.message ? e.message : e), 'err', true);
    }
    if (recs.length) { await app().applyItemPhotos(recs); snapshotSoon(); toast('已加入 ' + recs.length + ' 張照片', 'ok'); }
  }

  // ---------- 畫面 ----------
  function render() {
    renderSync();
    renderList();
    if (editingDetail()) { pendingDetail = true; return; }
    renderDetail();
  }

  function renderSync() {
    var box = $('q-sync');
    if (!box) return;
    box.textContent = '';
    var h = folder();
    var dot = el('span', { class: 'conn-dot' });
    var text = el('span', { class: 'conn-text' });
    if (!h) {
      dot.className = 'conn-dot warn';
      text.appendChild(el('strong', { text: '還沒和手機同步' }));
      text.appendChild(document.createTextNode(' — 品質異常紀錄目前只存在這台電腦。按上方「連結資料夾」選公司 OneDrive 裡的資料夾,才能和手機同步、加照片。'));
    } else {
      var notOneDrive = h.path && !/onedrive/i.test(h.path);
      dot.className = 'conn-dot ' + (sync.error || notOneDrive ? 'warn' : 'on');
      if (notOneDrive) {
        text.appendChild(el('strong', { text: '資料夾不在 OneDrive' }));
        text.appendChild(document.createTextNode(' — 「' + h.path + '」不會同步到手機。請把資料夾搬到「公司的 OneDrive」裡,再重新連結。'));
      } else if (sync.error) {
        text.appendChild(el('strong', { text: '同步失敗' }));
        text.appendChild(document.createTextNode(' — ' + sync.error));
      } else {
        text.appendChild(el('strong', { text: sync.running ? '同步中…' : '已和手機同步' }));
        text.appendChild(document.createTextNode(' — ' + (sync.lastAt ? '收件匣上次檢查 ' + hhmm(sync.lastAt) : '尚未檢查收件匣') +
          (sync.packages ? ',這次開啟後收到 ' + sync.packages + ' 包' : '') + '。手機要看最新進度,請載入「手機同步/手機清單.json」。'));
      }
    }
    var btns = el('span', { class: 'conn-buttons' }, [
      h ? el('button', { type: 'button', class: 'btn', text: '立即同步', disabled: sync.running, onclick: function () { syncNow('manual'); } }) : null,
      MOBILE_URL ? el('a', { class: 'btn', href: MOBILE_URL, target: '_blank', rel: 'noopener', text: '手機版網址' }) : null
    ]);
    box.appendChild(dot); box.appendChild(text); box.appendChild(btns);
  }

  function renderList() {
    var c = Q.counts(state, today());
    var labels = { open: '未結案 ' + c.open, overdue: '逾期 ' + c.overdue, all: '全部 ' + c.all };
    document.querySelectorAll('#q-filter button').forEach(function (b) {
      b.textContent = labels[b.dataset.f];
      b.setAttribute('aria-selected', String(b.dataset.f === view.filter));
    });
    // 來源篩選:選項後面附上筆數(跟著狀態篩選變)
    var oc = Q.originCounts(state, view.filter, today());
    var sel = $('q-origin-filter');
    sel.textContent = '';
    [['', '全部來源']].concat(Q.ORIGINS.map(function (o) { return [o, o]; }), [[Q.ORIGIN_NONE, '未指定']]).forEach(function (p) {
      sel.appendChild(el('option', { value: p[0], text: p[1] + ' ' + oc[p[0]] }));
    });
    sel.value = view.origin;
    var list = $('q-list');
    list.textContent = '';
    var rows = Q.filterIssues(state, view.filter, view.q, today(), view.origin);
    if (!rows.length) {
      var msg = c.all ? '沒有符合的異常' : '還沒有品質異常紀錄。按「新增異常」建立第一筆,或在手機版新增後傳送過來。';
      list.appendChild(el('p', { class: 'empty-state', text: msg }));
      return;
    }
    rows.forEach(function (it) {
      var b = Q.badge(it, today());
      var line2 = [it.defect, Q.qtyText(it)].filter(Boolean).join(' · ');
      var line3 = [it.scar && it.scar !== '未開' ? 'SCAR ' + it.scar : '', it.due ? '期限 ' + Q.shortDate(it.due) : ''].filter(Boolean).join(' · ');
      list.appendChild(el('button', {
        type: 'button', class: 'q-card tone-' + b.tone + (view.selected === it.id ? ' selected' : ''),
        'aria-pressed': String(view.selected === it.id),
        onclick: function () { view.selected = it.id; view.creating = false; renderList(); renderDetail(); }
      }, [
        el('span', { class: 'q-card-top' }, [el('span', { class: 'q-card-title', text: Q.title(it) }), el('span', { class: 'q-badge tone-' + b.tone, text: b.label })]),
        line2 ? el('span', { class: 'q-card-line', text: line2 }) : null,
        el('span', { class: 'q-card-meta' }, [
          el('span', { class: 'q-origin' + (it.origin ? '' : ' none'), text: '來源 ' + Q.originLabel(it) }),
          line3 ? el('span', { class: 'q-card-line muted', text: line3 }) : null
        ])
      ]));
    });
  }

  function renderDetailHead() {
    var it = state.issues[view.selected];
    var head = $('q-detail-head');
    if (!it || !head) return;
    var b = Q.badge(it, today());
    head.textContent = '';
    head.appendChild(el('h3', { text: Q.title(it) }));
    head.appendChild(el('span', { class: 'q-badge tone-' + b.tone, text: b.days ? '逾期 ' + b.days + ' 天' : b.label }));
  }

  function field(id, label, input) {
    return el('label', { class: 'q-field', for: id }, [el('span', { text: label }), input]);
  }
  function textInput(it, f, opts) {
    opts = opts || {};
    var inp = el(opts.multiline ? 'textarea' : 'input', {
      id: 'q-f-' + f, class: 'q-input', type: opts.multiline ? null : (opts.type || 'text'),
      placeholder: opts.placeholder || '', inputmode: opts.inputmode || null, min: opts.type === 'number' ? '0' : null, rows: opts.multiline ? '2' : null
    });
    inp.value = it[f] === '' || it[f] == null ? '' : String(it[f]);
    inp.addEventListener('change', function () { updateIssue(it.id, f, inp.value); });
    return inp;
  }
  function selectInput(it, f, options) {
    var s = el('select', { id: 'q-f-' + f, class: 'q-input' }, options.map(function (o) { return el('option', { value: o, text: o }); }));
    s.value = it[f];
    s.addEventListener('change', function () { updateIssue(it.id, f, s.value); });
    return s;
  }

  /**
   * 異常來源的下拉選單 + 「其他」自填格。
   * onCommit(origin, other) 在值確定時呼叫(編輯模式用來自動存檔);新增模式用 read() 取值。
   */
  function originControl(prefix, origin, other, opts) {
    opts = opts || {};
    var sel = el('select', { id: prefix + '-origin', class: 'q-input' },
      [el('option', { value: '', text: opts.isNew ? '請選擇…' : '未指定', disabled: !opts.isNew && !!origin })]
        .concat(Q.ORIGINS.map(function (o) { return el('option', { value: o, text: o }); })));
    sel.value = origin || '';
    var otherInput = el('input', {
      id: prefix + '-origin-other', class: 'q-input', type: 'text', maxlength: String(Q.ORIGIN_OTHER_MAX),
      placeholder: '請說明來源(最多 ' + Q.ORIGIN_OTHER_MAX + ' 字)', 'aria-label': '其他來源說明'
    });
    otherInput.value = origin === Q.ORIGIN_OTHER ? (other || '') : '';
    var err = el('span', { class: 'q-field-error', role: 'alert' });
    var otherWrap = el('div', { class: 'q-origin-other' }, [otherInput]);
    function sync() { otherWrap.hidden = sel.value !== Q.ORIGIN_OTHER; }
    function read() { return { origin: sel.value, other: sel.value === Q.ORIGIN_OTHER ? otherInput.value.trim() : '' }; }
    function check() {
      var v = read(), msg = Q.validateOrigin(v.origin, v.other, !!opts.isNew);
      err.textContent = msg;
      return msg ? null : v;
    }
    sel.addEventListener('change', function () {
      if (sel.value !== Q.ORIGIN_OTHER) otherInput.value = '';   // 切換成其他選項 → 清空自填
      sync();
      err.textContent = '';
      if (sel.value === Q.ORIGIN_OTHER) { otherInput.focus(); if (!otherInput.value.trim()) { if (!opts.isNew) err.textContent = '請填寫來源說明,填完才會儲存'; return; } }
      if (opts.onCommit) { var v = check(); if (v) opts.onCommit(v.origin, v.other); }
    });
    otherInput.addEventListener('change', function () {
      if (!opts.onCommit) return;
      var v = check();
      if (v) opts.onCommit(v.origin, v.other);
    });
    sync();
    var wrap = el('div', { class: 'q-field q-origin-field' }, [
      el('label', { for: prefix + '-origin', text: '異常來源' + (opts.isNew ? '(必填)' : '') }), sel, otherWrap, err
    ]);
    return { el: wrap, read: read, check: check };
  }

  /** 新增異常表單 */
  function renderNewForm(box) {
    box.appendChild(el('div', { class: 'q-detail-head' }, [el('h3', { text: '新增異常' })]));
    function inp(id, label, type, ph) {
      var i = el('input', { id: 'q-n-' + id, class: 'q-input', type: type || 'text', placeholder: ph || '', min: type === 'number' ? '0' : null });
      return { input: i, el: field('q-n-' + id, label, i) };
    }
    var f = {
      supplier: inp('supplier', '供應商', 'text', '例如:供應商名稱'),
      item: inp('item', '品名', 'text', '例如:蝶閥'),
      part_no: inp('part_no', '料號'),
      defect: inp('defect', '不良描述', 'text', '例如:把手鬆動'),
      defect_qty: inp('defect_qty', '不良數', 'number'),
      total_qty: inp('total_qty', '總數', 'number'),
      due: inp('due', '期限', 'date')
    };
    var oc = originControl('q-n', '', '', { isNew: true });
    var err = el('p', { class: 'q-field-error', role: 'alert' });
    var form = el('form', { class: 'q-form', id: 'q-newform' }, [
      f.supplier.el, f.item.el, f.part_no.el, f.defect.el, oc.el, f.defect_qty.el, f.total_qty.el, f.due.el
    ]);
    box.appendChild(form);
    box.appendChild(err);
    box.appendChild(el('div', { class: 'q-detail-foot' }, [
      el('button', { type: 'button', class: 'btn', text: '取消', onclick: function () { view.creating = false; renderDetail(); } }),
      el('button', { type: 'button', class: 'btn primary', text: '建立異常', onclick: submit })
    ]));
    form.addEventListener('submit', function (e) { e.preventDefault(); submit(); });
    function submit() {
      var o = oc.check();
      if (!f.supplier.input.value.trim() && !f.defect.input.value.trim()) { err.textContent = '至少填供應商或不良描述'; f.supplier.input.focus(); return; }
      if (!o) { err.textContent = ''; $('q-n-origin').focus(); return; }
      var fields = { origin: o.origin, origin_other: o.other };
      Object.keys(f).forEach(function (k) { fields[k] = f[k].input.value.trim(); });
      createIssue(fields);
    }
  }

  function renderDetail() {
    pendingDetail = false;
    var box = $('q-detail');
    box.textContent = '';
    if (view.creating) { renderNewForm(box); return; }
    var it = state.issues[view.selected];
    if (!it || it.deleted) {
      view.selected = null;
      box.appendChild(el('p', { class: 'empty-state', text: '從左邊選一筆異常,查看狀態與現場紀錄。' }));
      return;
    }
    box.appendChild(el('div', { class: 'q-detail-head', id: 'q-detail-head' }));
    renderDetailHead();

    box.appendChild(el('div', { class: 'q-form' }, [
      field('q-f-supplier', '供應商', textInput(it, 'supplier', { placeholder: '例如:供應商名稱' })),
      field('q-f-item', '品名', textInput(it, 'item', { placeholder: '例如:蝶閥' })),
      field('q-f-part_no', '料號', textInput(it, 'part_no')),
      field('q-f-defect', '不良描述', textInput(it, 'defect', { placeholder: '例如:把手鬆動' })),
      originControl('q-f', it.origin, it.origin_other, { onCommit: function (o, other) { updateOrigin(it.id, o, other); } }).el,
      field('q-f-defect_qty', '不良數', textInput(it, 'defect_qty', { type: 'number', inputmode: 'decimal' })),
      field('q-f-total_qty', '總數', textInput(it, 'total_qty', { type: 'number', inputmode: 'decimal' })),
      field('q-f-scar', 'SCAR', selectInput(it, 'scar', Q.SCAR_STATES)),
      field('q-f-due', '期限', textInput(it, 'due', { type: 'date' })),
      field('q-f-status', '處理狀態', selectInput(it, 'status', Q.STATUSES)),
      el('div', { class: 'q-field wide' }, [el('label', { for: 'q-f-containment', text: '暫時對策' }), textInput(it, 'containment', { multiline: true, placeholder: '例如:100% 拉力 >10kg' })])
    ]));

    // 現場紀錄
    var logs = Q.issueLogs(state, it.id);
    var timeline = el('ol', { class: 'q-timeline' });
    if (!logs.length) timeline.appendChild(el('li', { class: 'q-empty', text: '還沒有現場紀錄' }));
    logs.forEach(function (l) {
      timeline.appendChild(el('li', {}, [
        el('div', { class: 'q-log-meta' }, [
          el('span', { text: Q.shortDate(l.date) + ' ' + l.type }),
          l.source === 'phone' ? el('span', { class: 'q-src', text: '手機' }) : null,
          l.status_to ? el('span', { class: 'q-src', text: '狀態 → ' + l.status_to }) : null,
          el('button', { type: 'button', class: 'q-link', text: '刪除', onclick: function () { deleteLog(l.id); } })
        ]),
        l.note ? el('p', { class: 'q-log-note', text: l.note }) : null,
        l.photos.length ? el('div', { class: 'q-thumbs' }, l.photos.map(thumbImg)) : null
      ]));
    });

    var form = el('form', { class: 'q-logform', id: 'q-logform' }, [
      el('div', { class: 'q-logform-row' }, [
        field('q-l-date', '日期', el('input', { id: 'q-l-date', name: 'date', type: 'date', class: 'q-input', value: today() })),
        field('q-l-type', '類型', el('select', { id: 'q-l-type', name: 'type', class: 'q-input' }, Q.LOG_TYPES.map(function (t) { return el('option', { value: t, text: t }); }))),
        field('q-l-status', '更新狀態(選填)', el('select', { id: 'q-l-status', name: 'status', class: 'q-input' },
          [el('option', { value: '', text: '不變更' })].concat(Q.STATUSES.map(function (s) { return el('option', { value: s, text: s }); }))))
      ]),
      field('q-l-note', '註解', el('textarea', { id: 'q-l-note', name: 'note', class: 'q-input', rows: '2', placeholder: '例如:線邊倉剩 34 pcs,已貼隔離標籤' })),
      el('div', { class: 'q-logform-row end' }, [
        field('q-l-photos', '照片', el('input', { id: 'q-l-photos', name: 'photos', type: 'file', accept: 'image/*', multiple: true })),
        el('button', { type: 'submit', class: 'btn primary', text: '新增紀錄' })
      ])
    ]);
    form.addEventListener('submit', function (e) { e.preventDefault(); addLog(it.id, form); });

    box.appendChild(el('div', { class: 'q-section' }, [el('h4', { text: '現場紀錄' }), timeline]));
    box.appendChild(el('details', { class: 'q-section q-add', open: true }, [el('summary', { text: '新增現場紀錄' }), form]));
    box.appendChild(el('div', { class: 'q-detail-foot' }, [
      el('span', { class: 'footnote', text: '建立 ' + localStamp(it.created_at).slice(0, 10) + ' · 最後修改 ' + localStamp(it.updated_at) }),
      el('button', { type: 'button', class: 'btn danger small', text: '刪除這筆異常', onclick: function () { deleteIssue(it.id); } })
    ]));
  }

  // ---------- 事件 ----------
  document.querySelectorAll('#q-filter button').forEach(function (b) {
    b.addEventListener('click', function () { view.filter = b.dataset.f; renderList(); });
  });
  $('q-search').addEventListener('input', function (e) { view.q = e.target.value; renderList(); });
  $('q-origin-filter').addEventListener('change', function (e) { view.origin = e.target.value; renderList(); });
  $('q-add').addEventListener('click', addIssue);
  document.addEventListener('focusout', function () {
    setTimeout(function () { if (pendingDetail && !editingDetail()) renderDetail(); }, 0);
  });
  // 跨日時重算逾期
  setInterval(function () { if (isVisible() && !editingDetail()) render(); }, 600000);
  window.addEventListener('storage', function (e) {
    if (e.key !== LS_KEY || !e.newValue) return;
    try { state = Q.mergeStates(state, JSON.parse(e.newValue)); render(); } catch (err) { /* ignore */ }
  });

  // 看板的項目明細註記改了 → 2 秒後更新手機清單(連續修改只寫一次)
  var snapTimer = null;
  function snapshotSoon() {
    clearTimeout(snapTimer);
    snapTimer = setTimeout(function () {
      var h = folder();
      if (!h) return;
      writeChain = writeChain.then(function () { return writeSnapshot(h); }).catch(function () { /* 下次同步再寫 */ });
    }, 2000);
  }

  window.DWTQuality = {
    render: render,
    sync: syncNow,
    snapshotSoon: snapshotSoon,
    openItemPhotos: openItemPhotos,
    state: function () { return state; }
  };

  render();
  if (folder()) syncNow('startup');
})();
