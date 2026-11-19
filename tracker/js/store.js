/*
 * store.js — 本機儲存
 *
 * 兩種模式:
 *   A. 資料夾模式(Chrome / Edge):用 File System Access API 連結一個本機資料夾。
 *      - 讀:資料夾(含子資料夾,最多 3 層)裡所有 .xlsx → 交給 app.js 解析
 *      - 寫:資料夾根目錄的 tracker-data.json(週快照 + 註記 + 來源檔紀錄)
 *      - 資料夾的「授權」存在 IndexedDB,下次開啟只要按一下「重新連結」即可
 *   B. 瀏覽器模式(其他瀏覽器、或還沒連結資料夾):資料存在 localStorage,
 *      用「匯入WMS檔案」手動選檔;可下載/載入 tracker-data.json 備份
 *
 * 不論哪種模式,每次儲存都會同時寫一份到 localStorage 當作備援。
 */
(function (root) {
  'use strict';

  var LS_KEY = 'dwt-tracker-data-v1';
  var DATA_FILE = 'tracker-data.json';
  var BACKUP_FILE = 'tracker-data.backup.json';
  var IDB_NAME = 'dwt-folder-handles';
  var IDB_STORE = 'handles';
  var HANDLE_KEY = 'data-folder';
  var MAX_DEPTH = 3;
  var SKIP_DIRS = { 'node_modules': 1, '.git': 1, 'vendor': 1 };

  // C. 桌面版(desktop/,Electron):preload 提供 window.dwtDesktop,資料夾路徑由桌面程式記住,
  //    不需要每次重新授權。handle 是 { desktop: true, name, path },介面與資料夾模式相同,app.js 不用分辨。
  var desktop = typeof window !== 'undefined' && window.dwtDesktop ? window.dwtDesktop : null;
  var supportsFolder = !!desktop || (typeof window !== 'undefined' && typeof window.showDirectoryPicker === 'function');

  function desktopHandle(info) { return info ? { desktop: true, name: info.name, path: info.path } : null; }
  function abortError() { var e = new Error('已取消'); e.name = 'AbortError'; return e; }

  // ---------- localStorage ----------
  function loadLocal() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }
  function saveLocal(state) {
    try { localStorage.setItem(LS_KEY, JSON.stringify(state)); return true; }
    catch (e) { return false; }
  }

  // ---------- IndexedDB(只用來記住資料夾授權) ----------
  function idb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = function () { req.result.createObjectStore(IDB_STORE); };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idbOp(mode, fn) {
    return idb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(IDB_STORE, mode);
        var store = tx.objectStore(IDB_STORE);
        var req = fn(store);
        tx.oncomplete = function () { resolve(req && req.result); db.close(); };
        tx.onerror = function () { reject(tx.error); db.close(); };
      });
    });
  }
  function getSavedHandle() {
    if (desktop) return desktop.getFolder().then(desktopHandle).catch(function () { return null; });
    if (!supportsFolder) return Promise.resolve(null);
    return idbOp('readonly', function (s) { return s.get(HANDLE_KEY); }).catch(function () { return null; });
  }
  function saveHandle(h) { return idbOp('readwrite', function (s) { return s.put(h, HANDLE_KEY); }).catch(function () {}); }
  function clearHandle() {
    if (desktop) return desktop.clearFolder().catch(function () {});
    return idbOp('readwrite', function (s) { return s.delete(HANDLE_KEY); }).catch(function () {});
  }

  // ---------- 資料夾 ----------
  function pickFolder() {
    if (desktop) {
      return desktop.pickFolder().then(function (info) {
        if (!info) throw abortError();
        return desktopHandle(info);
      });
    }
    return window.showDirectoryPicker({ id: 'dwt-data', mode: 'readwrite' }).then(function (h) {
      return saveHandle(h).then(function () { return h; });
    });
  }

  /** request=true 時會跳出瀏覽器授權視窗(必須在使用者點擊事件中呼叫) */
  function ensurePermission(handle, request) {
    // 桌面版沒有授權問題,只確認資料夾還在(例如隨身碟拔掉、資料夾被改名)
    if (handle && handle.desktop) {
      return desktop.getFolder().then(function (info) { return !!(info && info.exists && info.path === handle.path); })
        .catch(function () { return false; });
    }
    var opts = { mode: 'readwrite' };
    return handle.queryPermission(opts).then(function (p) {
      if (p === 'granted') return true;
      if (!request) return false;
      return handle.requestPermission(opts).then(function (p2) { return p2 === 'granted'; });
    }).catch(function () { return false; });
  }

  function readDataFile(handle) {
    if (handle.desktop) return desktop.readText(DATA_FILE).then(function (txt) { return txt == null ? null : JSON.parse(txt); });
    return handle.getFileHandle(DATA_FILE).then(function (fh) { return fh.getFile(); })
      .then(function (f) { return f.text(); })
      .then(function (txt) { return JSON.parse(txt); })
      .catch(function (e) {
        if (e && e.name === 'NotFoundError') return null;
        throw e;
      });
  }

  function writeText(handle, name, text) {
    if (handle.desktop) return desktop.writeText(name, text);
    return handle.getFileHandle(name, { create: true }).then(function (fh) {
      return fh.createWritable().then(function (w) {
        return w.write(text).then(function () { return w.close(); });
      });
    });
  }

  var backedUpThisSession = false;
  /** 寫 tracker-data.json;本次開啟後第一次寫入前,先把舊檔複製成 tracker-data.backup.json */
  function writeDataFile(handle, state) {
    var json = JSON.stringify(state, null, 1);
    var pre = Promise.resolve();
    if (!backedUpThisSession) {
      var readOld = handle.desktop
        ? desktop.readText(DATA_FILE).then(function (t) { if (t == null) throw new Error('no file'); return t; })
        : handle.getFileHandle(DATA_FILE).then(function (fh) { return fh.getFile(); }).then(function (f) { return f.text(); });
      pre = readOld
        .then(function (old) { return writeText(handle, BACKUP_FILE, old); })
        .catch(function () { /* 第一次使用時還沒有舊檔 */ })
        .then(function () { backedUpThisSession = true; });
    }
    return pre.then(function () { return writeText(handle, DATA_FILE, json); });
  }

  /** 列出資料夾裡所有 Excel 檔:[{path, name, handle}]。略過 Excel 暫存檔(~$ 開頭)與隱藏檔。 */
  async function listExcelFiles(dirHandle) {
    if (dirHandle.desktop) {
      // 桌面版:清單由桌面程式產生;handle.getFile() 回傳真正的 File,app.js 的用法不變
      var list = await desktop.listExcel();
      return list.map(function (f) {
        return {
          path: f.path, name: f.name,
          handle: {
            getFile: function () {
              return desktop.readFile(f.path).then(function (r) {
                return new File([r.data], f.name, { lastModified: r.lastModified });
              });
            }
          }
        };
      });
    }
    var out = [];
    async function walk(dir, prefix, depth) {
      for await (var entry of dir.values()) {
        var name = entry.name;
        if (name.charAt(0) === '.' || name.indexOf('~$') === 0) continue;
        if (entry.kind === 'directory') {
          if (depth < MAX_DEPTH && !SKIP_DIRS[name]) await walk(entry, prefix + name + '/', depth + 1);
        } else if (/\.(xlsx|xlsm)$/i.test(name)) {
          out.push({ path: prefix + name, name: name, handle: entry });
        }
      }
    }
    await walk(dirHandle, '', 0);
    out.sort(function (a, b) { return a.path < b.path ? -1 : 1; });
    return out;
  }

  // ---------- 子資料夾檔案操作(品質異常紀錄用;rel 以 / 分隔,例如「品質異常/照片/a.jpg」) ----------
  function splitRel(rel) { return String(rel).split('/').filter(Boolean); }
  async function dirAt(handle, parts, create) {
    var d = handle;
    for (var i = 0; i < parts.length; i++) d = await d.getDirectoryHandle(parts[i], { create: !!create });
    return d;
  }
  function isNotFound(e) { return e && (e.name === 'NotFoundError' || e.name === 'TypeMismatchError'); }

  /** 讀文字檔;不存在回 null */
  async function readTextAt(handle, rel) {
    if (handle.desktop) return desktop.readText(rel);
    var parts = splitRel(rel);
    try {
      var d = await dirAt(handle, parts.slice(0, -1), false);
      var f = await (await d.getFileHandle(parts[parts.length - 1])).getFile();
      return await f.text();
    } catch (e) { if (isNotFound(e)) return null; throw e; }
  }
  /** 讀二進位檔(Uint8Array);不存在回 null */
  async function readBytesAt(handle, rel) {
    if (handle.desktop) return desktop.readBytes(rel);
    var parts = splitRel(rel);
    try {
      var d = await dirAt(handle, parts.slice(0, -1), false);
      var f = await (await d.getFileHandle(parts[parts.length - 1])).getFile();
      return new Uint8Array(await f.arrayBuffer());
    } catch (e) { if (isNotFound(e)) return null; throw e; }
  }
  /** 寫檔(文字或 Uint8Array),上層資料夾不存在會自動建立 */
  async function writeAt(handle, rel, data) {
    if (handle.desktop) return typeof data === 'string' ? desktop.writeText(rel, data) : desktop.writeBytes(rel, data);
    var parts = splitRel(rel);
    var d = await dirAt(handle, parts.slice(0, -1), true);
    var w = await (await d.getFileHandle(parts[parts.length - 1], { create: true })).createWritable();
    await w.write(data);
    await w.close();
  }
  /** 列出子資料夾:[{name, kind, size, lastModified}];不存在回 [] */
  async function listDirAt(handle, rel) {
    if (handle.desktop) return desktop.listDir(rel);
    var out = [];
    try {
      var d = await dirAt(handle, splitRel(rel), false);
      for await (var entry of d.values()) {
        if (entry.name.charAt(0) === '.') continue;
        var item = { name: entry.name, kind: entry.kind, size: 0, lastModified: 0 };
        if (entry.kind === 'file') { var f = await entry.getFile(); item.size = f.size; item.lastModified = f.lastModified; }
        out.push(item);
      }
    } catch (e) { if (isNotFound(e)) return []; throw e; }
    return out;
  }
  /** 搬移檔案(瀏覽器模式是複製後刪除) */
  async function moveAt(handle, fromRel, toRel) {
    if (handle.desktop) return desktop.move(fromRel, toRel);
    var bytes = await readBytesAt(handle, fromRel);
    if (bytes == null) return;
    await writeAt(handle, toRel, bytes);
    var parts = splitRel(fromRel);
    var d = await dirAt(handle, parts.slice(0, -1), false);
    await d.removeEntry(parts[parts.length - 1]);
  }

  /** 刪除檔案(只用在看板自己產生的檔案) */
  async function removeAt(handle, rel) {
    if (handle.desktop) return desktop.remove(rel);
    var parts = splitRel(rel);
    try {
      var d = await dirAt(handle, parts.slice(0, -1), false);
      await d.removeEntry(parts[parts.length - 1]);
    } catch (e) { if (!isNotFound(e)) throw e; }
  }

  // ---------- 下載 / 另存 ----------
  /**
   * 存檔:Chrome/Edge 用「另存新檔」視窗(檔名一定正確,可選位置);其他瀏覽器用一般下載。
   * makeBlob 在選好位置之後才執行 —— showSaveFilePicker 必須在點擊後立刻呼叫,不能先 await 其他事。
   * @returns {Promise<boolean>} false = 使用者取消
   */
  async function saveBlob(filename, makeBlob, mime, ext, description) {
    if (desktop) {
      var target = await desktop.chooseSavePath(filename, ext, description || filename);
      if (!target) return false;
      var b = await makeBlob();
      await desktop.writePath(target, new Uint8Array(await b.arrayBuffer()));
      return true;
    }
    if (typeof window.showSaveFilePicker === 'function') {
      var handle = null;
      try {
        var accept = {}; accept[mime] = [ext];
        handle = await window.showSaveFilePicker({ suggestedName: filename, types: [{ description: description || filename, accept: accept }] });
      } catch (e) {
        if (e && e.name === 'AbortError') return false;
        handle = null; // 其他錯誤 → 改用一般下載
      }
      if (handle) {
        var blob = await makeBlob();
        var w = await handle.createWritable();
        await w.write(blob);
        await w.close();
        return true;
      }
    }
    downloadBlob(filename, await makeBlob());
    return true;
  }

  function downloadBlob(filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  root.DWTStore = {
    DATA_FILE: DATA_FILE,
    supportsFolder: supportsFolder,
    isDesktop: !!desktop,
    loadLocal: loadLocal,
    saveLocal: saveLocal,
    getSavedHandle: getSavedHandle,
    clearHandle: clearHandle,
    pickFolder: pickFolder,
    ensurePermission: ensurePermission,
    readDataFile: readDataFile,
    writeDataFile: writeDataFile,
    listExcelFiles: listExcelFiles,
    readTextAt: readTextAt,
    readBytesAt: readBytesAt,
    writeAt: writeAt,
    listDirAt: listDirAt,
    moveAt: moveAt,
    removeAt: removeAt,
    saveBlob: saveBlob,
    downloadBlob: downloadBlob
  };
})(window);
