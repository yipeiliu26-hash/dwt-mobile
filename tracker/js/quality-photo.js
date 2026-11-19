/*
 * quality-photo.js — 照片處理(瀏覽器端,桌面版與手機版共用)
 * 壓縮成 JPEG(原圖最長邊 1600px、縮圖 320px),以及 Base64 / Uint8Array 轉換。
 */
(function (root) {
  'use strict';

  var FULL_SIDE = 1600, FULL_QUALITY = 0.8;
  var THUMB_SIDE = 320, THUMB_QUALITY = 0.7;

  function loadImage(blob) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('無法讀取這張照片(格式可能不支援)')); };
      img.src = url;
    });
  }
  function drawJpeg(img, maxSide, quality) {
    var w = img.naturalWidth || img.width, h = img.naturalHeight || img.height;
    var scale = Math.min(1, maxSide / Math.max(w, h));
    var c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(w * scale)); c.height = Math.max(1, Math.round(h * scale));
    var ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(img, 0, 0, c.width, c.height);
    return new Promise(function (resolve, reject) {
      c.toBlob(function (b) { b ? resolve(b) : reject(new Error('照片轉檔失敗')); }, 'image/jpeg', quality);
    });
  }
  /** 一張照片 → { full: Blob, thumb: Blob }(JPEG) */
  async function prepare(file) {
    var img = await loadImage(file);
    return { full: await drawJpeg(img, FULL_SIDE, FULL_QUALITY), thumb: await drawJpeg(img, THUMB_SIDE, THUMB_QUALITY) };
  }

  function blobToBase64(blob) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { var s = String(r.result); resolve(s.slice(s.indexOf(',') + 1)); };
      r.onerror = function () { reject(r.error); };
      r.readAsDataURL(blob);
    });
  }
  function base64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function bytesToBase64(bytes) {
    var s = '', chunk = 0x8000;
    for (var i = 0; i < bytes.length; i += chunk) s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return btoa(s);
  }
  function dataUrl(b64) { return 'data:image/jpeg;base64,' + b64; }

  root.DWTQualityPhoto = {
    prepare: prepare, blobToBase64: blobToBase64, base64ToBytes: base64ToBytes,
    bytesToBase64: bytesToBase64, dataUrl: dataUrl
  };
})(typeof self !== 'undefined' ? self : this);
