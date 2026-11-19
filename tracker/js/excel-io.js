/*
 * excel-io.js — Excel 讀/寫(用 vendor/exceljs.min.js)
 *
 * readSheetRows(ExcelJS, arrayBuffer)  → 找出 WMS 匯出格式的工作表,回傳 {sheetName, rows}
 * buildExportWorkbook(ExcelJS, view)    → 產生「匯出Excel(趨勢+明細)」的 workbook
 *
 * 瀏覽器:window.DWTExcel;Node:require('./js/excel-io.js')
 * ExcelJS 由呼叫端傳入(瀏覽器是全域 ExcelJS,Node 是 require('exceljs')),
 * 這樣 tests/ 可以用同一份程式碼驗證解析結果。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.DWTExcel = factory(root.DWTCore);
})(typeof self !== 'undefined' ? self : this, function (Core) {
  'use strict';

  var EXCEL_FONT = '微軟正黑體';

  /** ExcelJS 的儲存格值可能是物件(公式、富文字、超連結…),統一轉成純值 */
  function plainValue(v) {
    if (v == null) return null;
    if (v instanceof Date) return v;
    if (typeof v !== 'object') return v;
    if (Array.isArray(v.richText)) return v.richText.map(function (t) { return t.text || ''; }).join('');
    if ('result' in v) return plainValue(v.result);           // formula / sharedFormula
    if ('text' in v) return v.text;                          // hyperlink
    if ('error' in v) return null;
    return String(v);
  }

  function sheetToRows(ws) {
    var rows = [];
    ws.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      var vals = row.values || []; // 1-based
      var arr = [];
      for (var c = 1; c < vals.length; c++) arr.push(plainValue(vals[c]));
      rows[rowNumber - 1] = arr;
    });
    for (var i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  /**
   * 讀取 xlsx 並挑出 WMS 匯出格式的工作表(優先名為 "Data" 的,其次任何具備必要欄位的)。
   * 找不到時回傳第一個工作表的列,讓 Core.parseSnapshot 產生「缺少欄位」的說明。
   */
  async function readSheetRows(ExcelJS, arrayBuffer) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    var sheets = wb.worksheets.slice();
    sheets.sort(function (a, b) { return (b.name === 'Data') - (a.name === 'Data'); });
    var firstRows = null, firstName = null;
    for (var i = 0; i < sheets.length; i++) {
      var rows = sheetToRows(sheets[i]);
      if (firstRows == null) { firstRows = rows; firstName = sheets[i].name; }
      if (Core.findHeader(rows)) return { sheetName: sheets[i].name, rows: rows };
    }
    return { sheetName: firstName, rows: firstRows || [] };
  }

  /**
   * 匯出內容 = 目前畫面上的資料(趨勢只含目前選取區間,明細是全部項目)。
   * view = { rangeStart, rangeEnd, weeklyRows, itemRows, chartImage: {base64, width, height} | null }
   */
  function buildExportWorkbook(ExcelJS, view) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('不良品倉追蹤');

    var IMAGE_ROWS = 0;
    if (view.chartImage) {
      var imgId = wb.addImage({ base64: view.chartImage.base64, extension: 'png' });
      ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: view.chartImage.width, height: view.chartImage.height } });
      IMAGE_ROWS = 16;
    }
    for (var i = 0; i < IMAGE_ROWS; i++) ws.addRow([]);

    var rs = view.rangeStart, re = view.rangeEnd;
    var trendRows = view.weeklyRows
      .filter(function (d) { return (!rs || d.date >= rs) && (!re || d.date <= re); })
      .slice()
      .sort(function (a, b) { return (a.date + a.warehouse) < (b.date + b.warehouse) ? -1 : 1; })
      .map(function (d) { return [d.date, d.week, d.warehouse, d.usd_amount, d.quantity, d.line_count]; });
    var itemsRows = Core.sortItems(view.itemRows)
      .map(function (it) { return [it.part_no, it.name || '', it.warehouse, it.supplier || '', it.qty == null ? '' : it.qty, it.usd_amount == null ? '' : it.usd_amount, it.status || '', it.owner || '', it.eta || '', (it.photos || []).length || '']; });

    var rangeLabel = (rs || re) ? '（' + (rs || '最早') + ' ~ ' + (re || '最新') + '，與頁面目前選取區間一致）' : '';
    var titleRow1 = ws.addRow(['不良品倉金額趨勢' + rangeLabel]);
    var headerRow1 = ws.addRow(['日期', '週次', '倉別', '美金金額', '數量', '列數']);
    trendRows.forEach(function (r) { ws.addRow(r); });
    ws.addRow([]);
    var titleRow2 = ws.addRow(['項目明細']);
    var headerRow2 = ws.addRow(['料號', '品名', '倉別', '供應商', '數量', '金額(美金)', '處理狀態', 'Owner', '預計完成時間', '照片數']);
    itemsRows.forEach(function (r) { ws.addRow(r); });

    ws.columns = [{ width: 18 }, { width: 30 }, { width: 10 }, { width: 14 }, { width: 10 }, { width: 14 }, { width: 16 }, { width: 10 }, { width: 16 }, { width: 8 }];

    ws.eachRow({ includeEmpty: true }, function (row) {
      row.eachCell({ includeEmpty: true }, function (cell) { cell.font = { name: EXCEL_FONT, size: 11 }; });
    });
    [titleRow1, titleRow2].forEach(function (r) { r.font = { name: EXCEL_FONT, size: 13, bold: true }; });
    [headerRow1, headerRow2].forEach(function (r) { r.font = { name: EXCEL_FONT, size: 11, bold: true }; });

    return wb;
  }

  return { EXCEL_FONT: EXCEL_FONT, plainValue: plainValue, readSheetRows: readSheetRows, buildExportWorkbook: buildExportWorkbook };
});
