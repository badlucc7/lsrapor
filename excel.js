/**
 * excel.js - rapor nesnesinden (Engine.gs çıktısı) formüllü .xlsx üretir.
 * Tarayıcıda ExcelJS (CDN) ile, Node'da `exceljs` paketiyle aynı kod çalışır.
 */
(function (root) {
  var TZ_MS = 3 * 3600 * 1000;
  var FMT = { int: '0', pct: '0.0%', gbp: '#,##0"£"', d1: '0.0', txt: '@' };
  var FONT = 'Arial';
  var NAVY = 'FF1F4E78', LIGHT = 'FFD9E1F2';

  function colLetter(i) {
    var s = ''; i = i + 1;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  var thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  var BORDER = { top: thin, left: thin, bottom: thin, right: thin };

  function buildWorkbook(ExcelJS, rep) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Lara Smile';
    wb.calcProperties = { fullCalcOnLoad: true };
    var brand = rep.meta.brand;
    var period = rep.meta.from.split('-').reverse().join('.') + ' - ' + rep.meta.to.split('-').reverse().join('.');

    function head(ws, title) {
      ws.getCell('A1').value = brand + ' - ' + title;
      ws.getCell('A1').font = { name: FONT, bold: true, size: 14, color: { argb: NAVY } };
      ws.getCell('A2').value = 'Dönem: ' + period + (rep.meta.eskiMod === 'off' ? ' (Eski datalar dahil değil)' : rep.meta.eskiMod === 'all' ? ' (Eski datalar: booklama tarihi dönem dışı olanlar dahil)' : '') +
        ((rep.meta.sources && rep.meta.sources.length) ? ' | Kaynak: ' + rep.meta.sources.join(', ') : '');
      ws.getCell('A2').font = { name: FONT, italic: true, size: 10, color: { argb: 'FF595959' } };
    }

    function addSummary(sh) {
      var ws = wb.addWorksheet(sh.name);
      head(ws, sh.name);
      (sh.widths || []).forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
      var r = 4;
      sh.tables.forEach(function (t) {
        ws.getCell(r, 1).value = t.title;
        ws.getCell(r, 1).font = { name: FONT, bold: true, size: 11, color: { argb: NAVY } };
        r++;
        if (t.note) {
          ws.getCell(r, 1).value = t.note;
          ws.getCell(r, 1).font = { name: FONT, italic: true, size: 9, color: { argb: 'FF808080' } };
          r++;
        }
        t.headers.forEach(function (h, j) {
          var c = ws.getCell(r, j + 1);
          c.value = h;
          c.font = { name: FONT, bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          c.border = BORDER;
        });
        ws.getRow(r).height = 32;
        r++;
        function put(cells, isTotal) {
          cells.forEach(function (cell, j) {
            var c = ws.getCell(r, j + 1);
            if (cell.f) c.value = { formula: cell.f, result: cell.v };
            else c.value = cell.v;
            if (cell.fmt && FMT[cell.fmt]) c.numFmt = FMT[cell.fmt];
            c.font = { name: FONT, size: 10, bold: !!(isTotal || cell.b) };
            if (isTotal) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
            c.border = BORDER;
          });
          r++;
        }
        t.body.forEach(function (row) { put(row, false); });
        if (t.total) put(t.total, true);
        r++;
      });
      ws.views = [{ state: 'frozen', ySplit: 2 }];
      return ws;
    }

    function addHam() {
      var ws = wb.addWorksheet('Ham Veri');
      var cols = rep.hv.cols;
      cols.forEach(function (col, j) {
        var c = ws.getCell(1, j + 1);
        c.value = col.h;
        c.font = { name: FONT, bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        c.border = BORDER;
        ws.getColumn(j + 1).width = col.w;
      });
      ws.getRow(1).height = 32;
      rep.hv.rows.forEach(function (row, i) {
        cols.forEach(function (col, j) {
          var v = row[col.k];
          var c = ws.getCell(i + 2, j + 1);
          if (v === null || v === undefined || v === '') { c.border = BORDER; c.font = { name: FONT, size: 10 }; return; }
          if (col.date) { c.value = new Date(v + TZ_MS); c.numFmt = 'dd.mm.yyyy hh:mm'; }
          else c.value = v;
          if (col.fmt && FMT[col.fmt]) c.numFmt = FMT[col.fmt];
          c.font = { name: FONT, size: 10 };
          c.border = BORDER;
        });
      });
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
    }

    // Sıra: Genel Özet, Ham Veri, diğer özet sayfaları
    addSummary(rep.sheets[0]);
    addHam();
    for (var i = 1; i < rep.sheets.length; i++) addSummary(rep.sheets[i]);
    return wb;
  }

  var api = { buildWorkbook: buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.KommoExcel = api;
})(typeof window !== 'undefined' ? window : this);
