/* Shared logic: calculations and Excel report generation.
 * Works in the browser (window.MealReport) and in Node (module.exports). */
(function (root) {
  'use strict';

  var INSTITUTIONS = [
    { id: 'pninim', name: 'סמינר פנינים' },
    { id: 'keter', name: 'סמינר כתר חיה' }
  ];
  var DEFAULT_PRICE = 19.95;
  var DEFAULT_VAT = 18;
  var MONTHS = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי',
    'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];
  var WEEKDAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

  function daysInMonth(year, month) {
    return new Date(year, month + 1, 0).getDate();
  }

  function weekday(year, month, day) {
    return new Date(year, month, day).getDay();
  }

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  function entry(state, day, instId) {
    var d = state.days[day] || {};
    var e = d[instId] || {};
    var qty = Number(e.qty) || 0;
    var price = e.price === undefined || e.price === '' || e.price === null
      ? DEFAULT_PRICE : Number(e.price) || 0;
    return { qty: qty, price: price, total: round2(qty * price) };
  }

  /* Totals per institution and overall. Prices include VAT; VAT is extracted. */
  function computeTotals(state, year, month, vatRate) {
    var n = daysInMonth(year, month);
    var rate = Number(vatRate) / 100;
    var result = { institutions: {}, all: { qty: 0, total: 0 } };
    INSTITUTIONS.forEach(function (inst) {
      var qty = 0, total = 0;
      for (var day = 1; day <= n; day++) {
        var e = entry(state, day, inst.id);
        qty += e.qty;
        total += e.total;
      }
      total = round2(total);
      var beforeVat = round2(total / (1 + rate));
      result.institutions[inst.id] = {
        qty: qty, total: total, beforeVat: beforeVat, vat: round2(total - beforeVat)
      };
      result.all.qty += qty;
      result.all.total += total;
    });
    // Overall figures are sums of the per-institution rounded figures, as in the report.
    result.all.total = round2(result.all.total);
    result.all.beforeVat = 0;
    INSTITUTIONS.forEach(function (inst) { result.all.beforeVat += result.institutions[inst.id].beforeVat; });
    result.all.beforeVat = round2(result.all.beforeVat);
    result.all.vat = round2(result.all.total - result.all.beforeVat);
    return result;
  }

  /* Days that have meals or a note — these are the rows written to the report. */
  function reportDays(state, year, month) {
    var days = [];
    for (var day = 1; day <= daysInMonth(year, month); day++) {
      var hasMeals = INSTITUTIONS.some(function (inst) {
        return entry(state, day, inst.id).qty > 0;
      });
      var note = (state.days[day] && state.days[day].note || '').trim();
      if (hasMeals || note) days.push(day);
    }
    return days;
  }

  function reportFileName(year, month) {
    return 'meals-report-' + year + '-' + String(month + 1).padStart(2, '0') + '.xlsx';
  }

  /* ---------- Excel ---------- */

  var COLORS = {
    title: 'FF1F3A5F',
    header: 'FF2E5A88',
    pninim: 'FFDCE9F5',
    keter: 'FFE6F2E6',
    subHeader: 'FFF2F2F2',
    shabbat: 'FFFAF5E6',
    total: 'FFFFF2CC',
    border: 'FFB7C3D0'
  };
  var MONEY = '#,##0.00 "₪"';
  var thin = { style: 'thin', color: { argb: COLORS.border } };
  var BORDER = { top: thin, left: thin, bottom: thin, right: thin };

  function fill(argb) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } };
  }

  function styleRange(ws, row, fromCol, toCol, style) {
    for (var c = fromCol; c <= toCol; c++) {
      var cell = ws.getRow(row).getCell(c);
      Object.keys(style).forEach(function (k) { cell[k] = style[k]; });
    }
  }

  /* Columns: A date | B day | C-E Pninim (qty, price, total) |
   * F-H Keter Chaya (qty, price, total) | I daily total | J notes */
  function buildWorkbook(ExcelJS, state, year, month, vatRate) {
    var totals = computeTotals(state, year, month, vatRate);
    var days = reportDays(state, year, month);
    var monthLabel = MONTHS[month] + ' ' + year;

    var wb = new ExcelJS.Workbook();
    wb.creator = 'מערכת דוחות מנות';
    wb.created = new Date();
    var ws = wb.addWorksheet('דוח ' + monthLabel, {
      views: [{ rightToLeft: true, state: 'frozen', ySplit: 5 }],
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true,
        fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }
      },
      headerFooter: { oddFooter: '&Cעמוד &P מתוך &N' }
    });
    ws.properties.defaultRowHeight = 18;
    ws.columns = [
      { width: 12 }, { width: 9 },
      { width: 9 }, { width: 11 }, { width: 13 },
      { width: 9 }, { width: 11 }, { width: 13 },
      { width: 14 }, { width: 30 }
    ];
    var center = { vertical: 'middle', horizontal: 'center', wrapText: true };
    var font = 'Arial';

    // Title
    ws.mergeCells('A1:J1');
    var title = ws.getCell('A1');
    title.value = 'דוח סיכום מנות – ' + monthLabel;
    title.font = { name: font, size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    title.fill = fill(COLORS.title);
    title.alignment = center;
    ws.getRow(1).height = 34;

    ws.mergeCells('A2:J2');
    var sub = ws.getCell('A2');
    sub.value = 'סמינר פנינים | סמינר כתר חיה   •   הופק בתאריך: ' +
      new Date().toLocaleDateString('he-IL');
    sub.font = { name: font, size: 10, italic: true, color: { argb: 'FF555555' } };
    sub.alignment = center;

    // Header rows (4 & 5)
    var hFont = { name: font, bold: true, color: { argb: 'FFFFFFFF' } };
    ws.mergeCells('A4:A5'); ws.getCell('A4').value = 'תאריך';
    ws.mergeCells('B4:B5'); ws.getCell('B4').value = 'יום';
    ws.mergeCells('C4:E4'); ws.getCell('C4').value = INSTITUTIONS[0].name;
    ws.mergeCells('F4:H4'); ws.getCell('F4').value = INSTITUTIONS[1].name;
    ws.mergeCells('I4:I5'); ws.getCell('I4').value = 'סה"כ ליום';
    ws.mergeCells('J4:J5'); ws.getCell('J4').value = 'הערות';
    styleRange(ws, 4, 1, 10, { font: hFont, fill: fill(COLORS.header), alignment: center, border: BORDER });
    styleRange(ws, 5, 1, 10, { font: hFont, fill: fill(COLORS.header), alignment: center, border: BORDER });
    ['מנות', 'מחיר למנה', 'סה"כ', 'מנות', 'מחיר למנה', 'סה"כ'].forEach(function (t, i) {
      var cell = ws.getRow(5).getCell(3 + i);
      cell.value = t;
      cell.font = { name: font, bold: true, color: { argb: 'FF1F3A5F' } };
      cell.fill = fill(i < 3 ? COLORS.pninim : COLORS.keter);
    });
    ws.getRow(4).height = 22;
    ws.getRow(5).height = 20;

    // Data rows
    var first = 6;
    days.forEach(function (day, idx) {
      var r = first + idx;
      var p = entry(state, day, 'pninim');
      var k = entry(state, day, 'keter');
      var wd = weekday(year, month, day);
      var row = ws.getRow(r);
      row.getCell(1).value = new Date(Date.UTC(year, month, day));
      row.getCell(1).numFmt = 'dd/mm/yyyy';
      row.getCell(2).value = WEEKDAYS[wd];
      row.getCell(3).value = p.qty;
      row.getCell(4).value = p.price;
      row.getCell(5).value = { formula: 'C' + r + '*D' + r, result: p.total };
      row.getCell(6).value = k.qty;
      row.getCell(7).value = k.price;
      row.getCell(8).value = { formula: 'F' + r + '*G' + r, result: k.total };
      row.getCell(9).value = { formula: 'E' + r + '+H' + r, result: round2(p.total + k.total) };
      row.getCell(10).value = (state.days[day] && state.days[day].note || '').trim();
      for (var c = 1; c <= 10; c++) {
        var cell = row.getCell(c);
        cell.border = BORDER;
        cell.font = { name: font, size: 11, bold: c === 9 };
        cell.alignment = c === 10
          ? { vertical: 'middle', horizontal: 'right', wrapText: true }
          : center;
        if ([4, 5, 7, 8, 9].indexOf(c) !== -1) cell.numFmt = MONEY;
        if (wd === 6) cell.fill = fill(COLORS.shabbat);
      }
    });
    var last = first + days.length - 1;

    // Totals row
    var tr = last + 1;
    var tRow = ws.getRow(tr);
    ws.mergeCells('A' + tr + ':B' + tr);
    tRow.getCell(1).value = 'סה"כ חודשי';
    var sum = function (col, result) {
      return days.length
        ? { formula: 'SUM(' + col + first + ':' + col + last + ')', result: result }
        : result;
    };
    var tp = totals.institutions.pninim, tk = totals.institutions.keter;
    tRow.getCell(3).value = sum('C', tp.qty);
    tRow.getCell(5).value = sum('E', tp.total);
    tRow.getCell(6).value = sum('F', tk.qty);
    tRow.getCell(8).value = sum('H', tk.total);
    tRow.getCell(9).value = sum('I', totals.all.total);
    styleRange(ws, tr, 1, 10, {
      font: { name: font, bold: true, size: 12 }, fill: fill(COLORS.total),
      alignment: center, border: { top: { style: 'medium' }, bottom: { style: 'medium' }, left: thin, right: thin }
    });
    [5, 8, 9].forEach(function (c) { tRow.getCell(c).numFmt = MONEY; });
    tRow.height = 22;

    // Summary block
    var s = tr + 3;
    ws.mergeCells('C' + s + ':H' + s);
    var sTitle = ws.getCell('C' + s);
    sTitle.value = 'סיכום לתשלום';
    styleRange(ws, s, 3, 8, { font: { name: font, bold: true, size: 13, color: { argb: 'FFFFFFFF' } }, fill: fill(COLORS.title), alignment: center });
    ws.getRow(s).height = 24;

    var hr = s + 1;
    ws.mergeCells('C' + hr + ':D' + hr);
    ws.mergeCells('E' + hr + ':F' + hr);
    ws.getCell('E' + hr).value = INSTITUTIONS[0].name;
    ws.getCell('G' + hr).value = INSTITUTIONS[1].name;
    ws.getCell('H' + hr).value = 'סה"כ';
    styleRange(ws, hr, 3, 8, { font: { name: font, bold: true }, fill: fill(COLORS.subHeader), alignment: center, border: BORDER });
    ws.getCell('E' + hr).fill = fill(COLORS.pninim);
    ws.getCell('G' + hr).fill = fill(COLORS.keter);

    var rate = Number(vatRate) / 100;
    var rows = [
      ['מספר מנות', tp.qty, tk.qty, totals.all.qty, null],
      ['סכום לפני מע"מ', tp.beforeVat, tk.beforeVat, totals.all.beforeVat, 'before'],
      ['מע"מ (' + vatRate + '%)', tp.vat, tk.vat, totals.all.vat, 'vat'],
      ['סה"כ לתשלום (כולל מע"מ)', tp.total, tk.total, totals.all.total, 'total']
    ];
    var totalRowOf = s + 5; // row of "total incl. VAT"
    rows.forEach(function (def, i) {
      var r = hr + 1 + i;
      ws.mergeCells('C' + r + ':D' + r);
      ws.mergeCells('E' + r + ':F' + r);
      ws.getCell('C' + r).value = def[0];
      var srcQty = { E: 'C' + tr, G: 'F' + tr }, srcTotal = { E: 'E' + tr, G: 'H' + tr };
      [['E', def[1]], ['G', def[2]]].forEach(function (pair) {
        var col = pair[0], val = pair[1], f;
        if (def[4] === null) f = srcQty[col];
        else if (def[4] === 'total') f = srcTotal[col];
        else if (def[4] === 'before') f = 'ROUND(' + col + totalRowOf + '/(1+' + rate + '),2)';
        else f = col + totalRowOf + '-' + col + (totalRowOf - 2);
        ws.getCell(col + r).value = { formula: f, result: val };
      });
      ws.getCell('H' + r).value = { formula: 'E' + r + '+G' + r, result: def[3] };
      var isTotal = def[4] === 'total';
      styleRange(ws, r, 3, 8, {
        font: { name: font, bold: isTotal || i === 0, size: isTotal ? 13 : 11 },
        alignment: center, border: BORDER,
        fill: fill(isTotal ? COLORS.total : 'FFFFFFFF')
      });
      ws.getCell('C' + r).alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
      if (i > 0) ['E', 'G', 'H'].forEach(function (c) { ws.getCell(c + r).numFmt = MONEY; });
      ws.getRow(r).height = isTotal ? 24 : 20;
    });

    var note = (state.generalNote || '').trim();
    if (note) {
      var nr = hr + rows.length + 3;
      ws.getCell('A' + nr).value = 'הערות כלליות:';
      ws.getCell('A' + nr).font = { name: font, bold: true };
      ws.mergeCells('B' + nr + ':J' + (nr + 2));
      var nc = ws.getCell('B' + nr);
      nc.value = note;
      nc.font = { name: font };
      nc.alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
      nc.border = BORDER;
    }

    ws.pageSetup.printArea = 'A1:J' + ws.rowCount;
    ws.pageSetup.printTitlesRow = '4:5';
    return wb;
  }

  var api = {
    INSTITUTIONS: INSTITUTIONS,
    DEFAULT_PRICE: DEFAULT_PRICE,
    DEFAULT_VAT: DEFAULT_VAT,
    MONTHS: MONTHS,
    WEEKDAYS: WEEKDAYS,
    daysInMonth: daysInMonth,
    weekday: weekday,
    round2: round2,
    entry: entry,
    computeTotals: computeTotals,
    reportDays: reportDays,
    reportFileName: reportFileName,
    buildWorkbook: buildWorkbook
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MealReport = api;
})(this);
