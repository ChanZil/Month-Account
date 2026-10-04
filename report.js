/* Shared logic: calculations and Excel report generation.
 * Works in the browser (window.MealReport) and in Node (module.exports).
 *
 * Data model (one object per report period):
 *   { from: 'YYYY-MM-DD'|null, to: 'YYYY-MM-DD'|null,
 *     prices: { pninim: 19.95, keter: 19.95 },                 // default price per meal
 *     days: { 'YYYY-MM-DD': { special: true|false|undefined,   // Shabbat / holiday override
 *                             lines: [ { pninim: {qty, price}, keter: {qty, price}, note } ] } } }
 * Prices are before VAT; VAT is added on top at the end of the report. */
(function (root) {
  'use strict';

  var INSTITUTIONS = [
    { id: 'pninim', name: 'פנינים', fullName: 'סמינר פנינים' },
    { id: 'keter', name: 'כתר חיה', fullName: 'סמינר כתר חיה' }
  ];
  var DEFAULT_PRICE = 19.95;
  var DEFAULT_VAT = 18;
  var MAX_DAYS = 92;
  var MONTHS = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי',
    'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];
  var WEEKDAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];

  /* ---------- dates (ISO strings 'YYYY-MM-DD', local time) ---------- */

  function pad(n) { return String(n).padStart(2, '0'); }
  function daysInMonth(year, month) { return new Date(year, month + 1, 0).getDate(); }
  function parseIso(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function toIso(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function weekday(iso) { return parseIso(iso).getDay(); }
  function round2(n) { return Math.round(n * 100) / 100; }

  function monthRange(year, month) {
    return {
      from: year + '-' + pad(month + 1) + '-01',
      to: year + '-' + pad(month + 1) + '-' + pad(daysInMonth(year, month))
    };
  }

  /* The report covers the calendar month unless the user widened/narrowed the range. */
  function periodOf(state, year, month) {
    var def = monthRange(year, month);
    return { from: state.from || def.from, to: state.to || def.to };
  }

  function rangeDays(from, to) {
    var out = [], d = parseIso(from);
    while (toIso(d) <= to && out.length < MAX_DAYS) {
      out.push(toIso(d));
      d.setDate(d.getDate() + 1);
    }
    return out;
  }

  /* ---------- entries ---------- */

  function isSpecial(state, iso) {
    var d = state.days[iso];
    if (d && typeof d.special === 'boolean') return d.special;
    return weekday(iso) === 6;
  }

  function linesOf(state, iso) {
    var d = state.days[iso];
    return (d && d.lines) || [];
  }

  function hasValue(v) { return v !== undefined && v !== null && v !== ''; }

  function defaultPrice(state, instId) {
    var p = state.prices && state.prices[instId];
    return hasValue(p) ? Number(p) || 0 : DEFAULT_PRICE;
  }

  function lineEntry(state, line, instId) {
    var e = (line && line[instId]) || {};
    var qty = Number(e.qty) || 0;
    var price = hasValue(e.price) ? Number(e.price) || 0 : defaultPrice(state, instId);
    return { qty: qty, price: price, total: round2(qty * price), custom: hasValue(e.price) };
  }

  function lineHasContent(state, line) {
    return INSTITUTIONS.some(function (inst) { return lineEntry(state, line, inst.id).qty > 0; }) ||
      !!String(line && line.note || '').trim();
  }

  /* Lines written to the report; a date without any is still listed once (as in the existing reports). */
  function reportLines(state, iso) {
    var lines = linesOf(state, iso).filter(function (l) { return lineHasContent(state, l); });
    return lines.length ? lines : [null];
  }

  function hasData(state, from, to) {
    return rangeDays(from, to).some(function (iso) {
      return linesOf(state, iso).some(function (l) { return lineHasContent(state, l); });
    });
  }

  /* ---------- totals ---------- */

  function finish(o, rate) {
    var incl = o.total * (1 + rate);
    o.total = round2(o.total);
    o.special = round2(o.special);
    o.regular = round2(o.total - o.special);
    o.incl = round2(incl);
    o.vat = round2(o.incl - o.total);
    return o;
  }

  /* Totals per institution and overall: meals, regular days, Shabbat/holiday, before and after VAT. */
  function computeTotals(state, from, to, vatRate) {
    var rate = Number(vatRate) / 100;
    var days = rangeDays(from, to);
    var all = { qty: 0, total: 0, special: 0 };
    var institutions = {};
    INSTITUTIONS.forEach(function (inst) {
      var o = { qty: 0, total: 0, special: 0 };
      days.forEach(function (iso) {
        var special = isSpecial(state, iso);
        linesOf(state, iso).forEach(function (line) {
          var e = lineEntry(state, line, inst.id);
          o.qty += e.qty;
          o.total += e.total;
          if (special) o.special += e.total;
        });
      });
      all.qty += o.qty;
      all.total += o.total;
      all.special += o.special;
      institutions[inst.id] = finish(o, rate);
    });
    return { institutions: institutions, all: finish(all, rate) };
  }

  function reportFileName(year, month) {
    return 'meals-report-' + year + '-' + pad(month + 1) + '.xlsx';
  }

  /* ---------- Excel (mirrors the existing monthly reports) ---------- */

  var FILL = {
    blue: 'FFBDD7EE',   // day, date, line total, notes
    pninim: 'FFFCE4D6',
    keter: 'FFE2EFDA'
  };
  var MONEY = '"₪"\\ #,##0.00;[Red]"₪"\\ \\-#,##0.00';
  var thin = { style: 'thin' };
  var BORDER = { top: thin, left: thin, bottom: thin, right: thin };
  // Column groups: A,B,I,J blue | C-E Pninim | F-H Keter
  var COL_FILL = [FILL.blue, FILL.blue, FILL.pninim, FILL.pninim, FILL.pninim,
    FILL.keter, FILL.keter, FILL.keter, FILL.blue, FILL.blue];
  var MONEY_COLS = [4, 5, 7, 8, 9];

  function letter(col) { return String.fromCharCode(64 + col); }

  function styleRow(ws, r, bold) {
    for (var c = 1; c <= 10; c++) {
      var cell = ws.getCell(r, c);
      cell.font = { name: 'Arial', size: 11, bold: !!bold };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: c === 10 };
      cell.border = BORDER;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COL_FILL[c - 1] } };
    }
  }

  function buildWorkbook(ExcelJS, state, year, month, vatRate) {
    var period = periodOf(state, year, month);
    var days = rangeDays(period.from, period.to);
    var totals = computeTotals(state, period.from, period.to, vatRate);

    var wb = new ExcelJS.Workbook();
    wb.creator = 'מערכת דוחות מנות';
    wb.created = new Date();
    var ws = wb.addWorksheet('דוח ' + MONTHS[month] + ' ' + year, {
      views: [{ rightToLeft: true, state: 'frozen', ySplit: 2 }],
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }
      }
    });
    [11, 13, 12, 11, 13, 12.5, 11, 12, 14, 20].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });

    // Header (rows 1-2)
    styleRow(ws, 1, true);
    styleRow(ws, 2, true);
    ws.getCell('A1').value = 'יום בשבוע';
    ws.getCell('B1').value = 'תאריך ביצוע';
    ws.getCell('C1').value = INSTITUTIONS[0].name;
    ws.getCell('F1').value = INSTITUTIONS[1].name;
    ws.getCell('I1').value = 'מחיר';
    ws.getCell('J1').value = 'הערות';
    [3, 6].forEach(function (c) {
      ws.getCell(2, c).value = 'כמות סועדות';
      ws.getCell(2, c + 1).value = 'מחיר למנה';
      ws.getCell(2, c + 2).value = 'סה"כ';
    });
    ['A', 'B', 'I', 'J'].forEach(function (col) { ws.mergeCells(col + '1:' + col + '2'); });
    ws.mergeCells('C1:E1');
    ws.mergeCells('F1:H1');

    // Data: every date of the period; one row per line, date and day merged across a date's lines
    var r = 3;
    var specialRanges = [];
    days.forEach(function (iso) {
      var start = r;
      reportLines(state, iso).forEach(function (line) {
        styleRow(ws, r, false);
        MONEY_COLS.forEach(function (c) { ws.getCell(r, c).numFmt = MONEY; });
        var sum = 0, any = false;
        INSTITUTIONS.forEach(function (inst, i) {
          var e = line ? lineEntry(state, line, inst.id) : { qty: 0 };
          if (e.qty > 0) {
            var q = 3 + i * 3;
            ws.getCell(r, q).value = e.qty;
            ws.getCell(r, q + 1).value = e.price;
            ws.getCell(r, q + 2).value = {
              formula: letter(q) + r + '*' + letter(q + 1) + r, result: e.total
            };
            sum += e.total;
            any = true;
          }
        });
        if (any) ws.getCell(r, 9).value = { formula: 'E' + r + '+H' + r, result: round2(sum) };
        var note = line && String(line.note || '').trim();
        if (note) ws.getCell(r, 10).value = note;
        r++;
      });
      var end = r - 1;
      var d = parseIso(iso);
      ws.getCell(start, 1).value = WEEKDAYS[d.getDay()];
      ws.getCell(start, 2).value = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      ws.getCell(start, 2).numFmt = 'dd.mm.yyyy';
      if (end > start) {
        ws.mergeCells('A' + start + ':A' + end);
        ws.mergeCells('B' + start + ':B' + end);
      }
      if (isSpecial(state, iso)) specialRanges.push([start, end]);
    });
    var last = r - 1;

    // Summary rows (same labels and order as the existing reports)
    var rows = { qty: r, regular: r + 1, special: r + 2, net: r + 3, gross: r + 4 };
    var labels = {
      qty: 'סה"כ מנות', regular: 'סה"כ לתשלום ימי חול', special: 'סה"כ לתשלום ימי שבת וחג',
      net: 'סה"כ לתשלום ללא מע"מ', gross: 'סה"כ לתשלום כולל מע"מ'
    };
    Object.keys(rows).forEach(function (key) {
      var row = rows[key];
      styleRow(ws, row, false);
      var label = ws.getCell(row, 1);
      label.value = labels[key];
      label.font = { name: 'Arial', size: 11, bold: true };
      label.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      ws.mergeCells('A' + row + ':B' + row);
      ws.mergeCells('C' + row + ':E' + row);
      ws.mergeCells('F' + row + ':H' + row);
      if (key !== 'qty') [3, 6, 9].forEach(function (c) { ws.getCell(row, c).numFmt = MONEY; });
      if (key === 'regular' || key === 'special') ws.getRow(row).height = 30;
    });

    INSTITUTIONS.forEach(function (inst, i) {
      var c = 3 + i * 3;                 // merged value cell: C or F
      var vc = letter(c);
      var qtyCol = letter(c);            // quantity column of this institution
      var totalCol = letter(c + 2);      // line-total column: E or H
      var t = totals.institutions[inst.id];
      var specialRefs = specialRanges.map(function (p) {
        return totalCol + p[0] + (p[1] > p[0] ? ':' + totalCol + p[1] : '');
      });
      ws.getCell(rows.qty, c).value = { formula: 'SUM(' + qtyCol + '3:' + qtyCol + last + ')', result: t.qty };
      ws.getCell(rows.net, c).value = { formula: 'SUM(' + totalCol + '3:' + totalCol + last + ')', result: t.total };
      ws.getCell(rows.special, c).value = specialRefs.length
        ? { formula: 'SUM(' + specialRefs.join(',') + ')', result: t.special } : 0;
      ws.getCell(rows.regular, c).value = {
        formula: vc + rows.net + '-' + vc + rows.special, result: t.regular
      };
      ws.getCell(rows.gross, c).value = {
        formula: vc + rows.net + '*(1+' + Number(vatRate) + '%)', result: t.incl
      };
    });
    var all = totals.all;
    var allResults = { qty: all.qty, regular: all.regular, special: all.special, net: all.total, gross: all.incl };
    Object.keys(rows).forEach(function (key) {
      ws.getCell(rows[key], 9).value = {
        formula: 'C' + rows[key] + '+F' + rows[key], result: allResults[key]
      };
    });

    ws.pageSetup.printTitlesRow = '1:2';
    ws.pageSetup.printArea = 'A1:J' + rows.gross;
    return wb;
  }

  var api = {
    INSTITUTIONS: INSTITUTIONS,
    DEFAULT_PRICE: DEFAULT_PRICE,
    DEFAULT_VAT: DEFAULT_VAT,
    MAX_DAYS: MAX_DAYS,
    MONTHS: MONTHS,
    WEEKDAYS: WEEKDAYS,
    daysInMonth: daysInMonth,
    weekday: weekday,
    parseIso: parseIso,
    toIso: toIso,
    monthRange: monthRange,
    periodOf: periodOf,
    rangeDays: rangeDays,
    round2: round2,
    isSpecial: isSpecial,
    linesOf: linesOf,
    lineEntry: lineEntry,
    defaultPrice: defaultPrice,
    hasData: hasData,
    computeTotals: computeTotals,
    reportFileName: reportFileName,
    buildWorkbook: buildWorkbook
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MealReport = api;
})(this);
