(function () {
  'use strict';

  var R = window.MealReport;
  var $ = function (sel) { return document.querySelector(sel); };
  var money = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS' });

  var monthSel = $('#month'), yearSel = $('#year'), vatInput = $('#vat');
  var fromInput = $('#from'), toInput = $('#to'), tbody = $('#rows');
  var fileInput = $('#file-name');
  var year, month, state;

  /* ---------- storage (localStorage may be unavailable) ---------- */
  function load(key, fallback) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (e) { return fallback; }
  }
  function save(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* ignore */ }
  }
  function monthKey() { return 'meals:v2:' + year + '-' + (month + 1); }

  var autosave = $('#autosave'), saveTimer;
  function saveState() {
    save(monthKey(), state);
    autosave.classList.add('saving');
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () { autosave.classList.remove('saving'); }, 700);
  }

  var toastEl = $('#toast'), toastTimer;
  function toast(msg) {
    toastEl.textContent = msg;
    toastEl.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove('show'); }, 2600);
  }

  var settings = load('meals:settings', { vat: R.DEFAULT_VAT });
  function vatRate() { return vatInput.value === '' ? R.DEFAULT_VAT : Number(vatInput.value); }

  /* ---------- period selectors ---------- */
  var now = new Date();
  R.MONTHS.forEach(function (name, i) { monthSel.add(new Option(name, i)); });
  for (var y = now.getFullYear() - 3; y <= now.getFullYear() + 2; y++) yearSel.add(new Option(y, y));
  var last = load('meals:lastPeriod', null);
  monthSel.value = last ? last.month : now.getMonth();
  yearSel.value = last ? last.year : now.getFullYear();
  vatInput.value = settings.vat;

  /* ---------- state helpers ---------- */
  function period() { return R.periodOf(state, year, month); }

  function ensureDay(iso) {
    var d = state.days[iso] || (state.days[iso] = {});
    if (!d.lines) d.lines = [];
    return d;
  }
  function ensureLine(iso, idx) {
    var d = ensureDay(iso);
    while (d.lines.length <= idx) d.lines.push({});
    return d.lines[idx];
  }

  /* ---------- table ---------- */
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }
  function td(cls, content) {
    var c = el('td', cls);
    if (content instanceof Node) c.appendChild(content); else if (content !== undefined) c.textContent = content;
    return c;
  }
  function numInput(cls, value, step, inst, field, placeholder) {
    var input = el('input', cls);
    input.type = 'number';
    input.min = '0';
    input.step = step;
    input.inputMode = step === '1' ? 'numeric' : 'decimal';
    input.value = value;
    input.dataset.inst = inst;
    input.dataset.field = field;
    if (placeholder) input.placeholder = placeholder;
    input.setAttribute('enterkeyhint', 'next');
    return input;
  }
  function roundBtn(cls, label, title) {
    var b = el('button', 'row-btn ' + cls, label);
    b.type = 'button';
    b.title = title;
    b.setAttribute('aria-label', title);
    return b;
  }

  function buildRow(iso, idx, count, line) {
    var tr = el('tr');
    tr.dataset.iso = iso;
    tr.dataset.idx = idx;
    if (idx === 0) tr.classList.add('group-start');
    if (R.isSpecial(state, iso)) tr.classList.add('special');

    if (idx === 0) {
      var p = iso.split('-');
      var dateCell = td('date', p[2] + '.' + p[1]);
      dateCell.rowSpan = count;
      tr.appendChild(dateCell);

      var dayCell = td('day-cell');
      dayCell.rowSpan = count;
      dayCell.appendChild(el('div', 'day-name', R.WEEKDAYS[R.weekday(iso)]));
      var toggle = el('button', 'special-toggle');
      toggle.type = 'button';
      dayCell.appendChild(toggle);
      tr.appendChild(dayCell);
      setToggle(toggle, R.isSpecial(state, iso));
    }

    R.INSTITUTIONS.forEach(function (inst, i) {
      var e = line[inst.id] || {};
      var entry = R.lineEntry(state, line, inst.id);
      var qty = numInput('qty', e.qty === undefined || e.qty === '' ? '' : e.qty, '1', inst.id, 'qty', '0');
      var price = numInput('price', entry.price, '0.01', inst.id, 'price');
      qty.setAttribute('aria-label', inst.name + ' – כמות סועדות');
      price.setAttribute('aria-label', inst.name + ' – מחיר למנה');
      if (entry.custom) price.classList.add('custom');
      tr.appendChild(td((i ? 'sep ' : '') + 'c-qty c-' + inst.id, qty));
      tr.appendChild(td('c-price c-' + inst.id, price));
      tr.appendChild(td('money c-total c-' + inst.id + ' total-' + inst.id));
    });
    tr.appendChild(td('money line-total sep'));

    var note = el('input', 'note');
    note.type = 'text';
    note.dataset.field = 'note';
    note.value = line.note || '';
    note.setAttribute('aria-label', 'הערה');
    note.setAttribute('enterkeyhint', 'next');
    tr.appendChild(td('c-note', note));

    tr.appendChild(td('act', idx === 0
      ? roundBtn('add', '+', 'הוספת שורה לתאריך זה')
      : roundBtn('remove', '×', 'מחיקת שורה')));
    updateRow(tr);
    return tr;
  }

  function setToggle(btn, special) {
    btn.setAttribute('aria-pressed', special ? 'true' : 'false');
    btn.textContent = special ? 'שבת/חג' : 'סמן כחג';
    btn.title = special
      ? 'יום שבת/חג: הסכום מחושב בנפרד. לחצו לביטול'
      : 'סמנו כיום חג כדי שהסכום יחושב בשורת "שבת וחג"';
  }

  function drawRows() {
    var p = period();
    tbody.innerHTML = '';
    R.rangeDays(p.from, p.to).forEach(function (iso) {
      var lines = R.linesOf(state, iso);
      var count = Math.max(1, lines.length);
      for (var i = 0; i < count; i++) tbody.appendChild(buildRow(iso, i, count, lines[i] || {}));
    });
  }

  function render() {
    year = Number(yearSel.value);
    month = Number(monthSel.value);
    state = load(monthKey(), null) || { from: null, to: null, prices: {}, days: {} };
    if (!state.prices) state.prices = {};
    if (!state.days) state.days = {};
    save('meals:lastPeriod', { year: year, month: month });
    $('#period-title').textContent = R.MONTHS[month] + ' ' + year;

    var p = period();
    fromInput.value = p.from;
    toInput.value = p.to;
    fileInput.placeholder = R.defaultFileBase(year, month);
    fileInput.value = state.fileName || '';
    document.querySelectorAll('.setting[data-inst]').forEach(function (box) {
      box.querySelector('.bulk-price').value = R.defaultPrice(state, box.dataset.inst);
    });
    drawRows();
    updateTotals();
  }

  function updateRow(tr) {
    var iso = tr.dataset.iso;
    var line = R.linesOf(state, iso)[Number(tr.dataset.idx)] || {};
    var sum = 0, any = false;
    R.INSTITUTIONS.forEach(function (inst) {
      var e = R.lineEntry(state, line, inst.id);
      tr.querySelector('.total-' + inst.id).textContent = e.qty ? money.format(e.total) : '';
      tr.querySelector('input.price[data-inst="' + inst.id + '"]').classList.toggle('idle', !e.qty);
      if (e.qty) { sum += e.total; any = true; }
    });
    tr.querySelector('.line-total').textContent = any ? money.format(sum) : '';
    tr.classList.toggle('has-data', any);
    tr.classList.toggle('has-note', !!String(line.note || '').trim());
  }

  function updateTotals() {
    var p = period();
    var t = R.computeTotals(state, p.from, p.to, vatRate());
    R.INSTITUTIONS.forEach(function (inst) {
      $('#f-' + inst.id + '-qty').textContent = t.institutions[inst.id].qty.toLocaleString('he-IL');
      $('#f-' + inst.id + '-total').textContent = money.format(t.institutions[inst.id].total);
    });
    $('#f-all-total').textContent = money.format(t.all.total);

    var cards = R.INSTITUTIONS.map(function (inst) {
      return summaryCard(inst.fullName, inst.id, t.institutions[inst.id]);
    });
    cards.push(summaryCard('סה"כ לתשלום', 'grand', t.all));
    $('#summary').innerHTML = cards.join('');
  }

  function summaryCard(title, cls, s) {
    return '<article class="kpi ' + cls + '">' +
      '<div class="kpi-label">' + title + '</div>' +
      '<div class="kpi-value">' + money.format(s.incl) + '</div>' +
      '<div class="kpi-caption">כולל מע"מ</div><dl>' +
      '<dt>סה"כ מנות</dt><dd>' + s.qty.toLocaleString('he-IL') + '</dd>' +
      '<dt>ימי חול</dt><dd>' + money.format(s.regular) + '</dd>' +
      '<dt>ימי שבת וחג</dt><dd>' + money.format(s.special) + '</dd>' +
      '<dt>ללא מע"מ</dt><dd>' + money.format(s.total) + '</dd>' +
      '<dt>מע"מ (' + vatRate() + '%)</dt><dd>' + money.format(s.vat) + '</dd>' +
      '</dl></article>';
  }

  /* ---------- table events ---------- */
  tbody.addEventListener('input', function (ev) {
    var input = ev.target;
    if (input.tagName !== 'INPUT') return;
    var tr = input.closest('tr');
    var line = ensureLine(tr.dataset.iso, Number(tr.dataset.idx));
    var field = input.dataset.field;
    if (field === 'note') {
      line.note = input.value;
      tr.classList.toggle('has-note', !!input.value.trim());
    } else {
      var e = line[input.dataset.inst] || (line[input.dataset.inst] = {});
      if (input.value === '') delete e[field]; else e[field] = Number(input.value);
      if (field === 'price') input.classList.toggle('custom', input.value !== '');
      updateRow(tr);
      updateTotals();
    }
    saveState();
  });

  // An emptied price goes back to the default price.
  tbody.addEventListener('change', function (ev) {
    var input = ev.target;
    if (input.dataset.field !== 'price' || input.value !== '') return;
    var tr = input.closest('tr');
    var line = R.linesOf(state, tr.dataset.iso)[Number(tr.dataset.idx)] || {};
    input.value = R.lineEntry(state, line, input.dataset.inst).price;
    input.classList.remove('custom');
  });

  tbody.addEventListener('click', function (ev) {
    var btn = ev.target.closest('button');
    if (!btn) return;
    var tr = btn.closest('tr');
    var iso = tr.dataset.iso, idx = Number(tr.dataset.idx);

    if (btn.classList.contains('special-toggle')) {
      var d = ensureDay(iso);
      var next = !R.isSpecial(state, iso);
      if (next === (R.weekday(iso) === 6)) delete d.special; else d.special = next;
      setToggle(btn, next);
      var group = tbody.querySelectorAll('tr[data-iso="' + iso + '"]');
      Array.prototype.forEach.call(group, function (row) { row.classList.toggle('special', next); });
      updateTotals();
      saveState();
    } else if (btn.classList.contains('add')) {
      var day = ensureDay(iso);
      while (day.lines.length < 1) day.lines.push({});
      day.lines.push({});
      saveState();
      drawRows();
      focusRow(iso, day.lines.length - 1);
      updateTotals();
    } else if (btn.classList.contains('remove')) {
      state.days[iso].lines.splice(idx, 1);
      saveState();
      drawRows();
      updateTotals();
    }
  });

  function focusRow(iso, idx) {
    var row = tbody.querySelector('tr[data-iso="' + iso + '"][data-idx="' + idx + '"]');
    if (row) row.querySelector('input.qty').focus();
  }

  // Enter moves to the same field in the next row.
  tbody.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Enter' || ev.target.tagName !== 'INPUT') return;
    ev.preventDefault();
    var tr = ev.target.closest('tr');
    var idx = Array.prototype.indexOf.call(tr.querySelectorAll('input'), ev.target);
    var next = tr.nextElementSibling;
    if (next) {
      var target = next.querySelectorAll('input')[idx];
      // On a phone, a collapsed empty day shows only its quantity fields.
      if (!target.offsetParent) target = next.querySelector('input.qty');
      target.focus();
      if (target.select) target.select();
    }
  });

  // Prevent the mouse wheel from silently changing a focused number field.
  tbody.addEventListener('wheel', function (ev) {
    if (ev.target.type === 'number' && document.activeElement === ev.target) ev.target.blur();
  }, { passive: true });

  /* ---------- settings ---------- */
  document.querySelectorAll('.setting[data-inst]').forEach(function (box) {
    var input = box.querySelector('.bulk-price');
    box.querySelector('.apply-price').addEventListener('click', function () {
      var inst = box.dataset.inst;
      var price = input.value === '' ? R.DEFAULT_PRICE : Number(input.value);
      var custom = [];
      Object.keys(state.days).forEach(function (iso) {
        (state.days[iso].lines || []).forEach(function (line) {
          if (line[inst] && line[inst].price !== undefined) custom.push(line[inst]);
        });
      });
      state.prices[inst] = price;
      // Lines with their own price are replaced only if the user agrees.
      if (custom.length && confirm('יש ' + custom.length + ' שורות עם מחיר שונה. להחליף גם אותן?\n' +
          'אישור = להחליף את כולן, ביטול = להשאיר אותן ולשנות רק את שאר השורות.')) {
        custom.forEach(function (e) { delete e.price; });
      }
      saveState();
      drawRows();
      updateTotals();
      toast('המחיר ' + price + ' ₪ הוגדר לכל החודש');
    });
  });

  vatInput.addEventListener('input', function () {
    settings.vat = vatRate();
    save('meals:settings', settings);
    updateTotals();
  });

  function setRange(from, to) {
    if (!from || !to || from > to) { toast('טווח התאריכים אינו תקין'); return false; }
    if (Math.round((R.parseIso(to) - R.parseIso(from)) / 864e5) + 1 > R.MAX_DAYS) {
      toast('אפשר לדווח על עד ' + R.MAX_DAYS + ' ימים'); return false;
    }
    var def = R.monthRange(year, month);
    state.from = from === def.from ? null : from;
    state.to = to === def.to ? null : to;
    saveState();
    drawRows();
    updateTotals();
    return true;
  }
  function onRangeChange() {
    if (!setRange(fromInput.value, toInput.value)) {
      var p = period();
      fromInput.value = p.from;
      toInput.value = p.to;
    }
  }
  fromInput.addEventListener('change', onRangeChange);
  toInput.addEventListener('change', onRangeChange);
  $('#reset-range').addEventListener('click', function () {
    var def = R.monthRange(year, month);
    fromInput.value = def.from;
    toInput.value = def.to;
    setRange(def.from, def.to);
  });

  fileInput.addEventListener('input', function () {
    if (fileInput.value.trim()) state.fileName = fileInput.value; else delete state.fileName;
    saveState();
  });
  fileInput.addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter') { ev.preventDefault(); $('#export').click(); }
  });

  /* ---------- month navigation ---------- */
  monthSel.addEventListener('change', render);
  yearSel.addEventListener('change', render);

  function shiftMonth(delta) {
    var m = month + delta, yr = year;
    if (m < 0) { m = 11; yr--; }
    if (m > 11) { m = 0; yr++; }
    if (!yearSel.querySelector('option[value="' + yr + '"]')) {
      var opt = new Option(yr, yr);
      if (delta < 0) yearSel.insertBefore(opt, yearSel.firstChild); else yearSel.add(opt);
    }
    monthSel.value = m;
    yearSel.value = yr;
    render();
  }
  $('#prev').addEventListener('click', function () { shiftMonth(-1); });
  $('#next').addEventListener('click', function () { shiftMonth(1); });

  /* ---------- phone bottom bar ---------- */
  $('#jump-today').addEventListener('click', function () {
    var row = tbody.querySelector('tr[data-iso="' + R.toIso(new Date()) + '"]');
    if (!row) { toast('היום אינו בטווח התאריכים המוצג'); return; }
    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
    row.classList.add('flash');
    setTimeout(function () { row.classList.remove('flash'); }, 1600);
  });
  $('#export-mobile').addEventListener('click', function () { $('#export').click(); });

  /* ---------- clear & export ---------- */
  $('#clear').addEventListener('click', function () {
    var label = R.MONTHS[month] + ' ' + year;
    if (!confirm('למחוק את כל הנתונים של ' + label + '?')) return;
    state = { from: null, to: null, prices: {}, days: {} };
    saveState();
    render();
    toast('נתוני ' + label + ' נמחקו');
  });

  $('#export').addEventListener('click', function () {
    var p = period();
    if (!R.hasData(state, p.from, p.to)) {
      toast('לא הוזנו מנות לחודש ' + R.MONTHS[month] + ' ' + year);
      return;
    }
    var wb = R.buildWorkbook(window.ExcelJS, state, year, month, vatRate());
    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      var fileName = R.reportFileName(year, month, fileInput.value);
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
      // FSI/PDI keep a Hebrew name with ".xlsx" in the right order inside the Hebrew message.
      toast('הדוח \u2068' + fileName + '\u2069 הורד בהצלחה');
    });
  });

  /* ---------- table width ---------- */
  // The daily table grows with the page; if it is wider than its panel, scroll it sideways.
  var scroller = $('.table-scroll'), grid = $('#grid');
  function fitTable() {
    var wide = grid.offsetWidth > scroller.clientWidth + 1;
    if (wide !== scroller.classList.contains('scroll-x')) scroller.classList.toggle('scroll-x', wide);
  }
  if (window.ResizeObserver) {
    var ro = new ResizeObserver(fitTable);
    ro.observe(grid);
    ro.observe(scroller);
  } else {
    window.addEventListener('resize', fitTable);
  }

  render();
  fitTable();
})();
