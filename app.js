(function () {
  'use strict';

  var R = window.MealReport;
  var $ = function (sel) { return document.querySelector(sel); };
  var money = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS' });

  var monthSel = $('#month'), yearSel = $('#year'), vatInput = $('#vat');
  var tbody = $('#rows'), noteInput = $('#general-note');
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
  function monthKey() { return 'meals:' + year + '-' + (month + 1); }
  function saveState() { save(monthKey(), state); }

  var settings = load('meals:settings', { vat: R.DEFAULT_VAT });

  /* ---------- period selectors ---------- */
  var now = new Date();
  R.MONTHS.forEach(function (name, i) { monthSel.add(new Option(name, i)); });
  for (var y = now.getFullYear() - 3; y <= now.getFullYear() + 2; y++) yearSel.add(new Option(y, y));
  var last = load('meals:lastPeriod', null);
  monthSel.value = last ? last.month : now.getMonth();
  yearSel.value = last ? last.year : now.getFullYear();
  vatInput.value = settings.vat;

  /* ---------- table ---------- */
  function dayData(day) {
    return state.days[day] || (state.days[day] = {});
  }

  function cell(content, cls) {
    var td = document.createElement('td');
    if (cls) td.className = cls;
    if (content instanceof Node) td.appendChild(content); else td.textContent = content;
    return td;
  }

  function numberInput(cls, value, attrs) {
    var input = document.createElement('input');
    input.type = 'number';
    input.className = cls;
    input.min = '0';
    input.step = attrs.step;
    input.inputMode = attrs.step === '1' ? 'numeric' : 'decimal';
    input.value = value;
    if (attrs.placeholder) input.placeholder = attrs.placeholder;
    return input;
  }

  function render() {
    year = Number(yearSel.value);
    month = Number(monthSel.value);
    state = load(monthKey(), null) || { days: {}, generalNote: '' };
    save('meals:lastPeriod', { year: year, month: month });
    noteInput.value = state.generalNote || '';

    tbody.innerHTML = '';
    var n = R.daysInMonth(year, month);
    for (var day = 1; day <= n; day++) {
      var wd = R.weekday(year, month, day);
      var tr = document.createElement('tr');
      tr.dataset.day = day;
      if (wd === 6) tr.className = 'shabbat';
      tr.appendChild(cell(String(day).padStart(2, '0') + '/' + String(month + 1).padStart(2, '0'), 'date'));
      tr.appendChild(cell(R.WEEKDAYS[wd], 'day-name'));
      R.INSTITUTIONS.forEach(function (inst, i) {
        var d = (state.days[day] || {})[inst.id] || {};
        var qty = numberInput('qty', d.qty || '', { step: '1', placeholder: '0' });
        qty.dataset.inst = inst.id;
        qty.dataset.field = 'qty';
        var price = numberInput('price', d.price !== undefined ? d.price : R.DEFAULT_PRICE, { step: '0.01' });
        price.dataset.inst = inst.id;
        price.dataset.field = 'price';
        tr.appendChild(cell(qty, i ? 'sep' : ''));
        tr.appendChild(cell(price));
        tr.appendChild(cell('', 'money total-' + inst.id));
      });
      tr.appendChild(cell('', 'money day-total sep'));
      var note = document.createElement('input');
      note.type = 'text';
      note.className = 'note';
      note.dataset.field = 'note';
      note.value = (state.days[day] || {}).note || '';
      tr.appendChild(cell(note));
      tbody.appendChild(tr);
      updateRow(tr);
    }
    updateTotals();
  }

  function updateRow(tr) {
    var day = Number(tr.dataset.day);
    var sum = 0, any = false;
    R.INSTITUTIONS.forEach(function (inst) {
      var e = R.entry(state, day, inst.id);
      tr.querySelector('.total-' + inst.id).textContent = e.qty ? money.format(e.total) : '';
      sum += e.total;
      if (e.qty) any = true;
    });
    tr.querySelector('.day-total').textContent = any ? money.format(sum) : '';
    tr.classList.toggle('has-data', any);
  }

  function updateTotals() {
    var t = R.computeTotals(state, year, month, vatInput.value);
    R.INSTITUTIONS.forEach(function (inst) {
      $('#f-' + inst.id + '-qty').textContent = t.institutions[inst.id].qty;
      $('#f-' + inst.id + '-total').textContent = money.format(t.institutions[inst.id].total);
    });
    $('#f-all-total').textContent = money.format(t.all.total);

    var vat = vatInput.value;
    var cards = R.INSTITUTIONS.map(function (inst) {
      return summaryCard(inst.name, inst.id, t.institutions[inst.id], vat);
    });
    cards.push(summaryCard('סה"כ שני המוסדות', '', t.all, vat));
    $('#summary').innerHTML = cards.join('');
  }

  function summaryCard(title, cls, s, vat) {
    return '<div class="sum-card ' + cls + '"><h3>' + title + '</h3><dl>' +
      '<dt>מספר מנות</dt><dd>' + s.qty + '</dd>' +
      '<dt>לפני מע"מ</dt><dd>' + money.format(s.beforeVat) + '</dd>' +
      '<dt>מע"מ (' + vat + '%)</dt><dd>' + money.format(s.vat) + '</dd>' +
      '<dt>סה"כ לתשלום</dt><dd class="grand">' + money.format(s.total) + '</dd>' +
      '</dl></div>';
  }

  /* ---------- events ---------- */
  tbody.addEventListener('input', function (ev) {
    var input = ev.target;
    var tr = input.closest('tr');
    var d = dayData(Number(tr.dataset.day));
    if (input.dataset.field === 'note') {
      d.note = input.value;
    } else {
      var e = d[input.dataset.inst] || (d[input.dataset.inst] = {});
      e[input.dataset.field] = input.value === '' ? '' : Number(input.value);
      updateRow(tr);
      updateTotals();
    }
    saveState();
  });

  // Restore the default price if a price field is left empty.
  tbody.addEventListener('change', function (ev) {
    var input = ev.target;
    if (input.dataset.field === 'price' && input.value === '') {
      input.value = R.DEFAULT_PRICE;
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
  });

  // Enter moves to the same field on the next day.
  tbody.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Enter') return;
    ev.preventDefault();
    var tr = ev.target.closest('tr');
    var idx = Array.prototype.indexOf.call(tr.querySelectorAll('input'), ev.target);
    var next = tr.nextElementSibling;
    if (next) {
      var target = next.querySelectorAll('input')[idx];
      target.focus();
      if (target.select) target.select();
    }
  });

  // Number inputs: prevent the mouse wheel from silently changing values.
  tbody.addEventListener('wheel', function (ev) {
    if (ev.target.type === 'number' && document.activeElement === ev.target) ev.target.blur();
  }, { passive: true });

  document.querySelectorAll('.setting[data-inst]').forEach(function (box) {
    var input = box.querySelector('.bulk-price');
    input.value = R.DEFAULT_PRICE;
    box.querySelector('.apply-price').addEventListener('click', function () {
      var inst = box.dataset.inst;
      var price = input.value === '' ? R.DEFAULT_PRICE : Number(input.value);
      for (var day = 1; day <= R.daysInMonth(year, month); day++) {
        var d = dayData(day);
        (d[inst] || (d[inst] = {})).price = price;
      }
      saveState();
      render();
    });
  });

  vatInput.addEventListener('input', function () {
    settings.vat = vatInput.value === '' ? R.DEFAULT_VAT : Number(vatInput.value);
    save('meals:settings', settings);
    updateTotals();
  });

  noteInput.addEventListener('input', function () {
    state.generalNote = noteInput.value;
    saveState();
  });

  monthSel.addEventListener('change', render);
  yearSel.addEventListener('change', render);

  $('#clear').addEventListener('click', function () {
    var label = R.MONTHS[month] + ' ' + year;
    if (!confirm('למחוק את כל הנתונים של ' + label + '?')) return;
    state = { days: {}, generalNote: '' };
    saveState();
    render();
  });

  $('#export').addEventListener('click', function () {
    var vat = vatInput.value === '' ? R.DEFAULT_VAT : Number(vatInput.value);
    if (!R.reportDays(state, year, month).length) {
      alert('לא הוזנו מנות לחודש ' + R.MONTHS[month] + ' ' + year + '.');
      return;
    }
    var wb = R.buildWorkbook(window.ExcelJS, state, year, month, vat);
    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = R.reportFileName(year, month);
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    });
  });

  render();
})();
