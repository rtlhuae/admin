// Sheet code that was live before 22 Nov 2026 (kept for rollback).
const H = {
  Orders: ['id','date','customer','product','color','type','text','price','notes'],
  Expenses: ['id','date','category','description','amount','paidBy','settled']
};

const GREEN = '#b7e1cd';

function sh(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);

  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, 1, H[n].length).setValues([H[n]]);
  }

  return s;
}

function isGreen(hex) {
  const m = /^#(..)(..)(..)$/.exec(hex || '');
  if (!m) return false;

  const c = m.slice(1).map(x => parseInt(x, 16));
  return c[1] > c[0] + 12 && c[1] > c[2] + 12;
}

function rows(n) {
  const s = sh(n);
  const rg = s.getDataRange();

  const v = rg.getValues();
  const bg = rg.getBackgrounds();

  if (v.length <= 1) return [];

  const h = v.shift();
  bg.shift();

  return v
    .map((r, j) => ({ r: r, b: bg[j] ? bg[j][0] : '' }))
    .filter(x => x.r[0] !== '')
    .map(x => {
      const o = {};

      h.forEach((k, i) => {
        let value = x.r[i];

        if (value instanceof Date) {
          value = Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          );
        }

        o[k] = value;
      });

      if (n === 'Orders') {
        o.done = isGreen(x.b);
      }

      return o;
    });
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    return json({
      ok: true,
      orders: rows('Orders'),
      expenses: rows('Expenses')
    });
  } catch (err) {
    return json({
      ok: false,
      error: String(err)
    });
  }
}

function nextRow(s) {
  const last = Math.max(s.getLastRow(), 1);
  return last + 1;
}

function findRow(s, id) {
  const last = s.getLastRow();

  if (last < 2) return 0;

  const v = s.getRange(2, 1, last - 1, 1).getValues();

  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      return i + 2;
    }
  }

  return 0;
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({
        ok: false,
        error: 'No POST body received'
      });
    }

    const d = JSON.parse(e.postData.contents);
    const a = d.action;

    if (a === 'fulfill' || a === 'reopen') {
      const s = sh('Orders');
      const r = findRow(s, d.id);

      if (r) {
        s.getRange(r, 1, 1, H.Orders.length)
          .setBackground(a === 'fulfill' ? GREEN : null);
      }

      return json({ ok: true });
    }

    if (a === 'settle' || a === 'unsettle') {
      const s = sh('Expenses');
      const r = findRow(s, d.id);

      if (r) {
        const col = H.Expenses.indexOf('settled') + 1;
        s.getRange(r, col).setValue(a === 'settle');
      }

      return json({ ok: true });
    }

    if (
      a === 'update_order' ||
      a === 'update_expense' ||
      a === 'delete_order' ||
      a === 'delete_expense'
    ) {
      const n = a.includes('expense') ? 'Expenses' : 'Orders';
      const s = sh(n);
      const r = findRow(s, d.id);

      if (!r) {
        return json({
          ok: false,
          error: 'ID not found'
        });
      }

      if (a.startsWith('delete')) {
        s.deleteRow(r);
      } else {
        const values = H[n].map(k =>
          d[k] === undefined ? '' : d[k]
        );

        s.getRange(r, 1, 1, H[n].length)
          .setValues([values]);
      }

      return json({ ok: true });
    }

    if (a === 'order' || a === 'expense') {
      const n = a === 'expense' ? 'Expenses' : 'Orders';
      const s = sh(n);

      d.id = Date.now();

      const values = H[n].map(k =>
        d[k] === undefined ? '' : d[k]
      );

      s.getRange(nextRow(s), 1, 1, H[n].length)
        .setValues([values]);

      return json({
        ok: true,
        id: d.id
      });
    }

    return json({
      ok: false,
      error: 'Unknown action: ' + a
    });

  } catch (err) {
    return json({
      ok: false,
      error: String(err)
    });
  }
}

function setup() {
  sh('Orders');
  sh('Expenses');
}