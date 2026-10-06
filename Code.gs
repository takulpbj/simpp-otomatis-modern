/**
 * SIMPP Otomatis - Web App (Apps Script STANDALONE)
 * Input : SIMPP!F2, G2:G101, J2:J101, K2:K50, L2:L50, M2:M50 (opsional), P4
 * Output: SIMPP!A2:E213 + P2:P7 (dihitung formula, hanya dibaca)
 *
 * SETUP:
 *  1) isi SHEET_ID
 *  2) Buat sheet "Whitelist_User" dengan email di kolom A
 *  3) Project Settings > Script Properties: PIN = kode akses (opsional)
 *  4) Deploy > New deployment > Web app > Execute as: Me
 */

const SHEET_ID = '1NyxQWxpWcNtxxJ7zVPCaA0nc1bgNHz7228uQUrhwiIE';
const S = 'SIMPP';
const WHITELIST_SHEET = 'Whitelist_User';
const MAX_EVT = 100;
const MAX_ACT = 49;

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIMPP Otomatis')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sh_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(S);
}

function whitelistSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(WHITELIST_SHEET);
  if (!sheet) throw new Error('Sheet "Whitelist_User" belum dibuat.');
  return sheet;
}

function fmt_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd') : '';
}

function checkPin_(pin) {
  const p = PropertiesService.getScriptProperties().getProperty('PIN');
  if (p && String(pin) !== p) throw new Error('PIN salah.');
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isWhitelisted_(email) {
  const clean = normalizeEmail_(email);
  if (!clean) return false;

  const sheet = whitelistSheet_();
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  return values.some(row => normalizeEmail_(row[0]) === clean);
}

function verifyUserAccess(email) {
  const clean = normalizeEmail_(email);
  if (!clean) {
    return { success: false, message: 'Email wajib diisi.' };
  }

  if (!isWhitelisted_(clean)) {
    return { success: false, message: 'Email Anda tidak terdaftar pada sheet Whitelist_User kolom A.' };
  }

  return { success: true, email: clean };
}

function getData() {
  const userEmail = Session.getActiveUser().getEmail() || '';
  if (!userEmail) {
    throw new Error('Tidak dapat mendeteksi email pengguna. Pastikan aplikasi diakses dengan akun yang terdaftar.');
  }

  if (!isWhitelisted_(userEmail)) {
    throw new Error('Akses ditolak. Email Anda tidak ada di sheet Whitelist_User kolom A.');
  }

  const s = sh_();
  const month = s.getRange('F2').getValue();
  const limit = s.getRange('P4').getValue();

  const ev = s.getRange(2, 7, MAX_EVT, 4).getValues()
    .map((r, i) => ({ row: i + 2, date: fmt_(r[0]), note: r[3] }))
    .filter(x => x.date || x.note);

  const acts = s.getRange(2, 11, MAX_ACT, 3).getValues()
    .map((r, i) => ({ row: i + 2, dur: r[0], name: r[1], type: r[2] || 'Rutin' }))
    .filter(x => x.name || x.dur);

  const res = s.getRange(2, 1, 212, 5).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ date: fmt_(r[0]), act: r[1], vol: r[2], dur: r[3], tot: r[4] }));

  const d = s.getRange('P2:P7').getValues().map(r => r[0]);

  return {
    month: fmt_(month),
    limit: limit,
    events: ev,
    acts: acts,
    result: res,
    dash: {
      menit: d[0],
      jam: d[1],
      batas: d[2],
      sisa: d[3],
      persen: d[4],
      status: d[5]
    },
    email: userEmail
  };
}

function saveData(payload) {
  const userEmail = Session.getActiveUser().getEmail() || '';
  if (!userEmail || !isWhitelisted_(userEmail)) {
    throw new Error('Akses ditolak. Email Anda tidak ada di sheet Whitelist_User kolom A.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Server sibuk, coba lagi sebentar.');

  try {
    const s = sh_();
    const p = payload || {};
    if (p.events && p.events.length > MAX_EVT) throw new Error('Maksimal ' + MAX_EVT + ' kegiatan tambahan.');
    if (p.acts && p.acts.length > MAX_ACT) throw new Error('Maksimal ' + MAX_ACT + ' daftar kegiatan.');

    const lim = Number(p.limit);
    if (!(lim > 0)) throw new Error('Batas jam harus > 0.');

    const rutin = (p.acts || []).filter(a => !a.type || a.type === 'Rutin').length;
    if (rutin < 4) throw new Error('Minimal 4 kegiatan berjenis Rutin (syarat rotasi).');

    if ((p.acts || []).some(a => !(Number(a.dur) > 0) || !String(a.name).trim())) {
      throw new Error('Setiap kegiatan wajib punya nama dan durasi > 0.');
    }

    const d = v => {
      const [y, m, dd] = v.split('-').map(Number);
      return new Date(y, m - 1, dd);
    };

    s.getRange('F2').setValue(d(p.month));
    s.getRange('P4').setValue(lim);

    s.getRange(2, 7, MAX_EVT, 1).clearContent();
    s.getRange(2, 10, MAX_EVT, 1).clearContent();
    if ((p.events || []).length) {
      s.getRange(2, 7, p.events.length, 1).setValues(p.events.map(e => [d(e.date)]));
      s.getRange(2, 10, p.events.length, 1).setValues(p.events.map(e => [e.note]));
    }

    s.getRange(2, 11, MAX_ACT, 3).clearContent();
    if ((p.acts || []).length) {
      s.getRange(2, 11, p.acts.length, 3).setValues(
        (p.acts || []).map(a => [Number(a.dur), a.name.trim(), a.type === 'Rutin' ? '' : a.type])
      );
    }

    SpreadsheetApp.flush();
    log_(p.user || userEmail, 'Simpan', p.month + ' | ' + (p.events || []).length + ' tambahan, ' + (p.acts || []).length + ' kegiatan, batas ' + lim + ' jam');
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function log_(user, aksi, info) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const l = ss.getSheetByName('Log_Input') || ss.insertSheet('Log_Input');
  l.appendRow([new Date(), user || '(tanpa nama)', Session.getActiveUser().getEmail() || '-', aksi, info]);
}

function exportXlsx(name) {
  const userEmail = Session.getActiveUser().getEmail() || '';
  if (!userEmail || !isWhitelisted_(userEmail)) {
    throw new Error('Akses ditolak. Email Anda tidak ada di sheet Whitelist_User kolom A.');
  }

  const d = getData();
  const tz = 'Asia/Jakarta';
  const bulan = Utilities.formatDate(new Date(d.month + 'T00:00:00'), tz, 'yyyy-MM');
  const tmp = SpreadsheetApp.create('SIMPP_' + bulan + '_tmp');

  try {
    const h = tmp.getSheets()[0].setName('Hasil');
    const rows = d.result.map(r => [new Date(r.date + 'T00:00:00'), r.act, r.vol, r.dur, r.tot]);

    h.getRange(1, 1, 1, 5).setValues([['Tanggal', 'Kegiatan Tugas Jabatan', 'Volume', 'Durasi (menit)', 'Total (menit)']])
      .setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');

    if (rows.length) {
      h.getRange(2, 1, rows.length, 5).setValues(rows);
      h.getRange(2, 1, rows.length, 1).setNumberFormat('dd/mm/yyyy');
      h.getRange(rows.length + 2, 4, 1, 2).setValues([['TOTAL', '=SUM(E2:E' + (rows.length + 1) + ')']]).setFontWeight('bold');
    }

    h.setColumnWidth(1, 95).setColumnWidth(2, 520).setColumnWidths(3, 3, 95);
    h.getRange(2, 2, Math.max(rows.length, 1), 1).setWrap(true);
    h.setFrozenRows(1);

    const k = d.dash;
    const r = tmp.insertSheet('Ringkasan');
    r.getRange(1, 1, 8, 2).setValues([
      ['Bulan', Utilities.formatDate(new Date(d.month + 'T00:00:00'), tz, 'MMMM yyyy')],
      ['Total menit', k.menit],
      ['Total jam', k.jam],
      ['Batas (jam)', k.batas],
      ['Sisa (jam)', k.sisa],
      ['Persentase', k.persen],
      ['Status', k.status],
      ['Diekspor oleh', (name || userEmail || '-') + ' · ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')]
    ]);

    r.getRange('B6').setNumberFormat('0.0%');
    r.getRange('A1:A8').setFontWeight('bold');
    r.setColumnWidth(1, 130).setColumnWidth(2, 240);
    SpreadsheetApp.flush();

    const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      throw new Error('Gagal membuat file export (' + res.getResponseCode() + ').');
    }

    log_(name || userEmail, 'Export XLSX', bulan);
    return { name: 'SIMPP_' + bulan + '.xlsx', b64: Utilities.base64Encode(res.getContent()) };
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}
