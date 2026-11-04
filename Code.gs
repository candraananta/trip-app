/**
 * TRIP LEDGER: catat rencana & pengeluaran perjalanan (Google Apps Script)
 *
 * SETUP SEKALI:
 * 1. (Opsional) ubah CONFIG di bawah.
 * 2. Pilih fungsi setupDatabase di dropdown atas, klik Run, lalu Review permissions > Allow.
 * 3. Buka Execution log. Di sana ada link Spreadsheet database.
 * 4. Deploy > New deployment > Select type: Web app
 *    Execute as: Me | Who has access: Anyone with Google account > Deploy.
 * 5. Buka URL web app di Safari iPhone > Share > Add to Home Screen.
 *
 * File yang dibutuhkan: Code.gs (file ini) dan Index.html
 * (di editor: klik + di samping Files > HTML > beri nama "Index").
 */

const CONFIG = {
  TRIP_NAME: 'Trip KL 2026',
  START_DATE: '2027-03-12',                   // format yyyy-mm-dd, contoh '2026-12-20'. Bisa diubah juga di menu Data.
  END_DATE: '2027-03-17',
  KURS: { MYR: 4454, USD: 18143 },  // 1 unit mata uang asing = sekian IDR. Contoh, sesuaikan.
  PAYMENTS: ['Tunai', 'Kartu', 'QRIS', 'E-Wallet', 'Transfer'],
};

const TZ = 'Asia/Jakarta';
const DB_KEY = 'TRIP_DB_ID';
const TABLES = {
  destinasi: { sheet: 'Master_Destinasi', cols: ['id', 'nama', 'lokasi', 'catatan'] },
  kategori: { sheet: 'Master_Kategori', cols: ['id', 'nama', 'ikon', 'budget_idr'] },
  rencana: { sheet: 'Rencana', cols: ['id', 'tanggal', 'jam', 'destinasi_id', 'kategori_id', 'kegiatan', 'estimasi_idr', 'status', 'catatan'] },
  realisasi: { sheet: 'Realisasi', cols: ['id', 'tanggal', 'destinasi_id', 'kategori_id', 'keterangan', 'mata_uang', 'nominal', 'kurs', 'nominal_idr', 'pembayaran', 'rencana_id'] },
};
const SEED = {
  destinasi: [
    { id: 'DS1', nama: 'KLCC & Petronas', lokasi: 'Kuala Lumpur City Centre' },
    { id: 'DS2', nama: 'Batu Caves', lokasi: 'Gombak' },
    { id: 'DS3', nama: 'Bukit Bintang', lokasi: 'Kuala Lumpur' },
  ],
  kategori: [
    { id: 'KT1', nama: 'Transportasi', ikon: '🚆', budget_idr: 3000000 },
    { id: 'KT2', nama: 'Akomodasi', ikon: '🏨', budget_idr: 6000000 },
    { id: 'KT3', nama: 'Makan & Minum', ikon: '🍜', budget_idr: 4000000 },
    { id: 'KT4', nama: 'Tiket & Wisata', ikon: '🎟️', budget_idr: 2000000 },
    { id: 'KT5', nama: 'Belanja', ikon: '🛍️', budget_idr: 3000000 },
    { id: 'KT6', nama: 'Lainnya', ikon: '📦', budget_idr: 1000000 },
  ],
};

/* ---------- WEB APP ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.TRIP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------- SETUP (jalankan sekali dari editor, aman dijalankan ulang) ---------- */
function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.openById(props.getProperty(DB_KEY)); } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.TRIP_NAME + ' - Database');
    props.setProperty(DB_KEY, ss.getId());
  }
  ss.setSpreadsheetTimeZone(TZ);

  Object.values(TABLES).forEach(t => {
    const sh = getOrCreate_(ss, t.sheet);
    if (sh.getLastRow() === 0) {
      sh.appendRow(t.cols);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, t.cols.length).setFontWeight('bold');
    }
  });

  const st = getOrCreate_(ss, 'Settings');
  if (st.getLastRow() === 0) {
    writeSettings_(st, { trip: CONFIG.TRIP_NAME, mulai: CONFIG.START_DATE, selesai: CONFIG.END_DATE, kurs: CONFIG.KURS });
  }

  Object.entries(SEED).forEach(([key, list]) => {
    const t = TABLES[key];
    const sh = ss.getSheetByName(t.sheet);
    if (sh.getLastRow() === 1) {
      sh.getRange(2, 1, list.length, t.cols.length).setValues(list.map(o => t.cols.map(c => o[c] ?? '')));
    }
  });

  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  Logger.log('Database siap: ' + ss.getUrl());
  return ss.getUrl();
}

/* ---------- DATA (dipanggil dari halaman) ---------- */
function getData() {
  const exp = readAll_('realisasi');
  const plans = readAll_('rencana');
  plans.forEach(p => {
    p.aktual = exp.filter(e => e.rencana_id === p.id).reduce((t, e) => t + Number(e.nominal_idr || 0), 0);
  });
  return {
    settings: getSettings_(),
    destinasi: readAll_('destinasi'),
    kategori: readAll_('kategori'),
    rencana: plans,
    realisasi: exp,
    payments: CONFIG.PAYMENTS,
  };
}

function saveRecord(key, d) {
  return withLock_(() => {
    const rec = { ...d };
    if (key === 'realisasi') {
      rec.nominal = Number(d.nominal);
      if (!(rec.nominal > 0)) throw new Error('Nominal harus lebih dari 0.');
      rec.mata_uang = d.mata_uang || 'IDR';
      rec.kurs = Number(d.kurs) || getSettings_().kurs[rec.mata_uang] || 1;
      rec.nominal_idr = Math.round(rec.nominal * rec.kurs);
    } else if (key === 'rencana') {
      if (!d.kegiatan) throw new Error('Nama kegiatan wajib diisi.');
      rec.estimasi_idr = Number(d.estimasi_idr) || 0;
      rec.status = d.status || 'Rencana';
    } else {
      if (!d.nama) throw new Error('Nama wajib diisi.');
      if (key === 'kategori') rec.budget_idr = Number(d.budget_idr) || 0;
    }
    return upsert_(key, rec);
  });
}

function removeRecord(key, id) {
  return withLock_(() => {
    const sh = db_().getSheetByName(TABLES[key].sheet);
    const i = sh.getDataRange().getValues().findIndex((r, n) => n > 0 && r[0] === id);
    if (i > 0) sh.deleteRow(i + 1);
    return true;
  });
}

function saveSettings(s) {
  return withLock_(() => {
    writeSettings_(db_().getSheetByName('Settings'), s);
    return true;
  });
}

/* ---------- HELPER (nama diakhiri _ supaya tidak bisa dipanggil dari browser) ---------- */
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(DB_KEY);
  if (!id) throw new Error('Database belum dibuat. Jalankan setupDatabase() dulu.');
  return SpreadsheetApp.openById(id);
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function fmtVal_(v) {
  return v instanceof Date
    ? Utilities.formatDate(v, TZ, v.getFullYear() < 1950 ? 'HH:mm' : 'yyyy-MM-dd')
    : v;
}

function readAll_(key) {
  const [head, ...rows] = db_().getSheetByName(TABLES[key].sheet).getDataRange().getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(head.map((h, i) => [h, fmtVal_(r[i])])));
}

function getSettings_() {
  const s = {};
  db_().getSheetByName('Settings').getDataRange().getValues().forEach(([k, v]) => { s[k] = fmtVal_(v); });
  return {
    trip: s.TRIP_NAME || CONFIG.TRIP_NAME,
    mulai: s.START_DATE || '',
    selesai: s.END_DATE || '',
    kurs: { IDR: 1, MYR: Number(s.KURS_MYR) || 0, USD: Number(s.KURS_USD) || 0 },
  };
}

function writeSettings_(sh, s) {
  const v = x => x ?? '';
  sh.clear();
  sh.getRange(1, 1, 6, 2).setValues([
    ['key', 'value'],
    ['TRIP_NAME', v(s.trip)],
    ['START_DATE', v(s.mulai)],
    ['END_DATE', v(s.selesai)],
    ['KURS_MYR', v(s.kurs.MYR)],
    ['KURS_USD', v(s.kurs.USD)],
  ]);
}

function upsert_(key, rec) {
  const t = TABLES[key];
  const sh = db_().getSheetByName(t.sheet);
  rec.id = rec.id || uid_();
  const row = t.cols.map(c => rec[c] ?? '');
  const i = sh.getDataRange().getValues().findIndex((r, n) => n > 0 && r[0] === rec.id);
  if (i > 0) sh.getRange(i + 1, 1, 1, t.cols.length).setValues([row]);
  else sh.appendRow(row);
  return rec.id;
}

const uid_ = () => 'ID' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}