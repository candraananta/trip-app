/**
 * MODUL DOKUMEN PERJALANAN (Docs.gs)
 * Menyimpan KTP, paspor, tiket, bukti hotel, asuransi, dll.
 * File tersimpan di folder Google Drive pribadi "<nama trip> - Dokumen".
 * Butuh Code.gs di project yang sama (memakai db_, withLock_, uid_, fmtVal_, TZ, CONFIG).
 *
 * Jalankan setupDocs() sekali dari editor setelah setupDatabase() (memberi izin Drive).
 */
const DOCS = {
  sheet: 'Dokumen',
  cols: ['id', 'jenis', 'judul', 'pemilik', 'catatan', 'file_id', 'nama_file', 'mime', 'ukuran', 'dibuat'],
  maxBytes: 10 * 1024 * 1024,
};
const DOC_FOLDER_KEY = 'DOC_FOLDER_ID';

function setupDocs() {
  docSheet_();
  const url = docFolder_().getUrl();
  Logger.log('Modul dokumen siap: ' + url);
  return url;
}

/* ---------- DIPANGGIL DARI HALAMAN ---------- */
function listDocs() {
  const [head, ...rows] = docSheet_().getDataRange().getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => Object.fromEntries(head.map((h, i) => [h, fmtVal_(r[i])])));
}

function uploadDoc(meta, b64, mime, fileName) {
  return withLock_(() => {
    const bytes = Utilities.base64Decode(b64);
    if (bytes.length > DOCS.maxBytes) throw new Error('File lebih dari 10 MB.');
    const file = docFolder_().createFile(Utilities.newBlob(bytes, mime, fileName));
    const rec = {
      id: uid_(),
      jenis: meta.jenis || 'Lainnya',
      judul: meta.judul || fileName,
      pemilik: meta.pemilik || '',
      catatan: meta.catatan || '',
      file_id: file.getId(),
      nama_file: fileName,
      mime: mime,
      ukuran: bytes.length,
      dibuat: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    };
    docSheet_().appendRow(DOCS.cols.map(c => rec[c]));
    return rec.id;
  });
}

function saveDocMeta(d) {
  return withLock_(() => {
    const sh = docSheet_();
    const { i } = findDoc_(sh, d.id);
    if (i < 1) throw new Error('Dokumen tidak ditemukan.');
    ['jenis', 'judul', 'pemilik', 'catatan'].forEach(c => {
      sh.getRange(i + 1, DOCS.cols.indexOf(c) + 1).setValue(d[c] ?? '');
    });
    return true;
  });
}

function getDocFile(id) {
  const { row } = findDoc_(docSheet_(), id);
  if (!row) throw new Error('Dokumen tidak ditemukan.');
  const col = k => row[DOCS.cols.indexOf(k)];
  const blob = DriveApp.getFileById(col('file_id')).getBlob();
  return {
    name: col('nama_file'),
    mime: col('mime') || blob.getContentType(),
    b64: Utilities.base64Encode(blob.getBytes()),
  };
}

function removeDoc(id) {
  return withLock_(() => {
    const sh = docSheet_();
    const { i, row } = findDoc_(sh, id);
    if (i < 1) return true;
    try {
      DriveApp.getFileById(row[DOCS.cols.indexOf('file_id')]).setTrashed(true); // masuk Trash Drive
    } catch (e) { /* file sudah tidak ada, lanjutkan hapus baris */ }
    sh.deleteRow(i + 1);
    return true;
  });
}

/* ---------- HELPER ---------- */
function docSheet_() {
  const ss = db_();
  let sh = ss.getSheetByName(DOCS.sheet);
  if (!sh) {
    sh = ss.insertSheet(DOCS.sheet);
    sh.appendRow(DOCS.cols);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, DOCS.cols.length).setFontWeight('bold');
  }
  return sh;
}

function docFolder_() {
  const props = PropertiesService.getScriptProperties();
  try {
    const id = props.getProperty(DOC_FOLDER_KEY);
    if (id) return DriveApp.getFolderById(id);
  } catch (e) { /* folder hilang, buat baru */ }
  const folder = DriveApp.createFolder(CONFIG.TRIP_NAME + ' - Dokumen');
  props.setProperty(DOC_FOLDER_KEY, folder.getId());
  return folder;
}

function findDoc_(sh, id) {
  const vals = sh.getDataRange().getValues();
  const i = vals.findIndex((r, n) => n > 0 && r[0] === id);
  return { i: i, row: i > 0 ? vals[i] : null };
}