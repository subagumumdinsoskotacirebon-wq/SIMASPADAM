/**
 * SAPRAS RESCUE — Backend Google Apps Script (versi 2)
 *
 * Menu: Dashboard | Kondisi Unit | Kondisi Peralatan Rescue | Download Laporan
 *
 *  - Kondisi Unit            : dibaca dari spreadsheet lama (01_KONDISI_UNIT, 02_PERLENGKAPAN, MASTER_UNIT)
 *  - Kondisi Peralatan Rescue: dibaca & DITULIS ke spreadsheet "KONDISI UNIT" (Sheet1, header di baris 3)
 */

// ====== KONFIGURASI ======
const ID_UNIT   = '1hnN7TzFu6HOZiy0areo-iWiTwKZxsmEwgh5fQpCu1LI'; // spreadsheet kondisi unit (kemarin)
const ID_RESCUE = '1Y0M2IWwPkHwypDrVRY1mAs8KFMxpsLMnBCAQrFIOdRg'; // spreadsheet peralatan rescue (baru)

const SH_UNIT   = '01_KONDISI_UNIT';
const SH_PERL   = '02_PERLENGKAPAN';
const SH_MASTER = 'MASTER_UNIT';
const SH_RESCUE = 'Sheet1';

const RESCUE_FIRST_ROW = 4;   // data mulai baris 4 (baris 3 = judul kolom)
const RESCUE_ITEMS_N   = 4;   // Floating Pump, Rebar Cutter, Cutter, Hydraulic Spreader
const COL_TGL          = 20;  // kolom T = "Tanggal Input" (ditambahkan otomatis)

const PILIHAN_FISIK  = ['Baik', 'Rusak Ringan', 'Rusak Berat'];
const PILIHAN_FUNGSI = ['Berfungsi', 'Tidak Berfungsi'];

// ====== WEB APP ======
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sapras Rescue')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil dari Index.html */
function getData() {
  const out = { updated: '', unit: null, rescue: null, errors: [] };
  out.updated = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
  try { out.unit = readUnit_(); }
  catch (e) { out.errors.push('Data Kondisi Unit tidak bisa dibaca: ' + e.message); }
  try { out.rescue = readRescue_(); }
  catch (e) { out.errors.push('Data Peralatan Rescue tidak bisa dibaca: ' + e.message); }
  return out;
}

/** Simpan isian baru ke spreadsheet peralatan rescue */
function saveRescue(f) {
  // ---- validasi ----
  const pel = Number(f && f.peleton), reg = Number(f && f.regu);
  if (!(pel >= 1 && pel <= 6)) throw new Error('Peleton harus 1 sampai 6.');
  if (!(reg >= 1 && reg <= 6)) throw new Error('Regu harus 1 sampai 6.');
  if (!f.items || f.items.length !== RESCUE_ITEMS_N) throw new Error('Data alat tidak lengkap.');

  const vals = [pel, reg];
  f.items.forEach(function (it, i) {
    const jml = Number(it.jumlah);
    if (it.jumlah === '' || it.jumlah === null || isNaN(jml) || jml < 0 || Math.floor(jml) !== jml)
      throw new Error('Jumlah alat ke-' + (i + 1) + ' harus angka bulat 0 atau lebih.');
    var fisik = '-', fungsi = '-';
    if (jml > 0) {
      if (PILIHAN_FISIK.indexOf(it.fisik) < 0)   throw new Error('Pilih kondisi fisik untuk alat ke-' + (i + 1) + '.');
      if (PILIHAN_FUNGSI.indexOf(it.fungsi) < 0) throw new Error('Pilih kondisi fungsi untuk alat ke-' + (i + 1) + '.');
      fisik = it.fisik; fungsi = it.fungsi;
    }
    vals.push(jml, fisik, fungsi, String(it.ket || '').substring(0, 300));
  });

  // ---- tulis (dikunci agar 2 orang tidak menimpa baris yang sama) ----
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(ID_RESCUE);
    const sh = ss.getSheetByName(SH_RESCUE) || ss.getSheets()[0];

    if (sh.getMaxColumns() < COL_TGL) sh.insertColumnsAfter(sh.getMaxColumns(), COL_TGL - sh.getMaxColumns());
    if (!sh.getRange(3, COL_TGL).getValue()) sh.getRange(3, COL_TGL).setValue('Tanggal Input');

    const last = Math.max(sh.getLastRow(), RESCUE_FIRST_ROW);
    const cur = sh.getRange(RESCUE_FIRST_ROW, 1, last - RESCUE_FIRST_ROW + 1, 19).getValues();
    var target = -1;
    for (var i = 0; i < cur.length; i++) {
      if (rowBlank_(cur[i], 1, 18)) { target = RESCUE_FIRST_ROW + i; break; }
    }
    if (target < 0) {                       // semua baris terisi → tambah baris baru
      target = last + 1;
      if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
      sh.getRange(target, 1).setValue(target - RESCUE_FIRST_ROW + 1);
    }
    sh.getRange(target, 2, 1, vals.length).setValues([vals]);
    sh.getRange(target, COL_TGL).setValue(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'));
    SpreadsheetApp.flush();
    return { ok: true, row: target };
  } finally {
    lock.releaseLock();
  }
}

/** Ubah HTML laporan menjadi PDF (dikirim balik sebagai base64) */
function makePdf(html) {
  const blob = Utilities.newBlob(String(html), 'text/html', 'laporan.html').getAs('application/pdf');
  return Utilities.base64Encode(blob.getBytes());
}

// ====== PEMBACA DATA ======
function readUnit_() {
  const ss = SpreadsheetApp.openById(ID_UNIT);
  const tz = 'Asia/Jakarta';
  return {
    units: readMasterUnit_(ss),
    unit:  readWide_(ss, SH_UNIT, tz),
    perl:  readWide_(ss, SH_PERL, tz)
  };
}

function readMasterUnit_(ss) {
  const sh = ss.getSheetByName(SH_MASTER);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return r[0] && String(r[2]).toUpperCase() !== 'TIDAK'; })
    .map(function (r) { return { kode: String(r[0]), nama: String(r[1] || r[0]) }; });
}

function readWide_(ss, name, tz) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const vals = sh.getRange(1, 1, lr, lc).getValues();
  const heads = vals[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (var i = 1; i < vals.length; i++) {
    if (rowBlank_(vals[i], 0, vals[i].length - 1)) continue;
    const o = {};
    for (var j = 0; j < heads.length; j++) {
      if (!heads[j]) continue;
      var v = vals[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
      o[heads[j]] = v;
    }
    rows.push(o);
  }
  return rows;
}

/** Baca sheet peralatan rescue berdasarkan URUTAN KOLOM (A=No, B=Peleton, C=Regu, D..S = 4 alat x 4 kolom, T=Tanggal) */
function readRescue_() {
  const ss = SpreadsheetApp.openById(ID_RESCUE);
  const sh = ss.getSheetByName(SH_RESCUE) || ss.getSheets()[0];
  const last = sh.getLastRow();
  if (last < RESCUE_FIRST_ROW) return [];
  const nCol = Math.min(COL_TGL, sh.getMaxColumns());
  const vals = sh.getRange(RESCUE_FIRST_ROW, 1, last - RESCUE_FIRST_ROW + 1, nCol).getValues();
  const rows = [];
  vals.forEach(function (v, idx) {
    if (rowBlank_(v, 1, 18)) return;
    const items = [];
    for (var k = 0; k < RESCUE_ITEMS_N; k++) {
      const b = 3 + k * 4;
      items.push({ jumlah: v[b], fisik: v[b + 1], fungsi: v[b + 2], ket: v[b + 3] });
    }
    var tgl = v[COL_TGL - 1];
    if (tgl instanceof Date) tgl = Utilities.formatDate(tgl, 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss");
    rows.push({ row: RESCUE_FIRST_ROW + idx, no: v[0], peleton: v[1], regu: v[2], tgl: tgl || '', items: items });
  });
  return rows;
}

function rowBlank_(arr, from, to) {
  for (var i = from; i <= to && i < arr.length; i++) {
    if (arr[i] !== '' && arr[i] !== null) return false;
  }
  return true;
}

/*
 * ===================== CARA PASANG =====================
 * 1. Buka script.google.com → Proyek baru (atau Ekstensi → Apps Script dari salah satu Sheet).
 * 2. Tempel isi Code.gs ini. Klik "+" → HTML → beri nama  Index  → tempel isi Index.html.
 * 3. Jalankan fungsi getData sekali, setujui izin akses (harus akun yang BISA MENGEDIT kedua spreadsheet).
 * 4. Deploy → Deployment baru → Aplikasi web:
 *      Jalankan sebagai: Saya   |   Yang memiliki akses: sesuai kebutuhan
 * 5. Salin URL web app.
 * Setelah mengubah kode: Deploy → Kelola deployment → Edit → Versi baru.
 */
