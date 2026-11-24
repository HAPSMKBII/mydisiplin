/**
 * MyDisiplin – pelayan (Google Apps Script)
 * Laporan kes disiplin murid SMK Bugaya II Semporna
 * Versi 1.0.2 · Oleh Cikgu Henry Able Padtong
 *
 * Serasi dengan AppSheet MYDISIPLIN26: menulis ke helaian LAPORAN yang sama,
 * ID 8 aksara, TARIKH sebagai tarikh, gambar dalam folder LAPORAN_Images aplikasi AppSheet.
 */

var KONFIG = {
  NAMA_SEKOLAH: 'SMK Bugaya II Semporna',
  NAMA_SISTEM: 'MyDisiplin',
  VERSI: '1.0.2',
  TEMPOH_SESI_SAAT: 21600,          // 6 jam
  MAKS_CUBAAN_GAGAL: 5,
  KUNCI_MINIT: 15,
  // Folder data aplikasi AppSheet dalam Google Drive anda (nama folder, bukan laluan)
  FOLDER_APP: 'MYDISIPLIN26-302733627-26-04-26-3',
  FOLDER_GAMBAR: 'LAPORAN_Images',
  FOLDER_PDF: 'DISIPLINREPORT2026',
  PAUTAN_APLIKASI: 'https://hapsmkbii.github.io/mydisiplin/',
  PEMISAH: ' , ',                    // pemisah senarai yang digunakan AppSheet
  ADMIN_AWAL: ['HENRY ABLE PADTONG'] // dijadikan ADMIN semasa "Isi pengguna"
};

var HELAIAN = {
  LAPORAN: 'LAPORAN', MURID: 'SENARAI MURID', KESALAHAN: 'SENARAI KESALAHAN', UNIT: 'UNIT',
  PELAPOR: 'DILAPORKAN OLEH', TINDAKAN: 'TINDAKAN', PUNCA: 'PUNCA SALAH LAKU',
  TEMPAT: 'TEMPAT KEJADIAN', GURUKELAS: 'GURU KELAS / PENOLONG',
  PENGGUNA: 'PENGGUNA', PADAM: 'LAPORAN DIPADAM'
};

var LAJUR_PENGGUNA = ['ID', 'NAMA', 'PERANAN', 'AKTIF', 'KATA LALUAN BARU', 'MESTI TUKAR', 'HASH', 'SALT', 'LOG MASUK TERAKHIR'];

// Medan klien → tajuk lajur LAPORAN
var PETA = {
  id: 'ID', kelas: 'KELAS', nama: 'NAMA MURID', guruKelas: 'GURU KELAS / PENOLONG', kategori: 'KATEGORI KES',
  perincian: 'PERINCIAN KES', punca: 'PUNCA SALAH LAKU', keterangan: 'KETERANGAN KES', tarikh: 'TARIKH',
  tempat: 'TEMPAT', masa: 'MASA', unit: 'UNIT GURU DISIPLIN', guruBertugas: 'GURU DISIPLIN BERTUGAS',
  pelapor: 'DILAPORKAN OLEH', tindakan: 'TINDAKAN', gambar1: 'GAMBAR 1', gambar2: 'GAMBAR 2',
  tandatangan: 'TANDATANGAN PELAPOR', pdf: 'PDF', pautan: 'PAUTAN URL', dokumen: 'DOKUMEN SOKONGAN'
};
var MEDAN_TEKS = ['kelas', 'nama', 'guruKelas', 'kategori', 'perincian', 'punca', 'keterangan', 'tempat', 'unit', 'guruBertugas', 'tindakan'];
var MEDAN_FAIL = { gambar1: 'GAMBAR 1', gambar2: 'GAMBAR 2', tandatangan: 'TANDATANGAN PELAPOR', dokumen: 'DOKUMEN SOKONGAN' };

/* ===================== Pintu masuk ===================== */

var API = {
  log_masuk: log_masuk, profil: profil, tukar_kata_laluan: tukar_kata_laluan, log_keluar: log_keluar,
  rujukan: rujukan, senarai: senarai, butiran: butiran, gambar: gambar,
  hantar: hantar, ubah: ubah, padam: padam, jana_pdf: jana_pdf, pdf_fail: pdf_fail
};

function doPost(e) {
  var keluar;
  try {
    var req = JSON.parse(e.postData.contents);
    if (!req || !Object.prototype.hasOwnProperty.call(API, req.fn)) throw new Error('Fungsi tidak dikenali.');
    keluar = { ok: true, data: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    keluar = { ok: false, ralat: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(keluar)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var u = KONFIG.PAUTAN_APLIKASI;
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Buka <a href="' + u + '">' + KONFIG.NAMA_SISTEM + '</a>…</p>' +
    '<script>window.top.location.replace(' + JSON.stringify(u) + ')</script>')
    .setTitle(KONFIG.NAMA_SISTEM).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== Utiliti helaian ===================== */

function ss_() { return SpreadsheetApp.getActive(); }
function zon_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Kuala_Lumpur'; }
function helaian_(nama) {
  var sh = ss_().getSheetByName(nama);
  if (!sh) throw new Error('Helaian "' + nama + '" tidak dijumpai.');
  return sh;
}
function tajuk_(sh) {
  var lc = sh.getLastColumn();
  if (lc < 1) return [];
  return sh.getRange(1, 1, 1, lc).getValues()[0].map(function (h) { return String(h).trim().toUpperCase(); });
}
/** Baca semua baris sebagai objek ikut nama tajuk. */
function bacaSemua_(nama) {
  var sh = helaian_(nama);
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0].map(function (x) { return String(x).trim().toUpperCase(); });
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _baris: i + 1 }, kosong = true;
    for (var j = 0; j < h.length; j++) { if (h[j]) { o[h[j]] = v[i][j]; if (v[i][j] !== '' && v[i][j] !== null) kosong = false; } }
    if (!kosong) out.push(o);
  }
  return out;
}
function teks_(v, jenis) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, zon_(), jenis === 'masa' ? 'HH:mm' : 'dd/MM/yyyy');
  }
  return String(v).trim();
}
/** 'dd/MM/yyyy' atau Date → 'yyyy-MM-dd' ('' jika tidak sah) */
function iso_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, zon_(), 'yyyy-MM-dd');
  var m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return '';
  var d = +m[1], b = +m[2];
  if (b < 1 || b > 12 || d < 1 || d > 31) return '';
  return m[3] + '-' + ('0' + b).slice(-2) + '-' + ('0' + d).slice(-2);
}
function masa_(v) {
  if (v instanceof Date) return teks_(v, 'masa');
  var m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
/** Elak teks ditafsir sebagai formula. */
function selamat_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function pastikanLajur_(sh, senarai) {
  var h = tajuk_(sh);
  var tiada = senarai.filter(function (x) { return h.indexOf(x) < 0; });
  if (tiada.length) sh.getRange(1, h.length + 1, 1, tiada.length).setValues([tiada]).setFontWeight('bold');
  return tajuk_(sh);
}

/* ===================== Sesi & kata laluan ===================== */

function hash_(salt, kata) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + kata, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(b);
}
function penggunaIkutId_(id) {
  id = String(id || '').trim().toLowerCase();
  var L = bacaSemua_(HELAIAN.PENGGUNA);
  for (var i = 0; i < L.length; i++) if (String(L[i]['ID']).trim().toLowerCase() === id) return L[i];
  return null;
}
function tulisPengguna_(p, medan) {
  var sh = helaian_(HELAIAN.PENGGUNA), h = tajuk_(sh);
  Object.keys(medan).forEach(function (k) {
    var c = h.indexOf(k);
    if (c >= 0) sh.getRange(p._baris, c + 1).setValue(medan[k]);
  });
}
function bentukProfil_(p) {
  return { id: String(p['ID']).trim(), nama: String(p['NAMA']).trim(), peranan: String(p['PERANAN'] || 'GURU').trim().toUpperCase(),
    mestiTukar: String(p['MESTI TUKAR']).toUpperCase() === 'YA' };
}
function log_masuk(id, kata) {
  id = String(id || '').trim().toLowerCase(); kata = String(kata || '');
  if (!id || !kata) throw new Error('Isi ID dan kata laluan.');
  var c = CacheService.getScriptCache(), kk = 'gagal_' + id;
  var gagal = +(c.get(kk) || 0);
  if (gagal >= KONFIG.MAKS_CUBAAN_GAGAL) throw new Error('Terlalu banyak cubaan. Cuba lagi selepas ' + KONFIG.KUNCI_MINIT + ' minit.');
  var p = penggunaIkutId_(id);
  var lulus = false;
  if (p && String(p['AKTIF']).toUpperCase() !== 'TIDAK') {
    var baru = String(p['KATA LALUAN BARU'] || '');
    if (baru && baru === kata) {
      var salt = Utilities.getUuid();
      tulisPengguna_(p, { 'HASH': hash_(salt, kata), 'SALT': salt, 'KATA LALUAN BARU': '', 'MESTI TUKAR': 'YA' });
      p['MESTI TUKAR'] = 'YA';
      lulus = true;
    } else if (p['HASH'] && hash_(String(p['SALT']), kata) === String(p['HASH'])) {
      lulus = true;
    }
  }
  if (!lulus) {
    c.put(kk, String(gagal + 1), KONFIG.KUNCI_MINIT * 60);
    throw new Error('ID atau kata laluan salah.');
  }
  c.remove(kk);
  tulisPengguna_(p, { 'LOG MASUK TERAKHIR': new Date() });
  var token = Utilities.getUuid() + Utilities.getUuid();
  c.put('sesi_' + token, String(p['ID']).trim().toLowerCase(), KONFIG.TEMPOH_SESI_SAAT);
  return { token: token, profil: bentukProfil_(p) };
}
function sesi_(token) {
  var c = CacheService.getScriptCache();
  var id = token && c.get('sesi_' + token);
  if (!id) throw new Error('SESI_TAMAT');
  var p = penggunaIkutId_(id);
  if (!p || String(p['AKTIF']).toUpperCase() === 'TIDAK') throw new Error('SESI_TAMAT');
  c.put('sesi_' + token, id, KONFIG.TEMPOH_SESI_SAAT);
  return bentukProfil_(p);
}
function adminSahaja_(u) { if (u.peranan !== 'ADMIN') throw new Error('Hanya admin boleh melakukan tindakan ini.'); }
function profil(token) { return sesi_(token); }
function log_keluar(token) { CacheService.getScriptCache().remove('sesi_' + token); return true; }
function tukar_kata_laluan(token, lama, baru) {
  var u = sesi_(token);
  baru = String(baru || '');
  if (baru.length < 6) throw new Error('Kata laluan baharu mesti sekurang-kurangnya 6 aksara.');
  var p = penggunaIkutId_(u.id);
  if (hash_(String(p['SALT']), String(lama || '')) !== String(p['HASH'])) throw new Error('Kata laluan semasa salah.');
  var salt = Utilities.getUuid();
  tulisPengguna_(p, { 'HASH': hash_(salt, baru), 'SALT': salt, 'MESTI TUKAR': 'TIDAK' });
  return true;
}

/* ===================== Data rujukan ===================== */

function senaraiLajur_(nama, lajur) {
  var seen = {}, out = [];
  bacaSemua_(nama).forEach(function (r) {
    var v = teks_(r[lajur]);
    if (v && !seen[v]) { seen[v] = 1; out.push(v); }
  });
  return out;
}
function kumpul_(nama, kunci, nilai) {
  var o = {};
  bacaSemua_(nama).forEach(function (r) {
    var k = teks_(r[kunci]), v = teks_(r[nilai]);
    if (!k || !v) return;
    (o[k] = o[k] || []);
    if (o[k].indexOf(v) < 0) o[k].push(v);
  });
  return o;
}
function rujukan(token) {
  sesi_(token);
  var murid = bacaSemua_(HELAIAN.MURID).map(function (r) {
    return [teks_(r['KELAS']), teks_(r['NAMA']), teks_(r['GURU KELAS / PENOLONG'])];
  }).filter(function (m) { return m[0] && m[1]; });
  var guruKelas = bacaSemua_(HELAIAN.GURUKELAS).map(function (r) {
    return [teks_(r['KELAS']), teks_(r['GURU KELAS / PENOLONG'])];
  }).filter(function (m) { return m[0]; });
  return {
    murid: murid,
    guruKelas: guruKelas,
    kesalahan: kumpul_(HELAIAN.KESALAHAN, 'KATEGORI KES', 'PERINCIAN KES'),
    unit: kumpul_(HELAIAN.UNIT, 'UNIT GURU DISIPLIN', 'GURU DISIPLIN BERTUGAS'),
    tindakan: senaraiLajur_(HELAIAN.TINDAKAN, 'TINDAKAN'),
    punca: senaraiLajur_(HELAIAN.PUNCA, 'PUNCA SALAH LAKU'),
    tempat: senaraiLajur_(HELAIAN.TEMPAT, 'TEMPAT KEJADIAN'),
    pelapor: senaraiLajur_(HELAIAN.PELAPOR, 'DILAPORKAN OLEH')
  };
}

/* ===================== Laporan ===================== */

function bentukLaporan_(r) {
  var o = { baris: r._baris };
  Object.keys(PETA).forEach(function (k) { o[k] = teks_(r[PETA[k]]); });
  o.tarikh = teks_(r['TARIKH']);
  o.tarikhIso = iso_(r['TARIKH']);
  o.masa = masa_(r['MASA']);
  o.rosak = !o.tarikhIso;
  return o;
}
function semuaLaporan_() { return bacaSemua_(HELAIAN.LAPORAN).filter(function (r) { return teks_(r['ID']); }).map(bentukLaporan_); }
function cariLaporan_(id) {
  var L = bacaSemua_(HELAIAN.LAPORAN);
  for (var i = 0; i < L.length; i++) if (teks_(L[i]['ID']) === String(id)) return L[i];
  return null;
}
function senarai(token) {
  sesi_(token);
  return semuaLaporan_().map(function (o) {
    return { id: o.id, kelas: o.kelas, nama: o.nama, kategori: o.kategori, perincian: o.perincian, tarikh: o.tarikh,
      tarikhIso: o.tarikhIso, masa: o.masa, tempat: o.tempat, pelapor: o.pelapor, unit: o.unit,
      gambar1: o.gambar1, ada_pdf: !!(o.pdf || o.pautan), rosak: o.rosak };
  });
}
function butiran(token, id) {
  var u = sesi_(token);
  var r = cariLaporan_(id);
  if (!r) throw new Error('Laporan tidak dijumpai. Mungkin telah dipadam.');
  var o = bentukLaporan_(r);
  o.bilKesMurid = o.nama ? semuaLaporan_().filter(function (x) { return x.nama === o.nama && x.kelas === o.kelas; }).length : 0;
  o.bolehUbah = u.peranan === 'ADMIN';
  return o;
}

/* ===================== Fail dalam Drive (folder AppSheet) ===================== */

function folderApp_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_APP_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var it = DriveApp.getFoldersByName(KONFIG.FOLDER_APP);
  if (!it.hasNext()) throw new Error('Folder AppSheet "' + KONFIG.FOLDER_APP + '" tidak dijumpai dalam Drive.');
  var f = it.next();
  p.setProperty('FOLDER_APP_ID', f.getId());
  return f;
}
function subFolder_(nama) {
  var app = folderApp_(), it = app.getFoldersByName(nama);
  return it.hasNext() ? it.next() : app.createFolder(nama);
}
/** Folder asas untuk laluan relatif: folder yang mengandungi Sheet, kemudian folder aplikasi AppSheet. */
function folderAsas_() {
  var out = [], ids = {};
  try {
    var it = DriveApp.getFileById(ss_().getId()).getParents();
    while (it.hasNext()) { var f = it.next(); if (!ids[f.getId()]) { ids[f.getId()] = 1; out.push(f); } }
  } catch (e) { }
  try { var a = folderApp_(); if (!ids[a.getId()]) out.push(a); } catch (e) { }
  return out;
}
/** Folder LAPORAN_Images yang sama digunakan AppSheet (dicari sekali, kemudian diingati). */
function folderGambar_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_GAMBAR_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var asas = folderAsas_(), f = null;
  for (var i = 0; i < asas.length && !f; i++) { var it = asas[i].getFoldersByName(KONFIG.FOLDER_GAMBAR); if (it.hasNext()) f = it.next(); }
  if (!f) {
    // Cari melalui fail gambar sedia ada dalam LAPORAN
    var L = bacaSemua_(HELAIAN.LAPORAN);
    for (var j = L.length - 1; j >= 0 && !f; j--) {
      var nm = teks_(L[j]['GAMBAR 1']).split('/').pop();
      if (!nm) continue;
      var fi = DriveApp.getFilesByName(nm);
      if (fi.hasNext()) { var pi = fi.next().getParents(); if (pi.hasNext()) f = pi.next(); }
    }
  }
  if (!f) f = (asas[0] || folderApp_()).createFolder(KONFIG.FOLDER_GAMBAR);
  p.setProperty('FOLDER_GAMBAR_ID', f.getId());
  return f;
}
function ikutLaluan_(folder, bah, nama) {
  for (var i = 0; i < bah.length; i++) {
    var it = folder.getFoldersByName(bah[i]);
    if (!it.hasNext()) return null;
    folder = it.next();
  }
  var f = folder.getFilesByName(nama);
  while (f.hasNext()) { var x = f.next(); if (!x.isTrashed || !x.isTrashed()) return x; }
  return null;
}
/** 'LAPORAN_Images/abc.jpg' atau '/appsheet/data/<app>/DISIPLINREPORT2026/x.pdf' → File, atau null */
function failIkutLaluan_(laluan) {
  laluan = String(laluan || '').replace(/^\/+/, '');
  if (!laluan) return null;
  var bah = laluan.split('/');
  var nama = bah.pop(), x;
  var awal = bah.indexOf(KONFIG.FOLDER_APP);
  if (awal >= 0) {
    try { x = ikutLaluan_(folderApp_(), bah.slice(awal + 1), nama); if (x) return x; } catch (e) { }
  } else {
    if (bah.length === 1 && bah[0] === KONFIG.FOLDER_GAMBAR) {
      try { x = ikutLaluan_(folderGambar_(), [], nama); if (x) return x; } catch (e) { }
    }
    var asas = folderAsas_();
    for (var i = 0; i < asas.length; i++) { x = ikutLaluan_(asas[i], bah, nama); if (x) return x; }
  }
  // Pilihan terakhir: cari ikut nama fail (nama AppSheet mengandungi ID + masa, jadi unik)
  var g = DriveApp.getFilesByName(nama);
  while (g.hasNext()) { x = g.next(); if (!x.isTrashed || !x.isTrashed()) return x; }
  return null;
}
function dataUrl_(blob) {
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}
function gambar(token, laluan, kecil) {
  sesi_(token);
  laluan = String(laluan || '');
  if (laluan.indexOf('..') >= 0) throw new Error('Laluan tidak sah.');
  var f = failIkutLaluan_(laluan);
  if (!f) return '';
  if (kecil) { try { var t = f.getThumbnail(); if (t) return dataUrl_(t); } catch (e) { } }
  return dataUrl_(f.getBlob());
}
function simpanFail_(dataUrl, id, lajur) {
  var m = String(dataUrl || '').match(/^data:([\w\/+.-]+);base64,(.+)$/);
  if (!m) throw new Error('Fail "' + lajur + '" tidak sah.');
  var jenis = m[1];
  if (!/^image\/(jpeg|png|webp)$|^application\/pdf$/.test(jenis)) throw new Error('Jenis fail "' + lajur + '" tidak dibenarkan.');
  var ext = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'application/pdf': '.pdf' }[jenis];
  var nama = id + '.' + lajur + '.' + Utilities.formatDate(new Date(), zon_(), 'HHmmss') + ext;
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), jenis, nama);
  folderGambar_().createFile(blob);
  return KONFIG.FOLDER_GAMBAR + '/' + nama;
}

/* ===================== Tulis laporan ===================== */

function sahkan_(d) {
  var wajib = { kelas: 'Kelas', nama: 'Nama murid', kategori: 'Kategori kes', tarikh: 'Tarikh', masa: 'Masa' };
  Object.keys(wajib).forEach(function (k) { if (!String(d[k] || '').trim()) throw new Error(wajib[k] + ' wajib diisi.'); });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.tarikh)) throw new Error('Tarikh tidak sah.');
  if (!/^\d{2}:\d{2}$/.test(d.masa)) throw new Error('Masa tidak sah.');
}
function tarikhSel_(iso) { var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function nilaiBaris_(h, d, lama) {
  var row = lama ? lama.slice() : h.map(function () { return ''; });
  function set(k, v) { var c = h.indexOf(PETA[k]); if (c >= 0) row[c] = v; }
  MEDAN_TEKS.forEach(function (k) { if (d[k] !== undefined) set(k, selamat_(Array.isArray(d[k]) ? d[k].join(KONFIG.PEMISAH) : d[k])); });
  if (d.tarikh) set('tarikh', tarikhSel_(d.tarikh));
  if (d.masa) set('masa', d.masa + ':00');
  return row;
}
function hantar(token, d, fail) {
  var u = sesi_(token);
  d = d || {}; fail = fail || {};
  sahkan_(d);
  ['gambar1', 'gambar2', 'tandatangan'].forEach(function (k) { if (!fail[k]) throw new Error({ gambar1: 'Gambar 1', gambar2: 'Gambar 2', tandatangan: 'Tandatangan pelapor' }[k] + ' wajib.'); });
  var id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  // Simpan fail dahulu (di luar kunci; mengambil masa)
  var laluan = {};
  Object.keys(MEDAN_FAIL).forEach(function (k) { if (fail[k]) laluan[k] = simpanFail_(fail[k], id, MEDAN_FAIL[k]); });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh);
    var row = nilaiBaris_(h, d);
    var c = h.indexOf('ID'); if (c >= 0) row[c] = id;
    c = h.indexOf(PETA.pelapor); if (c >= 0) row[c] = selamat_(d.pelapor || u.nama);
    Object.keys(laluan).forEach(function (k) { var j = h.indexOf(MEDAN_FAIL[k]); if (j >= 0) row[j] = laluan[k]; });
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
  var hasil = { id: id, pdf: false, telegram: false };
  try { janaPdf_(id); hasil.pdf = true; } catch (e) { hasil.ralatPdf = String(e.message || e); }
  try { hasil.telegram = beritahuTelegram_(cariLaporan_(id)); } catch (e) { hasil.ralatTelegram = String(e.message || e); }
  return hasil;
}
function ubah(token, id, d, fail) {
  var u = sesi_(token); adminSahaja_(u);
  d = d || {}; fail = fail || {};
  sahkan_(d);
  var laluan = {};
  Object.keys(MEDAN_FAIL).forEach(function (k) { if (fail[k]) laluan[k] = simpanFail_(fail[k], id, MEDAN_FAIL[k]); });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var r = cariLaporan_(id);
    if (!r) throw new Error('Laporan tidak dijumpai.');
    var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh);
    var lama = sh.getRange(r._baris, 1, 1, h.length).getValues()[0];
    var row = nilaiBaris_(h, d, lama);
    if (d.pelapor) { var c = h.indexOf(PETA.pelapor); if (c >= 0) row[c] = selamat_(d.pelapor); }
    Object.keys(laluan).forEach(function (k) { var j = h.indexOf(MEDAN_FAIL[k]); if (j >= 0) row[j] = laluan[k]; });
    sh.getRange(r._baris, 1, 1, h.length).setValues([row]);
  } finally { lock.releaseLock(); }
  var hasil = { id: id, pdf: false };
  try { janaPdf_(id); hasil.pdf = true; } catch (e) { hasil.ralatPdf = String(e.message || e); }
  return hasil;
}
function padam(token, id) {
  var u = sesi_(token); adminSahaja_(u);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var r = cariLaporan_(id);
    if (!r) throw new Error('Laporan tidak dijumpai.');
    var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh);
    var nilai = sh.getRange(r._baris, 1, 1, h.length).getValues()[0];
    var arkib = ss_().getSheetByName(HELAIAN.PADAM);
    if (!arkib) {
      arkib = ss_().insertSheet(HELAIAN.PADAM);
      arkib.getRange(1, 1, 1, h.length + 2).setValues([h.concat(['DIPADAM OLEH', 'DIPADAM PADA'])]).setFontWeight('bold');
    }
    arkib.appendRow(nilai.concat([u.nama, new Date()]));
    sh.deleteRow(r._baris);
  } finally { lock.releaseLock(); }
  return true;
}
/** PDF sebagai data URL supaya semua guru boleh membukanya tanpa akses Drive pemilik. */
function pdf_fail(token, id) {
  sesi_(token);
  var r = cariLaporan_(id);
  if (!r) throw new Error('Laporan tidak dijumpai.');
  var f = teks_(r[PETA.pdf]) ? failIkutLaluan_(teks_(r[PETA.pdf])) : null;
  if (!f) {
    var m = teks_(r[PETA.pautan]).match(/\/d\/([\w-]{20,})/);
    if (m) { try { f = DriveApp.getFileById(m[1]); } catch (e) { f = null; } }
  }
  if (!f) f = janaPdf_(id);
  return { nama: 'Laporan_' + teks_(r['KELAS']) + '_' + id + '.pdf', data: dataUrl_(f.getBlob()) };
}
function jana_pdf(token, id) {
  sesi_(token);
  janaPdf_(id);
  var r = cariLaporan_(id);
  return { pautan: teks_(r[PETA.pautan]) };
}

/* ===================== PDF (ikut templat Unit Disiplin) ===================== */

function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function imejPdf_(laluan) {
  try { var f = failIkutLaluan_(laluan); return f ? dataUrl_(f.getBlob()) : ''; } catch (e) { return ''; }
}
function htmlPdf_(o) {
  var biru = '#3c78d8', latar = '#cfe2f3';
  var baris = [['KELAS:', o.kelas], ['NAMA MURID:', o.nama], ['MASA:', o.masa], ['GURU KELAS / PENOLONG:', o.guruKelas],
    ['KATEGORI KES:', o.kategori], ['PERINCIAN KES:', o.perincian], ['KETERANGAN KES', o.keterangan],
    ['PUNCA SALAH LAKU:', o.punca], ['TARIKH:', o.tarikh], ['TEMPAT:', o.tempat], ['UNIT GURU DISIPLIN:', o.unit],
    ['GURU DISIPLIN BERTUGAS:', o.guruBertugas], ['DILAPORKAN OLEH:', o.pelapor], ['TINDAKAN:', o.tindakan]];
  var td = 'border:1.5px solid #222;padding:6px 8px;font-size:10pt;vertical-align:top;';
  var jadual = baris.map(function (b) {
    return '<tr><td bgcolor="' + biru + '" style="' + td + 'width:32%;background:' + biru + ';color:#fff;font-weight:bold"><font color="#ffffff">' + esc_(b[0]) + '</font></td>' +
      '<td bgcolor="' + latar + '" style="' + td + 'background:' + latar + '">' + esc_(b[1]) + '</td></tr>';
  }).join('');
  var ttd = imejPdf_(o.tandatangan), g1 = imejPdf_(o.gambar1), g2 = imejPdf_(o.gambar2), dok = o.dokumen ? imejPdf_(o.dokumen) : '';
  function kotakGambar(tajuk, src) {
    return '<table style="width:100%;border-collapse:collapse;margin:0 0 18px;page-break-inside:avoid"><tr><td bgcolor="#1155cc" style="' + td + 'background:#1155cc;color:#fff;font-weight:bold"><font color="#ffffff">' + tajuk + '</font>' +
      '</td></tr><tr><td bgcolor="' + latar + '" style="' + td + 'background:' + latar + ';text-align:center;height:250px">' +
      (src ? '<img src="' + src + '" style="max-width:100%;max-height:245px">' : '(tiada gambar)') + '</td></tr></table>';
  }
  var kepala = '<div style="text-align:center;margin-top:10px"><div style="font-family:\'Times New Roman\',serif;font-weight:bold;font-size:15pt;color:#1a1aa6">LAPORAN KES DISIPLIN MURID ' + esc_(KONFIG.NAMA_SEKOLAH.toUpperCase()) + '</div>' +
    '<div style="text-align:right;font-style:italic;font-weight:bold;font-size:8.5pt;color:#555">Hak Milik Unit Disiplin SMK Bugaya II</div></div>';
  return '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:16mm}body{font-family:Arial,sans-serif;background:' + latar + ';color:#000}</style></head><body bgcolor="' + latar + '" style="background:' + latar + '">' +
    kepala +
    '<table style="width:100%;border-collapse:collapse;margin-top:22px">' + jadual + '</table>' +
    '<table style="width:52%;border-collapse:collapse;margin-top:40px"><tr><td bgcolor="' + biru + '" style="' + td + 'background:' + biru + ';color:#fff;font-weight:bold"><font color="#ffffff">TANDATANGAN PELAPOR:</font></td></tr>' +
    '<tr><td bgcolor="' + latar + '" style="' + td + 'background:' + latar + ';height:150px">' + (ttd ? '<img src="' + ttd + '" style="max-height:110px;max-width:100%"><br>' : '<br><br><br><br>') +
    '(' + esc_(o.pelapor) + ')</td></tr></table>' +
    '<div style="page-break-before:always"></div>' + kepala +
    '<div style="font-weight:bold;font-size:11pt;margin:28px 0 12px">LAPORAN BERGAMBAR</div>' +
    kotakGambar('GAMBAR 1:', g1) + kotakGambar('GAMBAR 2:', g2) +
    (o.dokumen ? kotakGambar('DOKUMEN SOKONGAN:', dok) : '') +
    '</body></html>';
}
function janaPdf_(id) {
  var r = cariLaporan_(id);
  if (!r) throw new Error('Laporan tidak dijumpai.');
  var o = bentukLaporan_(r);
  var pdf = HtmlService.createHtmlOutput(htmlPdf_(o)).getBlob().getAs('application/pdf').setName(id + '.pdf');
  var folder = subFolder_(KONFIG.FOLDER_PDF);
  var lama = folder.getFilesByName(id + '.pdf');
  while (lama.hasNext()) lama.next().setTrashed(true);
  var f = folder.createFile(pdf);
  var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh);
  var cp = h.indexOf(PETA.pdf), cu = h.indexOf(PETA.pautan);
  if (cp >= 0) sh.getRange(r._baris, cp + 1).setValue('/appsheet/data/' + KONFIG.FOLDER_APP + '/' + KONFIG.FOLDER_PDF + '/' + id + '.pdf');
  if (cu >= 0) sh.getRange(r._baris, cu + 1).setValue(f.getUrl());
  return f;
}

/* ===================== Telegram ===================== */

function telegram_(teks) {
  var p = PropertiesService.getScriptProperties();
  var tok = p.getProperty('TELEGRAM_TOKEN'), chat = p.getProperty('TELEGRAM_CHAT_ID');
  if (!tok || !chat) return false;
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + tok + '/sendMessage', {
    method: 'post', muteHttpExceptions: true,
    payload: { chat_id: chat, text: teks, parse_mode: 'HTML', disable_web_page_preview: 'true' }
  });
  var j = JSON.parse(res.getContentText() || '{}');
  if (!j.ok) throw new Error('Telegram: ' + (j.description || res.getResponseCode()));
  return true;
}
function beritahuTelegram_(r) {
  if (!r) return false;
  var o = bentukLaporan_(r);
  var t = '<b>LAPORAN KES DISIPLIN BAHARU</b>\n' +
    'Kelas: ' + esc_(o.kelas) + '\n' +
    'Kategori: ' + esc_(o.kategori) + '\n' +
    (o.perincian ? 'Perincian: ' + esc_(o.perincian) + '\n' : '') +
    'Tarikh: ' + esc_(o.tarikh) + ', ' + esc_(o.masa) + '\n' +
    (o.tempat ? 'Tempat: ' + esc_(o.tempat) + '\n' : '') +
    (o.unit ? 'Unit: ' + esc_(o.unit) + '\n' : '') +
    'Dilaporkan oleh: ' + esc_(o.pelapor) + '\n' +
    (o.tindakan ? 'Tindakan: ' + esc_(o.tindakan) + '\n' : '') +
    '\nButiran penuh dalam ' + KONFIG.NAMA_SISTEM + ': ' + KONFIG.PAUTAN_APLIKASI;
  return telegram_(t);
}
