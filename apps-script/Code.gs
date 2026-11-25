/**
 * MyDisiplin – pelayan (Google Apps Script)
 * Laporan kes disiplin murid SMK Bugaya II Semporna
 * Versi 1.0.11 · Oleh Cikgu Henry Able Padtong
 *
 * Serasi dengan AppSheet MYDISIPLIN26: menulis ke helaian LAPORAN yang sama,
 * ID 8 aksara, TARIKH sebagai tarikh, gambar dalam folder LAPORAN_Images aplikasi AppSheet.
 */

var KONFIG = {
  NAMA_SEKOLAH: 'SMK Bugaya II Semporna',
  NAMA_SISTEM: 'MyDisiplin',
  VERSI: '1.0.11',
  // Cache pelayan (saat). Data LAPORAN juga dibatalkan serta-merta apabila laporan dihantar/diubah/dipadam
  // atau Sheet disunting dengan tangan. Suntingan melalui AppSheet kelihatan selepas tempoh ini (atau butang Muat semula).
  CACHE_LAPORAN_SAAT: 600,
  CACHE_RUJUKAN_SAAT: 21600,
  CACHE_PROFIL_SAAT: 300,
  TEMPOH_SESI_SAAT: 21600,          // 6 jam
  MAKS_CUBAAN_GAGAL: 5,
  KUNCI_MINIT: 15,
  // Folder data aplikasi AppSheet dalam Google Drive anda (nama folder, bukan laluan)
  FOLDER_APP: 'MYDISIPLIN26-302733627-26-04-26-3',
  FOLDER_GAMBAR: 'LAPORAN_Images',
  FOLDER_PDF: 'DISIPLINREPORT2026',
  PAUTAN_APLIKASI: 'https://hapsmkbii.github.io/mydisiplin/',
  PEMISAH: ' , ',                    // pemisah senarai yang digunakan AppSheet
  ADMIN_AWAL: ['HENRY ABLE PADTONG'], // dijadikan ADMIN semasa "Isi pengguna"
  // Templat Google Docs yang sama digunakan oleh bot AppSheet (JANA LAPORAN). Kosongkan untuk guna templat HTML ringkas.
  TEMPLAT_DOC_ID: '1hMuf-u4pgCVmbQiRc1bLDjMjjnwreOTTtApWKmxpiOI'
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

var SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.getActive()); }
var ZON_ = null;
// Diingati sekali setiap panggilan: dahulu dipanggil untuk setiap sel tarikh/masa (beribu panggilan ke Sheet).
function zon_() { return ZON_ || (ZON_ = ss_().getSpreadsheetTimeZone() || 'Asia/Kuala_Lumpur'); }
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

/* ===================== Cache (mempercepatkan aplikasi) ===================== */

// CacheService menghadkan satu nilai kepada 100 KB, jadi data besar dipecah kepada cebisan.
var CEBISAN = 30000;
function cache_() { return CacheService.getScriptCache(); }
function cacheBaca_(kunci) {
  var c = cache_(), n = +(c.get(kunci + '_n') || 0);
  if (!n) return null;
  var ks = [];
  for (var i = 0; i < n; i++) ks.push(kunci + '_' + i);
  var m = c.getAll(ks), s = '';
  for (var j = 0; j < n; j++) { if (m[ks[j]] == null) return null; s += m[ks[j]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cacheSimpan_(kunci, obj, saat, gen) {
  try {
    var c = cache_();
    if (gen !== undefined && c.get(kunci + '_gen') !== gen) return; // data berubah semasa membaca: jangan simpan
    var s = JSON.stringify(obj), o = {}, n = 0;
    for (var i = 0; i < s.length; i += CEBISAN) o[kunci + '_' + (n++)] = s.substr(i, CEBISAN);
    c.putAll(o, saat);
    c.put(kunci + '_n', String(n), saat);
  } catch (e) { console.log('Cache gagal: ' + e); }
}
function cacheGen_(kunci) { return cache_().get(kunci + '_gen'); }
function cacheBuang_(kunci) {
  try { var c = cache_(); c.put(kunci + '_gen', Utilities.getUuid(), 21600); c.remove(kunci + '_n'); } catch (e) { }
}
/** Simple trigger: batalkan cache apabila Sheet disunting dengan tangan. */
function onEdit(e) {
  try {
    var sh = e && e.range && e.range.getSheet(), nama = sh ? sh.getName() : '';
    if (nama === HELAIAN.LAPORAN) cacheBuang_('lap');
    else if (nama === HELAIAN.PENGGUNA) {
      var c = tajuk_(sh).indexOf('ID');
      if (c >= 0 && e.range.getRow() > 1) {
        sh.getRange(e.range.getRow(), c + 1, e.range.getNumRows(), 1).getValues().forEach(function (v) {
          cache_().remove('prof_' + String(v[0]).trim().toLowerCase());
        });
      }
    } else if ([HELAIAN.MURID, HELAIAN.KESALAHAN, HELAIAN.UNIT, HELAIAN.PELAPOR, HELAIAN.TINDAKAN, HELAIAN.PUNCA,
      HELAIAN.TEMPAT, HELAIAN.GURUKELAS].indexOf(nama) >= 0) cacheBuang_('ruj');
  } catch (err) { }
}
/** Pemicu masa: isi semula cache setiap 10 minit supaya guru jarang menunggu bacaan Sheet. */
function panaskan() {
  try { laporanRingkas_(true); } catch (e) { console.log('panaskan lap: ' + e); }
  try { rujukanData_(false); } catch (e) { console.log('panaskan ruj: ' + e); }
}
/** Jalankan SEKALI dari editor (pilih pasangPemanas › Run). */
function pasangPemanas() {
  buangPemanas();
  ScriptApp.newTrigger('panaskan').timeBased().everyMinutes(10).create();
  panaskan();
  return 'Pemanas dipasang: cache diisi semula setiap 10 minit.';
}
function buangPemanas() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'panaskan') ScriptApp.deleteTrigger(t); });
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
  cache_().remove('prof_' + String(p['ID']).trim().toLowerCase());
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
  var prof = null;
  try { prof = JSON.parse(c.get('prof_' + id) || 'null'); } catch (e) { }
  if (!prof) {
    var p = penggunaIkutId_(id);
    if (!p || String(p['AKTIF']).toUpperCase() === 'TIDAK') throw new Error('SESI_TAMAT');
    prof = bentukProfil_(p);
    c.put('prof_' + id, JSON.stringify(prof), KONFIG.CACHE_PROFIL_SAAT);
  }
  c.put('sesi_' + token, id, KONFIG.TEMPOH_SESI_SAAT);
  return prof;
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
function rujukan(token, segar) {
  sesi_(token);
  return rujukanData_(!!segar);
}
function rujukanData_(segar) {
  if (!segar) { var ada = cacheBaca_('ruj'); if (ada) return ada; }
  var gen = cacheGen_('ruj');
  var d = rujukanBaca_();
  cacheSimpan_('ruj', d, KONFIG.CACHE_RUJUKAN_SAAT, gen);
  return d;
}
function rujukanBaca_() {
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
/** Cari satu laporan ikut ID dengan TextFinder (hanya baris itu dibaca, bukan seluruh helaian). */
function cariLaporan_(id) {
  id = String(id == null ? '' : id).trim();
  if (!id) return null;
  try {
    var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh), c = h.indexOf('ID'), n = sh.getLastRow();
    if (c >= 0 && n >= 2) {
      var jumpa = sh.getRange(2, c + 1, n - 1, 1).createTextFinder(id).matchCase(true).matchEntireCell(true).findAll();
      for (var k = 0; k < jumpa.length; k++) {
        var b = jumpa[k].getRow(), v = sh.getRange(b, 1, 1, h.length).getValues()[0];
        if (teks_(v[c]) !== id) continue;
        var o = { _baris: b };
        for (var j = 0; j < h.length; j++) if (h[j]) o[h[j]] = v[j];
        return o;
      }
    }
  } catch (e) { console.log('TextFinder gagal: ' + e); }
  return cariLaporanPenuh_(id);
}
function cariLaporanPenuh_(id) {
  var L = bacaSemua_(HELAIAN.LAPORAN);
  for (var i = 0; i < L.length; i++) if (teks_(L[i]['ID']) === String(id)) return L[i];
  return null;
}
/** Senarai ringkas semua laporan, disimpan dalam cache pelayan. segar = abaikan cache. */
function laporanRingkas_(segar) {
  if (!segar) { var ada = cacheBaca_('lap'); if (ada) return ada; }
  var gen = cacheGen_('lap');
  var L = semuaLaporan_().map(function (o) {
    return { id: o.id, kelas: o.kelas, nama: o.nama, kategori: o.kategori, perincian: o.perincian, tarikh: o.tarikh,
      tarikhIso: o.tarikhIso, masa: o.masa, tempat: o.tempat, pelapor: o.pelapor, unit: o.unit,
      gambar1: o.gambar1, tindakan: o.tindakan, ada_pdf: !!(o.pdf || o.pautan), rosak: o.rosak };
  });
  cacheSimpan_('lap', L, KONFIG.CACHE_LAPORAN_SAAT, gen);
  return L;
}
function senarai(token, segar) {
  sesi_(token);
  return laporanRingkas_(!!segar);
}
function butiran(token, id) {
  var u = sesi_(token);
  var r = cariLaporan_(id);
  if (!r) throw new Error('Laporan tidak dijumpai. Mungkin telah dipadam.');
  var o = bentukLaporan_(r);
  o.bilKesMurid = o.nama ? laporanRingkas_(false).filter(function (x) { return x.nama === o.nama && x.kelas === o.kelas; }).length : 0;
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
  var u = String(laluan || '').match(/^https?:\/\/[^ ]*?(?:\/d\/|[?&]id=)([\w-]{20,})/);
  if (u) { try { return DriveApp.getFileById(u[1]); } catch (e) { return null; } }
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
  var kunci = kecil ? 'th_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, laluan)) : '';
  if (kunci) { var ada = cacheBaca_(kunci); if (ada) return ada; }
  var f = failIkutLaluan_(laluan);
  if (!f) return '';
  if (kecil) {
    try { var t = f.getThumbnail(); if (t) { var u = dataUrl_(t); cacheSimpan_(kunci, u, 21600); return u; } } catch (e) { }
  }
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
    cacheBuang_('lap');
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
    cacheBuang_('lap');
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
    cacheBuang_('lap');
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
// Saiz imej dalam PDF (pt). lebar 0 = penuh lebar petak. isi: true = petak mengecut ikut imej (buang semua baris kosong).
var LAJUR_IMEJ = {
  'GAMBAR 1': { lebar: 0, tinggi: 250, isi: true },
  'GAMBAR 2': { lebar: 0, tinggi: 250, isi: true },
  'DOKUMEN SOKONGAN': { lebar: 0, tinggi: 600, isi: true },   // cukup untuk gambar satu muka surat A4
  'TANDATANGAN PELAPOR': { lebar: 150, tinggi: 55, isi: false }
};
var TINGGI_BARIS = 15; // anggaran tinggi satu baris kosong dalam templat (pt)
/** Lebar ruang dalam petak jadual (atau lebar halaman) dalam pt. */
function lebarPetak_(el, body) {
  var x = el, sel = null;
  for (var i = 0; i < 6 && x; i++) {
    if (x.getType && x.getType() === DocumentApp.ElementType.TABLE_CELL) { sel = x.asTableCell(); break; }
    x = x.getParent && x.getParent();
  }
  var w = null;
  try { if (sel && sel.getWidth()) w = sel.getWidth() - (sel.getPaddingLeft() || 5) - (sel.getPaddingRight() || 5); } catch (e) { }
  if (!w) { try { w = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight() - 12; } catch (e) { w = 440; } }
  return { lebar: w, sel: sel };
}
function reEsc_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function ganti_(s) { return String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/\$/g, '\\$'); }
/** Isi templat Google Docs AppSheet (<<[LAJUR]>>) dan pulangkan blob PDF. */
function pdfDariTemplat_(r, id) {
  var o = bentukLaporan_(r);
  var nilai = {};
  Object.keys(r).forEach(function (k) { if (k !== '_baris') nilai[k] = teks_(r[k]); });
  nilai['TARIKH'] = o.tarikh; nilai['MASA'] = o.masa;
  var salinan = DriveApp.getFileById(KONFIG.TEMPLAT_DOC_ID).makeCopy('tmp-' + id, subFolder_(KONFIG.FOLDER_PDF));
  try {
    var doc = DocumentApp.openById(salinan.getId());
    var bah = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(function (x) { return x; });
    bah.forEach(function (b) {
      Object.keys(LAJUR_IMEJ).forEach(function (lajur) {
        var re = '<<\\[' + reEsc_(lajur) + '\\]>>', jumpa, had = 5;
        while ((jumpa = b.findText(re)) && had-- > 0) {
          var t = jumpa.getElement().asText();
          t.deleteText(jumpa.getStartOffset(), jumpa.getEndOffsetInclusive());
          var f = nilai[lajur] ? failIkutLaluan_(nilai[lajur]) : null;
          if (!f) continue;
          var induk = t.getParent();
          var img = induk.insertInlineImage(induk.getChildIndex(t) + 1, f.getBlob());
          var c = LAJUR_IMEJ[lajur], ruang = lebarPetak_(induk, doc.getBody());
          var maksW = c.lebar || ruang.lebar, maksH = c.tinggi;
          var w = img.getWidth(), h = img.getHeight();
          var k = Math.min(maksW / w, maksH / h);            // besarkan atau kecilkan, kekalkan nisbah (tiada potongan)
          var tinggi = Math.round(h * k);
          img.setWidth(Math.round(w * k)).setHeight(tinggi);
          if (c.isi) { try { induk.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) { } }
          // Buang baris kosong selepas imej supaya petak tidak melimpah ke halaman lain
          var buang = c.isi ? 999 : Math.ceil((tinggi - TINGGI_BARIS) / TINGGI_BARIS);
          var nx = induk.getNextSibling();
          while (buang > 0 && nx && nx.getType() === DocumentApp.ElementType.PARAGRAPH &&
                 nx.asParagraph().getText() === '' && nx.asParagraph().getNumChildren() <= 1 && (ruang.sel || nx.getNextSibling())) {
            var seterusnya = nx.getNextSibling();
            nx.removeFromParent(); buang--; nx = seterusnya;
          }
        }
      });
      Object.keys(nilai).forEach(function (lajur) {
        if (LAJUR_IMEJ[lajur]) return;
        b.replaceText('<<\\[' + reEsc_(lajur) + '\\]>>', ganti_(nilai[lajur]));
      });
      b.replaceText('<<[^>]*>>', '');
    });
    doc.saveAndClose();
    return salinan.getAs('application/pdf').setName(id + '.pdf');
  } finally {
    salinan.setTrashed(true);
  }
}
function janaPdf_(id) {
  var r = cariLaporan_(id);
  if (!r) throw new Error('Laporan tidak dijumpai.');
  var o = bentukLaporan_(r);
  var pdf = null;
  if (KONFIG.TEMPLAT_DOC_ID) {
    try { pdf = pdfDariTemplat_(r, id); } catch (e) { console.log('Templat Docs gagal: ' + e); }
  }
  if (!pdf) pdf = HtmlService.createHtmlOutput(htmlPdf_(o)).getBlob().getAs('application/pdf').setName(id + '.pdf');
  var folder = subFolder_(KONFIG.FOLDER_PDF);
  var lama = folder.getFilesByName(id + '.pdf');
  while (lama.hasNext()) lama.next().setTrashed(true);
  var f = folder.createFile(pdf);
  var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh);
  var cp = h.indexOf(PETA.pdf), cu = h.indexOf(PETA.pautan);
  // Pautan Drive penuh (bukan laluan /appsheet/data/...): AppSheet tidak dapat melihat fail yang
  // dicipta di luar AppSheet melalui laluan, tetapi boleh membuka pautan terus.
  if (cp >= 0) sh.getRange(r._baris, cp + 1).setValue(f.getUrl());
  if (cu >= 0) sh.getRange(r._baris, cu + 1).setValue(f.getUrl());
  cacheBuang_('lap');
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
  var v = function (x) { return esc_(x || '–'); };
  var t = '<b>LAPORAN KES DISIPLIN BAHARU</b>\n' +
    'Nama murid: ' + v(o.nama) + '\n' +
    'Kelas: ' + v(o.kelas) + '\n' +
    'Kategori: ' + v(o.kategori) + '\n' +
    'Perincian: ' + v(o.perincian) + '\n' +
    'Keterangan kes: ' + v(o.keterangan) + '\n' +
    'Tarikh dan masa: ' + v([o.tarikh, o.masa].filter(String).join(', ')) + '\n' +
    'Dilaporkan oleh: ' + v(o.pelapor);
  return telegram_(t);
}
