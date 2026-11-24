/**
 * MyDisiplin – menu pentadbir dalam Google Sheet.
 * Jalankan fungsi ini dari menu "MyDisiplin" dalam Sheet (bukan dari editor).
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('MyDisiplin')
    .addItem('1. Sediakan helaian PENGGUNA', 'sediakanPengguna')
    .addItem('2. Isi pengguna daripada DILAPORKAN OLEH', 'isiPengguna')
    .addSeparator()
    .addItem('3. Tetapkan Telegram (token & chat)', 'tetapkanTelegram')
    .addItem('4. Cari Chat ID kumpulan Telegram', 'cariChatId')
    .addItem('5. Uji notifikasi Telegram', 'ujiTelegram')
    .addSeparator()
    .addItem('Semak folder AppSheet', 'semakFolder')
    .addItem('Semak baris LAPORAN bermasalah', 'semakBaris')
    .addItem('Jana PDF yang belum ada', 'janaPdfTiada')
    .addItem('Jana semula PDF buatan MyDisiplin', 'janaSemulaPdfMyDisiplin')
    .addItem('Betulkan pautan PDF untuk AppSheet', 'betulkanPautanPdf')
    .addItem('Tunjuk pautan aplikasi', 'tunjukPautan')
    .addToUi();
}

function sediakanPengguna() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(HELAIAN.PENGGUNA);
  if (!sh) sh = ss.insertSheet(HELAIAN.PENGGUNA);
  pastikanLajur_(sh, LAJUR_PENGGUNA);
  sh.setFrozenRows(1);
  sh.getRange('A:I').setNumberFormat('@');
  SpreadsheetApp.getUi().alert('Helaian PENGGUNA sedia. Seterusnya: menu 2 untuk mengisi guru secara automatik, atau isi sendiri (ID, NAMA, PERANAN ADMIN/GURU, AKTIF YA, KATA LALUAN BARU).');
}

function slug_(nama) {
  var buang = ['BIN', 'BINTI', 'BT', 'BTE', 'B', 'BT.', 'B.', '@', 'A/L', 'A/P', 'AL', 'AP', 'HJ', 'HJ.', 'HAJI', 'HAJAH', 'EN.', 'PN.', 'CIK', 'UST.', 'USTZ', 'USTAZ', 'USTAZAH'];
  var w = String(nama).toUpperCase().replace(/[^A-Z@\/. ]/g, ' ').split(/\s+/).filter(function (x) { return x && buang.indexOf(x) < 0; });
  return w.slice(0, 2).join('.').toLowerCase().replace(/[^a-z.]/g, '');
}
function kataSementara_() {
  var a = 'abcdefghjkmnpqrstuvwxyz23456789', s = '';
  for (var i = 0; i < 7; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}
function isiPengguna() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(HELAIAN.PENGGUNA);
  if (!sh) { sediakanPengguna(); sh = ss.getSheetByName(HELAIAN.PENGGUNA); }
  var h = pastikanLajur_(sh, LAJUR_PENGGUNA);
  var ada = bacaSemua_(HELAIAN.PENGGUNA);
  var namaAda = {}, idAda = {};
  ada.forEach(function (p) { namaAda[teks_(p['NAMA']).toUpperCase().replace(/\s+/g, ' ')] = 1; idAda[teks_(p['ID']).toLowerCase()] = 1; });
  var baru = [];
  senaraiLajur_(HELAIAN.PELAPOR, 'DILAPORKAN OLEH').forEach(function (nama) {
    var kunci = nama.toUpperCase().replace(/\s+/g, ' ');
    if (namaAda[kunci]) return;
    var id = slug_(nama) || 'guru', asal = id, n = 2;
    while (idAda[id]) id = asal + n++;
    idAda[id] = 1; namaAda[kunci] = 1;
    var admin = KONFIG.ADMIN_AWAL.map(function (x) { return x.toUpperCase(); }).indexOf(kunci) >= 0;
    var o = { 'ID': id, 'NAMA': nama.replace(/\s+/g, ' '), 'PERANAN': admin ? 'ADMIN' : 'GURU', 'AKTIF': 'YA', 'KATA LALUAN BARU': kataSementara_(), 'MESTI TUKAR': 'YA' };
    baru.push(h.map(function (k) { return o[k] || ''; }));
  });
  if (baru.length) sh.getRange(sh.getLastRow() + 1, 1, baru.length, h.length).setNumberFormat('@').setValues(baru);
  SpreadsheetApp.getUi().alert(baru.length + ' pengguna ditambah. Kata laluan sementara berada dalam lajur KATA LALUAN BARU; edarkan kepada setiap guru secara peribadi. Lajur itu dikosongkan secara automatik selepas guru log masuk.');
}

function tetapkanTelegram() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  var a = ui.prompt('Token bot Telegram', 'Tampal token daripada @BotFather (contoh 123456:ABC-…):', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var tok = a.getResponseText().trim();
  if (tok) p.setProperty('TELEGRAM_TOKEN', tok);
  var b = ui.prompt('Chat ID kumpulan', 'Tampal Chat ID (biasanya bermula dengan -100…). Biarkan kosong jika belum tahu; guna menu 4.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() === ui.Button.OK && b.getResponseText().trim()) p.setProperty('TELEGRAM_CHAT_ID', b.getResponseText().trim());
  ui.alert('Disimpan. Token disimpan dalam tetapan skrip, bukan dalam Sheet.');
}
function cariChatId() {
  var tok = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  var ui = SpreadsheetApp.getUi();
  if (!tok) { ui.alert('Tetapkan token dahulu (menu 3).'); return; }
  var j = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tok + '/getUpdates', { muteHttpExceptions: true }).getContentText() || '{}');
  var seen = {}, out = [];
  (j.result || []).forEach(function (u) {
    var m = u.message || u.channel_post || u.my_chat_member;
    var c = m && m.chat;
    if (c && !seen[c.id]) { seen[c.id] = 1; out.push(c.id + '  ·  ' + (c.title || c.username || c.first_name || '') + ' (' + c.type + ')'); }
  });
  ui.alert(out.length ? 'Chat yang dijumpai:\n\n' + out.join('\n') + '\n\nSalin Chat ID kumpulan, kemudian guna menu 3 sekali lagi.'
    : 'Tiada chat dijumpai. Tambah bot ke dalam kumpulan, hantar satu mesej dalam kumpulan itu (contoh: /mula), kemudian cuba lagi.');
}
function ujiTelegram() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ok = telegram_('<b>' + KONFIG.NAMA_SISTEM + '</b>: ujian notifikasi berjaya.');
    ui.alert(ok ? 'Mesej ujian dihantar. Semak kumpulan Telegram.' : 'Token atau Chat ID belum ditetapkan (menu 3).');
  } catch (e) { ui.alert('Gagal: ' + e.message); }
}

function semakFolder() {
  var ui = SpreadsheetApp.getUi();
  try {
    var f = folderApp_(), g = folderGambar_();
    var L = bacaSemua_(HELAIAN.LAPORAN), cuba = 0, jumpa = 0;
    for (var i = L.length - 1; i >= 0 && cuba < 5; i--) { var p = teks_(L[i]['GAMBAR 1']); if (!p) continue; cuba++; if (failIkutLaluan_(p)) jumpa++; }
    var induk = g.getParents(), nIn = induk.hasNext() ? induk.next().getName() : '?';
    ui.alert('Folder AppSheet: ' + f.getName() + '\nFolder gambar: ' + nIn + ' / ' + g.getName() + '\nUjian: ' + jumpa + ' daripada ' + cuba + ' gambar terkini dijumpai.\nPDF disimpan dalam ' + KONFIG.FOLDER_PDF + '.');
  } catch (e) { ui.alert(e.message + '\n\nSemak nama folder dalam KONFIG.FOLDER_APP (Code.gs).'); }
}
function semakBaris() {
  var rosak = semuaLaporan_().filter(function (o) { return o.rosak; }).map(function (o) { return 'Baris ' + o.baris + ' (ID ' + o.id + '): TARIKH = "' + o.tarikh + '"'; });
  SpreadsheetApp.getUi().alert(rosak.length ? rosak.length + ' baris dengan TARIKH tidak sah (tidak muncul dalam takwim/statistik bulanan):\n\n' + rosak.slice(0, 40).join('\n') : 'Semua baris LAPORAN mempunyai tarikh yang sah.');
}
function janaPdfTiada() {
  var mula = Date.now(), n = 0, gagal = 0;
  var L = semuaLaporan_().filter(function (o) { return !o.pdf && !o.pautan; });
  for (var i = 0; i < L.length; i++) {
    if (Date.now() - mula > 300000) break; // had masa Apps Script
    try { janaPdf_(L[i].id); n++; } catch (e) { gagal++; }
  }
  SpreadsheetApp.getUi().alert(n + ' PDF dijana' + (gagal ? ', ' + gagal + ' gagal' : '') + '. Baki: ' + (L.length - n - gagal) + (L.length - n - gagal > 0 ? ' (jalankan sekali lagi).' : '.'));
}
function janaSemulaPdfMyDisiplin() {
  var mula = Date.now(), n = 0, gagal = 0;
  var L = semuaLaporan_().filter(function (o) { return /drive\.google\.com/.test(o.pautan); });
  for (var i = 0; i < L.length; i++) {
    if (Date.now() - mula > 300000) break;
    try { janaPdf_(L[i].id); n++; } catch (e) { gagal++; }
  }
  SpreadsheetApp.getUi().alert(n + ' daripada ' + L.length + ' PDF buatan MyDisiplin dijana semula' + (gagal ? ', ' + gagal + ' gagal' : '') + '.' + (n + gagal < L.length ? ' Jalankan sekali lagi untuk baki.' : ''));
}
/** PDF yang dijana MyDisiplin sebelum v1.0.6: tukar lajur PDF daripada laluan /appsheet/... kepada pautan Drive. */
function betulkanPautanPdf() {
  var sh = helaian_(HELAIAN.LAPORAN), h = tajuk_(sh), cp = h.indexOf(PETA.pdf), n = 0;
  semuaLaporan_().forEach(function (o) {
    if (/drive\.google\.com/.test(o.pautan) && /^\/appsheet\//.test(o.pdf)) {
      sh.getRange(o.baris, cp + 1).setValue(o.pautan); n++;
    }
  });
  SpreadsheetApp.getUi().alert(n + ' baris dikemas kini. Lajur PDF kini memegang pautan Drive bagi PDF buatan MyDisiplin; PDF asal AppSheet tidak disentuh.');
}
function tunjukPautan() {
  SpreadsheetApp.getUi().alert('Aplikasi: ' + KONFIG.PAUTAN_APLIKASI + '\nPautan /exec: ' + ScriptApp.getService().getUrl());
}
