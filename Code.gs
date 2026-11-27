/**
 * e-Jawatankuasa MRSM Arau - pelayan (Google Apps Script)
 * Data: Google Sheet "e-Jawatankuasa MRSM Arau (Data)"
 * Tab: STAF, JK, LANTIKAN, PEMBERAT, SEMAKAN, PENGGUNA (dicipta oleh setup)
 */
var SHEET_ID = '1plzCK_iEmpx2fkB3pvUwAQrXcTFAjqOj2jRLivRdt24';
var SESI_SAAT = 21600;           // sesi log masuk 6 jam
var CUBA_MAKS = 5;               // percubaan salah sebelum dikunci
var KUNCI_SAAT = 600;            // dikunci 10 minit
var PERANAN = ['PENTADBIR', 'PENYELARAS', 'PEMBANTU'];
var H_PENGGUNA = ['ID', 'NAMA', 'ID_STAF', 'PERANAN', 'KOD_JK', 'HASH', 'STATUS', 'TUKAR_KATALALUAN', 'DICIPTA', 'LOG_MASUK_AKHIR'];

/* ---------- Jalankan SEKALI dari editor Apps Script ---------- */
function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid() + Utilities.getUuid());
  var sh = ss.getSheetByName('PENGGUNA');
  if (!sh) {
    sh = ss.insertSheet('PENGGUNA');
    sh.getRange(1, 1, 1, H_PENGGUNA.length).setValues([H_PENGGUNA]).setFontWeight('bold')
      .setBackground('#1F4E79').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@');
  }
  if (!cariPengguna_('admin')) {
    var tempPw = 'Arau@2026';
    var h = sha_('mrsm-jk:admin:' + tempPw);
    sh.appendRow(['admin', 'Pentadbir Sistem', '', 'PENTADBIR', '', simpanHash_(h), 'AKTIF', 1, new Date(), '']);
    Logger.log('Akaun admin dicipta. ID: admin  Kata laluan sementara: ' + tempPw + ' (wajib tukar semasa log masuk pertama)');
  } else {
    Logger.log('Akaun admin sudah wujud.');
  }
}

/* ---------- Pintu masuk web app (JSONP) ---------- */
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.cb) || '';
  var out;
  try {
    var body = JSON.parse((e && e.parameter && e.parameter.q) || '{}');
    out = handle_(body);
  } catch (err) {
    out = { ok: false, error: 'pelayan', msg: String(err && err.message || err) };
  }
  var json = JSON.stringify(out);
  if (/^[A-Za-z0-9_]{1,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function handle_(b) {
  switch (b.action) {
    case 'ping': return { ok: true };
    case 'login': return login_(b);
    case 'logout': if (b.token) CacheService.getScriptCache().remove('t_' + b.token); return { ok: true };
  }
  var u = sesi_(b.token);
  if (!u) return { ok: false, error: 'auth' };
  switch (b.action) {
    case 'me': return { ok: true, user: awam_(u) };
    case 'setpw': return setpw_(u, b);
  }
  if (u.mustChange) return { ok: false, error: 'tukarpw' };
  switch (b.action) {
    case 'dashboard': return dashboard_(u);
    case 'users': return wajibAdmin_(u) || senaraiPengguna_();
    case 'saveuser': return wajibAdmin_(u) || simpanPengguna_(u, b);
    case 'resetpw': return wajibAdmin_(u) || resetPw_(b);
  }
  return { ok: false, error: 'bad' };
}

/* ---------- Log masuk & sesi ---------- */
function login_(b) {
  var id = String(b.id || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{1,30}$/.test(id) || !/^[a-f0-9]{64}$/.test(String(b.h || ''))) return { ok: false, error: 'salah' };
  var cache = CacheService.getScriptCache();
  var kk = 'cuba_' + id, cuba = Number(cache.get(kk) || 0);
  if (cuba >= CUBA_MAKS) return { ok: false, error: 'locked' };
  var p = cariPengguna_(id);
  if (!p || p.STATUS !== 'AKTIF' || p.HASH !== simpanHash_(b.h)) {
    cache.put(kk, String(cuba + 1), KUNCI_SAAT);
    return { ok: false, error: 'salah' };
  }
  cache.remove(kk);
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var u = { id: p.ID, nama: p.NAMA, idStaf: p.ID_STAF, peranan: p.PERANAN, kodJk: senaraiKod_(p.KOD_JK), mustChange: String(p.TUKAR_KATALALUAN) === '1' || p.TUKAR_KATALALUAN === true };
  cache.put('t_' + token, JSON.stringify(u), SESI_SAAT);
  kemaskiniPengguna_(id, { LOG_MASUK_AKHIR: new Date() });
  return { ok: true, token: token, user: awam_(u) };
}

function sesi_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(token)) return null;
  var c = CacheService.getScriptCache(), v = c.get('t_' + token);
  if (!v) return null;
  c.put('t_' + token, v, SESI_SAAT); // panjangkan sesi bila aktif
  var u = JSON.parse(v); u.token = token; return u;
}

function setpw_(u, b) {
  if (!/^[a-f0-9]{64}$/.test(String(b.h || ''))) return { ok: false, error: 'bad' };
  kemaskiniPengguna_(u.id, { HASH: simpanHash_(b.h), TUKAR_KATALALUAN: 0 });
  u.mustChange = false;
  var t = u.token; delete u.token;
  CacheService.getScriptCache().put('t_' + t, JSON.stringify(u), SESI_SAAT);
  return { ok: true, user: awam_(u) };
}

function awam_(u) { return { id: u.id, nama: u.nama, idStaf: u.idStaf, peranan: u.peranan, kodJk: u.kodJk, mustChange: !!u.mustChange }; }
function wajibAdmin_(u) { return u.peranan === 'PENTADBIR' ? null : { ok: false, error: 'akses' }; }

/* ---------- Dashboard ---------- */
function dashboard_(u) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var staf = jadual_(ss, 'STAF'), jk = jadual_(ss, 'JK'), lant = jadual_(ss, 'LANTIKAN'), pb = jadual_(ss, 'PEMBERAT');
  var nilai = {}; pb.forEach(function (r) { nilai[String(r.JAWATAN).trim()] = Number(r.NILAI) || 0; });
  var jkMap = {}; jk.forEach(function (r) { jkMap[r.KOD_JK] = r; });
  var semua = u.peranan === 'PENTADBIR';
  var skop = semua ? null : {}; if (!semua) u.kodJk.forEach(function (k) { skop[k] = true; });

  var perStaf = {}, perJk = {}, perJaw = {}, perInduk = {};
  staf.forEach(function (s) { if (s.STATUS === 'AKTIF') perStaf[s.ID_STAF] = { id: s.ID_STAF, nama: s.NAMA, jabatan: s.JABATAN, jawatan: s.JAWATAN, susunan: Number(s.SUSUNAN) || 9999, bil: 0, pemberat: 0, semak: !!s.CATATAN }; });
  jk.forEach(function (r) {
    if (skop && !skop[r.KOD_JK]) return;
    var i = r.JK_INDUK || 'LAIN-LAIN';
    perInduk[i] = perInduk[i] || { induk: i, jk: 0, lantikan: 0 }; perInduk[i].jk++;
    perJk[r.KOD_JK] = { kod: r.KOD_JK, nama: r.NAMA_JK, induk: i, ahli: 0, pengerusi: '', penyelaras: '', setiausaha: '' };
  });
  var jumPemberat = 0, bilLant = 0;
  lant.forEach(function (l) {
    var w = nilai[String(l.JAWATAN).trim()] || 0;
    var s = perStaf[l.ID_STAF];
    if (s) { s.bil++; s.pemberat += w; }           // beban tugas kira semua JK
    if (skop && !skop[l.KOD_JK]) return;
    bilLant++; jumPemberat += w;
    var j = perJk[l.KOD_JK];
    if (j) {
      j.ahli++; perInduk[j.induk].lantikan++;
      var jw = String(l.JAWATAN);
      if (jw.indexOf('PENGERUSI') === 0 && l.JAWATAN_ASAL === 'PENGERUSI' && !j.pengerusi) j.pengerusi = l.NAMA_STAF;
      if (jw.indexOf('PENYELARAS') === 0 && !j.penyelaras) j.penyelaras = l.NAMA_STAF;
      if (jw === 'SETIAUSAHA' && !j.setiausaha) j.setiausaha = l.NAMA_STAF;
    }
    perJaw[l.JAWATAN] = (perJaw[l.JAWATAN] || 0) + 1;
  });
  var beban = Object.keys(perStaf).map(function (k) { return perStaf[k]; });
  if (skop) { // penyelaras: hanya ahli JK yang dijaga
    var ahli = {}; lant.forEach(function (l) { if (skop[l.KOD_JK]) ahli[l.ID_STAF] = true; });
    beban = beban.filter(function (s) { return ahli[s.id]; });
  }
  beban.sort(function (a, b) { return b.pemberat - a.pemberat || a.nama.localeCompare(b.nama); });
  var berTugas = beban.filter(function (s) { return s.bil > 0; });
  var purata = berTugas.length ? berTugas.reduce(function (t, s) { return t + s.pemberat; }, 0) / berTugas.length : 0;
  return {
    ok: true, user: awam_(u), dikemaskini: new Date().toISOString(),
    kpi: {
      staf: beban.length, jk: Object.keys(perJk).length, lantikan: bilLant, pemberat: jumPemberat,
      purata: Math.round(purata * 10) / 10, tiadaTugas: beban.length - berTugas.length,
      semakan: semua ? staf.filter(function (s) { return s.CATATAN; }).length : 0
    },
    induk: Object.keys(perInduk).map(function (k) { return perInduk[k]; }).sort(function (a, b) { return b.jk - a.jk; }),
    jawatan: Object.keys(perJaw).map(function (k) { return { jawatan: k, bil: perJaw[k], nilai: nilai[k] || 0 }; }).sort(function (a, b) { return b.bil - a.bil; }),
    beban: beban,
    jkSaya: semua ? [] : Object.keys(perJk).map(function (k) { return perJk[k]; })
  };
}

/* ---------- Urus pengguna (pentadbir) ---------- */
function senaraiPengguna_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rows = jadual_(ss, 'PENGGUNA').map(function (p) {
    return { id: p.ID, nama: p.NAMA, idStaf: p.ID_STAF, peranan: p.PERANAN, kodJk: senaraiKod_(p.KOD_JK), status: p.STATUS, akhir: p.LOG_MASUK_AKHIR ? new Date(p.LOG_MASUK_AKHIR).toISOString() : '' };
  });
  var staf = jadual_(ss, 'STAF').filter(function (s) { return s.STATUS === 'AKTIF'; }).map(function (s) { return { id: s.ID_STAF, nama: s.NAMA }; });
  var jk = jadual_(ss, 'JK').map(function (j) { return { kod: j.KOD_JK, nama: j.NAMA_JK }; });
  return { ok: true, users: rows, staf: staf, jk: jk };
}

function simpanPengguna_(admin, b) {
  var d = b.data || {};
  var id = String(d.id || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{1,30}$/.test(id)) return { ok: false, error: 'ID tidak sah (huruf kecil, nombor, . _ - sahaja)' };
  if (PERANAN.indexOf(d.peranan) < 0) return { ok: false, error: 'Peranan tidak sah' };
  var status = d.status === 'TIDAK AKTIF' ? 'TIDAK AKTIF' : 'AKTIF';
  if (id === admin.id && (d.peranan !== 'PENTADBIR' || status !== 'AKTIF')) return { ok: false, error: 'Tidak boleh buang akses pentadbir sendiri' };
  var kodList = senaraiKod_(d.kodJk), sah = {};
  jadual_(SpreadsheetApp.openById(SHEET_ID), 'JK').forEach(function (j) { sah[j.KOD_JK] = true; });
  var salah = kodList.filter(function (k) { return !sah[k]; });
  if (salah.length) return { ok: false, error: 'Kod JK tidak wujud: ' + salah.join(', ') };
  if (d.peranan !== 'PENTADBIR' && !kodList.length) return { ok: false, error: 'Penyelaras/Pembantu mesti ada sekurang-kurangnya satu kod JK' };
  var kod = kodList.join(',');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var ada = cariPengguna_(id);
    if (b.baru) {
      if (ada) return { ok: false, error: 'ID sudah digunakan' };
      if (!/^[a-f0-9]{64}$/.test(String(d.h || ''))) return { ok: false, error: 'Kata laluan sementara diperlukan' };
      SpreadsheetApp.openById(SHEET_ID).getSheetByName('PENGGUNA')
        .appendRow([id, String(d.nama || '').trim(), d.idStaf || '', d.peranan, kod, simpanHash_(d.h), status, 1, new Date(), '']);
    } else {
      if (!ada) return { ok: false, error: 'Pengguna tidak dijumpai' };
      kemaskiniPengguna_(id, { NAMA: String(d.nama || '').trim(), ID_STAF: d.idStaf || '', PERANAN: d.peranan, KOD_JK: kod, STATUS: status });
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function resetPw_(b) {
  var id = String(b.id || '').toLowerCase();
  if (!cariPengguna_(id) || !/^[a-f0-9]{64}$/.test(String(b.h || ''))) return { ok: false, error: 'bad' };
  kemaskiniPengguna_(id, { HASH: simpanHash_(b.h), TUKAR_KATALALUAN: 1 });
  CacheService.getScriptCache().remove('cuba_' + id);
  return { ok: true };
}

/* ---------- Pembantu ---------- */
function jadual_(ss, nama) {
  var sh = ss.getSheetByName(nama); if (!sh) return [];
  var v = sh.getDataRange().getValues(); if (v.length < 2) return [];
  var h = v[0].map(function (x) { return String(x).trim(); });
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i].join('') === '') continue;
    var o = {}; for (var c = 0; c < h.length; c++) o[h[c]] = typeof v[i][c] === 'string' ? v[i][c].trim() : v[i][c];
    out.push(o);
  }
  return out;
}
function cariPengguna_(id) {
  var r = jadual_(SpreadsheetApp.openById(SHEET_ID), 'PENGGUNA');
  for (var i = 0; i < r.length; i++) if (String(r[i].ID).toLowerCase() === id) return r[i];
  return null;
}
function kemaskiniPengguna_(id, ubah) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('PENGGUNA');
  var v = sh.getDataRange().getValues(), h = v[0];
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() !== id) continue;
    Object.keys(ubah).forEach(function (k) { var c = h.indexOf(k); if (c >= 0) sh.getRange(i + 1, c + 1).setValue(ubah[k]); });
    return true;
  }
  return false;
}
function senaraiKod_(s) {
  var a = Array.isArray(s) ? s : String(s || '').split(/[,\s]+/);
  return a.map(function (x) { return String(x).trim().toUpperCase(); }).filter(function (x) { return /^[A-Z]\d{2}$/.test(x); });
}
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function simpanHash_(h) { return sha_(PropertiesService.getScriptProperties().getProperty('SALT') + ':' + h); }
