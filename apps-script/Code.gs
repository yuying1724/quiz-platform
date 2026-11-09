/**
 * 測驗平台 — Google Apps Script 後端（JSON API）
 *
 * 資料：同一個試算表內的「題庫」與「作答紀錄」兩張工作表。
 * 前端：獨立的 PWA（web/index.html），以 fetch 呼叫此 API。
 * 驗證：每台裝置一組「裝置授權碼」（試算表選單「測驗平台 → 新增裝置授權碼」產生，工作表只存雜湊）；
 *       Script Properties 的 API_TOKEN 只給維運用（匯入、管理裝置），不放進前端。
 * 部署：網頁應用程式 → 執行身分「我」、存取權「任何人」（靠 token 擋）。
 *
 * API
 *   GET  ?token=…&action=meta
 *   GET  ?token=…&action=questions&subject=…&chapters=1,2&source=all|練習題|模擬試題&onlyWrong=0|1
 *   GET  ?token=…&action=history&subject=…
 *   GET  ?token=…&action=sessionDetail&subject=…&session=…
 *   GET  ?token=…&action=page&subject=…&page=12        （教材單頁原文）
 *   GET  ?token=…&action=search&subject=…&q=關鍵字      （教材全文搜尋）
 *   POST body(JSON): { token, action:'importPages', rows:[[科目,頁碼,章節,原文],…] }
 *   POST body(JSON): { token, action:'setKeys', items:[{id, keys:[{page,text}]},…] }
 *   POST body(JSON): { token, action:'answer',  rec:{…} }
 *   POST body(JSON): { token, action:'answers', recs:[{…}] }
 *   POST body(JSON): { token(admin), action:'newDevice', name }        → { code, setupUrl }
 *   POST body(JSON): { token(admin), action:'listDevices' } / { action:'setDevice', hash, enabled }
 */

var SHEET_BANK = '題庫';
var SHEET_LOG = '作答紀錄';
var SHEET_TEXT = '教材';
var SHEET_DEV = '裝置';
var DEV_COLS = ['授權碼雜湊', '裝置名稱', '建立時間', '最後使用', '啟用'];
var APP_URL = 'https://yuying1724.github.io/quiz-platform/web/';

var BANK_COLS = ['題號', '科目', '來源', '章節', '學習目標', '題目',
  '選項1', '選項2', '選項3', '選項4', '選項5', '答案', '教材頁碼', '啟用', '關鍵句'];
var TEXT_COLS = ['科目', '頁碼', '章節', '原文'];
var LOG_COLS = ['時間', '場次', '科目', '章節', '題號', '來源', '我的答案', '正確答案', '是否正確', '模式'];

// ---------- 一次性設定 ----------

/** 第一次使用執行：整理工作表名稱並建立作答紀錄表。 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bank = ss.getSheetByName(SHEET_BANK);
  if (!bank) {
    var first = ss.getSheets()[0];
    var header = first.getRange(1, 1, 1, BANK_COLS.length).getValues()[0];
    if (header[0] !== BANK_COLS[0]) {
      throw new Error('找不到「題庫」工作表，且第一張工作表的標題列不是題庫格式。');
    }
    first.setName(SHEET_BANK);
    bank = first;
  }
  bank.setFrozenRows(1);

  var log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]);
    log.setFrozenRows(1);
    log.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  return '設定完成';
}

/** 產生 API_TOKEN 並存入 Script Properties。執行後到「執行記錄」複製 token。 */
function generateToken() {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', token);
  Logger.log('API_TOKEN = ' + token);
  return token;
}

function getToken_() {
  return PropertiesService.getScriptProperties().getProperty('API_TOKEN') || '';
}

// ---------- 裝置授權碼 ----------
// 前端只持有「裝置授權碼」（每台裝置一組，輸入一次後存在該裝置）；試算表只存雜湊，可個別停用。
// Script Properties 的 API_TOKEN 保留給維運（匯入題庫、管理裝置），不放進任何前端檔案。

var ADMIN_ACTIONS = ['import', 'importPages', 'setKeys', 'newDevice', 'listDevices', 'setDevice', 'rotateToken'];

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function normCode_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

function devSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_DEV);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DEV);
    sh.getRange(1, 1, 1, DEV_COLS.length).setValues([DEV_COLS]);
    sh.setFrozenRows(1);
    sh.getRange('C:D').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sh;
}

/** 回傳 { ok, admin, row } */
function checkAuth_(token) {
  if (!token) return { ok: false };
  var admin = getToken_();
  if (admin && token === admin) return { ok: true, admin: true };
  var code = normCode_(token);
  if (code.length < 16) return { ok: false };
  var h = hash_(code);
  var sh = devSheet_();
  if (sh.getLastRow() < 2) return { ok: false };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, DEV_COLS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === h) {
      if (String(rows[i][4]).toUpperCase() === 'N') return { ok: false };
      // 最後使用時間：超過 10 分鐘才更新，避免每次請求都寫表
      var last = rows[i][3] instanceof Date ? rows[i][3].getTime() : 0;
      if (Date.now() - last > 600000) sh.getRange(i + 2, 4).setValue(new Date());
      return { ok: true, admin: false, row: i + 2 };
    }
  }
  return { ok: false };
}

/** 產生一組新的裝置授權碼（明碼只回傳這一次），格式 XXXXX-XXXXX-XXXXX-XXXXX */
function newDevice(name) {
  var alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var raw = '';
  for (var i = 0; i < 20; i++) raw += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  var code = raw.match(/.{5}/g).join('-');
  var sh = devSheet_();
  sh.appendRow([hash_(raw), String(name || '未命名裝置'), new Date(), '', 'Y']);
  return { code: code, name: String(name || '未命名裝置'), setupUrl: APP_URL + '#code=' + code };
}

function listDevices() {
  var sh = devSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, DEV_COLS.length).getValues().map(function (r) {
    return { hash: String(r[0]).slice(0, 8), name: String(r[1]), created: r[2] instanceof Date ? r[2].getTime() : null,
      lastUsed: r[3] instanceof Date ? r[3].getTime() : null, enabled: String(r[4]).toUpperCase() !== 'N' };
  });
}

/** 依雜湊前 8 碼啟用/停用裝置 */
function setDevice(hashPrefix, enabled) {
  var sh = devSheet_();
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).indexOf(String(hashPrefix)) === 0) { sh.getRange(i + 2, 5).setValue(enabled ? 'Y' : 'N'); return { updated: 1 }; }
  }
  return { updated: 0 };
}

// ---------- 試算表選單（開啟試算表時出現「測驗平台」選單） ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('測驗平台')
    .addItem('新增裝置授權碼', 'menuNewDevice')
    .addItem('查看裝置清單', 'menuListDevices')
    .addToUi();
}

function menuNewDevice() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('新增裝置授權碼', '幫這台裝置取個名字（例如：手機、iPad、公司電腦）：', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var d = newDevice(r.getResponseText().trim() || '未命名裝置');
  var qr = 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=' + encodeURIComponent(d.setupUrl);
  var html = '<div style="font:15px/1.6 sans-serif;padding:8px 12px">'
    + '<p>裝置「<b>' + d.name + '</b>」的授權碼（只會顯示這一次）：</p>'
    + '<p style="font:700 22px monospace;letter-spacing:1px;background:#f4efdf;padding:10px;border-radius:8px">' + d.code + '</p>'
    + '<p>新裝置的設定方式擇一：</p><ol><li>用相機掃下方 QR code，開啟後自動完成設定</li>'
    + '<li>或在新裝置開啟 App → 設定 → 貼上授權碼</li></ol>'
    + '<p style="text-align:center"><img src="' + qr + '" width="220" height="220" alt="QR"></p>'
    + '<p style="font-size:13px;color:#666">連結：<br>' + d.setupUrl + '</p></div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(560), '新增裝置授權碼');
}

function menuListDevices() {
  var list = listDevices();
  var rows = list.map(function (d) {
    var f = function (t) { return t ? Utilities.formatDate(new Date(t), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '—'; };
    return '<tr><td>' + d.name + '</td><td>' + f(d.created) + '</td><td>' + f(d.lastUsed) + '</td><td>' + (d.enabled ? '啟用' : '已停用') + '</td></tr>';
  }).join('');
  var html = '<div style="font:14px/1.6 sans-serif;padding:8px 12px"><table border="0" cellpadding="6" style="border-collapse:collapse;width:100%">'
    + '<tr style="background:#f4efdf"><th align="left">裝置</th><th align="left">建立</th><th align="left">最後使用</th><th align="left">狀態</th></tr>' + rows + '</table>'
    + '<p style="font-size:13px;color:#666">要停用某台裝置，直接在「裝置」工作表把該列的「啟用」改成 N。</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(360), '裝置清單');
}

// ---------- HTTP 入口 ----------

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {}, null);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  return handle_(body, body);
}

function handle_(p, body) {
  try {
    var auth = checkAuth_(String(p.token || ''));
    if (!auth.ok) return json_({ ok: false, error: 'unauthorized' });

    var action = String(p.action || '');
    var data;
    // 管理用動作（匯入、改題庫、管理裝置）只接受 Script Properties 的 API_TOKEN
    if (ADMIN_ACTIONS.indexOf(action) >= 0 && !auth.admin) return json_({ ok: false, error: 'unauthorized' });
    switch (action) {
      case 'newDevice':
        data = newDevice(body.name); break;
      case 'listDevices':
        data = listDevices(); break;
      case 'setDevice':
        data = setDevice(body.hash, body.enabled); break;
      case 'rotateToken':
        data = { token: generateToken() }; break;
      case 'meta':
        data = getMeta(); break;
      case 'questions':
        data = getQuestions({
          subject: p.subject,
          chapters: p.chapters ? String(p.chapters).split(',').filter(String).map(Number) : [],
          source: p.source || 'all',
          onlyWrong: String(p.onlyWrong) === '1' || p.onlyWrong === true,
          shuffle: String(p.shuffle) === '1' || p.shuffle === true
        }); break;
      case 'history':
        data = getHistory(p.subject); break;
      case 'sessionDetail':
        data = getSessionDetail(p.subject, p.session); break;
      case 'page':
        data = getPage(p.subject, p.page); break;
      case 'search':
        data = searchText(p.subject, p.q); break;
      case 'importPages':
        data = importPages(body.rows || []); break;
      case 'setKeys':
        data = setKeys(body.items || []); break;
      case 'answer':
        data = saveAnswers([body.rec]); break;
      case 'answers':
        data = saveAnswers(body.recs || []); break;
      case 'ping':
        data = { time: new Date().toISOString() }; break;
      case 'import':
        data = importQuestions(body.rows || []); break;
      case 'deleteSession':
        data = deleteSession(body.subject, body.session); break;
      default:
        return json_({ ok: false, error: 'unknown action: ' + action });
    }
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 讀資料 ----------

function readBank_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BANK);
  if (!sh) throw new Error('找不到「題庫」工作表，請先執行 setup()。');
  var values = sh.getDataRange().getValues();
  var header = values.shift().map(String);
  var idx = {};
  header.forEach(function (h, i) { idx[h.trim()] = i; });

  return values.filter(function (r) { return String(r[idx['題號']]).trim() !== ''; })
    .map(function (r) {
      var options = [];
      for (var k = 1; k <= 5; k++) {
        var v = r[idx['選項' + k]];
        if (v !== '' && v !== null && v !== undefined) options.push(String(v));
      }
      return {
        id: String(r[idx['題號']]).trim(),
        subject: String(r[idx['科目']]).trim(),
        source: String(r[idx['來源']]).trim(),
        chapter: Number(r[idx['章節']]),
        objective: String(r[idx['學習目標']] || ''),
        stem: String(r[idx['題目']]),
        options: options,
        answer: Number(r[idx['答案']]),
        page: String(r[idx['教材頁碼']] || ''),
        enabled: String(r[idx['啟用']] || 'Y').toUpperCase() !== 'N',
        keys: parseKeys_(idx['關鍵句'] !== undefined ? r[idx['關鍵句']] : '')
      };
    })
    .filter(function (q) { return q.enabled; });
}

function readLog_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_COLS.length).getValues();
  return values.map(function (r) {
    return {
      time: r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime(),
      session: String(r[1]), subject: String(r[2]), chapter: Number(r[3]),
      id: String(r[4]), source: String(r[5]), chosen: Number(r[6]),
      answer: Number(r[7]), correct: r[8] === true || String(r[8]).toUpperCase() === 'TRUE',
      mode: String(r[9])
    };
  });
}

// ---------- 業務邏輯 ----------

/** 首頁資料：科目、章節、各章題數與最近正確率、錯題數 */
function getMeta() {
  var bank = readBank_();
  var log = readLog_();

  var latest = {};
  log.forEach(function (l) {
    var key = l.subject + '|' + l.id;
    if (!latest[key] || l.time > latest[key].time) latest[key] = l;
  });

  var subjects = {};
  bank.forEach(function (q) {
    var s = subjects[q.subject] || (subjects[q.subject] = { name: q.subject, chapters: {}, total: 0, wrong: 0 });
    var c = s.chapters[q.chapter] || (s.chapters[q.chapter] = { chapter: q.chapter, total: 0, practice: 0, mock: 0, answered: 0, correct: 0, wrong: 0 });
    s.total++; c.total++;
    if (q.source === '練習題') c.practice++; else c.mock++;
    var l = latest[q.subject + '|' + q.id];
    if (l) {
      c.answered++;
      if (l.correct) c.correct++; else { c.wrong++; s.wrong++; }
    }
  });

  return {
    subjects: Object.keys(subjects).map(function (k) {
      var s = subjects[k];
      s.chapters = Object.keys(s.chapters).map(function (c) { return s.chapters[c]; })
        .sort(function (a, b) { return a.chapter - b.chapter; });
      return s;
    }),
    totalAnswered: log.length
  };
}

function getQuestions(opts) {
  opts = opts || {};
  var bank = readBank_().filter(function (q) { return q.subject === opts.subject; });
  if (opts.chapters && opts.chapters.length) {
    var set = {};
    opts.chapters.forEach(function (c) { set[Number(c)] = true; });
    bank = bank.filter(function (q) { return set[q.chapter]; });
  }
  if (opts.source && opts.source !== 'all') {
    bank = bank.filter(function (q) { return q.source === opts.source; });
  }
  if (opts.onlyWrong) {
    var latest = {};
    readLog_().filter(function (l) { return l.subject === opts.subject; }).forEach(function (l) {
      if (!latest[l.id] || l.time > latest[l.id].time) latest[l.id] = l;
    });
    bank = bank.filter(function (q) { return latest[q.id] && !latest[q.id].correct; });
  }
  if (opts.shuffle) {
    for (var i = bank.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = bank[i]; bank[i] = bank[j]; bank[j] = t;
    }
  }
  return bank;
}

/** 儲存多筆作答。rec = { session, subject, chapter, id, source, chosen, answer, mode } */
function saveAnswers(recs) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh) throw new Error('找不到「作答紀錄」工作表，請先執行 setup()。');
  var rows = (recs || []).filter(Boolean).map(function (rec) {
    return [new Date(), String(rec.session || ''), String(rec.subject || ''), Number(rec.chapter), String(rec.id || ''),
      String(rec.source || ''), Number(rec.chosen), Number(rec.answer), Number(rec.chosen) === Number(rec.answer), String(rec.mode || '')];
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_COLS.length).setValues(rows);
  return { saved: rows.length };
}

/**
 * 批次匯入題目到「題庫」。rows = [[題號,科目,來源,章節,學習目標,題目,選項1..5,答案,教材頁碼,啟用], ...]
 * 題號已存在的列會略過，回傳 { added, skipped }。
 */
function importQuestions(rows) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BANK);
  if (!sh) throw new Error('找不到「題庫」工作表，請先執行 setup()。');
  var existing = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[String(r[0]).trim()] = true; });
  }
  var toAdd = [], skipped = [];
  rows.forEach(function (r) {
    if (!r || r.length < 14) return;
    var id = String(r[0]).trim();
    if (!id || existing[id]) { skipped.push(id); return; }
    existing[id] = true;
    var row = r.slice(0, BANK_COLS.length).map(function (v) { return v === null || v === undefined ? '' : v; });
    while (row.length < BANK_COLS.length) row.push('');
    row[3] = Number(row[3]); row[11] = Number(row[11]);
    toAdd.push(row);
  });
  if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, BANK_COLS.length).setValues(toAdd);
  return { added: toAdd.length, skipped: skipped };
}

/** 刪除某科目某一場次的所有作答紀錄，回傳 { deleted } */
function deleteSession(subject, session) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return { deleted: 0 };
  if (!session) throw new Error('缺少場次');
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_COLS.length).getValues();
  var rows = [];
  values.forEach(function (r, i) {
    if (String(r[1]) === String(session) && (!subject || String(r[2]) === String(subject))) rows.push(i + 2);
  });
  rows.reverse().forEach(function (row) { sh.deleteRow(row); });
  return { deleted: rows.length };
}

// ---------- 教材原文 ----------

function parseKeys_(v) {
  if (!v) return [];
  try { var a = JSON.parse(String(v)); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}

function textSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TEXT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_TEXT);
    sh.getRange(1, 1, 1, TEXT_COLS.length).setValues([TEXT_COLS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 讀整本教材：[{subject,page,chapter,text}] */
function readText_(subject) {
  var sh = textSheet_();
  if (sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, TEXT_COLS.length).getValues();
  return values.filter(function (r) { return r[1] !== '' && (!subject || String(r[0]) === String(subject)); })
    .map(function (r) { return { subject: String(r[0]), page: Number(r[1]), chapter: Number(r[2]) || null, text: String(r[3]) }; });
}

/** 取單頁原文（頁碼可為字串或數字；找不到該頁時回傳最接近的前一頁並標 alias） */
function getPage(subject, page) {
  var n = Number(page);
  if (!n) throw new Error('缺少頁碼');
  var all = readText_(subject);
  var hit = all.filter(function (t) { return t.page === n; })[0];
  var alias = false;
  if (!hit) {
    var prev = all.filter(function (t) { return t.page < n; }).sort(function (a, b) { return b.page - a.page; })[0];
    if (!prev) throw new Error('沒有第 ' + n + ' 頁的原文');
    hit = prev; alias = true;
  }
  var pages = all.map(function (t) { return t.page; }).sort(function (a, b) { return a - b; });
  return { page: hit.page, requested: n, alias: alias, chapter: hit.chapter, text: hit.text,
    prev: pages.filter(function (p) { return p < hit.page; }).pop() || null,
    next: pages.filter(function (p) { return p > hit.page; })[0] || null };
}

/** 全文關鍵字搜尋：回傳每頁的命中片段（最多 50 頁） */
function searchText(subject, q) {
  q = String(q || '').trim();
  if (q.length < 2) throw new Error('請輸入至少 2 個字');
  var terms = q.split(/\s+/).filter(String);
  var out = [];
  readText_(subject).forEach(function (t) {
    var txt = t.text;
    if (!terms.every(function (k) { return txt.indexOf(k) >= 0; })) return;
    var i = txt.indexOf(terms[0]);
    var s = Math.max(0, i - 40), e = Math.min(txt.length, i + terms[0].length + 80);
    var count = txt.split(terms[0]).length - 1;
    out.push({ page: t.page, chapter: t.chapter, count: count, snippet: (s > 0 ? '…' : '') + txt.slice(s, e).replace(/\n/g, ' ') + (e < txt.length ? '…' : '') });
  });
  out.sort(function (a, b) { return b.count - a.count || a.page - b.page; });
  return { q: q, total: out.length, results: out.slice(0, 50) };
}

/** 批次匯入教材：rows = [[科目,頁碼,章節,原文], ...]；同科目同頁碼已存在則覆蓋 */
function importPages(rows) {
  var sh = textSheet_();
  var existing = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r, i) { existing[String(r[0]) + '|' + String(r[1])] = i + 2; });
  }
  var toAdd = [], updated = 0;
  rows.forEach(function (r) {
    if (!r || r.length < 4) return;
    var row = [String(r[0]), Number(r[1]), Number(r[2]) || '', String(r[3])];
    var key = row[0] + '|' + row[1];
    if (existing[key]) { sh.getRange(existing[key], 1, 1, 4).setValues([row]); updated++; }
    else toAdd.push(row);
  });
  if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 4).setValues(toAdd);
  return { added: toAdd.length, updated: updated };
}

/** 寫入題庫「關鍵句」欄：items = [{id, keys:[{page,text}]}] */
function setKeys(items) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BANK);
  if (!sh) throw new Error('找不到「題庫」工作表');
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var col = header.indexOf('關鍵句') + 1;
  if (!col) { col = header.length + 1; sh.getRange(1, col).setValue('關鍵句'); }
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  var n = 0, missing = [];
  items.forEach(function (it) {
    var i = ids.indexOf(String(it.id).trim());
    if (i < 0) { missing.push(it.id); return; }
    sh.getRange(i + 2, col).setValue(JSON.stringify(it.keys || []));
    n++;
  });
  return { updated: n, missing: missing };
}

/** 某一場次的逐題明細：題目、選項、我的答案、正確答案、對錯、教材頁碼（依作答順序） */
function getSessionDetail(subject, session) {
  if (!session) throw new Error('缺少場次');
  var bank = {};
  readBank_().forEach(function (q) { bank[q.subject + '|' + q.id] = q; });
  var items = readLog_().filter(function (l) {
    return String(l.session) === String(session) && (!subject || l.subject === subject);
  }).sort(function (a, b) { return a.time - b.time; }).map(function (l) {
    var q = bank[l.subject + '|' + l.id] || {};
    return {
      id: l.id, chapter: l.chapter, source: l.source, mode: l.mode, time: l.time,
      chosen: l.chosen, answer: l.answer, correct: l.correct,
      stem: q.stem || '（題目已不在題庫中）', options: q.options || [], page: q.page || '', objective: q.objective || '', keys: q.keys || []
    };
  });
  return { session: String(session), total: items.length,
    correct: items.filter(function (i) { return i.correct; }).length, items: items };
}

/** 某科目的歷次場次摘要（最近 20 場） */
function getHistory(subject) {
  var log = readLog_().filter(function (l) { return l.subject === subject; });
  var sessions = {};
  log.forEach(function (l) {
    var s = sessions[l.session] || (sessions[l.session] = { session: l.session, time: l.time, total: 0, correct: 0, chapters: {}, mode: l.mode });
    s.total++; if (l.correct) s.correct++;
    s.chapters[l.chapter] = true;
    if (l.time < s.time) s.time = l.time;
  });
  return Object.keys(sessions).map(function (k) {
    var s = sessions[k];
    s.chapters = Object.keys(s.chapters).map(Number).sort(function (a, b) { return a - b; });
    return s;
  }).sort(function (a, b) { return b.time - a.time; }).slice(0, 20);
}
