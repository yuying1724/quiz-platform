/**
 * 測驗平台 — Google Apps Script 後端（JSON API）
 *
 * 資料：同一個試算表內的「題庫」與「作答紀錄」兩張工作表。
 * 前端：獨立的 PWA（web/index.html），以 fetch 呼叫此 API。
 * 驗證：Script Properties 的 API_TOKEN（執行 generateToken() 產生）。
 * 部署：網頁應用程式 → 執行身分「我」、存取權「任何人」（靠 token 擋）。
 *
 * API
 *   GET  ?token=…&action=meta
 *   GET  ?token=…&action=questions&subject=…&chapters=1,2&source=all|練習題|模擬試題&onlyWrong=0|1
 *   GET  ?token=…&action=history&subject=…
 *   POST body(JSON): { token, action:'answer',  rec:{…} }
 *   POST body(JSON): { token, action:'answers', recs:[{…}] }
 *   POST body(JSON): { token, action:'deleteSession', subject, session }
 */

var SHEET_BANK = '題庫';
var SHEET_LOG = '作答紀錄';

var BANK_COLS = ['題號', '科目', '來源', '章節', '學習目標', '題目',
  '選項1', '選項2', '選項3', '選項4', '選項5', '答案', '教材頁碼', '啟用'];
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
    var token = getToken_();
    if (!token) return json_({ ok: false, error: '尚未設定 API_TOKEN，請先執行 generateToken()' });
    if (String(p.token || '') !== token) return json_({ ok: false, error: 'unauthorized' });

    var action = String(p.action || '');
    var data;
    switch (action) {
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
        enabled: String(r[idx['啟用']] || 'Y').toUpperCase() !== 'N'
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
    if (!r || r.length < BANK_COLS.length) return;
    var id = String(r[0]).trim();
    if (!id || existing[id]) { skipped.push(id); return; }
    existing[id] = true;
    var row = r.slice(0, BANK_COLS.length).map(function (v) { return v === null || v === undefined ? '' : v; });
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
