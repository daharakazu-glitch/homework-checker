// 提出物チェッカー（Google Apps Script 用・1ファイル版）
// このファイルの中身をすべてコピーして、script.google.com のプロジェクトに貼り付けてください。
// 設定手順: https://daharakazu-glitch.github.io/homework-checker/setup.html
// ※ tools/build-gas.js で gas/server.gs と index.html から自動生成しています。直接編集しないでください。

/**
 * 提出物チェッカー（英語課題・複数クラス対応）
 * ウェブアプリとしてデプロイすると、同じGoogleアカウントならどの端末からでも同じデータを使えます。
 * 「ウェブアプリにアクセスしているユーザー」として実行するようにデプロイすれば、1つのURLを複数の先生で使えます。
 * データは先生ごとに、その先生のドライブに自動で作成するスプレッドシート「提出物チェッカー データ」に保存し、
 * 保存先などの設定は先生ごとのユーザープロパティに記録します。メールもその先生のGmailから送ります。
 * （スプレッドシートにバインドした場合は、そのスプレッドシートに保存します。）
 *
 * シート構成
 *   設定   : A列=クラス名, B列=人数（2行目から最大10クラス）
 *   提出物 : ID / 提出物名 / 対象クラス（「、」区切り） / 伝えた日 / 締切 / 作成日時
 *   クラス名のシート : 1行目=見出し, 2行目(非表示)=提出物ID, 3行目以降=出席番号ごとの記録（✓＝提出）
 */

const SHEET_SETTINGS = '設定';
const SHEET_TASKS = '提出物';
const ID_MARK = '#ID';
const MAX_CLASSES = 10;
const MAX_STUDENTS = 45;
const TASK_HEADERS = ['ID', '提出物名', '対象クラス', '伝えた日', '締切', '作成日時'];
const MAIL_MARK = '#MAIL';
const SHEET_MAILLOG = 'メール送信記録';

// ===== メニュー =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ 提出物チェッカー')
    .addItem('提出物チェッカーを開く', 'openChecker')
    .addItem('クラスシートを生成', 'generateClassSheets')
    .addSeparator()
    .addItem('サンプルデータを消して初期化', 'menuReset')
    .addToUi();
}

function openChecker() {
  ensureBaseSheets_();
  const html = HtmlService.createHtmlOutput(appHtml_())
    .setWidth(1200)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '提出物チェッカー');
}

// ウェブアプリとして開いたとき
function doGet() {
  return HtmlService.createHtmlOutput(appHtml_())
    .setTitle('提出物チェッカー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function generateClassSheets() {
  ensureBaseSheets_();
  const settings = readSettings_();
  if (!settings.length) {
    SpreadsheetApp.getUi().alert('「設定」シートにクラス名と人数を入力してから実行してください。');
    return;
  }
  const tasks = readTasks_();
  settings.forEach(function (c) { ensureClassSheet_(c.name, c.count, tasks); });
  SpreadsheetApp.getUi().alert(settings.length + 'クラスのシートを用意しました。\nメニュー「提出物チェッカーを開く」から使い始められます。');
}

function menuReset() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('初期化', 'すべてのクラス・提出物・チェック記録を消します。元に戻せません。よろしいですか？', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  importState({ mode: getMode_(), classes: [], tasks: [], checks: {} });
  ui.alert('初期化しました。「設定」シートにクラス名と人数を入力し、「クラスシートを生成」を実行してください。');
}

// ===== 画面から呼ばれる関数 =====
function getData() {
  ensureBaseSheets_();
  const settings = readSettings_();
  const tasks = readTasks_();
  const checks = {};
  tasks.forEach(function (t) { checks[t.id] = {}; });
  const classes = settings.map(function (c) {
    const sh = ensureClassSheet_(c.name, c.count, tasks);
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(1, 1, c.count + 2, Math.max(lastCol, 2)).getValues();
    const names = [];
    for (let r = 2; r < c.count + 2; r++) names.push(String(vals[r][1] || '').trim());
    const mailCol = vals[1].indexOf(MAIL_MARK);
    const emails = [];
    for (let r = 2; r < c.count + 2; r++) emails.push(mailCol >= 0 ? String(vals[r][mailCol] || '').trim() : '');
    for (let col = 2; col < vals[1].length; col++) {
      const id = String(vals[1][col] || '');
      if (!checks[id]) continue;
      const m = {};
      for (let r = 2; r < c.count + 2; r++) if (String(vals[r][col]).trim() !== '') m[r - 1] = true;
      checks[id][c.name] = m;
    }
    const last = props_().getProperty('lastMail:' + c.name);
    return { name: c.name, count: c.count, names: names, emails: emails, lastMail: last ? JSON.parse(last) : null, mailStop: getMailStop_(c.name) };
  });
  return { mode: getMode_(), classes: classes, tasks: tasks, checks: checks, sheetUrl: ss_().getUrl(), user: userEmail_() };
}

function setCheck(clsName, taskId, no, on) {
  return withLock_(function () {
    const sh = ss_().getSheetByName(clsName);
    if (!sh || !isClassSheet_(sh)) throw new Error('クラスのシートが見つかりません：' + clsName);
    let col = colMap_(sh)[taskId];
    if (!col) {
      const t = readTasks_().filter(function (x) { return x.id === taskId; })[0];
      if (!t) throw new Error('提出物が見つかりません');
      col = addTaskColumn_(sh, t);
    }
    sh.getRange(Number(no) + 2, col).setValue(on ? '✓' : '');
    return true;
  });
}

function saveTask(task) {
  return withLock_(function () {
    const name = String(task.name || '').trim();
    if (!name || !task.id) throw new Error('提出物名が空です');
    const t = {
      id: String(task.id), name: name, classes: (task.classes || []).map(String),
      told: normDate_(task.told), due: normDate_(task.due), created: Number(task.created) || Date.now(),
    };
    const tasks = readTasks_();
    const i = tasks.map(function (x) { return x.id; }).indexOf(t.id);
    if (i >= 0) tasks[i] = t; else tasks.push(t);
    writeTasks_(tasks);
    const settings = readSettings_();
    settings.forEach(function (c) {
      const sh = ensureClassSheet_(c.name, c.count, t.classes.indexOf(c.name) >= 0 ? [t] : []);
      const col = colMap_(sh)[t.id];
      if (col) sh.getRange(1, col).setValue(t.name);
    });
    return true;
  });
}

function deleteTask(taskId) {
  return withLock_(function () {
    writeTasks_(readTasks_().filter(function (t) { return t.id !== taskId; }));
    classSheets_().forEach(function (sh) {
      const col = colMap_(sh)[taskId];
      if (col) sh.deleteColumn(col);
    });
    return true;
  });
}

function saveClasses(list) {
  withLock_(function () {
    list = validateClasses_(list);
    const ss = ss_();
    const oldNames = readSettings_().map(function (c) { return c.name; });
    const keep = list.map(function (c) { return c.oldName; }).filter(String);
    let tasks = readTasks_();

    // 削除されたクラス
    oldNames.filter(function (n) { return keep.indexOf(n) < 0; }).forEach(function (n) {
      const sh = ss.getSheetByName(n);
      if (sh && isClassSheet_(sh)) ss.deleteSheet(sh);
      tasks.forEach(function (t) { t.classes = t.classes.filter(function (x) { return x !== n; }); });
    });
    // 名前が変わったクラス（入れ替えにも対応するため、いったん仮の名前にする）
    const renamed = list.filter(function (c) { return c.oldName && c.oldName !== c.name; });
    renamed.forEach(function (c) {
      const sh = ss.getSheetByName(c.oldName);
      if (sh) sh.setName('__tmp__' + c.oldName);
    });
    // クラスごとの設定（メール停止・前回送信）も新しい名前に移す
    const props = props_();
    const moved = renamed.map(function (c) {
      return { name: c.name, values: ['mailStop:', 'lastMail:'].map(function (k) { return [k, props.getProperty(k + c.oldName)]; }) };
    });
    renamed.forEach(function (c) { props.deleteProperty('mailStop:' + c.oldName); props.deleteProperty('lastMail:' + c.oldName); });
    moved.forEach(function (m) { m.values.forEach(function (kv) { if (kv[1] !== null) props.setProperty(kv[0] + m.name, kv[1]); }); });
    renamed.forEach(function (c) {
      if (ss.getSheetByName(c.name)) throw new Error('同じ名前のシートがすでにあります：' + c.name);
      const sh = ss.getSheetByName('__tmp__' + c.oldName);
      if (sh) sh.setName(c.name);
    });
    tasks.forEach(function (t) {
      t.classes = t.classes.map(function (n) {
        const r = renamed.filter(function (c) { return c.oldName === n; })[0];
        return r ? r.name : n;
      });
    });
    writeTasks_(tasks);
    writeSettings_(list);

    list.forEach(function (c) {
      const sh = ensureClassSheet_(c.name, c.count, tasks);
      writeRoster_(sh, c);
      const last = sh.getLastRow();
      if (last > c.count + 2) sh.getRange(c.count + 3, 1, last - c.count - 2, sh.getLastColumn()).clearContent();
    });
  });
  return getData();
}

// 未提出の生徒へ一人ずつメールを送る（mails: [{no, to, subject, body}]）
function sendReminders(clsName, mails) {
  if (!mails || !mails.length) throw new Error('送信対象がいません');
  const remaining = MailApp.getRemainingDailyQuota();
  if (mails.length > remaining) throw new Error('今日送れるメールはあと' + remaining + '通です（Googleの1日の上限）。人数を減らすか、明日送ってください。');
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const log = [];
  const failed = [];
  const skipped = [];
  const stop = getMailStop_(clsName);
  let sent = 0;
  mails.forEach(function (m) {
    // ほかの端末でメール停止にした生徒には送らない
    if (stop.indexOf(Number(m.no)) >= 0) { skipped.push(Number(m.no)); return; }
    const to = String(m.to || '').trim();
    let result = '送信済み';
    try {
      if (!re.test(to)) throw new Error('メールアドレスの形が正しくありません');
      MailApp.sendEmail({ to: to, subject: String(m.subject).slice(0, 200), body: String(m.body) });
      sent++;
    } catch (e) {
      result = '失敗：' + e.message;
      failed.push({ no: m.no, to: to, error: e.message });
    }
    log.push([new Date(), clsName, m.no, to, m.subject, result]);
  });
  withLock_(function () {
    let sh = ss_().getSheetByName(SHEET_MAILLOG);
    if (!sh) {
      sh = ss_().insertSheet(SHEET_MAILLOG);
      sh.getRange(1, 1, 1, 6).setValues([['送信日時', 'クラス', '番号', '宛先', '件名', '結果']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    if (log.length) sh.getRange(sh.getLastRow() + 1, 1, log.length, 6).setValues(log);
  });
  let lastMail = { at: Date.now(), sent: sent };
  if (sent) props_().setProperty('lastMail:' + clsName, JSON.stringify(lastMail));
  else { const prev = props_().getProperty('lastMail:' + clsName); lastMail = prev ? JSON.parse(prev) : null; }
  return { sent: sent, failed: failed, skipped: skipped, remaining: MailApp.getRemainingDailyQuota(), lastMail: lastMail };
}

// 事情があって提出できない生徒など、未提出のメールを送らない番号を保存する
function setMailStop(clsName, nos) {
  props_().setProperty('mailStop:' + clsName, JSON.stringify((nos || []).map(Number).filter(function (n) { return n > 0; })));
  return true;
}

function getMailStop_(clsName) {
  const v = props_().getProperty('mailStop:' + clsName);
  try { return v ? JSON.parse(v).map(Number) : []; } catch (e) { return []; }
}

// メール送信の許可を出すため、エディタで一度だけ実行する
function authorize() {
  ss_();
  Logger.log('許可が完了しました。今日送れるメールはあと ' + MailApp.getRemainingDailyQuota() + ' 通です。');
}

function setMode(mode) {
  props_().setProperty('mode', mode === 'single' ? 'single' : 'multi');
  return true;
}

function importState(state) {
  withLock_(function () {
    ensureBaseSheets_();
    const ss = ss_();
    const list = validateClasses_((state.classes || []).map(function (c) {
      return { name: c.name, count: c.count, names: c.names || [], emails: c.emails || [] };
    }));
    // 既存のクラスシートを削除（シートが1枚だけにならないよう、基本シートは残る）
    classSheets_().forEach(function (sh) { ss.deleteSheet(sh); });
    const tasks = (state.tasks || []).map(function (t) {
      return {
        id: String(t.id), name: String(t.name), classes: (t.classes || []).map(String),
        told: normDate_(t.told), due: normDate_(t.due), created: Number(t.created) || Date.now(),
      };
    });
    writeSettings_(list);
    writeTasks_(tasks);
    if (state.mode) setMode(state.mode);
    (state.classes || []).forEach(function (c) { setMailStop(String(c.name).trim(), c.mailStop || []); });
    const checks = state.checks || {};
    list.forEach(function (c) {
      const sh = ensureClassSheet_(c.name, c.count, tasks);
      writeRoster_(sh, c);
      const map = colMap_(sh);
      tasks.forEach(function (t) {
        const col = map[t.id];
        const m = (checks[t.id] || {})[c.name] || {};
        if (!col) return;
        const vals = [];
        for (let n = 1; n <= c.count; n++) vals.push([m[n] ? '✓' : '']);
        sh.getRange(3, col, c.count, 1).setValues(vals);
      });
    });
  });
  return getData();
}

// 画面（index.html）は、ファイルを小さくするため圧縮して埋め込んでいる（APP_HTML_GZ）
function appHtml_() {
  const blob = Utilities.newBlob(Utilities.base64Decode(APP_HTML_GZ), 'application/x-gzip');
  return Utilities.ungzip(blob).getDataAsString('UTF-8');
}

// ===== シート操作 =====
// バインドされていればそのスプレッドシート、そうでなければ先生ごとに自動作成したデータ用スプレッドシートを使う
function ss_() {
  if (ss_.cache) return ss_.cache;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const props = PropertiesService.getUserProperties();
    const id = props.getProperty('SPREADSHEET_ID') || migrateFromScriptProps_();
    if (id) {
      try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
    }
    if (!ss) {
      ss = SpreadsheetApp.create('提出物チェッカー データ');
      props.setProperty('SPREADSHEET_ID', ss.getId());
    }
  }
  ss_.cache = ss;
  return ss;
}

// 以前の版（1人で使う版）は、保存先などをスクリプト全体の設定に記録していた。
// データ用スプレッドシートの持ち主本人が開いたときだけ、その人の設定に引き継ぐ（ほかの先生には引き継がない）
function migrateFromScriptProps_() {
  const sp = PropertiesService.getScriptProperties();
  const id = sp.getProperty('SPREADSHEET_ID');
  if (!id) return null;
  try {
    const owner = SpreadsheetApp.openById(id).getOwner();
    const me = userEmail_();
    if (!owner || !me || owner.getEmail() !== me) return null;
  } catch (e) {
    return null;
  }
  PropertiesService.getUserProperties().setProperties(sp.getProperties());
  sp.deleteAllProperties();
  return id;
}

function userEmail_() {
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

// 設定（使い方・メール停止・前回送信）の保存先。ウェブアプリでは先生ごとに分ける
function props_() {
  const doc = PropertiesService.getDocumentProperties();
  if (doc) return doc;
  ss_(); // 以前の版の設定があれば、先に引き継いでおく
  return PropertiesService.getUserProperties();
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getUserLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function getMode_() { return props_().getProperty('mode') || ''; }

function ensureBaseSheets_() {
  const ss = ss_();
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    s = ss.insertSheet(SHEET_SETTINGS, 0);
    s.getRange('A1:B1').setValues([['クラス名', '人数']]).setFontWeight('bold');
    s.getRange('D1').setValue('最大' + MAX_CLASSES + 'クラス、1クラス' + MAX_STUDENTS + '人まで。入力後、メニュー「⚙ 提出物チェッカー」→「クラスシートを生成」を実行します。');
    s.setColumnWidth(1, 140);
    s.setFrozenRows(1);
  }
  let t = ss.getSheetByName(SHEET_TASKS);
  if (!t) {
    t = ss.insertSheet(SHEET_TASKS, 1);
    t.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]).setFontWeight('bold');
    t.getRange('D:E').setNumberFormat('@');
    t.setColumnWidth(2, 220);
    t.setFrozenRows(1);
  }
  // 新規スプレッドシートの空の「シート1」を片付ける
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 2 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function readSettings_() {
  const s = ss_().getSheetByName(SHEET_SETTINGS);
  const vals = s.getRange(2, 1, MAX_CLASSES, 2).getValues();
  const out = [];
  vals.forEach(function (r) {
    const name = String(r[0] || '').trim();
    if (!name) return;
    const count = Math.max(1, Math.min(MAX_STUDENTS, parseInt(r[1], 10) || 1));
    out.push({ name: name, count: count });
  });
  return out;
}

function writeSettings_(list) {
  const s = ss_().getSheetByName(SHEET_SETTINGS);
  s.getRange(2, 1, MAX_CLASSES, 2).clearContent();
  if (list.length) s.getRange(2, 1, list.length, 2).setValues(list.map(function (c) { return [c.name, c.count]; }));
}

function readTasks_() {
  const s = ss_().getSheetByName(SHEET_TASKS);
  const last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, TASK_HEADERS.length).getValues()
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        name: String(r[1]).trim(),
        classes: String(r[2] || '').split(/[、,，]/).map(function (x) { return x.trim(); }).filter(String),
        told: normDate_(r[3]),
        due: normDate_(r[4]),
        created: r[5] instanceof Date ? r[5].getTime() : Number(r[5]) || 0,
      };
    });
}

function writeTasks_(tasks) {
  const s = ss_().getSheetByName(SHEET_TASKS);
  const last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, TASK_HEADERS.length).clearContent();
  if (!tasks.length) return;
  const rows = tasks.map(function (t) { return [t.id, t.name, t.classes.join('、'), t.told, t.due, t.created]; });
  s.getRange(2, 4, rows.length, 2).setNumberFormat('@');
  s.getRange(2, 1, rows.length, TASK_HEADERS.length).setValues(rows);
}

function normDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function isClassSheet_(sh) { return sh.getRange('A2').getValue() === ID_MARK; }

function classSheets_() { return ss_().getSheets().filter(isClassSheet_); }

function colMap_(sh) {
  const lastCol = sh.getLastColumn();
  const map = {};
  if (lastCol < 3) return map;
  sh.getRange(2, 1, 1, lastCol).getValues()[0].forEach(function (id, i) { if (i >= 2 && id) map[String(id)] = i + 1; });
  return map;
}

function addTaskColumn_(sh, t) {
  const col = Math.max(sh.getLastColumn(), 2) + 1;
  if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
  sh.getRange(1, col, 2, 1).setValues([[t.name], [t.id]]);
  sh.getRange(1, col).setFontWeight('bold').setWrap(true);
  sh.getRange(3, col, Math.max(sh.getMaxRows() - 2, 1), 1).setHorizontalAlignment('center');
  sh.setColumnWidth(col, 110);
  return col;
}

function ensureClassSheet_(name, count, tasks) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A1:C2').setValues([['番号', '氏名（任意）', 'メール（任意）'], [ID_MARK, '', MAIL_MARK]]);
    sh.getRange('A1:C1').setFontWeight('bold');
    sh.setColumnWidth(3, 200);
    sh.hideRows(2);
    sh.setFrozenRows(2);
    sh.setFrozenColumns(2);
    sh.setColumnWidth(1, 50);
    sh.setColumnWidth(2, 120);
  } else if (!isClassSheet_(sh)) {
    throw new Error('「' + name + '」という名前の別のシートがあります。クラス名を変えるか、そのシートの名前を変えてください。');
  }
  // 以前の版で作ったシートには、氏名の右にメール列を追加する
  if (sh.getRange(2, 1, 1, Math.max(sh.getLastColumn(), 2)).getValues()[0].indexOf(MAIL_MARK) < 0) {
    sh.insertColumnAfter(2);
    sh.getRange(1, 3, 2, 1).setValues([['メール（任意）'], [MAIL_MARK]]);
    sh.getRange(1, 3).setFontWeight('bold');
    sh.setColumnWidth(3, 200);
  }
  if (sh.getMaxRows() < count + 2) sh.insertRowsAfter(sh.getMaxRows(), count + 2 - sh.getMaxRows());
  const nums = [];
  for (let n = 1; n <= count; n++) nums.push([n]);
  sh.getRange(3, 1, count, 1).setValues(nums);
  const map = colMap_(sh);
  (tasks || []).forEach(function (t) {
    if (t.classes.indexOf(name) >= 0 && !map[t.id]) map[t.id] = addTaskColumn_(sh, t);
  });
  return sh;
}

function writeRoster_(sh, c) {
  const mailCol = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(MAIL_MARK) + 1;
  const names = [], emails = [];
  for (let i = 0; i < c.count; i++) {
    names.push([String((c.names || [])[i] || '').trim()]);
    emails.push([String((c.emails || [])[i] || '').trim()]);
  }
  sh.getRange(3, 2, c.count, 1).setValues(names);
  if (mailCol > 0) sh.getRange(3, mailCol, c.count, 1).setValues(emails);
}

function validateClasses_(list) {
  if (list.length > MAX_CLASSES) throw new Error('クラスは' + MAX_CLASSES + 'クラスまでです');
  const seen = {};
  return list.map(function (c) {
    const name = String(c.name || '').trim();
    const count = parseInt(c.count, 10);
    if (!name) throw new Error('クラス名が空です');
    if (name === SHEET_SETTINGS || name === SHEET_TASKS || /[\[\]\*\?\/\\:]/.test(name)) throw new Error('「' + name + '」はクラス名に使えません');
    if (!(count >= 1 && count <= MAX_STUDENTS)) throw new Error('人数は1〜' + MAX_STUDENTS + '人で入力してください');
    if (seen[name]) throw new Error('同じクラス名があります：' + name);
    seen[name] = true;
    return { name: name, count: count, names: c.names || [], emails: c.emails || [], oldName: c.oldName ? String(c.oldName) : '' };
  });
}

// ===== 画面（index.html の内容を gzip 圧縮して base64 にしたもの） =====
// Apps Script は1ファイルが大きすぎると保存できないため、圧縮して小さくしている
const APP_HTML_GZ = [
  "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"
].join('');
