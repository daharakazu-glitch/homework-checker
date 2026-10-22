// 提出物チェッカー（Google Apps Script 用・1ファイル版）
// このファイルの中身をすべてコピーして、script.google.com のプロジェクトに貼り付けてください。
// 設定手順: https://daharakazu-glitch.github.io/homework-checker/setup.html
// ※ tools/build-gas.js で gas/server.gs と index.html から自動生成しています。直接編集しないでください。

/**
 * 提出物チェッカー（英語課題・複数クラス対応）
 * ウェブアプリとしてデプロイすると、同じGoogleアカウントならどの端末からでも同じデータを使えます。
 * データは自動で作成するスプレッドシート「提出物チェッカー データ」に保存します。
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
  return { mode: getMode_(), classes: classes, tasks: tasks, checks: checks, sheetUrl: ss_().getUrl() };
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
// バインドされていればそのスプレッドシート、そうでなければ自動作成したデータ用スプレッドシートを使う
function ss_() {
  if (ss_.cache) return ss_.cache;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('SPREADSHEET_ID');
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

function props_() { return PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties(); }

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
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
