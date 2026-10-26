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
const BOARD_ID_ROW = 6;   // 掲示用シート：提出物IDの行（非表示）
const BOARD_FIRST = 7;    // 掲示用シート：1番の生徒の行

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
  const boards = {};
  settings.forEach(function (c) { const b = getBoard_(c.name); if (b) boards[c.name] = b; });
  return { mode: getMode_(), classes: classes, tasks: tasks, checks: checks, sheetUrl: ss_().getUrl(), user: userEmail_(), boards: boards };
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
    safely_(function () { updateBoardCell_(clsName, taskId, Number(no), on); });
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
    const affected = t.classes.concat(i >= 0 ? tasks[i].classes : []);
    if (i >= 0) tasks[i] = t; else tasks.push(t);
    writeTasks_(tasks);
    const settings = readSettings_();
    settings.forEach(function (c) {
      const sh = ensureClassSheet_(c.name, c.count, t.classes.indexOf(c.name) >= 0 ? [t] : []);
      const col = colMap_(sh)[t.id];
      if (col) sh.getRange(1, col).setValue(t.name);
    });
    renderBoards_(affected);
    return true;
  });
}

function deleteTask(taskId) {
  return withLock_(function () {
    const all = readTasks_();
    const old = all.filter(function (t) { return t.id === taskId; })[0];
    writeTasks_(all.filter(function (t) { return t.id !== taskId; }));
    classSheets_().forEach(function (sh) {
      const col = colMap_(sh)[taskId];
      if (col) sh.deleteColumn(col);
    });
    if (old) renderBoards_(old.classes);
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
      safely_(function () { removeBoard_(n); });
      tasks.forEach(function (t) { t.classes = t.classes.filter(function (x) { return x !== n; }); });
    });
    // 名前が変わったクラス（入れ替えにも対応するため、いったん仮の名前にする）
    const renamed = list.filter(function (c) { return c.oldName && c.oldName !== c.name; });
    renamed.forEach(function (c) {
      const sh = ss.getSheetByName(c.oldName);
      if (sh) sh.setName('__tmp__' + c.oldName);
    });
    // クラスごとの設定（メール停止・前回送信・掲示用）も新しい名前に移す
    const props = props_();
    const keys = ['mailStop:', 'lastMail:', 'board:'];
    const moved = renamed.map(function (c) {
      return { name: c.name, values: keys.map(function (k) { return [k, props.getProperty(k + c.oldName)]; }) };
    });
    renamed.forEach(function (c) { keys.forEach(function (k) { props.deleteProperty(k + c.oldName); }); });
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
    renderBoards_(list.map(function (c) { return c.name; }));
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
  DriveApp.getRootFolder();
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
    // 掲示用は止める（サンプルや復元したデータが生徒に見えないように。必要なら作り直す）
    readSettings_().forEach(function (c) { safely_(function () { removeBoard_(c.name); }); });
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

// ===== 掲示用（生徒が見るだけの、クラスごとのスプレッドシート） =====
// チェックのたびに書き換えるので、生徒はリンクを開くだけで最新の状況を見られる。
// 共有は「学校内のリンクを知っている全員が閲覧のみ」。表示するのは出席番号と ✓ だけ（名前・メールは出さない）。
function createBoard(clsName) {
  return withLock_(function () {
    if (!readSettings_().some(function (c) { return c.name === clsName; })) throw new Error('クラスが見つかりません：' + clsName);
    let info = getBoard_(clsName);
    let bss = null;
    if (info) { try { bss = SpreadsheetApp.openById(info.id); } catch (e) { bss = null; } }
    if (!bss) {
      bss = SpreadsheetApp.create(boardTitle_(clsName));
      info = { id: bss.getId(), url: bss.getUrl() };
    }
    info.shared = false;
    info.shareError = '';
    try {
      DriveApp.getFileById(info.id).setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      info.shared = true;
    } catch (e) {
      info.shareError = e.message;
    }
    props_().setProperty('board:' + clsName, JSON.stringify(info));
    renderBoard_(clsName);
    return info;
  });
}

function stopBoard(clsName) {
  return withLock_(function () { removeBoard_(clsName); return true; });
}

function getBoard_(clsName) {
  const v = props_().getProperty('board:' + clsName);
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

function removeBoard_(clsName) {
  const info = getBoard_(clsName);
  props_().deleteProperty('board:' + clsName);
  if (info) safely_(function () { DriveApp.getFileById(info.id).setTrashed(true); });
}

function boardTitle_(clsName) { return clsName + ' 提出物の状況'; }

function renderBoards_(names) {
  const seen = {};
  (names || []).forEach(function (n) {
    if (seen[n]) return;
    seen[n] = true;
    safely_(function () { renderBoard_(n); });
  });
}

function boardSheet_(clsName) {
  const info = getBoard_(clsName);
  if (!info) return null;
  const bss = SpreadsheetApp.openById(info.id);
  if (bss.getName() !== boardTitle_(clsName)) bss.rename(boardTitle_(clsName));
  const sh = bss.getSheets()[0];
  if (sh.getName() !== '提出状況') sh.setName('提出状況');
  return sh;
}

function nowText_() { return Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'M/d HH:mm'); }

// クラスの掲示用シートを、いまの記録から作り直す
function renderBoard_(clsName) {
  const c = readSettings_().filter(function (x) { return x.name === clsName; })[0];
  if (!c) return;
  const sh = boardSheet_(clsName);
  if (!sh) return;
  const tasks = readTasks_().filter(function (t) { return t.classes.indexOf(clsName) >= 0; })
    .sort(function (a, b) { return a.created - b.created; });
  const csh = ensureClassSheet_(clsName, c.count, tasks);
  const map = colMap_(csh);
  const cvals = csh.getRange(3, 1, c.count, Math.max(csh.getLastColumn(), 2)).getValues();
  const width = Math.max(tasks.length + 1, 2);
  const pad = function (row) { while (row.length < width) row.push(''); return row; };
  const last = BOARD_FIRST + c.count - 1;
  const colName = function (i) { let n = i + 1, s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  const rows = [
    pad([clsName + ' 提出物の状況']),
    pad(['最終更新：' + nowText_()]),
    pad(['✓＝提出済み　空欄＝まだ確認できていません（先生が確認すると、ここに反映されます）']),
    pad(['番号'].concat(tasks.map(function (t) { return t.name; }))),
    pad(['締切'].concat(tasks.map(function (t) { return t.due ? Number(t.due.slice(5, 7)) + '/' + Number(t.due.slice(8, 10)) : '−'; }))),
    pad(['#ID'].concat(tasks.map(function (t) { return t.id; }))),
  ];
  const bgs = [];
  for (let n = 1; n <= c.count; n++) {
    const row = [n];
    const bg = ['#ffffff'];
    tasks.forEach(function (t) {
      const col = map[t.id];
      const on = col ? String(cvals[n - 1][col - 1]).trim() !== '' : false;
      row.push(on ? '✓' : '');
      bg.push(on ? '#ffffff' : '#eeeeee');
    });
    rows.push(pad(row));
    while (bg.length < width) bg.push('#ffffff');
    bgs.push(bg);
  }
  rows.push(pad(['提出'].concat(tasks.map(function (t, i) {
    const L = colName(i + 1);
    return '=COUNTIF(' + L + BOARD_FIRST + ':' + L + last + ',"✓")&"/' + c.count + '"';
  }))));

  sh.clear();
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  if (sh.getMaxRows() < rows.length) sh.insertRowsAfter(sh.getMaxRows(), rows.length - sh.getMaxRows());
  sh.showRows(1, rows.length);
  sh.getRange(1, 1, rows.length, width).setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, 1, 3, width).setHorizontalAlignment('left');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1, 2, 1).setFontColor('#555555');
  sh.getRange(4, 1, 2, width).setFontWeight('bold').setBackground('#e6f2f0').setWrap(true);
  sh.getRange(BOARD_FIRST, 1, c.count, width).setBackgrounds(bgs);
  sh.getRange(BOARD_FIRST, 1, c.count, 1).setFontWeight('bold');
  sh.getRange(rows.length, 1, 1, width).setFontWeight('bold').setBackground('#e6f2f0');
  sh.getRange(4, 1, rows.length - 3, width).setBorder(true, true, true, true, true, true, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
  sh.hideRows(BOARD_ID_ROW);
  sh.setFrozenRows(BOARD_ID_ROW);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 60);
  if (width > 1) sh.setColumnWidths(2, width - 1, 120);
  if (!tasks.length) sh.getRange(4, 2).setValue('（まだ提出物がありません）');
}

// チェック1つ分だけ書き換える（列が見つからなければ全体を作り直す）
function updateBoardCell_(clsName, taskId, no, on) {
  const sh = boardSheet_(clsName);
  if (!sh) return;
  const ids = sh.getRange(BOARD_ID_ROW, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const col = ids.indexOf(taskId) + 1;
  const c = readSettings_().filter(function (x) { return x.name === clsName; })[0];
  if (col < 2 || !c || no > c.count || sh.getRange(BOARD_FIRST + c.count - 1, 1).getValue() !== c.count) { renderBoard_(clsName); return; }
  sh.getRange(BOARD_FIRST + no - 1, col).setValue(on ? '✓' : '').setBackground(on ? '#ffffff' : '#eeeeee');
  sh.getRange(2, 1).setValue('最終更新：' + nowText_());
}

// 掲示用の更新に失敗しても、記録の保存は止めない
function safely_(fn) {
  try { fn(); } catch (e) { console.warn('掲示用の更新に失敗：' + e.message); }
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
  "H4sIAAAAAAACA9W9eXcUR5Y4+vdPnyKQNa6qplRaAVsbTWPs9oy3Z+FzZg5mrFRVlpRNVmZ1ZRaCluscVclggYShscFmsTHdGDA0wlu7McLwXVwqSfzFV3jv3hsRGZFLSXL3zHtvFqPKjIy4sd19Gdnx0pv7D/7XWwfYtF+yxzpG4B9mG87UaOcfjE54YBqFsY6RScMzmW9Upkx/tPM93y3Du5LpGyw/bVQ8ePrOwZe7X5CPHaNkjnYetcyZslvxO1nedXzT8Uc7Z6yCPz1aMI9aebMbf2SZ5Vi+ZdjdXt6wzdE+6MS3fNscWzt7rvXhw/VTXzfn683Greb8fLNxtzn/aKSH3neMeP5x+Heo4ro+m+1grLt7cmqIPVccLO4q7hlm3d15o1KAB8Ui/PLNY/4Qe66v0D8w0AsPSlXfhPe7J/f0v4BPbMsxh9hzZr+5yzSGsUsjnzcd+K5/1+4Bc3JYPur23CI8NwfMyWKRGue7865tm3m9",
  "vXxK8JlmcSD4oGjODLHnJgd3DfS+OCye8IkUi7vNXtHQPWpWhthz+d39L/S/MCwf8aYFM2/mRdOCi9Po2/VC70BhWD7iw+8p7jEL1NR1oNluY2DQGMaf1KSQL+4yB4c7ah2/LZkFy2DpcsUsmhUPJuJWur38tFkyh1jBqBzJ4NIHmyC3oW+gb3dfXtmGvkJ/X/+Lyk6Ye8wXTVPdiRcNo88oKDvRbw4MDJjDvF+5F3sMYxdtqr4X/YMDvYNF0VzdjReKk/30RWg3+gr9ewZ7g09oP4qTk8VdRmg/Boz+Yt+eoCnfkWLxBeMFI7wjA2Zfvi8fNOZ7sts0B140wnvSt2dgV/8u0Rh3JWgodqVvcsDox+FrHbWO37BZNuke6/asP1nO1BCbdCsFGNs9NsxqHZNu4TibZSWjMmU5Q6x3mE0a+SNTFbfqFIbYUaOShl3KDDPcUPEE9iUzzIourKj1J3OI9e0qH4MR8VHRKFn28SHmHfd8s9Rd",
  "tbKs2yiXbbObnmRZ5++tijFlOS4bNxyvM8s633B9+sH+/S34/V9V9orrT1t59s6r8Pt106ocdzuzzDMcr9szK1YR4AfcY1bYLCu7nuVbsCKeb+WPHB9mvlvGCf2p23IK5rEh1h87Ozh0meFgWXzfLQ2xvvIx5rm2VeCt4JhlYIJlo1DAdXyhfIz17S4fG2YFyyvbxvEhVrTNY8P43+6ZilEeYvDfYTYFf+6G5oPQ3LCtKafb8s2SN8TgUJoVZSbTfWxWW1gcQtmfmWnLN7u9spE3h5jj0hC1Dsc4ymbDoODI/eU4oOiLyarvuw6cEGVdHNcxxYKIX3Lae2AefdAlX7GKUbCqHq5H6JTgZc2IUzHELGfarFi+fm5wSfDBjGlNTftDbHdv7zDLVyse9FR2LbFAAbw5I+9bR80Q2DSocs/Dp5ZeZaCvkmE5eOyPEX0ZYn39vb3aSjOj6rvKxAFS3A02iA1rHTk4OrFAaGcq6TCFF7APNyo6",
  "noApOJuD6vjT/aEDs0ufRi/r4/BOD+Q8Mx9qPai27oND3Zu8lTBoxZ1JOGgvxB60hOOeK0Z7wW8LVsXM003Ou3a15PDeB+RB4ZD3t4HScspVP8s8E/B3lgG+MiqmweeuHMU4tJaAAdvvZgeLvRByO+HgvED7oEAjDl9v778Ns5LldE/zS9AnNi036TtRsCM7uNlRSzqiMfPfykz44ocvaQJugjnkyhWrZFSOt7m0wZWIv7b8KbFpvNeC4Uwh/tc+Ifqaka28kmHboYOv37aBYHvgi6GC5RmTtgnX2y0becs/PsRyg7uCORfMolG1ffjA4HsEy9ddMPNuxaDzS6hTHnLLgd3onrTd/BEcCM9rGHZxiFVYB8SFnz6i3hreoXIFdyUSmEh3MRg2N20VCqajDgFTYDusEnDohoPTzZmlsg/7KBevH7EUTh/HDkZNuJ4dPb9hnm/4VY/BOfTYb3oIm3nq2FMVqzCM/+32zVLZNnyzm1CCN8Qq",
  "Ztk0/DQg6e6iZdtZuD4l41i6f1dv+ViW9RUrmQxfFXmXfI6xVVBts8gv1LbRwpbQfHCdYCS6PmoryWbGUIRenSL0Jt+87WDSQXU1hqbhqiCP2ObqybXL+W45irj/UPV8q3i8m0twQwxRQPek6c+YphNwQFtH4GI4kBHFxRUMwh5gECIUL/gm71YdP3Tb6QYxvcmk3qhfPya5kuV5sd0ELUCMnd36pCaNwpQZc4M5SlD76YvwRTjt4DAIhBU6My+++CKH0fPFyQp2F45gdwhVBsevxnKT6jfRA6+IRGHioXWEoxfNmfYjF80ZMSpvGzMiSVXR0fjHOJJ+hOOGEvQAxhKtYwbjYll0tICeeD7KY+2HgxZiONE6Zjgu2EWHE9/DkQFYBVuwO2bH8UxGe+eUHwAv2u7MECPkLjsdYwXrqNI18SBqP/lqpWI6/n4ADT+zzSnTKSRwf7sTxAzA9flpM38EMTlieh9IK7wH6irhI34bWHIBUt6w",
  "8+m+3t6j06yb9e/uLR/bPlPdG786HHXXOhCWHIKm4kDbKHvmEPPMslExfCkOIW+D9683lkIrvYHSSv2p9N9eyJRT4PsS30gigt54yhuCxTTgv23l5MRFCgToEBdIsoPlCBlqD5dWVKEKnhw1K76VN2wBI80/BGOO4IkFkSjnlmDsU5eGMGSU2ifwqZEVC2CS/Q8Q/p/u3hJlQmqAKF5es1z/rmE241YK3ZMV0zgyxPCfbsO2Rc9eNUSYiBgkc4li+EEYHthEE7ndiHChbZMAaECTH7jsryyzXzEcr2wAJhhuJ4bsURRA2npocLtOJp71BJhzEVWE+IiQpGgGnOCmCJw3lvyNW/VJV9gfulCSyeEtut1i0TP9IdbdzynpH6tW/ghDmTJY1hd7w/zMC0lcMKBOw/NL3lQcO6GJfbkBsySxJggtoB7n7LFtJ8vehOc0zh9xsecbFT+kvundoraLVAw0am5m2o097eq5CvQi8iMDGTGb30z1",
  "nbMp4ySpbpJY6bi+uWX2q61UnyDqoo5PUwvtRgWJEChgl8oVy/FZuWLKnSobU6a3Te0GKXqSdCXQY+jAkxQsROJe2Aq+Df19vaWSfqz6X9wDjxTdUqnE+vrhGaqGp40CUOBePAgwxcrUpJHuzcL/5vpeyMQgNQFVRP30Qlj9NKg2z7VHbs/t2rUrou96Qe0AcXQcqRZ/Det4D9t328Zxt+oPsaJ1zCwonU1nGf8rIM/qdSB4AkG3fIzYrbhbLu8wwhvLd/FBQwvQG68C1Xe7F/63HeHg0xBSi/a1CdYTuQXExpXjoOArjyzBbnVfBgaI8Hm2VQ6tVMHwps0Ce27Pnj2R67Q7SYANbfFuOIpiAXe9CL9iFxCHnx5QLRbihMUJhNi8aivNdxP+4yBxkZx2LE7Oww5yRbRZJR/awcFBaP1bfk2pyb5BTW0vTWSEL2ZRLVA4Hnurax2MkTkgy3KOi19kGeqsx5jH8Yfj+unnAOd04/tMFmRz",
  "w/Oz7LmSWzBfsqc20eMwxrXgKiupMG/ECkA7idEQVfUq2gapy+KNhK2JIxRhO5gyu/lhLfpwbAx7xjjuDTPtITQLLRnvdQioZ3d+2rILCG6kO5QcYh6B/QtwdMny8iR5wAppLCbHCOIgEsNKx2IX4g9gf4pupcQ5IeA9/jPdvav332JpCtcVaRQI6Ey85SiO+ERUu1L/2Muh4ZDz5yzX7w0LVqrbPGo6vidWXnKtu4kxxOnnvGlkIwK95ot4fHNw42zjeNzyWA5yRCGspFKJwUxECRVLzSKqIvFCwrpLU3Ps5jexYBm2O7W5tSWQ/AY13NPfG5JNBnfjE41YBOMgWVNxzO5Ybg91sq6V15Q6/FLoZCgig6ioNqzo65fGwmCAyZghYnAeyTC25fnbIJKJkiz2E0iy9HPrkqwqLcatQVgwLFmFgk2katq0OaZPtG1xa5U4IfiFa2cZ/YVoX8fz/f3q54NECsSXtqXs+IB495ztGtCDiibp",
  "4GxD3T3Sw11QRnqEr4xbOM49Z8wKy9uG5412ckzfOdbB2Mh0XxvPluk+bOOVDYdZhdFO77iTH/cN3+wUfZFpgShnJ8PhRzv5kcPFQARJVHygc2ykB/rCTrlJmPcDlg2ynoi+YLxJ33nbLFZMb1r2TV0NdjL0uhntbNZXmvXFZn15/e79tat3m/VbrRNftU5fadY/bdavtU6eaC3/1Gyc37hzr1l/svH4UbP+pFl/3Kxf6hz75eQKW7s6t3bxm2b97tqVH9YufjPSQ2AhhGAIBigc4yiuVQB0wfCNbqCJo51kXugca135onX2o/XTP659X1d7if0I9VSdY8p63yeoN//SsO3OsdZfv8FJ/7k5/3Fz/ivcq00+5FveOvNNa+EfzflPm/N/2/rHvuEd8TqDg7LpB3DSO8dWf37SrH+wdvEnbVl7HOOoOKJmBZy0DIvegKIQFpzfhc6xYNcaHzfrt5r1D2jvfpm7OdJTsI7S6SRGBb9EPoVv",
  "iTha/EDB2aOW8V/Rnmz3I9iObX5DO7HNj2gHtvkR7kKbb0Z6aOnlwnOGTn4jKLX4VmwSf000TNyN6f6xZv3rZv16s3GqWV+Wm9+s39148mfcu1vN+qVmffHZo2sjPdP9/Luyhks6x5r1RrN+G85341RzbinATnNnfjl5vjm31Gzcb85/3Wz8tHH7Xmv5cnPuDNz6G6fWrvyAQ5yhQ9Kca4z0lPUjqmAbIneddGph4qOdnuVM2Wbn2MjkWJ8cpFm/hVM5OdIzOUaoUId3bfHE6sqK0v4OLMNcXe1iOZhF4/z68vX1cycRQQGcHCfG36n2AJeqtm8hvBs3Ply78M02gL5waf3WnznosNqfNus3m3P1tcUTrZ8/htVrLIb7bJzfuL7UrJ9Z+/yrfwXwnlEq89VuNv7enP8esRJg8I2vb3KA6k+ajcXEOXzx4/rXD5v15QFl3rc3Fr/duPOFtuT1W+ufrDz9/C+w9n95uHHnjAT/2aMFcdxu",
  "tRY+X7t6rbV0UT1Ezx6dip9i6NQSefTdijFlvuECgdRo4RB62fSyXrh/ZUKDiL/EP8qtQuaZOqQ/x3ijjhEvX7HK/lhHTw8bhf9hrcULqz9fXT//Tesv8/QI3u0rlz02jm1hgo2fBL6fb86fajb+0Zx/1JxfePbo1OqD0zjZ+/Jhs3539cnnrXufwRLN1Zv1z5uNpdWVr1o3LvKW9Y8VYrvcnL+IK3+z2fhR+RaOT3Ou0ZF3Hc9nr46/98q+cTbK/ONl0y2yKdedsk22Y3SUpapOwSxajllIseefZzt2pOlljqYKz7QHuUrVyQzzfl8bf+8/DvwXG2Up37S86arvVQmNm5Xuo30p0ez1ff/53v7X9o2PHwAY+nqz+GT84DsvHXjjIDwa3CWajr/zyisHxg+yUXYo1Zy/D2sCR+si0ehUlqVaZz6D4zV/Epd1YR1wzqfqi49oIeHRxuK3qz9fXbv4IfxYfbiw/uS89n7t4odr31/AzbmD",
  "5x+fPr32PdA9ZMGa86dS2Q7GUtACBlyEdo1vNu589/T6ErTGp58goPfUd88eLTxd+u7pte+fPTqF4N24vHHn3sbtz54ufYdj376O12x549Zfm/UzwTw2Fr9du3H1af2j1kcrrQtnEaIL/1i7em310Z/XvroXjLy2dL/ZOL3x5JNm47Q8P6nDcikPvvfavt8deI2NkucJ+b2miE9afXAvlWXo2pqiC7hr9eHD1ZWvVh8sprKMHFlT6//4srXw4caPJ57WP0plGXmsplonbrc+vk63O8Vqwx0dtumjm4fJRhk6jSCPms4M45uqBSB0MFRxD7EUsQepLMvbIMBWbRtW2HMrAF6hagJglu2blf3wPpXKAvkzX8LBi4btmdA8P33k96GnzHXs469bnofMO3+GKv2DhneEj8RAxfC6N0UdF9181fu/oInatWHbcmhYiULVfNMBP1feZ8mw7PHq5B9oQRU0tyyw23Kzfvrp9c+a9Q9SvPkRqzzE",
  "Zmv0682y6QQzR37k4PEyrC2ioFSWnu2X60O/x/3qEEsZti0avGGUTE+CNWkbzpH9ro3WUP7zIDDpOJMOxsyC5ePaUJ841ksVAyQmeuKWTedt10PFCkFX41voFNxx38gfgYt5mJ7l3fLxg+Yx3+PPOopVh/gf9QiwWVYx/WoFvTDdAsFCQwPohw6DyOkd4X8i8vBgoViN9O4S0Z749uncZY5i6YB3sVGW9jJsdIwV3Hy1ZDp+7o9Vs3J8HL0B3Urak5jK9PJB63G/YjlTaY/t3ctSqUyuYpZtI2+mew49PzLWmTrcM5Vl6Tw2Tc+y1POpIZZ63iiVh+HWjeAv28cfY/hjin504o8/Vl382ZnqhJ/PDbw4nGK1Q/nDmcxwsERIWNIlbyqDlxNA9Nko60qnnsNXqcww83Mgeu4n/QkbZSVvCh7i2r1meX7OKBTSKVDyQOu8bRqVg1bJdKt+2s+9N409vDfNRpln+uJFGmeldlIxS+5RU/ST",
  "Zf29vb0ozEpYiyX/JdjLgrKZhdyU6b9cte3/Mo1KOsN2slR3iu0US4uvX3cdfxrf9WVyZaMwDiaqdH+WpXpT8V/gMJloYw0c3y0Yx8f9inq2BIiOOcN4J9o3RqHwknHcSzvBchfYKAuaD7NCzuMAKKCwnQzsmKFRCnrnBeP4S1axmDaybDLovxycN/Ho0PEsK2VZ4TDsSc4r25afTnWnMrmSUU6/US1NohGMDwYj5d45uD8NH7Fu1pdlOK5s8LrhT+dQP5ZOl9OTGdbNymkjk2E97IXdg729kX0sFQCeWWYVWXqHlxH9pFLDAr548OSIE12zO0u1nq7ZnYXahNa1Y868WlA3JOXD3uIcHHcmncn5Lt/ogd2wqgS84RTcUuglqOLzJuz8bnIhlINMmYAR02CPV0ZCupPjCCVXtGA56O7myakMsAd+o58I9CzfT5+lI/0BqsIvUySTpNje0EgEZm+W9WXYkP5OP66A3d4spvO290Yc5Pg+",
  "RwQvnfbV+2l6OcvJ29WC6cnPMzkglGl+2EbHmJHLV0wDlEDdbFL8rc/V8vYjU1ZIWwWkulnmuAogO3akOfyIfg9ZhcPA94WfHcrbXvKLQ457OHTeiBS/4XppP8vywdWoEMVgRbfC0kBKHGALh5nDRkZZnnz2hpmzc2eGzmoAv5/DGeDGZpmTybBKrlz1ptPKLa3oJ9P1gEiljUoFaO8xik7Cbo1KJWebzpQ/HdwFFB4/TQ3zNiXjGEw4aMnGqBPeHl7IgwAvcn9wLSedas6vIIqbWL9wh5GGrGtW6aUbGtdWHz6cgJGUzvTPU+sX7qQg6kpOx/OrcAzS+dAWpmlNPPb+++zQ4cwhxwWUcRh+wv3WezgAPEhcFya8SO6Ddu/A6/tefe29tw+wUdbz34f++13vt4d3/pb/+26O/9HVMwy0uzW3uPrwYWvhK+CNPrm9/sMlEPXmryOvSrpCIOsgLF788OnnfxECw/gr7736xsG33wTRAuXA",
  "LwVnRZLhTdRm3AH9hiJkks4hEDvGX3nv5TffBGGCM2mkk2w2TrVOLKx/cq1ZvwuKvMZ5rsuDbs8261826xeg/7kGidtP6Wf9FnWy9mABdZjLrS9/aJ1baDYaaxcuNesL0c/FoAQzQMs/qd9f++Txxp3vWo+X1pYXm/W/wreND/gfdfmH1ltqWEcqr1mOCfhCbqBP+G4nS/u5QtVke9nEs0cLxMOzrtlSgZ5nas8enZpgwIiF7qth2S+7FToYWby+WeYRn0vXhtYVzZmjQA3o0NWac3Nds55ffc2YNMW5qgHcoDR88K7zrjPBdqLUPkFgSAqeqa2ffdy6ers5V++albteUz8BMJBActyIl4PtVJdA3Lp3Hbo18HmKfy4PAT5PKddtFmfpu0PhKyEnPQTw8r9rzx4tyAmzrlnHra1fuAMrmaUFqYXvqboaEvU5JSCu2iWWiMspwZaJrmGMUg32ST6aCGH1N0k20Q7Bjh20+c8/z+iPEYVh",
  "0tmWqomY0VdQIh0QiQ3p7Cg4MWCdBNMTdJ6lATMCdxbYGOuVfU3EnkOu/OmaLdTWPv0KVjP4Gqiv8v3qyum1T7+iXlSsmdBx12w39knCK/WsMZF2IWb28FhljOKnjM2yyrpmFHgKbGyU9cJOrj76nNACnx7fSwQTB6qBwoYaNRZXHy60li/rUILB81Wn6Ar6KcFB34/RCIUFMEgWB5zHZX5JzHbgVRIUj9qASJ8aZqbtmdhGPVSiDYjAahulGyDYu0S7ojmj3y+ASeAR0h1w4g70L+gkC9/r1wfZnzfL8LeXpuAus5BlM5Y/vc+2aSUE2eIPYcVHXPyEHTXsqjnaCarsS836T4hLl6WOcqSHmo1xJMgxRYgjJJxDnGS4465Z08tzkpupdQq8gEdWQMv2spT8kcKRamP6hwEgAoOlMtoymI5XrZjA9Kqbn4etD0Er9zifyANXrVze9jIZ/gcbZXnvUO9htpf+pYYk9yu7yJlu/jFAZ4AR",
  "kgVbxdUAafSIQDB9jP4yZgzLZ45x1JoyfLeSy9tWedI1KoXcTMXyzeCbYS4Np5qN77kqjauHgVoH8ge42ecNPz/N0sBG1+SC+AbcTqECIBb4gG3Cr3RKBOGhRG3kcAdBE2oe8/EBqolzwhECzzG4QqSUl8IBYxSEUVgbORZg/pxRLptOYT94raR9g4ahjad9oeWQ35jHzPx+t1QynEI6BauX2sIKaDNn5YpbKkeaawxDioIhM+Rd4xtCzsctDNQqpC8GRKzokZlbYW2U1lwJE8hlhpcuOlmWy+WMypSn3U6QsN+quCXLM9PpiullWcX8A57MiFY5Bxd5vJrPm573e8Mp2GYFvsjg85cNy65WzOD5HzKHis7htBhTvzdgv3zNzRt2Wj2QASb/9/E338iVIVdF2oZm42Q3ALn/Vd8spUmvDToEIkXPP8/2VSrG8Zzl4b/pghDQMpJaFGKOJ3+lq0RrgfL8d/v2/8c7b8Xrz0FPCS4w1bLK",
  "+MHT3+FT1GdEpweYeGfSpGi8gOPw2V6SxIu261bS6UBeZ93M1zQJAjOETqLYZnql7gHB/jvDccxKOiCy3BTx/vtshyr/huWwEOkNTzvgEsCGAaPDHhXYCOsbDHUi2ATVwoNBXcILQxiKtLg4LfhI8UQKBSsJ06vSd8WdkV2GfaC0aLlOblPjbQPvF/DMAd+ED9Go8ETafHRTz/Lqg3t42/+smH0uNBtLQoTgJtiuWWKlcJH2CnGqOX+OTCrNxl/whn/arC8BM9Y4jQ2uAO8+1wC6NbF2da71eAkBiH4ERmLO34BoVb8u0VZzrjFRk3aqtQsnnl6+CkKOPgs5zbUfF1ofrYBI1hDS1Fw9ZH7mVkSY+2KzfhfhXQSDdBSwxnm+JnNnUMA7jazApWajLm2PaJ8W7jDJDjHcZjrlSjeMuHV4IM2TZCyk/+rsnGccNaNoSbusnoaBsoSpPFSOWcXjpKbJROiBoB6aaVAeCmlMvSK8OCRV6dA0",
  "JZWSYVt/MtMchQOL8CYKPznD86wpJ60isixDRcFsDS+jJxAiaj3lD9IkqKwUZYThjAzygqOEgiCkty/L/7actGoizDJE1q86fppzkFnW15uB7vuI+RZdkhJklKPrYsUtpWcZIZchPl4ty9LvZZmlmgLC6hOLaz0yOb9ildL6GFxL8msG0RUsbUdBqxEE4Y6CZkb+UlaUa4yl7tDBwRyQP/oAHZI6DcDhHQeCrwHJejhYnBvlMGQZX19cjyyfLDebGZYNPLz4G4Ahw5EEr4bZaPiRwOtCB4L+VI+DL+wrFsQeoV6PwPK5dk/aiHz9OGVReIPH8C+tYBak2SEu8fInXBU6xGid0r5UjkKLXlbLiJOLikxUevM/8VwrxMODawyCFTB64CQ4ynoVogymk/HjTl4hczuIzglaFFAz0+ZmHunKl0IwTDvGJkNuP7KBbg4SsOxlKbrpqw/u/TJ3E0WNX65+zFk70lalYhj3slnxLM9PTxney2Hu",
  "TZvCrIq4FGYcTJcEw86dw8oiDIdEAOAPQ4PQ1zEoTOCsu60b365d+FRlg589ugx6nzTqNsxcyfQ8cER//32GGJGD0t09zIetmMAGJkAMzVSII6vDv/5j1TL9CPfy/PNMvAjvZWiTUuTWyLeGM+O4NsJeHqBd3vleZdlSZIkyUsB/KVxtdAHxggOIEjC+osLBcjnGFROIYRxdUFZGWzyGs+XegdG56j6e29zC4VDfbW6DcqkAKtAzo/Pm07lPkClaFnzCfXCfAWP8Y/y/u+QX06EsFABStXIY0oXmHkRU6AoDMi+3mAOEVSsXmMcl25nB65GwXMp2oPwszO1sVJGynYKJmjn9NPb0sC061jbn6mK6d9cWVoTq+AxyY7fxj7vKjs91BOeYjg0se+Dn237l45q7Tt620DMgza2diD34RcuIm+RXqmiCEzoq6nOfT7hUkamNQuEABDcABCYID6mjlmdNWrblH89PQxqYVFYMJU+9/Jq6hV1R",
  "BgjkmvB5hm9lw+efZ5oEJF+MsQEQgtBBKn5enOqp5krP9NFslUaDG9I3F3xUQATOu5XCO07BBUUEfK9odyDAkI0m2OyE3EONRkeZ66gkBl4FnWcCtw0yks0yDRTspCb7DBpLU1c/CH7Bc2/aKnKdRsReOBp9JKlovCFxNOG5/AyAgvklmhv56nHVZMG0Td9Mbo2hOZzmpcTmpLIstD26JqFa9syKD/5Dac0GYknwhf3WKbwKYS3p9DE8msdyFsldwNrINbZQOZxRPwUucJT5fBJqn7hnxLsFcBtHTYAG1Ds6pHRXEVKrwBn4oLMIuLq52SJBmkOasJLwSnUECg6G6K6K3VVD3UngqV8OPr6MkF2Yn1BqQhBKmPQKSpdEQxXSqXSVyjLsLCOu/b+U8xDeIeCFRQPUaDNVWeeIaZa5x8m46ePcVLEon3PtAhr4xWL+znVt03B0oYA2uRDcHe4moPXkqN1wqWAHjJ+bNsAFhveI1nehNHKL",
  "6vkQq8wU3ntU8UoI9c2hCpwV5BjaKHkYhTYVNlROGVY1+AEHh88hYXScrBB2RrG5+HivUHgPgd8OglALT9YqMMvRDncwXWoBYkUUO2WZYx4DakLYaRuTc+3CIfngcAb7OUSAwt3XX4u+YzAsfKdOSjsDoCULHaq4WalrxXX7wdlT1P8x8qJwu0ABUdpyuKQoXkuRER4I4VGIh0MIxF74byDRkudfIF8GTaScSWAxvu41vrPR20/qkWFNXuFMYtiDUc4OKUgUEVEELek6FH2IqiQZ/rWYSek7lWVeW7T0K7GRjockOJ6+NnErE8+qBkYj+ZM70gLjH/JRAde/gpkGHyqVEJFPFXqBDicyxCqFhl7Afxb60c0WarxC2BwxWbXswjjGNnCBHE2iJqLNPYEkXnEKkmdN89f0729Y3+4Xevdk2L+x/r7BPYMvDOwe3APa8OBX0E1w/Q6Ji5Lqb/30Q9/638EIwy/LwAugEtIa9GsNBnsjDQb0",
  "Hl5ktcPBsBCpN8qS8X/QEtN1j+qcbR8q9YcVCxrxCIfwGEooXjM9z3XYIIt1xZcKGgNUQKSUEY6W3XsyQiejOAhAhqUh1pt7YTerZfWhpKv+s0cLrx9nL1f9asVkL1VMo0QO9IHLsLq8wUoejoDQ1ytgkI/6FSBeHIgAkRRb0Ham/eFRBpRBdu2KDBIKL3j2aKGvd+3qQuwsBxJmNhiZWZ8ctI9GPMzVbEIJqWtyfThroHlLURQQYBRL0ZnhodnJLKEl42ddaMqIBIY5CI134JQLeMvDCtFUSaaDH0lTWohe5ZXzTSZtwVlLk3YekwwpAx3KO4ejFHpzF8OKA96rI8zPwRpmIn0ecqSwodBfzifTR4LQhL3dMUQs2Fnu7i5c3cnvIUBsJMqD003j9NqVJ836Akdv8R7u+2w7nXouSJicyuSKbuWAkZ9Op8k9dDJGPpcKDzaZA9uGZ/r4YFhRR8xYTsGdAfOsa9sHXfBvRHdi1YldtJZq",
  "hDjtDTqpqwqEXz+ToFPfnZqyzXSKvCBS2dBEhL8D/J3JhI7qUTh1h4IolFSeC4KpPIVWpNAuCH+QMijLUhDImTqcwUliZCfcl6OZGIj4TLPsKHKyAgiAIT3Lc6EO8ZUbx1/8KIiHKJZmMQhFPjJkCIh49BalCuGhE/TsIB0sgFU8+j0EvNcyhzgYh9NShcMjTVNt57BjR0C5dWnTMc3CuOlXy6rOO+xBEzLqRk2waE8CLx00SQYxhEvrl1aeLn0XsWdyq+TIZGWM/j9isePJh2NsdvGWxFtqdKUY9RLFPKrmPOmX1s5qvckEFVfVTSaIf19u1u+T09vTKydBVdtYVINYVx/MNes3RD8iWJOYxMaHzfrJX7dQ9NsoFEY7+zrHnj1aZOqgG09+bp3+Mm5tFJt7raONLg9xESjwTCmi0LUERhLWFcqK5PK265nABR7i4B1OcWeMHVOqwivAZFMSAUy5Yq+CZ0ahgDpYoc9FQRzuC7Kf6Yym",
  "iu0AJKdEHimJBcK8Jmzy7/2SnbY0FV5eWjUpPsPKYW7QHmblhAfcb4AHy2gOCvomUeZZz+/umrVynl/jO5O3PeF/Bp1xDzR6BxGr0BrIVi3GNQFLtWiBvJSndlKOMdY1K8IGD8GTwzVuIQf/Xj0ugugaet3xFjpM/EPysavJbAE6RNC0U37qy09jG+PCYbgygAsrWoPIZNbD4DdfWHCq50cWvUDLeb/2bxhEHNvnpFGR3hq65wdfkkznGLbnbSibCu+2U7P1R/oGF8dOhFX1mNzLUmtX7wgIUZQTMQrUMMv6doEkrsVZav5gSQsEWYphQOHZa+X8TK1rVvq6wm90AeEe1NpzfZvkFZ8IKRkDuqX5JEqzppIAIqVwjg6KjgHRoMvsYDySaecs8BL6/cHXIVjV8XQhljOYIEYADxY4M8eIYDFekcD8YWoiKfiHtTaS/4zlZyF0h85lhvohvazumqsCA6fQ04Nue3moba+IrO0VDrG9pNSg",
  "jiXDQxw79UT3cOdOZQhcAxiCf8a1cfQVXCOwgqPuGRpGIoYmtfPYzQz1t1wzTAEUP0R6R9XKyRBd0DyIay8YL/kyg5Y0aC9ieKm551P0Ofofa8sHmnWxADIIqmAee7MIyEVpCbMyKx6PLWbUeogF88xbaQOE3bwFHOT77zMj52vBUvIXSWg+RgmrkVUJzaEvrXP6HgWy4PO0kQv8EV588cUXU5kc+vqY+90S5FlNT0YawJ9ttycrYyIwkHeLmxoLMFnigO/D88GXE1hF+POwcDpQr2bIp28nm4giaaNS4KSHUpIQCwMsDw8mCPJCIIIW8SAYlhJkJWEMEBec8wBxRvgrRK+8FUi7ghwH7tQcoyEViiQ6aQMa2o7D3nCQOSJFol8SidJm/2u8Aof1hImQMEx2iAsaymvRqzu5KxmCgpXUQaGskEqnIsFUiCfg2EtkVlp9cI91zRJOyvF3gsjHfl80Z0QumSC3QNBF0Zypqe53bSABnNk5",
  "piYjCLqBd23BAAzTOabS0uBjzkPoQChbG/6RsKHtt8yGOCDxWbFzbPXBzWb9JxLuR8hPnCcvGYfqeNpq6GEHhaoplqFZX6KcPk+/PCnDCELNOZboHJP8RrO+1LpxOfSVMl4bLi8V6hzn0zmmCG2fNOu3Zac84CH0EWBYJYtV6BtlB2hZxkZ6cPGCN23Z0PBCBwEf2irvtxFtaCEmKsnKkkm+FoGCs7IJO5ufPhI+EGDxgLwznWMjlFMa0q/w5GOT7jGRswbyWHR2zao0EmJHyBFcho4wjR1snH96+UtM/6OtkHZcu2YRu2+CPqXRaJvIc+P67fUbD0lY3jLajMnbE3uHAKsJBWmzcX7t9CNEcrd5RhwlmAddiMNZ3GB5Li7GJ6IKRFb03MJLl8rIuBBO/iLvXYd8T4A7MQM9GrRVBVfsRtWjSU8G6m2/7aUyGab8SOo54K827Z56g2OU3J08WqqUTScs0p9px+jmqBTKlhSMuNDSYBSo",
  "5vK2h4YfmZNFews1gAJ5nmvlFNgCF5vADBTZ+bCArujU2sgrNNa/UlwB5bUSvaXZbQ5yY0tIvJAaJtGEa70VHw3OXcv1o2AudTnFtxTYpfyCzwPrLjD4KKTALgWZdJRPQi4i4UhDqfoakl8Ec5w2MFcBhpiRUTjnuSVTujQICCrujMfNh1vS1AfuXuCMFyT7wWBXL2ceNSvHBcBtcgZksACt5QhlPoKxc5RNjPiVsRG/IBAU1YTgCgcHdQ1ds3Jue9kEi0uKRsoMNcQ3E1WG+IUxEd/MAHTFI3pW0hdutnc0b7DwbIZlc6lH8gthhR/UR+iadR2kK67DSUrX7A7XgbXTIomhCeQ+5o001RIpluiJ48ITp9bJjIpldCMZoicUvAyrIBQ6sCo0+i9XPxasQaBEhNUQ86gF5AdCtkd6/MpYSthWQldvYnv897+SjcgfiWMjFN1XlHWYwHOkx39Ojk204yfa8Q35IwHfoF/hGNZheWn1ISU8",
  "vKaqc2N5iO0DAsmxCBD1UsYBonCiyxS9s37hTuvsPyiQKQpIRG8thgT/xsR0twBL2JkRMj0BFEwU/EvVxn45+SNrnb0IprYfF5r1T/UMg7qQF/zZNStxZMBaxZ5ESj62hfM4MhkTBYaMo1ieWzI9rfKVch7/eBBZawU2BaVEw5fpItO/EarSJnLZbxO5HJz6AEI6NQjgG24nVVah5JNOtWRWrHwnJnDPu2CD9iG7cLHYyTAb17RrF8zKaCetgLpYMUiXL1PjfOvEV2RRYQcgU3NzrrH+42XK7bH2t+vgTK20RGfqUL5SVRZMFP54kRc6jH983ZuClQ9cYmrxB0dTxYvKWMHpoIIX/D3UJlLnjGWKxohATYcJFM1Jo02g41375mzr3BmJbv1pUJYk0BoZdRSTVYAf82GlcUBrpse0aVG5pM7IcUE1ud7Qq04iAg1SAkBqKaHGb86vwPGkPCap5txJ0FjLDA9sKJScQgwBs5wIIK1pxxNW",
  "r4dWUllZTNHdNYtcwPvvJ3EBoMyFs4G3R978nayvJtGQyBsOqex1jYOwz32gHDSg0gDRBABGacLFqcOTED0+kTy981+Q3VKKR+yXqx8/e7QgjAunMOThDohjc/Vmo9Gsn1x9MNd6eFN+AFlF4bpwqDb+fgOApH7r9/kKN84/rX/UhP+7puDvW4FgRdhvqJ3982NdZltq3X+88e31iMYN7cBRAfJ/x4xZmxjuUIQ1JPFCWOM/kqSrfKyYFnaSi5PbiIy36Vkh7VsQ3ahLIMjJXapEestdAsGNC/ZQnJOrmk942ZVy046qarNlQWRENYfO91Xkaas5cMCv5iAsIUve1Jrvsx/jch/xtq8Kb3umr/6ESudVV8quWQQCUvoAAMC8NueWumYhNF5BYYREm3NnOHJBIQS3fYdMrYG5cnm6jHCqDXwgojAVgVYx6vC0SmHhO50KqtWlpJCIrTP00das7NIdMNbOjiXUpI19Ut8vLhNDtMrOQF6H",
  "3YJV10V45QPvGMnYQRnIVIY7FL1mFqFq9/HEBgfdcuisCJFHRGrsCEXICInIzWSi68zawAHH6thwMhzw/ngQ1irQA3JcgB5mZWAcf5R071Q5fRuoAnt+w0XNSx7MrsPCETp4Ez0BR8zjBXfGCZ8Bgt/MHTGPkwEO2aQUeg/nLA8sU67Hw5uCA8CYidfSdPyXqBx2Wgqe3JnWmInMSsmr+uzRue5njy5RZtWCEhSN0dP7p43KfvAYLuTy/M99froXrFW9x14+cKA3I6KT9UHxRMrI8Ioxg1HheputYQ1VpzKsLNQOP7wQ+DTtuEGUtasoKjIiIhSYwlRGl1a1dA+6swHWv+sc62vOXVVYoBhuNZzmRSgWhoO1l6lnQqFuPAIpSZPgZjIKP6gjzzaQQ4IMYOM46gQifgmZg7uqck7kzLsVZKwPJH41WCb2yvPocJdr5odDGhKbLnFcMqy4yUBstFTSqEnaMjUG8ZQK2CqtALcsIAwqX9uc",
  "OwP5+xqnu2YBCM4VgguK9gCOifpzBJKgA1v77NGC6v8BbVDrAb7ARHGCVeqQcwmSVssAOB3XiQAJrikL2me4alntgIctxWEZ1LpoSZgjZU8S9K2GbbdTtxpIa/4V2lY+RZ6yGwM0lYxV9JTTYd5kVEsuE7iOKE32skPRTg6DqjPqXhLKgo2hLm/ZhhN4SXO9b6Gw3y0DwUv7x3yZFph/zGP9MF8Tl62Cd9Km36eEz3JuARIdzGDUiPBuCbJH7AcTz+ZeLypz4MfopUOaYPDlUJKiozuHmjRO5QCgOlQAQYzSdxtq31DqOz+soW6XJVZD3loWPF0THM4iEs496Qeaen2SlA5Twyfgz3C5a5bU3zyt61w9lQmuNK4O7TyV0VKgDLZ9bJT1w7lW3TngWYb2XUb2ZplDufbk9Vc125pghoVZOzV5fGbaVVTXyiw0b7ywFgPKvsJn4XWa0NFk12xEw9zGzk8Wfk4lBBrMaEsYBkntzJEACWy8",
  "8mPExN8+1Q9cPZDz+Z2l3amBJpinW4tzgBUoWg3PmWjOnVOypP55i6lsMfkpS9P5UFKbgp8Q13fwLZbXnI6B+Ikuqfks8z1x9hFdRE/CP+ut0jkW65Giq7mbc3PxhhKaYEA0ZaEX1QlIbYT6+RidcOyuGbaNmxbSPUAyqMdgSobkFzejWxp1WEVXJAWEtFAVRY3iin4kyRCeibWmx+lDlrbggTQhTCOEDLbt6ff/IYvK0QTHDE6A/wm3jHZuLUkuG0Lv3c7wcRSQGjd8KBQxzvARo0oLNjrG9qGrntv6TAy9ICobBaevcX79k2utx+fJwwbVgMqZV/SQzbmlAcxjfDka4tecq8uoPKzxtdxsLEBgQf1uc65OmkTIKyOZwvotkc32lKpVDPlecCMKHdhk5xRELWMbN+tr399e++CsqL+zAPpNxZ+pf3XlR3QtO/30Q0qyLXJoC9tSyK1Qcz+HQlFj6+dOrn/yLehRv32IK3QDJtNYxFTh",
  "dXQw4WOLCkzha9l68lkQxwEfLmNivAbGYiwRdqd6Tqs/X2jWF1cfLq7NnxA3PNo+6Hnj5uLGuZ9xDy+IfjhU61eWW2e+kbXDsNxYTNrwECKjZUdaPcONQ1vkDRLuELDcGu8MqH9GhAGANMOkvVljL2Zy+SybyW2RxZjJhZkMhcOI4w6iTMFMLp4t4JhUMRWoCFbBLl2zGLRO1fuIxoqgPkmIM6IpUFzi9SYtpwBh7GnT1jRIiPGkgvloewWzlF624AtEiKmNAlhBVr/WHeiQpLmHt+ITJLMES7HmkKJKxG4yEeeeoEzA+rWvsKjhlWZjkYucIDOEhS21ZA1UvEkbamp0We3F8KHkCAbBmzwJXuoPRve/v5XKYtymA3n8UtxCCbnloBKv8nvarUJ1rP7ugjVlQaBgyXKqvqk8Ah48nFo/xJgp0jGtPwlKoInAnAea3LYVAemfFY82E47gvaj9kPNBa6zlzQe/Gj4RIZMoAgnpeGhqsmJH",
  "kL9rdeV068oXRAmfgq3qFCE7TEys6GMa57EOHunuTwkah26B4rS06lfX7v1F5AGjBvefXrrRrJ8NQgeOWGUS+UVtLpmO4/33UaEQeR6kblGjFTxK0CheRfI0Ki1Np/AahVSIVQptEPSFGVr8nINhEzsATPlE6cs8hnlWCol96V29/z7TewqJtNizvErKjFGqFIBn5bDyiCrqjYo5Barnikk5RijxRziVjWJI5qKzgQuiZJTcyyZap87gWVh+OldffXIdpWhxn4OWOcPPiJoM8plnUgCYKHBB8hb1o9tggbwqaiDI7BYciDfxl8xDo2iLINH4RDgOT0bhTSTnGyayKhg3tWA6LykNldS7wfQwpHgFJEjo+fZxckQyadYkVzB0PRI7qUlc6JPEN1bhxqTM1kbpq19U0HVwPWZwUPr6w75v0Ar3KmassHmb3+i1q3cozFUbhLqJGyKG0RNuItgNKaNZ12xwaFXPg9WHD9H5AE5VLRrg1151",
  "UMKT1DWbNyDwjl/ONi5IXbP4AXjGXfyOPb14qln/rNlYpCSeFx8xXOPF9SsP1k4tBtVT51eeXrrRunFR8aLbEnCY8Y0DFzoNsdBRUR457tYGC9nkSzDQdgYN6H7jPE0fr3nM6UXOXtPEcN6Lr+lEhNePlZwGlXiNJGojKMld4P/n6iOTYyrJgmuEoGpSgcblI2seLQgk+sWSxdD1EvT7eKnZaKjXC8Uvnr6S6N3Grb8iUaMnj6HP8LoFwDx7tMDrjjc+iOrw5axAVcLFj1sIz2moL8r9/C9IkQ4cS+YaW4uJAXka3WE2O5qGbYOrBh0UJaRJ2Y9gr7fSV2+kL9zMS3EnRjs84tbG+EkFRL/Mk7ApdDaUNWu7BIB7D5EXY19/mxii/PQRBtRALrXldFNcct+u3jbRJvw6Au+DK6P68fo5B/wBd4i5RdiPGK0GiITUeq96g4XKI6pN9rmwF/byjBH3fEXcOybFvWNx4p4ORptQtAiHqOuY",
  "w0LoJocMWb/4dRzTQKLhwOsIryyOFnvJNZ9oVa+sSKYT0iwYki5ipVIM8klMSp2kc0kSvFCLsokuqrdT/Xx59cHc2sVTggFDzGmw6YpZHO30wNaXm/ZLdifnYkc738OatZ2QWhUQNqBxsAi8gqVCINl+4y4UDsBKzeATe25p7eo1woHrpxZGeowxxFyAklEdsqDpoJRF5Up65zgXtWRRRwhTkIm0OBerxS5k2tWUEGzbZoso/Ij5+HH+w8nxhYROI1GFKz+2zp1RHG5LvEJxp160SMg2vKxbLQaL9O8mLMLvaRsXxNDmU4l1hE9V0XFHb6R0qw/mgL+qf4fayNOkSly7+jd8uKx8hYUtGufXrjzgOi+FuiGpauCludasL78CE6JKFArdg9LyseUHW+eWmvXPOGuDHglE0MS311CBphJnIMvtGTGkz7eQwFxRKm9EFcMjk5CJZlNtMPa31GxgVXKAZRkgnatv/mHjfHDagsM/GSC5sjh9",
  "iBzCOi84gXh63qqYYKsXKS6Ei2dqO/cfqmOcghJC6PLJjV/tE9/oWXp+OXkeqnsIc8HG7Xut5ctcFT231Dp3pnUK9iAYEV7BOcCk2LR5ywnYYzkeziAL0KdxulRufomt7mGiEArSvyihCNmLZxUVgL+ZrkCmPKacNWksD0iJgU0/R/nG0jwPtNLBKDuUy0HtkMOR6rCQuSATSazIP0wFaqCgNz3Rk6q8VFPc2IbzZhHUezROoD7Qwv0Oq/pOgZGEyjP4naSoVDBV1LuMPMLQDrdJCfbN9Zgoth3ecva0QEkQURlo02d7KXHpUDgIUtGwZjYDDdiwJNDkokXcPGNUXNGdob2ZFJrfDOmJ+DHTXD3lMYQGcDJDb6OerZvOy7Dtw1sPJ93WrBBKLIiezqjzFC1RTEUjBlQijwNyQl38CQXKY0qO+2PaXkd2P6Os1jF9LWu/asWAqdz2kkEkKvdm0qHLxnj0au67UsyJqT8TpHp2VfWxKj7k",
  "XadoVUrpieTasFzujMqyWpxYslwLFKOx0GycQWz9AQ9KqC8+e3RtIhPxOQ9jZwEtnwAl3A0Y9xi40UwacLchzl71FuS1PmM64Kx/8qQmftXZANXOts8GYHFwoUVknp7MDCtOs/AwJ5Quh3oPa7WJwDeKVwvGdiTcwX8og1WIywzclQVfEfGNTdLUzExbvtmNTidD5YqJKtJh1VhPvOrkWGv5SuvEwrNHl7lO08unSznfzdTedVBbA+xx6K2obYz1hsVDdKwXer+wm8azRwuqXrVZX+5bffhw/coyyDqPF2VUTNj6LtIWSKYkzOnWl5DTXVTt5qjeKnMXj82Rg+kUNjkAlOJ6C8cgVIbKCx8HVTeynXOgWZpQttIzKAZegiXy1e2apW65B1V9mUywqtOQ9EbAvy+D9T7+bsGSvuvIg4DcuQJc7V1H5fE1XXPi5EGjkWWDvdKrD4tzQOuQTj2DMwIHLDhrqtSPeswFLkYovKzUUjx7dCoK",
  "TKBijxtcU7GrQ0d16s36LYRmSYEmPJhQtccNpdhuwuVTQu9zhs9G2MBuXk+SA4XFun+ZW+lbu9R4evHj1ZWvWidPIGUIhYw9iDPjXEI169LTD8+It3eb9U+4x0ikfLo8e4I0IQwEgtL7Y0lEUiEiMpkTWi7peT0ZLoJEHYmCT0FKeD1cAm6jUqrCdApvmyXIc1aBfLPKoSfXAk9qF8OLPsoq8sdw4CMu2SOlK82IKdoatlnxgUZXpPWsWb+rr4ao5igzFjCWruSAueG0Xv5QrDqwqNESRqHjjfslN/Qa+bYot2MR9ebXwqdSjicUYhkdtKJh2RFgxJyk8l3rnHwu5ad4x4vCmbWYkzXZiznfJc/eYs6sVNxK4JuKPpphaHBokWiMap1XoK6SYTmWM1V7OneZOB5xCfmhJmA3bn+GV/R+UjVeEoTlompfQeFL2eyu4rQE/U9o3sZtNzOj11NTQjBqkQJk4jyl9MUO1xhrX8JAcaZtl/yF",
  "WzT4srSNR8BsyW0CEijhs5KBJSGRcTQcQQ1aCMclKJEIOACmQ4TCt+HuJTMd8wEFL4if0RSXonL3cEcMw692pEQaJiSf2XryFhF+NXMQC6JMBFn8MDNnprZ+9nHr6u0JEaxQNqbMwOmfx2wgZAePl7k7GTJaKTXJyw4/4A9kDyqjCA87x0bKW48vjvOnJVaPo+qeHtbXnL+MnpAP8MJeg7DRq3OtG7f6+lsLn4K3H6ZQE8XV7gfNG+dbCyfJd1CPbfCwyBnzIMO8mBP83DkK5usQbSgbmFLK93KebeXNtAeVVndCQz3uAArGYSGX0WDPqJodriw+kFQDly/s/83Xb7o/5FfANH0KJj4m3b/yLY/epzMALt4Yev45fUkW9mePzq5//XDtbx88A6/QJM93vhlh6wvFwIfyHoSS8MoV2Mv6d7MhNjBYC3Jm8IQHSptUpIsXUXigPAnYPhXyJe6ahd3Qc2lgt1p2g8nK2Ca5Cnj+hYmkdAQj",
  "WgYAGDe+tmwuprqsBYcjo1ZFEikMeNoi/ENZB8gQJNYBmJdutEYOgeExMX0RJizS1yZudZKdzGj9C2O/XP0Ye8OVVtIsULJkehVepYhrJp0QnqYhyMugngc2BKkZeK5BvgYxe1nQM1DEZ75QvUNoKcIQKqkLaSdFEodIgAiPdozHgaCjdR3DToWishT/wG1HTakjjftVivqFADxAFjvDb5ykkKh/NuwqlI02FqJ2DoSRCBvPtsqe4gGH+6p+EnFuhi8A3Q1E41PiI5+EmurBSM/0QJz/eNF1fQ0L1rtmXx9/5b1X3zj49pvhEBI9CfjESNWOi5wasS2BXgzvyGuWg0F1YDS3QseuaovMjIGLg4p+pTJCkL0lJFknnz2q6xkZYyYEk3j5zTcPav5PE5qmD5Z/i+Q5JuuTtKyFKjuE6bE86BaRUIuNMHVoeLJzlA1m2pK4rln6hgiqlUWcORhNdDkcX7sPUExSmfK+3Vm8RWi53u9SOuy+",
  "XnH4xcK0pbvi84OWbyMXzJNpg+98QIoVT5iN67cR6HiaHDrbzbm5tasL/I9Pv2rOXdWDSzantAODEbKayHFsi852zaptp01ratof2tVP1mdoBikETAO8xm2vLcncDsFUyCPSTkEi281pezSTSLJON3GuBU7hYuYVQ+TaEJOYbHiTrlEpSCfz5OzczHFxaptG/qlOSuieqqeNjrggUEliNaarDKStXUbjcBbh1sJXEN/z0XfrNx6uf3IbNaDxyY0Fqewck0Z+jAwiT4K7aw+uk2E/6XvyNBlbv/QzegUsh67X/25S4vK2kwnqTGobTiIWCEK/GgTjfrUzvEbg7RTjyUGZreLzR2/nFoaT0tFl7Oya5TEZdItwPvyJwjVk2mSpU+l4cBXbZazbztLiyYlfV5AP68trF75R/HDKQBQ6uTOeUy1NmpVOCBhBb8eScWy0s2+34qOjkpLAOYfQ6B7NL0e9VWFANq7ffnp9SYUCaUvUFyigO+AK",
  "pCXeS5QJ9ZvSudXF26Gfy+efZ2ou1e3lnCzjd53BCJL2RFNOItUBIx06siZGjbbJLYYjog4pOeXkGKmj2qSQ3Hoc575BcCYi9Zaen/Ctl17W/IBJATFXp8aiqBWPnGidwIDFuSX6ihc4BrXg+af1B0IfpynJ4847KWZC+5UiHE2QtL78oXVugSAJLTfq+8CwwH13IPr6NLKrlzCK8VM5NckStEdmqb61Ly4363cHEeefDCxujZ/W/3EZ1g1yAqKzNZCAK+H+44JRw7wZ4mL8oxZKF44ETc8WLmENt0nwaVHmttW04WUlTrDcNkxQ0RPGZJiSEheXt5IHHPerckD8u+2A1OWWJ+Oqs3HbTSdgqrfePSIz2T//1XYAYru3PgLeATkC/9V2gYRCdUu59gjTxOTa4yUZy6TMFpGlvwO2j0JLVe04KogrrlsCZgjv6rNHC0LsWtq4uYgYhGzPyxg/dwf94O42G0+ajRvN+S8D5SlqAJ89OhWp",
  "tBviOJP8jNWEOzz9S1gFHqlaHSSro6rOORwMpavZWkYEqamW6q0EZ7FQfrvteuknJk2JCcmKOrYzEdmEytDVn68GweqxatAtRthMUlkiHu4gmWfY0fnvKY4FxoqvaRhok40KnN3UtO+XvaGenrw4QLkpdKDM5d1SD7baW63Yo6DiNJ28WzDfeftVTBvnmI6fnsxVK7Z0YPmXRML9MysOjSdzCDX64rerEHPv5tqX18n43Drxt6cXF9H2vvD04ncbN2+h5uIJyhNBzECbJDatE9+uXT3VrC9zuldfaj05sXGzrqVA03glFcxUOJmqqLnx4Z3W4gXMBk3dJ5nXFmiNeJ8HwE6J9ZXIxRZrY5wln+Xm/IVm4zre+LtB2Yy5euvs96sPTnNHWxpN+OOuPpjbWPjb+mVMWwNutfebjRUwuyBrwdfx3JnAtgsjwiJuzJ1oz36IRKz6nhoxMRgUQyBmCWeutmkYAU0bogQ6thTiMQlh6SKISLlM",
  "0ZwxyZAGV1WFGHdlCxArSJxv+RagZwUgQxUxiSBOZYyONc7gAyooGocRlHCTdrEOgRoBFFmx5CZKU7RcHJu5kotYglCNIR4zoJAm4DZDBE3Y5J49WhD+3/EWdLgUHD8G1nE9B6Bkti816+daZ8+sffYl91/nqTtuSQ6cDvzGd4+ajdNYPfySuPYxxkepBNWBJz//21D0Fb7Sr/hcQ4vH4wDcb334sPXgnkgWj3mgmdofd5VvfJDgnH8/EuWnKGR5vABQ7lqc47vGhSD5buOnxonVr/RUi6sos2XPHMhks3BuE88cyym6unMOAYwTxDLegS+OxpiMsiifEtfwELgpQ0kPp+iK9+R6Ck8CIjAB6iAPfN6SKfrpsL9p9AsNu4c+ggiZtnQK4ujjfD2Elwet6K/z8kjw8Njk8Gw734i29LqvOJKNTYfcgtP1tk+r5oQWu23SA2E5jLQFjlIxLiAc9GUKTkj9PpK7P6P0fZZngRdxUdqxAO+i",
  "H5rz19aXv23W767fWlFantqGhzXeJc2rzXfL2rVhFFkQvQ7D/Aak4iaLJW/leYk5hf/S80cyVOvcmfVvnoB778b1JViwuSWOXOfOQIwSKTnwby34CAB/ggm0VrB+2ffN+S9aSw8xZus0ouKGQNL316/W1y981TqraE3UUKb6fXKeaJ29/+zRqQDbYr7jt13IW4DWQUxOiMeOPG0qaCiGpzmvbFt+uudd//3s+88eLfVwMcujHCE8gEZ6PeGnwkAJImRfyApNyZZ9rCYluvZ29mTCaT5UscJAVxbMuvwqVHFVUi8Lp6fUb1OqYzB6ifZKAW0Wk3EOyZGzDKMyQZcj0mCHWvrSQq24SIDyEdL+EGfDUhnZj3/I8A8Hqar/G5767lCPlUVPvpqwtgBsPf/9bmFnVw9lvxEQZTJy1WkJycrYJ8tcUWFcOQOMWErx8SMfSd5LyXiN6IKaQ8Jpo2SGVg8c03hv6AsmhoRE4rWO0ArJmeM/bWeu",
  "KhTUqq3PHi2Gg/Lg6H6ISOtJjKZAr7au+OHlE6o2S4itAlUty/LdTSEywZZDLK+fV0jcm/eERxwk7IV8SlDmekipG52lQr10fiK1rbF4ThtXQfSbS221+DTvoUCqQl59XvGg43ohiN4Tnrnw9/g7r7xyYPxglkFUn5K8PJqK9nBGSfrr88Jy6id0rmJqQAeVjQ3xd3DYsG6B4Xm4YbyMgfgpBxBB04ruJi0uoQwrBMOPtADRJnpDgWchQcePJL4QQdfijWsX3lC6ZDU1E1KBA6QBGLw+annWpG3qYAdmOvD1sU3QKmM/fMzeLGC+IXqGFUH+uSSZW0yk2r7Gb0Jd365ZVWGwiQnD31codHIjo9k2Kw3USwnVSJmoxZcY3RHtLprc9MSCiNJY5B7LMSG9twIrMCgPbq8+fLh24ZstBORqlhy+IhMxm4G7JJYWSllPVdyqU+BZjSanMjyuCLNvkCUmNkfFNmzkMaY5uZMYhrLQenLi6Zcg",
  "hSrGOhMOfCfW5ganjqkp0wuH8ZsFodfUjXarjyFy5TXT81yH7WLRpJ5x8f6DiXZF4FyxEDxWquWQjIms1prLYdiaC1A6kIBXN7iKHtuuEU8jWr+29ulX6rocdO2CMKUWDF+zZxZygOhrW7OPUhy92vVLVTO550LVjOs45LMauVpjSJI0zUkxtkYwrj93WRaXQJ6N5vzKxo0P1y5887T+YO30F6pFENUpiakzE9LxDGqpZVSwFUS+DXssbwKFTkOHVOrecRHFKJJ1ESlC46y2YR0vN9rGOYK20+Ph1u4DNwbpIkerGFbdxaX73OJ6hnL1bG5KNseNo1QirQAFSMCYeuWHtYvfkKYJl4CQL39QiwKbXBrR3G84eRMm3LjXnP8rSoMrzfm7SRPuSJ49pELSqxziavAVAD130XZnuo+hAXw4siJ6QTu4850hLzMNFwp/7eQC5P50+I6ovmTwWsMb8A4e8rsufkqPYPpJf7T3xY4wX9G6ecjH",
  "+hpLRDEYQfkEnlUGWPjoVchkhiN9+i4mR/RyFbNQzZtpiASgTiAeQLJUvcDRUnmwTdomezhDL/H1/chTTpYu9ZXSpcKzPNn5Cd/rq2J6+XC23ljfcnBMpK8BrbO95E8Pf6ND/S8L53kdXdGMvO81r/twI1gl8OX2Xe7HHR4vxsOvAj6Jw2oUsuNyu1R76wEw+0GpzTHS6Mbd5K1q8gtUX5f31zp1+umlG3oRXfIjTKx+GOdLGJuMnuyTWy7qzhHJJlwx5ETSub6krNzJqY+IULatBIp5nz5Yu/gTxnqHyj5sQtMqRsFyO1FGGe2EAy3pGUkKMYnkAxEiSsP6FMPFLYTr5LZL7SaDhNesfWr7MEDERmwKVYx7cFssrvLtEq8S9y5/RnP1PHu0sLqysvbBWWRitoGNVaGOsyyaZ2+HjkL4cooVJtaaalqj02EtiWmJY6z7Wz/9MLD+9w9Cjnl9A5yD1nFKzPD7AQ+Hx2/jG9g1+/q+/3xv",
  "/OA7Lx144+C4CqsISEnwEYyBpD2+qqBKkYM0JqqWhIL/yDk9Li8aZu7FTQ4+jU//yxOsyoMA7dvmC4YPEFEpRdAi+fIi803C5cmIW0e3eY5vrXbIVmOjIdtX2XRIN4uX0SIqGApUGuwMof+ooR/Uzs36XUguAdY7xbr39Euqv5Cgg1ateqiPJlvk6spnVARZlZ+fPVo4cCxv2iSSc720KBscSau1DM6t4L73OZocHzfrj8N9NxqtG7dWHyy0btxVE6thKtHrGNK9iAkBfxJJzkjBf03qulEJ0Ail7ksyWWLiODWw5LSSZ++8mn+P8joIf8m7RCJk/3G2UVFHGr5aRPhP4h/3ybsBLcx3Rf93wfmxcRadDdtWXxBn0zzmGxXT0HCSF0YKGvZ5+vn5tau3WevGnafzH/0fwz1igSeZ+1vzmAGVs3NmIfeH8vPP9fUOr1/7vvWPz9nG6W9b9879H8ubto4YM0Zu2nCMI6H2uieNEOdLWXYE",
  "kekh+FNczUNHuN76cOSS8uB3SOulBMIHeuV3nZ1dPahSRi6ET17lWEI1npIDILaeczbEKSQT7uef17SESoWtvjaKNV4AhhRqhdCHgLH3v7ZvfPzAeGxmVNCuqfWIkjRtm0uSeZIkw9q0xnny9Y2R+WI4JgpuJkLDwa4F4M3VAzYmRIww0uIxWbXaFE7vmiV3wFjVIHdKDhIbLwUawSB7we2IC8iymttA88lQKtvEdNU4v37lh7WPvhIOFmAnVN2Xm3NLv5xcYWtX51Agv8slc0Slsmq2snkSiCW0rWPdG8g9SI7Y8ZCFlZZb55+FhUXjnMuRulrBessvOA8vU2aSTwxg06/RDHwRlzQpyYTwHdfnQYk0E9Ir3qEsjIg7ZfoPTM2IOJ6+DcAD3wQN67duXHw6f1tJ+0In5GPc/Rh5JCXP+FB42h8rECyjIvRrALTxIyaovSf8ZBImuU04+ECQuRTnI+kO0BdtaHXyWJf6EpAhShY1f06k",
  "IP8Lbsan8k7HqsEVt3bUybe5cCOTY1wMULKyJOwGpWT5gO4HSPz/syl8hVRoQJqt1sLnMr16c35l/YcPdrUWTpK3Vuwa8BwQSWuMSVzbrSt6NzKCZ/0Uj5yI+QAPcevx160T8zzr66MLzfqZ9R8vNevnE++4ggaff54boqqeWUEGsRyVYdc/uU0lTEJLBlLtpAxzEr2QPgYdO+R8b6uZ6igN1uqDlSDRe6RrfNLgDm1ct1yOqATDk/CmTdN/p2LTRMbk2VBBFG224PGoTCAZH3E3VXTljMKYXKQuiY4Xfmfkj1TLMHri+YihpZroDp2NtT8ummBftCAcC0d/2/R8t2J2MiOfN8v+aGfuD57rZI1y2bbyBlhVeuCB1LRPW4UCll6MKBOS5jeOXB8wCn/Hvf4UudfgtktuN2aWURZEiEkCeNNP7hkRG9zW1sLna1evtZYutomQks6FVBRIKqCU4AuwYYrQC/o7PkMrN7VTxivF9SBcWB1N",
  "bZHS6lgjvaC7v6CPl55EkTFuhuMOAqIzHj/CTVLiJRivtJeFqvzwpWrwnZI0IWpPyrDAksIdB2Kd13JgjhkSCqDMYTWT/jEaSDj2Ba43CMo+SMRAC0x/xywwrAUslgZLjBVJJBcKh7lQoAs3VWxhC6u2ndQHcMBxPQjPS4I0Ut0XgnFo70KV5dH9gexSXAnOXdNUW0WbguwpvSfF7qWnOpQeb7pFI8k5Xu9WnCzwa6NzFPw1Is5deCSeO7y+pFpJ0LnvNGc81OqJSMhSoZWj06y6l4BjV+CSYhVQuEK7lvwUUyrJ+c3/BZE4sIEKDw66A/APWLmoZvwDpcOp7wC+nz9GsO5DtfX6JQ0mwoZslJGZwLWDFReWFzJR7wgeA8zqCUWonUwGPYZkFr2yZ1Z8wB1pckMic9377wNKebUAvincx0W4I8FtJHDQcyfYe+mFxLeGuyDRjmUZd78ZUqfAvXOG1JSKQWFn2lJqLq2HQbI+lFqxWj2v",
  "PHxGiplaktqMkjEw8boFWKLW3lUWeth2Clo/9kBJJ8FCKK0yjMFB0mCmHaIMNMINMLQ7gR1K+hZx1zB1T8iQVAul4BWxd16+4tr2QRechHq35rNcMO3/6UUpmHYmzrEYT4Cfk2eACuWRAk8N4qZ8gY3zpOZUg3K36vtbMUvuURPvio93P+JXSyQ9pKnoUImr8NJSlyOetuUl0daScqMe5pCSats6LKizTvISekXd/Na6JXPqFvsVEaLhHOJ6RlCKi+QdxmnPpM9tZe+7DvflDTkDy/uctBQwBIxERLqCcFT0ovZxn5oiC3D0W3wlkATtduKFQP73ENqyQBGnXQvRX0JSe8pSycOuQ/o73RtR8b1Rs5CnYqxxXPGD2YPXL9zh8W5BM15QQ4lk5/7zQKe0lpqgiTzvglJ8C5Tg5ALOlQKhFLOzCqJRk0hS1nK3YIqpx9yrtriHbDrbQD/yHqZlLnBLzxFPXWI4uWrsYPHGD179wNom3Pnt",
  "Ic3NoYYOAwxJR9w6nOO+peF09eH3NRaLHCO6y3okD2Usnm0oB2Q7SJY7qZZFEqu+THQborS7ptccVuSnfJL8pK0njSoSsgU+2OTYO9grfXoPHQ7ceOFv6biLh6CmSQFt8fohTaHezfoO7w3JZwh9hPOPhuGoE1HIq011WNPJxnzg5FSzAPJzqpNwX+awcBHO6A7QHI3SSvG8NVwO4qKGXDzE3q86UAuG+/L09WbCPtJUcSHqIS2eK/7R4jgDiwou8xwpZ5RUaiJFrFvEVdDTBO4QHnm64KD75S4JU90y2iiX1MTCupTSc+jdQ+8efvc37+59t+fdd4cO87gJMUiQ5wwXnziCVPixFLtSIbCIu8kr3M19HdC7ikZZravA5yoWnUw94JTFf4+QBUeYOTLhcblfcv1+X3PuapxJ5FaSbKhDIMxukmAJABAW1T+ev8i0scXHd6L50sf3kukI+XzxWcqlffbosoyG2ri+hFNfojXg5Xdl2VTU",
  "zqI8CRlyz0blyYmQPDlpaOWC8R5J+2MJ0bJecboEfgZHIKUTG2K9CZFHvNqLUYgXubWpJRiXl1s//wVqMtz7K+4irwcWPfC4AjAS92ibX0llahvXl9avLKuxk5LyiCAPW9YR0BQlmZxn/YkyJdlBqdOIyoC09fqtjLuGMKLSDxtTTZOR6xRQrftJxkBeGfSWspUqKg6H4AYxXRgU6BlBdAwinwyvba+GlhABiydrYXwUMndqEnB8gJ+UQwS7ppBGrgUW1FH+bCupTdruJE+r/zvbnUwf+vfxN9/IeYj2reJxIjFZHsHUlzkMBeZBATzEUmH9bkrK9zx0jo2ygpuvlkzH5/qAA7YJv9IpA+Zn5EDVzkbZO2+/xhu8iaUs3nn7tTQAhm0K7owDidshcCUuhxe4MbwHPpEiPqqG6ucJ+BbnLWinZ2INbrfqp2kdYNiKedQ9ogxLMGWAkvUKl1IKDLXdvAHldyrGlAmlyl71zVL6tfH3frdv",
  "/3+881ZWkMpA3YExmmq8Z00IF3LHuOZcbpn8HZeNR1naopqLBzTxEDQ2THShqDNdQb0G2OOXLdt82zQUxQhILHx5o/yHHmAug5bYKMNjgvQ/DeUAvKrtKw64CAil0rM8/DctdUmUQj78EvUEmQzzpyvuDIKKST/S4U6lOLS6cjpkg4y3O9aXKa0j6Ot+Xm7Wz6ydvbIV5jVcOgNrAyAasEplt+KPw8VIg6ZAYDUyqYXusK6FUsPQKxU1BDgKuB75vnHnXrP+ZOPxIxFVrEUKB9hCKH5yFdMo7KNqE8VArMVTRtaUTVhPWO10rGT6/vuabkfgeE1Yje6ONKTANrU342x3a6LbMlm1bD5NCBAGdAzBjxBRCqBXvZSCjDeFJW4vg9VEA9IWFjNYnKA4tCprza8o2G0lVr8VNUap4hyevrtrCysia9wmFcY2X0ZCiznD86wpJ22Wyv5xepEBMgByx5AakwicurwNYQgTFlGN0JWuz/FFMH5v",
  "2mUe2CqjWadNu9yuFBcmeoFGSmVbOYyayWVMBOmhMBybQWWugdp7wRA2FrnHoUzeMtdYu/8R1klZbl35onX2I1EuIDBZr/2dqqSIzCbS1s6hG0AXBWCO6tD64cVW41KQvHvEDXycbWssVGxUZk/EKEUqmLKYGKsItzA+XFGc/Jsx3yreBsKhKNAq/fOOWKRqwlzhyjy5g2P9voCMMkLcjLpa6mlkgvQx2m7Orygl5FdE3hsYHnNTnVF8IAQcIz1i5WGLgkUPYjsvUWnkLW0UZaEKYlSjixmzCJr5rn473uxW/z5IkYHnbfXhEhQUEzXaeECe6tsKPiLhGEtyym2d+WzjzhfN+ZOInRYwHdCneHruoFfU/dbcpY2/XuWFi2H0nxKhh6xbqtVu7ozuo3r76Ykz4mqRk9Xy2qdfkaxGXm3c+hf5kEoT8bfzK/THxo8nsMDwEmbYxMvZWEAf2LtqJpE2exy9/8ruVsO7q7bmwrMeERzsD3pD",
  "c7+9e+fAiUkkPVLiVZZWH94Aj9zGYnP+C/o8dDzQ0+Xk6oO51sOb0q8PKrBpfnqhe3TyNK1ouACMdJ6Woj+Wb159/KS18LVEU+qxl5OROFBiDXYAsuCBJ9mPl8mveu1v14N5BpbnW/F3jWB9evHjp6BzWEAXrgDQX07+yFpnL4IqmoghrvP6lR8QQyyvfXwG/RjPK2Qw4Ty2lpdWH3IPSm1qlzFukfKkRBP/q7gFx+ba9cb59b9/Eb/6WOogwB5tiUi7U8aJCqUMo/OFpEVU5xJ3RCsatIRWgburD25K7IDH564wYy9vnIIKXSw2GWDetUHB2TlGY68+uMcdZeObF6HgNpUJ27X68OHqylerDxbbfkHZBtVb27Y5T3WICbxFICQ1V6hpdKsDfM/LbKnLF2xy43y47iWs1201a1Fzbmlg/cId1LxE0WZ9Y/FbyCl1kSMqNdhAZt3b5OSr+cr7V1d+xEU8/fTDM5o6KjiJcCs2btZ5zUDI",
  "WAfYGid4F3vh46yfO7n+ybdwNb59iDO6QZWeqGYdXw0irXP1jZuLG+d+xkW4QLoj0Wx5/cpy68w3McutH3M1fkGJa1gE18jAR1EmpUw+9f90PfW7ajH1WH2ZkmQv5KwI3prR6BUlocPj7R46rd5daG0EaxatrrsY9uDWYkeW1KIkopLod3hyT/Nsg7yuQKRuK/j5PuDbG0C1pPlag3KTU3mqbKhgubWLHyLTeV9WMcAqsbexKMFtrulLupo0NeE1vMT7apznfDN5p5L0jg8RiDAE//JKg4CgYzkEAnr14eLa/AmxRn+VRWuUwowfiONGjNJd6ZqvFmgN5jm/wqun8oSmCufROA/Pccuh7rJwKI4pIz23xAszNxpqdVZ+AxQbX7RYdbS3pEsdzZJJ/Wz7UsfWPpT3I2YYcTPaZAeOXhHVkLnFfJu4r23xsw7fLZlbFj8L0P3a6Qvrt5+E+voVeTmVXJyfNuv3Vx+ckdkN/zdTcLbF9XjZ",
  "727OJifvui6PRoqogJzWtkoKcP4JlVCajfNqPYL5lbdeelno228liXvq7FrLp1pPSLF3k/wuBC28mzxZxZ1Ci2f5tTEs6IvfJhQkVUsIBvkUJA4RAwJZFSOXXYVvSyEx839uNn5Gdub7/wd3WW8ZBYHpfmWsTLTMrmKBvCb0t/eTY5vqtyL6UVkyIRotIjbxLrK9/+PBICwp4iNVC+OWja8vrD6+3mx8sN74iZQxSJ3Por06JlhHUL1L8hbDLRPyJawpD1H9Ev8+k3ifNUVLcvQonKXNAkhbCzfWLt5LyHYVvmFKOt5AD7jx93+0Fi9wJaC0IEX9LZ4DreNL9hTjFeTDLlmbJzpFO6t0slEcKbiGXOYD3a5yOd4fSnhD6UMGWevi1O3cxihGoPR9gfYaNvdyG2VeUu4xYTEI1LFgvAO9qkcmrjdc3wxpVglNdLD/fwfmQaBdgEu5zlRDOHrQVT2MdePitDYNw0vAQlqE2v/b6Gi4g5fI",
  "zgx3jPR4+YpV9sc6RnoogcVIDwAz1vF/A3u/8gNnNQEA"
].join('');
