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
  "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"
].join('');
