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
