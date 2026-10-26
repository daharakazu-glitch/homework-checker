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
