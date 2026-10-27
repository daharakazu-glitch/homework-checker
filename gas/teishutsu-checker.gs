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
const TMP_PREFIX = '__tmp__';
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
    // 先にぶつかる名前がないか確かめてから変更する（途中で止まって仮の名前のシートが残らないように）
    const leaving = renamed.map(function (c) { return c.oldName; })
      .concat(oldNames.filter(function (n) { return keep.indexOf(n) < 0; }));
    renamed.forEach(function (c) {
      if (ss.getSheetByName(c.name) && leaving.indexOf(c.name) < 0) throw new Error('「' + c.name + '」という名前のシートがすでにあります。別のクラス名にするか、そのシートの名前を変えてください。');
    });
    const tmpName = {};
    renamed.forEach(function (c, i) {
      const sh = ss.getSheetByName(c.oldName);
      tmpName[c.oldName] = TMP_PREFIX + Date.now() + '_' + i;
      if (sh) sh.setName(tmpName[c.oldName]);
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
      const sh = ss.getSheetByName(tmpName[c.oldName]);
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
    if (log.length) {
      sh.getRange(sh.getLastRow() + 1, 2, log.length, 1).setNumberFormat('@');
      sh.getRange(sh.getLastRow() + 1, 1, log.length, 6).setValues(log);
    }
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
  // 「3-6」「2/1」のようなクラス名・提出物名が日付に変換されないよう、書式なしテキストにしておく
  s.getRange(2, 1, MAX_CLASSES, 1).setNumberFormat('@');
  t.getRange('B:C').setNumberFormat('@');
  // 新規スプレッドシートの空の「シート1」を片付ける
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 2 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function readSettings_() {
  const s = ss_().getSheetByName(SHEET_SETTINGS);
  const vals = s.getRange(2, 1, MAX_CLASSES, 2).getValues();
  const out = [];
  let repaired = false;
  vals.forEach(function (r, i) {
    // 以前の版で日付に変換されてしまったクラス名（例：「3-6」→ 3月6日）を元に戻す
    const fixed = dateClassName_(r[0]);
    if (fixed) {
      repairClass_(String(r[0]), fixed);
      s.getRange(i + 2, 1).setNumberFormat('@').setValue(fixed);
      r[0] = fixed;
      repaired = true;
    }
    const name = String(r[0] || '').trim();
    if (!name) return;
    const count = Math.max(1, Math.min(MAX_STUDENTS, parseInt(r[1], 10) || 1));
    out.push({ name: name, count: count });
  });
  if (repaired) writeTasks_(readTasks_());
  return out;
}

// 日付に変換されたクラス名なら「月-日」の形を返す（Date 型のセル、または Date を文字にしたもの）
function dateClassName_(v) {
  if (v instanceof Date) {
    const tz = ss_().getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz, 'M') + '-' + Utilities.formatDate(v, tz, 'd');
  }
  const m = String(v || '').match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (\d{2}) \d{4} \d{2}:\d{2}:\d{2} GMT/);
  if (!m) return null;
  return ('JanFebMarAprMayJunJulAugSepOctNovDec'.indexOf(m[1]) / 3 + 1) + '-' + Number(m[2]);
}

// 日付の文字になったシート・設定を、正しいクラス名にまとめ直す
function repairClass_(bad, good) {
  const ss = ss_();
  const names = [good, bad, TMP_PREFIX + bad];
  const sheets = ss.getSheets().filter(function (sh) { return names.indexOf(sh.getName()) >= 0 && isClassSheet_(sh); });
  if (sheets.length) {
    // ✓ がいちばん多いシートを正しい名前にする。ほかは、空なら削除し、記録があれば「（古い）」を付けて残す
    let best = sheets[0], bestN = -1;
    sheets.forEach(function (sh) { const n = countMarks_(sh); if (n > bestN) { best = sh; bestN = n; } });
    sheets.forEach(function (sh) {
      if (sh === best) return;
      if (isEmptyClassSheet_(sh)) ss.deleteSheet(sh);
      else sh.setName(uniqueSheetName_(good + '（古い）'));
    });
    if (best.getName() !== good) best.setName(good);
  }
  const props = props_();
  ['mailStop:', 'lastMail:', 'board:'].forEach(function (k) {
    const v = props.getProperty(k + bad);
    if (v !== null && props.getProperty(k + good) === null) props.setProperty(k + good, v);
    props.deleteProperty(k + bad);
  });
}

// ✓ も名前もメールも入っていないクラスシート
function isEmptyClassSheet_(sh) {
  const rows = sh.getLastRow() - 2, cols = sh.getLastColumn() - 1;
  if (rows < 1 || cols < 1) return true;
  return sh.getRange(3, 2, rows, cols).getValues().every(function (r) {
    return r.every(function (v) { return String(v).trim() === ''; });
  });
}

function countMarks_(sh) {
  const rows = sh.getLastRow() - 2, cols = sh.getLastColumn() - 2;
  if (rows < 1 || cols < 1) return 0;
  return sh.getRange(3, 3, rows, cols).getValues().reduce(function (n, r) {
    return n + r.filter(function (v) { return v === '✓'; }).length;
  }, 0);
}

function uniqueSheetName_(base) {
  const ss = ss_();
  let name = base, i = 2;
  while (ss.getSheetByName(name)) name = base + i++;
  return name;
}

function writeSettings_(list) {
  const s = ss_().getSheetByName(SHEET_SETTINGS);
  s.getRange(2, 1, MAX_CLASSES, 2).clearContent();
  s.getRange(2, 1, MAX_CLASSES, 1).setNumberFormat('@');
  if (list.length) s.getRange(2, 1, list.length, 2).setValues(list.map(function (c) { return [String(c.name), c.count]; }));
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
        name: r[1] instanceof Date ? Utilities.formatDate(r[1], ss_().getSpreadsheetTimeZone(), 'M/d') : String(r[1]).trim(),
        classes: (r[2] instanceof Date ? dateClassName_(r[2]) : String(r[2] || '')).split(/[、,，]/)
          .map(function (x) { return x.trim(); }).filter(String)
          .map(function (x) { return dateClassName_(x) || x; }),
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
  s.getRange(2, 2, rows.length, 4).setNumberFormat('@');
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
    const name = dateClassName_(c.name) || String(c.name || '').trim();
    if (c.oldName) c.oldName = dateClassName_(c.oldName) || String(c.oldName);
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
