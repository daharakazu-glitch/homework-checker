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
  const html = HtmlService.createHtmlOutput(APP_HTML)
    .setWidth(1200)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, '提出物チェッカー');
}

// ウェブアプリとして開いたとき
function doGet() {
  return HtmlService.createHtmlOutput(APP_HTML)
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
    for (let col = 2; col < vals[1].length; col++) {
      const id = String(vals[1][col] || '');
      if (!checks[id]) continue;
      const m = {};
      for (let r = 2; r < c.count + 2; r++) if (String(vals[r][col]).trim() !== '') m[r - 1] = true;
      checks[id][c.name] = m;
    }
    return { name: c.name, count: c.count, names: names };
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
      const names = [];
      for (let i = 0; i < c.count; i++) names.push([String((c.names || [])[i] || '').trim()]);
      sh.getRange(3, 2, c.count, 1).setValues(names);
      const last = sh.getLastRow();
      if (last > c.count + 2) sh.getRange(c.count + 3, 1, last - c.count - 2, sh.getLastColumn()).clearContent();
    });
  });
  return getData();
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
      return { name: c.name, count: c.count, names: c.names || [] };
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
    const checks = state.checks || {};
    list.forEach(function (c) {
      const sh = ensureClassSheet_(c.name, c.count, tasks);
      const names = [];
      for (let i = 0; i < c.count; i++) names.push([String((c.names || [])[i] || '').trim()]);
      sh.getRange(3, 2, c.count, 1).setValues(names);
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
    sh.getRange('A1:B2').setValues([['番号', '氏名（任意）'], [ID_MARK, '']]);
    sh.getRange('A1:B1').setFontWeight('bold');
    sh.hideRows(2);
    sh.setFrozenRows(2);
    sh.setFrozenColumns(2);
    sh.setColumnWidth(1, 50);
    sh.setColumnWidth(2, 120);
  } else if (!isClassSheet_(sh)) {
    throw new Error('「' + name + '」という名前の別のシートがあります。クラス名を変えるか、そのシートの名前を変えてください。');
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
    return { name: name, count: count, names: c.names || [], oldName: c.oldName ? String(c.oldName) : '' };
  });
}

// ===== 画面（index.html の内容） =====
const APP_HTML = "<!DOCTYPE html>\n<html lang=\"ja\">\n<head>\n<base target=\"_top\">\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>提出物チェッカー</title>\n<style>\n:root {\n  --bg: #f4f5f7; --card: #fff; --text: #1d2330; --muted: #6b7280; --line: #e2e5ea;\n  --accent: #2563eb; --accent-soft: #e3ebff;\n  --c-collect: #2563eb; --c-collect-bg: #eef3ff;\n  --c-few: #b45309; --c-few-bg: #fff6e0;\n  --c-over: #c62828; --c-over-bg: #fdecec;\n  --c-done: #15803d; --c-done-bg: #e7f7ed;\n  --on: #16a34a; --on-bg: #dcf5e4;\n}\n@media (prefers-color-scheme: dark) {\n  :root {\n    --bg: #13161c; --card: #1d2129; --text: #e7e9ee; --muted: #9aa1ad; --line: #2e333e;\n    --accent: #7aa5ff; --accent-soft: #24304f;\n    --c-collect: #8fb2ff; --c-collect-bg: #1d2740;\n    --c-few: #fbbf5a; --c-few-bg: #3a2f17;\n    --c-over: #ff8a8a; --c-over-bg: #3e1c1c;\n    --c-done: #6ee39a; --c-done-bg: #173525;\n    --on: #6ee39a; --on-bg: #1b3a27;\n  }\n}\n* { box-sizing: border-box; }\nbody { margin: 0; background: var(--bg); color: var(--text); font-size: 15px;\n  font-family: system-ui, -apple-system, \"Hiragino Sans\", \"Noto Sans JP\", \"Yu Gothic UI\", \"Meiryo\", sans-serif; }\nheader { position: sticky; top: 0; z-index: 20; background: var(--card); border-bottom: 1px solid var(--line);\n  padding: 8px 16px; display: flex; flex-wrap: wrap; gap: 6px 14px; align-items: center; }\nheader h1 { font-size: 16px; margin: 0; white-space: nowrap; }\nnav { display: flex; gap: 2px; flex-wrap: wrap; }\nnav button { background: none; border: none; padding: 7px 11px; border-radius: 8px; color: var(--muted);\n  font: inherit; font-size: 14px; font-weight: 600; cursor: pointer; }\nnav button.active { background: var(--accent-soft); color: var(--accent); }\nmain { max-width: 1200px; margin: 0 auto; padding: 14px 16px 40px; }\n.card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 14px 16px; margin-bottom: 14px; }\n.card h2 { font-size: 15px; margin: 0 0 10px; }\nh3.sec { font-size: 14px; margin: 18px 0 8px; color: var(--muted); }\n.row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }\n.f { display: flex; flex-direction: column; gap: 3px; font-size: 12px; color: var(--muted); }\ninput, select, textarea { font: inherit; color: var(--text); background: var(--bg); border: 1px solid var(--line);\n  border-radius: 8px; padding: 6px 8px; }\ntextarea { width: 100%; min-height: 110px; }\n.btn { font: inherit; font-size: 14px; border: 1px solid var(--line); background: var(--card); color: var(--text);\n  border-radius: 8px; padding: 6px 12px; cursor: pointer; white-space: nowrap; }\n.btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }\n.btn.danger { color: var(--c-over); }\n.btn.small { font-size: 12px; padding: 3px 8px; }\n.btn:disabled { opacity: .45; cursor: default; }\n.muted { color: var(--muted); font-size: 13px; }\n.chk { display: inline-flex; gap: 5px; align-items: center; font-size: 13px; cursor: pointer; }\n.hidden { display: none !important; }\n.empty { padding: 26px; text-align: center; color: var(--muted); }\n\n/* status cards */\n.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 10px; }\n.tcard { text-align: left; font: inherit; color: var(--text); background: var(--card); border: 1px solid var(--line);\n  border-left: 6px solid var(--c-collect); border-radius: 10px; padding: 10px 12px; cursor: pointer; display: flex; flex-direction: column; gap: 4px; }\n.tcard:hover { border-color: var(--accent); }\n.tcard .top { display: flex; justify-content: space-between; gap: 6px; font-size: 12px; color: var(--muted); }\n.tcard .name { font-weight: 700; font-size: 15px; }\n.tcard .count { font-size: 13px; } .tcard .count b { font-size: 20px; }\n.tcard .miss { font-size: 13px; }\n.tcard .meta { font-size: 12px; color: var(--muted); }\n.badge { display: inline-block; font-size: 11px; font-weight: 700; padding: 1px 8px; border-radius: 999px; }\n.st-collect { border-left-color: var(--c-collect); } .b-collect { background: var(--c-collect-bg); color: var(--c-collect); }\n.st-few { border-left-color: var(--c-few); } .b-few { background: var(--c-few-bg); color: var(--c-few); }\n.st-over { border-left-color: var(--c-over); } .b-over { background: var(--c-over-bg); color: var(--c-over); }\n.st-done { border-left-color: var(--c-done); } .b-done { background: var(--c-done-bg); color: var(--c-done); }\n.bar { height: 6px; border-radius: 3px; background: var(--line); overflow: hidden; }\n.bar > div { height: 100%; background: currentColor; }\n.legend { display: flex; gap: 6px; flex-wrap: wrap; }\n\n/* check grid */\n.tablewrap { overflow: auto; max-height: calc(100vh - 260px); border: 1px solid var(--line); border-radius: 10px; background: var(--card); }\ntable.grid { border-collapse: separate; border-spacing: 0; font-size: 13px; }\ntable.grid th, table.grid td { border-bottom: 1px solid var(--line); border-right: 1px solid var(--line); padding: 0; text-align: center; }\ntable.grid thead th { position: sticky; top: 0; background: var(--card); z-index: 2; padding: 6px 4px; min-width: 76px; max-width: 120px; vertical-align: bottom; }\ntable.grid .sticky { position: sticky; left: 0; background: var(--card); z-index: 1; padding: 0 8px; text-align: left; white-space: nowrap; }\ntable.grid thead .sticky { z-index: 3; }\n.th-name { font-weight: 700; font-size: 12px; line-height: 1.25; word-break: break-all; }\n.th-sub { font-size: 11px; color: var(--muted); font-weight: 400; }\n.cell { width: 100%; min-width: 76px; height: 34px; border: none; background: transparent; font: inherit; font-size: 17px;\n  font-weight: 700; color: var(--on); cursor: pointer; }\n.cell.on { background: var(--on-bg); }\n.cell.late { background: var(--c-over-bg); }\n.cell:hover { outline: 2px solid var(--accent); outline-offset: -2px; }\n.quick input { width: 90px; font-size: 18px; text-align: center; }\n.lastmsg { font-size: 13px; min-height: 1.3em; }\n\n/* call view */\n.callrow { display: flex; gap: 10px; align-items: flex-start; padding: 7px 0; border-bottom: 1px solid var(--line); }\n.callrow .who { font-weight: 700; min-width: 70px; }\n.callrow .what { flex: 1; }\n.callrow .n { font-size: 12px; color: var(--c-over); white-space: nowrap; }\n.note { font-size: 12px; color: var(--muted); background: var(--bg); border-radius: 8px; padding: 8px 10px; margin: 6px 0 10px; }\n\n/* print preview */\n.pages { display: flex; flex-direction: column; gap: 16px; align-items: center; }\n.page { background: #fff; color: #000; width: 210mm; min-height: 297mm; padding: 14mm 12mm; box-shadow: 0 1px 6px rgba(0,0,0,.18); font-size: 12px; }\n.page h2 { font-size: 18px; margin: 0 0 4px; }\n.page .sub { font-size: 11px; color: #555; margin-bottom: 8px; }\n.page table { border-collapse: collapse; width: 100%; table-layout: fixed; }\n.page th, .page td { border: 1px solid #555; padding: 2px 3px; text-align: center; height: 18px; overflow: hidden; }\n.page th { font-size: 10px; font-weight: 600; background: #f0f0f0; word-break: break-all; }\n.page td.miss { background: #eee; }\n.page .legendp { font-size: 10px; margin-top: 6px; color: #333; }\n.slip { border: 1px dashed #777; border-radius: 6px; padding: 10px 12px; margin-bottom: 6mm; height: 59mm; overflow: hidden; }\n.slip h3 { margin: 0 0 4px; font-size: 15px; }\n.slip ul { margin: 6px 0; padding-left: 18px; font-size: 13px; }\n.slip .foot { font-size: 11px; color: #444; }\n@page { size: A4; margin: 0; }\n@media print {\n  body { background: #fff; }\n  header, .noprint, main > section:not(#view-print), .toast, #modeDlg { display: none !important; }\n  main { padding: 0; max-width: none; }\n  .pages { gap: 0; display: block; }\n  .page { box-shadow: none; page-break-after: always; break-after: page; margin: 0; }\n  .page:last-child { page-break-after: auto; break-after: auto; }\n}\n\n/* misc */\n.toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); background: var(--text); color: var(--bg);\n  padding: 8px 16px; border-radius: 8px; font-size: 14px; opacity: 0; transition: opacity .2s; pointer-events: none; z-index: 60; }\n.toast.show { opacity: .94; }\n.overlay { position: fixed; inset: 0; background: rgba(0,0,0,.4); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 16px; }\n.dialog { background: var(--card); border-radius: 14px; padding: 20px; max-width: 460px; width: 100%; }\n.dialog h2 { margin: 0 0 6px; font-size: 18px; }\n.choice { display: block; width: 100%; text-align: left; margin-top: 10px; padding: 12px 14px; }\n.choice b { display: block; font-size: 15px; }\ntable.list { border-collapse: collapse; width: 100%; font-size: 13px; }\ntable.list th, table.list td { border-bottom: 1px solid var(--line); padding: 6px; text-align: left; vertical-align: middle; }\n.help h3 { font-size: 15px; margin: 18px 0 6px; }\n.help ol, .help ul { padding-left: 22px; margin: 4px 0; }\n.help li { margin: 3px 0; }\n#loading { padding: 60px; text-align: center; color: var(--muted); }\n</style>\n</head>\n<body>\n<header class=\"noprint\">\n  <h1>提出物チェッカー</h1>\n  <span id=\"syncState\" class=\"muted hidden\" style=\"margin-left:auto;order:3\"></span>\n  <button class=\"btn small hidden\" id=\"btnRefresh\" style=\"order:4\" title=\"ほかの端末で入力した内容を読み込みます\">↻ 最新に更新</button>\n  <nav id=\"nav\">\n    <button data-view=\"status\">回収状況</button>\n    <button data-view=\"check\">チェック入力</button>\n    <button data-view=\"call\">声かけビュー</button>\n    <button data-view=\"print\">印刷プレビュー</button>\n    <button data-view=\"tasks\">提出物</button>\n    <button data-view=\"help\">使い方</button>\n  </nav>\n</header>\n<main>\n  <div id=\"loading\">読み込んでいます…</div>\n  <section id=\"view-status\" class=\"hidden\"></section>\n  <section id=\"view-check\" class=\"hidden\"></section>\n  <section id=\"view-call\" class=\"hidden\"></section>\n  <section id=\"view-print\" class=\"hidden\"></section>\n  <section id=\"view-tasks\" class=\"hidden\"></section>\n  <section id=\"view-help\" class=\"hidden\"></section>\n</main>\n<div id=\"modeDlg\" class=\"overlay hidden\">\n  <div class=\"dialog\">\n    <h2>どちらの使い方に近いですか？</h2>\n    <p class=\"muted\">あとから「提出物」→「クラス設定」で変更できます。</p>\n    <button class=\"btn choice\" data-mode=\"single\"><b>1クラスで使う</b><span class=\"muted\">担任クラスなど、1クラスの提出物を管理します</span></button>\n    <button class=\"btn choice\" data-mode=\"multi\"><b>複数クラスで使う</b><span class=\"muted\">教科担任として、担当する複数クラスを行き来します</span></button>\n    <button class=\"btn choice\" data-mode=\"sample\"><b>サンプルで試してみる</b><span class=\"muted\">架空の3クラスと英語の提出物で画面を確認します（あとで初期化できます）</span></button>\n    <p class=\"muted\" id=\"storageNote\" style=\"margin:14px 0 0\"></p>\n  </div>\n</div>\n<div class=\"toast\" id=\"toast\"></div>\n\n<script>\n// ===== 動作環境 =====\n// Apps Script（スプレッドシート）上ではシートに保存し、それ以外ではこの端末のブラウザに保存する。\nconst IS_GAS = typeof google !== 'undefined' && !!(google.script && google.script.run);\nconst LS_KEY = 'teishutsu-checker-v1';\nconst MAX_CLASSES = 10, MAX_STUDENTS = 45;\nconst SUGGEST = ['ワークブック', '単語テスト直し', '単語ノート', '英作文', '予習ノート', '文法プリント', '音読カード',\n  'リスニング課題', 'スピーキング課題（録音）', '多読記録', '模試の解き直し', '英検過去問', '長期休業課題', '振り返りシート'];\nconst ST_LABEL = { collect: '回収中', few: 'あと5人以下', over: '締切超過', done: '全員提出' };\n\nlet state = emptyState();\nlet ui = {\n  view: 'status', cls: null,\n  sort: 'due', filterCls: '', hideDone: false,\n  chkHideDone: false, onlyMissing: false, quickTask: null, lastMsg: '', focusQuick: false,\n  callCls: '', overdueOnly: false,\n  printType: 'class', printCls: null, printStu: 'all', printNames: false, blankCols: 10, blankTitle: '',\n  editing: null, classDraft: null, openRoster: null,\n};\nlet undoStack = [];\nlet copyTexts = [];\n\nfunction emptyState() { return { mode: '', classes: [], tasks: [], checks: {} }; }\n\n// ===== 共通 =====\nconst $ = (s) => document.querySelector(s);\nconst esc = (s) => String(s ?? '').replace(/[&<>\"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c]));\nfunction toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('show'), 2000); }\nfunction fmtDate(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }\nfunction todayStr() { return fmtDate(new Date()); }\nfunction addDays(n) { const d = new Date(); d.setDate(d.getDate() + n); return fmtDate(d); }\nfunction dayDiff(a, b) { const p = (s) => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d); }; return Math.round((p(b) - p(a)) / 86400000); }\nfunction md(s) { if (!s) return ''; const [, m, d] = s.split('-'); return `${+m}/${+d}`; }\nfunction newId() { return 't' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }\n\nfunction getCls(name) { return state.classes.find((c) => c.name === name); }\nfunction activeClasses() { return state.mode === 'single' ? state.classes.slice(0, 1) : state.classes; }\nfunction tasksOf(clsName) { return state.tasks.filter((t) => t.classes.includes(clsName)).sort((a, b) => a.created - b.created); }\nfunction isChecked(id, cls, no) { return !!(state.checks[id] && state.checks[id][cls] && state.checks[id][cls][no]); }\nfunction missingNos(t, c) { const r = []; for (let n = 1; n <= c.count; n++) if (!isChecked(t.id, c.name, n)) r.push(n); return r; }\nfunction nosText(arr, max) {\n  if (!arr.length) return 'なし';\n  if (max && arr.length > max) return arr.slice(0, max).join('・') + `番 ほか${arr.length - max}人`;\n  return arr.join('・') + '番';\n}\nfunction stuName(c, no) { return (c.names || [])[no - 1] || ''; }\nfunction stuLabel(c, no) { const nm = stuName(c, no); return nm ? `${no}番 ${nm}` : `${no}番`; }\nfunction isOverdue(t) { return !!t.due && t.due < todayStr(); }\nfunction dueText(t) {\n  if (!t.due) return '締切なし';\n  const d = dayDiff(todayStr(), t.due);\n  if (d > 0) return `締切 ${md(t.due)}（あと${d}日）`;\n  if (d === 0) return `今日締切`;\n  return `締切 ${md(t.due)}（${-d}日超過）`;\n}\nfunction toldText(t) {\n  if (!t.told) return '';\n  const d = dayDiff(t.told, todayStr());\n  return d >= 0 ? `伝えて${d}日` : `${md(t.told)}に伝える予定`;\n}\nfunction cardInfo(t, c) {\n  const miss = missingNos(t, c);\n  let st = 'collect';\n  if (!miss.length) st = 'done'; else if (isOverdue(t)) st = 'over'; else if (miss.length <= 5) st = 'few';\n  return { t, c, miss, done: c.count - miss.length, st };\n}\nfunction classOptions(selected, withAll) {\n  return (withAll ? `<option value=\"\">すべてのクラス</option>` : '') +\n    activeClasses().map((c) => `<option value=\"${esc(c.name)}\"${c.name === selected ? ' selected' : ''}>${esc(c.name)}</option>`).join('');\n}\nfunction ensureCls() {\n  const cs = activeClasses();\n  if (!cs.find((c) => c.name === ui.cls)) ui.cls = cs[0] ? cs[0].name : null;\n  return getCls(ui.cls);\n}\nasync function copyText(text) {\n  try { await navigator.clipboard.writeText(text); toast('コピーしました'); return; } catch (e) {}\n  const ta = document.createElement('textarea'); ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';\n  document.body.appendChild(ta); ta.select();\n  try { document.execCommand('copy'); toast('コピーしました'); } catch (e) { prompt('コピーしてください', text); }\n  ta.remove();\n}\n\n// ===== 保存（ブラウザ or スプレッドシート） =====\nfunction gas(fn, ...args) {\n  return new Promise((res, rej) => google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...args));\n}\nfunction loadLocal() {\n  try { const d = JSON.parse(localStorage.getItem(LS_KEY)); if (d && Array.isArray(d.classes)) return d; } catch (e) {}\n  return emptyState();\n}\nconst LS_BACKUP = 'teishutsu-checker-lastbackup';\nfunction lastBackupDays() {\n  try { const t = +localStorage.getItem(LS_BACKUP); return t ? Math.floor((Date.now() - t) / 86400000) : null; } catch (e) { return null; }\n}\nfunction backupBanner() {\n  if (IS_GAS || !state.tasks.length) return '';\n  const d = lastBackupDays();\n  if (d !== null && d < 14) return '';\n  return `<div class=\"card noprint\" style=\"border-color:var(--c-few);background:var(--c-few-bg)\">\n    <div class=\"row\" style=\"justify-content:space-between\"><span style=\"font-size:13px\">データはこのブラウザの中だけに保存されています。${d === null ? 'まだバックアップがありません。' : `最後のバックアップから${d}日たちました。`}端末の故障やブラウザのデータ消去に備えて、「提出物」画面の下にある「バックアップを保存」をおすすめします。</span>\n    <button class=\"btn small\" data-go=\"tasks\">バックアップへ</button></div></div>`;\n}\nfunction saveLocal() {\n  try { localStorage.setItem(LS_KEY, JSON.stringify(state)); } catch (e) { toast('この端末に保存できませんでした'); }\n}\nfunction normalize(s) {\n  s = Object.assign(emptyState(), s || {});\n  s.classes = (s.classes || []).map((c) => {\n    const count = Math.max(1, Math.min(MAX_STUDENTS, parseInt(c.count, 10) || 1));\n    const names = Array.from({ length: count }, (_, i) => String((c.names || [])[i] || '').trim());\n    return { name: String(c.name).trim(), count, names };\n  });\n  s.tasks = (s.tasks || []).map((t) => ({ id: t.id, name: t.name, classes: t.classes || [], told: t.told || '', due: t.due || '', created: Number(t.created) || 0 }));\n  s.checks = s.checks || {};\n  return s;\n}\nlet pending = 0;\nfunction showSync() {\n  if (!IS_GAS) return;\n  const el = $('#syncState');\n  el.classList.remove('hidden');\n  el.textContent = pending ? '保存中…' : '✓ 保存済み';\n}\nasync function persist(gasFn, ...args) {\n  if (!IS_GAS) { saveLocal(); return; }\n  pending++; showSync();\n  try { await gas(gasFn, ...args); }\n  catch (e) { toast('保存に失敗しました：' + (e && e.message || e)); pending--; await reload(); return; }\n  pending--; showSync();\n}\nasync function reload(quiet) {\n  if (IS_GAS && quiet) { $('#syncState').textContent = '更新中…'; }\n  try { state = normalize(IS_GAS ? await gas('getData') : loadLocal()); }\n  catch (e) {\n    if (quiet) { toast('最新の内容を読み込めませんでした'); showSync(); return; }\n    $('#loading').textContent = '読み込みに失敗しました：' + (e && e.message || e); $('#loading').classList.remove('hidden'); return;\n  }\n  // 入力途中の画面はそのままにする\n  if (quiet && ui.view === 'tasks' && (ui.editing || ui.openRoster !== null)) { showSync(); return; }\n  if (quiet) ui.classDraft = null;\n  render(); showSync();\n}\n// ほかの端末で入力した内容を、画面に戻ってきたときに読み込む\nif (IS_GAS) {\n  $('#btnRefresh').classList.remove('hidden');\n  $('#btnRefresh').onclick = () => { if (!pending) reload(true); };\n  let hiddenAt = 0;\n  document.addEventListener('visibilitychange', () => {\n    if (document.hidden) { hiddenAt = Date.now(); return; }\n    if (hiddenAt && Date.now() - hiddenAt > 30000 && !pending) reload(true);\n  });\n}\n\nfunction setCheck(cls, id, no, on, recordUndo = true) {\n  const prev = isChecked(id, cls, no);\n  if (prev === on) return;\n  if (recordUndo) undoStack.push({ cls, id, no, prev });\n  if (undoStack.length > 200) undoStack.shift();\n  state.checks[id] = state.checks[id] || {};\n  state.checks[id][cls] = state.checks[id][cls] || {};\n  if (on) state.checks[id][cls][no] = true; else delete state.checks[id][cls][no];\n  persist('setCheck', cls, id, no, on);\n}\nfunction upsertTask(t) {\n  const i = state.tasks.findIndex((x) => x.id === t.id);\n  if (i >= 0) state.tasks[i] = t; else state.tasks.push(t);\n  persist('saveTask', t);\n}\nfunction removeTask(id) {\n  state.tasks = state.tasks.filter((t) => t.id !== id);\n  delete state.checks[id];\n  undoStack = undoStack.filter((u) => u.id !== id);\n  persist('deleteTask', id);\n}\nasync function saveClasses(list) {\n  if (IS_GAS) {\n    try { state = normalize(await gas('saveClasses', list)); }\n    catch (e) { toast('保存に失敗しました：' + (e && e.message || e)); return false; }\n  } else {\n    const keep = new Set(list.map((c) => c.oldName).filter(Boolean));\n    const removed = state.classes.map((c) => c.name).filter((n) => !keep.has(n));\n    for (const t of state.tasks) {\n      t.classes = t.classes.filter((n) => !removed.includes(n));\n      for (const c of list) if (c.oldName && c.oldName !== c.name) t.classes = t.classes.map((n) => n === c.oldName ? c.name : n);\n    }\n    for (const id in state.checks) {\n      const old = state.checks[id], next = {};\n      for (const c of list) if (c.oldName && old[c.oldName]) next[c.name] = old[c.oldName];\n      state.checks[id] = next;\n    }\n    state.classes = list.map((c) => ({ name: c.name, count: c.count, names: c.names }));\n    state = normalize(state); saveLocal();\n  }\n  undoStack = [];\n  return true;\n}\nasync function importState(s) {\n  s = normalize(s);\n  if (IS_GAS) {\n    try { state = normalize(await gas('importState', s)); }\n    catch (e) { toast('失敗しました：' + (e && e.message || e)); return; }\n  } else { state = s; saveLocal(); }\n  undoStack = []; ui.classDraft = null; ui.cls = null; ui.lastMsg = '';\n}\nfunction setMode(mode) {\n  state.mode = mode; ui.classDraft = null;\n  persist('setMode', mode);\n}\n\n// ===== サンプル =====\nfunction buildSample() {\n  let seed = 7;\n  const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;\n  const classes = [{ name: '2年1組', count: 38 }, { name: '2年2組', count: 40 }, { name: '2年3組', count: 39 }];\n  const all = classes.map((c) => c.name);\n  const base = Date.now() - 100000;\n  const tasks = [\n    { name: 'Lesson 4 ワークブック', classes: all, told: addDays(-7), due: todayStr(), rate: 0.86 },\n    { name: '英作文（My Future Dream）', classes: ['2年1組', '2年2組'], told: addDays(-10), due: addDays(-2), rate: 0.93 },\n    { name: '単語テスト直し', classes: all, told: addDays(-2), due: addDays(3), rate: 0.55 },\n    { name: '音読カード（10月）', classes: ['2年3組'], told: addDays(-14), due: addDays(-1), rate: 1 },\n  ].map((t, i) => Object.assign(t, { id: 'sample' + i, created: base + i }));\n  const checks = {};\n  for (const t of tasks) {\n    checks[t.id] = {};\n    for (const cn of t.classes) {\n      const c = classes.find((x) => x.name === cn); checks[t.id][cn] = {};\n      for (let n = 1; n <= c.count; n++) if (rnd() < t.rate) checks[t.id][cn][n] = true;\n    }\n    delete t.rate;\n  }\n  return { mode: 'multi', classes, tasks, checks };\n}\n\n// ===== 画面切り替え =====\ndocument.querySelectorAll('#nav button').forEach((b) => b.onclick = () => { ui.view = b.dataset.view; render(); window.scrollTo(0, 0); });\nfunction render() {\n  $('#loading').classList.add('hidden');\n  document.querySelectorAll('#nav button').forEach((b) => b.classList.toggle('active', b.dataset.view === ui.view));\n  for (const v of ['status', 'check', 'call', 'print', 'tasks', 'help']) $('#view-' + v).classList.toggle('hidden', v !== ui.view);\n  ({ status: renderStatus, check: renderCheck, call: renderCall, print: renderPrint, tasks: renderTasks, help: renderHelp })[ui.view]();\n  $('#modeDlg').classList.toggle('hidden', !!state.mode);\n}\nfunction needSetup() {\n  if (!activeClasses().length) return `<div class=\"card empty\">まだクラスが登録されていません。<br><br><button class=\"btn primary\" data-go=\"tasks\">「提出物」画面でクラスを登録する</button></div>`;\n  if (!state.tasks.length) return `<div class=\"card empty\">まだ提出物が登録されていません。まずは今日集める提出物を一つ登録してみましょう。<br><br><button class=\"btn primary\" data-go=\"tasks\" data-add=\"1\">＋ 提出物を追加</button></div>`;\n  return '';\n}\ndocument.addEventListener('click', (e) => {\n  const g = e.target.closest('[data-go]'); if (!g) return;\n  ui.view = g.dataset.go;\n  if (g.dataset.add) ui.editing = newTaskDraft();\n  render();\n});\n\n// ===== 回収状況 =====\nfunction cardHtml(i) {\n  const pct = Math.round(i.done / i.c.count * 100);\n  return `<button class=\"tcard st-${i.st}\" data-cls=\"${esc(i.c.name)}\" data-id=\"${i.t.id}\">\n    <div class=\"top\"><span class=\"badge b-${i.st}\">${ST_LABEL[i.st]}</span>${state.mode === 'multi' ? `<span>${esc(i.c.name)}</span>` : ''}</div>\n    <div class=\"name\">${esc(i.t.name)}</div>\n    <div class=\"count\"><b>${i.done}</b> / ${i.c.count}人 提出（${pct}%）</div>\n    <div class=\"bar\" style=\"color:var(--c-${i.st})\"><div style=\"width:${pct}%\"></div></div>\n    <div class=\"miss\">${i.miss.length ? '未提出：' + nosText(i.miss, 15) : '全員提出しました'}</div>\n    <div class=\"meta\">${dueText(i.t)}${toldText(i.t) ? '・' + toldText(i.t) : ''}</div>\n  </button>`;\n}\nfunction renderStatus() {\n  const el = $('#view-status');\n  const ns = needSetup(); if (ns) { el.innerHTML = ns; return; }\n  const today = todayStr();\n  const classes = activeClasses();\n  let items = [];\n  for (const c of classes) for (const t of tasksOf(c.name)) items.push(cardInfo(t, c));\n  const counts = { collect: 0, few: 0, over: 0, done: 0 };\n  items.forEach((i) => counts[i.st]++);\n  const todays = items.filter((i) => i.t.due === today).sort((a, b) => b.miss.length - a.miss.length);\n  let list = items.filter((i) => (!ui.filterCls || i.c.name === ui.filterCls) && (!ui.hideDone || i.st !== 'done'));\n  const ci = (i) => classes.indexOf(i.c);\n  const sorters = {\n    class: (a, b) => ci(a) - ci(b) || a.t.created - b.t.created,\n    task: (a, b) => a.t.created - b.t.created || ci(a) - ci(b),\n    due: (a, b) => (a.t.due || '9999').localeCompare(b.t.due || '9999') || b.miss.length - a.miss.length,\n    missing: (a, b) => b.miss.length - a.miss.length || ci(a) - ci(b),\n  };\n  list.sort(sorters[ui.sort]);\n  el.innerHTML = backupBanner() + `\n    <div class=\"card\">\n      <h2>今日が締切の提出物（${md(today)}）</h2>\n      ${todays.length ? `<div class=\"cards\">${todays.map(cardHtml).join('')}</div>` : '<p class=\"muted\">今日が締切の提出物はありません。</p>'}\n    </div>\n    <div class=\"card\">\n      <div class=\"row\" style=\"justify-content:space-between;margin-bottom:10px\">\n        <h2 style=\"margin:0\">すべての提出物</h2>\n        <div class=\"legend\">\n          <span class=\"badge b-collect\">回収中 ${counts.collect}</span><span class=\"badge b-few\">あと5人以下 ${counts.few}</span>\n          <span class=\"badge b-over\">締切超過 ${counts.over}</span><span class=\"badge b-done\">全員提出 ${counts.done}</span>\n        </div>\n      </div>\n      <div class=\"row\" style=\"margin-bottom:10px\">\n        <label class=\"f\">並べ替え<select id=\"stSort\">\n          <option value=\"due\">締切が近い順</option><option value=\"missing\">未提出が多い順</option>\n          ${state.mode === 'multi' ? '<option value=\"class\">クラスごと</option>' : ''}<option value=\"task\">提出物ごと</option>\n        </select></label>\n        ${state.mode === 'multi' ? `<label class=\"f\">クラス<select id=\"stCls\">${classOptions(ui.filterCls, true)}</select></label>` : ''}\n        <label class=\"chk\" style=\"margin-top:14px\"><input type=\"checkbox\" id=\"stHide\"${ui.hideDone ? ' checked' : ''}> 全員提出を隠す</label>\n      </div>\n      ${list.length ? `<div class=\"cards\">${list.map(cardHtml).join('')}</div>` : '<p class=\"muted\">表示する提出物はありません。</p>'}\n      <p class=\"muted\" style=\"margin-bottom:0\">カードを押すと、そのクラスのチェック入力を開きます。</p>\n    </div>`;\n  $('#stSort').value = ui.sort;\n  $('#stSort').onchange = (e) => { ui.sort = e.target.value; render(); };\n  if ($('#stCls')) $('#stCls').onchange = (e) => { ui.filterCls = e.target.value; render(); };\n  $('#stHide').onchange = (e) => { ui.hideDone = e.target.checked; render(); };\n  el.querySelectorAll('.tcard').forEach((b) => b.onclick = () => {\n    ui.cls = b.dataset.cls; ui.quickTask = b.dataset.id; ui.view = 'check'; render();\n  });\n}\n\n// ===== チェック入力 =====\nfunction renderCheck() {\n  const el = $('#view-check');\n  const ns = needSetup(); if (ns) { el.innerHTML = ns; return; }\n  const c = ensureCls();\n  const allTasks = tasksOf(c.name);\n  if (!allTasks.find((t) => t.id === ui.quickTask)) ui.quickTask = allTasks[0] ? allTasks[0].id : null;\n  let ts = ui.chkHideDone ? allTasks.filter((t) => missingNos(t, c).length) : allTasks;\n  const hasNames = c.names.some(Boolean);\n  let rows = '';\n  for (let n = 1; n <= c.count; n++) {\n    if (ui.onlyMissing && ts.every((t) => isChecked(t.id, c.name, n))) continue;\n    rows += `<tr><td class=\"sticky\"><b>${n}</b>${hasNames ? ` <span class=\"muted\">${esc(stuName(c, n))}</span>` : ''}</td>` +\n      ts.map((t) => {\n        const on = isChecked(t.id, c.name, n);\n        return `<td><button class=\"cell${on ? ' on' : ''}${!on && isOverdue(t) ? ' late' : ''}\" data-id=\"${t.id}\" data-no=\"${n}\" aria-label=\"${n}番 ${esc(t.name)}\">${on ? '✓' : ''}</button></td>`;\n      }).join('') + '</tr>';\n  }\n  el.innerHTML = `\n    <div class=\"card\">\n      <div class=\"row\">\n        ${state.mode === 'multi' ? `<label class=\"f\">クラス<select id=\"ckCls\">${classOptions(c.name)}</select></label>` : `<b>${esc(c.name)}</b>`}\n        <label class=\"chk\"><input type=\"checkbox\" id=\"ckHide\"${ui.chkHideDone ? ' checked' : ''}> 完了した提出物を隠す</label>\n        <label class=\"chk\"><input type=\"checkbox\" id=\"ckOnly\"${ui.onlyMissing ? ' checked' : ''}> 未提出のある番号だけ</label>\n        <button class=\"btn\" id=\"ckUndo\" style=\"margin-left:auto\"${undoStack.length ? '' : ' disabled'}>↶ 取り消し</button>\n      </div>\n    </div>\n    ${allTasks.length ? `\n    <div class=\"card quick\">\n      <div class=\"row\">\n        <b style=\"font-size:14px\">番号で入力</b>\n        <select id=\"qTask\">${allTasks.map((t) => `<option value=\"${t.id}\"${t.id === ui.quickTask ? ' selected' : ''}>${esc(t.name)}</option>`).join('')}</select>\n        <input id=\"qNo\" inputmode=\"numeric\" autocomplete=\"off\" placeholder=\"番号\">\n        <span class=\"muted\">番号を入れて Enter。続けて次の番号を入力できます。</span>\n      </div>\n      <div class=\"lastmsg\" id=\"qMsg\">${ui.lastMsg}</div>\n    </div>\n    <div class=\"tablewrap\">\n      <table class=\"grid\">\n        <thead><tr><th class=\"sticky\">番号${hasNames ? '・氏名' : ''}</th>${ts.map((t) => {\n          const m = missingNos(t, c).length;\n          return `<th><div class=\"th-name\">${esc(t.name)}</div><div class=\"th-sub\">${c.count - m}/${c.count}・${t.due ? '〆' + md(t.due) : '締切なし'}</div></th>`;\n        }).join('')}</tr></thead>\n        <tbody>${rows || `<tr><td class=\"sticky\" colspan=\"${ts.length + 1}\" style=\"padding:16px\">全員提出しています。</td></tr>`}</tbody>\n      </table>\n    </div>\n    <p class=\"muted\">マスを押すと ✓（提出）になり、もう一度押すと外れます。赤いマスは締切を過ぎた未提出です。</p>`\n    : `<div class=\"card empty\">このクラスが対象の提出物はまだありません。<br><br><button class=\"btn primary\" data-go=\"tasks\" data-add=\"1\">＋ 提出物を追加</button></div>`}`;\n\n  if ($('#ckCls')) $('#ckCls').onchange = (e) => { ui.cls = e.target.value; ui.lastMsg = ''; render(); };\n  $('#ckHide').onchange = (e) => { ui.chkHideDone = e.target.checked; render(); };\n  $('#ckOnly').onchange = (e) => { ui.onlyMissing = e.target.checked; render(); };\n  $('#ckUndo').onclick = () => {\n    const u = undoStack.pop(); if (!u) return;\n    setCheck(u.cls, u.id, u.no, u.prev, false);\n    const t = state.tasks.find((x) => x.id === u.id);\n    ui.lastMsg = `取り消しました：${u.cls} ${u.no}番「${t ? esc(t.name) : ''}」`;\n    if (u.cls !== ui.cls && getCls(u.cls)) ui.cls = u.cls;\n    render();\n  };\n  const tbody = el.querySelector('table.grid');\n  if (tbody) tbody.addEventListener('click', (e) => {\n    const b = e.target.closest('.cell'); if (!b) return;\n    const no = +b.dataset.no, id = b.dataset.id;\n    const sx = $('.tablewrap').scrollLeft, sy = $('.tablewrap').scrollTop;\n    setCheck(c.name, id, no, !isChecked(id, c.name, no));\n    render();\n    $('.tablewrap').scrollLeft = sx; $('.tablewrap').scrollTop = sy;\n  });\n  if ($('#qTask')) {\n    $('#qTask').onchange = (e) => { ui.quickTask = e.target.value; ui.lastMsg = ''; render(); $('#qNo').focus(); };\n    $('#qNo').addEventListener('keydown', (e) => {\n      if (e.key !== 'Enter' || e.isComposing) return;\n      e.preventDefault();\n      const raw = e.target.value.replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xFEE0)).trim();\n      const no = parseInt(raw, 10);\n      const t = state.tasks.find((x) => x.id === ui.quickTask);\n      if (!t) return;\n      if (!(no >= 1 && no <= c.count)) { $('#qMsg').innerHTML = `<span style=\"color:var(--c-over)\">1〜${c.count}の番号を入力してください</span>`; e.target.select(); return; }\n      if (isChecked(t.id, c.name, no)) {\n        ui.lastMsg = `<span style=\"color:var(--c-few)\">${no}番はすでにチェック済みです</span>`;\n      } else {\n        setCheck(c.name, t.id, no, true);\n        const left = missingNos(t, c);\n        ui.lastMsg = `✓ ${esc(stuLabel(c, no))} をチェックしました。「${esc(t.name)}」残り${left.length}人${left.length && left.length <= 10 ? '（' + nosText(left) + '）' : ''}`;\n      }\n      ui.focusQuick = true; render();\n    });\n    if (ui.focusQuick) { ui.focusQuick = false; $('#qNo').focus(); }\n  }\n}\n\n// ===== 声かけビュー =====\nfunction renderCall() {\n  const el = $('#view-call');\n  const ns = needSetup(); if (ns) { el.innerHTML = ns; return; }\n  if (ui.callCls && !getCls(ui.callCls)) ui.callCls = '';\n  const classes = ui.callCls ? [getCls(ui.callCls)] : activeClasses();\n  copyTexts = [];\n  const addCopy = (txt) => { copyTexts.push(txt); return copyTexts.length - 1; };\n  let body = '', watch = [];\n  for (const c of classes) {\n    const ts = tasksOf(c.name).filter((t) => !ui.overdueOnly || isOverdue(t));\n    const lines = [];\n    let rows = '';\n    for (let n = 1; n <= c.count; n++) {\n      const miss = ts.filter((t) => !isChecked(t.id, c.name, n));\n      if (!miss.length) continue;\n      const names = miss.map((t) => t.name);\n      const line = `${stuLabel(c, n)}：${names.join('、')}`;\n      lines.push(line);\n      if (ts.length >= 2 && miss.length >= 2) watch.push({ c, n, miss });\n      rows += `<div class=\"callrow\"><div class=\"who\">${esc(stuLabel(c, n))}</div>\n        <div class=\"what\">${miss.map((t) => `${esc(t.name)}${isOverdue(t) ? ' <span class=\"badge b-over\">超過</span>' : ''}`).join('、')}</div>\n        <span class=\"n\">${miss.length}件</span>\n        <button class=\"btn small\" data-copy=\"${addCopy(line)}\">コピー</button></div>`;\n    }\n    const all = `【${c.name}】まだ確認できていない提出物\\n` + (lines.join('\\n') || 'なし');\n    body += `<div class=\"card\">\n      <div class=\"row\" style=\"justify-content:space-between\"><h2 style=\"margin:0\">${esc(c.name)}　<span class=\"muted\">${lines.length}人</span></h2>\n      ${lines.length ? `<button class=\"btn\" data-copy=\"${addCopy(all)}\">このクラスをまとめてコピー</button>` : ''}</div>\n      ${ts.length ? (rows || '<p class=\"muted\">未提出はありません。</p>') : '<p class=\"muted\">対象の提出物がありません。</p>'}\n    </div>`;\n  }\n  watch.sort((a, b) => b.miss.length - a.miss.length);\n  el.innerHTML = `\n    <div class=\"card\">\n      <div class=\"row\">\n        ${state.mode === 'multi' ? `<label class=\"f\">クラス<select id=\"cvCls\">${classOptions(ui.callCls, true)}</select></label>` : ''}\n        <label class=\"chk\"${state.mode === 'multi' ? ' style=\"margin-top:14px\"' : ''}><input type=\"checkbox\" id=\"cvOver\"${ui.overdueOnly ? ' checked' : ''}> 締切を過ぎた提出物だけ</label>\n      </div>\n      <p class=\"muted\" style=\"margin:8px 0 0\">未提出を生徒ごとにまとめています。「3番：ワークブック、英作文」のように、一度の声かけで伝えられます。</p>\n    </div>\n    ${watch.length ? `<div class=\"card\"><h2>要注意リスト（未提出が2件以上重なっている番号）</h2>\n      <div class=\"note\">理由を決めつけるためのリストではありません。忘れているのかもしれないし、何か事情があるのかもしれません。見落とさないための目印として使ってください。</div>\n      ${watch.map((w) => `<div class=\"callrow\"><div class=\"who\">${state.mode === 'multi' && !ui.callCls ? esc(w.c.name) + ' ' : ''}${esc(stuLabel(w.c, w.n))}</div>\n        <div class=\"what\">${w.miss.map((t) => esc(t.name)).join('、')}</div><span class=\"n\">${w.miss.length}件</span></div>`).join('')}\n    </div>` : ''}\n    ${body}`;\n  if ($('#cvCls')) $('#cvCls').onchange = (e) => { ui.callCls = e.target.value; render(); };\n  $('#cvOver').onchange = (e) => { ui.overdueOnly = e.target.checked; render(); };\n  el.querySelectorAll('[data-copy]').forEach((b) => b.onclick = () => copyText(copyTexts[+b.dataset.copy]));\n}\n\n// ===== 印刷プレビュー =====\nfunction renderPrint() {\n  const el = $('#view-print');\n  if (!activeClasses().length) { el.innerHTML = needSetup(); return; }\n  if (!getCls(ui.printCls) || !activeClasses().includes(getCls(ui.printCls))) ui.printCls = activeClasses()[0].name;\n  const c = getCls(ui.printCls);\n  const ts = tasksOf(c.name);\n  const hasNames = c.names.some(Boolean);\n  const nowTxt = `${md(todayStr())}現在`;\n  let pages = '';\n  if (ui.printType === 'class') {\n    if (!ts.length) pages = '<div class=\"page\"><p>このクラスが対象の提出物はありません。</p></div>';\n    // 1ページあたり最大12列。多いときはページを分ける\n    for (let s = 0; s < ts.length; s += 12) {\n      const part = ts.slice(s, s + 12);\n      const showName = hasNames && ui.printNames;\n      pages += `<div class=\"page\"><h2>${esc(c.name)} 提出物の状況</h2><div class=\"sub\">${nowTxt}　✓＝提出済み／空欄＝まだ確認できていません</div>\n        <table><thead><tr><th style=\"width:${showName ? 26 : 34}px\">番号</th>${showName ? '<th style=\"width:90px\">氏名</th>' : ''}\n        ${part.map((t) => `<th>${esc(t.name)}<br>${t.due ? '〆' + md(t.due) : ''}</th>`).join('')}</tr></thead><tbody>\n        ${Array.from({ length: c.count }, (_, i) => i + 1).map((n) => `<tr><td>${n}</td>${showName ? `<td style=\"text-align:left\">${esc(stuName(c, n))}</td>` : ''}\n          ${part.map((t) => isChecked(t.id, c.name, n) ? '<td>✓</td>' : '<td class=\"miss\"></td>').join('')}</tr>`).join('')}\n        <tr><td colspan=\"${showName ? 2 : 1}\">提出</td>${part.map((t) => `<td>${c.count - missingNos(t, c).length}/${c.count}</td>`).join('')}</tr>\n        </tbody></table></div>`;\n    }\n  } else if (ui.printType === 'personal') {\n    const targets = [];\n    for (let n = 1; n <= c.count; n++) {\n      if (ui.printStu !== 'all' && +ui.printStu !== n) continue;\n      const miss = ts.filter((t) => !isChecked(t.id, c.name, n));\n      if (miss.length || ui.printStu !== 'all') targets.push({ n, miss });\n    }\n    const slips = targets.map(({ n, miss }) => `<div class=\"slip\"><h3>${esc(c.name)}　${esc(stuLabel(c, n))}さんへ</h3>\n      <div class=\"foot\">${nowTxt}、まだ確認できていない提出物です。</div>\n      ${miss.length ? `<ul>${miss.map((t) => `<li>${esc(t.name)}${t.due ? `（締切 ${md(t.due)}）` : ''}</li>`).join('')}</ul>` : '<p>すべて提出済みです。ありがとう！</p>'}\n      <div class=\"foot\">提出したら先生に声をかけてください。行き違いで提出済みの場合も教えてください。提出できない場合は放課後残ってやっていってください。</div></div>`);\n    if (!slips.length) pages = '<div class=\"page\"><p>未提出のある生徒はいません。</p></div>';\n    for (let i = 0; i < slips.length; i += 4) pages += `<div class=\"page\">${slips.slice(i, i + 4).join('')}</div>`;\n  } else {\n    const cols = Math.max(1, Math.min(16, +ui.blankCols || 10));\n    pages = `<div class=\"page\"><h2>${esc(ui.blankTitle || c.name + ' 提出物チェック表')}</h2><div class=\"sub\">${esc(c.name)}　　月　　日〜</div>\n      <table><thead><tr><th style=\"width:34px\">番号</th>${hasNames && ui.printNames ? '<th style=\"width:90px\">氏名</th>' : ''}${'<th style=\"height:52px\"></th>'.repeat(cols)}</tr></thead><tbody>\n      ${Array.from({ length: c.count }, (_, i) => `<tr><td>${i + 1}</td>${hasNames && ui.printNames ? `<td style=\"text-align:left\">${esc(stuName(c, i + 1))}</td>` : ''}${'<td></td>'.repeat(cols)}</tr>`).join('')}\n      </tbody></table></div>`;\n  }\n  el.innerHTML = `\n    <div class=\"card noprint\">\n      <div class=\"row\" style=\"align-items:flex-end\">\n        <label class=\"f\">内容<select id=\"pType\">\n          <option value=\"class\">クラス別（掲示用）</option><option value=\"personal\">個人別（本人に渡す）</option><option value=\"blank\">白紙のチェック表</option>\n        </select></label>\n        ${state.mode === 'multi' ? `<label class=\"f\">クラス<select id=\"pCls\">${classOptions(c.name)}</select></label>` : ''}\n        ${ui.printType === 'personal' ? `<label class=\"f\">生徒<select id=\"pStu\"><option value=\"all\">未提出がある全員</option>\n          ${Array.from({ length: c.count }, (_, i) => `<option value=\"${i + 1}\"${String(i + 1) === String(ui.printStu) ? ' selected' : ''}>${esc(stuLabel(c, i + 1))}</option>`).join('')}</select></label>` : ''}\n        ${ui.printType === 'blank' ? `<label class=\"f\">列の数<input id=\"pCols\" type=\"number\" min=\"1\" max=\"16\" value=\"${ui.blankCols}\" style=\"width:70px\"></label>\n          <label class=\"f\">表題<input id=\"pTitle\" value=\"${esc(ui.blankTitle)}\" placeholder=\"${esc(c.name)} 提出物チェック表\"></label>` : ''}\n        ${ui.printType !== 'personal' && hasNames ? `<label class=\"chk\"><input type=\"checkbox\" id=\"pNames\"${ui.printNames ? ' checked' : ''}> 氏名を表示</label>` : ''}\n        <button class=\"btn primary\" id=\"pPrint\" style=\"margin-left:auto\">印刷</button>\n      </div>\n      <p class=\"muted\" style=\"margin:8px 0 0\">A4で印刷できます。PDFにするときは、印刷画面の送信先で「PDFに保存」を選んでください。${ui.printType === 'class' && hasNames ? '掲示する場合は、氏名を表示しない設定をおすすめします。' : ''}${ui.printType === 'personal' ? '1枚に4人分です。点線で切って渡せます。' : ''}</p>\n    </div>\n    <div class=\"pages\">${pages}</div>`;\n  $('#pType').value = ui.printType;\n  $('#pType').onchange = (e) => { ui.printType = e.target.value; render(); };\n  if ($('#pCls')) $('#pCls').onchange = (e) => { ui.printCls = e.target.value; ui.printStu = 'all'; render(); };\n  if ($('#pStu')) $('#pStu').onchange = (e) => { ui.printStu = e.target.value; render(); };\n  if ($('#pCols')) $('#pCols').onchange = (e) => { ui.blankCols = e.target.value; render(); };\n  if ($('#pTitle')) $('#pTitle').onchange = (e) => { ui.blankTitle = e.target.value; render(); };\n  if ($('#pNames')) $('#pNames').onchange = (e) => { ui.printNames = e.target.checked; render(); };\n  $('#pPrint').onclick = () => window.print();\n}\n\n// ===== 提出物（＋クラス設定・データ） =====\nfunction newTaskDraft() {\n  const cs = activeClasses();\n  return { id: null, name: '', classes: cs.length === 1 ? [cs[0].name] : [], told: todayStr(), due: '' };\n}\nfunction renderTasks() {\n  const el = $('#view-tasks');\n  const classes = activeClasses();\n  const ed = ui.editing;\n  const names = [...new Set([...SUGGEST, ...state.tasks.map((t) => t.name)])];\n  const tasks = state.tasks.slice().sort((a, b) => b.created - a.created);\n  if (!ui.classDraft) ui.classDraft = state.classes.map((c) => ({ name: c.name, count: c.count, names: c.names.slice(), oldName: c.name }));\n  const draft = ui.classDraft;\n  const visibleDraft = state.mode === 'single' ? draft.slice(0, 1) : draft;\n\n  el.innerHTML = `\n    <div class=\"card\">\n      <div class=\"row\" style=\"justify-content:space-between\">\n        <h2 style=\"margin:0\">提出物</h2>\n        ${ed ? '' : `<button class=\"btn primary\" id=\"tAdd\"${classes.length ? '' : ' disabled'}>＋ 追加</button>`}\n      </div>\n      ${!classes.length ? '<p class=\"muted\">先に、下の「クラス設定」でクラス名と人数を登録してください。</p>' : ''}\n      ${ed ? `\n      <div class=\"card\" style=\"background:var(--bg);margin:12px 0 0\">\n        <div class=\"row\" style=\"align-items:flex-end\">\n          <label class=\"f\">提出物名（必須）<input id=\"eName\" list=\"suggest\" value=\"${esc(ed.name)}\" placeholder=\"例：Lesson 5 ワークブック\" style=\"min-width:240px\"></label>\n          <datalist id=\"suggest\">${names.map((n) => `<option value=\"${esc(n)}\">`).join('')}</datalist>\n          <label class=\"f\">伝えた日<input id=\"eTold\" type=\"date\" value=\"${ed.told}\"></label>\n          <label class=\"f\">締切<input id=\"eDue\" type=\"date\" value=\"${ed.due}\"></label>\n        </div>\n        ${classes.length > 1 ? `<div class=\"f\" style=\"margin-top:10px\">対象クラス（必須・複数選択できます）</div>\n        <div class=\"row\" style=\"margin-top:4px\">\n          ${classes.map((c) => `<label class=\"chk\"><input type=\"checkbox\" class=\"eCls\" value=\"${esc(c.name)}\"${ed.classes.includes(c.name) ? ' checked' : ''}> ${esc(c.name)}</label>`).join('')}\n          <button class=\"btn small\" id=\"eAll\">すべて選択</button>\n        </div>` : ''}\n        <div class=\"row\" style=\"margin-top:12px\">\n          <button class=\"btn primary\" id=\"eSave\">${ed.id ? '更新する' : '追加する'}</button>\n          <button class=\"btn\" id=\"eCancel\">キャンセル</button>\n        </div>\n      </div>` : ''}\n      ${tasks.length ? `<div style=\"overflow-x:auto;margin-top:12px\"><table class=\"list\"><thead><tr><th>提出物名</th>${state.mode === 'multi' ? '<th>対象クラス</th>' : ''}<th>伝えた日</th><th>締切</th><th>提出</th><th></th></tr></thead><tbody>\n        ${tasks.map((t) => {\n          const cs = t.classes.map(getCls).filter((c) => c && classes.includes(c));\n          const tot = cs.reduce((s, c) => s + c.count, 0), done = cs.reduce((s, c) => s + c.count - missingNos(t, c).length, 0);\n          return `<tr><td><b>${esc(t.name)}</b></td>${state.mode === 'multi' ? `<td>${t.classes.map(esc).join('、')}</td>` : ''}\n            <td>${t.told ? md(t.told) : '−'}</td><td>${t.due ? md(t.due) : '−'}</td><td>${done}/${tot}</td>\n            <td style=\"text-align:right;white-space:nowrap\"><button class=\"btn small\" data-edit=\"${t.id}\">編集</button>\n            <button class=\"btn small danger\" data-del=\"${t.id}\">削除</button></td></tr>`;\n        }).join('')}</tbody></table></div>` : '<p class=\"muted\">まだ提出物はありません。</p>'}\n    </div>\n\n    <div class=\"card\">\n      <h2>クラス設定</h2>\n      <div class=\"row\" style=\"margin-bottom:10px\">\n        <span class=\"muted\">使い方：</span>\n        <label class=\"chk\"><input type=\"radio\" name=\"mode\" value=\"single\"${state.mode === 'single' ? ' checked' : ''}> 1クラスで使う</label>\n        <label class=\"chk\"><input type=\"radio\" name=\"mode\" value=\"multi\"${state.mode === 'multi' ? ' checked' : ''}> 複数クラスで使う</label>\n      </div>\n      <table class=\"list\"><thead><tr><th>クラス名</th><th>人数</th><th>名前（任意）</th><th></th></tr></thead><tbody>\n        ${visibleDraft.map((c, i) => `<tr>\n          <td><input class=\"cName\" data-i=\"${i}\" value=\"${esc(c.name)}\" placeholder=\"例：2年3組\" style=\"width:130px\"></td>\n          <td><input class=\"cCount\" data-i=\"${i}\" type=\"number\" min=\"1\" max=\"${MAX_STUDENTS}\" value=\"${c.count}\" style=\"width:70px\"></td>\n          <td><button class=\"btn small\" data-roster=\"${i}\">${c.names.some(Boolean) ? `入力済み（${c.names.filter(Boolean).length}人）` : '入力しない'}</button></td>\n          <td style=\"text-align:right\"><button class=\"btn small danger\" data-cdel=\"${i}\">削除</button></td></tr>\n          ${ui.openRoster === i ? `<tr><td colspan=\"4\">\n            <div class=\"muted\">1行に1人、出席番号順に貼り付けてください（Excelの「番号・氏名」2列でも可）。空欄のままでも使えます。名前を入れるかどうかは学校のルールに合わせてください。</div>\n            <textarea class=\"cNames\" data-i=\"${i}\">${esc(c.names.join('\\n').replace(/\\n+$/, ''))}</textarea></td></tr>` : ''}`).join('')}\n      </tbody></table>\n      <div class=\"row\" style=\"margin-top:10px\">\n        ${state.mode === 'single' && visibleDraft.length >= 1 ? '' : `<button class=\"btn\" id=\"cAdd\"${draft.length >= MAX_CLASSES ? ' disabled' : ''}>＋ クラスを追加</button>`}\n        <button class=\"btn primary\" id=\"cSave\">クラス設定を保存</button>\n        <span class=\"muted\">最大${MAX_CLASSES}クラス、1クラス${MAX_STUDENTS}人まで。</span>\n      </div>\n      ${IS_GAS ? '<p class=\"muted\">保存すると「設定」シートとクラスごとのシートに反映されます。「設定」シートを直接編集した場合は、「↻ 最新に更新」を押すとクラスシートが作られ、画面に反映されます。</p>' : ''}\n    </div>\n\n    <div class=\"card\">\n      <h2>データ</h2>\n      <p class=\"muted\">${IS_GAS ? 'データはあなたのGoogleドライブのスプレッドシートに保存されます。同じGoogleアカウントなら、どの端末からでも同じデータを使えます。外部へ送信することはありません。'\n        : 'データはこの端末のブラウザの中だけに保存されます。外部へ送信することはありません。端末を替えるときやブラウザのデータを消す前に、バックアップを保存してください。'}</p>\n      ${!IS_GAS ? '<p class=\"muted\"><b>複数の端末で同じデータを使いたい場合</b>は、<a href=\"setup.html\" target=\"_blank\" rel=\"noopener\">Googleアカウントで同期する設定</a>（初回だけ・約5分）をしてください。このブラウザのデータは「バックアップを保存」→ 同期版で「バックアップから復元」で引き継げます。</p>' : ''}\n      ${IS_GAS && state.sheetUrl ? `<p><a href=\"${esc(state.sheetUrl)}\" target=\"_blank\" rel=\"noopener\">データのスプレッドシートを開く</a></p>` : ''}\n      <div class=\"row\">\n        <button class=\"btn\" id=\"dBackup\">バックアップを保存</button>\n        <label class=\"btn\">バックアップから復元<input type=\"file\" id=\"dRestore\" accept=\".json,application/json\" class=\"hidden\"></label>\n        <button class=\"btn\" id=\"dSample\">サンプルデータを入れる</button>\n        <button class=\"btn danger\" id=\"dReset\">サンプルデータを消して初期化</button>\n      </div>\n    </div>`;\n\n  // 提出物\n  if ($('#tAdd')) $('#tAdd').onclick = () => { ui.editing = newTaskDraft(); render(); $('#eName').focus(); };\n  if (ed) {\n    const sync = () => {\n      ed.name = $('#eName').value; ed.told = $('#eTold').value; ed.due = $('#eDue').value;\n      if (classes.length > 1) ed.classes = [...el.querySelectorAll('.eCls:checked')].map((x) => x.value);\n    };\n    if ($('#eAll')) $('#eAll').onclick = () => { sync(); ed.classes = classes.map((c) => c.name); render(); };\n    $('#eCancel').onclick = () => { ui.editing = null; render(); };\n    $('#eSave').onclick = () => {\n      sync();\n      const name = ed.name.trim();\n      if (!name) return toast('提出物名を入力してください');\n      if (!ed.classes.length) return toast('対象クラスを選んでください');\n      if (ed.told && ed.due && ed.due < ed.told) return toast('締切が伝えた日より前になっています');\n      const old = state.tasks.find((t) => t.id === ed.id);\n      // 1クラスモードで編集しても、他クラスへの割り当ては残す\n      const hidden = old ? old.classes.filter((n) => !classes.find((c) => c.name === n)) : [];\n      upsertTask({ id: ed.id || newId(), name, classes: [...hidden, ...ed.classes], told: ed.told, due: ed.due, created: old ? old.created : Date.now() });\n      toast(old ? '更新しました' : `「${name}」を追加しました`);\n      ui.editing = null; render();\n    };\n  }\n  el.querySelectorAll('[data-edit]').forEach((b) => b.onclick = () => {\n    const t = state.tasks.find((x) => x.id === b.dataset.edit);\n    ui.editing = { id: t.id, name: t.name, classes: t.classes.slice(), told: t.told, due: t.due };\n    render(); window.scrollTo(0, 0);\n  });\n  el.querySelectorAll('[data-del]').forEach((b) => b.onclick = () => {\n    const t = state.tasks.find((x) => x.id === b.dataset.del);\n    if (!confirm(`「${t.name}」と、そのチェック記録を削除します。よろしいですか？`)) return;\n    removeTask(t.id); render();\n  });\n\n  // クラス設定\n  const syncDraft = () => {\n    el.querySelectorAll('.cName').forEach((x) => draft[+x.dataset.i].name = x.value);\n    el.querySelectorAll('.cCount').forEach((x) => draft[+x.dataset.i].count = x.value);\n    el.querySelectorAll('.cNames').forEach((x) => {\n      draft[+x.dataset.i].names = x.value.split(/\\r?\\n/).map((l) => {\n        const p = l.split('\\t').map((s) => s.trim()).filter(Boolean);\n        return p.length >= 2 && /^\\d+$/.test(p[0]) ? p.slice(1).join(' ') : l.trim();\n      });\n    });\n  };\n  el.querySelectorAll('input[name=mode]').forEach((r) => r.onchange = () => {\n    if (r.value === 'single' && state.classes.length > 1 && !confirm('1クラスで使う場合、一番上のクラスだけを表示します（他のクラスのデータは消えません）。切り替えますか？')) { render(); return; }\n    setMode(r.value); render();\n  });\n  el.querySelectorAll('[data-roster]').forEach((b) => b.onclick = () => {\n    syncDraft(); const i = +b.dataset.roster; ui.openRoster = ui.openRoster === i ? null : i; render();\n  });\n  el.querySelectorAll('[data-cdel]').forEach((b) => b.onclick = () => {\n    syncDraft(); const i = +b.dataset.cdel;\n    if (draft[i].oldName && !confirm(`${draft[i].oldName} を削除します。保存すると、このクラスのチェック記録も消えます。よろしいですか？`)) return;\n    draft.splice(i, 1); ui.openRoster = null; render();\n  });\n  if ($('#cAdd')) $('#cAdd').onclick = () => { syncDraft(); draft.push({ name: '', count: 40, names: [], oldName: null }); render(); el.querySelectorAll('.cName')[draft.length - 1]?.focus(); };\n  $('#cSave').onclick = async () => {\n    syncDraft();\n    const list = (state.mode === 'single' ? [...visibleDraft, ...draft.slice(1)] : draft).map((c) => ({\n      name: String(c.name).trim(), count: parseInt(c.count, 10), names: c.names || [], oldName: c.oldName || null,\n    }));\n    for (const c of list) {\n      if (!c.name) return toast('クラス名が空欄の行があります');\n      if (/[\\[\\]\\*\\?\\/\\\\:]/.test(c.name) || c.name === '設定' || c.name === '提出物') return toast(`「${c.name}」はクラス名に使えません`);\n      if (!(c.count >= 1 && c.count <= MAX_STUDENTS)) return toast(`人数は1〜${MAX_STUDENTS}人で入力してください`);\n      if (c.names.filter(Boolean).length > c.count) return toast(`${c.name}：名前の数が人数より多くなっています`);\n    }\n    if (new Set(list.map((c) => c.name)).size !== list.length) return toast('同じクラス名があります');\n    if (list.length > MAX_CLASSES) return toast(`クラスは${MAX_CLASSES}クラスまでです`);\n    $('#cSave').disabled = true;\n    if (await saveClasses(list)) { ui.classDraft = null; ui.openRoster = null; toast('クラス設定を保存しました'); }\n    render();\n  };\n\n  // データ\n  if ($('#dBackup')) $('#dBackup').onclick = () => {\n    const blob = new Blob([JSON.stringify(state, null, 1)], { type: 'application/json' });\n    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `提出物チェッカー_${todayStr()}.json`; a.click();\n    setTimeout(() => URL.revokeObjectURL(a.href), 1000);\n    try { localStorage.setItem(LS_BACKUP, String(Date.now())); } catch (e) {}\n  };\n  if ($('#dRestore')) $('#dRestore').onchange = (e) => {\n    const f = e.target.files[0]; if (!f) return;\n    const r = new FileReader();\n    r.onload = async () => {\n      try {\n        const d = JSON.parse(r.result);\n        if (!Array.isArray(d.classes) || !Array.isArray(d.tasks)) throw new Error();\n        if (!confirm('今のデータを、バックアップの内容で置き換えます。よろしいですか？')) return;\n        await importState(d); toast('復元しました'); render();\n      } catch (err) { alert('バックアップファイルを読み込めませんでした'); }\n    };\n    r.readAsText(f);\n  };\n  $('#dSample').onclick = async () => {\n    if ((state.classes.length || state.tasks.length) && !confirm('今のデータを消して、サンプルデータを入れます。よろしいですか？')) return;\n    await importState(buildSample()); ui.view = 'status'; toast('サンプルデータを入れました'); render();\n  };\n  $('#dReset').onclick = async () => {\n    if (!confirm('すべてのクラス・提出物・チェック記録を消して初期化します。元に戻せません。よろしいですか？')) return;\n    await importState(Object.assign(emptyState(), { mode: state.mode })); toast('初期化しました'); render();\n  };\n}\n\n// ===== 使い方 =====\nfunction renderHelp() {\n  $('#view-help').innerHTML = `<div class=\"card help\">\n    <h2>使い方</h2>\n    <p>登録する。チェックする。残っている番号を見る。毎日の回収確認は、この流れだけです。</p>\n    <h3>はじめの準備</h3>\n    <ol>\n      <li>「提出物」画面の下にある「クラス設定」で、クラス名と人数を入れて「クラス設定を保存」を押します（最大${MAX_CLASSES}クラス、1クラス${MAX_STUDENTS}人まで）。</li>\n      <li>名前は入れなくても使えます。出席番号だけで、チェック・声かけ・印刷まで動きます。</li>\n    </ol>\n    <h3>提出物を登録する</h3>\n    <ol>\n      <li>「提出物」→「＋ 追加」を押します。</li>\n      <li>提出物名と対象クラスを選びます（この二つだけが必須です）。「ワークブック」「単語テスト直し」などは候補から選べます。</li>\n      <li>「伝えた日」を入れると配ってからの日数が、「締切」を入れると今日締切・締切超過が分かるようになります。</li>\n    </ol>\n    <h3>チェックする</h3>\n    <ul>\n      <li>「チェック入力」でクラスを選び、出した子の番号と提出物が交わるマスを押します。もう一度押すと外れます。</li>\n      <li>冊数が多いときは「番号で入力」が便利です。提出物を選び、番号を入れて Enter。続けて次の番号を入力できます。</li>\n      <li>間違えたときは「↶ 取り消し」で直前の操作を戻せます。</li>\n      <li>「完了した提出物を隠す」「未提出のある番号だけ」で表示を絞れます。</li>\n    </ul>\n    <h3>残っている番号を見る</h3>\n    <ul>\n      <li>「回収状況」では、今日が締切の提出物が上に並びます。カードの色は <span class=\"badge b-collect\">回収中</span> <span class=\"badge b-few\">あと5人以下</span> <span class=\"badge b-over\">締切超過</span> <span class=\"badge b-done\">全員提出</span> です。</li>\n      <li>「声かけビュー」では、未提出を生徒ごとにまとめます。「3番：ワークブック、英作文」をそのままコピーできます。</li>\n      <li>未提出が2件以上重なっている番号は「要注意リスト」に出ます。理由を決めつけるためではなく、見落とさないための目印です。</li>\n    </ul>\n    <h3>紙にする</h3>\n    <ul>\n      <li>「印刷プレビュー」で、クラス別（掲示用）・個人別（本人に渡す）・白紙のチェック表をA4で印刷・PDF保存できます。</li>\n    </ul>\n    <h3>安心して使うために</h3>\n    <ul>\n      <li>データは${IS_GAS ? 'あなたのGoogleドライブのスプレッドシート' : 'この端末のブラウザ'}の中だけに保存し、外部へは送りません。${IS_GAS ? '同じGoogleアカウントなら、パソコン・iPadなどどの端末からでも同じデータを使えます。ほかの端末で入力した内容は「↻ 最新に更新」で読み込めます。' : '複数の端末で使うには <a href=\"setup.html\" target=\"_blank\" rel=\"noopener\">Googleアカウントで同期する設定</a> をしてください。'}</li>\n      <li>評価や点数を書く欄はありません。記録するのは「出したか、まだか」だけです。</li>\n      <li>名前を入れるかどうかは、学校のルールに合わせて判断してください。</li>\n    </ul>\n  </div>`;\n}\n\n// ===== 起動 =====\ndocument.querySelectorAll('#modeDlg [data-mode]').forEach((b) => b.onclick = async () => {\n  if (b.dataset.mode === 'sample') { await importState(buildSample()); ui.view = 'status'; render(); return; }\n  setMode(b.dataset.mode);\n  if (!state.classes.length) { ui.view = 'tasks'; toast('まず、クラス名と人数を登録してください'); }\n  render();\n});\n$('#storageNote').innerHTML = IS_GAS\n  ? 'データはあなたのGoogleドライブのスプレッドシートに保存されます。同じGoogleアカウントなら、どの端末からでも同じデータを使えます。'\n  : 'この画面で入力したデータは、この端末のこのブラウザの中だけに保存されます。複数の端末で使いたい場合は <a href=\"setup.html\" target=\"_blank\" rel=\"noopener\">Googleアカウントで同期する設定</a> をしてください。';\nreload();\n</script>\n</body>\n</html>\n";
