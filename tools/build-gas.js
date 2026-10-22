// gas/server.gs と index.html から、貼り付け用の1ファイル版 gas/teishutsu-checker.gs を作る。
// 使い方: node tools/build-gas.js
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const server = fs.readFileSync(path.join(root, 'gas/server.gs'), 'utf8');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

const out = `// 提出物チェッカー（Google Apps Script 用・1ファイル版）
// このファイルの中身をすべてコピーして、script.google.com のプロジェクトに貼り付けてください。
// 設定手順: https://daharakazu-glitch.github.io/homework-checker/setup.html
// ※ tools/build-gas.js で gas/server.gs と index.html から自動生成しています。直接編集しないでください。

${server.trimEnd()}

// ===== 画面（index.html の内容） =====
// 貼り付けで途切れないよう、1行ずつの文字列に分けている
const APP_HTML = [
${html.split('\n').map((l) => '  ' + JSON.stringify(l)).join(',\n')}
].join('\\n');
`;

fs.writeFileSync(path.join(root, 'gas/teishutsu-checker.gs'), out);
console.log('gas/teishutsu-checker.gs を作成しました（' + Math.round(out.length / 1024) + ' KB）');
