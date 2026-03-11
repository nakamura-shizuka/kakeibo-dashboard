/**
 * build.js
 * GitHub Pages 用の静的 HTML を生成するビルドスクリプト。
 *
 * 使い方:
 *   1. GAS_URL に GAS Web App の URL を設定する
 *   2. node build.js を実行する
 *   3. docs/index.html が生成される
 */

const fs = require('fs');
const path = require('path');

// ============================================================
// GAS Web App URL（GASエディタ > デプロイ > 新しいデプロイ で取得）
// ============================================================
const GAS_URL = 'https://script.google.com/macros/s/AKfycbzsZ61S4xLGU8oMMkI4fALwUEEl5JPyyPNhILZlq3tCY7EBLxQk4JLG4Eb2jTYbOn9X/exec';
// ============================================================

if (GAS_URL.includes('REPLACE_WITH_YOUR_DEPLOYMENT_ID')) {
    console.error('❌ エラー: build.js の GAS_URL を実際の GAS Web App URL に書き換えてください。');
    process.exit(1);
}

const DIR = __dirname;
const PARTS = ['styles', 'helpers', 'charts', 'records', 'modals', 'scripts'];

// 各パーツファイルを読み込む
const contents = {};
PARTS.forEach(function (name) {
    const filePath = path.join(DIR, name + '.html');
    if (!fs.existsSync(filePath)) {
        console.error('❌ ファイルが見つかりません: ' + filePath);
        process.exit(1);
    }
    contents[name] = fs.readFileSync(filePath, 'utf8');
});

// index.html を読み込む
let html = fs.readFileSync(path.join(DIR, 'index.html'), 'utf8');

// GAS テンプレートインクルードを展開
PARTS.forEach(function (name) {
    html = html.replace("<?!= include('" + name + "'); ?>", contents[name]);
});

// GAS_API_URL のプレースホルダーを実際の URL に置換
html = html.replace("'%%GAS_URL%%'", "'" + GAS_URL + "'");

// 出力先を作成
const outDir = path.join(DIR, 'docs');
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'index.html'), html, 'utf8');

console.log('✅ Build complete: docs/index.html');
console.log('   GAS URL: ' + GAS_URL);
