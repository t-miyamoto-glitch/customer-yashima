// index.html と src/*.js を 1 つの HTML にまとめて dist/ に出力する（配布用）
//   node tools/build.js
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
let html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
html = html.replace(/<script src="(src\/[^"]+)"><\/script>/g, (_, src) =>
  '<script>\n' + fs.readFileSync(path.join(root, src), 'utf8').replace(/<\/script/gi, '<\\/script') + '\n</script>');
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
const out = path.join(root, 'dist', '請求書照合ツール.html');
fs.writeFileSync(out, html);
console.log('wrote', out);
