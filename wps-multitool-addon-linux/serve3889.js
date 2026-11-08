#!/usr/bin/env node
/**
 * Input: 无参数（服务本目录下 addon-wps / addon-et / addon-wpp）或 --root <目录>
 * Output: http://127.0.0.1:3889/<type>/ 的静态加载项文件服务
 * Pos: WPS Linux 加载项在线入口（jsplugins.xml 的 url 指向这里）。
 *      仅绑 127.0.0.1；带 WPS webview 需要的 CORS 头；缺 Default-Content-Type 时 js/maps 正常回。
 * 用法：node serve3889.js [--root <addon-linux 目录>] [--port 3889]
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

let root = __dirname;
const argv = process.argv.slice(2);
const iRoot = argv.indexOf('--root');
if (iRoot >= 0 && argv[iRoot + 1]) root = path.resolve(argv[iRoot + 1]);
const port = (() => { const i = argv.indexOf('--port'); const v = i >= 0 ? parseInt(argv[iRoot === i ? i + 1 : i + 1], 10) : NaN; return Number.isFinite(v) ? v : 3889; })();
if (!fs.existsSync(path.join(root, 'addon-wps'))) {
  console.error(`✗ ${root} 下没有 addon-wps/——请在 addon-linux 树目录运行，或用 --root 指定`);
  process.exit(1);
}

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'application/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8', '.xml': 'text/xml; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.ico': 'image/x-icon' };

http.createServer((req, res) => {
  const url = decodeURIComponent((req.url || '/').split('?')[0]);
  const type = url.replace(/^\/+/, '').split('/')[0];   // /wps/... → wps
  if (!['wps', 'et', 'wpp'].includes(type)) { res.writeHead(404); res.end('not found'); return; }
  let rel = url.slice(type.length + 1).replace(/^\/+/, '');
  if (rel === '' || rel.endsWith('/')) rel += 'index.html';
  const file = path.normalize(path.join(root, 'addon-' + type, rel));
  if (!file.startsWith(path.join(root, 'addon-' + type))) { res.writeHead(403); res.end('forbidden'); return; }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404, { 'Access-Control-Allow-Origin': '*' }); res.end('not found'); return; }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-store',
    });
    res.end(data);
  });
}).listen(port, '127.0.0.1', () => {
  console.log(`serve3889: http://127.0.0.1:${port}/{wps,et,wpp}/ ← ${root}（仅本机回环）`);
});
