/**
 * Pos: 清理探针遗留的未保存 Word 文档（只关名字匹配 文档N/文档 N 的未保存文档——误伤护栏）。
 * 一旦我被修改，请更新我的头部注释，以及所属目录的md。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-closeprobe-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
(async () => {
  const r = await call('wps_word_get_open_documents', {});
  const t = String((r && r.content && r.content[0] && r.content[0].text) || '');
  // 名单里找出 文档N 形态的未保存文档
  const names = [...t.matchAll(/(文档\s?\d+)/g)].map((m) => m[1]);
  const uniq = [...new Set(names)];
  console.log('探针遗留候选：', uniq.join(', ') || '无');
  for (const n of uniq) {
    try {
      const c = await call('wps_word_close_document', { save: false, document: n });
      console.log(`closed ${n}:`, String((c && c.content && c.content[0] && c.content[0].text) || '').slice(0, 50));
    } catch (e) {
      console.log(`close ${n} 失败:`, e.message);
    }
  }
  process.exit(0);
})().catch((e) => { console.error(e); process.exit(1); });
