/**
 * B3 探针：照着错误提示里写的参数名去关"从未保存过"的文档，看能不能关掉。
 *
 * 背景：三个 close 工具里，只有 Word 侧的 PS 错误提示写的是 saveChanges=false，
 * 而 wps_word_close_document 的【公开 schema 参数名】是 save。
 * 照提示改 → 工具收到一个它不认的键 → save 还是 undefined →
 * TS 兜成 saveChanges=true → 护栏照样触发 → 用户卡死在同一个错误里。
 *
 * 这个探针走两遍：
 *   第 1 遍：照提示传 saveChanges=false（错的）→ 预期关不掉
 *   第 2 遍：传 save=false（对的）→ 预期能关掉
 *
 * 纪律：自己建的文档自己关；建之前先记打开列表差集。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-b3-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '').replace(/\s+/g, ' ');

let pass = 0, fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

(async () => {
  console.log('\n=== 0. 认领临时文档（打开列表差集）===');
  const before = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const beforeNames = new Set(((before.data && before.data.documents) || []).map((d) => d.name));
  const created = await call('wps_word_create_document', {});
  const after = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const mine = ((after.data && after.data.documents) || []).filter((d) => !beforeNames.has(d.name));
  ok('差集 = 恰好 1 个新文档', mine.length === 1, mine.map((d) => d.name).join(','));
  const myName = mine.length ? mine[0].name : null;
  if (!myName) { console.log('  ✗ 认领失败'); process.exit(1); }
  await call('wps_word_insert_text', { text: 'B3 探针' });

  console.log('\n=== 1. 照【错误提示里写的名字】传 saveChanges=false ===');
  let r = await call('wps_word_close_document', { name: myName, saveChanges: false });
  console.log('  回包：', txt(r).slice(0, 200));
  const stillOpen = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const stillThere = ((stillOpen.data && stillOpen.data.documents) || []).some((d) => d.name === myName);
  console.log(`  文档还在吗：${stillThere}`);
  if (stillThere) {
    console.log('  → 照提示改【关不掉】，这就是 B3 现场：提示里的参数名工具根本不认，');
    console.log('    save 仍是 undefined → TS 兜成 saveChanges=true → 护栏照样触发。');
  }
  ok('照提示传 saveChanges=false 关不掉（复现 B3）', stillThere, `还在=${stillThere}`);

  console.log('\n=== 2. 传【schema 里真的有的】save=false ===');
  r = await call('wps_word_close_document', { name: myName, save: false });
  console.log('  回包：', txt(r).slice(0, 160));
  const after2 = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const left = ((after2.data && after2.data.documents) || []).map((d) => d.name);
  ok('传 save=false 能关掉', !left.includes(myName), `剩余 ${left.length} 个`);
  ok('没动别人的文档', left.every((n) => beforeNames.has(n)), `开工前 ${beforeNames.size} 个 / 现在 ${left.length} 个`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(0);
})().catch((e) => {
  console.error('探针异常：', e && e.message ? e.message : e);
  process.exit(2);
});
