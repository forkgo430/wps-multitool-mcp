/**
 * P0 · structuredContent 体积闸门（提质小卡 2026-11-08）
 *
 * 判据（质量红线）：
 *   1. 小 JSON 回包（≤2048 字节）→ structuredContent 镜像在
 *   2. 大 JSON 回包（>2048 字节）→ 只有 text（无 structuredContent）
 *   3. text 块（数据主载体）逐字节不变——闸门只影响 structuredContent 副本
 *   4. 闸门边界：恰好 2048 字节仍在镜像内
 *
 * 全部 mock 注入（registry.register 自定义 handler 返回可控 content），不起 WPS。
 */
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));

let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

const GATE = 2048;

/** 注册一个返回指定文本的测试工具（唯一名），跑 handleToolCall 拿回包 */
let probeSeq = 0;
async function callWithText(text) {
  probeSeq++;
  const name = `__gateProbe${probeSeq}`;
  toolRegistry.register(
    { name, description: '闸门探针', inputSchema: { type: 'object', properties: {} }, category: 'common' },
    async () => ({ id: 'probe', success: true, content: [{ type: 'text', text }] })
  );
  const server = new WpsMcpServer();
  // handleToolCall 是 TS private；JS 运行时可直接访问（FAST mock 层惯例）
  return server.handleToolCall(name, {});
}

(async () => {
  console.log('=== P0 structuredContent 体积闸门 ===');

  // 1. 小 JSON（≤2KB）→ 镜像在
  const small = JSON.stringify({ ok: true, rows: Array.from({ length: 8 }, (_, i) => ({ i, v: 'x' + i })) });
  if (Buffer.byteLength(small, 'utf8') > GATE) throw new Error('探针数据超过闸门，测试失效');
  const r1 = await callWithText(small);
  ok('小 JSON（≤2048B）structuredContent 镜像在', !!r1.structuredContent && r1.structuredContent.ok === true);
  // text 可能带快照层前缀（探针工具无 annotations）；JSON 原文完整保留在尾部即主载体不变
  ok('小 JSON 的 text 完整含原文（主载体不动）', r1.content[0].text.endsWith(small), `${Buffer.byteLength(small)}B`);

  // 2. 大 JSON（>2KB）→ 只有 text
  const big = JSON.stringify({ ok: true, rows: Array.from({ length: 300 }, (_, i) => ({ i, v: 'padding-' + i + '-xxxxxxxxxx' })) });
  if (Buffer.byteLength(big, 'utf8') <= GATE) throw new Error('探针数据没超过闸门，测试失效');
  const r2 = await callWithText(big);
  console.log('  [debug] r2 text head=', JSON.stringify(r2.content[0].text.slice(0, 120)));
  ok('大 JSON（>2048B）无 structuredContent（只留 text）', r2.structuredContent === undefined);
  // text 可能带快照层前缀（探针工具无 annotations）；JSON 原文完整保留在尾部即主载体不变
  ok('大 JSON 的 text 完整含原文（不截断不改写）', r2.content[0].text.endsWith(big));

  // 3. 边界：恰好 2048 字节 → 仍在镜像内；+1B → 消失
  const base = JSON.stringify({ v: '' });
  const fill = GATE - Buffer.byteLength(base, 'utf8');   // 插入后总长恰为 2048
  const edge = JSON.stringify({ v: 'a'.repeat(fill) });
  const edgeBytes = Buffer.byteLength(edge, 'utf8');
  const r3 = await callWithText(edge);
  ok(`边界 ${edgeBytes}B（=2048）镜像在`, edgeBytes === GATE && !!r3.structuredContent);
  // +1B 用例：JSON 体内多一字符（trim 后仍 2049——闸门按 trimmed 长度判，与镜像对象一致）
  const over = JSON.stringify({ v: 'a'.repeat(fill + 1) });
  const overBytes = Buffer.byteLength(over, 'utf8');
  const r4 = await callWithText(over);
  ok(`边界 +1B（${overBytes}B）镜像消失`, overBytes === GATE + 1 && r4.structuredContent === undefined && r4.content[0].text.endsWith(over));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
