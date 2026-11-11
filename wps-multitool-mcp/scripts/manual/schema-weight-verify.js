/**
 * B 段 · 默认档 40 schema 减重 · 量化 + 门禁（提质合并卡）
 *
 * 口径（ARCH-AUDIT）：tools/list payload = 默认档 40 工具的 JSON 字节（与判定实测 36,456B 同口径）。
 * 门禁：payload ≤ 36000B（防回弹）；同时输出 top 体积排行（施工参考，非断言）。
 * 参数名/类型/必填/枚举一字不动——本脚本只量测，压缩发生在 src/tools/*.ts 源码。
 */
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));

let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== 默认档 schema 体积量测 ===');
  const server = new WpsMcpServer();
  const { tools } = server.registry.listTools();
  const payload = JSON.stringify({ tools: tools.map((t) => ({
    name: t.name,
    description: t.description,
    inputSchema: t.inputSchema,
    // 判定阶段修正（2026-11-11）：量测必须用 getToolAnnotations()（tools/list 实际发送的全量注解），
    // 定义内联 t.annotations 是轻量版——两者差约 4.7KB 盲区，门禁用轻量版会漏判回弹。
    ...(getToolAnnotations(t.name) ? { annotations: getToolAnnotations(t.name) } : {}),
  })) });
  const bytes = Buffer.byteLength(payload, 'utf8');
  console.log(`默认档工具数：${tools.length}`);
  console.log(`payload（name+description+inputSchema+annotations）：${bytes}B`);

  // 门禁：≤36000B（提质合并卡 B 段目标 = 40,093 → ≤36,000，约 -10%；描述压缩大头）
  ok(`默认档 payload ≤ 36000B（门禁防回弹）`, bytes <= 36000, `实测 ${bytes}B`);

  // top 排行（施工参考）
  const rows = tools.map((t) => ({
    name: t.name,
    desc: Buffer.byteLength(t.description || '', 'utf8'),
    schema: Buffer.byteLength(JSON.stringify(t.inputSchema || {}), 'utf8'),
  })).sort((a, b) => (b.desc + b.schema) - (a.desc + a.schema));
  console.log('top 10（description+schema 字节）：');
  for (const r of rows.slice(0, 10)) console.log(`  ${r.name}  desc=${r.desc} schema=${r.schema}`);
  const totalDesc = rows.reduce((a, r) => a + r.desc, 0);
  const totalSchema = rows.reduce((a, r) => a + r.schema, 0);
  console.log(`合计 description=${totalDesc}B schema=${totalSchema}B`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
