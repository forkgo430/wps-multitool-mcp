/**
 * 注解端到端验证
 * --------------
 * 前面的 annotation-verify 只验了"注解算得对不对"。
 * 这一条验的是"客户端真的能在 tools/list 里看到它" —— 因为中间隔着
 * MCP 协议的序列化，字段名拼错、类型不对，客户端那边就静默忽略了。
 *
 * 做法：起一个真的 stdio server 子进程，用官方 SDK 的 Client 连上去，
 * 调一次 tools/list，看返回的每个 tool 上有没有 annotations。
 *
 * 用法：node scripts/manual/annotation-e2e.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

(async () => {
  const { Client } = require('@modelcontextprotocol/sdk/client/index.js');
  const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');

  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [path.join(ROOT, 'dist', 'index.js')],
    cwd: ROOT,
    stderr: 'inherit',
  });

  const client = new Client({ name: 'annotation-probe', version: '1.0.0' }, { capabilities: {} });
  await client.connect(transport);

  const { tools } = await client.listTools();

  let pass = 0;
  let fail = 0;
  const check = (label, cond, extra) => {
    if (cond) {
      pass++;
      console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
    } else {
      fail++;
      console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
    }
  };

  console.log(`\ntools/list 返回 ${tools.length} 个工具（默认基础档）`);

  const withAnn = tools.filter((t) => t.annotations && typeof t.annotations === 'object');
  check('每个工具都带 annotations', withAnn.length === tools.length, `${withAnn.length}/${tools.length}`);

  // 字段名必须与 MCP 规范一致，拼错了客户端会直接忽略
  const sample = tools.find((t) => t.annotations)?.annotations || {};
  check(
    '字段名是规范里的四个',
    ['readOnlyHint', 'destructiveHint', 'idempotentHint', 'openWorldHint'].every((k) => k in sample),
    Object.keys(sample).join(', ')
  );

  const ro = tools.filter((t) => t.annotations?.readOnlyHint === true).length;
  check('只读工具存在（不是全部标 false 的空壳）', ro > 0, `只读 ${ro} / 共 ${tools.length}`);

  // 抽一个已知只读的工具，看注解是不是真的传到了
  const status = tools.find((t) => t.name === 'wps_common_status');
  check('wps_common_status 标注为只读', status?.annotations?.readOnlyHint === true,
    status ? JSON.stringify(status.annotations) : '该工具不在基础档里，跳过此条');

  // wps_execute_method 由 mcp-server 直接注册，不在自动生成的映射表里，
  // 注解是手写的 EXTRA_ANNOTATIONS —— 别让以后有人以为它漏了
  const exec = tools.find((t) => t.name === 'wps_execute_method');
  check('wps_execute_method（server 直注册）也带注解', !!exec?.annotations,
    exec ? JSON.stringify(exec.annotations) : '不在基础档里');

  console.log('\n一个工具的完整返回样例：');
  const demo = tools.find((t) => t.name === 'wps_common_status') || tools[0];
  console.log(JSON.stringify({ name: demo.name, annotations: demo.annotations }, null, 2));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);

  await client.close();
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('E2E 失败:', e && e.message ? e.message : e);
  process.exit(1);
});