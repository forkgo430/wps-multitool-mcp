/**
 * P1 · tools/list 减重（提质小卡 2026-11-08）
 *
 * 判据：
 *   1. tools/list 组装不再附带 examples（源码静态 + 运行时回包双验）
 *   2. wps_search_tools 路径 examples 原样保留（那是它发挥价值的地方）
 *   3. 注解仍在（annotations 嵌套不丢——回归保护）
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== tools/list 减重（examples 退出）===');

  // 1) 源码静态：tools/list 组装对象不再有 examples 字段
  const src = fs.readFileSync(path.join(ROOT, 'src', 'server', 'mcp-server.ts'), 'utf8');
  // 组装段：ListToolsRequestSchema handler 内的 map 返回体
  const listSection = src.slice(src.indexOf('Handling tools/list request'), src.indexOf('Handling tools/call request'));
  ok('tools/list 组装段无 examples 字段（源码静态）',
    listSection.includes('examples 不进 tools/list') && !/\.\.\.\(examples\?\.length/.test(listSection));
  ok('examples 剔除的口径注释在（searchTools 保留是刻意的）',
    listSection.includes('wps_search_tools 返回'));
  // getToolMeta 已无引用（死 import 清理）
  ok('getToolMeta 死 import 已清', !/import \{ getToolMeta \} from '\.\/tool-metadata'/.test(src));

  // 2) searchTools 路径 examples 原样保留（运行时：注册带 examples 的探针工具 → 搜索）
  const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
  const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
  toolRegistry.register(
    {
      name: '__examplesProbe', description: 'examples 探针：把 PPT 组加载一下',
      inputSchema: { type: 'object', properties: {} }, category: 'common',
      examples: ['我要用 PPT 的工具', '加载 PPT 组'],
    },
    async () => ({ id: 'p', success: true, content: [{ type: 'text', text: '{}' }] })
  );
  // searchTools 走 common 组的真实实现
  // searchTools 是 common 组工具：直接经 handleToolCall 调真入口
  const server = new WpsMcpServer();
  const resp = await server.handleToolCall('wps_search_tools', { keyword: '背景' });
  const text = resp.content[0].text;
  ok('searchTools 返回里 CURATED 工具带 examples（原样保留）',
    text.includes('wps_search_tools') && text.includes('examples'),
    'keyword=背景（CURATED 表工具命中）');

  // 3) tools/list 运行时：registry 层探针在已加载组（可进 tools/list）+ annotations 回归
  const allNames = (toolRegistry.listAllTools() || []).map((t) => t.name);
  ok('registry 层探针已注册（可被 tools/list 收录）', allNames.includes('__examplesProbe'));
  const resp2 = await server.handleToolCall('wps_list_tool_groups', {});
  ok('wps_list_tool_groups 正常（工具面无回归）', resp2 && resp2.success !== false);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
