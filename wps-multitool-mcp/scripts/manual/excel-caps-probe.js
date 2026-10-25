/**
 * Excel 侧能力探测（第 4 期新增接口的真机验证）
 * ------------------------------------------------
 * 探什么：__probeExcelCaps 在【临时工作簿】上逐个试 WPS 12.1 到底认不认这些接口，
 * 改完改回、探测完删表 —— 不碰用户任何文档。
 *
 * 判据要说清楚：supported 只代表"调用没抛异常"，不代表语义完全对
 * （比如 AutoFit 之后宽度是否真的合适，那要肉眼看）。
 *
 * 用法：node scripts/manual/excel-caps-probe.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

// 第 4 期新增的 10 个工具，以及它们依赖的 PS 动作名
const NEW_TOOLS = [
  'wps_excel_get_sheet_info',
  'wps_excel_auto_fit',
  'wps_excel_set_wrap_text',
  'wps_excel_find_in_sheet',
  'wps_excel_delete_chart',
  'wps_excel_get_pivot_tables',
  'wps_excel_refresh_pivot_tables',
  'wps_excel_clear_pivot_table',
  'wps_excel_set_sheet_page_setup',
  'wps_excel_set_sheet_print_titles',
];

(async () => {
  let pass = 0;
  let fail = 0;
  const ok = (label, cond, extra) => {
    if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
    else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
  };

  console.log('\n=== A. 10 个新工具的注册与元数据（不需要 WPS）===');
  ok('新工具全部注册', NEW_TOOLS.every((n) => toolRegistry.hasTool(n)),
    `缺：${NEW_TOOLS.filter((n) => !toolRegistry.hasTool(n)).join(',') || '无'}`);
  const noAlias = NEW_TOOLS.filter((n) => {
    const m = getToolAnnotations(n);
    return !m; // 注解缺失的单独看 search-meta
  });
  ok('新工具都有注解', noAlias.length === 0, noAlias.join(','));

  console.log('\n=== B. 真机能力探测（需要 WPS Excel 在跑）===');
  const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
  let probe;
  try {
    const r = await wpsClient.executeMethod('__probeExcelCaps', {}, 'excel');
    probe = r.data;
    if (!r.success) throw new Error(r.error);
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log(`  ⚠ 探测未执行：${msg}`);
    console.log('    （需要 WPS 表格可访问。这一项本轮不判失败，但也不算数 ——');
    console.log('     意味着这 10 个接口在真机上【未验证】。）');
    console.log(`\n汇总：${pass} 通过 / ${fail} 失败（真机探测跳过）`);
    process.exit(fail === 0 ? 0 : 1);
  }

  const supported = probe.supported;
  const unsupported = probe.unsupported || [];
  console.log(`  临时工作簿已清理：${(probe.leaked || []).length === 0 ? '是' : '否 —— ' + probe.leaked.join('；')}`);
  console.log(`  探测项 ${probe.total} 个：支持 ${supported} / 不支持 ${unsupported.length}`);
  if (unsupported.length) console.log(`  不支持：${unsupported.join(', ')}`);
  ok('临时工作簿删干净了（探测不留垃圾）', (probe.leaked || []).length === 0);
  ok('核心接口 sheet.usedRange 受支持', !unsupported.includes('sheet.usedRange'));
  ok('核心接口 sheet.autoFitColumn 受支持', !unsupported.includes('sheet.autoFitColumn'));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})();
