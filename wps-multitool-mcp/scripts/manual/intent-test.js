/**
 * Input: 无
 * Output: 意图识别（wps_route_intent）的断言结果
 * Pos: 意图识别的手工验证脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用法（先 npm run build）：node scripts/manual/intent-test.js
 *
 * 说明：classifyIntent 是纯函数，直接断言；handler 那条路会真的 load 工具组，也一并验。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { classifyIntent, routeIntentHandler, matchDirectRoute, matchTaskPackage } = require(path.join(ROOT, 'dist', 'tools', 'common', 'intent.js'));
const { resolveToolAlias } = require(path.join(ROOT, 'dist', 'server', 'tool-aliases.generated.js'));

toolRegistry.registerAll(allTools);

const results = [];
function check(name, cond, detail) {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✅' : '❌'} ${name}${detail ? '  → ' + detail : ''}`);
}
const groupsOf = (t) => classifyIntent(t).map((m) => m.group);

console.log('=== 单意图 ===');
const cases = [
  ['做一份产品发布会的 PPT', 'ppt'],
  ['帮我把这份 excel 的透视表更新一下', 'excel'],
  ['把论文的题注和交叉引用整理一下', 'word'],
  ['导出成 PDF 给我', 'common'],
];
for (const [task, expect] of cases) {
  const g = groupsOf(task);
  check(`"${task}" → 命中 ${expect}`, g.includes(expect), g.join(',') || '未命中');
}

console.log('\n=== 多意图 / 排除 ===');
const multi = groupsOf('把 word 里的表格做成 PPT');
check('"word 里表格做成 PPT" 同时命中 word 与 ppt', multi.includes('word') && multi.includes('ppt'), multi.join(','));
const none = groupsOf('今天天气不错，随便聊聊');
check('无关语句不误报', none.length === 0, none.join(',') || '未命中');

console.log('\n=== 置信度与建议工具 ===');
const r = classifyIntent('我要做 PPT，需要配色和动画');
console.log('  ' + JSON.stringify(r.map((m) => ({ group: m.group, score: m.score, conf: m.confidence, suggested: m.suggested.slice(0, 3) }))));
const ppt = r.find((m) => m.group === 'ppt');
check('命中 ppt 且给了建议工具', !!ppt && ppt.suggested.length > 0, ppt ? ppt.suggested.slice(0, 3).join(', ') : 'n/a');
check('置信度在 0~1 之间', !!ppt && ppt.confidence > 0 && ppt.confidence <= 1, ppt ? String(ppt.confidence) : 'n/a');
check('命中时会带 matched 关键词', !!ppt && ppt.matched.length > 0, ppt ? ppt.matched.join(',') : 'n/a');

async function main() {
console.log('\n=== handler：默认只建议，不自动加载 ===');
const before = toolRegistry.getGroups().find((g) => g.group === 'ppt').fullyLoaded;
let out = await routeIntentHandler({ task: '做一份 PPT' });
let payload = JSON.parse(out.content[0].text);
const afterSuggest = toolRegistry.getGroups().find((g) => g.group === 'ppt').fullyLoaded;
check('默认不加载（只建议）', before === false && afterSuggest === false);
check('type=group 且返回 needLoad 含 ppt', payload.type === 'group' && payload.result.needLoad.includes('ppt'), payload.result.needLoad.join(','));
check('没有 autoLoaded', payload.result.autoLoaded.length === 0);

console.log('\n=== handler：autoLoad=true 会真的加载 ===');
out = await routeIntentHandler({ task: '做一份 PPT', autoLoad: true });
payload = JSON.parse(out.content[0].text);
const afterAuto = toolRegistry.getGroups().find((g) => g.group === 'ppt').fullyLoaded;
check('ppt 组已完整加载', afterAuto === true);
check('返回 autoLoaded 含 ppt', payload.result.autoLoaded.includes('ppt'), payload.result.autoLoaded.join(','));
check('nextStep 提示已刷新', /已自动加载/.test(payload.nextStep), payload.nextStep);

console.log('\n=== handler：没识别出来时指向 search_tools ===');
out = await routeIntentHandler({ task: '随手写点什么' });
payload = JSON.parse(out.content[0].text);
check('明确说没识别出来', payload.type === 'group' && payload.result.matches.length === 0 && !!payload.result.note);
check('nextStep 指向 wps_search_tools', /wps_search_tools/.test(payload.nextStep));

console.log('\n=== 三层意图：短路口（直觉直答，≥15 条）===');
const directCases = [
  ['帮我把这份文件保存一下', 'wps_common_save'],
  ['这份合同另存为一份新文件', 'wps_common_save_as'],
  ['把这份报告导出pdf', 'wps_convert_to_pdf'],
  ['读取单元格 A1 的内容', 'wps_excel_get_cell_value'],
  ['写入单元格 B2，值是 42', 'wps_excel_set_cell_value'],
  ['把这段文字加粗', 'wps_word_set_font'],
  ['把标题设为斜体', 'wps_word_set_font'],
  ['把字号改成 16', 'wps_word_set_font'],
  ['把字体颜色改成蓝色', 'wps_word_set_text_color'],
  ['在第 3 行前插入一行', 'wps_excel_insert_rows'],
  ['在 B 列前插入一列', 'wps_excel_insert_columns'],
  ['把第 5 行删除一行', 'wps_excel_delete_rows'],
  ['删掉一列 D', 'wps_excel_delete_columns'],
  ['冻结首行', 'wps_excel_freeze_panes'],
  ['冻结首列', 'wps_excel_freeze_panes'],
  ['给文档加页码', 'wps_word_page_number'],
  ['新建工作簿', 'wps_excel_create_workbook'],
  ['建个文档开始写', 'wps_word_create_document'],
  ['新建ppt开始做演示', 'wps_ppt_create_presentation'],
  ['给演示加一页', 'wps_ppt_add_slide'],
  ['统计一下这篇文档多少字', 'wps_word_get_document_stats'],
  ['先做个连接测试', 'wps_check_connection'],
];
for (const [text, tool] of directCases) {
  const d = matchDirectRoute(text);
  check('短路口「' + text + '」→ ' + tool, !!d && d.tool === tool, d ? 'matched=' + d.matched.join(',') : '未命中');
}
{
  const d = matchDirectRoute('写入单元格 A1，值是 42');
  check('短路口参数提取：单元格地址进 params.cell', !!d && d.params && d.params.cell === 'A1', JSON.stringify(d && d.params));
  const h = await routeIntentHandler({ task: '冻结首行' });
  const pl = JSON.parse(h.content[0].text);
  check('handler 走通：type=direct 且 confidence=0.95', pl.type === 'direct' && pl.confidence === 0.95,
    JSON.stringify({ type: pl.type, conf: pl.confidence }));
  check('直答回包带 nextStep 指向工具', pl.nextStep.includes(pl.result.tool), pl.nextStep.slice(0, 60));
}

console.log('\n=== 三层意图：任务包（导航层，≥10 条）===');
const pkgCases = [
  ['帮我写一份本周工作周报', 'weekly_report'],
  ['按公文格式排一下这个红头文件', 'official_document'],
  ['审阅这份合同并开修订模式', 'contract_review'],
  ['给毕业论文排个版', 'paper_format'],
  ['帮我把这些数据录入表格', 'data_entry'],
  ['把这份演示美化一下', 'demo_beautify'],
  ['把这组数据图表化', 'data_charts'],
  ['把全表的格式统一一下', 'format_unify'],
  ['整理今天的会议纪要', 'meeting_notes'],
  ['做一个项目进度表', 'project_tracker'],
  ['对两个月的数据做对比分析', 'compare_analysis'],
  ['按模板填字段生成合同', 'template_fill'],
];
for (const [text, taskName] of pkgCases) {
  const m = matchTaskPackage(text);
  check('任务包「' + text + '」→ ' + taskName,
    !!m.best && m.best.task === taskName && m.best.steps.length >= 2,
    m.best ? 'steps=' + m.best.steps.length + ' matched=' + m.best.matched.join(',') : '未命中');
}
{
  const m = matchTaskPackage('写周报');
  check('任务包 steps 每步带 app 路由', !!m.best && m.best.steps.every((x) => x.app), m.best ? m.best.steps.map((x) => x.app).join(',') : 'n/a');
  const h = await routeIntentHandler({ task: '帮我整理会议纪要' });
  const pl = JSON.parse(h.content[0].text);
  check('handler 走通：type=task_package 且 confidence=0.85',
    pl.type === 'task_package' && pl.confidence === 0.85, JSON.stringify({ type: pl.type, conf: pl.confidence }));
  check('confirm_before 在回包里', Array.isArray(pl.result.confirm_before), JSON.stringify(pl.result.confirm_before));
}

console.log('\n=== 三层意图：置信度 / 降级回归 ===');
{
  const h = await routeIntentHandler({ task: '随手写点什么' });
  const pl = JSON.parse(h.content[0].text);
  check('无匹配降级到 group 层（type=group）', pl.type === 'group', pl.type);
  check('低置信度（<0.5）时 alternatives 非空', pl.confidence < 0.5 && pl.alternatives.length > 0,
    'conf=' + pl.confidence + ' alts=' + pl.alternatives.length);
  check('定组原字段保留在 result（回归不断）', Array.isArray(pl.result.matches) && 'nextStep' in pl.result);
  const h2 = await routeIntentHandler({ task: '做一份产品发布会的 PPT' });
  const pl2 = JSON.parse(h2.content[0].text);
  check('高置信度定组：type=group 且不误入短路口/任务包',
    pl2.type === 'group' && pl2.confidence >= 0.5 && pl2.result.matches.some((m) => m.group === 'ppt'));
}

console.log('\n=== 工具合并：ping / wire_check → wps_check_connection ===');
check('别名解析：wps_common_ping → wps_check_connection', resolveToolAlias('wps_common_ping') === 'wps_check_connection');
check('别名解析：wps_common_wire_check → wps_check_connection', resolveToolAlias('wps_common_wire_check') === 'wps_check_connection');
check('规范名已注册（registry 可取到）', !!toolRegistry.getTool('wps_check_connection'));
check('旧名不在列表面（不出现在 tools/list）',
  !toolRegistry.listAllTools().some((t) => t.name === 'wps_common_wire_check' || t.name === 'wps_common_ping'));

const passed = results.filter((x) => x.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
}

main().catch((e) => {
  console.error('测试异常:', e && e.message ? e.message : e);
  process.exit(1);
});
