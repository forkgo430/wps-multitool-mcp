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
const { classifyIntent, routeIntentHandler } = require(path.join(ROOT, 'dist', 'tools', 'common', 'intent.js'));

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
check('返回 needLoad 含 ppt', payload.needLoad.includes('ppt'), payload.needLoad.join(','));
check('没有 autoLoaded', payload.autoLoaded.length === 0);

console.log('\n=== handler：autoLoad=true 会真的加载 ===');
out = await routeIntentHandler({ task: '做一份 PPT', autoLoad: true });
payload = JSON.parse(out.content[0].text);
const afterAuto = toolRegistry.getGroups().find((g) => g.group === 'ppt').fullyLoaded;
check('ppt 组已完整加载', afterAuto === true);
check('返回 autoLoaded 含 ppt', payload.autoLoaded.includes('ppt'), payload.autoLoaded.join(','));
check('nextStep 提示已刷新', /已自动加载/.test(payload.nextStep), payload.nextStep);

console.log('\n=== handler：没识别出来时指向 search_tools ===');
out = await routeIntentHandler({ task: '随手写点什么' });
payload = JSON.parse(out.content[0].text);
check('明确说没识别出来', payload.matches.length === 0 && !!payload.note);
check('nextStep 指向 wps_search_tools', /wps_search_tools/.test(payload.nextStep));

const passed = results.filter((x) => x.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
}

main().catch((e) => {
  console.error('测试异常:', e && e.message ? e.message : e);
  process.exit(1);
});
