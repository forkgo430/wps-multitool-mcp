/**
 * 任务引擎 mock 单测（意图一期 · 服务端执行，2026-11-05 决策者）
 * 原则：不起真 WPS、不依赖 task-packages.json——包与执行器全部经 deps 注入
 *       （引擎的依赖注入设计就是为这一刻准备的）。
 * 覆盖：
 *   1) resolveValue：{{args.x}} / {{steps.N.data.x}} / {{steps.N.text}} / 整串引用保类型 /
 *      内嵌替换 / 缺引用原样返回 / 深遍历对象数组
 *   2) briefOf：base64 图片压缩为占位符（引擎存在的意义之一）/ 120 字符截断
 *   3) happy path：调用序列 + 参数已解析（含跨步引用）+ briefs + 进度回调 N/M
 *   4) confirm 闸门：命中 confirm_before 未带 confirm → 停在前面 needConfirm；带上 → 跑完
 *   5) fail-fast：单步失败立即停，failedStep/已完成清单/错误原文
 *   6) 未知包：明确报错不装
 * 跑法：node scripts/manual/task-engine-test.js（verify-all FAST 档包含它）。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const engine = require(path.join(ROOT, 'dist', 'server', 'task-engine.js'));
const { runTaskPackage, resolveValue, briefOf } = engine;
const { runWithProgress } = require(path.join(ROOT, 'dist', 'server', 'progress-context.js'));

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};
const eq = (label, got, want) => ok(`${label} = ${JSON.stringify(want)}`, JSON.stringify(got) === JSON.stringify(want), `got ${JSON.stringify(got)}`);

const mkCtx = () => ({
  args: { body: '本周完成X', n: 5 },
  steps: new Map([
    [1, { data: { path: 'C:/a.png', rows: 5, range: 'A1:B5' } }],
    [2, { text: 'hello', data: { path: 'C:/chart.png' } }],
  ]),
});

/* ---- 1. resolveValue：参数模板解析 ---- */
console.log('\n=== 参数模板解析（resolveValue）===');
const ctx = mkCtx();
eq('整串引用 {{args.body}} 保字符串', resolveValue('{{args.body}}', ctx), '本周完成X');
eq('整串引用 {{args.n}} 保数字类型', resolveValue('{{args.n}}', ctx), 5);
eq('steps.N.data.x 深路径', resolveValue('{{steps.1.data.path}}', ctx), 'C:/a.png');
eq('steps.N.data.x 第二步', resolveValue('{{steps.2.data.path}}', ctx), 'C:/chart.png');
eq('steps.N.text 原文', resolveValue('{{steps.2.text}}', ctx), 'hello');
eq('steps.N.x 简写视同 data.x', resolveValue('{{steps.1.rows}}', ctx), 5);
ok('缺引用原样返回（不造假不抛错）', resolveValue('{{steps.1.nothere}}', ctx) === '{{steps.1.nothere}}');
eq('内嵌替换（部分引用）', resolveValue('共 {{args.n}} 行', ctx), '共 5 行');
const deep = resolveValue({ a: ['{{args.body}}', { b: '{{steps.2.data.path}}' }] }, ctx);
eq('深遍历对象数组', deep, { a: ['本周完成X', { b: 'C:/chart.png' }] });

/* ---- 2. briefOf：base64 压缩 + 截断 ---- */
console.log('\n=== 回包压缩（briefOf）===');
const hugeB64 = 'data:image/png;base64,' + 'A'.repeat(5000);
const b1 = briefOf({ id: 'x', success: true, content: [{ type: 'text', text: `图好了 ${hugeB64} 完毕` }] });
ok('base64 压缩为占位符', b1.includes('[inline-image]') && !b1.includes('AAAA'), b1);
const long = 'x'.repeat(300);
const b2 = briefOf({ id: 'x', success: true, content: [{ type: 'text', text: long }] });
ok('超 120 字符截断', b2.length <= 120 && b2.endsWith('…'), `len=${b2.length}`);
const b3 = briefOf({ id: 'x', success: false, error: '炸了', content: [] });
ok('无文本回包退回 error', b3 === '炸了');

/* ---- 3~6（异步段：happy path / confirm / fail-fast / 未知包）---- */
(async () => {
console.log('\n=== happy path（注入执行器，3 步链）===');
const pkg1 = {
  task: 'fake_report',
  title: '假周报',
  match: [],
  execute: 'server',
  confirm_before: [],
  steps: [
    { action: 'read', tool: 'fake_read', app: 'excel', params: {}, note: '读数据' },
    { action: 'chart', tool: 'fake_chart', app: 'excel', params: { source: '{{steps.1.data.range}}' }, note: '成图' },
    { action: 'write', tool: 'fake_write', app: 'word', params: { text: '共 {{steps.1.data.rows}} 行，图在 {{steps.2.data.path}}' }, note: '写入' },
  ],
};
const calls = [];
const progress = [];
const deps1 = {
  callTool: async (name, args) => {
    calls.push({ name, args });
    if (name === 'fake_read') return { id: 'x', success: true, content: [{ type: 'text', text: JSON.stringify({ data: { range: 'A1:B5', rows: 5 } }) }] };
    if (name === 'fake_chart') return { id: 'x', success: true, content: [{ type: 'text', text: JSON.stringify({ data: { path: 'C:/chart.png' } }) }] };
    if (name === 'fake_write') return { id: 'x', success: true, content: [{ type: 'text', text: JSON.stringify({ data: { saved: true } }) }] };
    return { id: 'x', success: false, error: 'unknown tool' };
  },
  findPackage: () => pkg1,
};
const r1 = await runWithProgress(
  (p, t, m) => progress.push([p, t, m]),
  () => runTaskPackage('fake_report', { body: '正文' }, {}, deps1)
);
ok('整包成功', r1.success === true, `executed=${r1.executed}`);
eq('执行步数', r1.executed, 3);
eq('调用序列', calls.map((c) => c.name), ['fake_read', 'fake_chart', 'fake_write']);
eq('第2步参数已解析跨步引用', calls[1].args.source, 'A1:B5');
eq('第3步参数混合引用（数据+文字）', calls[2].args.text, '共 5 行，图在 C:/chart.png');
ok('最终产物 = 最后一步 data', r1.data && r1.data.saved === true, JSON.stringify(r1.data));
eq('进度回调 N/M', progress.map((p) => p[0]), [1, 2, 3]);
eq('进度总数', progress[0][1], 3);
ok('briefs 全 ok', r1.steps.every((s) => s.ok));

/* ---- 4. confirm 闸门 ---- */
console.log('\n=== confirm 闸门（破坏性动作先征询）===');
const pkg2 = {
  task: 'fake_confirm',
  title: '假确认包',
  match: [],
  execute: 'server',
  confirm_before: ['fake_doc'],
  steps: [
    { action: 'read', tool: 'fake_read', app: 'excel', params: {} },
    { action: 'fake_doc', tool: 'fake_doc', app: 'word', params: {} },
    { action: 'write', tool: 'fake_write', app: 'word', params: {} },
  ],
};
const calls2 = [];
const deps2 = {
  callTool: async (name) => { calls2.push(name); return { id: 'x', success: true, content: [{ type: 'text', text: 'ok' }] }; },
  findPackage: () => pkg2,
};
const r2 = await runTaskPackage('fake_confirm', {}, {}, deps2);
ok('未带 confirm：停在破坏性步骤前', r2.needConfirm === true && r2.confirmAction === 'fake_doc');
eq('未带 confirm：只执行了第 1 步', r2.executed, 1);
eq('未带 confirm：只调了 1 次', calls2.length, 1);
const r3 = await runTaskPackage('fake_confirm', {}, { confirm: true }, deps2);
ok('带 confirm：跑完整包', r3.success === true && r3.executed === 3, `executed=${r3.executed}`);

/* ---- 5. fail-fast ---- */
console.log('\n=== fail-fast（失败即停，绝不重放）===');
const pkg3 = {
  task: 'fake_fail',
  title: '假失败包',
  match: [],
  execute: 'server',
  confirm_before: [],
  steps: [
    { action: 'read', tool: 'fake_read', app: 'excel', params: {} },
    { action: 'boom', tool: 'fake_boom', app: 'excel', params: {} },
    { action: 'write', tool: 'fake_write', app: 'word', params: {} },
  ],
};
const calls3 = [];
const deps3 = {
  callTool: async (name) => {
    calls3.push(name);
    if (name === 'fake_read') return { id: 'x', success: true, content: [{ type: 'text', text: 'ok' }] };
    return { id: 'x', success: false, error: '模拟失败' };
  },
  findPackage: () => pkg3,
};
const r4 = await runTaskPackage('fake_fail', {}, {}, deps3);
ok('失败即停', r4.success === false && r4.failedStep === 2);
eq('失败后不再执行后续步骤', calls3.length, 2);
ok('已完成步骤如实列出', r4.steps.length === 2 && r4.steps[0].ok === true && r4.steps[1].ok === false);
ok('错误带步号与原文', (r4.error || '').includes('第 2 步') && (r4.error || '').includes('模拟失败'));

/* ---- 6. 未知包 ---- */
console.log('\n=== 未知包（明确报错不装）===');
const r5 = await runTaskPackage('no_such_task', {}, {}, { callTool: async () => ({ id: 'x', success: true, content: [] }) });
ok('未知包明确报错', r5.success === false && (r5.error || '').includes('没有找到') && (r5.error || '').includes('server'));

/* ---- 收官 ---- */
console.log(`\n═══ 结果：${pass} 通过 / ${fail} 失败 ═══`);
process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('测试框架异常:', e); process.exit(1); });
