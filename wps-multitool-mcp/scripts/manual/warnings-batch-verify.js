/**
 * warnings 语义 + 批量执行验证
 * ----------------------------
 * 验六件事：
 *  1) warnings 通道：PS 侧的 data.warnings 能提到顶层，且对脏数据容错
 *  2) 渲染：success=true 但有 warnings 时，正文前必须挂 [部分成功|...]，
 *     不挂的话模型会当全好了（它不会主动去看 warnings）
 *  3) 批量上限是 3（赢容定的）：给 4 步必须被拒，且拒绝文案要说清为什么
 *  4) 批量遇错不中断：第 1 步失败，第 3 步照样执行
 *  5) 部分成功语义：1 步失败 + 2 步成功 → success 仍为 true，但 warnings 里
 *     必须点名是哪一步
 *  6) 全部失败 → success=false
 *
 * 用法：node scripts/manual/warnings-batch-verify.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

const { extractWarnings } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { prefixWithWarnings } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { BATCH_MAX_STEPS } = require(path.join(ROOT, 'dist', 'tools', 'common', 'batch.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));

const _s = new WpsMcpServer();
_s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

console.log('\n=== A. warnings 解析容错 ===');
ok('对象带 step', JSON.stringify(extractWarnings({ warnings: [{ step: 'a.b', message: 'x' }] })) === '[{"step":"a.b","message":"x"}]');
ok('裸字符串', JSON.stringify(extractWarnings({ warnings: ['裸提示'] })) === '[{"message":"裸提示"}]');
ok('垃圾项被丢掉', JSON.stringify(extractWarnings({ warnings: [{}, null, '', { message: 'ok' }] })) === '[{"message":"ok"}]');
ok('非数组 → 空', extractWarnings({ warnings: 'x' }).length === 0);
ok('无 data → 空', extractWarnings(null).length === 0);
ok('data 不是对象 → 空', extractWarnings(42).length === 0);

console.log('\n=== B. 渲染 ===');
const withW = {
  id: '1', success: true,
  content: [{ type: 'text', text: '已建样式' }],
  warnings: [{ step: 'addStyle.属性', message: 'alignment 不支持' }],
};
const rendered = prefixWithWarnings(withW)[0].text;
ok('正文前有 [部分成功|…] 前缀', rendered.startsWith('[部分成功|'), rendered.split('\n')[0]);
ok('前缀里写清了处数', /1 处/.test(rendered));
ok('列出了具体项', rendered.includes('addStyle.属性：alignment 不支持'));
ok('原始正文仍在', rendered.includes('已建样式'));
ok('没有 warnings 时原样返回', prefixWithWarnings({ id: '1', success: true, content: [{ type: 'text', text: 'x' }] })[0].text === 'x');

console.log('\n=== C. 批量上限 ===');
ok('上限就是 3', BATCH_MAX_STEPS === 3, `BATCH_MAX_STEPS=${BATCH_MAX_STEPS}`);
const batchHandler = toolRegistry.getTool('wps_common_batch').handler;
ok('wps_common_batch 已注册', !!batchHandler);

console.log('\n=== D/E/F. 批量执行语义（用 stub 工具，不碰 WPS）===');
// 注册两个假工具：一个必成功、一个必失败、一个带 warnings
const reg = (def, handler) => toolRegistry.register(def, handler);
reg({ name: 'zz_fake_ok', description: '测试用', inputSchema: { type: 'object', properties: {} }, category: 'common' },
  async () => ({ id: 'a', success: true, content: [{ type: 'text', text: 'ok' }] }));
reg({ name: 'zz_fake_bad', description: '测试用', inputSchema: { type: 'object', properties: {} }, category: 'common' },
  async () => ({ id: 'b', success: false, content: [{ type: 'text', text: '炸了' }], error: '模拟失败' }));
reg({ name: 'zz_fake_warn', description: '测试用', inputSchema: { type: 'object', properties: {} }, category: 'common' },
  async () => ({ id: 'c', success: true, content: [{ type: 'text', text: '半成' }], warnings: [{ message: '属性没设上' }] }));

(async () => {
  // 超过上限
  const over = await batchHandler({ steps: [{}, {}, {}, {}] });
  ok('4 步被拒', over.success === false);
  ok('拒绝文案说明了上限与原因', /最多 3 步/.test(over.content[0].text));

  // 空 steps
  const empty = await batchHandler({ steps: [] });
  ok('空 steps 被拒', empty.success === false);

  // 3 步全成功
  const allOk = await batchHandler({ steps: [{ tool: 'zz_fake_ok' }, { tool: 'zz_fake_ok' }, { tool: 'zz_fake_ok' }] });
  ok('3 步全成功 → success=true', allOk.success === true);
  ok('3 步全成功 → 无 warnings', !allOk.warnings || allOk.warnings.length === 0);
  ok('汇总写清了 3/3', /3\/3 步成功/.test(allOk.content[0].text));

  // 中间失败不中断
  const midFail = await batchHandler({ steps: [{ tool: 'zz_fake_bad' }, { tool: 'zz_fake_ok' }, { tool: 'zz_fake_ok' }] });
  ok('中间失败不中断（后续仍执行）', /2\/3 步成功/.test(midFail.content[0].text), midFail.content[0].text.split('\n')[0]);
  ok('部分成功 → success 仍为 true', midFail.success === true);
  ok("warnings 点名了失败的那步", (midFail.warnings || []).some((w) => w.step === 'zz_fake_bad'), JSON.stringify(midFail.warnings));
  ok('失败原因带上了', midFail.warnings.some((w) => /模拟失败/.test(w.message)));

  // 带 warnings 的步骤会被提升
  const withWarnStep = await batchHandler({ steps: [{ tool: 'zz_fake_warn' }] });
  ok('步骤的 warnings 被提升到批次级', withWarnStep.warnings?.some((w) => /属性没设上/.test(w.message)));

  // 全失败
  const allBad = await batchHandler({ steps: [{ tool: 'zz_fake_bad' }, { tool: 'zz_fake_bad' }] });
  ok('全部失败 → success=false', allBad.success === false);
  ok('全失败文案提醒状态要自查', /自己确认/.test(allBad.content[0].text));

  // 工具不存在
  const missing = await batchHandler({ steps: [{ tool: 'zz_not_exist' }] });
  ok('不存在的工具算失败并给出线索', missing.warnings?.some((w) => /不存在/.test(w.message)));

  // 注解：batch 能包含任意写入 → destructive
  const ann = getToolAnnotations('wps_common_batch');
  ok('batch 注解：非只读', ann?.readOnlyHint === false);
  ok('batch 注解：破坏性（可含任意写入）', ann?.destructiveHint === true);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})();
