#!/usr/bin/env node
/**
 * Input: 无（只读 dist/ 与源码）
 * Output: word target 参数链路的静态自检结果（不碰 WPS、不起桥）
 * Pos: 第 1 期第 3 章轮 1 的自检脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 验的是"参数有没有被正确理解与路由"，不是"WPS 有没有真的改对文档"。
 * 后者必须真机（A1–A10），不归本脚本。
 *
 * 手法：把 wpsClient.invokeAction 换掉，截获 executeMethod 最终要发出去的东西：
 *   - params 里有没有 documentName（doc 合流）
 *   - opts.instance 是什么（路由键）
 * 这样连 ALS 继承、schema 校验、错误文案都一并覆盖，且全程零 COM。
 */

'use strict';

const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { runWithTarget } = require(path.join(ROOT, 'dist/server/target-context.js'));

// ---- mock：只换 invokeAction（executeMethod 仍是真代码）----
/** @type {Array<{action:string, params:Record<string,unknown>, app:string, instance:string}>} */
const calls = [];
wpsClient.invokeAction = async (action, params, app, opts) => {
  calls.push({ action, params: params || {}, app: app || '', instance: (opts && opts.instance) || 'user' });
  return { success: true, data: {} };
};

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (name, args) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));
const last = () => calls[calls.length - 1];

let pass = 0;
let fail = 0;
const ok = (label, cond, detail) => {
  if (cond) pass++;
  else fail++;
  console.log(`  ${cond ? '✓' : '✗'} ${label}${detail ? ' — ' + String(detail).slice(0, 150) : ''}`);
};

(async () => {
  console.log('=== 1. 不传 target：旧行为必须一字不变 ===');
  calls.length = 0;
  let r = await call('wps_word_insert_text', { text: 'hello' });
  ok('调用成功', r.success === true, r.error);
  ok('params 里没有 documentName（没有凭空注入）', last().params.documentName === undefined, JSON.stringify(last().params));
  ok('params 里没有 target', last().params.target === undefined, JSON.stringify(Object.keys(last().params)));
  ok('路由到 user', last().instance === 'user', last().instance);

  console.log('\n=== 2. target.doc 合流进 documentName ===');
  calls.length = 0;
  r = await call('wps_word_insert_text', { text: 'hello', target: { doc: '报告.docx' } });
  ok('调用成功', r.success === true, r.error);
  ok('documentName = 报告.docx', last().params.documentName === '报告.docx', JSON.stringify(last().params));
  ok('没有新造第二个字段（params 里没有 docName/targetDoc 之类）',
    !Object.keys(last().params).some((k) => /^(docName|targetDoc|doc)$/i.test(k)),
    JSON.stringify(Object.keys(last().params)));

  console.log('\n=== 3. target.doc 优先级最高（覆盖会话级 documentName）===');
  // 这条曾经写成"已有 documentName 就不覆盖"，真机 A1/A3 立刻暴露它是错的：
  // executeMethod 会把 switch_document / create_document 记下的会话文档注入 documentName，
  // 于是每次调用的 target.doc 被静默忽略 —— 所有写动作都落到同一份，还不报错。
  // 现在的契约：**每次调用的 target.doc 赢**（会话级只是兜底）。这条不能用工具调用验
  // （handler 不转发 documentName），所以直接验规则函数。
  {
    const { applyWordTarget } = require(path.join(ROOT, 'dist/tools/word/target.js'));
    const a = applyWordTarget({ documentName: 'B.docx', target: { doc: 'A.docx' } });
    ok('target.doc 覆盖已有的 documentName', a.params.documentName === 'A.docx', a.params.documentName);
    const b = applyWordTarget({ target: { doc: 'A.docx' } });
    ok('没有 documentName 时补上', b.params.documentName === 'A.docx', b.params.documentName);
    const c = applyWordTarget({ documentName: 'B.docx' });
    ok('没有 target 时不动 documentName', c.params.documentName === 'B.docx', c.params.documentName);
  }

  console.log('\n=== 4. instance 走路由，不混进 COM 参数语义 ===');
  calls.length = 0;
  r = await call('wps_word_insert_text', { text: 'h', target: { instance: 'sandbox-2' } });
  ok('调用成功', r.success === true, r.error);
  ok('instance=sandbox-2', last().instance === 'sandbox-2', last().instance);
  ok('target 仍然随 params 下发（PS 侧要靠它拒绝选区类动作）', last().params.target !== undefined,
    JSON.stringify(last().params.target));

  calls.length = 0;
  r = await call('wps_word_insert_text', { text: 'h', target: { instance: 'sandbox', doc: 'A.docx' } });
  ok('sandbox 无后缀原样下发（由桥挑第一个空闲的）', last().instance === 'sandbox', last().instance);
  ok('doc 同时合流', last().params.documentName === 'A.docx', last().params.documentName);

  console.log('\n=== 5. ALS 继承：内层 handler 不转发 target，也要能钉住 ===');
  calls.length = 0;
  r = await runWithTarget({ instance: 'sandbox-1', doc: 'C.docx' }, () =>
    call('wps_word_insert_text', { text: 'h' })
  );
  ok('调用成功', r.success === true, r.error);
  ok('继承外层 instance=sandbox-1', last().instance === 'sandbox-1', last().instance);
  ok('继承外层 doc=C.docx', last().params.documentName === 'C.docx', last().params.documentName);

  console.log('\n=== 6. 非法输入必须报错（不许静默回落） ===');
  const expectFail = async (label, args, needle) => {
    const res = await call('wps_word_insert_text', args);
    const msg = String(res.error || '');
    ok(label, res.success === false && msg.includes(needle), msg || '(没有报错)');
  };
  await expectFail('sandbox-9 超范围', { text: 'h', target: { instance: 'sandbox-9' } }, '超出范围');
  await expectFail('不认识的 instance', { text: 'h', target: { instance: 'pool-1' } }, '不认识的 target.instance');
  await expectFail('target 不是对象', { text: 'h', target: 'sandbox' }, 'target 必须是对象');
  await expectFail('doc 不是字符串', { text: 'h', target: { doc: 123 } }, 'target.doc 必须是字符串');

  console.log('\n=== 7. excel / ppt 传 target：明确拒绝（本轮只做 word）===');
  {
    // 选 set_zoom：它只要求 percent，能真正走到 executeMethod。
    // （先前用 get_cell_value 时被 "Missing required parameter: sheet" 先拦住了，
    //   那说明测的不是 target 这条路 —— 换工具，不是改断言。）
    const res = await call('wps_excel_set_zoom', { percent: 100, target: { instance: 'sandbox' } });
    ok('excel 传 target → 报"只支持 word"', res.success === false && /只支持 word/.test(String(res.error)),
      String(res.error));
  }

  console.log('\n=== 8. schema：67 个 word 工具都声明了 target ===');
  {
    const wordTools = toolRegistry.listAllTools().filter((t) => /^wps_word_/.test(t.name));
    const missing = wordTools.filter((t) => !t.inputSchema?.properties?.target);
    ok(`word 工具 ${wordTools.length} 个全部声明 target`, missing.length === 0,
      missing.map((t) => t.name).join(','));
    const t0 = wordTools[0];
    const prop = t0.inputSchema.properties.target;
    ok('target 是 object 且有 instance/doc 两个子字段',
      prop && prop.type === 'object' && !!prop.properties.instance && !!prop.properties.doc,
      JSON.stringify(prop && Object.keys(prop.properties || {})));
    // 不传 target 时旧调用不该被 required 卡住
    ok('target 不在任何 required 里', wordTools.every((t) => !(t.inputSchema.required || []).includes('target')));
  }

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  process.exit(fail ? 1 : 0);
})().catch((e) => {
  console.error('脚本自身异常：', e && e.stack ? e.stack : e);
  process.exit(1);
});
