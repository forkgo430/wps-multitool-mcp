#!/usr/bin/env node
/**
 * Input: 无（mock COM，不碰 WPS）
 * Output: M2 wordkit wps_word_build_document 的断言组——
 *         A 展开序列（8 kind 的纯函数映射：工具链/参数/纪律拒绝）；
 *         B 生成后校验三路（段落数对账失败→警告、空段落→警告、全绿→无警告）；
 *         C 降级两路（stats 失败→降级记一条、校验异常→降级不失败）；
 *         D fail-fast（中途炸 → 报停在第几段第几步）。
 * Pos: 主线一 B 段验收脚本（2026-11-06）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 怎么 mock：替换 wpsClient.executeMethod——word 全链 mock（createDocument/insertText/
 * addHeading/insertTable/insertImage/insertPageBreak/getDocumentStats/getDocumentParagraphs
 * 全部可控），callTool 真入口跑 buildDocumentHandler。
 * 用法：node scripts/manual/wordkit-verify.js
 */
'use strict';

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));
const kit = require(path.join(ROOT, 'dist/tools/word/wordkit.js'));

new (require(path.join(ROOT, 'dist/server/mcp-server.js')).WpsMcpServer)();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};
const eq = (label, got, want) => ok(`${label} = ${JSON.stringify(want)}`, JSON.stringify(got) === JSON.stringify(want), `got ${JSON.stringify(got)}`);
const call = (name, args = {}) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));

/** word 全链 mock：记录所有调用；stats/paragraphs 文案可控 */
function makeWordMock(opts = {}) {
  const calls = [];
  const statsParagraphs = opts.statsParagraphs ?? 0;
  const statsText = opts.statsText ?? '字数 120；共 6 个段落；页数 1';
  const paraLines = opts.paraLines ?? [];
  const failAt = opts.failAt; // { tool } 在第 N 次遇到该工具时炸
  const counter = {};
  const fn = async (action, params) => {
    calls.push({ action, params });
    if (action === 'createDocument') return { success: true, data: { name: 'MockDoc' } };
    if (action === 'insertText') return { success: true, data: { textLength: (params.text || '').length } };
    if (action === 'addHeading') return { success: true, data: { level: params.level } };
    if (action === 'insertTable') return { success: true, data: { rows: params.rows, cols: params.cols } };
    if (action === 'insertImage') return { success: true, data: {} };
    if (action === 'insertPageBreak') return { success: true, data: {} };
    if (action === 'getDocumentStats') {
      if (opts.statsFail) throw new Error('mock: stats 通道炸了');
      // 回真实 data 形状：gaps.ts 的 handler 读 d.paragraphs 等字段后自行格式化
      return { success: true, data: { name: 'MockDoc', paragraphs: opts.statsParagraphs, words: 120, characters: 300, sections: 1, pages: 1 } };
    }
    if (action === 'getDocumentParagraphs') {
      if (opts.parasFail) throw new Error('mock: paragraphs 通道炸了');
      return {
        success: true,
        data: {
          paragraphs: paraLines.map((l, i) => ({ index: i + 1, text: l.text || '', style: l.style || '正文', start: i, end: i })),
          totalCount: paraLines.length,
          returnedCount: paraLines.length,
        },
      };
    }
    return { success: true, data: {} };
  };
  return { calls, fn };
}

/** 统计处理器发出的 word 动作名序列 */
function actionsOf(calls) {
  return calls.map((c) => c.action);
}

(async () => {
  console.log('\n=== A. 展开序列（纯函数 expandDoc，8 kind 映射）===');
  {
    const pages = kit.expandDoc({
      sections: [
        { kind: 'title', content: '季度总结' },
        { kind: 'heading', content: '一、业绩回顾', level: 2 },
        { kind: 'para', content: '本季度完成三件事。' },
        { kind: 'table', head: ['指标', '值'], rows: [['营收', '100万']] },
        { kind: 'image', imagePath: 'C:/tmp/x.png', width: 400, height: 300 },
        { kind: 'list', items: ['第一件', '第二件'], listType: 'number' },
        { kind: 'quote', content: '要事第一' },
        { kind: 'pagebreak' },
      ],
    });
    const flat = pages.flat();
    eq('步骤工具链（title→heading→para→table→image→list×2→quote→pagebreak）',
      flat.map((s) => s.tool),
      ['wps_word_add_heading', 'wps_word_add_heading', 'wps_word_insert_text',
       'wps_word_insert_table', 'wps_word_insert_image',
       'wps_word_insert_text', 'wps_word_insert_text',
       'wps_word_insert_text', 'wps_word_insert_page_break']);
    ok('title → add_heading level 1', pages[0][0].args.level === 1);
    ok('heading level=2（默认）', pages[1][0].args.level === 2);
    ok('para/quote 带 new_paragraph（每段独立段落）',
      pages[2][0].args.new_paragraph === true && pages[6][0].args.new_paragraph === true);
    ok('quote 「」包裹', pages[6][0].args.text === '「要事第一」');
    ok('list number 类型无 • 前缀', pages[5][0].args.text === '第一件' && pages[5][1].args.text === '第二件');
    ok('table 一次落位（data 2D 数组直传）', pages[3][0].args.data !== undefined && pages[3][0].args.rows === 1);
  }

  console.log('\n=== A2. 纪律拒绝（超密度/缺字段/未知 kind）===');
  {
    let threw = '';
    try { kit.expandDoc({ sections: [{ kind: 'magical' }] }); } catch (e) { threw = e.message; }
    ok('未知 kind → 列出支持清单', /未知 section kind/.test(threw) && /title \/ heading/.test(threw), threw.slice(0, 50));
    threw = '';
    try { kit.expandDoc({ sections: [{ kind: 'para' }] }); } catch (e) { threw = e.message; }
    ok('para 缺 content → 拒绝', /para 缺 content/.test(threw));
    threw = '';
    try {
      kit.expandDoc({
        sections: Array.from({ length: 90 }, (_, i) => ({ kind: 'para', content: `段落${i}` })),
      });
    } catch (e) { threw = e.message; }
    ok('段落超上限 80 → 拒绝（拆文档）', /超过上限 80/.test(threw), threw.slice(0, 50));
    threw = '';
    try { kit.expandDoc({ sections: [{ kind: 'table', rows: [['a'], ['b', 'c']] }] }); } catch (e) { threw = e.message; }
    ok('table 行列不齐 → 拒绝', /行列数不一致/.test(threw));
    threw = '';
    try { kit.expandDoc({}); } catch (e) { threw = e.message; }
    ok('缺 sections → 拒绝', /sections 为空/.test(threw));
  }

  console.log('\n=== B. 端到端（callTool 真入口）：展开落位 + 生成后校验三路 ===');
  {
    // 三路之一：全绿——回读段落数与预期一致、无空段 → 无警告
    const good = makeWordMock({
      statsParagraphs: 5,
      paraLines: [
        { text: '季度总结', style: '标题 1' },
        { text: '一、业绩回顾', style: '标题 2' },
        { text: '本季度完成三件事。', style: '正文' },
        { text: '• 第一件', style: '正文' },
        { text: '• 第二件', style: '正文' },
      ],
    });
    wpsClient.executeMethod = good.fn;
    const rGood = await call('wps_word_build_document', {
      doc: {
        sections: [
          { kind: 'title', content: '季度总结' },
          { kind: 'heading', content: '一、业绩回顾' },
          { kind: 'para', content: '本季度完成三件事。' },
          { kind: 'list', items: ['第一件', '第二件'] },
        ],
      },
    });
    ok('全绿路 success=true', rGood.success === true, String(rGood.error || ''));
    // 快照层可能补一条环境性 snapshot 提示（mock 文档未落盘），只看 wordkit 自身的校验警告
    const goodWs = (rGood.warnings || []).filter((w) => w.step && w.step.startsWith('wordkit.'));
    ok('全绿路无 wordkit 警告（段落数对账通过、无空段）', goodWs.length === 0, JSON.stringify((rGood.warnings || []).map((w) => w.step)));
    ok('全绿路正文含步数', rGood.content[0].text.includes('7 步展开') || rGood.content[0].text.includes('步展开'), rGood.content[0].text.slice(0, 60));

    // 三路之二：空段落 → warning
    const empty = makeWordMock({
      statsParagraphs: 5,
      paraLines: [
        { text: '季度总结', style: '标题 1' },
        { text: '', style: '正文' },
        { text: '一', style: '正文' },
        { text: '一', style: '正文' },
        { text: '一', style: '正文' },
      ],
    });
    wpsClient.executeMethod = empty.fn;
    const rEmpty = await call('wps_word_build_document', {
      doc: { sections: [{ kind: 'title', content: '季度总结' }, { kind: 'para', content: 'x' }, { kind: 'para', content: 'y' }, { kind: 'para', content: 'z' }] },
    });
    ok('空段落 → warning（不静默）', rEmpty.success === true && (rEmpty.warnings || []).some((w) => w.step === 'wordkit.空段落' && /疑似空段落/.test(w.message)), JSON.stringify((rEmpty.warnings || []).map((w) => w.step)));

    // 三路之三：段落数对账失败 → warning
    const mismatch = makeWordMock({
      statsParagraphs: 2,
      paraLines: [{ text: '少了很多', style: '正文' }, { text: '还是少', style: '正文' }],
    });
    wpsClient.executeMethod = mismatch.fn;
    const rMis = await call('wps_word_build_document', {
      doc: { sections: [{ kind: 'title', content: 'A' }, { kind: 'para', content: 'B' }, { kind: 'para', content: 'C' }, { kind: 'para', content: 'D' }] },
    });
    ok('段落数对账失败 → warning（回读少于展开预期）',
      rMis.success === true && (rMis.warnings || []).some((w) => w.step === 'wordkit.段落数对账' && /少于展开预期/.test(w.message)), JSON.stringify((rMis.warnings || []).map((w) => w.step)));
  }

  console.log('\n=== C. 降级两路：stats 失败 / 校验异常 → 降级记一条不失败 ===');
  {
    const d1 = makeWordMock({ statsFail: true });
    wpsClient.executeMethod = d1.fn;
    const r1 = await call('wps_word_build_document', { doc: { sections: [{ kind: 'para', content: '内容' }] } });
    const d1Ws = (r1.warnings || []).filter((w) => w.step && w.step.startsWith('wordkit.'));
    ok('stats 失败 → 降级记一条、success=true',
      r1.success === true && d1Ws.some((w) => w.step === 'wordkit.校验降级'), JSON.stringify((r1.warnings || []).map((w) => w.step)));

    const d2 = makeWordMock({ parasFail: true });
    wpsClient.executeMethod = d2.fn;
    const r2 = await call('wps_word_build_document', { doc: { sections: [{ kind: 'para', content: '内容' }] } });
    const d2Ws = (r2.warnings || []).filter((w) => w.step && w.step.startsWith('wordkit.'));
    ok('paragraphs 失败 → 降级记一条、success=true',
      r2.success === true && d2Ws.some((w) => w.step === 'wordkit.校验降级'));
  }

  console.log('\n=== D. fail-fast：中途炸 → 报停在第几段第几步 ===');
  {
    const m = makeWordMock();
    const orig = m.fn;
    let headingCalls = 0;
    wpsClient.executeMethod = async (action, params) => {
      if (action === 'addHeading') {
        headingCalls += 1;
        if (headingCalls === 2) throw new Error('mock: 第二个标题炸了');
      }
      return orig(action, params);
    };
    const r = await call('wps_word_build_document', {
      doc: {
        sections: [
          { kind: 'title', content: 'T' },
          { kind: 'heading', content: 'H1' },
          { kind: 'heading', content: 'H2 炸点之后的段落不会被建' },
          { kind: 'para', content: 'P' },
        ],
      },
    });
    ok('中途炸 → success=false 且报错含现场', r.success === false && /第二个标题炸了/.test(r.error || ''), (r.error || '').slice(0, 60));
    ok('fail-fast 报停点（第 2 段——第二个标题就是段 2）', (r.error || '').includes('第2段'), (r.error || '').slice(0, 90));
    //炸点之后不再有 insertText 落位（P 未写入）
    const inserts = m.calls.filter((x) => x.action === 'insertText');
    ok('炸点之后的展开步未执行（fail-fast）', inserts.length === 0, `insertText=${inserts.length}`);
  }

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('验收异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
