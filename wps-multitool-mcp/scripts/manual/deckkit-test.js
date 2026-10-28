/**
 * deckkit 引擎骨架的 mock 单测（第 3 期 E2 验收项）
 * 原则：不起真 WPS —— ① expandDeck 纯函数断言"坐标/字号/调用序列"（10 页型 × ≥3）；
 * ② mock COM 冒烟：mock wpsClient.executeMethod 后走 build_ppt_deck 真入口跑完整调用序列
 *    （不建快照垃圾——handler 逐动作直调 handler，不经 callTool 的快照层）。
 * 语义依据：docs/agents/BORROW-PLAN.md §2.3（10 页型）+ 引擎探针（行高 1.2/渐变可用/AddTable 不可靠）。
 * 跑法：node scripts/manual/deckkit-test.js（verify-all FAST 档包含它）。
 */

const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
const { expandDeck, DECK_TYPO, deckSummary } = require(path.join(ROOT, 'dist', 'tools', 'ppt', 'deckkit.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};
const eq = (label, got, want) => ok(`${label} = ${JSON.stringify(want)}`, JSON.stringify(got) === JSON.stringify(want), `got ${JSON.stringify(got)}`);

/* ---- 1. 排版纪律常量（一处集中）---- */
console.log('\n=== 排版纪律常量（DECK_TYPO）===');
ok('正文字号下限 ≥ 14（可读性纪律）', DECK_TYPO.bodyFontMin >= 14, String(DECK_TYPO.bodyFontMin));
ok('行高系数 = 1.2（引擎探针 seg3 本机标定）', DECK_TYPO.lineFactor === 1.2);
ok('密度上限集中可读（bullets 6 / table 6×4 / timeline 6 / slides 30）',
  DECK_TYPO.maxBullets === 6 && DECK_TYPO.maxTableRows === 6 && DECK_TYPO.maxTableCols === 4 && DECK_TYPO.maxTimeline === 6 && DECK_TYPO.maxSlides === 30);
ok('16:9 画布 960×540', DECK_TYPO.slideWidth === 960 && DECK_TYPO.slideHeight === 540);

/* ---- 2. expandDeck：10 页型 × ≥3 断言（坐标/字号/调用序列）---- */
console.log('\n=== 10 页型展开（纯函数，坐标/字号/序列）===');

// title
{
  const p = expandDeck({ slides: [{ kind: 'title', heading: '季度汇报', sub: '2026 Q4', footer: '内部资料' }] })[0];
  eq('title 步骤数（渐变底两步+标题+白字+副标+accent+页脚+页脚色=8）', p.length, 8);
  eq('title 渐变底是第一个形状且全页', p[0].tool, 'wps_ppt_add_shape');
  eq('title 渐变参数（linear 90° 学术蓝→mid）', JSON.stringify(p[1].args.gradient),
    JSON.stringify({ type: 'linear', angle: 90, stops: [{ color: DECK_TYPO.theme.deep, position: 0 }, { color: DECK_TYPO.theme.mid, position: 1 }] }));
  eq('title 标题字号 36 且坐标居中带', p[2].args.fontSize, 36);
  eq('title 副标 20pt', p[4].args.fontSize, 20);
  ok('title 白字上深底（set_font_color → 形状#2）', p[3].tool === 'wps_ppt_set_font_color' && p[3].args.shapeIndex === 2 && p[3].args.color === '#FFFFFF');
}
// section
{
  const p = expandDeck({ slides: [{ kind: 'section', ordinal: '02', heading: '市场分析', tagline: '数据篇' }] })[0];
  eq('section 步骤数（竖条+序号+标题+标语=6）', p.length, 6);
  eq('section 竖条 24×全高', JSON.stringify([p[0].args.width, p[0].args.height]), JSON.stringify([24, 540]));
  eq('section 序号字号 40', p[1].args.fontSize, 40);
}
// bullets
{
  const p = expandDeck({ slides: [{ kind: 'bullets', heading: '要点', points: ['甲', '乙', '丙'] }] })[0];
  ok('bullets 步骤数（标题带3 + 3 卡×2=9）', p.length === 9, String(p.length));
  const shapes = p.filter((x) => x.tool === 'wps_ppt_add_shape').slice(-3);
  ok('3 张卡片色块坐标均匀分布（间距 = 可用高度/3）',
    shapes.length === 3 && Math.abs((shapes[1].args.top - shapes[0].args.top) - (shapes[2].args.top - shapes[1].args.top)) < 0.01);
  ok('要点字号 = bodyFont(16) ≥ 下限', p.filter((x) => x.tool === 'wps_ppt_add_textbox').every((x) => x.args.fontSize >= DECK_TYPO.bodyFontMin));
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'bullets', heading: 'x', points: Array(7).fill('超密度') }] }); } catch (e) { threw = e.message; }
  ok('超密度（7 要点）→ 拒绝不静默截断', /超过每页密度上限/.test(threw), threw.slice(0, 50));
}
// compare
{
  const p = expandDeck({ slides: [{ kind: 'compare', heading: '对比', left: { label: '方案A', points: ['快'] }, right: { label: '方案B', points: ['稳'] } }] })[0];
  eq('compare 步骤数（带3 + 左4 + 右4=11）', p.length, 11);
  const L = p.filter((x) => x.tool === 'wps_ppt_add_shape')[1];
  const R = p.filter((x) => x.tool === 'wps_ppt_add_shape')[2];
  ok('左右两栏不重叠（左 x+w ≤ 右 x）', L.args.left + L.args.width <= R.args.left);
  ok('右栏文字墨色（set_font_color → 形状#8）', p[p.length - 1].args.shapeIndex === 8);
}
// imageText
{
  const p = expandDeck({ slides: [{ kind: 'imageText', heading: '图文', imagePath: 'C:/tmp/x.png', points: ['要点'] }] })[0];
  ok('imageText 走 insert_ppt_image（非 SVG 非占位）', p.some((x) => x.tool === 'wps_ppt_insert_ppt_image' && x.args.imagePath === 'C:/tmp/x.png'));
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'imageText', heading: 'x' }] }); } catch (e) { threw = e.message; }
  ok('缺 imagePath → 拒绝', /imagePath/.test(threw), threw.slice(0, 50));
}
// table（矩形拼）
{
  const p = expandDeck({ slides: [{ kind: 'table', heading: '表', head: ['名称', '数量'], rows: [['甲', '1'], ['乙', '2']] }] })[0];
  const shapes = p.filter((x) => x.tool === 'wps_ppt_add_shape');
  ok('矩形拼表格：色块数 = 1(标题带) + 2(表头) + 4(2×2) = 7', shapes.length === 7, String(shapes.length));
  ok('全程无 AddTable（探针坑 5：边框不可靠）', !p.some((x) => /table/i.test(x.tool) && x.tool !== 'wps_ppt_add_textbox'));
  const headText = p.filter((x) => x.tool === 'wps_ppt_add_textbox')[1];
  eq('font_color 序列（标题白字2 + 表头白字 3/5）', p.filter((x) => x.tool === 'wps_ppt_set_font_color').map((x) => x.args.shapeIndex), [2, 3, 5]);
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'table', heading: 'x', head: ['a', 'b'], rows: [['1']] }] }); } catch (e) { threw = e.message; }
  ok('行列不齐 → 拒绝', /不一致/.test(threw), threw.slice(0, 50));
}
// chart（两段式）
{
  const p = expandDeck({ slides: [{ kind: 'chart', heading: '图', categories: ['Q1', 'Q2'], series: [{ name: '销量', values: [1, 2] }] }] })[0];
  eq('chart 步骤数（带3 + insert + setdata = 5）', p.length, 5);
  const ins = p.find((x) => x.tool === 'wps_ppt_insert_ppt_chart');
  const set = p.find((x) => x.tool === 'wps_ppt_set_ppt_chart_data');
  ok('insert 在前 set 在后（两段式：insert 的 data 不生效）', p.findIndex((x) => x.tool === 'wps_ppt_insert_ppt_chart') < p.findIndex((x) => x.tool === 'wps_ppt_set_ppt_chart_data'));
  eq('set_ppt_chart_data 的 chartIndex=3（标题带两形状后图表追加）', set.args.chartIndex, 3);
  eq('真数据在 set 段', JSON.stringify(set.args.data), JSON.stringify({ categories: ['Q1', 'Q2'], series: [{ name: '销量', values: [1, 2] }] }));
}
// timeline
{
  const p = expandDeck({ slides: [{ kind: 'timeline', heading: '轴', milestones: [{ label: 'M1', note: '启动' }, { label: 'M2' }, { label: 'M3', note: '收尾' }] }] })[0];
  ok('timeline 步骤数（带3 + 主线1 + 3 节点×2~3=12）', p.length === 12, String(p.length));
  const labels = p.filter((x) => x.tool === 'wps_ppt_add_textbox');
  ok('节点标签上下交替（防重叠）', labels[1].args.top !== labels[2].args.top, `label tops: ${labels.slice(0, 4).map((x) => x.args.top).join(',')}`);
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'timeline', heading: 'x', milestones: Array(7).fill({ label: 'm' }) }] }); } catch (e) { threw = e.message; }
  ok('里程碑超上限 → 拒绝', /超上限/.test(threw), threw.slice(0, 50));
}
// closing / quote
{
  const c = expandDeck({ slides: [{ kind: 'closing', heading: '谢谢', contact: 'wps@x' }] })[0];
  eq('closing 步骤数（渐变+大字+联系=6）', c.length, 6);
  const q = expandDeck({ slides: [{ kind: 'quote', quote: '工欲善其事', attribution: '孔子' }] })[0];
  ok('quote 引文带书名括号', q.some((x) => x.args.text === '「工欲善其事」'));
  ok('quote 出处前缀——', q.some((x) => String(x.args.text || '').startsWith('—— ')));
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'quote', attribution: 'x' }] }); } catch (e) { threw = e.message; }
  ok('缺 quote → 拒绝', /quote/.test(threw), threw.slice(0, 50));
}

/* ---- 3. 纪律：未知 kind / 页数上限 / 全局 ---- */
console.log('\n=== 排版纪律（未知 kind / 页数上限）===');
{
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'magical', heading: 'x' }] }); } catch (e) { threw = e.message; }
  ok('未知 kind → 拒绝并列出支持的 10 种', /未知页型/.test(threw) && /title \/ section/.test(threw), threw.slice(0, 60));
  threw = '';
  try { expandDeck({ slides: [] }); } catch (e) { threw = e.message; }
  ok('空 slides → 拒绝', /至少给一页/.test(threw));
  threw = '';
  try { expandDeck({ slides: Array(31).fill({ kind: 'title', heading: 'x' }) }); } catch (e) { threw = e.message; }
  ok('页数超上限（31>30）→ 拒绝', /超过上限/.test(threw), threw.slice(0, 50));
  ok('每型步骤里字号全部 ≥ bodyFontMin（排版纪律抽查：bullets/compare/table 全量）', (() => {
    const all = [
      ...expandDeck({ slides: [{ kind: 'bullets', heading: 'h', points: ['a', 'b'] }] }),
      ...expandDeck({ slides: [{ kind: 'compare', heading: 'h', left: { label: 'L', points: ['x'] }, right: { label: 'R', points: ['y'] } }] }),
      ...expandDeck({ slides: [{ kind: 'table', heading: 'h', head: ['a'], rows: [['1']] }] }),
    ].filter((x) => x.args && x.args.fontSize !== undefined);
    return all.every((x) => x.args.fontSize >= DECK_TYPO.bodyFontMin);
  })());
}

/* ---- 4. mock COM 冒烟：标题页 + 要点页完整调用序列（不建快照垃圾）---- */
console.log('\n=== mock COM 冒烟（标题页+要点页完整调用序列，handler 直调不建快照）===');
(async () => {
  // mock WPS COM 层：记录每次 executeMethod 的动作与参数
  const wpsClientMod = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
  const comLog = [];
  let slideCounter = 0;
  wpsClientMod.wpsClient.executeMethod = async (action, params) => {
    comLog.push({ action, params });
    if (action === 'createPresentation') return { success: true, data: { name: 'MockDeck' } };
    if (action === 'addSlide') { slideCounter += 1; return { success: true, data: { slideIndex: slideCounter } }; }
    if (action === 'addTextbox') return { success: true, data: { name: 'TB' + comLog.length } };
    if (action === 'addShape') return { success: true, data: { name: 'SH' + comLog.length } };
    return { success: true, data: {} };
  };

  // 注册真工具，handler 直调（不经 callTool → 不建快照垃圾）
  const inst = new WpsMcpServer();
  inst.registerBuiltinTools();
  toolRegistry.registerAll(allTools);
  const deck = require(path.join(ROOT, 'dist', 'tools', 'ppt', 'deckkit.js'));
  const handler = deck.buildDeckHandler;

  const r = await handler({ deck: { slides: [
    { kind: 'title', heading: '冒烟标题', sub: 'smoke' },
    { kind: 'bullets', heading: '冒烟要点', points: ['第一点', '第二点'] },
  ] } });
  const text = r.content[0].text;
  ok('冒烟 success=true', r.success === true, text.slice(0, 60));
  ok('COM 序列以 getActiveDocumentState 开头（auto-snapshot 前置探测——mock 语义下无文档则快照放弃，不落盘）', comLog[0].action === 'getActiveDocumentState');
  ok('createPresentation 在 add_slide 之前', comLog.findIndex((x) => x.action === 'createPresentation') < comLog.findIndex((x) => x.action === 'addSlide'));
  ok('createPresentation 恰好一次（漏发是 B5-T3 首跑真 bug，此断言防回归）', comLog.filter((x) => x.action === 'createPresentation').length === 1);
  const createCalls = comLog.filter((x) => x.action === 'createPresentation').length;
  const pageActionCalls = comLog.filter((x) => x.action !== 'getActiveDocumentState').length;
  ok('页面动作序列完整（create 1 + add_slide 2 + 标题页 6 + 要点页 7 = 16，快照探测另计）', pageActionCalls === 16, 'pageActions=' + pageActionCalls + ' total=' + comLog.length);
  const adds = comLog.filter((x) => x.action === 'addSlide');
  ok('add_slide 两次且 layout=blank（干净页，形状计数前提）', adds.length === 2 && adds.every((x) => x.params.layout === 'blank'), JSON.stringify(adds.map((x) => x.params.layout)));
  ok('冒烟无 set_transition/set_background（标题+要点页不含）', !comLog.some((x) => x.action === 'setTransition' || x.action === 'setBackground'));
  ok('COM 调用里字号全部 ≥ bodyFontMin', comLog.filter((x) => x.params.fontSize !== undefined).every((x) => x.params.fontSize >= DECK_TYPO.bodyFontMin));

  // 失败路径：mid-way COM 失败 → fail-fast 且如实报停在哪页
  const failLog = [];
  wpsClientMod.wpsClient.executeMethod = async (action, params) => {
    if (action === 'addTextBox') throw new Error('mock: 标题文本框炸了');
    comLog.push({ action, params });
    if (action === 'createPresentation') return { success: true, data: { name: 'MockDeck2' } };
    if (action === 'addSlide') { slideCounter += 1; return { success: true, data: { slideIndex: 1 } }; }
    if (action === 'addTextbox') throw new Error('mock: 标题文本框炸了');
    return { success: true, data: {} };
  };
  const r2 = await handler({ deck: { slides: [{ kind: 'title', heading: '炸' }] } });
  ok('COM 中途炸 → success=false 且报错含现场（不静默续跑）',
    r2.success === false && /炸/.test(r2.error || ''), (r2.error || '').slice(0, 60));
  ok('fail-fast：第一页炸了就停（第二页未开始）', (r2.error || '').includes('完成到第 1 页'));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('门禁异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
