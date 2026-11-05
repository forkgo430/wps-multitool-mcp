/**
 * deckkit 引擎的 mock 单测（第 3 期 E2 骨架 + E3 主题/生成后溢出检查 + 质量修复轮断言）
 * 原则：不起真 WPS —— ① expandDeck 纯函数断言"坐标/字号/调用序列/主题"（10 页型 × ≥3 + 三主题）；
 * ② mock COM 冒烟：mock wpsClient.executeMethod 后走 build_ppt_deck 真入口跑完整调用序列
 *    （不建快照垃圾——handler 逐动作直调 handler，不经 callTool 的快照层）；
 * ③ 溢出检查三路（缩字修复 / 缩字仍超 / 已达下限不强改）+ 降级两路（bound 字段缺失 / 读回失败）：
 *    mock COM 按"行高 = 字号 × 1.2"模型回 BoundHeight（与引擎探针 seg3 同一标定数值）；
 * ④ 质量修复轮（审计三.1/三.2）：0 坐标原样到 COM + add_shape 发 type 键。
 * 语义依据：docs/agents/BORROW-PLAN.md §2.3（10 页型）+ 引擎探针（行高 1.2/渐变可用/BoundHeight 判溢出）。
 * 跑法：node scripts/manual/deckkit-test.js（verify-all FAST 档包含它）。
 */

const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const dk = require(path.join(ROOT, 'dist', 'tools', 'ppt', 'deckkit.js'));
const { expandDeck, DECK_TYPO, DECK_THEMES, DECK_OVERFLOW, deckSummary, pageTextBoxRecords, buildDeckHandler } = dk;
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

/* ---- 1. 排版纪律常量（一处集中）+ 主题常量 ---- */
console.log('\n=== 排版纪律常量（DECK_TYPO / DECK_THEMES / DECK_OVERFLOW）===');
ok('正文字号下限 ≥ 14（可读性纪律）', DECK_TYPO.bodyFontMin >= 14, String(DECK_TYPO.bodyFontMin));
ok('行高系数 = 1.2（引擎探针 seg3 本机标定）', DECK_TYPO.lineFactor === 1.2);
ok('密度上限集中可读（bullets 6 / table 6×4 / timeline 6 / slides 30）',
  DECK_TYPO.maxBullets === 6 && DECK_TYPO.maxTableRows === 6 && DECK_TYPO.maxTableCols === 4 && DECK_TYPO.maxTimeline === 6 && DECK_TYPO.maxSlides === 30);
ok('16:9 画布 960×540', DECK_TYPO.slideWidth === 960 && DECK_TYPO.slideHeight === 540);
ok('主题三套齐（学术蓝/政务红/商务灰）', !!DECK_THEMES['academic-blue'] && !!DECK_THEMES['gov-red'] && !!DECK_THEMES['business-gray'],
  Object.keys(DECK_THEMES).join('/'));
ok('政务红/商务灰带主题级间距覆盖（自创数值）',
  DECK_THEMES['gov-red'].spacing?.margin === 48 && DECK_THEMES['business-gray'].spacing?.cardGap === 20);
ok('溢出检查常量在案（tolerance 容差 + 缩字档位 2pt）', DECK_OVERFLOW.tolerancePt === 1 && DECK_OVERFLOW.fontStep === 2);

/* ---- 2. expandDeck：10 页型 × ≥3 断言（坐标/字号/调用序列）---- */
console.log('\n=== 10 页型展开（纯函数，坐标/字号/序列）===');

// title
{
  const p = expandDeck({ slides: [{ kind: 'title', heading: '季度汇报', sub: '2026 Q4', footer: '内部资料' }] })[0];
  eq('title 步骤数（渐变底两步+标题+白字+副标+accent+页脚+页脚色=8）', p.length, 8);
  eq('title 渐变底是第一个形状且全页', p[0].tool, 'wps_ppt_add_shape');
  eq('title 渐变参数（linear 90° 学术蓝→mid）', JSON.stringify(p[1].args.gradient),
    JSON.stringify({ type: 'linear', angle: 90, stops: [{ color: DECK_THEMES['academic-blue'].deep, position: 0 }, { color: DECK_THEMES['academic-blue'].mid, position: 1 }] }));
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
  ok('imageText 走 insert_image（大项②同步：新规范名 + path 参数）', p.some((x) => x.tool === 'wps_ppt_insert_image' && x.args.path === 'C:/tmp/x.png'));
  // 合并轮 A 段缺陷钉死：图片步必须计入形状序号——图片=#3，其后的要点文本框=#4
  // （缺陷复盘：SHAPE_CREATING_TOOLS 漏了新名 → 图片步不计数 → 文本框 shapeIndex 少算一格 → 溢出检查测错形状）
  eq('imageText 文本框记录 shapeIndex（图片步被计入：#2 标题、#4 要点）',
    pageTextBoxRecords(p).map((r) => r.shapeIndex), [2, 4]);
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
  eq('font_color 序列（标题白字2 + 表头白字 3/5）', p.filter((x) => x.tool === 'wps_ppt_set_font_color').map((x) => x.args.shapeIndex), [2, 3, 5]);
  let threw = '';
  try { expandDeck({ slides: [{ kind: 'table', heading: 'x', head: ['a', 'b'], rows: [['1']] }] }); } catch (e) { threw = e.message; }
  ok('行列不齐 → 拒绝', /不一致/.test(threw), threw.slice(0, 50));
}
// chart（两段式）
{
  const p = expandDeck({ slides: [{ kind: 'chart', heading: '图', categories: ['Q1', 'Q2'], series: [{ name: '销量', values: [1, 2] }] }] })[0];
  eq('chart 步骤数（带3 + insert + setdata = 5）', p.length, 5);
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

/* ---- 3. E3 主题（纯函数：缺省=学术蓝 / gov-red 边距与字体 / business-gray 卡距）---- */
console.log('\n=== E3 三主题（纯函数）===');
{
  const spec1 = { slides: [{ kind: 'bullets', heading: 'h', points: ['a', 'b'] }] };
  eq('缺省主题 = academic-blue（与显式传入全等）',
    JSON.stringify(expandDeck(spec1)), JSON.stringify(expandDeck(spec1, 'academic-blue')));
  const g = expandDeck({ slides: [{ kind: 'bullets', heading: 'h', points: ['a'] }] }, 'gov-red')[0];
  ok('gov-red 边距覆盖生效（卡片色块 left=48、paper 色）',
    g.some((x) => x.tool === 'wps_ppt_add_shape' && x.args.left === 48 && x.args.fillColor === '#FBF6EF'));
  ok('gov-red 字体对（黑体标题 + 仿宋正文）',
    g.some((x) => x.tool === 'wps_ppt_add_textbox' && x.args.fontName === '黑体')
    && g.some((x) => x.tool === 'wps_ppt_add_textbox' && x.args.fontName === '仿宋'));
  const bg = expandDeck(spec1, 'business-gray')[0];
  const ac = expandDeck(spec1)[0];
  const hbg = bg.filter((x) => x.tool === 'wps_ppt_add_shape')[1].args.height;
  const hac = ac.filter((x) => x.tool === 'wps_ppt_add_shape')[1].args.height;
  ok('business-gray 卡距加大（卡片高度差 = cardGap 20-18 = 2）', Math.round((hac - hbg) * 100) / 100 === 2, `academic=${hac} business=${hbg}`);
  ok('business-gray 用自己的色板（accent #B7791F）',
    expandDeck({ slides: [{ kind: 'quote', quote: 'q' }] }, 'business-gray')[0].some((x) => x.args.fillColor === '#B7791F'));
  let threw = '';
  try { expandDeck(spec1, 'solar-flare'); } catch (e) { threw = e.message; }
  ok('未知主题 → 拒绝并列出三套', /未知主题/.test(threw) && /gov-red/.test(threw) && /business-gray/.test(threw), threw.slice(0, 60));
  eq('文本框静态记录序号（bullets：标题#2 + 卡文 #4/#6）',
    pageTextBoxRecords(expandDeck(spec1)[0]).map((r) => r.shapeIndex), [2, 4, 6]);
  ok('静态记录带字号与帧几何（缩字档位/溢出判据的输入）',
    pageTextBoxRecords(expandDeck(spec1)[0]).every((r) => r.fontSize >= DECK_TYPO.bodyFontMin && r.width > 0 && r.height > 0));
}

/* ---- 4. 纪律：未知 kind / 页数上限 / 全局 ---- */
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

/* ================= 以下为 mock COM 冒烟（handler 直调不建快照）================= */
(async () => {
  /**
   * 有状态 mock COM：addShape/addTextBox 按页记形状序号，getTextBoxes 按
   * "行高 = 字号 × 1.2（引擎探针 seg3 同值）"模型回 BoundHeight —— 溢出检查的语义替身。
   * mode: 'model'（正常）| 'noBoundFields'（老桥回包缺 bound 字段）| 'fail'（读回通道炸）
   */
  function makeDeckMock(mode = 'model') {
    const comLog = [];
    const slides = [null];
    let slideCount = 0;
    const fn = async (action, params) => {
      comLog.push({ action, params });
      if (action === 'getActiveDocumentState') return { success: true, data: { hasDocument: false } };
      if (action === 'createPresentation') return { success: true, data: { name: 'MockDeck' } };
      if (action === 'addSlide') { slideCount += 1; slides[slideCount] = new Map(); return { success: true, data: { slideIndex: slideCount } }; }
      if (action === 'addShape') {
        const s = slides[params.slideIndex]; const n = s.size + 1;
        s.set(n, { kind: 'shape' });
        return { success: true, data: { name: 'SH' + n, slideIndex: params.slideIndex } };
      }
      if (action === 'addTextBox') {
        const s = slides[params.slideIndex]; const n = s.size + 1;
        s.set(n, { kind: 'text', text: params.text || '', fontSize: params.fontSize || 16, width: params.width, height: params.height });
        return { success: true, data: { name: 'TB' + n, slideIndex: params.slideIndex } };
      }
      if (action === 'getTextBoxes') {
        if (mode === 'fail') return { success: false, error: 'mock: 读回通道不可用' };
        const s = slides[params.slideIndex] || new Map();
        const out = [];
        for (const [i, b] of s) {
          if (b.kind !== 'text') continue;
          const rec = { index: i, name: 'TB' + i, text: b.text, left: 0, top: 0, width: b.width, height: b.height };
          if (mode === 'noBoundFields') { out.push(rec); continue; }
          // CJK 字宽 ≈ 字号，ASCII ≈ 0.55 × 字号；行数 = 估宽 / 可用宽（内边距 8pt 假设）
          const fs = b.fontSize;
          const estWidth = Array.from(b.text).reduce((w, ch) => w + (ch.charCodeAt(0) > 255 ? fs : fs * 0.55), 0);
          const usable = Math.max(b.width - 8, 20);
          const lines = Math.max(1, Math.ceil(estWidth / usable));
          rec.boundWidth = Math.round(Math.min(estWidth + 8, b.width) * 10) / 10;
          rec.boundHeight = Math.round(lines * fs * 1.2 * 10) / 10;
          rec.fontSize = fs;
          out.push(rec);
        }
        return { success: true, data: { slideIndex: params.slideIndex, textBoxes: out } };
      }
      if (action === 'setTextBoxStyle') {
        const s = slides[params.slideIndex];
        const b = s && s.get(params.shapeIndex);
        if (b && b.kind === 'text' && params.fontSize) b.fontSize = params.fontSize;
        return { success: true, data: { name: 'TB' + params.shapeIndex } };
      }
      return { success: true, data: {} };
    };
    return { fn, comLog };
  }

  // 注册真工具，handler 直调（不经 callTool → 不建快照垃圾）
  const inst = new WpsMcpServer();
  inst.registerBuiltinTools();
  toolRegistry.registerAll(allTools);
  const wpsClientMod = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
  const handler = buildDeckHandler;
  const jsonTail = (r) => JSON.parse(r.content[0].text.slice(r.content[0].text.lastIndexOf('\n') + 1));

  /* ---- 5. mock COM 冒烟：标题页 + 要点页完整调用序列 ---- */
  console.log('\n=== mock COM 冒烟（标题页+要点页完整调用序列，E3 口径含溢出读回）===');
  {
    const m = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = m.fn;
    const r = await handler({ deck: { slides: [
      { kind: 'title', heading: '冒烟标题', sub: 'smoke' },
      { kind: 'bullets', heading: '冒烟要点', points: ['第一点', '第二点'] },
    ] } });
    const text = r.content[0].text;
    ok('冒烟 success=true', r.success === true, text.slice(0, 60));
    ok('COM 序列以 getActiveDocumentState 开头（auto-snapshot 前置探测——mock 语义下无文档则快照放弃，不落盘）', m.comLog[0].action === 'getActiveDocumentState');
    ok('createPresentation 在 add_slide 之前', m.comLog.findIndex((x) => x.action === 'createPresentation') < m.comLog.findIndex((x) => x.action === 'addSlide'));
    ok('createPresentation 恰好一次（漏发是 B5-T3 首跑真 bug，此断言防回归）', m.comLog.filter((x) => x.action === 'createPresentation').length === 1);
    const pageActionCalls = m.comLog.filter((x) => x.action !== 'getActiveDocumentState').length;
    // E3 口径：create 1 + add_slide 2 + 标题页 6 + 要点页 7 + 溢出读回 2（每页一次 getTextBoxes）
    // 体验修复轮 A 口径：每个 add_shape/add_textbox 后一次 getShapes 索引回读（9 次：标题 3 + 要点 6）
    ok('页面动作序列完整（create 1 + add_slide 2 + 标题页 6 + 要点页 7 + 溢出读回 2 + 索引回读 9 = 27，快照探测另计）', pageActionCalls === 27, 'pageActions=' + pageActionCalls);
    ok('索引回读每 add 一次（getShapes ×9）', m.comLog.filter((x) => x.action === 'getShapes').length === 9);
    ok('回读落空不打印 undefined（回包无 undefined 字样）', !/undefined/.test(text));
    const adds = m.comLog.filter((x) => x.action === 'addSlide');
    ok('add_slide 两次且 layout=blank（干净页，形状计数前提）', adds.length === 2 && adds.every((x) => x.params.layout === 'blank'), JSON.stringify(adds.map((x) => x.params.layout)));
    ok('溢出读回每页恰好一次（getTextBoxes ×2）', m.comLog.filter((x) => x.action === 'getTextBoxes').length === 2);
    // 0 坐标修复断言（审计三.1）：deckkit 的全页底/标题色带发 left:0，COM 收到的必须还是 0
    const fullBleed = m.comLog.find((x) => x.action === 'addShape' && x.params.width === DECK_TYPO.slideWidth && x.params.height === DECK_TYPO.slideHeight);
    ok('0 坐标原样到 COM（全页底 left=0/top=0，不被 || 默认值吞）',
      !!fullBleed && fullBleed.params.left === 0 && fullBleed.params.top === 0,
      fullBleed ? `left=${fullBleed.params.left} top=${fullBleed.params.top}` : '未找到全页底形状');
    // shapeType→type 参数名错位修复断言（审计三.2）：deckkit 发 type 键，COM 不再收 undefined
    const shapes = m.comLog.filter((x) => x.action === 'addShape');
    ok('add_shape 步骤发 type 键且无 shapeType 残留（参数名错位已修）',
      shapes.length > 0 && shapes.every((x) => x.params.type === 'rectangle' && x.params.shapeType === undefined),
      `shapes=${shapes.length}`);
    ok('冒烟无溢出（遗留 0 处）', /遗留 0 处溢出/.test(text), (text.match(/溢出检查：[^\n]*/) || [''])[0]);
    ok('冒烟无 set_transition/set_background（标题+要点页不含）', !m.comLog.some((x) => x.action === 'setTransition' || x.action === 'setBackground'));
    ok('COM 调用里字号全部 ≥ bodyFontMin', m.comLog.filter((x) => x.params.fontSize !== undefined).every((x) => x.params.fontSize >= DECK_TYPO.bodyFontMin));
    const j = jsonTail(r);
    eq('回包 JSON 尾巴可解析（pages + overflow 结构）', [j.pages.length, j.overflow.checkedPages, j.overflow.degraded], [2, 2, false]);
    ok('deckSummary 工具集合含 setShapeGradient（标题页渐变走 COM）', j.pages[0].tools.includes('wps_ppt_set_shape_gradient'));
  }

  // 失败路径：mid-way COM 失败 → fail-fast 且如实报停在哪页（溢出检查不跑——页没铺完）
  {
    const mFail = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = async (action, params) => {
      if (action === 'addTextBox') throw new Error('mock: 标题文本框炸了');
      return mFail.fn(action, params);
    };
    const r2 = await handler({ deck: { slides: [{ kind: 'title', heading: '炸' }] } });
    ok('COM 中途炸 → success=false 且报错含现场（不静默续跑）',
      r2.success === false && /炸/.test(r2.error || ''), (r2.error || '').slice(0, 60));
    ok('fail-fast：第一页炸了就停（第二页未开始）', (r2.error || '').includes('完成到第 1 页'));
    ok('页没铺完不跑溢出检查（无 getTextBoxes）', mFail.comLog.every((x) => x.action !== 'getTextBoxes'));
  }

  // 0 坐标修复直调断言（审计三.1）：不经 deckkit，直接打 add_shape handler——
  // left:0/top:0 必须原样落 COM（修复前 || 会吞成 100）。
  // mock 用无状态版：makeDeckMock 的形状表按"先 addSlide 再加形状"建模，直调没有前情会踩它自己的空引用
  {
    const m0com = [];
    wpsClientMod.wpsClient.executeMethod = async (action, params) => {
      m0com.push({ action, params });
      if (action === 'getActiveDocumentState') return { success: true, data: { hasDocument: false } };
      return { success: true, data: { name: 'X', slideIndex: params.slideIndex || 1 } };
    };
    const r0 = await toolRegistry.callTool({
      id: 'audit-zero-coord',
      name: 'wps_ppt_add_shape',
      arguments: { slideIndex: 1, left: 0, top: 0, width: 100, height: 50 },
    });
    const a0 = m0com.find((x) => x.action === 'addShape');
    ok('0 坐标修复（add_shape 直调）：left:0/top:0 原样到 COM',
      r0.success === true && !!a0 && a0.params.left === 0 && a0.params.top === 0,
      a0 ? `left=${a0.params.left} top=${a0.params.top} err=${r0.error || '-'}` : '未捕获');
  }

  // 顺车（微轮 · PPT 定位去 Active 化）：get_table_cell 读回 value——
  // ps1 回的是 value，TS 历史读 text → 恒 undefined。两个键都接（value 优先），此处 mock 钉死
  {
    wpsClientMod.wpsClient.executeMethod = async (action, params) => {
      if (action === 'getPptTableCell') return { success: true, data: { value: '格子里的值' } };
      return { success: true, data: {} };
    };
    const rv = await toolRegistry.callTool({
      id: 'cell-value-priority',
      name: 'wps_ppt_get_table_cell',
      arguments: { slideIndex: 1, tableIndex: 1, row: 1, col: 1 },
    });
    ok('get_table_cell 读回 value（ps1 键，非历史 text 键）',
      rv.success === true && rv.content[0].text.includes('格子里的值') && !rv.content[0].text.includes('undefined'),
      (rv.content[0].text || '').split('\n').pop());
  }


  /* ---- 6. E3 溢出检查：三路 + 降级两路 ---- */
  console.log('\n=== E3 溢出检查（超框→警告→缩字一档复测一次）===');
  {
    // A. 缩字后不超：55 个 CJK 字在 16pt 时估宽 880 > 可用 836 → 溢出；缩到 14pt 估宽 770 ≤ 836 → 修复
    const a = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = a.fn;
    const ra = await handler({ deck: { slides: [{ kind: 'bullets', heading: '溢出冒烟', points: ['短要点', '长'.repeat(55)] }] } });
    ok('A 缩字修复：success=true', ra.success === true);
    eq('A 缩字落 COM（shapeIndex=6、fontSize 16→14）',
      a.comLog.filter((x) => x.action === 'setTextBoxStyle').map((x) => [x.params.shapeIndex, x.params.fontSize]), [[6, 14]]);
    ok('A 复测读回一次（getTextBoxes 共 2 次：初测+复测）', a.comLog.filter((x) => x.action === 'getTextBoxes').length === 2);
    ok('A 回包报自动修复且遗留 0 处', /遗留 0 处溢出/.test(ra.content[0].text) && /16→14pt/.test(ra.content[0].text), (ra.content[0].text.match(/溢出检查：[^\n]*/) || [''])[0]);
    const ja = jsonTail(ra);
    eq('A warnings 空（修完不留警告）', ja.overflow.warnings, []);
    eq('A issues fixed=true / finalFontSize=14', [ja.overflow.issues[0].fixed, ja.overflow.issues[0].finalFontSize], [true, 14]);

    // B. 缩字仍超：100 个 CJK 字在 14pt 时估宽 1400 仍 > 836 → 只警告不强改
    const b = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = b.fn;
    const rb = await handler({ deck: { slides: [{ kind: 'bullets', heading: '仍超', points: ['超'.repeat(100)] }] } });
    ok('B 缩字已尝试（setTextBoxStyle fontSize=14）', b.comLog.some((x) => x.action === 'setTextBoxStyle' && x.params.fontSize === 14));
    ok('B 仍超 → 只警告不强改，success 仍 true', rb.success === true && /缩字一档后仍溢出/.test(rb.content[0].text));
    const jb = jsonTail(rb);
    eq('B warnings 恰 1 条且 issues fixed=false', [jb.overflow.warnings.length, jb.overflow.issues[0].fixed], [1, false]);
    ok('B 警告文案带页/块号与帧/实际值（任务卡格式；单要点页溢出的是卡文#4）', /第 1 页第 4 块文字溢出（帧高/.test(jb.overflow.warnings[0]), jb.overflow.warnings[0].slice(0, 50));

    // C. 已达下限不强改：table 单元格字号就是 bodyFontMin(14)，62 字溢出但无档可缩
    const c = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = c.fn;
    const rc = await handler({ deck: { slides: [{ kind: 'table', heading: '下限', head: ['列'], rows: [['数'.repeat(62)]] }] } });
    ok('C 已达下限：不发缩字（无 setTextBoxStyle）', c.comLog.every((x) => x.action !== 'setTextBoxStyle'));
    ok('C 警告文案含下限提示，success 仍 true', rc.success === true && /已达正文字号下限/.test(rc.content[0].text));
    const jc = jsonTail(rc);
    eq('C 遗留 1 处、修复 0 处', [jc.overflow.issues.length, jc.overflow.issues.filter((x) => x.fixed).length], [1, 0]);

    // D1. 老桥 bound 字段缺失 → 降级记一条，不算失败
    const d1 = makeDeckMock('noBoundFields');
    wpsClientMod.wpsClient.executeMethod = d1.fn;
    const rd1 = await handler({ deck: { slides: [{ kind: 'bullets', heading: 'h', points: ['一'] }] } });
    const jd1 = jsonTail(rd1);
    ok('D1 bound 字段缺失 → 降级且 success=true', rd1.success === true && /降级/.test(rd1.content[0].text));
    eq('D1 checkedPages=0 / degraded=true', [jd1.overflow.checkedPages, jd1.overflow.degraded], [0, true]);

    // D2. 读回失败（通道炸）→ 同样降级
    const d2 = makeDeckMock('fail');
    wpsClientMod.wpsClient.executeMethod = d2.fn;
    const rd2 = await handler({ deck: { slides: [{ kind: 'bullets', heading: 'h', points: ['一'] }] } });
    const jd2 = jsonTail(rd2);
    ok('D2 读回失败 → 降级且 success=true', rd2.success === true && /降级/.test(rd2.content[0].text));
    eq('D2 degraded=true', jd2.overflow.degraded, true);
    eq('D2 降级警告只记一条（不逐页刷屏）', jd2.overflow.warnings.length, 1);

    // E. 读不到≠没超（体验修复轮 A）：单个形状 boundHeight 回 null（老桥对个别形状量不出）——
    // 该块不计入溢出判定，但按页出"无法测量"警告；其余可测的照常检查（不整体降级）
    const e = makeDeckMock();
    const origFn = e.fn;
    wpsClientMod.wpsClient.executeMethod = async (action, params) => {
      const r = await origFn(action, params);
      if (action === 'getTextBoxes' && r.success && r.data?.textBoxes?.length) {
        r.data.textBoxes[0] = { ...r.data.textBoxes[0], boundHeight: null, boundWidth: null };
      }
      return r;
    };
    const re = await handler({ deck: { slides: [{ kind: 'bullets', heading: 'h', points: ['可测的一', '长'.repeat(55)] }] } });
    ok('E 读不到≠没超：success=true 且按页出"无法测量"警告',
      re.success === true && /无法测量/.test(re.content[0].text));
    const je = jsonTail(re);
    ok('E 其余可测的照常检查（不整体降级，遗留 1 处）',
      je.overflow.degraded === false && je.overflow.checkedPages === 1 && je.overflow.issues.length === 1,
      `degraded=${je.overflow.degraded} checked=${je.overflow.checkedPages} issues=${je.overflow.issues.length}`);
    // 不可测 1 条；55 字点走缩字修复路（A 案），修完不留溢出警告 → 共 1 条
    eq('E 不可测警告恰 1 条（缩字修复路不留溢出警告）', je.overflow.warnings.length, 1);
  }

  /* ---- 7. E3 主题冒烟：三套各出一份（title 页走真入口）---- */
  console.log('\n=== E3 三主题 mock 冒烟（theme 参数 → COM 到色到字）===');
  {
    for (const [name, th] of Object.entries(DECK_THEMES)) {
      const m = makeDeckMock();
      wpsClientMod.wpsClient.executeMethod = m.fn;
      const r = await handler({ deck: { slides: [{ kind: 'title', heading: '主题冒烟', sub: 's' }] }, theme: name });
      ok(`${name} 冒烟 success=true`, r.success === true);
      ok(`${name} 渐变底用主题深色（addShape fillColor=${th.deep}）`, m.comLog.some((x) => x.action === 'addShape' && x.params.fillColor === th.deep));
      const grad = m.comLog.find((x) => x.action === 'setShapeGradient');
      ok(`${name} 渐变第二色到 mid`, !!grad && grad.params.gradient.stops[1].color === th.mid);
      ok(`${name} 标题字体到 COM（${th.fontPair.heading}）`, m.comLog.some((x) => x.action === 'addTextBox' && x.params.fontName === th.fontPair.heading));
    }
    // handler 的 theme 未知 → 如实报错
    const mBad = makeDeckMock();
    wpsClientMod.wpsClient.executeMethod = mBad.fn;
    const rBad = await handler({ deck: { slides: [{ kind: 'title', heading: 'x' }] }, theme: 'solar-flare' });
    ok('handler 未知主题 → success=false 且列出支持清单',
      rBad.success === false && /未知主题/.test(rBad.error || '') && /gov-red/.test(rBad.error || ''));
    ok('未知主题在展开期就拒绝（没起 COM）', mBad.comLog.length === 0 || mBad.comLog.every((x) => x.action === 'getActiveDocumentState'));
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('门禁异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
