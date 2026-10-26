/**
 * 加载项动作分发器的 mock 单测（Linux B4 验收项，异步 jsAsynCall 形态）
 * 原则：不起真 WPS —— mock 全局 cefQuery（B3 §3.3：webview 唯一原生桥），断言每个动作
 * 【发了什么 method、什么 params】而不是"返回了什么 mock 值"——防假测试的规矩不变。
 * B3 抓包范式（kwebpack_api_invoker.js）：request = "jsAsynCall('" + JSON + "')"，
 * 原生侧回调全局 jsCallbackN(result)。
 *
 * 断言保留策略：B2.5 的 79 条语义等价迁移（COM 属性赋值断言 → params 键值断言），
 * 新增：超时 reject、回调清理（无残留全局）、并发两调用不串线、cefQuery onFailure 原文。
 * 语义依据：docs/agents/linux/S1-REPORT.md §3/§5（method 出处见各断言标签与动作头注）。
 * 跑法：node scripts/manual/actions-dispatcher-test.js（verify-all FAST 档包含它）。
 */

const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const ADDON = path.resolve(ROOT, '..', 'wps-multitool-addon-linux');
const jsapi = require(path.join(ADDON, 'js', 'jsapi.js'));
const wordActions = require(path.join(ADDON, 'js', 'actions', 'word.js'));
const pptActions = require(path.join(ADDON, 'js', 'actions', 'ppt.js'));
const clientJs = require(path.join(ADDON, 'js', 'channel-client.js'));

// 超时收紧：超时用例不用等 30s（jsapi 每次调用时重读 env）
process.env.WPS_JSAPI_TIMEOUT = '120';

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};
const eq = (label, got, want) => ok(`${label} = ${JSON.stringify(want)}`, JSON.stringify(got) === JSON.stringify(want),
  `got ${JSON.stringify(got)}`);

/**
 * mock 全局 cefQuery：解析 jsAsynCall 请求，记录 {method, params, callback}，
 * 按 responses[method]（或默认响应）异步回调全局 jsCallbackN。
 * onCall(payload) 可覆盖默认响应；throw 则走"不回调"（供超时用例）。
 */
function installCefQuery(opts) {
  opts = opts || {};
  const sent = [];
  global.cefQuery = function (req) {
    const m = /^jsAsynCall\('(.*)'\)$/.exec(req.request);
    if (!m) { sent.push({ method: '<解析失败>', params: { raw: req.request } }); return; }
    const payload = JSON.parse(m[1]);
    const { method, callback, ...params } = payload;
    sent.push({ method, callback, params });
    const delay = opts.delay !== undefined ? opts.delay : 5;
    setTimeout(() => {
      if (opts.onFailureOnce && !opts.onFailureOnce.done) {
        opts.onFailureOnce.done = true;
        req.onFailure('JSAPI: 原生侧失败（mock）');
        return;
      }
      const cb = global[payload.callback];
      if (typeof cb !== 'function') { if (opts.onMissingCallback) opts.onMissingCallback(payload); return; }
      if (opts.onCall) {
        try {
          const r = opts.onCall(payload);
          cb(typeof r === 'string' ? r : JSON.stringify(r));
        } catch (e) { /* onCall 抛错 = 不回调（超时用例路径） */ }
        return;
      }
      const def = (opts.responses && opts.responses[payload.method]) !== undefined
        ? opts.responses[payload.method]
        : { success: true };
      cb(JSON.stringify(def));
    }, delay);
  };
  return sent;
}

function uninstallCefQuery() { delete global.cefQuery; }
const sentOf = (sent, method) => sent.filter((s) => s.method === method);
const lastSent = (sent) => sent[sent.length - 1];

(async () => {
  console.log('\n=== Word（wps 型）5 动作（异步 jsAsynCall）===');

  console.log('\n--- openDocument（S1:134 §3#108，method Documents.Open）---');
  {
    const sent = installCefQuery({ responses: { 'Documents.Open': { documentName: 'Doc1' } } });
    const r1 = await wordActions.openDocument({ filePath: '/tmp/a.docx', path: '/tmp/a.docx' });
    ok('正路 ok:true', r1.ok === true);
    const s1 = sentOf(sent, 'Documents.Open');
    eq('发了 method=Documents.Open 且 params.path 正确', s1.length === 1 && s1[0].params.path, '/tmp/a.docx');
    eq('回包透传原生结果', r1.data && r1.data.documentName, 'Doc1');
    const r2 = await wordActions.openDocument({});
    ok('缺 path → ok:false 且不再发送', r2.ok === false && sentOf(sent, 'Documents.Open').length === 1, r2.error);
    uninstallCefQuery();
    const r3 = await wordActions.openDocument({ path: '/tmp/a.docx' });
    ok('宿主无 cefQuery → ok:false 且 error 是原文（等价原"JSAPI 抛错原文"）',
      r3.ok === false && /cefQuery/.test(r3.error || ''), (r3.error || '').slice(0, 70));
  }

  console.log('\n--- getActiveDocumentState（S1:129 §3#103，method ActiveDocument）---');
  {
    const sent = installCefQuery({ responses: { ActiveDocument: { name: 'Doc1', hasPath: true, saved: false } } });
    const r1 = await wordActions.getActiveDocumentState({ app: 'word' });
    ok('正路 ok:true 且透传原生字段', r1.ok === true && r1.data.name === 'Doc1' && r1.data.hasPath === true, JSON.stringify(r1.data).slice(0, 80));
    eq('method=ActiveDocument', lastSent(sent).method, 'ActiveDocument');
    const m2 = installCefQuery({ responses: {} });
    const r2 = await wordActions.getActiveDocumentState({});
    ok('响应无字段 → 仍 ok（透传语义，数据以真机为准）', r2.ok === true, JSON.stringify(r2.data).slice(0, 60));
    uninstallCefQuery(); uninstallCefQuery();
  }

  console.log('\n--- getSelectedText（S1:303 §3#277，method Selection.Text）---');
  {
    const sent = installCefQuery({ responses: { 'Selection.Text': { text: '选中文字' } } });
    const r1 = await wordActions.getSelectedText({});
    eq('正路取回选区文本（响应 text 键优先）', r1.ok && r1.data && r1.data.text, '选中文字');
    eq('method=Selection.Text', lastSent(sent).method, 'Selection.Text');
    uninstallCefQuery();
    const m2 = installCefQuery({ responses: { 'Selection.Text': '裸文本' } });
    const r2 = await wordActions.getSelectedText({});
    eq('响应是裸字符串也能取出', r2.ok && r2.data.text, '裸文本');
    uninstallCefQuery();
  }

  console.log('\n--- replaceRange（S1:163 §3#137，method Range.Text）---');
  {
    const sent = installCefQuery({});
    const r1 = await wordActions.replaceRange({ startPos: 1, endPos: 3, text: '新文字' });
    ok('正路 ok:true', r1.ok === true);
    const s = lastSent(sent);
    ok('method=Range.Text 且 params 三键换算正确',
      s.method === 'Range.Text' && s.params.startPos === 1 && s.params.endPos === 3 && s.params.text === '新文字',
      JSON.stringify(s.params));
    const r2 = await wordActions.replaceRange({ startPos: 1, endPos: 3 });
    ok('缺 text → ok:false 且不重复发送', r2.ok === false && sent.filter((x) => x.method === 'Range.Text').length === 1, r2.error);
    const r3 = await wordActions.replaceRange({ text: 'x' });
    ok('缺坐标 → ok:false', r3.ok === false && sent.filter((x) => x.method === 'Range.Text').length === 1, r3.error);
    uninstallCefQuery();
  }

  console.log('\n--- setFont（S1:145 §3#119，method Font；换算在 JS 侧）---');
  {
    const sent = installCefQuery({});
    const r1 = await wordActions.setFont({ fontName: 'SimSun', fontSize: 24, bold: true, color: '#00FF00' });
    ok('正路 ok:true 且回包带 settings', r1.ok === true && r1.data.settings.colorRgb === 65280);
    const s1 = lastSent(sent);
    ok('params：fontName/fontSize/bold 正确', s1.params.fontName === 'SimSun' && s1.params.fontSize === 24 && s1.params.bold === true);
    ok('#00FF00 → params.colorRgb=65280（COM RGB 编码，S0 #18 背书）', s1.params.colorRgb === 65280);
    ok("缺省 range='selection'", s1.params.range === 'selection');
    const r2 = await wordActions.setFont({ fontName: 'KaiTi', range: 'all' });
    ok("range='all' → params.range='all'（走 Content 语义）", r2.ok === true && lastSent(sent).params.range === 'all');
    const r3 = await wordActions.setFont({ fontName: 'X', documentName: 'Doc1' });
    ok('钉文档+选区范围 → 拒绝（ps1 同款护栏）且不发送', r3.ok === false && /改到别人/.test(r3.error || '') && sent.filter((x) => x.method === 'Font').length === 2, r3.error);
    const r4 = await wordActions.setFont({ color: 'purple' });
    ok('非法色名 → ok:false 且未发送', r4.ok === false && sent.filter((x) => x.method === 'Font').length === 2, r4.error);
    uninstallCefQuery();
  }

  console.log('\n=== PPT（wpp 型）15 动作（异步 jsAsynCall）===');

  console.log('\n--- createPresentation（S1:309 §3#283，method Presentations.Add）---');
  {
    const sent = installCefQuery({ responses: { 'Presentations.Add': { name: 'Presentation2' } } });
    const r1 = await pptActions.createPresentation({});
    ok('正路 ok:true 且调 Presentations.Add', r1.ok === true && sentOf(sent, 'Presentations.Add').length === 1);
    eq('回包带新演示名', r1.data && r1.data.name, 'Presentation2');
    uninstallCefQuery();
    const r2 = await pptActions.createPresentation({});
    ok('宿主无 cefQuery → ok:false 原文', r2.ok === false && /cefQuery/.test(r2.error || ''), (r2.error || '').slice(0, 60));
  }

  console.log('\n--- getSlideCount（S1:202 §3#176，method Slides.Count）---');
  {
    const sent = installCefQuery({ responses: { 'Slides.Count': { count: 7 } } });
    const r1 = await pptActions.getSlideCount({});
    ok('正路回数字 count=7', r1.ok === true && r1.data.count === 7 && sentOf(sent, 'Slides.Count').length === 1);
    uninstallCefQuery();
    const m2 = installCefQuery({ responses: { 'Slides.Count': 3 } });
    const r2 = await pptActions.getSlideCount({});
    ok('响应是裸数字也取出', r2.ok === true && r2.data.count === 3);
    uninstallCefQuery();
  }

  console.log('\n--- addSlide（S1:210 §3#184，method Slides.Add；layout 编号换算在 JS 侧）---');
  {
    const sent = installCefQuery({ responses: { 'Slides.Add': { slideIndex: 4 } } });
    const r1 = await pptActions.addSlide({ layout: 'blank', title: 'T', content: 'C' });
    ok('正路 ok:true', r1.ok === true, JSON.stringify(r1.data));
    const s1 = lastSent(sent);
    ok("'blank' → params.layout=12（ps1 $layouts 表）", s1.params.layout === 12);
    ok('title/content 原样下发', s1.params.title === 'T' && s1.params.content === 'C');
    eq('回包 slideIndex 用原生回传', r1.data && r1.data.slideIndex, 4);
    const r2b = await pptActions.addSlide({ layout: 'title_content', content: '第二页' });
    const s2 = sentOf(sent, 'Slides.Add')[1];
    ok('缺省布局路径：显式 title_content → 2 且 content 下发',
      !!s2 && s2.params.layout === 2 && s2.params.content === '第二页', s2 ? JSON.stringify(s2.params) : '（没有第二次发送）');
    const r3 = await pptActions.addSlide({ layout: '不存在的版式' });
    ok('认不出的版式回落 2（ps1 同款）', r3.ok === true && lastSent(sent).params.layout === 2);
    uninstallCefQuery();
  }

  console.log('\n--- deleteSlide（S1:198 §3#172，method Slides.Delete）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.deleteSlide({ slideIndex: 2 });
    ok('正路发 method=Slides.Delete 且 params.index=2',
      r1.ok === true && sentOf(sent, 'Slides.Delete').length === 1 && lastSent(sent).params.index === 2);
    const r2 = await pptActions.deleteSlide({});
    ok('缺 slideIndex → ok:false 且不再发送', r2.ok === false && sentOf(sent, 'Slides.Delete').length === 1, r2.error);
    uninstallCefQuery();
  }

  console.log('\n--- getShapes（S1:222 §3#196，method Slide.Shapes）---');
  {
    const sent = installCefQuery({ responses: { 'Slide.Shapes': [{ name: 'TextBox 1', index: 1 }] } });
    const r1 = await pptActions.getShapes({ slideIndex: 2 });
    ok('正路透传原生形状数组', r1.ok === true && r1.data.shapes.length === 1 && r1.data.shapes[0].name === 'TextBox 1');
    eq('params.slideIndex 正确', lastSent(sent).params.slideIndex, 2);
    uninstallCefQuery();
    const m2 = installCefQuery({ responses: {} });
    const r2 = await pptActions.getShapes({ slideIndex: 1 });
    ok('响应缺省也 ok（透传）', r2.ok === true);
    uninstallCefQuery();
  }

  console.log('\n--- setShapeText（S1:226 §3#200，method Shape.TextFrame.TextRange.Text）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.setShapeText({ slideIndex: 1, shapeIndex: 1, text: '新形状文字' });
    ok('正路 method 正确且 params 带 text/shapeIndex',
      r1.ok === true && lastSent(sent).method === 'Shape.TextFrame.TextRange.Text' &&
      lastSent(sent).params.text === '新形状文字' && lastSent(sent).params.shapeIndex === 1);
    const r2 = await pptActions.setShapeText({ slideIndex: 1 });
    ok('缺 text → ok:false 且不发送', r2.ok === false && sent.filter((x) => x.method === 'Shape.TextFrame.TextRange.Text').length === 1, r2.error);
    const r3 = await pptActions.setShapeText({ slideIndex: 1, text: 'x' });
    ok('形状没给 → ok:false（ps1 坏对象坑：Item(null) 毁整页）且不发送',
      r3.ok === false && /找不到目标形状/.test(r3.error || '') && sent.filter((x) => x.method === 'Shape.TextFrame.TextRange.Text').length === 1);
    uninstallCefQuery();
  }

  console.log('\n--- addTextBox（S1:211 §3#185，method Shapes.AddTextbox；缺省镜像 ps1）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.addTextBox({ slideIndex: 1, text: '框内容', fontSize: 20 });
    const s1 = lastSent(sent);
    ok('orientation=1 + 缺省 100/100/400/50（ps1 同款）',
      s1.method === 'Shapes.AddTextbox' && s1.params.orientation === 1 &&
      s1.params.left === 100 && s1.params.top === 100 && s1.params.width === 400 && s1.params.height === 50,
      JSON.stringify(s1.params));
    ok('text/fontSize 进 params', s1.params.text === '框内容' && s1.params.fontSize === 20);
    const r2 = await pptActions.addTextBox({ slideIndex: 1, left: 10, top: 20, width: 30, height: 40 });
    const s2 = lastSent(sent);
    ok('显式坐标直传', s2.params.left === 10 && s2.params.top === 20 && s2.params.width === 30 && s2.params.height === 40);
    const r3 = await pptActions.addTextBox({ slideIndex: 1 });
    ok('不传 text 也不炸（文本可选）', r3.ok === true && !('text' in lastSent(sent).params));
    uninstallCefQuery();
  }

  console.log('\n--- setShapePosition（S1:227 §3#201，method Shape.Position）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.setShapePosition({ slideIndex: 1, shapeIndex: 1, left: 300, top: 250 });
    const s1 = lastSent(sent);
    ok('只下发给出的维度：left/top 在 params', s1.params.left === 300 && s1.params.top === 250);
    ok('width/height 不在 params（没给就不发）', !('width' in s1.params) && !('height' in s1.params));
    const r2 = await pptActions.setShapePosition({ slideIndex: 1, left: 1 });
    ok('形状没给 → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
    const r3 = await pptActions.setShapePosition({ slideIndex: 1, shapeIndex: 1 });
    ok('一个维度都没给 → ok:false', r3.ok === false && /至少给一个/.test(r3.error || ''), (r3.error || '').slice(0, 60));
    uninstallCefQuery();
  }

  console.log('\n--- insertPptTable（S1:233 §3#207，method Shapes.AddTable）---');
  {
    const sent = installCefQuery({ responses: { 'Shapes.AddTable': { name: 'Table 4', tableIndex: 2 } } });
    const r1 = await pptActions.insertPptTable({ slideIndex: 1, rows: 4, cols: 2, left: 5, top: 6 });
    const s1 = lastSent(sent);
    ok('AddTable(4,2,5,6,400,200)（缺省尺寸 ps1 同款）',
      s1.params.rows === 4 && s1.params.cols === 2 && s1.params.left === 5 && s1.params.top === 6 &&
      s1.params.width === 400 && s1.params.height === 200, JSON.stringify(s1.params));
    eq('tableIndex 透传原生回传', r1.data && r1.data.tableIndex, 2);
    uninstallCefQuery();
    const m2 = installCefQuery({});
    const r2 = await pptActions.insertPptTable({ slideIndex: 1 });
    ok('全缺省 → 3×3 @ 100/100/400/200', r2.ok === true &&
      lastSent(m2).params.rows === 3 && lastSent(m2).params.height === 200);
    uninstallCefQuery();
  }

  console.log('\n--- getPptTableCell / setPptTableCell（S1:235/234 §3#209/#208，method Table.Cell.Text）---');
  {
    const sent = installCefQuery({ responses: { 'Table.Cell.Text': { text: '旧内容2,3' } } });
    const r1 = await pptActions.getPptTableCell({ slideIndex: 1, tableIndex: 1, row: 2, col: 3 });
    const s1 = lastSent(sent);
    ok('读路 method/params（tableIndex 缺省 1）', s1.method === 'Table.Cell.Text' && s1.params.row === 2 && s1.params.col === 3);
    eq('读回 value/text 双键', r1.data && r1.data.value, '旧内容2,3');
    const r3 = await pptActions.setPptTableCell({ slideIndex: 1, tableIndex: 2, row: 1, col: 1, text: '新格值' });
    const s3 = lastSent(sent);
    ok('写路 params 带 text 与 tableIndex', s3.params.text === '新格值' && s3.params.tableIndex === 2);
    const r4 = await pptActions.setPptTableCell({ slideIndex: 1, row: 1, col: 1 });
    ok('缺 text → ok:false 且不发送', r4.ok === false && sent.filter((x) => x.method === 'Table.Cell.Text').length === 2, r4.error);
    uninstallCefQuery();
  }

  console.log('\n--- addAnimation（S1:252/271/272 §3#226/#245/#246 + §5 行 355；只做加铁律）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.addAnimation({ slideIndex: 1, shapeIndex: 1, effect: 'fadeIn' });
    const s1 = lastSent(sent);
    ok("effect 'fadeIn' → params.effectId=10（映射表抄 ps1）+ level 恒 0",
      s1.method === 'TimeLine.MainSequence.AddEffect' && s1.params.effectId === 10 && s1.params.level === 0,
      JSON.stringify(s1.params));
    ok('回包没有 count 字段（S1 §5：读侧不稳，异步模型下同样绝不发起读侧调用）',
      r1.ok === true && !('count' in (r1.data || {})));
    ok('发送列表里没有任何动画读侧 method（无 Count/Delete/MoveTo）',
      !sent.some((x) => /Count|Delete|MoveTo/.test(x.method)));
    const r2 = await pptActions.addAnimation({ slideIndex: 1, shapeIndex: 1, effect: 29 });
    ok('数字 effect 直通 → params.effectId=29（S1 §5 编号恒等）',
      r2.ok === true && lastSent(sent).params.effectId === 29 && lastSent(sent).params.level === 0);
    const r3 = await pptActions.addAnimation({ slideIndex: 1, shapeIndex: 1, effect: 'magicFly' });
    ok('认不出的效果名 → 明确报错（ps1 同款，不猜）且不发送',
      r3.ok === false && /不认识的 effect/.test(r3.error || '') && sent.filter((x) => x.method === 'TimeLine.MainSequence.AddEffect').length === 2);
    const r4 = await pptActions.addAnimation({ slideIndex: 1, effect: 'fadeIn' });
    ok('形状没给 → ok:false（ps1 坏对象坑）', r4.ok === false && /找不到要加动画的形状/.test(r4.error || ''));
    uninstallCefQuery();
  }

  console.log('\n--- setSlideTransition（S1:253 §3#227 + §5 行 354，method SlideShowTransition.EntryEffect）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.setSlideTransition({ slideIndex: 1, effect: 'push', duration: 2 });
    let s = lastSent(sent);
    ok("'push' → params.entryEffect=2（映射表抄 ps1）+ duration 下发",
      s.method === 'SlideShowTransition.EntryEffect' && s.params.entryEffect === 2 && s.params.duration === 2);
    const r2 = await pptActions.setSlideTransition({ slideIndex: 1, effect: 'none' });
    ok("'none' → 0（取消切换不再炸，ps1 历史坑）", lastSent(sent).params.entryEffect === 0);
    const r3 = await pptActions.setSlideTransition({ slideIndex: 1, effect: '77' });
    ok('数字串直通 → 77', lastSent(sent).params.entryEffect === 77);
    const before = sent.length;
    const r4 = await pptActions.setSlideTransition({ slideIndex: 1, effect: 'starWipe' });
    ok('认不出的切换 → ok:false 且未发送', r4.ok === false && sent.length === before, r4.error);
    uninstallCefQuery();
  }

  console.log('\n--- setSlideBackground（S1:254 §3#228 + S0 #18，method Background.Fill.Solid）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.setSlideBackground({ slideIndex: 1, color: '#00FF00' });
    const s1 = lastSent(sent);
    ok('params.colorRgb=65280（Solid+RGB 两步编码进单 method，次序由原生侧承担——B3 复核点）',
      s1.method === 'Background.Fill.Solid' && s1.params.colorRgb === 65280, JSON.stringify(s1.params));
    const before = sent.length;
    const r2 = await pptActions.setSlideBackground({ slideIndex: 1, color: 'zzz' });
    ok('非法颜色 → ok:false（ps1 静默跳过仍回 success——有意差异）且未发送',
      r2.ok === false && sent.length === before, (r2.error || '').slice(0, 60));
    const r3 = await pptActions.setSlideBackground({ slideIndex: 1, imagePath: '/tmp/x.png' });
    ok('图片路径本轮未实现 → 明确报"纯色"，不静默成功', r3.ok === false && /纯色/.test(r3.error || ''), (r3.error || '').slice(0, 60));
    uninstallCefQuery();
  }

  console.log('\n--- exportSlideAsImage（S1:232 §3#206 + S0 #19；B3 §5 明确属异步范畴）---');
  {
    const sent = installCefQuery({});
    const r1 = await pptActions.exportSlideAsImage({ slideIndex: 1, outputPath: '/tmp/s1.png' });
    const s1 = lastSent(sent);
    ok("method=Slide.Export 且 params.path/filter（PNG）",
      s1.method === 'Slide.Export' && s1.params.path === '/tmp/s1.png' && s1.params.filter === 'PNG');
    const r2 = await pptActions.exportSlideAsImage({ slideIndex: 1, path: '/tmp/s2.png', format: 'JPG', width: 1334, height: 750 });
    const s2 = lastSent(sent);
    ok('path 别名键也认 + 带尺寸四参下发（S0 实测带尺寸成功）',
      s2.params.path === '/tmp/s2.png' && s2.params.filter === 'JPG' && s2.params.width === 1334 && s2.params.height === 750);
    const before = sent.length;
    const r3 = await pptActions.exportSlideAsImage({ slideIndex: 1 });
    ok('缺 path → ok:false 且不发送', r3.ok === false && sent.length === before, r3.error);
    uninstallCefQuery();
  }

  console.log('\n=== 分发器接线（注册表查找 + Promise 化 + 回落）===');
  {
    clientJs.WpsAddonActionsWord = wordActions;
    clientJs.WpsAddonActionsPpt = pptActions;
    const d = clientJs.WpsAddonChannel.dispatch;
    const ping = await d('ping', {}, 'word');
    ok('ping 仍真实现（Promise 化）', ping.ok === true && ping.data.pong === true);
    const wc = await d('wireCheck', {}, 'wpp');
    ok('wireCheck 仍真实现（文案 B4 异步动作层）', wc.ok === true && /B4/.test(wc.data.message || ''));
    const ghost = await d('setValue', { x: 1 }, 'word');
    ok('注册表外的动作仍回落 unsupported:B3', ghost.ok === false && ghost.unsupported === 'B3');
    // 注册表命中但宿主无 cefQuery（Node 侧未 mock）→ 动作 reject → dispatch 兜成 ok:false 原文
    const host = await d('getSlideCount', {}, 'ppt');
    ok('注册表命中但无 cefQuery → 兜成 ok:false 原文（不静默成功）',
      host.ok === false && /cefQuery/.test(host.error || ''), (host.error || '').slice(0, 70));
    uninstallCefQuery();
  }

  console.log('\n=== jsapi.js 专项：超时 / 回调清理 / 并发 / onFailure ===');
  {
    // 超时：onCall 抛错 = 永不回调 → 120ms 后 reject（await 会 throw，用 .then 捕成值）
    installCefQuery({ onCall: () => { throw new Error('never call back'); } });
    const t0 = Date.now();
    const r1 = await jsapi.jsAsynCall('slowMethod', {}).then(() => 'resolved', (e) => e);
    const dt = Date.now() - t0;
    ok('超时 → reject 且文案含 method', r1 instanceof Error && /jsAsynCall 超时/.test(r1.message) && /slowMethod/.test(r1.message),
      r1.message.slice(0, 60));
    ok('超时时长 ≈ WPS_JSAPI_TIMEOUT(120ms)', dt >= 110 && dt < 1000, `${dt}ms`);
    ok('超时后全局无残留 jsCallback（清理到位）', jsapi.pendingCallbacks().length === 0, JSON.stringify(jsapi.pendingCallbacks()));
    uninstallCefQuery();

    // 正常调用后同样无残留
    installCefQuery({ responses: { m1: { a: 1 } } });
    await jsapi.jsAsynCall('m1', {});
    ok('正常调用后无残留 jsCallback', jsapi.pendingCallbacks().length === 0);
    uninstallCefQuery();

    // 并发：两个调用交错，各自回调不串线
    const m2 = installCefQuery({
      responses: { mA: { which: 'A' }, mB: { which: 'B' } },
      delay: 15,
    });
    const pa = jsapi.jsAsynCall('mA', {});
    const pb = jsapi.jsAsynCall('mB', {});
    const [ra, rb] = await Promise.all([pa, pb]);
    ok('并发 A：拿到 A 的结果（不串线）', ra && ra.which === 'A');
    ok('并发 B：拿到 B 的结果（不串线）', rb && rb.which === 'B');
    eq('并发发送两条且 method 各自正确', sentOf(m2, 'mA').length + sentOf(m2, 'mB').length, 2);
    ok('并发完成后无残留 jsCallback', jsapi.pendingCallbacks().length === 0);
    // 回调名自增不重复
    const names = m2.map((x) => x.callback);
    ok('回调名自增不重复', new Set(names).size === names.length, names.join(','));
    uninstallCefQuery();

    // onFailure：原生侧失败 → reject 原文（等价原"JSAPI 抛错原文带回"）
    installCefQuery({ onFailureOnce: {} });
    const rf = await jsapi.jsAsynCall('failingMethod', {}).then(() => 'resolved', (e) => e);
    ok('cefQuery onFailure → reject 且 error 是原文', rf instanceof Error && /原生侧失败/.test(rf.message), rf.message.slice(0, 60));
    ok('onFailure 后无残留 jsCallback', jsapi.pendingCallbacks().length === 0);
    uninstallCefQuery();
  }

  console.log('\n=== Excel（et 型）B5-T1：57 实现 + 3 PivotTables 桩（表驱动，断调用序列）===');
  {
    const ex = require(path.join(ADDON, 'js', 'actions', 'excel.js'));
    const A = ex.actions;
    const AD = ex.adapter;

    // ---- ExcelAdapter 纯函数直测（集合双语义的编码层）----
    eq('colToLetters(1)=A', AD.colToLetters(1), 'A');
    eq('colToLetters(26)=Z', AD.colToLetters(26), 'Z');
    eq('colToLetters(27)=AA', AD.colToLetters(27), 'AA');
    eq('colToLetters(52)=AZ', AD.colToLetters(52), 'AZ');
    eq('a1(3,2)=B3（Item 语义的 A1 编码）', AD.a1(3, 2), 'B3');
    eq('refOf：row/col 组 A1', AD.refOf({ row: 5, col: 3 }), 'C5');
    eq('refOf：range 优先', AD.refOf({ range: 'B2:C9', row: 1, col: 1 }), 'B2:C9');
    let threwV = '';
    try { AD.writeValue2({ ref: 'A1', value: 1 }); } catch (e) { threwV = e.message; }
    ok('writeValue2 防呆：value 键直接抛（S1 §4 #3 静默失效的编码层禁止）', /禁止 value 键/.test(threwV), threwV.slice(0, 60));
    eq('withValue2：value 自动改道 value2', AD.withValue2({ ref: 'A1' }, 42).value2, 42);
    ok('cellRef：Item 语义定位参数', JSON.stringify(AD.cellRef(2, 3)) === JSON.stringify({ row: 2, col: 3 }));

    // ---- 57 实现动作：正路（method + 关键 params）表驱动 ----
    const cases = [
      ['getActiveWorkbook', {}, 'ActiveWorkbook', {}],
      ['getRangeData', { range: 'A1:B3' }, 'Range.Value2', { ref: 'A1:B3' }],
      ['setRangeData', { range: 'A2:B3', value2: [[1, 2], [3, 4]] }, 'Range.Value2', { ref: 'A2:B3', value2: [[1, 2], [3, 4]] }],
      ['setRangeData', { range: 'A2:B3', value: '旧键改道' }, 'Range.Value2', { value2: '旧键改道' }],
      ['clearFormats', { range: 'B2:D4' }, 'Range.ClearFormats', { ref: 'B2:D4' }],
      ['addConditionalFormat', { range: 'A1:A9', type: 1, operator: 1, formula: '>0' }, 'FormatConditions.Add', { type: 1, operator: 1, formula: '>0' }],
      ['removeConditionalFormat', { range: 'A1:A9' }, 'FormatConditions.Delete', { ref: 'A1:A9' }],
      ['getConditionalFormats', { range: 'A1:A9' }, 'FormatConditions.Count', { ref: 'A1:A9' }],
      ['addDataValidation', { range: 'B1:B9', type: 1, operator: 1, alertStyle: 1, formula: '10' }, 'Validation.Add', { formula: '10' }],
      ['removeDataValidation', { range: 'B1:B9' }, 'Validation.Delete', { ref: 'B1:B9' }],
      ['getDataValidations', { range: 'B1:B9' }, 'Validation', { ref: 'B1:B9' }],
      ['setArrayFormula', { range: 'C1:C3', formula: '=A1:A3&"!"' }, 'Range.FormulaArray', { formula: '=A1:A3&"!"' }],
      ['refreshLinks', {}, 'Workbook.UpdateLink', {}],
      ['consolidate', { range: 'E1:E9' }, 'Range.Consolidate', { ref: 'E1:E9' }],
      ['calculateSheet', { sheet: 'Sheet1' }, 'Worksheet.Calculate', { sheet: 'Sheet1' }],
      ['getExcelContext', {}, 'Context', { includeSelection: false }],
      ['getContext', {}, 'Context', { includeSelection: true }],
      ['sortRange', { range: 'A1:C9', key: 'B', order: 2 }, 'Range.Sort', { key: 'B', order: 2 }],
      ['autoFilter', { range: 'A1:C9', field: 2, criteria: 'x' }, 'Range.AutoFilter', { field: 2, criteria: 'x' }],
      ['updateChart', { chartIndex: 1, chartType: 51 }, 'Chart', { chartIndex: 1, chartType: 51 }],
      ['exportChartAsImage', { chartIndex: 1, outputPath: '/tmp/c.png', format: 'PNG' }, 'Chart.Export', { path: '/tmp/c.png', filter: 'PNG' }],
      ['removeDuplicates', { range: 'A1:C9', columns: 1, header: true }, 'Range.RemoveDuplicates', { columns: 1, header: true }],
      ['deleteSheet', { name: 'Sheet2' }, 'Worksheet.Delete', { name: 'Sheet2' }],
      ['deleteSheet', { index: 2 }, 'Worksheet.Delete', { index: 2 }],
      ['renameSheet', { target: 'Sheet1', name: '探针表' }, 'Worksheet.Name', { name: '探针表', target: 'Sheet1' }],
      ['copySheet', { target: 'Sheet1' }, 'Worksheet.Copy', { target: 'Sheet1' }],
      ['getSheetList', {}, 'Worksheets', {}],
      ['getSheetInfo', { sheet: 'Sheet1' }, 'Worksheet.Info', { sheet: 'Sheet1' }],
      ['autoFit', { sheet: 'Sheet1' }, 'UsedRange.AutoFit', { sheet: 'Sheet1' }],
      ['findInSheet', { what: '探针', sheet: 'Sheet1' }, 'Cells.Find', { what: '探针', sheet: 'Sheet1' }],
      ['setSheetPageSetup', { orientation: 2 }, 'PageSetup', { orientation: 2 }],
      ['setSheetPrintTitles', { rows: '$1:$1' }, 'PageSetup.PrintTitleRows', { rows: '$1:$1' }],
      ['switchSheet', { index: 2 }, 'Worksheets.Activate', { index: 2 }],
      ['switchSheet', { name: 'Sheet2' }, 'Worksheets.Activate', { name: 'Sheet2' }],
      ['moveSheet', { target: 'Sheet1', before: 1 }, 'Worksheet.Move', { target: 'Sheet1', before: 1 }],
      ['mergeCells', { range: 'A1:B2' }, 'Range.Merge', { ref: 'A1:B2' }],
      ['unmergeCells', { range: 'A1:B2' }, 'Range.UnMerge', { ref: 'A1:B2' }],
      ['autoFitAll', {}, 'UsedRange.AutoFit', {}],
      ['setPrintArea', { area: '$A$1:$D$5' }, 'PageSetup.PrintArea', { area: '$A$1:$D$5' }],
      ['getSelection', {}, 'Selection.Address', {}],
      ['clearRange', { range: 'A1:B2' }, 'Range.ClearContents', { ref: 'A1:B2' }],
      ['freezePanes', {}, 'ActiveWindow.FreezePanes', { freeze: true }],
      ['unfreezePanes', {}, 'ActiveWindow.FreezePanes', { freeze: false }],
      ['copyRange', { range: 'A1:B2', destination: 'D1' }, 'Range.Copy', { ref: 'A1:B2', destination: 'D1' }],
      ['pasteRange', { range: 'D1', pasteMode: 1 }, 'Range.PasteSpecial', { ref: 'D1', pasteMode: 1 }],
      ['fillSeries', { range: 'A1:A9', direction: 'down' }, 'Range.DataSeries', { ref: 'A1:A9', direction: 'down' }],
      ['textToColumns', { range: 'A1:A9', delimiter: ',' }, 'Range.TextToColumns', { ref: 'A1:A9', delimiter: ',' }],
      ['subtotal', { range: 'A1:C9', groupBy: 1, 'function': 1, totalList: [2, 3] }, 'Range.Subtotal', { groupBy: 1 }],
      ['createNamedRange', { name: '探针名', refersTo: 'Sheet1!$A$1:$B$2' }, 'Names.Add', { name: '探针名', refersTo: 'Sheet1!$A$1:$B$2' }],
      ['getNamedRanges', {}, 'Names', {}],
      ['protectSheet', { sheet: 'Sheet1', password: 'p' }, 'Worksheet.Protect', { sheet: 'Sheet1', password: 'p' }],
      ['unprotectSheet', { sheet: 'Sheet1' }, 'Worksheet.Unprotect', { sheet: 'Sheet1' }],
      ['protectWorkbook', { structureOnly: true }, 'Workbook.Protect', { structureOnly: true }],
      ['insertExcelImage', { path: '/tmp/x.png', left: 10, top: 20 }, 'Shapes.AddPicture', { path: '/tmp/x.png', left: 10, top: 20 }],
      ['setHyperlink', { anchorRef: 'A1', address: 'https://example.com' }, 'Hyperlinks.Add', { anchor: 'A1', address: 'https://example.com' }],
      ['openWorkbook', { path: '/tmp/wb.xlsx' }, 'Workbooks.Open', { path: '/tmp/wb.xlsx' }],
      ['getOpenWorkbooks', {}, 'Workbooks', {}],
      ['switchWorkbook', { index: 2 }, 'Workbooks.Activate', { index: 2 }],
      ['closeWorkbook', { index: 2, saveChanges: false }, 'Workbooks.Close', { index: 2, saveChanges: false }],
      ['createWorkbook', {}, 'Workbooks.Add', {}],
    ];
    for (const [action, input, method, wantParams] of cases) {
      const sent = installCefQuery({});
      let r;
      try { r = await A[action](input); } catch (e) { r = { ok: false, error: String(e) }; }
      ok(`${action} 正路 ok:true`, r && r.ok === true, r && r.error);
      const hits = sentOf(sent, method);
      ok(`${action} 恰好发一条 method=${method}`, hits.length === 1,
        hits.length ? JSON.stringify(hits[0].params).slice(0, 90) : `实际发送：${sent.map((x) => x.method).join(',') || '（无）'}`);
      for (const [k, v] of Object.entries(wantParams)) {
        eq(`${action} params.${k}`, hits.length ? hits[0].params[k] : undefined, v);
      }
      uninstallCefQuery();
    }

    // ---- 错误路（校验拒绝 + 绝不发送）----
    const errCases = [
      ['getRangeData', {}, /缺少 range/],
      ['clearFormats', {}, /缺少 range/],
      ['setRangeData', { range: 'A1:B2' }, /缺少 value2/],
      ['setArrayFormula', { range: 'A1' }, /缺少 formula/],
      ['sortRange', { range: 'A1:C9' }, /缺少 key/],
      ['updateChart', {}, /chartIndex\/chartName/],
      ['exportChartAsImage', {}, /outputPath\/path/],
      ['deleteSheet', {}, /name\/index/],
      ['renameSheet', {}, /缺少 name/],
      ['findInSheet', {}, /缺少 what/],
      ['setSheetPageSetup', {}, /至少给一个/],
      ['setSheetPrintTitles', {}, /缺少 rows/],
      ['switchSheet', {}, /name\/index/],
      ['setPrintArea', {}, /缺少 area/],
      ['clearRange', {}, /缺少 range/],
      ['copyRange', {}, /缺少 range/],
      ['pasteRange', {}, /缺少 range/],
      ['fillSeries', {}, /缺少 range/],
      ['textToColumns', {}, /缺少 range/],
      ['subtotal', {}, /缺少 range/],
      ['createNamedRange', { name: 'x' }, /refersTo 必填/],
      ['insertExcelImage', {}, /path\/imagePath/],
      ['setHyperlink', { anchorRef: 'A1' }, /address 必填/],
      ['openWorkbook', {}, /path\/filePath/],
      ['switchWorkbook', {}, /name\/index/],
      ['closeWorkbook', {}, /name\/index/],
      ['mergeCells', {}, /缺少 range/],
      ['consolidate', {}, /缺少 range/],
    ];
    for (const [action, input, match] of errCases) {
      const sent = installCefQuery({});
      const r = await A[action](input);
      ok(`${action} 错误路 → ok:false 且匹配「${match.source.slice(0, 30)}」`,
        r && r.ok === false && match.test(r.error || ''), (r && r.error || '').slice(0, 60));
      ok(`${action} 错误路零发送`, sent.length === 0, `实际发送 ${sent.length} 条`);
      uninstallCefQuery();
    }

    // ---- PivotTables 桩 3 条：unsupported:'待S1补测'，绝不猜 method ----
    for (const action of ['createPivotTable', 'getPivotTables', 'refreshPivotTables']) {
      const sent = installCefQuery({});
      const r = await A[action]({});
      ok(`${action} → unsupported:'待S1补测'（S1 §4 #6：Count=undefined、Item 不可调）`,
        r && r.ok === false && r.unsupported === '待S1补测', (r && r.error || '').slice(0, 60));
      ok(`${action} 桩零发送`, sent.length === 0);
      uninstallCefQuery();
    }

    // ---- 分发器接线：excel 表命中（et 型动作经 channel-client 分发）----
    clientJs.WpsAddonActionsExcel = A;
    installCefQuery({ responses: { 'Slides.Count': { count: 9 } } });
    const viaDispatch = await clientJs.WpsAddonChannel.dispatch('getSlideCount', {}, 'ppt');
    ok('分发器接线：ppt 表不受 excel 表影响（回归守护）', viaDispatch.ok === true && viaDispatch.data.count === 9);
    uninstallCefQuery();
    installCefQuery({ responses: { 'Workbooks.Add': { name: 'Workbook2' } } });
    const viaExcel = await clientJs.WpsAddonChannel.dispatch('createWorkbook', {}, 'excel');
    ok('分发器接线：excel 动作表已注册（createWorkbook 走通道）',
      viaExcel.ok === true && !viaExcel.unsupported, JSON.stringify(viaExcel.data).slice(0, 60));
    const ghost = await clientJs.WpsAddonChannel.dispatch('setCellValue', { row: 1, col: 1 }, 'excel');
    ok('部分支持的动作（setCellValue）本批未实现 → 仍回落 unsupported:B3（不越批）',
      ghost.ok === false && ghost.unsupported === 'B3');
    uninstallCefQuery();
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('门禁异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
