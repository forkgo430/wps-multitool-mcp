/**
 * 加载项动作分发器的 mock 单测（Linux B2.5 验收项）
 * 原则：不起真 WPS —— 造一个会录像的 mock app 对象，断言每个动作【调用了什么 JSAPI 方法、
 * 赋了什么属性、参数换算对不对】，而不是断言"返回了 mock 值"——防"永远绿"的假测试。
 * 语义依据：docs/agents/linux/S1-REPORT.md §3（逐条明细）/§4（语义差异）/§5（枚举映射），
 * 每段断言前的注释都标了 S1 行号。跑法：node scripts/manual/actions-dispatcher-test.js
 * （verify-all FAST 档包含它）。
 */

const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const ADDON = path.resolve(ROOT, '..', 'wps-multitool-addon-linux');
const wordActions = require(path.join(ADDON, 'js', 'actions', 'word.js'));
const pptActions = require(path.join(ADDON, 'js', 'actions', 'ppt.js'));
const clientJs = require(path.join(ADDON, 'js', 'channel-client.js'));

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};
const eq = (label, got, want) => ok(`${label} = ${JSON.stringify(want)}`, JSON.stringify(got) === JSON.stringify(want),
  `got ${JSON.stringify(got)}`);

/** 属性赋值/读取录像：把 obj 的指定字段包上 getter/setter，读写都进 log */
function watched(obj, fields, log, prefix) {
  const store = {};
  fields.forEach((f) => {
    store[f] = obj[f];
    Object.defineProperty(obj, f, {
      get() { log.push(`${prefix}.${f} (read)`); return store[f]; },
      set(v) { log.push(`${prefix}.${f}=${JSON.stringify(v)}`); store[f] = v; },
      configurable: true,
    });
  });
  return obj;
}

/** 方法调用录像 */
function recording(obj, name, log, prefix, impl) {
  obj[name] = function (...args) {
    log.push(`${prefix}.${name}(${args.map((a) => JSON.stringify(a)).join(', ')})`);
    return impl ? impl(...args) : undefined;
  };
  return obj;
}

const inLog = (log, needle) => log.some((l) => l.includes(needle));

/* ============================================================
 * Word mock（wps 型）
 * ============================================================ */
function makeWordMock(opts) {
  opts = opts || {};
  const log = [];
  const font = watched({ Name: '宋体', Size: 12, Bold: false, Italic: false, Underline: false, Color: 0 },
    ['Name', 'Size', 'Bold', 'Italic', 'Underline', 'Color'], log, 'Font');
  const selFont = watched({ Name: '宋体', Size: 12, Bold: false, Italic: false, Underline: false, Color: 0 },
    ['Name', 'Size', 'Bold', 'Italic', 'Underline', 'Color'], log, 'SelFont');
  const contentFont = watched({ Name: '宋体', Size: 12, Bold: false, Italic: false, Underline: false, Color: 0 },
    ['Name', 'Size', 'Bold', 'Italic', 'Underline', 'Color'], log, 'ContentFont');
  const rangeText = { text: '旧文字' };
  const range = {};
  Object.defineProperty(range, 'Text', {
    get() { log.push('Range.Text (read)'); return rangeText.text; },
    set(v) { log.push(`Range.Text=${JSON.stringify(v)}`); rangeText.text = v; },
  });
  const doc = {
    Name: opts.noDoc ? undefined : 'Doc1',
    FullName: '/tmp/Doc1.docx',
    Path: '/tmp',
    Saved: false,
    ReadOnly: false,
    Content: { Font: contentFont },
    Range: (a, b) => { log.push(`ActiveDocument.Range(${JSON.stringify(a)}, ${JSON.stringify(b)})`); return range; },
  };
  const selection = { Text: '选中文字', Font: selFont };
  const documents = {};
  recording(documents, 'Open', log, 'Documents', (p) => {
    if (opts.openThrows) throw new Error('JSAPI: Open 炸了（mock）');
    return doc;
  });
  const app = {
    ActiveDocument: opts.noDoc ? null : doc,
    Documents: documents,
    Selection: opts.noSelection ? undefined : selection,
  };
  if (opts.selectionTextThrows) {
    Object.defineProperty(app, 'Selection', {
      get() { throw new Error('JSAPI: Selection 炸了（mock）'); },
      configurable: true,
    });
  }
  return { app, log, doc, range };
}

/* ============================================================
 * PPT mock（wpp 型）
 * ============================================================ */
function makeTextChain(log, prefix, text) {
  const chain = {};
  let t = text;
  const textRange = {};
  Object.defineProperty(textRange, 'Text', {
    get() { log.push(`${prefix}.TextRange.Text (read)`); return t; },
    set(v) { log.push(`${prefix}.TextRange.Text=${JSON.stringify(v)}`); t = v; },
  });
  const font = watched({ Size: 18, Name: '宋体' }, ['Size', 'Name'], log, `${prefix}.Font`);
  textRange.Font = font;
  chain.TextFrame = { TextRange: textRange };
  return chain;
}

function makePptMock(opts) {
  opts = opts || {};
  const log = [];
  const mkShape = (i, spec) => {
    const s = { Name: spec.name, Type: spec.type, HasTable: !!spec.hasTable, Table: spec.table };
    watched(s, ['Left', 'Top', 'Width', 'Height'], log, `Shape${i}`);
    if (spec.text !== undefined) Object.assign(s, makeTextChain(log, `Shape${i}`, spec.text));
    return s;
  };
  const tableShape = mkShape(2, {
    name: 'Table 1', type: 19, hasTable: true, text: undefined,
    table: (() => {
      const cells = {};
      return {
        Cell(r, c) {
          log.push(`Table.Cell(${JSON.stringify(r)}, ${JSON.stringify(c)})`);
          const key = r + ',' + c;
          if (!cells[key]) cells[key] = { Shape: makeTextChain(log, `Cell(${r},${c}).Shape`, `旧内容${key}`) };
          return cells[key];
        },
      };
    })(),
  });
  const shapes = [
    mkShape(1, { name: 'TextBox 1', type: 17, text: '初始文本' }),
    tableShape,
    mkShape(3, { name: 'Title 1', type: 13, text: '' }),
  ];
  const shapesObj = {
    get Count() { log.push('Shapes.Count (read)'); return shapes.length; },
    Item(i) {
      log.push(`Shapes.Item(${JSON.stringify(i)})`);
      const s = shapes[i - 1];
      if (!s) throw new Error('JSAPI: Shapes.Item 越界（mock）');
      return s;
    },
    AddTextbox(orientation, l, t, w, h) {
      log.push(`Shapes.AddTextbox(${orientation}, ${l}, ${t}, ${w}, ${h})`);
      const s = mkShape(shapes.length + 1, { name: 'TextBox ' + (shapes.length + 1), type: 17, text: '' });
      shapes.push(s);
      return s;
    },
    AddTable(rows, cols, l, t, w, h) {
      log.push(`Shapes.AddTable(${rows}, ${cols}, ${l}, ${t}, ${w}, ${h})`);
      if (opts.addTableThrows) throw new Error('JSAPI: AddTable 炸了（mock）');
      const s = mkShape(shapes.length + 1, { name: 'Table ' + (shapes.length + 1), type: 19, hasTable: true });
      s.Table = { Cell() { return { Shape: makeTextChain(log, 'NewTable.Cell.Shape', '') }; } };
      shapes.push(s);
      return s;
    },
  };
  // 动画"只做加"的铁律 mock：MainSequence 的 Count 一读就炸（S1 §5：读侧不稳会桥死锁）
  const mainSequence = {};
  Object.defineProperty(mainSequence, 'Count', {
    get() { throw new Error('S1 §5：MainSequence.Count 读侧不稳（读了就桥死锁），动作层禁止读'); },
    configurable: true,
  });
  recording(mainSequence, 'AddEffect', log, 'TimeLine.MainSequence');
  const slide = {
    Shapes: shapesObj,
    Delete() { log.push('Slide.Delete()'); },
    TimeLine: { MainSequence: mainSequence },
    SlideShowTransition: watched({ EntryEffect: 0, Duration: 0 }, ['EntryEffect', 'Duration'], log, 'Transition'),
    Background: {
      Fill: {
        solidCalls: 0,
        Solid() { log.push('Background.Fill.Solid()'); this.solidCalls += 1; },
        ForeColor: watched({ RGB: 0 }, ['RGB'], log, 'Background.Fill.ForeColor'),
      },
    },
    Export(p, f, w, h) { log.push(`Slide.Export(${JSON.stringify(p)}, ${JSON.stringify(f)}${w !== undefined ? ', ' + w + ', ' + h : ''})`); },
  };
  if (opts.placeholderSlide) {
    // addSlide 的 content 落点：PlaceholderFormat.Type===2 的正文占位符
    const ph = mkShape(1, { name: 'Content Placeholder 1', type: 7, text: '' });
    ph.PlaceholderFormat = { Type: 2 };
    slide.Shapes = {
      get Count() { return 1; },
      Item(i) { return ph; },
      AddTextbox: shapesObj.AddTextbox,
      AddTable: shapesObj.AddTable,
    };
    slide.Shapes.HasTitle = false;
    slide._ph = ph;
  }
  const pres = {
    Name: 'Pres1',
    Slides: {
      get Count() { log.push('Slides.Count (read)'); return opts.slideCount !== undefined ? opts.slideCount : 3; },
      Item(i) {
        log.push(`Slides.Item(${JSON.stringify(i)})`);
        if (opts.slidesItemThrows) throw new Error('JSAPI: Slides.Item 炸了（mock）');
        if (opts.slidesItemNull && i === 99) return null;
        return slide;
      },
      Add(position, layout) {
        log.push(`Slides.Add(${JSON.stringify(position)}, ${JSON.stringify(layout)})`);
        if (opts.slidesAddThrows) throw new Error('JSAPI: Slides.Add 炸了（mock）');
        return slide;
      },
    },
  };
  const app = {
    ActivePresentation: opts.noPres ? null : pres,
    Presentations: {
      Add() { log.push('Presentations.Add()'); return { Name: 'Presentation2' }; },
    },
  };
  return { app, log, slide, shapes, pres };
}

/* ============================================================
 * Word 5 动作
 * ============================================================ */
console.log('\n=== Word（wps 型）5 动作 ===');

(function () {
  console.log('\n--- openDocument（S1:134 §3#108）---');
  const m = makeWordMock();
  const r1 = wordActions.openDocument(m.app, { filePath: '/tmp/a.docx', path: '/tmp/a.docx' });
  ok('正路 ok:true', r1.ok === true, JSON.stringify(r1.data));
  ok('调了 Documents.Open 且路径参数正确', inLog(m.log, 'Documents.Open("/tmp/a.docx")'), m.log.join(' | ').slice(0, 100));
  eq('回包带文档名', r1.data && r1.data.documentName, 'Doc1');
  const r2 = wordActions.openDocument(m.app, {});
  ok('缺 path → ok:false 且不再调 Open', r2.ok === false && !inLog(m.log.slice(m.log.indexOf('Documents.Open') + 1) || [], 'Documents.Open("/tmp/a.docx")') === false && m.log.filter((l) => l.includes('Documents.Open')).length === 1, r2.error);
  const m2 = makeWordMock({ openThrows: true });
  const r3 = wordActions.openDocument(m2.app, { path: '/tmp/a.docx' });
  ok('JSAPI 抛错 → ok:false 且 error 是原文', r3.ok === false && /Open 炸了/.test(r3.error || ''), (r3.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- getActiveDocumentState（S1:129 §3#103）---');
  const m = makeWordMock();
  const r1 = wordActions.getActiveDocumentState(m.app, { app: 'word' });
  ok('正路 ok:true 且字段齐', r1.ok === true && r1.data.name === 'Doc1' && r1.data.hasPath === true && r1.data.saved === false, JSON.stringify(r1.data).slice(0, 90));
  const m2 = makeWordMock({ noDoc: true });
  const r2 = wordActions.getActiveDocumentState(m2.app, {});
  ok('空壳 → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- getSelectedText（S1:303 §3#277）---');
  const m = makeWordMock();
  const r1 = wordActions.getSelectedText(m.app, {});
  eq('正路取回选区文本', r1.ok && r1.data && r1.data.text, '选中文字');
  const m2 = makeWordMock({ selectionTextThrows: true });
  const r2 = wordActions.getSelectedText(m2.app, {});
  ok('Selection 抛错 → ok:false 且 error 是原文', r2.ok === false && /Selection 炸了/.test(r2.error || ''), (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- replaceRange（S1:163 §3#137）---');
  const m = makeWordMock();
  const r1 = wordActions.replaceRange(m.app, { startPos: 1, endPos: 3, text: '新文字' });
  ok('正路 ok:true', r1.ok === true, JSON.stringify(r1.data));
  ok('调了 ActiveDocument.Range(1, 3)', inLog(m.log, 'ActiveDocument.Range(1, 3)'), m.log.join(' | ').slice(0, 100));
  ok('对 Range.Text 赋了新值', inLog(m.log, 'Range.Text="新文字"'));
  const r2 = wordActions.replaceRange(m.app, { startPos: 1, endPos: 3 });
  ok('缺 text → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
  ok('错误路上没有第二次 Range 调用', m.log.filter((l) => l.includes('ActiveDocument.Range')).length === 1);
})();

(function () {
  console.log('\n--- setFont（S1:145 §3#119 + ps1 setFont case）---');
  const m = makeWordMock();
  const r1 = wordActions.setFont(m.app, { fontName: 'SimSun', fontSize: 24, bold: true, color: '#00FF00' });
  ok('正路 ok:true', r1.ok === true);
  ok('Selection.Font.Name 赋值', inLog(m.log, 'SelFont.Name="SimSun"'));
  ok('Selection.Font.Size=24', inLog(m.log, 'SelFont.Size=24'));
  ok('Selection.Font.Bold=true', inLog(m.log, 'SelFont.Bold=true'));
  ok('#00FF00 → Font.Color=65280（COM RGB 编码，S0 #18 背书）', inLog(m.log, 'SelFont.Color=65280'));
  const mAll = makeWordMock();
  const r2 = wordActions.setFont(mAll.app, { fontName: 'KaiTi', range: 'all' });
  ok("range='all' 走 Content.Font（不是选区）", inLog(mAll.log, 'ContentFont.Name="KaiTi"') && !inLog(mAll.log, 'SelFont.Name="KaiTi"'));
  const mPin = makeWordMock();
  const r3 = wordActions.setFont(mPin.app, { fontName: 'X', documentName: 'Doc1' });
  ok('钉文档+选区范围 → 拒绝（ps1 同款护栏）', r3.ok === false && /改到别人/.test(r3.error || ''), (r3.error || '').slice(0, 60));
  const r4 = wordActions.setFont(mPin.app, { color: 'purple' });
  ok('非法色名 → ok:false 且未赋 Color', r4.ok === false && !inLog(mPin.log, 'Color=') , r4.error);
})();

/* ============================================================
 * PPT 15 动作
 * ============================================================ */
console.log('\n=== PPT（wpp 型）15 动作 ===');

(function () {
  console.log('\n--- createPresentation（S1:309 §3#283）---');
  const m = makePptMock();
  const r1 = pptActions.createPresentation(m.app, {});
  ok('正路 ok:true 且调了 Presentations.Add()', r1.ok === true && inLog(m.log, 'Presentations.Add()'));
  eq('回包带新演示名', r1.data && r1.data.name, 'Presentation2');
  const r2 = pptActions.createPresentation({ ActivePresentation: null, Presentations: { Add() { throw new Error('JSAPI: Add 炸了（mock）'); } } }, {});
  ok('JSAPI 抛错 → ok:false 原文', r2.ok === false && /Add 炸了/.test(r2.error || ''), (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- getSlideCount（S1:202 §3#176）---');
  const m = makePptMock({ slideCount: 7 });
  const r1 = pptActions.getSlideCount(m.app, {});
  ok('正路读 Slides.Count 且回数字', r1.ok === true && r1.data.count === 7 && inLog(m.log, 'Slides.Count (read)'));
  const m2 = makePptMock({ noPres: true });
  const r2 = pptActions.getSlideCount(m2.app, {});
  ok('无活动演示 → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- addSlide（S1:210 §3#184 + ps1 addSlide 映射表）---');
  const m = makePptMock({ slideCount: 3, placeholderSlide: true });
  const r1 = pptActions.addSlide(m.app, { layout: 'blank', title: 'T', content: 'C' });
  ok('正路 ok:true', r1.ok === true, JSON.stringify(r1.data));
  ok("layout 'blank' → 编号 12（ps1 $layouts 表）", inLog(m.log, 'Slides.Add(4, 12)'));
  ok('缺省 position = Count+1 = 4', inLog(m.log, 'Slides.Add(4, 12)'));
  const m2 = makePptMock({ slideCount: 2, placeholderSlide: true });
  const r2 = pptActions.addSlide(m2.app, { layout: 'title_content', content: '正文内容' });
  ok("'title_content' → 2 且落正文占位符（PlaceholderFormat.Type===2）",
    inLog(m2.log, 'Slides.Add(3, 2)') && inLog(m2.log, 'Shape1.TextRange.Text="正文内容"'),
    m2.log.filter((l) => l.includes('TextRange.Text')).join(' | ').slice(0, 80));
  const m3 = makePptMock({ slideCount: 5, placeholderSlide: true });
  const r3 = pptActions.addSlide(m3.app, { layout: '不存在的版式' });
  ok('认不出的版式回落 2（ps1 同款）', r3.ok === true && inLog(m3.log, 'Slides.Add(6, 2)'));
  const m4 = makePptMock({ slidesAddThrows: true });
  const r4 = pptActions.addSlide(m4.app, { layout: 'blank' });
  ok('Slides.Add 抛错 → ok:false 原文', r4.ok === false && /Add 炸了/.test(r4.error || ''), (r4.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- deleteSlide（S1:198 §3#172）---');
  const m = makePptMock();
  const r1 = pptActions.deleteSlide(m.app, { slideIndex: 2 });
  ok('正路调 Slides.Item(2).Delete()', r1.ok === true && inLog(m.log, 'Slides.Item(2)') && r1.data.deleted === 2);
  const r2 = pptActions.deleteSlide(m.app, {});
  ok('缺 slideIndex → ok:false 且不多调 Item', r2.ok === false && m.log.filter((l) => l.includes('Slides.Item')).length === 1, r2.error);
})();

(function () {
  console.log('\n--- getShapes（S1:222 §3#196）---');
  const m = makePptMock();
  const r1 = pptActions.getShapes(m.app, { slideIndex: 1 });
  ok('正路返回 3 个形状', r1.ok === true && r1.data.shapes.length === 3 && r1.data.count === 3);
  eq('index 从 1 计且字段齐（S1 证据：#1:T14 #2:T14/#19 #3:T13）',
    r1.data.shapes.map((s) => [s.index, s.name]), [[1, 'TextBox 1'], [2, 'Table 1'], [3, 'Title 1']]);
  const m2 = makePptMock({ slidesItemThrows: true });
  const r2 = pptActions.getShapes(m2.app, { slideIndex: 9 });
  ok('Item 抛错 → ok:false 原文', r2.ok === false && /Item 炸了/.test(r2.error || ''), (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- setShapeText（S1:226 §3#200）---');
  const m = makePptMock();
  const r1 = pptActions.setShapeText(m.app, { slideIndex: 1, shapeIndex: 1, text: '新形状文字' });
  ok('正路对 Item(1).TextFrame.TextRange.Text 赋值', r1.ok === true && inLog(m.log, 'Shapes.Item(1)') && inLog(m.log, 'Shape1.TextRange.Text="新形状文字"'));
  const r2 = pptActions.setShapeText(m.app, { slideIndex: 1 });
  ok('缺 text → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
  const r3 = pptActions.setShapeText(m.app, { slideIndex: 1, shapeIndex: 99, text: 'x' });
  ok('形状越界 → ok:false 原文', r3.ok === false && /越界/.test(r3.error || ''), (r3.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- addTextBox（S1:211 §3#185 + ps1 addTextBox 缺省）---');
  const m = makePptMock();
  const r1 = pptActions.addTextBox(m.app, { slideIndex: 1, text: '框内容', fontSize: 20 });
  ok('orientation 固定 1，缺省 100/100/400/50（ps1 同款）',
    inLog(m.log, 'Shapes.AddTextbox(1, 100, 100, 400, 50)'), m.log.filter((l) => l.includes('AddTextbox')).join(''));
  ok('text 与 fontSize 落到新文本框', inLog(m.log, 'Shape4.TextRange.Text="框内容"') && inLog(m.log, 'Shape4.Font.Size=20'));
  const m2 = makePptMock();
  const r2 = pptActions.addTextBox(m2.app, { slideIndex: 1, left: 10, top: 20, width: 30, height: 40 });
  ok('显式坐标直传', inLog(m2.log, 'Shapes.AddTextbox(1, 10, 20, 30, 40)'));
  const r3 = pptActions.addTextBox(m2.app, { slideIndex: 1 });
  ok('不传 text 也不炸（文本可选）', r3.ok === true);
})();

(function () {
  console.log('\n--- setShapePosition（S1:227 §3#201）---');
  const m = makePptMock();
  const r1 = pptActions.setShapePosition(m.app, { slideIndex: 1, shapeIndex: 1, left: 300, top: 250 });
  ok('只设给了的维度：Left/Top 赋值', inLog(m.log, 'Shape1.Left=300') && inLog(m.log, 'Shape1.Top=250'));
  ok('Width/Height 未赋值（录像里没有）', !inLog(m.log, 'Shape1.Width=') && !inLog(m.log, 'Shape1.Height='));
  const r2 = pptActions.setShapePosition(m.app, { slideIndex: 1, left: 1 });
  ok('形状没给 → ok:false', r2.ok === false, (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- insertPptTable（S1:233 §3#207）---');
  const m = makePptMock();
  const r1 = pptActions.insertPptTable(m.app, { slideIndex: 1, rows: 4, cols: 2, left: 5, top: 6 });
  ok('AddTable(4, 2, 5, 6, 400, 200)（缺省尺寸 ps1 同款）',
    inLog(m.log, 'Shapes.AddTable(4, 2, 5, 6, 400, 200)'), m.log.filter((l) => l.includes('AddTable')).join(''));
  eq('tableIndex 按 HasTable 计数（页上已有 1 表，新表是第 2 个）', r1.data && r1.data.tableIndex, 2);
  const m2 = makePptMock({ addTableThrows: true });
  const r2 = pptActions.insertPptTable(m2.app, { slideIndex: 1, rows: 2, cols: 2 });
  ok('AddTable 抛错 → ok:false 原文', r2.ok === false && /AddTable 炸了/.test(r2.error || ''), (r2.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- getPptTableCell / setPptTableCell（S1:235/234 §3#209/#208）---');
  const m = makePptMock();
  const r1 = pptActions.getPptTableCell(m.app, { slideIndex: 1, tableIndex: 1, row: 2, col: 3 });
  ok('正路走 Table.Cell(2, 3).Shape.TextFrame.TextRange.Text', r1.ok === true && inLog(m.log, 'Table.Cell(2, 3)'));
  eq('读回单元格文本', r1.data && r1.data.value, '旧内容2,3');
  const r2 = pptActions.getPptTableCell(m.app, { slideIndex: 1, tableIndex: 9, row: 1, col: 1 });
  ok('tableIndex 越界 → ok:false', r2.ok === false && /找不到第 9 个表格/.test(r2.error || ''), (r2.error || '').slice(0, 60));
  const r3 = pptActions.setPptTableCell(m.app, { slideIndex: 1, tableIndex: 1, row: 1, col: 1, text: '新格值' });
  ok('写路径对 Cell(1,1).Shape…Text 赋值', r3.ok === true && inLog(m.log, 'Cell(1,1).Shape.TextRange.Text="新格值"'),
    m.log.filter((l) => l.includes('Cell(1,1)')).join(' | ').slice(0, 90));
  const r4 = pptActions.setPptTableCell(m.app, { slideIndex: 1, tableIndex: 1, row: 1, col: 1 });
  ok('缺 text → ok:false', r4.ok === false, (r4.error || '').slice(0, 60));
})();

(function () {
  console.log('\n--- addAnimation（S1:252/271/272 §3#226/#245/#246 + §5 行 355）---');
  const m = makePptMock();
  const r1 = pptActions.addAnimation(m.app, { slideIndex: 1, shapeIndex: 1, effect: 'fadeIn' });
  ok("effect 'fadeIn' → AddEffect(shape, 10, 0)（映射表抄 ps1；level 恒 0 防 ps1 实测的双效果坑）",
    inLog(m.log, 'TimeLine.MainSequence.AddEffect'), m.log.filter((l) => l.includes('AddEffect')).join(' | '));
  ok('回包没有 count 字段（S1 §5：Count 读侧不稳，刻意不读）', r1.ok === true && !('count' in (r1.data || {})));
  // "只做加"的强断言：mock 的 MainSequence.Count 一读就炸——实现若读了，r1 就不会是 ok:true
  ok('全程没碰 MainSequence.Count（读了就会炸）', r1.ok === true);
  const r2 = pptActions.addAnimation(m.app, { slideIndex: 1, shapeIndex: 1, effect: 29 });
  ok('数字 effect 直通（S1 §5：编号恒等）', r2.ok === true && inLog(m.log, 'AddEffect') && /, 29, 0\)/.test(m.log.filter((l) => l.includes('AddEffect')).pop() || ''));
  const r3 = pptActions.addAnimation(m.app, { slideIndex: 1, shapeIndex: 1, effect: 'magicFly' });
  ok('认不出的效果名 → 明确报错（ps1 同款，不猜）', r3.ok === false && /不认识的 effect/.test(r3.error || ''), (r3.error || '').slice(0, 60));
  const r4 = pptActions.addAnimation(m.app, { slideIndex: 1, effect: 'fadeIn' });
  ok('形状没给 → ok:false（ps1 实测过的坏对象坑：Item(null) 会毁整页）', r4.ok === false && /找不到要加动画的形状/.test(r4.error || ''));
  const effectCalls = m.log.filter((l) => l.includes('AddEffect')).length;
  ok('错误路径没有产生任何 AddEffect 调用', effectCalls === 2, `实际 ${effectCalls} 次`);
})();

(function () {
  console.log('\n--- setSlideTransition（S1:253 §3#227 + §5 行 354）---');
  const m = makePptMock();
  const r1 = pptActions.setSlideTransition(m.app, { slideIndex: 1, effect: 'push', duration: 2 });
  ok("'push' → EntryEffect=2（映射表抄 ps1）", inLog(m.log, 'Transition.EntryEffect=2'));
  ok('Duration 赋值', inLog(m.log, 'Transition.Duration=2'));
  const r2 = pptActions.setSlideTransition(m.app, { slideIndex: 1, effect: 'none' });
  ok("'none' → 0（取消切换不再炸，ps1 注释里的历史坑）", inLog(m.log, 'Transition.EntryEffect=0'));
  const r3 = pptActions.setSlideTransition(m.app, { slideIndex: 1, effect: '77' });
  ok('数字串直通 → 77', inLog(m.log, 'Transition.EntryEffect=77'));
  const r4 = pptActions.setSlideTransition(m.app, { slideIndex: 1, effect: 'starWipe' });
  ok('认不出的切换 → ok:false 且未赋值', r4.ok === false && m.log.filter((l) => l.includes('EntryEffect=')).length === 3, r4.error);
})();

(function () {
  console.log('\n--- setSlideBackground（S1:254 §3#228 + S0 #18）---');
  const m = makePptMock();
  const r1 = pptActions.setSlideBackground(m.app, { slideIndex: 1, color: '#00FF00' });
  ok('先 Solid() 再 ForeColor.RGB', inLog(m.log, 'Background.Fill.Solid()') && inLog(m.log, 'Background.Fill.ForeColor.RGB=65280'));
  const r2 = pptActions.setSlideBackground(m.app, { slideIndex: 1, color: 'zzz' });
  ok('非法颜色 → ok:false（ps1 是静默跳过仍回 success——这是与 ps1 的有意差异）', r2.ok === false, (r2.error || '').slice(0, 60));
  const r3 = pptActions.setSlideBackground(m.app, { slideIndex: 1, imagePath: '/tmp/x.png' });
  ok('图片路径本轮未实现 → 明确报"纯色"，不静默成功', r3.ok === false && /纯色/.test(r3.error || ''), (r3.error || '').slice(0, 60));
  const solidCalls = m.log.filter((l) => l.includes('Solid()')).length;
  ok('两条错误路径都没有碰 Fill', solidCalls === 1, `Solid 实际 ${solidCalls} 次`);
})();

(function () {
  console.log('\n--- exportSlideAsImage（S1:232 §3#206 + S0 #19）---');
  const m = makePptMock();
  const r1 = pptActions.exportSlideAsImage(m.app, { slideIndex: 1, outputPath: '/tmp/s1.png' });
  ok("Export(path, 'PNG') 两参调用", inLog(m.log, 'Slide.Export("/tmp/s1.png", "PNG")'));
  const r2 = pptActions.exportSlideAsImage(m.app, { slideIndex: 1, path: '/tmp/s2.png', format: 'JPG', width: 1334, height: 750 });
  ok('path 别名键也认 + 带尺寸四参调用（S0 实测带尺寸成功）',
    inLog(m.log, 'Slide.Export("/tmp/s2.png", "JPG", 1334, 750)'), m.log.filter((l) => l.includes('Export')).join(' | '));
  const r3 = pptActions.exportSlideAsImage(m.app, { slideIndex: 1 });
  ok('缺 path → ok:false 且 Export 未多调', r3.ok === false && m.log.filter((l) => l.includes('Slide.Export')).length === 2, r3.error);
})();

/* ============================================================
 * 分发器接线（channel-client.js 的注册表路由）
 * ============================================================ */
console.log('\n=== 分发器接线（注册表查找 + 回落）===');
(function () {
  // channel-client.js 在 CommonJS 里挂 module.exports；把动作注册表挂上去，
  // 等价模拟浏览器里 manifest 按序加载（channel-client → actions/*.js）后的 window 状态
  clientJs.WpsAddonActionsWord = wordActions;
  clientJs.WpsAddonActionsPpt = pptActions;
  const d = clientJs.WpsAddonChannel.dispatch;
  const ping = d('ping', {}, 'word');
  ok("ping 仍真实现", ping.ok === true && ping.data.pong === true);
  const wc = d('wireCheck', {}, 'wpp');
  ok("wireCheck 仍真实现（文案升级 B2.5）", wc.ok === true && /B2\.5/.test(wc.data.message || ''));
  const ghost = d('setValue', { x: 1 }, 'word');
  ok('注册表外的动作仍回落 unsupported:B3（显式拒绝）', ghost.ok === false && ghost.unsupported === 'B3');
  // Node 侧 hostApp 拿不到 window.Application → 注册表内动作应报"拿不到宿主"，而不是静默成功
  const host = d('getSlideCount', {}, 'ppt');
  ok('注册表命中但无宿主 → 明确报错（不静默成功）', host.ok === false && /宿主 Application/.test(host.error || ''), (host.error || '').slice(0, 70));
})();

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
