/**
 * Pos: 定位层真机四锚探针 v2（文本断言版）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 锚①：第三章第三段第三句句末——dryRun 预览 → 真插 → 域结果三对账
 * 锚②：29 处批量一次插完
 * 锚③：坏书签 → 自检如实红 + 人话提示
 * 锚④：Sentences 探针结论入档（via=sentences-native / fallback-split）
 * 锚⑤：清稿
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-locator-anchors-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '180000';

const path = require('path');
const fs = require('fs');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '');

let pass = 0, fail = 0;
const ok = (label, cond, extra = '') => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

(async () => {
  console.log('== 定位层真机四锚 ==');

  // ---- 造场：三章节（每章 3 段、每段 3 句）----
  const cr = await call('wps_word_create_document', {});
  if (cr.success === false) { console.log('造场失败：', txt(cr).slice(0, 120)); process.exit(1); }
  const paras = [];
  for (let c = 1; c <= 3; c++) {
    paras.push({ kind: 'heading', content: `第${['一', '二', '三'][c - 1]}章 测试章标题${c}` });
    for (let pIdx = 1; pIdx <= 3; pIdx++) {
      paras.push({ kind: 'para', content: `第${c}章第${pIdx}段的首句内容。这是第二句，用来凑满段内句序。这是第三句，句末是插入点。` });
    }
  }
  // 造场 v8：全文平铺（\r 产生真实段落）+ 两步套样式：
  // ①先全文档套'正文'（一切压回正文级 lv=10）②再倒序逐章套'标题 1'（覆盖为 lv=1）
  // 两步分离避免交错偏移漂移；倒序套章标题确保前方偏移稳定
  const chapterLines = paras.filter((x) => x.kind === 'heading').map((x) => x.content);
  const allText = paras.map((x) => x.content).join('\r') + '\r';
  const ins = await call('wps_word_insert_text', { text: allText, position: 'end' });
  if (ins.success === false) throw new Error('全文插入失败: ' + txt(ins).slice(0, 80));
  {
    const dump = await call('wps_word_get_document_text', {});
    const bodyText = txt(dump).split('\n\n').slice(1).join('\n\n');
    // ①全文档套正文（一切压回 lv=10；Range 覆盖全部插入文本，不拆段）
    await call('wps_word_apply_style', { style_name: '正文', range: { start: 0, end: allText.length + 2 } });
    // ②倒序逐章套标题（后面的先套，前方偏移不受影响）
    let outlineOk = false;
    for (let base = 0; base <= 2 && !outlineOk; base++) {
      for (let ci = chapterLines.length - 1; ci >= 0; ci--) {
        const line = chapterLines[ci];
        const at = bodyText.indexOf(line);
        if (at < 0) throw new Error('章行在 dump 中未找到: ' + line);
        await call('wps_word_apply_style', { style_name: '标题 1', range: { start: base + at, end: base + at + line.length + 1 } });
      }
      const o = txt(await call('wps_word_get_outline', {}));
      const allFound = chapterLines.every((line) => o.includes(line));
      if (allFound) {
        outlineOk = true;
        console.log(`  造场：两步套样式成功（base=${base}，3 章 + 正文）`);
      }
    }
    if (!outlineOk) throw new Error('造场失败：章样式基座 3 轮仍未让 getOutline 含全部 3 个章标题');
  }

  const structDump = await call('wps_word_get_document_text', {});
  console.log('造场结构:', txt(structDump).slice(0, 260));
  const oDump = await call('wps_word_get_outline', {});
  console.log('造场大纲:', txt(oDump).slice(0, 200));
  const targets = await call('wps_word_cross_reference', { op: 'listTargets' });
  const tText = txt(targets);
  const headingCount = (tText.match(/heading: (\d+) 个/) || [])[1];
  console.log(`  可引用 heading 项：${headingCount || '?'} 个`);
  const refIdx = Math.min(3, Math.max(1, Number(headingCount || 1)));

  // ---- 锚④（先探）：via 字段 ----
  const dry = await call('wps_word_cross_reference', {
    op: 'insert', dryRun: true,
    anchors: [{ locate: { chapter: 3, paragraph: 3, sentence: 3, edge: 'end' }, refType: 'heading', itemIndex: refIdx }],
  });
  const dryText = txt(dry);
  const via = (dryText.match(/\[via=([a-z-]+)\]/) || [])[1] || '';
  ok('锚④ Sentences 探针结论入档（via 明确：sentences-native 或 fallback-split）',
    via === 'sentences-native' || via === 'fallback-split', `via=${via}`);
  const evidence = { ts: new Date().toISOString(), anchors: { via } };

  // ---- 锚①：dryRun → 真插 → 三对账 ----
  ok('锚①a dryRun 预览返回（未写文档）', dryText.includes('dryRun 预览') && dryText.includes('未写文档'), dryText.slice(0, 90));
  ok('锚①b 落点是第三章（章标题原文）', dryText.includes('第3章') && dryText.includes('测试章标题3'), '');
  ok('锚①c 落点是段3·句3', /段3·句3/.test(dryText), '');
  ok('锚①d 句全文含"第三句…插入点"', dryText.includes('第三句') && dryText.includes('插入点'), '');

  const real = await call('wps_word_cross_reference', {
    op: 'insert',
    anchors: [{ locate: { chapter: 3, paragraph: 3, sentence: 3, edge: 'end' }, refType: 'heading', itemIndex: refIdx }],
  });
  const realText = txt(real);
  ok('锚①e 真插 filled=1 failed=0', realText.includes('已插 1 项，失败 0 项'), realText.slice(0, 90));
  ok('锚①f 域结果非空（插后自检通过）', realText.includes('域结果：'), '');
  ok('锚①g 三对账：dryRun 与真插落点同一地址（同一章段句）',
    dryText.includes('第3章') && realText.includes('第3章') && /段3·句3/.test(realText), '');
  evidence.anchors.anchor1 = { dry: dryText.slice(0, 400), real: realText.slice(0, 400), via };

  // ---- 锚②：29 处批量一次插完 ----
  const anchors29 = [];
  // text 寻址（不依赖段位——field 插入改变段落结构后仍可靠）：
  // 29 个唯一文本片段 = 9 body paras × 3 sentences + 3 chapter headings − 1 overlap
  const allParaTexts = paras.filter((x) => x.kind === 'para').map((x) => x.content);
  const fragments = [];
  for (const pt of allParaTexts) {
    const sents = pt.split(/(?<=。)/);
    for (const sent of sents) { if (sent.trim()) fragments.push(sent.trim().replace(/。$/, '')); }
  }
  for (const cl of chapterLines) { fragments.push(cl); }
  for (let i = 0; i < 29; i++) {
    const frag = fragments[i % fragments.length];
    const occ = Math.floor(i / fragments.length) + 1;
    anchors29.push({ locate: { text: frag, occurrence: occ, edge: 'end' }, refType: 'heading', itemIndex: refIdx });
  }
  const batch = await call('wps_word_cross_reference', { op: 'insert', anchors: anchors29 });
  const batchText = txt(batch);
  ok('锚② 29 处批量一次插完（filled=29）', batchText.includes('已插 29 项，失败 0 项'), batchText.slice(0, 120));
  evidence.anchors.anchor2 = { summary: (batchText.match(/已插 \d+ 项，失败 \d+ 项/) || [''])[0] };
  console.log('锚② FULL:', batchText.slice(0, 1100));

  // ---- 锚③：坏书签（空段书签 → REF 结果空 → 自检红 + 人话）----
  let anchor3done = false;
  try {
    await call('wps_word_insert_text', { text: '\n' });
    const bm = await call('wps_word_insert_bookmark', { name: 'BadBookmark' });
    const tgt2 = await call('wps_word_cross_reference', { op: 'listTargets' });
    const bms = (txt(tgt2).match(/bookmark: (\d+) 个/) || [])[1];
    if (Number(bms) >= 1) {
      // 找 BadBookmark 在 bookmark 列表里的序号：listTargets 文本截断了，逐个试不可取——
      // 直接扫 bookmark 列表工具拿顺序
      const bmList = await call('wps_word_get_bookmarks', {});
      const bmText = txt(bmList);
      const order = bmText.split('\n').map((l) => l.trim()).filter((l) => l.length > 0);
            const idx = 0; // 只有 1 个书签（BadBookmark），可引用项序号=1
      if (Number(bms) >= 1) {
        const bad = await call('wps_word_cross_reference', {
          op: 'insert',
          anchors: [{ locate: { paragraph: 1, sentence: 1, edge: 'end' }, refType: 'bookmark', itemIndex: idx + 1 }],
        });
        const badText = txt(bad);
        const isRed = badText.includes('失败') && badText.includes('被引对象');
        ok('锚③ 坏书签 → 自检如实红 + 人话提示', isRed, badText.slice(0, 140));
        anchor3done = true;
        evidence.anchors.anchor3 = { text: badText.slice(0, 400) };
      }
    }
  } catch (e) {
    console.log('  ⚠ 锚③ 异常：', e.message);
  }
  if (!anchor3done) {
    fail++;
    console.log('  ✗ 锚③ 坏书签造场未成（bookmark 列表未含 BadBookmark 或插入失败）——如实记录');
    evidence.anchors.anchor3 = { skipped: true };
  }

  // ---- 锚⑤：清稿 ----
  const cl = await call('wps_word_close_document', { save: false });
  ok('锚⑤ 清稿（不保存关闭）', cl.success !== false, txt(cl).slice(0, 60));

  const evFile = path.join(ROOT, 'scripts', 'manual', 'locator-anchors-evidence.json');
  fs.writeFileSync(evFile, JSON.stringify(evidence, null, 2), 'utf8');
  console.log(`证据：${evFile}`);
  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch(async (e) => {
  // 失败路径也要清稿：探针文档泄漏会让下一轮的 Resolve-Doc 落到 Active（多文档混沌的根因之一）
  console.error('PROBE FAIL:', e.message || e);
  try { await call('wps_word_close_document', { save: false }); console.log('失败路径已尝试关稿（未保存）'); } catch { /* 已关/无文档 */ }
  process.exit(1);
});
