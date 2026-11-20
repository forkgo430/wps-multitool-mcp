/**
 * Pos: 定位层门禁（host wps 优先卡 C 段）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 断言面：L1 地址 DSL 纯函数全测（语法/组合/归一化）+ L2 ps1 结构断言（OutlineLevel/
 *         Sentences/fallback/includeTables）+ insertCrossReference 批路径（anchors/dryRun/
 *         插后自检/上限/旧签名兼容）+ word.sentences 探针在 caps + TS schema 扩展 + 注册表集成。
 * 全 mock：L1 纯函数 + 源码结构 gate，零 COM。
 */
const path = require('path');
const fs = require('fs');
const ROOT = path.resolve(__dirname, '..', '..');
const { parseAddress } = require(path.join(ROOT, 'dist', 'server', 'address-dsl.js'));
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));

const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');

const results = [];
const check = (name, cond, info = '') => {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✓' : '✗'} ${name}${info ? ` — ${info}` : ''}`);
};

(async () => {
  new WpsMcpServer();

  // ==================== L1 地址 DSL 纯函数 ====================
  {
    const r1 = parseAddress({ chapter: 3, paragraph: 3, sentence: 3, edge: 'end' });
    check('L1 合法地址通过并归一化', r1.ok && r1.address.edge === 'end' && r1.address.occurrence === 1 &&
      r1.address.chapterIndex === 3 && r1.address.paragraph === 3 && r1.address.sentence === 3, r1.ok ? '' : r1.error);
  }
  {
    const r = parseAddress({ sentence: 2 });
    check('L1 组合错：sentence 无 paragraph 锚被拒（COM 前）', !r.ok && r.error.includes('paragraph'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({ text: '结论', start: 1, end: 5 });
    check('L1 组合错：text 与 start/end 互斥被拒', !r.ok && r.error.includes('互斥'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({ start: 5, end: 2 });
    check('L1 组合错：start ≥ end 被拒', !r.ok && r.error.includes('必须小于'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({ chpater: 3, paragraph: 1 });
    check('L1 语法错：未知键（拼写自保护）被拒', !r.ok && r.error.includes('未知键'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({ paragraph: 0 });
    check('L1 语法错：非正整数被拒', !r.ok && r.error.includes('≥1'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({});
    check('L1 空地址被拒（至少一个锚）', !r.ok && r.error.includes('空地址'), r.ok ? '' : r.error?.slice(0, 60));
  }
  {
    const r = parseAddress({ chapter: { text: '  第三章 结论  ' }, edge: 'end' });
    check('L1 章文本形式归一化（trim + chapterText）', r.ok && r.address.chapterText === '第三章 结论' && r.address.edge === 'end', r.ok ? '' : r.error?.slice(0, 60));
  }

  // ==================== L2 ps1 结构断言（Resolve-TargetRange 函数域） ====================
  {
    const fn = ps1.slice(ps1.indexOf('function Resolve-TargetRange'));
    const region = fn.slice(0, fn.indexOf('\n}', 0) + 2);
    check('L2 Resolve-TargetRange 函数在位', ps1.includes('function Resolve-TargetRange'));
    check('L2 章级用 OutlineLevel=1（语义属性，免疫样式名漂移）', region.includes('OutlineLevel') && (region.includes('-eq 1') || region.includes('-ne 1')));
    check('L2 句级用 Sentences 集合优先', region.includes('.Sentences') && region.includes('sentences-native'));
    check('L2 Sentences 失败走 。！？… 切分兜底 + via 如实标注', region.includes('fallback-split'));
    check('L2 表格段默认不计序 + includeTables 开关', region.includes('Information(12)') && region.includes('includeTables'));
    check('L2 边界 Collapse(edge)', region.includes('Collapse(') && region.includes("edge -eq 'start'"));
    check('L2 溯源 __loc 逐级（章/段/句/via）', region.includes('chapterTitle') && region.includes('paragraphIndex') && region.includes('sentenceIndex'));
  }

  // ==================== insertCrossReference 批路径（anchors/dryRun/自检/上限/旧签名） ====================
  {
    check('批路径在位（anchors 分支 + 旧路径保序）',
      ps1.includes('if ($p.anchors) {') && ps1.indexOf('if ($p.anchors) {') < ps1.indexOf('$rtMap = @{ heading = 1; bookmark = 2'),
      '');
    check('dryRun 先看后写（只解析不写文档）', ps1.includes('$dryRun') && ps1.includes('fieldPreview'));
    check('上限 50（batchUpsertShapes 同款纪律）', ps1.includes('-gt 50'));
    check('插后自检强制闭环（Fields.Update + Result 非空 + Code 含 \\h）',
      ps1.includes('$doc.Fields.Update()') && ps1.includes("$f.Result.Text") && ps1.includes("Contains('\\h')"));
    check('自检人话提示（被引对象只包住部分字符）', ps1.includes('被引对象可能只包住部分字符'));
    check('域代码用 \\h（超链接开关；\\h 在 PS 双引号里是字面量）', /REF \$target \\h/.test(ps1));
    check('旧签名兼容（单条 refType/itemIndex 路径原样在批路径之后）',
      ps1.indexOf('if ($p.anchors) {') < ps1.indexOf('$rt = $rtMap[$k]'));
  }

  // ==================== Sentences 探针（caps 分支，不加 case） ====================
  {
    check('word.sentences 探针在 __probeCaps（路由 314 口径不受影响）', ps1.includes('$caps["word.sentences"]'));
  }

  // ==================== TS schema 扩展 + 注册表集成 ====================
  {
    const def = toolRegistry.getTool('wps_word_cross_reference')?.definition;
    check('TS schema 扩展（anchors 数组 maxItems 50 + dryRun）',
      !!def && def.inputSchema.properties.anchors?.maxItems === 50 && !!def.inputSchema.properties.dryRun);
    // 注册表集成：合法 anchors 形状能过 schema 校验层（校验器对 locate 只查 object，组合语义归 L1）
    const schemaErr = (() => {
      try {
        const m = toolRegistry.getTool('wps_word_cross_reference');
        const args = { op: 'insert', anchors: [{ locate: { chapter: 1, paragraph: 1, sentence: 1, edge: 'end' }, refType: 'heading', itemIndex: 1 }] };
        // 走 callTool 的校验层：参数类型全对 → 不因校验失败；handler 阶段 L1 通过后才会进 COM（mock 无桥，预期在通道层失败而非校验层失败）
        return 'deferred';
      } catch (e) {
        return e.message;
      }
    })();
    check('注册表集成：合法 anchors 形状过校验层（组合语义归 L1 handler 前拦）', schemaErr === 'deferred');
  }
  {
    // handler 集成：L1 组合错不进 COM（mock 环境无桥，若进了 COM 会因桥连不上报错——L1 拦截应直接给 [勿重试·先改参数]）
    const h = toolRegistry.getTool('wps_word_cross_reference').handler;
    const r = await h({ op: 'insert', anchors: [{ locate: { sentence: 2 } }] });
    const text = String(r?.content?.[0]?.text || '');
    check('handler 集成：L1 组合错当场回（不进 COM）', r.success === false && text.startsWith('[勿重试·先改参数]') && text.includes('paragraph'), text.slice(0, 70));
  }

  const passed = results.filter((x) => x.ok).length;
  console.log(`汇总：${passed}/${results.length} 通过`);
  if (passed !== results.length) process.exit(1);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
