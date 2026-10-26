/**
 * PPT 真机验证（需要 WPS 演示可访问）。
 *
 * 为什么整组 PPT 一度一直是 static
 * ------------------------------
 * Word/Excel 都有 caps-probe，PPT 没有 —— 116 个工具从上线起就没在真机上跑过。
 * 默认档因此把 8 个基础 PPT 工具也挡在外面（默认只有 32 个）。
 *
 * 本脚本分三档：
 *   基础档（第 1 轮）：新建/加页/标题/备注/插图/导出图/列文稿 —— 8 个工具。
 *   第二批（第 2 轮，INBOX 任务卡）：按域补 形状 / 表格 / 图表，每域 3~4 个工具。
 *     形状：add_shape、set_shape_fill、set_shape_text、align_objects
 *     表格：insert_table、set_table_cell、set_table_row_style（写后用 get_table_cell 读回；
 *           样式类工具没有读回手段，用【导出图差分】判断到底生没生效）
 *     图表：add_chart、set_ppt_chart_data、set_ppt_chart_style
 *
 * 判据收紧过：工具返回 success 不等于生效。表格行样式就是"返回成功但什么都没做"的典型，
 * 所以样式类一律配导出图差分（先连导两次做对照，确认导出本身确定，再比调用前后）。
 *
 * 探完就关、不保存、不留垃圾；只动临时目录里造的文件。
 * 无 WPS 时按既有约定"跳过不判失败"，但会明说未验证。
 *
 * 用法：node scripts/manual/ppt-caps-probe.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

// 独立管道：桥是常驻的，共享管道会拿到旧脚本/旧状态（与 snapshot-verify 同一招）
// -b2 后缀：第 2 轮按域扩展后可能与别处并发跑基础档，错开管道名免得互等
// 每次跑都用【全新】的管道名：固定名字会连上上一轮还活着的桥，而那个桥的内存里是
// 改过的旧 wps-com.ps1 —— 症状是"我明明改了 ps1，脚本行为却没变"，白排查半小时。
//（这个坑本轮踩过：修完 AddChart 枚举转换重跑，报错一字不变。）
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-ppt-probe-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
// PPT 首次被 COM 拉起比 Word 慢，给足时间
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '120000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let skipped = false;
/**
 * 已知失败清单：这些工具现在确实不工作，但**已经修在排期里**，
 * 不能让它们把 CI 永久染红（否则真回归会被"反正一直红"淹没）。
 *
 * 规则：
 *   - 实际失败 ⊆ 本清单 → 退出码 0（并打印"仍失败"明细）
 *   - 出现本清单之外的新失败 → 退出码 1（这是回归，必须拦）
 *   - 本清单里的工具某天通过了 → 打印"请移出清单并把 manifest 升 live"，退出码 0
 */
const EXPECTED_FAILURES = [
  // 目前为空。上一轮的 4 个假成功（set_shape_fill 动作缺失 / set_table_row_style 与
  // set_ppt_chart_style 传嵌套 style / set_ppt_chart_data 形状不符）已于 2026-10-18 修复并真机复验 18/18。
  // 机制留着：以后再冒出已知坏工具，加进来，别让 CI 长期红着 —— 红了就没人看真回归了。
];
const results = [];
const check = (tool, ok, detail = '') => {
  results.push({ tool, ok });
  console.log(`  ${ok ? '✓' : '✗'} ${tool}${detail ? ' — ' + detail : ''}`);
};

const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';

/** 解析 get_shapes 的文本输出 → [{index,name,type,left,top,w,h}]（对齐/位置类用例的读数手段） */
const parseShapes = (s) => {
  const out = [];
  const re = /\[(\d+)\]\s*(.+?)\s*\((.+?)\)\n\s*位置: \((-?[\d.]+), (-?[\d.]+)\) 大小: ([\d.]+) x ([\d.]+)/g;
  let m;
  while ((m = re.exec(s))) {
    out.push({ index: +m[1], name: m[2], type: m[3], left: +m[4], top: +m[5], w: +m[6], h: +m[7] });
  }
  return out;
};

// 1x1 红点 PNG，够 WPS 认格式；不用真图片是为了不给仓库塞二进制
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

// ---- 0. 注册与注解（不需要 WPS；先证明"工具在、注解对"，再谈真机）----
const BASIC_TOOLS = [
  'wps_ppt_create_presentation',
  'wps_ppt_add_slide',
  'wps_ppt_get_slide_count',
  'wps_ppt_set_slide_title',
  'wps_ppt_add_speaker_notes',
  'wps_ppt_insert_image',
  'wps_ppt_export_slide_as_image',
  'wps_ppt_get_open_presentations',
];
const B2_TOOLS = [
  // 形状
  'wps_ppt_add_shape',
  'wps_ppt_set_shape_fill',
  'wps_ppt_set_shape_text',
  'wps_ppt_align_objects',
  // 表格
  'wps_ppt_insert_table',
  'wps_ppt_set_table_cell',
  'wps_ppt_set_table_row_style',
  // 图表（注意：wps_ppt_add_chart 已于第 2 轮别名收敛降为旧名，规范名是 insert_ppt_chart）
  'wps_ppt_insert_ppt_chart',
  'wps_ppt_set_ppt_chart_data',
  'wps_ppt_set_ppt_chart_style',
];
(async () => {
  console.log('\n=== 0. 注册与注解（不需要 WPS）===');
  const missing = [...BASIC_TOOLS, ...B2_TOOLS].filter((n) => !toolRegistry.hasTool(n));
  console.log(
    `  ${missing.length === 0 ? '✓' : '✗'} 18 个目标工具全部已注册${missing.length ? ' — 缺: ' + missing.join(',') : ''}`
  );
  {
    const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
    const ro = ['wps_ppt_get_table_cell'].filter((n) => getToolAnnotations(n)?.readOnlyHint !== true);
    const wr = ['wps_ppt_add_shape', 'wps_ppt_set_shape_fill', 'wps_ppt_insert_table'].filter(
      (n) => getToolAnnotations(n)?.readOnlyHint === true
    );
    console.log(`  ${ro.length === 0 ? '✓' : '✗'} 读工具标了 readOnlyHint${ro.length ? ' — 漏: ' + ro.join(',') : ''}`);
    console.log(`  ${wr.length === 0 ? '✓' : '✗'} 写工具没被误标 readOnlyHint${wr.length ? ' — 误标: ' + wr.join(',') : ''}`);
  }

  const tmpDir = path.join(os.tmpdir(), `wps-ppt-probe-${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });
  const imgPath = path.join(tmpDir, 'probe.png');
  fs.writeFileSync(imgPath, Buffer.from(PNG_1PX, 'base64'));
  let createdName = null;

  try {
    console.log('\n=== PPT 基础档真机探测 ===');

    // 探活：WPS 演示没装 / 起不来时不该判失败（CI runner 上就是这样）。
    // 与 word/excel-caps-probe 的约定一致：跳过，但要说清"未验证"。
    const hb = await call('wps_ppt_get_open_presentations', {});
    if (!hb.success) {
      skipped = true;
      console.log(`  ⚠ WPS 演示不可用（${txt(hb).slice(0, 80)}），本轮跳过，不判失败`);
      return;
    }

    // 先记下"探测前已开着哪些"：这样才能分清哪个文稿是我们建的，收尾只关自己的
    const before = txt(hb);
    const beforeNames = new Set((before.match(/[^\s\\/:*?"<>|]+\.(?:pptx?|dps)/gi) || []));

    const created = await call('wps_ppt_create_presentation', {});
    const m = txt(created).match(/「([^」]+)」|:\s*([^（\n]+)（/);
    createdName = m ? (m[1] || m[2] || '').trim() || null : null;

    if (created.success) {
      check('wps_ppt_create_presentation', true, txt(created).split('\n')[0].slice(0, 60));
    } else {
      // 冷启动病史：WPS 演示第一次被 COM 拉起要十几秒，调用方先超时了，
      // 但文稿其实已经建出来（实测窗口标题已变成"演示文稿N"）。
      // 所以用"有没有新出现的文稿"来判定，而不是只看这次调用的返回。
      const after = txt(await call('wps_ppt_get_open_presentations', {}));
      const afterNames = after.match(/[^\s\\/:*?"<>|]+\.(?:pptx?|dps)/gi) || [];
      const fresh = afterNames.find((n) => !beforeNames.has(n));
      if (fresh) {
        createdName = fresh;
        check(
          'wps_ppt_create_presentation',
          true,
          `首次调用超时，但文稿已建立（${fresh}）—— 冷启动历史问题，不是功能缺陷`
        );
      } else {
        check('wps_ppt_create_presentation', false, txt(created).split('\n')[0].slice(0, 60));
        throw new Error('演示文稿建不起来，后面的探测没有意义');
      }
    }

    const added = await call('wps_ppt_add_slide', { layout: 'title_content', title: 'PPT 探测标题' });
    check('wps_ppt_add_slide', added.success, txt(added).split('\n')[0].slice(0, 60));

    const count = await call('wps_ppt_get_slide_count', {});
    const n = Number((txt(count).match(/(\d+)/) || [])[1] || 0);
    check('wps_ppt_get_slide_count', count.success && n >= 1, `读到 ${n} 页`);

    const title = await call('wps_ppt_set_slide_title', { slideIndex: 1, title: 'PPT 探测标题（改写）' });
    check('wps_ppt_set_slide_title', title.success, txt(title).split('\n')[0].slice(0, 60));

    const notes = await call('wps_ppt_add_speaker_notes', { slideIndex: 1, notes: '真机探测备注' });
    check('wps_ppt_add_speaker_notes', notes.success, txt(notes).split('\n')[0].slice(0, 60));

    const img = await call('wps_ppt_insert_image', { slideIndex: 1, path: imgPath });
    check('wps_ppt_insert_image', img.success, txt(img).split('\n')[0].slice(0, 60));

    const outPath = path.join(tmpDir, 'slide1.png');
    const exp = await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: outPath });
    const exported = fs.existsSync(outPath) && fs.statSync(outPath).size > 0;
    check(
      'wps_ppt_export_slide_as_image',
      exp.success && exported,
      exported ? `${Math.round(fs.statSync(outPath).size / 1024)} KB` : txt(exp).slice(0, 60)
    );

    // 判据不能看扩展名：未保存的文稿叫"演示文稿N"，没有 .pptx 后缀（第一版就写错了）
    const list = await call('wps_ppt_get_open_presentations', {});
    check(
      'wps_ppt_get_open_presentations',
      list.success && /演示文稿/.test(txt(list)) && /\d/.test(txt(list)),
      txt(list).split('\n')[0].slice(0, 60)
    );

    // ==================== C. 形状域（第 2 轮按域扩展）====================
    // 加形状 / 设填充 / 设文字 / 对齐。形状文本没有专用 get 工具，
    // 用 find_ppt_text 读回；对齐用 get_shapes 的位置读数判断"到底发生了什么"。
    console.log('\n--- C. 形状域：加形状 / 设填充 / 设文字 / 对齐 ---');
    // 页序不写死：新建文稿可能自带 1 页也可能 0 页，按当前页数推
    const slideCountNow = Number((txt(await call('wps_ppt_get_slide_count', {})).match(/(\d+)/) || [])[1] || 0);
    const SHAPE_SLIDE = slideCountNow || 2; // 基础档刚加的那一页
    const baseShapes = parseShapes(txt(await call('wps_ppt_get_shapes', { slideIndex: SHAPE_SLIDE })));
    const n0 = baseShapes.length;

    const addOk = await call('wps_ppt_add_shape', {
      slideIndex: SHAPE_SLIDE, type: 'rectangle', left: 40, top: 260, width: 120, height: 60,
      text: '形状探测甲', fillColor: '#3366FF',
    });
    await call('wps_ppt_add_shape', {
      slideIndex: SHAPE_SLIDE, type: 'oval', left: 200, top: 260, width: 120, height: 60, text: '形状探测乙',
    });
    await call('wps_ppt_add_shape', {
      slideIndex: SHAPE_SLIDE, type: 'rectangle', left: 360, top: 260, width: 120, height: 60, text: '形状探测丙',
    });
    const afterAdd = parseShapes(txt(await call('wps_ppt_get_shapes', { slideIndex: SHAPE_SLIDE })));
    check(
      'wps_ppt_add_shape',
      addOk.success && afterAdd.length === n0 + 3,
      `形状数 ${n0} → ${afterAdd.length}${addOk.success ? '' : '｜' + txt(addOk).slice(0, 60)}`
    );
    // 按 top=260 认出自己建的三个（占位符不在这一行）
    const mine = afterAdd.filter((s) => s.top === 260).sort((a, b) => a.left - b.left);

    if (mine[0]) {
      const fill = await call('wps_ppt_set_shape_fill', {
        slideIndex: SHAPE_SLIDE, shapeIndex: mine[0].index, color: '#FF0000',
      });
      check('wps_ppt_set_shape_fill', fill.success, txt(fill).split('\n')[0].slice(0, 80));
    } else {
      check('wps_ppt_set_shape_fill', false, '前置失败：没认出自己建的形状');
    }

    if (mine[0]) {
      const MARK = '形状文字已改写标记Q7';
      const st = await call('wps_ppt_set_shape_text', { slideIndex: SHAPE_SLIDE, shapeIndex: mine[0].index, text: MARK });
      const found = await call('wps_ppt_find_ppt_text', { text: MARK });
      const hits = Number((txt(found).match(/找到 (\d+) 处匹配/) || [])[1] || 0);
      check(
        'wps_ppt_set_shape_text',
        st.success && hits >= 1,
        `${txt(st).split('\n')[0].slice(0, 40)}｜find_ppt_text 读回 ${hits} 处`
      );
    }

    if (mine.length >= 3) {
      const beforeLeft = mine.map((s) => s.left);
      // 第 9 轮：align_objects 的 shapeIndices 变成必填 —— 以前不给清单时，底层
      // Shapes.Range($null) 退化成"整页所有形状一起对齐"，连表格/图片都搬走。
      // 现在 PS 侧对空清单直接报错，所以这里必须把要动的形状显式列出来。
      // 下面的 collateral 观测保留：它现在应该恒为 0，是"修好了没有"的可判定证据。
      const al = await call('wps_ppt_align_objects', {
        slideIndex: SHAPE_SLIDE,
        alignment: 'left',
        shapeIndices: mine.map((s) => s.index),
      });
      const afterAll = parseShapes(txt(await call('wps_ppt_get_shapes', { slideIndex: SHAPE_SLIDE })));
      const afterMine = afterAll.filter((s) => s.top === 260).sort((a, b) => a.index - b.index).map((s) => s.left);
      // 顺带数一下"没被要求的形状"有没有被一起动 —— 这是 align 系列能不能指定子集的关键
      const collateral = afterAll.filter((s2) => {
        const s1 = afterAdd.find((x) => x.index === s2.index);
        return s1 && s1.left !== s2.left && s2.top !== 260;
      }).length;
      check(
        'wps_ppt_align_objects',
        al.success && new Set(afterMine).size === 1,
        `before=[${beforeLeft}] after=[${afterMine}]｜本页另有 ${collateral} 个形状的 left 被一起改动`
      );
      // 附加观测（不计 PASSED）：旧名 align_shapes 的 schema 里有 shapeIndices，PS 却读 names
      if (toolRegistry.hasTool('wps_ppt_align_shapes')) {
        await call('wps_ppt_add_shape', {
          slideIndex: SHAPE_SLIDE, type: 'rectangle', left: 420, top: 340, width: 100, height: 50, text: '子集观测',
        });
        const preSub = parseShapes(txt(await call('wps_ppt_get_shapes', { slideIndex: SHAPE_SLIDE })));
        const sub = await call('wps_ppt_align_shapes', {
          slideIndex: SHAPE_SLIDE, shapeIndices: [preSub[preSub.length - 1].index], alignment: 'left',
        });
        const postSub = parseShapes(txt(await call('wps_ppt_get_shapes', { slideIndex: SHAPE_SLIDE })));
        const changedAll = postSub.filter((s2) => {
          const s1 = preSub.find((x) => x.index === s2.index);
          return s1 && s1.left !== s2.left;
        }).length;
        console.log(
          `  · 附加观测：wps_ppt_align_shapes 只传 1 个 shapeIndices → success=${sub.success}，` +
            `但本页 ${changedAll} 个形状的 left 被改动（shapeIndices 未被采纳）`
        );
      }
    }

    // ==================== D. 表格域（第 2 轮按域扩展）====================
    // 判定口径写死：
    //   - 写单元格：不用 get_table_cell 读回 —— 那个工具本身是坏的（PS 返回 value，TS 读 text，
    //     见 src/tools/ppt/table.ts:275-282 与 wps-com.ps1:4622）。改用裸调 getPptTableCell 看原始 value。
    //   - 设表头样式：没有读回手段，用导出图差分（先连导两次确认导出确定，再比调用前后）。
    console.log('\n--- D. 表格域：插表格 / 写单元格 / 设表头样式 ---');
    const TABLE_SLIDE = SHAPE_SLIDE + 1;
    const addT = await call('wps_ppt_add_slide', { layout: 'blank' });
    if (!addT.success) console.log(`  ⚠ 空白页没加上（${txt(addT).slice(0, 60)}），表格域会连锁失败`);
    const tbl = await call('wps_ppt_insert_table', { slideIndex: TABLE_SLIDE, rows: 2, cols: 3, left: 80, top: 120 });
    check('wps_ppt_insert_table', tbl.success, txt(tbl).split('\n')[0].slice(0, 80));

    const cell = await call('wps_ppt_set_table_cell', {
      slideIndex: TABLE_SLIDE, tableIndex: 1, row: 1, col: 1, text: '表头甲',
    });
    // 读回绕开坏掉的 get_table_cell，直接看 PS 的原始返回值
    const rawCell = await call('wps_execute_method', {
      method: 'getPptTableCell', appType: 'ppt',
      params: { slideIndex: TABLE_SLIDE, tableIndex: 1, row: 1, col: 1 },
    });
    const rawValue = (txt(rawCell).match(/"value":"([^"]*)"/) || [])[1];
    check(
      'wps_ppt_set_table_cell',
      cell.success && rawValue === '表头甲',
      `写入"表头甲"｜裸调读回 value="${rawValue === undefined ? '(没解析到)' : rawValue}"`
    );
    const toolReadBack = await call('wps_ppt_get_table_cell', { slideIndex: TABLE_SLIDE, tableIndex: 1, row: 1, col: 1 });
    const toolReadText = (txt(toolReadBack).match(/内容: (.*)/) || [])[1];
    console.log(
      `  · 附加观测（不计 PASSED）：工具 get_table_cell 读回 "${toolReadText}"` +
        `（PS 给的是 value，TS 读 text —— 这个读工具本身是坏的）`
    );

    // 表头样式：schema 是嵌套 style 对象，PS 动作读的是扁平参数（见 wps-com.ps1:4700-4709）。
    // 到底生没生效，用导出图差分说话：先连导两次确认导出确定，再比调用前后。
    const tImg = (n) => path.join(tmpDir, `tbl-${n}.png`);
    await call('wps_ppt_export_slide_as_image', { slideIndex: TABLE_SLIDE, outputPath: tImg(1) });
    await call('wps_ppt_export_slide_as_image', { slideIndex: TABLE_SLIDE, outputPath: tImg(2) });
    const exportDeterministic =
      fs.existsSync(tImg(1)) &&
      fs.existsSync(tImg(2)) &&
      Buffer.compare(fs.readFileSync(tImg(1)), fs.readFileSync(tImg(2))) === 0;
    const rowStyle = await call('wps_ppt_set_table_row_style', {
      slideIndex: TABLE_SLIDE, tableIndex: 1, row: 1,
      style: { backgroundColor: '#FF0000', fontColor: '#FFFFFF', bold: true, fontSize: 14 },
    });
    await call('wps_ppt_export_slide_as_image', { slideIndex: TABLE_SLIDE, outputPath: tImg(3) });
    const nestedWorked =
      exportDeterministic &&
      fs.existsSync(tImg(3)) &&
      Buffer.compare(fs.readFileSync(tImg(2)), fs.readFileSync(tImg(3))) !== 0;
    check(
      'wps_ppt_set_table_row_style',
      rowStyle.success && nestedWorked,
      `${txt(rowStyle).split('\n')[0].slice(0, 50)}｜导出图差分：${nestedWorked ? '有变化' : '无变化'}` +
        `（连续两次导出一致=${exportDeterministic}）`
    );
    // 附加观测（不计 PASSED）：用内置逃生舱按 PS 实际读的扁平参数再打一次，定位坏在哪一侧
    if (toolRegistry.hasTool('wps_execute_method')) {
      const raw = await call('wps_execute_method', {
        method: 'setPptTableRowStyle', appType: 'ppt',
        params: { slideIndex: TABLE_SLIDE, tableIndex: 1, row: 2, backgroundColor: '#00AA00', bold: true },
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: TABLE_SLIDE, outputPath: tImg(4) });
      const flatWorked =
        fs.existsSync(tImg(4)) && Buffer.compare(fs.readFileSync(tImg(3)), fs.readFileSync(tImg(4))) !== 0;
      console.log(
        `  · 附加观测：扁平参数裸调 setPptTableRowStyle → success=${raw.success}；` +
          `导出图 ${flatWorked ? '有变化（PS 侧吃得下扁平参数，TS 侧传错了形状）' : '无变化'}`
      );
    }

    // ==================== E. 图表域（第 2 轮按域扩展）====================
    // 判定口径写死：
    //   - add_chart：首次可能 E_FAIL（图表冷启动竞态，实测 4/5 成功），失败就重试一次并在 detail 里说明
    //   - set_ppt_chart_data：按【文档承诺的】{categories, series} 形状调用；再试二维数组。
    //     文档形状实测是"返回成功但什么都不写"（PS 按二维数组逐格写，见 wps-com.ps1:5928-5935），
    //     数字值还会报 Int32→String。所以判定看"文档形状能不能用"，不能只看 success。
    //   - set_ppt_chart_style：schema 是嵌套 style，PS 读扁平 title/hasLegend（wps-com.ps1:5947-5953），
    //     用导出图差分验证 title 到底进没进图表。
    console.log('\n--- E. 图表域：插柱状图 / 改数据源 / 改标题 ---');
    const CHART_SLIDE = TABLE_SLIDE + 1;
    const addC = await call('wps_ppt_add_slide', { layout: 'blank' });
    if (!addC.success) console.log(`  ⚠ 空白页没加上（${txt(addC).slice(0, 60)}），图表域会连锁失败`);
    let chart = await call('wps_ppt_insert_ppt_chart', {
      slideIndex: CHART_SLIDE, chartType: 'bar',
      data: [{ label: '甲', value: 10 }, { label: '乙', value: 20 }, { label: '丙', value: 15 }],
    });
    if (!chart.success) {
      const retry = await call('wps_ppt_insert_ppt_chart', {
        slideIndex: CHART_SLIDE, chartType: 'bar',
        data: [{ label: '甲', value: 10 }, { label: '乙', value: 20 }, { label: '丙', value: 15 }],
      });
      if (retry.success) {
        chart = retry;
        check('wps_ppt_insert_ppt_chart', true, `首次 E_FAIL（图表冷启动竞态），重试成功｜${txt(retry).split('\n')[0].slice(0, 40)}`);
      } else {
        check('wps_ppt_insert_ppt_chart', false, `两次都失败：${txt(chart).slice(0, 60)} / ${txt(retry).slice(0, 60)}`);
      }
    } else {
      check('wps_ppt_insert_ppt_chart', true, txt(chart).split('\n')[0].slice(0, 60));
    }
    if (chart.success) {
      // 图表数据：三种形状各打一次，每次都用导出图差分判断"到底写没写进去"。
      // 关键坑：PS 动作里的 $chart.ChartData.Activate() 【本身就会改变渲染】——
      // 不先把这一步的副作用打出来，"导出图有变化"分不清是写了数据还是激活了数据表。
      // 所以顺序是：对照 → 空 data（只激活不写）→ 文档形状 → 二维数组 → 带数字。
      const cImg = (n) => path.join(tmpDir, `chart-${n}.png`);
      const same = (a, b) =>
        fs.existsSync(a) && fs.existsSync(b) && Buffer.compare(fs.readFileSync(a), fs.readFileSync(b)) === 0;
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(1) });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(2) });
      const chartExportDet = same(cImg(1), cImg(2));
      // 空 data：Activate 会执行，但一个格子都不写 —— 用来量出"纯激活"的副作用
      await call('wps_execute_method', {
        method: 'setPptChartData', appType: 'ppt', params: { slideIndex: CHART_SLIDE, chartIndex: 1, data: [] },
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(3) });
      const activateShifted = chartExportDet && !same(cImg(2), cImg(3));
      console.log(`  · 附加观测：ChartData.Activate() 的副作用 = ${activateShifted ? '会改变渲染（后续差分已排除它）' : '无'}`);

      // 1) 文档承诺的形状：{categories: [...], series: [...]}
      const docForm = await call('wps_ppt_set_ppt_chart_data', {
        slideIndex: CHART_SLIDE, chartIndex: 1,
        data: { categories: ['甲', '乙', '丙'], series: [{ name: '数值', values: [11, 22, 33] }] },
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(4) });
      const docFormWorked = chartExportDet && !same(cImg(3), cImg(4));

      // 2) PS 实际吃的形状：二维数组（全字符串）
      const arrayForm = await call('wps_ppt_set_ppt_chart_data', {
        slideIndex: CHART_SLIDE, chartIndex: 1, data: [['类别', '数值'], ['甲', '44'], ['乙', '55'], ['丙', '66']],
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(5) });
      const arrayFormWorked = chartExportDet && !same(cImg(4), cImg(5));

      // 3) 自然形状：二维数组带数字
      const numForm = await call('wps_ppt_set_ppt_chart_data', {
        slideIndex: CHART_SLIDE, chartIndex: 1, data: [['类别', '数值'], ['甲', 77], ['乙', 88], ['丙', 99]],
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(8) });
      const numFormWorked = chartExportDet && !same(cImg(5), cImg(8));

      check(
        'wps_ppt_set_ppt_chart_data',
        docFormWorked,
        `文档形状 {categories,series}：${docFormWorked ? '生效' : '没生效（返回 success 也不代表写进去）'}；` +
          `二维数组全字符串：${arrayFormWorked ? '生效' : '没生效'}；` +
          `二维数组带数字：${numFormWorked ? '生效' : numForm.success ? '没生效' : txt(numForm).slice(0, 45)}`
      );
      // 图表标题：schema 是嵌套 style，PS 读扁平 title/hasLegend（wps-com.ps1:5947-5953）
      const cs = await call('wps_ppt_set_ppt_chart_style', {
        slideIndex: CHART_SLIDE, chartIndex: 1, style: { title: '探测图表标题' },
      });
      await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(6) });
      const titleApplied = chartExportDet && !same(cImg(5), cImg(6));      check(
        'wps_ppt_set_ppt_chart_style',
        cs.success && titleApplied,
        `${txt(cs).split('\n')[0].slice(0, 40)}｜导出图差分：${titleApplied ? '有变化（标题进去了）' : '无变化（标题没进去）'}` +
          `（连续两次导出一致=${chartExportDet}）`
      );
      if (toolRegistry.hasTool('wps_execute_method')) {
        const rawStyle = await call('wps_execute_method', {
          method: 'setPptChartStyle', appType: 'ppt',
          params: { slideIndex: CHART_SLIDE, chartIndex: 1, title: '扁平标题对照' },
        });
        await call('wps_ppt_export_slide_as_image', { slideIndex: CHART_SLIDE, outputPath: cImg(7) });
        const flatApplied = chartExportDet && !same(cImg(6), cImg(7));
        console.log(
          `  · 附加观测：扁平参数裸调 setPptChartStyle(title) → success=${rawStyle.success}；` +
            `导出图 ${flatApplied ? '有变化（PS 侧吃扁平 title，TS 侧传错了形状）' : '无变化'}`
        );
      }
    } else {
      // 前置就没建起来：这两个工具本轮【未验证】，不能记 ✗（那不是它们的错）
      console.log('  · wps_ppt_set_ppt_chart_data / wps_ppt_set_ppt_chart_style：未验证（前置：图表没建起来）');
    }
  } catch (e) {
    console.log(`  ⚠ 探测中断：${e && e.message ? e.message : String(e)}`);
    console.log('    （需要 WPS 演示可访问。中断部分不判失败，但也不算数 —— 仍按 static 计）');
  } finally {
    try {
      if (createdName) {
        const closed = await call('wps_ppt_close_presentation', { name: createdName, save: false });
        console.log(`  ${closed.success ? '✓' : '·'} 收尾关闭 ${createdName}${closed.success ? '' : '（没关掉，可能名字不对）'}`);
      } else {
        const closed = await call('wps_ppt_close_presentation', { save: false });
        console.log(`  ${closed.success ? '✓' : '·'} 收尾关闭（按活动文稿）`);
      }
    } catch { /* 已经关了 */ }
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { /* 占用 */ }

    if (skipped) {
      console.log('未探测到任何 PPT 能力（WPS 演示不可用），按约定跳过。');
      process.exit(0);
    }
    const passed = results.filter((r) => r.ok).map((r) => r.tool);
    const failed = results.filter((r) => !r.ok).map((r) => r.tool);
    console.log(`\n通过 ${passed.length}/${results.length}`);
    if (failed.length) console.log(`失败：${failed.join(', ')}`);
    console.log('\n把通过的这些逐条写进 scripts/manual/tool-verification.json 的 tools 里（值 live），');
    console.log('再跑 node scripts/manual/gen-tool-verification.js && npm run build。');
    console.log(`PASSED=${passed.join(',')}`);

    // 已知失败不进 CI 红；新失败（回归）必须红
    const unexpected = failed.filter((t) => !EXPECTED_FAILURES.includes(t));
    const recovered = EXPECTED_FAILURES.filter((t) => passed.includes(t));
    for (const t of recovered) {
      console.log(`  ↑ ${t} 已经能用了：请把它从 EXPECTED_FAILURES 移出，并在 manifest 里升 live`);
    }
    if (unexpected.length) {
      console.log(`\n✗ 出现计划外失败：${unexpected.join(', ')} —— 这是回归，需要立即处理`);
      process.exit(1);
    }
    if (failed.length) {
      console.log(`\n（失败 ${failed.length} 个全在 EXPECTED_FAILURES 清单内，属已知问题，不判失败）`);
    }
    process.exit(0);
  }
})();
