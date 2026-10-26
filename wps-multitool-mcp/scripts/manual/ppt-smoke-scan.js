#!/usr/bin/env node
/**
 * Input: docs/agents/PPT_SMOKE_MATRIX.md（探路者产出的 95 行表：工具 | 安全参数 JSON | 判据 | 备注）
 * Output: PASSED= / FAILED= 两份清单（失败的带复现参数）
 * Pos: PPT 冒烟的驱动器（第 5 轮）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 这个脚本和 ppt-caps-probe.js 的分工
 * --------------------------------
 * caps-probe 是**精验**：每个工具写死的用例 + 严格判据（读回比对具体值、导出图差分）。
 * 本脚本是**扫雷**：95 个 static 工具一次过，只判两件事 —— 不报错、有变化。
 * 所以它故意宽松：冒烟通过 ≠ live，通过项在 manifest 里仍然记 static（本脚本不碰那个文件）。
 * 它要抓的不是"结果对不对"，而是"调用会不会直接炸 / 是不是压根没生效"。
 *
 * 为什么表驱动而不是把 95 个用例写进脚本
 * ----------------------------------
 * 参数和判据是**会随实现变**的知识：某个工具的必填项改了、某个形状参数换了名字，
 * 该改的是矩阵这张表，不是驱动器。写死在脚本里的后果是：改一次参数要动两个文件，
 * 而且驱动器会悄悄和矩阵不一致（谁也不会去对）。
 *
 * 判据只有三种（矩阵 §2 的"判据"列）
 * -------------------------------
 *   数量差   —— 调用前后比"页数 / 形状数 / 动画数 / 打开文稿数"，变了才算生效。
 *   导出图差分 —— 先把同一页连导两次确认导出本身稳定（不稳定的差分没有意义），
 *                再比调用前后的字节。
 *   读回     —— 调用成功 + 返回内容不是空壳（没有 undefined / 空串）。
 *   跳过     —— 见矩阵 §4（要外部文件 / 语义是等用户操作 / 无可靠观测手段）。
 *
 * 顺序：常规段 → 末段（改页数或改"当前文稿"）→ open/close/switch 单独最末。
 * 末段跑完重开一份干净文稿，否则后面的差分基线全被污染。
 *
 * 用法：node scripts/manual/ppt-smoke-scan.js
 *      node scripts/manual/ppt-smoke-scan.js --limit 10   # 只跑前 10 条（统筹者抽查用）
 */

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

// ---- 独立管道 + 时间戳（照 ppt-caps-probe.js）--------------------------------
// 固定名字会连上上一轮还活着的桥，而那个桥内存里是旧版 wps-com.ps1 ——
// 症状是"改了 ps1 行为却没变"，排查成本极高（第 4 轮踩过）。
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-ppt-smoke-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '120000';

const ROOT = path.resolve(__dirname, '..', '..');
const REPO = path.resolve(ROOT, '..');
const MATRIX = path.join(REPO, 'docs', 'agents', 'PPT_SMOKE_MATRIX.md');

const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const LIMIT = (() => {
  const i = process.argv.indexOf('--limit');
  return i > 0 ? Number(process.argv[i + 1]) : 0;
})();

// ---- 矩阵解析 ---------------------------------------------------------------
/**
 * 从 markdown 表里解析出 [{ tool, args, judge, note }]。
 * 只认 §2 主表的行（`| wps_xxx | \`{...}\` | 判据 | 备注 |`）。
 * 备注里的【末段】用于分段，所以必须原样带出来。
 *
 * 第 9 轮：矩阵多了第 5 列（setup / check），多步用例与"读回值"判据**本脚本跑不了**。
 * 顺序跑本来就分不清"工具坏"和"靶子已经不是当初那个靶子"（第 7 轮的教训），
 * 所以这里不猜、不硬跑 —— 直接跳过并说明该走 scripts/manual/ppt-isolated-verify.js。
 */
function parseMatrix(file) {
  const text = fs.readFileSync(file, 'utf8');
  const rows = [];
  for (const line of text.split(/\r?\n/)) {
    if (!line.startsWith('| wps_')) continue;
    const cells = line.split('|').map((s) => s.trim());
    // ['', tool, args, judge, note, extra, ''] —— 备注可能为空
    const tool = cells[1];
    const rawArgs = (cells[2] || '').replace(/^`|`$/g, '').trim() || '{}';
    const judge = cells[3] || '';
    const note = cells[4] || '';
    const extraRaw = (cells[5] || '').replace(/`/g, '').trim();
    if (extraRaw) {
      rows.push({ tool, args: null, judge: '交给隔离验证器', note: '第 5 列有多步前置/读回，顺序版跑不了：' + extraRaw });
      continue;
    }
    if (judge.includes('读回值')) {
      rows.push({ tool, args: null, judge: '交给隔离验证器', note: '「读回值」判据要第 5 列的 check，顺序版跑不了' });
      continue;
    }
    let args;
    try {
      args = JSON.parse(rawArgs);
    } catch {
      rows.push({ tool, args: null, judge: '跳过', note: '矩阵里的 JSON 解析失败：' + rawArgs });
      continue;
    }
    rows.push({ tool, args, judge, note });
  }
  return rows;
}

if (!fs.existsSync(MATRIX)) {
  console.error(`✗ 找不到矩阵文件：${MATRIX}\n  这是探路者的产出，没有它就无法表驱动。`);
  process.exit(1);
}
const ALL_ROWS = parseMatrix(MATRIX);
if (!ALL_ROWS.length) {
  console.error('✗ 矩阵里解析到 0 行 —— 表格格式变了？期望 `| 工具 | \`{json}\` | 判据 | 备注 |`');
  process.exit(1);
}

// 分段：常规 / 末段（备注带【末段】或属于结构类）/ 文稿切换（最末）
const STRUCTURAL = new Set([
  'wps_ppt_delete_slide', 'wps_ppt_move_slide', 'wps_ppt_copy_slide', 'wps_ppt_add_textbox',
  'wps_ppt_beautify_all_slides', 'wps_ppt_add_title_decoration', 'wps_ppt_add_page_indicator',
  'wps_ppt_duplicate_shape', 'wps_ppt_group_shapes',
]);
const PRES_SWITCH = new Set(['wps_ppt_open_presentation', 'wps_ppt_close_presentation', 'wps_ppt_switch_presentation']);

const isLast = (r) => r.note.includes('【末段】') || STRUCTURAL.has(r.tool) || /^wps_ppt_create_/.test(r.tool);
const normal = ALL_ROWS.filter((r) => !PRES_SWITCH.has(r.tool) && !isLast(r) && r.judge !== '跳过');
const last = ALL_ROWS.filter((r) => !PRES_SWITCH.has(r.tool) && isLast(r) && r.judge !== '跳过');
const pres = ALL_ROWS.filter((r) => PRES_SWITCH.has(r.tool) && r.judge !== '跳过');
const skipped = ALL_ROWS.filter((r) => r.judge === '跳过');

// --parse-only：只验"矩阵能不能被解析成分段计划"，全程不碰工具、不起桥。
// 存在的理由：矩阵格式一变（多一列、JSON 里带 `|`、判据写错），脚本会静默少跑一批。
// 这个开关让"表坏了"这件事能在 CI 里被发现，而不是等真机跑完才发现少了一半。
if (process.argv.includes('--parse-only')) {
  const judges = {};
  for (const r of ALL_ROWS) judges[r.judge] = (judges[r.judge] || 0) + 1;
  const badJudge = ALL_ROWS.filter((r) => !['读回', '导出图差分', '数量差', '跳过'].includes(r.judge));
  console.log(`矩阵：${MATRIX}`);
  console.log(`  总行数 ${ALL_ROWS.length}`);
  console.log(`  分段：常规 ${normal.length} / 末段 ${last.length} / 文稿切换 ${pres.length} / 跳过 ${skipped.length}`);
  console.log(`  判据分布：${Object.entries(judges).map(([k, v]) => `${k} ${v}`).join(' / ')}`);
  if (badJudge.length) {
    console.log(`  ✗ 非法判据：${badJudge.map((r) => `${r.tool}="${r.judge}"`).join(', ')}`);
    process.exit(1);
  }
  const missingArgs = ALL_ROWS.filter((r) => r.args === null);
  if (missingArgs.length) {
    console.log(`  ✗ 参数 JSON 解析失败：${missingArgs.map((r) => r.tool).join(', ')}`);
    process.exit(1);
  }
  console.log('  ✓ 判据合法、参数 JSON 全部可解析');
  process.exit(0);
}

// ---- 基础设施 ---------------------------------------------------------------
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const results = [];
const note = (tool, ok, detail, repro) => {
  results.push({ tool, ok, detail, repro });
  console.log(`  ${ok ? '✓' : '✗'} ${tool}${detail ? ' — ' + detail : ''}`);
};

/**
 * 直接调 handler，不走 toolRegistry.callTool。
 * 理由：callTool 外面裹了自动快照（第 4 轮），95 条里 30 多条是破坏性工具 →
 * 会给一份**临时靶子文稿**建几十份快照，纯属浪费（况且靶子跑完就删）。
 * 代价是不再顺带验路由与参数校验 —— 那两件事各有专门的门禁（param-shape / alias）。
 */
const call = (name, args = {}) => {
  const t = toolRegistry.getTool(name);
  if (!t) throw new Error(`工具未注册：${name}`);
  return t.handler(args);
};
const txt = (r) => (r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '';

/** 从 get_shapes 文本里数形状个数（比解析坐标稳，格式一变只会读到 0，不会误判成通过） */
const countEntries = (s) => (String(s).match(/^\s*\[\d+\]/gm) || []).length;

let tmpDir = null;
const imgPath = () => path.join(tmpDir, 'smoke-1px.png');
const shotPath = (tag) => path.join(tmpDir, `shot-${tag}.png`);

/** 状态指纹：页数 + 每页形状数 + 第 1 页动画数 */
async function fingerprint() {
  const countR = await call('wps_ppt_get_slide_count', {});
  const slideCount = Number((txt(countR).match(/(\d+)/) || [])[1] || 0);
  const shapeCounts = [];
  for (let i = 1; i <= Math.min(slideCount, 3); i++) {
    const s = await call('wps_ppt_get_shapes', { slideIndex: i });
    shapeCounts.push(countEntries(txt(s)));
  }
  let animCount = 0;
  try {
    const a = await call('wps_ppt_get_animations', { slideIndex: 1 });
    animCount = countEntries(txt(a));
  } catch {
    /* 动画读回失败不致命，指纹少一维 */
  }
  const openR = await call('wps_ppt_get_open_presentations', {});
  const openCount = countEntries(txt(openR));
  return { slideCount, shapeCounts, animCount, openCount };
}

const samePrint = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/** 导出第 1 页并取字节哈希 */
async function shot(tag) {
  const out = shotPath(tag);
  try { fs.unlinkSync(out); } catch {}
  const r = await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: out });
  if (!r.success || !fs.existsSync(out)) return null;
  return crypto.createHash('md5').update(fs.readFileSync(out)).digest('hex');
}

// ---- 主流程 -----------------------------------------------------------------
(async () => {
  tmpDir = path.join(os.tmpdir(), `wps-ppt-smoke-${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });
  fs.writeFileSync(imgPath(), Buffer.from(PNG_1PX, 'base64'));

  console.log(`\n=== 0. 矩阵与注册表对账（不需要 WPS）===`);
  const missing = ALL_ROWS.filter((r) => !toolRegistry.hasTool(r.tool)).map((r) => r.tool);
  console.log(`  矩阵 ${ALL_ROWS.length} 行：常规 ${normal.length} / 末段 ${last.length} / 文稿切换 ${pres.length} / 跳过 ${skipped.length}`);
  console.log(`  ${missing.length ? '✗' : '✓'} 表里的工具全部已注册${missing.length ? ' — 缺: ' + missing.join(',') : ''}`);

  console.log(`\n=== 1. 探活 ===`);
  const hb = await call('wps_ppt_get_open_presentations', {});
  if (!hb.success) {
    console.log(`  ⚠ WPS 演示不可用（${txt(hb).slice(0, 60)}）—— 按约定【跳过不判失败】，但这批工具本次没有被验证。`);
    console.log('\nPASSED=0');
    console.log('FAILED=0');
    process.exit(0);
  }

  console.log(`\n=== 2. 建靶子文稿（2 页；每页 2 形状 + 1 文本框 + 1 表格 + 1 图片 + 1 超链接形状）===`);
  let created = null;
  // 建靶要重试：WPS 对新加的那一页有冷启动期，第一次往它上面加任何元素都会 E_FAIL
  // （实测：第 1 页 6 个元素全成功，第 2 页 6 个全 E_FAIL，第二次就好）。
  // 不重试的话靶子只有 1 页，后面所有工具都在残靶上失败 —— 79 个"失败"里有 58 个是这么来的。
  const mk = async (fn, label) => {
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const r = await fn();
        if (r.success) {
          console.log(`  ✓ ${label}${attempt > 1 ? `（第 ${attempt} 次才成）` : ''}`);
          return true;
        }
        if (attempt === 3) {
          console.log(`  ✗ ${label} — ${txt(r).slice(0, 80)}`);
          return false;
        }
        await new Promise((res) => setTimeout(res, 300 * attempt));
      } catch (e) {
        if (attempt === 3) {
          console.log(`  ✗ ${label} — ${e.message}`);
          return false;
        }
        await new Promise((res) => setTimeout(res, 300 * attempt));
      }
    }
    return false;
  };
  const cr = await mk(() => call('wps_ppt_create_presentation', {}), '新建空白文稿');
  if (!cr) {
    console.log('\n⚠ 靶子建不起来，冒烟无法进行。');
    console.log('PASSED=0');
    console.log('FAILED=0');
    process.exit(0);
  }
  {
    const r = await call('wps_ppt_get_open_presentations', {});
    created = (txt(r).match(/([^\s\\/]+\.pptx)/) || [])[1] || null;
  }
  // 新建的 WPS 文稿是 0 页（不是 1 页）—— 之前只加了一次页，于是"第 2 页"根本不存在，
  // 往它上面加元素全 E_FAIL，后面 58 条工具在残靶上失败，看起来像 58 个工具坏了。
  await mk(() => call('wps_ppt_add_slide', { layout: 'blank' }), '第 1 页');
  await mk(() => call('wps_ppt_add_slide', { layout: 'blank' }), '第 2 页');
  const built = Number((txt(await call('wps_ppt_get_slide_count', {})).match(/(\d+)/) || [])[1] || 0);
  for (const page of [1, 2]) {
    // 参数名必须与 schema 一致：add_shape 要的是 type（不是 shapeType）。
    // 写错过一次：PS 读不到 type → AddShape($null) → E_FAIL，靶子只建出第 1 页，
    // 后面 58 条工具全在空靶子上失败，看起来像"58 个工具坏了"，实际是建靶失败。
    await mk(() => call('wps_ppt_add_shape', { slideIndex: page, type: 'rectangle', left: 60, top: 60, width: 120, height: 60 }), `p${page} 形状1`);
    await mk(() => call('wps_ppt_add_shape', { slideIndex: page, type: 'oval', left: 240, top: 60, width: 120, height: 60 }), `p${page} 形状2`);
    await mk(() => call('wps_ppt_add_textbox', { slideIndex: page }), `p${page} 文本框`);
    await mk(() => call('wps_ppt_insert_table', { slideIndex: page, rows: 2, cols: 2 }), `p${page} 表格`);
    await mk(() => call('wps_ppt_insert_image', { slideIndex: page, path: imgPath() }), `p${page} 图片`);
    // 同理：超链接工具的参数是 url，不是 address
    await mk(() => call('wps_ppt_add_ppt_hyperlink', { slideIndex: page, shapeIndex: 1, url: 'https://example.com' }), `p${page} 超链接形状`);
  }
  const builtCount = built * 6;
  console.log(`  靶子页数 = ${built}${built >= 2 ? ' ✓' : ' ⚠ 不足 2 页，末段用例会失真'}`);
  console.log(`  靶子元素 = ${builtCount}/12${builtCount >= 12 ? ' ✓' : ' ⚠ 元素不足'}`);

  // 给靶子写文字：unify_font / beautify / beautify_all_slides / set_font_color 这些
  // 全是"改文字样式"的工具，靶子上一旦没有文字，它们改不了任何东西 ——
  // 导出图当然没变化，于是被判成"假成功"。上一轮 4 条里 3 条就是这么误判的。
  // 这不是改判据，是把靶子补到"这些工具本来该作用的对象存在"。
  console.log('  写入文字（让文字类工具有作用对象）：');
  let textOk = 0;
  for (const page of [1, 2]) {
    // 形状索引 1 = 矩形，3 = 文本框（顺序：矩形/椭圆/文本框/表格/图片/超链接）
    if (await mk(() => call('wps_ppt_set_shape_text', { slideIndex: page, shapeIndex: 1, text: `冒烟靶子 第${page}页 标题文字` }), `p${page} 形状1 文字`)) textOk++;
    if (await mk(() => call('wps_ppt_set_shape_text', { slideIndex: page, shapeIndex: 3, text: `第${page}页正文样例文字，用于验证字体与配色类工具` }), `p${page} 文本框 文字`)) textOk++;
  }
  console.log(`  带文字的元素 = ${textOk}/4${textOk >= 4 ? ' ✓' : ' ⚠ 文字类工具的结论会不准'}`);

  // 导出确定性：连导两次，字节必须一致，否则"差分"这个判据本身就不可信
  const d1 = await shot('det1');
  const d2 = await shot('det2');
  const exportStable = d1 !== null && d1 === d2;
  console.log(`  ${exportStable ? '✓' : '⚠'} 导出确定性（连导两次一致）${exportStable ? '' : ' — 不稳定，导出图差分类用例会标"无法判定"'}`);

  // --build-only：只建靶、验靶，然后退出。
  // 存在的理由：交付前必须证明"靶子真的能建起来"（上一轮 58 个失败全是因为靶子只有 1 页），
  // 而跑完整冒烟要 10 分钟且属于车道 C 的正式占用。这个开关让"靶子是否健康"能单独、快速地验。
  if (process.argv.includes('--build-only')) {
    console.log('\n=== --build-only：靶子建好即退出（不跑冒烟用例）===');
    console.log(`  页数=${built} 元素=${builtCount}/12 带文字=${textOk}/4 导出稳定=${exportStable}`);
    const healthy = built >= 2 && builtCount >= 12 && textOk >= 4;
    console.log(`  ${healthy ? '✓ 靶子健康' : '✗ 靶子不健康 —— 这状态下跑出来的失败数没有意义'}`);
    // 不保存直接关，别留垃圾
    try { await call('wps_ppt_close_presentation', { name: created || undefined, save: false }); } catch {}
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
    process.exit(healthy ? 0 : 1);
  }

  console.log(`\n=== 3. 常规段（${normal.length} 条）===`);
  let basePrint = await fingerprint();
  let baseShot = exportStable ? d1 : null;

  const runRow = async (row) => {
    const { tool, args, judge } = row;
    const repro = `${tool} ${JSON.stringify(args)}`;
    if (!toolRegistry.hasTool(tool)) {
      note(tool, false, '工具未注册', repro);
      return;
    }
    let r;
    try {
      r = await call(tool, args);
    } catch (e) {
      note(tool, false, `抛异常: ${e.message}`, repro);
      return;
    }
    if (!r || r.success !== true) {
      note(tool, false, `报错: ${String(r && (r.error || txt(r))).slice(0, 100)}`, repro);
      return;
    }

    if (judge === '读回') {
      const body = txt(r).trim();
      const shell = !body || /undefined|null|NaN/.test(body);
      note(tool, !shell, shell ? `读回是空壳: ${body.slice(0, 60)}` : `读回 ${body.length} 字`, repro);
      return;
    }
    if (judge === '数量差') {
      // 页数 / 形状数 / 动画数 / 打开文稿数，任一变化即"生效"
      const after = PRES_SWITCH.has(tool) ? null : await fingerprint();
      const changed = after ? !samePrint(basePrint, after) : false;
      if (changed) {
        basePrint = after;
        note(tool, true, '计数有变化', repro);
      } else {
        note(tool, false, '计数无变化（返回 success 但没生效）', repro);
      }
      return;
    }
    if (judge === '导出图差分') {
      if (!exportStable || !baseShot) {
        note(tool, false, '导出图不稳定，无法判定（不算通过）', repro);
        return;
      }
      const after = await shot('diff');
      const changed = after !== null && after !== baseShot;
      if (changed) baseShot = after;
      // 差分没变化时补一条诊断（不改结论，只把线索摆出来）：
      // 有些操作是"结构变了但静态图看不见"（空占位符、切换效果、页脚覆盖层），
      // 也有的是"矩阵参数根本不是 schema 的参数"（工具其实没收到值）。
      // 这两种的最优判据完全不同，靠一句"无变化"没法分诊 —— 第 6 轮为此白判过一次。
      let hint = '';
      if (!changed) {
        const afterPrint = await fingerprint();
        if (!samePrint(basePrint, afterPrint)) {
          hint = `（注：导出图没变，但计数变了 ${JSON.stringify(basePrint)} → ${JSON.stringify(afterPrint)}：` +
            '可能是不可见的结构变化，判据该换读回/数量差，不是工具坏）';
          basePrint = afterPrint;
        } else {
          hint = '（注：图与计数都没变，检查矩阵参数是否与 schema 一致 —— 参数名不对时工具会静默什么都不做）';
        }
      }
      note(tool, changed, (changed ? '导出图有变化' : '导出图无变化（返回 success 但没生效）') + hint, repro);
      return;
    }
    note(tool, false, `未知判据 "${judge}"（矩阵只允许 读回/导出图差分/数量差/跳过）`, repro);
  };

  for (const row of (LIMIT ? normal.slice(0, LIMIT) : normal)) await runRow(row);

  console.log(`\n=== 4. 末段（改页数/结构，${last.length} 条）===`);
  for (const row of last) await runRow(row);

  console.log(`\n=== 5. 文稿切换（最末，${pres.length} 条）===`);
  for (const row of pres) await runRow(row);

  console.log(`\n=== 6. 收尾 ===`);
  // 关掉靶子（不保存）+ 删临时目录。存不存在都兜住，别让收尾把结论带崩。
  try {
    const r = await call('wps_ppt_close_presentation', { name: created || undefined, save: false });
    console.log(`  ${r.success ? '✓' : '⚠'} 关闭临时文稿${r.success ? '' : ' — ' + txt(r).slice(0, 80)}`);
  } catch (e) {
    console.log(`  ⚠ 关闭临时文稿异常：${e.message}`);
  }
  try {
    const left = await call('wps_ppt_get_open_presentations', {});
    const n = countEntries(txt(left));
    console.log(`  ${n <= 0 ? '✓' : '⚠'} 剩余打开的文稿 ${n} 个${n > 0 ? '（可能不是自己建的，不擅自关）' : ''}`);
  } catch {}
  try {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  } catch {}

  // ---- 结果清单（INBOX 要求的两行机器可读输出）----
  const passed = results.filter((r) => r.ok).map((r) => r.tool);
  const failed = results.filter((r) => !r.ok);
  console.log(`\n=== 结果 ===`);
  console.log(`  跑完 ${results.length} 条，通过 ${passed.length}，失败 ${failed.length}，跳过 ${skipped.length}`);
  if (failed.length) {
    console.log('  失败明细（复现参数直接可喂）：');
    for (const f of failed) console.log(`    ✗ ${f.repro}  → ${f.detail}`);
  }
  console.log(`\nPASSED=${passed.length}`);
  console.log(`PASSED= ${passed.join(', ')}`);
  console.log(`FAILED=${failed.length}`);
  console.log(`FAILED= ${failed.map((f) => f.repro).join(' | ') || '(无)'}`);
  console.log('\n注：冒烟通过 ≠ live —— 这些工具在 manifest 里仍记 static，本脚本不改 tool-verification.json。');

  process.exit(failed.length ? 1 : 0);
})().catch((e) => {
  console.error(`\n✗ 冒烟驱动器自己崩了：${e && e.stack ? e.stack : e}`);
  console.log('PASSED=0');
  console.log(`FAILED=1`);
  console.log(`FAILED= (驱动器异常) ${e && e.message ? e.message : e}`);
  process.exit(1);
});
