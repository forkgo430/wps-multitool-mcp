/**
 * Input: docs/agents/PPT_SMOKE_MATRIX.md（工具 + 安全参数 + 判据 + 可选前置/读回）
 * Output: 每个工具在【全新靶子】上单独跑一次的结果（PASSED=/FAILED= 清单）
 * Pos: 冒烟脚本的"隔离版"。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ----------
 * ppt-smoke-scan.js 是顺序跑的：前一条留下的状态会影响后一条（版式被换、形状被删、
 * 尺寸被改），于是"图没变化"既可能是工具坏，也可能是靶子已经不是当初那个靶子了。
 * 实测过：wps_ppt_set_background 在序列里判失败，单独跑三个（含两个别名）全部成功。
 * 所以这里每条用例都重建一次靶子 —— 慢，但结论可信。
 *
 * 走 callTool 而不是直接拿 handler（第 8 轮改的）
 * --------------------------------------------------
 * callTool 才是 MCP 客户端真正走的那条路：别名解析、别名参数映射、必填校验、自动快照都在里面。
 * 直接调 handler 等于绕过了整条链 —— 矩阵里那些"用旧名 + 旧参数"的行会被判成成功，
 * 而线上其实早就 Missing required parameter 了。走真入口，旧名用例才有意义。
 *
 * 第 9 轮加的：一条用例可以有多步（第 9 轮还债的主体）
 * ---------------------------------------------------------
 * 光"重建靶子再调一次"表达不了一些真实前置，21 条跳过里有一半是被这个卡住的，
 * 结论写成了"靶子没有 / 需要外部素材"，其实只是【用例设计】没跟上：
 *   · remove_animation / set_animation_order —— 先在同一条用例里 set_animation(op='add') 加两个动画
 *   · set_ppt_chart_data / set_ppt_chart_style —— 先 insert_ppt_chart 插一个图表
 *   · set_slide_subtitle / set_slide_content —— 先 add_slide(layout:'title' / 'title_content')
 *   · add_master_element —— 判据是"母版元素数"，所以指纹里必须带母版
 *   · close_presentation / switch_presentation / set_active_target —— 需要第二份文稿
 * 于是矩阵第 5 列给一段 JSON：{"setup":[...],"check":{...}}。
 *
 * 判据多了一种：读回值
 * --------------------
 * 原来的"读回"只断言"回包不是空话"，太松（工具回一句"成功"就算过）。
 * "读回值"要求调用完再去调一个【别的工具】把结果读回来，断言目标字符串在里面。
 * 表格单元格、形状文字、母版元素、动画顺序、文稿锁定都靠它 —— 这才是真正的读回。
 *
 * 素材从哪来（不要"需外部素材"这个借口）
 * ------------------------------------
 * PPT 是这轮唯一能自造素材的一边，而且不用手搓 OOXML：
 *   · 1×1 PNG：内联 base64（见 png 常量）
 *   · .pptx：启动时用 WPS 自己存一份 —— create_presentation → add_slide → save_as → close
 *     存出来的一定是合法 pptx，比手写 OOXML 靠谱，也不会在别人机器上缺字体/缺主题。
 * 矩阵里用 {{PNG}} / {{SOURCE_PPTX}} / {{TARGET}} 三个占位符引用，运行时替换。
 * {{TARGET}} 的存在是因为 WPS 的新建文稿叫"演示文稿83"这种带序号的随机名字，写死必错。
 *
 * 只关自己建的文稿
 * ----------------
 * owned 集合记下【本进程新建/新打开】的文稿名，收尾只关它们。
 * 以前是"关掉所有叫 演示文稿\d+ 的" —— 万一用户正开着别的演示文稿，那是在动用户的文件。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-isolated-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';

const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const REPO = path.resolve(ROOT, '..');
const MATRIX = path.join(REPO, 'docs', 'agents', 'PPT_SMOKE_MATRIX.md');
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-iso-'));

// 1×1 PNG：内联 base64，只用来证明"插进来了"（占位素材）。
const png = path.join(tmpDir, 'one.png');
fs.writeFileSync(
  png,
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )
);

/**
 * 再造一张【看得见】的图：16×16 纯红。
 * 为什么不用上面那张 1×1：替换/设背景都要求"渲染结果真的变了"才算数，
 * 而 1×1 那张插进去以后原尺寸就是 0.75×0.75，再拿它替换等于什么都没变 ——
 * 判据必须建在"换了东西"上，不能建在"调用没报错"上。
 */
function makeSolidPng(file, size, rgb) {
  const zlib = require('zlib');
  const crcTable = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;   // bit depth
  ihdr[9] = 2;   // truecolor
  const raw = Buffer.alloc(size * (size * 3 + 1));
  let o = 0;
  for (let y = 0; y < size; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      raw[o++] = rgb[0]; raw[o++] = rgb[1]; raw[o++] = rgb[2];
    }
  }
  fs.writeFileSync(
    file,
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', ihdr),
      chunk('IDAT', zlib.deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ])
  );
}
const png2 = path.join(tmpDir, 'red16.png');
makeSolidPng(png2, 16, [255, 0, 0]);

const sourcePptx = path.join(tmpDir, 'source.pptx');

const call = (name, args = {}) => toolRegistry.callTool({ id: name, name, arguments: args });
const txt = (r) => String(((r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '')).replace(/\s+/g, ' ');

/** 本进程新建/新打开的文稿名。收尾只关这些，不碰用户自己的文稿。 */
const owned = new Set();

function parseMatrix(file) {
  const rows = [];
  // 只认 §2 主表。§4 的结论表里也有 `| wps_ppt_xxx |` 开头的行，
  // 照单全收会把"实测依据"那一长段当成安全参数 JSON 去跑 —— 所以按标题切段。
  let inMain = false;
  for (const line of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    if (/^##\s/.test(line)) {
      inMain = /^##\s*2\./.test(line);
      continue;
    }
    if (!inMain) continue;
    if (!/^\|\s*`?wps_/.test(line)) continue;
    const cells = line.split('|').map((c) => c.trim());
    const tool = (cells[1] || '').replace(/`/g, '');
    const argRaw = (cells[2] || '').replace(/`/g, '');
    const judge = cells[3] || '';
    const note = cells[4] || '';
    const extraRaw = (cells[5] || '').replace(/`/g, '');
    const row = { tool, args: {}, judge: judge.replace(/\s/g, ''), note, setup: [], check: null, skipped: judge.replace(/\s/g, '') === '跳过' };
    let args = {};
    try { args = JSON.parse(argRaw); } catch { /* 跳过行里的 <占位符> 不是 JSON，交给下面的 extra */ }
    row.args = args;
    if (extraRaw) {
      try {
        const extra = JSON.parse(extraRaw);
        row.setup = extra.setup || [];
        row.check = extra.check || null;
      } catch { /* 保持默认值 */ }
    }
    rows.push(row);
  }
  return rows;
}

/** 运行时占位符：文稿名是 WPS 随机给的（演示文稿83），写死必错 */
function fillTokens(v, target) {
  if (typeof v === 'string') {
    return v
      .replace(/\{\{TARGET\}\}/g, target || '')
      .replace(/\{\{PNG\}\}/g, png)
      .replace(/\{\{PNG2\}\}/g, png2)
      .replace(/\{\{SOURCE_PPTX\}\}/g, sourcePptx);
  }
  if (Array.isArray(v)) return v.map((x) => fillTokens(x, target));
  if (v && typeof v === 'object') {
    const out = {};
    for (const [k, x] of Object.entries(v)) out[k] = fillTokens(x, target);
    return out;
  }
  return v;
}

/** 列出当前打开的演示文稿名（去重） */
async function openNames() {
  try {
    const t = txt(await call('wps_ppt_get_open_presentations', {}));
    const out = new Set();
    for (const m of t.matchAll(/(\d+)\.\s*(\S+)\s*(?:\[当前活动\])?\s*路径:/g)) out.add(m[2]);
    return out;
  } catch {
    return new Set();
  }
}

async function closeOwned() {
  const names = await openNames();
  for (const n of names) {
    if (!owned.has(n)) continue;
    try { await call('wps_ppt_close_presentation', { name: n, save: false }); } catch { /* ignore */ }
    owned.delete(n);
  }
}

/**
 * 素材：启动时让 WPS 自己存一份 pptx。
 * 存完立刻关掉（save:false），后面每条用例要开要插都用这份文件。
 */
async function makeSourcePptx() {
  await closeOwned();
  const before = await openNames();
  await call('wps_ppt_create_presentation', {});
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_textbox', { slideIndex: 1, left: 40, top: 40, width: 300, height: 50, text: '素材页' });
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_textbox', { slideIndex: 2, left: 40, top: 40, width: 300, height: 50, text: '素材页2' });
  await call('wps_common_save_as', { filePath: sourcePptx, format: 'pptx' });
  // 存盘之后要【重新取一次名字】：saveAs 会把文稿从"演示文稿83"改成"source.pptx"，
  // 拿旧名字去 closeOwned 是关不掉的 —— 那样这份素材文稿会一路开到整个验证跑完，
  // 而且会让后面每条用例的"打开文稿数"基线都多 1。差集重新认领一次就干净了。
  for (const n of await openNames()) if (!before.has(n)) owned.add(n);
  await closeOwned();
  return fs.existsSync(sourcePptx);
}

async function rebuild() {
  await closeOwned();
  // 靶子锁要是上一条用例留下的（set_active_target 是全局状态），后面每条都会定位错文稿。
  await call('wps_ppt_set_active_target', { clear: true });
  // 只把【这次新建出来的】文稿记成自己的。
  // 以前是"建完之后把所有打开的文稿都记上" —— 万一用户正开着别的演示文稿，
  // 收尾就会把它当自己的关掉。差集才是安全的。
  const before = await openNames();
  await call('wps_ppt_create_presentation', {});
  const after = await openNames();
  for (const n of after) if (!before.has(n)) owned.add(n);
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_slide', { layout: 'blank' });
  for (const p of [1, 2]) {
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'rectangle', left: 60, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'oval', left: 240, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_textbox', { slideIndex: p, left: 60, top: 160, width: 240, height: 60, text: `文字${p}` });
    await call('wps_ppt_insert_table', { slideIndex: p, rows: 2, cols: 2 });
    await call('wps_ppt_insert_image', { slideIndex: p, path: png });
    await call('wps_ppt_set_shape_text', { slideIndex: p, shapeIndex: 1, text: `文字${p}` });
  }
  await call('wps_ppt_set_active_target', { clear: true });
  // TARGET = 本次新建的那个（不是列表里第一个 —— 用户自己的文稿可能排在前面）
  const nowOpen = await openNames();
  for (const n of nowOpen) if (!before.has(n)) return n;
  return undefined;
}

/**
 * 状态指纹：页数 / 每页形状数 / 每页动画数 / 母版元素数 / 打开的文稿数。
 * 母版和文稿数是第 9 轮加的 —— add_master_element 只动母版、close/open_presentation 只动文稿数，
 * 旧指纹看不见这两类变化，加了它们不会让别的用例误判（这两项对其余工具恒定不变）。
 */
async function fingerprint() {
  const parts = [];
  const c = await call('wps_ppt_get_slide_count', {});
  parts.push(txt(c).match(/(\d+)/)?.[1] || '0');
  for (const i of [1, 2]) {
    const s = await call('wps_ppt_get_shapes', { slideIndex: i });
    parts.push(String((txt(s).match(/共 (\d+) 个/) || [])[1] || '0'));
    const a = await call('wps_ppt_get_animations', { slideIndex: i });
    parts.push(String((txt(a).match(/共 (\d+) 个|(\d+) 个动画/) || [, '0'])[1] || '0'));
  }
  const m = await call('wps_ppt_get_slide_master', {});
  parts.push(String((txt(m).match(/母版元素: (\d+) 个/) || [])[1] || '0'));
  parts.push(String((await openNames()).size));
  return parts.join('/');
}

/** 坐标指纹：只取第 1 页形状的 (left, top)，用于"只改位置"这类工具的判定 */
async function coordFingerprint() {
  const s = await call('wps_ppt_get_shapes', { slideIndex: 1 });
  return (txt(s).match(/\((\d+(?:\.\d+)?),\s*(\d+(?:\.\d+)?)\)/g) || []).join(' ');
}

async function shot(tag) {
  const p = path.join(tmpDir, `s-${tag}.png`);
  await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: p });
  return fs.existsSync(p) ? fs.readFileSync(p) : null;
}

(async () => {
  const rows = parseMatrix(MATRIX);
  const haveSource = await makeSourcePptx();
  console.log(`素材：1×1 PNG=${png}\n素材：source.pptx=${sourcePptx}（${haveSource ? '已生成' : '生成失败'}）`);
  console.log(`=== 隔离验证：${rows.filter((r) => !r.skipped).length} 条执行 / ${rows.length} 行（跳过 ${rows.filter((r) => r.skipped).length}）===\n`);

  const passed = [];
  const failed = [];
  const notes = [];

  for (const row of rows) {
    if (row.skipped) {
      notes.push(`跳过 ${row.tool}：${row.note}`);
      continue;
    }
    const target = await rebuild();
    const openBefore = await openNames();
    const args = fillTokens(row.args, target);

    // 前置步骤：先跑完 setup，再取基线（不然 setup 造成的变化会被算成工具的效果）
    let setupErr = '';
    for (const s of row.setup) {
      const r = await call(s.tool, fillTokens(s.args, target));
      if (!r || r.success === false) setupErr = `前置 ${s.tool} 失败：${txt(r).slice(0, 80)}`;
      if (s.catch) continue;
      if (setupErr) break;
    }

    const fpBefore = await fingerprint();
    const coordBefore = await coordFingerprint();
    const imgBefore = await shot('a');
    const imgCheck = await shot('b');
    const stable = imgBefore && imgCheck && Buffer.compare(imgBefore, imgCheck) === 0;

    let r = null;
    let err = setupErr;
    if (!err) {
      try {
        r = await call(row.tool, args);
        if (!r || r.success === false) err = txt(r).slice(0, 90) || 'success=false';
      } catch (e) {
        err = `THROW ${e.message}`.slice(0, 90);
      }
    }

    let ok = !err;
    let why = '';
    if (ok) {
      if (row.judge.includes('数量差')) {
        const fpAfter = await fingerprint();
        ok = fpAfter !== fpBefore;
        why = `状态 ${fpBefore} → ${fpAfter}`;
      } else if (row.judge.includes('坐标差')) {
        const coordAfter = await coordFingerprint();
        ok = coordAfter !== coordBefore;
        why = ok ? '形状坐标有变化' : `坐标未变 ${coordAfter.slice(0, 60)}`;
      } else if (row.judge.includes('读回值')) {
        if (!row.check) { ok = false; why = '判据是读回值但矩阵没给 check'; }
        else {
          const cr = await call(row.check.tool, fillTokens(row.check.args || {}, target));
          const ct = txt(cr);
          // expect 也要填占位符 —— 断言"活动文稿名 == {{TARGET}}"这类用例就靠它
          const expect = String(fillTokens(row.check.expect, target));
          ok = !!cr && cr.success !== false && ct.includes(expect);
          why = ok ? `读回「${expect}」命中` : `读回未命中：${ct.slice(0, 80)}`;
        }
      } else if (row.judge.includes('读回')) {
        const t = txt(r);
        ok = t.length > 8 && !/undefined|null/.test(t.slice(0, 40));
        why = `读回 ${t.length} 字`;
      } else {
        const imgAfter = await shot('c');
        ok = !!stable && !!imgAfter && Buffer.compare(imgBefore, imgAfter) !== 0;
        why = stable ? (ok ? '导出图有变化' : '导出图无变化') : '导出不稳定';
      }
    }
    console.log(`${ok ? '✓' : '✗'} ${row.tool.padEnd(34)} ${ok ? why : err || why}`);
    (ok ? passed : failed).push(row.tool);

    // 用例自己新建的文稿（"再建一份"的前置、或 open_presentation 打开的素材）
    // 也要记进 owned，收尾才关得掉；同时只收【这条用例期间新出现的】名字，
    // 用户本来就开着的文稿不在其中，不会被误关。
    for (const n of await openNames()) {
      if (!openBefore.has(n)) owned.add(n);
    }
  }

  console.log(`\n通过 ${passed.length}/${passed.length + failed.length}`);
  if (failed.length) console.log(`失败：${failed.join(', ')}`);
  console.log('\n--- 跳过的用例及理由 ---');
  for (const n of notes) console.log('· ' + n);
  console.log('\nPASSED=' + passed.join(','));
  console.log('FAILED=' + failed.join(','));

  await closeOwned();
  // 兜底：确认没留下演示文稿（只报，不关用户的）
  const left = [...(await openNames())];
  if (left.length) console.log(`\n⚠ 收尾后仍打开：${left.join(' / ')}`);
  fs.rmSync(tmpDir, { recursive: true, force: true });
  process.exit(failed.length ? 1 : 0);
})();