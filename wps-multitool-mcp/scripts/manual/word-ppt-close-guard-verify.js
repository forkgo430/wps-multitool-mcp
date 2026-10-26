// Word / PPT 关闭护栏专项验证（第 5 轮 验证者 任务 2）
// 背景：统筹者把三处护栏判定信号从 $X.FullName 改成 $X.Path（wps-com.ps1:3373/3465/4283）。
// Excel 已验过；本脚本补验 Word 与 PPT。
//
// 用例（两个应用各一份）：
//   A：新建【未保存】文档/演示文稿 → close({save:true})
//      → 必须 success===false、10 秒内返回、消息说清"先另存为"、且文档/演示文稿仍在（没被静默关掉）
//   B：同一个未保存对象 → close({save:false}) → 必须能正常关掉（逃生口得留着）
//
// 隔离：独立时间戳管道 + EPHEMERAL；收尾只按名字关自己建的；用户已开的绝不碰。
// 用法：node scripts/manual/word-ppt-close-guard-verify.js

const fs = require('fs');
const os = require('os');
const path = require('path');

const PIPE_SUFFIX = `agentC-wordpptguard-${Date.now()}`;
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-${PIPE_SUFFIX}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = '20000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => (r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '';
const jq = (s, key) => {
  try { return JSON.parse(s)[key]; } catch { return undefined; }
};

// 列出某应用已打开对象的名字（裸调，拿得到 name）
async function namesOf(appType, action) {
  const r = await call('wps_execute_method', { method: action, appType, params: {} });
  const names = [];
  const re = /"name":"([^"]*)"/g;
  let m;
  while ((m = re.exec(txt(r)))) names.push(m[1]);
  return names;
}
// 列出带 path 的明细（用于确认"仍在/已关"以及 save_as 后改名）
async function listOf(appType, action) {
  const r = await call('wps_execute_method', { method: action, appType, params: {} });
  try {
    const k = appType === 'word' ? 'documents' : 'presentations';
    return jq(txt(r), 'data')[k] || [];
  } catch { return []; }
}

// 新建并认出名字：用 TOOL 建（稳），再轮询列表差分认名字
async function createMine(toolName, appType, action, before, mine) {
  const created = await call(toolName, {});
  if (!created.success) return null;
  for (let i = 0; i < 8; i++) {
    const after = await namesOf(appType, action);
    const fresh = after.find((n) => !before.includes(n) && !mine.includes(n));
    if (fresh) { mine.push(fresh); return fresh; }
    await new Promise((r) => setTimeout(r, 200));
  }
  return null;
}

const LIMIT_MS = 10000;
const results = [];
const check = (name, pass, detail) => {
  results.push({ name, pass, detail });
  console.log(`${pass ? '  ✓' : '  ✗'} ${name} — ${detail}`);
};

async function timed(toolName, args) {
  const t0 = Date.now();
  let r;
  try { r = await call(toolName, args); } catch (e) { r = { success: false, error: String(e && e.message || e) }; }
  const ms = Date.now() - t0;
  return { r, ms };
}

async function runOneApp({ appType, createAction, listAction, createTool, closeTool, label }) {
  console.log(`\n=== ${label} ===`);
  const before = await namesOf(appType, listAction);
  const mine = [];
  console.log(`探测前已打开的${label}：${before.length ? before.join(' / ') : '(无)'}`);

  const name = await createMine(createTool, appType, listAction, before, mine);
  if (!name) {
    check(`${label} 前置：新建未保存${label}（A 前置）`, false, '建不起来');
    return mine;
  }

  // ---- A：未保存 + save:true 必须报错、10s 内、文档仍在 ----
  const a = await timed(closeTool, { name, save: true });
  const stillOpen = (await namesOf(appType, listAction)).includes(name);
  const msg = txt(a.r).split('\n')[0].slice(0, 60);
  check(
    `${label} A1 save:true 必须报错（护栏触发）`,
    a.r.success === false,
    `success=${a.r.success}｜${msg}`
  );
  check(`${label} A2 不卡死（<10s）`, a.ms < LIMIT_MS, `${a.ms}ms（上限 ${LIMIT_MS}ms）`);
  check(
    `${label} A3 报错说清"先另存为"`,
    /另存为|先保存|save[_ ]?as/i.test(txt(a.r)),
    msg
  );
  check(`${label} A4 护栏生效后${label}仍在（没被静默关掉）`, stillOpen === true, `仍在=${stillOpen}`);

  // ---- B：同一个未保存对象 + save:false 必须能关掉（逃生口）----
  let nameB = name;
  // A 若护栏生效，name 仍在；A 若失效（bug），name 已被关 → B 无对象可测
  if (!stillOpen) {
    check(`${label} B1 save:false 逃生口`, false, 'A 段护栏没生效，对象已被关掉，逃生口无从测（护栏本身即 bug）');
  } else {
    const b = await timed(closeTool, { name: nameB, save: false });
    const closed = !(await namesOf(appType, listAction)).includes(nameB);
    check(
      `${label} B1 save:false 逃生口能关掉且不卡死`,
      b.r.success === true && closed && b.ms < LIMIT_MS,
      `success=${b.r.success}｜${b.ms}ms｜已关闭=${closed}｜${txt(b.r).split('\n')[0].slice(0, 50)}`
    );
    if (closed) {
      const i = mine.indexOf(nameB);
      if (i >= 0) mine.splice(i, 1);
    }
  }
  return mine;
}

(async () => {
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-wpg-'));
  const wordMine = [];
  const pptMine = [];
  try {
    Object.assign(process.env, {});
    const wm = await runOneApp({
      appType: 'word',
      createAction: 'createDocument',
      listAction: 'getOpenDocuments',
      createTool: 'wps_word_create_document',
      closeTool: 'wps_word_close_document',
      label: 'Word 文档',
    });
    wm.forEach((n) => { if (!wordMine.includes(n)) wordMine.push(n); });

    const pm = await runOneApp({
      appType: 'presentation',
      createAction: 'createPresentation',
      listAction: 'getOpenPresentations',
      createTool: 'wps_ppt_create_presentation',
      closeTool: 'wps_ppt_close_presentation',
      label: 'PPT 演示文稿',
    });
    pm.forEach((n) => { if (!pptMine.includes(n)) pptMine.push(n); });
  } catch (e) {
    console.error('运行异常:', e && e.message ? e.message : String(e));
  } finally {
    // 收尾：只按名字关自己建的；save:false 不保存
    const cleanup = async (appType, listAction, names) => {
      for (const n of names) {
        try {
          await call('wps_execute_method', {
            method: appType === 'word' ? 'closeDocument' : 'closePresentation',
            appType, params: { name: n, saveChanges: false },
          });
        } catch { /* 已关 */ }
      }
    };
    await cleanup('word', 'getOpenDocuments', wordMine);
    await cleanup('presentation', 'getOpenPresentations', pptMine);
    try { fs.rmSync(tmpRoot, { recursive: true, force: true }); } catch { /* 占用 */ }

    const wRemain = await namesOf('word', 'getOpenDocuments');
    const pRemain = await namesOf('presentation', 'getOpenPresentations');
    console.log(`\n剩余 Word 文档: ${wRemain.length} 个｜剩余 PPT 文稿: ${pRemain.length} 个`);
  }

  const pass = results.filter((r) => r.pass).length;
  console.log(`\nWord/PPT 护栏专项：${pass}/${results.length} 通过`);
  process.exit(0);
})();
