/**
 * Input: 一个【没保存过】的 Excel 工作簿 + wps_excel_close_workbook({save:true})
 * Output: 「勿重试·先另存为」护栏是否生效；有没有弹框/卡死（秒表计，>10s 判失败）
 * Pos: 第 3 轮 #5 修复的专项验收（INBOX 第 3 轮 · 验证者 · 第 2 件事）。一旦我被修改，请更新头部注释与所属目录的md。
 *
 * 为什么这是本轮最关键的验收
 * ------------------------
 * 病根：TS 发 `save`、PS 读 `saveChanges`（缺省 $true）→ 关未保存工作簿时 `Close($true)`
 * 弹「另存为」模态框，而 COM 是 STA —— 弹框会把桥通道卡死，实测要等满 120s 超时。
 * 修法（任务卡要求）：照 Word 侧 `closeDocument` 的护栏（scripts/wps-com.ps1:3214-3218 附近），
 * 未保存 + 要求保存 → 返回明确的 `[勿重试·先另存为]`，而不是弹框。
 *
 * 三个用例：
 *   A. 未保存 + save:true   → 必须【快速失败】并提示先另存为（不弹框、不卡死）
 *   B. 未保存 + save:false  → 必须能关掉（逃生口得留着）
 *   C. 已 save_as 到临时目录 + save:true → 必须正常保存并关闭（护栏不许误伤正常路径）
 *
 * 计时口径：每次调用单独计时；单次超过 10 秒判失败（任务卡原文）。
 * 超时兜底：WPS_DAEMON_TIMEOUT 压到 20s —— 真卡死时 20s 超时也能被 10s 判据抓住，
 * 且比默认 120s 少浪费 100 秒。
 *
 * 收尾：只关自己新建的工作簿（按新建时从列表里认出的新名字），用户已开的一个都不碰。
 *
 * 用法（需先 npm run build）：node scripts/manual/excel-close-guard-verify.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

// 独立管道 + EPHEMERAL：卡死也只卡我自己这条桥，不碰用户正在用的那条
// 管道名带时间戳：固定名会连上还活着的旧桥，桥内存里是旧版 wps-com.ps1（本轮已踩过）
const PIPE_SUFFIX = `agentC-excelguard-${Date.now()}`;
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

/** 单次调用超过这么久算失败（任务卡口径） */
const LIMIT_MS = 10000;

const results = [];
const check = (label, ok, detail = '') => {
  results.push({ label, ok });
  console.log(`  ${ok ? '✓' : '✗'} ${label}${detail ? ' — ' + detail : ''}`);
};
const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';

/** 计时调用：返回 { r, ms }；超时会走 tool 的 catch 分支返回 success=false */
async function timed(name, args = {}) {
  const t0 = Date.now();
  const r = await call(name, args);
  return { r, ms: Date.now() - t0 };
}

/** 当前打开的工作簿名列表（裸调，拿得到 name） */
async function workbookNames() {
  const r = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  const names = [];
  const re = /"name":"([^"]*)"/g;
  let m;
  while ((m = re.exec(txt(r)))) names.push(m[1]);
  return names;
}

/** 打开的工作簿明细（name + path），用于 save_as 后重新认名字 */
async function workbookList() {
  const r = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  try { return JSON.parse(txt(r)).data.workbooks || []; } catch { return []; }
}

(async () => {
  const tmpDir = path.join(os.tmpdir(), `wps-excel-guard-${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });
  console.log('\n=== Excel 关闭护栏专项验收（#5 修复）===');
  console.log(`管道：${process.env.WPS_BRIDGE_PIPE}｜单次调用上限：${LIMIT_MS}ms`);

  // 探活：WPS 表格起不来时按约定跳过，不判失败
  const hb = await call('wps_common_ping', {});
  if (!hb.success) {
    console.log(`  ⚠ WPS 不可用（${txt(hb).slice(0, 80)}），本轮跳过，不判失败`);
    console.log('    （#5 的护栏在真机上【未验证】）');
    fs.rmSync(tmpDir, { recursive: true, force: true });
    process.exit(0);
  }

  // 记账：探测前已开的工作簿 + 我新建的（收尾只关自己的）
  const before = await workbookNames();
  const mine = [];
  console.log(`  探测前已打开的工作簿：${before.length ? before.join(' / ') : '(无)'}`);

  /** 新建工作簿并认出它的名字：用 TOOL 建（稳），再轮询 getOpenWorkbooks 差分认名字（避免单拍 race） */
  const createMine = async () => {
    const created = await call('wps_excel_create_workbook', {});
    if (!created.success) return null;
    for (let i = 0; i < 8; i++) {
      const after = await workbookNames();
      const fresh = after.find((n) => !before.includes(n) && !mine.includes(n));
      if (fresh) {
        mine.push(fresh);
        return fresh;
      }
      await new Promise((r) => setTimeout(r, 200));
    }
    return null;
  };

  /** 收尾：按名字关掉我建的；用户的一个都不碰 */
  const cleanup = async () => {
    // 上一轮泄漏的临时工作簿（guard-probe.xlsx）也一并收掉
    const allMine = [...mine];
    if (before.includes('guard-probe.xlsx')) allMine.push('guard-probe.xlsx');
    for (const n of allMine) {
      try {
        await call('wps_execute_method', {
          method: 'closeWorkbook', appType: 'excel', params: { name: n, saveChanges: false },
        });
      } catch { /* 已经关了 */ }
    }
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { /* 占用 */ }
  };

  try {
    // 通道健康对照：证明"慢"是 close 调用造成的，不是环境慢
    const warm = await timed('wps_excel_get_open_workbooks', {});
    console.log(`  · 对照：get_open_workbooks 耗时 ${warm.ms}ms（success=${warm.r.success}）`);

    // ---------- A. 未保存 + save:true（本轮最关键）----------
    console.log('\n--- A. 未保存工作簿 + close_workbook({save:true}) ---');
    const nameA = await createMine();
    if (!nameA) {
      check('新建未保存工作簿（A 前置）', false, '建不起来（createMine 未认出新建工作簿）');
    } else {
      const a = await timed('wps_excel_close_workbook', { name: nameA, save: true });
      const msg = txt(a.r);
      console.log(`    返回：success=${a.r.success}｜耗时 ${a.ms}ms`);
      console.log(`    原文：${msg.split('\n')[0].slice(0, 120)}`);
      check(
        'A1 未保存 + 要求保存 → 明确失败（不是弹框、不是假成功）',
        a.r.success === false,
        `success=${a.r.success}`
      );
      check(
        'A2 不卡死：10 秒内返回',
        a.ms < LIMIT_MS,
        `实际 ${a.ms}ms（上限 ${LIMIT_MS}ms）${a.ms >= LIMIT_MS ? '｜超时原文：' + msg.slice(0, 80) : ''}`
      );
      check(
        'A3 报错说清"先另存为"',
        /另存为/.test(msg),
        msg.split('\n')[0].slice(0, 80)
      );
      const hasMarker = /\[勿重试[·・]?\s*先另存为\]/.test(msg);
      console.log(
        `  · 任务卡要求的标记 [勿重试·先另存为]：${hasMarker ? '有' : '没有（措辞不一致请统筹者裁决）'}`
      );
      // 护栏生效后工作簿应该还开着（没被关掉）
      const stillOpen = (await workbookNames()).includes(nameA);
      check('A4 护栏生效后工作簿仍在（没有被硬关）', stillOpen, stillOpen ? nameA : '已不在列表');
    }

    // ---------- B. 未保存 + save:false（逃生口）----------
    // 注意：必须用【新建】的工作簿，不能复用 A 的（A 若护栏生效则仍在、若失效则已被关，都不是干净起点）
    console.log('\n--- B. 未保存工作簿 + close_workbook({save:false}) ---');
    const nameB = await createMine();
    if (!nameB) {
      check('新建未保存工作簿（B 前置）', false, '建不起来');
    } else {
      const b = await timed('wps_excel_close_workbook', { name: nameB, save: false });
      const closed = !(await workbookNames()).includes(nameB);
      check(
        'B1 save:false 能关掉且不卡死',
        b.r.success === true && closed && b.ms < LIMIT_MS,
        `success=${b.r.success}｜${b.ms}ms｜已关闭=${closed}｜${txt(b.r).split('\n')[0].slice(0, 60)}`
      );
      if (closed) {
        const i = mine.indexOf(nameB);
        if (i >= 0) mine.splice(i, 1);
      }
    }

    // ---------- C. 已保存 + save:true（护栏不许误伤正常路径）----------
    console.log('\n--- C. 已另存到临时目录 + close_workbook({save:true}) ---');
    const nameC = await createMine();
    if (!nameC) {
      check('新建工作簿（C 前置）', false, '建不起来');
    } else {
      const savePath = path.join(tmpDir, 'guard-probe.xlsx');
      let sv = await call('wps_common_save_as', { filePath: savePath });
      if (!sv.success) sv = await call('wps_common_save_as', { filePath: savePath }); // save_as 偶发超时，重试一次
      check('C0 先另存到临时目录', sv.success && fs.existsSync(savePath), txt(sv).split('\n')[0].slice(0, 60));
      if (sv.success) {
        // save_as 会把工作簿改名为文件名，必须重新认名字再去关
        const renamed = (await workbookList()).find((w) => w.path === savePath);
        const realName = renamed ? renamed.name : nameC;
        if (renamed) { const i = mine.indexOf(nameC); if (i >= 0) mine[i] = realName; }
        const c = await timed('wps_excel_close_workbook', { name: realName, save: true });
        const closed = !(await workbookNames()).includes(realName);
        check(
          'C1 已保存 + save:true 正常关闭（护栏没误伤）',
          c.r.success === true && closed && c.ms < LIMIT_MS,
          `success=${c.r.success}｜${c.ms}ms｜已关闭=${closed}｜${txt(c.r).split('\n')[0].slice(0, 60)}`
        );
        check('C2 另存出来的文件还在', fs.existsSync(savePath), savePath);
        if (closed) {
          const i = mine.indexOf(nameC);
          if (i >= 0) mine.splice(i, 1);
        }
      }
    }
  } catch (e) {
    console.log(`  ⚠ 探测中断：${e && e.message ? e.message : String(e)}`);
    console.log('    （中断部分不判失败，但也不算数 —— 按未验证计）');
  } finally {
    await cleanup();
    const after = await workbookNames();
    const leaked = after.filter((n) => mine.includes(n));
    console.log(`\n收尾：我建的 ${mine.length ? '已按名关闭' : '无需关闭'}；剩余工作簿 ${after.length} 个（${after.join(' / ') || '无'}）`);
    check('收尾没留下我建的工作簿', leaked.length === 0, leaked.join(',') || '无');

    const passed = results.filter((r) => r.ok).length;
    console.log(`\n通过 ${passed}/${results.length}`);
    // A2 是硬指标：卡死就是没修好，无论别的过没过
    const a2 = results.find((r) => r.label.startsWith('A2'));
    if (a2 && !a2.ok) {
      console.log('✗ A2 未过：close 仍会卡死通道 —— #5 视为未修复');
    }
    console.log(`PASSED=${passed}/${results.length}`);
    process.exit(passed === results.length ? 0 : 1);
  }
})().catch((e) => {
  console.error('探测失败:', e && e.message ? e.message : String(e));
  process.exit(1);
});
