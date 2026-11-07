/**
 * 孤儿桥生命周期验收（Office 收口合并卡 C 段）
 *
 * 病灶：WPS_BRIDGE_EPHEMERAL=1 时父 node 客户端异常退出（dispose 没跑到）→ 桥/pwsh 守护成孤儿
 * （二期执行者轮 9+ 条残留；判定实战孤儿桥占默认管道误报一轮）。
 * 修复：桥进程在 EPHEMERAL 模式下每秒探测 process.ppid，连续 2 秒 ESRCH = 父亡 → 收子进程退出。
 *
 * 验收（10 连跑，交替 crash/clean 两种父退方式）：每次起一个子 node 客户端（EPHEMERAL 桥 +
 * ping 全链——pwsh 守护会真实拉起），客户端死后轮询桥 pid 与 pwsh 守护数必须在 ≤5 秒内归零。
 * **不发真 COM 动作**（ping 在 ps1 层直接回话，不碰 WPS 实例）→ 可挂 FAST。
 * 默认桥不受影响的断言在 office-host-verify.js（源码静态：parentWatch 在 EPHEMERAL 条件内）。
 *
 * 判据：10/10 归零。判定归决策者。
 */
const { spawn, execFileSync } = require('child_process');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const ROUNDS = 10;
const ZERO_MS = 5000;   // 爪亡后桥/守护归零上限（父亡监测 2s + 收尾 1.5s，留足余量）

function pidAlive(pid) {
  try { process.kill(pid, 0); return true; } catch (e) { return e.code !== 'ESRCH'; }
}

function daemonCount() {
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command',
      "(Get-CimInstance Win32_Process -Filter \"Name='powershell.exe' or Name='pwsh.exe'\" | Where-Object { $_.CommandLine -match 'wps-comd' }).Count"],
      { encoding: 'utf8', timeout: 20000 });
    return parseInt(out.trim(), 10) || 0;
  } catch { return 0; }
}

/** 子客户端脚本：EPHEMERAL 桥 + ping（拉起 pwsh 守护），然后按 mode 退出 */
function childScript(pipe, mode) {
  return `
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_PIPE = ${JSON.stringify(pipe)};
const path = require('path');
const { psDaemon } = require(${JSON.stringify(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'))});
(async () => {
  const r = await psDaemon.call('ping', {}, 'common');
  if (!r.success) throw new Error('ping failed: ' + (r.error || ''));
  console.log('BRIDGEPID=' + (psDaemon.stats.bridgePid || 'none'));
  await new Promise((res) => setTimeout(res, 1200));
  if (${JSON.stringify(mode)} === 'crash') { process.kill(process.pid); }
  psDaemon.dispose();
  process.exit(0);
})().catch((e) => { console.error('child fail: ' + e.message); process.exit(2); });
`;
}

(async () => {
  console.log('=== 孤儿桥生命周期验收（10 连跑，crash/clean 交替）===');
  const preDaemons = daemonCount();
  console.log(`开场 pwsh 守护数：${preDaemons}（应接近 0）`);
  let okCount = 0, failCount = 0;
  for (let i = 1; i <= ROUNDS; i++) {
    const mode = i % 2 === 1 ? 'crash' : 'clean';
    const pipe = `\\\\.\\pipe\\wps-multitool-mcp-orphan-${Date.now()}-${i}`;
    const child = spawn(process.execPath, ['-e', childScript(pipe, mode)], { stdio: ['ignore', 'pipe', 'pipe'] });
    let bridgePid = null;
    await new Promise((resolve) => {
      let buf = '';
      const done = () => resolve();
      child.stdout.on('data', (c) => {
        buf += c.toString();
        const m = buf.match(/BRIDGEPID=(\d+)/);
        if (m && !bridgePid) bridgePid = parseInt(m[1], 10);
        if (bridgePid) done();
      });
      child.stderr.on('data', () => {});
      child.on('exit', () => setTimeout(done, 200));
      setTimeout(done, 30000);
    });
    // 等子客户端退出（crash 自杀 1.2s；clean dispose 后退）
    await new Promise((r) => setTimeout(r, mode === 'crash' ? 2200 : 3000));
    if (!bridgePid) {
      failCount++; console.log(`  ✗ 轮 ${i}（${mode}）：未拿到桥 pid`);
      continue;
    }
    // 轮询桥归零（≤ZERO_MS）；守护数同步降回基准
    const t0 = Date.now();
    let bridgeGone = false, daemonGone = false;
    while (Date.now() - t0 < ZERO_MS) {
      bridgeGone = !pidAlive(bridgePid);
      daemonGone = daemonCount() <= preDaemons;
      if (bridgeGone && daemonGone) break;
      await new Promise((r) => setTimeout(r, 300));
    }
    if (bridgeGone && daemonGone) {
      okCount++; console.log(`  ✓ 轮 ${i}（${mode}）：桥 ${bridgePid} + 守护在 ${Date.now() - t0}ms 内归零`);
    } else {
      failCount++; console.log(`  ✗ 轮 ${i}（${mode}）：bridgeGone=${bridgeGone} daemonGone=${daemonGone}（超 ${ZERO_MS}ms，孤儿）`);
      try { if (pidAlive(bridgePid)) process.kill(bridgePid); } catch { /* 收尾 */ }
    }
  }
  const leftover = daemonCount() - preDaemons;
  console.log(`\n汇总：${okCount}/10 归零；残留守护差值 ${leftover} —— ${failCount === 0 && leftover <= 0 ? '验收通过' : '验收失败'}`);
  process.exit(failCount === 0 && leftover <= 0 ? 0 : 1);
})();
