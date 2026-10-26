/**
 * ①A 验证：常驻守护进程 vs 直跑模式，逐次计时对比。
 * 用法：
 *   node bench-daemon.js            # 守护进程
 *   node bench-daemon.js --legacy   # 强制直跑（WPS_USE_DAEMON=0）
 */
const path = require('path');
const { execSync } = require('child_process');

if (process.argv.includes('--legacy')) {
  process.env.WPS_USE_DAEMON = '0';
}

const wps = require('../../../dist/client/wps-client.js');
const client = wps.wpsClient || wps.default;

function countPowerShell() {
  try {
    const out = execSync(
      'tasklist /FI "IMAGENAME eq powershell.exe" /FO CSV',
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }
    );
    return out.split(/\r?\n/).filter((l) => l.toLowerCase().startsWith('"powershell.exe"')).length;
  } catch {
    return -1;
  }
}

const MODE = process.env.WPS_USE_DAEMON === '0' ? '直跑（每次新起 powershell）' : '常驻守护进程';

(async () => {
  console.log(`=== 模式: ${MODE} ===`);
  console.log('起始 powershell 进程数:', countPowerShell());

  const times = [];
  const N = 6;
  for (let i = 1; i <= N; i++) {
    const t = Date.now();
    const r = await client.invokeAction('ping', {});
    const dt = Date.now() - t;
    times.push(dt);
    console.log(`  #${i} ping  ${String(dt).padStart(5)}ms  success=${r.success}`);
  }
  console.log('跑完 N 次后 powershell 进程数:', countPowerShell());
  console.log(`总耗时 ${times.reduce((a, b) => a + b, 0)}ms  平均 ${Math.round(times.reduce((a, b) => a + b, 0) / N)}ms`);
  console.log(`第 1 次(含冷启动) ${times[0]}ms   第 2 次起平均 ${Math.round(times.slice(1).reduce((a, b) => a + b, 0) / (N - 1))}ms`);

  // 一个真实 COM 动作，确认走的还是同一套
  const t2 = Date.now();
  const r2 = await client.invokeAction('getOpenWorkbooks', {});
  console.log(`getOpenWorkbooks ${Date.now() - t2}ms ->`, JSON.stringify(r2));

  // 收尾：显式 dispose，确认守护进程会跟着退（不留孤儿）
  const { psDaemon } = require('../../../dist/client/ps-daemon.js');
  psDaemon.dispose();
  await new Promise((r) => setTimeout(r, 1500));
  console.log('dispose 后 powershell 进程数:', countPowerShell());
  process.exit(0);
})();
