/**
 * Input: 命令行模式（cold / timeout / idle / clean）
 * Output: 桥的端到端行为实测结果（打印到 stdout）
 * Pos: ①B 桥的手工验证脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用法（务必从包根目录跑，需先 npm run build）：
 *   node scripts/manual/bridge-test.js clean     # 清场：清掉测试桥与全部 wps-comd 子进程
 *   node scripts/manual/bridge-test.js cold      # 冷启动 + 连接复用 + 按应用拆子进程
 *   node scripts/manual/bridge-test.js timeout   # 超时只重建那一路，另一路 PID 不动
 *   node scripts/manual/bridge-test.js idle      # 空闲回收：桥与子进程都不留
 *
 * 用独立的管道名 + 独立空闲时长，避免影响正在用的生产桥。
 *
 * 关于"桥能否活过 MCP server 重启"
 * --------------------------------
 * 本题在本机沙箱里【不可验证】：实测连最小的 detached node 进程都活不过父进程退出
 * （见 scripts/manual/detached-survive-probe.js 的对照），是环境的 Job Object 行为。
 * 所以这里不测这一条；真实部署下是否成立需另测，别拿本脚本的结果当依据。
 */

const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '../..');
const MODE = process.argv[2] || 'cold';

// 必须在 require 客户端之前设好（模块加载时读环境变量）
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-selftest';
process.env.WPS_BRIDGE_START_TIMEOUT = '40000';
process.env.WPS_BRIDGE_IDLE_MS = MODE === 'idle' ? '6000' : '300000';

const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

function psCount(script) {
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script], {
      encoding: 'utf8',
    });
    return parseInt(out.trim(), 10) || 0;
  } catch (e) {
    if (process.env.WPS_TEST_DEBUG) {
      console.error('[psCount 失败]', e.message.slice(0, 200));
      if (e.stderr) console.error('[stderr]', String(e.stderr).slice(0, 400));
    }
    return -1;
  }
}

const countBridges = () =>
  psCount(
    "@(Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" -EA SilentlyContinue |" +
      " Where-Object { $_.CommandLine -like '*bridge*index.js*' }).Count"
  );

const killBridges = () =>
  psCount(
    "Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" -EA SilentlyContinue |" +
      " Where-Object { $_.CommandLine -like '*bridge*index.js*' -and $_.ProcessId -ne $PID } |" +
      ' ForEach-Object { Stop-Process -Id $_.ProcessId -Force -EA SilentlyContinue; 1 } |' +
      ' Measure-Object | Select-Object -ExpandProperty Count'
  );

const killPsChildren = () =>
  psCount(
    "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" -EA SilentlyContinue |" +
      // 关键：必须排除自己 —— 本命令的 CommandLine 里就含 'wps-comd.ps1' 字面量，
      // 不排除的话会把自己的 powershell 杀掉（实测踩过，表现为 PS 退出码非 0）
      " Where-Object { $_.CommandLine -like '*wps-comd.ps1*' -and $_.ProcessId -ne $PID } |" +
      ' ForEach-Object { Stop-Process -Id $_.ProcessId -Force -EA SilentlyContinue; 1 } |' +
      ' Measure-Object | Select-Object -ExpandProperty Count'
  );

const appPid = async (app) => {
  const st = await psDaemon.bridgeStatus();
  const c = st.data.children.find((x) => x.app === app);
  return c ? c.pid : null;
};

(async () => {
  if (MODE === 'clean') {
    console.log('=== 清场（测试桥 + 全部 wps-comd 子进程）===');
    console.log('  清掉桥进程:', killBridges(), '个');
    console.log('  清掉常驻 powershell:', killPsChildren(), '个');
    await new Promise((r) => setTimeout(r, 1500));
    console.log(`  清场后：桥 ${countBridges()} 个，子进程 ${countPsChildrenCount()}`);
    process.exit(0);
  }

  if (MODE === 'idle') {
    console.log('=== 空闲回收测试（空闲阈值 6 秒）===');
    await psDaemon.call('ping', {}, 'common');
    console.log('  调一次 ping 后：桥 =', countBridges(), ' 常驻 powershell =', countPsChildrenCount());
    // 关键：桥在【有客户端连着】时不该回收，所以必须先断开，再等超时
    psDaemon.dispose();
    console.log('  已断开客户端连接，等 12 秒（>2 倍阈值）...');
    await new Promise((r) => setTimeout(r, 12000));
    const b = countBridges();
    const c = countPsChildrenCount();
    console.log(`  回收后：桥 = ${b}（期望 0）  常驻 powershell = ${c}（期望 0）`);
    console.log(b === 0 && c === 0 ? '  ✅ 无残留' : '  ❌ 有残留');
    process.exit(b === 0 && c === 0 ? 0 : 1);
  }

  if (MODE === 'isolate') {
    console.log('=== 按应用隔离测试：外部干掉 excel 那一路，看 common 是否受影响 ===');
    await psDaemon.call('ping', {}, 'common');
    await psDaemon.call('getOpenWorkbooks', {}, 'excel');
    const commonBefore = await appPid('common');
    const excelBefore = await appPid('excel');
    console.log(`  初始：common PID=${commonBefore}  excel PID=${excelBefore}`);

    // 模拟"表格那一路挂了"：外部强杀它的 powershell
    const killed = psCount(
      `Stop-Process -Id ${excelBefore} -Force -EA SilentlyContinue; 1`
    );
    console.log(`  已强杀 excel 子进程（${killed >= 0 ? 'ok' : '失败'}），等 800ms`);
    await new Promise((r) => setTimeout(r, 800));

    const pingOk = await psDaemon
      .call('ping', {}, 'common')
      .then((r) => r.success === true)
      .catch(() => false);
    const commonAfter = await appPid('common');
    console.log(`  common 仍可用: ${pingOk ? '✅ 是' : '❌ 否'}`);
    console.log(
      `  common PID: ${commonBefore} -> ${commonAfter} ` +
        (commonBefore === commonAfter ? '✅ 未被牵连' : '❌ 被重建了')
    );

    const excelRetry = await psDaemon
      .call('getOpenWorkbooks', {}, 'excel')
      .then((r) => r.success === true)
      .catch((e) => '失败:' + e.message.slice(0, 60));
    const excelAfter = await appPid('excel');
    console.log(`  excel 自动重建后可再用: ${excelRetry === true ? '✅ 是' : '❌ ' + excelRetry}`);
    console.log(
      `  excel PID: ${excelBefore} -> ${excelAfter} ` +
        (excelAfter && excelAfter !== excelBefore ? '✅ 已换新进程' : '❌ 没重建')
    );

    const pass =
      pingOk && commonBefore === commonAfter && excelRetry === true && excelAfter !== excelBefore;
    process.exit(pass ? 0 : 1);
  }

  if (MODE === 'clienttimeout') {
    console.log('=== 客户端超时与重连测试 ===');
    await psDaemon.call('ping', {}, 'common');
    const pidBefore = await appPid('common');

    // 把 common 那一路干掉，让下一次调用必须冷启动新子进程（约 4 秒）——
    // 这样 1ms 超时必然先触发，避免"1ms vs 2ms 往返"的竞态（上一版就是这么 flake 的）
    psCount(`Stop-Process -Id ${pidBefore} -Force -EA SilentlyContinue; 1`);
    await new Promise((r) => setTimeout(r, 500));

    process.env.WPS_DAEMON_TIMEOUT = '1'; // 运行时收紧，下次调用必超时
    let msg = '';
    try {
      await psDaemon.call('ping', {}, 'common');
    } catch (e) {
      msg = e.message;
    }
    console.log(`  超时是否如实抛出: ${/超时/.test(msg) ? '✅ 是' : '❌ 否'}  → "${msg.slice(0, 70)}"`);
    console.log(`    是否自动重放: ${/不自动重放/.test(msg) ? '✅ 未重放（符合红线）' : '❌ 有重放风险'}`);

    process.env.WPS_DAEMON_TIMEOUT = '120000';
    const recovered = await psDaemon
      .call('ping', {}, 'common')
      .then((r) => r.success === true)
      .catch(() => false);
    console.log(`  断开后自动重连恢复: ${recovered ? '✅ 是' : '❌ 否'}`);
    const pidAfter = await appPid('common');
    console.log(`  common 子进程: ${pidBefore} -> ${pidAfter}（被我们杀过，应换新）`);
    process.exit(recovered && /超时/.test(msg) ? 0 : 1);
  }

  // cold
  console.log('=== 冷启动 ===');
  console.log('  起始：桥 =', countBridges(), ' 常驻 powershell =', countPsChildrenCount());

  const t0 = Date.now();
  const r1 = await psDaemon.call('ping', {}, 'common');
  console.log(`  第 1 次 ping: ${Date.now() - t0}ms  success=${r1.success}`);

  const times = [];
  for (let i = 0; i < 5; i++) {
    const a = Date.now();
    const r = await psDaemon.call('ping', {}, 'common');
    times.push(Date.now() - a);
    if (!r.success) throw new Error('ping 失败: ' + JSON.stringify(r));
  }
  console.log(`  之后 5 次 ping: ${times.join(' / ')} ms`);

  const t2 = Date.now();
  const r2 = await psDaemon.call('getOpenWorkbooks', {}, 'excel');
  console.log(`  excel getOpenWorkbooks: ${Date.now() - t2}ms  success=${r2.success}`);

  const st = await psDaemon.bridgeStatus();
  console.log('  桥自身状态:');
  console.log('    bridgePid =', st.data.bridgePid, ' clients =', st.data.clients);
  console.log(
    '    children  =',
    JSON.stringify(st.data.children.map((c) => ({ app: c.app, pid: c.pid, ready: c.ready })))
  );

  const pidA = await appPid('common');
  await psDaemon.call('ping', {}, 'common');
  const pidB = await appPid('common');
  console.log(`  common 子进程复用: ${pidA} -> ${pidB} ${pidA === pidB ? '✅ 同一进程' : '❌ 被重建'}`);

  const pids = st.data.children.map((c) => c.pid);
  console.log(
    `  按应用拆子进程：apps=${JSON.stringify(st.data.children.map((c) => c.app))} 各 PID 互不相同=` +
      (new Set(pids).size === pids.length ? '✅' : '❌')
  );

  console.log('  当前：桥 =', countBridges(), ' 常驻 powershell =', countPsChildrenCount());
  process.exit(0);
})().catch((err) => {
  console.error('测试失败:', err && err.message ? err.message : err);
  process.exit(1);
});

function countPsChildrenCount() {
  return psCount(
    "@(Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" -EA SilentlyContinue |" +
      " Where-Object { $_.CommandLine -like '*wps-comd.ps1*' -and $_.ProcessId -ne $PID }).Count"
  );
}
