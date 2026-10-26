/**
 * Input: 命令行参数 spawn | check
 * Output: 验证"detached 子进程能否活过父进程"（用最小 node 循环做对照）
 * Pos: 沙箱行为探测。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 背景：桥（dist/bridge/index.js）以 detached+unref 拉起，但实测父进程一退出桥就没了，
 * 且没有任何错误输出 —— 说明不是崩，是被"收走"。这里用最小对照把原因钉死：
 * 如果连 `node -e "setInterval(()=>{},1000)"` 都活不过父进程，那是环境（Job Object）行为，
 * 不是我们的桥写错了。
 */

const { spawn, execFileSync } = require('child_process');

const MARK = 'survive-probe-marker';

function countProbes() {
  const s =
    "@(Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" -EA SilentlyContinue |" +
    ` Where-Object { $_.CommandLine -like '*${MARK}*' }).Count`;
  try {
    return parseInt(execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', s], { encoding: 'utf8' }).trim(), 10);
  } catch {
    return -1;
  }
}

if ((process.argv[2] || 'spawn') === 'check') {
  console.log(`check：带标记的探测进程数 = ${countProbes()}（期望 1 表示"活过父进程"）`);
  process.exit(0);
}

// 用一个自打印标记的常驻 node 进程做对照
const code = `console.log('${MARK}'); setInterval(()=>{}, 1000);`;
const child = spawn(process.execPath, ['-e', code], { detached: true, stdio: 'ignore', windowsHide: true });
child.unref();
console.log(`spawn：探测进程 PID=${child.pid}（父进程即将先等 3 秒再正常退出）`);

setTimeout(() => {
  console.log(`  3 秒后仍在：${countProbes()} 个`);
  console.log('  父进程正常退出（exit 0）。随后用 check 模式看它还在不在。');
  process.exit(0);
}, 3000);
