/**
 * Input: 命令行参数 launch | check
 * Output: 桥进程的启动/存活诊断（含 stdout/stderr 落盘）
 * Pos: 桥存活诊断工具。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用途：桥是 detached 进程，默认 stdio 被丢弃，崩了什么都看不到。
 * 本脚本把它挂到文件上，用来判断"桥到底有没有活过父进程"。
 *
 * 用法：
 *   node scripts/manual/bridge-daemon-check.js launch   # 拉起桥，等 6 秒，报告存活与日志
 *   node scripts/manual/bridge-daemon-check.js check    # 另起一个进程，看它是否还在
 */

const { spawn } = require('child_process');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '../..');
const PIPE = '\\\\.\\pipe\\wps-multitool-mcp-daemoncheck';
const LOGDIR = path.join(os.tmpdir(), 'wps-bridge-diag');

function psCount() {
  const s =
    "@(Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" -EA SilentlyContinue |" +
    " Where-Object { $_.CommandLine -like '*bridge*index.js*' }).Count";
  try {
    return parseInt(execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', s], { encoding: 'utf8' }).trim(), 10);
  } catch {
    return -1;
  }
}

const mode = process.argv[2] || 'launch';

if (mode === 'check') {
  console.log(`check：当前桥进程数 = ${psCount()}`);
  if (fs.existsSync(LOGDIR)) {
    for (const f of fs.readdirSync(LOGDIR)) {
      const p = path.join(LOGDIR, f);
      console.log(`--- ${f} (${fs.statSync(p).size} B) ---`);
      console.log(fs.readFileSync(p, 'utf8').slice(0, 1500));
    }
  }
  process.exit(0);
}

fs.rmSync(LOGDIR, { recursive: true, force: true });
fs.mkdirSync(LOGDIR, { recursive: true });
const outFd = fs.openSync(path.join(LOGDIR, 'bridge.out.log'), 'a');
const errFd = fs.openSync(path.join(LOGDIR, 'bridge.err.log'), 'a');

const child = spawn(process.execPath, [path.join(ROOT, 'dist', 'bridge', 'index.js')], {
  detached: true,
  stdio: ['ignore', outFd, errFd],
  windowsHide: true,
  env: { ...process.env, WPS_BRIDGE_PIPE: PIPE, WPS_BRIDGE_IDLE_MS: '300000' },
});
console.log(`launch：已拉起桥 PID=${child.pid}，管道=${PIPE}`);
child.unref();

setTimeout(() => {
  console.log(`  6 秒后：child.exitCode=${child.exitCode} signalCode=${child.signalCode}`);
  console.log(`  6 秒后：系统里桥进程数 = ${psCount()}`);
  for (const f of fs.readdirSync(LOGDIR)) {
    const p = path.join(LOGDIR, f);
    const body = fs.readFileSync(p, 'utf8');
    console.log(`  --- ${f} (${body.length} B) ---`);
    if (body) console.log(body.slice(0, 1200));
  }
  console.log('  父进程现在退出；随后用 check 模式看桥是否还在');
  process.exit(0);
}, 6000);
