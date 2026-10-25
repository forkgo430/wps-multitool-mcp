/**
 * Input: 无
 * Output: 直接把原始 JSON 行喂给 wps-comd.ps1，看参数在 PS 侧是否完整
 * Pos: 参数链路定位脚本（绕过 Node 桥）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 */

const { spawn } = require('child_process');
const path = require('path');

const PS = path.join(__dirname, '..', 'wps-comd.ps1');

const reqs = [
  { id: 'r1', action: '__probeParams', params: { count: 2, text: '中文', arr: [1, 2, 3] } },
  { id: 'r2', action: '__probeParams', params: {} },
  { id: 'r3', action: 'ping', params: { anything: 1 } },
];

const child = spawn('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', PS], {
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true,
});

let out = '';
let err = '';
child.stdout.on('data', (d) => (out += d.toString('utf8')));
child.stderr.on('data', (d) => (err += d.toString('utf8')));

child.on('close', () => {
  console.log('=== stdout 逐行 ===');
  out.split('\n').filter(Boolean).forEach((l, i) => console.log(`  [${i + 1}] ${l}`));
  if (err.trim()) console.log('=== stderr ===\n' + err.slice(0, 800));
  process.exit(0);
});

// 先写一行空行触发不了握手，这里直接按协议逐行写
for (const r of reqs) {
  child.stdin.write(JSON.stringify(r) + '\n');
}
setTimeout(() => child.stdin.end(), 1500);
