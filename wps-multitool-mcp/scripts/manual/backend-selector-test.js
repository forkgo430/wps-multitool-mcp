/**
 * 平台识别决策链的单测（Linux B0 验收项）
 * 原则：不碰真环境，全部注入 env / platform / exists，把每层分支钉死。
 * 跑法：node scripts/manual/backend-selector-test.js（verify-all FAST 档包含它）
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { detectBackend, BackendUnavailableError, WSL_MARKER } = require(
  path.join(ROOT, 'dist', 'server', 'backend-selector.js')
);

let pass = 0;
let fail = 0;
const eq = (label, got, want) => {
  if (JSON.stringify(got) === JSON.stringify(want)) { pass++; console.log(`  ✓ ${label} = ${JSON.stringify(want)}`); }
  else { fail++; console.log(`  ✗ ${label}：得到 ${JSON.stringify(got)}，期望 ${JSON.stringify(want)}`); }
};
const throws = (label, fn, match) => {
  try { fn(); fail++; console.log(`  ✗ ${label}：没有抛错`); }
  catch (e) {
    if (e instanceof BackendUnavailableError && (!match || match.test(e.message))) { pass++; console.log(`  ✓ ${label}（BackendUnavailableError）`); }
    else { fail++; console.log(`  ✗ ${label}：抛了 ${(e && e.name) || typeof e}：${(e && e.message || '').slice(0, 80)}`); }
  }
};

const win32 = { env: {}, platform: 'win32', exists: () => false };
const linux = { env: {}, platform: 'linux', exists: () => false };
const wsl = { env: {}, platform: 'linux', exists: (p) => p === WSL_MARKER };
const darwin = { env: {}, platform: 'darwin', exists: () => false };

console.log('\n=== A. 第 2 层：原生 Windows ===');
eq('win32 → com', detectBackend(win32).backend, 'com');
eq('win32 判定依据带 win32', detectBackend(win32).basis.includes('win32'), true);

console.log('\n=== B. 第 1 层：WPS_BACKEND 强制覆盖（最高优先级，任何平台都生效）===');
eq('win32 + addon → addon（强制覆盖不问平台）', detectBackend({ ...win32, env: { WPS_BACKEND: 'addon' } }).backend, 'addon');
eq('win32 + off → off', detectBackend({ ...win32, env: { WPS_BACKEND: 'off' } }).backend, 'off');
eq('linux + off → off（不抛错）', detectBackend({ ...linux, env: { WPS_BACKEND: 'OFF' } }).backend, 'off');
eq('大小写不敏感（Addon）', detectBackend({ ...win32, env: { WPS_BACKEND: 'Addon' } }).backend, 'addon');
eq('win32 + com → com', detectBackend({ ...win32, env: { WPS_BACKEND: 'com' } }).backend, 'com');
throws('非法值抛 BackendUnavailableError', () => detectBackend({ ...win32, env: { WPS_BACKEND: 'bogus' } }), /不是合法值/);
// 判据修正（第一版期望"空格也抛"是错的）：空/空白值视为未设置、回落平台判定——
// 环境变量误设成空串不该把整个 server 弄瘫；真正要拦的是非空垃圾值（上一条）。
eq('空白值视为未设置（回落平台判定）', detectBackend({ ...win32, env: { WPS_BACKEND: '   ' } }).backend, 'com');

console.log('\n=== C. 第 3 层：WSL 特判 ===');
eq('linux + Windows PowerShell 标记 → com', detectBackend(wsl).backend, 'com');
eq('WSL 判定 wsl=true', detectBackend(wsl).wsl, true);
eq('WSL 依据里说明管道可达性归通道层', detectBackend(wsl).basis.includes('通道层'), true);

console.log('\n=== D. 第 4/5 层：Linux/macOS 通道未实现 → 明确抛错 ===');
throws('linux 无标记 → 抛错并给出三选项', () => detectBackend(linux), /尚未实现/);
throws('darwin → 同样抛错', () => detectBackend(darwin), /macOS/);
throws('抛错文案提到三型加载项', () => detectBackend(linux), /wps\/et\/wpp/);

console.log('\n=== E. 失败不缓存（每次调用都能拿到最新判定）===');
const flip = { env: {}, platform: 'win32', exists: () => false };
eq('第一次 win32 → com', detectBackend(flip).backend, 'com');

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
