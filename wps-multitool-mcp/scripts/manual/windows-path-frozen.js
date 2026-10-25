/**
 * Windows 链路冻结门禁（Linux 计划 C1 约束的执法工具，B1）
 * ------------------------------------------------------
 * C1：Windows COM 链路零改动。为 Linux 加文件可以，为了 Linux 去改 Windows 文件不行。
 * 本门禁把 Windows 专用文件的 SHA256 锁进基线，任何 drift 都红。
 *
 * 基线：windows-frozen-baseline.json（路径 → sha256）。
 *   重立基线（仅限 Windows 自身的正当修复）：node scripts/manual/windows-path-frozen.js --update
 *   —— 必须在同一个提交里写明改动理由；CI 跑的是不带 --update 的校验模式。
 *
 * wps-com.ps1 有意不在清单里：它是持续演进的 COM 真相源（修 bug 加动作都动它），
 * 整文件哈希锁不现实；它的 COM 分支保护靠评审纪律 + B5 的 backend-parity 门禁。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.resolve(__dirname, '..', '..'); // → wps-multitool-mcp/
const BASELINE = path.join(__dirname, 'windows-frozen-baseline.json');

/** 冻结清单：Windows COM 链路的全部纯 Windows 文件（相对 wps-multitool-mcp/） */
const FROZEN = [
  'src/bridge/index.ts',
  'src/bridge/ps-child.ts',
  'src/bridge/cancellable.ts',
  'src/client/ps-daemon.ts',
  'scripts/wps-comd.ps1',
];

// 哈希前做换行归一化（CRLF/LF）：git 的 autocrlf 会在 checkout 时改行尾，
// 字节级哈希会在"git 认为没变"的文件上漂移（2026-10-25 红证时踩到）。
// 归一化后，行尾变化不再误报，真正的内容改动照样红。
const sha256 = (p) => {
  const normalized = fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');
  return crypto.createHash('sha256').update(normalized, 'utf8').digest('hex');
};

if (process.argv.includes('--update')) {
  const files = {};
  for (const rel of FROZEN) {
    const abs = path.join(ROOT, rel);
    if (!fs.existsSync(abs)) { console.error(`✗ 基线更新失败：文件不存在 ${rel}`); process.exit(1); }
    files[rel] = sha256(abs);
  }
  const doc = {
    _note: [
      'C1 冻结门禁基线。这些文件是 Windows COM 链路——为 Linux 改它们门禁就红。',
      'Windows 自身修 bug 需要改时：node scripts/manual/windows-path-frozen.js --update 重立基线，',
      '并在同一提交里写明理由。wps-com.ps1 不进本表（持续演进的 COM 真相源，保护靠纪律 + B5 parity 门禁）。',
      '基线建立：2026-10-25 第五任（pwsh7 兼容与退出护栏已包含在内）。',
    ],
    files,
  };
  fs.writeFileSync(BASELINE, JSON.stringify(doc, null, 2) + '\n');
  console.log(`✓ 基线已重立（${FROZEN.length} 个文件）`);
  process.exit(0);
}

let fail = 0;
if (!fs.existsSync(BASELINE)) {
  console.error('✗ 基线文件不存在，先跑 --update 建立');
  process.exit(1);
}
const base = JSON.parse(fs.readFileSync(BASELINE, 'utf8'));
for (const rel of FROZEN) {
  const abs = path.join(ROOT, rel);
  if (!fs.existsSync(abs)) {
    fail++;
    console.error(`  ✗ ${rel} —— 文件被删除！`);
    continue;
  }
  const now = sha256(abs);
  const want = (base.files || {})[rel];
  if (!want) {
    fail++;
    console.error(`  ✗ ${rel} —— 不在基线里（基线损坏或清单与基线不同步）`);
  } else if (now !== want) {
    fail++;
    console.error(`  ✗ ${rel} —— 内容与冻结基线不一致（C1 违规嫌疑）`);
    console.error(`      基线 ${want.slice(0, 16)}… / 当前 ${now.slice(0, 16)}…`);
  } else {
    console.log(`  ✓ ${rel}`);
  }
}
// 基线里有、清单里没有的（清单被删项 = 有人想绕过冻结）
for (const rel of Object.keys(base.files || {})) {
  if (!FROZEN.includes(rel)) {
    fail++;
    console.error(`  ✗ 基线里的 ${rel} 不在冻结清单中（清单被改动）`);
  }
}

if (fail) {
  console.error(`\n✗ C1 冻结门禁：${fail} 处违规。`);
  console.error('  Windows 自身正当修复 → 跑 --update 重立基线并在提交里写明理由；');
  console.error('  为 Linux 而改 → 停手，Linux 侧只允许新增文件（C1 约束原文见 linux/PLAN-LINUX-BACKEND.md §1）。');
  process.exit(1);
}
console.log(`\n✓ C1 冻结门禁：${FROZEN.length} 个 Windows 链路文件与基线一致`);
