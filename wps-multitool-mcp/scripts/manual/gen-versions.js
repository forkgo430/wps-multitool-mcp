/**
 * gen-versions.js —— 版本号单源分发（2026-10-25，外部评审 P1-8）
 *
 * 读 wps-multitool-mcp/package.json 的 version，回写三处漂移点：
 *   1) .claude-plugin/marketplace.json → metadata.version
 *   2) skills/wps-multitool/SKILL.md   → frontmatter metadata.version
 *   3) src/version.ts                  → FALLBACK.version
 * 已挂进 npm run gen:docs 链：改完版本号跑一次 gen:docs，三处自动跟上。
 *
 * root 的 package.json（private 聚合层）**有意不同步**：它不是发版物，
 * 版本独立演进。本脚本不改它，也不对它报警。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..'); // → wps-multitool-mcp/
const REPO = path.resolve(ROOT, '..');            // → 仓库根

const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
const version = pkg.version;
if (!/^\d+\.\d+\.\d+/.test(version)) {
  console.error(`✗ 包版本号异常：${version}`);
  process.exit(1);
}

const changes = [];

// 1) marketplace.json
const mpPath = path.join(REPO, '.claude-plugin', 'marketplace.json');
const mp = JSON.parse(fs.readFileSync(mpPath, 'utf8'));
if (!mp.metadata || mp.metadata.version !== version) {
  mp.metadata = mp.metadata || {};
  mp.metadata.version = version;
  fs.writeFileSync(mpPath, JSON.stringify(mp, null, 2) + '\n');
  changes.push(`marketplace.json → ${version}`);
}

// 2) SKILL.md frontmatter（metadata.version）
const skillPath = path.join(ROOT, 'skills', 'wps-multitool', 'SKILL.md');
const skillRaw = fs.readFileSync(skillPath, 'utf8');
const skillRe = /(\n  version: ")[^"]*(")/;
if (skillRe.test(skillRaw)) {
  const updated = skillRaw.replace(skillRe, `$1${version}$2`);
  if (updated !== skillRaw) {
    fs.writeFileSync(skillPath, updated);
    changes.push(`SKILL.md → ${version}`);
  }
} else {
  console.error('✗ SKILL.md 里找不到 metadata.version 行（格式变了？手工核对）');
  process.exit(1);
}

// 3) src/version.ts 的 FALLBACK
const vtPath = path.join(ROOT, 'src', 'version.ts');
const vtRaw = fs.readFileSync(vtPath, 'utf8');
const vtRe = /(const FALLBACK = \{[\s\S]*?version: ')[^']*(',)/;
if (vtRe.test(vtRaw)) {
  const updated = vtRaw.replace(vtRe, `$1${version}$2`);
  if (updated !== vtRaw) {
    fs.writeFileSync(vtPath, updated);
    changes.push(`src/version.ts → ${version}`);
  }
} else {
  console.error('✗ version.ts 里找不到 FALLBACK.version 行（格式变了？手工核对）');
  process.exit(1);
}

if (changes.length) {
  console.log(`gen-versions：${changes.join('；')}（源 = package.json ${version}）`);
} else {
  console.log(`gen-versions：三处已是 ${version}，无需改动`);
}
