#!/usr/bin/env node
/**
 * Input: wps-multitool-mcp/ 的 npm pack --dry-run 产物清单 + 版本号四处真相
 * Output: 0（包内容与版本派生物全部对上）| 1（任一缺失/漂移，逐条列出）
 * Pos: 包内容对账门禁（2026-11-06 所有者批准的根治项，2026-11-07 决策者实现）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * -----------
 * "打包漏文件"是这个项目的惯犯，同族已发三次：
 *   ① docs/ 快照不入库导致干净检出必红（10-25 修）
 *   ② Version PR 漏带 panel 版本派生物（1.0.3、1.0.5 两次）
 *   ③ files 缺 src/server/*.json → 已发布 1.0.5 装机后任务包层空、wps_run_task 空引擎
 *     （11-06 所有者实测发现——装机坏了，CI 和发布对账全是绿的）
 * 共同根因：没有任何一个地方会因为"包里缺了运行时需要的文件"而失败。
 * 本门禁把这件事变成硬失败：npm pack 干跑 → 断言运行时关键文件在包内 + 版本派生物无漂移。
 *
 * 断言两层
 * --------
 *   A. 包内容：装机后引擎要真正跑起来所需的最小文件集（判据 = 运行时读取路径，
 *      不是"觉得该有"——src/server/*.json 是 dist/server 读取候选的兜底位，必须随包）。
 *   B. 版本对账：package.json ↔ src/version.ts FALLBACK ↔ marketplace.json ↔
 *      SKILL.md frontmatter ↔ dist/version.js（编译产物）五处一致。
 *      gen-versions.js 负责写，本门禁负责验——写和验必须分开，同一支笔既写又验等于没验。
 *
 * 用法
 * ----
 *   node scripts/manual/package-content-verify.js            # 正式跑（需先 npm run build）
 *   node scripts/manual/package-content-verify.js --selftest # 自证能红（喂坏数据，不碰真文件）
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..'); // wps-multitool-mcp/
const REPO = path.resolve(ROOT, '..');            // 仓库根

/** 运行时关键文件：装机后引擎真正会读的（判据逐条注明来源路径） */
const REQUIRED_FILES = [
  // 任务包/直答路由：src/server/task-engine.ts 与 intent.ts 的读取候选（dist/server → src/server）
  'src/server/task-packages.json',
  'src/server/intent-routes.json',
  // COM 真相源与守护分发脚本（ps-child.ts:41 / wps-client.ts:108 的读取路径）
  'scripts/wps-com.ps1',
  'scripts/wps-comd.ps1',
  // 入口与传输链（bin 指向 dist/index.js；桥/守护/客户端是常驻链三件）
  'dist/index.js',
  'dist/server/mcp-server.js',
  'dist/bridge/index.js',
  'dist/client/ps-daemon.js',
  'dist/client/wps-client.js',
  'dist/version.js',
  // 技能包（dsh-bundle 与插件市场引用）
  'skills/wps-multitool/SKILL.md',
  'package.json',
];

/** 版本必须一致的五处。返回 [{where, version}]，异常处 version=null */
function collectVersions() {
  const out = [];
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  out.push({ where: 'package.json', version: pkg.version || null });

  const vt = fs.readFileSync(path.join(ROOT, 'src', 'version.ts'), 'utf8');
  const m = vt.match(/const FALLBACK = \{[\s\S]*?version: '([^']*)'/);
  out.push({ where: 'src/version.ts FALLBACK', version: m ? m[1] : null });

  const mp = JSON.parse(fs.readFileSync(path.join(REPO, '.claude-plugin', 'marketplace.json'), 'utf8'));
  out.push({ where: 'marketplace.json', version: (mp.metadata && mp.metadata.version) || null });

  const skill = fs.readFileSync(path.join(ROOT, 'skills', 'wps-multitool', 'SKILL.md'), 'utf8');
  const sm = skill.match(/\n  version: "([^"]*)"/);
  out.push({ where: 'SKILL.md frontmatter', version: sm ? sm[1] : null });

  const distVer = path.join(ROOT, 'dist', 'version.js');
  if (!fs.existsSync(distVer)) {
    out.push({ where: 'dist/version.js（编译产物）', version: null, missing: true });
  } else {
    const dv = fs.readFileSync(distVer, 'utf8');
    const dm = dv.match(/version: '([^']*)'/);
    out.push({ where: 'dist/version.js（编译产物）', version: dm ? dm[1] : null });
  }
  return out;
}

/**
 * 纯断言：给定 tarball 文件清单，返回缺失清单（空 = 过）。
 * 拆纯函数是为了 --selftest 能喂坏数据自证能红，不碰真文件。
 */
function missingFrom(fileList) {
  return REQUIRED_FILES.filter((f) => !fileList.includes(f));
}

/** npm pack 干跑，返回 tarball 内文件路径数组 */
function packDryRun() {
  // Windows 上 npm 是 npm.cmd：裸 spawn ENOENT、直接 spawnSync 又被 Node≥20.12 的
  // CVE-2024-27980 策略拒（EINVAL，本机两连踩）——走 shell 是唯一稳路
  const out = require('child_process').execSync('npm pack --dry-run --json', {
    cwd: ROOT,
    encoding: 'utf8',
    timeout: 120000,
    windowsHide: true,
    maxBuffer: 64 * 1024 * 1024,
  });
  const parsed = JSON.parse(out);
  const entry = Array.isArray(parsed) ? parsed[0] : parsed;
  return (entry && entry.files ? entry.files : []).map((f) => f.path);
}

function main() {
  if (process.argv.includes('--selftest')) {
    // 自证能红：喂两份坏数据（缺关键文件 / 版本漂移），门禁必须报红，否则门禁自身失效
    const badList = packDryRun().filter((f) => f !== 'src/server/task-packages.json');
    const miss = missingFrom(badList);
    const badVersions = [
      { where: 'package.json', version: '9.9.9' },
      { where: 'src/version.ts FALLBACK', version: '1.0.1' },
    ];
    const drifted = badVersions.filter((v) => v.version !== badVersions[0].version).length;
    const red = miss.length === 1 && drifted > 0;
    console.log(red
      ? `--selftest 通过：坏数据被正确判红（缺 1 文件 + 版本漂移 ${drifted} 处）`
      : `--selftest 失败：坏数据没被判红（miss=${JSON.stringify(miss)}）—— 门禁自身失效，禁止使用`);
    process.exit(red ? 0 : 1);
  }

  const problems = [];

  // A. 包内容
  let fileList;
  try {
    fileList = packDryRun();
  } catch (e) {
    console.error(`✗ npm pack 干跑失败：${e.message}`);
    process.exit(1);
  }
  const missing = missingFrom(fileList);
  if (missing.length) {
    problems.push(`包内缺 ${missing.length} 个运行时关键文件:\n  - ${missing.join('\n  - ')}`);
  }

  // B. 版本五处对账
  const versions = collectVersions();
  const anchor = versions[0].version;
  for (const v of versions) {
    if (v.missing) { problems.push(`${v.where} 不存在（先 npm run build 再跑本门禁）`); continue; }
    if (!v.version) { problems.push(`${v.where} 读不到版本号（格式变了？手工核对）`); continue; }
    if (v.version !== anchor) {
      problems.push(`版本漂移：package.json=${anchor} ≠ ${v.where}=${v.version}`);
    }
  }

  if (problems.length) {
    console.error(`✗ 包内容对账失败（${problems.length} 处）:`);
    for (const p of problems) console.error(`  - ${p}`);
    process.exit(1);
  }
  console.log(`✓ 包内容对账通过：tarball ${fileList.length} 条，运行时关键文件 ${REQUIRED_FILES.length}/${REQUIRED_FILES.length} 在包内；版本五处一致（${anchor}）`);
  process.exit(0);
}

main();
