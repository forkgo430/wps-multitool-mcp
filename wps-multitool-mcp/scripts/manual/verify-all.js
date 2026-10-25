/**
 * 一键跑全部验证 —— CI 与本地回归的门禁入口。
 *
 * 为什么要有这个
 * --------------
 * 之前 9 个 verify 脚本全靠人记得逐个跑。今天已经吃过一次亏：
 * 工具数在注册表 / 注解表 / 文档三处漂了四轮才发现，
 * 根因就是"没有任何一个地方会因为数字对不上而失败"。
 *
 * 所以这里的定位是**门禁**：任何一个 verify 挂了，这个脚本就非 0 退出，
 * CI 随之红。人不记得跑没关系，红了会来找你。
 *
 * 用法：
 *   node scripts/manual/verify-all.js          # 快测（不启动子进程，秒级）
 *   node scripts/manual/verify-all.js --all    # 全量（含真跑 PowerShell / 起 server）
 *
 * 分快慢两档的原因：cancel-verify 会 spawn PowerShell 子进程、annotation-e2e 会
 * 真起一个 stdio server 跑协议握手，它们慢（十几秒到几十秒），本地随手跑会烦，
 * 但 CI 上必须跑（那才是防"单测全绿但功能是坏的"的关键）。
 */

const { spawnSync } = require('child_process');
const path = require('path');

const DIR = __dirname;
/**
 * 档位开关。
 *
 * 2026-10-20 修：这里原本硬编码 `const ALL = true;`，于是 `npm run verify` 与
 * `npm run verify:all` 行为完全一样 —— 文件头与 package.json 都承诺"无参是快测档"，
 * 实际不存在这个档位。改回读 argv 后：
 *   npm run verify     → 快测（FAST 15 项，不起子进程）
 *   npm run verify:all → 全量（FAST + HEAVY 7 项，含真 WPS）
 * CI 走的是 verify:all，行为不变。
 */
const ALL = process.argv.includes('--all');

/** 不需要 spawn 子进程的快测：本地随手跑、CI 每提交都跑 */
const FAST = [
  ['error-map-test.js', '错误分类映射'],
  ['error-type-verify.js', '错误类型与可重试标签'],
  ['search-meta-verify.js', '工具检索元数据（别名/反例）'],
  ['verification-verify.js', '验证状态覆盖与默认档'],
  ['basic-tier-verify.js', '默认档 = 白名单 ∩ 真机验证'],
  ['alias-verify.js', '工具别名（旧名可解析/不列出/不计入）'],
  ['param-shape-verify.js', '参数形状（TS 传了 PS 不读）'],
  ['auto-snapshot-verify.js', '单工具写入自动快照（mock COM）'],
  ['word-target-verify.js', 'word target 参数链路（mock COM）'],
  ['tool-groups-test.js', '工具组加载'],
  ['intent-test.js', '意图识别词表'],
  ['annotation-verify.js', '工具注解与三方对账'],
  ['warnings-batch-verify.js', 'warnings 语义 + 批量执行'],
  ['install-verify.js', 'install 子命令（合并/备份/幂等）'],
  ['skill-verify.js', 'skill 包规范'],
  ['dsh-config-verify.js', 'dsh 适配约束（命名/超时）'],
  ['windows-path-frozen.js', 'Windows 链路冻结门禁（C1）'],
  ['backend-selector-test.js', '平台识别决策链（Linux B0）'],
  ['addon-channel-test.js', '加载项通道骨架（Linux B2：握手三态/端口绑定/路由/失败语义/轮询闭环）'],
];

/** 要起子进程 / 真跑 PowerShell 的重测。第三列 = 仅 Windows 能跑（Linux 上跳过不判失败，B8） */
const HEAVY = [
  ['excel-caps-probe.js', 'Excel 能力探测（需 WPS）', true],
  ['word-caps-probe.js', 'Word 能力探测（需 WPS）', true],
  ['snapshot-verify.js', '快照 / 恢复（需 WPS）', true],
  ['ppt-caps-probe.js', 'PPT 基础档（需 WPS 演示）', true],
  ['annotation-e2e.js', '注解端到端（真起 stdio server）'],
  ['bridge-test.js', '桥进程', true],
  ['cancel-verify.js', '取消机制（真跑 PowerShell）', true],
];

const list = [...FAST, ...(ALL ? HEAVY : [])];

console.log(`\n═══ 验证门禁（共 ${list.length} 项${ALL ? '，含重测' : '，快测档'}）═══\n`);

const results = [];
const skipped = [];
for (const [file, label, windowsOnly] of list) {
  // B8：Windows-only 门禁在非 Windows 平台显式跳过（不判失败），Linux 上 npm test 才能跑完
  if (windowsOnly && process.platform !== 'win32') {
    skipped.push(file);
    console.log(` 跳过  ${label} —— 仅 Windows（Linux 门禁）`);
    continue;
  }
  const started = Date.now();
  process.stdout.write(`  ${label.padEnd(28, ' ')} …`);
  const r = spawnSync(process.execPath, [path.join(DIR, file)], {
    encoding: 'utf8',
    timeout: 300000,
  });
  const ms = Date.now() - started;
  const okFlag = r.status === 0;
  results.push({ file, label, ok: okFlag, ms });

  if (okFlag) {
    const tail = (r.stdout || '').trim().split('\n').pop() || '';
    console.log(` 通过  ${String(ms).padStart(5)}ms  ${tail.slice(0, 60)}`);
  } else {
    console.log(` 失败  ${String(ms).padStart(5)}ms`);
    // 失败时把细节打全：门禁的价值就在这行
    const out = `${r.stdout || ''}\n${r.stderr || ''}`;
    out.trim()
      .split('\n')
      .filter((l) => l.trim())
      .slice(-25)
      .forEach((l) => console.log(`        ${l}`));
  }
}

const bad = results.filter((r) => !r.ok);
const skipNote = skipped.length ? `，跳过 ${skipped.length}（仅 Windows）` : '';
console.log(`\n═══ 结果：${results.length - bad.length}/${results.length} 通过${skipNote} ═══`);
if (bad.length) {
  console.log(`失败项：${bad.map((b) => b.file).join(', ')}`);
  console.log('\n门禁不通过 —— 别绕过它，先看失败原因。');
  process.exit(1);
}
process.exit(0);
