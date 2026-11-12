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
  ['collab-guard-verify.js', '协作安全网：人工编辑检测与修订建议（mock COM）'],
  ['batch-shapes-verify.js', '⑤ 批量形状原语（主线一 A 段）：upsert 两路/逐项失败聚合/上限 50（mock COM）'],
  ['wordkit-verify.js', 'M2 wordkit 文档引擎（主线一 B 段）：8 kind 展开/生成后校验三路/降级两路（mock COM）'],
  ['readback-slim-verify.js', '回包瘦身（体验修复轮 B）：read_range maxRows/brief 摘要/inline 开关（mock COM）'],
  ['office-host-verify.js', 'Office host 维度（一期+二期+收口轮）：探测纯函数/ps1 结构 gate/对账表一致性/绑定一致性/find51/孤儿桥监测（全 mock）'],
  ['orphan-bridge-verify.js', '孤儿桥生命周期：EPHEMERAL 父亡监测 10 连跑归零（不起 WPS COM，ping 全链）'],
  ['response-gate-verify.js', '提质小卡 P0：structuredContent 体积闸门（≤2048 镜像/text 逐字节不变/大回包只留 text）'],
  ['toolslist-shape-verify.js', '提质小卡 P1：tools/list 无 examples + searchTools 仍有 examples + 注解回归'],
  ['shape-cleanup-verify.js', '提质小卡 P3：handleToolCall 孤儿 {} 块清理'],
  ['lint-run-verify.js', '提质小卡：lint 真跑 0 错 + eslint 精确锁版本 + 最小 config'],
  ['quality-anchors-verify.js', '提质合并卡：R1/R2/R4/R5/R6 锚断言（绑定/归一/降级/raw，静态+坏HOME运行时）'],
  ['schema-weight-verify.js', '提质合并卡 B 段：默认档 payload ≤36000B 门禁（防回弹）'],
  ['slow-actions-verify.js', '提质合并卡 R4：超时分级表断言（五动作慢档/快动作不放宽）'],
  ['janitor-verify.js', '看门狗 janitor：journal 落盘/四条件决策/UI 零误伤/PID 复用免疫/销账/滚动截断（pwsh7 注入，零 WPS COM）'],
  ['render-template-verify.js', '模板层大卡 A：renderTemplate 结构/双轨/缺 key/落盘校验/慢档'],
  ['extract-template-verify.js', '模板层大卡 B：extractTemplate 复制保样式/标题保序/raw 上限/慢档'],
  ['template-packages-verify.js', '模板层大卡 C：配方追加 19 包/既有 17 包键集合对账/SKILL 两轨'],
  ['starter-assets-verify.js', '模板层大卡 D：starter 三件在仓+files 字段+OOXML 锚点预置解包验证'],
  ['template-gate-verify.js', '模板层 gate：R 类教训防复发（归一/慢档/空文本/text 主载体/schema 短描述）+ 双宿主台账标记'],
  ['dual-host-verify.js', '模板层大卡：双宿主真机渲染/提取产物在案 + 台账双宿主证据（live PASS 固化锚）'],
  ['render-ppt-xlsx-verify.js', '模板三格式补全卡：ppt 组合递归/Replace 两参/xlsx Areas 枚举/缺 key 列全/导图像素'],
  ['word-target-verify.js', 'word target 参数链路（mock COM）'],
  ['tool-groups-test.js', '工具组加载'],
  ['intent-test.js', '意图识别词表'],
  ['annotation-verify.js', '工具注解与三方对账'],
  ['warnings-batch-verify.js', 'warnings 语义 + 批量执行'],
  ['install-verify.js', 'install 子命令（合并/备份/幂等）'],
  ['skill-verify.js', 'skill 包规范'],
  ['dsh-config-verify.js', 'dsh 适配约束（命名/超时）'],
  ['windows-path-frozen.js', 'Windows 链路冻结门禁（C1）'],
  ['backend-selector-test.js', '平台识别决策链（Linux B0）+ Office host 维度（WPS_OFFICE 矩阵/二进制解析）'],
  ['addon-channel-test.js', '加载项通道骨架（Linux B2：握手三态/端口绑定/路由/失败语义/轮询闭环）'],
  ['actions-dispatcher-test.js', '加载项动作分发器 20 动作 mock 单测（Linux B2.5）'],
  ['actions-sync-verify.js', 'Linux 落仓轮：actions-sync 字节指纹钉（云端 B5-T4 验证版）/ 规格计数 13+15+95 / model-detect 双分支 / 三型接线（全 mock）'],
  ['deckkit-test.js', 'deckkit 引擎（E2 骨架 + E3 主题/溢出检查 + 质量修复轮：0 坐标/参数名对齐，10 页型展开 + 排版纪律 + mock 冒烟）'],
  ['task-engine-test.js', '任务引擎（意图一期：服务端执行/参数模板跨步引用/confirm 闸门/fail-fast，mock 注入）'],
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
