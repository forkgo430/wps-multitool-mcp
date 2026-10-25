#!/usr/bin/env node
/**
 * Input: 真实注册路径（dist 里的注册表 + 注解计算）
 * Output: README.md 的两处 AUTO 区 + wps-multitool-addon/panel.html 的数字/版本
 * Pos: 文档统计的单一真相源。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 要解决的问题
 * ------------
 * "工具总数/分布/注解统计/面板版本"这些数字曾经在 README 三处、panel.html、
 * SKILL.md 各抄一遍，漂了两轮没人发现 —— 根因是没有任何一处跟注册表挂钩，
 * 也没有任何一个地方会因为"数字对不上"而失败。
 *
 * 这个脚本让 README 与面板的数字【只能来自注册表】：
 *   - README `<!-- AUTO:tool-stats:start/end -->`：总数、分布、默认列出数
 *   - README `<!-- AUTO:tool-table:start/end -->`：工具清单表（计数列）
 *   - panel.html：工具数那行 + 版本号（版本以 package.json 为准）
 * SKILL.md 不归这个脚本管 —— 按约定它【不含数字】（"近 300 个，以 tools/list 为准"），
 * 背数字必腐化且没有门禁会红。
 *
 * 覆盖面必须走【真实的注册路径】（与 gen-tool-annotations.js 同一套教训）：
 *   registerBuiltinTools() 的 12 个内置 + registerAll(allTools) 的专业工具。
 * 少做第一步，总数就少 12 个 —— 这种"两处维护"的漂移已经吃过一次亏。
 *
 * 用法：node scripts/manual/gen-doc-stats.js
 * CI：跑完本脚本后 `git diff --exit-code` README.md 与 panel.html ——
 * 改了工具忘了重跑 → diff 非空 → 红。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const REPO_ROOT = path.resolve(ROOT, '..');
const README = path.join(REPO_ROOT, 'README.md');
const README_EN = path.join(REPO_ROOT, 'README.en.md');
const PANEL = path.join(REPO_ROOT, 'wps-multitool-addon', 'panel.html');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { annotationSummary } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
const { verificationSummary } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));

// ---- 1. 真实注册路径，量出数字 ----------------------------------------------
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();
const total = all.length;
const builtinCount = total - allTools.length; // registerBuiltinTools 注册的那批
const professionalCount = allTools.length;

// 按组统计（内置工具计入其 category 对应的组 —— 与 list_tool_groups 的口径一致）
const byGroup = {};
for (const t of all) {
  const g = groupOfTool(t);
  byGroup[g] = (byGroup[g] || 0) + 1;
}
const ann = annotationSummary(all.map((t) => t.name));
// 别叫 ver —— 文件末尾 panel 版本号那个变量也叫 ver，重名会在运行期直接炸（已经踩过一次）
const vstat = verificationSummary(all.map((t) => t.name));
const listed = toolRegistry.listTools().tools.length;
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));

// ---- 2. 生成两块 AUTO 内容 ---------------------------------------------------

/** 各组在清单表里的展示名与能力描述（能力描述是给人看的，属于数据不是逻辑） */
const TABLE_ROWS = [
  {
    group: 'excel',
    label: 'Excel',
    labelEn: 'Excel',
    caps: '公式/数据/图表/透视表/工作表/格式/工作簿/行列/批注保护/图片导出',
    capsEn:
      'formulas, data, charts, pivot tables, sheets, formatting, workbooks, rows/columns, comments & protection, image export',
  },
  {
    group: 'word',
    label: 'Word',
    labelEn: 'Word',
    caps: '格式/内容/文档管理/页眉页脚/批注/书签/表格/模板填写/学术排版/修订/索引/交叉引用',
    capsEn:
      'formatting, content, document management, headers/footers, comments, bookmarks, tables, templates, academic layout, revisions, index, cross-references',
  },
  {
    group: 'ppt',
    label: 'PPT',
    labelEn: 'PPT',
    caps: '幻灯片/形状/图片/表格/美化/动画/图表/3D/数据可视化/图片导出',
    capsEn: 'slides, shapes, images, tables, beautify, animations, charts, 3D, data viz, image export',
  },
  {
    // common + core：core 是工具组管理本身，与通用工具合一行展示
    group: 'common',
    label: '通用',
    labelEn: 'Common',
    caps: '保存/连接检测/文本选取/格式转换/工具组管理/意图路由/批量执行',
    capsEn:
      'save, connection check, text selection, format conversion, tool groups, intent routing, batch',
    includeCore: true,
  },
];

const commonTotal = (byGroup.common || 0) + (byGroup.core || 0);

/** 按语言生成三块 AUTO 内容（中文 README 与英文 README 共用同一份实测数字） */
function makeBlocks(lang) {
  const en = lang === 'en';
  const tableLines = en
    ? ['| App | Tools | Highlights |', '|-----|-------|-----------|']
    : ['| 应用 | 工具数 | 主要能力 |', '|------|--------|---------|'];
  // 口径与统计区【完全一致】：内置计入对应组、core 计入通用。
  // 曾经表格用"专业互斥 + 单列内置"口径，数字虽对但和统计区长得不像，
  // 读者一律当成"旧的没同步"—— 口径不统一比数字错更费沟通成本。
  for (const row of TABLE_ROWS) {
    let n = byGroup[row.group] || 0;
    if (row.includeCore) n += byGroup.core || 0;
    tableLines.push(
      en
        ? `| ${row.labelEn} | ${n} | ${row.capsEn} |`
        : `| ${row.label} | ${n} | ${row.caps} |`
    );
  }
  tableLines.push(
    en
      ? `| Built-in | ${builtinCount} (already counted above) | connection check, generic method, data cache |`
      : `| 内置 | ${builtinCount}（已计入上方各组） | 连接检查/万能方法调用/数据缓存 |`
  );

  const statsBlock = en
    ? [
        `- **${total} tools** (${builtinCount} built-in + ${professionalCount} specialized); ` +
          `**only ${listed} listed by default** (essential ∩ live-verified, load on demand — see \`scripts/manual/basic-tier-verify.js\`)`,
        `- Distribution (by tool group, built-ins counted in their app group): ` +
          `Common ${commonTotal} (incl. core) / Word ${byGroup.word || 0} / Excel ${byGroup.excel || 0} / PPT ${byGroup.ppt || 0}`,
        `- Annotations (measured): read-only ${ann.readOnly} / destructive ${ann.destructive} / ` +
          `non-idempotent ${ann.nonIdempotent} — coverage ${ann.annotated}/${ann.total}`,
      ].join('\n')
    : [
        `- **${total} 个工具**（${builtinCount} 个内置 + ${professionalCount} 个专业/组工具）；` +
          `**默认只列 ${listed} 个**（基础档 ∩ 真机验证，按需加载 —— 口径与校验见 \`scripts/manual/basic-tier-verify.js\`）`,
        `- 分布（按工具组，内置计入对应组）：` +
          `通用 ${commonTotal}（含 core）/ Word ${byGroup.word || 0} / Excel ${byGroup.excel || 0} / PPT ${byGroup.ppt || 0}`,
        `- 注解（实测）：只读 ${ann.readOnly} / 破坏性 ${ann.destructive} / 非幂等 ${ann.nonIdempotent}，` +
          `覆盖 ${ann.annotated}/${ann.total}`,
      ].join('\n');

  // 徽章行同样不许手抄：版本取自 package.json，工具数取自注册表。
  // CI 徽章在 run #5 真跑绿之后才挂（没绿就挂只会显示 no status / failing，比不挂难看）。
  const badgesBlock = [
    '[![CI](https://github.com/forkgo430/wps-multitool-mcp/actions/workflows/ci.yml/badge.svg)]' +
      '(https://github.com/forkgo430/wps-multitool-mcp/actions/workflows/ci.yml)',
    '[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)',
    // npm 版本徽章走 shields 动态读 registry：比手写版本号诚实（本地改版本但没发 npm 时不会骗人）
    '[![npm](https://img.shields.io/npm/v/wps-multitool-mcp)](https://www.npmjs.com/package/wps-multitool-mcp)',
    // release 徽章同样动态 —— 发新版本自动变，不会漂
    '[![release](https://img.shields.io/github/v/release/forkgo430/wps-multitool-mcp)](https://github.com/forkgo430/wps-multitool-mcp/releases)',
    `[![tools](https://img.shields.io/badge/tools-${total}-brightgreen.svg)](${en ? '#capabilities' : '#能力一览'})`,
    '[![platform](https://img.shields.io/badge/platform-Windows%20only-lightgrey.svg)](' +
      (en ? '#requirements' : '#安装') +
      ')',
  ].join(' ');

  return { statsBlock, tableBlock: tableLines.join('\n'), badgesBlock };
}

// ---- 3. 替换 README 的 AUTO 区 -------------------------------------------------

/** 把 start/end 标记之间的内容换掉。标记缺失 = 调用方没按约定接好，直接失败。 */
function replaceAutoRegion(text, startMarker, endMarker, block, label) {
  const startIdx = text.indexOf(startMarker);
  const endIdx = text.indexOf(endMarker);
  if (startIdx < 0 || endIdx < 0 || endIdx < startIdx) {
    console.error(`✗ README 里找不到 ${label} 的 AUTO 标记（${startMarker} … ${endMarker}）`);
    process.exit(1);
  }
  return text.slice(0, startIdx + startMarker.length) + '\n' + block + '\n' + text.slice(endIdx);
}

/** 中英两份 README 用同一套数据生成，各自有独立的 AUTO 标记 */
const DOCS = [
  { file: README, lang: 'zh', name: 'README.md' },
  { file: README_EN, lang: 'en', name: 'README.en.md' },
];

for (const doc of DOCS) {
  if (!fs.existsSync(doc.file)) {
    console.error(`✗ 找不到 ${doc.name}`);
    process.exit(1);
  }
  const blocks = makeBlocks(doc.lang);
  let text = fs.readFileSync(doc.file, 'utf8');
  const before = text;
  text = replaceAutoRegion(text, '<!-- AUTO:badges:start -->', '<!-- AUTO:badges:end -->', blocks.badgesBlock, 'badges');
  text = replaceAutoRegion(text, '<!-- AUTO:tool-stats:start -->', '<!-- AUTO:tool-stats:end -->', blocks.statsBlock, 'tool-stats');
  text = replaceAutoRegion(text, '<!-- AUTO:tool-table:start -->', '<!-- AUTO:tool-table:end -->', blocks.tableBlock, 'tool-table');
  if (text !== before) {
    fs.writeFileSync(doc.file, text, 'utf8');
    console.log(`已更新 ${doc.name}（AUTO 区三块）`);
  } else {
    console.log(`${doc.name} 已是最新（AUTO 区无需改动）`);
  }
}

// ---- 4. panel.html：工具数那行 + 版本号 ----------------------------------------
// 版本以 package.json 为唯一真相源（与 src/version.ts 同源），面板不许自己背一个。
let panel = fs.readFileSync(PANEL, 'utf8');
const panelBefore = panel;
const ver = pkg.version;
panel = panel.replace(
  /<li>\d+ 个 MCP 工具，默认只列 \d+ 个基础工具<\/li>/,
  `<li>${total} 个 MCP 工具，默认只列 ${listed} 个基础工具</li>`
);
panel = panel.replace(/(<span class="ver">v)\d+\.\d+\.\d+(<\/span>)/, `$1${ver}$2`);
panel = panel.replace(
  /(<span class="k">版本<\/span><span>)\d+\.\d+\.\d+(<\/span>)/,
  `$1${ver}$2`
);
if (panel !== panelBefore) {
  fs.writeFileSync(PANEL, panel, 'utf8');
  console.log(`已更新 panel.html（工具数 / 版本 → ${ver}）`);
} else {
  console.log('panel.html 已是最新');
}

// ---- 5. 报告 -------------------------------------------------------------------
console.log(`\n注册表实测（真实注册路径）:`);
console.log(`  总数 ${total} = 内置 ${builtinCount} + 专业 ${professionalCount}`);
console.log(
  `  按组: common ${byGroup.common || 0} + core ${byGroup.core || 0} / word ${byGroup.word || 0} / ` +
    `excel ${byGroup.excel || 0} / ppt ${byGroup.ppt || 0}`
);
console.log(`  默认列出 ${listed} 个（基础档）`);
console.log(
  `  注解: 只读 ${ann.readOnly} / 破坏性 ${ann.destructive} / 非幂等 ${ann.nonIdempotent} / 覆盖 ${ann.annotated}/${ann.total}`
);
process.exit(0);
