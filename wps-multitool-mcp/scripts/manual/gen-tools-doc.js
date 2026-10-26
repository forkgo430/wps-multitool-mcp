#!/usr/bin/env node
/**
 * Input: 真实注册路径（dist 注册表 + 注解计算 + 验证清单）
 * Output: 仓库根 TOOLS.md —— 全量工具清单（按组分区，含默认档/注解/验证状态）
 * Pos: 工具清单的单一真相源。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有这个文件
 * ---------------
 * 306 个工具只有运行期才看得见（tools/list 默认只列 40 个），GitHub 页面上等于"什么都没有"。
 * 想搜"WPS 能不能做交叉引用"的人搜不到；搜索引擎也索引不到。
 *
 * 为什么是生成物而不是手写清单
 * --------------------------
 * 手写清单必然腐化 —— 这个仓库的"数字漂移"已经吃过四轮亏（见 .github/workflows/ci.yml 头注释）。
 * 所以清单和 README 的 AUTO 区同一套路：从注册表算，CI 重跑再 git diff --exit-code。
 * 修改工具后忘了重跑 → CI 红，而不是让读者看到一份过期清单。
 *
 * 两条硬约束
 * ----------
 *   1) 不许出现时间戳/版本号/机器名 —— 否则每次生成都 diff 非空，门禁永远红。
 *   2) 落地必须在【仓库根】：.gitignore 里有 `docs/` 与 `/scripts/`，
 *      放进 docs/ 会直接被忽略，CI 的 diff 门禁就成了摆设。
 *
 * 用法：node scripts/manual/gen-tools-doc.js
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..'); // wps-multitool-mcp/
const REPO_ROOT = path.resolve(ROOT, '..');
const OUT = path.join(REPO_ROOT, 'TOOLS.md');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
const { getToolMeta } = require(path.join(ROOT, 'dist', 'server', 'tool-metadata.js'));

// 验证状态是可选依赖：生成器先于它落地时也能跑（整列显示 —）。
// 不写成硬依赖的原因：先生成清单、后补验证表是允许的顺序，不该互相卡住。
let verificationOf = null;
try {
  const mod = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));
  verificationOf = mod.verificationOf;
} catch (e) {
  verificationOf = null;
}

// ---- 1. 真实注册路径（少走一步就少算 12 个内置） -----------------------------
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();
const builtinCount = all.length - allTools.length;
const listedCount = toolRegistry.listTools().tools.length;

const GROUP_ORDER = ['core', 'common', 'word', 'excel', 'ppt'];
const GROUP_LABEL = { core: 'core 组 · 常驻管理', common: '通用', word: 'Word', excel: 'Excel', ppt: 'PPT' };

/** 描述取首行、去掉 markdown 痕迹、压到一行 —— 表格里塞多行会被渲染成断行 */
function shortDesc(desc) {
  const first = String(desc || '').split('\n')[0].replace(/\*\*|`/g, '').trim();
  return first.length > 64 ? `${first.slice(0, 63)}…` : first;
}

/** 注解压成三字母：R 只读 / D 破坏 / N 非幂等 / — 无注解 */
function annotationCell(name) {
  const a = getToolAnnotations(name);
  if (!a) return '—';
  const parts = [];
  if (a.readOnlyHint) parts.push('R');
  if (a.destructiveHint) parts.push('D');
  if (a.idempotentHint === false) parts.push('N');
  return parts.length ? parts.join('') : '—';
}

function verificationCell(name) {
  if (!verificationOf) return '—';
  return verificationOf(name) === 'live' ? 'live' : 'static';
}

// ---- 2. 分组 ---------------------------------------------------------------
const byGroup = new Map();
for (const def of all) {
  const g = groupOfTool(def);
  if (!byGroup.has(g)) byGroup.set(g, []);
  byGroup.get(g).push(def);
}

const groups = [...byGroup.keys()].sort(
  (a, b) => (GROUP_ORDER.indexOf(a) + 1 || 99) - (GROUP_ORDER.indexOf(b) + 1 || 99) || a.localeCompare(b)
);

// ---- 3. 统计（全部实测，不手写） --------------------------------------------
const liveCount = verificationOf ? all.filter((d) => verificationOf(d.name) === 'live').length : 0;
const annotated = all.filter((d) => getToolAnnotations(d.name) !== undefined).length;

const lines = [];
lines.push('# 工具清单');
lines.push('');
lines.push('> 本文件由 `wps-multitool-mcp/scripts/manual/gen-tools-doc.js` 从**运行时注册表**生成，请勿手改。');
lines.push('> 改完工具跑一次 `node scripts/manual/gen-tools-doc.js`；CI 会重跑并 `git diff --exit-code`，对不上就红。');
lines.push('');
lines.push(
  `**${all.length} 个工具**（${builtinCount} 内置 + ${all.length - builtinCount} 专业/组工具）· ` +
    `默认列出 ${listedCount} 个 · 注解覆盖 ${annotated}/${all.length}` +
    (verificationOf ? ` · 真机验证 ${liveCount}/${all.length}` : '')
);
lines.push('');
lines.push('| 标记 | 含义 |');
lines.push('|---|---|');
lines.push('| 默认档 `是` | 不加载组、直接出现在 `tools/list` 里 |');
lines.push('| 注解 `R` / `D` / `N` | 只读 / 破坏性 / 非幂等（`—` = 无法从动作名判定） |');
lines.push('| 验证 `live` | 该工具在真机上跑通过 |');

for (const g of groups) {
  const defs = byGroup.get(g).slice().sort((a, b) => {
    // 默认档的排在前面，其次按名字 —— 读者最想先看到"开箱能用什么"
    const la = toolRegistry.isListed ? Number(!toolRegistry.isListed(a)) : 0;
    const lb = toolRegistry.isListed ? Number(!toolRegistry.isListed(b)) : 0;
    return la - lb || a.name.localeCompare(b.name);
  });
  lines.push('');
  lines.push(`## ${GROUP_LABEL[g] || g}（${defs.length}）`);
  lines.push('');
  lines.push('| 工具 | 默认档 | 注解 | 验证 | 说明 |');
  lines.push('|---|---|---|---|---|');
  for (const def of defs) {
    const meta = getToolMeta(def) || {};
    const aliases = (meta.aliases || []).slice(0, 3).join(' / ');
    const desc = shortDesc(def.description);
    const cell = aliases ? `${desc}<br><sub>别名：${aliases}</sub>` : desc;
    lines.push(
      `| \`${def.name}\` | ${toolRegistry.isListed(def) ? '是' : ''} | ${annotationCell(def.name)} | ` +
        `${verificationCell(def.name)} | ${cell} |`
    );
  }
}

lines.push('');
lines.push('---');
lines.push('');
lines.push('想按中文关键词找工具，别翻这份表：直接用 `wps_search_tools`（支持"插入表格""批注""改背景"这类说法，');
lines.push('返回里带 `notFor` 反例，能分辨近名工具）。这份表是给人看的，也是给搜索引擎看的。');
lines.push('');

const text = lines.join('\n');
const before = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
if (before === text) {
  console.log('TOOLS.md 已是最新');
} else {
  fs.writeFileSync(OUT, text, 'utf8');
  console.log(`已生成 TOOLS.md（${all.length} 个工具 / 默认 ${listedCount} 个）`);
}
process.exit(0);
