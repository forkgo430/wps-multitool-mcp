#!/usr/bin/env node
/**
 * Input: scripts/manual/tool-verification.json + 真实注册路径（dist 注册表）
 * Output: src/server/tool-verification.generated.ts（每个注册工具 → live / static）
 * Pos: 验证状态表的生成器（与 gen-tool-annotations.js 同套路）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有这张表
 * --------------
 * 306 个工具里有整整一组（PPT 116 个）没有真机探测。把它和"真机跑通过"的工具混在一个列表里，
 * 用户没法判断结果能不能信 —— 而 PPT 偏偏是最容易"没报错但做错了"的一类。
 *
 * 规则
 * ----
 *   组级默认（groups）+ 逐工具覆盖（tools）。命名一律用【工具名】，不用动作名。
 *   注册表里有、表里没有 → 直接非 0 退出（新加组忘了表态就红，不允许"默认当验证过"）。
 *   表里有、注册表没有 → 报警告（多半是工具被删了，配置留了死条目）。
 *
 * 为什么生成成 TS 常量而不是运行期读 JSON
 * ------------------------------------
 * 包发布只带 dist/ 与 scripts/wps-com.ps1，scripts/manual/ 不进包。
 * 运行期读 JSON 会在用户机器上找不到文件。
 *
 * 用法：node scripts/manual/gen-tool-verification.js
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const MANIFEST = path.join(__dirname, 'tool-verification.json');
const OUT = path.join(ROOT, 'src', 'server', 'tool-verification.generated.ts');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));

const VALID = new Set(['live', 'static']);

function loadManifest() {
  const raw = JSON.parse(fs.readFileSync(MANIFEST, 'utf8'));
  const groups = raw.groups || {};
  const tools = raw.tools || {};
  for (const [k, v] of Object.entries(groups)) {
    if (!VALID.has(v)) {
      console.error(`✗ groups.${k} 的值 "${v}" 不合法（只能是 live / static）`);
      process.exit(1);
    }
  }
  for (const [k, v] of Object.entries(tools)) {
    if (!VALID.has(v)) {
      console.error(`✗ tools.${k} 的值 "${v}" 不合法（只能是 live / static）`);
      process.exit(1);
    }
  }
  return { groups, tools, evidence: raw.evidence || {} };
}

const manifest = loadManifest();

// 真实注册路径：漏了 registerBuiltinTools 就会少 12 个（这个坑已经吃过一次）
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();

const unknownGroup = [];
const table = {};
for (const def of all) {
  const group = groupOfTool(def);
  const override = manifest.tools[def.name];
  const status = override || manifest.groups[group];
  if (!status) {
    unknownGroup.push(`${def.name}（组 ${group} 没在 manifest.groups 里表态）`);
    continue;
  }
  table[def.name] = status;
}

if (unknownGroup.length) {
  console.error('✗ 下列工具没有验证状态，先在 tool-verification.json 的 groups 里给它的组表态：');
  for (const u of unknownGroup) console.error(`   - ${u}`);
  process.exit(1);
}

// 死条目：manifest 里写了但注册表没有 —— 不致命，但要提示，避免配置越积越假
const registered = new Set(all.map((d) => d.name));
const dead = Object.keys(manifest.tools).filter((n) => !registered.has(n));
if (dead.length) {
  console.warn(`⚠ tool-verification.json 里这些工具已不存在（死条目）：${dead.join(', ')}`);
}

// 稳定输出：键排序，无时间戳（否则每次生成都 diff 非空）
const names = Object.keys(table).sort();
const body = names.map((n) => `  ${JSON.stringify(n)}: ${JSON.stringify(table[n])},`).join('\n');
const liveCount = names.filter((n) => table[n] === 'live').length;

const content = `/* eslint-disable */
/**
 * 本文件由 scripts/manual/gen-tool-verification.js 生成，请勿手改。
 * 真相源：scripts/manual/tool-verification.json（人维护：组级默认 + 逐工具覆盖 + 证据）
 *
 * 语义：live = 有真机证据；static = 只有静态校验。拿不准一律 static ——
 * 把没验过的说成验过，是在拿用户的文档赌。
 */
export type ToolVerification = 'live' | 'static';

/** 每个注册工具的验证状态（覆盖 ${names.length}/${all.length}） */
export const TOOL_VERIFICATION: Readonly<Record<string, ToolVerification>> = {
${body}
};

/** 生成时的总量，供自检/文档使用（别手写这些数字） */
export const VERIFICATION_TOTALS = { live: ${liveCount}, static: ${names.length - liveCount}, total: ${names.length} } as const;
`;

const before = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
if (before === content) {
  console.log(`tool-verification.generated.ts 已是最新（live ${liveCount} / static ${names.length - liveCount} / 共 ${names.length}）`);
} else {
  fs.writeFileSync(OUT, content, 'utf8');
  console.log(`已生成 tool-verification.generated.ts（live ${liveCount} / static ${names.length - liveCount} / 共 ${names.length}）`);
}
process.exit(0);
