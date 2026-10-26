#!/usr/bin/env node
/**
 * Input: scripts/manual/tool-aliases.json（人维护的真相源：旧名 → 规范名）
 * Output: src/server/tool-aliases.generated.ts（运行时用的别名表 + 解析函数）
 * Pos: 别名表生成器。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么生成成 TS 常量，而不是运行期读 JSON
 * ----------------------------------------
 * package.json 的 files 只发 dist + scripts/wps-com.ps1 + skills，scripts/manual/*.json 不进包。
 * 运行期读它会在用户机器上 ENOENT —— 而这个表恰好是每次 callTool 都要查的，出错即全站不可用。
 *
 * 生成器只做【形状校验】，不查目标是否存在
 * --------------------------------------
 * "别名目标必须在注册表里"这条要在 alias-verify.js 里查，因为这里拿不到可信的注册表：
 * 生成器的输入就是 registry 要 import 的表，先 require(dist) 会把旧的编译产物当真相源（鸡生蛋）。
 * 所以：形状问题在这里非 0 退出；语义问题（目标不存在、链式别名）交给门禁。
 *
 * 幂等：输出按 key 排序，输入行序变动不会造成 diff（CI 有 git diff 门禁）。
 *
 * 用法：node scripts/manual/gen-tool-aliases.js
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const SRC = path.join(__dirname, 'tool-aliases.json');
const OUT = path.join(ROOT, 'src', 'server', 'tool-aliases.generated.ts');

const NAME_RE = /^wps_[a-z0-9]+(_[a-z0-9]+)*$/;

function die(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}

if (!fs.existsSync(SRC)) die(`找不到真相源 ${path.relative(ROOT, SRC)}`);

const raw = fs.readFileSync(SRC, 'utf8').replace(/^\uFEFF/, '');

// JSON.parse 会把重复 key 静默并成最后一个 —— 而重复 key 恰恰是最容易犯的编辑错误。
// 这里先扫一遍原始文本，把同一层里出现两次的 key 抓出来。
function findDuplicateKeys(text) {
  // 只扫 "aliases" 那一层的直接子键：缩进两格的 "key":
  const inAliases = text.indexOf('"aliases"');
  if (inAliases < 0) return [];
  const body = text.slice(inAliases);
  const seen = new Map();
  const dups = [];
  for (const m of body.matchAll(/^ {4}"([^"]+)"\s*:/gm)) {
    const k = m[1];
    if (seen.has(k)) dups.push(k);
    else seen.set(k, true);
  }
  return dups;
}
const dups = findDuplicateKeys(raw);
if (dups.length) die(`tool-aliases.json 里有重复的键：${dups.join(', ')}`);

let parsed;
try {
  parsed = JSON.parse(raw);
} catch (e) {
  die(`tool-aliases.json 不是合法 JSON：${e.message}`);
}

const aliases = parsed.aliases || {};
const rawEntries = Object.entries(aliases);
if (!rawEntries.length) die('aliases 是空的 —— 没有别名就别跑生成器（会把表清空）');

// 第 8 轮：value 可以是字符串（纯改名），也可以是 { to, rename, set }（带参数映射）。
// 统一在这里摊平成 [from, to, map]，后面所有校验与生成只看摊平后的形态。
const entries = [];
for (const [from, spec] of rawEntries) {
  if (typeof spec === 'string') {
    entries.push({ from, to: spec, rename: null, set: null });
  } else if (spec && typeof spec === 'object' && typeof spec.to === 'string') {
    entries.push({
      from,
      to: spec.to,
      rename: spec.rename && typeof spec.rename === 'object' ? spec.rename : null,
      set: spec.set && typeof spec.set === 'object' ? spec.set : null,
    });
  } else {
    entries.push({ from, to: spec, rename: null, set: null, bad: true });
  }
}

const problems = [];
const targets = new Map();
for (const { from, to, rename, bad } of entries) {
  if (bad || typeof to !== 'string') {
    problems.push(`${from}: 目标必须是字符串或 { to, rename?, set? }，收到 ${JSON.stringify(to)}`);
    continue;
  }
  if (!NAME_RE.test(from)) problems.push(`${from}: 别名不合 wps_<组>_<动词>_<名词> 形状`);
  if (!NAME_RE.test(to)) problems.push(`${to}: 规范名不合 wps_<组>_<动词>_<名词> 形状（来自 ${from}）`);
  if (from === to) problems.push(`${from}: 自己别名自己`);
  // rename 的两端都得是合法的标识符名，否则生成的 TS 会在对象字面量里拼出非法键
  for (const [oldKey, newKey] of Object.entries(rename || {})) {
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(oldKey) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(newKey)) {
      problems.push(`${from}: rename 的键名非法 ${oldKey} → ${newKey}`);
    }
    if (oldKey === newKey) problems.push(`${from}: rename 把 ${oldKey} 映射给了自己（等于没写）`);
  }
  // 第 5 轮：允许多个旧名指向同一个规范名 —— 族合并就是"多个旧名 → 一个规范名"。
  // 以前这里禁止是第 2 轮的口径（那时一条规范名只吸收一个重复），现在会挡住 background 那种 4→1。
  // 真正要防的是"别名指向不存在的名字"，那条由 alias-verify.js 对着注册表查。
  targets.set(to, (targets.get(to) || 0) + 1);
}
// 链式别名：a→b 且 b→c。解析只做一跳，链式会让 a 解析到不存在的 b。
for (const { from, to } of entries) {
  if (typeof to === 'string' && Object.prototype.hasOwnProperty.call(aliases, to)) {
    problems.push(`${from} → ${to}，但 ${to} 本身也是别名（解析只做一跳，链式会指向不存在的名字）`);
  }
}
if (problems.length) {
  console.error('✗ 别名表有问题：');
  for (const p of problems) console.error(`  - ${p}`);
  process.exit(1);
}

const sorted = entries.slice().sort((a, b) => a.from.localeCompare(b.from));

const lines = [];
lines.push('/**');
lines.push(' * Input: 无（纯常量）');
lines.push(' * Output: 旧工具名 → 规范工具名的映射与解析函数');
lines.push(' * Pos: 别名表【生成物】。一旦我被修改，请更新我的头部注释，以及所属目录的md。');
lines.push(' *');
lines.push(' * 别手改这个文件：真相源是 scripts/manual/tool-aliases.json，');
lines.push(` * 生成器是 scripts/manual/gen-tool-aliases.js，共 ${sorted.length} 条。`);
lines.push(' *');
lines.push(' * 别名不注册成工具：它不出现在 tools/list、不计入工具总数，只在 registry 的入口被解析成规范名。');
lines.push(' * 为什么只做一跳解析：链式别名（a→b→c）会让"到底谁是真的"变成谜，一份表里就该一步到位。');
lines.push(' */');
lines.push('');
lines.push('/** 旧名 → 规范名。键是已下线的旧工具名，值是仍在注册表里的规范名 */');
lines.push('export const TOOL_ALIASES: Readonly<Record<string, string>> = {');
for (const { from, to } of sorted) {
  lines.push(`  ${JSON.stringify(from)}: ${JSON.stringify(to)},`);
}
lines.push('};');
lines.push('');
const mapped = sorted.filter((e) => e.rename || e.set);
lines.push('/**');
lines.push(' * 旧名 → 该旧名的参数映射。只收带映射的那些（纯改名的旧名不需要任何处理）。');
lines.push(' *');
lines.push(' * 族合并的失效方式有一半在参数上：光把名字指向合并工具，旧调用会 Missing required parameter，');
lines.push(' * 或者更糟 —— newOrder 这类改名被静默忽略，工具按默认值把动画排到一个错的位置。');
lines.push(' * rename = 旧参数名 → 新参数名（调用方没给新键时才搬，不覆盖规范名自己收的参数）；');
lines.push(' * set = 【缺了才补】的固定值（调用方显式给了就听调用方的）。');
lines.push(' */');
lines.push('export interface ToolAliasParamMap {');
lines.push('  rename?: Readonly<Record<string, string>>;');
lines.push('  set?: Readonly<Record<string, unknown>>;');
lines.push('}');
lines.push('');
lines.push(`export const TOOL_ALIAS_PARAMS: Readonly<Record<string, ToolAliasParamMap>> = {`);
for (const { from, rename, set } of mapped) {
  const parts = [];
  if (rename) parts.push(`rename: ${JSON.stringify(rename)}`);
  if (set) parts.push(`set: ${JSON.stringify(set)}`);
  lines.push(`  ${JSON.stringify(from)}: { ${parts.join(', ')} },`);
}
lines.push('};');
lines.push('');
lines.push('/** 全部别名（旧名）。用于"这个请求是不是旧名"的判断与门禁遍历 */');
lines.push('export const ALIAS_NAMES: readonly string[] = Object.keys(TOOL_ALIASES);');
lines.push('');
lines.push('/**');
lines.push(' * 把（可能是旧名的）工具名解析成规范名。');
lines.push(' * 不是别名就原样返回 —— 调用方不用先判断再解析，少一个分支就少一处漏判。');
lines.push(' */');
lines.push('export function resolveToolAlias(name: string): string {');
lines.push('  return TOOL_ALIASES[name] ?? name;');
lines.push('}');
lines.push('');
lines.push('/** 这个名字是不是已下线的旧名（用于在结果里告诉调用方"你用的是旧名"） */');
lines.push('export function isToolAlias(name: string): boolean {');
lines.push('  return Object.prototype.hasOwnProperty.call(TOOL_ALIASES, name);');
lines.push('}');
lines.push('');
lines.push('/**');
lines.push(' * 把旧名的参数搬成规范名的参数。非别名、或没有参数映射时原样返回同一个对象');
lines.push(' * （不复制：绝大多数调用走的是这条路，复制纯属白花 CPU）。');
lines.push(' */');
lines.push('export function applyAliasParams(');
lines.push('  name: string,');
lines.push('  args: Record<string, unknown>');
lines.push('): Record<string, unknown> {');
lines.push('  const map = TOOL_ALIAS_PARAMS[name];');
lines.push('  if (!map) return args;');
lines.push('');
lines.push('  const out: Record<string, unknown> = { ...(args || {}) };');
lines.push('  for (const [oldKey, newKey] of Object.entries(map.rename || {})) {');
lines.push('    if (oldKey in out && !(newKey in out)) {');
lines.push('      out[newKey] = out[oldKey];');
lines.push('      delete out[oldKey];');
lines.push('    }');
lines.push('  }');
lines.push('  for (const [key, value] of Object.entries(map.set || {})) {');
lines.push('    if (out[key] === undefined) out[key] = value;');
lines.push('  }');
lines.push('  return out;');
lines.push('}');
lines.push('');

const content = lines.join('\n');
const eolContent = content.replace(/\n/g, detectEol());
function detectEol() {
  if (!fs.existsSync(OUT)) return '\n';
  return fs.readFileSync(OUT, 'utf8').includes('\r\n') ? '\r\n' : '\n';
}

const before = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
if (before === eolContent) {
  console.log(`= ${path.relative(ROOT, OUT)} 无变化（${sorted.length} 条别名）`);
  process.exit(0);
}
fs.writeFileSync(OUT, eolContent, 'utf8');
console.log(`✓ ${path.relative(ROOT, OUT)}：${sorted.length} 条别名，其中 ${mapped.length} 条带参数映射`);
for (const { from, to } of sorted) {
  const m = mapped.find((e) => e.from === from);
  const extra = m && (m.rename || m.set) ? '  [+参数映射]' : '';
  console.log(`    ${from} → ${to}${extra}`);
}
