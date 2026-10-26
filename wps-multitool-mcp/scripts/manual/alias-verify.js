#!/usr/bin/env node
/**
 * Input: src/server/tool-aliases.generated.ts + 真实注册路径（dist 注册表）
 * Output: 别名层的四条断言，任一不过即非 0 退出
 * Pos: 别名门禁。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么单独一个门禁，而不是塞进 verification-verify
 * ------------------------------------------------
 * 别名的失效方式全是【静默】的，这比崩掉更难发现：
 *   别名指向一个不存在的名字  → 旧名调用报 ToolNotFound，而工具数、注解数全对得上；
 *   别名自己也被注册          → 旧名回到 tools/list，上下文白省；
 *   别名进了工具总数          → 文档数字和注册表对不上，而两边都"像是对的"；
 *   生成物与真相源漂          → 人改了 JSON 没重跑生成器，线上用的是旧表。
 * 所以这四条必须每次提交都跑，而不是等人拿旧名去试。
 *
 * 注意：本脚本只读 dist，不起桥、不 spawn 子进程 —— 实现者与验证者都能跑。
 */

'use strict';

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist/server/tool-annotations.js'));
const { TOOL_ALIASES, TOOL_ALIAS_PARAMS } = require(path.join(ROOT, 'dist/server/tool-aliases.generated.js'));

// 走真实注册路径：registerBuiltinTools() 的内置 + registerAll(allTools) 的专业工具。
// 漏一个都会让"目标存在"的断言变成假的。
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const registered = new Set(toolRegistry.listAllTools().map((d) => d.name));
const entries = Object.entries(TOOL_ALIASES);
const aliasNames = Object.keys(TOOL_ALIASES);

console.log('\n=== A. 别名表自身 ===');
ok('别名表非空', entries.length > 0, `${entries.length} 条`);
ok(
  '没有自引用',
  entries.every(([a, c]) => a !== c),
  entries.filter(([a, c]) => a === c).join(',')
);
// 多个旧名指向同一规范名是【合法的】：族合并就是"多个旧名 → 一个规范名"
// （background 6→1 有 4 个旧名指向 wps_ppt_set_background）。
// 这里只统计分布，不做断言 —— 真正该拦的是"别名指向不存在的名字"，见下面 B 段。
const targets = new Map();
for (const [, c] of entries) targets.set(c, (targets.get(c) || 0) + 1);
const merged = [...targets.entries()].filter(([, n]) => n > 1);
ok(
  '别名表结构可读：' + targets.size + ' 个规范名，其中 ' + merged.length + ' 个吸收了多个旧名（族合并）',
  true,
  merged.map(([c, n]) => `${c}×${n}`).join(',') || '(无合并族)'
);
ok('别名名本身不重复', new Set(aliasNames).size === aliasNames.length, `${aliasNames.length} 条`);
// 链式别名（a→b 且 b 也是别名）：解析只做一跳，链式会让 a 落到不存在的名字上
const chained = entries.filter(([, c]) => aliasNames.includes(c)).map(([a]) => a);
ok('没有链式别名', chained.length === 0, chained.join(','));

console.log('\n=== B. 别名目标与解析 ===');
const missing = entries.filter(([, c]) => !registered.has(c));
ok('每个别名目标都在注册表里', missing.length === 0, missing.map(([a, c]) => `${a}→${c}`).join(','));

const wrong = entries.filter(([a, c]) => {
  const got = toolRegistry.getTool(a);
  return !got || got.definition.name !== c;
});
ok('旧名入口解析到规范名', wrong.length === 0, wrong.map(([a, c]) => `${a}→${c}`).join(','));

const notHas = aliasNames.filter((a) => !toolRegistry.hasTool(a));
ok('旧名 hasTool 为 true（mcp-server 与 batch 的守卫）', notHas.length === 0, notHas.join(','));

// 注解必须跟随规范名：否则旧名没有破坏性提示，客户端会把它当安全操作自动执行
const annDiff = aliasNames.filter((a) => {
  const got = toolRegistry.getTool(a);
  return got && JSON.stringify(getToolAnnotations(a)) !== JSON.stringify(getToolAnnotations(got.definition.name));
});
ok('旧名拿到与规范名同一份注解', annDiff.length === 0, annDiff.join(','));

console.log('\n=== C. 别名不许污染对外面 ===');
const listed = toolRegistry.listTools().tools.map((d) => d.name);
const aliasListed = aliasNames.filter((a) => listed.includes(a));
ok('别名不出现在 tools/list', aliasListed.length === 0, aliasListed.join(','));

const aliasRegistered = aliasNames.filter((a) => registered.has(a));
ok('别名自己没被注册成工具（否则会双重计数）', aliasRegistered.length === 0, aliasRegistered.join(','));

const size = toolRegistry.size;
const unique = new Set([...registered, ...aliasNames]).size;
ok('注册表大小 = 真实工具数（别名不计入）', size === registered.size, `${size} vs ${registered.size}`);

console.log('\n=== D. 生成物与真相源一致 ===');
// 真相源 scripts/manual/tool-aliases.json 的 aliases 键值集合必须与 generated 完全一致。
// value 有两种形态：字符串（纯改名）或 { to, rename, set }（带参数映射，第 8 轮加的），两种都要比。
const fs = require('fs');
const SRC_JSON = path.join(ROOT, 'scripts/manual/tool-aliases.json');
const truth = JSON.parse(fs.readFileSync(SRC_JSON, 'utf8')).aliases || {};
const truthKeys = Object.keys(truth).sort();
const genKeys = aliasNames.slice().sort();
const sameKeys = JSON.stringify(truthKeys) === JSON.stringify(genKeys);
const sameVals = truthKeys.every((k) => {
  const t = truth[k];
  return (typeof t === 'string' ? t : t && t.to) === TOOL_ALIASES[k];
});
ok('generated 与 tool-aliases.json 一致（改了 JSON 要重跑生成器）', sameKeys && sameVals,
  sameKeys && sameVals ? `${genKeys.length} 条` : `JSON ${truthKeys.length} / generated ${genKeys.length}`);

console.log('\n=== E. 参数映射不许指向不存在的参数 ===');
// 参数映射是新的失效面：rename/set 里的键如果目标工具 schema 不认，
// 旧调用会"看起来成功了"但参数被忽略 —— 比直接报错更难查，所以在这里挡掉。
const params = Object.entries(TOOL_ALIAS_PARAMS || {});
const unknownKeys = [];
for (const [alias, map] of params) {
  const canonical = TOOL_ALIASES[alias];
  const props = (toolRegistry.getTool(canonical) || { definition: { inputSchema: { properties: {} } } })
    .definition.inputSchema.properties || {};
  const wanted = [
    ...Object.values(map.rename || {}),
    ...Object.keys(map.set || {}),
  ];
  for (const key of wanted) {
    if (!(key in props)) unknownKeys.push(`${alias}:${key}→${canonical}`);
  }
}
ok(
  '映射到的参数名都在目标工具的 schema 里',
  unknownKeys.length === 0,
  unknownKeys.length ? unknownKeys.join(',') : `${params.length} 个旧名带映射`
);

const oldSideGone = Object.entries(TOOL_ALIASES).filter(([a]) =>
  toolRegistry.listAllTools().some((d) => d.name === a)
);
ok(
  '映射表的旧名自己都没被注册（旧名必须在映射生效的一侧）',
  oldSideGone.length === 0,
  oldSideGone.map(([a]) => a).join(',')
);

console.log(`\n汇总：${pass}/${pass + fail} 通过`);
if (fail) {
  console.log('\n别名门禁不通过 —— 别名失效是静默的，别绕过。');
  process.exit(1);
}
