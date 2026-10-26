/**
 * 验证"每个工具都有验证状态、默认档里没有未验证的工具"这两个承诺。
 *
 * 为什么需要门禁而不是靠自觉
 * ------------------------
 * 276 个工具时没人管验证状态也没出事；到 306 个时 PPT 已经占了一整组，
 * 却一个都没在真机上跑过。这种"沉默的缺口"不会自己报警 —— 必须有人因为数字对不上而失败。
 *
 * 查七件事：
 *   1) 覆盖率：注册表里每个工具都能查到状态（漏 1 个就红）
 *   2) 生成物与 manifest 一致（改了 json 忘了跑生成器）
 *   3) 默认档里没有 static 工具（默认档不能替未验证的东西背书）
 *   4) 默认档数量 = core 组 + 基础档白名单里 live 的那批（数字不许手写）
 *   5) ppt 组里的 live 必须在 manifest.tools 里逐条点名，不能靠组默认浑过去
 *   6) manifest 里没有死条目
 *   7) 源码头部注释里的工具数/PPT 数不许漂（这两处不在任何生成物覆盖范围里，只能靠断言守）
 */
const path = require('path');
const fs = require('fs');
const ROOT = path.resolve(__dirname, '../..');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool, ALWAYS_LOADED_GROUP } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { TOOL_VERIFICATION } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.generated.js'));
const { verificationOf, isLiveVerified } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));
const { ESSENTIAL_TOOLS } = require(path.join(ROOT, 'dist', 'server', 'tool-tiers.js'));

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
};

// 真实注册路径：漏 registerBuiltinTools 就少 12 个（老坑）
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();
const names = all.map((d) => d.name);
const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, 'tool-verification.json'), 'utf8'));

// 1) 覆盖率
const missing = names.filter((n) => TOOL_VERIFICATION[n] === undefined);
check(`验证状态覆盖 ${names.length - missing.length}/${names.length}`, missing.length === 0, missing.slice(0, 8).join(', '));

// 2) 生成物与 manifest 一致
const expected = {};
for (const d of all) {
  const g = groupOfTool(d);
  const s = manifest.tools[d.name] || manifest.groups[g];
  if (s) expected[d.name] = s;
}
const drift = names.filter((n) => expected[n] && TOOL_VERIFICATION[n] !== expected[n]);
check(
  'tool-verification.generated.ts 与 manifest 一致',
  drift.length === 0,
  drift.length ? `${drift.slice(0, 5).join(', ')}（跑 node scripts/manual/gen-tool-verification.js）` : ''
);

// 3) 默认档里不许有 static
const listed = toolRegistry.listTools().tools.map((d) => d.name);
const staticListed = listed.filter((n) => !isLiveVerified(n));
check('默认档里没有未验证的工具', staticListed.length === 0, staticListed.join(', '));

// 4) 默认档数量 = core 组 + 基础档 live
const coreNames = all.filter((d) => groupOfTool(d) === ALWAYS_LOADED_GROUP).map((d) => d.name);
const essentialLive = ESSENTIAL_TOOLS.filter((n) => isLiveVerified(n) && names.includes(n));
const expectedListed = new Set([...coreNames, ...essentialLive]);
check(
  `默认列出 ${listed.length} = core ${coreNames.length} + 基础档 live ${essentialLive.length}`,
  listed.length === expectedListed.size && listed.every((n) => expectedListed.has(n)),
  `实际 ${listed.length}，期望 ${expectedListed.size}`
);

// 5) ppt 组里的 live 必须逐条点名
const pptLive = all
  .filter((d) => groupOfTool(d) === 'ppt' && isLiveVerified(d.name))
  .map((d) => d.name);
const pptUnnamed = pptLive.filter((n) => !manifest.tools[n]);
check(
  'PPT 的 live 全部逐条写进 manifest.tools（不能靠组默认）',
  pptUnnamed.length === 0,
  pptUnnamed.join(', ')
);

// 6) manifest 里的死条目（写了但工具不存在）
const registered = new Set(names);
const dead = Object.keys(manifest.tools || {}).filter((n) => !registered.has(n));
check('manifest.tools 没有死条目', dead.length === 0, dead.join(', '));

// ---------------------------------------------------------------------------
// 7) 源码头部注释里的数字不许漂（第 1 期第 2 章第 1 轮加）
// ---------------------------------------------------------------------------
// 为什么要有这条
// --------------
// tool-tiers.ts 与 tool-verification.ts 的头部注释里写着"全量 N 个工具""PPT 那 N 个"。
// 这两个文件【不在任何生成物的覆盖范围里】—— 改注释不会让任何门禁变红，
// 也不会被生成器自动改回来。所以它们只能靠这条断言守住。
// 这不是假设：第 1 期第 2 章开工前，它们里面写的是 306 / PPT 116，
// 而实测是 284 / PPT 101（PPT 组数含内置工具，与 list_tool_groups 同口径）。
//
// 断言怎么写才不会自己变成新的数字来源
// ----------------------------------------
// 期望值全部从【实测】来，不写死：总数取注册表 listAllTools()，
// ppt 组数取 groupOfTool()（内置工具按 category 计入对应组，与 list_tool_groups 同口径）。
// 注释里匹配到的【每一个】数字都要对 —— 两个地方写了不同的数，同样算漂。
const headerOf = (src) => {
  const end = src.indexOf('*/');
  return end === -1 ? src : src.slice(0, end + 2);
};
const allNumbers = (src, re) => {
  const out = [];
  let m;
  const r = new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g');
  while ((m = r.exec(src)) !== null) out.push(Number(m[1]));
  return out;
};

const tiersSrc = fs.readFileSync(path.join(ROOT, 'src', 'server', 'tool-tiers.ts'), 'utf8');
const verSrc = fs.readFileSync(path.join(ROOT, 'src', 'server', 'tool-verification.ts'), 'utf8');
const tiersHeader = headerOf(tiersSrc);
const verHeader = headerOf(verSrc);

const total = names.length;
const pptGroupCount = all.filter((d) => groupOfTool(d) === 'ppt').length;
const pptEssentialCount = ESSENTIAL_TOOLS.filter((n) => n.startsWith('wps_ppt_')).length;

const numCheck = (label, values, expected) => {
  const ok = values.length > 0 && values.every((v) => v === expected);
  check(label, ok, values.length ? `注释里写的是 ${values.join('/')}，期望 ${expected}` : '注释里没找到这个数');
};

numCheck(
  `tool-tiers.ts 头部注释的「全量 N 个工具」= 注册表实测 ${total}`,
  allNumbers(tiersHeader, /全量\s*(\d+)\s*个工具/),
  total
);
numCheck(
  `tool-verification.ts 头部注释的「N 个工具」= 注册表实测 ${total}`,
  allNumbers(verHeader, /(\d+)\s*个工具里/),
  total
);
numCheck(
  `tool-verification.ts 头部注释的「PPT 那 N 个」= ppt 组实测 ${pptGroupCount}`,
  allNumbers(verHeader, /PPT 那\s*(\d+)\s*个/),
  pptGroupCount
);
numCheck(
  `tool-tiers.ts 头部注释的「PPT 那 N 个基础工具」= 基础档白名单里 ppt 工具实测 ${pptEssentialCount}`,
  allNumbers(tiersHeader, /PPT 那\s*(\d+)\s*个基础工具/),
  pptEssentialCount
);

const live = names.filter((n) => verificationOf(n) === 'live').length;
const passed = results.filter((r) => r.ok).length;
console.log(`\n真机验证 ${live}/${names.length}（static ${names.length - live}）`);
console.log(`汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
