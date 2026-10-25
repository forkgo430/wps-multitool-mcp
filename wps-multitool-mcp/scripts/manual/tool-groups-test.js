/**
 * Input: 无
 * Output: 工具档位（基础档 / 全量）+ 工具组加载 的断言结果
 * Pos: 工具档位与工具组的手工验证脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用法（先 npm run build）：node scripts/manual/tool-groups-test.js
 *
 * 注意：除了 allTools，这里还给内置的 wps_execute_method 补了一个 stub ——
 * 它在生产里由 mcp-server 注册，本脚本不起 server，补出来才和生产口径一致。
 * 默认档数字不再写死：= core + （白名单 ∩ 真机验证），PPT 验完一个就涨一个。
 * 通知（notifications/tools/list_changed）不在这里测 —— 那需要真客户端，
 * 这里只断言"变更事件确实被触发"。
 */

const path = require('path');
const { execFileSync } = require('child_process');
const ROOT = path.resolve(__dirname, '../..');
const NODE = process.execPath;

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { ToolCategory } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { onToolGroupsChanged } = require(path.join(ROOT, 'dist', 'tools', 'common', 'toolgroups.js'));
const { ESSENTIAL_TOOL_NAMES, findUnknownEssentialTools, isEssentialTier } = require(
  path.join(ROOT, 'dist', 'server', 'tool-tiers.js')
);
const { isLiveVerified } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));

const results = [];
function check(name, cond, detail) {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✅' : '❌'} ${name}${detail ? '  → ' + detail : ''}`);
}

toolRegistry.registerAll(allTools);
if (!toolRegistry.hasTool('wps_execute_method')) {
  toolRegistry.register(
    {
      name: 'wps_execute_method',
      description: 'stub（生产里由 mcp-server 注册，这里补出来以对齐口径）',
      inputSchema: { type: 'object', properties: {} },
      category: ToolCategory.COMMON,
    },
    async () => ({ id: '', success: true, content: [{ type: 'text', text: 'stub' }] })
  );
}

const registeredNames = new Set(toolRegistry.listAllTools().map((t) => t.name));

console.log('=== 档位与默认列出数 ===');
const groups0 = toolRegistry.getGroups();
const coreGroup = groups0.find((g) => g.group === 'core');
const pptGroup = groups0.find((g) => g.group === 'ppt');
console.log(`  档位: ${toolRegistry.tier}（isEssentialTier=${isEssentialTier()}）`);
console.log(
  '  ' +
    groups0
      .map((g) => `${g.group} 列出${g.listed}/${g.total}${g.fullyLoaded ? '[整组]' : ''}`)
      .join('  ')
);
const listed0 = toolRegistry.listTools().tools.length;
const total0 = toolRegistry.listAllTools().length;
// 默认档不再是一个固定数字：= core + （白名单 ∩ 真机验证过）。
// 写死 40 的门禁会在 PPT 验完那天开始"红得莫名其妙"，所以这里现算。
const essentialLive = [...ESSENTIAL_TOOL_NAMES].filter((n) => isLiveVerified(n) && registeredNames.has(n));
const essentialStatic = [...ESSENTIAL_TOOL_NAMES].filter((n) => !isLiveVerified(n) && registeredNames.has(n));
const expectedListed = coreGroup.total + essentialLive.length;
console.log(`  默认列出 ${listed0} / 共 ${total0}`);
console.log(
  `  计算值 = core ${coreGroup.total} + 白名单里 live ${essentialLive.length} = ${expectedListed}` +
    `（白名单还有 ${essentialStatic.length} 个 static 不占名额）`
);

check('档位是 essential', toolRegistry.tier === 'essential');
check(`默认列出 ${listed0} = core + 白名单 live`, listed0 === expectedListed, `${expectedListed} vs ${listed0}`);
check('总工具数 > 250', total0 > 250, `${total0}`);

console.log('\n=== 白名单自检（防手误）===');
const unknown = findUnknownEssentialTools(registeredNames);
check('基础档里没有"写了但不存在"的工具名', unknown.length === 0, unknown.length ? unknown.join(', ') : '无');

const listedNames = new Set(toolRegistry.listTools().tools.map((t) => t.name));
const pptEssentialLive = [...ESSENTIAL_TOOL_NAMES].filter(
  (n) => n.startsWith('wps_ppt_') && isLiveVerified(n) && registeredNames.has(n)
).length;
check(
  `ppt 组默认只列出基础档里 live 的 ${pptEssentialLive} 个` + '（未整组列出）',
  pptGroup.listed === pptEssentialLive,
  `${pptGroup.listed}`
);
check('core 组全部列出', coreGroup.listed === coreGroup.total, `${coreGroup.listed}/${coreGroup.total}`);
check('core 里有 route_intent（意图识别）', listedNames.has('wps_route_intent'));
check('核心管理工具都在', listedNames.has('wps_load_tool_group') && listedNames.has('wps_search_tools'));

console.log('\n=== 搜索（含未列出）===');
const hits = toolRegistry.searchTools('ppt', { limit: 5 });
check('搜得到 ppt 工具', hits.length > 0 && hits.some((h) => h.group === 'ppt'));
// ⚠ 第 10 轮 T3b 改了同分排序，这两条断言跟着改了判据 —— 理由写在这里：
//
// 【原断言为什么不成立】
//   原来是 `hits.some(h => h.group==='ppt' && h.listed===false)`，limit=5。
//   T3b 把同分改成 (score, isListed, isLive, name) 四元之后，查询「ppt」的并列项里
//   8 个 listed+live 的基础档工具（add_slide / get_slide_count / insert_image …）
//   全部排到了前面，limit=5 截出来的 5 条**全是 listed**，于是原断言失败。
//   但它想验的性质——「searchTools 能搜到【未列出】的工具」——并没有变，
//   变的是"用哪个探针去验"。原来那个探针把"未列出的必须挤进前 5"当成了性质本身，
//   而那从来不是性质，是排序实现的副作用。
//
// 【现在的判据】
//   性质用【不受排序影响】的探针验：放大 limit，看未列出的 ppt 工具在不在结果集里。
//   排序本身单开一条断言钉住，免得以后有人把 T3b 改回去。
const allPptHits = toolRegistry.searchTools('ppt', { limit: 200 });
const unlistedPpt = allPptHits.filter((h) => h.group === 'ppt' && h.listed === false);
check(
  '搜索能覆盖未列出的工具（放大 limit，不受同分排序影响）',
  unlistedPpt.length > 0,
  `未列出的 ppt 工具 ${unlistedPpt.length} 个`
);
const top5 = hits.filter((h) => h.group === 'ppt');
const listedFirst = top5.every(
  (h, i) => i === 0 || !top5[i - 1].listed || top5[i - 1].listed === h.listed
);
check(
  '同分时已列出的排前面（T3b 四元排序）',
  listedFirst,
  top5.map((h) => `${h.listed ? 'listed' : '     '}:${h.name}`).join(' ')
);

console.log('\n=== 加载 ppt 组 ===');
let notified = 0;
onToolGroupsChanged(() => {
  notified += 1;
});

const load = toolRegistry.loadGroup('ppt');
const listed1 = toolRegistry.listTools().tools.length;
check('loadGroup 返回 ok 且有变化', load.ok && load.changed);
check('变更事件触发一次', notified === 1, `notified=${notified}`);
check(
  `整组加载后列出数量 = ${listed0} - ${pptEssentialLive} + ${pptGroup.total} = ${listed0 - pptEssentialLive + pptGroup.total}`,
  listed1 === listed0 - pptEssentialLive + pptGroup.total,
  `${listed0} → ${listed1}`
);
check(
  `ppt 整组 ${pptGroup.total} 个都已列出`,
  toolRegistry.getGroups().find((g) => g.group === 'ppt').listed === pptGroup.total
);

const load2 = toolRegistry.loadGroup('ppt');
check('重复 load 不再触发通知', load2.changed === false && notified === 1);

console.log('\n=== 保护 core / 卸载 ppt ===');
check('core 不允许卸载', toolRegistry.unloadGroup('core').ok === false);
const un = toolRegistry.unloadGroup('ppt');
check('卸载触发通知', notified === 2, `notified=${notified}`);
check(
  `卸载后回到默认档 ${listed0}`,
  toolRegistry.listTools().tools.length === listed0,
  `${toolRegistry.listTools().tools.length}`
);
check('不存在的组报错', toolRegistry.loadGroup('nope').ok === false && toolRegistry.unloadGroup('nope').ok === false);

console.log('\n=== 全量档（子进程带 WPS_TOOL_TIER=full 跑一遍）===');
let fullOut = '';
try {
  const code =
    `const r=require(${JSON.stringify(path.join(ROOT, 'dist', 'server', 'tool-registry.js'))});` +
    `const t=require(${JSON.stringify(path.join(ROOT, 'dist', 'tools', 'index.js'))});` +
    `r.toolRegistry.registerAll(t.allTools);` +
    `console.log(JSON.stringify({tier:r.toolRegistry.tier, listed:r.toolRegistry.listTools().tools.length}));`;
  fullOut = execFileSync(NODE, ['-e', code], {
    encoding: 'utf8',
    env: { ...process.env, WPS_TOOL_TIER: 'full' },
  }).trim();
} catch (e) {
  fullOut = 'ERR: ' + String(e.message).slice(0, 120);
}
console.log('  子进程输出:', fullOut);
let full = null;
try {
  full = JSON.parse(fullOut);
} catch {
  /* 保持 null */
}
check('full 档位生效', !!full && full.tier === 'full', full ? `tier=${full.tier}` : fullOut);
check('full 档列出数远大于 40（回到按组切）', !!full && full.listed > 100, full ? `${full.listed}` : 'n/a');

const passed = results.filter((r) => r.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
