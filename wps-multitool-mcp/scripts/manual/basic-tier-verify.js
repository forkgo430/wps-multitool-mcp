/**
 * 验证"默认档 = 基础档白名单 ∩ 真机验证过"这个承诺。
 *
 * 坑：wps_execute_method 等 12 个是 mcp-server 在启动时【单独注册】的内置工具，
 * 不在 src/tools/index.ts 的 allTools 里。只注册 allTools 会少算 1 个，
 * 看起来像"基础档掉了" —— 实际生产数字要用真实注册路径量。
 *
 * 第二个坑：数字不能再写死。白名单回答"常不常用"，验证状态回答"跑通过没有"，
 * 两个条件都会变（PPT 验完就会涨）。所以这里全部按注册表 + manifest 现算，
 * 只断言"关系成立"，不写 40 也不写 32。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { ESSENTIAL_TOOLS } = require(path.join(ROOT, 'dist', 'server', 'tool-tiers.js'));
const { isLiveVerified, verificationSummary } = require(path.join(ROOT, 'dist', 'server', 'tool-verification.js'));
const { groupOfTool, ALWAYS_LOADED_GROUP } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
};

// 走真实注册路径（内置 12 个也要注册）—— 只注册 allTools 会少算内置工具，
// 数字差 1 却看不出原因，这个坑吃过一次。
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const allDefs = toolRegistry.listAllTools();
const all = allDefs.map((t) => t.name);
const listed = toolRegistry.listTools().tools.map((t) => t.name);
const ess = new Set(ESSENTIAL_TOOLS);
const core = allDefs.filter((d) => groupOfTool(d) === ALWAYS_LOADED_GROUP).map((d) => d.name);

const ghost = ESSENTIAL_TOOLS.filter((n) => !all.includes(n));
const essLive = ESSENTIAL_TOOLS.filter((n) => isLiveVerified(n) && all.includes(n));
const essStatic = ESSENTIAL_TOOLS.filter((n) => !isLiveVerified(n) && all.includes(n));

const ver = verificationSummary(all);
console.log(`ESSENTIAL 白名单 ${ESSENTIAL_TOOLS.length} 个：live ${essLive.length} / static ${essStatic.length}`);
console.log(`core 组（永远列出）: ${core.length}`);
console.log(`默认列出: ${listed.length} = core ${core.length} + 基础档 live ${essLive.length}`);
console.log(`真机验证覆盖 ${ver.live}/${ver.total}\n`);

check('白名单没有"写了但不存在"的工具名', ghost.length === 0, ghost.join(', '));
check(
  `默认档全部真机验证过（${listed.filter((n) => isLiveVerified(n)).length}/${listed.length}）`,
  listed.every((n) => isLiveVerified(n)),
  listed.filter((n) => !isLiveVerified(n)).join(', ')
);
check(
  `默认列出 ${listed.length} = core ${core.length} + 基础档 live ${essLive.length}`,
  listed.length === core.length + essLive.length,
  `${listed.length} vs ${core.length + essLive.length}`
);
check(
  '默认档里的每个工具都在 core 或基础档白名单里',
  listed.every((n) => core.includes(n) || ess.has(n))
);
// 组级验证覆盖：PPT 整组 static 时，默认档里就不该有 PPT 工具；验一个涨一个，不写死 0
const pptEssentialLive = ESSENTIAL_TOOLS.filter(
  (n) => n.startsWith('wps_ppt_') && isLiveVerified(n) && all.includes(n)
);
const pptListed = listed.filter((n) => n.startsWith('wps_ppt_'));
check(
  `PPT 组默认列出 ${pptListed.length} = 基础档里 live 的 PPT 工具 ${pptEssentialLive.length}`,
  pptListed.length === pptEssentialLive.length,
  pptListed.join(', ')
);

// repair 是本轮新增的，确认它已注册但默认不列出（需要时按组加载）
check('wps_common_repair 已注册', all.includes('wps_common_repair'));
check('wps_common_repair 默认不列出', !listed.includes('wps_common_repair'));

const passed = results.filter((r) => r.ok).length;
console.log(`\n汇总：${passed}/${results.length} 通过`);
process.exit(passed === results.length ? 0 : 1);
