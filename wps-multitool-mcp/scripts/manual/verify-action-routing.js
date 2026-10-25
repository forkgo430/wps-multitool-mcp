// 第 1 期第 2 章第 1 轮 · 验证者 交叉核对（第 7 条）
// 第 1 期第 2 章第 2 轮 · 由实现者改为复用 gen-action-routes.js 的解析器。
//
// 做的事：
//   ① 把注解表里每个工具的动作名，逐一比对 wps-com.ps1 的【顶层】switch($Action) case；
//      对不上的就是"这工具调用时会回 Unknown action（等于从未生效）"。
//   ② 找出"深度 1 上的 case 标签但花括号后还带内容"的那些（生成器的 CASE_RE 收不了它们），
//      确认有没有真动作被漏出路由表。
//
// 为什么不再自己写一遍解析：
//   原来的做法是"从 switch 行扫到文件末尾"，于是嵌套 switch 的分支标签
//   （trim / remove_duplicates / unify_date / remove_empty_rows / placeholder / underline /
//   afterColon / afterLabel / acceptOne / rejectOne）也被算成了 ps1 的合法动作 ——
//   那会让①永远查不出问题（拿幽灵名去比幽灵名），而路由表同时虚高 10。
//   现在与 gen-action-routes.js 共用 parseTopLevelCases（括号深度判顶层），口径只有一份。
//
// 注意：这是补充核对，不是门禁。门禁在 scripts/manual/gen-tool-annotations.js 里
//      （生成时校验，查不到就非 0 退出）。

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..'); // wps-multitool-mcp 根

const { parseTopLevelCases } = require(path.join(ROOT, 'scripts', 'gen-action-routes.js'));

// 1) ps1 顶层 case（与路由表生成器同一份解析器）
const parsed = parseTopLevelCases(path.join(ROOT, 'scripts', 'wps-com.ps1'));
const ps1Cases = new Set(parsed.cases.map((c) => c.name));

// 2) 注解表（dist 已构建）
const ANN = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.generated.js'));
const TOOL_ANNOTATIONS = ANN.TOOL_ACTIONS || (ANN.default && ANN.default.TOOL_ACTIONS);

// 3) 路由表实际收录的动作键
const RT = require(path.join(ROOT, 'dist', 'client', 'action-routes.generated.js'));
const ACTION_ROUTES = RT.ACTION_ROUTES || (RT.default && RT.default.ACTION_ROUTES);

// 4) 交叉核对：注解里的动作 → ps1 顶层
const tools = Object.keys(TOOL_ANNOTATIONS);
const missingByTool = {};
const missingGlobal = {};
let annotatedActionTotal = 0;
for (const tool of tools) {
  const ann = TOOL_ANNOTATIONS[tool];
  const acts = Array.isArray(ann) ? ann : (ann && ann.actions) || [];
  for (const a of acts) {
    annotatedActionTotal++;
    if (!ps1Cases.has(a)) {
      (missingByTool[tool] = missingByTool[tool] || []).push(a);
      (missingGlobal[a] = missingGlobal[a] || []).push(tool);
    }
  }
}

// 5) 深度 1 上"格式不纯"的 case 标签（花括号后有内容）且不在路由表 → 可能漏了真动作
const dropped = parsed.labels.filter((l) => !l.pure && !ACTION_ROUTES[l.name]);

console.log(`工具总数(注解表): ${tools.length}`);
console.log(`ps1 顶层 case(生成器能收): ${ps1Cases.size}  （深度 1 标签共 ${parsed.labels.length} 个，其中格式不纯 ${parsed.labels.filter((l) => !l.pure).length} 个）`);
console.log(`注解里出现过的动作总数: ${annotatedActionTotal}`);
console.log(`路由表 ACTION_ROUTES 收录动作数: ${Object.keys(ACTION_ROUTES).length}`);
const dupInfo = parsed.dups.length ? `重复 ${parsed.dups.length} 个：${[...new Set(parsed.dups)].join(', ')}` : '无重复';
console.log(`顶层 case 同名重复: ${dupInfo}`);
console.log('');
console.log(`=== 注解动作不在 ps1 顶层 case 里的：共 ${Object.keys(missingGlobal).length} 个动作 / ${Object.keys(missingByTool).length} 个工具 ===`);
for (const a of Object.keys(missingGlobal).sort()) {
  console.log(`  动作 "${a}"  ←  工具: ${missingGlobal[a].join(', ')}`);
}
console.log('');
console.log(`=== 生成器正则漏收的顶层 case（花括号后带内容，且不在路由表）：共 ${dropped.length} 个 ===`);
for (const d of dropped) {
  console.log(`  :${d.line}  "${d.name}"  (检查是否真动作被漏)`);
}

// 退出码：查不到的动作名是硬问题（与生成器门禁一致），漏收标签只是线索。
process.exit(Object.keys(missingGlobal).length ? 1 : 0);
