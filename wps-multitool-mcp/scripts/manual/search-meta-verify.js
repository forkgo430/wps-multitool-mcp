/**
 * Input: 无
 * Output: 中文关键词检索 + 别名/反例元数据的独立验证
 * Pos: 工具检索元数据核验脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 验三件事
 * --------
 * 1) 中文关键词能不能搜到（注意：工具描述本来就是中文，所以这条【不是】从 0 到 1，
 *    别写成"以前搜不到"——实测以前也搜得到）。这里量的是真实增量：关掉别名后
 *    少捞到多少。
 * 2) 近名工具（6 个 background / 5 个 animation / 2 个 toc / 2 个 zoom）
 *    搜出来之后，靠前的命中带不带 notFor（模型分辨的依据，这是本层的【主力价值】）。
 * 3) 打分顺序：别名是机器拼的，不能压过人写的描述命中。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { getToolMeta, __clearToolMetaCache, expandQuery } = require(path.join(ROOT, 'dist', 'server', 'tool-metadata.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));

// 走真实注册路径：registerBuiltinTools() 的 12 个 + src/tools 的 273 个。
// 只 registerAll 会漏掉内置那一批 —— wps_check_connection 的别名就是这么漏成 0 的，
// 搜「检查连接」命中 0 个。
const _inst = new WpsMcpServer();
_inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
function ok(label, cond, extra) {
  if (cond) {
    pass++;
    console.log(`  ✅ ${label}${extra ? '  ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ❌ ${label}${extra ? '  ' + extra : ''}`);
  }
}

const ALL = toolRegistry.listAllTools();

// 对照组：只用 name + description（模拟加别名之前的行为）
function legacySearch(kw) {
  const k = kw.toLowerCase();
  return ALL.filter((t) => {
    const def = toolRegistry.getTool(t.name)?.definition;
    if (!def) return false;
    return def.name.toLowerCase().includes(k) || (def.description || '').toLowerCase().includes(k);
  }).map((t) => t.name);
}

console.log('=== A. 中文关键词：别名到底补上了什么 ===');
const CN_CASES = [
  ['插入表格', 'wps_word_insert_table'],
  ['批注', 'wps_word_insert_comment'],
  // 第 5 轮 background 6→1：旧名 wps_ppt_set_background_color 已并入 wps_ppt_set_background（别名）
  ['背景色', 'wps_ppt_set_background'],
  // 第 8 轮 animation/transition 两族合并：加动画进 wps_ppt_set_animation（op='add'），旧名成别名
  ['动画', 'wps_ppt_set_animation'],
  ['切换', 'wps_ppt_set_transition'],
  ['转PDF', 'wps_convert_to_pdf'],
  ['目录', 'wps_word_generate_toc'],
  ['合并单元格', 'wps_excel_merge_cells'],
  ['公式', 'wps_excel_set_formula'],
  ['页眉', 'wps_word_insert_header'],
  ['幻灯片', 'wps_ppt_add_slide'],
];
for (const [kw, expect] of CN_CASES) {
  const before = legacySearch(kw);
  const after = toolRegistry.searchTools(kw, { limit: 50 }).map((h) => h.name);
  const hitIdx = after.indexOf(expect);
  ok(
    `搜"${kw}" → ${expect}`,
    hitIdx >= 0,
    `（旧: ${before.length} 条命中 / 新: ${after.length} 条，目标排第 ${hitIdx + 1}）`
  );
}

console.log('\n=== B. 真实增量：别名多捞到了什么（对照组）===');
// 注意前提：工具描述是中文，所以旧逻辑本来就有命中。这里量"只靠别名才捞到"的那部分。
function viaAlias(kw) {
  const k = kw.toLowerCase();
  return ALL
    .map((t) => toolRegistry.getTool(t.name)?.definition)
    .filter(Boolean)
    .filter((d) => getToolMeta(d).aliases.join(' ').toLowerCase().includes(k))
    .map((d) => d.name);
}
let gained = 0;
for (const [kw] of CN_CASES) {
  const L = new Set(legacySearch(kw));
  const only = viaAlias(kw).filter((n) => !L.has(n));
  gained += only.length;
  console.log(`  "${kw}"：描述命中 ${L.size}，别名额外 +${only.length}${only.length ? ' → ' + only.slice(0, 3).join('、') : ''}`);
}
ok('别名确实带来增量（不吹牛：是锦上添花，不是从 0 到 1）', gained > 0, `（合计 +${gained} 对）`);
ok('增量不靠吹：至少半数关键词本来就有描述命中', CN_CASES.filter(([kw]) => legacySearch(kw).length > 0).length >= CN_CASES.length / 2);

console.log('\n=== C. 近名工具：notFor 有没有给到 ===');
const AMBIG = [
  ['背景', ['wps_ppt_set_background', 'wps_ppt_set_background_color', 'wps_ppt_set_master_background']],
  ['动画', ['wps_ppt_set_animation', 'wps_ppt_add_animation']],
  ['目录', ['wps_word_generate_toc', 'wps_word_generate_doc_toc']],
  ['缩放', ['wps_excel_zoom', 'wps_excel_set_zoom']],
];
// 第 2 轮别名收敛后，上面这些"旧名"已不在注册表里：wps_word_generate_doc_toc 与
// wps_excel_zoom 降为别名；第 5 轮 background 6→1 之后，_color/_image/_gradient 也成了别名、
// _master 下线 —— 所以"背景"这一族在册的只剩一个。
// 这里按"当前在册的名字"重列，否则断言的是不存在的东西。
const AMBIG_LIVE = [
  ['背景', ['wps_ppt_set_background']],
  // 第 8 轮动画族 5→1 后，在册的是三个：设置（含 add/reorder）+ 查询 + 删除。
  // 覆盖断言要按"在册的"列，否则断言的是已经下线成别名的名字。
  ['动画', ['wps_ppt_set_animation', 'wps_ppt_get_animations', 'wps_ppt_remove_animation']],
  ['切换', ['wps_ppt_set_transition']],
  ['目录', ['wps_word_generate_toc']],
  ['缩放', ['wps_excel_set_zoom']],
];
for (const [kw, names] of AMBIG_LIVE) {
  const hits = toolRegistry.searchTools(kw, { limit: 30 });
  const withNotFor = hits.filter((h) => h.notFor && h.notFor.length);
  ok(`搜"${kw}" 靠前命中带 notFor`, withNotFor.length > 0,
     `（${withNotFor.length} 条带反例：${withNotFor.map((h) => h.name).slice(0, 3).join('、')}）`);
  // 目录/缩放/背景三族合并后只剩一个在册名字，覆盖断言改成"在册的都捞得到"
  const covered = names.filter((n) => hits.some((h) => h.name === n));
  ok(`搜"${kw}" 能捞到在册的全部（${names.length} 个中 ${covered.length} 个）`, covered.length === names.length,
     names.filter((n) => !covered.includes(n)).join(','));
}

console.log('\n=== D. 命中字段 ===');
const h1 = toolRegistry.searchTools('插入表格', { limit: 3 })[0];
ok('命中带 aliases', Array.isArray(h1.aliases) && h1.aliases.length > 0, JSON.stringify(h1.aliases));
ok('命中带 matchedBy', typeof h1.matchedBy === 'string', `matchedBy=${h1.matchedBy}`);
const hNoHint = toolRegistry.searchTools('背景', { limit: 10, hints: false });
ok('hints=false 时不给 examples/notFor', hNoHint.every((h) => !h.examples && !h.notFor));

// 别名是机器拼的（会串味），不能压过人写的描述命中
const mixed = toolRegistry.searchTools('合并单元格', { limit: 20 });
const ranks = mixed.map((h) => h.matchedBy);
const firstAlias = ranks.indexOf('alias');
const lastDesc = ranks.lastIndexOf('description');
ok(
  '打分：所有 description 命中都排在 alias 命中之前',
  firstAlias === -1 || lastDesc === -1 || lastDesc < firstAlias,
  `（顺序：${ranks.join(',')}）`
);

console.log('\n=== E. 别名覆盖率（不是只给 40 个基础档打的补丁）===');
let withZh = 0;
for (const t of ALL) {
  const def = toolRegistry.getTool(t.name)?.definition;
  if (!def) continue;
  const m = getToolMeta(def);
  if (m.aliases.some((a) => /[一-龥]/.test(a))) withZh++;
}
const rate = ((withZh / ALL.length) * 100).toFixed(1);
ok(`中文别名覆盖 ≥90% 的工具`, withZh / ALL.length >= 0.9, `（${withZh}/${ALL.length} = ${rate}%）`);

// 内置的（mcp-server 手写，不在 src/tools 下）必须有中文别名。
// 这批以前没进覆盖统计，wps_check_connection 的别名是 0 —— 搜「检查连接」命中 0 个。
// 第 2 轮后剩 8 个：4 个历史短名降为别名，别名不进注册表，其"可搜到"由别名层负责（见 alias-verify）。
const BUILTIN = [
  'wps_check_connection',
  'wps_get_active_workbook',
  'wps_get_active_presentation', 'wps_execute_method',
  'wps_cache_data', 'wps_get_cached_data', 'wps_list_cache', 'wps_clear_cache',
];
const noAlias = BUILTIN.filter((n) => {
  const def = toolRegistry.getTool(n)?.definition;
  if (!def) return true;
  return !getToolMeta(def).aliases.some((a) => /[一-龥]/.test(a));
});
ok('内置工具都有中文别名', noAlias.length === 0, noAlias.join(','));
ok('搜「检查连接」能命中 wps_check_connection',
   toolRegistry.searchTools('检查连接', { limit: 5, hints: false }).some((h) => h.name === 'wps_check_connection'));

console.log('\n=== F. 查询侧同义词扩展（第 9 轮：词表不能只用来拼别名）===');
// 背景：190 个词根的 TOKEN_ZH 一直是【正向】的（英文→中文），只用来拼别名，
// 从不参与查询。结果是"排版""格式设置"这类用户用词搜不到对应工具。
// 这里断言三件事：扩展得出来、分得清高低、不许压过人写的描述。
const SYNONYM_GROUPS = [['排版', '版式', '布局'], ['页眉', '页脚']];
for (const group of SYNONYM_GROUPS) {
  const linked = group.every((a) => expandQuery(a).some((t) => group.includes(t)));
  ok(`同义词组互相可扩展：${group.join('/')}`, linked,
    group.map((w) => `${w}→${expandQuery(w).join(',')}`).join(' ; '));
}
ok('英文词根能展开出中文（layout → 版式/布局）',
   expandQuery('layout').some((t) => /[一-龥]/.test(t)), expandQuery('layout').join(','));
ok('中文说法能反查英文词根',
   expandQuery('背景色').includes('background'), expandQuery('背景色').join(','));

// 分档红线：synonym 必须排在 alias 与 description 之后。
// 注意不能去读 h.score —— 返回前就把 score 剥掉了（不给模型看打分），
// 所以这里断言的是【排序行为】，那才是模型实际看到的。
const orderViolations = [];
let synSeen = 0;
for (const kw of ['排版', '格式', '背景', '表格', '版式', '页眉', '目录']) {
  const ranks = toolRegistry.searchTools(kw, { limit: 40, hints: false }).map((h) => h.matchedBy);
  const firstSyn = ranks.indexOf('synonym');
  if (firstSyn === -1) continue;
  synSeen++;
  const lastReal = Math.max(ranks.lastIndexOf('description'), ranks.lastIndexOf('alias'), ranks.lastIndexOf('name'));
  if (lastReal > firstSyn) {
    orderViolations.push(`${kw}(真命中@${lastReal + 1} > syn@${firstSyn + 1})`);
  }
}
ok('synonym 不插到 description/alias/name 命中前面（7 个词实测）',
   orderViolations.length === 0, orderViolations.length ? orderViolations.join(' ') : `走扩展的词 ${synSeen} 个`);

// 扩展必须真的捞到东西，否则上面那些断言全是在测一个没接上的机制
let synGain = 0;
for (const kw of ['排版', '版式', '格式', '页脚', '合并']) {
  synGain += toolRegistry.searchTools(kw, { limit: 30, hints: false })
    .filter((h) => h.matchedBy === 'synonym').length;
}
ok('同义词扩展带来了真实命中（不是空转）', synGain > 0, `合计 ${synGain} 条 synonym 命中`);

console.log('\n=== G. 中文错名纠偏（第 9 轮：原来中文一律返回空）===');
// scoreSimilar 原来只切 ASCII 词元 → 模型写中文错名时拿不到任何线索。
const TYPO_CASES = [
  ['合并单元格', 'wps_excel_merge_cells'],
  ['设置背景色', 'wps_ppt_set_background'],
];
for (const [bad, expect] of TYPO_CASES) {
  const near = toolRegistry.suggestSimilar(bad, 3);
  ok(`中文错名「${bad}」能纠到 ${expect}`, near[0]?.name === expect,
     near.length ? near.map((x) => x.name).join(',') : '（返回空 —— 阈值太高或没生效）');
}
ok('英文错名没被改坏（回归）',
   toolRegistry.suggestSimilar('set_merg_cells', 2)[0]?.name === 'wps_excel_merge_cells');
ok('无关中文不瞎猜（宁可空也别误导模型）',
   toolRegistry.suggestSimilar('今天天气怎么样', 3).length === 0,
   JSON.stringify(toolRegistry.suggestSimilar('今天天气怎么样', 3).map((x) => x.name)));

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
