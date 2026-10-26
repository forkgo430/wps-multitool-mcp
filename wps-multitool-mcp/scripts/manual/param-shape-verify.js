#!/usr/bin/env node
/**
 * Input: param-shape-core 的扫描结果 + 下面的白名单
 * Output: 0 = 没有新的"TS 传了、PS 不读"；1 = 有新候选，或白名单里有失效条目
 * Pos: 参数形状门禁（第 3 轮新增）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 门禁的形状：白名单之外出现新候选就红。
 *
 * 为什么不能像别的 verify 那样"报出来让人看"
 * ------------------------------------------
 * 这一类 bug 的特点是【静默】：工具返回 success，文档也改了，但改的不是用户说的那处。
 * 它不会让任何现有测试变红 —— 注解、验证状态、工具数全都对得上。所以只能靠"新出现的候选"
 * 作为信号：一旦有人新写一个工具、新加一个参数，而忘了在 PS 侧接住，这里立刻红。
 *
 * 白名单为什么必须小
 * ------------------
 * 白名单是"已知且接受"的清单，每条都要写清为什么接受。条目一多就等于没有门禁：
 * 谁都可以往里加一条，问题就永远不会暴露。所以这里同时断言【白名单条目必须仍然被用到】——
 * 修好之后忘记删条目的，直接红。这是逼着白名单只减不增。
 *
 * 分类含义
 * --------
 *   已修-入口   入口归一化已经把值转成 PS 要的形状（见 wps-com.ps1 归一化块第 3 步），
 *               扫描器看不到入口的转换，所以仍会报 —— 留着是为了提醒别把它删了。
 *   OK-误报     PS 确实不读，但这个键是"无害冗余"（例如 TS 故意多发一个别名兜底）。
 *   DEPRECATED  属于第 2 轮标了 [DEPRECATED] 的旧工具，参数形状的历史遗留，等族合并轮一起收。
 *   延后        真缺实现，但不是"静默丢弃"而是"响亮失败"或"能力缺失"，下轮补。
 *
 * 用法：node scripts/manual/param-shape-verify.js
 */

'use strict';

const path = require('path');
const { scanCandidates } = require('./param-shape-core');

/** 'action|参数键' → { class, why } */
const ALLOWLIST = {
  // ===== 已修-入口：入口归一化把值转成了 PS 要的形状，扫描器看不到那一层 =====
  'hideRows|startRow': { class: '已修-入口', why: '入口第 3 步把 startRow/endRow 转成 PS 的 rows 数组（hideRows 只认 row/rows）' },
  'hideRows|endRow': { class: '已修-入口', why: '同上：转成 rows 数组的右端点' },
  'showRows|startRow': { class: '已修-入口', why: '入口第 3 步把 startRow/endRow 转成 PS 的 rows 数组' },
  'showRows|endRow': { class: '已修-入口', why: '同上：转成 rows 数组的右端点' },
  'showColumns|startColumn': { class: '已修-入口', why: '入口第 3 步把 startColumn/endColumn 转成 PS 的 columns 数字数组' },
  'showColumns|endColumn': { class: '已修-入口', why: '同上：字母列号在入口已 Convert-ColumnLetterToNumber' },

  // ===== OK-误报：TS 故意多发的冗余键，PS 不读也不影响结果 =====
  'convertToPDF|filePath': { class: 'OK-误报', why: 'TS 把 filePath/path 当作 outputPath 的别名兜底多发一份（见 convert.ts 注释），PS 读 outputPath' },
  'convertFormat|filePath': { class: 'OK-误报', why: '同上：filePath 是 outputPath 的别名兜底' },
  'findReplace|replaceMode': { class: 'OK-误报', why: '信息性布尔（"是不是替换模式"）；真正决定行为的是 replaceText 为空 + replaceAll' },

  // ===== DEPRECATED：第 2 轮标了 [DEPRECATED] 的旧工具，等族合并轮 =====
  // （addAnimation|animationType 已删：第 7 轮让 PS 侧直接认这个旧键，条目失效，门禁要求修好就删）
  'fillSeries|sourceRange': { class: 'DEPRECATED', why: '来自 wps_excel_auto_fill（已标 DEPRECATED，改用 wps_excel_fill_series 的 range/direction）' },
  'fillSeries|targetRange': { class: 'DEPRECATED', why: '同上：auto_fill 的旧参数名' },
  'hideRows|count': { class: 'DEPRECATED', why: '来自 wps_excel_hide_row（已标 DEPRECATED，改用 wps_excel_hide_rows 的 startRow/endRow）' },
  'hideRows|hide': { class: 'DEPRECATED', why: '同上：hide_row 的 show/hide 开关，hide_rows 只管隐藏' },
  'hideColumns|count': { class: 'DEPRECATED', why: '来自 wps_excel_hide_column（已标 DEPRECATED）' },
  'hideColumns|hide': { class: 'DEPRECATED', why: '同上：hide_column 的 show/hide 开关' },

  // ===== 延后：真缺实现，但要么响亮失败、要么是能力缺口，不是静默丢弃 =====
  'addAnimation|trigger': { class: '延后', why: 'PS 的 AddEffect 写死了 Trigger=1（单击），文档承诺的"开始方式"没实现' },
  'addCaption|newParagraph': { class: '延后', why: 'PS 用 position 决定插到段前/段后，newParagraph 是冗余表达' },
  'applyColorScheme|slideIndex': { class: '延后', why: 'PS 只支持整份文稿换配色；单页换色没实现（不会静默改错页）' },
  'copyRange|source': { class: '延后', why: 'PS 的 copyRange 读 range（剪贴板复制），source/destination 形态没实现' },
  'copyRange|destination': { class: '延后', why: '同上：目标区域形态没实现' },
  'copySheet|newName': { class: '延后', why: 'PS 的 copySheet 不支持改名（只 before/after 定位）' },
  'copySheet|position': { class: '延后', why: 'PS 用 before/after，TS 的 position（before/after 枚举）没映射' },
  'createDonutChart|data': { class: '延后', why: 'PS 的 createDonutChart 读 value/centerText，不吃 data 数组' },
  'createFlowChart|nodes': { class: '延后', why: 'PS 读 steps（已拍平的节点数组），TS 的 nodes/connections 形态没映射' },
  'createFlowChart|connections': { class: '延后', why: '同上：连线信息 PS 侧未实现' },
  'createGauge|max': { class: '延后', why: 'PS 的 createGauge 不支持自定义量程' },
  'createSheet|position': { class: '延后', why: 'PS 的 createSheet 不支持插入位置（永远追加到末尾）' },
  'duplicateSlide|targetIndex': { class: '延后', why: 'PS 的 duplicateSlide 不支持指定目标位置' },
  'fillSeries|direction': { class: '延后', why: 'PS 读 type/range/step，TS 的 direction（down/right/up/left）没映射' },
  'getCellComments|range': { class: '延后', why: 'PS 的 getCellComments 按 sheet 读全表批注，range 过滤未实现' },
  'insertExcelImage|cell': { class: '延后', why: 'PS 用 left/top 定位，TS 的 cell（单元格地址）没转换' },
  'insertPptChart|data': { class: '延后', why: 'PS 的 insertPptChart 不写数据（要先插再 set_ppt_chart_data），TS 声称能一次传 data' },
  'setPptDateTime|format': { class: '延后', why: 'PS 只写 text（固定文本），日期格式串没实现' },
  'setPptFooter|show': { class: '延后', why: 'PS 只写 text，不支持显示/隐藏开关' },
  'setSlideNumber|startFrom': { class: '延后', why: 'PS 只写 visible，startFrom（起始页号）没实现' },
  'startSlideShow|fromSlide': { class: '延后', why: 'PS 只能从头放映，fromSlide 没实现' },
  'transpose|destination': { class: '延后', why: 'PS 读 targetCell/destinationCell，TS 的 destination 形态没映射' },
};

const { rows, caseCount, aliasCount, flattenCount, helperCount } = scanCandidates();

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

console.log('\n=== A. 扫描器自身 ===');
console.log(`  case 体 ${caseCount} 个（同名 case 只算第一个）、助手 ${helperCount} 个、` +
  `入口改名 ${aliasCount} 条、入口摊平 ${flattenCount} 条`);
ok('扫描器认得入口归一化表（改名/摊平）', aliasCount > 0 && flattenCount > 0,
  `${aliasCount} / ${flattenCount}`);
ok('扫描器认得助手函数里读的键', helperCount > 0, `${helperCount} 个`);

console.log('\n=== B. 白名单之外不许有新候选 ===');
const seen = new Set();
const newOnes = [];
const used = new Set();
for (const r of rows) {
  for (const k of r.missing.split(',')) {
    const id = `${r.act}|${k}`;
    seen.add(id);
    const rule = ALLOWLIST[id];
    if (rule) {
      used.add(id);
    } else {
      newOnes.push(`${id}  （${r.file}）`);
    }
  }
}
ok('没有新的"TS 传了、PS 不读"', newOnes.length === 0, newOnes.join(' | '));

console.log('\n=== C. 白名单本身不许发胖 ===');
const stale = Object.keys(ALLOWLIST).filter((k) => !seen.has(k));
ok('白名单没有失效条目（修好了就删条目）', stale.length === 0, stale.join(','));
ok('白名单规模 <= 60 条', Object.keys(ALLOWLIST).length <= 60, `${Object.keys(ALLOWLIST).length} 条`);

const byClass = {};
for (const k of used) {
  const c = ALLOWLIST[k].class;
  byClass[c] = (byClass[c] || 0) + 1;
}
console.log(`\n  白名单分布：${Object.entries(byClass).map(([k, v]) => `${k} ${v}`).join(' / ')}`);
console.log(`  当前候选 ${rows.length} 个调用点 / ${seen.size} 个键，其中已登记 ${used.size} 个`);

console.log(`\n汇总：${pass}/${pass + fail} 通过`);
if (fail) {
  console.log('\n参数形状门禁不通过 —— 新出现"TS 传了、PS 不读"的键。' +
    '\n要么在 wps-com.ps1 的入口归一化表（$ParamAlias / $StyleFlatten）里接住，' +
    '\n要么在对应 case 里读它，实在接不住就写进 ALLOWLIST 并写清为什么。');
  process.exit(1);
}