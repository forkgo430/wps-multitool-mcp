/**
 * Input: 工作表详情 / 自适应 / 自动换行 / 表内查找 / 图表删除 / 透视表三件 / 页面设置 / 打印标题
 * Output: 各自的查询或修改结果
 * Pos: 第 4 期补齐的高频缺口。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 这一批补的是什么
 * ----------------
 * 对着同类 WPS 自动化工具的公开能力清单逐项比对后挑出来的**高频缺口**
 * （同类 118、本仓 85，其中真缺的约 20 个，这里取最常被用到的 10 个）。
 * 刻意不补的：ListObject（表格对象）8 个、迷你图、goal_seek、公式审计 ——
 * 真实办公里用得极少，补了只是让模型的选择负担变大。
 *
 * 每个工具都写了"补它是因为什么会卡住"，不是凑数：
 *   get_sheet_info       以前只有表名列表，"这张表多大/用到哪/藏了哪些行"无从回答
 *   auto_fit             列宽只能手填，中文表头一换就得逐列重调（最高频重复劳动）
 *   set_wrap_text        表格排版最常碰的动作，之前完全没有
 *   find_in_sheet        以前只有 find_replace（改），没有"找"。用户问"这个数在哪"答不上
 *   delete_chart         只有 create/update，建错了删不掉 —— COM 没有撤销
 *   get/refresh/clear_pivot_tables  建完看不到、源数据变了刷不了、改错了清不掉
 *   set_sheet_page_setup 只有 set_print_area，"每页重复表头/横向打印/纸张"全做不了
 *   set_sheet_print_titles 多页表刚需
 *
 * 真机支持度：全部走标准 Excel COM 接口（WPS 为兼容 Office 而实现），
 * 但 WPS 12.1 有个别属性与 Excel 不同 —— 每个动作都做了逐属性容错，
 * 不支持的那个走 warnings 而不是整次失败。没验过的接口在能力探测里有对应项，
 * 见 scripts/manual/caps-probe.js。
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
  ToolWarning,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

type Args = Record<string, unknown>;

/** 统一调用：一次失败一个成功，warnings 一路带回 */
async function call<T>(action: string, params: Args): Promise<{
  ok: boolean;
  data?: T;
  error?: string;
  warnings?: ToolWarning[];
}> {
  const r = await wpsClient.executeMethod<T>(action, params, WpsAppType.SPREADSHEET);
  return { ok: r.success, data: r.data, error: r.error, warnings: r.warnings };
}

const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: true,
  content: [{ type: 'text', text }],
  ...(warnings && warnings.length ? { warnings } : {}),
});

const fail = (text: string, error?: string): ToolCallResult => ({
  id: uuidv4(),
  success: false,
  content: [{ type: 'text', text }],
  error: error ?? text,
});

/** 把"行/列表/散列"的查询结果排成可读文本 */
function bulletList(items: Array<Record<string, unknown>>, keys: string[]): string {
  return items.map((it) => '  · ' + keys.map((k) => `${k}=${it[k]}`).join('  ')).join('\n');
}

// ============================================================================
// 1. 表详情
// ============================================================================

export const getSheetInfoDefinition: ToolDefinition = {
  name: 'wps_excel_get_sheet_info',
  description:
    '查工作表的实际使用范围与规模：用到哪一块、多少行多少列、有没有隐藏行列、有几个图表。\n\n' +
    '什么时候用：动一张表之前先看清它的规模（比 get_sheet_list 更深一层 —— 那个只给表名和可见性）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
    },
  },
};

export const getSheetInfoHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<Record<string, unknown>>('getSheetInfo', { sheet: args.sheet });
  if (!r.ok) return fail(`读取工作表详情失败: ${r.error}`, r.error);
  const d = r.data || {};
  return ok(
    `工作表「${d.sheet}」\n` +
      `  已用区域 : ${d.usedRange}\n` +
      `  规模     : ${d.rows} 行 × ${d.columns} 列\n` +
      `  起始位置 : 第 ${d.firstRow} 行 第 ${d.firstColumn} 列\n` +
      `  隐藏行列 : 行 ${d.hiddenRows} / 列 ${d.hiddenColumns}\n` +
      `  图表数量 : ${d.charts}\n` +
      `  有数据   : ${d.hasData ? '是' : '否（空表）'}`
  );
};

// ============================================================================
// 2. 行高列宽自适应
// ============================================================================

export const autoFitDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fit',
  description:
    '按内容自动调整行高/列宽（等同 Excel 里双击列边界）。\n\n' +
    '什么时候用：中文表头经常把列撑爆、或者改完内容后宽度对不上 —— 这是表格排版最高频的重复劳动。\n' +
    '不填 dimension 就是行和列都做。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '要自适应的区域，如 A1:H50；不填则整张表的已用区域' },
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
      dimension: {
        type: 'string',
        enum: ['column', 'row', 'both'],
        description: '只调列宽 / 只调行高 / 两个都调（默认 both）',
      },
    },
  },
};

export const autoFitHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ sheet: string; dimension: string; applied: string[]; note?: string }>(
    'autoFit',
    { range: args.range, sheet: args.sheet, dimension: args.dimension }
  );
  if (!r.ok) return fail(`自适应失败: ${r.error}`, r.error);
  const applied = (r.data?.applied || []).join(' + ') || '（无）';
  const text =
    `已按内容自适应「${r.data?.sheet}」：${applied}\n` +
    (r.data?.note ? `提示：${r.data.note}` : '');
  return ok(text, r.warnings);
};

// ============================================================================
// 3. 自动换行
// ============================================================================

export const setWrapTextDefinition: ToolDefinition = {
  name: 'wps_excel_set_wrap_text',
  description:
    '设置/取消单元格自动换行。表格里最常碰的排版动作。\n\n' +
    '什么时候用：文字被截断、或者想让长文本在格子里折行显示（不影响实际内容，只影响显示）。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: { type: 'string', description: '区域，如 A1:F20；不填则整张表的已用区域' },
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
      enabled: { type: 'boolean', description: 'true 开自动换行，false 关（默认 true）' },
    },
  },
};

export const setWrapTextHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ range: string; wrapText: boolean }>('setWrapText', {
    range: args.range,
    sheet: args.sheet,
    enabled: args.enabled,
  });
  if (!r.ok) return fail(`设置自动换行失败: ${r.error}`, r.error);
  return ok(
    `${r.data?.wrapText ? '已开启' : '已关闭'}自动换行：${r.data?.range}`,
    r.warnings
  );
};

// ============================================================================
// 4. 表内查找（只读）
// ============================================================================

export const findInSheetDefinition: ToolDefinition = {
  name: 'wps_excel_find_in_sheet',
  description:
    '在工作表里找值，返回所有命中的单元格地址。**只读，不改任何内容。**\n\n' +
    '什么时候用：用户问"这个数在哪一行""这个编号出现在哪些地方"。\n' +
    '与 wps_excel_find_replace 的区别：那个是【改】，这个是【只找不改】。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      value: { type: 'string', description: '要找的内容' },
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
      range: { type: 'string', description: '限定查找范围，不填则整张表的已用区域' },
      matchCase: { type: 'boolean', description: 'true 区分大小写（默认 false）' },
      wholeCell: { type: 'boolean', description: 'true 要求整个单元格内容完全相同（默认 false，即包含即命中）' },
      maxResults: { type: 'number', description: '最多返回几个，默认不限' },
    },
    required: ['value'],
  },
};

export const findInSheetHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{
    keyword: string;
    count: number;
    matches: Array<Record<string, unknown>>;
    scope: string;
  }>('findInSheet', {
    value: args.value,
    sheet: args.sheet,
    range: args.range,
    matchCase: args.matchCase,
    wholeCell: args.wholeCell,
    maxResults: args.maxResults,
  });
  if (!r.ok) return fail(`查找失败: ${r.error}`, r.error);
  const n = r.data?.count ?? 0;
  if (n === 0) return ok(`在 ${r.data?.scope} 里没找到「${args.value}」`);
  const head = `在 ${r.data?.scope} 里找到 ${n} 处「${args.keyword}」：\n`;
  return ok(head + bulletList(r.data?.matches || [], ['address', 'value', 'formula']), r.warnings);
};

// ============================================================================
// 5. 删除图表
// ============================================================================

export const deleteChartDefinition: ToolDefinition = {
  name: 'wps_excel_delete_chart',
  description:
    '删除工作表上的图表。\n\n' +
    '什么时候用：图表建错了要重来。以前只有创建与更新，没有删除 —— COM 写入没有撤销，删不掉就只能留着。\n' +
    '不指定 index/name 时删最后一个（相当于"撤销刚插入的图表"）。\n' +
    '先用 wps_excel_get_sheet_info 看有几个图表，或直接报错信息里会列出图表名。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      index: { type: 'number', description: '图表序号（1 起）。与 name 二选一' },
      name: { type: 'string', description: '图表对象名。与 index 二选一' },
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
    },
  },
};

export const deleteChartHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ deleted: string[]; remaining: number; note?: string }>('deleteChart', {
    index: args.index,
    name: args.name,
    sheet: args.sheet,
  });
  if (!r.ok) return fail(`删除图表失败: ${r.error}`, r.error);
  const warnings = [...(r.warnings || [])];
  if (r.data?.note) warnings.push({ step: 'deleteChart', message: String(r.data.note) });
  return ok(
    `已删除图表：${(r.data?.deleted || []).join(', ')}。该表还剩 ${r.data?.remaining} 个图表。`,
    warnings
  );
};

// ============================================================================
// 6-8. 透视表三件
// ============================================================================

export const getPivotTablesDefinition: ToolDefinition = {
  name: 'wps_excel_get_pivot_tables',
  description:
    '列出工作簿/工作表里的所有透视表：名字、所在表、占的区域、各有几行/列/值字段。\n\n' +
    '什么时候用：刷新或清理之前先看清有哪些 —— 以前只能建不能查。\n' +
    '本机限制：WPS 12.1 的工作簿级 PivotTables() 不通（实测），已改为逐个工作表扫描，结果等价，' +
    '但会在 warnings 里说明这一点。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '只看某张表；不填则整个工作簿' },
    },
  },
};

export const getPivotTablesHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ count: number; pivotTables: Array<Record<string, unknown>>; scope: string }>(
    'getPivotTables',
    { sheet: args.sheet }
  );
  if (!r.ok) return fail(`读取透视表失败: ${r.error}`, r.error);
  if (!r.data?.count) return ok(`${r.data?.scope}里没有透视表。`);
  return ok(
    `${r.data.scope}里共 ${r.data.count} 个透视表：\n` +
      bulletList(r.data.pivotTables, ['name', 'sheet', 'tableRange1', 'rowFields', 'columnFields', 'dataFields']),
    r.warnings
  );
};

export const refreshPivotTablesDefinition: ToolDefinition = {
  name: 'wps_excel_refresh_pivot_tables',
  description:
    '刷新透视表（源数据改了之后重新统计）。\n\n' +
    '什么时候用：改了原始数据但透视表数字没变 —— 以前没有刷新出口，只能删了重建。\n' +
    '不填 name 则刷新全部。个别刷不了会走 warnings，不整次失败。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: { type: 'string', description: '透视表名；不填则刷新全部' },
    },
  },
};

export const refreshPivotTablesHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ scope: string; refreshed: string[]; failedCount: number }>(
    'refreshPivotTables',
    { name: args.name }
  );
  if (!r.ok) return fail(`刷新透视表失败: ${r.error}`, r.error);
  return ok(
    `已刷新 ${r.data?.refreshed?.length ?? 0} 个透视表（${r.data?.scope}）：` +
      `${(r.data?.refreshed || []).join(', ') || '无'}`,
    r.warnings
  );
};

export const clearPivotTableDefinition: ToolDefinition = {
  name: 'wps_excel_clear_pivot_table',
  description:
    '清空指定透视表的数据与字段布局，**保留透视表本身的位置与格式**。\n\n' +
    '什么时候用：透视表结构建错了想重来，但不想把它删掉重建（删了要重新指定源区域）。\n' +
    '与删除的区别：这个是"清空重填"，不是"删掉"。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: { type: 'string', description: '透视表名（先用 wps_excel_get_pivot_tables 查）' },
    },
    required: ['name'],
  },
};

export const clearPivotTableHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ cleared: string; note?: string }>('clearPivotTable', { name: args.name });
  if (!r.ok) return fail(`清空透视表失败: ${r.error}`, r.error);
  return ok(`已清空透视表「${r.data?.cleared}」。${r.data?.note || ''}`, r.warnings);
};

// ============================================================================
// 9-10. 页面设置与打印标题
// ============================================================================

export const setSheetPageSetupDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_page_setup',
  description:
    '设置打印页面：方向、缩放比例、纸张大小、按宽高缩放到几页。\n\n' +
    '什么时候用：要打印/导出 PDF 之前。以前只有 set_print_area（只管打印区域），\n' +
    '"横向打印""一页宽""A4 纸"这些全做不了。逐属性容错：不支持的那个走 warnings，不整次失败。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
      orientation: {
        type: 'string',
        enum: ['portrait', 'landscape'],
        description: 'portrait 纵向 / landscape 横向（横向是宽表的默认选择）',
      },
      zoom: { type: 'number', description: '缩放百分比，如 80' },
      paperSize: { type: 'number', description: '纸张枚举值，如 9=A4、1=Letter' },
      zoomFitWidth: { type: 'number', description: '按宽度缩放到几页（宽表常用 1）' },
      zoomFitHeight: { type: 'number', description: '按高度缩放到几页' },
    },
  },
};

export const setSheetPageSetupHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<{ sheet: string; applied: string[]; failed: string[] }>('setSheetPageSetup', {
    sheet: args.sheet,
    orientation: args.orientation,
    zoom: args.zoom,
    paperSize: args.paperSize,
    zoomFitWidth: args.zoomFitWidth,
    zoomFitHeight: args.zoomFitHeight,
  });
  if (!r.ok) return fail(`设置页面失败: ${r.error}`, r.error);
  return ok(
    `已设置「${r.data?.sheet}」页面：${(r.data?.applied || []).join('、') || '（无）'}` +
      (r.data?.failed?.length ? `\n未生效：${r.data.failed.join('/')}` : ''),
    r.warnings
  );
};

export const setSheetPrintTitlesDefinition: ToolDefinition = {
  name: 'wps_excel_set_sheet_print_titles',
  description:
    '设置每页重复打印的行/列（等同"打印标题行"），或清除。\n\n' +
    '什么时候用：一张表打印出来好几页，第二页开始没表头 —— 这是多页表最常见的诉求。\n' +
    '例：rows="$1:$1" 表示第 1 行每页重复；columns="$A:$B" 表示 A、B 两列每页重复。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名，不填则当前活动表' },
      rows: { type: 'string', description: '重复行，如 "$1:$1" 或 "$1:$2"' },
      columns: { type: 'string', description: '重复列，如 "$A:$B"' },
      clear: { type: 'boolean', description: 'true 清除打印标题' },
    },
  },
};

export const setSheetPrintTitlesHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const r = await call<Record<string, unknown>>('setSheetPrintTitles', {
    sheet: args.sheet,
    rows: args.rows,
    columns: args.columns,
    clear: args.clear,
  });
  if (!r.ok) return fail(`设置打印标题失败: ${r.error}`, r.error);
  const d = r.data || {};
  const what = d.printTitleRows
    ? `行 ${d.printTitleRows}`
    : d.printTitleColumns
      ? `列 ${d.printTitleColumns}`
      : '（已清除）';
  return ok(`「${d.sheet}」每页重复打印：${what}`, r.warnings);
};

export const sheetGapsTools: RegisteredTool[] = [
  { definition: getSheetInfoDefinition, handler: getSheetInfoHandler },
  { definition: autoFitDefinition, handler: autoFitHandler },
  { definition: setWrapTextDefinition, handler: setWrapTextHandler },
  { definition: findInSheetDefinition, handler: findInSheetHandler },
  { definition: deleteChartDefinition, handler: deleteChartHandler },
  { definition: getPivotTablesDefinition, handler: getPivotTablesHandler },
  { definition: refreshPivotTablesDefinition, handler: refreshPivotTablesHandler },
  { definition: clearPivotTableDefinition, handler: clearPivotTableHandler },
  { definition: setSheetPageSetupDefinition, handler: setSheetPageSetupHandler },
  { definition: setSheetPrintTitlesDefinition, handler: setSheetPrintTitlesHandler },
];
