/**
 * Input: 工作表高级操作工具参数
 * Output: 工作表高级操作结果
 * Pos: Excel 工作表高级操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * - wps_excel_delete_row: 删除指定行
 * - wps_excel_set_named_range: 设置命名范围
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';

import { call, ok, fail } from './helpers';

export const getSelectionDefinition: ToolDefinition = {
  name: 'wps_excel_get_selection',
  description: '获取当前Excel中选中区域的信息，包括范围地址、行列数等。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const getSelectionHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      address: string;
      rowCount: number;
      columnCount: number;
      sheet: string;
    }>('getSelection', {});

    if (!response.ok || !response.data) {
      return fail(`获取选中区域失败: ${response.error}`, response.error);
    }

    const { address, rowCount, columnCount, sheet } = response.data;

    return ok(`当前选中区域信息：\n工作表: ${sheet}\n范围: ${address}\n行数: ${rowCount}\n列数: ${columnCount}`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`获取选中区域出错: ${errMsg}`, errMsg);
  }
};

/**
 * 删除指定行
 */
export const deleteRowDefinition: ToolDefinition = {
  name: 'wps_excel_delete_row',
  description:
    '[DEPRECATED·改用 wps_excel_delete_rows] 删除指定行。可指定起始行号和删除行数。\n\n' +
    '什么时候用：已被 wps_excel_delete_rows 取代（同一个 deleteRows 动作），' +
    '旧参数名是 row，新工具用 startRow —— 别两个混着用。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '要删除的起始行号（从1开始）',
      },
      count: {
        type: 'number',
        description: '要删除的行数，默认1',
      },
    },
    required: ['row'],
  },
};

export const deleteRowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, count = 1 } = args as { row: number; count?: number };

  try {
    const response = await call<{
      deletedRows: number;
    }>('deleteRows', { row, count });

    if (!response.ok) {
      return fail(`删除行失败: ${response.error}`, response.error);
    }

    return ok(`已成功删除第${row}行起共${count}行`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`删除行出错: ${errMsg}`, errMsg);
  }
};

export const freezePanesDefinition: ToolDefinition = {
  name: 'wps_excel_freeze_panes',
  description: '冻结/取消冻结窗格。可指定冻结的行和列位置。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      row: {
        type: 'number',
        description: '冻结到第几行（从1开始），不填则不冻结行',
      },
      column: {
        type: 'number',
        description: '冻结到第几列（从1开始），不填则不冻结列',
      },
      freeze: {
        type: 'boolean',
        description: '是否冻结，默认true。设为false则取消冻结',
      },
    },
  },
};

export const freezePanesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { row, column, freeze = true } = args as {
    row?: number;
    column?: number;
    freeze?: boolean;
  };

  try {
    const response = await call<{
      frozen: boolean;
    }>('freezePanes', { row, column, freeze });

    if (!response.ok) {
      return fail(`冻结窗格操作失败: ${response.error}`, response.error);
    }

    const action = freeze ? '冻结' : '取消冻结';
    return ok(`窗格${action}成功！${row ? `\n冻结行: ${row}` : ''}${column ? `\n冻结列: ${column}` : ''}`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`冻结窗格操作出错: ${errMsg}`, errMsg);
  }
};

export const autoFillDefinition: ToolDefinition = {
  name: 'wps_excel_auto_fill',
  description:
    '[DEPRECATED·改用 wps_excel_fill_series] 自动填充单元格区域。根据源区域的数据模式自动填充到目标区域。\n\n' +
    '什么时候用：已被 wps_excel_fill_series 取代（同一个 fillSeries 动作），' +
    '旧参数是 sourceRange/targetRange，新工具用 range/direction/step/type —— 别两个混着用。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sourceRange: {
        type: 'string',
        description: '源数据区域，如 "A1:A5"',
      },
      targetRange: {
        type: 'string',
        description: '目标填充区域，如 "A1:A20"',
      },
    },
    required: ['sourceRange', 'targetRange'],
  },
};

export const autoFillHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { sourceRange, targetRange } = args as {
    sourceRange: string;
    targetRange: string;
  };

  try {
    const response = await call<{
      filled: boolean;
    }>('fillSeries', { sourceRange, targetRange });

    if (!response.ok) {
      return fail(`自动填充失败: ${response.error}`, response.error);
    }

    return ok(`自动填充成功！\n源区域: ${sourceRange}\n目标区域: ${targetRange}`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`自动填充出错: ${errMsg}`, errMsg);
  }
};

/**
 * 设置命名范围
 */
export const setNamedRangeDefinition: ToolDefinition = {
  name: 'wps_excel_set_named_range',
  description: '设置命名范围。为指定单元格区域创建或更新命名范围。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: '命名范围的名称',
      },
      range: {
        type: 'string',
        description: '单元格区域，如 "A1:D10"',
      },
    },
    required: ['name', 'range'],
  },
};

export const setNamedRangeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { name, range } = args as {
    name: string;
    range: string;
  };

  try {
    const response = await call<{
      name: string;
      range: string;
    }>('createNamedRange', { name, range });

    if (!response.ok) {
      return fail(`设置命名范围失败: ${response.error}`, response.error);
    }

    return ok(`命名范围设置成功！\n名称: ${name}\n范围: ${range}`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`设置命名范围出错: ${errMsg}`, errMsg);
  }
};

export const hideColumnDefinition: ToolDefinition = {
  name: 'wps_excel_hide_column',
  description: '隐藏或显示指定列。可指定起始列号、列数和隐藏/显示状态。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      column: {
        type: 'number',
        description: '起始列号（从1开始）',
      },
      count: {
        type: 'number',
        description: '列数，默认1',
      },
      hide: {
        type: 'boolean',
        description: '是否隐藏，true为隐藏，false为显示',
      },
    },
    required: ['column', 'count', 'hide'],
  },
};

export const hideColumnHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { column, count = 1, hide } = args as {
    column: number;
    count: number;
    hide: boolean;
  };

  try {
    const response = await call<{
      column: number;
      count: number;
      hidden: boolean;
    }>('hideColumns', { column, count, hide });

    if (!response.ok) {
      return fail(`${hide ? '隐藏' : '显示'}列失败: ${response.error}`, response.error);
    }

    const action = hide ? '隐藏' : '显示';
    return ok(`已成功${action}第${column}列起共${count}列`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`隐藏/显示列出错: ${errMsg}`, errMsg);
  }
};

export const autoSumDefinition: ToolDefinition = {
  name: 'wps_excel_auto_sum',
  description:
    '把一段区域求和的结果写进一个单元格（等价于在该格写 =SUM(区域)）。\n\n' +
    '与 wps_excel_set_formula 的关系：**同一个底层动作（setFormula）**，本工具是它的【求和子集】——' +
    '只拼 =SUM(range) 这一种公式。要写别的公式（AVERAGE/IF/VLOOKUP/自拼表达式）用 wps_excel_set_formula；' +
    '要算一个值但不落到单元格里，用 wps_excel_evaluate_formula。别两个混着用，它们改的是同一批单元格。\n\n' +
    '写进去的是【公式】不是死数字：源区域改了，目标格会跟着变（这是特性，不是差异）。' +
    '要死数字就先算再写值 —— 用 evaluate_formula 取结果，或 set_formula 里写死。\n\n' +
    '使用场景：\n- "把 A1:A10 的和写在 A11"\n- "给这个月的数据加一行合计"',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range: {
        type: 'string',
        description: '要求和的【源】数据范围，如 "A1:A10" 或 "B2:F2"',
      },
      targetCell: {
        type: 'string',
        description: '求和结果写入的目标单元格，如 "A11" 或 "G2"',
      },
    },
    required: ['range', 'targetCell'],
  },
};

export const autoSumHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { range, targetCell } = args as {
    range: string;
    targetCell: string;
  };

  // 这两个参数拼的是公式文本，空值会拼出 =SUM() —— 那是语法错，写进去只会得到 #NAME?。
  // 在 TS 侧先拦住，比让 WPS 抛一句看不懂的错更省事。
  if (!range || !range.trim() || !targetCell || !targetCell.trim()) {
    return fail('[勿重试·先改参数] range 与 targetCell 都不能为空', 'range/targetCell 为空');
  }

  try {
    const formula = `=SUM(${range.trim()})`;
    const response = await call<{
      range: string;
      formula: string;
      value: unknown;
    }>('setFormula', { range: targetCell.trim(), formula });

    if (!response.ok) {
      return fail(`自动求和失败: ${response.error}`, response.error);
    }

    // 读回值：公式写进去了不等于算出了数（引用空区域得 0，引用坏区域得 #REF!）。
    // 有值就报，没读到就明说 —— 不能让用户以为一定成功了。
    const v = response.data?.value;
    const shown = v === undefined || v === null || v === '' ? '（没读到结果值，公式已写入，请用 wps_excel_get_cell_value 核对）' : String(v);
    return ok(`自动求和完成！\n求和范围: ${range}\n写入: ${targetCell} = ${formula}\n当前值: ${shown}`, response.warnings);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`自动求和出错: ${errMsg}`, errMsg);
  }
};

// ============================================================================
// Excel 表格（ListObjects）与切片器（2026-10-26 对标 excel-mcp-server）：
// "Excel 表格"是结构化引用区（Sort/筛选/结构化引用跟随），与"插入网格表格"不同物。
// 真机探针（2026-10-26）：ListObjects.Add / SlicerCaches.Add / Slicers.Add 全部可用。
// ============================================================================

export const createExcelTableDefinition: ToolDefinition = {
  name: 'wps_excel_create_table',
  description: `把一块区域创建为"Excel 表格"（结构化引用区）。

与 insert_table 的网格表格不同：创建后获得筛选按钮、结构化引用、随数据扩展的汇总能力。
range 需包含表头行。可选 tableName 自定义表名（默认自动命名）。`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名，默认当前表' },
      range: { type: 'string', description: '区域，如 A1:C10（含表头行）' },
      tableName: { type: 'string', description: '表名，默认自动命名' },
    },
    required: ['range'],
  },
};

export const createExcelTableHandler: ToolHandler = async (args) => {
  const { sheet, range, tableName } = args as { sheet?: string; range?: string; tableName?: string };
  try {
    const response = await call<{ name: string; range: string }>('createExcelTable', { sheet, range, tableName });
    if (response.ok) {
      const d = response.data as { name?: string; range?: string } | undefined;
      return ok(`已创建表格「${d?.name ?? tableName ?? ''}」（${d?.range ?? range}）。之后可 addSlicer 加切片器，或对表内数据排序/筛选。`, response.warnings);
    }
    return fail(`创建表格失败: ${response.error}`, response.error);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`创建表格出错: ${errMsg}`, errMsg);
  }
};

export const addSlicerDefinition: ToolDefinition = {
  name: 'wps_excel_add_slicer',
  description: `给已创建的 Excel 表格添加切片器（按某列筛选的按钮面板）。

需要先 createExcelTable。field 是列名（表头文字）。`,
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      sheet: { type: 'string', description: '工作表名，默认当前表' },
      tableName: { type: 'string', description: '表名；不传默认当前表的第一张表格' },
      field: { type: 'string', description: '列名（表头文字）' },
      name: { type: 'string', description: '切片器名称，默认"切片_列名"' },
    },
    required: ['field'],
  },
};

export const addSlicerHandler: ToolHandler = async (args) => {
  const { sheet, tableName, field, name } = args as {
    sheet?: string;
    tableName?: string;
    field?: string;
    name?: string;
  };
  try {
    const response = await call<{ slicer: string; table: string; field: string }>('addSlicer', { sheet, tableName, field, name });
    if (response.ok) {
      const d = response.data as { slicer?: string; table?: string; field?: string } | undefined;
      return ok(`已添加切片器「${d?.slicer ?? name ?? field}」（表 ${d?.table ?? tableName ?? ''} · 列 ${d?.field ?? field}）`, response.warnings);
    }
    return fail(`添加切片器失败: ${response.error}`, response.error);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`添加切片器出错: ${errMsg}`, errMsg);
  }
};

export const sheetAdvancedTools: RegisteredTool[] = [
  { definition: getSelectionDefinition, handler: getSelectionHandler },
  { definition: deleteRowDefinition, handler: deleteRowHandler },
  { definition: freezePanesDefinition, handler: freezePanesHandler },
  { definition: autoFillDefinition, handler: autoFillHandler },
  { definition: setNamedRangeDefinition, handler: setNamedRangeHandler },
  { definition: hideColumnDefinition, handler: hideColumnHandler },
  { definition: autoSumDefinition, handler: autoSumHandler },
  { definition: createExcelTableDefinition, handler: createExcelTableHandler },
  { definition: addSlicerDefinition, handler: addSlicerHandler },
];

export default sheetAdvancedTools;
