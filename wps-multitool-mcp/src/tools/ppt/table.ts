/**
 * Input: PPT 表格操作参数
 * Output: 表格操作结果
 * Pos: PPT 表格工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT表格Tools - 表格管理模块
 * 处理表格的插入、单元格读写、样式设置等操作
 *
 * 包含：
 * - wps_ppt_insert_table: 在幻灯片中插入表格
 * - wps_ppt_set_table_cell: 设置表格单元格文本
 * - wps_ppt_get_table_cell: 获取表格单元格文本
 * - wps_ppt_set_table_style: 设置表格整体样式
 * - wps_ppt_set_table_cell_style: 设置表格单元格样式
 * - wps_ppt_set_table_row_style: 设置表格行样式
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText } from './helpers';

/**
 * 在幻灯片中插入表格
 * 支持指定行列数和位置
 */
export const insertPptTableDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_table',
  description: `在幻灯片中插入表格。

支持指定：
- 表格行列数
- 表格位置（左上角坐标）

使用场景：
- "在第2页插入一个3行4列的表格"
- "添加一个表格到当前页"
- "插入数据表格"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      rows: {
        type: 'number',
        description: '表格行数',
      },
      cols: {
        type: 'number',
        description: '表格列数',
      },
      left: {
        type: 'number',
        description: '表格左上角X坐标（磅），可选',
      },
      top: {
        type: 'number',
        description: '表格左上角Y坐标（磅），可选',
      },
    },
    required: ['slideIndex', 'rows', 'cols'],
  },
};

export const insertPptTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, rows, cols, left, top } = args as {
    slideIndex: number;
    rows: number;
    cols: number;
    left?: number;
    top?: number;
  };

  const r = await call<{
    // PS 侧回 name 与本页表格序号 tableIndex（第 3 轮补的 tableIndex）。
    // 历史 TS 只读 tableIndex，而 PS 当时没回这个键 → 永远打印"表格索引: undefined"。
    name?: string;
    tableIndex?: number;
  }>('insertPptTable', { slideIndex, rows, cols, left, top });
  if (!r.ok) return fail(errText('插入表格失败', r), r.error);
  const idx = r.data?.tableIndex ?? '?';
  return ok(
    `表格插入成功！\n幻灯片: 第 ${slideIndex} 页\n规格: ${rows} 行 × ${cols} 列\n表格索引: ${idx}${r.data?.name ? `\n对象名: ${r.data.name}` : ''}`,
    r.warnings
  );
};

/**
 * 设置表格单元格文本
 * 修改指定单元格的内容
 */
export const setPptTableCellDefinition: ToolDefinition = {
  name: 'wps_ppt_set_table_cell',
  description: `设置PPT表格中指定单元格的文本内容。

使用场景：
- "把表格第1行第2列的内容改成'销售额'"
- "设置单元格文本"
- "填写表格内容"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      tableIndex: {
        type: 'number',
        description: '表格索引（从1开始）',
      },
      row: {
        type: 'number',
        description: '行号（从1开始）',
      },
      col: {
        type: 'number',
        description: '列号（从1开始）',
      },
      text: {
        type: 'string',
        description: '要设置的文本内容',
      },
    },
    required: ['slideIndex', 'tableIndex', 'row', 'col', 'text'],
  },
};

export const setPptTableCellHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, tableIndex, row, col, text } = args as {
    slideIndex: number;
    tableIndex: number;
    row: number;
    col: number;
    text: string;
  };

  const r = await call('setPptTableCell', { slideIndex, tableIndex, row, col, text });
  if (!r.ok) return fail(errText('设置单元格文本失败', r), r.error);
  return ok(
    `单元格文本设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n内容: ${text}`,
    r.warnings
  );
};

/**
 * 获取表格单元格文本
 * 读取指定单元格的内容
 */
export const getPptTableCellDefinition: ToolDefinition = {
  name: 'wps_ppt_get_table_cell',
  description: `获取PPT表格中指定单元格的文本内容。

使用场景：
- "读取表格第1行第1列的内容"
- "获取单元格文本"
- "查看表格数据"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      tableIndex: {
        type: 'number',
        description: '表格索引（从1开始）',
      },
      row: {
        type: 'number',
        description: '行号（从1开始）',
      },
      col: {
        type: 'number',
        description: '列号（从1开始）',
      },
    },
    required: ['slideIndex', 'tableIndex', 'row', 'col'],
  },
};

export const getPptTableCellHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, tableIndex, row, col } = args as {
    slideIndex: number;
    tableIndex: number;
    row: number;
    col: number;
  };

  const r = await call<{
    // PS 侧 getPptTableCell 的 Output-Json 给的是 value（单元格的值），
    // 历史 TS 读的是 text → 永远 undefined，读回形同虚设。两个键都接，value 优先。
    value?: string;
    text?: string;
  }>('getPptTableCell', { slideIndex, tableIndex, row, col });
  if (!r.ok) return fail(errText('获取单元格内容失败', r), r.error);
  const cellText = r.data?.value ?? r.data?.text ?? '';
  return ok(
    `单元格内容获取成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n内容: ${cellText}`,
    r.warnings
  );
};

/**
 * 设置表格整体样式
 * 支持边框、背景色、字体等整体样式配置
 */
export const setPptTableStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_table_style',
  description: `设置PPT表格的整体样式。

支持的样式属性（通过style对象传入）：
- borderColor: 边框颜色（如 "#000000"）
- borderWidth: 边框宽度（磅）
- backgroundColor: 背景色（如 "#FFFFFF"）
- fontName: 字体名称（如 "微软雅黑"）
- fontSize: 字体大小（磅）
- fontColor: 字体颜色（如 "#333333"）
- headerBackground: 表头行背景色
- alternateRowColor: 隔行变色颜色

使用场景：
- "设置表格为蓝色主题"
- "修改表格边框和背景"
- "美化表格样式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      tableIndex: {
        type: 'number',
        description: '表格索引（从1开始）',
      },
      style: {
        type: 'object',
        description: '样式配置对象，包含borderColor、borderWidth、backgroundColor、fontName、fontSize、fontColor、headerBackground、alternateRowColor等属性',
      },
    },
    required: ['slideIndex', 'tableIndex', 'style'],
  },
};

export const setPptTableStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, tableIndex, style } = args as {
    slideIndex: number;
    tableIndex: number;
    style: Record<string, unknown>;
  };

  const r = await call('setPptTableStyle', { slideIndex, tableIndex, style });
  if (!r.ok) return fail(errText('设置表格样式失败', r), r.error);
  const styleKeys = Object.keys(style).join('、');
  return ok(
    `表格样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n已设置属性: ${styleKeys}`,
    r.warnings
  );
};

/**
 * 设置表格单元格样式
 * 修改指定单元格的样式（字体、颜色、背景等）
 */
export const setPptTableCellStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_table_cell_style',
  description: `设置PPT表格中指定单元格的样式。

支持的样式属性（通过style对象传入）：
- backgroundColor: 单元格背景色
- fontName: 字体名称
- fontSize: 字体大小（磅）
- fontColor: 字体颜色
- bold: 是否加粗（boolean）
- italic: 是否斜体（boolean）
- alignment: 文本对齐方式（left/center/right）
- verticalAlignment: 垂直对齐方式（top/middle/bottom）

使用场景：
- "把表格第1行第1列的背景设为蓝色"
- "加粗标题单元格"
- "设置单元格居中对齐"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      tableIndex: {
        type: 'number',
        description: '表格索引（从1开始）',
      },
      row: {
        type: 'number',
        description: '行号（从1开始）',
      },
      col: {
        type: 'number',
        description: '列号（从1开始）',
      },
      style: {
        type: 'object',
        description: '样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、verticalAlignment等属性',
      },
    },
    required: ['slideIndex', 'tableIndex', 'row', 'col', 'style'],
  },
};

export const setPptTableCellStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, tableIndex, row, col, style } = args as {
    slideIndex: number;
    tableIndex: number;
    row: number;
    col: number;
    style: Record<string, unknown>;
  };

  const r = await call('setPptTableCellStyle', { slideIndex, tableIndex, row, col, style });
  if (!r.ok) return fail(errText('设置单元格样式失败', r), r.error);
  const styleKeys = Object.keys(style).join('、');
  return ok(
    `单元格样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n位置: 第 ${row} 行第 ${col} 列\n已设置属性: ${styleKeys}`,
    r.warnings
  );
};

/**
 * 设置表格行样式
 * 批量修改指定行的样式
 */
export const setPptTableRowStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_table_row_style',
  description: `设置PPT表格中指定行的样式。

支持的样式属性（通过style对象传入）：
- backgroundColor: 行背景色
- fontName: 字体名称
- fontSize: 字体大小（磅）
- fontColor: 字体颜色
- bold: 是否加粗（boolean）
- italic: 是否斜体（boolean）
- alignment: 文本对齐方式（left/center/right）
- height: 行高（磅）

使用场景：
- "把表头行设为蓝色背景白色字"
- "加粗第一行"
- "设置表格行高"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      tableIndex: {
        type: 'number',
        description: '表格索引（从1开始）',
      },
      row: {
        type: 'number',
        description: '行号（从1开始）',
      },
      style: {
        type: 'object',
        description: '样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、height等属性',
      },
    },
    required: ['slideIndex', 'tableIndex', 'row', 'style'],
  },
};

export const setPptTableRowStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, tableIndex, row, style } = args as {
    slideIndex: number;
    tableIndex: number;
    row: number;
    style: Record<string, unknown>;
  };

  const r = await call('setPptTableRowStyle', { slideIndex, tableIndex, row, style });
  if (!r.ok) return fail(errText('设置行样式失败', r), r.error);
  const styleKeys = Object.keys(style).join('、');
  return ok(
    `表格行样式设置成功！\n幻灯片: 第 ${slideIndex} 页\n表格: 第 ${tableIndex} 个\n行: 第 ${row} 行\n已设置属性: ${styleKeys}`,
    r.warnings
  );
};

/**
 * 导出所有表格相关的Tools
 */
export const tableTools: RegisteredTool[] = [
  { definition: insertPptTableDefinition, handler: insertPptTableHandler },
  { definition: setPptTableCellDefinition, handler: setPptTableCellHandler },
  { definition: getPptTableCellDefinition, handler: getPptTableCellHandler },
  { definition: setPptTableStyleDefinition, handler: setPptTableStyleHandler },
  { definition: setPptTableCellStyleDefinition, handler: setPptTableCellStyleHandler },
  { definition: setPptTableRowStyleDefinition, handler: setPptTableRowStyleHandler },
];

export default tableTools;
