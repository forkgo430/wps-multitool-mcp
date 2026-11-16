/**
 * Input: PPT 形状基础操作参数（添加、删除、获取、位置、透明度、文字、标题、对齐、分布、组合）
 * Output: 形状操作结果
 * Pos: PPT 形状基础工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 处理形状的添加、删除、获取、定位、文字、对齐分布和组合操作
 * - wps_ppt_delete_shape: 删除形状
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText, shapeIndexByName } from './helpers';

// ==================== 1. 添加形状 ====================

export const addShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_add_shape',
  description: `在幻灯片中添加形状。

支持的形状类型：
- rectangle: 矩形
- oval: 椭圆
- triangle: 三角形
- diamond: 菱形
- pentagon: 五边形
- hexagon: 六边形
- arrow: 箭头
- star: 星形
- heart: 心形
- cloud: 云形

使用场景：
- "在第1页添加一个矩形"
- "插入一个蓝色的圆形"
- "添加一个带文字的箭头"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      type: {
        type: 'string',
        description: '形状类型',
        enum: ['rectangle', 'oval', 'triangle', 'diamond', 'pentagon', 'hexagon', 'arrow', 'star', 'heart', 'cloud'],
      },
      left: {
        type: 'number',
        description: '左边距（磅），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（磅），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（磅），默认100',
      },
      height: {
        type: 'number',
        description: '高度（磅），默认100',
      },
      text: {
        type: 'string',
        description: '形状内的文本',
      },
      fillColor: {
        type: 'string',
        description: '填充颜色，十六进制如 #FF0000',
      },
    },
    required: [],
  },
};

export const addShapeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, type, left, top, width, height, text, fillColor } = args as {
    slideIndex?: number;
    type?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
    text?: string;
    fillColor?: string;
  };

  const r = await call<{
    name?: string;
    slideIndex?: number;
  }>(
    'addShape',
    {
      slideIndex: slideIndex || 1,
      type: type || 'rectangle',
      // 坐标用 ?? 不用 ||：0 是合法坐标（全页底/竖条都从 0 起），|| 会把它吞成 100
      // （deckkit 引擎的全页渐变底/标题色带/章节引文竖条全踩——审计三.1）
      left: left ?? 100,
      top: top ?? 100,
      width: width ?? 100,
      height: height ?? 100,
      text,
      fillColor,
    }
  );
  if (!r.ok) return fail(errText('添加形状失败', r), r.error);
  const shapeNameMap: Record<string, string> = {
    rectangle: '矩形', oval: '椭圆', triangle: '三角形', diamond: '菱形',
    pentagon: '五边形', hexagon: '六边形', arrow: '箭头', star: '星形',
    heart: '心形', cloud: '云形',
  };
  const typeName = shapeNameMap[type || 'rectangle'] || type;

  // 回包逐项拼：数据缺失就不打印该行（体验修复轮 A——此前"第 undefined 页、名称 undefined"就是这么来的）
  const parts: string[] = [];
  if (r.data?.slideIndex !== undefined) parts.push(`幻灯片: 第 ${r.data.slideIndex} 页`);
  parts.push(`形状类型: ${typeName}`);
  if (r.data?.name) {
    parts.push(`形状名称: ${r.data.name}`);
    const idx = await shapeIndexByName(r.data.slideIndex ?? slideIndex ?? 1, r.data.name);
    if (idx !== undefined) parts.push(`形状索引: ${idx}（后续 set_shape_text / set_animation 用）`);
  }
  if (text) parts.push(`文本: ${text}`);
  if (fillColor) parts.push(`填充色: ${fillColor}`);

  return ok(`形状添加成功！\n${parts.join('\n')}`, r.warnings);
};

// ==================== 2. 添加文本框 ====================

export const addTextboxDefinition: ToolDefinition = {
  name: 'wps_ppt_add_textbox',
  description: `在幻灯片中添加文本框。

使用场景：
- "在第1页添加一个文本框"
- "插入一个写着标题的文本框"
- "添加文本框并设置字号"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      left: {
        type: 'number',
        description: '左边距（磅），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（磅），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（磅），默认200',
      },
      height: {
        type: 'number',
        description: '高度（磅），默认50',
      },
      text: {
        type: 'string',
        description: '文本框内容',
      },
      fontSize: {
        type: 'number',
        description: '字号大小',
      },
      fontName: {
        type: 'string',
        description: '字体名称，如 "微软雅黑"',
      },
    },
    required: [],
  },
};

export const addTextboxHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, left, top, width, height, text, fontSize, fontName } = args as {
    slideIndex?: number;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
    text?: string;
    fontSize?: number;
    fontName?: string;
  };

  const r = await call<{
    name?: string;
    slideIndex?: number;
  }>(
    'addTextBox',
    {
      slideIndex: slideIndex || 1,
      // 坐标用 ?? 不用 ||：0 是合法坐标（同 add_shape，审计三.1）
      left: left ?? 100,
      top: top ?? 100,
      width: width ?? 200,
      height: height ?? 50,
      text,
      fontSize,
      fontName,
    }
  );
  if (!r.ok) return fail(errText('添加文本框失败', r), r.error);
  // 回包逐项拼：数据缺失就不打印该行（同 add_shape，体验修复轮 A）
  const parts: string[] = [];
  if (r.data?.slideIndex !== undefined) parts.push(`幻灯片: 第 ${r.data.slideIndex} 页`);
  if (r.data?.name) {
    parts.push(`名称: ${r.data.name}`);
    const idx = await shapeIndexByName(r.data.slideIndex ?? slideIndex ?? 1, r.data.name);
    if (idx !== undefined) parts.push(`文本框索引: ${idx}（后续 set_textbox_text / set_textbox_style 用）`);
  }
  if (text) parts.push(`内容: "${text}"`);
  if (fontSize) parts.push(`字号: ${fontSize}`);
  if (fontName) parts.push(`字体: ${fontName}`);

  return ok(`文本框添加成功！\n${parts.join('\n')}`, r.warnings);
};

// ==================== 3. 删除形状 ====================

export const deleteShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_shape',
  description: `删除幻灯片中指定的形状。

使用场景：
- "删除第2页的第3个形状"
- "移除这个形状"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
};

export const deleteShapeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex } = args as {
    slideIndex: number;
    shapeIndex: number;
  };

  const r = await call('deleteShape', { slideIndex, shapeIndex });
  if (!r.ok) return fail(errText('删除形状失败', r), r.error);
  return ok(`形状删除成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`, r.warnings);
};

// ==================== 4. 获取形状列表 ====================

export const getShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_shapes',
  description: `获取幻灯片中所有形状的列表信息。

brief 模式（默认开，回包瘦身）：只回序号/名称/几何摘要；全量字段传 brief:false。

使用场景：
- "这页PPT有哪些形状"
- "列出第3页的所有元素"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      brief: {
        type: 'boolean',
        description: '回包瘦身开关（体验修复轮 B），默认 true 只回 index/name/几何摘要；false 回全量字段',
      },
    },
    required: ['slideIndex'],
  },
};

export const getShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, brief = true } = args as {
    slideIndex: number;
    brief?: boolean;
  };

  const r = await call<{
    shapes: Array<{
      index: number;
      name: string;
      type: string;
      left: number;
      top: number;
      width: number;
      height: number;
      text?: string;
    }>;
  }>('getShapes', { slideIndex });
  if (!r.ok || !r.data) return fail(errText('获取形状列表失败', r), r.error);
  const shapes = r.data.shapes;
  let output = `幻灯片第 ${slideIndex} 页的形状列表（共 ${shapes.length} 个${brief ? '，brief 摘要模式' : ''}）：\n\n`;
  shapes.forEach((s) => {
    if (brief) {
      output += `[${s.index}] ${s.name} — (${s.left}, ${s.top}) ${s.width} x ${s.height}\n`;
    } else {
      output += `[${s.index}] ${s.name} (${s.type})\n`;
      output += `    位置: (${s.left}, ${s.top}) 大小: ${s.width} x ${s.height}\n`;
    }
  });

  return ok(output, r.warnings);
};

// ==================== 5. 设置形状位置 ====================

export const setShapePositionDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_position',
  description: `设置幻灯片中指定形状的位置和大小。

使用场景：
- "把这个形状移到左上角"
- "调整形状大小"
- "设置形状位置为(100, 200)"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      left: {
        type: 'number',
        description: '左边距（磅）',
      },
      top: {
        type: 'number',
        description: '上边距（磅）',
      },
      width: {
        type: 'number',
        description: '宽度（磅），可选',
      },
      height: {
        type: 'number',
        description: '高度（磅），可选',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'left', 'top'],
  },
};

export const setShapePositionHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, left, top, width, height } = args as {
    slideIndex: number;
    shapeIndex: number;
    left: number;
    top: number;
    width?: number;
    height?: number;
  };

  const r = await call('setShapePosition', { slideIndex, shapeIndex, left, top, width, height });
  if (!r.ok) return fail(errText('设置形状位置失败', r), r.error);
  let text = `形状位置设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n位置: (${left}, ${top})`;
  if (width !== undefined) text += `\n宽度: ${width}`;
  if (height !== undefined) text += `\n高度: ${height}`;

  return ok(text, r.warnings);
};


// ==================== 10. 设置形状透明度 ====================

export const setShapeTransparencyDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_transparency',
  description: `设置幻灯片中指定形状的透明度。

使用场景：
- "把这个形状设为半透明"
- "设置形状透明度为50%"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）',
      },
      transparency: {
        type: 'number',
        description: '透明度值 (0-100)，0为完全不透明，100为完全透明',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'transparency'],
  },
};

export const setShapeTransparencyHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, transparency } = args as {
    slideIndex: number;
    shapeIndex: number;
    transparency: number;
  };

  const r = await call('setShapeTransparency', { slideIndex, shapeIndex, transparency });
  if (!r.ok) return fail(errText('设置形状透明度失败', r), r.error);
  return ok(
    `形状透明度设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n透明度: ${transparency}%`,
    r.warnings
  );
};

// ==================== 11. 设置形状文字 ====================

export const setShapeTextDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_text',
  description: `设置幻灯片中指定形状的文字内容。

使用场景：
- "把第1页的第2个形状文字改成'销售额'"
- "修改形状里的文字"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
      text: { type: 'string', description: '要设置的文字内容' },
    },
    required: ['slideIndex', 'shapeIndex', 'text'],
  },
};

export const setShapeTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, text } = args as {
    slideIndex: number;
    shapeIndex: number;
    text: string;
  };

  const r = await call<{ name: string }>('setShapeText', { slideIndex, shapeIndex, text });
  if (!r.ok || !r.data) return fail(errText('设置形状文字失败', r), r.error);
  return ok(`形状文字设置成功！\n形状: ${r.data.name}\n文字: "${text}"`, r.warnings);
};


// ==================== 13. 设置幻灯片标题 ====================

export const setSlideTitleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_title',
  description: `设置幻灯片的标题文本。

使用场景：
- "把第1页标题改成'年度总结'"
- "设置标题为'项目进展'"
- "修改幻灯片标题"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      title: {
        type: 'string',
        description: '标题文本',
      },
    },
    required: ['title'],
  },
};

export const setSlideTitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, title } = args as {
    slideIndex?: number;
    title: string;
  };

  const r = await call<{
    slideIndex: number;
    title: string;
  }>('setSlideTitle', { slideIndex: slideIndex || 1, title });
  if (!r.ok || !r.data) return fail(errText('设置幻灯片标题失败', r), r.error);
  return ok(
    `幻灯片标题设置成功！\n幻灯片: 第 ${r.data.slideIndex} 页\n标题: "${r.data.title}"`,
    r.warnings
  );
};

// ==================== 14. 对齐多个形状 ====================

export const alignShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_align_shapes',
  description: `对齐幻灯片中的多个形状。

支持的对齐方式：
- left: 左对齐
- center: 水平居中
- right: 右对齐
- top: 顶部对齐
- middle: 垂直居中
- bottom: 底部对齐

使用场景：
- "把这几个形状左对齐"
- "让第1、3、5个形状垂直居中"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要对齐的形状索引数组（从1开始）',
      },
      alignment: {
        type: 'string',
        description: '对齐方式',
        enum: ['left', 'center', 'right', 'top', 'middle', 'bottom'],
      },
    },
    required: ['slideIndex', 'shapeIndices', 'alignment'],
  },
};

export const alignShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices, alignment } = args as {
    slideIndex: number;
    shapeIndices: number[];
    alignment: string;
  };

  const r = await call<{ count?: number }>('alignShapes', { slideIndex, shapeIndices, alignment });
  if (!r.ok) return fail(errText('对齐形状失败', r), r.error);
  const alignName: Record<string, string> = {
    left: '左对齐',
    center: '水平居中',
    right: '右对齐',
    top: '顶部对齐',
    middle: '垂直居中',
    bottom: '底部对齐',
  };

  return ok(
    `形状对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}\n形状数量: ${shapeIndices.length} 个`,
    r.warnings
  );
};

// ==================== 15. 等距分布多个形状 ====================

export const distributeShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_distribute_shapes',
  description: `等距分布幻灯片中的多个形状。

支持的分布方向：
- horizontal: 水平等距分布
- vertical: 垂直等距分布

使用场景：
- "让这些形状水平等距排列"
- "垂直均匀分布这些元素"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要分布的形状索引数组（从1开始）',
      },
      direction: {
        type: 'string',
        description: '分布方向',
        enum: ['horizontal', 'vertical'],
      },
    },
    required: ['slideIndex', 'shapeIndices', 'direction'],
  },
};

export const distributeShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices, direction } = args as {
    slideIndex: number;
    shapeIndices: number[];
    direction: string;
  };

  const r = await call('distributeShapes', { slideIndex, shapeIndices, direction });
  if (!r.ok) return fail(errText('分布形状失败', r), r.error);
  const dirName = direction === 'horizontal' ? '水平等距' : '垂直等距';

  return ok(
    `形状分布完成！\n幻灯片: 第 ${slideIndex} 页\n分布方向: ${dirName}\n形状数量: ${shapeIndices.length} 个`,
    r.warnings
  );
};

// ==================== 16. 组合多个形状 ====================

export const groupShapesDefinition: ToolDefinition = {
  name: 'wps_ppt_group_shapes',
  description: `将幻灯片中的多个形状组合为一个组。

使用场景：
- "把这几个形状组合在一起"
- "将第1、2、3个形状编组"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description: '要组合的形状索引数组（从1开始，至少2个）',
      },
    },
    required: ['slideIndex', 'shapeIndices'],
  },
};

export const groupShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndices } = args as {
    slideIndex: number;
    shapeIndices: number[];
  };

  const r = await call<{ groupIndex?: number }>('groupShapes', { slideIndex, shapeIndices });
  if (!r.ok) return fail(errText('组合形状失败', r), r.error);
  let text = `形状组合成功！\n幻灯片: 第 ${slideIndex} 页\n组合形状: ${shapeIndices.length} 个`;
  if (r.data?.groupIndex) {
    text += `\n组合索引: ${r.data.groupIndex}`;
  }

  return ok(text, r.warnings);
};


export const shapeBasicTools: RegisteredTool[] = [
  { definition: addShapeDefinition, handler: addShapeHandler },
  { definition: addTextboxDefinition, handler: addTextboxHandler },
  { definition: deleteShapeDefinition, handler: deleteShapeHandler },
  { definition: getShapesDefinition, handler: getShapesHandler },
  { definition: setShapePositionDefinition, handler: setShapePositionHandler },
  { definition: setShapeTransparencyDefinition, handler: setShapeTransparencyHandler },
  { definition: setShapeTextDefinition, handler: setShapeTextHandler },
  { definition: setSlideTitleDefinition, handler: setSlideTitleHandler },
  { definition: alignShapesDefinition, handler: alignShapesHandler },
  { definition: distributeShapesDefinition, handler: distributeShapesHandler },
  { definition: groupShapesDefinition, handler: groupShapesHandler },
];

export default shapeBasicTools;
