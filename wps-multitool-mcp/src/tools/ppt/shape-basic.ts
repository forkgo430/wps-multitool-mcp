/**
 * Input: PPT 形状操作参数（删除、获取、位置、阴影、渐变、边框、透明度、对齐、分布、组合、添加、样式、文本框、标题、文字、填充）
 * Output: 形状操作结果
 * Pos: PPT 形状基础工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT形状基础Tools - 形状管理模块
 * 处理形状的添加、删除、获取、定位、样式设置、对齐分布和组合操作
 *
 * 包含（16 个；add_shape/set_shape_style/add_textbox/set_slide_title/set_shape_text/set_shape_fill
 * 自 slide-ops.ts 迁入——风格统一+拆分轮，见 ppt-audit §三之二）：
 * - wps_ppt_add_shape: 添加形状
 * - wps_ppt_add_textbox: 添加文本框
 * - wps_ppt_delete_shape: 删除形状
 * - wps_ppt_get_shapes: 获取幻灯片中的形状列表
 * - wps_ppt_set_shape_position: 设置形状位置和大小
 * - wps_ppt_set_shape_style: 设置形状样式
 * - wps_ppt_set_shape_shadow: 设置形状阴影
 * - wps_ppt_set_shape_gradient: 设置形状渐变填充
 * - wps_ppt_set_shape_border: 设置形状边框
 * - wps_ppt_set_shape_transparency: 设置形状透明度
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 * - wps_ppt_set_slide_title: 设置幻灯片标题
 * - wps_ppt_align_shapes: 对齐多个形状
 * - wps_ppt_distribute_shapes: 等距分布多个形状
 * - wps_ppt_group_shapes: 组合多个形状
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText } from './helpers';

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
        description: '左边距（像素），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（像素），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（像素），默认100',
      },
      height: {
        type: 'number',
        description: '高度（像素），默认100',
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
    name: string;
    slideIndex: number;
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
  if (!r.ok || !r.data) return fail(errText('添加形状失败', r), r.error);
  const shapeNameMap: Record<string, string> = {
    rectangle: '矩形', oval: '椭圆', triangle: '三角形', diamond: '菱形',
    pentagon: '五边形', hexagon: '六边形', arrow: '箭头', star: '星形',
    heart: '心形', cloud: '云形',
  };
  const typeName = shapeNameMap[type || 'rectangle'] || type;

  return ok(
    `形状添加成功！\n幻灯片: 第 ${r.data.slideIndex} 页\n形状类型: ${typeName}\n形状名称: ${r.data.name}${text ? `\n文本: ${text}` : ''}${fillColor ? `\n填充色: ${fillColor}` : ''}`,
    r.warnings
  );
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
        description: '左边距（像素），默认100',
      },
      top: {
        type: 'number',
        description: '上边距（像素），默认100',
      },
      width: {
        type: 'number',
        description: '宽度（像素），默认200',
      },
      height: {
        type: 'number',
        description: '高度（像素），默认50',
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
    name: string;
    slideIndex: number;
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
  if (!r.ok || !r.data) return fail(errText('添加文本框失败', r), r.error);
  return ok(
    `文本框添加成功！\n幻灯片: 第 ${r.data.slideIndex} 页\n名称: ${r.data.name}${text ? `\n内容: "${text}"` : ''}${fontSize ? `\n字号: ${fontSize}` : ''}${fontName ? `\n字体: ${fontName}` : ''}`,
    r.warnings
  );
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

返回每个形状的类型、位置、大小等属性。

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
    },
    required: ['slideIndex'],
  },
};

export const getShapesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as {
    slideIndex: number;
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
    }>;
  }>('getShapes', { slideIndex });
  if (!r.ok || !r.data) return fail(errText('获取形状列表失败', r), r.error);
  const shapes = r.data.shapes;
  let output = `幻灯片第 ${slideIndex} 页的形状列表（共 ${shapes.length} 个）：\n\n`;
  shapes.forEach((s) => {
    output += `[${s.index}] ${s.name} (${s.type})\n`;
    output += `    位置: (${s.left}, ${s.top}) 大小: ${s.width} x ${s.height}\n`;
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

// ==================== 6. 设置形状样式 ====================

export const setShapeStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_style',
  description: `设置幻灯片中形状的样式，包括填充颜色、边框颜色和边框粗细。

使用场景：
- "把矩形改成红色"
- "设置形状的边框为蓝色"
- "修改形状样式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      name: {
        type: 'string',
        description: '形状名称（通过getSlideInfo获取）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（与name二选一）',
      },
      fillColor: {
        type: 'string',
        description: '填充颜色，十六进制如 #FF0000',
      },
      lineColor: {
        type: 'string',
        description: '边框颜色，十六进制如 #000000',
      },
      lineWidth: {
        type: 'number',
        description: '边框粗细（磅）',
      },
    },
    required: [],
  },
};

export const setShapeStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, name, shapeIndex, fillColor, lineColor, lineWidth } = args as {
    slideIndex?: number;
    name?: string;
    shapeIndex?: number;
    fillColor?: string;
    lineColor?: string;
    lineWidth?: number;
  };

  // 定位参数至少给一个：都缺时 ps1 会 Shapes.Item($null) 直接炸（审计三.6），前置拦
  if (!name && !shapeIndex) {
    return fail(
      `要给 name（形状名，getSlideInfo 里看）或 shapeIndex（形状序号）至少一个，否则不知道改哪个形状`,
      'Missing required parameter: name / shapeIndex（二选一）'
    );
  }

  const r = await call<{ name: string }>(
    'setShapeStyle',
    {
      slideIndex: slideIndex || 1,
      name,
      shapeIndex,
      fillColor,
      lineColor,
      lineWidth,
    }
  );
  if (!r.ok || !r.data) return fail(errText('设置形状样式失败', r), r.error);
  let output = `形状样式设置成功！\n形状: ${r.data.name}`;
  if (fillColor) output += `\n填充色: ${fillColor}`;
  if (lineColor) output += `\n边框色: ${lineColor}`;
  if (lineWidth) output += `\n边框粗细: ${lineWidth}pt`;

  return ok(output, r.warnings);
};

// ==================== 7. 设置形状阴影 ====================

export const setShapeShadowDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_shadow',
  description: `设置幻灯片中指定形状的阴影效果。

shadow对象属性：
- enabled: 是否启用阴影 (boolean)
- color: 阴影颜色，如 "#000000"
- blur: 模糊半径（磅）
- offsetX: 水平偏移（磅）
- offsetY: 垂直偏移（磅）
- opacity: 透明度 (0-1)

**⚠️ 阴影只进对象模型，不进 WPS 的导出渲染**。实测：设阴影前后把这一页导出成图片，
两次导出的字节完全一样。所以**别用"导出图有没有变"来判断阴影设上没有** ——
看不出变化不代表没生效，要看形状属性本身。

使用场景：
- "给这个形状加阴影"
- "设置阴影效果"`,
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
      shadow: {
        type: 'object',
        description: '阴影配置对象',
        properties: {
          enabled: { type: 'boolean', description: '是否启用阴影' },
          color: { type: 'string', description: '阴影颜色' },
          blur: { type: 'number', description: '模糊半径（磅）' },
          offsetX: { type: 'number', description: '水平偏移（磅）' },
          offsetY: { type: 'number', description: '垂直偏移（磅）' },
          opacity: { type: 'number', description: '透明度 (0-1)' },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'shadow'],
  },
};

export const setShapeShadowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, shadow } = args as {
    slideIndex: number;
    shapeIndex: number;
    shadow: Record<string, unknown>;
  };

  const r = await call('setShapeShadow', { slideIndex, shapeIndex, shadow });
  if (!r.ok) return fail(errText('设置形状阴影失败', r), r.error);
  return ok(`形状阴影设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`, r.warnings);
};

// ==================== 8. 设置形状渐变填充 ====================

export const setShapeGradientDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_gradient',
  description: `设置幻灯片中指定形状的渐变填充效果。

gradient对象属性：
- type: 渐变类型，"linear"(线性) 或 "radial"(径向)
- angle: 渐变角度（0-360，仅线性渐变）
- stops: 渐变色标数组 [{color: "#FF0000", position: 0}, {color: "#0000FF", position: 1}]

使用场景：
- "给形状加渐变色"
- "设置从红到蓝的渐变"`,
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
      gradient: {
        type: 'object',
        description: '渐变配置对象',
        properties: {
          type: { type: 'string', description: '渐变类型: linear 或 radial', enum: ['linear', 'radial'] },
          angle: { type: 'number', description: '渐变角度（0-360）' },
          stops: {
            type: 'array',
            description: '渐变色标数组',
            items: {
              type: 'object',
              properties: {
                color: { type: 'string', description: '颜色值' },
                position: { type: 'number', description: '位置 (0-1)' },
              },
            },
          },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'gradient'],
  },
};

export const setShapeGradientHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, gradient } = args as {
    slideIndex: number;
    shapeIndex: number;
    gradient: Record<string, unknown>;
  };

  const r = await call('setShapeGradient', { slideIndex, shapeIndex, gradient });
  if (!r.ok) return fail(errText('设置形状渐变失败', r), r.error);
  return ok(`形状渐变填充设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`, r.warnings);
};

// ==================== 9. 设置形状边框 ====================

export const setShapeBorderDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_border',
  description: `设置幻灯片中指定形状的边框样式。

border对象属性：
- enabled: 是否启用边框 (boolean)
- color: 边框颜色，如 "#000000"
- weight: 边框粗细（磅）
- style: 边框样式，如 "solid"(实线)、"dash"(虚线)、"dot"(点线)

使用场景：
- "给形状加边框"
- "设置红色虚线边框"`,
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
      border: {
        type: 'object',
        description: '边框配置对象',
        properties: {
          enabled: { type: 'boolean', description: '是否启用边框' },
          color: { type: 'string', description: '边框颜色' },
          weight: { type: 'number', description: '边框粗细（磅）' },
          style: { type: 'string', description: '边框样式', enum: ['solid', 'dash', 'dot', 'dash_dot', 'dash_dot_dot'] },
        },
      },
    },
    required: ['slideIndex', 'shapeIndex', 'border'],
  },
};

export const setShapeBorderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, border } = args as {
    slideIndex: number;
    shapeIndex: number;
    border: Record<string, unknown>;
  };

  const r = await call('setShapeBorder', { slideIndex, shapeIndex, border });
  if (!r.ok) return fail(errText('设置形状边框失败', r), r.error);
  return ok(`形状边框设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个`, r.warnings);
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

// ==================== 12. 设置形状填充颜色 ====================

export const setShapeFillDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_fill',
  description: `设置幻灯片中指定形状的填充颜色。

使用场景：
- "把第1页的第2个形状填充为红色"
- "修改形状背景色为#00FF00"
- "设置形状的填充颜色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      shapeIndex: { type: 'number', description: '形状索引（从1开始）' },
      color: { type: 'string', description: '填充颜色，十六进制如 #FF0000' },
    },
    required: ['slideIndex', 'shapeIndex', 'color'],
  },
};

export const setShapeFillHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, color } = args as {
    slideIndex: number;
    shapeIndex: number;
    color: string;
  };

  const r = await call(
    // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
    'setShapeFill',
    { slideIndex, shapeIndex, color }
  );
  if (!r.ok) return fail(errText('设置形状填充失败', r), r.error);
  return ok(
    `形状填充颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n填充颜色: ${color}`,
    r.warnings
  );
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

/**
 * 导出所有形状基础相关的Tools
 */
export const shapeBasicTools: RegisteredTool[] = [
  { definition: addShapeDefinition, handler: addShapeHandler },
  { definition: addTextboxDefinition, handler: addTextboxHandler },
  { definition: deleteShapeDefinition, handler: deleteShapeHandler },
  { definition: getShapesDefinition, handler: getShapesHandler },
  { definition: setShapePositionDefinition, handler: setShapePositionHandler },
  { definition: setShapeStyleDefinition, handler: setShapeStyleHandler },
  { definition: setShapeShadowDefinition, handler: setShapeShadowHandler },
  { definition: setShapeGradientDefinition, handler: setShapeGradientHandler },
  { definition: setShapeBorderDefinition, handler: setShapeBorderHandler },
  { definition: setShapeTransparencyDefinition, handler: setShapeTransparencyHandler },
  { definition: setShapeTextDefinition, handler: setShapeTextHandler },
  { definition: setShapeFillDefinition, handler: setShapeFillHandler },
  { definition: setSlideTitleDefinition, handler: setSlideTitleHandler },
  { definition: alignShapesDefinition, handler: alignShapesHandler },
  { definition: distributeShapesDefinition, handler: distributeShapesHandler },
  { definition: groupShapesDefinition, handler: groupShapesHandler },
];

export default shapeBasicTools;
