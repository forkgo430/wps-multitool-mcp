/**
 * Input: PPT 形状样式参数（样式、阴影、渐变、边框、填充）
 * Output: 形状样式操作结果
 * Pos: PPT 形状样式工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText } from './helpers';

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

export const shapeStyleTools: RegisteredTool[] = [
  { definition: setShapeStyleDefinition, handler: setShapeStyleHandler },
  { definition: setShapeShadowDefinition, handler: setShapeShadowHandler },
  { definition: setShapeGradientDefinition, handler: setShapeGradientHandler },
  { definition: setShapeBorderDefinition, handler: setShapeBorderHandler },
  { definition: setShapeFillDefinition, handler: setShapeFillHandler },
];

export default shapeStyleTools;
