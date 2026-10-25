/**
 * Input: PPT 背景、页脚、形状操作参数
 * Output: 背景/页脚/形状操作结果
 * Pos: PPT 背景与页面信息工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT背景Tools - 背景、页面信息、高级形状模块
 * 处理幻灯片背景设置、页码/页脚/日期时间、形状复制与层级调整
 *
 * 包含：
 * - wps_ppt_set_slide_background: 设置幻灯片背景（通用）
 * - wps_ppt_set_background_color: 设置幻灯片背景颜色
 * - wps_ppt_set_background_image: 设置幻灯片背景图片
 * - wps_ppt_set_slide_number: 设置幻灯片页码显示
 * - wps_ppt_set_ppt_footer: 设置PPT页脚
 * - wps_ppt_set_ppt_date_time: 设置PPT日期时间
 * - wps_ppt_duplicate_shape: 复制形状
 * - wps_ppt_set_shape_z_order: 设置形状层级顺序
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

// 第 5 轮：background 族 6→1 合并。
// set_slide_background / set_background_color / set_background_image 三者的能力
// 已并入 wps_ppt_set_background（见 slide-ops.ts 的 target/type 两个维度），
// 旧名留在 tool-aliases.json 里做别名，这里不再各留一份实现。
// 保留：页码、页脚、日期时间、形状复制、层级 —— 它们与背景无关。

/**
 * 设置幻灯片页码显示
 * 控制页码的显示/隐藏和起始编号
 */
export const setSlideNumberDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_number',
  description: `设置幻灯片页码的显示状态和起始编号。

使用场景：
- "显示页码"
- "隐藏幻灯片编号"
- "页码从第2页开始编号"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      show: {
        type: 'boolean',
        description: '是否显示页码',
      },
      startFrom: {
        type: 'number',
        description: '页码起始编号，默认为1',
      },
    },
    required: ['show'],
  },
};

export const setSlideNumberHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { show, startFrom } = args as {
    show: boolean;
    startFrom?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideNumber',
      { show, startFrom },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `幻灯片页码设置成功！\n显示状态: ${show ? '显示' : '隐藏'}`;
      if (startFrom !== undefined) {
        text += `\n起始编号: ${startFrom}`;
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页码失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页码出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置PPT页脚
 * 控制页脚文本的显示/隐藏
 */
export const setPptFooterDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_footer',
  description: `设置演示文稿页脚文本。

使用场景：
- "添加页脚'公司名称'"
- "设置页脚为'机密文件'"
- "隐藏页脚"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '页脚文本内容',
      },
      show: {
        type: 'boolean',
        description: '是否显示页脚，默认为true',
      },
    },
    required: ['text'],
  },
};

export const setPptFooterHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text, show } = args as {
    text: string;
    show?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptFooter',
      { text, show: show !== false },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `页脚设置成功！\n页脚文本: ${text}\n显示状态: ${show !== false ? '显示' : '隐藏'}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置页脚失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置页脚出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置PPT日期时间
 * 控制日期时间占位符的显示和格式
 */
export const setPptDateTimeDefinition: ToolDefinition = {
  name: 'wps_ppt_set_ppt_date_time',
  description: `设置演示文稿日期时间显示。

支持的日期格式（format）：
- "YYYY-MM-DD": 如 2026-03-21
- "YYYY/MM/DD": 如 2026/03/21
- "MM/DD/YYYY": 如 03/21/2026
- "DD/MM/YYYY": 如 21/03/2026

使用场景：
- "显示日期时间"
- "设置自动更新日期"
- "隐藏日期"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      show: {
        type: 'boolean',
        description: '是否显示日期时间',
      },
      autoUpdate: {
        type: 'boolean',
        description: '是否自动更新日期时间，默认为true',
      },
      format: {
        type: 'string',
        description: '日期时间格式，如 "YYYY-MM-DD"',
      },
    },
    required: ['show'],
  },
};

export const setPptDateTimeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { show, autoUpdate, format } = args as {
    show: boolean;
    autoUpdate?: boolean;
    format?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setPptDateTime',
      { show, autoUpdate, format },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `日期时间设置成功！\n显示状态: ${show ? '显示' : '隐藏'}`;
      if (show) {
        text += `\n自动更新: ${autoUpdate !== false ? '是' : '否'}`;
        if (format) {
          text += `\n格式: ${format}`;
        }
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置日期时间失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置日期时间出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ==================== 高级形状 (2) ====================

/**
 * 复制形状
 * 在同一幻灯片内复制指定形状
 */
export const duplicateShapeDefinition: ToolDefinition = {
  name: 'wps_ppt_duplicate_shape',
  description: `复制幻灯片中的指定形状。

使用场景：
- "复制这个形状"
- "把第1页的第2个形状复制一份"
- "克隆这个元素"`,
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
        description: '要复制的形状索引（从1开始）',
      },
    },
    required: ['slideIndex', 'shapeIndex'],
  },
};

export const duplicateShapeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex } = args as {
    slideIndex: number;
    shapeIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      newShapeIndex?: number;
    }>(
      'duplicateShape',
      { slideIndex, shapeIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      let text = `形状复制成功！\n幻灯片: 第 ${slideIndex} 页\n源形状: 第 ${shapeIndex} 个`;
      if (response.data?.newShapeIndex) {
        text += `\n新形状索引: ${response.data.newShapeIndex}`;
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `复制形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `复制形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置形状层级顺序
 * 调整形状在幻灯片中的前后层级
 */
export const setShapeZOrderDefinition: ToolDefinition = {
  name: 'wps_ppt_set_shape_z_order',
  description: `设置形状在幻灯片中的层级顺序（Z轴排列）。

支持的层级操作（order）：
- front: 置于顶层
- back: 置于底层
- forward: 上移一层
- backward: 下移一层

使用场景：
- "把这个形状移到最前面"
- "将形状置于底层"
- "上移一层"`,
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
      order: {
        type: 'string',
        description: '层级操作类型',
        enum: ['front', 'back', 'forward', 'backward'],
      },
    },
    required: ['slideIndex', 'shapeIndex', 'order'],
  },
};

export const setShapeZOrderHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, order } = args as {
    slideIndex: number;
    shapeIndex: number;
    order: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setShapeZOrder',
      { slideIndex, shapeIndex, order },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const orderName: Record<string, string> = {
        front: '置于顶层',
        back: '置于底层',
        forward: '上移一层',
        backward: '下移一层',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状层级调整成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n操作: ${orderName[order] || order}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `调整形状层级失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `调整形状层级出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有背景、页面信息、高级形状相关的Tools
 */
export const backgroundTools: RegisteredTool[] = [
  { definition: setSlideNumberDefinition, handler: setSlideNumberHandler },
  { definition: setPptFooterDefinition, handler: setPptFooterHandler },
  { definition: setPptDateTimeDefinition, handler: setPptDateTimeHandler },
  { definition: duplicateShapeDefinition, handler: duplicateShapeHandler },
  { definition: setShapeZOrderDefinition, handler: setShapeZOrderHandler },
];

export default backgroundTools;
