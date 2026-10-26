/**
 * Input: PPT 幻灯片操作参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT幻灯片Tools - 幻灯片管理模块
 * 处理幻灯片的添加、美化、字体统一等操作
 *
 * 包含：
 * - wps_ppt_add_slide: 添加新幻灯片
 * - wps_ppt_beautify: 美化幻灯片（核心功能）
 * - wps_ppt_unify_font: 统一字体
 * - wps_ppt_set_font_color: 设置文字颜色
 * - wps_ppt_align_objects: 对齐幻灯片中的对象
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

/**
 * 添加新幻灯片
 * 可以指定布局、位置、标题和内容
 */
export const addSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_add_slide',
  description: `添加新幻灯片到演示文稿。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容（最常用）
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局

使用场景：
- "新建一页PPT"
- "添加一个标题页"
- "在第3页后面插入一页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      layout: {
        type: 'string',
        description: '幻灯片布局类型',
        enum: ['title', 'title_content', 'blank', 'two_column', 'comparison'],
      },
      position: {
        type: 'number',
        description: '插入位置（页码），不填则在末尾添加',
      },
      title: {
        type: 'string',
        description: '幻灯片标题',
      },
      content: {
        type: 'string',
        description: '幻灯片内容（针对有内容区域的布局）',
      },
    },
    required: [],
  },
};

export const addSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { layout, position, title, content } = args as {
    layout?: string;
    position?: number;
    title?: string;
    content?: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      layout: string;
    }>(
      'addSlide',
      {
        layout: layout || 'title_content',
        position,
        title,
        content,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const layoutName = {
        title: '标题页',
        title_content: '标题+内容',
        blank: '空白页',
        two_column: '两栏内容',
        comparison: '对比布局',
      }[response.data.layout] || response.data.layout;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片添加成功！\n位置: 第 ${response.data.slideIndex} 页\n布局: ${layoutName}${title ? `\n标题: ${title}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 美化幻灯片
 * 一键优化排版、配色、字体和间距
 */
export const beautifyDefinition: ToolDefinition = {
  name: 'wps_ppt_beautify',
  description: `一键美化幻灯片，优化排版、配色、字体和间距。

支持的配色方案：
- business: 商务风（深蓝+灰色）
- tech: 科技风（蓝色+绿色）
- creative: 创意风（珊瑚红+金色）
- minimal: 简约风（黑白灰）

美化【实际会做】的操作：
- 统一字体（改所有有文字的文本框的字体名）
- 应用配色方案（按字号分标题/正文上色）

注意：不做自动对齐，也不动行距/间距 —— 这些改起来没有唯一正确答案，
一键"优化"很容易把用户排版搞乱。要精确调版面用 align_shapes / set_shape_position。

使用场景：
- "美化这页PPT"
- "用商务风格优化一下"
- "把PPT弄好看点"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slide_index: {
        type: 'number',
        description: '要美化的幻灯片页码，不填则美化当前页',
      },
      color_scheme: {
        type: 'string',
        description: '配色方案',
        enum: ['business', 'tech', 'creative', 'minimal'],
      },
      font: {
        type: 'string',
        description: '统一使用的字体，如 "微软雅黑"、"思源黑体"',
      },
      beautify_all: {
        type: 'boolean',
        description: '是否美化所有幻灯片，默认false只美化指定页',
      },
    },
    required: [],
  },
};

export const beautifyHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slide_index, color_scheme, font, beautify_all } = args as {
    slide_index?: number;
    color_scheme?: string;
    font?: string;
    beautify_all?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number | string;
      operations: Array<{
        operation: string;
        count: number;
        details?: string;
      }>;
    }>(
      'beautifySlide',
      {
        slideIndex: beautify_all ? 'all' : slide_index,
        style: {
          colorScheme: color_scheme || 'business',
          font: font || '微软雅黑',
        },
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const result = response.data;
      const schemeName = {
        business: '商务风',
        tech: '科技风',
        creative: '创意风',
        minimal: '简约风',
      }[color_scheme || 'business'] || color_scheme;

      let output = `幻灯片美化完成！\n`;
      output += `范围: ${beautify_all ? '所有幻灯片' : `第 ${result.slideIndex} 页`}\n`;
      output += `配色方案: ${schemeName}\n`;
      output += `字体: ${font || '微软雅黑'}\n\n`;

      // operations 以前不存在（PS 只回 {style,count}），直接 .forEach 会把整个工具打崩
      // ——冒烟第一条失败就是它。这里既兜空数组，也如实说"改了多少个文本"：
      // 0 个不是"成功"，是"这一页没有可美化的文字"，必须让用户看出来。
      const ops = Array.isArray(result.operations) ? result.operations : [];
      const touched = ops.reduce((sum, op) => sum + (op.count || 0), 0);
      if (touched === 0) {
        output +=
          '没有任何文本被改动 —— 这一页上没有带文字的形状/文本框。\n' +
          '先写入文字再美化（wps_ppt_set_shape_text / wps_ppt_set_textbox_text），或换有文字的页。';
      } else {
        output += `优化详情：\n`;
        for (const op of ops) {
          const opName =
            {
              unify_font: '统一字体',
              apply_color_scheme: '应用配色',
              align: '对齐元素',
              optimize_spacing: '优化间距',
            }[op.operation] || op.operation;
          output += `- ${opName}: 处理了 ${op.count} 个元素\n`;
        }
      }

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `美化幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `美化幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 统一字体
 * 整个演示文稿的字体统一是PPT美观的基础
 */
export const unifyFontDefinition: ToolDefinition = {
  name: 'wps_ppt_unify_font',
  description: `统一演示文稿中所有幻灯片的字体。

使用场景：
- "把所有页面的字体都改成微软雅黑"
- "统一字体"
- "换个字体"

常用字体推荐：
- 微软雅黑：现代简洁，适合商务
- 思源黑体：开源免费，适合各种场合
- 黑体：传统正式
- 宋体：适合正式文档`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      font_name: {
        type: 'string',
        description: '要统一使用的字体名称，如 "微软雅黑"、"思源黑体"',
      },
      slide_index: {
        type: 'number',
        description: '只处理指定页，不填则处理所有页',
      },
      include_title: {
        type: 'boolean',
        description: '是否包含标题，默认true',
      },
      include_body: {
        type: 'boolean',
        description: '是否包含正文，默认true',
      },
    },
    required: ['font_name'],
  },
};

export const unifyFontHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { font_name, slide_index, include_title, include_body } = args as {
    font_name: string;
    slide_index?: number;
    include_title?: boolean;
    include_body?: boolean;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      font: string;
      count: number;
      slideCount: number;
    }>(
      // 调 slide.unifyFont 而不是 unifyFont（第 3 轮修）：
      // PS 里 unifyFont 这个 case 会遍历【所有页、所有文本框】，压根不读
      // slideIndex / includeTitle / includeBody —— 工具描述承诺的三个开关全是摆设。
      // 认这三个参数的是另一个 case：slide.unifyFont。
      'slide.unifyFont',
      {
        fontName: font_name,
        slideIndex: slide_index,
        includeTitle: include_title !== false,
        includeBody: include_body !== false,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const result = response.data;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `字体统一完成！\n字体: ${result.font}\n处理幻灯片: ${result.slideCount} 页\n修改文本框: ${result.count} 个`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `统一字体失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `统一字体出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 设置文字颜色
 * 修改幻灯片中指定形状的文字颜色
 */
export const setFontColorDefinition: ToolDefinition = {
  name: 'wps_ppt_set_font_color',
  description: `设置幻灯片中指定形状的文字颜色。

使用场景：
- "把标题改成红色"
- "设置第2页第1个文本框的文字颜色为蓝色"
- "修改文字颜色"`,
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
      color: {
        type: 'string',
        description: '颜色值，支持十六进制如 "#FF0000" 或颜色名如 "red"',
      },
    },
    required: ['slideIndex', 'shapeIndex', 'color'],
  },
};

export const setFontColorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, shapeIndex, color } = args as {
    slideIndex: number;
    shapeIndex: number;
    color: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setFontColor', // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
      { slideIndex, shapeIndex, color },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文字颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n颜色: ${color}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置文字颜色失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置文字颜色出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 对齐幻灯片中的对象
 * 支持多种对齐方式
 */
export const alignObjectsDefinition: ToolDefinition = {
  name: 'wps_ppt_align_objects',
  description:
    '[DEPRECATED·改用 wps_ppt_align_shapes] 对齐幻灯片中的对象。\n\n' +
    '什么时候用：已被 wps_ppt_align_shapes 取代（同一个 alignShapes 动作、同一套参数），' +
    '新调用一律用 wps_ppt_align_shapes，别两个混着用。\n\n' +
    '注意：shapeIndices 现在是**必填**。这个工具原来没有形状清单参数，底层拿到空清单时' +
    '不是报错，而是"把整页所有形状一起对齐" —— 实测会把同页的表格和图片也搬走，' +
    '而工具照样回 success。要对齐的形状必须显式列出来。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      alignment: {
        type: 'string',
        description: '对齐方式',
        enum: ['left', 'center', 'right', 'top', 'middle', 'bottom', 'distribute_h', 'distribute_v'],
      },
      shapeIndices: {
        type: 'array',
        items: { type: 'number' },
        description:
          '要参与对齐的形状索引数组（从1开始）。**必填**：不给的话底层会退化成"整页所有形状一起对齐"，' +
          '把没打算动的表格/图片也搬走。要对齐全部就别用这个工具，用 wps_ppt_align_shapes 并显式列出。',
      },
    },
    required: ['slideIndex', 'alignment', 'shapeIndices'],
  },
};

export const alignObjectsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, alignment, shapeIndices } = args as {
    slideIndex: number;
    alignment: string;
    shapeIndices: number[];
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count?: number;
    }>(
      'alignShapes',
      { slideIndex, shapeIndices, alignment },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const alignName: Record<string, string> = {
        left: '左对齐',
        center: '水平居中',
        right: '右对齐',
        top: '顶部对齐',
        middle: '垂直居中',
        bottom: '底部对齐',
        distribute_h: '水平等距分布',
        distribute_v: '垂直等距分布',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `对象对齐完成！\n幻灯片: 第 ${slideIndex} 页\n对齐方式: ${alignName[alignment] || alignment}${response.data?.count ? `\n处理对象: ${response.data.count} 个` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `对齐对象失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `对齐对象出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

/**
 * 导出所有幻灯片相关的Tools
 */
export const slideTools: RegisteredTool[] = [
  { definition: addSlideDefinition, handler: addSlideHandler },
  { definition: beautifyDefinition, handler: beautifyHandler },
  { definition: unifyFontDefinition, handler: unifyFontHandler },
  { definition: setFontColorDefinition, handler: setFontColorHandler },
  { definition: alignObjectsDefinition, handler: alignObjectsHandler },
];

export default slideTools;
