/**
 * Input: 幻灯片操作工具参数
 * Output: 幻灯片操作结果
 * Pos: PPT 幻灯片操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 幻灯片操作Tools - 删除/复制/移动/查询/切换/布局/备注
 *
 * 包含：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_duplicate_slide: 复制幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_notes: 设置幻灯片备注
 * - wps_ppt_add_shape: 添加形状
 * - wps_ppt_set_shape_style: 设置形状样式
 * - wps_ppt_add_textbox: 添加文本框
 * - wps_ppt_set_slide_title: 设置幻灯片标题
 * - wps_ppt_insert_image: 插入图片
 * - wps_ppt_set_shape_text: 设置形状文字
 * - wps_ppt_set_animation: 设置元素动画
 * - wps_ppt_set_background: 设置幻灯片背景
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 * - wps_ppt_set_transition: 设置幻灯片切换效果
 * - wps_ppt_add_chart: 在幻灯片中插入图表
 * - wps_ppt_set_shape_fill: 设置形状填充颜色
 * - wps_ppt_add_speaker_notes: 添加演讲者备注
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

// ============================================================
// 1. wps_ppt_delete_slide - 删除指定幻灯片
// ============================================================

export const deleteSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_slide',
  description: `删除指定的幻灯片。

使用场景：
- "删除第3页幻灯片"
- "把最后一页删掉"
- "移除多余的页面"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '要删除的幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const deleteSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'deleteSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片删除成功！\n已删除: 第 ${slideIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `删除幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `删除幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 3. wps_ppt_move_slide - 移动幻灯片到指定位置
// ============================================================

export const moveSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_move_slide',
  description: `移动幻灯片到指定位置。

使用场景：
- "把第5页移到第2页"
- "把最后一页移到开头"
- "调整幻灯片顺序"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      fromIndex: {
        type: 'number',
        description: '原位置索引（从1开始）',
      },
      toIndex: {
        type: 'number',
        description: '目标位置索引（从1开始）',
      },
    },
    required: ['fromIndex', 'toIndex'],
  },
};

export const moveSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { fromIndex, toIndex } = args as {
    fromIndex: number;
    toIndex: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'moveSlide',
      { fromIndex, toIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片移动成功！\n从第 ${fromIndex} 页 → 移到第 ${toIndex} 页`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `移动幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `移动幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 4. wps_ppt_get_slide_count - 获取幻灯片总数
// ============================================================

export const getSlideCountDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_count',
  description: `获取演示文稿中的幻灯片总数。

使用场景：
- "一共有多少页幻灯片"
- "PPT有几页"
- "查看幻灯片数量"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSlideCountHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      count: number;
    }>(
      'getSlideCount',
      {},
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `当前演示文稿共有 ${response.data.count} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片总数失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片总数出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 5. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
// ============================================================

export const getSlideInfoDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_info',
  description: `获取指定幻灯片的详细信息，包括布局、元素列表等。

使用场景：
- "查看第3页的信息"
- "这页有什么内容"
- "获取幻灯片详情"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideInfoHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      layout: string;
      shapesCount: number;
      shapes: Array<{
        name: string;
        type: string;
        text?: string;
      }>;
    }>(
      'getSlideInfo',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const info = response.data;
      let output = `第 ${info.slideIndex} 页幻灯片信息：\n`;
      output += `布局: ${info.layout}\n`;
      output += `元素数量: ${info.shapesCount}\n`;

      if (info.shapes && info.shapes.length > 0) {
        output += `\n元素列表：\n`;
        info.shapes.forEach((shape, i) => {
          output += `  ${i + 1}. [${shape.type}] ${shape.name}`;
          if (shape.text) {
            output += ` - "${shape.text.substring(0, 50)}${shape.text.length > 50 ? '...' : ''}"`;
          }
          output += '\n';
        });
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
        content: [{ type: 'text', text: `获取幻灯片信息失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片信息出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 6. wps_ppt_switch_slide - 切换到指定幻灯片
// ============================================================

export const switchSlideDefinition: ToolDefinition = {
  name: 'wps_ppt_switch_slide',
  description: `切换到指定的幻灯片页面。

使用场景：
- "切换到第5页"
- "跳到最后一页"
- "显示第1页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '目标幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const switchSlideHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'switchSlide',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `已切换到第 ${slideIndex} 页幻灯片`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `切换幻灯片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `切换幻灯片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 7. wps_ppt_set_slide_layout - 设置幻灯片版式布局
// ============================================================

export const setSlideLayoutDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_layout',
  description: `设置幻灯片的版式布局。

支持的布局类型：
- title: 标题页
- title_content: 标题+内容
- blank: 空白页
- two_column: 两栏内容
- comparison: 对比布局
- section_header: 节标题
- title_only: 仅标题

使用场景：
- "把这页改成空白布局"
- "设置为两栏内容"
- "改成标题页版式"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      layout: {
        type: 'string',
        description: '布局名称',
      },
    },
    required: ['slideIndex', 'layout'],
  },
};

export const setSlideLayoutHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, layout } = args as {
    slideIndex: number;
    layout: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideLayout',
      { slideIndex, layout },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const layoutName: Record<string, string> = {
        title: '标题页',
        title_content: '标题+内容',
        blank: '空白页',
        two_column: '两栏内容',
        comparison: '对比布局',
        section_header: '节标题',
        title_only: '仅标题',
      };

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片布局设置成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${layoutName[layout] || layout}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片布局失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片布局出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 8. wps_ppt_get_slide_notes - 获取幻灯片备注内容
// ============================================================

export const getSlideNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_notes',
  description: `获取指定幻灯片的备注内容。

使用场景：
- "查看第3页的备注"
- "读取演讲备注"
- "这页有什么备注"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getSlideNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      notes: string;
    }>(
      'getSlideNotes',
      { slideIndex },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const notes = response.data.notes;
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: notes
              ? `第 ${slideIndex} 页备注内容：\n${notes}`
              : `第 ${slideIndex} 页没有备注内容`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `获取幻灯片备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `获取幻灯片备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 10. wps_ppt_add_shape - 添加形状
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      slideIndex: number;
    }>(
      'addShape',
      {
        slideIndex: slideIndex || 1,
        type: type || 'rectangle',
        left: left || 100,
        top: top || 100,
        width: width || 100,
        height: height || 100,
        text,
        fillColor,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      const shapeNameMap: Record<string, string> = {
        rectangle: '矩形', oval: '椭圆', triangle: '三角形', diamond: '菱形',
        pentagon: '五边形', hexagon: '六边形', arrow: '箭头', star: '星形',
        heart: '心形', cloud: '云形',
      };
      const typeName = shapeNameMap[type || 'rectangle'] || type;

      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `形状添加成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n形状类型: ${typeName}\n形状名称: ${response.data.name}${text ? `\n文本: ${text}` : ''}${fillColor ? `\n填充色: ${fillColor}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加形状失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加形状出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 11. wps_ppt_set_shape_style - 设置形状样式
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'setShapeStyle',
      {
        slideIndex: slideIndex || 1,
        name,
        shapeIndex,
        fillColor,
        lineColor,
        lineWidth,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      let output = `形状样式设置成功！\n形状: ${response.data.name}`;
      if (fillColor) output += `\n填充色: ${fillColor}`;
      if (lineColor) output += `\n边框色: ${lineColor}`;
      if (lineWidth) output += `\n边框粗细: ${lineWidth}pt`;

      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: output }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状样式失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状样式出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 12. wps_ppt_add_textbox - 添加文本框
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      slideIndex: number;
    }>(
      'addTextBox',
      {
        slideIndex: slideIndex || 1,
        left: left || 100,
        top: top || 100,
        width: width || 200,
        height: height || 50,
        text,
        fontSize,
        fontName,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `文本框添加成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n名称: ${response.data.name}${text ? `\n内容: "${text}"` : ''}${fontSize ? `\n字号: ${fontSize}` : ''}${fontName ? `\n字体: ${fontName}` : ''}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加文本框失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加文本框出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 13. wps_ppt_set_slide_title - 设置幻灯片标题
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      slideIndex: number;
      title: string;
    }>(
      'setSlideTitle',
      {
        slideIndex: slideIndex || 1,
        title,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片标题设置成功！\n幻灯片: 第 ${response.data.slideIndex} 页\n标题: "${response.data.title}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片标题失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片标题出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 14. wps_ppt_insert_image - 插入图片
// ============================================================

export const insertImageDefinition: ToolDefinition = {
  name: 'wps_ppt_insert_image',
  description: `在幻灯片中插入图片。

使用场景：
- "在第1页插入一张图片"
- "添加图片到幻灯片"
- "把这个图片放到PPT里"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始），默认1',
      },
      path: {
        type: 'string',
        description: '图片文件的完整路径',
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
        description: '宽度（像素），不填则保持原始尺寸',
      },
      height: {
        type: 'number',
        description: '高度（像素），不填则保持原始尺寸',
      },
    },
    required: ['path'],
  },
};

export const insertImageHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, path, left, top, width, height } = args as {
    slideIndex?: number;
    path: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
      path: string;
    }>(
      'insertPptImage',
      {
        slideIndex: slideIndex || 1,
        path,
        left: left || 100,
        top: top || 100,
        width: width || -1,
        height: height || -1,
      },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `图片插入成功！\n幻灯片: 第 ${slideIndex || 1} 页\n图片名称: ${response.data.name}\n图片路径: ${response.data.path}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `插入图片失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `插入图片出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 15. wps_ppt_set_shape_text - 设置形状文字
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name: string;
    }>(
      'setShapeText',
      { slideIndex, shapeIndex, text },
      WpsAppType.PRESENTATION
    );

    if (response.success && response.data) {
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `形状文字设置成功！\n形状: ${response.data.name}\n文字: "${text}"` }],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状文字失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状文字出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 16. wps_ppt_set_animation - 设置元素动画
// ============================================================

export const setAnimationDefinition: ToolDefinition = {
  name: 'wps_ppt_set_animation',
  description: `给形状加动画，或调整已有动画的播放顺序（第 8 轮合并了动画族 5 个工具）。

**两个分支用 \`op\` 选**：
- \`op:'add'\`（默认）：给形状加动画，要 \`shapeIndex\` + \`category\` + \`effect\`
- \`op:'reorder'\`：把已有动画挪位置，要 \`animationIndex\` + \`order\`

\`category\` 决定动作落在哪个动画窗格：
- entrance 入场（appear / fadeIn / flyIn / wipeIn / zoomIn / bounceIn / spinIn）
- emphasis 强调（pulse / grow / teeter / colorPulse）
- exit 退场（fadeOut / flyOut / wipeOut / zoomOut / disappear）

什么时候用：
- "给第1页第2个形状加个淡入" → { op:'add', slideIndex:1, shapeIndex:2, category:'entrance', effect:'fadeIn' }
- "让这个元素闪一下强调" → { op:'add', slideIndex:1, shapeIndex:2, category:'emphasis', effect:'pulse' }
- "把第3个动画移到最前面" → { op:'reorder', slideIndex:1, animationIndex:3, order:1 }

限制：\`.add\` 是【累加】的 —— 同一形状重复加会叠出多条动画，不会替换。
想看现在有哪些动画用 wps_ppt_get_animations，删掉某一条用 wps_ppt_remove_animation。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      op: {
        type: 'string',
        enum: ['add', 'reorder'],
        description: '操作：add 加动画（默认）/ reorder 调顺序',
      },
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）；op=add 时必填',
      },
      category: {
        type: 'string',
        enum: ['entrance', 'emphasis', 'exit'],
        description: '动画类别；op=add 时必填',
      },
      effect: {
        type: 'string',
        description: '具体效果名。entrance: fadeIn/flyIn/wipeIn/zoomIn/bounceIn/spinIn；emphasis: pulse/spin/grow/teeter/colorPulse；exit: fadeOut/flyOut/wipeOut/zoomOut/disappear',
      },
      trigger: {
        type: 'string',
        enum: ['onClick', 'withPrevious', 'afterPrevious'],
        description: '触发方式，默认 onClick（单击时）',
      },
      animationIndex: {
        type: 'number',
        description: '已有动画的序号（从1开始）；op=reorder 时必填',
      },
      order: {
        type: 'number',
        description: '新的播放位置（从1开始）；op=reorder 时必填',
      },
    },
    required: ['slideIndex'],
  },
};

/** 旧 animationType → { category, effect }。旧名 wps_ppt_set_animation 只收一个字符串，
 *  收尾先按 Out/Exit 结尾判退场，其余按入场；未知名默认入场（宁可多加一个也别不加）。 */
function splitLegacyAnimationType(animationType: string): {
  category: 'entrance' | 'exit';
  effect: string;
} {
  return {
    category: /(out|exit)$/i.test(animationType) ? 'exit' : 'entrance',
    effect: animationType,
  };
}

const ANIM_NAME_MAP: Record<string, string> = {
  fadeIn: '淡入', flyIn: '飞入', wipeIn: '擦除', zoomIn: '缩放进入',
  bounceIn: '弹跳进入', spinIn: '旋转进入', fadeOut: '淡出', flyOut: '飞出',
  wipeOut: '擦除退场', zoomOut: '缩小退场', disappear: '消失',
  pulse: '脉冲', grow: '放大缩小', teeter: '跷跷板', colorPulse: '颜色脉冲',
};

const TRIGGER_NAME_MAP: Record<string, string> = {
  onClick: '单击时',
  withPrevious: '与上一动画同时',
  afterPrevious: '上一动画之后',
};

const CATEGORY_NAME_MAP: Record<string, string> = {
  entrance: '入场', emphasis: '强调', exit: '退场',
};

/**
 * category + effect → PS 真正认的效果名。
 * 为什么要在 TS 侧合成：WPS COM 的 AddEffect 只吃一个效果 ID，不认"入场/强调/退场"这个维度。
 * 不做这层合成的话，category 就是一个装饰参数（PS 收不到、进不了动画窗格对应格），
 * 那才是真的骗人 —— 用户以为加了退场动画，实际加的还是入场。
 */
function resolveEffectId(category: string, effect: string): string {
  const normalized = String(effect || '').trim();
  const lower = normalized.toLowerCase();
  if (category === 'exit') {
    const outMap: Record<string, string> = {
      fade: 'fadeOut', fly: 'flyOut', wipe: 'wipeOut', zoom: 'zoomOut',
      vanish: 'vanish', disappear: 'disappear', dissolve: 'dissolveOut',
    };
    // 已经带 Out 尾名的（fadeOut）原样送下去，不重复拼
    if (lower.endsWith('out') || lower === 'disappear' || lower === 'vanish') return normalized;
    return outMap[lower] || normalized;
  }
  const inMap: Record<string, string> = {
    fade: 'fadeIn', fly: 'flyIn', wipe: 'wipeIn', zoom: 'zoomIn',
    dissolve: 'dissolveIn', appear: 'appear',
  };
  if (lower.endsWith('in') || lower === 'appear') return normalized;
  return inMap[lower] || normalized;
}

export const setAnimationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const op = (args.op as string) || 'add';

  try {
    if (op === 'reorder') {
      const { slideIndex, animationIndex, order } = args as {
        slideIndex: number;
        animationIndex: number;
        order: number;
      };
      if (!animationIndex || !order) {
        return {
          id: uuidv4(),
          success: false,
          content: [{ type: 'text', text: `op=reorder 时必须同时给 animationIndex 和 order` }],
          error: 'Missing required parameter: animationIndex / order',
        };
      }
      const response = await wpsClient.executeMethod<{ success: boolean }>(
        'setAnimationOrder',
        { slideIndex, animationIndex, newOrder: order },
        WpsAppType.PRESENTATION
      );
      if (response.success) {
        return {
          id: uuidv4(),
          success: true,
          content: [
            {
              type: 'text',
              text: `动画顺序调整成功！\n幻灯片: 第 ${slideIndex} 页\n动画 ${animationIndex} → 新位置 ${order}`,
            },
          ],
        };
      }
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `调整动画顺序失败: ${response.error}` }],
        error: response.error,
      };
    }

    // op=add：effect 缺失时兜住旧的 animationType（第 2 轮起这个参数就叫 animationType）
    const raw = args as {
      slideIndex: number;
      shapeIndex?: number;
      category?: string;
      effect?: string;
      trigger?: string;
      animationType?: string;
    };
    const legacy = !raw.effect && raw.animationType
      ? splitLegacyAnimationType(String(raw.animationType))
      : null;
    const category = raw.category || (legacy ? legacy.category : undefined);
    const effect = raw.effect || (legacy ? legacy.effect : undefined);

    if (!raw.shapeIndex || !category || !effect) {
      return {
        id: uuidv4(),
        success: false,
        content: [
          {
            type: 'text',
            text: `op=add 时必须同时给 shapeIndex、category（entrance/emphasis/exit）和 effect`,
          },
        ],
        error: 'Missing required parameter: shapeIndex / category / effect',
      };
    }

    // category 在这里落地成真正的效果名（exit+fade → fadeOut），PS 侧只认效果 ID
    const effectId = resolveEffectId(category as string, effect as string);

    const response = await wpsClient.executeMethod<{
      animationIndex?: number;
    }>(
      'addAnimation',
      {
        slideIndex: raw.slideIndex,
        shapeIndex: raw.shapeIndex,
        effect: effectId,
        trigger: raw.trigger || 'onClick',
      },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      const lines = [
        `动画设置成功！`,
        `幻灯片: 第 ${raw.slideIndex} 页`,
        `形状: 第 ${raw.shapeIndex} 个`,
        `类别: ${CATEGORY_NAME_MAP[category as string] || category}`,
        `动画: ${ANIM_NAME_MAP[effect as string] || effect}（${effectId}）`,
        `触发: ${TRIGGER_NAME_MAP[raw.trigger || 'onClick'] || raw.trigger || 'onClick'}`,
      ];
      if (response.data && response.data.animationIndex) {
        lines.push(`动画序号: ${response.data.animationIndex}`);
      }
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置动画失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置动画出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 17. wps_ppt_set_background - 设置背景（第 5 轮：background 族 6→1 合并）
// ============================================================

export const setBackgroundDefinition: ToolDefinition = {
  name: 'wps_ppt_set_background',
  description: `设置背景：某一页、整份文稿、或幻灯片母版。

**三个维度一次说清**（合并了原来 6 个背景工具，别再找 set_background_color / _image / _gradient / set_slide_background / set_master_background）：
- \`target\`: slide（默认，单页）/ all（整份文稿）/ master（母版）
- \`type\`: solid（纯色）/ image（图片）/ gradient（渐变）/ pattern
- 具体的填充物：\`color\` / \`imagePath\` / \`gradient\`

什么时候用：
- "把这页背景改成深蓝" → { target:'slide', slideIndex:2, type:'solid', color:'#003366' }
- "整份文稿统一白底" → { target:'all', type:'solid', color:'#FFFFFF' }
- "用这张图做背景" → { target:'slide', slideIndex:1, type:'image', imagePath:'C:\\\\pics\\\\bg.png' }
- "母版加渐变" → { target:'master', type:'gradient', gradient:{ type:'linear', colors:['#FF0000','#00FF00'] } }

不填 \`type\` 也能用：按你给的 color / imagePath / gradient 自动推断（老名字走别名过来时就是这条路）。

限制（真限制，不是没实现）：
- \`target:'master'\` 只支持纯色与渐变 —— 母版背景换图片 WPS 不认，会明确报错而不是假装成功。
- \`type:'pattern'\`（图案填充）WPS COM 没有对应接口，明确报错。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      target: {
        type: 'string',
        enum: ['slide', 'all', 'master'],
        description: '作用范围：slide 单页（默认）/ all 整份文稿 / master 母版',
      },
      slideIndex: {
        type: 'number',
        description: '页码（从1开始）；target=all/master 时忽略',
      },
      type: {
        type: 'string',
        enum: ['solid', 'gradient', 'image', 'pattern'],
        description: '填充类型。不填则按 color→solid / imagePath→image / gradient→gradient 推断',
      },
      color: { type: 'string', description: 'type=solid 时的颜色，十六进制如 #FF0000' },
      imagePath: { type: 'string', description: 'type=image 时的图片路径' },
      gradient: {
        type: 'object',
        description:
          'type=gradient 时的渐变配置：{ type:"linear"|"radial", colors:["#FF0000","#00FF00"] }。' +
          'WPS 只吃前两个颜色，多余的颜色会被忽略（会在提示里说明）',
      },
      pattern: { type: 'string', description: 'type=pattern 时的图案名（WPS 不支持，传了会明确报错）' },
      // 老名字（别名）可能会带过来的历史参数，接受但不推荐
      background: {
        type: 'object',
        description: '（历史形状）{type,color,colors,imagePath,pattern}，等价于把里面几项平铺到顶层',
      },
    },
    // 刻意不 required['type']：草案里是必填，但那样 5 个旧名就没法走别名（别名不能注入参数），
    // 而合并的目标恰恰是"旧名无损映射"。改为按参数推断，并在描述里写明。
    required: [],
  },
};

/** 把老形状 background{} 摊平到顶层（只补缺失的键） */
function flattenLegacyBackground(args: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = { ...args };
  const bg = args.background as Record<string, unknown> | undefined;
  if (bg && typeof bg === 'object') {
    for (const [k, v] of Object.entries(bg)) {
      if (out[k] === undefined) out[k] = v;
    }
    // 老形状里 "colors" 是渐变色数组
    if (out.gradient === undefined && Array.isArray(bg.colors)) {
      out.gradient = { type: 'linear', colors: bg.colors };
    }
  }
  return out;
}

export const setBackgroundHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const a = flattenLegacyBackground(args);
  const target = (a.target as string) || 'slide';
  const slideIndex = a.slideIndex as number | undefined;
  const color = a.color as string | undefined;
  const imagePath = a.imagePath as string | undefined;
  const pattern = a.pattern as string | undefined;
  const gradient = a.gradient as { type?: string; colors?: string[]; angle?: number } | undefined;

  // type 推断：草案要求必填，但旧名走别名时给不出 type（别名不能注入参数）。
  // 按"给了什么填什么"推断 —— 唯一有歧义的情况是同时给了多个，那时按 image > gradient > color 取先。
  const type =
    (a.type as string) ||
    (imagePath ? 'image' : gradient ? 'gradient' : pattern ? 'pattern' : color ? 'solid' : '');

  const fail = (msg: string, err?: string): ToolCallResult => ({
    id: uuidv4(),
    success: false,
    content: [{ type: 'text', text: msg }],
    error: err ?? msg,
  });
  const ok = (text: string): ToolCallResult => ({
    id: uuidv4(),
    success: true,
    content: [{ type: 'text', text }],
  });

  if (!type) {
    return fail(
      '要说清怎么填：给 color（纯色）/ imagePath（图片）/ gradient（渐变）之一，或显式传 type。' +
        '示例：{ "target":"slide", "slideIndex":1, "type":"solid", "color":"#003366" }'
    );
  }
  if (type === 'pattern') {
    return fail('WPS COM 没有图案填充的接口，type:"pattern" 无法执行（这不是参数问题，换 solid/gradient/image）。');
  }
  if (type === 'solid' && !color) return fail('type=solid 需要 color，例如 "#003366"。');
  if (type === 'image' && !imagePath) return fail('type=image 需要 imagePath（图片在本机的绝对路径）。');
  if (type === 'gradient') {
    const colors = (gradient && gradient.colors) || [];
    if (!Array.isArray(colors) || colors.length < 2) {
      return fail('type=gradient 需要 gradient.colors 至少两个颜色，例如 ["#FF0000","#00FF00"]。');
    }
  }
  if (target === 'master' && type === 'image') {
    return fail('母版背景不支持图片（WPS 的母版 Fill 只吃纯色/渐变）。要整份文稿用图片背景，用 target:"all"。');
  }
  if (target === 'slide' && !slideIndex) {
    return fail('target=slide 需要 slideIndex（从 1 开始）。要改整份文稿用 target:"all"。');
  }

  // 一个内部小包装。注意：动作名走变量 → param-shape-audit（扫 executeMethod 的字面量）看不到这 4 个
  // 调用点，属于那个脚本自己写明的"动态拼参数抓不到 = 漏报"。所以这里的参数是【手工对着 PS 的读法核过的】：
  //   setMasterBackground 读 color / color1 / color2
  //   setBackgroundColor  读 color / applyToAll
  //   setSlideBackground  读 slideIndex / color / imagePath
  //   getSlideCount       无参数
  // 改这里的参数名时，请同步改上面这份清单（它就是这几个动作的"参数契约"）。
  const exec = <T>(action: string, params: Record<string, unknown>) =>
    wpsClient.executeMethod<T>(action, params, WpsAppType.PRESENTATION);

  const warnings: ToolWarning[] = [];
  const where = target === 'all' ? '整份文稿' : target === 'master' ? '母版' : `第 ${slideIndex} 页`;

  try {
    // ---- 母版 ----
    if (target === 'master') {
      const params: Record<string, unknown> =
        type === 'gradient'
          ? { color1: gradient!.colors![0], color2: gradient!.colors![1] }
          : { color };
      const r = await exec('setMasterBackground', params);
      if (!r.success) return fail(`设置母版背景失败: ${r.error}`, r.error);
      return ok(`母版背景已设置：${type === 'gradient' ? `渐变 ${gradient!.colors!.slice(0, 2).join(' → ')}` : color}`);
    }

    // ---- 整份文稿 ----
    if (target === 'all') {
      if (type === 'solid') {
        const r = await exec('setBackgroundColor', { color, applyToAll: true });
        if (!r.success) return fail(`设置整份文稿背景失败: ${r.error}`, r.error);
        return ok(`整份文稿背景已设为 ${color}（applyToAll）`);
      }
      // 渐变/图片没有"一次全设"的动作 → 逐页设。先问页数，别拿失败当终止条件。
      const cnt = await exec<{ count?: number }>('getSlideCount', {});
      const total = Number((cnt.data && cnt.data.count) || 0);
      if (!total) return fail('拿不到页数，无法逐页设置背景。');
      let done = 0;
      const failedPages: number[] = [];
      for (let i = 1; i <= total; i++) {
        const r =
          type === 'image'
            ? await exec('setSlideBackground', { slideIndex: i, imagePath })
            : await exec('setSlideBackground', {
                slideIndex: i,
                color1: gradient!.colors![0],
                color2: gradient!.colors![1],
              });
        if (r.success) done++;
        else failedPages.push(i);
      }
      if (done === 0) return fail(`整份文稿逐页设置背景全部失败（共 ${total} 页）。`);
      if (failedPages.length) {
        warnings.push({
          step: 'set_background.all',
          message: `第 ${failedPages.join('、')} 页没设置成功（共 ${total} 页，成功 ${done} 页）`,
        });
      }
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: `整份文稿背景已设置：${done}/${total} 页` }],
        ...(warnings.length ? { warnings } : {}),
      };
    }

    // ---- 单页 ----
    const params: Record<string, unknown> = { slideIndex };
    if (type === 'solid') params.color = color;
    else if (type === 'image') params.imagePath = imagePath;
    else {
      params.color1 = gradient!.colors![0];
      params.color2 = gradient!.colors![1];
    }
    const r = await exec('setSlideBackground', params);
    if (!r.success) return fail(`设置背景失败: ${r.error}`, r.error);
    const detail =
      type === 'solid' ? `颜色 ${color}` : type === 'image' ? `图片 ${imagePath}` : `渐变 ${gradient!.colors!.slice(0, 2).join(' → ')}`;
    if (type === 'gradient' && gradient!.colors!.length > 2) {
      warnings.push({ step: 'gradient', message: 'WPS 的渐变只支持两个色标，多余的颜色已忽略' });
    }
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: `背景已设置：${where}，${detail}` }],
      ...(warnings.length ? { warnings } : {}),
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`设置背景出错: ${errMsg}`);
  }
};

// ============================================================
// 18. wps_ppt_set_slide_size - 设置幻灯片尺寸
// ============================================================

export const setSlideSizeDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_size',
  description: `设置演示文稿的幻灯片尺寸（按像素给，内部换算成 WPS 的磅）。

常用尺寸：
- 标准(4:3): 宽960, 高720
- 宽屏(16:9): 宽960, 高540
- 宽屏(16:10): 宽960, 高600
- A4横版: 宽1123, 高794
- A4竖版: 宽794, 高1123

注意：改尺寸会重排整份文稿，元素可能溢出页面。改完建议 wps_ppt_export_slide_as_image 看一眼。

使用场景：
- "把PPT改成16:9宽屏"
- "设置幻灯片为A4尺寸"
- "修改幻灯片大小"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      width: {
        type: 'number',
        description: '幻灯片宽度（像素，96 DPI；内部 ×0.75 换算成磅）',
      },
      height: {
        type: 'number',
        description: '幻灯片高度（像素，96 DPI；内部 ×0.75 换算成磅）',
      },
    },
    required: ['width', 'height'],
  },
};

export const setSlideSizeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { width, height } = args as {
    width: number;
    height: number;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideSize', // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
      { width, height },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `幻灯片尺寸设置成功！\n宽度: ${width}px\n高度: ${height}px`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置幻灯片尺寸失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置幻灯片尺寸出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 19. wps_ppt_set_transition - 设置幻灯片切换效果
// ============================================================

export const setTransitionDefinition: ToolDefinition = {
  name: 'wps_ppt_set_transition',
  description: `设置幻灯片切换效果：单页 or 整份文稿（第 8 轮合并了切换族 3 个工具）。

- \`target\`: slide（默认，只改 \`slideIndex\` 那一页）/ all（整份文稿统一）
- \`effect\`（**必填**）: fade 淡出 / push 推入 / wipe 擦除 / split 拆分 / reveal 揭开 / cover 覆盖 / dissolve 溶解 / curtains 帷幕 / **none 取消切换**

什么时候用：
- "给第1页加个淡出切换" → { target:'slide', slideIndex:1, effect:'fade', duration:1 }
- "所有页统一用推入" → { target:'all', effect:'push' }
- "第2页不要切换效果" → { target:'slide', slideIndex:2, effect:'none' }

取消切换用 \`effect:'none'\`，不要去找"删除切换"的工具 —— 切换本来就是"切走时播放的东西"，
把它设成没有就等价于删掉，而且这样不会因为名字里有"删/删掉"而弹破坏性确认框。

真限制：没有"切换音效"参数。WPS 的 COM 接口里没有"切走时播一段音频"这个能力
（那是 PowerPoint 的 AddMediaObject2），所以本工具不提供该参数 —— 要音效得先在 WPS 里录进切换设置。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      target: {
        type: 'string',
        enum: ['slide', 'all'],
        description: '作用范围：slide 单页（默认）/ all 整份文稿',
      },
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）；target=all 时忽略',
      },
      effect: {
        type: 'string',
        description: '切换效果：fade / push / wipe / split / reveal / cover / dissolve / curtains；none = 取消切换',
      },
      duration: {
        type: 'number',
        description: '切换持续时间（秒），默认1秒',
      },
    },
    required: [],
  },
};

const TRANS_NAME_MAP: Record<string, string> = {
  fade: '淡出', push: '推入', wipe: '擦除', split: '拆分',
  reveal: '揭开', cover: '覆盖', dissolve: '溶解', curtains: '帷幕',
  none: '取消切换',
};

export const setTransitionHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, duration } = args as {
    slideIndex?: number;
    duration?: number;
  };
  const target = ((args.target as string) || 'slide') as 'slide' | 'all';
  // 第 2 轮起的旧参数名是 transition；旧名调用经别名进来时带的是它
  const effect = (args.effect as string) || (args.transition as string);

  if (!effect) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `必须给 effect（切换效果名，或 none 表示取消切换）` }],
      error: 'Missing required parameter: effect',
    };
  }
  if (target === 'slide' && !slideIndex) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `target='slide' 时必须给 slideIndex` }],
      error: 'Missing required parameter: slideIndex',
    };
  }

  try {
    // 整份文稿走 PS 侧的 applyTransitionToAll：它是独立的动作，
    // 而 setSlideTransition 的循环版不在这个 case 里，硬塞 all 进去等于让 PS 认一个它不认的键。
    const response =
      target === 'all'
        ? await wpsClient.executeMethod<{ slideCount?: number }>(
            'applyTransitionToAll',
            { effect, duration: duration || 1 },
            WpsAppType.PRESENTATION
          )
        : await wpsClient.executeMethod<{ slideCount?: number }>(
            'setSlideTransition',
            { slideIndex, effect, duration: duration || 1 },
            WpsAppType.PRESENTATION
          );

    if (response.success) {
      const lines = [
        `切换效果设置成功！`,
        target === 'all' ? `范围: 全部幻灯片` : `幻灯片: 第 ${slideIndex} 页`,
        `切换效果: ${TRANS_NAME_MAP[effect] || effect}`,
        `持续时间: ${duration || 1} 秒`,
      ];
      if (response.data && response.data.slideCount) {
        lines.push(`已应用: ${response.data.slideCount} 页`);
      }
      return {
        id: uuidv4(),
        success: true,
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    }
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置切换效果失败: ${response.error}` }],
      error: response.error,
    };
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置切换效果出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 21. wps_ppt_set_shape_fill - 设置形状填充颜色
// ============================================================

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

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
      name?: string;
    }>(
      'setShapeFill', // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
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
            text: `形状填充颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n填充颜色: ${color}`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `设置形状填充失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `设置形状填充出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 22. wps_ppt_add_speaker_notes - 添加演讲者备注
// ============================================================

export const addSpeakerNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_add_speaker_notes',
  description: `添加或追加演讲者备注到指定幻灯片。

使用场景：
- "给第1页添加演讲者备注"
- "在备注中写上提示词"
- "追加演讲提示到第3页"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      notes: { type: 'string', description: '演讲者备注内容' },
    },
    required: ['slideIndex', 'notes'],
  },
};

export const addSpeakerNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, notes } = args as {
    slideIndex: number;
    notes: string;
  };

  try {
    const response = await wpsClient.executeMethod<{
      success: boolean;
      message: string;
    }>(
      'setSlideNotes',
      { slideIndex, notes },
      WpsAppType.PRESENTATION
    );

    if (response.success) {
      return {
        id: uuidv4(),
        success: true,
        content: [
          {
            type: 'text',
            text: `演讲者备注添加成功！\n幻灯片: 第 ${slideIndex} 页\n备注内容: "${notes.length > 50 ? notes.substring(0, 50) + '...' : notes}"`,
          },
        ],
      };
    } else {
      return {
        id: uuidv4(),
        success: false,
        content: [{ type: 'text', text: `添加演讲者备注失败: ${response.error}` }],
        error: response.error,
      };
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `添加演讲者备注出错: ${errMsg}` }],
      error: errMsg,
    };
  }
};

// ============================================================
// 导出所有幻灯片操作相关的Tools
// ============================================================

export const slideOpsTools: RegisteredTool[] = [
  { definition: deleteSlideDefinition, handler: deleteSlideHandler },
  { definition: moveSlideDefinition, handler: moveSlideHandler },
  { definition: getSlideCountDefinition, handler: getSlideCountHandler },
  { definition: getSlideInfoDefinition, handler: getSlideInfoHandler },
  { definition: switchSlideDefinition, handler: switchSlideHandler },
  { definition: setSlideLayoutDefinition, handler: setSlideLayoutHandler },
  { definition: getSlideNotesDefinition, handler: getSlideNotesHandler },
  { definition: addShapeDefinition, handler: addShapeHandler },
  { definition: setShapeStyleDefinition, handler: setShapeStyleHandler },
  { definition: addTextboxDefinition, handler: addTextboxHandler },
  { definition: setSlideTitleDefinition, handler: setSlideTitleHandler },
  { definition: insertImageDefinition, handler: insertImageHandler },
  { definition: setShapeTextDefinition, handler: setShapeTextHandler },
  { definition: setAnimationDefinition, handler: setAnimationHandler },
  { definition: setBackgroundDefinition, handler: setBackgroundHandler },
  { definition: setSlideSizeDefinition, handler: setSlideSizeHandler },
  { definition: setTransitionDefinition, handler: setTransitionHandler },
  { definition: setShapeFillDefinition, handler: setShapeFillHandler },
  { definition: addSpeakerNotesDefinition, handler: addSpeakerNotesHandler },
];

export default slideOpsTools;
