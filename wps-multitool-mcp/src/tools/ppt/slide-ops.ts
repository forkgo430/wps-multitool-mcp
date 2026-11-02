/**
 * Input: 幻灯片页操作参数
 * Output: 幻灯片页操作结果
 * Pos: PPT 幻灯片页操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT幻灯片页操作Tools（风格统一+拆分轮后本文件只留"页"这一群，见 ppt-audit §三之二）：
 *
 * 包含（8 个；形状/文本框 6 个迁 shape-basic.ts，动画/背景/切换/备注迁 page-effects.ts，
 * insert_image 迁 image.ts）：
 * - wps_ppt_delete_slide: 删除指定幻灯片
 * - wps_ppt_move_slide: 移动幻灯片到指定位置
 * - wps_ppt_get_slide_count: 获取幻灯片总数
 * - wps_ppt_get_slide_info: 获取指定幻灯片的详细信息
 * - wps_ppt_switch_slide: 切换到指定幻灯片
 * - wps_ppt_set_slide_layout: 设置幻灯片版式布局
 * - wps_ppt_get_slide_notes: 获取幻灯片备注内容
 * - wps_ppt_set_slide_size: 设置幻灯片尺寸
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText } from './helpers';

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

  const r = await call('deleteSlide', { slideIndex });
  if (!r.ok) return fail(errText('删除幻灯片失败', r), r.error);
  return ok(`幻灯片删除成功！\n已删除: 第 ${slideIndex} 页`, r.warnings);
};

// ============================================================
// 2. wps_ppt_move_slide - 移动幻灯片到指定位置
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

  const r = await call('moveSlide', { fromIndex, toIndex });
  if (!r.ok) return fail(errText('移动幻灯片失败', r), r.error);
  return ok(`幻灯片移动成功！\n从第 ${fromIndex} 页 → 移到第 ${toIndex} 页`, r.warnings);
};

// ============================================================
// 3. wps_ppt_get_slide_count - 获取幻灯片总数
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
  const r = await call<{ count: number }>('getSlideCount', {});
  if (!r.ok || !r.data) return fail(errText('获取幻灯片总数失败', r), r.error);
  return ok(`当前演示文稿共有 ${r.data.count} 页幻灯片`, r.warnings);
};

// ============================================================
// 4. wps_ppt_get_slide_info - 获取指定幻灯片的详细信息
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

  const r = await call<{
    slideIndex: number;
    layout: string;
    shapesCount: number;
    shapes: Array<{
      name: string;
      type: string;
      text?: string;
      // ps1 特意回传的两个排障字段：placeholderType 区分占位符种类（setSlideSubtitle/
      // setSlideContent 按 PlaceholderFormat.Type 找目标，报"placeholder not found"时靠它判断
      // 该版式有没有这种占位符）；hasText 区分"有文本框没文字"。
      placeholderType?: string;
      hasText?: boolean;
    }>;
  }>('getSlideInfo', { slideIndex });
  if (!r.ok || !r.data) return fail(errText('获取幻灯片信息失败', r), r.error);
  const info = r.data;
  let output = `第 ${info.slideIndex} 页幻灯片信息：\n`;
  output += `布局: ${info.layout}\n`;
  output += `元素数量: ${info.shapesCount}\n`;

  if (info.shapes && info.shapes.length > 0) {
    output += `\n元素列表：\n`;
    info.shapes.forEach((shape, i) => {
      output += `  ${i + 1}. [${shape.type}] ${shape.name}`;
      if (shape.placeholderType) {
        output += ` 占位符类型:${shape.placeholderType}`;
      }
      if (shape.hasText !== undefined) {
        output += shape.hasText ? ' 有文本框' : ' 无文本框';
      }
      if (shape.text) {
        output += ` - "${shape.text.substring(0, 50)}${shape.text.length > 50 ? '...' : ''}"`;
      }
      output += '\n';
    });
  }

  return ok(output, r.warnings);
};

// ============================================================
// 5. wps_ppt_switch_slide - 切换到指定幻灯片
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

  const r = await call('switchSlide', { slideIndex });
  if (!r.ok) return fail(errText('切换幻灯片失败', r), r.error);
  return ok(`已切换到第 ${slideIndex} 页幻灯片`, r.warnings);
};

// ============================================================
// 6. wps_ppt_set_slide_layout - 设置幻灯片版式布局
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

  const r = await call('setSlideLayout', { slideIndex, layout });
  if (!r.ok) return fail(errText('设置幻灯片布局失败', r), r.error);
  const layoutName: Record<string, string> = {
    title: '标题页',
    title_content: '标题+内容',
    blank: '空白页',
    two_column: '两栏内容',
    comparison: '对比布局',
    section_header: '节标题',
    title_only: '仅标题',
  };

  return ok(
    `幻灯片布局设置成功！\n幻灯片: 第 ${slideIndex} 页\n布局: ${layoutName[layout] || layout}`,
    r.warnings
  );
};

// ============================================================
// 7. wps_ppt_get_slide_notes - 获取幻灯片备注内容
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

  const r = await call<{ notes: string }>('getSlideNotes', { slideIndex });
  if (!r.ok || !r.data) return fail(errText('获取幻灯片备注失败', r), r.error);
  const notes = r.data.notes;
  return ok(
    notes
      ? `第 ${slideIndex} 页备注内容：\n${notes}`
      : `第 ${slideIndex} 页没有备注内容`,
    r.warnings
  );
};

// ============================================================
// 8. wps_ppt_set_slide_size - 设置幻灯片尺寸
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

  const r = await call(
    // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
    'setSlideSize',
    { width, height }
  );
  if (!r.ok) return fail(errText('设置幻灯片尺寸失败', r), r.error);
  return ok(`幻灯片尺寸设置成功！\n宽度: ${width}px\n高度: ${height}px`, r.warnings);
};

// ============================================================
// 导出所有幻灯片页操作相关的Tools
// ============================================================

export const slideOpsTools: RegisteredTool[] = [
  { definition: deleteSlideDefinition, handler: deleteSlideHandler },
  { definition: moveSlideDefinition, handler: moveSlideHandler },
  { definition: getSlideCountDefinition, handler: getSlideCountHandler },
  { definition: getSlideInfoDefinition, handler: getSlideInfoHandler },
  { definition: switchSlideDefinition, handler: switchSlideHandler },
  { definition: setSlideLayoutDefinition, handler: setSlideLayoutHandler },
  { definition: getSlideNotesDefinition, handler: getSlideNotesHandler },
  { definition: setSlideSizeDefinition, handler: setSlideSizeHandler },
];

export default slideOpsTools;
