/**
 * Input: PPT 高级美化操作参数
 * Output: 高级美化操作结果
 * Pos: PPT 高级美化工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT高级美化Tools - 高级美化模块
 * 处理配色方案、自动美化、KPI卡片、装饰元素等高级美化操作
 *
 * 包含（5 个；set_background_gradient 已并入 set_background 第 5 轮；auto_beautify_slide/beautify_all_slides 已于 2026-11-04 op-merge 进 wps_ppt_beautify 并转别名）：
 * - wps_ppt_apply_color_scheme: 应用配色方案
 * - wps_ppt_auto_beautify_slide: 已转别名指向 beautify（op:'auto'，2026-11-04 大项③）
 * - wps_ppt_beautify_all_slides: 已转别名指向 beautify（op:'style'，2026-11-04 大项③）
 * - wps_ppt_create_kpi_cards: 创建KPI指标卡片
 * - wps_ppt_create_styled_table: 创建带样式的表格
 * - wps_ppt_add_title_decoration: 添加标题装饰元素
 * - wps_ppt_add_page_indicator: 添加页码指示器
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
 * 应用配色方案
 * 为幻灯片或全文档应用统一配色方案
 */
export const applyColorSchemeDefinition: ToolDefinition = {
  name: 'wps_ppt_apply_color_scheme',
  description: `为幻灯片应用统一配色方案。

支持的配色方案：
- ocean: 海洋蓝（#0078D4主色）
- forest: 森林绿（#107C10主色）
- sunset: 日落橙（#FF8C00主色）
- lavender: 薰衣草紫（#7B68EE主色）
- corporate: 企业蓝灰（#2B579A主色）
- warm: 暖色调（#E74C3C主色）
- cool: 冷色调（#3498DB主色）
- monochrome: 单色黑白灰

使用场景：
- "给PPT换个海洋蓝配色"
- "用企业风格的配色方案"
- "统一所有页面的颜色"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始），不填则应用到所有页',
      },
      scheme: {
        type: 'string',
        description: '配色方案名称',
      },
    },
    required: ['scheme'],
  },
};

export const applyColorSchemeHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, scheme } = args as {
    slideIndex?: number;
    scheme: string;
  };

  const r = await call<{ appliedSlides: number }>('applyColorScheme', { slideIndex, scheme });
  if (!r.ok || !r.data) return fail(errText('应用配色方案失败', r), r.error);
  const schemeNames: Record<string, string> = {
    ocean: '海洋蓝',
    forest: '森林绿',
    sunset: '日落橙',
    lavender: '薰衣草紫',
    corporate: '企业蓝灰',
    warm: '暖色调',
    cool: '冷色调',
    monochrome: '单色黑白灰',
  };

  return ok(
    `配色方案应用成功！\n方案: ${schemeNames[scheme] || scheme}\n范围: ${slideIndex ? `第 ${slideIndex} 页` : '所有幻灯片'}\n已应用页数: ${r.data.appliedSlides}`,
    r.warnings
  );
};

// wps_ppt_auto_beautify_slide / wps_ppt_beautify_all_slides 已于 2026-11-04（大项③ op-merge）
// 转别名指向 wps_ppt_beautify（op:'auto' / op:'style'），旧定义删除——能力在 canonical 的
// op 三分支里逐字保留（含 operations 数组兜底与风格名映射），mock 断言在 deckkit-test? 不，
// 在 slide.ts 侧由 alias-verify 与旧调用逐项一致判据覆盖。

/**
 * 创建KPI指标卡片
 * 在幻灯片上生成数据指标卡片，适合展示关键业绩数据
 */
export const createKpiCardsDefinition: ToolDefinition = {
  name: 'wps_ppt_create_kpi_cards',
  description: `在幻灯片上创建KPI指标卡片。

每个KPI卡片数据项包含：
- title: 指标名称（如"销售额"、"增长率"）
- value: 指标数值（如"¥128万"、"+15.3%"）
- trend: 趋势方向（up/down/flat），可选
- color: 卡片颜色（十六进制），可选

使用场景：
- "创建一组KPI展示卡片"
- "在第2页展示销售数据指标"
- "添加业绩看板"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      data: {
        type: 'array',
        description: 'KPI数据数组，每项包含title（指标名称）、value（指标数值）、trend（趋势:up/down/flat，可选）、color（颜色，可选）',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string', description: '指标名称' },
            value: { type: 'string', description: '指标数值' },
            trend: { type: 'string', description: '趋势方向', enum: ['up', 'down', 'flat'] },
            color: { type: 'string', description: '卡片颜色（十六进制）' },
          },
          required: ['title', 'value'],
        },
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createKpiCardsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data } = args as {
    slideIndex: number;
    data: Array<{
      title: string;
      value: string;
      trend?: string;
      color?: string;
    }>;
  };

  const r = await call<{
    // PS createKpiCards 回的是 count；cardCount 是历史键，两个都接（cardCount 优先）。
    count?: number;
    cardCount?: number;
    cards?: string[];
  }>('createKpiCards', { slideIndex, data });
  if (!r.ok || !r.data) return fail(errText('创建KPI卡片失败', r), r.error);
  const n = r.data.cardCount ?? r.data.count ?? 0;
  let output = `KPI卡片创建成功！\n幻灯片: 第 ${slideIndex} 页\n卡片数量: ${n} 个\n\n指标列表：\n`;
  data.forEach((item) => {
    const trendIcon = item.trend === 'up' ? '↑' : item.trend === 'down' ? '↓' : '→';
    output += `- ${item.title}: ${item.value} ${item.trend ? trendIcon : ''}\n`;
  });

  return ok(output, r.warnings);
};

/**
 * 创建带样式的表格
 * 插入预设样式的精美表格
 */
export const createStyledTableDefinition: ToolDefinition = {
  name: 'wps_ppt_create_styled_table',
  description: `在幻灯片中创建带预设样式的精美表格。

表格数据格式：二维数组，第一行为表头。
例如: [["姓名","分数"],["张三","95"],["李四","88"]]

支持的样式：
- blue: 蓝色主题（默认）
- green: 绿色主题
- orange: 橙色主题
- gray: 灰色主题
- dark: 深色主题
- striped: 斑马纹样式

使用场景：
- "创建一个蓝色主题的数据表格"
- "插入带样式的表格"
- "在第3页添加成绩表"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      data: {
        type: 'array',
        description: '表格数据，二维数组格式，第一行为表头',
        items: {
          type: 'array',
          items: { type: 'string' },
        },
      },
      style: {
        type: 'string',
        description: '表格样式，不填默认使用 blue',
      },
    },
    required: ['slideIndex', 'data'],
  },
};

export const createStyledTableHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, data, style } = args as {
    slideIndex: number;
    data: string[][];
    style?: string;
  };

  const r = await call<{ rows: number; cols: number }>(
    'createStyledTable',
    { slideIndex, data, style: style || 'blue' }
  );
  if (!r.ok || !r.data) return fail(errText('创建带样式表格失败', r), r.error);
  const styleNames: Record<string, string> = {
    blue: '蓝色主题',
    green: '绿色主题',
    orange: '橙色主题',
    gray: '灰色主题',
    dark: '深色主题',
    striped: '斑马纹',
  };

  return ok(
    `带样式表格创建成功！\n幻灯片: 第 ${slideIndex} 页\n规格: ${r.data.rows} 行 × ${r.data.cols} 列\n样式: ${styleNames[style || 'blue'] || style}`,
    r.warnings
  );
};

/**
 * 添加标题装饰元素
 * 为幻灯片标题添加装饰性元素（下划线、色块、图标等）
 */
export const addTitleDecorationDefinition: ToolDefinition = {
  name: 'wps_ppt_add_title_decoration',
  description: `为幻灯片标题添加装饰性元素。

支持的装饰样式：
- underline: 标题下划线装饰
- block: 色块背景装饰
- sidebar: 侧边竖条装饰
- bracket: 括号框装饰
- gradient_bar: 渐变条装饰

使用场景：
- "给标题加个装饰线"
- "美化标题样式"
- "添加侧边装饰条"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      style: {
        type: 'string',
        description: '装饰样式，不填默认使用 underline',
      },
    },
    required: ['slideIndex'],
  },
};

export const addTitleDecorationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, style } = args as {
    slideIndex: number;
    style?: string;
  };

  const r = await call('addTitleDecoration', { slideIndex, style: style || 'underline' });
  if (!r.ok) return fail(errText('添加标题装饰失败', r), r.error);
  const styleNames: Record<string, string> = {
    underline: '下划线装饰',
    block: '色块背景装饰',
    sidebar: '侧边竖条装饰',
    bracket: '括号框装饰',
    gradient_bar: '渐变条装饰',
  };

  return ok(
    `标题装饰添加成功！\n幻灯片: 第 ${slideIndex} 页\n装饰样式: ${styleNames[style || 'underline'] || style}`,
    r.warnings
  );
};

/**
 * 添加页码指示器
 * 在幻灯片上添加页码/进度指示器
 */
export const addPageIndicatorDefinition: ToolDefinition = {
  name: 'wps_ppt_add_page_indicator',
  description: `为演示文稿添加页码指示器。

支持的位置：
- bottom-right: 右下角（默认）
- bottom-center: 底部居中
- bottom-left: 左下角

使用场景：
- "给PPT加上页码"
- "在右下角添加页码指示"
- "添加页面进度指示器"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      position: {
        type: 'string',
        description: '页码位置',
        enum: ['bottom-right', 'bottom-center', 'bottom-left'],
      },
    },
    required: [],
  },
};

export const addPageIndicatorHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { position } = args as {
    position?: string;
  };

  const r = await call<{ slideCount: number }>('addPageIndicator', { position: position || 'bottom-right' });
  if (!r.ok || !r.data) return fail(errText('添加页码指示器失败', r), r.error);
  const posNames: Record<string, string> = {
    'bottom-right': '右下角',
    'bottom-center': '底部居中',
    'bottom-left': '左下角',
  };

  return ok(
    `页码指示器添加成功！\n位置: ${posNames[position || 'bottom-right'] || position}\n已添加页数: ${r.data.slideCount} 页`,
    r.warnings
  );
};

// 第 5 轮：wps_ppt_set_background_gradient 并入 wps_ppt_set_background
// （target:'slide'|'all'|'master' + type:'gradient' + gradient{colors}）。
// 旧名留作别名：它的参数 {slideIndex, gradient} 能被新工具无损接受。

/**
 * 导出所有高级美化相关的Tools
 */
export const beautifyAdvancedTools: RegisteredTool[] = [
  { definition: applyColorSchemeDefinition, handler: applyColorSchemeHandler },
  { definition: createKpiCardsDefinition, handler: createKpiCardsHandler },
  { definition: createStyledTableDefinition, handler: createStyledTableHandler },
  { definition: addTitleDecorationDefinition, handler: addTitleDecorationHandler },
  { definition: addPageIndicatorDefinition, handler: addPageIndicatorHandler },
];

export default beautifyAdvancedTools;
