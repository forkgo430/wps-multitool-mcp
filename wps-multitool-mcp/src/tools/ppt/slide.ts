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
 * - wps_ppt_align_objects: 已转别名指向 align_shapes（2026-11-04 大项②裁决砍值）
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

  const r = await call<{ slideIndex: number; layout: string }>(
    'addSlide',
    { layout: layout || 'title_content', position, title, content }
  );
  if (!r.ok || !r.data) return fail(errText('添加幻灯片失败', r), r.error);
  const layoutName = {
    title: '标题页',
    title_content: '标题+内容',
    blank: '空白页',
    two_column: '两栏内容',
    comparison: '对比布局',
  }[r.data.layout] || r.data.layout;

  return ok(
    `幻灯片添加成功！\n位置: 第 ${r.data.slideIndex} 页\n布局: ${layoutName}${title ? `\n标题: ${title}` : ''}`,
    r.warnings
  );
};

/**
 * 美化幻灯片
 * 一键优化排版、配色、字体和间距
 */
export const beautifyDefinition: ToolDefinition = {
  name: 'wps_ppt_beautify',
  description: `美化幻灯片（合并轮后是美化族的唯一入口；op 三分支，覆盖原 beautify / auto_beautify_slide / beautify_all_slides 三工具）。

**op 三分支**：
- \`op:'quick'\`（默认）：快速美化——统一字体 + 配色，可单页（slideIndex）或全篇（beautifyAll:true）
- \`op:'auto'\`：单页深度优化——字体/间距/对齐/配色/布局/字号逐项处理（原 auto_beautify_slide）
- \`op:'style'\`：全篇风格统一——business/tech/creative/minimal/academic 五种（原 beautify_all_slides）

支持的配色方案（quick）：
- business: 商务风（深蓝+灰色）
- tech: 科技风（蓝色+绿色）
- creative: 创意风（珊瑚红+金色）
- minimal: 简约风（黑白灰）

quick【实际会做】的操作：
- 统一字体（改所有有文字的文本框的字体名）
- 应用配色方案（按字号分标题/正文上色）

注意：quick 不做自动对齐，也不动行距/间距 —— 这些改起来没有唯一正确答案，
一键"优化"很容易把用户排版搞乱。要精确调版面用 align_shapes / set_shape_position；
要深度排版优化用 op:'auto'。

使用场景：
- "美化这页PPT" → { op:'quick', slideIndex:1 }
- "用商务风格优化一下" → { op:'quick', slideIndex:1, colorScheme:'business' }
- "这页深度优化一下" → { op:'auto', slideIndex:1 }
- "整个PPT统一成商务风" → { op:'style', style:'business' }`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      op: {
        type: 'string',
        enum: ['quick', 'auto', 'style'],
        description: '美化方式：quick 快速（默认）/ auto 单页深度 / style 全篇风格',
      },
      slideIndex: {
        type: 'number',
        description: '要美化的幻灯片页码（quick：不填美化当前页；auto：必填；style：忽略）',
      },
      colorScheme: {
        type: 'string',
        description: '配色方案（quick）',
        enum: ['business', 'tech', 'creative', 'minimal'],
      },
      font: {
        type: 'string',
        description: '统一使用的字体，如 "微软雅黑"、"思源黑体"（quick）',
      },
      beautifyAll: {
        type: 'boolean',
        description: '是否美化所有幻灯片，默认false只美化指定页（quick；老参数，与 op:"style" 等效但风格来自 colorScheme）',
      },
      style: {
        type: 'string',
        description: '全篇风格（style）：business/tech/creative/minimal/academic，不填默认 business',
      },
    },
    required: [],
  },
};

export const beautifyHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, colorScheme, font, beautifyAll, style } = args as {
    slideIndex?: number;
    colorScheme?: string;
    font?: string;
    beautifyAll?: boolean;
    style?: string;
  };
  const op = (args.op as string) || 'quick';

  // ---- op:'auto' —— 原 wps_ppt_auto_beautify_slide（单页深度优化，合并轮 B 段吸收）----
  if (op === 'auto') {
    if (!slideIndex) {
      return fail(`op='auto' 必须给 slideIndex（要深度优化的页码）`, 'Missing required parameter: slideIndex');
    }
    const ra = await call<{
      operations?: Array<{ operation: string; count: number; details?: string }>;
    }>('autoBeautifySlide', { slideIndex });
    if (!ra.ok || !ra.data) return fail(errText('自动美化失败', ra), ra.error);
    let outputAuto = `幻灯片自动美化完成！\n幻灯片: 第 ${slideIndex} 页\n\n优化详情：\n`;

    const autoOpNames: Record<string, string> = {
      unify_font: '统一字体',
      optimize_spacing: '优化间距',
      align_elements: '对齐元素',
      adjust_colors: '调整配色',
      fix_layout: '修正布局',
      adjust_text_size: '调整文字大小',
    };

    // PS 侧只回 {style,count}（老版本）时 operations 不存在，直接 .forEach 会把整个工具打崩
    // —— 与 wps_ppt_beautify 同一个坑，那边已经兜过，这里漏了。
    const autoOps = Array.isArray(ra.data.operations) ? ra.data.operations : [];
    if (autoOps.length === 0) {
      outputAuto += '- 没有可优化的文本（本页可能是空占位符，或元素都没有文字）\n';
    }
    autoOps.forEach((opItem) => {
      outputAuto += `- ${autoOpNames[opItem.operation] || opItem.operation}: 处理了 ${opItem.count} 个元素${opItem.details ? ` (${opItem.details})` : ''}\n`;
    });

    return ok(outputAuto, ra.warnings);
  }

  // ---- op:'style' —— 原 wps_ppt_beautify_all_slides（全篇风格统一，合并轮 B 段吸收）----
  if (op === 'style') {
    const rs = await call<{ slideCount: number; style: string }>(
      'beautifyAllSlides',
      { style: style || 'business' }
    );
    if (!rs.ok || !rs.data) return fail(errText('批量美化失败', rs), rs.error);
    const styleNames: Record<string, string> = {
      business: '商务风',
      tech: '科技风',
      creative: '创意风',
      minimal: '极简风',
      academic: '学术风',
    };

    return ok(
      `全部幻灯片美化完成！\n风格: ${styleNames[rs.data.style] || rs.data.style}\n处理页数: ${rs.data.slideCount} 页`,
      rs.warnings
    );
  }

  // ---- op:'quick'（默认）—— 原主路径：统一字体 + 配色 ----
  const r = await call<{
    slideIndex: number | string;
    operations?: Array<{ operation: string; count: number; details?: string }>;
  }>(
    'beautifySlide',
    {
      slideIndex: beautifyAll ? 'all' : slideIndex,
      style: {
        colorScheme: colorScheme || 'business',
        font: font || '微软雅黑',
      },
    }
  );
  if (!r.ok || !r.data) return fail(errText('美化幻灯片失败', r), r.error);
  const result = r.data;
  const schemeName = {
    business: '商务风',
    tech: '科技风',
    creative: '创意风',
    minimal: '简约风',
  }[colorScheme || 'business'] || colorScheme;

  let output = `幻灯片美化完成！\n`;
  output += `范围: ${beautifyAll ? '所有幻灯片' : `第 ${result.slideIndex} 页`}\n`;
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

  return ok(output, r.warnings);
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
      fontName: {
        type: 'string',
        description: '要统一使用的字体名称，如 "微软雅黑"、"思源黑体"',
      },
      slideIndex: {
        type: 'number',
        description: '只处理指定页，不填则处理所有页',
      },
      includeTitle: {
        type: 'boolean',
        description: '是否包含标题，默认true',
      },
      includeBody: {
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
  const { fontName, slideIndex, includeTitle, includeBody } = args as {
    fontName: string;
    slideIndex?: number;
    includeTitle?: boolean;
    includeBody?: boolean;
  };

  const r = await call<{ font: string; count: number; slideCount: number }>(
    // 调 slide.unifyFont 而不是 unifyFont（第 3 轮修）：
    // PS 里 unifyFont 这个 case 会遍历【所有页、所有文本框】，压根不读
    // slideIndex / includeTitle / includeBody —— 工具描述承诺的三个开关全是摆设。
    // 认这三个参数的是另一个 case：slide.unifyFont。
    'slide.unifyFont',
    {
      fontName: fontName,
      slideIndex: slideIndex,
      includeTitle: includeTitle !== false,
      includeBody: includeBody !== false,
    }
  );
  if (!r.ok || !r.data) return fail(errText('统一字体失败', r), r.error);
  return ok(
    `字体统一完成！\n字体: ${r.data.font}\n处理幻灯片: ${r.data.slideCount} 页\n修改文本框: ${r.data.count} 个`,
    r.warnings
  );
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

  const r = await call(
    // 第 1 期第 2 章第 2 轮：这句原为「仅Windows支持」，是假注释 —— 该动作曾长期不存在于 ps1。第 6 轮已补 case，现在 Windows 上真能用。
    'setFontColor',
    { slideIndex, shapeIndex, color }
  );
  if (!r.ok) return fail(errText('设置文字颜色失败', r), r.error);
  return ok(
    `文字颜色设置成功！\n幻灯片: 第 ${slideIndex} 页\n形状: 第 ${shapeIndex} 个\n颜色: ${color}`,
    r.warnings
  );
};

// wps_ppt_align_objects 已于 2026-11-04（大项②裁决=砍值）转别名指向 wps_ppt_align_shapes：
// 参数同名零映射。裁决依据：align_objects schema 里的 distribute_h/v 两值从未正确工作——
// ps1 alignShapes 的 alignMap 只有 6 个对齐值，未知值兜底成左对齐（真机：success 但坐标零变化）。
// 转别名后老调用传 distribute_h/v 会拿到清晰的 enum 校验错误；分布需求用 wps_ppt_distribute_shapes。

/**
 * 导出所有幻灯片相关的Tools
 */
export const slideTools: RegisteredTool[] = [
  { definition: addSlideDefinition, handler: addSlideHandler },
  { definition: beautifyDefinition, handler: beautifyHandler },
  { definition: unifyFontDefinition, handler: unifyFontHandler },
  { definition: setFontColorDefinition, handler: setFontColorHandler },
];

export default slideTools;
