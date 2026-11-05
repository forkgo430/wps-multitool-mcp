/**
 * Input: 文本框与标题工具参数
 * Output: 文本框操作结果
 * Pos: PPT 文本框与标题工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
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
// 1. wps_ppt_delete_textbox - 删除文本框
// ============================================================

export const deleteTextboxDefinition: ToolDefinition = {
  name: 'wps_ppt_delete_textbox',
  description: '删除幻灯片上指定的文本框。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
    },
    required: ['slideIndex', 'textboxIndex'],
  },
};

export const deleteTextboxHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex } = args as { slideIndex: number; textboxIndex: number };
  // 桥侧按扁平 shapeIndex 定位（读 $p.shapeIndex），此前发 textboxIndex 桥读不到，工具恒败（E3 顺修）
  const r = await call('deleteTextBox', { slideIndex, shapeIndex: textboxIndex });
  if (!r.ok) return fail(errText('删除文本框失败', r), r.error);
  return ok(`第${slideIndex}页的文本框${textboxIndex}已删除`, r.warnings);
};

// ============================================================
// 2. wps_ppt_get_textboxes - 获取文本框列表
// ============================================================

export const getTextboxesDefinition: ToolDefinition = {
  name: 'wps_ppt_get_textboxes',
  description: '获取幻灯片上所有文本框的列表。brief 模式（默认开，回包瘦身）只回 index/名称/几何摘要；全量传 brief:false。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      brief: {
        type: 'boolean',
        description: '回包瘦身开关（体验修复轮 B），默认 true 只回 index/name/几何摘要；false 回全量字段',
      },
    },
    required: ['slideIndex'],
  },
};

export const getTextboxesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, brief = true } = args as { slideIndex: number; brief?: boolean };
  // PS 回的键是 textBoxes（大写 B），这里历史上读的是 textboxes → 永远拿到 undefined，
  // 打印出来永远是"第N页文本框: []"，读回形同虚设。两个键都接。
  const r = await call<{ textBoxes?: Array<Record<string, unknown>>; textboxes?: Array<Record<string, unknown>> }>('getTextBoxes', { slideIndex });
  if (!r.ok) return fail(errText('获取文本框列表失败', r), r.error);
  const rawList = (r.data?.textBoxes ?? r.data?.textboxes ?? []) as Array<Record<string, unknown>>;
  // brief 模式（体验修复轮 B）：只留 index/name/几何摘要，正文与详细属性瘦身掉——全量传 brief:false
  const list = brief
    ? rawList.map((b) => {
        const { text, ...rest } = b;
        return rest;
      })
    : rawList;
  return ok(`第${slideIndex}页文本框（${list.length} 个）:\n${JSON.stringify(list, null, 2)}`, r.warnings);
};

// ============================================================
// 3. wps_ppt_set_textbox_text - 设置文本框文本
// ============================================================

export const setTextboxTextDefinition: ToolDefinition = {
  name: 'wps_ppt_set_textbox_text',
  description: '设置指定文本框的文本内容。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
      text: { type: 'string', description: '文本内容' },
    },
    required: ['slideIndex', 'textboxIndex', 'text'],
  },
};

export const setTextboxTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex, text } = args as {
    slideIndex: number; textboxIndex: number; text: string;
  };
  // 桥侧按扁平 shapeIndex 定位（读 $p.shapeIndex），此前发 textboxIndex 桥读不到，工具恒败（E3 顺修）
  const r = await call('setTextBoxText', { slideIndex, shapeIndex: textboxIndex, text });
  if (!r.ok) return fail(errText('设置文本框内容失败', r), r.error);
  return ok(`文本框内容已更新`, r.warnings);
};

// ============================================================
// 4. wps_ppt_set_textbox_style - 设置文本框样式
// ============================================================

export const setTextboxStyleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_textbox_style',
  description: '设置文本框样式（字体大小、颜色、粗体等）。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      textboxIndex: { type: 'number', description: '文本框索引' },
      style: {
        type: 'object',
        description: '样式对象，包含 fontSize/fontColor/bold/italic/align 等',
      },
    },
    required: ['slideIndex', 'textboxIndex', 'style'],
  },
};

export const setTextboxStyleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, textboxIndex, style } = args as {
    slideIndex: number; textboxIndex: number; style: Record<string, unknown>;
  };
  // 桥侧读扁平 $p.shapeIndex/$p.fontSize/...（不是 textboxIndex、也不是 style 嵌套）——
  // 此前发 { textboxIndex, style:{...} } 桥全部读不到，工具恒败（E3 顺修，deckkit 缩字重排走这里）。
  // 键名映射：fontColor→color、align→alignment，其余同名直传。
  // style 缺省当空对象：注解生成器观测 handler 时传空参，这里一炸动作名就记不到（实测 288/289 的教训）。
  const s = (style ?? {}) as Record<string, unknown>;
  const wire: Record<string, unknown> = { slideIndex, shapeIndex: textboxIndex };
  if (s.fontSize !== undefined) wire.fontSize = s.fontSize;
  if (s.fontName !== undefined) wire.fontName = s.fontName;
  if (s.bold !== undefined) wire.bold = s.bold;
  if (s.italic !== undefined) wire.italic = s.italic;
  if (s.fontColor !== undefined) wire.color = s.fontColor;
  if (s.align !== undefined) wire.alignment = s.align;
  const r = await call('setTextBoxStyle', wire);
  if (!r.ok) return fail(errText('设置文本框样式失败', r), r.error);
  return ok(`文本框样式已更新`, r.warnings);
};

// ============================================================
// 5. wps_ppt_get_slide_title - 获取幻灯片标题
// ============================================================

export const getSlideTitleDefinition: ToolDefinition = {
  name: 'wps_ppt_get_slide_title',
  description: '获取指定幻灯片的标题。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
    },
    required: ['slideIndex'],
  },
};

export const getSlideTitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };
  const r = await call<{ title: string }>('getSlideTitle', { slideIndex });
  if (!r.ok) return fail(errText('获取标题失败', r), r.error);
  return ok(`第${slideIndex}页标题: ${r.data?.title || '无标题'}`, r.warnings);
};

// ============================================================
// 6. wps_ppt_set_slide_subtitle - 设置幻灯片副标题
// ============================================================

export const setSlideSubtitleDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_subtitle',
  description: '设置幻灯片的副标题。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      subtitle: { type: 'string', description: '副标题内容' },
    },
    required: ['slideIndex', 'subtitle'],
  },
};

export const setSlideSubtitleHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, subtitle } = args as { slideIndex: number; subtitle: string };
  const r = await call('setSlideSubtitle', { slideIndex, subtitle });
  if (!r.ok) return fail(errText('设置副标题失败', r), r.error);
  return ok(`第${slideIndex}页副标题已设置`, r.warnings);
};

// ============================================================
// 7. wps_ppt_set_slide_content - 设置幻灯片内容
// ============================================================

export const setSlideContentDefinition: ToolDefinition = {
  name: 'wps_ppt_set_slide_content',
  description: '设置幻灯片的正文内容区域。',
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      content: { type: 'string', description: '正文内容' },
    },
    required: ['slideIndex', 'content'],
  },
};

export const setSlideContentHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, content } = args as { slideIndex: number; content: string };
  const r = await call('setSlideContent', { slideIndex, content });
  if (!r.ok) return fail(errText('设置正文内容失败', r), r.error);
  return ok(`第${slideIndex}页正文内容已设置`, r.warnings);
};

/**
 * 导出所有文本框相关的Tools
 */
export const textboxTools: RegisteredTool[] = [
  { definition: deleteTextboxDefinition, handler: deleteTextboxHandler },
  { definition: getTextboxesDefinition, handler: getTextboxesHandler },
  { definition: setTextboxTextDefinition, handler: setTextboxTextHandler },
  { definition: setTextboxStyleDefinition, handler: setTextboxStyleHandler },
  { definition: getSlideTitleDefinition, handler: getSlideTitleHandler },
  { definition: setSlideSubtitleDefinition, handler: setSlideSubtitleHandler },
  { definition: setSlideContentDefinition, handler: setSlideContentHandler },
];

export default textboxTools;
