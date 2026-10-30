/**
 * Input: Word 长文档/学术排版类工具参数
 * Output: Word 工具注册数组
 * Pos: Word 长文档能力（标题层级/题注/脚注/图表目录/章节编号/分栏/列表/表格样式/刷新域/大纲）。
 *      一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么单独一组
 * --------------
 * 原有 32 个 word 工具覆盖的是"零散编辑"（插入文本、改字体、查找替换、页眉页脚…），
 * 缺的是【认真写一篇论文/报告】那条链路。对标 JulianPoleszczuk/office-mcp 的 40 个 doc_*，
 * 这里补上它多出来而我们没有的那部分：
 *   标题层级(add_heading) / 大纲(get_outline) / 段落级增删 / 列表
 *   题注(add_caption) / 脚注(add_footnote) / 图表目录(insert_table_of_figures)
 *   章节自动编号(set_heading_numbering) / 刷新域(update_fields) / 分栏(set_columns)
 *   默认字体(set_default_font) / 表格样式(format_table)
 * 已有的不重复造：导出 PDF 走 convertToPDF，页边距/纸张走 setPageSetup，字数统计走 getDocumentStats。
 *
 * 说明：本组工具全部走 WpsAppType.WRITER，桥会据此路由到 word 那一路常驻子进程。
 * 2026-10-30：收尾风格统一到 helpers（本文件的 toResult/errText 局部助手并入），
 * 并补上 warnings 透传——此前本组会把 PS 侧的警告丢掉。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { WORD_TARGET_PROPERTY } from './target';
import { call, ok, fail, errMsg } from './helpers';

// ==================== 1. 标题与大纲 ====================

export const addHeadingDefinition: ToolDefinition = {
  name: 'wps_word_add_heading',
  description: `按级别插入一个标题段落（使用 Word 内置 Heading 样式，因此能自动进目录）。

使用场景：
- "加一个一级标题：第一章 绪论"
- "这里插个二级标题"
- 写论文/报告时会先搭骨架

参数：
- text: 标题文字
- level: 级别 1-9，默认 1`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: { type: 'string', description: '标题文字' },
      level: { type: 'number', description: '标题级别 1-9，默认 1' },
    },
    required: ['text'],
  },
};

export const addHeadingHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ level: number; text: string }>(
      'addHeading',
      { text: String(args.text ?? ''), level: args.level ?? 1 }
    );
    if (!r.ok) return fail(`插入标题失败：${r.error}`, r.error);
    return ok(`已插入 ${r.data!.level} 级标题：${r.data!.text}`, r.warnings);
  } catch (e) {
    return fail(`插入标题异常：${errMsg(e)}`, errMsg(e));
  }
};

export const getOutlineDefinition: ToolDefinition = {
  name: 'wps_word_get_outline',
  description: `抽取文档的标题大纲（按段落的大纲级别 1-9），返回层级 + 段落序号。

使用场景：
- "这篇文档的目录结构是什么"
- "看看标题都是几级的"
- 改章节编号前先确认结构

参数：
- limit: 最多返回多少个标题，默认 500`,
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { target: WORD_TARGET_PROPERTY,  limit: { type: 'number', description: '最多返回标题数，默认 500' } } },
};

export const getOutlineHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ headings: Array<{ index: number; level: number; text: string }>; count: number }>(
      'getOutline',
      { limit: args.limit ?? 500 }
    );
    if (!r.ok) return fail(`抽取大纲失败：${r.error}`, r.error);
    const hs = r.data?.headings ?? [];
    if (!hs.length) return ok('文档里没有带大纲级别的标题（先用 add_heading 建立层级）', r.warnings);
    const lines = hs.map((h) => `${'  '.repeat(Math.max(0, h.level - 1))}${h.level}. ${h.text}   [段 ${h.index}]`);
    return ok(`大纲（共 ${hs.length} 个标题）：\n${lines.join('\n')}`, r.warnings);
  } catch (e) {
    return fail(`抽取大纲异常：${errMsg(e)}`, errMsg(e));
  }
};

// ==================== 2. 段落级编辑 ====================

export const insertParagraphDefinition: ToolDefinition = {
  name: 'wps_word_insert_paragraph',
  description: `插入一个正文段落。给了 index 就插在该段之前，不给就追加到文末。

参数：
- text: 段落文字
- index: 插到第几段之前（1 基），省略=追加到文末`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: { type: 'string', description: '段落文字' },
      index: { type: 'number', description: '插到第几段之前（1 基），省略=文末' },
    },
    required: ['text'],
  },
};

export const insertParagraphHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ paragraphs: number }>(
      'insertParagraph',
      { text: String(args.text ?? ''), index: args.index }
    );
    if (!r.ok) return fail(`插入段落失败：${r.error}`, r.error);
    return ok(`已插入段落，当前共 ${r.data?.paragraphs} 段`, r.warnings);
  } catch (e) {
    return fail(`插入段落异常：${errMsg(e)}`, errMsg(e));
  }
};

export const deleteParagraphDefinition: ToolDefinition = {
  name: 'wps_word_delete_paragraph',
  description: `删除段落。从 index 起删 count 段（1 基，越界自动停下并如实报告删了几段）。

参数：
- index: 起始段号，省略=最后一段
- count: 删除段数，默认 1`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      index: { type: 'number', description: '起始段号（1 基），省略=最后一段' },
      count: { type: 'number', description: '删除段数，默认 1' },
    },
  },
};

export const deleteParagraphHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ deleted: number; paragraphs: number }>(
      'deleteParagraph',
      { index: args.index, count: args.count ?? 1 }
    );
    if (!r.ok) return fail(`删除段落失败：${r.error}`, r.error);
    return ok(`已删除 ${r.data?.deleted} 段，当前共 ${r.data?.paragraphs} 段`, r.warnings);
  } catch (e) {
    return fail(`删除段落异常：${errMsg(e)}`, errMsg(e));
  }
};

// ==================== 3. 学术排版：题注 / 脚注 / 图表目录 ====================

export const addCaptionDefinition: ToolDefinition = {
  name: 'wps_word_add_caption',
  description: `在当前位置插入题注（自动编号，如"图 1""表 2"）。

实现说明：优先用 WPS 原生题注接口（生成活编号域，图表目录能自动收集）；
实测部分 WPS 版本（如本机 12.1）原生接口不可用，此时自动退化为【手工题注】——
按已有同标签题注数自增序号插入纯文本，序号不是域，图表目录收集不到，返回值里会明确标注。

使用场景：
- "给这张图加题注：系统架构图"
- "表格下面加个题注"
- 之后可用 insert_table_of_figures 生成图表目录

参数：
- text: 题注文字（编号是自动的，不用自己写）
- label: 标签，默认"图"（也可传"表"）
- position: above | below，默认 below
- newParagraph: 是否插完另起一段`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: { type: 'string', description: '题注文字' },
      label: { type: 'string', description: '标签，默认"图"' },
      position: { type: 'string', description: 'above | below，默认 below' },
      newParagraph: { type: 'boolean', description: '插完是否另起一段' },
    },
    required: ['text'],
  },
};

export const addCaptionHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ label: string; text: string; mode?: string; caption?: string; note?: string; nativeError?: string }>(
      'addCaption',
      {
        text: String(args.text ?? ''),
        label: args.label ?? '图',
        position: args.position ?? 'below',
        newParagraph: args.newParagraph ?? false,
      }
    );
    if (!r.ok) return fail(`插入题注失败：${r.error}`, r.error);
    const d = r.data!;
    const shown = d.caption ?? d.text ?? '';
    const text = d.mode === 'manual'
      ? `已插入题注（手工模式）：${shown}\n注意：${d.note ?? '原生题注接口不可用，已退化为文本'}`
      : `已插入题注：${shown}`;
    return ok(text, r.warnings);
  } catch (e) {
    return fail(`插入题注异常：${errMsg(e)}`, errMsg(e));
  }
};

export const addFootnoteDefinition: ToolDefinition = {
  name: 'wps_word_add_footnote',
  description: `插入脚注（页脚注释，自动编号）。

参数：
- text: 脚注内容
- atEnd: true=把脚注标记放在文末，默认放在当前光标处`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      text: { type: 'string', description: '脚注内容' },
      atEnd: { type: 'boolean', description: '标记是否放在文末（默认当前光标处）' },
    },
    required: ['text'],
  },
};

export const addFootnoteHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ total: number; index: number }>(
      'addFootnote',
      { text: String(args.text ?? ''), atEnd: args.atEnd ?? false }
    );
    if (!r.ok) return fail(`插入脚注失败：${r.error}`, r.error);
    return ok(`已插入脚注（文档现有 ${r.data?.total} 条脚注）`, r.warnings);
  } catch (e) {
    return fail(`插入脚注异常：${errMsg(e)}`, errMsg(e));
  }
};

export const insertTableOfFiguresDefinition: ToolDefinition = {
  name: 'wps_word_insert_table_of_figures',
  description: `在文末生成"图表目录"（基于题注标签；所以要先有 add_caption 加的题注）。

参数：
- label: 收集哪个标签，默认"图"（"表"则生成表格目录）`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY,  label: { type: 'string', description: '标签，默认"图"' } },
  },
};

export const insertTableOfFiguresHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ label: string; total: number }>('insertTableOfFigures', { label: args.label ?? '图' });
    if (!r.ok) return fail(`生成图表目录失败：${r.error}`, r.error);
    return ok(
      `已在文末生成"${r.data?.label}"目录（文档现有 ${r.data?.total} 个图表目录）。若内容为空，先给图/表加题注再 update_fields。`,
      r.warnings
    );
  } catch (e) {
    return fail(`生成图表目录异常：${errMsg(e)}`, errMsg(e));
  }
};

export const updateFieldsDefinition: ToolDefinition = {
  name: 'wps_word_update_fields',
  description: `刷新文档里的所有域：目录、图表目录、交叉引用、页码等。

使用场景：
- "目录页码不对了"
- "改完正文后刷新一下目录"
- 插完题目/题注后收尾`,
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { target: WORD_TARGET_PROPERTY } },
};

export const updateFieldsHandler: ToolHandler = async () => {
  try {
    const r = await call<{ fields: number; toc: number; tof: number; errors: string[] }>('updateFields', {});
    if (!r.ok) return fail(`刷新域失败：${r.error}`, r.error);
    const d = r.data;
    const errs = d?.errors?.length ? `\n部分项刷新报错：\n  - ${d.errors.join('\n  - ')}` : '';
    return ok(
      `已刷新：域 ${d?.fields ?? 0} 个，目录 ${d?.toc ?? 0} 个，图表目录 ${d?.tof ?? 0} 个${errs}`,
      r.warnings
    );
  } catch (e) {
    return fail(`刷新域异常：${errMsg(e)}`, errMsg(e));
  }
};

// ==================== 4. 版式：分栏 / 默认字体 / 章节编号 ====================

export const setColumnsDefinition: ToolDefinition = {
  name: 'wps_word_set_columns',
  description: `设置分栏（整节等宽多栏）。

参数：
- count: 栏数 1-6，默认 2
- spacing: 栏间距（厘米）
- lineBetween: 是否显示分隔线`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      count: { type: 'number', description: '栏数 1-6，默认 2' },
      spacing: { type: 'number', description: '栏间距（厘米）' },
      lineBetween: { type: 'boolean', description: '是否显示栏间分隔线' },
    },
  },
};

export const setColumnsHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ columns: number }>(
      'setColumns',
      { count: args.count ?? 2, spacing: args.spacing, lineBetween: args.lineBetween }
    );
    if (!r.ok) return fail(`设置分栏失败：${r.error}`, r.error);
    return ok(`已设为 ${r.data?.columns} 栏`, r.warnings);
  } catch (e) {
    return fail(`设置分栏异常：${errMsg(e)}`, errMsg(e));
  }
};

export const setDefaultFontDefinition: ToolDefinition = {
  name: 'wps_word_set_default_font',
  description: `设置文档默认字体（改 Normal 样式，正文与后续新建段落都会跟着变）。

参数：
- eastAsia: 中文字体名，如"宋体"（论文常用）
- ascii: 西文字体名，如"Times New Roman"
- size: 字号（磅）
- name: 直接指定字体名（同时作用于中西文时用）`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      eastAsia: { type: 'string', description: '中文字体名' },
      ascii: { type: 'string', description: '西文字体名' },
      name: { type: 'string', description: '字体名（通用）' },
      size: { type: 'number', description: '字号（磅）' },
    },
  },
};

export const setDefaultFontHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ name: string; eastAsia: string; size: number }>(
      'setDefaultFont',
      { eastAsia: args.eastAsia, ascii: args.ascii, name: args.name, size: args.size }
    );
    if (!r.ok) return fail(`设置默认字体失败：${r.error}`, r.error);
    const d = r.data!;
    return ok(`默认字体已设为：中文 ${d.eastAsia || '-'} / 西文 ${d.name || '-'}，字号 ${d.size || '-'}`, r.warnings);
  } catch (e) {
    return fail(`设置默认字体异常：${errMsg(e)}`, errMsg(e));
  }
};

export const setHeadingNumberingDefinition: ToolDefinition = {
  name: 'wps_word_set_heading_numbering',
  description: `给标题加章节自动编号（1. / 1.1. / 1.1.1. 这种）。

说明：这是【近似实现】，不同 WPS 版本的多级列表接口有差异，效果可能不一致；
失败会如实报错，不会假装成功。

参数：
- levels: 编到第几级，默认 3`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY,  levels: { type: 'number', description: '编到第几级，默认 3' } },
  },
};

export const setHeadingNumberingHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ levels: number; mode?: string }>('setHeadingNumbering', { levels: args.levels ?? 3 });
    if (!r.ok) return fail(`设置章节编号失败：${r.error}`, r.error);
    return ok(
      `已给 1-${r.data?.levels} 级标题挂上自动编号（用大纲编号库现成模板，方式：${r.data?.mode ?? 'unknown'}）。` +
        '若编号样式不合意，可在 WPS 里用「开始 → 多级列表」手动挑一个。',
      r.warnings
    );
  } catch (e) {
    return fail(`设置章节编号异常：${errMsg(e)}`, errMsg(e));
  }
};

// ==================== 5. 列表与表格 ====================

export const addListDefinition: ToolDefinition = {
  name: 'wps_word_add_list',
  description: `在当前位置插入项目符号或编号列表（一次写多条）。

参数：
- items: 列表项文字数组
- type: bullet（项目符号，默认）| number（编号）`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      items: { type: 'array', items: { type: 'string' }, description: '列表项' },
      type: { type: 'string', description: 'bullet | number，默认 bullet' },
    },
    required: ['items'],
  },
};

export const addListHandler: ToolHandler = async (args) => {
  try {
    const items = Array.isArray(args.items) ? (args.items as string[]) : [];
    const r = await call<{ type: string; count: number }>(
      'addList',
      { items, type: args.type ?? 'bullet' }
    );
    if (!r.ok) return fail(`插入列表失败：${r.error}`, r.error);
    return ok(`已插入 ${r.data?.count} 条${r.data?.type === 'number' ? '编号' : '项目符号'}列表`, r.warnings);
  } catch (e) {
    return fail(`插入列表异常：${errMsg(e)}`, errMsg(e));
  }
};

export const formatTableDefinition: ToolDefinition = {
  name: 'wps_word_format_table',
  description: `设置文档里某个表格的样式 / 表头加粗 / 边框。

参数：
- tableIndex: 第几个表格（1 基，默认 1）
- style: 表格样式名（需文档里存在该样式）
- headerBold: 首行加粗
- borders: 加内部+外部单线边框`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      tableIndex: { type: 'number', description: '第几个表格（1 基，默认 1）' },
      style: { type: 'string', description: '表格样式名' },
      headerBold: { type: 'boolean', description: '首行加粗' },
      borders: { type: 'boolean', description: '加边框' },
    },
  },
};

export const formatTableHandler: ToolHandler = async (args) => {
  try {
    const r = await call<{ tableIndex: number; applied: { style: boolean; headerBold: boolean; borders: boolean } }>(
      'formatTable',
      {
        tableIndex: args.tableIndex ?? 1,
        style: args.style,
        headerBold: args.headerBold ?? false,
        borders: args.borders ?? false,
      }
    );
    if (!r.ok) return fail(`设置表格样式失败：${r.error}`, r.error);
    const a = r.data!.applied;
    const done = [
      a.style ? '套用样式' : null,
      a.headerBold ? '表头加粗' : null,
      a.borders ? '加边框' : null,
    ].filter(Boolean);
    return ok(
      done.length
        ? `第 ${r.data?.tableIndex} 个表格已应用：${done.join('、')}`
        : `第 ${r.data?.tableIndex} 个表格没有可应用的项（未指定 style/headerBold/borders，或应用失败）`,
      r.warnings
    );
  } catch (e) {
    return fail(`设置表格样式异常：${errMsg(e)}`, errMsg(e));
  }
};

/** Word 长文档/学术排版工具集合 */
export const academicTools: RegisteredTool[] = [
  { definition: addHeadingDefinition, handler: addHeadingHandler },
  { definition: getOutlineDefinition, handler: getOutlineHandler },
  { definition: insertParagraphDefinition, handler: insertParagraphHandler },
  { definition: deleteParagraphDefinition, handler: deleteParagraphHandler },
  { definition: addCaptionDefinition, handler: addCaptionHandler },
  { definition: addFootnoteDefinition, handler: addFootnoteHandler },
  { definition: insertTableOfFiguresDefinition, handler: insertTableOfFiguresHandler },
  { definition: updateFieldsDefinition, handler: updateFieldsHandler },
  { definition: setColumnsDefinition, handler: setColumnsHandler },
  { definition: setDefaultFontDefinition, handler: setDefaultFontHandler },
  { definition: setHeadingNumberingDefinition, handler: setHeadingNumberingHandler },
  { definition: addListDefinition, handler: addListHandler },
  { definition: formatTableDefinition, handler: formatTableHandler },
];

export default academicTools;
