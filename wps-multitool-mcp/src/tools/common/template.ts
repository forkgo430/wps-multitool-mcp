/**
 * Input: 模板渲染/提取参数（templatePath/data/outputPath；path/outputPath/maxSlots）
 * Output: 渲染回包（filled/skipped/落盘校验）与提取回包（anchors/placeholderCount）
 * Pos: 模板层大卡（第 5 期主线 2026-11-11）A/B 段工具壳。一旦我被修改，请更新我的头部注释，
 *      以及所属目录的 md。
 *
 * 锚点双轨：①书签锚点优先（WPS 开放平台正统做法）②{{key}} 文本占位兜底。
 * 缺 key 不静默：模板里有、data 没给的锚点 → warnings 列全清单。
 * 实地教训防复发（R2/R4）：路径参数 ps1 侧 Normalize-WinPath 归一；renderTemplate 进慢档
 * （ps-child DEFAULT_SLOW_ACTIONS 已补 renderTemplate/extractTemplate）。
 * schema 纪律（提质小卡先例）：参数描述按"名+类型+约束"写短。
 */
import { ToolDefinition, ToolCategory, ToolHandler, ToolCallResult } from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { ok, fail, errText } from '../word/helpers';

const normalizeWinPath = (p: string): string => p.replace(/\//g, '\\');

export const renderTemplateDefinition: ToolDefinition = {
  name: 'wps_render_template',
  description: '数据模板渲染：打开模板（docx），书签锚点优先 + {{key}} 占位兜底填充，另存 outputPath。缺 key 保留原文并警告。',
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      templatePath: { type: 'string', description: '模板文件路径（docx）' },
      data: { type: 'object', description: '锚点名 → 值（JSON 对象）' },
      outputPath: { type: 'string', description: '渲染产物路径（已存在则拒绝）' },
      format: { type: 'string', enum: ['auto', 'docx', 'pptx', 'xlsx'], description: '模板格式，auto 缺省按扩展名分发（docx/pptx/xlsx 三路）' },
    },
    required: ['templatePath', 'data', 'outputPath'],
  },
};

export const renderTemplateHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { templatePath, data, outputPath, format = 'auto' } = args as {
    templatePath?: string;
    data?: Record<string, unknown>;
    outputPath?: string;
    format?: 'auto' | 'docx' | 'pptx' | 'xlsx';
  };
  if (!templatePath) return fail(errText('templatePath 必填', { error: 'templatePath required' }));
  if (!outputPath) return fail(errText('outputPath 必填', { error: 'outputPath required' }));
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return fail(errText('data 必须是 JSON 对象（锚点名 → 值）', { error: 'data must be object' }));
  }
  if (Object.keys(data).length === 0) {
    return fail(errText('data 为空对象——至少给一个锚点值', { error: 'data empty' }));
  }
  // R2 教训防复发：Windows host 路径入口归一（/ → \）
  // format 分发（提质三格式卡）：auto 缺省按扩展名（docx/pptx/xlsx 三路）；docx 路径行为不变
  const ext = (templatePath.split('.').pop() || '').toLowerCase();
  const fmt = format === 'auto' ? (ext === 'pptx' ? 'pptx' : ext === 'xlsx' ? 'xlsx' : 'docx') : format;
  if (fmt === 'pptx' || fmt === 'xlsx') {
    const action = fmt === 'pptx' ? 'renderTemplatePpt' : 'renderTemplateXlsx';
    const pptApp = fmt === 'pptx' ? 'ppt' : 'excel';
    const r = await wpsClient.executeMethod<{
      outputPath: string;
      size: number;
      filled: Array<{ anchor: string; via: string }>;
      skipped: string[];
      filledCount: number;
    }>(action, {
      templatePath: normalizeWinPath(templatePath),
      outputPath: normalizeWinPath(outputPath),
      data,
    }, pptApp as never);
    if (!r.success || !r.data) return fail(errText('模板渲染失败', r), r.error);
    const pd = r.data;
    const plines = [
      `模板渲染完成（${fmt}）：${pd.outputPath}（${pd.size}B）`,
      `填充 ${pd.filledCount} 个锚点`,
      pd.skipped && pd.skipped.length ? `未替换锚点（data 缺值，保留原文）：${pd.skipped.join('、')}` : '锚点全部填充',
    ];
    return ok(plines.join('\n'), r.warnings);
  }
  const r = await wpsClient.executeMethod<{
    outputPath: string;
    size: number;
    filled: Array<{ anchor: string; via: string }>;
    skipped: string[];
    filledCount: number;
  }>('renderTemplate', {
    templatePath: normalizeWinPath(templatePath),
    outputPath: normalizeWinPath(outputPath),
    data,
  }, 'word' as never);
  if (!r.success || !r.data) return fail(errText('模板渲染失败', r), r.error);
  const d = r.data;
  const lines = [
    `模板渲染完成：${d.outputPath}（${d.size}B）`,
    `填充 ${d.filledCount} 个锚点（书签 ${d.filled.filter((f) => f.via === 'bookmark').length} / 占位 ${d.filled.filter((f) => f.via === 'placeholder').length}）`,
    d.skipped && d.skipped.length ? `未替换锚点（data 缺值，保留原文）：${d.skipped.join('、')}` : '锚点全部填充',
  ];
  return ok(lines.join('\n'), r.warnings);
};

export const extractTemplateDefinition: ToolDefinition = {
  name: 'wps_extract_template',
  description: '模板提取：读现有 Word 文档结构，正文段转 {{占位}}（样式随段落保留），标题层级保序，存为带锚点模板。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      path: { type: 'string', description: '源文档路径' },
      outputPath: { type: 'string', description: '模板落点（已存在则拒绝）' },
      maxSlots: { type: 'number', description: '占位上限，默认 200（≤500）' },
    },
    required: ['path', 'outputPath'],
  },
};

export const extractTemplateHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { path: srcPath, outputPath, maxSlots } = args as {
    path?: string;
    outputPath?: string;
    maxSlots?: number;
  };
  if (!srcPath) return fail(errText('path 必填（源文档）', { error: 'path required' }));
  if (!outputPath) return fail(errText('outputPath 必填（模板落点）', { error: 'outputPath required' }));
  const r = await wpsClient.executeMethod<{
    outputPath: string;
    anchors: Array<{ paragraph: number; style: string; kind: string; anchor?: string; original?: string; reason?: string }>;
    placeholderCount: number;
    headingCount: number;
  }>('extractTemplate', {
    path: normalizeWinPath(srcPath),
    outputPath: normalizeWinPath(outputPath),
    ...(maxSlots ? { maxSlots } : {}),
  }, 'word' as never);
  if (!r.success || !r.data) return fail(errText('模板提取失败', r), r.error);
  const d = r.data;
  const lines = [
    `模板已生成：${d.outputPath}`,
    `占位 ${d.placeholderCount} 个 / 标题保留 ${d.headingCount} 个（跳过 ${d.anchors.filter((a) => a.kind === 'skipped').length} 段）`,
    '锚点清单（占位前 50）：',
    ...d.anchors
      .filter((a) => a.kind === 'placeholder')
      .slice(0, 50)
      .map((a) => `  {{${a.anchor}}}（第 ${a.paragraph} 段，样式 ${a.style}，原文「${a.original || ''}」）`),
    d.anchors.some((a) => a.kind === 'skipped') ? '注：部分段落替换失败已如实跳过（见 anchors.kind=skipped）' : '',
  ].filter(Boolean);
  return ok(lines.join('\n'), r.warnings);
};

export const templateTools = [
  { definition: renderTemplateDefinition, handler: renderTemplateHandler },
  { definition: extractTemplateDefinition, handler: extractTemplateHandler },
];
