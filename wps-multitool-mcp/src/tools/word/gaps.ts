/**
 * Input: Word 深水区工具的参数（统计/批注/书签/表格读取、建关文档、超链接、表转文本）
 * Output: 各自的查询或修改结果
 * Pos: 第 5 期补齐的 Word 缺口。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 这一批补的是什么
 * ----------------
 * 对着同类 WPS 自动化工具的公开能力清单逐项比对后挑出来的**真缺口**：
 * 同类 61 / 本仓 58，数量不差，差在"深水区"三批 ——
 *
 *   一、读侧太弱（模型只能盲写）。现只能插批注/建书签/插表格，
 *       读不了删不掉："文档里有几个表、表里有什么"答不上，
 *        skill 里"先读后写"的原则没有工具支撑。→ get_document_stats /
 *        get_comments / get_bookmarks / get_tables / get_table_data
 *   二、建不了新文档。PS 动作 createDocument 早就有（还修过"零等待返回空壳"的
 *       病史），但一直没有 TS 工具暴露 —— Word 建不了新文档，而 Excel 有
 *        create_workbook。→ create_document / close_document
 *   三、批注管理无闭环、超链接完全没有。→ delete_comment / insert_hyperlink
 *        外加 convert_table_to_text（表格贴去别处前的常规动作）。
 *
 * 刻意不补的（对方有我也不做）：
 *   get_revisions / accept / reject —— wps_word_revisions 的 mode
 *   （list/acceptAll/rejectAll/acceptOne/rejectOne）已全覆盖，拆三个只增加模型选择负担；
 *   mail_merge / content_control —— WPS 12.1 支持度存疑或场景边缘。
 *
 * close_document 的"未保存怎么办"（设计答案，不是实现细节）
 * ------------------------------------------------------
 * COM 是 STA：关闭未保存文档会弹「另存为」模态框，弹框就把桥通道卡死。
 * PS 侧对此有确定性守卫 —— 文档从未保存过时直接报错，让调用方先 save_as，
 * 或显式 save=false。工具描述里必须把这条说给模型听。
 *
 * 真机支持度：新 PS 动作逐个探过（__probeWordCaps，在临时文档上探、探完删），
 * 见 scripts/manual/word-caps-probe.js。
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
import { wpsClient, setWordTarget, getWordTarget } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';
import { WORD_TARGET_PROPERTY } from './target';

type Args = Record<string, unknown>;

/** 统一调用：PS 侧 data.warnings 已被 wpsClient 提到顶层，一路带回 */
async function call<T>(action: string, params: Args): Promise<{
  ok: boolean;
  data?: T;
  error?: string;
  warnings?: ToolWarning[];
}> {
  const r = await wpsClient.executeMethod<T>(action, params, WpsAppType.WRITER);
  return { ok: r.success, data: r.data, error: r.error, warnings: r.warnings };
}

const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: true,
  content: [{ type: 'text', text }],
  ...(warnings && warnings.length ? { warnings } : {}),
});

const fail = (text: string, error?: string): ToolCallResult => ({
  id: uuidv4(),
  success: false,
  content: [{ type: 'text', text }],
  error: error ?? text,
});

// ============================================================================
// 批次 A：只读 5 个（低风险，直接支撑"先读后写"）
// ============================================================================

// ---------------------------------------------------------------------------
// 1. 文档统计
// ---------------------------------------------------------------------------

export const getDocumentStatsDefinition: ToolDefinition = {
  name: 'wps_word_get_document_stats',
  description:
    '读文档规模：字数 / 页数 / 段落数 / 字符数 / 节数。\n\n' +
    '什么时候用：动文档前先看清它多大（决定要不要分段处理、读多少）；' +
    '用户问"这篇多少字"时直接答。比 get_active_document 深一层 —— ' +
    '那个只给当前文档的基本信息，这个给全文统计口径。\n' +
    '页数在个别 WPS 版本上取不到，会走 warnings 说明，其余项不受影响。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const getDocumentStatsHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  const r = await call<{
    name: string;
    path: string;
    paragraphs: number;
    words: number;
    characters: number;
    sections: number;
    pages: number;
  }>('getDocumentStats', {});
  if (!r.ok) return fail(`读取文档统计失败: ${r.error}`, r.error);
  const d = r.data || ({} as never);
  const lines = [
    `文档「${d.name}」统计：`,
    `  字数     : ${d.words}`,
    `  字符数   : ${d.characters}`,
    `  段落数   : ${d.paragraphs}`,
    `  节数     : ${d.sections}`,
    `  页数     : ${d.pages}`,
  ];
  if (d.path) lines.push(`  路径     : ${d.path}`);
  return ok(lines.join('\n'), r.warnings);
};

// ---------------------------------------------------------------------------
// 2. 批注列表
// ---------------------------------------------------------------------------

export const getCommentsDefinition: ToolDefinition = {
  name: 'wps_word_get_comments',
  description:
    '列出文档里所有批注：序号、作者、日期、内容、所批注的原文片段。\n\n' +
    '什么时候用：处理审阅意见前先看有哪些（以前只能插批注，读不了）；' +
    '删批注前先用这个确认序号。\n' +
    '与 wps_word_revisions 的分工：那个管"修订"（track changes），这个管"批注"。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const getCommentsHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  const r = await call<{
    count: number;
    comments: Array<{
      index: number;
      author: string;
      date: string;
      text: string;
      reference: string;
    }>;
  }>('getComments', {});
  if (!r.ok) return fail(`读取批注失败: ${r.error}`, r.error);
  const list = r.data?.comments || [];
  if (!list.length) return ok('文档里没有批注');
  const lines = list.map(
    (c) =>
      `  #${c.index} [${c.author || '未知作者'}${c.date ? ' · ' + c.date : ''}]\n` +
      `     内容: ${c.text}${c.reference ? `\n     针对: ${c.reference}` : ''}`
  );
  return ok(`文档批注（共 ${list.length} 条）:\n${lines.join('\n')}`);
};

// ---------------------------------------------------------------------------
// 3. 删除批注
// ---------------------------------------------------------------------------

export const deleteCommentDefinition: ToolDefinition = {
  name: 'wps_word_delete_comment',
  description:
    '删除一条批注：按序号（get_comments 返回的 #n）或按内容片段定位。\n\n' +
    '什么时候用：清理审阅意见、撤掉自己插错的批注。\n' +
    '注意：删除不可恢复（COM 没有撤销），删前先用 wps_word_get_comments 确认目标。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      index: {
        type: 'number',
        description: '批注序号（从 1 开始，见 get_comments 的 #n）',
      },
      text: {
        type: 'string',
        description: '批注内容片段（按内容匹配，index 与 text 二选一）',
      },
    },
  },
};

export const deleteCommentHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const { index, text } = args as { index?: number; text?: string };
  if (index === undefined && !text) {
    return fail('需要 index（序号）或 text（内容片段）之一');
  }
  const params: Args = {};
  if (index !== undefined) params.index = index;
  if (text) params.text = text;
  const r = await call<{ deleted: string; remaining: number }>('deleteComment', params);
  if (!r.ok) return fail(`删除批注失败: ${r.error}`, r.error);
  const preview = (r.data?.deleted || '').slice(0, 60);
  return ok(`批注已删除:「${preview}」，剩余 ${r.data?.remaining ?? '?'} 条`);
};

// ---------------------------------------------------------------------------
// 4. 书签列表
// ---------------------------------------------------------------------------

export const getBookmarksDefinition: ToolDefinition = {
  name: 'wps_word_get_bookmarks',
  description:
    '列出文档里所有书签：名称、位置、锚定文本。\n\n' +
    '什么时候用：模板填空前看有哪些书签可用（以前只能建书签，读不到有哪些）；' +
    '确认某次插入是否成功。\n' +
    '配套：wps_word_replace_bookmark_content 按名称填内容，名称从这儿查。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const getBookmarksHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  const r = await call<{
    count: number;
    bookmarks: Array<{ name: string; start: number; text: string }>;
  }>('getBookmarks', {});
  if (!r.ok) return fail(`读取书签失败: ${r.error}`, r.error);
  const list = r.data?.bookmarks || [];
  if (!list.length) return ok('文档里没有书签');
  const lines = list.map((b) => `  · ${b.name}（位置 ${b.start}）${b.text ? ` → 「${b.text}」` : ''}`);
  return ok(`文档书签（共 ${list.length} 个）:\n${lines.join('\n')}`);
};

// ---------------------------------------------------------------------------
// 5. 表格列表
// ---------------------------------------------------------------------------

export const getTablesDefinition: ToolDefinition = {
  name: 'wps_word_get_tables',
  description:
    '列出文档里所有表格：序号、行数、列数、第一行预览。\n\n' +
    '什么时候用：回答"文档里有几个表"；操作表格前确认序号（现有工具都按序号定位）。\n' +
    '与 wps_word_get_cell_text 的分工：那个读单个格子，这个先给你全表轮廓，' +
    '要看整表内容再用 wps_word_get_table_data。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const getTablesHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  const r = await call<{
    count: number;
    tables: Array<{ index: number; rows: number; columns: number; firstRow: string }>;
  }>('getTables', {});
  if (!r.ok) return fail(`读取表格列表失败: ${r.error}`, r.error);
  const list = r.data?.tables || [];
  if (!list.length) return ok('文档里没有表格');
  const lines = list.map(
    (t) => `  #${t.index} ${t.rows} 行 × ${t.columns} 列${t.firstRow ? `\n     首行: ${t.firstRow}` : ''}`
  );
  return ok(`文档表格（共 ${list.length} 张）:\n${lines.join('\n')}`);
};

// ---------------------------------------------------------------------------
// 6. 表格数据
// ---------------------------------------------------------------------------

export const getTableDataDefinition: ToolDefinition = {
  name: 'wps_word_get_table_data',
  description:
    '读整张表格的内容（二维文本，按序号定位）。\n\n' +
    '什么时候用：把 Word 表格的数据读出来做分析/搬运；核对表格填对了没有。\n' +
    '大表保护：超过 100 行或 20 列只返回前一部分，并在 warnings 里说明 —— ' +
    '别把一次调用撑爆。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      tableIndex: {
        type: 'number',
        description: '表格序号（从 1 开始，见 get_tables），默认 1',
      },
    },
  },
};

export const getTableDataHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const { tableIndex } = args as { tableIndex?: number };
  const params: Args = {};
  if (tableIndex !== undefined) params.tableIndex = tableIndex;
  const r = await call<{
    tableIndex: number;
    rows: number;
    columns: number;
    returnedRows: number;
    returnedColumns: number;
    data: string[][];
  }>('getTableData', params);
  if (!r.ok || !r.data) return fail(`读取表格数据失败: ${r.error}`, r.error);
  const d = r.data;
  const lines = d.data.map((row, i) => `  R${i + 1}: ${row.join(' | ')}`);
  return ok(
    `表格 #${d.tableIndex}（${d.rows} 行 × ${d.columns} 列）:\n${lines.join('\n')}`,
    r.warnings
  );
};

// ============================================================================
// 批次 B：写入 5 个
// ============================================================================

// ---------------------------------------------------------------------------
// 7. 新建文档
// ---------------------------------------------------------------------------

export const createDocumentDefinition: ToolDefinition = {
  name: 'wps_word_create_document',
  description:
    '新建一个空白 WPS 文字文档，并把它设为后续 word 操作的目标。\n\n' +
    '什么时候用：从零开始写文档（以前 Word 建不了新文档，Excel 却有 create_workbook）；' +
    '没有打开任何文档时，后续全部 word 工具都会失败，先建一个。\n' +
    '建完会自动锁定目标文档，紧接着的插入/设置直接生效，不用再切换。\n' +
    '注意：新建的文档【尚未保存】，关它之前要先另存为（见 close_document 的说明）。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: { target: WORD_TARGET_PROPERTY, },
  },
};

export const createDocumentHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  const r = await call<{ count: number; name: string }>('createDocument', {});
  if (!r.ok) return fail(`新建文档失败: ${r.error}`, r.error);
  // 锁定目标：ActiveDocument 是弱引用（实测 Count>0 时也可能取到空壳），
  // 锁定名字是唯一不依赖 UI 焦点的定位方式。拿不到名字就不锁，让后续调用走遍历兜底。
  const name = r.data?.name;
  if (name) setWordTarget(name);
  return ok(
    `新文档已创建${name ? `: ${name}` : ''}（当前共 ${r.data?.count ?? '?'} 个文档）` +
      (name ? '，已锁定为后续操作目标' : '') +
      '。文档尚未保存，需要落盘请用 wps_word_save_as。',
    r.warnings
  );
};

// ---------------------------------------------------------------------------
// 8. 关闭文档
// ---------------------------------------------------------------------------

export const closeDocumentDefinition: ToolDefinition = {
  name: 'wps_word_close_document',
  description:
    '关闭指定的 WPS 文字文档（不填关当前目标文档），可选是否保存。\n\n' +
    '什么时候用：做完一个文件收尾；批量处理时逐个关闭。\n' +
    '**未保存怎么办（重要）**：从未保存过的文档，直接关会弹「另存为」对话框，' +
    '而对话框会把自动化通道卡死。所以：从未保存的文档必须先 wps_word_save_as，' +
    '或者显式 save=false 放弃保存（改动全部丢失）。存过盘的文档按 save 参数正常关。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      name: {
        type: 'string',
        description: '要关闭的文档名称，不填则关闭当前目标文档',
      },
      save: {
        type: 'boolean',
        description: '关闭前是否保存，默认 true（从未保存过的文档会要求先另存为）',
      },
    },
  },
};

export const closeDocumentHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const { name, save } = args as { name?: string; save?: boolean };
  const params: Args = {};
  if (name) params.name = name;
  params.saveChanges = save !== false;
  const r = await call<{ closed: string }>('closeDocument', params);
  if (!r.ok) return fail(`关闭文档失败: ${r.error}`, r.error);
  // 目标被关掉了就别再锁着：后续 word 调用会拿着一个已不存在的名字定位，
  // 报错信息还很难懂。这里清掉（只清指向被关文档的那个目标 —— 关了别的文档不动）。
  const target = getWordTarget();
  const closed = r.data?.closed || name;
  if (target && (!name || closed === target)) {
    setWordTarget(undefined);
  }
  return ok(
    `文档已关闭: ${closed}${save !== false ? '（已保存）' : '（未保存，改动已放弃）'}`
  );
};

// ---------------------------------------------------------------------------
// 9. 插入超链接
// ---------------------------------------------------------------------------

export const insertHyperlinkDefinition: ToolDefinition = {
  name: 'wps_word_insert_hyperlink',
  description:
    '在文档中插入超链接（可指定显示文字与插入位置）。\n\n' +
    '什么时候用：给文档加网址链接、参考资料跳转 —— Word 高频动作，以前完全没有。\n' +
    'position：cursor（光标处，默认）/ start（文档开头）/ end（文档末尾）。\n' +
    '兼容性：个别 WPS 版本（实测 12.1 部分装机）不认 5 参签名，工具会自动退回 2 参调用；' +
    '两种签名都失败时会明确报错，别重试。\n' +
    '与书签跳转的区别：这个链到 URL；链到文档内位置请用书签 + 交叉引用。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      url: {
        type: 'string',
        description: '链接地址，如 https://example.com',
      },
      text: {
        type: 'string',
        description: '显示文字，不填则直接显示 url',
      },
      position: {
        type: 'string',
        description: '插入位置：cursor（光标处）/ start（开头）/ end（末尾），默认 cursor',
        enum: ['cursor', 'start', 'end'],
      },
    },
    required: ['url'],
  },
};

export const insertHyperlinkHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const { url, text, position } = args as {
    url: string;
    text?: string;
    position?: string;
  };
  if (!url || !url.trim()) {
    return fail('链接地址不能为空');
  }
  const params: Args = { url };
  if (text) params.text = text;
  if (position) params.position = position;
  const r = await call<{ url: string; text: string; position: string }>('insertHyperlink', params);
  if (!r.ok) return fail(`插入超链接失败: ${r.error}`, r.error);
  const posLabel =
    r.data?.position === 'start' ? '文档开头' : r.data?.position === 'end' ? '文档末尾' : '光标处';
  return ok(`超链接已插入（${posLabel}）: 「${r.data?.text || url}」 → ${url}`);
};

// ---------------------------------------------------------------------------
// 10. 表格转文本
// ---------------------------------------------------------------------------

export const convertTableToTextDefinition: ToolDefinition = {
  name: 'wps_word_convert_table_to_text',
  description:
    '把文档里的表格转换成纯文本（按分隔符拼接），表格本身消失。\n\n' +
    '什么时候用：表格数据要贴去别处（邮件/聊天/纯文本报告）前的常规动作；' +
    '拆掉不再需要的表格结构。\n' +
    '注意：破坏性操作且不可恢复（COM 没有撤销），转换前先 wps_word_get_table_data ' +
    '确认内容，必要时先存快照。',
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      tableIndex: {
        type: 'number',
        description: '表格序号（从 1 开始，见 get_tables），默认 1',
      },
      separator: {
        type: 'string',
        description: '单元格之间的分隔符，默认英文逗号',
      },
    },
  },
};

export const convertTableToTextHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const { tableIndex, separator } = args as { tableIndex?: number; separator?: string };
  const params: Args = {};
  if (tableIndex !== undefined) params.tableIndex = tableIndex;
  if (separator) params.separator = separator;
  const r = await call<{ tableIndex: number; separator: string; remainingTables: number }>(
    'convertTableToText',
    params
  );
  if (!r.ok) return fail(`表格转文本失败: ${r.error}`, r.error);
  return ok(
    `表格 #${r.data?.tableIndex ?? 1} 已转为文本（分隔符: ${r.data?.separator ?? ','}），` +
      `文档中剩余表格 ${r.data?.remainingTables ?? '?'} 张`
  );
};

// ============================================================================
// 导出
// ============================================================================

export const gapTools: RegisteredTool[] = [
  // 批次 A：只读
  { definition: getDocumentStatsDefinition, handler: getDocumentStatsHandler },
  { definition: getCommentsDefinition, handler: getCommentsHandler },
  { definition: deleteCommentDefinition, handler: deleteCommentHandler },
  { definition: getBookmarksDefinition, handler: getBookmarksHandler },
  { definition: getTablesDefinition, handler: getTablesHandler },
  { definition: getTableDataDefinition, handler: getTableDataHandler },
  // 批次 B：写入
  { definition: createDocumentDefinition, handler: createDocumentHandler },
  { definition: closeDocumentDefinition, handler: closeDocumentHandler },
  { definition: insertHyperlinkDefinition, handler: insertHyperlinkHandler },
  { definition: convertTableToTextDefinition, handler: convertTableToTextHandler },
];

export default gapTools;
