/**
 * Input: DocSpec（文档质量声明，自有 schema——字段名刻意不与外部引擎同源）
 * Output: wps_word_build_document 工具：纯函数展开成既有 word 工具调用序列（callTool 真入口）+
 *         生成后校验（页数/段落/空段三路）——deckkit 同构（L1 优秀档，INTENT-DESIGN §4.3）
 * Pos: M2 wordkit 文档质量引擎（主线一 B 段 2026-11-06）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *   展开成的是既有 word 工具（add_heading/insert_text/insert_table/insert_image/add_list/
 *   insert_page_break），逐动作经 toolRegistry.callTool 走真入口——每个子动作各自的
 *   快照/协作/注解护栏生效。ps1 零改动。
 * 密度纪律（WORDKIT_TYPO，自创数值，须真机复核）：正文字号下限、每 doc 段落上限——
 *   超出直接拒绝（不静默截断），与 DECK_TYPO 同范式。
 * 生成后校验（deckkit 溢出检查范式）：展开执行完回读 get_document_stats + get_paragraphs——
 *   ①段落数回读（与展开步数对账）②空段落检出 ③校验通道不可用 → 降级记一条不失败。
 */
import { v4 as uuidv4 } from 'uuid';
import { ToolDefinition, ToolHandler, ToolCategory, RegisteredTool } from '../../types/tools';
import { toolRegistry } from '../../server/tool-registry';
import { WORD_TARGET_PROPERTY } from './target';
import { reportProgress } from '../../server/progress-context';

/** 密度纪律常量（自创数值，须真机复核） */
export const WORDKIT_TYPO = {
  /** 正文段文字下限：低于此值视为可疑空段 */
  paraTextMin: 2,
  /** 段落上限（质量引擎一次建整篇，超限建议拆文档） */
  maxParas: 80,
  /** list 项数上限 */
  maxListItems: 20,
  /** table 行列上限（4 行 × 6 列起） */
  maxTableRows: 30,
  maxTableCols: 8,
} as const;

export const DOC_KINDS = ['title', 'heading', 'para', 'table', 'image', 'list', 'quote', 'pagebreak'] as const;
export type DocKind = (typeof DOC_KINDS)[number];

export interface DocSection {
  kind: DocKind;
  /** title/heading/para/quote/list 项/table 单元格 通用文字字段（不叫 text——刻意不与外部引擎同源） */
  content?: string;
  /** heading 级别（1-9，title 固定 1，heading 默认 2） */
  level?: number;
  /** table 专有：首行表头 + 数据行（content 不用） */
  head?: string[];
  rows?: string[][];
  /** image 专有 */
  imagePath?: string;
  width?: number;
  height?: number;
  /** list 专有 */
  items?: string[];
  listType?: 'bullet' | 'number';
}

export interface DocSpec {
  sections: DocSection[];
}

/** 展开产物：一步 = 一次既有 word 工具调用（callTool 真入口） */
export interface DocStep {
  tool: string;
  args: Record<string, unknown>;
  /** 该步产出的段落语义（校验用：expectedParas 计数） */
  paraKind?: 'title' | 'heading' | 'para' | 'quote' | 'list-item' | 'table-cell' | 'none';
}

const KNOWN_KINDS = DOC_KINDS.join(' / ');

function quoteWrap(s: string): string {
  return '「' + s + '」';
}

/** 纯函数：DocSpec → 步骤序列。抛错 = 描述违反纪律（未知 kind/缺字段/超密度），handler 转如实报错 */
export function expandDoc(spec: DocSpec): DocStep[][] {
  if (!spec || !Array.isArray(spec.sections)) {
    throw new Error('doc.sections 为空——至少给一个 section');
  }
  if (!spec.sections.length) {
    throw new Error('doc.sections 为空——至少给一个 section');
  }
  const paraCount = spec.sections.filter((s) => s.kind === 'para' || s.kind === 'quote').length +
    spec.sections.reduce((n, s) => n + (s.kind === 'list' ? (s.items?.length || 0) : 0), 0);
  if (paraCount > WORDKIT_TYPO.maxParas) {
    throw new Error(`段落数 ${paraCount} 超过上限 ${WORDKIT_TYPO.maxParas}——拆文档`);
  }
  return spec.sections.map((s, i) => {
    const at = `第 ${i + 1} 段`;
    switch (s.kind) {
      case 'title':
        if (!s.content?.trim()) throw new Error(`${at} title 缺 content`);
        return [{ tool: 'wps_word_add_heading', args: { text: s.content, level: 1 }, paraKind: 'title' }];
      case 'heading': {
        const level = Math.min(Math.max(s.level ?? 2, 1), 9);
        if (!s.content?.trim()) throw new Error(`${at} heading 缺 content`);
        return [{ tool: 'wps_word_add_heading', args: { text: s.content, level }, paraKind: 'heading' }];
      }
      case 'para':
        if (!s.content?.trim()) throw new Error(`${at} para 缺 content`);
        return [{ tool: 'wps_word_insert_text', args: { text: s.content, position: 'end', new_paragraph: true }, paraKind: 'para' }];
      case 'quote':
        if (!s.content?.trim()) throw new Error(`${at} quote 缺 content`);
        return [{ tool: 'wps_word_insert_text', args: { text: quoteWrap(s.content), position: 'end', new_paragraph: true }, paraKind: 'quote' }];
      case 'table': {
        if (!s.rows?.length) throw new Error(`${at} table 缺 rows（数据行，首行可为表头）`);
        const cols = s.rows[0]?.length || 0;
        if (!cols) throw new Error(`${at} table rows 为空`);
        if (s.rows.some((r) => r.length !== cols)) throw new Error(`${at} table 行列数不一致`);
        if (s.rows.length > WORDKIT_TYPO.maxTableRows) throw new Error(`${at} table 行数 ${s.rows.length} 超上限 ${WORDKIT_TYPO.maxTableRows}`);
        if (cols > WORDKIT_TYPO.maxTableCols) throw new Error(`${at} table 列数 ${cols} 超上限 ${WORDKIT_TYPO.maxTableCols}`);
        // ps1 insertTable 原生吃 data（一次落位全部单元格）——不逐格写
        return [{ tool: 'wps_word_insert_table', args: { rows: s.rows.length, cols, data: s.rows }, paraKind: 'none' }];
      }
      case 'image':
        if (!s.imagePath?.trim()) throw new Error(`${at} image 缺 imagePath`);
        return [{
          tool: 'wps_word_insert_image',
          args: {
            imagePath: s.imagePath,
            ...(s.width !== undefined ? { width: s.width } : {}),
            ...(s.height !== undefined ? { height: s.height } : {}),
          },
          paraKind: 'none',
        }];
      case 'list': {
        if (!s.items?.length) throw new Error(`${at} list 缺 items`);
        if (s.items.length > WORDKIT_TYPO.maxListItems) throw new Error(`${at} list 项数 ${s.items.length} 超上限 ${WORDKIT_TYPO.maxListItems}`);
        return s.items.map((item) => ({
          tool: 'wps_word_insert_text',
          args: { text: (s.listType === 'number' ? '' : '• ') + item, position: 'end', new_paragraph: true },
          paraKind: 'list-item' as const,
        }));
      }
      case 'pagebreak':
        return [{ tool: 'wps_word_insert_page_break', args: {}, paraKind: 'none' }];
      default:
        throw new Error(`未知 section kind「${(s as { kind?: string }).kind}」（${at}）。支持：${KNOWN_KINDS}`);
    }
  });
}

/** 展开序列里的预期段落数（title/heading/para/quote/list-item 各占一段；table-cell 不经段落） */
export function expectedParaCount(pages: DocStep[][]): number {
  return pages.flat().filter((s) => s.paraKind && s.paraKind !== 'none').length;
}

export const buildDocumentDefinition: ToolDefinition = {
  name: 'wps_word_build_document',
  description: `按 DocSpec 声明一次生成整篇 Word 文档（wordkit 引擎，deckkit 同构：声明 → 纯函数展开 → 生成后校验）。

sections[] 的 kind（8 种）：title / heading / para / table / image / list / quote / pagebreak。
- title: { content } —— 一级标题（heading level 1）
- heading: { content, level? } —— 章节标题（默认 2 级，1-9）
- para: { content } —— 正文段落
- table: { head?, rows } —— 表格（rows 首行可为表头；逐格写入）
- image: { imagePath, width?, height? } —— 插图
- list: { items, listType? } —— 列表（bullet • / number）
- quote: { content } —— 引文段落（「」包裹）
- pagebreak: {} —— 分页

生成后自动回读校验：段落数对账、空段落检出——异常进 warnings 不静默；
校验通道不可用 → 降级记一条不失败（deckkit 溢出检查同范式）。
新建文档语义（不追加到已有文档）；坐标/尺寸口径与 1.0.5 一致。

使用场景：
- "按这份提纲建一篇周报文档"（title + para + list 一口气）
- "生成带表格和配图的说明文档"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      doc: {
        type: 'object',
        description: '文档声明：{ sections: [ {kind:"title",content}, {kind:"para",content}, ... ] }，kind 见描述',
      },
    },
    required: ['doc'],
  },
};

export const buildDocumentHandler: ToolHandler = async (args) => {
  const spec = args.doc as DocSpec | undefined;
  let pages: DocStep[][];
  try {
    pages = expandDoc(spec as DocSpec);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `doc 声明不合法：${msg}` }],
      error: msg,
    };
  }
  const steps = pages.flat();
  const expected = expectedParaCount(pages);
  const total = steps.length + 1; // + createDocument
  const failures: Array<{ page: number; step: number; tool: string; error: string }> = [];
  let executed = 0;

  reportProgress(0, total, '新建文档');

  // 建文档（新建语义；追加到已有文档归编辑类工具）
  const create = await toolRegistry.callTool({ id: uuidv4(), name: 'wps_word_create_document', arguments: {} });
  executed += 1;
  if (!create.success) {
    const msg = (create.content?.[0] as { text?: string })?.text || 'create_document 失败';
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: msg }], error: msg };
  }
  reportProgress(1, total, '文档已建');

  // 逐步展开（fail-fast：报停在第几段第几步，不静默续跑）
  let done = 0;
  outer: for (let p = 0; p < pages.length; p++) {
    for (let j = 0; j < pages[p].length; j++) {
      const step = pages[p][j];
      const r = await toolRegistry.callTool({ id: uuidv4(), name: step.tool, arguments: step.args });
      executed += 1;
      done += 1;
      reportProgress(executed, total, `第 ${p + 1} 段 · ${step.tool}`);
      if (!r.success) {
        failures.push({
          page: p + 1,
          step: j + 1,
          tool: step.tool,
          error: (r.content?.[0] as { text?: string })?.text || r.error || '未知错误',
        });
        break outer;
      }
    }
  }

  // ---- 生成后校验（deckkit 溢出检查范式）：读回段落数对账 + 空段落检出 ----
  const warnings: Array<{ step: string; message: string; severity?: 'info' | 'warning' }> = [];
  let checkedParaCount: number | null = null;
  try {
    const stats = await toolRegistry.callTool({ id: uuidv4(), name: 'wps_word_get_document_stats', arguments: {} });
    if (stats.success) {
      const st = stats.content?.[0] as { text?: string };
      const m = st?.text?.match(/(\d+)\s*个段落/) || st?.text?.match(/段落[^\d]*(\d+)/);
      if (m) checkedParaCount = Number(m[1]);
    } else {
      warnings.push({ step: 'wordkit.校验降级', severity: 'info', message: '文档统计回读失败，段落数对账未执行（不影响已生成内容）' });
    }
    const paras = await toolRegistry.callTool({ id: uuidv4(), name: 'wps_word_get_paragraphs', arguments: {} });
    if (paras.success) {
      const pText = (paras.content?.[0] as { text?: string })?.text || '';
      const paraLines = pText.split('\n').filter((l) => /^\[\d+\]/.test(l));
      const emptyParas = paraLines.filter((l) => {
        const body = l.replace(/^\[\d+\]\s*\([^)]*\)\s*/, '').trim();
        return body.length < WORDKIT_TYPO.paraTextMin;
      });
      if (emptyParas.length) {
        warnings.push({
          step: 'wordkit.空段落',
          severity: 'warning',
          message: `检出 ${emptyParas.length} 个疑似空段落（${emptyParas.slice(0, 3).join('；')}${emptyParas.length > 3 ? '…' : ''}）——请确认是否多余`,
        });
      }
    } else {
      warnings.push({ step: 'wordkit.校验降级', severity: 'info', message: '段落结构回读失败，空段落检查未执行（不影响已生成内容）' });
    }
    if (checkedParaCount !== null && expected > 0 && checkedParaCount < expected) {
      warnings.push({
        step: 'wordkit.段落数对账',
        severity: 'warning',
        message: `回读段落数 ${checkedParaCount} 少于展开预期的 ${expected}——部分段落可能未落位，建议 wps_word_get_paragraphs 复核`,
      });
    }
  } catch (e) {
    warnings.push({
      step: 'wordkit.校验降级',
      severity: 'info',
      message: `生成后校验异常（不影响已生成内容）：${e instanceof Error ? e.message : String(e)}`,
    });
  }

  const body = failures.length
    ? `生成中止：完成到第 ${pages[failures[0].page - 1] ? failures[0].page : pages.length} 段（执行 ${executed}/${total} 步）。失败明细：${failures.map((f) => `第${f.page}段#${f.step} ${f.tool}: ${f.error.slice(0, 80)}`).join('；')}`
    : `已生成整篇文档（${pages.length} 段声明、${steps.length} 步展开、${done} 步落位）。`;

  return {
    id: uuidv4(),
    success: failures.length === 0,
    content: [{
      type: 'text',
      text: body + (warnings.length ? `\nwarnings:\n- ${warnings.map((w) => w.message).join('\n- ')}` : ''),
    }],
    ...(warnings.length ? { warnings } : {}),
    ...(failures.length ? { error: body } : {}),
  };
};

export const wordkitTools: RegisteredTool[] = [
  { definition: buildDocumentDefinition, handler: buildDocumentHandler },
];

export default wordkitTools;
