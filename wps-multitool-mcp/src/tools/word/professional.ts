/**
 * Input: 各工具定义的参数（见各自 description）
 * Output: WPS 文字的动作结果
 * Pos: Word P0/P1 补齐工具（阶段 6 第二步）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 这一批补的是之前盘点出的缺口，全部按 __probeCaps 的真机探测结果实现：
 *   - P0 修订：之前只能开/查询，没有"处理"这个出口（半个功能）
 *   - P0 表格单元格级：之前只有整表插入和样式
 *   - P1 样式 新建/修改/删除：之前只能套用现有样式
 *   - P1 交叉引用 / 索引：完全没有
 *   - P1 页码域 / 奇偶页 / 首页不同：页眉页脚只有统一那种
 *
 * 探测依据：scripts/manual/caps-probe.js 在本机 WPS 12.1 上 36/39 支持。
 * 唯一报不支持的 MarkEntry 已单独标注（见 markIndexEntry 的描述）。
 *
 * 坐标约定：表格行列一律 1-based，与 Word UI 一致。
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
import { WORD_TARGET_PROPERTY } from './target';

type Args = Record<string, unknown>;

const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: true,
  content: [{ type: 'text', text }],
  ...(warnings && warnings.length ? { warnings } : {}),
});

const fail = (text: string, err?: string): ToolCallResult => ({
  id: uuidv4(),
  success: false,
  content: [{ type: 'text', text }],
  error: err ?? text,
});

/** 统一调用：Word 应用，走常驻桥。warnings 一路带回（主操作成功但有保留） */
async function call<T>(action: string, params: Args): Promise<{ ok: boolean; data?: T; error?: string; warnings?: ToolWarning[] }> {
  const r = await wpsClient.executeMethod<T>(action, params, WpsAppType.WRITER);
  return { ok: r.success, data: r.data, error: r.error, warnings: r.warnings };
}

/** 失败时统一文案：把 WPS 的原始报错带上，别让调用方只看到"失败" */
function errText(prefix: string, r: { error?: string }): string {
  return `${prefix}：${r.error ?? '未知错误'}`;
}

// ============================================================================
// P0：修订
// ============================================================================

export const revisionsDefinition: ToolDefinition = {
  name: 'wps_word_revisions',
  description: `处理文档的修订（tracked changes）：列出 / 全部接受 / 全部拒绝 / 单条接受或拒绝。

使用场景：
- "把这份文档的所有修订都接受"
- "看看我改了哪些地方"（mode=list）
- "第 3 条修订不要了"

参数：
- mode: list | acceptAll | rejectAll | acceptOne | rejectOne
- index: 单条处理时必填，从 1 开始

说明：
- 接受 = 保留改动；拒绝 = 撤销改动
- 单条操作前建议先 mode=list 看看作者和内容，别搞错
- 文档没开启"修订"时 list 返回空数组`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      mode: {
        type: 'string',
        enum: ['list', 'acceptAll', 'rejectAll', 'acceptOne', 'rejectOne'],
        description: '操作方式，默认 list',
      },
      index: { type: 'number', description: '单条处理的序号（从 1 开始），acceptOne/rejectOne 必填' },
    },
    required: ['mode'],
  },
};

export const revisionsHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const mode = String(args.mode ?? 'list');
  try {
    if (mode === 'list') {
      const r = await call<{ count: number; revisions: Array<Record<string, unknown>> }>('listRevisions', {});
      if (!r.ok) return fail(errText('读取修订失败', r), r.error);
      const revs = r.data?.revisions ?? [];
      if (!revs.length) return ok('文档里没有修订记录。');
      const lines = revs.map((x) => {
        const t = Number(x.type);
        const kind = t === 1 ? '插入' : t === 2 ? '删除' : t === 3 ? '属性' : t === 4 ? '段落' : `类型${t}`;
        const text = String(x.text ?? '').replace(/\s+/g, ' ').slice(0, 50);
        return `  [${x.index}] ${kind} · ${x.author ?? '?'} · 第${x.page ?? '?'}页 · "${text}"`;
      });
      return ok(`共 ${r.data?.count ?? 0} 条修订：\n${lines.join('\n')}`);
    }
    const params: Args = { mode };
    if (mode === 'acceptOne' || mode === 'rejectOne') {
      if (args.index === undefined) return fail('单条处理必须传 index（从 1 开始）');
      params.index = Number(args.index);
    }
    const r = await call<{ before: number; after: number; handled: number }>('handleRevisions', params);
    if (!r.ok) return fail(errText('处理修订失败', r), r.error);
    const d = r.data;
    if (!d || d.before === 0) return ok('文档里没有修订，无需处理。');
    return ok(
      `已${mode === 'acceptAll' || mode === 'acceptOne' ? '接受' : '拒绝'}：` +
        `${d.handled} 条（${d.before} → ${d.after}）`
    );
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`处理修订异常：${m}`, m);
  }
};

// ============================================================================
// P0：表格单元格级
// ============================================================================

const tableProps = {
  tableIndex: { type: 'number', description: '第几个表格，从 1 开始，默认 1' },
  row: { type: 'number', description: '行号，从 1 开始' },
  col: { type: 'number', description: '列号，从 1 开始' },
} as const;

export const getCellTextDefinition: ToolDefinition = {
  name: 'wps_word_get_cell_text',
  description: `读取表格里某个单元格的文本。表格类操作前先读一下，确认坐标对不对。

参数：tableIndex（第几个表格，默认 1）、row、col（都从 1 开始）`,
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { target: WORD_TARGET_PROPERTY, ...tableProps }, required: ['row', 'col'] },
};

export const getCellTextHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call<{ row: number; col: number; text: string }>('getCellText', {
      tableIndex: args.tableIndex, row: args.row, col: args.col,
    });
    if (!r.ok) return fail(errText('读取单元格失败', r), r.error);
    return ok(`(${r.data?.row},${r.data?.col}) = "${r.data?.text ?? ''}"`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`读取单元格异常：${m}`, m);
  }
};

export const mergeCellsDefinition: ToolDefinition = {
  name: 'wps_word_merge_table_cells',
  description: `合并表格里的一块矩形区域。

参数：tableIndex、row1/col1（左上角）、row2/col2（右下角），行列都从 1 开始

注意：不能跨越已有的合并区，否则 WPS 会报错（这属于约束，不是 bug）。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      ...tableProps,
      row1: { type: 'number', description: '起始行' },
      col1: { type: 'number', description: '起始列' },
      row2: { type: 'number', description: '结束行' },
      col2: { type: 'number', description: '结束列' },
    },
    required: ['row1', 'col1', 'row2', 'col2'],
  },
};

export const mergeCellsHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call('mergeTableCells', {
      tableIndex: args.tableIndex, row1: args.row1, col1: args.col1, row2: args.row2, col2: args.col2,
    });
    if (!r.ok) return fail(errText('合并单元格失败', r), r.error);
    return ok(`已合并 ${args.row1},${args.col1} 到 ${args.row2},${args.col2}`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`合并单元格异常：${m}`, m);
  }
};

export const splitCellDefinition: ToolDefinition = {
  name: 'wps_word_split_cell',
  description: `拆分一个已合并的单元格。

参数：tableIndex、row、col（该单元格的位置）、numRows（拆成几行，默认 1）、numCols（拆成几列，默认 1）

注意：只能拆"已合并过"的单元格；没合并过的会报错。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      ...tableProps,
      numRows: { type: 'number', description: '拆成几行，默认 1' },
      numCols: { type: 'number', description: '拆成几列，默认 1' },
    },
    required: ['row', 'col'],
  },
};

export const splitCellHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call('splitCell', {
      tableIndex: args.tableIndex, row: args.row, col: args.col,
      numRows: args.numRows, numCols: args.numCols,
    });
    if (!r.ok) return fail(errText('拆分单元格失败', r), r.error);
    return ok(`已拆分 ${args.row},${args.col}（${args.numRows ?? 1} 行 × ${args.numCols ?? 1} 列）`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`拆分单元格异常：${m}`, m);
  }
};

export const tableRowColDefinition: ToolDefinition = {
  name: 'wps_word_table_row_col',
  description: `在表格里插入或删除行/列。

参数：
- dir: rows（行）| cols（列）
- op: insert | delete
- index: 在第几行/列之前或之后
- where: before | after（insert 时用；默认 after）
- count: 一次插几行/列，默认 1

使用场景：
- "在这个表格上面加两行"
- "把第 3 列删掉"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      tableIndex: tableProps.tableIndex,
      dir: { type: 'string', enum: ['rows', 'cols'], description: '操作行还是列' },
      op: { type: 'string', enum: ['insert', 'delete'], description: '插入还是删除' },
      index: { type: 'number', description: '第几行/列（从 1 开始）' },
      where: { type: 'string', enum: ['before', 'after'], description: '插在该行/列之前还是之后，默认 after' },
      count: { type: 'number', description: '数量，默认 1' },
    },
    required: ['dir', 'op', 'index'],
  },
};

export const tableRowColHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const dir = String(args.dir);
  const op = String(args.op);
  const action = op === 'delete' ? 'deleteTableRowCol' : 'insertTableRowCol';
  try {
    const r = await call<{ before?: { rows: number; cols: number }; after: { rows: number; cols: number } }>(
      action,
      { tableIndex: args.tableIndex, dir, index: args.index, where: args.where, count: args.count }
    );
    if (!r.ok) return fail(errText(`${op === 'delete' ? '删除' : '插入'}${dir === 'rows' ? '行' : '列'}失败`, r), r.error);
    const b = r.data?.before;
    const a = r.data?.after;
    const shape = a ? `现在 ${a.rows} 行 × ${a.cols} 列` : '';
    const from = b ? `（原 ${b.rows} 行 × ${b.cols} 列）` : '';
    return ok(`已${op === 'delete' ? '删除' : '插入'}${dir === 'rows' ? '行' : '列'} ${from}，${shape}`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`操作行列异常：${m}`, m);
  }
};

export const setCellBordersDefinition: ToolDefinition = {
  name: 'wps_word_set_cell_borders',
  description: `设置表格单元格的边框。

参数：
- tableIndex、row、col
- side: top | bottom | left | right | horizontal | vertical | all
- style: 线条样式，0=无边框，1=单线（默认），其他值见 Word 的 WdLineStyle

使用场景：
- "给这个表加框线"
- "去掉中间的横线"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      ...tableProps,
      side: {
        type: 'string',
        enum: ['top', 'bottom', 'left', 'right', 'horizontal', 'vertical', 'all'],
        description: '哪一侧的边框',
      },
      style: { type: 'number', description: '线条样式，0=无 1=单线（默认 1）' },
    },
    required: ['row', 'col', 'side'],
  },
};

export const setCellBordersHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call<{ applied: string[]; style: number }>('setCellBorders', {
      tableIndex: args.tableIndex, row: args.row, col: args.col, side: args.side, style: args.style,
    });
    if (!r.ok) return fail(errText('设置边框失败', r), r.error);
    const applied = r.data?.applied ?? [];
    return ok(
      `已设置 (${args.row},${args.col}) 的边框：${applied.join('/')}` +
        (r.data?.style === 0 ? '（无边框）' : '')
    );
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`设置边框异常：${m}`, m);
  }
};

// ============================================================================
// P1：样式 新建 / 修改 / 删除 / 列出
// ============================================================================

const styleFormatProps = {
  fontName: { type: 'string', description: '字体名（中英文都设）' },
  fontSize: { type: 'number', description: '字号（磅）' },
  bold: { type: 'boolean', description: '加粗' },
  italic: { type: 'boolean', description: '斜体' },
  color: { type: 'number', description: '颜色，0xBBGGRR 十进制（与 Word 的 BGR 一致，如蓝色 0xFF0000）' },
  spaceBefore: { type: 'number', description: '段前间距（磅）' },
  spaceAfter: { type: 'number', description: '段后间距（磅）' },
  lineSpacing: { type: 'number', description: '行距（磅）' },
  alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify'], description: '对齐方式' },
} as const;

export const listStylesDefinition: ToolDefinition = {
  name: 'wps_word_list_styles',
  description: `列出文档里的样式。

参数：onlyUserDefined（true=只看自定义样式，默认 false=全部）

使用场景：
- "我之前建的样式还在吗"
- "这个文档都有哪些样式"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      onlyUserDefined: { type: 'boolean', description: '只看自定义样式' },
    },
  },
};

export const listStylesHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call<{ count: number; styles: Array<{ name: string; builtIn: boolean; inUse?: boolean }> }>(
      'listStyles', { onlyUserDefined: args.onlyUserDefined }
    );
    if (!r.ok) return fail(errText('列出样式失败', r), r.error);
    const st = r.data?.styles ?? [];
    if (!st.length) return ok(args.onlyUserDefined ? '文档里没有自定义样式。' : '文档里没有样式。');
    const show = st.slice(0, 80);
    const lines = show.map((s) => `  ${s.builtIn ? '[内置]' : '[自定义]'} ${s.name}${s.inUse ? '（在用）' : ''}`);
    const more = st.length > show.length ? `\n  …还有 ${st.length - show.length} 个` : '';
    return ok(`共 ${st.length} 个样式：\n${lines.join('\n')}${more}`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`列出样式异常：${m}`, m);
  }
};

export const addStyleDefinition: ToolDefinition = {
  name: 'wps_word_add_style',
  description: `新建一个样式（不直接应用到任何段落）。

参数：
- name: 样式名（必填）
- baseOn: 基于哪个已有样式派生（可选，如"正文"）
- type: 1=段落样式（默认），2=字符样式
- 其余可选格式属性：fontName / fontSize / bold / italic / color / spaceBefore / spaceAfter / lineSpacing / alignment

注意：新建后要用 wps_word_apply_style 套到段落上才会生效。

⚠️ 颜色是 BGR 不是 RGB（Word 的 COM 接口如此）：蓝色是 0xFF0000，红色是 0x0000FF。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      name: { type: 'string', description: '样式名' },
      baseOn: { type: 'string', description: '基于哪个样式派生' },
      styleType: { type: 'string', enum: ['paragraph', 'character'], description: '段落样式（默认）或字符样式' },
      ...styleFormatProps,
    },
    required: ['name'],
  },
};

export const addStyleHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call<{ name: string; applied: string[] }>('addStyle', args as Args);
    if (!r.ok) return fail(errText('新建样式失败', r), r.error);
    const applied = r.data?.applied ?? [];
    return ok(
      `已新建样式「${r.data?.name ?? args.name}」` +
        (applied.length ? `，设置：${applied.join('/')}` : '（未指定格式属性，是空样式）') +
        '。接下来用 wps_word_apply_style 套到段落上。',
      // 样式建成了、但有几个属性本机 WPS 不接受 —— 这是"部分成功"不是失败，
      // 必须让调用方知道，否则模型会当全好了。
      r.warnings
    );
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`新建样式异常：${m}`, m);
  }
};

export const modifyStyleDefinition: ToolDefinition = {
  name: 'wps_word_modify_style',
  description: `修改已有样式的属性。

参数：
- name: 要改的样式名（必填）
- newName: 改成新名字（可选；内置样式通常不能改名，会如实告知已跳过）
- 其余格式属性见 add_style，不传的属性保持不变

使用场景：
- "把标题1的字号改成 16"
- "自定义样式里的中文和英文字体不一样，统一一下"`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      name: { type: 'string', description: '样式名' },
      newName: { type: 'string', description: '改成新名字' },
      ...styleFormatProps,
    },
    required: ['name'],
  },
};

export const modifyStyleHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call<{ name: string; applied: string[] }>('modifyStyle', args as Args);
    if (!r.ok) return fail(errText('修改样式失败', r), r.error);
    const applied = r.data?.applied ?? [];
    if (!applied.length) return fail(`样式「${args.name}」没有任何可改的属性被应用（属性名可能拼错了）`);
    return ok(`已修改样式「${r.data?.name}」：${applied.join('/')}`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`修改样式异常：${m}`, m);
  }
};

export const deleteStyleDefinition: ToolDefinition = {
  name: 'wps_word_delete_style',
  description: `删除一个自定义样式。

参数：name（样式名）

注意：内置样式不能删（会明确报错）。样式正被段落使用时也可能删不掉。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: { type: 'object', properties: { target: WORD_TARGET_PROPERTY, name: { type: 'string', description: '样式名' } }, required: ['name'] },
};

export const deleteStyleHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  try {
    const r = await call('deleteStyle', { name: args.name });
    if (!r.ok) return fail(errText('删除样式失败', r), r.error);
    return ok(`已删除样式「${args.name}」`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`删除样式异常：${m}`, m);
  }
};

// ============================================================================
// P1：交叉引用
// ============================================================================

export const crossRefDefinition: ToolDefinition = {
  name: 'wps_word_cross_reference',
  description: `在【当前光标位置】插入交叉引用（指向某个标题/书签），或列出可引用的目标。

参数：
- op: insert（插入，默认）| listTargets（列出可引用项）
- refType: heading | bookmark | numberedItem | footnote（默认 heading）
- itemIndex: 第几个可引用项，从 1 开始（默认 1）

使用场景：
- "在结论这段开头插一个对'第二章'的引用"
- 先用 op=listTargets 看有哪些能引的，再 itemIndex 选中

注意：
- 引用插在【光标处】，所以先 wps_word_insert_text 定位好光标再调用
- 插完建议调 wps_word_update_fields 让引用显示正确`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      op: { type: 'string', enum: ['insert', 'listTargets'], description: '插入引用还是列出可引用项' },
      refType: { type: 'string', enum: ['heading', 'bookmark', 'numberedItem', 'footnote'], description: '引用类型，默认 heading' },
      itemIndex: { type: 'number', description: '第几个可引用项（从 1 开始），默认 1' },
    },
    required: ['op'],
  },
};

export const crossRefHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const op = String(args.op);
  try {
    if (op === 'listTargets') {
      const r = await call<{ targets: Record<string, string[]> }>('listCrossReferenceTargets', {});
      if (!r.ok) return fail(errText('列出可引用项失败', r), r.error);
      const t = r.data?.targets ?? {};
      const lines = Object.entries(t).map(([k, v]) => `  ${k}: ${v.length} 个${v.length ? ` → ${v.slice(0, 8).join('、')}` : ''}`);
      const total = Object.values(t).reduce((a, b) => a + b.length, 0);
      if (!total) return ok('文档里没有可引用的目标。\n提示：heading 类型需要先有标题样式的段落。');
      return ok(`可引用项：\n${lines.join('\n')}\n用 itemIndex 选中其中一个再 op=insert。`);
    }
    const r = await call<{ target: string; total: number; itemIndex: number }>('insertCrossReference', {
      refType: args.refType, itemIndex: args.itemIndex,
    });
    if (!r.ok) return fail(errText('插入交叉引用失败', r), r.error);
    return ok(
      `已在光标处插入对「${r.data?.target}」的引用（第 ${r.data?.itemIndex}/${r.data?.total} 项）。` +
        '建议接着调 wps_word_update_fields 让引用显示正确。'
    );
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`交叉引用异常：${m}`, m);
  }
};

// ============================================================================
// P1：索引
// ============================================================================

export const indexDefinition: ToolDefinition = {
  name: 'wps_word_index',
  description: `给文字编索引条目，或在光标处插入索引目录。

参数：
- op: markEntry（标记索引项，默认）| insert（插入索引目录）| count（看已有多少索引项）
- text: 索引项文字（默认取选中内容）
- columns: 索引分几栏（insert 时，默认 1）

使用场景：
- "把选中的'深度学习'编入索引"
- "在文档末尾插入索引目录"

⚠️ 已知限制：本机 WPS 12.1 的 Indexes.MarkEntry 在探测中报 E_FAIL，
   若报"不支持"属预期情况，可改用 wps_word_index op=insert 插一个空索引再手工填。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      op: { type: 'string', enum: ['markEntry', 'insert', 'count'], description: '操作方式' },
      text: { type: 'string', description: '索引项文字（默认取选中内容）' },
      columns: { type: 'number', description: '索引分几栏（insert 时）' },
    },
    required: ['op'],
  },
};

export const indexHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const op = String(args.op);
  try {
    if (op === 'count') {
      const r = await call<{ markedEntries: number }>('insertIndex', {});
      // insertIndex 会真的插一个索引，这里只用它的计数做参考；改用 markEntry 的探测更干净
      return ok(`当前索引项数量：${r.ok ? (r.data?.markedEntries ?? 0) : '读取失败'}`);
    }
    if (op === 'insert') {
      const r = await call<{ markedEntries: number; columns: number }>('insertIndex', { columns: args.columns });
      if (!r.ok) return fail(errText('插入索引失败', r), r.error);
      return ok(
        `已在光标处插入索引目录（${r.data?.columns ?? 1} 栏）。` +
          `当前文档有 ${r.data?.markedEntries ?? 0} 个索引项。` +
          '接着调 wps_word_update_fields 更新。'
      );
    }
    const r = await call<{ entry: string }>('markIndexEntry', { entry: args.text });
    if (!r.ok) return fail(errText('标记索引项失败', r), r.error);
    return ok(`已编入索引项：「${r.data?.entry}」`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`索引操作异常：${m}`, m);
  }
};

// ============================================================================
// P1：页码 / 奇偶页 / 首页不同
// ============================================================================

export const pageNumberDefinition: ToolDefinition = {
  name: 'wps_word_page_number',
  description: `在页眉里插入页码，或设置"首页不同/奇偶页不同"。

参数：
- op: addPageNumber（插页码，默认）| setLayout（设置版式）
- which: primary（默认，正文页）| first（首页）| even（偶数页）
- alignment: left | center（默认）| right
- differentFirstPage: true/false（setLayout 时）—— 首页不显示页码就靠它
- oddAndEvenPages: true/false（setLayout 时）—— 奇偶页用不同页眉就靠它

使用场景：
- "加页码，页脚居中"→ wps_word_insert_header 之后调这个（页码实际插在页眉）
- "首页不要页码"→ setLayout differentFirstPage=true`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      op: { type: 'string', enum: ['addPageNumber', 'setLayout'], description: '操作方式' },
      which: { type: 'string', enum: ['primary', 'first', 'even'], description: '页码加在哪部分' },
      alignment: { type: 'string', enum: ['left', 'center', 'right'], description: '对齐，默认 right' },
      sectionIndex: { type: 'number', description: '第几个节，默认 1' },
      differentFirstPage: { type: 'boolean', description: '首页页眉页脚不同' },
      oddAndEvenPages: { type: 'boolean', description: '奇偶页页眉页脚不同' },
    },
    required: ['op'],
  },
};

export const pageNumberHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const op = String(args.op);
  try {
    if (op === 'setLayout') {
      if (args.differentFirstPage === undefined && args.oddAndEvenPages === undefined) {
        return fail('setLayout 至少要传 differentFirstPage 或 oddAndEvenPages 之一');
      }
      const r = await call<{ differentFirstPage: boolean; oddAndEvenPages: boolean }>('setHeaderFooterLayout', {
        differentFirstPage: args.differentFirstPage, oddAndEvenPages: args.oddAndEvenPages,
      });
      if (!r.ok) return fail(errText('设置版式失败', r), r.error);
      return ok(`已设置：首页不同=${r.data?.differentFirstPage}，奇偶页不同=${r.data?.oddAndEvenPages}`);
    }
    const r = await call<{ which: string; alignment: string; count: number }>('addPageNumber', {
      which: args.which, alignment: args.alignment, sectionIndex: args.sectionIndex,
    });
    if (!r.ok) return fail(errText('添加页码失败', r), r.error);
    return ok(
      `已插入页码（${r.data?.which ?? 'primary'}，${r.data?.alignment ?? 'right'}），` +
        `该部分现有 ${r.data?.count} 个页码。`
    );
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`页码操作异常：${m}`, m);
  }
};

// ============================================================================
// 修订模式开关（对标 word-mcp-live 的 native tracked changes 卖点）：
// 开启后 AI 的所有写入自动带修订留痕，用户逐条审阅、可接受可拒绝。
// 会话级状态（TrackRevisions 是文档属性，守护进程存活期间保持）。
// ============================================================================

export const reviewModeDefinition: ToolDefinition = {
  name: 'wps_word_review_mode',
  description: `开启/关闭/查询这份文档的修订模式。

开启后，AI 对文档的所有写入自动以修订留痕呈现——用户逐条审阅、可接受可拒绝，
适合合同、公文等"改了什么必须看得见"的场景。不传 enabled 为查询当前状态。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      enabled: { type: 'boolean', description: 'true=开启修订模式，false=关闭；不传=查询当前状态' },
    },
  },
};

export const reviewModeHandler: ToolHandler = async (args) => {
  const payload = args.enabled === undefined ? {} : { enabled: Boolean(args.enabled) };
  try {
    const r = await call<{ reviewMode: boolean; doc?: string }>('reviewMode', payload);
    if (!r.ok) return fail(errText('设置修订模式失败', r), r.error);
    return ok(`修订模式：${r.data?.reviewMode ? '已开启（之后的写入带修订留痕）' : '已关闭'}`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`修订模式异常：${m}`, m);
  }
};

// ============================================================================
// 逐步撤销（对标 word-mcp-live 的 per-action undo）：
// 按步撤销最近的编辑。AI 的破坏性写入已包进撤销记录（ps1 Start/Stop-AiUndoRecord），
// 用户或 AI 撤销时一次退回一整个动作。
// ============================================================================

export const undoDefinition: ToolDefinition = {
  name: 'wps_word_undo',
  description: `按步撤销这份文档里最近的编辑。

参数：
- steps: 撤销几步，默认 1，最多 100

适合 AI 刚做完一组修改、发现方向错了要回退的场景。
每步对应一个撤销记录；AI 的破坏性动作已按动作分组，一次退回一整步。`,
  category: ToolCategory.DOCUMENT,
  inputSchema: {
    type: 'object',
    properties: {
      target: WORD_TARGET_PROPERTY,
      steps: { type: 'number', description: '撤销步数，默认 1，最大 100' },
    },
  },
};

export const undoHandler: ToolHandler = async (args) => {
  const steps = args.steps === undefined ? 1 : Math.max(1, Math.min(100, Number(args.steps) || 1));
  try {
    const r = await call<{ requested: number; undone: number; doc?: string }>('undoSteps', { steps });
    if (!r.ok) return fail(errText('撤销失败', r), r.error);
    return ok(`已撤销 ${r.data?.undone ?? 0} 步（请求 ${r.data?.requested ?? steps} 步）`);
  } catch (e) {
    const m = e instanceof Error ? e.message : String(e);
    return fail(`撤销异常：${m}`, m);
  }
};

/** 本批工具集合 */
export const p01Tools: RegisteredTool[] = [
  { definition: revisionsDefinition, handler: revisionsHandler },
  { definition: getCellTextDefinition, handler: getCellTextHandler },
  { definition: mergeCellsDefinition, handler: mergeCellsHandler },
  { definition: splitCellDefinition, handler: splitCellHandler },
  { definition: tableRowColDefinition, handler: tableRowColHandler },
  { definition: setCellBordersDefinition, handler: setCellBordersHandler },
  { definition: listStylesDefinition, handler: listStylesHandler },
  { definition: addStyleDefinition, handler: addStyleHandler },
  { definition: modifyStyleDefinition, handler: modifyStyleHandler },
  { definition: deleteStyleDefinition, handler: deleteStyleHandler },
  { definition: crossRefDefinition, handler: crossRefHandler },
  { definition: indexDefinition, handler: indexHandler },
  { definition: pageNumberDefinition, handler: pageNumberHandler },
  // 2026-10-26 对标补齐：修订模式开关 + 逐步撤销（见文件尾部的两个定义）
  { definition: reviewModeDefinition, handler: reviewModeHandler },
  { definition: undoDefinition, handler: undoHandler },
];

export default p01Tools;