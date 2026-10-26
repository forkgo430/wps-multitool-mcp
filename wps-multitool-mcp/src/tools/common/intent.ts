/**
 * Input: 一句任务描述（可带 autoLoad）
 * Output: 该开哪些工具组 + 该先用哪些工具
 * Pos: 快速意图识别（俗称"多面手"的路由器）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ------------
 * 默认只列 40 个基础工具（见 src/server/tool-tiers.ts）。当任务超出手里的工具时，
 * 模型需要知道"该去开哪个组"。如果让它自己猜，往往是先瞎调几个不存在的工具、
 * 再靠报错回退 —— 浪费好几轮。这里给一个**一次调用就能拿到答案**的入口。
 *
 * 它是规则，不是 NLP（别吹）
 * -------------------------
 * 纯关键词匹配，零依赖、零延迟（<1ms）。**会漏、会误判**。所以：
 *   - 默认只给【建议】，不自动改工具列表（自动会让列表在对话中突变，缓存和上下文都乱）
 *   - 识别为空时明确说"没识别出来"，并指向 wps_search_tools（全量搜索）
 *   - 设置 WPS_INTENT_AUTOLOAD=1 或传 autoLoad=true 才真的加载并广播 list_changed
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { toolRegistry } from '../../server/tool-registry';

interface IntentRule {
  group: 'word' | 'excel' | 'ppt' | 'common';
  /** 命中即计分；【强信号】权重 2，弱信号 1 */
  strong: string[];
  weak: string[];
}

/**
 * 意图规则表。
 *
 * 为什么是关键词而不是 embedding / 小模型分类
 * ------------------------------------------
 * 调研过（GitHub + 业内实践，2026-10）三种做法的适用边界：
 *   关键词匹配 —— 简单、零延迟、零依赖；组间语义差异大时够准
 *   模型预分类 —— 多一次模型调用；只有"组间语义很接近"才划得来
 *   两阶段调用 —— 最准但延迟翻倍；只适合极复杂场景
 * 本项目的组是 ppt / excel / word / common，语义区分非常明显（幻灯片 vs 单元格 vs 文档），
 * 所以关键词匹配是对的选择 —— 上 embedding 属于过度设计。
 *
 * 能改进的是【词表质量】而不是【算法】：
 *   1) 同义词/口语说法：用户说"做个表"而不是"工作簿"，说"东西"而不是"文档"
 *   2) 强/弱分级：强词（如 "PPT"）几乎不会误判，弱词（如"文件"）容易
 *   3) 常见动词不列进去（"做个""加个"不带组信息，列了只会增加误匹配）
 *
 * 兜底（已实现）：识别为空时明确说"没识别出来"并指向 wps_search_tools，
 * 不装 —— 这是业内推荐的做法，比猜一个组要好。
 */
const INTENT_RULES: IntentRule[] = [
  {
    group: 'ppt',
    strong: ['ppt', 'pptx', '幻灯片', '演示文稿', 'slide', 'presentation', 'keynote'],
    weak: [
      '母版', '版式', '动画', '切换', '配色', '主题', '封面', '演讲', '汇报', '路演', '美化',
      '播放', '放映', '备注', '页脚', '占位符',
    ],
  },
  {
    group: 'excel',
    strong: [
      'excel', 'xlsx', 'xls', '单元格', '工作表', '工作簿', 'spreadsheet', 'workbook', 'sheet',
      '表格文件', '电子表格', '表格',
    ],
    weak: [
      '公式', '函数', '透视', '筛选', '排序', '行列', '台账', '报表', '数据表', '图表',
      '计算', '去重', '汇总', '求和', '平均值', 'vlookup', '条件格式', '冻结',
      '统计', '销售额', '占比', '增长率',
    ],
  },
  {
    group: 'word',
    strong: [
      'word', 'docx', 'doc', '文档', '正文', 'document', '文字文稿', '文稿', '表格',
    ],
    weak: [
      '标题', '目录', '题注', '脚注', '尾注', '样式', '排版', '页眉', '页脚', '页码',
      '论文', '摘要', '参考文献', '批注', '修订', '书签', '分栏', '大纲', '交叉引用', '索引',
      '封面', '致谢', '字数', '段落', '宋体', '黑体', '字体', '行距', '缩进', '超链接',
    ],
  },
  {
    group: 'common',
    strong: ['pdf', '导出', '转换', 'convert', '另存为'],
    weak: ['保存', '另存', '打印', '快照', '置前', '窗口', '复制一份', '备份'],
  },
];

export interface IntentMatch {
  group: string;
  score: number;
  /** 归一化到 0~1 的粗置信度（score/4 封顶） */
  confidence: number;
  matched: string[];
  /** 该组当前是否已"完整加载"（整组列出） */
  fullyLoaded: boolean;
  toolCount: number;
  /** 推荐先看的工具（该组内命中关键词的） */
  suggested: string[];
}

export interface IntentResult {
  task: string;
  tier: 'essential' | 'full';
  matches: IntentMatch[];
  needLoad: string[];
  autoLoaded: string[];
  note?: string;
  nextStep: string;
}

/** 纯函数：把一句话映射成候选组（不含任何 IO，便于测试） */
export function classifyIntent(task: string): IntentMatch[] {
  const text = (task || '').toLowerCase();
  const groups = new Map(toolRegistry.getGroups().map((g) => [g.group, g]));

  const out: IntentMatch[] = [];
  for (const rule of INTENT_RULES) {
    const matched: string[] = [];
    let score = 0;
    for (const kw of rule.strong) {
      if (text.includes(kw.toLowerCase())) {
        score += 2;
        matched.push(kw);
      }
    }
    for (const kw of rule.weak) {
      if (text.includes(kw.toLowerCase())) {
        score += 1;
        matched.push(kw);
      }
    }
    if (score === 0) continue;

    const g = groups.get(rule.group);
    // 组内推荐：用命中的关键词在该组里搜
    const seen = new Set<string>();
    const suggested: string[] = [];
    for (const kw of matched.slice(0, 3)) {
      for (const hit of toolRegistry.searchTools(kw, { group: rule.group, limit: 3 })) {
        if (!seen.has(hit.name)) {
          seen.add(hit.name);
          suggested.push(hit.name);
        }
        if (suggested.length >= 5) break;
      }
      if (suggested.length >= 5) break;
    }

    out.push({
      group: rule.group,
      score,
      confidence: Math.min(1, Number((score / 4).toFixed(2))),
      matched: [...new Set(matched)],
      fullyLoaded: g?.fullyLoaded ?? false,
      toolCount: g?.total ?? 0,
      suggested,
    });
  }

  return out.sort((a, b) => b.score - a.score);
}

export const routeIntentDefinition: ToolDefinition = {
  name: 'wps_route_intent',
  description: `意图路由（三层）：一句话进来，返回【能直接行动】的结果，而不是只给一个组名。

三层（按顺序匹配，命中即返回）：
1. 直答（direct，置信度 0.95）——高频单工具操作（保存/导出PDF/改单元格/冻结首行/加页码…），
   直接返回工具名 + 已从话里提取的参数，跳过"开组→搜索→选工具"的链路；
2. 任务包（task_package，置信度 0.85）——匹配到高频任务模板（周报/论文排版/合同修订/会议纪要…），
   返回步骤序列（每步带 app 与说明），按序执行即可；
3. 定组（group）——都匹配不上时降级到原有的组识别 + 搜索流程（原行为完整保留）。

低置信度（<0.5）时 alternatives 必非空，可以据此向用户追问。

注意：这是关键词规则，会漏判。定组层没识别出来时仍会建议 wps_search_tools（全量搜索）。`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: {
    type: 'object',
    properties: {
      task: { type: 'string', description: '任务描述（一句话即可，中文英文都行）' },
      autoLoad: {
        type: 'boolean',
        description: '仅定组层生效：是否直接加载命中的组（默认 false）。设 WPS_INTENT_AUTOLOAD=1 可全局打开',
      },
    },
    required: ['task'],
  },
};

// ============================================================================
// 三层意图（B5 意图升级）：短路口 / 任务包 / 定组
// ============================================================================

import * as fs from 'fs';
import * as path from 'path';

/** 短路口：一条话 → 一个工具（标准：一个工具能干完、参数可从话里提取、无歧义）。
 * 跨两个以上工具的（周报/论文…）不在此列——走任务包。 */
interface DirectRoute {
  /** 命中关键词（对 lower-case 后的话做 includes） */
  match: string[];
  tool: string;
  app: 'word' | 'excel' | 'ppt' | 'common';
  /** 参数模板：命中的 extract 键会从话里提取；固定键直接带 */
  params: Record<string, unknown>;
  /** 人话说明（写进 result.note） */
  note: string;
}

const CELL_ADDR = /\b([a-z]{1,3}\d{1,7})\b/i;

/** 短路口表。工具名全部经注册表核对真实存在（2026-10-27，见意图升级交付报告）。 */
const DIRECT_ROUTES: DirectRoute[] = [
  { match: ['保存一下', '保存文件', '保存文档', '保存表格', '保存这份'], tool: 'wps_common_save', app: 'common', params: {}, note: '保存当前文档' },
  { match: ['另存为', '存成新文件', '另存一份'], tool: 'wps_common_save_as', app: 'common', params: {}, note: '另存为（需要目标路径）' },
  { match: ['导出pdf', '转成pdf', '转pdf', '转成 pdf'], tool: 'wps_convert_to_pdf', app: 'common', params: {}, note: '原地转 PDF' },
  { match: ['读取单元格', '读单元格', '取单元格内容', '看看单元格'], tool: 'wps_excel_get_cell_value', app: 'excel', params: {}, note: '读单元格（话里的地址会提取进 params.cell）' },
  { match: ['写入单元格', '改单元格', '设置单元格', '填单元格'], tool: 'wps_excel_set_cell_value', app: 'excel', params: {}, note: '写单元格（地址提取进 params.cell，值需要补）' },
  { match: ['加粗'], tool: 'wps_word_set_font', app: 'word', params: { bold: true }, note: '选区加粗' },
  { match: ['斜体'], tool: 'wps_word_set_font', app: 'word', params: { italic: true }, note: '选区斜体' },
  { match: ['改字号', '字号改成', '设置字号'], tool: 'wps_word_set_font', app: 'word', params: {}, note: '改字号（具体号数需补）' },
  { match: ['字体颜色', '文字颜色'], tool: 'wps_word_set_text_color', app: 'word', params: {}, note: '改文字颜色（颜色值需补）' },
  { match: ['插入一行', '加一行'], tool: 'wps_excel_insert_rows', app: 'excel', params: {}, note: '插一行（位置需补）' },
  { match: ['插入一列', '加一列'], tool: 'wps_excel_insert_columns', app: 'excel', params: {}, note: '插一列（位置需补）' },
  { match: ['删除一行', '删掉一行'], tool: 'wps_excel_delete_rows', app: 'excel', params: {}, note: '删一行（位置需补）' },
  { match: ['删除一列', '删掉一列'], tool: 'wps_excel_delete_columns', app: 'excel', params: {}, note: '删一列（位置需补）' },
  { match: ['冻结首行'], tool: 'wps_excel_freeze_panes', app: 'excel', params: { rows: 1 }, note: '冻结首行' },
  { match: ['冻结首列'], tool: 'wps_excel_freeze_panes', app: 'excel', params: { columns: 1 }, note: '冻结首列' },
  { match: ['冻结窗口', '冻结表头'], tool: 'wps_excel_freeze_panes', app: 'excel', params: {}, note: '冻结（位置需补）' },
  { match: ['加页码', '插入页码'], tool: 'wps_word_page_number', app: 'word', params: {}, note: '插入页码' },
  { match: ['新建工作簿', '建个工作簿', '新建表格文件'], tool: 'wps_excel_create_workbook', app: 'excel', params: {}, note: '新建工作簿' },
  { match: ['新建文档', '建个文档', '新建空白文档'], tool: 'wps_word_create_document', app: 'word', params: {}, note: '新建文字文档' },
  { match: ['新建演示文稿', '新建ppt', '建个ppt'], tool: 'wps_ppt_create_presentation', app: 'ppt', params: {}, note: '新建演示文稿' },
  { match: ['加一页', '加页幻灯片', '插入幻灯片', '加一张幻灯片'], tool: 'wps_ppt_add_slide', app: 'ppt', params: {}, note: '加一页幻灯片' },
  { match: ['字数统计', '统计字数', '多少字'], tool: 'wps_word_get_document_stats', app: 'word', params: {}, note: '文档统计（字数/段落等）' },
  { match: ['连接测试', '测试连接', '检查连接'], tool: 'wps_check_connection', app: 'common', params: {}, note: '通道健康检查（合并后的规范名）' },
];

interface TaskPackage {
  task: string;
  title: string;
  match: string[];
  steps: Array<{ action: string; app: string; note?: string; inline?: boolean; snapshot?: boolean }>;
  confirm_before: string[];
}

/** 任务包真相源：src/server/task-packages.json（match 关键词的场景依据见该文件 _note）。
 * 加载多候选：tsc 不复制 json，dist 态与仓库源码态各试一个；全失败则任务包层跳过（定组层仍在）。 */
let taskPackagesCache: TaskPackage[] | null = null;
function loadTaskPackages(): TaskPackage[] {
  if (taskPackagesCache) return taskPackagesCache;
  const candidates = [
    path.join(__dirname, '..', '..', 'server', 'task-packages.json'), // dist/tools/common → dist/server
    path.join(__dirname, '..', '..', '..', 'src', 'server', 'task-packages.json'), // dist → 仓库 src
  ];
  for (const file of candidates) {
    try {
      taskPackagesCache = JSON.parse(fs.readFileSync(file, 'utf8')).packages as TaskPackage[];
      break;
    } catch {
      /* 试下一个候选 */
    }
  }
  if (!taskPackagesCache) taskPackagesCache = [];
  return taskPackagesCache;
}

export interface DirectMatch {
  tool: string;
  app: string;
  params: Record<string, unknown>;
  matched: string[];
  note: string;
}

export interface TaskPackageMatch {
  task: string;
  title: string;
  steps: TaskPackage['steps'];
  confirm_before: string[];
  matched: string[];
}

/** 纯函数：短路口匹配（不含 IO，便于测试）。命中即返回，多条命中取关键词数最多的。 */
export function matchDirectRoute(task: string): DirectMatch | null {
  const text = (task || '').toLowerCase();
  let best: { route: DirectRoute; hits: string[] } | null = null;
  for (const route of DIRECT_ROUTES) {
    const hits = route.match.filter((kw) => text.includes(kw.toLowerCase()));
    if (!hits.length) continue;
    if (!best || hits.length > best.hits.length) best = { route, hits };
  }
  if (!best) return null;
  const params: Record<string, unknown> = { ...best.route.params };
  // 参数提取：单元格地址（读/写单元格类话术里最常见的可提取参数）
  if (best.route.tool.startsWith('wps_excel_get_cell_value') || best.route.tool === 'wps_excel_set_cell_value') {
    const m = CELL_ADDR.exec(task);
    if (m) params.cell = m[1].toUpperCase();
  }
  return { tool: best.route.tool, app: best.route.app, params, matched: best.hits, note: best.route.note };
}

/** 纯函数：任务包匹配。多个命中取命中关键词最多的一个，其余进 alternatives。 */
export function matchTaskPackage(task: string): { best: TaskPackageMatch | null; alternatives: TaskPackageMatch[] } {
  const text = (task || '').toLowerCase();
  const scored: Array<{ pkg: TaskPackage; hits: string[] }> = [];
  for (const pkg of loadTaskPackages()) {
    const hits = pkg.match.filter((kw) => text.includes(kw.toLowerCase()));
    if (hits.length) scored.push({ pkg, hits });
  }
  scored.sort((a, b) => b.hits.length - a.hits.length);
  const toMatch = ({ pkg, hits }: { pkg: TaskPackage; hits: string[] }): TaskPackageMatch => ({
    task: pkg.task,
    title: pkg.title,
    steps: pkg.steps,
    confirm_before: pkg.confirm_before,
    matched: hits,
  });
  if (!scored.length) return { best: null, alternatives: [] };
  return {
    best: toMatch(scored[0]),
    alternatives: scored.slice(1).map(toMatch),
  };
}

export const routeIntentHandler: ToolHandler = async (args) => {
  const task = String(args.task ?? '').trim();
  if (!task) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '缺少 task 参数' }],
      error: 'missing task',
    };
  }

  const respond = (payload: Record<string, unknown>) => ({
    id: uuidv4(),
    success: true as const,
    content: [{ type: 'text' as const, text: JSON.stringify(payload, null, 2) }],
  });

  // ── 第 1 层：短路口（直觉直答）——一个工具能干完的高频操作，置信度 0.95
  const direct = matchDirectRoute(task);
  if (direct) {
    return respond({
      type: 'direct',
      confidence: 0.95,
      result: direct,
      nextStep: `直接调 ${direct.tool}。params 里是已从话里提取的参数，其余参数按该工具 schema 补全。`,
    });
  }

  // ── 第 2 层：任务包（导航）——多步骤高频任务，置信度 0.85
  const { best: pkg, alternatives: pkgAlts } = matchTaskPackage(task);
  if (pkg) {
    return respond({
      type: 'task_package',
      confidence: 0.85,
      result: pkg,
      alternatives: pkgAlts,
      nextStep: `按任务包「${pkg.title}」的 steps 逐个执行（每步带 app 与说明）；confirm_before 里的动作执行前先向用户确认。`,
    });
  }

  // ── 第 3 层：定组（原有行为完整保留，回归不断）
  const matches = classifyIntent(task);
  const needLoad = matches.filter((m) => !m.fullyLoaded).map((m) => m.group);
  const autoLoadRequested =
    args.autoLoad === true || ['1', 'true', 'yes', 'on'].includes(String(process.env.WPS_INTENT_AUTOLOAD || '').toLowerCase());

  const autoLoaded: string[] = [];
  const loadErrors: string[] = [];
  if (autoLoadRequested) {
    for (const g of needLoad) {
      const r = toolRegistry.loadGroup(g);
      if (r.ok) {
        if (r.changed) autoLoaded.push(g);
      } else {
        loadErrors.push(`${g}: ${r.error}`);
      }
    }
  }

  const nextStep = !matches.length
    ? '没识别出任何工具组。改用 wps_search_tools（在全部工具里搜）或 wps_list_tool_groups 看全貌。'
    : autoLoaded.length
      ? `已自动加载 ${autoLoaded.join('、')}，工具列表已刷新，直接用下面建议的工具即可。`
      : needLoad.length
        ? `先调 wps_load_tool_group（${needLoad.join(' 或 ')}）把工具拉出来；只在基础档里的话下面这些就够用。`
        : '所需工具组都已在列，直接用下面建议的工具。';

  const result: IntentResult & { loadErrors?: string[] } = {
    task,
    tier: toolRegistry.tier,
    matches,
    needLoad,
    autoLoaded,
    nextStep,
    ...(loadErrors.length ? { loadErrors } : {}),
    ...(matches.length
      ? {}
      : { note: '关键词规则未命中 —— 是它的已知短板，不是"这个任务没有工具"，请改用 wps_search_tools' }),
  };

  // 定组层的新包装：type/confidence/alternatives 与三层格式对齐（result 即原 IntentResult，
  // 老调用方读 result.matches / result.needLoad 的路径不变——回归不断）。
  const confidence = matches.length ? matches[0].confidence : 0;
  let alternatives: Array<{ group: string; reason: string }> = matches
    .slice(1)
    .map((m) => ({ group: m.group, reason: `命中关键词：${m.matched.join('、')}` }));
  if (confidence < 0.5 && alternatives.length === 0) {
    // 低置信度必须给出可追问的备选——组关键词没命中时给全组粗信息
    alternatives = [
      { group: 'word', reason: '涉及文档/文字内容' },
      { group: 'excel', reason: '涉及表格/数据计算' },
      { group: 'ppt', reason: '涉及演示/幻灯片' },
      { group: 'common', reason: '涉及保存/导出/转换' },
    ];
  }

  return respond({
    type: 'group',
    confidence,
    result,
    alternatives,
    // nextStep 提到顶层，与 direct/task_package 两层格式对齐（result 内保留原字段，老读法不断）
    nextStep,
  });
};

/** 意图识别工具（core 组，永远列出） */
export const intentTools: RegisteredTool[] = [
  { definition: routeIntentDefinition, handler: routeIntentHandler },
];

export default intentTools;
