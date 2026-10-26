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
  description: `快速意图识别：给一句任务描述，告诉你需要开启哪些工具组、先用哪些工具。

为什么需要它：默认只列出 40 个基础工具（够"打开/读写/导出"），任务一旦超出手里的工具，
不要瞎猜工具名 —— 先调这个，它会说要开哪个组，再用 wps_load_tool_group 拉起来。

它同时有 **plan_tool_groups** 的称呼，本质是一回事：把任务映射到工具组。

使用场景：
- "我要做一份产品发布会的 PPT" → 会告诉你 ppt 组没加载
- "这份 excel 要做透视表" → 指向 excel 组
- "帮我把论文的题注和交叉引用整理一下" → 指向 word 组

注意：这是关键词规则，会漏判。没识别出来时会让你改用 wps_search_tools（全量搜索）。`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: {
    type: 'object',
    properties: {
      task: { type: 'string', description: '任务描述（一句话即可，中文英文都行）' },
      autoLoad: {
        type: 'boolean',
        description: '是否直接加载命中的组（默认 false，只给建议）。设 WPS_INTENT_AUTOLOAD=1 可全局打开',
      },
    },
    required: ['task'],
  },
};

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

  return {
    id: uuidv4(),
    success: true,
    content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
  };
};

/** 意图识别工具（core 组，永远列出） */
export const intentTools: RegisteredTool[] = [
  { definition: routeIntentDefinition, handler: routeIntentHandler },
];

export default intentTools;
