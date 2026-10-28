/**
 * Input: DeckSpec（JSON 页面描述，自有 schema，字段名自定）——见 DeckSpec/各页型注释
 * Output: build_ppt_deck 工具：把页面描述逐动作落成真 PPT（走 toolRegistry.callTool 真入口）
 * Pos: deckkit 引擎骨架（第 3 期 E2，方案 docs/agents/BORROW-PLAN.md §2.5 第二步）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 架构（方案 §2.2）：布局计算（坐标/字号/密度）全部在本 TS 层——可 mock、可进门禁；
 *   本模块【不碰 ps1】：展开成的是既有 ppt 动作（add_textbox/add_shape/set_shape_gradient/…），
 *   逐动作经 toolRegistry.callTool 走真入口——快照护栏照常（红线），性能优化归 E3。
 *
 * shapeIndex 约定（重要）：每页从空白页（layout=blank，0 形状）开始，展开序列里
 *   add_shape / add_textbox / insert_ppt_chart 每个都 +1；引用型动作
 *   （set_font_color / set_shape_gradient / set_ppt_chart_data）用【字面序号】——
 *   因为每页型的形状序列是静态可数的，字面值比运行时计数更可断言（mock 单测逐个钉）。
 *
 * 素材路线（§2.2/第一步探针）：COM 原生形状优先、零新增依赖——
 *   渐变横幅 = set_shape_gradient（TwoColorGradient，探针 seg1 可用）；
 *   表格 = 矩形拼（第一步探针坑 5：AddTable 边框不可靠，禁走真表格）；
 *   行高系数 = 1.2（探针 seg3 本机标定——【改动前须真机复核】，勿抄别处数值）。
 *
 * 数据约定（chart 页）：insert_ppt_chart 的 data 参数不生效（chart-flow.ts 头注），必须
 *   insert → set_ppt_chart_data 两段式；chartIndex = 页内形状序号（图表追加在最后）。
 *
 * schema 是自研的（防抄态势 BORROW-PLAN §4.4）：字段名 kind/heading/points/milestones 等
 *   与任何外部引擎不同源；页型命名体系自定。
 */

import { v4 as uuidv4 } from 'uuid';
import { ToolCategory } from '../../types/tools';
import { toolRegistry } from '../../server/tool-registry';
import type { ToolDefinition, ToolHandler, RegisteredTool } from '../../types/tools';

// ============================================================================
// 排版纪律常量（一处集中；【本机标定值，改动前须真机复核】——lineFactor 来自引擎探针 seg3）
// ============================================================================

export const DECK_TYPO = {
  slideWidth: 960,
  slideHeight: 540,
  margin: 42,
  titleBandTop: 30,
  titleBandHeight: 64,
  titleFont: 30,
  sectionFont: 40,
  closingFont: 44,
  heroFont: 36,
  subFont: 20,
  bodyTop: 128,
  bodyFont: 16,
  /** 正文字号下限（排版纪律：低于此值可读性崩） */
  bodyFontMin: 14,
  cardGap: 18,
  /** 多行行高系数：引擎探针 seg3 实测 24 号字行高 28.8pt → 1.2。改动前须真机复核 */
  lineFactor: 1.2,
  /** 每页信息密度上限（超出直接拒绝，不静默截断） */
  maxBullets: 6,
  maxComparePoints: 4,
  maxTableRows: 6,
  maxTableCols: 4,
  maxTimeline: 6,
  maxSlides: 30,
  /** 学术蓝（首发主题；政务红/商务灰归 E3） */
  theme: {
    deep: '#1F3A5F',
    mid: '#2E5C8A',
    accent: '#C8A24B',
    paper: '#F5F7FA',
    ink: '#1A1A1A',
    fontPair: { heading: '微软雅黑', body: '微软雅黑' },
  },
} as const;

// ============================================================================
// DeckSpec（自有 schema）
// ============================================================================

export interface DeckSlideBase {
  heading?: string;
  footer?: string;
}
export type DeckSlide =
  | ({ kind: 'title'; heading: string; sub?: string } & DeckSlideBase)
  | ({ kind: 'section'; ordinal: string; heading: string; tagline?: string } & DeckSlideBase)
  | ({ kind: 'bullets'; heading: string; points: string[] } & DeckSlideBase)
  | ({ kind: 'compare'; heading: string; left: { label: string; points: string[] }; right: { label: string; points: string[] } } & DeckSlideBase)
  | ({ kind: 'imageText'; heading: string; imagePath: string; points?: string[] } & DeckSlideBase)
  | ({ kind: 'table'; heading: string; head: string[]; rows: string[][] } & DeckSlideBase)
  | ({ kind: 'chart'; heading: string; categories: string[]; series: Array<{ name: string; values: number[] }>; chartType?: number } & DeckSlideBase)
  | ({ kind: 'timeline'; heading: string; milestones: Array<{ label: string; note?: string }> } & DeckSlideBase)
  | ({ kind: 'closing'; heading: string; contact?: string } & DeckSlideBase)
  | ({ kind: 'quote'; quote: string; attribution?: string } & DeckSlideBase);

export interface DeckSpec {
  slides: DeckSlide[];
}

/** 展开产物：一步 = 一次既有 ppt 工具调用（走 callTool 真入口） */
export interface DeckStep {
  tool: string;
  args: Record<string, unknown>;
  note?: string;
}

const KNOWN_KINDS = ['title', 'section', 'bullets', 'compare', 'imageText', 'table', 'chart', 'timeline', 'closing', 'quote'] as const;

// ============================================================================
// 展开器（纯函数：DeckSpec → DeckStep[][]，坐标/字号/shapeIndex 全在此算）
// ============================================================================

function textbox(slide: number, o: {
  left: number; top: number; width: number; height: number; text: string;
  fontSize?: number; fontName?: string;
}): DeckStep {
  return {
    tool: 'wps_ppt_add_textbox',
    args: {
      slideIndex: slide, left: o.left, top: o.top, width: o.width, height: o.height, text: o.text,
      ...(o.fontSize ? { fontSize: o.fontSize } : {}), ...(o.fontName ? { fontName: o.fontName } : {}),
    },
  };
}

function color(slide: number, shapeIndex: number, hex: string): DeckStep {
  return { tool: 'wps_ppt_set_font_color', args: { slideIndex: slide, shapeIndex, color: hex } };
}

/** 标题色带（形状#1）+ 标题文本框（形状#2，白字） */
function headingBand(slide: number, heading: string, T: typeof DECK_TYPO): DeckStep[] {
  return [
    { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: 0, top: T.titleBandTop, width: T.slideWidth, height: T.titleBandHeight, fillColor: T.theme.deep }, note: '标题色带' },
    { tool: 'wps_ppt_add_textbox', args: { slideIndex: slide, left: T.margin, top: T.titleBandTop + 8, width: T.slideWidth - T.margin * 2, height: T.titleBandHeight - 16, text: heading, fontSize: T.titleFont, fontName: T.theme.fontPair.heading } },
    color(slide, 2, '#FFFFFF'),
  ];
}

/** 全页渐变底（形状#1） */
function fullBleedGradient(slide: number, from: string, to: string, T: typeof DECK_TYPO): DeckStep[] {
  return [
    { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: 0, top: 0, width: T.slideWidth, height: T.slideHeight, fillColor: from }, note: '全页底' },
    { tool: 'wps_ppt_set_shape_gradient', args: { slideIndex: slide, shapeIndex: 1, gradient: { type: 'linear', angle: 90, stops: [{ color: from, position: 0 }, { color: to, position: 1 }] } } },
  ];
}

const pageSteps: Record<string, (slide: number, s: any, T: typeof DECK_TYPO) => DeckStep[]> = {
  // 形状序列（每页）：#1 渐变底，#2 标题，#3 副标，#4 页脚
  title: (slide, s, T) => {
    const steps: DeckStep[] = [
      ...fullBleedGradient(slide, T.theme.deep, T.theme.mid, T),
      textbox(slide, { left: 80, top: 200, width: 800, height: 90, text: s.heading, fontSize: T.heroFont, fontName: T.theme.fontPair.heading }),
      color(slide, 2, '#FFFFFF'),
    ];
    if (s.sub) {
      steps.push(textbox(slide, { left: 80, top: 300, width: 800, height: 50, text: s.sub, fontSize: T.subFont, fontName: T.theme.fontPair.body }));
      steps.push(color(slide, 3, T.theme.accent));
    }
    if (s.footer) {
      steps.push(textbox(slide, { left: 80, top: 480, width: 800, height: 30, text: s.footer, fontSize: 12 }));
      steps.push(color(slide, 4, '#DDE6F2'));
    }
    return steps;
  },

  // #1 竖条，#2 序号，#3 标题，#4 标语
  section: (slide, s, T) => {
    const steps: DeckStep[] = [
      { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: 0, top: 0, width: 24, height: T.slideHeight, fillColor: T.theme.accent }, note: '章节竖条' },
      textbox(slide, { left: 70, top: 200, width: 160, height: 90, text: s.ordinal, fontSize: T.sectionFont, fontName: T.theme.fontPair.heading }),
      color(slide, 2, T.theme.accent),
      textbox(slide, { left: 240, top: 210, width: 640, height: 80, text: s.heading, fontSize: T.titleFont, fontName: T.theme.fontPair.heading }),
    ];
    if (s.tagline) {
      steps.push(textbox(slide, { left: 240, top: 300, width: 640, height: 40, text: s.tagline, fontSize: T.bodyFont, fontName: T.theme.fontPair.body }));
      steps.push(color(slide, 4, '#666666'));
    }
    return steps;
  },

  // #1 色带，#2 标题（白）；要点卡 i：色块 2i+1、文本 2i+2
  bullets: (slide, s, T) => {
    if (!s.points?.length) throw new Error('bullets 页需要 points');
    if (s.points.length > T.maxBullets) throw new Error(`要点数 ${s.points.length} 超过每页密度上限 ${T.maxBullets}——拆页`);
    const gap = (T.slideHeight - T.bodyTop - 30) / s.points.length;
    const cards = s.points.map((pt: string, i: number) => {
      const top = T.bodyTop + i * gap;
      return [
        { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: T.margin, top, width: T.slideWidth - T.margin * 2, height: gap - T.cardGap, fillColor: T.theme.paper } },
        textbox(slide, { left: T.margin + 16, top: top + (gap - T.cardGap - T.bodyFont * T.lineFactor) / 2, width: T.slideWidth - T.margin * 2 - 32, height: T.bodyFont * T.lineFactor, text: pt, fontSize: Math.max(T.bodyFont, T.bodyFontMin), fontName: T.theme.fontPair.body }),
      ];
    }).flat();
    return [...headingBand(slide, s.heading, T), ...cards];
  },

  // #1/#2 标题带；左：色块3/label4/points5(墨)；右：色块6/label7/points8(墨)
  compare: (slide, s, T) => {
    const check = (col_: { label?: string; points?: string[] }): { label: string; points: string[] } => {
      if (!col_?.label || !col_?.points?.length) throw new Error('compare 页左右两栏各需要 label 与 points');
      if (col_.points.length > T.maxComparePoints) throw new Error(`每栏要点 ${col_.points.length} 超上限 ${T.maxComparePoints}`);
      return col_ as { label: string; points: string[] };
    };
    const L = check(s.left); const R = check(s.right);
    const w = (T.slideWidth - T.margin * 2 - T.cardGap) / 2;
    const col = (x: number, col_: { label: string; points: string[] }, color0: string, base: number): DeckStep[] => [
      { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: x, top: T.bodyTop, width: w, height: 60, fillColor: color0 } },
      textbox(slide, { left: x + 12, top: T.bodyTop + 10, width: w - 24, height: 40, text: col_.label, fontSize: 18, fontName: T.theme.fontPair.heading }),
      textbox(slide, { left: x + 12, top: T.bodyTop + 70, width: w - 24, height: 280, text: col_.points.map((p, i) => `${i + 1}. ${p}`).join('\n'), fontSize: T.bodyFontMin, fontName: T.theme.fontPair.body }),
      color(slide, base, T.theme.ink),
    ];
    return [
      ...headingBand(slide, s.heading, T),
      ...col(T.margin, L, T.theme.mid, 5),
      ...col(T.margin + w + T.cardGap, R, T.theme.accent, 8),
    ];
  },

  // #1/#2 标题带，#3 图
  imageText: (slide, s, T) => {
    if (!s.imagePath) throw new Error('imageText 页需要 imagePath');
    return [
      ...headingBand(slide, s.heading, T),
      { tool: 'wps_ppt_insert_ppt_image', args: { slideIndex: slide, imagePath: s.imagePath, left: T.margin, top: T.bodyTop, width: 420, height: 340 } },
      textbox(slide, { left: T.margin + 440, top: T.bodyTop, width: T.slideWidth - T.margin * 2 - 440, height: 340, text: (s.points || []).map((p: string, i: number) => `${i + 1}. ${p}`).join('\n'), fontSize: T.bodyFont, fontName: T.theme.fontPair.body }),
    ];
  },

  // #1/#2 标题带；表头 j：色块 3+2j、文本 4+2j；数据行 (i,j)：色块、文本（无单独上色）
  table: (slide, s, T) => {
    if (!s.head?.length) throw new Error('table 页需要 head（表头行）');
    if (s.head.length > T.maxTableCols) throw new Error(`列数 ${s.head.length} 超上限 ${T.maxTableCols}——矩形拼表格的密度纪律`);
    if (!s.rows?.length || s.rows.length > T.maxTableRows) throw new Error(`数据行需 1~${T.maxTableRows} 行（got ${s.rows?.length ?? 0}）`);
    if (s.rows.some((r: string[]) => r.length !== s.head.length)) throw new Error('数据行列数与表头不一致');
    const w = (T.slideWidth - T.margin * 2) / s.head.length;
    const headH = 40; const rowH = 40;
    const steps: DeckStep[] = [...headingBand(slide, s.heading, T)];
    s.head.forEach((h: string, j: number) => {
      steps.push({ tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: T.margin + j * w, top: T.bodyTop, width: w, height: headH, fillColor: T.theme.deep } });
      steps.push(textbox(slide, { left: T.margin + j * w + 6, top: T.bodyTop + 6, width: w - 12, height: headH - 12, text: h, fontSize: T.bodyFont, fontName: T.theme.fontPair.heading }));
      steps.push(color(slide, 3 + 2 * j, '#FFFFFF'));
    });
    s.rows.forEach((row: string[], i: number) => {
      const top = T.bodyTop + headH + i * rowH;
      row.forEach((cell: string, j: number) => {
        steps.push({ tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: T.margin + j * w, top, width: w, height: rowH, fillColor: i % 2 === 0 ? T.theme.paper : '#FFFFFF' } });
        steps.push(textbox(slide, { left: T.margin + j * w + 6, top: top + 6, width: w - 12, height: rowH - 12, text: String(cell), fontSize: T.bodyFontMin, fontName: T.theme.fontPair.body }));
      });
    });
    return steps;
  },

  // #1/#2 标题带，#3 图表（insert 后 #3 就是图表本身——追加在页尾）
  chart: (slide, s, T) => {
    if (!s.categories?.length || !s.series?.length) throw new Error('chart 页需要 categories 与 series');
    return [
      ...headingBand(slide, s.heading, T),
      { tool: 'wps_ppt_insert_ppt_chart', args: { slideIndex: slide, chartType: s.chartType ?? 51, left: T.margin, top: T.bodyTop, width: T.slideWidth - T.margin * 2, height: T.slideHeight - T.bodyTop - 30, data: { categories: s.categories, series: s.series } }, note: 'data 在此不生效（chart-flow.ts 头注），仅建形状' },
      { tool: 'wps_ppt_set_ppt_chart_data', args: { slideIndex: slide, chartIndex: 3, data: { categories: s.categories, series: s.series } }, note: '真数据二段式（chartIndex=3：标题带两形状后图表追加页尾）' },
    ];
  },

  // #1/#2 标题带，#3 主线；里程碑 i：节点 4+2i、label 5+2i、note（交替上/下，note 时 +1）
  timeline: (slide, s, T) => {
    if (!s.milestones?.length) throw new Error('timeline 页需要 milestones');
    if (s.milestones.length > T.maxTimeline) throw new Error(`里程碑 ${s.milestones.length} 超上限 ${T.maxTimeline}`);
    const axisY = T.bodyTop + 140;
    const span = T.slideWidth - T.margin * 2 - 60;
    const step = span / Math.max(s.milestones.length - 1, 1);
    const steps: DeckStep[] = [
      ...headingBand(slide, s.heading, T),
      { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: T.margin + 30, top: axisY, width: span, height: 4, fillColor: T.theme.mid }, note: '时间轴主线' },
    ];
    s.milestones.forEach((m: { label: string; note?: string }, i: number) => {
      const x = T.margin + 30 + i * step;
      steps.push({ tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: x - 8, top: axisY - 10, width: 16, height: 24, fillColor: T.theme.accent }, note: `节点${i + 1}` });
      const above = i % 2 === 0;
      steps.push(textbox(slide, { left: x - 55, top: above ? axisY - 60 : axisY + 30, width: 130, height: 40, text: m.label, fontSize: T.bodyFont, fontName: T.theme.fontPair.heading }));
      if (m.note) steps.push(textbox(slide, { left: x - 55, top: above ? axisY + 30 : axisY - 60, width: 130, height: 40, text: m.note, fontSize: T.bodyFontMin, fontName: T.theme.fontPair.body }));
    });
    return steps;
  },

  // #1 渐变底，#2 heading，#3 contact
  closing: (slide, s, T) => {
    const steps: DeckStep[] = [
      ...fullBleedGradient(slide, T.theme.mid, T.theme.deep, T),
      textbox(slide, { left: 80, top: 210, width: 800, height: 100, text: s.heading, fontSize: T.closingFont, fontName: T.theme.fontPair.heading }),
      color(slide, 2, '#FFFFFF'),
    ];
    if (s.contact) {
      steps.push(textbox(slide, { left: 80, top: 330, width: 800, height: 40, text: s.contact, fontSize: T.bodyFont }));
      steps.push(color(slide, 3, T.theme.accent));
    }
    return steps;
  },

  // #1 竖条，#2 引文，#3 出处
  quote: (slide, s, T) => {
    if (!s.quote) throw new Error('quote 页需要 quote（引文内容）');
    const steps: DeckStep[] = [
      { tool: 'wps_ppt_add_shape', args: { slideIndex: slide, shapeType: 'rectangle', left: 0, top: 0, width: 14, height: T.slideHeight, fillColor: T.theme.accent }, note: '引文竖条' },
      textbox(slide, { left: 90, top: 180, width: 780, height: 160, text: '「' + s.quote + '」', fontSize: 28, fontName: T.theme.fontPair.heading }),
      color(slide, 2, T.theme.ink),
    ];
    if (s.attribution) {
      steps.push(textbox(slide, { left: 90, top: 360, width: 780, height: 40, text: '—— ' + s.attribution, fontSize: T.bodyFont, fontName: T.theme.fontPair.body }));
      steps.push(color(slide, 3, '#666666'));
    }
    return steps;
  },
};

/** 纯函数：DeckSpec → 每页步骤。抛错 = 描述违反排版纪律（密度/缺字段/未知 kind），handler 转如实报错 */
export function expandDeck(spec: DeckSpec): DeckStep[][] {
  if (!spec?.slides?.length) throw new Error('deck.slides 为空：至少给一页');
  if (spec.slides.length > DECK_TYPO.maxSlides) {
    throw new Error(`页数 ${spec.slides.length} 超过上限 ${DECK_TYPO.maxSlides}`);
  }
  return spec.slides.map((s, i) => {
    const fn = pageSteps[s.kind];
    if (!fn) {
      throw new Error(`未知页型「${s.kind}」（第 ${i + 1} 页）。支持：${KNOWN_KINDS.join(' / ')}`);
    }
    return fn(i + 1, s, DECK_TYPO);
  });
}

/** 诊断：整套 deck 的步骤总览（handler 返回里带，便于客户端复核） */
export function deckSummary(pages: DeckStep[][]): Array<{ slide: number; steps: number; tools: string[] }> {
  return pages.map((steps, i) => ({
    slide: i + 1,
    steps: steps.length,
    tools: [...new Set(steps.map((x) => x.tool))],
  }));
}

// ============================================================================
// 工具：wps_ppt_build_deck（不进默认档；逐动作走 callTool 真入口，快照护栏照常）
// ============================================================================

export const buildDeckDefinition: ToolDefinition = {
  name: 'wps_ppt_build_deck',
  description: `按 JSON 页面描述一次性生成整套演示文稿（deckkit 引擎）。

页面 kind（10 种）：title / section / bullets / compare / imageText / table / chart / timeline / closing / quote。
- title: { heading, sub? } —— 渐变全页 + 大标题
- section: { ordinal, heading, tagline? } —— 章节隔页
- bullets: { heading, points[≤6] } —— 要点卡片（超出密度上限会拒绝，拆页）
- compare: { heading, left:{label,points[≤4]}, right:{...} } —— 两栏对比
- imageText: { heading, imagePath, points? } —— 左图右文
- table: { heading, head[], rows[≤6][][] } —— 矩形拼表格（真表格边框不可靠，引擎不走 AddTable）
- chart: { heading, categories[], series[{name,values}] } —— 插入图表并写入真数据（两段式）
- timeline: { heading, milestones[{label,note?}][≤6] } —— 时间轴
- closing: { heading, contact? } —— 结尾页
- quote: { quote, attribution? } —— 引文页

行为：新建演示文稿 → 逐页铺排（每页从空白页起）→ 逐步骤走真工具入口（快照护栏照常）。
任一步失败即停并如实报完成到第几页第几步——不静默续跑。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      deck: {
        type: 'object',
        description: '页面描述：{ slides: [ {kind:"title",heading,...}, {kind:"bullets",heading,points}, ... ] }，kind 与字段见工具描述；slides 上限 30 页',
      },
    },
    required: ['deck'],
  },
};

export const buildDeckHandler: ToolHandler = async (args) => {
  const spec = args.deck as DeckSpec | undefined;
  let pages: DeckStep[][];
  try {
    pages = expandDeck(spec as DeckSpec);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: `deck 描述不合法：${msg}` }],
      error: msg,
    };
  }
  const total = pages.reduce((n, p) => n + p.length, 0) + pages.length; // 每页一次 add_slide
  const failures: Array<{ slide: number; index: number; tool: string; error: string }> = [];
  let executed = 0;
  let stoppedAtSlide = 0;

  // 第一步：新建演示文稿（deck 是新建语义；追加到已有文稿归 render/编辑类工具）
  const create = await toolRegistry.callTool({ id: uuidv4(), name: 'wps_ppt_create_presentation', arguments: {} });
  if (!create.success) {
    const msg = (create.content?.[0] as any)?.text || 'create_presentation 失败';
    return { id: uuidv4(), success: false, content: [{ type: 'text', text: msg }], error: msg };
  }
  executed += 1;

  for (let i = 0; i < pages.length; i++) {
    stoppedAtSlide = i + 1;
    // 每页先铺空白页（layout blank = 无占位符，形状计数从 0 起——展开器按此推 shapeIndex）
    const add = await toolRegistry.callTool({ id: uuidv4(), name: 'wps_ppt_add_slide', arguments: { layout: 'blank' } });
    executed += 1;
    if (!add.success) {
      failures.push({ slide: i + 1, index: 0, tool: 'wps_ppt_add_slide', error: (add.content?.[0] as any)?.text || 'add_slide 失败' });
      break;
    }
    let pageBroken = false;
    for (let j = 0; j < pages[i].length; j++) {
      const step = pages[i][j];
      const r = await toolRegistry.callTool({ id: uuidv4(), name: step.tool, arguments: step.args });
      executed += 1;
      if (!r.success) {
        failures.push({ slide: i + 1, index: j, tool: step.tool, error: (r.content?.[0] as any)?.text || '步骤失败' });
        pageBroken = true;
        break;
      }
    }
    if (pageBroken) break; // fail-fast：如实报停在哪页哪步，不静默续跑
  }

  const body = failures.length
    ? `生成中止：完成到第 ${stoppedAtSlide} 页（执行 ${executed}/${total} 步）。失败明细：${failures.map((f) => `第${f.slide}页#${f.index} ${f.tool}: ${f.error.slice(0, 80)}`).join('；')}`
    : `已生成 ${pages.length} 页（共 ${total} 步）。可用 wps_ppt_export_slide_as_image 逐页出图复核。`;

  return {
    id: uuidv4(),
    success: failures.length === 0,
    content: [{ type: 'text', text: body + '\n' + JSON.stringify(deckSummary(pages)) }],
    ...(failures.length ? { error: body } : {}),
  };
};

export const deckTools: RegisteredTool[] = [
  { definition: buildDeckDefinition, handler: buildDeckHandler },
];

export default deckTools;
