/**
 * Input: 调用方给的结构化地址对象（不可信）
 * Output: 解析结果——ok=true 带归一化地址；ok=false 带给模型看的完整错误消息
 * Pos: 统一文档定位层 L1 地址 DSL（TS 纯函数）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 分工（设计方案变更三）：人话 → 地址的翻译归模型（"第三章第三段第三句"→{chapter:3,paragraph:3,sentence:3}），
 * DSL 只收结构化对象并做语法/组合校验——语法错、组合错在进 COM 前就死（解析/执行分离，编译器前端手法）。
 * 地址 → Range 的可靠落地归 L2（ps1 Resolve-TargetRange）。
 *
 * 地址形状：
 *   {
 *     chapter?:  number(≥1) | { text: string }   —— 章（OutlineLevel=1；index 或标题文本含糊匹配）
 *     paragraph?: number(≥1)                      —— 段（章内正文段计数，表格段默认不计序）
 *     sentence?: number(≥1)                       —— 句（必须带 paragraph 锚——组合校验）
 *     edge?:     'start' | 'end'                  —— 收尾对齐（缺省 'start'）
 *     text?:     string                           —— 文本寻址（与 start/end 互斥）
 *     occurrence?: number(≥1)                     —— text 的第几次出现（缺省 1）
 *     start?:    number(≥1)                       —— 字符偏移区间起点（与 text 互斥）
 *     end?:      number(≥1)                       —— 字符偏移区间终点（与 start 成对）
 *   }
 * 校验面：未知键拒绝（拼写自保护）/ 数字必须正整数 / sentence 须带 paragraph 锚 /
 *         text 与 start|end 互斥 / start+end 须成对 / 空地址拒绝。
 */

/** 归一化后的地址（L2 消费的形状） */
export interface ParsedAddress {
  chapterIndex?: number;
  chapterText?: string;
  paragraph?: number;
  sentence?: number;
  edge: 'start' | 'end';
  text?: string;
  occurrence: number;
  start?: number;
  end?: number;
}

export type ParseAddressResult =
  | { ok: true; address: ParsedAddress }
  | { ok: false; error: string };

const KNOWN_KEYS = ['chapter', 'paragraph', 'sentence', 'edge', 'text', 'occurrence', 'start', 'end', 'locate'] as const;

/** 正整数（≥1 的有限整数） */
function isPositiveInt(v: unknown): v is number {
  return typeof v === 'number' && Number.isFinite(v) && Number.isInteger(v) && v >= 1;
}

export function parseAddress(input: unknown): ParseAddressResult {
  if (input === null || input === undefined || typeof input !== 'object' || Array.isArray(input)) {
    return { ok: false, error: `地址必须是对象（如 {chapter:3, paragraph:3, sentence:3, edge:'end'}），收到 ${describeInput(input)}` };
  }
  const raw = input as Record<string, unknown>;
  const unknownKeys = Object.keys(raw).filter((k) => !(KNOWN_KEYS as readonly string[]).includes(k));
  if (unknownKeys.length > 0) {
    return { ok: false, error: `地址含未知键: ${unknownKeys.join(', ')}（可用键：${KNOWN_KEYS.join('/')}）` };
  }

  const out: ParsedAddress = { edge: 'start', occurrence: 1 };

  // chapter：数字或 {text}
  if (raw.chapter !== undefined) {
    const c = raw.chapter;
    if (typeof c === 'number') {
      if (!isPositiveInt(c)) return { ok: false, error: `chapter 必须是 ≥1 的整数，收到 ${describeInput(c)}` };
      out.chapterIndex = c;
    } else if (c !== null && typeof c === 'object' && !Array.isArray(c)) {
      const t = (c as Record<string, unknown>).text;
      if (typeof t !== 'string' || !t.trim()) {
        return { ok: false, error: "chapter 的文本形式必须是 { text: '非空标题文本' }" };
      }
      out.chapterText = t.trim();
    } else {
      return { ok: false, error: `chapter 必须是数字（第几章）或 { text: '标题文本' }，收到 ${describeInput(c)}` };
    }
  }

  // paragraph / sentence / occurrence：正整数
  if (raw.paragraph !== undefined) {
    if (!isPositiveInt(raw.paragraph)) return { ok: false, error: `paragraph 必须是 ≥1 的整数（第几个正文段），收到 ${describeInput(raw.paragraph)}` };
    out.paragraph = raw.paragraph as number;
  }
  if (raw.sentence !== undefined) {
    if (!isPositiveInt(raw.sentence)) return { ok: false, error: `sentence 必须是 ≥1 的整数（段内第几句），收到 ${describeInput(raw.sentence)}` };
    out.sentence = raw.sentence as number;
  }
  if (raw.occurrence !== undefined) {
    if (!isPositiveInt(raw.occurrence)) return { ok: false, error: `occurrence 必须是 ≥1 的整数（text 的第几次出现），收到 ${describeInput(raw.occurrence)}` };
    out.occurrence = raw.occurrence as number;
  }

  // 组合校验：sentence 须带 paragraph 锚（句在段内，段是句的坐标系）
  if (out.sentence !== undefined && out.paragraph === undefined) {
    return { ok: false, error: 'sentence 必须与 paragraph 同给（句的坐标系是段；只有章锚无法确定段内句序）' };
  }

  // edge
  if (raw.edge !== undefined) {
    if (raw.edge !== 'start' && raw.edge !== 'end') {
      return { ok: false, error: `edge 只能是 'start' 或 'end'，收到 ${describeInput(raw.edge)}` };
    }
    out.edge = raw.edge;
  }

  // text 与 start/end 互斥；start+end 成对
  const hasText = raw.text !== undefined;
  const hasStart = raw.start !== undefined;
  const hasEnd = raw.end !== undefined;
  if (hasText && (hasStart || hasEnd)) {
    return { ok: false, error: "text 与 start/end 互斥（文本寻址和字符偏移寻址二选一）" };
  }
  if (hasStart !== hasEnd) {
    return { ok: false, error: 'start 与 end 必须成对给（字符偏移区间）' };
  }
  if (hasStart) {
    if (!isPositiveInt(raw.start)) return { ok: false, error: `start 必须是 ≥1 的整数（字符偏移），收到 ${describeInput(raw.start)}` };
    if (!isPositiveInt(raw.end)) return { ok: false, error: `end 必须是 ≥1 的整数（字符偏移），收到 ${describeInput(raw.end)}` };
    if ((raw.start as number) >= (raw.end as number)) {
      return { ok: false, error: `start（${raw.start}）必须小于 end（${raw.end}）` };
    }
    out.start = raw.start as number;
    out.end = raw.end as number;
  }
  if (hasText) {
    if (typeof raw.text !== 'string' || !raw.text.trim()) {
      return { ok: false, error: 'text 必须是非空字符串（要定位的文本）' };
    }
    out.text = raw.text;
  }

  // 空地址拒绝（至少一个锚）
  if (out.chapterIndex === undefined && out.chapterText === undefined && out.paragraph === undefined &&
      out.text === undefined && out.start === undefined) {
    return { ok: false, error: '空地址：至少要给 chapter/paragraph/text/start 中的一个锚' };
  }

  // occurrence 只跟 text 搭配
  if (out.occurrence !== 1 && !hasText) {
    return { ok: false, error: 'occurrence 只在 text 寻址时有意义（第几次出现的文本）' };
  }

  return { ok: true, address: out };
}

function describeInput(v: unknown): string {
  if (v === null) return 'null';
  if (v === undefined) return 'undefined';
  if (Array.isArray(v)) return `array(长度 ${(v as unknown[]).length})`;
  const t = typeof v;
  if (t === 'string') {
    const s = v as string;
    return `string("${s.length > 20 ? s.slice(0, 20) + '…' : s}")`;
  }
  if (t === 'number' || t === 'boolean') return `${t}(${String(v)})`;
  if (t === 'object') return `object(键: ${Object.keys(v as object).slice(0, 5).join(', ') || '无'})`;
  return t;
}
