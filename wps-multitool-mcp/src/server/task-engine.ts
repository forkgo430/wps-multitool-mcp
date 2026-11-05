/**
 * Input: task-packages.json 里 execute:'server' 的包 + 运行参数
 * Output: 服务端一次跑完整包，中间结果不出服务端，只回最终摘要
 * Pos: 意图 v2 的服务端执行引擎（设计：docs/agents/INTENT-DESIGN-2026-11-04.md §4.2）。
 *      决策者 2026-11-05 亲自实现（所有者指定）。
 *
 * 为什么要有它
 * ------------
 * 任务包 v1 是"给模型的步骤清单"——每步由模型自己调用，每步的中间结果（含图表
 * base64 内联，上限 1.5MB）都完整流经模型上下文。Anthropic 实测这种模式是最大
 * token 浪费源（code execution with MCP：示例场景 -98.7%；PTC：-37%）。
 * 本引擎把包的执行搬进 server：逐步走 toolRegistry.callTool（与 deckkit/batch
 * 同一范式），中间结果留在引擎内存里供后续步骤用 {{steps.N.data.x}} 引用，
 * 模型只收到一步摘要 + 最终产物。
 *
 * 语义要点（都在头注里，改前先读）
 * ------------
 * 1. **不重放红线**：单步失败立即停止，已完成的步骤如实列出；不做部分续跑
 *    （v1 从头重跑——COM 不是事务，盲目续跑可能写两遍）。
 * 2. **破坏性闸门**：confirm_before 里的动作执行到它之前停下来，回 needConfirm
 *    让模型征询用户；带 confirm=true 才继续。闸门在引擎层，不依赖各工具自己的快照。
 * 3. **中间结果引用**：{{steps.N.data.字段}}（N 从 1 起，取该步回包 JSON 的 data）；
 *    {{steps.N.text}}（原始文本）；{{args.x}}（运行参数）。深解析 params 里所有字符串。
 * 4. **快照**：引擎自己不建快照——每步走 callTool 会触发各工具自己的
 *    withAutoSnapshot（破坏性/非幂等照常建）；wps_run_task 本体因无注解被
 *    fail-closed 判为破坏性，会在最外层多试建一次快照，失败降级为 warnings。
 *    这是"宁可多快照"纪律的体现，不是 bug。
 * 5. **与 intent.ts 的关系**：本文件自带包加载器（与 intent.ts 的 loader 并存）。
 *    不共用是有意的：intent.ts 归执行者改造（意图路由数据化），本文件归决策者，
 *    分文件避免同树互踩；合并时机由两轮都判定后再说。
 */

import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';
import { toolRegistry } from './tool-registry';
import { reportProgress } from './progress-context';
import { ToolCallResult } from '../types/tools';

// ---------------------------------------------------------------------------
// 数据形状（task-packages.json v2 schema 的 server 执行面）
// ---------------------------------------------------------------------------

export interface TaskStepV2 {
  /** 语义名（保留兼容；confirm_before 与 it 对齐） */
  action?: string;
  /** 注册工具名（v2 必填——消灭"语义名→工具名"断点） */
  tool: string;
  app: string;
  /** 字面量 + {{refs}}；缺省 {} */
  params?: Record<string, unknown>;
  note?: string;
  inline?: boolean;
  snapshot?: boolean;
}

export interface TaskPackageV2 {
  task: string;
  title: string;
  match: string[];
  not_for?: string[];
  /** 缺省 'model'——v1 行为（步骤清单交给模型逐步执行） */
  execute?: 'server' | 'model';
  steps: TaskStepV2[];
  confirm_before: string[];
}

export interface StepBrief {
  i: number;
  tool: string;
  ok: boolean;
  /** ≤120 字符的摘要；图片 base64 一律替换为占位符（这是本引擎存在的意义） */
  brief: string;
}

export interface TaskRunResult {
  success: boolean;
  task: string;
  title: string;
  /** 实际执行的步数 */
  executed: number;
  total: number;
  steps: StepBrief[];
  /** 最后一个成功步骤的 data（最常见的产物位置） */
  data?: unknown;
  needConfirm?: boolean;
  confirmAction?: string;
  failedStep?: number;
  error?: string;
  warnings?: Array<{ step?: string; message: string; severity?: 'info' | 'warning' }>;
}

// ---------------------------------------------------------------------------
// 包加载（与 intent.ts 的 loader 并存，理由见头注第 5 条）
// ---------------------------------------------------------------------------

let pkgCache: TaskPackageV2[] | null = null;

export function loadServerPackages(): TaskPackageV2[] {
  if (pkgCache) return pkgCache;
  const candidates = [
    path.join(__dirname, '..', 'server', 'task-packages.json'), // dist/server ← dist/tools/common 调用
    path.join(__dirname, '..', '..', 'src', 'server', 'task-packages.json'), // 仓库源码态
  ];
  for (const file of candidates) {
    try {
      pkgCache = JSON.parse(fs.readFileSync(file, 'utf8')).packages as TaskPackageV2[];
      break;
    } catch {
      /* 试下一个候选 */
    }
  }
  if (!pkgCache) pkgCache = [];
  return pkgCache;
}

/** 测试注入口：替换包缓存（测试用，生产不调） */
export function setPackagesForTest(pkgs: TaskPackageV2[]): void {
  pkgCache = pkgs;
}

// ---------------------------------------------------------------------------
// 参数模板解析：{{steps.N.data.x.y}} / {{steps.N.text}} / {{args.x}}
// ---------------------------------------------------------------------------

type StepRecord = { data?: unknown; text?: string };

export interface RunContext {
  args: Record<string, unknown>;
  steps: Map<number, StepRecord>;
}

const REF_RE = /\{\{\s*(args|steps)\.([\w.]+?)\s*\}\}/g;

function lookup(path: string, ctx: RunContext): unknown {
  const parts = path.split('.');
  const head = parts[0];
  let cur: unknown;
  if (head === 'args') {
    cur = ctx.args;
    parts.shift();
  } else if (head === 'steps') {
    const n = parseInt(parts[1], 10);
    if (Number.isNaN(n)) return undefined;
    cur = ctx.steps.get(n);
    parts.splice(0, 2);
    if (parts[0] === 'data' || parts[0] === 'text') {
      // steps.N.data.x / steps.N.text —— data/text 本身不是字段名
    } else {
      parts.unshift('data'); // {{steps.N.x}} 视为 {{steps.N.data.x}} 的简写
    }
  } else {
    return undefined;
  }
  for (const k of parts) {
    if (cur === null || cur === undefined) return undefined;
    cur = (cur as Record<string, unknown>)[k];
  }
  return cur;
}

/** 深遍历：对象/数组里所有字符串做模板替换；纯引用的字符串直接替换为解析值（保留类型） */
export function resolveValue<T>(v: T, ctx: RunContext): T {
  if (typeof v === 'string') {
    const whole = v.match(/^\{\{\s*([\w.]+?)\s*\}\}$/);
    if (whole) {
      const resolved = lookup(whole[1], ctx);
      return (resolved === undefined ? v : resolved) as unknown as T;
    }
    return v.replace(REF_RE, (raw, ns: string, rest: string) => {
      const resolved = lookup(`${ns}.${rest}`, ctx);
      return resolved === undefined ? raw : String(resolved);
    }) as unknown as T;
  }
  if (Array.isArray(v)) {
    return v.map((item) => resolveValue(item, ctx)) as unknown as T;
  }
  if (v && typeof v === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
      out[k] = resolveValue(val, ctx);
    }
    return out as unknown as T;
  }
  return v;
}

// ---------------------------------------------------------------------------
// 回包压缩：base64 图片一律替换占位符；文本截 120 字符
// ---------------------------------------------------------------------------

const IMAGE_B64_RE = /data:image\/[a-z]+;base64,[A-Za-z0-9+/=]+/g;

export function briefOf(result: ToolCallResult): string {
  const raw = (result.content?.[0] as { type: string; text?: string } | undefined)?.text;
  if (typeof raw !== 'string' || !raw) return result.error || '(无文本回包)';
  const stripped = raw.replace(IMAGE_B64_RE, '[inline-image]').replace(/\s+/g, ' ').trim();
  return stripped.length > 120 ? stripped.slice(0, 117) + '…' : stripped;
}

/** 从回包提取可供后续步骤引用的数据：优先 JSON 解析 content[0].text，失败只给 text */
export function extractStepData(result: ToolCallResult): StepRecord {
  const text = (result.content?.[0] as { type: string; text?: string } | undefined)?.text;
  const rec: StepRecord = { text: typeof text === 'string' ? text : undefined };
  if (typeof text === 'string' && text.trim().startsWith('{')) {
    try {
      const parsed = JSON.parse(text) as Record<string, unknown>;
      if (parsed && typeof parsed === 'object' && 'data' in parsed) rec.data = parsed.data;
      else rec.data = parsed;
    } catch {
      /* 非 JSON 文本——只留 text */
    }
  }
  return rec;
}

// ---------------------------------------------------------------------------
// 执行引擎
// ---------------------------------------------------------------------------

export interface TaskEngineDeps {
  /** 可注入的执行器（测试用）；生产默认走真注册表 */
  callTool?: (name: string, args: Record<string, unknown>) => Promise<ToolCallResult>;
  /** 可注入的包查找（测试用）；生产默认按 task 精确匹配 */
  findPackage?: (task: string) => TaskPackageV2 | undefined;
}

function defaultCallTool(name: string, args: Record<string, unknown>): Promise<ToolCallResult> {
  return toolRegistry.callTool({ id: uuidv4(), name, arguments: args });
}

function defaultFindPackage(task: string): TaskPackageV2 | undefined {
  return loadServerPackages().find((p) => p.task === task && p.execute === 'server');
}

/**
 * 执行一个 server 任务包。
 * @param args 运行参数（供 {{args.x}} 引用；confirm 不在此列——见独立形参）
 */
export async function runTaskPackage(
  task: string,
  args: Record<string, unknown> = {},
  options: { confirm?: boolean } = {},
  deps: TaskEngineDeps = {}
): Promise<TaskRunResult> {
  const callTool = deps.callTool ?? defaultCallTool;
  const findPackage = deps.findPackage ?? defaultFindPackage;

  const pkg = findPackage(task);
  if (!pkg) {
    return {
      success: false,
      task,
      title: task,
      executed: 0,
      total: 0,
      steps: [],
      error: `没有找到可服务端执行的任务包「${task}」。execute:'server' 的包才能用本入口；其余任务仍走 route_intent 的清单模式。`,
    };
  }

  const ctx: RunContext = { args, steps: new Map() };
  const briefs: StepBrief[] = [];
  const warnings: TaskRunResult['warnings'] = [];
  const total = pkg.steps.length;
  let executed = 0;
  let lastData: unknown;

  for (let i = 0; i < total; i++) {
    const step = pkg.steps[i];

    // 破坏性闸门：confirm_before 命中且未带 confirm → 停在它前面征询
    if (step.action && pkg.confirm_before.includes(step.action) && !options.confirm) {
      return {
        success: false,
        task,
        title: pkg.title,
        executed,
        total,
        steps: briefs,
        data: lastData,
        needConfirm: true,
        confirmAction: step.action,
        warnings: [
          {
            step: '任务引擎',
            message: `第 ${i + 1} 步「${step.note || step.action}」在 confirm_before 里——确认后带 confirm:true 重跑（已完成的 ${executed} 步会从头重执行，包语义要求幂等）`,
            severity: 'info',
          },
        ],
        error: 'need-confirm',
      };
    }

    try {
      reportProgress(i + 1, total, step.note || step.tool);
    } catch {
      /* 进度通知失败不影响执行 */
    }

    const params = resolveValue(step.params ?? {}, ctx);
    let result: ToolCallResult;
    try {
      result = await callTool(step.tool, params as Record<string, unknown>);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return {
        success: false,
        task,
        title: pkg.title,
        executed,
        total,
        steps: briefs,
        data: lastData,
        failedStep: i + 1,
        error: `第 ${i + 1} 步 ${step.tool} 抛异常：${msg}`,
      };
    }

    briefs.push({ i: i + 1, tool: step.tool, ok: result.success, brief: briefOf(result) });
    for (const w of result.warnings ?? []) warnings.push(w);

    if (!result.success) {
      return {
        success: false,
        task,
        title: pkg.title,
        executed,
        total,
        steps: briefs,
        data: lastData,
        failedStep: i + 1,
        error: `第 ${i + 1} 步 ${step.tool} 失败：${result.error || briefOf(result)}`,
        warnings: warnings.length ? warnings : undefined,
      };
    }

    executed += 1;
    const rec = extractStepData(result);
    ctx.steps.set(i + 1, rec);
    lastData = rec.data;
  }

  return {
    success: true,
    task,
    title: pkg.title,
    executed,
    total,
    steps: briefs,
    data: lastData,
    warnings: warnings.length ? warnings : undefined,
  };
}
