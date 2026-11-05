/**
 * Input: route_intent 的未命中记录（recordMiss）与后续实际工具调用（noteToolUse）
 * Output: 会话内缓冲 + 进程退出时一次性落盘 scripts/manual/intent-miss-log.jsonl；top 未命中聚合查询
 * Pos: 意图自适应回流的数据面（体验修复轮 A · 方案 §4.1 自适配零人力版）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么存在：关键词意图规则必然有漏——漏了什么只有真实会话知道。本模块把
 * “route_intent 未命中/低置信的话术”与“该会话随后实际调用的工具名”配对记下来，
 * 落盘后由人（或脚本）聚合出 top 未命中，在 route_intent 回包里提示“可补充词条”。
 * 只记录与提示，不改任何路由表——补词条永远是人的决定。
 *
 * 落盘守卫：目标是仓库的 scripts/manual/intent-miss-log.jsonl。目录不存在（npm 安装态、
 * 任意 cwd）就静默放弃——装出去的包不该往用户磁盘的包目录里写运行日志。
 */

import * as fs from 'fs';
import * as path from 'path';

interface MissPair {
  task: string;
  tool: string;
  ts: string;
}

const buffer: MissPair[] = [];
/** 最近一次未命中的话术；下一条实际工具调用到达时配对并清空 */
let pendingTask: string | null = null;
let flushed = false;

const LOG_RELATIVE = path.join('scripts', 'manual', 'intent-miss-log.jsonl');

function logPath(): string | null {
  try {
    const dir = path.resolve(process.cwd(), LOG_RELATIVE, '..');
    if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) return null;
  } catch {
    return null;
  }
  return path.resolve(process.cwd(), LOG_RELATIVE);
}

/** route_intent 未命中/低置信时调：记下这句待配对话术 */
export function recordMiss(task: string): void {
  if (!task.trim()) return;
  pendingTask = task.trim();
}

/** 每次实际工具调用后调：若存在待配对话术，与本次工具名配对进缓冲（一次配对即清） */
export function noteToolUse(toolName: string): void {
  if (!pendingTask) return;
  const task = pendingTask;
  pendingTask = null;
  if (!toolName) return;
  buffer.push({ task, tool: toolName, ts: new Date().toISOString() });
}

/** 进程退出时一次性落盘（append；只在仓库工作树里落，安装态放弃）。幂等：多次调只写一次 */
export function flushOnExit(): void {
  if (flushed || !buffer.length) return;
  const p = logPath();
  if (!p) return;
  flushed = true;
  try {
    const lines = buffer.map((m) => JSON.stringify(m)).join('\n') + '\n';
    fs.appendFileSync(p, lines, 'utf8');
  } catch {
    /* 落盘失败静默——回流是增益不是依赖 */
  }
}

/** 启动后首次 route_intent 调用读取历史日志，聚合 top N 未命中话术（按 task 计数） */
export function topMisses(n = 3): Array<{ task: string; count: number }> {
  const p = logPath();
  if (!p) return [];
  try {
    if (!fs.existsSync(p)) return [];
    const counts = new Map<string, number>();
    const raw = fs.readFileSync(p, 'utf8');
    for (const line of raw.split('\n')) {
      const s = line.trim();
      if (!s) continue;
      try {
        const rec = JSON.parse(s) as { task?: string };
        if (rec.task) counts.set(rec.task, (counts.get(rec.task) || 0) + 1);
      } catch {
        /* 跳过脏行 */
      }
    }
    return [...counts.entries()]
      .map(([task, count]) => ({ task, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, n);
  } catch {
    return [];
  }
}

/** 进程退出钩子（幂等注册） */
export function registerExitFlush(): void {
  if (flushed) return;
  process.on('exit', flushOnExit);
}
