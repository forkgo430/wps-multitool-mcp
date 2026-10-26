/**
 * Input: 应用键（excel / word / ppt / common）与动作名、参数
 * Output: 该动作的 COM 结果（{ success, data?, error? }）
 * Pos: 桥进程内部的 PowerShell 子进程管理器 —— 一个应用一个实例。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么按应用拆进程（per-app 隔离）
 * ----------------------------------
 * COM 是 STA：一个进程同一时刻只能跑一个动作。改造前所有应用共用一个 powershell，
 * 于是"WPS 文字卡在一个慢动作上"会把"表格"的调用一起堵死。
 * 拆成一应用一进程后，某一路卡死只影响它自己；桥还能【只杀那一路】并重建，
 * 另外两路完全不受影响。对标 JulianPoleszczuk/office-mcp 的 per-app 连接模型。
 *
 * 超时语义
 * --------
 * 超时即判定【该应用的这一路已死】：杀掉该子进程、拒绝它在途的请求，下次调用重建。
 * 绝不能"超时了还继续往这个进程塞请求"。
 * 另外：超时后【不自动重放】—— 写动作（保存/导出）可能已经执行了一半，
 * 重放会变成执行两次。由调用方决定怎么办。
 *
 * 子进程生命周期
 * --------------
 * wps-comd.ps1 自带父进程看门狗，本进程（桥）一死，子进程会自己走 —— 不需要额外清理。
 *
 * PowerShell 宿主
 * ---------------
 * 优先 PowerShell 7（pwsh），未安装自动回退 Windows PowerShell 5.1；WPS_PWSH_PATH
 * 可强制指定（完整路径或裸命令名，写 'powershell' 即强制回 5.1）。
 * 为什么要切 pwsh：5.1 读无 BOM 的 UTF-8 脚本会把中文解码成乱码（wps-comd.ps1 头部
 * 的 BOM 说明就是为它写的），PS7 默认 UTF-8，这一类坑整个消失。
 * COM 是 STA：5.1 默认就是 STA，pwsh 在 Windows 上默认也是，但这里显式传 -STA，不赌默认值。
 * 回退会往 stderr 记一行宿主是谁，不是无声的。
 */

import { ChildProcess, spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { CancelError, isReadOnlyAction } from './cancellable';

/** 常驻分发脚本（随包分发，见 .npmignore） */
const PS_SCRIPT = path.join(__dirname, '../../scripts/wps-comd.ps1');

/**
 * 选 PowerShell 宿主：pwsh7 优先，回退 5.1。模块加载时算一次，进程生命周期内不变 ——
 * 同一个桥里两路应用跑在不同宿主上没有任何好处。
 */
function resolvePowerShellHost(): { exe: string; sta: boolean; label: string } {
  const forced = process.env.WPS_PWSH_PATH;
  if (forced) {
    return { exe: forced, sta: /pwsh(\.exe)?$/i.test(forced), label: `WPS_PWSH_PATH=${forced}` };
  }
  const pf = process.env.ProgramFiles || 'C:\\Program Files';
  const candidates = [path.join(pf, 'PowerShell', '7', 'pwsh.exe')];
  for (const dir of (process.env.PATH || '').split(path.delimiter)) {
    if (dir) candidates.push(path.join(dir, 'pwsh.exe'));
  }
  for (const c of candidates) {
    try {
      if (fs.existsSync(c)) return { exe: c, sta: true, label: c };
    } catch {
      /* PATH 里可能有非法条目，跳过 */
    }
  }
  return { exe: 'powershell', sta: false, label: 'powershell.exe（未找到 PowerShell 7，回退 5.1）' };
}

const PS_HOST = resolvePowerShellHost();
console.error(`[ps-child] PowerShell 宿主: ${PS_HOST.label}`);

/** 起进程 + 握手就绪的上限。冷启动实测约 5 秒，给足余量 */
const READY_TIMEOUT_MS = 30000;

/**
 * 单次动作上限。COM 动作可能很慢（大文档导出、首次插图表要拉起 Excel），默认 120 秒。
 * 每次调用时重读环境变量，便于按场景临时收紧/放宽，不用重启桥。
 *
 * 第 1 期第 2 章第 4 轮：按【动作性质】分级（见 docs/agents/TIMEOUT-120S.md §4.1）。
 * 原来是一刀切：读一个单元格和导出一张图，上限都是 120 秒。
 * 后果是"快动作卡死"要等满 120 秒才恢复。
 * 现在：慢动作仍 120 秒，其余 30 秒 —— 卡死恢复时间从 120 秒降到 30 秒。
 *
 * 为什么不给"快动作"也留 120 秒的余量：宁可早失败也不要长时间占着 COM（STA，在途恒为 1，
 * 一个动作卡住 = 整个应用这条道全堵）。而且超时不重放（写动作可能已改了一半）。
 */
function callTimeoutMs(): number {
  const v = parseInt(process.env.WPS_DAEMON_TIMEOUT || '120000', 10);
  return Number.isFinite(v) && v > 0 ? v : 120000;
}

/**
 * 慢动作清单 —— 按【动作性质】分，不按实测数字分。
 *
 * 判据（见 TIMEOUT-120S.md §4.1）：这类动作要么"要渲染"、要么"要整篇转格式"、
 * 要么"首次建对象会拉起别的组件"、要么"全文级扫描/重算"。
 * 清单是可调的常量：发现漏了直接加，或用 WPS_SLOW_ACTIONS 环境变量覆盖。
 *
 * 名字对着 scripts/wps-com.ps1 的真实 case 核过 —— 文档草案里的
 * "exportPptImage" / "create3DPieChart" / "convertToPdf" 三个在 ps1 里【并不存在】
 * （真实的是 convertToPDF），照抄会写成一张永远匹配不上的表。
 */
const DEFAULT_SLOW_ACTIONS = [
  // ① 要渲染的（导出成图）
  'exportChartAsImage',
  'exportRangeAsImage',
  'exportSlideAsImage',
  // ② 要整篇转格式的
  'convertToPDF',
  'convertFormat',
  // ③ 首次建对象 / 重算的（可能拉起 Excel，或触发透视表重算）
  'createChart',
  'createPivotTable',
  'updatePivotTable',
  'refreshPivotTables',
  'insertPptChart',
  'createDonutChart',
  'createMiniCharts',
  'createKpiCards',
  // ④ 全文级扫描 / 更新的
  'updateFields',
  'generateTOC',
  'findReplace',
  'insertSlidesFromFile',
  'convertTableToText',
];

/** 慢动作清单：环境变量可覆盖（逗号分隔），没设就用默认那张表 */
function slowActions(): Set<string> {
  const raw = process.env.WPS_SLOW_ACTIONS;
  if (raw !== undefined) {
    return new Set(raw.split(',').map((s) => s.trim()).filter(Boolean));
  }
  return new Set(DEFAULT_SLOW_ACTIONS);
}

/** 快动作的单独上限：卡死时早恢复。设 WPS_DAEMON_FAST_TIMEOUT 覆盖。 */
function fastTimeoutMs(): number {
  const v = parseInt(process.env.WPS_DAEMON_FAST_TIMEOUT || '30000', 10);
  return Number.isFinite(v) && v > 0 ? v : 30000;
}

/**
 * 按动作性质取超时上限。
 * 慢动作 → WPS_DAEMON_TIMEOUT（默认 120000）；其余 → WPS_DAEMON_FAST_TIMEOUT（默认 30000）。
 * 动作名不认识（空/非字符串）时按慢的算 —— 宁可多等，不可误杀。
 */
export function callTimeoutMsFor(action?: string): number {
  const base = callTimeoutMs();
  if (!action) return base;
  return slowActions().has(action) ? base : fastTimeoutMs();
}

/**
 * 每应用的排队上限。超了立刻回 busy —— 把"无声地等几分钟"换成"明确被拒"。
 * 注意：这不等于并发度（COM 是 STA，在途恒为 1），它限的是"等待位"的数量。
 */
export function queueMax(): number {
  const v = parseInt(process.env.WPS_QUEUE_MAX || '4', 10);
  return Number.isFinite(v) && v > 0 ? v : 4;
}

/** 排队上限。排太久不如早拒 —— 与动作超时是两个独立维度。 */
export function queueTimeoutMs(): number {
  const v = parseInt(process.env.WPS_QUEUE_TIMEOUT || '10000', 10);
  return Number.isFinite(v) && v > 0 ? v : 10000;
}

/**
 * 需要「插队」的动作：会开/关/另存文档、切目标、退出 WPS，或弹出对话框。
 * 这些是用户主动发起的，等待位里已经排了一堆读操作时，让它们先走。
 * 饥饿风险可控 —— 这类动作很稀疏，且跳过后队列恢复 FIFO。
 * 可用 WPS_EXCLUSIVE_ACTIONS="a,b,c" 覆盖。
 */
const DEFAULT_EXCLUSIVE_ACTIONS = [
  'openFile',
  'openDocument',
  'openWorkbook',
  'openPresentation',
  'closeDocument',
  'closeWorkbook',
  'closePresentation',
  'switchDocument',
  'switchWorkbook',
  'switchPresentation',
  'createDocument',
  'createWorkbook',
  'createPresentation',
  'save',
  'saveAs',
  'convertToPDF',
  'convertFormat',
  'quitWps',
];

export function exclusiveSet(): Set<string> {
  const override = process.env.WPS_EXCLUSIVE_ACTIONS;
  const list = override
    ? override.split(',').map((s) => s.trim()).filter(Boolean)
    : DEFAULT_EXCLUSIVE_ACTIONS;
  return new Set(list);
}

/** 子进程压根没起来（请求从未送达）时抛这个，调用方可以安全重建。 */
export class ChildUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ChildUnavailableError';
  }
}

/** 队列满 / 排队超时：请求【从未送达】PowerShell，属于可安全重试的基础设施问题。 */
export class QueueBusyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'QueueBusyError';
  }
}

interface PendingCall {
  resolve: (value: Record<string, unknown>) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
  /** 调用方已放弃等待（只读动作专属）：回话到了直接丢，不再交给调用方 */
  abandoned?: boolean;
  /**
   * 通知调用方"不等了"的专用通道。
   * 不能直接用 reject —— reject 是 dispatch 自己的 promise 的 reject，
   * 一 reject 就会让 drain 收尾、把 inFlight 置 false，
   * 于是下一个请求会被写进一个【还在跑上一个动作】的 STA 进程。
   */
  onAbandon?: (err: Error) => void;
}

/** 队列里等位的一个请求 */
interface QueuedTask {
  action: string;
  params: Record<string, unknown>;
  exclusive: boolean;
  enqueuedAt: number;
  resolve: (value: Record<string, unknown>) => void;
  reject: (err: Error) => void;
  /** 排队超时定时器 */
  timer?: NodeJS.Timeout;
  /** 调用方（桥）给的关联 id，用于跨进程取消 */
  externalId?: string;
  /** 调用方的取消信号 */
  signal?: AbortSignal;
}

/** call() 的可选参数 */
export interface CallOptions {
  /** 取消信号：来自 MCP 客户端的 notifications/cancelled */
  signal?: AbortSignal;
  /** 桥侧的请求 id，用于后续按 id 取消 */
  externalId?: string;
}

export interface ChildStats {
  app: string;
  /** 这个子进程对应的实例：user（用户的 WPS）或 sandbox-N（隔离实例） */
  instance: string;
  /** 当前活着的请求数（在途 + 在等）。池子按它挑"第一个空闲的" */
  load: number;
  ready: boolean;
  pid: number | null;
  uptimeMs: number;
  restarts: number;
  /** 在途（已送出、等 WPS 回话） */
  pending: number;
  /** 队列里等位的数量 */
  queued: number;
  /** 是否已有一个动作在途（STA：恒为 0 或 1） */
  inFlight: boolean;
  /** 队列上限 */
  queueMax: number;
  /** 排队时长上限 */
  queueTimeoutMs: number;
  /** 最近一次排队等待的毫秒数（无则 0） */
  lastQueueWaitMs: number;
  /** 累计因队列满/排队超时被拒的次数 */
  rejected: number;
  /** 累计被取消的次数（排队中出队 + 只读动作放弃等待） */
  cancelled: number;
}

export class PsChildDaemon {
  private proc: ChildProcess | null = null;
  private stdoutBuf = '';
  private pending = new Map<string, PendingCall>();
  private queue: QueuedTask[] = [];
  private starting: Promise<void> | null = null;
  private ready = false;
  private disposed = false;
  private startedAt: number | null = null;
  private restarts = 0;
  private lastQueueWaitMs = 0;
  private rejected = 0;
  /**
   * 是否已有一个动作被 drain 送出、还没回话。
   *
   * 不能用 pending.size 判断：pending 是【送出之后】才登记的，
   * 而 ensureStarted() 是异步的 —— 冷启动那几秒里 pending 一直是 0，
   * 结果一串并发请求会全部"穿过"队列被同时写进子进程（既破了并发上限，
   * 也让 STA 约束失效）。所以这里用一个同步置位的标志位。
   */
  private inFlight = false;
  /** 在途请求的桥侧 id / 动作名 / 子进程侧 id —— 取消时要靠它们定位 */
  private inFlightId: string | null = null;
  private inFlightAction: string | null = null;
  private inFlightPendingId: string | null = null;
  /** 累计被取消的次数 */
  private cancelled = 0;

  constructor(private readonly appKey: string, private readonly instance: string = 'user') {}

  get isReady(): boolean {
    return this.proc !== null && this.ready;
  }

  /**
   * 这个实例当前有多少活在身上（在途 + 排队）。
   * 池子用它挑"第一个空闲的"—— 不能只看 inFlight：
   * 一个实例可能正在跑一个长动作，但队列里没人，那它也不该再被塞东西。
   */
  get load(): number {
    return (this.inFlight ? 1 : 0) + this.pending.size + this.queue.length;
  }

  get stats(): ChildStats {
    return {
      app: this.appKey,
      instance: this.instance,
      load: this.load,
      ready: this.ready,
      pid: this.proc?.pid ?? null,
      uptimeMs: this.startedAt === null ? 0 : Date.now() - this.startedAt,
      restarts: this.restarts,
      pending: this.pending.size,
      queued: this.queue.length,
      inFlight: this.inFlight,
      queueMax: queueMax(),
      queueTimeoutMs: queueTimeoutMs(),
      lastQueueWaitMs: this.lastQueueWaitMs,
      rejected: this.rejected,
      cancelled: this.cancelled,
    };
  }

  /**
   * 调用一个动作。业务成功/失败都以响应对象 resolve；
   * 只有【基础设施问题】（起不来、超时、进程退出、队列满）才 reject。
   *
   * 并发语义（重要）
   * ----------------
   * COM 是 STA：一个应用同一时刻【只能有一个动作在跑】。所以这里不是"提高并发度"，
   * 而是把"无声排队"变成"有界的、看得见的等待"：
   *   - 队列上限 queueMax（默认 4），超了立刻回 QueueBusyError，不让调用方干等；
   *   - 排队超时 queueTimeoutMs（默认 10s），与动作超时独立；
   *   - 排他的动作（open/save/quit 等）插队，避免用户主动操作被一堆读操作压住。
   *
   * 取消语义（详见 bridge/cancellable.ts）
   * --------------------------------------
   *   排队中被取消 → 直接出队，请求从未送达，任何动作都安全
   *   已在途中被取消 → 只有【只读动作】才放弃等待（回话到了丢弃，子进程继续跑完）；
   *                    写入动作一律不放弃，老实等它跑完如实返回
   */
  async call(
    action: string,
    params: Record<string, unknown> = {},
    opts: CallOptions = {}
  ): Promise<Record<string, unknown>> {
    if (this.disposed) {
      throw new ChildUnavailableError(`[${this.appKey}] 子进程已关闭`);
    }

    const max = queueMax();
    if (this.queue.length >= max) {
      this.rejected += 1;
      throw new QueueBusyError(
        `WPS ${this.appKey} 队列已满（${max} 个在等，动作=${action}）。` +
          '该应用同一时刻只能跑一个动作（COM 是 STA），请稍后重试或先降低并发。'
      );
    }

    const exclusive = exclusiveSet().has(action);

    const signal = opts.signal;
    if (signal?.aborted) {
      throw new CancelError(`[${this.appKey}] 请求在送出前已被取消（动作=${action}），从未送达`);
    }

    return new Promise<Record<string, unknown>>((resolve, reject) => {
      const task: QueuedTask = {
        action,
        params,
        exclusive,
        enqueuedAt: Date.now(),
        resolve,
        reject,
        externalId: opts.externalId,
        signal,
      };
      const qTimeout = queueTimeoutMs();
      task.timer = setTimeout(() => {
        const i = this.queue.indexOf(task);
        if (i >= 0) this.queue.splice(i, 1);
        this.rejected += 1;
        reject(
          new QueueBusyError(
            `WPS ${this.appKey} 排队超时（${qTimeout}ms，动作=${action}）。` +
              '前面还有动作在跑；这个请求从未送出，可以安全重试。'
          )
        );
      }, qTimeout);

      if (exclusive) this.queue.unshift(task);
      else this.queue.push(task);

      // 取消信号：只管"还在排队"这一档（已送出的由 dispatch 里的监听器按只读/写入分开处理）
      if (signal) {
        signal.addEventListener(
          'abort',
          () => {
            const i = this.queue.indexOf(task);
            if (i < 0) return; // 已经送出了，不归这里管
            this.queue.splice(i, 1);
            if (task.timer) clearTimeout(task.timer);
            this.cancelled += 1;
            reject(
              new CancelError(
                `WPS ${this.appKey} 请求在排队中被取消（动作=${action}）。该请求从未送达，可以安全重发。`
              )
            );
          },
          { once: true }
        );
      }

      void this.drain();
    });
  }

  /**
   * 按桥侧 id 取消一个请求（跨进程取消入口）。
   * @returns 是否真的取消了。false = 找不到 / 已送出且是写入动作（不许取消）
   */
  cancelByExternalId(externalId: string): { ok: boolean; reason: string } {
    const i = this.queue.findIndex((t) => t.externalId === externalId);
    if (i >= 0) {
      const task = this.queue[i];
      this.queue.splice(i, 1);
      if (task.timer) clearTimeout(task.timer);
      this.cancelled += 1;
      task.reject(
        new CancelError(
          `WPS ${this.appKey} 请求在排队中被取消（动作=${task.action}）。该请求从未送达，可以安全重发。`
        )
      );
      return { ok: true, reason: '排队中，已出队（从未送达）' };
    }
    if (this.inFlightId === externalId) {
      if (!isReadOnlyAction(this.inFlightAction ?? '')) {
        return {
          ok: false,
          reason: `动作 ${this.inFlightAction} 已送出且是【写入】动作，中途放弃会让调用方误判成败，不予取消`,
        };
      }
      const p = this.pending.get(this.inFlightPendingId ?? '');
      if (!p) return { ok: false, reason: '已在途但找不到在途记录' };
      this.abandonPending(this.inFlightPendingId ?? '');
      this.cancelled += 1;
      return { ok: true, reason: `动作 ${this.inFlightAction} 已在途（只读）：放弃等待，子进程会跑完，回话丢弃` };
    }
    return { ok: false, reason: '没有找到该请求（已完成或从未进入本应用）' };
  }

  /**
   * 放弃一个在途请求：只通知调用方"不等了"，【不动 dispatch 的 promise】。
   * dispatch 必须继续等到子进程真的回话，否则 inFlight 会提前释放 ——
   * 那等于允许下一个请求闯进一个还在忙的 STA 进程。
   */
  private abandonPending(pendingId: string): void {
    const p = this.pending.get(pendingId);
    if (!p || p.abandoned) return;
    p.abandoned = true;
    p.onAbandon?.(
      new CancelError(
        `WPS ${this.appKey} 请求已被放弃等待（动作=${this.inFlightAction}）。` +
          '动作仍在子进程里跑完（只读，无副作用），其回话会被丢弃。'
      )
    );
  }

  /**
   * 队列泵：把队首送进子进程，回话后继续送下一个。
   * 在途只允许 1 个 —— 这是 STA 的硬约束，不是策略选择。
   */
  private async drain(): Promise<void> {
    if (this.disposed || this.inFlight) return;
    const task = this.queue.shift();
    if (!task) return;
    if (task.timer) {
      clearTimeout(task.timer);
      task.timer = undefined;
    }
    this.inFlight = true;
    this.inFlightId = task.externalId ?? null;
    this.inFlightAction = task.action;
    this.inFlightPendingId = null;
    this.lastQueueWaitMs = Date.now() - task.enqueuedAt;

    try {
      // 注意：dispatch 的 promise 要一直等到子进程【真的】回话才 settle ——
      // 哪怕调用方中途放弃了等待，inFlight 也不能提前变 false，
      // 否则下一个请求会被写进一个还在忙的 STA 进程里。
      const msg = await this.dispatch(task.action, task.params, {
        signal: task.signal,
        onPending: (pendingId) => {
          this.inFlightPendingId = pendingId;
        },
        // 只读动作被放弃等待时走这条；写入动作永远不会触发
        onAbandon: (err) => task.reject(err),
      });
      task.resolve(msg);
    } catch (err) {
      task.reject(err as Error);
    } finally {
      this.inFlight = false;
      this.inFlightId = null;
      this.inFlightAction = null;
      this.inFlightPendingId = null;
      // 无论成败都继续放下一个，否则一次失败会把整个队列堵死
      void this.drain();
    }
  }

  /** 真正把一行 JSON 写进子进程并等它的响应 */
  private dispatch(
    action: string,
    params: Record<string, unknown>,
    opts: {
      signal?: AbortSignal;
      onPending?: (pendingId: string) => void;
      onAbandon?: (err: Error) => void;
    } = {}
  ): Promise<Record<string, unknown>> {
    return new Promise<Record<string, unknown>>((resolve, reject) => {
      // 冷启动放在真正送出时才做：排队中的请求不必陪着等握手
      this.ensureStarted()
        .then(() => {
          const id = uuidv4();
          const line = JSON.stringify({ id, action, params });
          const timeoutMs = callTimeoutMsFor(action);

          const timer = setTimeout(() => {
            this.pending.delete(id);
            // 卡死就是这一路死了：只丢弃重建【本应用】的进程，其它应用不受影响
            this.kill(`动作超时（${action}）`);
            reject(
              new Error(
                `WPS 动作超时（${timeoutMs}ms）: ${action}（应用=${this.appKey}）。` +
                  '该应用的常驻进程已重建；因原请求可能执行了一半，这里不自动重放。'
              )
            );
          }, timeoutMs);

          this.pending.set(id, {
            resolve,
            reject,
            timer,
            onAbandon: opts.onAbandon,
          });
          opts.onPending?.(id);

          // 已送出之后的取消：只读动作才能"放弃等待"，写入动作一律不许
          if (opts.signal) {
            opts.signal.addEventListener(
              'abort',
              () => {
                if (!isReadOnlyAction(action)) {
                  // 不取消，也不抛错 —— 让它老实跑完，调用方如实拿到结果
                  this.logStderr(
                    `[${this.appKey}] 收到取消信号，但动作 ${action} 是写入动作：` +
                      '中途放弃会让调用方误判成败，已忽略取消，等它跑完'
                  );
                  return;
                }
                this.abandonPending(id);
                this.cancelled += 1;
              },
              { once: true }
            );
          }

          try {
            if (!this.proc?.stdin || this.proc.stdin.destroyed) {
              throw new Error('stdin 不可写');
            }
            this.proc.stdin.write(line + '\n');
          } catch (err) {
            clearTimeout(timer);
            this.pending.delete(id);
            reject(new Error(`写入常驻进程失败: ${(err as Error).message}`));
          }
        })
        .catch((err: unknown) => reject(err as Error));
    });
  }

  /** 杀掉当前子进程并清理在途请求（下次调用重建）。reason 仅用于诊断。 */
  kill(reason = '手动'): void {
    const p = this.proc;
    this.proc = null;
    this.ready = false;
    if (p) {
      try {
        p.stdin?.end();
      } catch {
        /* stdin 可能已关，忽略 */
      }
      try {
        p.kill();
      } catch {
        /* 进程可能已没，忽略 */
      }
    }
    this.failAll(new Error(`[${this.appKey}] WPS 常驻进程已退出（${reason}），请重试`));
  }

  dispose(): void {
    this.disposed = true;
    this.failQueued(new QueueBusyError(`[${this.appKey}] 桥正在关闭，队列中的请求已取消`));
    this.kill('dispose');
  }

  private failAll(err: Error): void {
    for (const [, p] of this.pending) {
      clearTimeout(p.timer);
      p.reject(err);
    }
    this.pending.clear();
  }

  /** 子进程没了，队列里等位的也没意义了 —— 立刻拒掉，不要让它们干等到排队超时 */
  private failQueued(err: Error): void {
    const tasks = this.queue.splice(0, this.queue.length);
    for (const t of tasks) {
      if (t.timer) clearTimeout(t.timer);
      this.rejected += 1;
      t.reject(err);
    }
  }

  private ensureStarted(): Promise<void> {
    if (this.proc && this.ready) return Promise.resolve();
    if (this.starting) return this.starting;
    this.starting = this.start().finally(() => {
      this.starting = null;
    });
    return this.starting;
  }

  private start(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      let proc: ChildProcess;
      try {
        // -Instance / -App 必须显式传给守护脚本：
        //   Instance 决定 PS 侧是 GetActiveObject（user）还是 New-Object 一个自己的实例（sandbox-N），
        //   App 让脚本知道自己是哪一路（只有 word 支持 target，别的路不该做文档寻址前置检查）。
        // 宿主优先 pwsh7，-STA 保证 COM 套间；没装 pwsh 回退 5.1 —— 见 resolvePowerShellHost()。
        proc = spawn(
          PS_HOST.exe,
          [
            '-NoProfile',
            '-NonInteractive',
            ...(PS_HOST.sta ? ['-STA'] : []),
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            PS_SCRIPT,
            '-Instance',
            this.instance,
            '-App',
            this.appKey,
          ],
          { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] }
        );
      } catch (err) {
        reject(new ChildUnavailableError(`无法启动 PowerShell（${this.appKey}）: ${(err as Error).message}`));
        return;
      }

      this.proc = proc;
      this.ready = false;
      this.stdoutBuf = '';
      if (this.startedAt !== null) this.restarts += 1;
      this.startedAt = Date.now();

      let settled = false;
      const readyTimer = setTimeout(() => {
        if (settled) return;
        settled = true;
        this.kill('握手超时');
        reject(new ChildUnavailableError(`[${this.appKey}] 常驻进程 ${READY_TIMEOUT_MS}ms 内未完成握手`));
      }, READY_TIMEOUT_MS);

      const onReady = (): void => {
        if (settled) return;
        settled = true;
        clearTimeout(readyTimer);
        resolve();
      };

      proc.stdout?.on('data', (chunk: Buffer) => this.onStdout(chunk, onReady));

      proc.stderr?.on('data', (chunk: Buffer) => {
        const s = chunk.toString('utf8').trim();
        if (s) {
          // 桥进程 stdio 被忽略，只能靠文件日志
          this.logStderr(s);
        }
      });

      proc.on('error', (err: Error) => {
        clearTimeout(readyTimer);
        this.proc = null;
        this.ready = false;
        this.failAll(new Error(`[${this.appKey}] 常驻进程出错: ${err.message}`));
        if (!settled) {
          settled = true;
          reject(new ChildUnavailableError(`无法启动 PowerShell 常驻进程（${this.appKey}）: ${err.message}`));
        }
      });

      proc.on('exit', (code, signal) => {
        clearTimeout(readyTimer);
        if (this.proc === proc) {
          this.proc = null;
          this.ready = false;
        }
        this.failAll(new Error(`[${this.appKey}] 常驻进程已退出（code=${code} signal=${signal}），请重试`));
        if (!settled) {
          settled = true;
          reject(new ChildUnavailableError(`[${this.appKey}] 常驻进程启动后立即退出（code=${code}）`));
        }
      });
    });
  }

  /** 按行切分 stdout，各自当一条响应处理 */
  private onStdout(chunk: Buffer, onReady: () => void): void {
    this.stdoutBuf += chunk.toString('utf8');

    let idx = this.stdoutBuf.indexOf('\n');
    while (idx >= 0) {
      const line = this.stdoutBuf.slice(0, idx).trim();
      this.stdoutBuf = this.stdoutBuf.slice(idx + 1);
      if (line) this.handleLine(line, onReady);
      idx = this.stdoutBuf.indexOf('\n');
    }
  }

  private handleLine(line: string, onReady: () => void): void {
    let msg: Record<string, unknown>;
    try {
      msg = JSON.parse(line) as Record<string, unknown>;
    } catch {
      this.logStderr(`常驻进程输出非 JSON，已忽略: ${line.substring(0, 200)}`);
      return;
    }

    const data = msg.data as Record<string, unknown> | undefined;
    if (!this.ready && data && data.ready === true) {
      this.ready = true;
      onReady();
      return;
    }

    const id = typeof msg.id === 'string' ? msg.id : undefined;
    const pending = id ? this.pending.get(id) : undefined;
    if (!pending || !id) {
      this.logStderr(`收到无对应请求的响应（${this.appKey}）: ${line.substring(0, 200)}`);
      return;
    }
    clearTimeout(pending.timer);
    this.pending.delete(id);
    // 调用方已放弃等待：这条回话只用来放行 drain（让队列继续），不交给调用方
    if (pending.abandoned) {
      this.logStderr(`[${this.appKey}] 已放弃的请求回话到达，丢弃（动作=${this.inFlightAction}）`);
      pending.resolve({ cancelled: true });
      return;
    }
    pending.resolve(msg);
  }

  /** 桥是 detached 进程，stdio 被忽略 —— 落文件日志，出问题有据可查 */
  private logStderr(text: string): void {
    try {
      // 延迟 require，避免在不需要日志时也加载 winston
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const { log } = require('../utils/logger') as { log: { warn: (m: string, ctx?: Record<string, unknown>) => void } };
      log.warn(`[bridge:${this.appKey}] ${text}`);
    } catch {
      /* 日志失败不能影响主流程 */
    }
  }
}
