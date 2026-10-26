/**
 * Input: 动作名、参数、应用键
 * Output: 动作结果（与直跑模式同形：{ success, data?, error? }）
 * Pos: 常驻桥进程的客户端（MCP server 侧）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 这一版和上一版的区别（①A -> ①B）
 * -------------------------------
 * ①A：MCP server 自己 spawn 一个常驻 powershell，进程随 server 一起生灭。
 * ①B：改成一个【独立、常驻】的桥进程（dist/bridge/index.js），server 只是它的客户端。
 *      桥里再按应用（excel / word / ppt / common）各持一个常驻 powershell。
 *
 * 换来什么
 * --------
 *   - MCP server 被客户端重启（改配置/切会话/崩溃重拉）时，COM 连接与已打开的文档不受打扰
 *   - 一应用一进程：WPS 文字卡死不会堵住表格；桥只重建那一路
 *   - 不同应用可并发（同一连接上多请求并行，各走各的子进程）
 *
 * 谁负责收尾
 * ----------
 * 桥自己按空闲超时退出（默认 10 分钟无客户端），子进程由 wps-comd.ps1 的父进程看门狗兜底。
 * 所以本客户端 dispose() 【只断连接，不杀桥】—— 让桥的存活不依赖 server 的存活，这正是它的意义。
 * 需要"server 一退桥就退"的场合（例如自动化测试）设 WPS_BRIDGE_EPHEMERAL=1。
 *
 * 失败语义（红线）
 * --------------
 *   DaemonUnavailableError 只能表示【请求从未送出】（连不上桥、桥起不来）—— 调用方可安全退回直跑。
 *   请求已经写进管道之后出的任何错（超时、桥挂了），都可能已经在 WPS 里执行了一半，
 *   绝不能自动重放，一律如实抛出。
 */

import * as net from 'net';
import * as os from 'os';
import * as path from 'path';
import { ChildProcess, spawn } from 'child_process';
import { v4 as uuidv4 } from 'uuid';
import { log } from '../utils/logger';
import { CancelError } from '../bridge/cancellable';
import { RUNTIME_ID } from '../runtime-identity';

/** 桥进程入口（与客户端同时随包分发） */
const BRIDGE_ENTRY = path.join(__dirname, '..', 'bridge', 'index.js');

/** 桥起不来的情况下，等它监听管道的上限 */
const BRIDGE_START_TIMEOUT_MS = parseInt(process.env.WPS_BRIDGE_START_TIMEOUT || '30000', 10);
/** 单次动作上限（桥侧还有一层同样的超时）；每次调用时重读，便于按场景临时收紧 */
function callTimeoutMs(): number {
  const v = parseInt(process.env.WPS_DAEMON_TIMEOUT || '120000', 10);
  return Number.isFinite(v) && v > 0 ? v : 120000;
}
/** 重连尝试间隔 */
const RECONNECT_DELAY_MS = 300;

function defaultPipeName(): string {
  let user = 'default';
  try {
    user = os.userInfo().username || 'default';
  } catch {
    /* ignore */
  }
  const safe = user.replace(/[^A-Za-z0-9_.-]/g, '_');
  return `\\\\.\\pipe\\${RUNTIME_ID}-${safe}`;
}

export const BRIDGE_PIPE = process.env.WPS_BRIDGE_PIPE || defaultPipeName();

/** 桥【整条】不可用（请求从未送出）时抛这个；调用方可安全退回直跑。 */
export class DaemonUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'DaemonUnavailableError';
  }
}

interface PendingCall {
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

export interface BridgeClientStats {
  enabled: boolean;
  connected: boolean;
  pipe: string;
  bridgePid: number | null;
  uptimeMs: number;
  restarts: number;
  pending: number;
  disposed: boolean;
  /** 累计被取消的次数 */
  cancelled: number;
}

export class PsDaemon {
  private sock: net.Socket | null = null;
  private buf = '';
  private pending = new Map<string, PendingCall>();
  private connecting: Promise<void> | null = null;
  private helloOk = false;
  private disposed = false;
  private connectedAt: number | null = null;
  private restarts = 0;
  private bridgeProc: ChildProcess | null = null;
  /** 累计被取消的次数 */
  private cancelled = 0;

  get isReady(): boolean {
    return this.sock !== null && this.helloOk;
  }

  get pendingCount(): number {
    return this.pending.size;
  }

  get stats(): BridgeClientStats {
    return {
      enabled: process.env.WPS_USE_DAEMON !== '0',
      connected: this.isReady,
      pipe: BRIDGE_PIPE,
      bridgePid: this.bridgeProc?.pid ?? null,
      uptimeMs: this.connectedAt === null ? 0 : Date.now() - this.connectedAt,
      restarts: this.restarts,
      pending: this.pending.size,
      disposed: this.disposed,
      cancelled: this.cancelled,
    };
  }

  /**
   * 调用一个动作。业务成功/失败都以响应对象 resolve；只有基础设施问题才 reject。
   *
   * 取消：传 signal 即可。server 侧收到 notifications/cancelled 时这个 signal 会 abort，
   * 此时本端立刻 reject（不再等），同时给桥发一条 __cancel —— 让桥那边也把请求撤了
   * （排队中的直接出队；已在途且是只读动作则放弃等待）。
   */
  async call(
    action: string,
    params: Record<string, unknown> = {},
    app: string = 'common',
    opts: { signal?: AbortSignal; instance?: string } = {}
  ): Promise<unknown> {
    if (this.disposed) {
      throw new DaemonUnavailableError('桥客户端已关闭（dispose 之后不再接受调用）');
    }
    if (opts.signal?.aborted) {
      throw new CancelError(`请求在送出前已被取消（动作=${action}），从未送达`);
    }
    await this.ensureConnected();

    const id = uuidv4();
    // instance 是【路由键】不是 COM 参数：桥据此决定投给哪个子进程（user / sandbox-N）。
    // 缺省写 user，让桥端只认一个字段，不用到处判 undefined。
    const instance = (opts.instance || 'user').trim() || 'user';
    const line = JSON.stringify({ id, action, params, app, instance });
    const timeoutMs = callTimeoutMs();

    return new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        // 卡住就判定这条连接死了：断开重连（下次调用重建桥连接），但不重放本请求
        log.error('桥动作超时，断开连接', undefined, { action, app, id, timeoutMs });
        this.dropSocket(new Error('桥连接因超时被断开'));
        reject(
          new Error(
            `WPS 动作超时（${timeoutMs}ms）: ${action}（应用=${app}）。` +
              '桥连接已断开重连；因原请求可能执行了一半，这里不自动重放。'
          )
        );
      }, timeoutMs);

      this.pending.set(id, { resolve, reject, timer });

      // 取消：本端立刻不等了，同时请桥把对应的请求撤掉
      if (opts.signal) {
        opts.signal.addEventListener(
          'abort',
          () => {
            clearTimeout(timer);
            this.pending.delete(id);
            this.cancelled += 1;
            reject(new CancelError(`请求已被取消（动作=${action}，应用=${app}）`));
            // 通知桥（fire-and-forget：桥回不回都不影响本端已经放弃）
            // 必须带上 instance：桥按 (app, instance) 定位子进程，漏了会撤到别人的队列里。
            void this.sendCancel(id, app, instance).catch(() => {
              /* 通知失败无所谓，本端已经不等了 */
            });
          },
          { once: true }
        );
      }

      try {
        if (!this.sock || this.sock.destroyed) throw new Error('桥连接不可写');
        this.sock.write(line + '\n');
      } catch (err) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(new Error(`写入桥失败: ${(err as Error).message}`));
      }
    });
  }

  /** 给桥发一条取消指令（不等待结果） */
  private async sendCancel(targetId: string, app: string, instance: string = 'user'): Promise<void> {
    try {
      await this.call('__cancel', { targetId }, app, { instance });
    } catch {
      /* 忽略：取消通知是尽力而为 */
    }
  }

  /** 查桥自身状态（不落到 PowerShell）：桥 PID、各应用子进程状态等 */
  async bridgeStatus(): Promise<unknown> {
    return this.call('__bridgeStatus', {}, 'common');
  }

  /**
   * 只断开客户端连接。默认【不杀桥】—— 桥要活过 server 重启，这是设计意图。
   * WPS_BRIDGE_EPHEMERAL=1 时才顺手把桥停掉（测试用）。
   */
  dispose(): void {
    this.disposed = true;
    const ephemeral = process.env.WPS_BRIDGE_EPHEMERAL === '1';
    if (ephemeral) {
      try {
        this.sock?.end();
      } catch {
        /* ignore */
      }
      try {
        this.bridgeProc?.kill();
      } catch {
        /* ignore */
      }
      this.bridgeProc = null;
    }
    this.dropSocket(new Error('桥客户端已关闭'));
    this.sock = null;
    this.helloOk = false;
  }

  private dropSocket(err: Error): void {
    const s = this.sock;
    this.sock = null;
    this.helloOk = false;
    if (s) {
      try {
        s.destroy();
      } catch {
        /* ignore */
      }
    }
    this.failAll(err);
  }

  private failAll(err: Error): void {
    for (const [, p] of this.pending) {
      clearTimeout(p.timer);
      p.reject(err);
    }
    this.pending.clear();
  }

  private ensureConnected(): Promise<void> {
    if (this.sock && this.helloOk) return Promise.resolve();
    if (this.connecting) return this.connecting;
    this.connecting = this.connect().finally(() => {
      this.connecting = null;
    });
    return this.connecting;
  }

  private async connect(): Promise<void> {
    // 先试直接连（桥可能已经在跑了）
    try {
      await this.attachToPipe(1500);
      return;
    } catch {
      /* 没桥，往下走：拉一个起来 */
    }

    this.spawnBridge();
    const deadline = Date.now() + BRIDGE_START_TIMEOUT_MS;
    let lastErr: Error = new Error('未知');
    while (Date.now() < deadline) {
      await sleep(RECONNECT_DELAY_MS);
      try {
        await this.attachToPipe(2000);
        return;
      } catch (err) {
        lastErr = err as Error;
      }
    }
    throw new DaemonUnavailableError(
      `桥进程 ${BRIDGE_START_TIMEOUT_MS}ms 内未就绪（管道 ${BRIDGE_PIPE}）：${lastErr.message}`
    );
  }

  /** 连上管道并等到 hello 行（确认对面是个活桥，而不是半死进程） */
  private attachToPipe(timeoutMs: number): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const sock = net.connect(BRIDGE_PIPE);
      let settled = false;
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        try {
          sock.destroy();
        } catch {
          /* ignore */
        }
        reject(new Error('连接桥超时（hello 未到）'));
      }, timeoutMs);

      const finish = (err?: Error): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (err) {
          try {
            sock.destroy();
          } catch {
            /* ignore */
          }
          reject(err);
          return;
        }
        this.attachSocket(sock);
        resolve();
      };

      sock.setNoDelay(true);
      let handshakeBuf = '';
      sock.on('data', function onHello(chunk: Buffer) {
        handshakeBuf += chunk.toString('utf8');
        const idx = handshakeBuf.indexOf('\n');
        if (idx < 0) return;
        const line = handshakeBuf.slice(0, idx).trim();
        sock.off('data', onHello);
        try {
          const msg = JSON.parse(line) as Record<string, unknown>;
          if (msg.type !== 'hello') throw new Error('桥未按协议发 hello');
        } catch (err) {
          finish(new Error(`桥握手失败: ${(err as Error).message}`));
          return;
        }
        finish();
      });

      sock.on('error', (err: Error) => finish(err));
    });
  }

  /** 把一个已完成握手的 socket 接进收发循环 */
  private attachSocket(sock: net.Socket): void {
    const old = this.sock;
    if (old) {
      try {
        old.destroy();
      } catch {
        /* ignore */
      }
    }
    if (this.connectedAt !== null) this.restarts += 1;
    this.connectedAt = Date.now();
    this.sock = sock;
    this.helloOk = true;
    this.buf = '';

    sock.on('data', (chunk: Buffer) => {
      this.buf += chunk.toString('utf8');
      let idx = this.buf.indexOf('\n');
      while (idx >= 0) {
        const line = this.buf.slice(0, idx).trim();
        this.buf = this.buf.slice(idx + 1);
        if (line) this.handleLine(line);
        idx = this.buf.indexOf('\n');
      }
    });

    sock.on('close', () => {
      if (this.sock === sock) {
        this.helloOk = false;
        this.sock = null;
        this.failAll(new Error('桥连接已关闭，请重试'));
      }
    });
    sock.on('error', (err: Error) => {
      log.warn('桥连接出错', { error: err.message });
    });

    log.info('已连上 WPS 桥', { pipe: BRIDGE_PIPE });
  }

  private handleLine(line: string): void {
    let msg: Record<string, unknown>;
    try {
      msg = JSON.parse(line) as Record<string, unknown>;
    } catch {
      log.warn('桥输出非 JSON，已忽略', { line: line.substring(0, 200) });
      return;
    }
    const id = typeof msg.id === 'string' ? msg.id : undefined;
    const pending = id ? this.pending.get(id) : undefined;
    if (!pending || !id) {
      log.warn('收到无对应请求的桥响应', { id, body: line.substring(0, 200) });
      return;
    }
    clearTimeout(pending.timer);
    this.pending.delete(id);
    pending.resolve(msg);
  }

  private spawnBridge(): void {
    try {
      const child = spawn(process.execPath, [BRIDGE_ENTRY], {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
        env: process.env,
      });
      child.unref();
      this.bridgeProc = child;
      log.info('已拉起 WPS 桥进程', { pid: child.pid, entry: BRIDGE_ENTRY, pipe: BRIDGE_PIPE });
    } catch (err) {
      log.error('拉起桥进程失败', err as Error);
    }
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

/** 全局单例：一个 MCP server 进程只维护一个桥客户端（连接） */
export const psDaemon = new PsDaemon();

export default psDaemon;
