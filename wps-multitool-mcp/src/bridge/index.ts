/**
 * Input: 命名管道上的请求 {"id","action","params","app"}
 * Output: 同一管道上的响应 {"id","success","data?","error?"}（与直跑模式同形）
 * Pos: WPS COM 常驻桥进程（独立于 MCP server 生命周期）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有独立桥进程
 * --------------------
 * MCP server 会被客户端随时重启（改配置、切会话、崩溃重拉）。如果 COM 连接挂在 server 进程里，
 * 一重启就断，用户手上开着的文档也跟着受影响；而且每次重启都留下一批抓着 COM 的 powershell 孤儿。
 * 把连接放进一个【独立、常驻】的桥进程后：
 *   - MCP server 重启，COM 与已打开的文档不受打扰（对标 office-mcp 的 Bridge）
 *   - 收尾只有一处：桥自己按空闲超时退出，子进程由 wps-comd.ps1 的父进程看门狗兜底
 *
 * 为什么走命名管道而不是 TCP
 * --------------------------
 * 命名管道是用户会话内、按用户 ACL 隔离的内核对象：不占端口、不会撞端口、也不对外暴露。
 * office-mcp 用 TCP 是为了"一行命令就能手测"，代价是端口管理；这里用管道，
 * 手测用 node -e 一样能做（见 README）。
 *
 * 协议
 * ----
 * 一行一个 JSON。连接建立后桥先发一条 hello（含 pid / 已用协议版本），
 * 之后每个请求回一条带同 id 的响应。除这些之外桥不往管道写任何东西。
 *
 * 应用路由
 * --------
 * 请求里的 app 决定用哪一路子进程：excel / word / ppt / common。
 * 路由由客户端算好（优先用调用点声明的 appType，其次查生成的路由表），桥只负责按 app 分派。
 *
 * 生命周期
 * --------
 *   空闲超时（默认 10 分钟无客户端消息且无连接）→ 收掉全部子进程后退出。
 *   已有桥在监听同名管道 → 自己直接退出（不重复起）。
 *   WPS_BRIDGE_IDLE_MS 可调；WPS_BRIDGE_PIPE 可换管道名。
 */

import * as net from 'net';
import * as os from 'os';
import * as path from 'path';
import * as fs from 'fs';
import { PsChildDaemon, queueMax, queueTimeoutMs, exclusiveSet } from './ps-child';
import { RUNTIME_ID } from '../runtime-identity';

/** 默认管道名：按用户区分，避免多用户/多会话互相串。前缀用 RUNTIME_ID（跨进程契约，见该文件） */
function defaultPipeName(): string {
  let user = 'default';
  try {
    user = os.userInfo().username || 'default';
  } catch {
    /* 取不到就用 default */
  }
  const safe = user.replace(/[^A-Za-z0-9_.-]/g, '_');
  return `\\\\.\\pipe\\${RUNTIME_ID}-${safe}`;
}

const PIPE_NAME = process.env.WPS_BRIDGE_PIPE || defaultPipeName();
const IDLE_MS = parseInt(process.env.WPS_BRIDGE_IDLE_MS || '600000', 10); // 10 分钟
/** 巡检间隔：自适应，短空闲用短间隔，否则测空闲回收要等到天荒地老 */
const IDLE_TICK_MS = Math.max(1000, Math.min(30000, Math.floor(IDLE_MS / 4)));

/** 已知的应用键。common = 跨应用/基础设施动作（save / openFile / ping / wpsStatus 等） */
const APP_KEYS = ['excel', 'word', 'ppt', 'common'] as const;
type AppKey = (typeof APP_KEYS)[number];

function normalizeApp(v: unknown): AppKey {
  const s = String(v ?? '').toLowerCase();
  if (s === 'excel' || s === 'word' || s === 'ppt') return s;
  return 'common';
}

/** 确保日志目录存在 —— 桥是 detached 进程，stdio 被忽略，只能落文件。
 *  目录名用 RUNTIME_ID：与 logger.ts、ps-daemon.ts 的管道名是同一套约定（见该文件）。 */
try {
  const home = process.env.HOME || process.env.USERPROFILE || os.homedir();
  fs.mkdirSync(path.join(home, `.${RUNTIME_ID}`, 'logs'), { recursive: true });
} catch {
  /* 建不了目录也不能拦住桥 */
}

// ==================== 实例池（第 1 期第 3 章轮 1）====================
//
// 为什么要"一个 instance 一个子进程"
// ----------------------------------
// COM 是 STA：一个进程同一时刻只能跑一个动作。所以"真并行"只有一条路 —— 多进程，
// 每个进程各自持有一个 COM 实例（user 是用户的 WPS，sandbox-N 是我们 New-Object 出来的）。
// 这不是并发加速，是**隔离**：沙箱写坏了不心疼，也不会碰用户正打开的文档。
//
// 容量口径（默认 1，最多 3）
// -------------------------
// 不设 WPS_SANDBOX_MAX → 1（等于最多额外开 1 个实例；只跑 user 时一个都不开）
// 设了 → min(值, 3)；超过 3 按 3 算并告警（实测单实例约 130MB，3 个约 400MB）
const MAX_SANDBOX = 3;

function sandboxCapacity(): number {
  const raw = process.env.WPS_SANDBOX_MAX;
  if (raw === undefined || raw === null || String(raw).trim() === '') return 1;
  const v = parseInt(String(raw), 10);
  if (!Number.isFinite(v) || v <= 0) return 1;
  return v > MAX_SANDBOX ? MAX_SANDBOX : v;
}

/** 启动时把"超上限按 3"这件事说一次，别让它静默生效 */
try {
  const raw = parseInt(String(process.env.WPS_SANDBOX_MAX ?? ''), 10);
  if (Number.isFinite(raw) && raw > MAX_SANDBOX) {
    console.warn(`[bridge] WPS_SANDBOX_MAX=${raw} 超过上限，已按 ${MAX_SANDBOX} 处理`);
  }
} catch {
  /* 环境变量乱七八糟时按默认 1，不拦桥启动 */
}

/** 子进程按 (app, instance) 唯一：同一个 app 的 user 与 sandbox-1 是两个独立进程 */
const children = new Map<string, PsChildDaemon>();

function childKey(app: AppKey, instance: string): string {
  return `${app}:${instance}`;
}

function getChild(app: AppKey, instance: string = 'user'): PsChildDaemon {
  const key = childKey(app, instance);
  let c = children.get(key);
  if (!c) {
    c = new PsChildDaemon(app, instance);
    children.set(key, c);
  }
  return c;
}

/**
 * 把请求里的 instance 解析成【具体实例】。
 *
 * - 缺省 / user        → 'user'
 * - sandbox            → 池里第一个空闲的（load 最小；都忙就挑最闲的排队，不失败）
 * - sandbox-<n>        → 第 n 个；n 越界 → 报错并列出可用范围
 *
 * 为什么 sandbox 不固定映射到 sandbox-1：那样"3 个实例各改一份"永远只会用到第 1 个，
 * 池子形同虚设。按 load 挑是最小惊讶的做法 —— 空闲优先，忙则排到最闲的那个。
 */
function resolveInstance(requested: unknown, app: AppKey = 'word'): string {
  const s = String(requested ?? 'user').trim() || 'user';
  if (s === 'user') return 'user';

  const cap = sandboxCapacity();

  if (s === 'sandbox') {
    let best = 1;
    let bestLoad = Number.POSITIVE_INFINITY;
    for (let i = 1; i <= cap; i++) {
      const c = children.get(childKey(app, `sandbox-${i}`));
      // 没起过的实例 load 视为 0（它就是空闲的）；起过的看真实负载
      const load = c ? c.load : 0;
      if (load < bestLoad) {
        bestLoad = load;
        best = i;
      }
      // 完全空闲（0）就直接用，不用再比
      if (load === 0) break;
    }
    return `sandbox-${best}`;
  }

  const m = s.match(/^sandbox-(\d+)$/);
  if (!m) {
    throw new Error(
      `[勿重试·先改参数] 不认识的 instance「${s}」。可用：user / sandbox / sandbox-1 … sandbox-${cap}`
    );
  }
  const n = Number(m[1]);
  if (n < 1 || n > cap) {
    throw new Error(
      `[勿重试·先改参数] sandbox-${n} 不存在：当前池容量 ${cap}（WPS_SANDBOX_MAX 未设=1，最多 ${MAX_SANDBOX}）。` +
        `可用：sandbox-1 … sandbox-${cap}；不写后缀（sandbox）会自动挑第一个空闲的。`
    );
  }
  return `sandbox-${n}`;
}

const startedAt = Date.now();
let lastActivity = Date.now();
let activeConns = 0;

function writeLine(sock: net.Socket, obj: unknown): void {
  if (sock.destroyed || !sock.writable) return;
  try {
    sock.write(JSON.stringify(obj) + '\n');
  } catch {
    /* 客户端可能已断，忽略 */
  }
}

function childStats(): unknown[] {
  // 按 (app, instance) 遍历真实存在的子进程。不能只按 APP_KEYS 取 ——
  // 那样 sandbox 那几路永远不会出现在状态里，排查"实例到底起没起"时看不到。
  const out: unknown[] = [];
  for (const c of children.values()) out.push(c.stats);
  return out;
}

// ==================== 改动日志（供任务窗格「最近改动」用）====================
// 桥在管道上看得见【每一个】请求，所以由它来记最省事 —— 不需要 MCP server
// 额外上报，也不会漏。只记摘要，不记全文（避免把文档内容泄到日志里）。
interface LogEntry {
  t: number;        // 时间戳 ms
  app: string;
  action: string;
  ok: boolean;
  summary: string;  // 一行摘要，如 "insertText(46 字)"
}

const LOG: LogEntry[] = [];
const LOG_MAX = 200;

/** 从参数里抽一行人话摘要。刻意只取少量关键字段，避免把文档内容整个记下来。 */
function summarize(action: string, params: Record<string, unknown>): string {
  const pick = (...keys: string[]): string => {
    for (const k of keys) {
      const v = params[k];
      if (v === undefined || v === null) continue;
      if (typeof v === 'string') {
        const s = v.replace(/\s+/g, ' ').trim();
        if (s) return s.length > 24 ? s.slice(0, 24) + '…' : s;
      } else if (typeof v === 'number' || typeof v === 'boolean') {
        return String(v);
      }
    }
    return '';
  };
  const detail = pick('text', 'name', 'path', 'label', 'entry', 'mode', 'range', 'sheet', 'slide', 'level');
  const doc = pick('documentName', 'workbookName', 'presentationName');
  const parts = [action];
  if (detail) parts.push(detail);
  if (doc) parts.push(`@${doc}`);
  return parts.join(' ');
}

function recordLog(app: string, action: string, params: Record<string, unknown>, ok: boolean): void {
  LOG.push({ t: Date.now(), app, action, ok, summary: summarize(action, params) });
  if (LOG.length > LOG_MAX) LOG.splice(0, LOG.length - LOG_MAX);
}

// ==================== HTTP 状态端点（供任务窗格轮询）====================
// 为什么需要 HTTP：任务窗格跑在内嵌浏览器里，碰不到命名管道（那是内核对象，
// 只有本进程能连）。而任务面板要显示"桥进程/队列/工具组/最近改动"，
// 所以给桥开一个只读的 localhost HTTP 端点，面板轮询它。
//
// 安全边界：
//   - 只绑 127.0.0.1，不对局域网/公网暴露
//   - 只读（GET），没有写操作
//   - 返回内容不含文档正文（摘要已截断到 24 字）
const HTTP_PORT = parseInt(process.env.WPS_BRIDGE_HTTP_PORT || '27731', 10);

function bridgeStatusPayload(): Record<string, unknown> {
  const childrenList = childStats() as Array<Record<string, unknown>>;
  const inFlight = childrenList.reduce((a, c) => a + Number(c.pending ?? 0), 0);
  const queued = childrenList.reduce((a, c) => a + Number(c.queued ?? 0), 0);
  const running = childrenList
    .filter((c) => c.ready)
    .map((c) => `${String(c.app)}:${String(c.instance ?? 'user')}`);
  return {
    bridgePid: process.pid,
    pipe: PIPE_NAME,
    httpPort: HTTP_PORT,
    startedAt,
    uptimeMs: Date.now() - startedAt,
    clients: activeConns,
    inFlight,
    queued,
    queueMax: queueMax(),
    queueTimeoutMs: queueTimeoutMs(),
    appsRunning: running,
    children: childrenList,
    recent: LOG.slice(-50).reverse(),
  };
}

function startHttpEndpoint(): void {
  const http = require('http') as typeof import('http');
  const server = http.createServer((req, res) => {
    const url = (req.url || '/').split('?')[0];
    if (url === '/status' || url === '/') {
      res.writeHead(200, {
        'Content-Type': 'application/json; charset=utf-8',
        // 任务窗格会频繁轮询，禁缓存
        'Cache-Control': 'no-store',
        // 只允许 file:// 页面（面板是本机文件，Origin 为 "null"）。
        // 曾经这里是 '*'：任何网页都能读这份状态（含子进程 pid 与最近日志行，日志里可能有文档名）。
        // 回环地址不等于只有自己人可访问 —— 用户浏览器里开着的任意站点都能打 127.0.0.1。
        'Access-Control-Allow-Origin': 'null',
      });
      res.end(JSON.stringify(bridgeStatusPayload()));
      return;
    }
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('not found');
  });
  server.on('error', (err: NodeJS.ErrnoException) => {
    // 端口被占就安静退开 —— HTTP 端点只是给面板用的，不该拦住桥本身
    if (err.code === 'EADDRINUSE') {
      try {
        require('../utils/logger') as { log: { warn: (m: string) => void } };
      } catch {
        /* 日志失败不影响 */
      }
    }
  });
  server.listen(HTTP_PORT, '127.0.0.1', () => {
    // 不往 stdout 打印（桥的 stdio 是 ignore 的）
  });
}

async function handleRequest(
  sock: net.Socket,
  req: Record<string, unknown>
): Promise<void> {
  const id = typeof req.id === 'string' ? req.id : '';
  const action = String(req.action ?? '');
  const params = (req.params as Record<string, unknown>) ?? {};

  // 桥自身的状态查询：不落到 PowerShell，直接回
  if (action === '__bridgeStatus') {
    writeLine(sock, {
      id,
      success: true,
      data: {
        bridgePid: process.pid,
        pipe: PIPE_NAME,
        uptimeMs: Date.now() - startedAt,
        clients: activeConns,
        idleMs: Date.now() - lastActivity,
        queue: {
          max: queueMax(),
          timeoutMs: queueTimeoutMs(),
          exclusiveActions: [...exclusiveSet()],
        },
        children: childStats(),
      },
    });
    return;
  }

  if (!action) {
    writeLine(sock, { id, success: false, error: '请求缺少 action' });
    return;
  }

  // 取消：按【原请求的 id】撤。排队中的一定撤得掉；已在途的只有只读动作会"放弃等待"。
  if (action === '__cancel') {
    const targetId = String(params.targetId ?? '');
    const app = normalizeApp(req.app);
    if (!targetId) {
      writeLine(sock, { id, success: false, error: '__cancel 缺少 targetId' });
      return;
    }
    // 取消也要按 instance 定位：同一个 app 的 user 与 sandbox-1 是两个独立队列，
    // 撤错队列等于"以为取消了、其实还在跑"。
    let instance = 'user';
    try {
      instance = resolveInstance(req.instance, app);
    } catch {
      /* 实例名非法时按 user 处理：取消是尽力而为，不该因为名字错而整个失败 */
    }
    const r = getChild(app, instance).cancelByExternalId(targetId);
    writeLine(sock, { id, success: r.ok, data: { cancelled: r.ok, reason: r.reason } });
    return;
  }

  const app = normalizeApp(req.app);
  try {
    const instance = resolveInstance(req.instance, app);
    const msg = await getChild(app, instance).call(action, params, { externalId: id });
    recordLog(app, action, params, msg.success !== false);
    writeLine(sock, {
      id,
      success: msg.success !== false,
      data: msg.data,
      error: msg.error,
      instance,
    });
  } catch (err) {
    // 基础设施问题（起不来/超时/退出/队列满/实例名非法）：如实回错误，不假装成功
    recordLog(app, action, params, false);
    writeLine(sock, {
      id,
      success: false,
      error: err instanceof Error ? err.message : String(err),
    });
  }
}

const server = net.createServer((sock) => {
  activeConns += 1;
  lastActivity = Date.now();
  sock.setNoDelay(true);

  // hello：客户端据此确认桥是活的（而不是连上了一个半死的进程）
  writeLine(sock, {
    type: 'hello',
    bridgePid: process.pid,
    pipe: PIPE_NAME,
    protocol: 1,
    apps: APP_KEYS,
  });

  let buf = '';
  sock.on('data', (chunk: Buffer) => {
    lastActivity = Date.now();
    buf += chunk.toString('utf8');
    let idx = buf.indexOf('\n');
    while (idx >= 0) {
      const line = buf.slice(0, idx).trim();
      buf = buf.slice(idx + 1);
      if (line) {
        let req: Record<string, unknown>;
        try {
          req = JSON.parse(line) as Record<string, unknown>;
        } catch {
          writeLine(sock, { id: '', success: false, error: '请求不是合法 JSON' });
          idx = buf.indexOf('\n');
          continue;
        }
        // 不 await：允许同一连接上并发多路请求（不同应用各走各的子进程）
        void handleRequest(sock, req);
      }
      idx = buf.indexOf('\n');
    }
  });

  const onGone = (): void => {
    activeConns = Math.max(0, activeConns - 1);
    lastActivity = Date.now();
  };
  sock.on('close', onGone);
  sock.on('error', onGone);
});

server.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EADDRINUSE') {
    // 已经有一个桥在跑（管道被占）—— 正常退出，让客户端连那个
    process.exit(0);
  }
  process.exit(1);
});

server.listen(PIPE_NAME, () => {
  // 管道已就绪。注意：不要往 stdout 打印，桥的 stdio 是 ignore 的。
});

// HTTP 状态端点（任务窗格轮询用）。放在 listen 之后，确保桥本身先就绪。
startHttpEndpoint();

// 空闲回收：没人连、且久无消息 → 收掉子进程后退出，不留孤儿
const idleTimer = setInterval(() => {
  if (activeConns > 0) return;
  if (Date.now() - lastActivity < IDLE_MS) return;
  for (const c of children.values()) c.dispose();
  children.clear();
  server.close(() => process.exit(0));
  // 兜底：close 回调没来也要走
  setTimeout(() => process.exit(0), 2000).unref();
}, IDLE_TICK_MS);
idleTimer.unref();

// 桥退出前收掉子进程（子进程另有父进程看门狗兜底）
const shutdown = (): void => {
  for (const c of children.values()) c.dispose();
  process.exit(0);
};
process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
