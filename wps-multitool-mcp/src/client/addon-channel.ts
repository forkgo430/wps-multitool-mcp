/**
 * Input: 动作名、参数、应用键、实例名（与 B0 转发点签名一致，见 wps-client.ts execWpsAction）
 * Output: 动作结果（Promise<unknown>，与 Windows 通道同形）
 * Pos: Linux/macOS 加载项通道的 Node 侧服务（B2 骨架，长轮询版）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么由本模块持有服务（计划 §3.1）
 * --------------------------------
 * 桥（src/bridge/**）是 Windows 专用：spawn PowerShell、监听 \\.\pipe\，Linux 上两者都不存在。
 * 所以加载项通道由 MCP server 进程本身持有：独立端口、独立认证，与面板只读端口 27731
 * 物理分开（C2）——执行通道能改文档，暴露面必须小于只读通道。
 *
 * 连接方向（计划 §4.4）
 * ------------------
 * 加载项跑在 WPS 进程内的 webview 里，主动【出站】连 Node；Node 不去连 WPS
 * （让 WPS 监听端口要动防火墙，且嵌入式 webview 未必能监听）。
 *
 * 三型同端口 + 握手声明型别（计划 §12.1，B2 定案）
 * ---------------------------------------------
 * S1 实测三型对象模型互相看不见（wps 只有 Documents / et 只有 Workbooks / wpp 只有
 * Presentations），Linux 侧必须部署三个加载项，各自连到同一个端口，靠握手 body 的
 * type 字段记会话。同型后来者覆盖前者（WPS 重启后加载项重连的场景）。
 *
 * 传输：长轮询（计划 §4.6）
 * ----------------------
 * 加载项 GET /poll 取任务（无任务挂起到 10s）、POST /result 交结果。
 * 轮询的已知风险照记：WPS 最小化/省电时 webview 定时器可能被节流到分钟级——
 * 不报错、不掉线、只是慢（R6）。spike 用轮询，上线按 §4.6 评估换 WS；
 * 协议帧与 Windows 管道同构（{ id, action, params, app }，ps-daemon.ts:150），
 * 换 transport 不改动作层、不用重测覆盖率。
 *
 * 失败语义（红线，照抄 ps-daemon.ts:24-28 的契约）
 * ---------------------------------------------
 *   AddonUnavailableError 只能表示【请求从未送出】（无该型会话 / 会话在途忙 /
 *   未被取走的任务被顶替）。调用方可安全重试或降级。
 *   任务已被 /poll 的响应送达加载项之后出的任何错（超时、无回包），都可能已经在
 *   WPS 里执行了一半 —— 绝不自动重放，一律如实抛普通 Error。
 *   "已送出"的边界 = /poll 的 HTTP 响应写出到 socket；poll 请求在写出前断开，
 *   任务退回轮询池。写出失败无法回退，由超时兜底（不重放）。
 *
 * 排队（计划 §4.5 / 任务卡 §0）
 * --------------------------
 * 每会话同时只挂 1 个待执行任务（与桥"STA 在途恒为 1"对齐）：
 *   - 轮询池里已有未取走任务，新调用到达 → 旧任务从未送出，判 AddonUnavailableError，新任务顶上；
 *   - 已有任务在途（已送达加载项未回）→ 新请求从未送出，判 AddonUnavailableError（不排队堆积）。
 *
 * instance 参数：v1 轮询版忽略。S1 实测 Linux WPS 单实例（R15），Windows 的实例池隔离
 * 不存在，多目标语义由动作层"同实例内按名钉文档"承接。
 *
 * 本模块由谁启动：MCP server 进程按需调用 start()；测试可传 { port, token } 随机端口自起自停。
 */

import * as http from 'http';
import { AddressInfo } from 'net';
import { v4 as uuidv4 } from 'uuid';

/** 执行通道端口（C2：与面板只读端口 27731 物理分开）。只 bind 127.0.0.1。 */
export const ADDON_PORT = (() => {
  const v = parseInt(process.env.WPS_ADDON_PORT || '27740', 10);
  return Number.isFinite(v) && v > 0 ? v : 27740;
})();

/** 长轮询挂起上限：无任务时 /poll 最多挂这么久就回 idle，让加载项重新发起 */
const POLL_HANG_MS = 10_000;

/** 单动作上限沿用桥口径（ps-daemon.ts:46-49）；每次调用时重读，便于按场景临时收紧 */
function callTimeoutMs(): number {
  const v = parseInt(process.env.WPS_DAEMON_TIMEOUT || '120000', 10);
  return Number.isFinite(v) && v > 0 ? v : 120000;
}

/** 通道【整条】不可用（请求从未送出）时抛这个；调用方可安全重试/降级。 */
export class AddonUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AddonUnavailableError';
  }
}

export type AddonType = 'wps' | 'et' | 'wpp';
const ADDON_TYPES: readonly AddonType[] = ['wps', 'et', 'wpp'];

/**
 * 动作路由（任务卡 §0 拍死的 v1 简化）：word→wps 会话、excel→et、ppt→wpp、common→wps。
 * common 挂在文字型上是 v1 的简化 —— Windows 侧 common 动作多是进程级（quitWps/wpsStatus 等），
 * B3 动作层实现时若发现需要按动作名细分会话，改这里一个函数就够。
 */
function sessionTypeOf(app: string): AddonType {
  switch (app) {
    case 'excel':
      return 'et';
    case 'ppt':
      return 'wpp';
    default:
      return 'wps';
  }
}

/** 一个待执行/在途任务 */
interface Task {
  id: string;
  action: string;
  params: Record<string, unknown>;
  app: string;
  createdAt: number;
  /** 超时计时器：pending 阶段超时=未送出；inflight 阶段超时=可能执行了一半，绝不重放 */
  timer: NodeJS.Timeout;
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
}

/** 一个已握手的加载项会话（每型至多一个，同型后来者覆盖前者） */
interface Session {
  sessionId: string;
  type: AddonType;
  addonVersion: string;
  connectedAt: string;
  /** 轮询池：已挂单、未被 /poll 取走（此阶段的任务 = 从未送出） */
  pending: Task | null;
  /** 在途：已随 /poll 响应送达加载项、等 /result（此阶段的失败绝不重放） */
  inflight: Task | null;
  /** 挂起中的长轮询：reply(status, body) 写响应；timer 到点回 idle */
  pollWait: {
    reply: (status: number, body: Record<string, unknown>) => void;
    timer: NodeJS.Timeout;
  } | null;
}

const sessions = new Map<AddonType, Session>();
const sessionById = new Map<string, Session>();
let server: http.Server | null = null;
let currentPort = 0;
let currentToken = '';

function disposeTask(task: Task, err: Error): void {
  clearTimeout(task.timer);
  task.reject(err);
}

/** 会话终结（被同型覆盖 / 通道停止）：pending=从未送出；inflight=可能执行了一半，绝不判"未送出" */
function disposeSession(s: Session, reason: string): void {
  if (s.pending) {
    disposeTask(s.pending, new AddonUnavailableError(`WPS 加载项通道：任务 ${s.pending.action} 在送达前会话失效（${reason}），请求从未送出，可安全重试`));
    s.pending = null;
  }
  if (s.inflight) {
    disposeTask(s.inflight, new Error(`WPS 加载项通道：任务 ${s.inflight.action} 已送达加载项后会话失效（${reason}），可能已在 WPS 里执行了一半 —— 不重放，请人工核实`));
    s.inflight = null;
  }
  if (s.pollWait) {
    clearTimeout(s.pollWait.timer);
    try {
      s.pollWait.reply(404, { error: `会话已失效（${reason}），请重新握手` });
    } catch {
      /* 响应可能已不可写 */
    }
    s.pollWait = null;
  }
  sessionById.delete(s.sessionId);
}

/** 任务从轮询池转入在途：重置计时器只计剩余额度（总上限 = 单动作超时，不因排队翻倍） */
function promoteToInflight(s: Session, task: Task): void {
  clearTimeout(task.timer);
  const remain = Math.max(task.createdAt + callTimeoutMs() - Date.now(), 1);
  task.timer = setTimeout(() => onTaskTimeout(task, s), remain);
  s.inflight = task;
}

/** 把轮询池顶端的任务随【挂起中的 /poll】送达：从这一刻起任务"已送出" */
function deliverPending(s: Session): void {
  const task = s.pending;
  const wait = s.pollWait;
  if (!task || !wait) return;
  s.pending = null;
  s.pollWait = null;
  clearTimeout(wait.timer);
  promoteToInflight(s, task);
  try {
    wait.reply(200, { id: task.id, action: task.action, params: task.params, app: task.app });
  } catch {
    /* 写出失败无法回退：任务按已送出处理，由超时兜底如实抛（不重放） */
  }
}

function onTaskTimeout(task: Task, s: Session): void {
  if (s.pending === task) {
    s.pending = null;
    task.reject(
      new AddonUnavailableError(
        `WPS 加载项通道：动作 ${task.action} 等待 ${callTimeoutMs()}ms 仍无人取走（type=${s.type}）。` +
          '请求从未送出，可安全重试；若持续出现，检查加载项是否已从 WPS UI 打开（Linux 必须手动打开）'
      )
    );
    return;
  }
  if (s.inflight === task) {
    s.inflight = null;
    task.reject(
      new Error(
        `WPS 加载项通道：动作 ${task.action} 已送达加载项后超时（${callTimeoutMs()}ms，type=${s.type}）。` +
          '可能已在 WPS 里执行了一半 —— 绝不自动重放，请人工核实'
      )
    );
  }
}

/* ---------------------------------------------------------------------------
 * HTTP 端点（三个：/handshake、/poll、/result）
 * 执行通道不做 CORS（计划 §4.4）：不回 Access-Control-* 头，OPTIONS 直接 405。
 * ------------------------------------------------------------------------ */

function readBody(req: http.IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks: Buffer[] = [];
    req.on('data', (c: Buffer) => {
      size += c.length;
      if (size > 2 * 1024 * 1024) {
        reject(new Error('请求体超过 2MB 上限'));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function sendJson(res: http.ServerResponse, status: number, body: Record<string, unknown>): void {
  // 刻意不带 CORS 头（C2）：面板只读通道的宽松策略不许套到能写文档的执行通道上
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(body));
}

async function handleHandshake(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
  let body: Record<string, unknown>;
  try {
    body = JSON.parse((await readBody(req)) || '{}') as Record<string, unknown>;
  } catch {
    sendJson(res, 400, { error: 'handshake 请求体必须是 JSON' });
    return;
  }
  if (String(body.token ?? '') !== currentToken) {
    sendJson(res, 401, { error: '令牌不匹配：加载项侧的 token 与服务端不一致（WPS_ADDON_TOKEN / 启动日志）' });
    return;
  }
  const type = body.type;
  if (typeof type !== 'string' || !ADDON_TYPES.includes(type as AddonType)) {
    sendJson(res, 400, { error: `type 必须是 ${ADDON_TYPES.join('|')} 之一（三型对象模型隔离，S1 实测）` });
    return;
  }
  const old = sessions.get(type as AddonType);
  if (old) disposeSession(old, `被同型（${type}）新加载项覆盖`);
  const s: Session = {
    sessionId: uuidv4(),
    type: type as AddonType,
    addonVersion: String(body.addonVersion ?? ''),
    connectedAt: new Date().toISOString(),
    pending: null,
    inflight: null,
    pollWait: null,
  };
  sessions.set(s.type, s);
  sessionById.set(s.sessionId, s);
  sendJson(res, 200, { sessionId: s.sessionId, pollIntervalMs: 50 });
}

function handlePoll(_req: http.IncomingMessage, res: http.ServerResponse, url: URL): void {
  const sessionId = url.searchParams.get('sessionId') || '';
  const s = sessionById.get(sessionId);
  if (!s) {
    sendJson(res, 404, { error: '会话不存在或已失效，请重新握手' });
    return;
  }
  // 上一轮挂起的 poll 还没回话就又来 poll：先回掉旧的（idle），避免连接堆积
  if (s.pollWait) {
    clearTimeout(s.pollWait.timer);
    try {
      s.pollWait.reply(200, { idle: true });
    } catch {
      /* 旧连接可能已断 */
    }
    s.pollWait = null;
  }
  if (s.pending) {
    // 任务先挂、poll 后到：没有挂起中的 poll 可借，直接用当前这条响应送帧
    const task = s.pending;
    s.pending = null;
    promoteToInflight(s, task);
    sendJson(res, 200, { id: task.id, action: task.action, params: task.params, app: task.app });
    return;
  }
  const wait = {
    reply: (status: number, body: Record<string, unknown>) => {
      if (res.writableEnded || res.destroyed) return;
      sendJson(res, status, body);
    },
    timer: null as unknown as NodeJS.Timeout,
  };
  wait.timer = setTimeout(() => {
    if (s.pollWait && s.pollWait.reply === wait.reply) {
      s.pollWait = null;
      wait.reply(200, { idle: true });
    }
  }, POLL_HANG_MS);
  s.pollWait = wait;
  // 客户端在响应写出前断开：撤下挂起等待即可，任务（若有）本来就不在这条连接上。
  // 为什么挂 res 不挂 req：Node ≥13 的 IncomingMessage 'close' 在【消息完成】时就触发
  // （GET 无 body，完成得极早），会把正常挂起的长轮询误判成断开（实测 routeProbe
  // 的第 4 个用例因此"无人取走"）；ServerResponse 'close' 只在连接真断或响应写完时
  // 触发，配合 writableEnded 就能区分"还没写出就断了"与"已写出正常收尾"。
  res.on('close', () => {
    if (res.writableEnded) return; // 响应已写出（帧或 idle 已送达），不是断开
    if (s.pollWait && s.pollWait.reply === wait.reply) {
      clearTimeout(wait.timer);
      s.pollWait = null;
    }
  });
}

async function handleResult(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
  let body: Record<string, unknown>;
  try {
    body = JSON.parse((await readBody(req)) || '{}') as Record<string, unknown>;
  } catch {
    sendJson(res, 400, { error: 'result 请求体必须是 JSON' });
    return;
  }
  const s = sessionById.get(String(body.sessionId ?? ''));
  if (!s) {
    sendJson(res, 404, { error: '会话不存在或已失效' });
    return;
  }
  const task = s.inflight;
  if (!task || task.id !== String(body.id ?? '')) {
    sendJson(res, 409, { error: '没有匹配的在途任务（可能已超时或会话被覆盖）' });
    return;
  }
  s.inflight = null;
  clearTimeout(task.timer);
  if (body.ok === false) {
    const why = String(body.error ?? '加载项返回失败');
    const unsupported = body.unsupported ? `（unsupported=${String(body.unsupported)}）` : '';
    task.reject(new Error(`WPS 加载项通道：动作 ${task.action} 执行失败${unsupported}：${why}`));
  } else {
    task.resolve(body.data);
  }
  sendJson(res, 200, { received: true });
}

async function route(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
  const url = new URL(req.url || '/', 'http://127.0.0.1');
  try {
    if (req.method === 'POST' && url.pathname === '/handshake') {
      await handleHandshake(req, res);
      return;
    }
    if (req.method === 'GET' && url.pathname === '/poll') {
      handlePoll(req, res, url);
      return;
    }
    if (req.method === 'POST' && url.pathname === '/result') {
      await handleResult(req, res);
      return;
    }
    if (req.method === 'OPTIONS') {
      // 执行通道不做 CORS（C2）：预检一律拒绝
      sendJson(res, 405, { error: '执行通道不做 CORS，浏览器来源请走 27731 只读面板' });
      return;
    }
    sendJson(res, 404, { error: `未知端点 ${req.method} ${url.pathname}（可用：POST /handshake、GET /poll、POST /result）` });
  } catch (e) {
    if (!res.writableEnded) {
      sendJson(res, 500, { error: `通道内部错误：${e instanceof Error ? e.message : String(e)}` });
    }
  }
}

export interface StartOptions {
  /** 覆盖 WPS_ADDON_PORT / 默认 27740（测试用随机端口） */
  port?: number;
  /** 覆盖 WPS_ADDON_TOKEN / 随机生成（测试专令牌走这里，不落任何文件） */
  token?: string;
}

export interface StartInfo {
  port: number;
  token: string;
  /** 令牌是否为本次启动随机生成（生成时已打到 stderr，不进文件、不进 git） */
  tokenGenerated: boolean;
}

/**
 * 起通道服务。幂等：已在跑时返回现有端口与令牌。
 * 令牌来源优先级：显式参数 > WPS_ADDON_TOKEN > 启动时随机生成（打到 stderr `[addon-channel] token=…`，
 * 不进任何文件、不进 git —— 任务卡 §0 红线）。
 */
export function start(opts: StartOptions = {}): Promise<StartInfo> {
  if (server) return Promise.resolve({ port: currentPort, token: currentToken, tokenGenerated: false });
  let generated = false;
  if (opts.token) {
    currentToken = opts.token;
  } else if (process.env.WPS_ADDON_TOKEN) {
    currentToken = process.env.WPS_ADDON_TOKEN;
  } else {
    currentToken = uuidv4().replace(/-/g, '');
    generated = true;
    // 令牌只在 stderr 出现一次；不写日志文件、不写任何被 git 跟踪的文件
    console.error(`[addon-channel] token=${currentToken}`);
  }
  return new Promise<StartInfo>((resolve, reject) => {
    const srv = http.createServer((req, res) => {
      void route(req, res);
    });
    srv.on('error', (err) => {
      server = null;
      reject(err);
    });
    srv.listen(opts.port ?? ADDON_PORT, '127.0.0.1', () => {
      server = srv;
      currentPort = (srv.address() as AddressInfo).port;
      resolve({ port: currentPort, token: currentToken, tokenGenerated: generated });
    });
  });
}

/** 停通道并终结全部会话（测试收尾 / 进程退出前调用） */
export function stop(): Promise<void> {
  for (const s of Array.from(sessions.values())) {
    disposeSession(s, '通道停止');
  }
  sessions.clear();
  sessionById.clear();
  if (!server) return Promise.resolve();
  const srv = server;
  server = null;
  currentPort = 0;
  return new Promise((resolve) => srv.close(() => resolve()));
}

export interface AddonSessionStatus {
  connected: boolean;
  sessionId?: string;
  addonVersion?: string;
  connectedAt?: string;
  pending?: boolean;
  inflight?: boolean;
}

export interface AddonChannelStatus {
  transport: 'long-poll';
  port: number;
  started: boolean;
  sessions: Record<AddonType, AddonSessionStatus>;
}

/** 三型各自的连接状态（给 B6 的 wps_backend_status 用；本轮只导出数据，不注册工具） */
export function status(): AddonChannelStatus {
  const one = (t: AddonType): AddonSessionStatus => {
    const s = sessions.get(t);
    if (!s) return { connected: false };
    return {
      connected: true,
      sessionId: s.sessionId,
      addonVersion: s.addonVersion,
      connectedAt: s.connectedAt,
      pending: s.pending !== null,
      inflight: s.inflight !== null,
    };
  };
  return {
    transport: 'long-poll',
    port: currentPort,
    started: server !== null,
    sessions: { wps: one('wps'), et: one('et'), wpp: one('wpp') },
  };
}

/** 在途单任务位的原始计数（给 status 之外的诊断用，测试也用它断言"只送达一次"） */
export function sessionCounters(type: AddonType): { pending: number; inflight: number } {
  const s = sessions.get(type);
  return { pending: s && s.pending ? 1 : 0, inflight: s && s.inflight ? 1 : 0 };
}

/**
 * 加载项通道的动作入口（B0 转发点签名：wps-client.ts:280-283）。
 * 全程不排队：轮询池单任务位，顶替/在途忙一律判"从未送出"。
 */
export function call(
  action: string,
  params: Record<string, unknown> = {},
  app: string = 'common',
  instance: string = 'user'
): Promise<unknown> {
  void instance; // v1 忽略：Linux 单实例（S1 R15），多目标由动作层按名钉文档承接
  const type = sessionTypeOf(String(app));
  const s = sessions.get(type);
  if (!s) {
    return Promise.reject(
      new AddonUnavailableError(
        `WPS 加载项通道：没有已连接的 ${type} 型加载项会话（动作=${action}，app=${app}）。` +
          '请求从未送出，可安全重试。若持续出现：确认对应加载项已从 WPS「工具 → WPS 加载项」打开，' +
          '且服务端口/令牌一致（WPS_ADDON_PORT / WPS_ADDON_TOKEN）'
      )
    );
  }
  if (s.inflight) {
    return Promise.reject(
      new AddonUnavailableError(
        `WPS 加载项通道：${type} 型会话上一动作（${s.inflight.action}）仍在途，单任务位不排队。` +
          `动作 ${action} 从未送出，可安全重试`
      )
    );
  }
  if (s.pending) {
    const old = s.pending;
    s.pending = null;
    disposeTask(
      old,
      new AddonUnavailableError(
        `WPS 加载项通道：任务 ${old.action} 被更新的调用顶替（单任务位），从未送达加载项，可安全重试`
      )
    );
  }
  const task: Task = {
    id: uuidv4(),
    action,
    params,
    app: String(app),
    createdAt: Date.now(),
    timer: null as unknown as NodeJS.Timeout,
    resolve: () => undefined,
    reject: () => undefined,
  };
  task.timer = setTimeout(() => onTaskTimeout(task, s), callTimeoutMs());
  s.pending = task;
  // 已有挂起中的 /poll：立即随它送达（任务"已送出"的边界 = 响应写出）
  if (s.pollWait) deliverPending(s);
  return new Promise<unknown>((resolve, reject) => {
    task.resolve = resolve;
    task.reject = reject;
  });
}
