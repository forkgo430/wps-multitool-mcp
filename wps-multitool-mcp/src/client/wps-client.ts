/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * WPS通信客户端 - 仅 Windows
 * 通过 PowerShell 调用 WPS 的 COM 接口；优先走常驻桥，桥起不来退回直跑。
 */

import { spawn } from 'child_process';
import * as path from 'path';
import {
  WpsEndpointConfig,
  WpsApiRequest,
  WpsApiResponse,
  WpsAppType,
  WpsClientStatus,
  DocumentInfo,
  WorkbookInfo,
  PresentationInfo,
} from '../types/wps';
import { log, logRequest, logResponse } from '../utils/logger';
import { errorUtils } from '../utils/error';
import { psDaemon, DaemonUnavailableError } from './ps-daemon';
import { AppKey, resolveAppKey } from './action-routes';
import { applyWordTarget } from '../tools/word/target';

/**
 * 实例级动作：它们的语义就是"对哪个实例做"，与具体应用无关。
 *
 * 为什么单独列：这些动作在路由表里落 common（它们天然跨应用），
 * 但 target 必须被接受；而且在 sandbox 下必须改道到 word 子进程 ——
 * 沙箱句柄每进程一份，common 自己去 New-Object 会开出第二个沙箱实例。
 */
const INSTANCE_SCOPED_ACTIONS = new Set(['quitWps', 'showWps', 'getSelectedText', 'setSelectedText']);
import { currentSignal } from '../server/cancel-context';
import { currentTarget } from '../server/target-context';
import { CancelError } from '../bridge/cancellable';
import { ToolWarning } from '../types/tools';

/**
 * 从 data 里挑出部分成功提示。
 *
 * 容错是有意的：PS 侧是 6400 行的手写脚本，它给 warnings 的形状不保证统一
 * （可能是对象、可能是纯字符串）。解析失败就当没提示 —— 宁可少报，
 * 也不能让一条脏数据把整次调用的结果判成失败。
 */
export function extractWarnings(data: unknown): ToolWarning[] {
  if (!data || typeof data !== 'object') return [];
  const raw = (data as Record<string, unknown>).warnings;
  if (!Array.isArray(raw)) return [];
  const out: ToolWarning[] = [];
  for (const item of raw) {
    if (typeof item === 'string' && item.trim()) {
      out.push({ message: item.trim() });
    } else if (item && typeof item === 'object') {
      const o = item as Record<string, unknown>;
      const message = typeof o.message === 'string' ? o.message.trim() : '';
      if (!message) continue;
      const step = typeof o.step === 'string' ? o.step : undefined;
      out.push(step ? { step, message } : { message });
    }
  }
  return out;
}

// 仅 Windows：优先走常驻桥（省掉每次 spawn powershell + 解析 5000 行脚本）。
// 设 WPS_USE_DAEMON=0 可强制退回直跑模式。
const USE_DAEMON = process.env.WPS_USE_DAEMON !== '0';
// 守护进程一旦起不来就本次全程直跑，不再每次调用都白等 20 秒握手
let daemonBroken = false;

/**
 * 通道种类（第 10 轮 T4 引入）
 *
 * 本轮【只拆结构、不加实现】：'addon' 是给 Linux 侧留的插入点，
 * 现在没有任何地方会返回它。等真的要做 Linux 后端时，
 * 只要往 pickTransport 里加一个判断、往 runTransport 里加一个 case，
 * Windows 那两个分支一个字都不用碰（对应计划的约束 C1）。
 */
type TransportKind = 'daemon' | 'direct' | 'addon';

/**
 * 选通道 —— 一次决定，结果缓存（第 10 轮 T4）
 *
 * 口径与拆分前完全一致：USE_DAEMON 为假、或闩锁已落下 → 'direct'；否则 'daemon'。
 * "缓存"指的是这个判断不再每次重算（原来每次调用都重跑一遍 `USE_DAEMON && !daemonBroken`）。
 * 两者等价，是因为 daemonBroken 一旦为 true 就再也不会回 false（见 markDaemonBroken）。
 */
let cachedTransport: TransportKind | null = null;

function pickTransport(): TransportKind {
  if (cachedTransport) return cachedTransport;
  // ★ Linux 侧的插入点就在这一行下面：
  //   if (process.platform === 'linux' && USE_ADDON) cachedTransport = 'addon';
  //   本轮【故意不加】—— 没有 Linux 环境，写了就是没有证据的代码。
  cachedTransport = USE_DAEMON && !daemonBroken ? 'daemon' : 'direct';
  return cachedTransport;
}

/** 闩锁：桥起不来 → 本进程后续一律直跑，并把缓存的通道选择一起翻过去 */
function markDaemonBroken(reason: string): void {
  daemonBroken = true;
  cachedTransport = 'direct';
  log.warn('桥不可用，本进程后续改用直跑模式', { reason });
}

// PowerShell脚本路径 (Windows)
const PS_SCRIPT_PATH = path.join(__dirname, '../../scripts/wps-com.ps1');

// PowerShell 默认超时（毫秒）
const PS_TIMEOUT = 30000;

// ==================== 目标文稿锁定（避免多文档打开时 Active* 漂移）====================
// 通过各 app 的 set*Target 设置后，所有该 app 的调用自动注入目标名，
// 由 wps-com.ps1 的 Resolve-Doc / Resolve-Book / Get-TargetPres 精确定位；
// 单次调用显式传 documentName / workbookName / presentationName 可覆盖。
//
// 为什么三个都要有：ActiveDocument 是【弱引用】，实测会在 Documents.Count > 0 时
// 返回一个取不到内容的空壳（见 wps_common_repair 的诊断）。锁定目标名是唯一
// 不依赖 UI 焦点、也不受空壳影响的定位方式。
//
// 用法：
//   wps_word_switch_document / wps_excel_switch_workbook 已经在做这件事（它们会调 set*Target）
//   这里只是让"锁定"这件事对所有工具生效，而不只是切换类工具。
let pptTargetName: string | undefined;
let wordTargetName: string | undefined;
let excelTargetName: string | undefined;

export function setPptTarget(name?: string): void {
  pptTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getPptTarget(): string | undefined {
  return pptTargetName;
}

export function setWordTarget(name?: string): void {
  wordTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getWordTarget(): string | undefined {
  return wordTargetName;
}

export function setExcelTarget(name?: string): void {
  excelTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getExcelTarget(): string | undefined {
  return excelTargetName;
}

/**
 * 执行PowerShell命令 (Windows)
 */
async function execPowerShell(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  return new Promise((resolve, reject) => {
    // JSON参数通过spawn args数组传递，Node自动处理Windows引号转义
    const paramsJson = JSON.stringify(params);
    const args = [
      // -NoProfile/-NonInteractive：直跑是回退路径，不该再花时间读用户 profile
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', PS_SCRIPT_PATH,
      '-Action', action,
      '-Params', paramsJson
    ];

    log.debug('Executing PowerShell', { action, params });

    const ps = spawn('powershell', args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';
    let killed = false;

    // 超时保护：防止PowerShell进程挂起，超时时主动 SIGTERM 并记录 PID
    const timeoutHandle = setTimeout(() => {
      killed = true;
      log.warn('PowerShell timeout, killing process', { pid: ps.pid, action, timeoutMs: PS_TIMEOUT });
      ps.kill('SIGTERM');
      reject(new Error(`PowerShell 执行超时（${PS_TIMEOUT}ms）: ${action} (PID: ${ps.pid})`));
    }, PS_TIMEOUT);

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      clearTimeout(timeoutHandle);
      if (killed) return; // 已超时处理，忽略后续事件

      if (code !== 0) {
        // 非零退出码：分别处理 stderr 有内容 / 为空两种路径，避免空 stderr 误入 JSON.parse
        if (stderr) {
          log.error('PowerShell error', { stderr, code, pid: ps.pid, action });
          reject(new Error(stderr));
        } else {
          log.error('PowerShell exited with non-zero code', { code, stdout: stdout.substring(0, 200), pid: ps.pid, action });
          reject(new Error(`PowerShell 退出码 ${code}: ${stdout.substring(0, 200) || '(无输出)'}`));
        }
        return;
      }

      try {
        const result = JSON.parse(stdout.trim());
        resolve(result);
      } catch (_e) {
        log.error('Failed to parse PowerShell output', { stdout: stdout.substring(0, 200), pid: ps.pid, action });
        reject(new Error(`PowerShell 输出解析失败（非有效JSON）: ${stdout.substring(0, 200)}`));
      }
    });

    ps.on('error', (err) => {
      clearTimeout(timeoutHandle);
      if (killed) return;
      log.error('PowerShell spawn error', { error: err.message, pid: ps.pid, action });
      reject(new Error(`无法启动 PowerShell 进程: ${err.message}`));
    });
  });
}

/**
 * 跑通道（第 10 轮 T4）—— 只负责"把请求送出去"，不做任何选择/降级判断。
 *
 * 'addon' 分支本轮**不实现**：真要做的时候在这里加一个 case 即可，
 * 上面 pickTransport 决定什么时候走它。故意留空而不是抛错占位，
 * 是因为 TypeScript 的穷尽检查会逼着写一个跑不通的分支，那比没有更危险。
 */
async function runTransport(
  kind: TransportKind,
  action: string,
  params: Record<string, unknown>,
  app: AppKey,
  instance: string = 'user'
): Promise<unknown> {
  switch (kind) {
    case 'daemon':
      // 取消信号从 MCP 的 notifications/cancelled 一路传到这里（见 server/cancel-context）
      return psDaemon.call(action, params, app, { signal: currentSignal(), instance });
    case 'direct':
      // 直跑模式每次都是一次性 powershell：没有常驻句柄，沙箱实例根本留不住
      // （每次 New-Object 都会多开一个 WPS 进程，且没人回收）。这里明确拒绝，
      // 而不是"看起来能跑"地泄漏一堆进程。
      if (instance !== 'user') {
        throw new Error(
          `[勿重试·先改配置] target.instance=${instance} 需要常驻桥（一个实例一个子进程）。` +
            '当前是直跑模式（WPS_USE_DAEMON=0 或桥起不来），沙箱实例无法保持。' +
            '请去掉 target.instance（用 user），或恢复常驻桥后重试。'
        );
      }
      return execPowerShell(action, params);
    default:
      throw new Error(`未实现的传输通道：${kind}`);
  }
}

/**
 * 统一执行接口（仅 Windows）
 *   常驻桥优先；桥起不来则退回"一调用一进程"的直跑方式
 *
 * 第 10 轮 T4 拆成两层：pickTransport（选一次，缓存）/ runTransport（只管跑）。
 * 【Windows 行为逐字未变】：延迟、降级时机、daemonBroken 闩锁语义、超时分类，
 * 都在下面这几行里，与拆分前逐条对应。
 */
async function execWpsAction(
  action: string,
  params: Record<string, unknown> = {},
  app: AppKey = 'common',
  instance: string = 'user'
): Promise<unknown> {
  const kind = pickTransport();
  if (kind === 'direct') {
    return runTransport('direct', action, params, app, instance);
  }

  try {
    return await runTransport(kind, action, params, app, instance);
  } catch (err) {
    // ↓↓↓ 降级口径，全仓只有这一份，拆分后仍只此一处 ↓↓↓
    // 只有"请求压根没送出去"（桥起不来）才允许退回直跑。
    // 超时那类错误请求可能已经在 WPS 里执行了一半，自动重放会变成执行两次 —— 一律如实抛出。
    if (err instanceof DaemonUnavailableError) {
      markDaemonBroken(err.message);
    } else {
      throw err;
    }
  }

  return runTransport('direct', action, params, app);
}

/**
 * WPS客户端类
 * 仅 Windows：PowerShell + COM 桥接
 */
export class WpsClient {
  private status: WpsClientStatus;

  constructor(_config?: Partial<WpsEndpointConfig>) {
    this.status = { connected: false };
    log.info('WPS Client initialized', { method: 'COM bridge (Windows only)' });
  }

  /**
   * @param app 显式指定走哪一路子进程；不传则按动作名自动解析（见 action-routes）
   */
  async invokeAction<T = unknown>(
    action: string,
    params: Record<string, unknown> = {},
    app?: AppKey,
    opts: { instance?: string } = {}
  ): Promise<WpsApiResponse<T>> {
    const startTime = Date.now();
    logRequest(action, params);

    try {
      const result = (await execWpsAction(
        action,
        params,
        app ?? resolveAppKey(action),
        opts.instance ?? 'user'
      )) as WpsApiResponse<T>;
      const duration = Date.now() - startTime;
      logResponse(action, result.success, duration);

      if (result.success) {
        this.status.connected = true;
        this.status.lastHeartbeat = new Date();
      }

      // PS 侧把部分成功提示放在 data.warnings（约定位置），这里提到顶层，
      // handler 不用每次去 data 里翻 —— 200 多个 handler 没法逐个记得翻。
      const warnings = extractWarnings(result.data);
      if (warnings.length) result.warnings = warnings;

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logResponse(action, false, duration);
      // 取消不是 WPS 出错，别被 wrap 成"WPS 调用失败"；原样往上抛
      if (error instanceof CancelError) throw error;
      this.status.connected = false;
      throw errorUtils.wrap(error, `WPS 调用失败（COM 桥）: ${action}`);
    }
  }

  /**
   * 兼容旧API
   */
  async callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>> {
    const actionMap: Record<string, string> = {
      // Excel 旧API
      'workbook.getActive': 'getActiveWorkbook',
      'cell.getValue': 'getCellValue',
      'cell.setValue': 'setCellValue',
      'range.getData': 'getRangeData',
      'range.setData': 'setRangeData',
      // Word 旧API
      'document.getActive': 'getActiveDocument',
      'document.getText': 'getDocumentText',
      'document.insertText': 'insertText',
      // PPT 旧API
      'presentation.getActive': 'getActivePresentation',
      'presentation.addSlide': 'addSlide',
      // 通用旧API
      'file.save': 'save',
      'file.saveAs': 'saveAs',
      'file.open': 'openFile',
      'ping': 'ping',
    };
    const action = actionMap[request.method] || request.method;
    return this.invokeAction<T>(action, request.params || {});
  }

  /**
   * 检查WPS连接状态
   */
  async checkConnection(): Promise<boolean> {
    try {
      const result = await this.invokeAction('ping');
      this.status.connected = result.success;
      return result.success;
    } catch {
      this.status.connected = false;
      this.status.error = 'Connection check failed';
      return false;
    }
  }

  /**
   * 获取客户端状态
   */
  getStatus(): WpsClientStatus {
    return { ...this.status };
  }

  // ==================== 表格操作 (WPS表格) ====================

  async getActiveWorkbook(): Promise<WorkbookInfo | null> {
    const response = await this.invokeAction<WorkbookInfo>('getActiveWorkbook');
    return response.success ? response.data || null : null;
  }

  async getRangeData(sheet: string | number, range: string): Promise<unknown[][]> {
    const response = await this.invokeAction<{ data: unknown[][] }>('getRangeData', { sheet, range });
    return response.data?.data || [];
  }

  async setRangeData(sheet: string | number, range: string, data: unknown[][]): Promise<boolean> {
    const response = await this.invokeAction('setRangeData', { sheet, range, data });
    return response.success;
  }

  async setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean> {
    const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
    return response.success;
  }

  // ==================== 文档操作 (WPS文字) ====================

  async getActiveDocument(): Promise<DocumentInfo | null> {
    const response = await this.invokeAction<DocumentInfo>('getActiveDocument');
    return response.success ? response.data || null : null;
  }

  async createDocument(): Promise<boolean> {
    const response = await this.invokeAction('createDocument');
    return response.success;
  }

  async getDocumentText(): Promise<string> {
    const response = await this.invokeAction<{ text: string }>('getDocumentText');
    return response.data?.text || '';
  }

  // ==================== 演示操作 (WPS演示) ====================

  async getActivePresentation(): Promise<PresentationInfo | null> {
    const response = await this.invokeAction<PresentationInfo>('getActivePresentation');
    return response.success ? response.data || null : null;
  }

  async createPresentation(): Promise<boolean> {
    const response = await this.invokeAction('createPresentation');
    return response.success;
  }

  async addSlide(layout?: string): Promise<boolean> {
    const response = await this.invokeAction('addSlide', { layout });
    return response.success;
  }

  // ==================== 通用操作 ====================

  /**
   * 通用方法调用入口。
   * appType 是调用点声明的目标应用，会被翻译成桥的应用键（excel/word/ppt），
   * 从而落到对应的常驻子进程上 —— 这是 per-app 隔离的关键，别把它丢掉。
   */
  async executeMethod<T = unknown>(
    method: string,
    params?: Record<string, unknown>,
    appType?: WpsAppType
  ): Promise<WpsApiResponse<T>> {
    let finalParams: Record<string, unknown> = params || {};
    // 锁定目标文稿后自动注入目标名（已显式传入则不覆盖）。
    // 三个 app 各有自己的参数名：presentationName / documentName / workbookName ——
    // 名字与 wps-com.ps1 里 Resolve-Doc / Resolve-Book / Get-TargetPres 的读取一致。
    if (appType === WpsAppType.PRESENTATION && pptTargetName && finalParams.presentationName === undefined) {
      finalParams = { ...finalParams, presentationName: pptTargetName };
    }
    // 会话级 target 只在【本次调用没有显式 target.doc】时兜底注入。
    // 否则会出现：用户每个调用都写了 target.doc，却被 switch_document 记下的会话文档接管
    // —— 而且完全不报错（真机 A1/A3 就是这么假的）。
    // 只要本次调用【给了 target】（哪怕只有 instance、没有 doc），就不继承会话文档。
    // 为什么：会话 target 是"上一次 switch/create 落在哪个实例"的记忆，它是【跨实例全局】的；
    // 一个发往 sandbox 的调用如果继承了 user 实例的文档名，PS 会在沙箱里找不到那份文档
    // 直接报错（A4 的预热就是这么被卡死的：get_open_documents 都被注入了一个 user 侧文档名）。
    // 注意：target 可能在 **ALS** 里而不在 finalParams 里 —— 67 个 handler 都不转发 args.target。
    // 只看 finalParams 会让守卫永远为 false（第一版就是这么继续注入的）。
    const alsTarget = currentTarget();
    const hasExplicitTarget =
      (finalParams.target !== undefined && finalParams.target !== null) ||
      (alsTarget !== undefined && alsTarget !== null);
    if (
      appType === WpsAppType.WRITER &&
      wordTargetName &&
      finalParams.documentName === undefined &&
      !hasExplicitTarget
    ) {
      finalParams = { ...finalParams, documentName: wordTargetName };
    }
    if (appType === WpsAppType.SPREADSHEET && excelTargetName && finalParams.workbookName === undefined) {
      finalParams = { ...finalParams, workbookName: excelTargetName };
    }

    // ==================== target（第 3 章轮 1：只做 word）====================
    // 三件事在这里一次做完，下游（daemon/bridge/ps1）只认"instance"这一个路由键：
    //   ① target.doc 合流进现成的 documentName（不新造第二个名字，否则两套必然漂）
    //   ② target 形状校验（错形状必须在这里就报，别等到 PS 里被静默忽略）
    //   ③ instance 抽出来走路由，不混进 COM 参数语义
    const resolvedApp = resolveAppKey(method, appType);
    let instance = 'user';
    // target 可能来自两处：本次调用显式传的，或上层工具（callTool 入口用 ALS 兜住的）。
    // 67 个 handler 不会转发 args.target，所以 ALS 那条路是主路径。
    const effectiveTarget = finalParams.target !== undefined ? finalParams.target : currentTarget();
    let appForCall = resolvedApp;
    if (effectiveTarget !== undefined && effectiveTarget !== null) {
      const allowTarget =
        resolvedApp === 'word' || appType === WpsAppType.WRITER || INSTANCE_SCOPED_ACTIONS.has(method);
      if (!allowTarget) {
        // 静默忽略最坏：模型以为钉住了，实际还在改活动文档。
        throw new Error(
          `[勿重试·先改参数] target 目前只支持 word（本调用落在 ${resolvedApp}）。` +
            'excel / ppt 的 target 还没实现，请去掉 target 再试。'
        );
      }
      const applied = applyWordTarget({ ...finalParams, target: effectiveTarget });
      finalParams = applied.params;
      instance = applied.instance;

      // 实例级动作（退出/显示/选区）在路由表里落 common，但**必须和写入走同一个子进程**：
      // 沙箱句柄是"每进程一份"的，common 子进程自己去 New-Object 会开出第二个沙箱实例 ——
      // 于是 quitWps 退掉的不是刚才写文档的那个实例（A6 会假通过）。
      // 只在 instance != user 时改道：user 下 quitWps 仍要退全部应用（旧行为）。
      if (instance !== 'user' && INSTANCE_SCOPED_ACTIONS.has(method)) {
        appForCall = 'word';
      }
    }

    return this.invokeAction<T>(method, finalParams, appForCall, { instance });
  }

  async openFile(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('openFile', { path: filePath });
    return response.success;
  }

  async saveFile(_appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('save');
    return response.success;
  }

  async saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('saveAs', { path: filePath });
    return response.success;
  }
}

// 导出单例
export const wpsClient = new WpsClient();

export default WpsClient;
