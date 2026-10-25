/**
 * Input: 错误信息与错误码
 * Output: 标准化错误对象
 * Pos: MCP 错误处理模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 错误处理工具
 * 报错见了这个模块都得绕道走，艹
 */

import { log } from './logger';

/**
 * 错误码枚举 - 各种SB错误都有对应的码
 */
export enum ErrorCode {
  // 通用错误 1xxx
  UNKNOWN = 1000,
  INVALID_PARAMS = 1001,
  TIMEOUT = 1002,
  INTERNAL_ERROR = 1003,

  // WPS连接错误 2xxx
  WPS_CONNECTION_FAILED = 2001,
  WPS_NOT_RUNNING = 2002,
  WPS_API_ERROR = 2003,
  WPS_TIMEOUT = 2004,

  // Tool错误 3xxx
  TOOL_NOT_FOUND = 3001,
  TOOL_EXECUTION_FAILED = 3002,
  TOOL_INVALID_ARGS = 3003,
  TOOL_ALREADY_REGISTERED = 3004,

  // MCP协议错误 4xxx
  MCP_INVALID_REQUEST = 4001,
  MCP_METHOD_NOT_FOUND = 4002,
  MCP_PARSE_ERROR = 4003,
}

/**
 * 稳定的 WPS 错误类型标识 - 给调用方分辨"到底哪一层坏了"。
 * 底层错误五花八门（COM / 桥 / 守护进程 / 超时 / 协议），这里收敛成一组固定值。
 */
export enum WpsErrorType {
  /** 桥起不来 / 命名管道连不上 / 桥进程未就绪 */
  BRIDGE_UNAVAILABLE = 'bridge_unavailable',
  /** 常驻守护进程不可用 / 已关闭 / 已退出 */
  DAEMON_UNAVAILABLE = 'daemon_unavailable',
  /** 动作超时（等太久，连接被断开） */
  TIMEOUT = 'timeout',
  /** WPS 没打开 / 类未注册（MK_E_UNAVAILABLE） */
  APP_NOT_RUNNING = 'app_not_running',
  /** 没有活动文档 / 目标块不存在 */
  DOCUMENT_NOT_FOUND = 'document_not_found',
  /** WPS 不支持该接口 / 未知动作 / 取值超范围 */
  UNSUPPORTED_OPERATION = 'unsupported_operation',
  /** 入参不合法 / 缺必填参数 */
  INVALID_PARAMS = 'invalid_params',
  /** 队列已满 / 排队超时 —— 请求从未送出，可安全重试 */
  BUSY = 'busy',
  /** 其余 COM 错误（含 RPC_E_CALL_REJECTED 等） */
  COM_ERROR = 'com_error',
  /** 输出不是合法 JSON / PowerShell 异常退出码 */
  PROTOCOL_ERROR = 'protocol_error',
  /** 调用方主动取消（排队中出队，或只读动作放弃等待）—— 不是 WPS 出错 */
  CANCELLED = 'cancelled',
  /** 无法归类 */
  UNKNOWN = 'unknown',
}

/**
 * 把底层错误信息归类成稳定的 WpsErrorType。
 * 纯函数：只看字符串，不碰 IO，规则按"越具体越靠前"排列。
 */
export function classifyWpsError(message: string): WpsErrorType {
  const msg = typeof message === 'string' ? message : String(message ?? '');
  if (!msg.trim()) {
    return WpsErrorType.UNKNOWN;
  }
  const lower = msg.toLowerCase();

  // 0) 取消必须排第一：CancelError 的文案里含"取消"也可能撞上别的规则
  if (msg.includes('被取消') || msg.includes('放弃等待') || lower.includes('cancelled by client')) {
    return WpsErrorType.CANCELLED;
  }

  // 1) RPC_E_CALL_REJECTED / 0x80010001 —— COM 调用被拒，明确不是超时，别被 "call" 误导
  if (
    msg.includes('0x80010001') ||
    msg.includes('RPC_E_CALL_REJECTED') ||
    lower.includes('call was rejected')
  ) {
    return WpsErrorType.COM_ERROR;
  }

  // 2) MK_E_UNAVAILABLE / 0x800401E3 —— WPS 实例没在跑 / 类未注册
  if (msg.includes('0x800401E3') || msg.includes('MK_E_UNAVAILABLE')) {
    return WpsErrorType.APP_NOT_RUNNING;
  }

  // 3) 队列满 / 排队超时 —— 必须排在"超时"前面：排队超时的文案里也含"超时"，
  //    但语义完全不同（请求从未送出、可安全重试），不能混进 TIMEOUT。
  if (
    msg.includes('队列已满') ||
    msg.includes('排队超时') ||
    lower.includes('queue is full') ||
    lower.includes('queue busy')
  ) {
    return WpsErrorType.BUSY;
  }

  // 4) 超时：WPS 动作超时 / timed out
  if (msg.includes('超时') || lower.includes('timed out') || lower.includes('timeout')) {
    return WpsErrorType.TIMEOUT;
  }

  // 5) 桥起不来：桥进程未就绪 / 管道连不上
  if (
    msg.includes('未就绪') ||
    msg.includes('桥进程') ||
    msg.includes('桥连接不可写') ||
    msg.includes('写入桥失败') ||
    (lower.includes('bridge') && (lower.includes('not ready') || lower.includes('unavailable')))
  ) {
    return WpsErrorType.BRIDGE_UNAVAILABLE;
  }

  // 6) 守护进程 / 桥已关闭
  if (
    msg.includes('守护进程已关闭') ||
    msg.includes('守护进程已退出') ||
    msg.includes('桥连接已关闭') ||
    msg.includes('桥客户端已关闭') ||
    lower.includes('daemon unavailable')
  ) {
    return WpsErrorType.DAEMON_UNAVAILABLE;
  }

  // 7) WPS 没开
  if (
    msg.includes('not running') ||
    msg.includes('未运行') ||
    msg.includes('未启动') ||
    msg.includes('未打开')
  ) {
    return WpsErrorType.APP_NOT_RUNNING;
  }

  // 8) 没有活动文档 / 目标块不存在
  if (
    msg.includes('No active document') ||
    msg.includes('没有活动文档') ||
    msg.includes('无活动文档') ||
    msg.includes('块未找到') ||
    msg.includes('块不存在') ||
    msg.includes('文档未找到') ||
    (msg.includes('未找到') && (lower.includes('doc') || msg.includes('块') || msg.includes('文档')))
  ) {
    return WpsErrorType.DOCUMENT_NOT_FOUND;
  }

  // 9) WPS 不支持该接口 / 未知动作 / 取值超范围
  if (
    msg.includes('Unknown action') ||
    msg.includes('值不在预期的范围内') ||
    msg.includes('不支持') ||
    msg.includes('未实现') ||
    lower.includes('not implemented') ||
    lower.includes('unsupported')
  ) {
    return WpsErrorType.UNSUPPORTED_OPERATION;
  }

  // 10) 入参问题
  if (
    msg.includes('Missing required parameter') ||
    msg.includes('参数错误') ||
    msg.includes('参数无效') ||
    msg.includes('缺少必填参数') ||
    lower.includes('invalid param')
  ) {
    return WpsErrorType.INVALID_PARAMS;
  }

  // 11) 协议层：输出不是合法 JSON / PowerShell 异常退出码
  if (
    msg.includes('输出解析失败') ||
    msg.includes('非有效JSON') ||
    msg.includes('非有效 JSON') ||
    msg.includes('PowerShell 退出码') ||
    (lower.includes('powershell') && msg.includes('退出码')) ||
    (msg.includes('解析失败') && lower.includes('json'))
  ) {
    return WpsErrorType.PROTOCOL_ERROR;
  }

  // 10.5) 缺参 / 参数不合法
  //   实测本仓自己会产出这些消息：PS 侧 "Missing outputPath"、"Missing range"、
  //   "items 为空"，它们都该归 invalid_params 而不是掉进下面的兜底。
  if (
    /^\s*missing\b/i.test(msg) ||
    msg.includes('缺少') ||
    msg.includes('为空') ||
    (msg.includes('参数') && (msg.includes('无效') || msg.includes('不合法') || msg.includes('错误')))
  ) {
    return WpsErrorType.INVALID_PARAMS;
  }

  // 12) 其余 COM 错误：HRESULT / 异常来自 / 中文兜底（带中文前缀的错误不应崩）
  if (
    msg.includes('HRESULT') ||
    msg.includes('异常来自') ||
    msg.includes('不能对 Null 值表达式调用方法') ||
    lower.includes('0x8') ||
    lower.includes('callee') ||
    /[\u4e00-\u9fff]/.test(msg)
  ) {
    return WpsErrorType.COM_ERROR;
  }

  return WpsErrorType.UNKNOWN;
}

/**
 * 自定义错误基类 - 所有错误的祖宗
 */
/**
 * 这类错误可以【安全重试】—— 请求从未真正送进 WPS（或只做了只读探测），
 * 重试不会造成"执行两次"。
 * 刻意【不含】 TIMEOUT：动作超时意味着请求很可能已经执行了一半，自动重放会重复执行。
 */
const RETRYABLE_TYPES = new Set<WpsErrorType>([
  WpsErrorType.BUSY,
  WpsErrorType.BRIDGE_UNAVAILABLE,
  WpsErrorType.DAEMON_UNAVAILABLE,
  WpsErrorType.APP_NOT_RUNNING,
]);

/** 该错误类型能否安全重试（请求未执行，重试不会重复生效） */
export function isRetryableWpsError(type: WpsErrorType): boolean {
  return RETRYABLE_TYPES.has(type);
}

/**
 * 每种错误类型的下一步提示 —— 光说"勿重试"模型不知道该干什么，
 * 这里补一句人话（为什么不能重放 / 要先做什么）。
 */
const RETRY_HINT: Partial<Record<WpsErrorType, string>> = {
  [WpsErrorType.TIMEOUT]: '可能已执行一半',
  [WpsErrorType.COM_ERROR]: '可能已执行一半',
  [WpsErrorType.INVALID_PARAMS]: '先改参数',
  [WpsErrorType.DOCUMENT_NOT_FOUND]: '先打开文档',
  [WpsErrorType.APP_NOT_RUNNING]: '先启动WPS',
  [WpsErrorType.BRIDGE_UNAVAILABLE]: '先拉起桥',
  [WpsErrorType.DAEMON_UNAVAILABLE]: '先拉起桥',
  [WpsErrorType.BUSY]: '稍后退避',
  [WpsErrorType.UNSUPPORTED_OPERATION]: '本机不支持该接口',
  [WpsErrorType.CANCELLED]: '调用方主动取消，未执行',
};

/**
 * 把错误类型渲染成给模型看的短标签，如 `[busy|可重试·稍后退避]`。
 * 前半段是"重放安全吗"，后半段是"下一步先做什么"。
 * 只做展示，不改 isError 语义。
 */
export function formatWpsErrorTag(
  type: WpsErrorType | string,
  retryable?: boolean
): string {
  const t = String(type ?? WpsErrorType.UNKNOWN);
  const canRetry =
    typeof retryable === 'boolean' ? retryable : isRetryableWpsError(t as WpsErrorType);
  const hint = RETRY_HINT[t as WpsErrorType];
  return `[${t}|${canRetry ? '可重试' : '勿重试'}${hint ? `·${hint}` : ''}]`;
}

export class McpError extends Error {
  public readonly code: ErrorCode;
  public readonly details?: Record<string, unknown>;
  public readonly timestamp: Date;
  /** 稳定的 WPS 错误类型标识（由 message 自动分类，可显式覆盖） */
  public readonly wpsType: WpsErrorType;

  constructor(
    message: string,
    code: ErrorCode = ErrorCode.UNKNOWN,
    details?: Record<string, unknown>,
    wpsType?: WpsErrorType
  ) {
    super(message);
    this.name = 'McpError';
    this.code = code;
    this.details = details;
    this.timestamp = new Date();
    this.wpsType = wpsType ?? classifyWpsError(message);

    // 保持正确的原型链
    Object.setPrototypeOf(this, McpError.prototype);
  }

  /** 能否安全重试。调用方据此决定是退避重试还是直接放弃。 */
  get retryable(): boolean {
    return isRetryableWpsError(this.wpsType);
  }

  /**
   * 转换为JSON格式 - 方便传输
   */
  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      message: this.message,
      code: this.code,
      wpsType: this.wpsType,
      retryable: this.retryable,
      details: this.details,
      timestamp: this.timestamp.toISOString(),
      stack: this.stack,
    };
  }
}

/**
 * WPS连接错误 - 连不上WPS就用这个骂
 */
export class WpsConnectionError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.WPS_CONNECTION_FAILED, details);
    this.name = 'WpsConnectionError';
    Object.setPrototypeOf(this, WpsConnectionError.prototype);
  }
}

/**
 * WPS API错误 - WPS API返回的SB错误
 */
export class WpsApiError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.WPS_API_ERROR, details);
    this.name = 'WpsApiError';
    Object.setPrototypeOf(this, WpsApiError.prototype);
  }
}

/**
 * Tool不存在错误 - 找不到Tool就骂这个
 */
export class ToolNotFoundError extends McpError {
  constructor(toolName: string) {
    super(`Tool not found: ${toolName}`, ErrorCode.TOOL_NOT_FOUND, { toolName });
    this.name = 'ToolNotFoundError';
    Object.setPrototypeOf(this, ToolNotFoundError.prototype);
  }
}

/**
 * Tool执行错误 - Tool跑出问题了
 */
export class ToolExecutionError extends McpError {
  constructor(toolName: string, originalError: Error, details?: Record<string, unknown>) {
    super(
      `Tool execution failed: ${toolName} - ${originalError.message}`,
      ErrorCode.TOOL_EXECUTION_FAILED,
      { toolName, originalError: originalError.message, ...details }
    );
    this.name = 'ToolExecutionError';
    Object.setPrototypeOf(this, ToolExecutionError.prototype);
  }
}

/**
 * 参数验证错误 - 参数传错了就骂这个
 */
export class InvalidParamsError extends McpError {
  constructor(message: string, details?: Record<string, unknown>) {
    super(message, ErrorCode.INVALID_PARAMS, details);
    this.name = 'InvalidParamsError';
    Object.setPrototypeOf(this, InvalidParamsError.prototype);
  }
}

/**
 * 超时错误 - 等太久了，不等了
 */
export class TimeoutError extends McpError {
  constructor(operation: string, timeoutMs: number) {
    super(
      `Operation timed out: ${operation} (${timeoutMs}ms)`,
      ErrorCode.TIMEOUT,
      { operation, timeoutMs }
    );
    this.name = 'TimeoutError';
    Object.setPrototypeOf(this, TimeoutError.prototype);
  }
}

/**
 * 错误处理工具函数
 */
export const errorUtils = {
  /**
   * 包装错误 - 把各种SB错误统一包装成McpError
   */
  wrap(error: unknown, defaultMessage = 'An error occurred'): McpError {
    if (error instanceof McpError) {
      return error;
    }

    if (error instanceof Error) {
      return new McpError(error.message, ErrorCode.INTERNAL_ERROR, {
        originalName: error.name,
        stack: error.stack,
      });
    }

    return new McpError(
      typeof error === 'string' ? error : defaultMessage,
      ErrorCode.UNKNOWN,
      { originalError: error }
    );
  },

  /**
   * 记录并重新抛出错误
   */
  logAndThrow(error: unknown, context?: string): never {
    const mcpError = errorUtils.wrap(error);
    log.error(context ? `${context}: ${mcpError.message}` : mcpError.message, mcpError);
    throw mcpError;
  },

  /**
   * 安全执行函数 - 出错了返回默认值，不会炸
   */
  async safeExecute<T>(
    fn: () => Promise<T>,
    defaultValue: T,
    context?: string
  ): Promise<T> {
    try {
      return await fn();
    } catch (error) {
      const mcpError = errorUtils.wrap(error);
      log.error(
        context ? `${context}: ${mcpError.message}` : mcpError.message,
        mcpError
      );
      return defaultValue;
    }
  },

  /**
   * 判断是否是特定类型的错误
   */
  isErrorCode(error: unknown, code: ErrorCode): boolean {
    return error instanceof McpError && error.code === code;
  },
};

/**
 * 格式化错误信息给用户看
 */
export const formatErrorForUser = (error: McpError): string => {
  switch (error.code) {
    case ErrorCode.WPS_CONNECTION_FAILED:
    case ErrorCode.WPS_NOT_RUNNING:
      return '无法连接到WPS Office，请确保WPS已启动并且加载项已安装';
    case ErrorCode.WPS_TIMEOUT:
    case ErrorCode.TIMEOUT:
      return `操作超时，请稍后重试${error.details?.operation ? '（' + error.details.operation + '）' : ''}`;
    case ErrorCode.WPS_API_ERROR:
      return `WPS接口调用失败: ${error.message}`;
    case ErrorCode.TOOL_NOT_FOUND:
      return `找不到指定的工具: ${error.details?.toolName}`;
    case ErrorCode.TOOL_EXECUTION_FAILED:
      return `工具执行失败${error.details?.toolName ? '（' + error.details.toolName + '）' : ''}: ${error.message}`;
    case ErrorCode.TOOL_INVALID_ARGS:
    case ErrorCode.INVALID_PARAMS:
      return `参数错误: ${error.message}`;
    case ErrorCode.TOOL_ALREADY_REGISTERED:
      return `工具已注册，请勿重复注册: ${error.details?.toolName}`;
    case ErrorCode.MCP_INVALID_REQUEST:
      return '请求格式无效，请检查请求参数';
    case ErrorCode.MCP_METHOD_NOT_FOUND:
      return '请求的方法不存在，请检查方法名称';
    case ErrorCode.MCP_PARSE_ERROR:
      return '消息解析失败，请检查数据格式';
    case ErrorCode.INTERNAL_ERROR:
      return `内部错误: ${error.message}`;
    default:
      return `操作失败: ${error.message}`;
  }
};

export default McpError;
