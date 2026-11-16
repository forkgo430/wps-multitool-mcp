/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 */

import winston from 'winston';
import path from 'path';
import { RUNTIME_ID } from '../runtime-identity';

/**
 * 日志级别枚举
 */
export enum LogLevel {
  ERROR = 'error',
  WARN = 'warn',
  INFO = 'info',
  DEBUG = 'debug',
}

/**
 * 自定义日志格式
 */
const wangFormat = winston.format.printf(({ level, message, timestamp, ...meta }) => {
  const metaStr = Object.keys(meta).length ? ` | ${JSON.stringify(meta)}` : '';
  return `[${timestamp}] [${level.toUpperCase()}] ${message}${metaStr}`;
});

/**
 * 解析布尔环境变量
 */
const isEnvTrue = (value?: string): boolean => {
  if (!value) return false;
  return ['1', 'true', 'yes', 'on'].includes(value.toLowerCase());
};

/**
 * 单文件上限 5 MB、保留 3 份。
 *
 * 为什么必须有：之前 ~/.wps-multitool-mcp/logs 涨到 10.5 MB 且没有任何上限 ——
 * 这个进程是常驻的（MCP server 跟着 IDE 活），一次调用记几行，几个月就能把用户
 * 家目录写出一个几十 MB 的文件。winston 自带 maxsize/maxFiles，不需要引新依赖。
 *
 * 语义（别搞反）：winston 是**写满当前文件就改名**——combined.log 写满 5 MB →
 * 变成 combined1.log，新的写进 combined.log；超过 maxFiles 份的最老那份直接删。
 * 所以"保留 3 份"= 最多在盘上留下 3 个 5 MB 的文件，上限 15 MB。
 */
const LOG_MAX_SIZE = 5 * 1024 * 1024;
const LOG_MAX_FILES = 3;

/**
 * 创建Logger实例
 */
const createLogger = (name: string): winston.Logger => {
  const homeDir = process.env.HOME || process.env.USERPROFILE || require('os').homedir();
  const logDir = path.join(homeDir, `.${RUNTIME_ID}`, 'logs');
  // MCP 默认走 stdio 协议，stdout 必须保持纯净；仅在显式开启时才输出 Console 日志
  const enableConsoleLog = isEnvTrue(process.env.MCP_CONSOLE_LOG);
  // R5（提质合并卡 2026-11-11）：日志目录建不出（坏 HOME / 只读盘）时【降级为无文件日志】
  // ——不能让 winston File transport 在模块加载期裸抛 ENOENT 把整个服务带死。
  // 实测锚：HOME=Q:/nonexistent 时服务必须照常起、tools/list 照常回。
  let fileTransports: winston.transport[] = [];
  try {
    require('fs').mkdirSync(logDir, { recursive: true });
    fileTransports = [
      // 文件输出 - 错误单独记录，出问题好找
      new winston.transports.File({
        filename: path.join(logDir, 'error.log'),
        level: 'error',
        maxsize: LOG_MAX_SIZE,
        maxFiles: LOG_MAX_FILES,
      }),
      // 所有日志
      new winston.transports.File({
        filename: path.join(logDir, 'combined.log'),
        maxsize: LOG_MAX_SIZE,
        maxFiles: LOG_MAX_FILES,
      }),
    ];
  } catch {
    // 降级：无文件日志；console 仍按 MCP_CONSOLE_LOG 走
  }
  const transports: winston.transport[] = [
    ...fileTransports,
  ];

  if (enableConsoleLog) {
    transports.unshift(
      new winston.transports.Console({
        format: winston.format.combine(
          winston.format.colorize({ all: true }),
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          wangFormat
        ),
      })
    );
  }

  return winston.createLogger({
    level: process.env.LOG_LEVEL || 'info',
    format: winston.format.combine(
      winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
      winston.format.errors({ stack: true }),
      wangFormat
    ),
    defaultMeta: { service: name },
    transports,
  });
};

export const logger = createLogger('wps-mcp');

export const createChildLogger = (moduleName: string): winston.Logger => {
  return logger.child({ module: moduleName });
};

/**
 * 快捷日志函数 - 懒人专用
 */
export const log = {
  /**
   * 记录信息日志
   */
  info: (message: string, meta?: Record<string, unknown>): void => {
    logger.info(message, meta);
  },

  /**
   * 记录警告日志
   */
  warn: (message: string, meta?: Record<string, unknown>): void => {
    logger.warn(message, meta);
  },

  /**
   * 记录错误日志 - 艹，又出错了
   */
  error: (message: string, error?: Error | unknown, meta?: Record<string, unknown>): void => {
    if (error instanceof Error) {
      logger.error(message, { ...meta, error: error.message, stack: error.stack });
    } else {
      logger.error(message, { ...meta, error });
    }
  },

  /**
   * 记录调试日志 - 排查问题用的
   */
  debug: (message: string, meta?: Record<string, unknown>): void => {
    logger.debug(message, meta);
  },
};

/**
 * 参数摘要：只记键名与大小，不记值。
 * params 里是单元格数据、文档正文这类用户内容，整份写盘属于隐私泄漏
 * （2026-10-25 外部评审 P1：日志"可能含文档路径"的说法低估了实际风险）。
 * 排障要完整参数时设 WPS_LOG_PARAMS=1。
 */
const summarizeParams = (params?: Record<string, unknown>): Record<string, unknown> | undefined => {
  if (!params) return undefined;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) { out[k] = String(v); continue; }
    if (typeof v === 'string') out[k] = `str(${v.length})`;
    else if (typeof v === 'number' || typeof v === 'boolean') out[k] = v;
    else if (Array.isArray(v)) out[k] = `arr(${v.length})`;
    else out[k] = `obj(${JSON.stringify(v).length})`;
  }
  return out;
};

export const logRequest = (
  method: string,
  params?: Record<string, unknown>
): void => {
  if (isEnvTrue(process.env.WPS_LOG_PARAMS)) {
    logger.info(`[REQUEST] ${method}`, { params });
    return;
  }
  logger.info(`[REQUEST] ${method}`, { params: summarizeParams(params) });
};

export const logResponse = (
  method: string,
  success: boolean,
  duration: number,
  data?: unknown
): void => {
  const level = success ? 'info' : 'error';
  logger.log(level, `[RESPONSE] ${method} - ${success ? 'SUCCESS' : 'FAILED'} (${duration}ms)`, {
    success,
    duration,
    ...(process.env.LOG_LEVEL === 'debug' ? { data } : {}),
  });
};

export default logger;
