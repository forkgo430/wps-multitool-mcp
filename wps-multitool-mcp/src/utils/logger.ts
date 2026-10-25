/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 日志工具
 * 写日志比写代码还规范，报错了至少知道是哪个SB搞出来的
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
  const transports: winston.transport[] = [
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

/**
 * 主Logger - MCP Server专用
 */
export const logger = createLogger('wps-mcp');

/**
 * 创建子Logger - 各模块用自己的
 */
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
 * 请求日志中间件风格的函数
 */
export const logRequest = (
  method: string,
  params?: Record<string, unknown>
): void => {
  logger.info(`[REQUEST] ${method}`, { params });
};

/**
 * 响应日志
 */
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
