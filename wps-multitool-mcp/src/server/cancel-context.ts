/**
 * Input: 一次 tools/call 的取消信号
 * Output: 同一次调用链深处也能拿到它
 * Pos: 取消信号的传递上下文。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么用 AsyncLocalStorage 而不是改函数签名
 * ------------------------------------------
 * 工具 handler 的签名是 `(args) => Promise<ToolCallResult>`，全仓 285 个工具都是这个形状。
 * 要把 signal 传下去就得改 285 处签名 + 每个 client 调用点 —— 收益极低、破坏面极大。
 * 用 AsyncLocalStorage 在【请求入口】放一次，调用链深处任何地方都能取到，
 * 签名一个字都不用动。
 *
 * 语义边界
 * --------
 * 这里只负责"把信号递下去"。该不该真取消、取消到哪一档，由 bridge/cancellable.ts 裁决。
 */

import { AsyncLocalStorage } from 'async_hooks';

export interface CancelContext {
  signal?: AbortSignal;
  /** 工具名，仅用于日志 */
  toolName?: string;
}

const storage = new AsyncLocalStorage<CancelContext>();

/**
 * 在一次调用外层包一下，之后这条链上任意深度都能用 currentSignal() 取到信号。
 */
export function runWithCancel<T>(ctx: CancelContext, fn: () => Promise<T>): Promise<T> {
  return storage.run(ctx, fn);
}

/** 当前调用链上的取消信号（没有就是 undefined，即"不可取消"） */
export function currentSignal(): AbortSignal | undefined {
  return storage.getStore()?.signal;
}

/** 当前工具名（日志用） */
export function currentToolName(): string | undefined {
  return storage.getStore()?.toolName;
}
