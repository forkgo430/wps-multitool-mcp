/**
 * Input: 进度回调函数（由 mcp-server.ts 的 CallTool handler 注入）
 * Output: 无返回值——调用 reportProgress 侧发进度通知给 MCP 客户端
 * Pos: 长时间操作（batch / 导出 / 大文档处理）在执行过程中报告进度的通道。
 * 加上进度通知后，客户端能显示"正在做第 3/5 步"，用户体验不再是一屏空白。
 */
import { AsyncLocalStorage } from 'async_hooks';

export type ProgressReporter = (progress: number, total: number, message?: string) => void;

const storage = new AsyncLocalStorage<ProgressReporter>();

/** 在 CallTool handler 里包住工具执行——把进度回调挂进上下文 */
export function runWithProgress<T>(reporter: ProgressReporter, fn: () => Promise<T>): Promise<T> {
  return storage.run(reporter, fn);
}

/** 工具执行链深处调这个报进度。客户端没传 progressToken 时是 no-op。 */
export function reportProgress(progress: number, total: number, message?: string): void {
  const reporter = storage.getStore();
  if (reporter) reporter(progress, total, message);
}
