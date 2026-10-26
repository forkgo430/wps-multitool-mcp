/**
 * Input: 进度回调函数（由 mcp-server.ts 的 CallTool handler 注入）
 * Output: 无返回值——调用 reportProgress 侧发进度通知给 MCP 客户端
 * Pos: 长时间操作（batch / 导出 / 大文档处理）在执行过程中报告进度的通道。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有这个
 * --------------
 * 批量执行 20 步、导出大 PDF，客户端面对的是黑盒等待。
 * 加上进度通知后，客户端能显示"正在做第 3/5 步"，用户体验不再是一屏空白。
 *
 * 为什么用 AsyncLocalStorage 而不是改工具签名
 * ------------------------------------------
 * 同 cancel-context.ts 的理由：288 个工具一个签名都不改。
 * 服务器在 CallTool handler 里调 runWithProgress 把回调挂进上下文，
 * 工具深处（batch.ts / 导出工具）调 reportProgress 取出来用。
 * 客户端没传 progressToken 时 reporter 是 no-op，零开销。
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
