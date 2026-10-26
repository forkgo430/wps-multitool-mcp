/**
 * Input: 工具调用参数里的 target（可选）
 * Output: 当前调用链上"要钉住哪一份"的上下文
 * Pos: target 的跨层传递。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么需要这一层
 * ----------------
 * 67 个 word handler 都是这样写的：从 args 里取自己认识的字段，然后**显式**拼一个新对象
 * 交给 wpsClient.executeMethod。它们不会把 args.target 一起转发 —— 这是好事
 * （handler 不该关心路由），但意味着 target 在 handler 边界就丢了。
 *
 * 三种可选做法，选了第三种：
 *   ① 改 67 个 handler 逐个转发 —— 必然漏，而漏了不报错，只表现为"target 没生效"；
 *   ② 让 executeMethod 从全局变量读 —— MCP 的 callTool 是并发可达的，
 *      两个并发调用会互相覆盖，一个任务的 target 会漂到另一个任务上；
 *   ③ AsyncLocalStorage（本文件）—— 跟着各自的调用链走，并发也不串。
 *      auto-snapshot.ts 用同一个理由选了同一条路，两次踩坑经验一致。
 *
 * 嵌套语义（batch 里再调工具）
 * ---------------------------
 * 内层没传 target 时**继承外层**，不是清空。否则 batch 带 target 跑 3 步，
 * 每一步都会退回"改活动文档"—— 那正是要消灭的行为。
 */

import { AsyncLocalStorage } from 'async_hooks';

interface TargetStore {
  /** 原始 target 值（未校验）。undefined = 本层没传，沿用上层的 */
  target?: unknown;
}

const als = new AsyncLocalStorage<TargetStore>();

/**
 * 在当前调用链上设置 target，并执行 fn。
 *
 * @param target 本次调用的 target（可为 undefined —— 表示"不改变上层设定"）
 */
export function runWithTarget<T>(target: unknown, fn: () => T): T {
  if (target === undefined || target === null) {
    // 本层没传：沿用上层上下文（若也没有，就是 undefined，等于旧行为）
    const prev = als.getStore();
    return prev ? als.run(prev, fn) : fn();
  }
  return als.run({ target }, fn);
}

/** 当前调用链上的 target（可能来自上层工具，也可能是本层） */
export function currentTarget(): unknown {
  return als.getStore()?.target;
}
