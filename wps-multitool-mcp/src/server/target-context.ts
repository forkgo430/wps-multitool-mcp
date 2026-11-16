/**
 * Input: 工具调用参数里的 target（可选）
 * Output: 当前调用链上"要钉住哪一份"的上下文
 * Pos: target 的跨层传递。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 67 个 word handler 都是这样写的：从 args 里取自己认识的字段，然后**显式**拼一个新对象
 *      两个并发调用会互相覆盖，一个任务的 target 会漂到另一个任务上；
 *      auto-snapshot.ts 用同一个理由选了同一条路，两次踩坑经验一致。
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
