/**
 * Input: 无
 * Output: "工具组已变化"的订阅/触发
 * Pos: 工具组变更事件（避免工具模块与注册表互相依赖）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要单独一个模块
 * --------------------
 * 事件必须在【注册表】里触发（loadGroup/unloadGroup 才是唯一真相），
 * 而订阅方是 mcp-server（要发 notifications/tools/list_changed）。
 * 如果事件表写在 tools/common/toolgroups.ts 里，就会变成
 *   tool-registry → toolgroups → tool-registry 的循环依赖。
 * 所以抽成这个零依赖的小模块，两边都只依赖它。
 */

export type ToolGroupsChangedListener = () => void;

const listeners = new Set<ToolGroupsChangedListener>();

/** 订阅"工具组变了"。mcp-server 用它来发 notifications/tools/list_changed */
export function onToolGroupsChanged(fn: ToolGroupsChangedListener): void {
  listeners.add(fn);
}

/** 取消订阅（测试用） */
export function offToolGroupsChanged(fn: ToolGroupsChangedListener): void {
  listeners.delete(fn);
}

/** 触发"工具组变了"。由 ToolRegistry.loadGroup/unloadGroup 在真的发生变化后调用 */
export function emitToolGroupsChanged(): void {
  for (const fn of listeners) {
    try {
      fn();
    } catch {
      /* 订阅方出错不该影响工具本身的结果 */
    }
  }
}

/** 当前订阅者数量（诊断用） */
export function listenerCount(): number {
  return listeners.size;
}
