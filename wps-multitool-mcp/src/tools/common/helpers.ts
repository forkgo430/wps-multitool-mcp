/**
 * Input: 各 common 工具 handler 的收尾调用
 * Output: 统一的 ToolCallResult 构造 + 带警告透传的 executeMethod 封装
 * Pos: Common 工具收尾风格的统一出口（与 word/helpers.ts、ppt/helpers.ts、excel/helpers.ts 同范式）。
 * 与三组不同：common 跨 word/excel/ppt 三应用，executeMethod 多数不传 app（按动作路由），
 * 所以 call 的 app 参数【无默认值】，调用处有什么传什么。
 * 迁移纪律：纯机械替换，不改任何行为——工具名/动作名/参数键/成功与校验文案逐字保留；
 * uuid 生成收敛到 ok/fail 各一处（callTool 会用请求 id 覆盖它，handler 里反复生成是纯浪费）；
 */

import { v4 as uuidv4 } from 'uuid';
import { ToolCallResult, ToolWarning } from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { AppKey } from '../../client/action-routes';
import { WpsAppType } from '../../types/wps';

type Args = Record<string, unknown>;

/** 统一调用：走常驻桥，PS 侧 data.warnings 已提到顶层，一路带回；app 不传=按动作路由 */
export async function call<T>(
  action: string,
  params: Args,
  app?: WpsAppType | AppKey
): Promise<{ ok: boolean; data?: T; error?: string; warnings?: ToolWarning[] }> {
  const r = await wpsClient.executeMethod<T>(action, params, app);
  return { ok: r.success, data: r.data, error: r.error, warnings: r.warnings };
}

export const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: true,
  content: [{ type: 'text', text }],
  ...(warnings && warnings.length ? { warnings } : {}),
});

export const fail = (text: string, error?: string): ToolCallResult => ({
  id: uuidv4(),
  success: false,
  content: [{ type: 'text', text }],
  error: error ?? text,
});

/** 失败时统一文案：后端 error 优先，其次异常消息 */
export function errText(prefix: string, r: { error?: string }): string {
  return `${prefix}：${r.error ?? '未知错误'}`;
}

/** 异常消息归一：Error 对象取 message，其余 String 化 */
export function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
