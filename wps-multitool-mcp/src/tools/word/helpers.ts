/**
 * Input: 各 word 工具 handler 的收尾调用
 * Output: 统一的 ToolCallResult 构造 + 带警告透传的 executeMethod 封装
 * Pos: Word 工具收尾风格的统一出口（word-audit-2026-10-29 §一：三套风格并存 → 收口到这一套）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么以 gaps/professional 的那套为基准：它是 A9 轮"warnings 必须到模型眼前"教训的产物，
 * 老样板（content/format/document/proofread）会把 PS 侧特意准备的 warnings 丢掉。
 * 迁移纪律：纯机械替换，不改任何行为——uuid 生成删掉（callTool 会用请求 id 覆盖它，
 * handler 里生成是纯浪费），PS warnings 一路带回。
 *
 * 迁移进度（word-audit 执行顺序）：proofread ✓ / format ✓ / content ▢ / document ▢
 */

import { v4 as uuidv4 } from 'uuid';
import { ToolCallResult, ToolWarning } from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';

type Args = Record<string, unknown>;

/** 统一调用：走常驻桥，PS 侧 data.warnings 已提到顶层，一路带回 */
export async function call<T>(
  action: string,
  params: Args,
  app: WpsAppType = WpsAppType.WRITER
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
