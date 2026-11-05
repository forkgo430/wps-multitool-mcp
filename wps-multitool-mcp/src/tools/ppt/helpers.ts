/**
 * Input: 各 ppt 工具 handler 的收尾调用
 * Output: 统一的 ToolCallResult 构造 + 带警告透传的 executeMethod 封装
 * Pos: PPT 工具收尾风格的统一出口（ppt-audit-2026-10-30 §一：102 个工具 0 个用助手 → 收口到这一套；
 *      范式照搬 word/helpers.ts，默认 app 为 PRESENTATION）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么以 word/gaps→helpers 的那套为基准：warnings 必须到模型眼前（A9 教训）——
 * 老样板会把 PS 侧特意准备的 warnings 丢掉（ppt 组 102 个 handler 此前无一透传）。
 * 迁移纪律：纯机械替换，不改任何行为——工具名/动作名/参数键/成功与校验文案逐字保留；
 * uuid 生成收敛到 ok/fail 各一处（callTool 会用请求 id 覆盖它，handler 里反复生成是纯浪费）；
 * 老样板的 try/catch 撤掉，异常统一由 toolRegistry.callTool 兜底（与 word 轮 03a534e 同一裁决）。
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
  app: WpsAppType = WpsAppType.PRESENTATION
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

/** add_* 后按回传的形状名回读它在页内的序号（体验修复轮 A）。
 *  为什么走 TS 回读而不是 ps1 加回传：ps1 本轮冻结（只许动 setSlideSize 一处），
 *  而形状追加在页尾、getShapes 的 index 就是 Shapes.Item 序号——按 name 精确命中，一次廉价读回。
 *  找不到/读回失败一律 undefined，调用方按缺失不打印——绝不为了 index 造假数据。 */
export async function shapeIndexByName(
  slideIndex: number,
  name: string
): Promise<number | undefined> {
  const g = await call<{ shapes?: Array<{ index: number; name: string }> }>('getShapes', { slideIndex });
  const hit = g.data?.shapes?.find((s) => s.name === name);
  return hit?.index;
}
