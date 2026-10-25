/**
 * Input: 工具名
 * Output: 该工具的真机验证状态（live / static）与组级汇总
 * Pos: 验证状态的运行期查询层。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ------------
 * 284 个工具里，PPT 那 101 个大部分从来没在真机上跑过（台账里逐条记 live 的只有 18 个）。
 * 用户看到"284 个工具"会默认它们都能用，而 PPT 恰恰是最容易"不报错但做错了"的一类。
 * 状态必须由工具自己带出来，不能只写在文档里 —— 文档里写的东西，模型和用户都不会在选工具时去翻。
 *
 * 为什么默认是 static（fail-safe）
 * ------------------------------
 * 未登记的工具一律当 static：宁可让一个真验过的工具被标成静态，
 * 也不要让一个没验过的工具被当成可信。缺状态是有门禁兜底的
 * （scripts/manual/verification-verify.js + 生成器自身的退出码）。
 */

import {
  TOOL_VERIFICATION,
  ToolVerification,
  VERIFICATION_TOTALS,
} from './tool-verification.generated';

/** 单工具的验证状态；未登记 = static（fail-safe） */
export function verificationOf(toolName: string): ToolVerification {
  return TOOL_VERIFICATION[toolName] ?? 'static';
}

/** 是否真机验证过 —— 默认档只收这类工具 */
export function isLiveVerified(toolName: string): boolean {
  return verificationOf(toolName) === 'live';
}

/** 一组工具的验证覆盖：给 list_tool_groups / README 统计用，别手写数字 */
export function verificationSummary(toolNames: readonly string[]): {
  live: number;
  static: number;
  total: number;
} {
  let live = 0;
  for (const n of toolNames) if (isLiveVerified(n)) live++;
  return { live, static: toolNames.length - live, total: toolNames.length };
}

/** 表里的总量（生成期算好，避免每次全量遍历） */
export const VERIFICATION_COVERAGE = VERIFICATION_TOTALS;
