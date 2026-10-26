/**
 * Input: 1~3 个待执行步骤（工具名 + 参数，或裸动作名 + 参数）
 * Output: 每步结果摘要 + 汇总 + 聚合后的部分成功提示
 * Pos: 批量执行工具（上限 3 步，赢容定的）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么有这个
 * ------------
 * "把销售表按地区汇总→降序→画柱状图→导出 PDF"，在没有批量的时候要发 4 次往返，
 * 每次都得重新定位文档、重新过一遍路由。往返不是贵，是**容易错**：中间用户切了一下
 * 窗口，后面几步就写到别的文件里去了。
 *
 * 为什么上限是 3
 * --------------
 * COM 是 STA，同一时刻只能跑一个动作，串行是硬约束。步数越多，单次调用占着通道
 * 越久，撞上"用户正在手动操作 WPS"的机会越大，也越容易中途超时留个半成品。
 * 3 步是"一次往返能收掉一个小任务"的上限；再往上应该是多发几次，
 * 或者交给 skill 规划成多阶段 —— 而不是把一把梭塞进一次调用。
 *
 * 失败语义（与 warnings 约定一致）
 * --------------------------------
 *   全成功        → success=true，无 warnings
 *   部分成功      → success=true + warnings（每条写清是哪一步没成、为什么）
 *   全部失败      → success=false
 * 遇错【不中断】后续步骤：前面失败不代表后面也会失败（比如只是"这个属性不支持"）。
 * 但绝不重试失败的步骤 —— COM 不是事务，重放可能写两遍。
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
  ToolWarning,
} from '../../types/tools';
import { toolRegistry, ToolRegistry } from '../../server/tool-registry';
// 风险判据只有一份：与单工具写入的自动快照共用（口径不一致会让人以为单调用也安全）
import { snapshotRiskOf } from '../../server/auto-snapshot';
import { isDestructiveAction, isIdempotentAction } from '../../bridge/cancellable';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';
import { classifyWpsError } from '../../utils/error';
import { createSnapshot, appendOpLog } from './snapshot';

/** 赢容定的上限。不是"随手写的保守值"，是通道占用的现实约束（见头部注释） */
export const BATCH_MAX_STEPS = 3;

/**
 * 这一步值不值得为它存快照。
 * 破坏性（删、覆盖、原地转格式）最痛；非幂等（重放会叠两层）次之。
 * 只读的一律不存 —— 快照是保险，不是每次调用的仪式。
 *
 * 工具步骤的判据走 snapshotRiskOf（与单工具自动快照同一份规则）；
 * 只有裸动作名（batch 的逃生舱）才在这里按动作名判。
 */
function riskOf(step: BatchStep): string {
  if (step.tool) return snapshotRiskOf(step.tool);
  if (step.action) {
    if (isDestructiveAction(step.action)) return `${step.action} 破坏性`;
    if (!isIdempotentAction(step.action)) return `${step.action} 非幂等`;
  }
  return '';
}

type Args = Record<string, unknown>;

interface BatchStep {
  /** 走注册表调工具（推荐，模型熟悉且能拿到工具的 warnings） */
  tool?: string;
  /** 裸动作名（逃生舱：注册表里没有对应工具时直接打 PS 动作） */
  action?: string;
  params?: Args;
  /** action 模式下的目标应用 */
  app?: 'word' | 'excel' | 'ppt' | 'common';
}

const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: true,
  content: [{ type: 'text', text }],
  ...(warnings && warnings.length ? { warnings } : {}),
});

/** 失败也要能带 warnings：全批失败时"哪一步为什么失败"就在里面，不能丢 */
const fail = (text: string, error?: string, warnings?: ToolWarning[]): ToolCallResult => ({
  id: uuidv4(),
  success: false,
  content: [{ type: 'text', text }],
  error: error ?? text,
  ...(warnings && warnings.length ? { warnings } : {}),
});

function textOf(result: ToolCallResult): string {
  const first = result.content.find((c) => c.type === 'text');
  const s = first && first.type === 'text' ? String(first.text) : '';
  return s.length > 160 ? `${s.slice(0, 160)}…` : s;
}

export const batchDefinition: ToolDefinition = {
  name: 'wps_common_batch',
  description: `一次执行 1~${BATCH_MAX_STEPS} 个操作，省掉逐个调用的往返。

**什么时候用**：任务能拆成几步且几步连着做才对（汇总→排序→出图；改样式→套样式→插目录）。
**什么时候不用**：要等用户确认中间结果、需要长时间跑、或者只是想"多试几种方案"——
那属于多发几次调用，不该塞进一批里。

步骤写法二选一：
1) \`{"tool": "wps_excel_sort_range", "params": {...}}\` —— 走工具名（推荐）
2) \`{"action": "readRange", "params": {...}, "app": "excel"}\` —— 裸动作名，注册表里没对应工具时用

**失败语义**：某步失败不会中断后续步骤。全成功=成功；部分成功=success 仍为 true 但会带
warnings 标明哪步没成；全部失败才算失败。失败的步骤【不会自动重试】。

**自动快照**：batch 里有破坏性或非幂等的步骤时，执行前会自动给当前文档存一份快照
（可用 snapshot=false 关掉）。快照建不成功不会拦住执行，但会在 warnings 里明说
"这次没有回滚保险" —— 别把它当没事。

注意：快照是【补偿】不是【事务】。它只保证"改坏了能退回磁盘上那一版"，
救不了"写到一半的中间态"，也救不了未保存改动（那种情况会标 dirty）。`,
  category: ToolCategory.COMMON,
  group: 'common',
  inputSchema: {
    type: 'object',
    properties: {
      steps: {
        type: 'array',
        description: `要依次执行的步骤，最多 ${BATCH_MAX_STEPS} 个。按顺序执行，不并发。`,
        minItems: 1,
        maxItems: BATCH_MAX_STEPS,
        items: {
          type: 'object',
          properties: {
            tool: { type: 'string', description: '工具名，如 wps_excel_sort_range' },
            action: { type: 'string', description: '裸动作名（tool 不给时用这个）' },
            params: { type: 'object', description: '该步的参数' },
            app: {
              type: 'string',
              enum: ['word', 'excel', 'ppt', 'common'],
              description: 'action 模式的目标应用；tool 模式会忽略',
            },
          },
        },
      },
      snapshot: {
        type: 'boolean',
        description:
          '执行前是否自动给当前文档建快照（默认 true）。只在有破坏性/非幂等步骤时才真的建；' +
          '设 false = 你自己承担没有回滚的风险。',
      },
    },
    required: ['steps'],
  },
};

export const batchHandler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {
  const rawSteps = args.steps;
  if (!Array.isArray(rawSteps) || rawSteps.length === 0) {
    return fail('steps 必须是非空数组。', 'invalid_params');
  }
  if (rawSteps.length > BATCH_MAX_STEPS) {
    return fail(
      `一次最多 ${BATCH_MAX_STEPS} 步，你给了 ${rawSteps.length} 步。` +
        '步数多不一定是好事：COM 串行，一次占通道太久容易中途超时留半成品。' +
        `请拆成多次调用（每批 ≤${BATCH_MAX_STEPS} 步），或先确认前一批的结果再发下一批。`,
      'invalid_params'
    );
  }

  const steps = rawSteps as BatchStep[];
  const lines: string[] = [];
  const warnings: ToolWarning[] = [];
  let okCount = 0;

  // ---- 执行前：有风险的批次先存一份快照 ----------------------------------
  // 判据不是"写不写"，而是"改坏了痛不痛"：破坏性（删/覆盖）与非幂等（重放会叠两层）
  // 才值得存。每个写入都存会让人一天攒几百份快照，最后等于没存。
  const risky = steps
    .map((s, i) => ({ index: i, step: s, why: riskOf(s) }))
    .filter((x) => x.why);
  const wantsSnapshot = args.snapshot !== false && risky.length > 0;
  let snapshotPath: string | null = null;

  if (wantsSnapshot) {
    const labels = risky.map((x) => x.step.tool || x.step.action).join(', ');
    const snap = await createSnapshot(`batch: ${labels}`);
    if (snap.ok) {
      snapshotPath = snap.record.snapshotPath ?? null;
      lines.push(
        `已自动建快照（${risky.map((x) => x.why).join(' / ')}）→ ${snapshotPath}`,
        snap.dirty ? '  注意：快照取自磁盘，不含未保存改动。' : '',
        ''
      );
      if (snap.dirty) {
        warnings.push({
          step: 'snapshot.dirty',
          message: '文档有未保存改动，自动快照是磁盘版本；要连未保存内容一起保命，先 wps_common_save 再重试',
        });
      }
    } else {
      // 建不了快照不阻断执行（用户可能就是要改一份还没保存的新文档），但必须说清
      warnings.push({
        step: 'snapshot',
        message: `没能建快照：${snap.error}。这次执行没有回滚保险。`,
      });
      lines.push(`没能建快照：${snap.error}（继续执行，但没有回滚保险）`, '');
    }
  }

  for (let i = 0; i < steps.length; i++) {
    const step = steps[i];
    const label = step.tool || step.action || `第${i + 1}步`;

    if (!step.tool && !step.action) {
      warnings.push({ step: `第${i + 1}步`, message: '没给 tool 也没给 action，跳过' });
      lines.push(`${i + 1}. ${label} — 跳过（没给 tool 也没给 action）`);
      continue;
    }

    try {
      let result: ToolCallResult;
      if (step.tool) {
        result = await runTool(step.tool, step.params ?? {});
      } else {
        result = await runAction(step.action as string, step.params ?? {}, step.app ?? 'common');
      }

      appendOpLog({
        ts: new Date().toISOString(),
        kind: 'batch',
        tool: label,
        success: result.success,
        snapshot: snapshotPath ?? undefined,
      });

      if (result.success) {
        okCount++;
        lines.push(`${i + 1}. ${label} — 成功${result.warnings?.length ? '（有保留，见下）' : ''}`);
        for (const w of result.warnings ?? []) {
          warnings.push({ step: `${label}：${w.step ?? '提示'}`, message: w.message });
        }
      } else {
        // content 是给人看的摘要，error 是原始错误（WPS 原文常在 error 里）。
        // 两个都留、去重后拼接 —— 只取其一会丢信息：只取 error 会把可读提示换成短码
        // （比如 'tool_not_found'），只取 content 可能丢掉 WPS 的原始报错。
        const detail =
          [textOf(result), result.error]
            .filter(Boolean)
            .filter((v, i, arr) => arr.indexOf(v) === i)
            .join('｜') || '未知失败';
        warnings.push({ step: label, message: detail });
        lines.push(`${i + 1}. ${label} — 失败：${detail}`);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      // 分类只用来决定"这句话怎么写"，不做任何重试 —— COM 不是事务，重放可能写两遍
      const kind = classifyWpsError(message);
      warnings.push({ step: label, message: `${message}（归类：${kind}）` });
      lines.push(`${i + 1}. ${label} — 异常：${message}`);
      appendOpLog({
        ts: new Date().toISOString(),
        kind: 'batch',
        tool: label,
        success: false,
        snapshot: snapshotPath ?? undefined,
        error: message.slice(0, 200),
      });
    }
  }

  const summary = `批量执行完成：${okCount}/${steps.length} 步成功。`;
  const body = [summary, '', ...lines].join('\n');

  if (okCount === 0) {
    // 全失败：success=false，但 warnings 必须一起交出去 —— 逐步失败原因就在里面
    return fail(
      `${body}\n没有任何一步成功，文档状态请自己确认（可能已部分改动）。`,
      'batch_all_failed',
      warnings
    );
  }
  // 部分成功走 success=true + warnings：批次本身跑完了，具体哪步没成交给 warnings 说清
  return ok(body, warnings);
};

/** 走注册表：拿到工具自己的 warnings，组没加载时给出可操作的提示 */
async function runTool(name: string, params: Args): Promise<ToolCallResult> {
  if (!toolRegistry.hasTool(name)) {
    const similar = toolRegistry.suggestSimilar(name);
    const hint = similar.length
      ? `组 ${similar[0].group} 可能未加载，先 wps_load_tool_group。`
      : '用 wps_search_tools 搜一下，或 wps_list_tool_groups 看有哪些组。';
    return fail(`工具 ${name} 不存在。${hint}`, 'tool_not_found');
  }
  return toolRegistry.callTool(ToolRegistry.createRequest(name, params));
}

/** 走裸动作名：注册表里没有对应工具时的逃生舱 */
async function runAction(action: string, params: Args, app: string): Promise<ToolCallResult> {
  const typeMap: Record<string, WpsAppType> = {
    word: WpsAppType.WRITER,
    excel: WpsAppType.SPREADSHEET,
    ppt: WpsAppType.PRESENTATION,
    common: undefined as unknown as WpsAppType,
  };
  const r = await wpsClient.executeMethod(action, params, typeMap[app]);
  if (!r.success) return fail(r.error ?? `${action} 执行失败`, r.error);
  return ok(`${action} 完成${r.warnings?.length ? '（有保留）' : ''}`, r.warnings);
}

export const batchTools: RegisteredTool[] = [
  { definition: batchDefinition, handler: batchHandler },
];
