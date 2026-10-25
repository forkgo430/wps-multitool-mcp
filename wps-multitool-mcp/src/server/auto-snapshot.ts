/**
 * Input: 工具名（规范名）+ 一次执行体
 * Output: 执行结果（可能在前面自动建了一份快照，失败时附 warnings）
 * Pos: 单工具写入的"可退"保险 —— 挂在 toolRegistry.callTool 这个唯一入口上。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有这一层
 * --------------
 * 以前只有 wps_common_batch 会在执行前存快照。用户直接调一个破坏性工具
 * （删页、覆盖样式、原地转格式）时没有任何保险 —— 而 COM 没有撤销，
 * 删掉就是删掉了。快照是唯一的补偿手段，它不该只在"批量"这条路上存在。
 *
 * 为什么挂在注册表而不是每个 handler 里
 * ----------------------------------
 * 296 个工具逐个加一遍，等于 296 个地方都可能漏。注册表是所有工具调用的唯一入口
 * （batch 的子调用也走它），挂在这里是唯一不会漏的位置。
 *
 * 判据（沿用 batch 的口径，别另发明）
 * ------------------------------
 *   只读（readOnlyHint）        → 不建。快照是保险，不是每次调用的仪式。
 *   破坏性（destructiveHint）    → 建。
 *   非幂等（idempotentHint === false）→ 建。
 *   其余写入（幂等的普通写入）    → 不建。重复执行结果一致，改坏了再写一次就行。
 * 每个写入都建会让用户一天攒几百份快照，最后等于没存（还会触发保留策略把有用的挤掉）。
 *
 * 为什么用 AsyncLocalStorage 防重，而不是一个深度计数器
 * ------------------------------------------------
 * MCP server 里 callTool 是**并发可达**的（stdio 上多个请求可以交错），batch 内部还会
 * 嵌套调子工具。全局 depth 计数会把两条独立调用链的层级算到一起 ——
 * A 链的 batch 跑着的时候，B 链的一次删除会以为自己也是子调用，于是**该保命时没保**。
 * ALS 跟着各自的调用链走，天然隔离；同一个理由在 gen-tool-annotations.js 里也用过。
 * 防重是必须的：batch 3 步会各自走一遍注册表，不防重就是 4 份快照（batch 1 + 3 步）。
 */

import { AsyncLocalStorage } from 'async_hooks';
import { getToolAnnotations } from './tool-annotations';
import { ToolWarning } from '../types/tools';
import { createSnapshot } from '../tools/common/snapshot';
import { createChildLogger } from '../utils/logger';

const logger = createChildLogger('AutoSnapshot');

/**
 * 从工具名推出它作用在哪个应用。
 *
 * 为什么要推而不是"让 PS 猜"：三个应用可以同时开着，而 ActiveDocument 是弱引用。
 * PS 兜底规则只能按"未保存优先"猜，一旦猜错，快照会把另一个应用里残留的文档复制下来 ——
 * 名义上有回滚保险，实际保的是别人的文件。工具名是现成的、确定的，不用白不用。
 */
export function appHintOf(toolName: string): 'word' | 'excel' | 'ppt' | undefined {
  if (toolName.startsWith('wps_word_')) return 'word';
  if (toolName.startsWith('wps_excel_')) return 'excel';
  if (toolName.startsWith('wps_ppt_')) return 'ppt';
  return undefined;
}

/** 当前调用链的嵌套深度。只有最外层（depth 0）建快照。 */
const chain = new AsyncLocalStorage<{ depth: number }>();

/**
 * 自己管快照、或被本层管理会很荒谬的工具。
 *   wps_common_batch        —— 它按【每个步骤】判风险，比按工具判更准；入口再建一次是重复。
 *   wps_common_snapshot_*   —— 快照/恢复工具本身：给"建快照"再建一份快照没有意义。
 */
const SELF_MANAGED = new Set([
  'wps_common_batch',
  'wps_common_snapshot_create',
  'wps_common_snapshot_restore',
  'wps_common_snapshot_list',
]);

/** 环境变量兜底开关：不想被自动快照的人可以整机关掉（不给 296 个工具加参数位） */
function disabledByEnv(): boolean {
  const v = (process.env.WPS_NO_AUTO_SNAPSHOT || '').toLowerCase();
  return ['1', 'true', 'yes', 'on'].includes(v);
}

/**
 * 这个工具值不值得存快照。返回空串 = 不必建；非空串 = 风险描述（写进快照 label 与日志）。
 * 判据与 batch 的 riskOf 一致 —— 两处口径必须一样，否则"批量有保险、单调用没有"这种不一致
 * 会让人以为单调用是安全的。
 */
export function snapshotRiskOf(toolName: string): string {
  if (SELF_MANAGED.has(toolName)) return '';
  const a = getToolAnnotations(toolName);
  if (!a || a.readOnlyHint) return '';
  if (a.destructiveHint) return `${toolName} 破坏性`;
  if (a.idempotentHint === false) return `${toolName} 非幂等`;
  return '';
}

/**
 * 在调用链最外层、且工具够危险时，先建一份快照，再执行。
 *
 * 失败语义：建不了快照【绝不阻断】执行（用户可能就是要改一份还没保存的新文档），
 * 但必须把"这次没有回滚保险"塞进 warnings —— 不吭声地少一层保险，比没有更坏。
 */
export async function withAutoSnapshot<T extends { warnings?: ToolWarning[] }>(
  toolName: string,
  run: () => Promise<T>
): Promise<T> {
  const outer = chain.getStore();
  const depth = outer ? outer.depth + 1 : 0;

  if (depth > 0) {
    // 嵌套调用（batch 的子步骤）：直接跑，绝不再建快照
    return chain.run({ depth }, run);
  }

  const why = disabledByEnv() ? '' : snapshotRiskOf(toolName);
  if (!why) return chain.run({ depth }, run);

  let snapWarnings: ToolWarning[] = [];
  try {
    const snap = await createSnapshot(`auto: ${why}`, appHintOf(toolName));
    if (snap.ok) {
      logger.info('auto snapshot created', { tool: toolName, path: snap.record.snapshotPath });
      if (snap.dirty) {
        snapWarnings = [
          {
            step: 'snapshot.dirty',
            message:
              '这次改动前已自动存快照，但快照取自磁盘、不含未保存改动；要连未保存内容一起保命，先 wps_common_save 再重做',
          },
        ];
      }
    } else {
      snapWarnings = [
        {
          step: 'snapshot',
          message: `改动前没能自动建快照：${snap.error}。这次执行没有回滚保险。`,
        },
      ];
      logger.warn('auto snapshot failed', { tool: toolName, error: snap.error, code: snap.code });
    }
  } catch (e) {
    // createSnapshot 自己已经兜了大部分异常，这里防的是更外层（比如 fs 模块本身炸了）
    snapWarnings = [
      {
        step: 'snapshot',
        message: `改动前没能自动建快照：${e instanceof Error ? e.message : String(e)}。这次执行没有回滚保险。`,
      },
    ];
  }

  const result = await chain.run({ depth }, run);
  if (snapWarnings.length) {
    result.warnings = [...(result.warnings ?? []), ...snapWarnings];
  }
  return result;
}
