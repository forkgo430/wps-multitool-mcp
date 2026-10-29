/**
 * Input: tool-registry 的调用包裹层——写入类工具（annotations.readOnlyHint !== true，
 *        缺注解按 fail-closed 也算写入类）执行前传 app，执行后传 app 记账
 * Output: warnings 字符串数组（空 = 没有协作提醒；只提醒、不阻断）
 * Pos: 用户与 AI 协作同一份文档时的安全网（2026-10-29 所有者批准的三件协作增强之一，
 *      决策者亲自实现）。原理：记住每个应用"上次 AI 操作后"的文档指纹
 *      （getActiveDocumentState 新增的 editProbe 字段，ps1 三型各一种廉价探针），
 *      下一个写入类工具执行前重新探测——指纹变了 = 期间有人工编辑，
 *      往 warnings 写一条提醒（不阻断写入）。
 *
 * 两条铁的教训写进了实现：
 *   1. 探测失败必须静默降级——CI 的 Linux 没有 WPS，快照 getActiveState 在那里
 *      抛过一次异常把 CI 炸红（2026-10-28）；这里任何异常都吞掉，绝不外抛。
 *   2. 写入完成后必须再探测一次存档——否则 AI 自己的修改会被下一次检查
 *      误报成"人工编辑"。误报比没有更毁信任。
 *
 * 指纹口径（诚实声明，都有漏报，这是提醒不是保证）：
 *   word = 正文流末位（插入/删除会动，纯排版不动）
 *   excel = UsedRange 地址+行列数（结构性增删会动，改单元格值不动）
 *   ppt = 页数（增删页会动，改页内文字不动）
 */
import { wpsClient } from '../client/wps-client';

const lastProbe = new Map<string, string>();
const reviewHinted = new Set<string>();

interface DocStateWithProbe {
  app?: string;
  name?: string;
  editProbe?: string | null;
}

/** 从工具名推应用：wps_word_x → word；common/认不出 → ''（ps1 侧自动判型） */
export function appOfToolName(name: string): string {
  if (name.startsWith('wps_word_')) return 'word';
  if (name.startsWith('wps_excel_')) return 'excel';
  if (name.startsWith('wps_ppt_')) return 'ppt';
  return '';
}

/** 探测当前活动文档指纹。失败返回 null（降级），绝不抛出 */
async function probeNow(app: string): Promise<{ probe: string | null; name: string }> {
  try {
    const r = await wpsClient.executeMethod<DocStateWithProbe>(
      'getActiveDocumentState',
      app ? { app } : {}
    );
    if (r.success && r.data) {
      const data = r.data as DocStateWithProbe;
      return { probe: data.editProbe ?? null, name: data.name ?? '' };
    }
  } catch {
    /* 无 WPS/桥不可达：静默降级 */
  }
  return { probe: null, name: '' };
}

/**
 * 写入类工具执行前调：与上次存档的指纹比对，变了就给协作提醒。
 * 没有存档（本会话第一次动这个应用）不提醒——没有"上次"可比。
 */
export async function collabGuardCheckBefore(app: string): Promise<string[]> {
  const warnings: string[] = [];
  const prev = lastProbe.get(app);
  if (prev === undefined) return warnings;
  const { probe, name } = await probeNow(app);
  if (!probe) return warnings;
  lastProbe.set(app, probe);
  if (probe !== prev) {
    warnings.push(
      `文档「${name || app}」在上次 AI 操作后发生了变化（可能是人工编辑）——写入前建议重新读取目标位置，别按旧坐标下笔`
    );
    if (app === 'word' && !reviewHinted.has(app)) {
      reviewHinted.add(app);
      warnings.push(
        '检测到人工与 AI 同时在编辑——可开启修订模式（wps_word_review_mode）让 AI 的修改带修订痕迹，由人工逐条审阅'
      );
    }
  }
  return warnings;
}

/** 写入类工具执行后调：把"AI 操作后"的指纹存档，给下一次比对用。失败静默。 */
export async function collabGuardNoteAfter(app: string): Promise<void> {
  try {
    const { probe } = await probeNow(app);
    if (probe) lastProbe.set(app, probe);
  } catch {
    /* 不外抛 */
  }
}

/** 测试用：清空记忆 */
export function resetCollabGuardForTest(): void {
  lastProbe.clear();
  reviewHinted.clear();
}
