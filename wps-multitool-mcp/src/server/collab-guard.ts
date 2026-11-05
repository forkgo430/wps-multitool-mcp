/**
 * Input: tool-registry 的调用包裹层——写入类工具（annotations.readOnlyHint !== true，
 *        缺注解按 fail-closed 也算写入类）执行前传 app，执行后传 app 记账
 * Output: warnings 字符串数组（空 = 没有协作提醒；只提醒、不阻断）
 * Pos: 用户与 AI 协作同一份文档时的安全网（2026-10-29 所有者批准的三件协作增强之一，
 *      决策者亲自实现）。原理：记住每份文档"上次 AI 操作后"的指纹
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
 * 第三条教训（体验修复轮 B，2026-11-03）：存档键必须是【文档级】而不是应用级——
 *   原来键是 app、指纹却是"活动文档"的：save_as 另存/切换文档后，新文档的指纹
 *   对比旧文档的存档 → 必然不等 → 必误报"人工编辑"（所有者 140 次实测踩中）。
 *   修法（决策者预备报告预裁）：键改 app+文档身份；换了文档 = 没有"上次"可比，
 *   静默重建档不提醒；提醒文案带上/本两侧指纹，误报可自查。
 *
 * 指纹口径（诚实声明，都有漏报，这是提醒不是保证）：
 *   word = 正文流末位（插入/删除会动，纯排版不动）
 *   excel = UsedRange 地址+行列数（结构性增删会动，改单元格值不动）
 *   ppt = 页数（增删页会动，改页内文字不动）
 */
import { wpsClient } from '../client/wps-client';

interface ProbeRecord {
  probe: string;
  name: string;
}

const lastProbe = new Map<string, ProbeRecord>();
const reviewHinted = new Set<string>();

interface DocStateWithProbe {
  app?: string;
  name?: string;
  fullName?: string;
  editProbe?: string | null;
}

/** 从工具名推应用：wps_word_x → word；common/认不出 → ''（ps1 侧自动判型） */
export function appOfToolName(name: string): string {
  if (name.startsWith('wps_word_')) return 'word';
  if (name.startsWith('wps_excel_')) return 'excel';
  if (name.startsWith('wps_ppt_')) return 'ppt';
  return '';
}

/** 存档键：应用 + 文档身份（fullName 优先，其次 name；都拿不到退回 app——老行为兜底）。
 *  键必须是文档级：同应用开多份文档时，各存各的指纹互不误伤（体验修复轮 B）。 */
function probeKey(app: string, fullName?: string, name?: string): string {
  const doc = fullName || name;
  return doc ? `${app}|${doc}` : app;
}

/** 探测当前活动文档指纹。失败返回 null（降级），绝不抛出 */
async function probeNow(app: string): Promise<{ probe: string | null; name: string; fullName: string }> {
  try {
    const r = await wpsClient.executeMethod<DocStateWithProbe>(
      'getActiveDocumentState',
      app ? { app } : {}
    );
    if (r.success && r.data) {
      const data = r.data as DocStateWithProbe;
      return {
        probe: data.editProbe ?? null,
        name: data.name ?? '',
        fullName: data.fullName ?? '',
      };
    }
  } catch {
    /* 无 WPS/桥不可达：静默降级 */
  }
  return { probe: null, name: '', fullName: '' };
}

/**
 * 写入类工具执行前调：与【本文档】上次存档的指纹比对，变了就给协作提醒。
 * 没有存档（本会话第一次动这份文档，或刚从别的文档切过来）→ 静默重建档、
 * 不提醒——没有"上次"可比（体验修复轮 B：这正是 save_as 误报的修复点）。
 */
export async function collabGuardCheckBefore(app: string): Promise<string[]> {
  const warnings: string[] = [];
  const { probe, name, fullName } = await probeNow(app);
  if (!probe) return warnings;
  const key = probeKey(app, fullName, name);
  const prev = lastProbe.get(key);
  lastProbe.set(key, { probe, name });
  if (!prev) return warnings;
  if (probe !== prev.probe) {
    const ctx = `上次存档：「${prev.name || app}」${prev.probe.slice(0, 16)}… → 本次：「${name || app}」${probe.slice(0, 16)}…`;
    warnings.push(
      `文档「${name || app}」在上次 AI 操作后发生了变化（可能是人工编辑；${ctx}）——写入前建议重新读取目标位置，别按旧坐标下笔`
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

/** 写入类工具执行后调：把"AI 操作后"的指纹按【文档】存档，给下一次比对用。失败静默。 */
export async function collabGuardNoteAfter(app: string): Promise<void> {
  try {
    const { probe, name, fullName } = await probeNow(app);
    if (probe) lastProbe.set(probeKey(app, fullName, name), { probe, name });
  } catch {
    /* 不外抛 */
  }
}

/** 测试用：清空记忆 */
export function resetCollabGuardForTest(): void {
  lastProbe.clear();
  reviewHinted.clear();
}
