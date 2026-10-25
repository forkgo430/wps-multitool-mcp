/**
 * Input: 无
 * Output: 三件套的实例/文档健康度诊断 + 挡路对话框 + 修掉的残留
 * Pos: 诊断与修复工具（阶段 4 的 D 部分）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么需要它
 * ------------
 * 自动化改 WPS 文档时有两类"说不清"的问题，之前只能靠翻日志猜：
 *   1) **ActiveDocument 空壳**：Documents.Count > 0，但活动文档是个取不到内容的空壳，
 *      后续调用要么诡异地失败，要么 RPC_E_CALL_REJECTED。实测踩过。
 *   2) **模态对话框挡路**：另存确认框、字体替换提示、文档恢复提示……
 *      COM 调用一律被拒，而且干等永远不会好（用户不点就不通）。
 *
 * 本工具一次性把这两件事说清楚，并且顺手修掉一类确定的残留：
 * **无名文档**（既没标题也没路径的空壳，通常是自动化脚本遗留的）。
 * 只关无名文档 —— 用户自己的文件一律不动，已命名的文档一律保留。
 *
 * 它是"诊断 + 有限修复"，不是万能修复：
 *   - 挡路的对话框只【报告】不自动关（关对话框等于替用户做决定，可能丢数据）
 *   - 卡死的 COM 连接要靠 wps_common_status 看桥那一层，超时要等桥自己重建
 *
 * 用法：工具报 com_error / document_not_found / RPC_E_CALL_REJECTED 时先跑这个。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';


/** PowerShell 侧 wpsRepair 返回的数据形状 */
interface RepairDoc {
  name?: string;
  saved?: boolean;
  healthy?: boolean;
}

interface RepairApp {
  running?: boolean;
  name?: string;
  path?: string;
  count?: number;
  activeHealthy?: boolean;
  documents?: RepairDoc[];
}

interface RepairData {
  apps?: Record<string, RepairApp>;
  dialogs?: string[];
  closed?: string[];
  hint?: string | null;
}

/** 文档健康度的三态描述（布尔不够用，得说清是哪种不健康） */
function describeApp(key: string, app: RepairApp): string {
  if (!app.running) return `${key}: 未运行`;
  const docs = app.documents ?? [];
  const parts: string[] = [];
  parts.push(`文档 ${app.count ?? docs.length} 个`);
  if (app.activeHealthy === false && (app.count ?? 0) > 0) {
    parts.push('活动文档【不健康】(空壳) — 已自动改用集合里第一个健康文档');
  } else {
    parts.push('活动文档正常');
  }
  const bad = docs.filter((d) => d.healthy === false);
  if (bad.length) {
    parts.push(`不健康文档 ${bad.length} 个: ${bad.map((d) => d.name || '<无名>').join(', ')}`);
  }
  const unsaved = docs.filter((d) => d.saved === false);
  if (unsaved.length) {
    parts.push(`未保存文档 ${unsaved.length} 个: ${unsaved.map((d) => d.name || '<无名>').join(', ')}`);
  }
  return `${key}: ${parts.join('，')}`;
}

export const repairDefinition: ToolDefinition = {
  name: 'wps_common_repair',
  description: `诊断并在【有限范围内】修复 WPS 自动化常见问题。

报告三件事：
1) 三件套各自的实例状态、文档总数，以及【每个文档的健康度】——
   "活动文档不健康(空壳)"这行专门用来解释"为什么 Documents.Count > 0 但工具取不到内容"
2) 挡路的对话框列表 —— RPC_E_CALL_REJECTED 的最常见来源就是有模态对话框在等用户点
3) 修掉的残留：只关【无名文档】（既没标题也没路径的空壳，通常是脚本遗留的）

会做的修改（仅此一类）：
- 关闭无名文档。已命名的文档一律不动，用户自己的文件不会被碰。

不会做的：
- 不自动关对话框（关对话框等于替你做决定，可能丢数据）—— 只报告，由你决定
- 不动任何已保存的文档

使用场景：
- 工具报 document_not_found / com_error / RPC_E_CALL_REJECTED 时【先跑这个】
- "WPS 里是不是卡住了"
- 怀疑有上次的自动化留下的空文档`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const repairHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  let r;
  try {
    r = await wpsClient.executeMethod<RepairData>('wpsRepair', {});
  } catch (err) {
    const m = err instanceof Error ? err.message : String(err);
    return {
      id: '',
      success: false,
      content: [{ type: 'text', text: `诊断失败：${m}` }],
      error: m,
    };
  }

  if (!r.success) {
    const m = String(r.error ?? '未知错误');
    return {
      id: '',
      success: false,
      content: [{ type: 'text', text: `诊断失败：${m}` }],
      error: m,
    };
  }

  const d = r.data ?? {};
  const apps = d.apps ?? {};
  const dialogs = d.dialogs ?? [];
  const closed = d.closed ?? [];

  const lines: string[] = ['=== 三件套与文档健康度 ==='];
  for (const key of ['word', 'excel', 'ppt']) {
    const app = apps[key];
    lines.push(app ? describeApp(key, app) : `${key}: (无数据)`);
  }

  lines.push('', '=== 挡路对话框 ===');
  if (dialogs.length) {
    lines.push(`检测到 ${dialogs.length} 个 —— COM 调用被拒多半是这个原因，请手动关闭：`);
    dialogs.forEach((x) => lines.push(`  · ${x}`));
  } else {
    lines.push('无（若仍报 RPC_E_CALL_REJECTED，说明不是对话框，桥那一层可能卡了，先看 wps_common_status）');
  }

  lines.push('', '=== 已清理的残留 ===');
  lines.push(closed.length ? closed.map((x) => `  · ${x}`).join('\n') : '无');

  const unhealthy = Object.entries(apps).filter(([, a]) => a.activeHealthy === false);
  if (unhealthy.length) {
    lines.push(
      '',
      `注意：${unhealthy.map(([k]) => k).join('/')} 的活动文档不健康。` +
        '本版本已让所有动作自动改用集合里第一个健康文档，所以工具通常还能用；' +
        '但如果你想确认操作的是不是"你正在看的那个文档"，请用带 documentName 参数的工具显式指定。'
    );
  }
  if (d.hint) lines.push('', d.hint);

  return {
    id: '',
    success: true,
    content: [{ type: 'text', text: lines.join('\n') }],
  };
};

export const repairTools: RegisteredTool[] = [
  { definition: repairDefinition, handler: repairHandler },
];

export default repairTools;