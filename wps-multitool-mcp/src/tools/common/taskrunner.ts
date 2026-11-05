/**
 * Input: 任务包名（task）+ 运行参数 + confirm
 * Output: 服务端一次执行整包，只回最终摘要（中间结果不出服务端）
 * Pos: 任务引擎的工具壳（wps_run_task）。引擎本体在 src/server/task-engine.ts——
 *      本文件只做 definition + 参数转发 + 回包包装（与全组 call/ok/fail 风格一致）。
 *      决策者 2026-11-05 亲自实现（设计：docs/agents/INTENT-DESIGN-2026-11-04.md §4.2）。
 *
 * 与 route_intent 的分工
 * ------------
 * route_intent 是"发现"：一句话 → 该用什么（direct / task_package 清单 / 组）。
 * wps_run_task 是"执行"：execute:'server' 的包在这里一次跑完。
 * 清单型（execute:'model'，v1 行为）的包不归本工具——仍由模型按 steps 逐步执行。
 * 发现 → 执行的完整链路：route_intent 命中 server 包时，nextStep 指向本工具
 * （由 intent.ts 改造时接线；本文件不反向依赖 intent.ts）。
 *
 * 台账口径
 * ------------
 * wps_run_task 目前 static（mock 断言 + 引擎单测；真机证据待 PPT 定位微轮修复后补）。
 */

import { v4 as uuidv4 } from 'uuid';
import { ToolDefinition, ToolHandler, ToolCategory, RegisteredTool } from '../../types/tools';
import { runTaskPackage } from '../../server/task-engine';

export const runTaskDefinition: ToolDefinition = {
  name: 'wps_run_task',
  description: `执行一个"服务端任务包"：一次调用跑完整个多步骤流程，只返回最终摘要。

适用：task-packages.json 里 execute:'server' 的包（如 weekly_report 周报）。
不适用：清单型包（仍按 route_intent 返回的 steps 由模型逐步执行）。

价值：中间结果（表格数据/图表图片等）留在服务端供步骤间引用，不占用模型上下文——
多步骤任务的 token 开销从"每步都过一遍"降为"只看最终结果"。

破坏性闸门：包里 confirm_before 标记的动作执行前会停下来要求确认（needConfirm:true），
向用户征得同意后带 confirm:true 重跑。`,
  category: ToolCategory.COMMON,
  group: 'common',
  inputSchema: {
    type: 'object',
    properties: {
      task: {
        type: 'string',
        description: "任务包名（route_intent 返回的 task 字段，如 'weekly_report'）",
      },
      args: {
        type: 'object',
        description: "运行参数（包内步骤用 {{args.键}} 引用），如 {\"body\": \"本周完成…\"}",
      },
      confirm: {
        type: 'boolean',
        description: '确认为 true 时才执行 confirm_before 里的破坏性动作',
      },
    },
    required: ['task'],
  },
};

export const runTaskHandler: ToolHandler = async (args) => {
  const task = String(args.task ?? '').trim();
  if (!task) {
    return {
      id: uuidv4(),
      success: false,
      content: [{ type: 'text', text: '缺少 task 参数' }],
      error: 'missing task',
    };
  }

  const runArgs = (args.args && typeof args.args === 'object' ? args.args : {}) as Record<string, unknown>;
  const confirm = args.confirm === true;

  const r = await runTaskPackage(task, runArgs, { confirm });

  const lines: string[] = [];
  lines.push(`${r.success ? '任务完成' : r.needConfirm ? '任务暂停待确认' : '任务中断'}：${r.title}`);
  lines.push(`进度: ${r.executed}/${r.total} 步`);
  for (const s of r.steps) {
    lines.push(`  ${s.ok ? '✓' : '✗'} 第${s.i}步 ${s.tool} — ${s.brief}`);
  }
  if (r.needConfirm) {
    lines.push(`待确认动作: ${r.confirmAction} —— 向用户征得同意后带 confirm:true 重新调用本工具。`);
  }
  if (r.error && !r.needConfirm) lines.push(`失败原因: ${r.error}`);
  if (r.data !== undefined) {
    const dataText = JSON.stringify(r.data);
    lines.push(`产物数据: ${dataText.length > 400 ? dataText.slice(0, 397) + '…' : dataText}`);
  }

  return {
    id: uuidv4(),
    success: r.success,
    content: [{ type: 'text', text: lines.join('\n') }],
    error: r.error,
    warnings: r.warnings,
  };
};

export const taskrunnerTools: RegisteredTool[] = [
  { definition: runTaskDefinition, handler: runTaskHandler },
];

export default taskrunnerTools;
