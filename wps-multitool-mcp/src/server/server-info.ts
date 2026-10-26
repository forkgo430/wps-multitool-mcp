/**
 * Input: 版本与署名信息（version.ts）
 * Output: MCP initialize 响应里报给客户端的 instructions（介绍 + 作者署名）
 * Pos: 客户端展示与模型首次感知这个服务器的全部依据。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么介绍要写在这里，而不是各写各的
 * ------------------------------------
 * 介绍会随 initialize 响应发给客户端，出现在用户面前，也进模型上下文。
 * 写长了白占 token，写短了用户不知道这个 MCP 怎么用。
 * 而"作者是谁""到底多少个工具""怎么用才不出错"这三件事本来就散在
 * README、CHANGELOG、各处注释里 —— 抄一遍就会抄出四个版本。
 *
 * 所以这里只放【客户端第一眼需要知道的】，并且全部从别处取真值：
 * 名字/版本/作者从 version.ts 拿，工具数从注册表数（不写死）。
 *
 * MCP 协议本身没有 author 字段，署名只能写进 instructions —— 这是唯一的出口，
 * 不是偷懒。
 */

import { MCP_AUTHOR, MCP_NAME, MCP_VERSION } from '../version';

/**
 * 工具总数在运行时才知道（registerBuiltinTools 的 12 个 + src/tools 的 273 个）。
 * 这里延迟求值，且【故意不写死数字】：写死的数字一定会随版本腐化，
 * 而文档里已经因为这个问题出过三次不一致。
 */
let toolCount = 0;
export function setToolCount(n: number): void {
  toolCount = n;
}

export function serverInstructions(): string {
  return [
    `${MCP_NAME} v${MCP_VERSION} —— 操作本机 WPS Office（文字 / 表格 / 演示）的 MCP 工具集。`,
    `作者：${MCP_AUTHOR}。`,
    toolCount ? `共 ${toolCount} 个工具，默认只列出 40 个基础工具，其余按需加载。` : '',
    '',
    '使用要点：',
    '1. 动手前先 wps_common_status 探活 —— 桥通不通、WPS 在不在、有没有活动文档，三态分开处理。',
    '2. wps_route_intent 判断该用哪个组，wps_load_tool_group 把它加载出来。',
    '   "工具不存在"通常不是没这个能力，是组没加载。',
    '3. 拿不准用哪个工具时用 wps_search_tools（支持中文关键词），',
    '   命中结果里的 notFor 必读 —— 本工具集有一批近名工具，选错不会报错，只会把事情改到别处。',
    '4. 每个工具都带 readOnlyHint / destructiveHint 等注解，是真实属性：',
    '   只读工具结果不对，通常说明前置步骤没做（组没加载、文档没开、路径不对），重复调用没有用。',
    '5. 失败文本的方括号前缀写明能不能重试：',
    '   [可重试] 可以直接重发；[勿重试·可能已执行一半] 必须先读回来看结果，绝不重放。',
    '',
    '写操作没有回滚：COM 写入不是事务。destructiveHint 为 true 的工具（删除、另存为、',
    '格式转换、关闭文档等）执行前先跟用户确认。',
  ]
    .filter(Boolean)
    .join('\n');
}