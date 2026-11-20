/**
 * Input: 工具组名 / 搜索关键词
 * Output: 工具组概况、加载/卸载结果、搜索结果
 * Pos: 工具组管理（按需加载，省上下文）。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *   - 默认只列出 core / common / word / excel
 *   - ppt（115 个）默认【不列出】，需要时 load
 * 否则客户端缓存里还是老列表。本模块用一个小订阅表把"变了"这件事抛出去，
 * 由 mcp-server 订阅后发通知（避免工具模块反向依赖 server）。
 *   WPS_TOOL_GROUPS=common,word   初始加载哪些组（逗号分隔，覆盖默认值）
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { toolRegistry } from '../../server/tool-registry';
import { verificationSummary } from '../../server/tool-verification';
import { essentialButUnverified } from '../../server/tool-tiers';
import { ok, fail } from './helpers';

// 变更事件放在零依赖的 tool-group-events 里，由 ToolRegistry 在 load/unload 后触发。
// 这里重新导出，方便订阅方只从工具模块引。
export { onToolGroupsChanged, offToolGroupsChanged, listenerCount } from '../../server/tool-group-events';

// ==================== 工具实现 ====================

function text(result: unknown): ToolCallResult {
  return ok(typeof result === 'string' ? result : JSON.stringify(result, null, 2));
}

export const listToolGroupsDefinition: ToolDefinition = {
  name: 'wps_list_tool_groups',
  description: `列出所有工具组及当前是否已在列。

默认是【基础档】：只列最常用且【真机验证过】的那批（工具组管理 + 常用三件套读写），
其余靠 wps_search_tools 搜、wps_load_tool_group 拉。
每组返回 total / listed / live 三个数：live 是该组真机验证过的工具数，
live < total 的组（目前是 ppt）结果要如实告诉用户"这部分未在真机验证"。
加载某个组 = 该组全部工具开始列出；卸载 = 回到基础档。
（WPS_TOOL_TIER=full 可关掉基础档过滤，改回"按组切"。）

使用场景：
- "有哪些工具组"、"PPT 工具怎么没了"
- 排查"这个工具为什么调不到"`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: { type: 'object', properties: {} },
};

export const listToolGroupsHandler: ToolHandler = async () => {
  const groups = toolRegistry.getGroups();
  const listed = toolRegistry.listTools().tools.length;
  const allNames = toolRegistry.listAllTools().map((d) => d.name);
  const total = allNames.length;
  const ver = verificationSummary(allNames);
  const unverifiedEssential = essentialButUnverified();
  return text({
    tier: toolRegistry.tier,
    groups,
    summary: {
      loadedGroups: toolRegistry.getLoadedGroups(),
      listedTools: listed,
      totalTools: total,
      verifiedTools: ver.live,
      note:
        toolRegistry.tier === 'essential'
          ? `基础档：列出的 = （基础档白名单 ∩ 真机验证过）∪ 已完整加载的组。` +
            `真机验证 ${ver.live}/${total}，未验证的组（ppt）默认不列。` +
            (unverifiedEssential.length
              ? `这批基础工具因未真机验证暂不默认列出：${unverifiedEssential.join(', ')}。`
              : '') +
            `未列出的工具仍可直接调用，也可以用 wps_search_tools 搜到再加载。`
          : '全量档：列出的 = 所有已加载组的工具。',
    },
  });
};

export const loadToolGroupDefinition: ToolDefinition = {
  name: 'wps_load_tool_group',
  description: `加载一个工具组，加载后该组工具才会出现在工具列表里。

常用组：word / excel / ppt / common（core 永远加载，无需手动加载）。
加载完成会通知客户端刷新工具列表。

使用场景：
- "我要做 PPT" → load ppt
- "刚才那些 PPT 工具怎么不见了"`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: {
    type: 'object',
    properties: {
      group: { type: 'string', description: '组名：word / excel / ppt / common' },
    },
    required: ['group'],
  },
};

export const loadToolGroupHandler: ToolHandler = async (args) => {
  const group = String(args.group ?? '').trim();
  if (!group) {
    return fail('缺少 group 参数', 'missing group');
  }
  const r = toolRegistry.loadGroup(group);
  if (!r.ok) {
    return fail(`加载失败：${r.error}。可用组见 wps_list_tool_groups`, r.error);
  }
  // 变更事件由 ToolRegistry 在 load/unload 后统一触发，这里不用管
  const listed = toolRegistry.listTools().tools.length;
  return text({
    group,
    alreadyLoaded: !r.changed,
    groupToolCount: r.toolCount,
    listedTools: listed,
    tier: toolRegistry.tier,
    message: r.changed
      ? `已加载 "${group}" 组（${r.toolCount} 个工具），工具列表已刷新，当前共列出 ${listed} 个`
      : `"${group}" 组本来就已经完整加载（${r.toolCount} 个工具），当前共列出 ${listed} 个`,
  });
};

export const unloadToolGroupDefinition: ToolDefinition = {
  name: 'wps_unload_tool_group',
  description: `卸载一个工具组，卸载后该组工具不再出现在工具列表里（仍可直接调用，但不推荐）。
注意：core 组不允许卸载（工具组管理本身在里面）。`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: {
    type: 'object',
    properties: { group: { type: 'string', description: '组名：word / excel / ppt / common' } },
    required: ['group'],
  },
};

export const unloadToolGroupHandler: ToolHandler = async (args) => {
  const group = String(args.group ?? '').trim();
  if (!group) {
    return fail('缺少 group 参数', 'missing group');
  }
  const r = toolRegistry.unloadGroup(group);
  if (!r.ok) {
    return fail(`卸载失败：${r.error}`, r.error);
  }
  // 变更事件由 ToolRegistry 在 load/unload 后统一触发，这里不用管
  const listed = toolRegistry.listTools().tools.length;
  return text({
    group,
    alreadyUnloaded: !r.changed,
    groupToolCount: r.toolCount,
    listedTools: listed,
    tier: toolRegistry.tier,
    message: r.changed
      ? `已卸载 "${group}" 组（${r.toolCount} 个工具不再列出，基础档里的仍会列出），当前列出 ${listed} 个`
      : `"${group}" 组本来就没完整加载，当前列出 ${listed} 个`,
  });
};

export const searchToolsDefinition: ToolDefinition = {
  name: 'wps_search_tools',
  description: `任何能力不确定，也可以先调我——在所有工具里搜索（**包括未加载组的工具**），返回名字 + 所属组 + 是否已列出 + 中文别名 + verification（live=真机跑通过 / static=只有静态校验，static 不是不能用，做完事如实告知即可）。

关键词直接用中文。**重点看 notFor**：本仓近名工具多（6 background、5 animation、4 transition、3 插入图片、2 zoom），名字几乎一样；靠前的命中带 examples 和 notFor（它不干什么、该用谁），同族工具务必按 notFor 挑，别瞎猜。

用途：组没加载时工具不在列表里，先搜出来，再 wps_load_tool_group 把它所在的组拉起来。`,
  category: ToolCategory.COMMON,
  group: 'core',
  inputSchema: {
    type: 'object',
    properties: {
      keyword: {
        type: 'string',
        description: '关键词（中文也行，会匹配工具名、描述与中文别名；留空则列出前 N 个）',
      },
      group: { type: 'string', description: '只在某个组里搜：word / excel / ppt / common' },
      limit: { type: 'number', description: '最多返回几项，默认 20' },
      hints: {
        type: 'boolean',
        description: '是否给靠前的命中补 examples / notFor，默认 true（内容多时设 false 省上下文）',
      },
    },
  },
};

export const searchToolsHandler: ToolHandler = async (args) => {
  const keyword = String(args.keyword ?? '');
  const group = args.group ? String(args.group) : undefined;
  const limit = typeof args.limit === 'number' ? args.limit : 20;
  const hints = args.hints === undefined ? true : Boolean(args.hints);
  const hits = toolRegistry.searchTools(keyword, { group, limit, hints });
  if (!hits.length) {
    return text({
      keyword,
      group: group ?? null,
      count: 0,
      message:
        '没搜到。中文关键词会去匹配别名，仍搜不到就换个说法（如"表格""背景""动画"）；也可先用 wps_list_tool_groups 看有哪些组。',
    });
  }
  const unloaded = hits.filter((h) => !h.listed);
  return text({
    keyword,
    group: group ?? null,
    count: hits.length,
    results: hits,
    ...(unloaded.length
      ? {
          nextStep: `其中 ${unloaded.length} 个当前未列出，调用前先 wps_load_tool_group（组名见各自 group 字段）`,
        }
      : {}),
  });
};

/** 工具组管理工具集合（属 core 组，永远加载） */
export const toolGroupTools: RegisteredTool[] = [
  { definition: listToolGroupsDefinition, handler: listToolGroupsHandler },
  { definition: loadToolGroupDefinition, handler: loadToolGroupHandler },
  { definition: unloadToolGroupDefinition, handler: unloadToolGroupHandler },
  { definition: searchToolsDefinition, handler: searchToolsHandler },
];

export default toolGroupTools;
