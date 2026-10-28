/**
 * Input: MCP 工具名
 * Output: MCP 标准工具注解（readOnlyHint / destructiveHint / idempotentHint / openWorldHint）
 * Pos: 让客户端【从 tools/list 直接读到】每个工具的读写属性，而不是让模型靠描述去猜。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么必须是注解，不能靠文档教
 * ------------------------------
 * 实测：几百个工具里，靠名字能一眼确认只读的只有一小部分（见 cancellable.ts 的判定逻辑），
 * 其余是写入或判不出。指望模型读描述去判断"这个能不能动我的文档"，
 * 就是让它在大量模糊项里赌。注解是机器可读的，客户端可以据此决定要不要弹确认、
 * 要不要允许自动执行 —— 这比在 skill 里写一句"注意只读"可靠得多。
 *
 * 数据来源
 * --------
 * 动作名映射由 scripts/manual/gen-tool-annotations.js 运行时观测生成：
 * 把 wpsClient 的出口换成假实现，逐个跑 handler 记下它发出的动作名。
 * 注解本身在这里【运行时算】，不写死进生成文件 —— 判定规则改了不必重新生成。
 */

import {
  isReadOnlyAction,
  isDestructiveAction,
  isIdempotentAction,
} from '../bridge/cancellable';
import { TOOL_ACTIONS } from './tool-annotations.generated';
import { resolveToolAlias } from './tool-aliases.generated';

/** MCP 规范里的工具注解（字段名与规范一致，别改） */
export interface ToolAnnotations {
  /** 人类可读的短标题，给不支持富渲染的客户端用 */
  title?: string;
  /** 只读：不改任何状态。客户端可放心自动执行 */
  readOnlyHint?: boolean;
  /** 可能造成破坏性更新：会删内容、丢未保存改动、原地改格式 */
  destructiveHint?: boolean;
  /** 重复执行结果一致（可安全重放） */
  idempotentHint?: boolean;
  /** 会与外部系统交互。本 MCP 只碰本机 WPS，一律 false */
  openWorldHint?: boolean;
}

const cache = new Map<string, ToolAnnotations | undefined>();

/**
 * 自动生成覆盖不到的工具 —— 手写在这里，并说明为什么。
 *
 * 有两类，覆盖不到的原因完全不同：
 *
 * 【一】动作名是运行期参数，静态推不出来。
 *   wps_execute_method 在 mcp-server.ts 里直接 this.registry.register()，
 *   不走 allTools，所以 gen 脚本观测不到它；它的动作名还是调用时传进来的字符串。
 *   它能执行【任意】WPS 动作（删东西、关文档、原地转格式），破坏性按 true 给。
 *
 * 【二】它本身不是"一个动作"，而是"编排别人的容器"。
 *   wps_common_batch 是批量执行器：它的 handler 不直接调 executeAction，
 *   而是转手调注册表里的其它工具，所以观测不到任何动作名，
 *   落到"按工具名推断"会推出 'batch' 这种毫无意义的名字 → 判成非破坏性、幂等。
 *   但它能装下任意写入型工具（含删除、另存为、关闭），这三项都必须按最坏情况给：
 *   非只读、破坏性、非幂等。名字推断在这里不是"保守"，是"猜错了"。
 */
const EXTRA_ANNOTATIONS: Record<string, ToolAnnotations> = {
  // deckkit 引擎（E2）：无 ps1 action 的编排工具——自身不直接调 COM，逐动作经 callTool
  //  走真入口（每个子动作各自的注解/快照护栏生效）。入口本身按最坏情况给：非只读、非幂等
  //  （中途失败会留下半成品文稿）；破坏性与 wps_common_batch 同口径：不标（子动作各自定）。
  wps_ppt_build_deck: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: false,
  },
  wps_execute_method: {
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: false,
  },
  wps_common_batch: {
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: false,
  },
  /**
   * 【三】观测到的动作名骗人：它只发了 getActiveDocumentState（只读），
   * 真正的副作用是 Node 侧写文件（建快照 / 覆盖原文件）。按动作名推断会得出
   * "只读"或"非破坏性"，而客户端可能拿它决定要不要弹确认 —— 这里必须手写。
   */
  wps_common_snapshot_create: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: false,
    openWorldHint: false,
  },
  wps_common_snapshot_restore: {
    // 默认写副本、不覆盖；但 overwrite=true 时会原地盖掉原文件 —— 按最坏情况给
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: false,
  },
  // list 压根不发动作（只读 oplog），名字推断会推成 snapshotList → 判成非只读，必须手写
  wps_common_snapshot_list: {
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },
  /**
   * 【四】动作名带点，观测器按"合法标识符"判定把它归成 unresolved，推不出来。
   *   wps_ppt_unify_font 调的是 slide.unifyFont（带点的动作名），注解生成器的
   *   /^[A-Za-z_][A-Za-z0-9_]*$/ 不收它 → 只能手写。
   *   值是人工判定的：它把一页/整份文稿里所有文本框的字体名改掉 —— 写入、原地改格式，
   *   但不改内容、不删东西（非破坏），重复执行结果一致（幂等），只碰本机（openWorld=false）。
   */
  wps_ppt_unify_font: {
    readOnlyHint: false,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },

  /**
   * 【五】压根不发 COM 动作的工具（10 个）—— 副作用不在 WPS 里。
   *   cache 系列写在 Node 侧内存；工具组三个改的是本进程的注册表；
   *   wps_search_tools / wps_route_intent 只做字符串匹配；
   *   wps_word_proofread_basic 是纯本地正则（描述里就写着"无需 API 调用"）。
   *   按工具名推断会推出 cacheData / searchTools / proofreadBasic 这类名字，
   *   而它们在 wps-com.ps1 里一个 case 都没有 —— 那不是"保守"，是拿假输入做安全判定。
   */
  wps_cache_data: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_get_cached_data: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_list_cache: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  // 清的是自己建的缓存，但按"会丢数据"给 destructive —— 客户端该弹确认
  wps_clear_cache: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
  wps_word_proofread_basic: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_list_tool_groups: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_load_tool_group: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_unload_tool_group: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_search_tools: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  wps_route_intent: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },

  /**
   * 【六】动作名【查不到】但工具确实要改文档（1 个）+ 1 个改的是服务器状态。
   *
   * wps_ppt_set_active_target 会锁定后续所有 PPT 操作的目标文稿：它不改文档，
   * 但决定了后面每一次写入落在哪份文件上。判成只读会让客户端跳过确认。
   *
   * 第 1 期第 2 章第 1 轮移出：wps_word_set_text_color / wps_word_set_line_spacing
   * 原先也在这张表里，原因是它们调的动作名在 ps1 里查不到（setTextColor / setLineSpacing，
   * 实测零命中）—— 动作缺失是另一条 bug，已单列修掉。动作名修对之后（setFont / setParagraph），
   * 它们的注解由真动作名算出来，分类与原来手写的【逐字段相同】
   * （readOnly=false / destructive=false / idempotent=true），所以这里不再需要特例。
   */
  wps_ppt_set_active_target: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },

  /**
   * 【七】重复执行后果不同的工具（1 个）。
   *   wps_common_quit 第一次调用可能丢弃未保存内容（confirm=true 时），第二次调用却无事可做
   *   —— 两次执行的后果不同，按非幂等给。破坏性由动作分类给出（quitWps 已在 DESTRUCTIVE_EXACT）。
   *   2026-10-24 起 quitWps 不再全量 taskkill（历史行为会误杀用户自己开的 WPS 窗口），
   *   改为"未保存默认拒绝 + confirm 放行 + 优雅退出"，但"会丢未保存改动"这个性质不变。
   */
  wps_common_quit: {
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: false,
  },
};

/**
 * 算一个工具的注解。
 * 拿不到动作名（工具没在映射表里、也不在 EXTRA 里）就返回 undefined ——
 * 宁可没有注解，也不要给一个猜出来的错注解（客户端可能拿它做安全决策）。
 *
 * 旧名（别名）先解析成规范名：注解挂在规范名上，旧名必须拿到同一份 ——
 * 否则会出现"旧名没有破坏性提示"这种安全缺口，而客户端正是拿这个决定要不要弹确认。
 */
export function getToolAnnotations(toolName: string): ToolAnnotations | undefined {
  const name = resolveToolAlias(toolName);
  if (cache.has(name)) return cache.get(name);

  const extra = EXTRA_ANNOTATIONS[name];
  if (extra) {
    cache.set(name, extra);
    return extra;
  }

  const actions = TOOL_ACTIONS[name];
  if (!actions || !actions.length) {
    cache.set(name, undefined);
    return undefined;
  }

  // 一个工具可能发多条命令，只要有一条会写，它就不能算只读
  const readOnly = actions.every(isReadOnlyAction);
  const destructive = actions.some(isDestructiveAction);

  const ann: ToolAnnotations = {
    readOnlyHint: readOnly,
    // 只读的按定义不破坏；写入里只有删/丢数据/原地转换那批算破坏
    destructiveHint: readOnly ? false : destructive,
    idempotentHint: actions.every(isIdempotentAction),
    // 本 MCP 的全部副作用都在本机 WPS，不访问外部系统
    openWorldHint: false,
  };

  cache.set(name, ann);
  return ann;
}

/** 注解覆盖了多少工具（诊断用） */
export function annotationCoverage(toolNames: string[]): {
  total: number;
  annotated: number;
} {
  const annotated = toolNames.filter((n) => getToolAnnotations(n) !== undefined).length;
  return { total: toolNames.length, annotated };
}

/** 汇总分布（诊断/文档用，数字必须实测别手写） */
export function annotationSummary(toolNames: string[]): {
  total: number;
  annotated: number;
  readOnly: number;
  destructive: number;
  nonIdempotent: number;
} {
  let readOnly = 0;
  let destructive = 0;
  let nonIdempotent = 0;
  let annotated = 0;
  for (const n of toolNames) {
    const a = getToolAnnotations(n);
    if (!a) continue;
    annotated++;
    if (a.readOnlyHint) readOnly++;
    if (a.destructiveHint) destructive++;
    if (a.idempotentHint === false) nonIdempotent++;
  }
  return { total: toolNames.length, annotated, readOnly, destructive, nonIdempotent };
}