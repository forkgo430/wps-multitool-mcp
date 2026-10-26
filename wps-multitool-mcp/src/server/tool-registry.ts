/**
 * Input: Tool 定义与调用请求
 * Output: Tool 执行结果
 * Pos: MCP Tool 注册与调度中心。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tool注册管理 - Tool 管理系统
 * 所有Tool都得在这儿注册，不注册的Tool就是野鸡Tool
 * 这个设计遵循OCP原则：扩展开放，修改关闭
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  RegisteredTool,
  ToolCallRequest,
  ToolCallResult,
  ToolCategory,
  ListToolsResponse,
  DEFAULT_LOADED_GROUPS,
  ALWAYS_LOADED_GROUP,
  groupOfTool,
} from '../types/tools';
import { createChildLogger } from '../utils/logger';
import { emitToolGroupsChanged } from './tool-group-events';
import { ESSENTIAL_TOOL_NAMES, isEssentialTier, isListedByDefault } from './tool-tiers';
import { isLiveVerified } from './tool-verification';
import {
  ToolNotFoundError,
  ToolExecutionError,
  InvalidParamsError,
  McpError,
} from '../utils/error';
import { getToolMeta, expandQuery } from './tool-metadata';
import { resolveToolAlias, isToolAlias, applyAliasParams, TOOL_ALIASES } from './tool-aliases.generated';
import { withAutoSnapshot } from './auto-snapshot';
import { runWithTarget } from './target-context';

/**
 * 中文相似度用的字符 bigram 集合。
 * 为什么要字符级而不是词级：中文没有空格，"合并单元格"没法切词；而模型写错的名字
 * 往往只错一两个字，字符 bigram 对这种错字最稳。
 */
function bigramSet(text: string): Set<string> {
  const clean = (text || '').replace(/[\s\p{P}]+/gu, '');
  const out = new Set<string>();
  if (clean.length === 1) out.add(clean);
  for (let i = 0; i + 1 < clean.length; i++) out.add(clean.slice(i, i + 2));
  return out;
}

/**
 * 覆盖率：查询的 bigram 有多少能在目标里找到（0~1）。
 *
 * 为什么不用 Dice：Dice 的分母两侧长度相近才合理，而这里查询只有 4 个 bigram、
 * 工具描述有几百个 —— 分母被描述撑爆后连"一字不差"都只有 0.03，阈值根本没法定。
 * 纠偏要回答的是"模型写的这个名字，有多少能在候选里找到"，那是覆盖率的问题。
 */
function coverage(query: Set<string>, target: string): number {
  if (!query.size) return 0;
  const t = bigramSet(target);
  if (!t.size) return 0;
  let hit = 0;
  for (const g of query) if (t.has(g)) hit++;
  return hit / query.size;
}

/** 搜索命中项（比之前多了别名与正/反例） */
export interface ToolSearchHit {
  name: string;
  group: string;
  listed: boolean;
  description: string;
  /** 中文别名（模型是用中文在找，这栏是命中关键） */
  aliases?: string[];
  /** 什么样的诉求该用它（只给靠前的命中，省上下文） */
  examples?: string[];
  /** 它不干什么 / 该用谁（近名工具的分辨依据） */
  notFor?: string[];
  /** 命中的是哪一段：name / description / alias */
  matchedBy?: string;
  /**
   * 真机验证状态。static 不是"不能用"，是"没人在这台机器上跑通过"——
   * 结果要如实转述，别把"没报错"当成"做对了"。
   */
  verified?: 'live' | 'static';
}

const logger = createChildLogger('ToolRegistry');

/**
 * Tool注册表 - 单例模式，全局唯一
 */
export class ToolRegistry {
  private static instance: ToolRegistry;
  private readonly tools: Map<string, RegisteredTool>;
  private readonly categories: Map<ToolCategory, Set<string>>;
  /** 当前已加载的工具组。未加载的组不出现在 tools/list（省上下文） */
  private readonly loadedGroups: Set<string>;
  /** 基础档模式下，被【完整加载】的组（该组全部工具都列出，而不是只列基础档里的那几个） */
  private readonly fullyLoadedGroups: Set<string>;
  /** 是否处于基础档模式（默认是；WPS_TOOL_TIER=full 关闭） */
  private readonly essentialOnly: boolean;

  private constructor() {
    this.tools = new Map();
    this.categories = new Map();

    // 初始化分类
    Object.values(ToolCategory).forEach((category) => {
      this.categories.set(category, new Set());
    });

    // 初始加载的组：环境变量 WPS_TOOL_GROUPS 可覆盖（逗号分隔），默认见 DEFAULT_LOADED_GROUPS
    const fromEnv = (process.env.WPS_TOOL_GROUPS || '')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    this.loadedGroups = new Set(fromEnv.length ? fromEnv : DEFAULT_LOADED_GROUPS);
    // core 永远加载
    this.loadedGroups.add(ALWAYS_LOADED_GROUP);
    // 基础档默认开启：默认只列 40 个（core 4 + 基础档 36），其余靠 search + load 拉
    this.essentialOnly = isEssentialTier();
    this.fullyLoadedGroups = new Set();

    logger.info('ToolRegistry initialized', {
      loadedGroups: [...this.loadedGroups],
      tier: this.essentialOnly ? 'essential' : 'full',
      essentialCount: ESSENTIAL_TOOL_NAMES.size,
    });
  }

  /**
   * 获取单例实例
   */
  static getInstance(): ToolRegistry {
    if (!ToolRegistry.instance) {
      ToolRegistry.instance = new ToolRegistry();
    }
    return ToolRegistry.instance;
  }

  /**
   * 注册Tool - 把Tool加到注册表里
   * 遵循SRP：只负责注册，不负责其他
   */
  register(definition: ToolDefinition, handler: ToolHandler): void {
    const { name, category } = definition;

    // 检查是否已注册 - 跳过重复注册而非崩溃，防止过期编译产物导致启动失败
    if (this.tools.has(name)) {
      logger.warn(`Tool already registered, skipping: ${name}`);
      return;
    }

    // 注册Tool
    this.tools.set(name, { definition, handler });

    // 添加到分类
    const toolCategory = category || ToolCategory.COMMON;
    this.categories.get(toolCategory)?.add(name);

    logger.info(`Tool registered: ${name}`, { category: toolCategory });
  }

  /**
   * 批量注册Tools - 一次注册一堆
   */
  registerAll(tools: Array<{ definition: ToolDefinition; handler: ToolHandler }>): void {
    tools.forEach(({ definition, handler }) => {
      this.register(definition, handler);
    });
    logger.info(`Batch registered ${tools.length} tools`);
  }

  /**
   * 注销Tool - 把Tool从注册表里删掉
   */
  unregister(name: string): boolean {
    const tool = this.tools.get(name);
    if (!tool) {
      return false;
    }

    this.tools.delete(name);

    // 从分类中移除
    const category = tool.definition.category || ToolCategory.COMMON;
    this.categories.get(category)?.delete(name);

    logger.info(`Tool unregistered: ${name}`);
    return true;
  }

  /**
   * 获取Tool定义。
   *
   * 旧名（见 tool-aliases.generated.ts）在这里解析成规范名 —— 只解析一次、只在一个地方解析，
   * 调用方不用"先判断是不是别名再取"。别名本身不在 tools 表里，所以先查真名再查别名不会误伤。
   */
  getTool(name: string): RegisteredTool | undefined {
    return this.tools.get(name) ?? this.tools.get(resolveToolAlias(name));
  }

  /**
   * 检查Tool是否存在。
   * 旧名算存在（会解析成规范名）—— 否则 mcp-server 与 batch 的入口守卫会把旧名挡在门外。
   */
  hasTool(name: string): boolean {
    return this.tools.has(resolveToolAlias(name));
  }

  /**
   * 旧名解析结果（给"你用的名字已合并"这类提示用；不是别名则返回 undefined）。
   * 为什么不直接暴露 TOOL_ALIASES：调用方要的是"这个名字要转成谁"，不是整张表。
   */
  aliasTargetOf(name: string): string | undefined {
    return isToolAlias(name) ? TOOL_ALIASES[name] : undefined;
  }

  /**
   * 获取【已加载组】的 Tool 定义 —— MCP 的 tools/list 用的。
   * 未加载的组不出现在这里（这就是省上下文的手段）。
   */
  listTools(): ListToolsResponse {
    const tools = Array.from(this.tools.values())
      .filter((t) => this.isListed(t.definition))
      .map((t) => t.definition);
    return { tools };
  }

  /** 取全部 Tool（含未加载组）—— 供搜索/统计用，不要拿去当 tools/list */
  listAllTools(): ToolDefinition[] {
    return Array.from(this.tools.values()).map((t) => t.definition);
  }

  /**
   * 该工具当前是否会被列出。
   *   基础档模式：core 组 ∪ （基础档白名单 ∩ 真机验证过）∪ 已完整加载的组
   *   全量模式：所属组已加载
   *
   * 为什么要叠"验证过"这一层：默认档是模型最先看到的 30 多个工具，
   * 等于我们在替用户做"这些能信"的背书。没在真机上跑过的（目前是 PPT 那批）
   * 不该混在里面 —— 它们仍可搜到、仍可加载，只是不占默认名额。
   */
  isListed(def: ToolDefinition): boolean {
    const group = groupOfTool(def);
    if (!this.essentialOnly) {
      return this.loadedGroups.has(group);
    }
    if (group === ALWAYS_LOADED_GROUP) return true;
    if (isListedByDefault(def.name)) return true;
    return this.fullyLoadedGroups.has(group);
  }

  /** 当前档位：'essential'（默认，只列基础档）| 'full' */
  get tier(): 'essential' | 'full' {
    return this.essentialOnly ? 'essential' : 'full';
  }

  /** 按分类获取Tools */
  getToolsByCategory(category: ToolCategory): ToolDefinition[] {
    const toolNames = this.categories.get(category);
    if (!toolNames) {
      return [];
    }

    return Array.from(toolNames)
      .map((name) => this.tools.get(name)?.definition)
      .filter((t): t is ToolDefinition => t !== undefined);
  }

  // ==================== 工具组 ====================

  /**
   * 各组概况：
   *   total       组里一共多少工具
   *   listed      其中当前会出现在 tools/list 里的数量
   *   loaded      该组是否已加载
   *   fullyLoaded 基础档模式下，是否被"完整加载"（整组列出）
   */
  getGroups(): Array<{
    group: string;
    loaded: boolean;
    fullyLoaded: boolean;
    total: number;
    listed: number;
    /** 该组真机验证过的工具数 —— 和 total 一起看才知道"这组的工具能不能信" */
    live: number;
  }> {
    const counts = new Map<string, number>();
    const listedCounts = new Map<string, number>();
    const liveCounts = new Map<string, number>();
    for (const t of this.tools.values()) {
      const g = groupOfTool(t.definition);
      counts.set(g, (counts.get(g) ?? 0) + 1);
      if (this.isListed(t.definition)) listedCounts.set(g, (listedCounts.get(g) ?? 0) + 1);
      if (isLiveVerified(t.definition.name)) liveCounts.set(g, (liveCounts.get(g) ?? 0) + 1);
    }
    // core 组一定存在（哪怕暂时没有工具）
    if (!counts.has(ALWAYS_LOADED_GROUP)) counts.set(ALWAYS_LOADED_GROUP, 0);

    return Array.from(counts.entries())
      .map(([group, total]) => ({
        group,
        loaded: this.loadedGroups.has(group),
        fullyLoaded: this.fullyLoadedGroups.has(group),
        total,
        listed: listedCounts.get(group) ?? 0,
        live: liveCounts.get(group) ?? 0,
      }))
      .sort((a, b) => a.group.localeCompare(b.group));
  }

  getLoadedGroups(): string[] {
    return [...this.loadedGroups];
  }

  /** 加载一个组。返回是否真的发生了变化（用于决定要不要发 list_changed 通知） */
  loadGroup(name: string): { ok: boolean; changed: boolean; toolCount: number; error?: string } {
    const exists = this.getGroups().some((g) => g.group === name);
    if (!exists) {
      return { ok: false, changed: false, toolCount: 0, error: `没有名为 "${name}" 的工具组` };
    }
    const changed = this.essentialOnly
      ? !this.fullyLoadedGroups.has(name)
      : !this.loadedGroups.has(name);
    this.loadedGroups.add(name);
    if (this.essentialOnly) this.fullyLoadedGroups.add(name);
    const toolCount = this.getGroups().find((g) => g.group === name)?.total ?? 0;
    if (changed) {
      logger.info(`Tool group loaded: ${name}`, { toolCount, tier: this.tier });
      emitToolGroupsChanged();
    }
    return { ok: true, changed, toolCount };
  }

  /** 卸载一个组。core 不允许卸载（否则连加载工具都找不到了） */
  unloadGroup(name: string): { ok: boolean; changed: boolean; toolCount: number; error?: string } {
    if (name === ALWAYS_LOADED_GROUP) {
      return {
        ok: false,
        changed: false,
        toolCount: 0,
        error: `"${ALWAYS_LOADED_GROUP}" 组包含工具组管理本身，不允许卸载`,
      };
    }
    const exists = this.getGroups().some((g) => g.group === name);
    if (!exists) {
      return { ok: false, changed: false, toolCount: 0, error: `没有名为 "${name}" 的工具组` };
    }
    const changed = this.essentialOnly
      ? this.fullyLoadedGroups.has(name)
      : this.loadedGroups.has(name);
    this.loadedGroups.delete(name);
    this.fullyLoadedGroups.delete(name);
    const toolCount = this.getGroups().find((g) => g.group === name)?.total ?? 0;
    if (changed) {
      logger.info(`Tool group unloaded: ${name}`, { toolCount, tier: this.tier });
      emitToolGroupsChanged();
    }
    return { ok: true, changed, toolCount };
  }

  /**
   * 搜索工具（含未加载组）—— 这是"用到再拉"的关键：
   * 组没加载时工具不在列表里，但可以先搜出名字再 load_tool_group。
   */
  searchTools(
    keyword: string,
    opts: { group?: string; limit?: number; hints?: boolean } = {}
  ): ToolSearchHit[] {
    const kw = (keyword || '').trim().toLowerCase();
    const limit = opts.limit && opts.limit > 0 ? opts.limit : 20;
    // 正/反例默认只给前 5 条命中：285 个工具全带上会撑爆上下文
    const hintTop = opts.hints === false ? 0 : 5;
    const out: Array<ToolSearchHit & { score: number; matchedBy: string }> = [];
    // 查询扩展（第 9 轮）：同义词/反查词根也算命中，但分比 alias 还低一档 ——
    // 扩展词本身就是猜的，必须排在人写的描述之后。matchedBy 单独标 synonym，
    // 这样搜索结果自带解释（"这条是靠扩展词命中的"），便于门禁和排障。
    const terms = expandQuery(kw);
    const isSynonymOnly = terms.length > 1;

    for (const t of this.tools.values()) {
      const def = t.definition;
      const group = groupOfTool(def);
      if (opts.group && group !== opts.group) continue;
      const name = def.name.toLowerCase();
      const desc = (def.description || '').toLowerCase();
      const meta = getToolMeta(def);
      const aliasBlob = meta.aliases.join(' ').toLowerCase();

      let score = 0;
      let matchedBy = '';
      if (!kw) {
        score = 1;
        matchedBy = 'all';
      } else {
        // 原词按原权重匹配
        if (name === kw) {
          score = 100;
          matchedBy = 'name';
        } else if (name.includes(kw)) {
          score = 60;
          matchedBy = 'name';
        } else if (desc.includes(kw)) {
          score = 40;
          matchedBy = 'description';
        } else if (aliasBlob.includes(kw)) {
          // 别名是【机器拼出来的】，质量不如人写的描述，所以分排最低：
          // 实测 66 个中文关键词里只有 17 个靠别名多捞到东西，别让它压过真命中。
          score = 25;
          matchedBy = 'alias';
        }
        // 原词没命中，才让扩展词上场（不会覆盖真命中，只补空缺）
        if (score === 0 && isSynonymOnly) {
          for (const term of terms) {
            if (term === kw) continue;
            if (name === term || name.includes(term)) {
              score = 20;
              matchedBy = 'synonym';
              break;
            }
            if (desc.includes(term)) {
              score = 18;
              matchedBy = 'synonym';
              break;
            }
            if (aliasBlob.includes(term)) {
              score = 12;
              matchedBy = 'synonym';
              break;
            }
          }
        }
      }
      if (score === 0) continue;

      out.push({
        name: def.name,
        group,
        listed: this.isListed(def),
        description: (def.description || '').split('\n')[0].slice(0, 120),
        aliases: meta.aliases.slice(0, 5),
        matchedBy,
        verified: isLiveVerified(def.name) ? 'live' : 'static',
        score,
      });
    }

    // 同分时不能只按名字排（第 10 轮 T3b）。
    // 同分意味着"跟这句诉求的相关度差不多"，这时候真正该拉开差距的是
    // **能不能直接用**和**验证过没有**：
    //   - 已列入 tools/list 的（默认档/已加载组）比没列入的更可能被直接调用
    //   - live 的比 static 的更可信
    // 原来只按 name 字母序，于是 static、未列入的工具会和 listed+live 的混排，
    // 模型看到的前几条反而是不该首选的那批。
    // 口径复用现成判定：isListed() 与 isLiveVerified()，不另写一套。
    return out
      .sort(
        (a, b) =>
          b.score - a.score ||
          Number(b.listed) - Number(a.listed) ||
          Number(b.verified === 'live') - Number(a.verified === 'live') ||
          a.name.localeCompare(b.name)
      )
      .slice(0, limit)
      .map((h, i) => {
        const { score: _score, ...rest } = h;
        // 前 N 条补正/反例；没有就不带这个字段（别塞空数组）
        if (i >= hintTop) return rest;
        const meta = getToolMeta(this.tools.get(h.name)!.definition);
        return {
          ...rest,
          ...(meta.examples?.length ? { examples: meta.examples } : {}),
          ...(meta.notFor?.length ? { notFor: meta.notFor } : {}),
        };
      });
  }

  /**
   * 调用未知名工具时给个线索：是不是"组没加载"导致的？
   * 返回最近似的几个工具名与它们所在组。
   */
  suggestSimilar(name: string, limit = 3): Array<{ name: string; group: string; loaded: boolean }> {
    const near = this.scoreSimilar(name, limit);

    // 旧名直接给答案：它已经不是"像哪个"的问题，而是"就是哪个"。
    // 排在第一位，后面的近似项照旧给（旧名残片可能还指向别的工具，删掉反而少信息）。
    const canonical = resolveToolAlias(name);
    if (canonical !== name) {
      const hit = this.tools.get(canonical);
      if (hit) {
        const group = groupOfTool(hit.definition);
        return [
          { name: canonical, group, loaded: this.loadedGroups.has(group) },
          ...near.filter((n) => n.name !== canonical),
        ].slice(0, limit);
      }
    }
    return near;
  }

  /** 按名字词元算近似度（suggestSimilar 的打分部分，拆出来是为了别名那条路能复用） */
  private scoreSimilar(
    name: string,
    limit: number
  ): Array<{ name: string; group: string; loaded: boolean }> {
    // 中文也得能纠偏（第 9 轮）：原来只切 ASCII 词元，模型写"合并单元格"或"设置背景色"
    // 这种错名时直接返回空 —— 模型拿不到任何线索，只能瞎猜工具名。
    // 中文没有空格可切，用字符 bigram 的 Dice 系数（0~1）；英文走原来的词元命中数。
    const query = (name || '').trim().toLowerCase();
    const tokens = query.split(/[^a-z0-9]+/).filter((t) => t.length > 2);
    const zhQuery = /[一-龥]/.test(query);
    if (!tokens.length && !zhQuery) return [];
    const bigrams = zhQuery ? bigramSet(query) : new Set<string>();

    const scored: Array<{
      name: string; group: string; loaded: boolean; score: number; listed: boolean; live: boolean;
    }> = [];
    for (const t of this.tools.values()) {
      const def = t.definition;
      const lower = def.name.toLowerCase();
      const meta = getToolMeta(def);
      // 中文要跟【中文面】比：工具名是英文，光比 def.name 一个中文都匹配不上。
      // 只取别名 + 描述首段：描述全文动辄几百字，算出来永远是"什么都像一点"。
      const zhAlias = coverage(bigrams, meta.aliases.join(' '));
      const zhHead = coverage(bigrams, (def.description || '').split('\n').slice(0, 2).join(' '));
      const score = zhQuery
        ? Math.round(Math.max(zhAlias, zhHead * 0.9) * 100)
        : tokens.reduce((acc, tok) => (lower.includes(tok) ? acc + 1 : acc), 0);
      const group = groupOfTool(def);
      if (score > 0) {
        scored.push({
          name: def.name,
          group,
          loaded: this.loadedGroups.has(group),
          score,
          // 同分排序的两个辅键，口径与 searchTools 完全一致（不许另写一套）：
          //   建议"调不上的名字"时，模型下一步要么直接调用（那要 listed），
          //   要么读它的 verified 字段再决定信不信。listed=false 的先靠后。
          listed: this.isListed(def),
          live: isLiveVerified(def.name),
        });
      }
    }
    // 中文是 0~100 的连续覆盖率，55% 意味着"模型写的名字大半能在候选里对上"；
    // 再低就是噪声，列出来只会误导模型（比不给建议更糟）。
    const min = zhQuery ? 55 : 0;
    return scored
      .filter((s) => s.score >= min)
      // 同分四元：score → listed → live → name（第 10 轮 T3b，与 searchTools 同口径）
      .sort(
        (a, b) =>
          b.score - a.score ||
          Number(b.listed) - Number(a.listed) ||
          Number(b.live) - Number(a.live) ||
          a.name.localeCompare(b.name)
      )
      .slice(0, limit)
      .map(({ score: _score, listed: _listed, live: _live, ...rest }) => rest);
  }

  /**
   * 调用Tool - 执行Tool的handler
   * 这是核心方法，处理Tool调用请求
   */
  async callTool(request: ToolCallRequest): Promise<ToolCallResult> {
    const { id, name: requestedName, arguments: args } = request;
    const startTime = Date.now();

    // 旧名在这里收敛成规范名：注解、验证状态、warnings、日志【全部】挂在规范名上。
    // 若在这里分叉（旧名单独跑一套），两套必然漂，而漂的那一套通常是没人看的那套。
    const name = resolveToolAlias(requestedName);
    const viaAlias = name !== requestedName;

    logger.debug(`Calling tool: ${requestedName}${viaAlias ? ` (alias → ${name})` : ''}`, { id, args });

    // 检查Tool是否存在
    const tool = this.tools.get(name);
    if (!tool) {
      throw new ToolNotFoundError(requestedName);
    }

    try {
      // 旧名的参数要在【校验之前】搬：族合并里旧参数名（preset/newOrder/...）对新工具是必填项，
      // 先校验再搬 = 旧调用必报 Missing required parameter。先搬再校验，旧调用就真的无损了。
      const callArgs: Record<string, unknown> = viaAlias
        ? applyAliasParams(requestedName, args || {})
        : args || {};

      // 验证参数
      this.validateArguments(tool.definition, callArgs);

      // 执行handler —— 外面裹一层"可退"保险：破坏性/非幂等的工具在动手前自动存快照。
      // 防重与失败降级都在 withAutoSnapshot 里（嵌套调用不再建、建不了只进 warnings）。
      // 挂在这里而不是各 handler：这是所有工具调用的唯一入口，296 个工具逐个加必漏。
      // target 也要在这一层兜住：67 个 word handler 都只挑自己认识的字段，
      // 不会把 args.target 转发下去（那是好事，handler 不该管路由）。
      // 用 ALS 而不是全局变量：callTool 并发可达，全局变量会把 A 的 target 漂给 B。
      const result = await runWithTarget(callArgs.target, () =>
        withAutoSnapshot(name, () => tool.handler(callArgs))
      );

      const duration = Date.now() - startTime;
      logger.info(`Tool executed: ${name}`, { id, duration, success: result.success });

      return {
        ...result,
        id,
      };
    } catch (error) {
      const duration = Date.now() - startTime;

      if (error instanceof McpError) {
        // 旧名调用最常见的失败不是"工具没了"，而是"参数名也一起变了"。
        // 只说 Missing required parameter: startRow，调用方会以为是新工具有毛病 ——
        // 所以这里补一句"你用的是旧名、它已合并到谁、现在的必填项是什么"，让人能直接改对。
        const message =
          viaAlias && error instanceof InvalidParamsError
            ? `${error.message}\n旧名 ${requestedName} 已合并为 ${name}（别名，仍可调用），` +
              `它现在的必填参数：${(tool.definition.inputSchema.required || []).join(', ') || '无'}。` +
              '旧参数名不再生效，请按规范名的参数表重发一次；不要重试原参数。'
            : error.message;

        logger.error(
          `Tool execution failed: ${requestedName}${viaAlias ? ` (alias → ${name})` : ''}`,
          error,
          { id, duration }
        );
        return {
          id,
          success: false,
          content: [{ type: 'text', text: message }],
          error: message,
          errorType: error.wpsType,
          retryable: error.retryable,
        };
      }

      const execError = new ToolExecutionError(
        name,
        error instanceof Error ? error : new Error(String(error))
      );

      logger.error(`Tool execution failed: ${name}`, execError, { id, duration });

      return {
        id,
        success: false,
        content: [{ type: 'text', text: execError.message }],
        error: execError.message,
        errorType: execError.wpsType,
        retryable: execError.retryable,
      };
    }
  }

  /**
   * 验证参数 - 检查必填参数是否都有
   */
  private validateArguments(
    definition: ToolDefinition,
    args: Record<string, unknown>
  ): void {
    const { inputSchema } = definition;
    const required = inputSchema.required || [];

    for (const param of required) {
      if (!(param in args) || args[param] === undefined || args[param] === null) {
        throw new InvalidParamsError(`Missing required parameter: ${param}`, {
          toolName: definition.name,
          missingParam: param,
        });
      }
    }

    // TODO: 可以加更详细的类型验证，但现在先这样，够用
  }

  /**
   * 获取Tool数量
   */
  get size(): number {
    return this.tools.size;
  }

  /**
   * 清空所有Tools - 测试用，生产环境别乱用
   */
  clear(): void {
    this.tools.clear();
    this.categories.forEach((set) => set.clear());
    logger.warn('All tools cleared');
  }

  /**
   * 创建Tool调用请求 - 辅助方法
   */
  static createRequest(
    name: string,
    args: Record<string, unknown>
  ): ToolCallRequest {
    return {
      id: uuidv4(),
      name,
      arguments: args,
    };
  }
}

// 导出单例
export const toolRegistry = ToolRegistry.getInstance();

/**
 * 装饰器：注册Tool
 * 用法：@RegisterTool(definition)
 */
export function RegisterTool(definition: ToolDefinition) {
  return function (
    _target: unknown,
    _propertyKey: string,
    descriptor: PropertyDescriptor
  ): PropertyDescriptor {
    const originalMethod = descriptor.value as ToolHandler;
    toolRegistry.register(definition, originalMethod);
    return descriptor;
  };
}

/**
 * 快捷注册函数
 */
export const registerTool = (
  definition: ToolDefinition,
  handler: ToolHandler
): void => {
  toolRegistry.register(definition, handler);
};

export default ToolRegistry;
