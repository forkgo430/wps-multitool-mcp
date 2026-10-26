/**
 * Input: MCP 协议请求
 * Output: Tool 调用响应
 * Pos: MCP Server 核心实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * MCP Server实现 - WPS 的 MCP 协议实现
 * 这是整个系统的核心，处理MCP协议通信
 * 基于stdio传输，Claude Desktop就是这么连的
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  ListPromptsRequestSchema,
  GetPromptRequestSchema,
  ErrorCode as McpErrorCode,
  McpError as SdkMcpError,
} from '@modelcontextprotocol/sdk/types.js';

import { toolRegistry, ToolRegistry } from './tool-registry';
import { wpsClient } from '../client/wps-client';
import { ToolCallResult, ToolCategory, ToolContent } from '../types/tools';
import { allTools } from '../tools';
import { onToolGroupsChanged } from '../tools/common/toolgroups';
import { createChildLogger } from '../utils/logger';
import { McpError, formatWpsErrorTag, WpsErrorType } from '../utils/error';
import { runWithCancel } from './cancel-context';
import { getToolAnnotations } from './tool-annotations';
import { runWithProgress } from './progress-context';
import { getToolMeta } from './tool-metadata';
import { serverInstructions, setToolCount } from './server-info';
import { CancelError } from '../bridge/cancellable';
import { WpsAppType } from '../types/wps';
import { MCP_NAME, MCP_VERSION } from '../version';

const logger = createChildLogger('McpServer');

/**
 * MCP Server配置
 */
export interface McpServerConfig {
  /** 服务器名称 */
  name: string;
  /** 服务器版本 */
  version: string;
  /** 是否启用调试模式 */
  debug?: boolean;
}

/**
 * 默认配置
 */
const DEFAULT_CONFIG: McpServerConfig = {
  name: MCP_NAME,
  // 版本号只有 package.json 一个真相源，见 src/version.ts
  version: MCP_VERSION,
  debug: false,
};

/**
 * WPS MCP Server - 核心服务器类
 */
export class WpsMcpServer {
  private readonly config: McpServerConfig;
  private readonly server: Server;
  private readonly registry: ToolRegistry;
  private isRunning: boolean = false;
  /** 工具组变更通知是否已接上（避免重复订阅） */
  private groupsNotifyWired: boolean = false;

  private static dataCache: Map<string, { data: unknown; timestamp: number; appType: string }> = new Map();

  constructor(config?: Partial<McpServerConfig>) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.registry = toolRegistry;

    // 工具注册放在【构造期】而不是 start() 里 —— 因为 instructions（客户端第一眼
    // 看到的介绍，含作者署名与真实工具数）在 new Server() 时就要定下来。
    // registerBuiltinTools() 只依赖 this.registry，不碰 this.server，提前调用安全。
    this.registerBuiltinTools();
    this.registry.registerAll(allTools);
    // 介绍里的工具数从注册表数，不写死 —— 写死必随版本腐化
    const toolCount = this.registry.listAllTools().length;
    setToolCount(toolCount);
    logger.info(`Registered ${toolCount} tools in total (builtin + professional)`);

    // 创建MCP Server实例
    this.server = new Server(
      {
        name: this.config.name,
        version: this.config.version,
      },
      {
        capabilities: {
          tools: {},
          resources: {},
          prompts: {},
        },
        // 介绍 + 作者署名：随 initialize 响应发给客户端。
        instructions: serverInstructions(),
      }
    );

    // 注册请求处理器
    this.setupRequestHandlers();
    this.setupResourceAndPromptHandlers();

    // 错误处理
    this.server.onerror = (error) => {
      logger.error('MCP Server error', error);
    };

    logger.info('MCP Server created', {
      name: this.config.name,
      version: this.config.version,
    });
  }

  /**
   * Resources 与 Prompts（2026-10-26 对标头部 MCP 项目补齐的协议面）
   *
   * Resources：把"当前打开的文档"暴露为资源——客户端（Claude Desktop 等）可以直接浏览
   *   与读取，不必先调工具。列表实时查询三应用的打开文档；读失败的某应用在 JSON 里如实标注。
   * Prompts：四个场景模板（周报/公文/合同/论文），把 skill 的五步链路预编排成用户可选的
   *   一等条目。失败语义与工具一致：Linux 等无通道环境抛 BackendUnavailableError。
   */
  private setupResourceAndPromptHandlers(): void {
    // Resources：当前打开的文档
    this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
      return {
        resources: [
          {
            uri: 'wps://documents',
            name: '当前打开的文档',
            description: 'WPS 文字/表格/演示里此刻打开的全部文档（实时查询）',
            mimeType: 'application/json',
          },
        ],
        // 资源模板：客户端填入应用类型与文档名后，可直接读取该文档的正文
        resourceTemplates: [
          {
            uriTemplate: 'wps://doc/{app}/{name}',
            name: '打开的文档内容',
            description: '按应用类型与文档名读取一份已打开文档的正文文本',
            mimeType: 'text/plain',
            examples: ['wps://doc/word/报告.docx', 'wps://doc/excel/数据.xlsx'],
          },
        ],
      };
    });

    this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const uri = request.params.uri;

      // 静态资源：三应用打开文档汇总
      if (uri === 'wps://documents') {
        const apps: Array<{ app: 'word' | 'excel' | 'ppt'; appType: WpsAppType; method: string }> = [
          { app: 'word', appType: WpsAppType.WRITER, method: 'getOpenDocuments' },
          { app: 'excel', appType: WpsAppType.SPREADSHEET, method: 'getOpenWorkbooks' },
          { app: 'ppt', appType: WpsAppType.PRESENTATION, method: 'getOpenPresentations' },
        ];
        const out: Record<string, unknown> = {};
        for (const { app, appType, method } of apps) {
          try {
            const r = await wpsClient.executeMethod(method, {}, appType);
            out[app] = r.data ?? null;
          } catch (e) {
            out[app] = { error: e instanceof Error ? e.message : String(e) };
          }
        }
        return {
          contents: [{ uri, mimeType: 'application/json', text: JSON.stringify(out, null, 2) }],
        };
      }

      // 资源模板：wps://doc/{app}/{name} → 读该文档的正文
      const docMatch = uri.match(/^wps:\/\/doc\/(word|excel|ppt)\/(.+)$/);
      if (docMatch) {
        const [, app, name] = docMatch;
        const appType = app === 'word' ? WpsAppType.WRITER : app === 'excel' ? WpsAppType.SPREADSHEET : WpsAppType.PRESENTATION;
        const method = app === 'word' ? 'getDocumentText' : app === 'excel' ? 'getRangeData' : 'getSlideCount';
        try {
          const decoded = decodeURIComponent(name);
          const r = await wpsClient.executeMethod(method, { documentName: decoded }, appType);
          return {
            contents: [{ uri, mimeType: 'text/plain', text: JSON.stringify(r.data ?? r, null, 2) }],
          };
        } catch (e) {
          throw new SdkMcpError(McpErrorCode.InternalError, `读取文档失败: ${e instanceof Error ? e.message : String(e)}`);
        }
      }

      throw new SdkMcpError(McpErrorCode.InvalidRequest, `未知资源: ${uri}`);
    });

    // Prompts：场景模板
    const scenario = (
      name: string,
      title: string,
      description: string,
      argsSpec: Array<{ name: string; description: string }>,
      render: (a: Record<string, string>) => string
    ) => ({ name, title, description, arguments: argsSpec, render });

    const prompts = [
      scenario('wps-weekly-report', '周报生成', '把表格里的数据整理成一页周报文档', [
        { name: 'sheet', description: '数据所在的工作簿/表名（可省略，默认当前）' },
        { name: 'focus', description: '本周重点，一句话' },
      ], (a) => `我在 WPS 表格里有本周的工作数据${a.sheet ? `（${a.sheet}）` : ''}。请：1) 探活并读取数据；2) 提炼三到五条要点${a.focus ? `，围绕「${a.focus}」展开` : ''}；3) 在新的 WPS 文字文档里生成一页周报（标题加粗、要点分级）；4) 导出 PDF 给我。动笔前先让我确认要点清单。`),
      scenario('wps-official-doc', '公文排版', '按中文公文规范重排一份 WPS 文字文档', [
        { name: 'doc', description: '文档名（默认当前活动文档）' },
      ], (a) => `请把${a.doc ? `「${a.doc}」` : '当前打开的公文'}按党政机关公文格式重排：标题二号小标宋居中、正文三号仿宋、一级标题黑体、二级标题楷体、行距 28 磅。重排前先建快照；完成后逐项报告改了哪些格式。`),
      scenario('wps-contract-review', '合同修订', '以修订模式逐条修改合同，全部改动可审阅', [
        { name: 'doc', description: '合同文档名（默认当前活动文档）' },
        { name: 'changes', description: '要改的要点，每行一条' },
      ], (a) => `请先开启修订模式，再把${a.doc ? `「${a.doc}」` : '这份合同'}按以下要点逐条修改：\n${a.changes || '（等我说改动要点）'}\n每条改动走修订留痕，全部完成后给出改动清单；不要接受任何修订——留给我审。`),
      scenario('wps-academic-paper', '学术论文排版', '给论文生成目录、交叉引用与参考文献整理', [
        { name: 'doc', description: '论文文档名（默认当前活动文档）' },
      ], (a) => `请给${a.doc ? `「${a.doc}」` : '当前这篇论文'}做学术排版：1) 按标题层级生成目录；2) 为"如图 X 所示"类表述补交叉引用；3) 整理文末参考文献为规范格式；4) 全文统一字体字号。动格式前先建快照。`),
    ];

    this.server.setRequestHandler(ListPromptsRequestSchema, async () => ({
      prompts: prompts.map(({ name, title, description, arguments: args }) => ({
        name,
        title,
        description,
        arguments: args.map(({ name: an, description: ad }) => ({ name: an, description: ad })),
      })),
    }));

    this.server.setRequestHandler(GetPromptRequestSchema, async (request) => {
      const prompt = prompts.find((p) => p.name === request.params.name);
      if (!prompt) {
        throw new SdkMcpError(McpErrorCode.InvalidRequest, `未知模板: ${request.params.name}`);
      }
      const args = (request.params.arguments || {}) as Record<string, string>;
      return {
        description: prompt.description,
        messages: [
          { role: 'user' as const, content: { type: 'text', text: prompt.render(args) } },
        ],
      };
    });
  }

  /**
   * 设置请求处理器
   */
  private setupRequestHandlers(): void {
    // 处理 tools/list 请求
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      logger.debug('Handling tools/list request');

      // 只返回【已加载组】里的工具；未加载组的不列出（省上下文，见 wps_list_tool_groups）
      const { tools } = this.registry.listTools();
      const total = this.registry.listAllTools().length;

      logger.info(`Returning ${tools.length} tools (已加载组) / 共 ${total}`, {
        loadedGroups: this.registry.getLoadedGroups(),
      });

      return {
        tools: tools.map((tool) => {
          const annotations = getToolAnnotations(tool.name);
          // examples 是【独立字段】，不拼进 description：
          //   ① 搜索路径把 description 截成首行 120 字（tool-registry.ts），拼进去会被吃掉；
          //   ② 拼进描述会污染模型的工具选择（描述是"这工具能干啥"，不是"用户怎么说"）。
          // 定义上没写 examples 的，优先退回 CURATED 表（同 searchTools 的口径，两条路径一致）。
          const meta = getToolMeta(tool);
          const examples = tool.examples?.length ? tool.examples : meta.examples;
          return {
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema,
            // 注解必须【嵌套】在 annotations 里。铺在顶层客户端会整个忽略 ——
            // 这坑踩过：单测全绿，端到端才发现 40 个工具一个注解都没透出去。
            ...(annotations ? { annotations } : {}),
            ...(examples?.length ? { examples } : {}),
          };
        }),
      };
    });

    // 处理 tools/call 请求
    this.server.setRequestHandler(CallToolRequestSchema, async (request, extra) => {
      const { name, arguments: args } = request.params;

      logger.debug('Handling tools/call request', { name, args });

      // 客户端发 notifications/cancelled 时 extra.signal 会 abort。
      // 把它放进 AsyncLocalStorage，调用链深处（wps-client → 桥 → 子进程）自己取，
      // 不用改 285 个工具的签名。真正"能不能取消"由 bridge/cancellable.ts 裁决。
      const signal = extra?.signal;
      if (signal) {
        signal.addEventListener(
          'abort',
          () => logger.info('Tool call cancelled by client', { name }),
          { once: true }
        );
      }
      const runIt = () => this.handleToolCall(name, args ?? {});
      const execute = signal
        ? () => runWithCancel({ signal, toolName: name }, runIt)
        : runIt;

      // 进度通知（2026-10-26 对标）：客户端传 progressToken 时，长操作发 notifications/progress
      const meta = (request.params as Record<string, unknown> & { _meta?: Record<string, unknown> })._meta;
      const progressToken = meta?.progressToken;
      const sendNotification = extra?.sendNotification;

      if (progressToken && typeof sendNotification === 'function') {
        return runWithProgress(
          (p: number, t: number, msg?: string) => {
            sendNotification({
              method: 'notifications/progress' as const,
              params: { progressToken, progress: p, total: t, ...(msg ? { message: msg } : {}) },
            } as never);
          },
          execute
        );
      }

      return execute();
    });
  }

  /** tools/call 的实际处理（拆出来是为了好包取消上下文） */
  private async handleToolCall(
    name: string,
    args: Record<string, unknown>
  ): Promise<{ content: ToolContent[]; isError: boolean }> {
    {
      // 检查Tool是否存在。不存在时给个线索：是不是"组没加载"？
      if (!this.registry.hasTool(name)) {
        const similar = this.registry.suggestSimilar(name);
        const hint = similar.length
          ? ` 近似工具：${similar
              .map((s) => `${s.name}（${s.loaded ? '已列出' : `组 ${s.group} 未加载`}）`)
              .join('、')}`
          : ' 可用 wps_search_tools 搜一下，或 wps_list_tool_groups 看有哪些组';
        throw new SdkMcpError(
          McpErrorCode.MethodNotFound,
          `Unknown tool: ${name}。${hint}`
        );
      }

      try {
        // 创建调用请求并执行
        const callRequest = ToolRegistry.createRequest(name, args || {});
        const result: ToolCallResult = await this.registry.callTool(callRequest);

        // 失败时把错误类型 + 可重试性透出去（前缀形式，不改动 isError 语义）
        // 形如 [busy|可重试·稍后退避] / [timeout|勿重试·可能已执行一半]
        const content = result.success
          ? prefixWithWarnings(result)
          : result.content.map((c, i) =>
              i === 0 && c.type === 'text' && typeof c.text === 'string' && result.errorType
                ? {
                    ...c,
                    text: `${formatWpsErrorTag(result.errorType, result.retryable)} ${c.text}`,
                  }
                : c
            );

        return {
          content,
          isError: !result.success,
          // 结构化输出（2026-10-26 对标）：工具文本里如果是 JSON 对象，同步给 structuredContent
          // 让客户端拿到结构化数据而非纯文本。仅试探首块，不递归。
          ...(() => {
            const first = result.content.find((c) => c.type === 'text');
            if (!first || first.type !== 'text') return {};
            const trimmed = (first as { text: string }).text.trim();
            if (!trimmed.startsWith('{') && !trimmed.startsWith('[')) return {};
            try {
              const parsed = JSON.parse(trimmed);
              if (parsed && typeof parsed === 'object') return { structuredContent: parsed };
            } catch { }
            return {};
          })(),
        };
      } catch (error) {
        // 取消不是故障：单独走一条，别记成 error 吓唬人
        if (error instanceof CancelError) {
          logger.info('Tool call cancelled', { name, reason: error.message });
          return {
            content: [
              {
                type: 'text',
                text: `${formatWpsErrorTag(WpsErrorType.CANCELLED, false)} ${error.message}`,
              },
            ],
            isError: true,
          };
        }

        logger.error('Tool call failed', error);

        if (error instanceof McpError) {
          return {
            content: [
              {
                type: 'text',
                text: `${formatWpsErrorTag(error.wpsType, error.retryable)} ${error.message}`,
              },
            ],
            isError: true,
          };
        }

        throw new SdkMcpError(
          McpErrorCode.InternalError,
          error instanceof Error ? error.message : 'Unknown error'
        );
      }
    }
  }

  /**
   * 注册内置Tools - 一些基础的WPS操作Tool
   */
  registerBuiltinTools(): void {
    logger.info('Registering built-in tools');

    // 连接状态检查Tool
    this.registry.register(
      {
        name: 'wps_check_connection',
        description: '检查WPS Office连接状态',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const connected = await wpsClient.checkConnection();
        const status = wpsClient.getStatus();

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                connected,
                status,
              }),
            },
          ],
        };
      }
    );

    // 获取当前工作簿信息
    this.registry.register(
      {
        name: 'wps_get_active_workbook',
        description: '获取当前打开的WPS表格工作簿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.SPREADSHEET,
      },
      async () => {
        const workbook = await wpsClient.getActiveWorkbook();

        return {
          id: '',
          success: workbook !== null,
          content: [
            {
              type: 'text',
              text: workbook
                ? JSON.stringify(workbook)
                : '没有打开的工作簿',
            },
          ],
        };
      }
    );

    // 获取当前演示文稿信息
    this.registry.register(
      {
        name: 'wps_get_active_presentation',
        description: '获取当前打开的WPS演示文稿信息',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.PRESENTATION,
      },
      async () => {
        const presentation = await wpsClient.getActivePresentation();

        return {
          id: '',
          success: presentation !== null,
          content: [
            {
              type: 'text',
              text: presentation
                ? JSON.stringify(presentation)
                : '没有打开的演示文稿',
            },
          ],
        };
      }
    );

    // 执行自定义WPS方法
    this.registry.register(
      {
        name: 'wps_execute_method',
        description: '执行自定义WPS API方法',
        inputSchema: {
          type: 'object',
          properties: {
            method: {
              type: 'string',
              description: 'API方法名',
            },
            params: {
              type: 'object',
              description: '方法参数',
            },
            appType: {
              type: 'string',
              description: '应用类型：wps（文字）、et（表格）、wpp（演示）',
              enum: ['wps', 'et', 'wpp'],
            },
          },
          required: ['method'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const method = args.method as string;
        const params = args.params as Record<string, unknown> | undefined;
        const appType = args.appType as string | undefined;

        const response = await wpsClient.executeMethod(
          method,
          params,
          appType as any
        );

        return {
          id: '',
          success: response.success,
          content: [
            {
              type: 'text',
              text: JSON.stringify(response),
            },
          ],
        };
      }
    );

    // ==================== 跨应用数据缓存工具 ====================
    // Excel读取数据 → 缓存到MCP Server → PPT获取缓存 → 创建演示文稿

    // 缓存数据
    this.registry.register(
      {
        name: 'wps_cache_data',
        description: '缓存数据到MCP Server，用于跨应用数据传递。例如：从Excel读取数据后缓存，然后在PPT中使用。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名，用于后续获取数据',
            },
            data: {
              type: 'object',
              description: '要缓存的数据（任意JSON对象）',
            },
            appType: {
              type: 'string',
              description: '数据来源应用类型：et（表格）、wps（文字）、wpp（演示）',
              enum: ['et', 'wps', 'wpp'],
            },
          },
          required: ['key', 'data'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const data = args.data;
        const appType = (args.appType as string) || 'unknown';

        WpsMcpServer.dataCache.set(key, {
          data,
          timestamp: Date.now(),
          appType,
        });

        logger.info(`Data cached: key=${key}, appType=${appType}`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                message: `数据已缓存，可在其他应用中通过 wps_get_cached_data 获取`,
                cacheSize: WpsMcpServer.dataCache.size,
              }),
            },
          ],
        };
      }
    );

    // 获取缓存数据
    this.registry.register(
      {
        name: 'wps_get_cached_data',
        description: '从MCP Server获取缓存的数据，用于跨应用数据传递。',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '缓存键名',
            },
          },
          required: ['key'],
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string;
        const cached = WpsMcpServer.dataCache.get(key);

        if (!cached) {
          return {
            id: '',
            success: false,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: false,
                  error: `缓存键 "${key}" 不存在`,
                  availableKeys: Array.from(WpsMcpServer.dataCache.keys()),
                }),
              },
            ],
          };
        }

        logger.info(`Cache hit: key=${key}, age=${Date.now() - cached.timestamp}ms`);

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                key,
                data: cached.data,
                appType: cached.appType,
                cachedAt: new Date(cached.timestamp).toISOString(),
              }),
            },
          ],
        };
      }
    );

    // 列出所有缓存
    this.registry.register(
      {
        name: 'wps_list_cache',
        description: '列出MCP Server中所有缓存的数据键名',
        inputSchema: {
          type: 'object',
          properties: {},
        },
        category: ToolCategory.COMMON,
      },
      async () => {
        const cacheList = Array.from(WpsMcpServer.dataCache.entries()).map(([key, value]) => ({
          key,
          appType: value.appType,
          cachedAt: new Date(value.timestamp).toISOString(),
          ageMs: Date.now() - value.timestamp,
        }));

        return {
          id: '',
          success: true,
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: true,
                count: cacheList.length,
                caches: cacheList,
              }),
            },
          ],
        };
      }
    );

    // 清除缓存
    this.registry.register(
      {
        name: 'wps_clear_cache',
        description: '清除MCP Server中的缓存数据',
        inputSchema: {
          type: 'object',
          properties: {
            key: {
              type: 'string',
              description: '要清除的缓存键名，不指定则清除所有缓存',
            },
          },
        },
        category: ToolCategory.COMMON,
      },
      async (args) => {
        const key = args.key as string | undefined;

        if (key) {
          const deleted = WpsMcpServer.dataCache.delete(key);
          logger.info(`Cache cleared: key=${key}, deleted=${deleted}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: deleted ? `缓存 "${key}" 已清除` : `缓存 "${key}" 不存在`,
                }),
              },
            ],
          };
        } else {
          const count = WpsMcpServer.dataCache.size;
          WpsMcpServer.dataCache.clear();
          logger.info(`All cache cleared: count=${count}`);

          return {
            id: '',
            success: true,
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `已清除所有缓存，共 ${count} 条`,
                }),
              },
            ],
          };
        }
      }
    );

    logger.info(`Registered ${this.registry.size} built-in tools`);
  }

  /**
   * 启动服务器
   */
  async start(): Promise<void> {
    if (this.isRunning) {
      logger.warn('Server is already running');
      return;
    }

    logger.info('Starting MCP Server...');

    // 工具已在构造期注册完（registerBuiltinTools + registerAll），
    // 因为 instructions 要在 new Server() 时就带上真实的工具数。

    // 创建stdio传输层
    const transport = new StdioServerTransport();

    // 连接传输层
    await this.server.connect(transport);

    // 工具组 load/unload 后通知客户端刷新 tools/list（否则客户端缓存里还是老列表）
    if (!this.groupsNotifyWired) {
      this.groupsNotifyWired = true;
      onToolGroupsChanged(() => {
        void this.notifyToolsChanged();
      });
    }

    this.isRunning = true;
    logger.info('MCP Server started successfully');
  }

  /**
   * 发 notifications/tools/list_changed，让客户端重新拉工具列表。
   * 失败只记日志 —— 通知失败不该影响工具本身的结果。
   */
  private async notifyToolsChanged(): Promise<void> {
    try {
      await this.server.notification({ method: 'notifications/tools/list_changed' });
      logger.info('已通知客户端工具列表变化', {
        listed: this.registry.listTools().tools.length,
        loadedGroups: this.registry.getLoadedGroups(),
      });
    } catch (err) {
      logger.warn('发送 tools/list_changed 通知失败', { error: (err as Error).message });
    }
  }

  /**
   * 停止服务器
   */
  async stop(): Promise<void> {
    if (!this.isRunning) {
      logger.warn('Server is not running');
      return;
    }

    logger.info('Stopping MCP Server...');

    await this.server.close();

    this.isRunning = false;
    logger.info('MCP Server stopped');
  }

  /**
   * 获取服务器状态
   */
  getStatus(): { running: boolean; toolCount: number } {
    return {
      running: this.isRunning,
      toolCount: this.registry.size,
    };
  }
}

// 导出单例创建函数
export const createMcpServer = (
  config?: Partial<McpServerConfig>
): WpsMcpServer => {
  return new WpsMcpServer(config);
};

export default WpsMcpServer;

/**
 * 成功但有保留时，在正文前挂一段提示。
 *
 * 为什么必须显式写出来：模型看到 success=true 就会当"全好了"，
 * 不说的话它不会去看 warnings —— 而 warnings 里的内容往往正是用户需要知道的
 * （"某个属性没设上""这个锚点插不进去"）。宁可啰嗦一点，也不能让用户以为全好。
 */
export function prefixWithWarnings(result: ToolCallResult): ToolContent[] {
  const ws = result.warnings;
  if (!ws || !ws.length) return result.content;
  const lines = ws.map((w) => (w.step ? `  - ${w.step}：${w.message}` : `  - ${w.message}`));
  const head = `[部分成功|主操作已完成，${ws.length} 处未按预期]`;
  const hint = '以下是没成功的部分（不要当成全好，也不要重复执行整批）';
  return result.content.map((c, i) =>
    i === 0 && c.type === 'text' && typeof c.text === 'string'
      ? { ...c, text: `${head} ${hint}\n${lines.join('\n')}\n\n${c.text}` }
      : c
  );
}
