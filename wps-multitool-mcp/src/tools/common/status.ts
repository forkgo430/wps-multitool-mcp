/**
 * Input: 无
 * Output: 桥/守护进程 + WPS 三件套的运行状态
 * Pos: 通用状态探针工具。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 为什么需要它
 * ------------
 * 改造前没有任何地方能回答"现在到底连上没有"。出了问题只能看日志翻，
 * 或者盲猜。对标 office-mcp 的 office_status() 与 ghidra-mcp 的 mcp_health。
 *
 * 报告三层：
 *   1) 桥客户端（本进程内）：连接是否就绪、管道名、桥 PID、重连次数、在途请求数
 *   2) 桥进程内部：桥自己的 PID/存活时长/客户端数，以及【每个应用】的常驻子进程状态
 *   3) WPS 这一层（PowerShell 侧）：三件套分别有没有【已附着】的实例
 *
 * 要点：本工具【只读】且【不产生副作用】。
 *   PowerShell 侧走 wpsStatus 动作，内部只用 Get-ActiveWpsApp（只取已运行实例），
 *   不会像 getAppInfo 那样顺手 New-Object 把 WPS 拉起来。
 *   WPS 没开时也要能正常返回 —— 状态探针不该因为目标没运行就报错。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { psDaemon } from '../../client/ps-daemon';
import { versionLine, MCP_VERSION, MCP_AUTHOR } from '../../version';

/** PowerShell 侧 wpsStatus 动作返回的数据形状 */
interface AppState {
  running: boolean;
  /** 我们请求的 ProgID 对应的应用名（身份权威来源） */
  app?: string;
  /** 我们请求的 ProgID，如 Ket.Application */
  progId?: string;
  /**
   * 应用自报的名字（Application.Name）。
   * 注意：WPS 为兼容 Excel/Word 宏会自报成 "Microsoft Excel"/"Microsoft Word"，
   * 这是兼容名，【不能】当作身份判据。
   */
  reportedName?: string;
  /** Application.Version，同样是兼容版本号（如 12.0），仅供参考 */
  version?: string;
  /** Application.Path —— 安装目录，用来交叉验证到底是 WPS 还是 MS Office */
  path?: string;
  /** 由安装目录反推的产品名：WPS Office / Microsoft Office / unknown */
  product?: string;
  documents?: number;
}

interface WpsStatusData {
  apps?: Record<string, AppState>;
  psPid?: number;
  psVersion?: string;
  daemon?: boolean;
}

export const statusDefinition: ToolDefinition = {
  name: 'wps_common_status',
  description: `查看 WPS MCP 的运行状态：桥（守护进程）这一层 + WPS 三件套这一层。

报告内容：
- 桥：是否就绪、进程号、已存活多久、重连次数、在途请求数、队列上限与排队超时
- 每个应用（表格/文字/演示/通用）的子进程：在途数、排队数、最近一次排队等待、累计被拒次数
- WPS 三件套：分别有没有已连接的实例，以及实例名、版本、打开文档数

关于并发（重要）：
COM 是 STA，同一个应用同一时刻只能跑一个动作。所以并发不是"提高吞吐"，
而是把等待变成有界的、可见的：队列默认上限 4、排满立刻回 [busy]、
排队超过 10 秒也回 [busy]（这两个都是【可安全重试】的，请求从未送出）。
看到 queued 长期不为 0，说明前面有慢动作在占着这一路。

关于"实例名"的一个重要提醒：
WPS 为了兼容 Excel/Word 宏，会故意把 Application.Name 自报成 "Microsoft Excel" /
"Microsoft Word"，Version 也报成 Office 的兼容版本号。所以本工具把身份拆成三个字段：
  app / progId  —— 权威身份（我们请求的就是 WPS 自家的 Ket/Kwps/Kwpp）
  product       —— 由 Application.Path 反推的产品（WPS Office / Microsoft Office）
  reportedName  —— 应用自报名，仅供参考，别拿它判断"到底在用谁"
（实测本机 Ket.Application 的 Path 就在 WPS 安装目录，且没装 EXCEL.EXE。）

使用场景：
- "WPS 连上了吗"
- "为什么刚才那个工具没反应 / 报 busy"
- "现在开了几个文档"
- 排查问题时先跑这个，比翻日志快

本工具只读，且不会顺手把 WPS 启动起来。`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const statusHandler: ToolHandler = async (): Promise<ToolCallResult> => {
  // 桥这一层：本进程内的客户端视图（永远拿得到）
  const client = { ...psDaemon.stats };

  // 桥进程自己的视图（PID/各应用子进程）—— 拿不到也不影响，如实标注
  let bridgeInternal: unknown = null;
  let bridgeInternalError: string | null = null;
  try {
    const r = (await psDaemon.bridgeStatus()) as { success?: boolean; data?: unknown; error?: string };
    if (r?.success) bridgeInternal = r.data ?? null;
    else bridgeInternalError = String(r?.error ?? '未知');
  } catch (err) {
    bridgeInternalError = err instanceof Error ? err.message : String(err);
  }

  const bridge = { ...client, internal: bridgeInternal, internalError: bridgeInternalError };

  // WPS 这一层：拿不到就如实标注，不要把失败伪装成"没运行"
  let wps: WpsStatusData | null = null;
  let wpsError: string | null = null;
  try {
    const response = await wpsClient.executeMethod<WpsStatusData>('wpsStatus', {});
    if (response.success) {
      wps = response.data ?? null;
    } else {
      wpsError = String(response.error ?? '未知错误');
    }
  } catch (err) {
    wpsError = err instanceof Error ? err.message : String(err);
  }

  // 身份自检：自报名说是 MS Office、但安装目录在 WPS 下 —— 这是 WPS 的兼容名，提示一下
  const apps: Record<string, AppState> = wps?.apps ?? {};
  const identityNotes: string[] = [];
  for (const [key, st] of Object.entries(apps)) {
    if (!st?.running) continue;
    const reported = st.reportedName ?? '';
    const isMsName = /microsoft\s+(excel|word|powerpoint)/i.test(reported);
    if (isMsName && st.product === 'WPS Office') {
      identityNotes.push(
        `${key}: 自报名为 "${reported}"（WPS 的 Office 兼容名），实际产品是 ${st.product}（${st.path ?? '路径未知'}）`
      );
    } else if (isMsName && st.product === 'Microsoft Office') {
      identityNotes.push(`${key}: 自报 "${reported}" 且安装目录为 ${st.path} —— 真的是 Microsoft Office`);
    }
  }

  const payload = {
    version: {
      line: versionLine(),
      version: MCP_VERSION,
      author: MCP_AUTHOR,
    },
    bridge,
    wps: wps
      ? {
          apps,
          powershell: { pid: wps.psPid, version: wps.psVersion, daemonMode: wps.daemon },
          identityNotes,
        }
      : null,
    wpsError,
  };

  return {
    id: '',
    success: true,
    content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
  };
};

/** 状态类工具集合 */
export const statusTools: RegisteredTool[] = [
  { definition: statusDefinition, handler: statusHandler },
];

export default statusTools;
