/**
 * Input: 通用操作工具参数
 * Output: 操作结果
 * Pos: 通用操作工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * - wps_common_wire_check: 检查通信线路
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import {
  setWordTarget,
  setExcelTarget,
  setPptTarget,
} from '../../client/wps-client';
import { call, ok, fail } from './helpers';

// ============================================================
// 1. wps_common_save - 保存当前文档
// ============================================================

export const saveDefinition: ToolDefinition = {
  name: 'wps_common_save',
  description: `保存当前文档。

使用场景：
- "保存一下"
- "Ctrl+S"
- "把修改存起来"

特点：
- 自动保存当前活动文档
- 如果文档从未保存过，可能会提示选择保存路径`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const saveHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      success: boolean;
      message: string;
    }>('save', {});

    if (response.ok) {
      return ok(`文档保存成功！${response.data?.message || ''}`, response.warnings);
    } else {
      return fail(`保存失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`保存出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 2. wps_common_save_as - 另存为
// ============================================================

export const saveAsDefinition: ToolDefinition = {
  name: 'wps_common_save_as',
  description: `将当前文档另存为指定路径和格式。

使用场景：
- "另存为到桌面"
- "换个名字保存"
- "保存一份副本到指定位置"

特点：
- 支持指定完整文件路径
- 可选指定保存格式`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      filePath: {
        type: 'string',
        description: '目标文件完整路径，包含文件名和扩展名',
      },
      format: {
        type: 'string',
        description: '保存格式（可选），如 docx, xlsx, pptx 等',
      },
    },
    required: ['filePath'],
  },
};

export const saveAsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { filePath, format } = args as {
    filePath: string;
    format?: string;
  };

  if (!filePath || filePath.trim() === '') {
    return fail('文件路径不能为空，请指定保存路径', '文件路径为空');
  }

  try {
    const params: Record<string, unknown> = {
      filePath,
      path: filePath,
      outputPath: filePath,
    };
    if (format) {
      params.format = format.toLowerCase().replace(/^\./, '');
    }

    const response = await call<{
      success: boolean;
      message: string;
      outputPath: string;
      appType?: string;
    }>('saveAs', params);

    if (response.ok && response.data) {
      // 另存为会【改掉文档的 Name】。之前锁定的目标名（比如"文字文稿1"）就此失效，
      // 后续调用会报一句莫名其妙的 "No active document" —— 实测踩过一次。
      // 这里按新文件名重新锁定；错锁比不锁更危险，所以类型不明时宁可不锁。
      const outPath = response.data.outputPath || filePath;
      const base = outPath.split(/[\\/]/).pop();
      const app = (response.data.appType || '').toLowerCase();
      if (base) {
        if (app === 'word') setWordTarget(base);
        else if (app === 'excel') setExcelTarget(base);
        else if (app === 'ppt') setPptTarget(base);
      }
      return ok(`另存为成功！\n输出路径: ${outPath}`, response.warnings);
    } else {
      return fail(`另存为失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`另存为出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 4. wps_check_connection - 检查通信线路（B5 意图升级：ping / check_connection /
//    wire_check 三名合一的规范名；wps_common_wire_check / wps_common_ping 降为别名，
//    映射见 scripts/manual/tool-aliases.json）
// ============================================================

export const wireCheckDefinition: ToolDefinition = {
  name: 'wps_check_connection',
  description: `检查与WPS加载项之间的通信线路状态（连接健康检查的唯一规范入口）。

使用场景：
- "检查通信状态"
- "诊断连接问题"
- "通信线路是否正常"

特点：
- 原 wps_common_ping / wps_common_wire_check 两个名字已合并到这里（旧名走别名仍可调用）
- 返回线路延迟、协议状态等信息`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const wireCheckHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      success: boolean;
      message: string;
      latency?: number;
      protocol?: string;
    }>('wireCheck', {});

    if (response.ok && response.data) {
      const data = response.data;
      let info = '通信线路正常！';
      if (data.latency !== undefined) {
        info += `\n延迟: ${data.latency}ms`;
      }
      if (data.protocol) {
        info += `\n协议: ${data.protocol}`;
      }
      if (data.message) {
        info += `\n${data.message}`;
      }

      return ok(info, response.warnings);
    } else {
      return fail(`通信线路检查失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`通信线路检查出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 5. wps_common_get_app_info - 获取WPS应用信息
// ============================================================

export const getAppInfoDefinition: ToolDefinition = {
  name: 'wps_common_get_app_info',
  description: `获取WPS应用的基本信息。

使用场景：
- "WPS是什么版本"
- "查看WPS信息"
- "获取应用状态"

特点：
- 返回WPS版本号、构建信息
- 返回当前打开的文档信息
- 返回运行平台信息`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getAppInfoHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      success: boolean;
      message: string;
      version?: string;
      build?: string;
      platform?: string;
      activeDocument?: string;
    }>('getAppInfo', {});

    if (response.ok && response.data) {
      const data = response.data;
      const lines: string[] = ['WPS应用信息：'];
      if (data.version) lines.push(`版本: ${data.version}`);
      if (data.build) lines.push(`构建号: ${data.build}`);
      if (data.platform) lines.push(`平台: ${data.platform}`);
      if (data.activeDocument) lines.push(`当前文档: ${data.activeDocument}`);
      if (data.message) lines.push(data.message);

      return ok(lines.join('\n'), response.warnings);
    } else {
      return fail(`获取应用信息失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`获取应用信息出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 6. wps_common_get_selected_text - 获取选中文本
// ============================================================

export const getSelectedTextDefinition: ToolDefinition = {
  name: 'wps_common_get_selected_text',
  description: `获取当前文档中选中的文本内容。

使用场景：
- "读取我选中的内容"
- "获取当前选区的文字"
- "看看我选了什么"

特点：
- 返回当前选区的纯文本内容
- 适用于Word、Excel、PPT`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {},
    required: [],
  },
};

export const getSelectedTextHandler: ToolHandler = async (
  _args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      success: boolean;
      message: string;
      text: string;
    }>('getSelectedText', {});

    if (response.ok && response.data) {
      const text = response.data.text;
      if (!text || text.length === 0) {
        return ok('当前没有选中任何文本。', response.warnings);
      }
      return ok(`选中的文本内容：\n${text}`, response.warnings);
    } else {
      return fail(`获取选中文本失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`获取选中文本出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 7. wps_common_set_selected_text - 替换选中文本
// ============================================================

export const setSelectedTextDefinition: ToolDefinition = {
  name: 'wps_common_set_selected_text',
  description: `替换当前文档中选中的文本内容。

使用场景：
- "把选中的内容替换成xxx"
- "修改选区的文字"
- "用新内容替换选中部分"

特点：
- 将当前选区的文本替换为指定内容
- 适用于Word、Excel、PPT`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      text: {
        type: 'string',
        description: '要替换为的新文本内容',
      },
    },
    required: ['text'],
  },
};

export const setSelectedTextHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { text } = args as { text: string };

  if (text === undefined || text === null) {
    return fail('替换文本不能为空，请指定要替换的内容', '替换文本为空');
  }

  try {
    const response = await call<{
      success: boolean;
      message: string;
    }>('setSelectedText', { text });

    if (response.ok) {
      return ok(`选中文本已替换成功！${response.data?.message || ''}`, response.warnings);
    } else {
      return fail(`替换选中文本失败: ${response.error}`, response.error);
    }
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`替换选中文本出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 8. wps_common_show_window - 将 WPS 置于前台
// ============================================================

export const showWindowDefinition: ToolDefinition = {
  name: 'wps_common_show_window',
  description: `把 WPS 窗口显示出来并置于前台。

使用场景：
- "WPS 看不见了"
- "把 WPS 调出来"
- "显示 WPS 窗口"
- "WPS 跑到后台了"

特点：
- 激活已有 WPS 窗口并置于最前
- COM 自动化实例本身无窗口时，会自动启动一个可见的 WPS 实例
- 传 launch=false 可只做激活、不启动新实例`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      launch: {
        type: 'boolean',
        description: '无可激活窗口时是否启动可见 WPS 实例，默认 true',
      },
    },
    required: [],
  },
};

export const showWindowHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const launch = args.launch === undefined ? true : Boolean(args.launch);
  try {
    const response = await call<{
      apps?: string[];
      visible?: number;
      activated?: number;
      foreground?: boolean;
      launched?: boolean;
      message?: string;
    }>('showWps', { launch });

    if (response.ok && response.data) {
      const d = response.data;
      const lines: string[] = ['WPS 窗口已唤起：'];
      if (d.apps && d.apps.length) lines.push(`应用: ${d.apps.join(' / ')}`);
      lines.push(`设为可见: ${d.visible ?? 0} 个`);
      lines.push(`激活窗口: ${d.activated ?? 0} 个`);
      lines.push(`启动可见实例: ${d.launched ? '是' : '否'}`);
      lines.push(`置前成功: ${d.foreground ? '是' : '否'}`);
      if (d.message) lines.push(d.message);
      return ok(lines.join('\n'), response.warnings);
    }
    return fail(`唤起 WPS 窗口失败: ${response.error || '未找到运行中的 WPS 实例'}`, response.error);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`唤起 WPS 窗口出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 9. wps_common_quit - 结束 WPS
// ============================================================

export const quitDefinition: ToolDefinition = {
  name: 'wps_common_quit',
  description: `结束 WPS（Word / Excel / PPT 全部）。

使用场景：
- "关掉 WPS"
- "退出 WPS"

特点：
- 只做优雅退出（Quit），不做全量杀进程 —— 历史版本会 taskkill /IM wps.exe，把用户自己开着的其它 WPS 窗口一起误杀，已移除
- 本会话自己创建的 WPS 实例（含沙箱）按 PID 直接收掉并计入 ownKilled —— 任务结束或进程异常都能清干净，不留自建残留
- 有未保存的文档/工作簿/演示文稿时默认拒绝并列出清单（只约束附着到的用户实例）；确认要放弃未保存内容，传 confirm=true
- 用户自己开的其它 WPS 不受影响，剩余进程会列在结果里
- force 参数已废弃：传不传都一样，只保留字段名兼容旧调用`,
  category: ToolCategory.COMMON,
  inputSchema: {
    type: 'object',
    properties: {
      confirm: {
        type: 'boolean',
        description:
          '允许丢弃未保存内容并退出。不传且有未保存文档时，调用会失败并列出文档清单',
      },
      force: {
        type: 'boolean',
        description:
          '已废弃，无效果。历史行为是全量强杀 WPS 进程（会误杀用户自己开的窗口），已移除',
      },
    },
    required: [],
  },
};

export const quitHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  try {
    const response = await call<{
      quit?: string[];
      forced?: boolean;
      remaining?: number;
      message?: string;
    }>('quitWps', { confirm: args.confirm === true });

    if (response.ok && response.data) {
      const d = response.data;
      return ok(`${d.message || 'WPS 已退出'}\n优雅退出: ${
              (d.quit || []).join(', ') || '无活动实例'
            }\n剩余进程: ${d.remaining ?? 0}${
              d.remaining ? '（多为用户自己打开的实例，未强行关闭）' : ''
            }`, response.warnings);
    }
    return fail(`结束 WPS 失败: ${response.error}`, response.error);
  } catch (error) {
    const errMsg = error instanceof Error ? error.message : String(error);
    return fail(`结束 WPS 出错: ${errMsg}`, errMsg);
  }
};

// ============================================================
// 导出所有通用操作Tools
// ============================================================

export const generalTools: RegisteredTool[] = [
  { definition: saveDefinition, handler: saveHandler },
  { definition: saveAsDefinition, handler: saveAsHandler },
  { definition: wireCheckDefinition, handler: wireCheckHandler },
  { definition: getAppInfoDefinition, handler: getAppInfoHandler },
  { definition: getSelectedTextDefinition, handler: getSelectedTextHandler },
  { definition: setSelectedTextDefinition, handler: setSelectedTextHandler },
  { definition: showWindowDefinition, handler: showWindowHandler },
  { definition: quitDefinition, handler: quitHandler },
];

export default generalTools;
