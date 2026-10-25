/**
 * Input: 通用工具定义
 * Output: 通用工具注册数组
 * Pos: Common Tools 汇总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Common Tools入口 - 通用工具汇总模块
 * 整合跨应用的通用Tools
 *
 * 导出所有通用的Tools，Word/Excel/PPT均可使用
 * 主要包括文档转换、格式互转等跨应用功能
 */

import { RegisteredTool } from '../../types/tools';
import { convertTools } from './convert';
import { generalTools } from './general';
import { statusTools } from './status';
import { toolGroupTools } from './toolgroups';
import { intentTools } from './intent';
import { repairTools } from './repair';
import { batchTools } from './batch';
import { snapshotTools } from './snapshot';

/**
 * 所有通用的Tools
 * 包含：
 * - 转换Tools: convert_to_pdf, convert_format
 * - 通用Tools: save, save_as, ping, wire_check, get_app_info, get_selected_text, set_selected_text
 * - 状态Tools: status（桥 + 三件套运行状态，只读、无副作用）
 * - 修复Tools: repair（诊断文档健康度 / 挡路对话框 / 清理无名残留文档）
 * - 工具组管理（core 组，永远列出）: list_tool_groups, load_tool_group, unload_tool_group, search_tools
 * - 意图识别（core 组）: route_intent —— 一句话判断该开哪个组
 * - 批量执行: batch —— 一次 1~3 步，省掉逐个调用的往返，遇错不中断
 * - 快照: snapshot_create / snapshot_list / snapshot_restore —— COM 非事务的补偿层
 *
 * 默认只列基础档里【真机验证过】的工具（WPS_TOOL_TIER=full 可关闭），其余靠 search + load 按需拉
 */
export const commonTools: RegisteredTool[] = [
  ...convertTools,
  ...generalTools,
  ...statusTools,
  ...repairTools,
  ...batchTools,
  ...snapshotTools,
  ...toolGroupTools,
  ...intentTools,
];

// 分别导出，方便按需使用
export { convertTools } from './convert';
export { generalTools } from './general';
export { statusTools, statusDefinition, statusHandler } from './status';
export { repairTools, repairDefinition, repairHandler } from './repair';
export { batchTools, batchDefinition, batchHandler, BATCH_MAX_STEPS } from './batch';
export {
  snapshotTools,
  snapshotCreateDefinition,
  snapshotCreateHandler,
  snapshotListDefinition,
  snapshotListHandler,
  snapshotRestoreDefinition,
  snapshotRestoreHandler,
  createSnapshot,
  readOpLog,
  appendOpLog,
  SNAPSHOT_ROOT,
} from './snapshot';
export {
  intentTools,
  routeIntentDefinition,
  routeIntentHandler,
  classifyIntent,
} from './intent';
export {
  toolGroupTools,
  onToolGroupsChanged,
  listToolGroupsDefinition,
  listToolGroupsHandler,
  loadToolGroupDefinition,
  loadToolGroupHandler,
  unloadToolGroupDefinition,
  unloadToolGroupHandler,
  searchToolsDefinition,
  searchToolsHandler,
} from './toolgroups';

// 导出单独的定义和处理器，方便测试
export {
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './convert';

// 导出通用操作工具的定义和处理器
export {
  saveDefinition,
  saveHandler,
  saveAsDefinition,
  saveAsHandler,
  wireCheckDefinition,
  wireCheckHandler,
  getAppInfoDefinition,
  getAppInfoHandler,
  getSelectedTextDefinition,
  getSelectedTextHandler,
  setSelectedTextDefinition,
  setSelectedTextHandler,
} from './general';

export default commonTools;
