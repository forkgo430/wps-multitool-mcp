/**
 * Input: Tool 定义集合
 * Output: Tool 注册数组
 * Pos: MCP Tools 总入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Tools总入口 - MCP工具汇总注册模块
 * 整合Excel、Word、PPT、Common的所有Tools
 *
 * 使用方法：
 * import { allTools } from './tools';
 * toolRegistry.registerAll(allTools);
 *
 * 或者按需导入：
 * import { excelTools, wordTools, pptTools, commonTools } from './tools';
 *
 * 数字不在这里写死：工具总数/分布以注册表实测为准（node scripts/manual/gen-doc-stats.js
 * 会把它写进 README，CI 有 diff 门禁）。各子模块的分文件清单见对应目录的 md。
 */

import { RegisteredTool } from '../types/tools';
import { excelTools } from './excel';
import { wordTools } from './word';
import { pptTools } from './ppt';
import { commonTools } from './common';

/**
 * 所有专业工具（不含 mcp-server.ts 里 registerBuiltinTools 的 12 个内置工具）。
 * 结构：excel/ 公式·数据·图表·透视表·工作表·格式·工作簿·第4期缺口；
 *       word/ 格式·内容·文档管理·校对·学术·专业·第5期缺口；
 *       ppt/ 幻灯片·形状·动画·图表·3D 等；common/ 转换·通用·状态·修复·批量·工具组·意图。
 */
export const allTools: RegisteredTool[] = [
  ...excelTools,
  ...wordTools,
  ...pptTools,
  ...commonTools,
];

// 按应用类型分别导出
export { excelTools } from './excel';
export { wordTools } from './word';
export { pptTools } from './ppt';
export { commonTools } from './common';

// Excel相关导出
export {
  formulaTools,
  dataTools,
  setFormulaDefinition,
  setFormulaHandler,
  generateFormulaDefinition,
  generateFormulaHandler,
  diagnoseFormulaDefinition,
  diagnoseFormulaHandler,
  readRangeDefinition,
  readRangeHandler,
  writeRangeDefinition,
  writeRangeHandler,
  cleanDataDefinition,
  cleanDataHandler,
  removeDuplicatesDefinition,
  removeDuplicatesHandler,
} from './excel';

// Word相关导出
export {
  formatTools,
  contentTools,
  proofreadTools,
  applyStyleDefinition,
  applyStyleHandler,
  setFontDefinition,
  setFontHandler,
  generateTocDefinition,
  generateTocHandler,
  insertTextDefinition,
  insertTextHandler,
  findReplaceDefinition,
  findReplaceHandler,
} from './word';

// PPT相关导出
export {
  slideTools,
  addSlideDefinition,
  addSlideHandler,
  beautifyDefinition,
  beautifyHandler,
  unifyFontDefinition,
  unifyFontHandler,
} from './ppt';

// Common相关导出
export {
  convertTools,
  convertToPdfDefinition,
  convertToPdfHandler,
  convertFormatDefinition,
  convertFormatHandler,
  getAppTypeByExtension,
  getFormatCode,
} from './common';

/**
 * 获取所有Tool的数量
 */
export const getToolCount = (): number => allTools.length;

/**
 * 获取按应用分类的Tool数量
 */
export const getToolCountByApp = (): { excel: number; word: number; ppt: number; common: number } => ({
  excel: excelTools.length,
  word: wordTools.length,
  ppt: pptTools.length,
  common: commonTools.length,
});

export default allTools;
