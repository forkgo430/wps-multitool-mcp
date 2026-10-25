/**
 * Input: 无
 * Output: 对错误类型表的【独立】验证（不依赖子代理自己的用例）
 * Pos: 错误分类的独立核验脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要另写一个
 * ----------------
 * 别人（或子代理）自报"26/26 PASS"不能作为结论。这里用**我自己挑的**输入复核，
 * 并补上它自己说"没验"的那一段：errorType 是否真的透到 ToolCallResult 上。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');

const {
  classifyWpsError,
  WpsErrorType,
  errorUtils,
  isRetryableWpsError,
  formatWpsErrorTag,
  McpError,
} = require(path.join(ROOT, 'dist', 'utils', 'error.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { ToolCategory } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));

let pass = 0;
let fail = 0;
function eq(label, actual, expected) {
  const ok = actual === expected;
  if (ok) pass++;
  else fail++;
  console.log(`  ${ok ? '✅' : '❌'} ${label}\n      期望 ${expected}  实际 ${actual}`);
}

console.log('=== A. 分类（输入是我自己从本轮真实运行里抓的字符串）===');
const T = WpsErrorType;
eq('超时', classifyWpsError('WPS 动作超时（120000ms）: getCellValue（应用=excel）。桥连接已断开重连；因原请求可能执行了一半，这里不自动重放。'), T.TIMEOUT);
eq('RPC_E_CALL_REJECTED 必须不是 timeout', classifyWpsError('Call was rejected by callee. (异常来自 HRESULT:0x80010001 (RPC_E_CALL_REJECTED))'), T.COM_ERROR);
eq('MK_E_UNAVAILABLE', classifyWpsError('使用"1"个参数调用"GetActiveObject"时发生异常:"操作无法使用 (异常来自 HRESULT:0x800401E3 (MK_E_UNAVAILABLE))"'), T.APP_NOT_RUNNING);
eq('桥未就绪', classifyWpsError('桥进程 30000ms 内未就绪（管道 \\\\.\\pipe\\wps-multitool-mcp-x）：connect ENOENT'), T.BRIDGE_UNAVAILABLE);
eq('守护进程已关闭', classifyWpsError('守护进程已关闭（dispose 之后不再接受调用）'), T.DAEMON_UNAVAILABLE);
eq('WPS 没开', classifyWpsError('WPS Word not running'), T.APP_NOT_RUNNING);
eq('没有活动文档', classifyWpsError('No active document'), T.DOCUMENT_NOT_FOUND);
eq('未知动作', classifyWpsError('Unknown action: noSuchActionXyz'), T.UNSUPPORTED_OPERATION);
eq('取值超范围（WPS 不支持）', classifyWpsError('插入题注失败: 值不在预期的范围内。'), T.UNSUPPORTED_OPERATION);
eq('非 JSON 输出', classifyWpsError('PowerShell 输出解析失败（非有效JSON）: <html>'), T.PROTOCOL_ERROR);
eq('空字符串', classifyWpsError(''), T.UNKNOWN);
eq('null', classifyWpsError(null), T.UNKNOWN);
// 注：设计上「含中文的兜底」一律归 com_error（不是 unknown），这里按该口径断言
  eq('中文乱码兜底为 com_error（设计口径）', classifyWpsError('锟斤拷锟斤拷'), T.COM_ERROR);
  eq('缺参 Missing outputPath', classifyWpsError('Missing outputPath'), T.INVALID_PARAMS);
  eq('缺参 items 为空', classifyWpsError('items 为空'), T.INVALID_PARAMS);

console.log('\n=== B. errorUtils.wrap 是否带上 wpsType ===');
const wrapped = errorUtils.wrap(new Error('WPS Excel not running'), '调用失败');
eq('wrap 后 wpsType', wrapped.wpsType, T.APP_NOT_RUNNING);

console.log('\n=== C. 端到端：errorType 是否真的透到 ToolCallResult（子代理说这段没验）===');
toolRegistry.register(
  {
    name: 'wps_test_failing_tool',
    description: '仅用于测试',
    inputSchema: { type: 'object', properties: {} },
    category: ToolCategory.COMMON,
    group: 'core',
  },
  async () => {
    throw new Error('WPS 动作超时（120000ms）: ping（应用=common）');
  }
);
toolRegistry.register(
  {
    name: 'wps_test_failing_mcperror',
    description: '仅用于测试',
    inputSchema: { type: 'object', properties: {} },
    category: ToolCategory.COMMON,
    group: 'core',
  },
  async () => {
    throw errorUtils.wrap(new Error('No active document'), '取文档失败');
  }
);

(async () => {
  const r1 = await toolRegistry.callTool({ id: 't1', name: 'wps_test_failing_tool', arguments: {} });
  console.log(`  普通 Error 路径: success=${r1.success} errorType=${r1.errorType} text="${String(r1.content[0].text).slice(0, 60)}"`);
  eq('普通 Error 失败也被分类为 timeout', r1.errorType, T.TIMEOUT);

  const r2 = await toolRegistry.callTool({ id: 't2', name: 'wps_test_failing_mcperror', arguments: {} });
  console.log(`  McpError 路径: success=${r2.success} errorType=${r2.errorType}`);
  eq('McpError 路径分类为 document_not_found', r2.errorType, T.DOCUMENT_NOT_FOUND);

  // —— retryable：调用方靠它决定"退避重试"还是"别试了"（阶段 3 新增）——
  console.log('\n=== retryable 语义 ===');
  eq('busy 可安全重试（请求从未送出）', isRetryableWpsError(T.BUSY), true);
  eq('动作超时【不可】重试（可能已执行一半）', isRetryableWpsError(T.TIMEOUT), false);
  eq('com_error 不可重试', isRetryableWpsError(T.COM_ERROR), false);
  eq('桥不可用可重试', isRetryableWpsError(T.BRIDGE_UNAVAILABLE), true);
  const m = new McpError('WPS common 队列已满（4 个在等）');
  eq('McpError.retryable 对 busy 为 true', m.retryable, true);
  eq('McpError.toJSON 带上 retryable', m.toJSON().retryable, true);

  // —— 错误标签：模型实际看到的那串前缀（P0 新增）——
  console.log('\n=== formatWpsErrorTag 渲染 ===');
  eq('busy 标签', formatWpsErrorTag(T.BUSY), '[busy|可重试·稍后退避]');
  eq('timeout 标签', formatWpsErrorTag(T.TIMEOUT), '[timeout|勿重试·可能已执行一半]');
  eq('com_error 标签', formatWpsErrorTag(T.COM_ERROR), '[com_error|勿重试·可能已执行一半]');
  eq('app_not_running 标签', formatWpsErrorTag(T.APP_NOT_RUNNING), '[app_not_running|可重试·先启动WPS]');
  eq('invalid_params 标签', formatWpsErrorTag(T.INVALID_PARAMS), '[invalid_params|勿重试·先改参数]');
  eq('unknown 标签（无提示后缀）', formatWpsErrorTag(T.UNKNOWN), '[unknown|勿重试]');
  eq('显式 retryable 覆盖', formatWpsErrorTag(T.TIMEOUT, true), '[timeout|可重试·可能已执行一半]');

  // —— ToolCallResult 是否带上 retryable（前缀要用它）——
  eq('ToolCallResult.retryable（timeout 路径）', r1.retryable, false);
  eq('ToolCallResult.retryable（document_not_found 路径）', r2.retryable, false);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})();
