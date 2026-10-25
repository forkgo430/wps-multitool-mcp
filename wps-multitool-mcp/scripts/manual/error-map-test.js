/**
 * Input: 无（内置实测错误字符串用例）
 * Output: 错误分类断言结果；全通过打印 PASS n/n，失败打印期望/实际并以非 0 退出
 * Pos: 手动证据脚本（读 dist/utils/error.js）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 */

const {
  classifyWpsError,
  WpsErrorType,
  McpError,
  errorUtils,
} = require('../..' + '/dist/utils/error.js');

// 每一行对应任务里给出的"实测错误字符串 → 期望类型"表。
// expect 为数组时表示该行语义允许落入其中任一类型。
const cases = [
  { input: 'WPS 动作超时（120000ms）: getCellValue', expect: 'timeout' },
  { input: 'WPS 守护进程已退出', expect: ['daemon_unavailable', 'timeout'] },
  { input: '桥连接已关闭', expect: ['daemon_unavailable', 'timeout'] },
  { input: 'WPS 动作超时', expect: 'timeout' },
  { input: 'WPS 守护进程已关闭', expect: 'daemon_unavailable' },
  { input: '桥进程 30000ms 内未就绪', expect: 'bridge_unavailable' },
  { input: 'WPS Word not running', expect: 'app_not_running' },
  { input: 'No active document', expect: 'document_not_found' },
  { input: 'Unknown action: xxx', expect: 'unsupported_operation' },
  { input: '值不在预期的范围内', expect: 'unsupported_operation' },
  {
    input: '操作无法使用 (异常来自 HRESULT:0x800401E3 (MK_E_UNAVAILABLE))',
    expect: 'app_not_running',
  },
  {
    input: 'Call was rejected by callee. (异常来自 HRESULT:0x80010001 (RPC_E_CALL_REJECTED))',
    expect: 'com_error',
  },
  { input: 'PowerShell 输出解析失败（非有效JSON）: {"a":', expect: 'protocol_error' },
  { input: 'PowerShell 退出码 1: (无输出)', expect: 'protocol_error' },
  { input: '块未找到', expect: 'document_not_found' },
  { input: '未找到名为 doc-1 的块', expect: 'document_not_found' },
  { input: '不能对 Null 值表达式调用方法', expect: 'com_error' },
  { input: '插入题注失败: 对象引用未设置到对象的实例。', expect: 'com_error' },
  { input: 'Missing required parameter: text', expect: 'invalid_params' },
  // —— 阶段 3 新增：并发控制的 busy。两条文案都含"超时"，必须判成 busy 而不是 timeout ——
  {
    input:
      'WPS common 队列已满（4 个在等，动作=ping）。该应用同一时刻只能跑一个动作（COM 是 STA），请稍后重试或先降低并发。',
    expect: 'busy',
  },
  {
    input:
      'WPS word 排队超时（10000ms，动作=insertText）。前面还有动作在跑；这个请求从未送出，可以安全重试。',
    expect: 'busy',
  },
  { input: '', expect: 'unknown' },
  { input: 'Something totally unrelated', expect: 'unknown' },
];

const checks = [];

// 1) classifyWpsError 逐行断言
for (const c of cases) {
  checks.push({
    name: `classifyWpsError(${JSON.stringify(c.input)})`,
    actual: classifyWpsError(c.input),
    expect: c.expect,
  });
}

// 2) 枚举值必须就是那串稳定标识
checks.push({
  name: 'WpsErrorType.TIMEOUT 值',
  actual: WpsErrorType.TIMEOUT,
  expect: 'timeout',
});
checks.push({
  name: 'WpsErrorType.BRIDGE_UNAVAILABLE 值',
  actual: WpsErrorType.BRIDGE_UNAVAILABLE,
  expect: 'bridge_unavailable',
});

// 3) 类型要能从 errorUtils.wrap / McpError 透出来
checks.push({
  name: "errorUtils.wrap(new Error('WPS 动作超时（120000ms）: x')).wpsType",
  actual: errorUtils.wrap(new Error('WPS 动作超时（120000ms）: x')).wpsType,
  expect: 'timeout',
});
checks.push({
  name: "errorUtils.wrap('PowerShell 退出码 1: x').wpsType",
  actual: errorUtils.wrap('PowerShell 退出码 1: x').wpsType,
  expect: 'protocol_error',
});
checks.push({
  name: "new McpError('桥进程 30000ms 内未就绪').wpsType",
  actual: new McpError('桥进程 30000ms 内未就绪').wpsType,
  expect: 'bridge_unavailable',
});

function ok(actual, expect) {
  return Array.isArray(expect) ? expect.includes(actual) : actual === expect;
}

let pass = 0;
const failures = [];
for (const c of checks) {
  if (ok(c.actual, c.expect)) {
    pass++;
  } else {
    const exp = Array.isArray(c.expect) ? c.expect.join(' | ') : c.expect;
    failures.push(`  FAIL ${c.name}\n       期望: ${exp}\n       实际: ${c.actual}`);
  }
}

console.log(`--- 错误分类断言（共 ${checks.length} 条）---`);
if (failures.length === 0) {
  console.log(`PASS ${pass}/${checks.length}`);
  process.exit(0);
} else {
  console.log(failures.join('\n'));
  console.log(`FAIL ${pass}/${checks.length}`);
  process.exit(1);
}
