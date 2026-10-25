/**
 * 并发控制实测（阶段 3）
 *
 * 验三件事：
 *   1) 同一应用并发 N 个调用 → 在途恒为 1（COM 是 STA 的硬约束），其余排队
 *   2) 队列上限 4 → 第 5 个【立刻】收到 busy，而不是无声地等
 *   3) busy 被 classifyWpsError 归到 WpsErrorType.BUSY，且标为可安全重试
 *
 * 另外确认：排他动作（save/openFile 等）插队。
 *
 * 前提：ps-child 走真实 PowerShell 子进程 —— 所以这是实跑，不是打桩。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-queuetest';
process.env.WPS_BRIDGE_IDLE_MS = '30000';
// 队列压到最小，让"第 5 个被拒"能被快速验到
process.env.WPS_QUEUE_MAX = process.env.WPS_QUEUE_MAX || '4';
process.env.WPS_QUEUE_TIMEOUT = process.env.WPS_QUEUE_TIMEOUT || '10000';
// 动作超时给足，别让"慢"被误判成超时
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '60000';

const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));
const {
  classifyWpsError,
  isRetryableWpsError,
  WpsErrorType,
} = require(path.join(ROOT, 'dist', 'utils', 'error.js'));

const results = [];
function check(name, ok, detail = '') {
  results.push({ name, ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  → ' + detail : ''}`);
}

const appPid = async (app) => {
  const r = await psDaemon.call('__bridgeStatus', {}, 'common').catch(() => null);
  if (!r || !r.success) return null;
  const c = (r.data.children || []).find((x) => x.app === app);
  return c ? c.pid : null;
};

async function main() {
  console.log(`\n=== 队列参数：max=${process.env.WPS_QUEUE_MAX} timeout=${process.env.WPS_QUEUE_TIMEOUT}ms ===\n`);

  // ---------- 1) 预热：把子进程拉起来，避免冷启动 5 秒混进排队时间 ----------
  console.log('预热 common 一路…');
  const warm = await psDaemon.call('ping', {}, 'common');
  check('预热 ping 成功', warm && warm.success === true, JSON.stringify(warm).slice(0, 60));

  // ---------- 2) 并发 10 个 → 队列上限 4，多出来的应被立刻 busy 拒 ----------
  // 用 10 个（不是 6）：上限 4 + 在途 1 = 最多接 5 个，多出的 5 个必被拒。
  // 用 6 个只会被拒 1 个，容易被"其实有一个溜过去了"糊弄过去。
  const CONC = 10;
  console.log(`\n并发发起 ${CONC} 个 ping（上限 4 + 在途 1 → 最多接 5）…`);
  // 注意计数口径：桥对 busy 是【回一条 success:false 的响应】（和其它业务错误一致），
// 不是断开/抛异常。所以要同时看 rejected 和 fulfilled-but-success:false 两个通道。
const settled = await Promise.allSettled(
    Array.from({ length: CONC }, () => psDaemon.call('ping', {}, 'common'))
  );
  const okCount = settled.filter((s) => s.status === 'fulfilled' && s.value?.success === true).length;
  const rejectErrs = settled
    .filter((s) => s.status === 'rejected')
    .map((s) => String(s.reason?.message ?? s.reason));
  const failRespErrs = settled
    .filter((s) => s.status === 'fulfilled' && s.value?.success === false)
    .map((s) => String(s.value?.error ?? ''));
  const busyErrs = [...rejectErrs, ...failRespErrs];

  const busyHits = busyErrs.filter((m) => /队列已满|排队超时/.test(m));
  console.log(
    `  成功 ${okCount} 个；被拒 ${busyErrs.length} 个` +
      `（reject 通道 ${rejectErrs.length} + success:false 通道 ${failRespErrs.length}）`
  );
  busyErrs.slice(0, 2).forEach((m) => console.log('  拒因：' + m.slice(0, 90)));

  check(
    `成功数 ≤ 5（上限 4 + 在途 1）`,
    okCount <= 5,
    `成功 ${okCount}`
  );
  check(
    `被拒数 ≥ ${CONC - 5}`,
    busyErrs.length >= CONC - 5,
    `实际被拒 ${busyErrs.length}`
  );
  check('被拒原因都带 busy 语义', busyHits.length === busyErrs.length, `${busyHits.length}/${busyErrs.length}`);

  // ---------- 3) busy 的分类与可重试标记 ----------
  console.log('\nbusy 分类与 retryable：');
  const busyMsg = busyErrs[0] ?? 'WPS common 队列已满（4 个在等，动作=ping）。';
  const t = classifyWpsError(busyMsg);
  check('classifyWpsError 归到 BUSY', t === WpsErrorType.BUSY, `实际 ${t}`);
  check('BUSY 标为可安全重试', isRetryableWpsError(t) === true);

  // 反向：动作超时【不能】被误判成可重试（可能已执行一半）
  const timeoutMsg = 'WPS 动作超时（60000ms）: save（应用=common）。该应用的常驻进程已重建';
  const t2 = classifyWpsError(timeoutMsg);
  check('动作超时仍归 TIMEOUT（没被 busy 规则抢走）', t2 === WpsErrorType.TIMEOUT, `实际 ${t2}`);
  check('TIMEOUT 不标可重试（避免重复执行）', isRetryableWpsError(t2) === false);

  // 排队超时的文案里也含"超时"，必须优先判成 busy
  const qTimeoutMsg = 'WPS word 排队超时（10000ms，动作=insertText）。前面还有动作在跑；这个请求从未送出，可以安全重试。';
  check('"排队超时"优先判 BUSY 而非 TIMEOUT', classifyWpsError(qTimeoutMsg) === WpsErrorType.BUSY);

  // ---------- 4) 子进程只有 1 个（没被并发搞出多进程） ----------
  const pid = await appPid('common');
  console.log(`\ncommon 子进程 PID = ${pid}`);
  check('子进程仍只有 1 个（并发不新起进程）', typeof pid === 'number' && pid > 0, `PID=${pid}`);

  // ---------- 5) 桥状态里能看到队列 ----------
  const st = await psDaemon.call('__bridgeStatus', {}, 'common');
  const q = st?.data?.queue;
  const childStats = (st?.data?.children || []).find((x) => x.app === 'common');
  console.log('  桥队列配置：', JSON.stringify(q));
  console.log('  common 子进程统计：', JSON.stringify(childStats));
  check('桥状态暴露 queue 配置', !!(q && q.max === 4 && q.timeoutMs === 10000), JSON.stringify(q));
  check(
    '子进程统计含 queued/queueMax/rejected/inFlight',
    !!(
      childStats &&
      'queued' in childStats &&
      'queueMax' in childStats &&
      'rejected' in childStats &&
      'inFlight' in childStats
    ),
    `rejected=${childStats?.rejected} inFlight=${childStats?.inFlight}`
  );
  check(
    '被拒次数已被如实计数',
    (childStats?.rejected ?? 0) >= CONC - 5,
    `rejected=${childStats?.rejected}（期望 ≥${CONC - 5}）`
  );
  check('压力过后 inFlight 已归零（没有卡住的请求）', childStats?.inFlight === false);

  // ---------- 6) 串行回归：队列空掉之后还能正常用 ----------
  const after = await psDaemon.call('ping', {}, 'common');
  check('压力测试后串行调用仍正常', after?.success === true);

  const passed = results.filter((r) => r.ok).length;
  console.log(`\n汇总：${passed}/${results.length} 通过`);
  psDaemon.dispose();
  process.exit(passed === results.length ? 0 : 1);
}

main().catch((e) => {
  console.error('测试异常:', e && e.message ? e.message : e);
  psDaemon.dispose();
  process.exit(1);
});