/**
 * Input: 无
 * Output: 取消机制的独立验证（真跑 PowerShell 子进程，不打桩）
 * Pos: 取消机制核验脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 要验的是这三条【安全红线】，错一条就会造成不可逆的后果：
 *   1) 排队中取消 → 一定成功，且请求从未送达（任何动作都安全）
 *   2) 已在途 + 只读 → 可以"放弃等待"（调用方立刻返回，子进程跑完，回话丢弃）
 *   3) 已在途 + 写入 → 【拒绝取消】，必须老实跑完如实返回
 *      （否则调用方以为没做成，实际做了一半 —— 比等着更糟）
 *
 * 另外两条容易出事的时序：
 *   4) 放弃等待之后 inFlight 不能提前变 false（否则下一个请求会写进还在忙的 STA 进程）
 *   5) 取消不是故障：要被分类成 cancelled，不能混进 timeout / com_error
 *
 * 依赖：wps-com.ps1 里的 __sleep（判为写入）与 __sleepRead（判为只读）两个时序钩子。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_QUEUE_MAX = process.env.WPS_QUEUE_MAX || '4';
process.env.WPS_QUEUE_TIMEOUT = process.env.WPS_QUEUE_TIMEOUT || '10000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '60000';

const { PsChildDaemon } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));
const { CancelError, isReadOnlyAction } = require(path.join(ROOT, 'dist', 'bridge', 'cancellable.js'));
const { classifyWpsError, WpsErrorType } = require(path.join(ROOT, 'dist', 'utils', 'error.js'));

let pass = 0;
let fail = 0;
function ok(label, cond, extra) {
  if (cond) { pass++; console.log(`  ✅ ${label}${extra ? '  ' + extra : ''}`); }
  else { fail++; console.log(`  ❌ ${label}${extra ? '  ' + extra : ''}`); }
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const d = new PsChildDaemon('common');

  console.log('=== A. 只读判定（纯函数，先过一遍真动作名）===');
  const RO = ['getCellValue', 'getOpenDocuments', 'getDocumentText', 'listStyles', 'readRange',
    'findInDocument', 'ping', 'wpsStatus', '__probeCaps', '__sleepRead'];
  const WR = ['setCellValue', 'insertText', 'addHeading', 'deleteStyle', 'save', 'saveAs',
    'convertToPDF', 'mergeCells', 'replaceRange', '__sleep'];
  for (const a of RO) ok(`只读: ${a}`, isReadOnlyAction(a) === true);
  for (const a of WR) ok(`写入: ${a}`, isReadOnlyAction(a) === false);

  console.log('\n=== B. 排队中取消（任何动作都该成功，且从未送达）===');
  // 先用一个慢的占住在途，后面的才会真的排队
  const inflight = d.call('__sleep', { ms: 1500 }, { externalId: 'HOLD' });
  await sleep(400);
  const queued = d.call('ping', {}, { externalId: 'Q1' });
  await sleep(100);
  const r1 = d.cancelByExternalId('Q1');
  ok('排队中的请求能取消', r1.ok === true, r1.reason);
  let q1Err = null;
  await queued.catch((e) => { q1Err = e; });
  ok('被取消的请求抛 CancelError', q1Err instanceof CancelError, q1Err && q1Err.name);
  ok('文案说明"从未送达"', !!q1Err && /从未送达/.test(q1Err.message));
  const v = await inflight;
  ok('占位的慢动作正常跑完（取消没波及它）', v && v.success === true, JSON.stringify(v && v.data));

  console.log('\n=== C. 已在途 + 写入 → 拒绝取消 ===');
  const w = d.call('__sleep', { ms: 1200 }, { externalId: 'W1' });
  await sleep(300);
  const r2 = d.cancelByExternalId('W1');
  ok('写入动作不予取消', r2.ok === false, r2.reason);
  const wv = await w;
  ok('写入动作老实跑完并如实返回', wv && wv.success === true, JSON.stringify(wv && wv.data));

  console.log('\n=== D. 已在途 + 只读 → 放弃等待，且 inFlight 不提前释放 ===');
  const ac = new AbortController();
  const ro = d.call('__sleepRead', { ms: 1500 }, { externalId: 'R1', signal: ac.signal });
  await sleep(300);
  const statsMid = d.stats;
  ok('此刻确实有一个动作在途', statsMid.inFlight === true, `inFlight=${statsMid.inFlight}`);
  ac.abort();
  let roErr = null;
  const t0 = Date.now();
  await ro.catch((e) => { roErr = e; });
  const dt = Date.now() - t0;
  ok('只读动作被放弃等待（立刻返回）', roErr instanceof CancelError, `${dt}ms 后返回`);
  ok('返回明显早于动作本身（1500ms）', dt < 800, `${dt}ms`);
  // 关键：子进程还在跑，这时 inFlight 必须仍然是 true
  const statsAfter = d.stats;
  ok('放弃后 inFlight 仍为 true（没被提前放行）', statsAfter.inFlight === true,
     `inFlight=${statsAfter.inFlight}`);
  await sleep(1500);
  ok('子进程跑完后队列恢复', d.stats.inFlight === false, `inFlight=${d.stats.inFlight}`);

  console.log('\n=== E. 取消不是故障：分类要对 ===');
  ok('CancelError 文案被归为 cancelled',
     classifyWpsError('WPS common 请求在排队中被取消（动作=ping）。该请求从未送达，可以安全重发。') === WpsErrorType.CANCELLED);
  ok('"放弃等待"文案也被归为 cancelled',
     classifyWpsError('WPS common 请求已被放弃等待（动作=getDocumentText）。') === WpsErrorType.CANCELLED);
  ok('没把取消误判成 timeout',
     classifyWpsError('请求已被取消（动作=getDocumentText）') !== WpsErrorType.TIMEOUT);

  console.log('\n=== F. 送出前就取消 ===');
  const ac2 = new AbortController();
  ac2.abort();
  let preErr = null;
  await d.call('ping', {}, { signal: ac2.signal }).catch((e) => { preErr = e; });
  ok('已 abort 的 signal 直接抛 CancelError', preErr instanceof CancelError, preErr && preErr.message);

  console.log('\n=== G. 取消计数进了 stats ===');
  ok('stats.cancelled 有累计', d.stats.cancelled >= 2, `cancelled=${d.stats.cancelled}`);

  d.kill('测试结束');
  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})();
