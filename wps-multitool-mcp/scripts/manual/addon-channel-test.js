/**
 * 加载项通道骨架的门禁（Linux B2 验收项）
 * 原则：不碰真 WPS —— 用 Node http 客户端模拟加载项，验计划 §8-B2 的验收三态
 * （握手成功/失败/超时）+ 两条硬约束（端口只绑 127.0.0.1、失败语义照抄 ps-daemon 契约）。
 * 跑法：node scripts/manual/addon-channel-test.js（verify-all FAST 档包含它）
 *
 * 为什么能不起 WPS：通道本体（HTTP 服务 + 会话/任务位/失败语义）全在 Node 侧，
 * 加载项侧只是"会 poll、会回 result"的远端 —— 用 http 模拟即可钉死协议。
 * 加载项侧的分发器（channel-client.js）是纯 JS，直接 require 进来单测。
 */

// callTimeoutMs() 每次调用时重读 env —— 在 require dist 之前收紧，测试不用等 120s
process.env.WPS_DAEMON_TIMEOUT = '700';

const path = require('path');
const http = require('http');
const net = require('net');
const os = require('os');

const ROOT = path.resolve(__dirname, '..', '..');
const ch = require(path.join(ROOT, 'dist', 'client', 'addon-channel.js'));
// 加载项侧分发器：浏览器 IIFE 在 CommonJS 里挂到 module.exports（window 不存在时取 this）
const clientJs = require(path.join(ROOT, '..', 'wps-multitool-addon-linux', 'js', 'channel-client.js'));

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};

/** 测试专令牌（红线 4：只允许硬编码测试专令牌） */
const TEST_TOKEN = 'addon-channel-test-token';
const CALL_TIMEOUT = parseInt(process.env.WPS_DAEMON_TIMEOUT, 10);

function randomPort() {
  return 27000 + Math.floor(Math.random() * 1000);
}

function request(method, port, pathName, body) {
  return new Promise((resolve) => {
    const data = body === undefined ? null : JSON.stringify(body);
    const req = http.request(
      {
        host: '127.0.0.1',
        port,
        path: pathName,
        method,
        headers: data ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) } : {},
      },
      (res) => {
        let buf = '';
        res.on('data', (c) => (buf += c));
        res.on('end', () => {
          let parsed = null;
          try { parsed = JSON.parse(buf); } catch { /* 非 JSON */ }
          resolve({ status: res.statusCode, body: parsed, raw: buf });
        });
      }
    );
    req.on('error', (e) => resolve({ status: 0, error: e }));
    if (data) req.write(data);
    req.end();
  });
}

/** 给任何 Promise 加观察超时：门禁自己不许挂死 */
function withTimeout(p, ms, label) {
  return Promise.race([
    p,
    new Promise((resolve) => setTimeout(() => resolve({ __hung: true, label }), ms)),
  ]);
}

function tryConnect(host, port) {
  return new Promise((resolve) => {
    const sock = net.connect({ host, port });
    const done = (r) => { sock.destroy(); resolve(r); };
    sock.setTimeout(1500, () => done('timeout'));
    sock.on('connect', () => done('connected'));
    sock.on('error', (e) => done('error:' + (e.code || e.message)));
  });
}

/**
 * 模拟加载项会话：握手后开 poll 循环。
 * onFrame(frame) 返回要回给 /result 的 body（不含 sessionId/id），返回 null 则不回（挂住）。
 */
function makeFakeAddon(port, type, onFrame) {
  const st = { sessionId: '', frames: [], running: true, resultBodies: [] };
  const handshake = () =>
    request('POST', port, '/handshake', { type, token: TEST_TOKEN, addonVersion: 'test-0' }).then((r) => {
      st.sessionId = (r.body && r.body.sessionId) || '';
      return st.sessionId;
    });
  const poll = () => {
    if (!st.running || !st.sessionId) return;
    request('GET', port, '/poll?sessionId=' + encodeURIComponent(st.sessionId)).then((r) => {
      if (!st.running) return;
      if (r.body && r.body.error) { st.running = false; return; } // 404/会话失效：停止（否则死循环刷 404）
      if (!r.body || r.body.idle) { setTimeout(poll, 5); return; }
      const frame = r.body;
      st.frames.push(frame);
      const reply = onFrame(frame);
      if (reply === null) return; // 挂住：收了帧不回结果（失败语义用例）
      const body = { sessionId: st.sessionId, id: frame.id, ...reply };
      st.resultBodies.push(body);
      request('POST', port, '/result', body).then(() => setTimeout(poll, 0));
    });
  };
  return handshake().then((sessionId) => {
    setTimeout(poll, 0);
    return st;
  });
}

(async () => {
  console.log('\n=== 0. 起服务（随机端口，显式测试令牌）===');
  const rnd = randomPort();
  let info = null;
  for (let i = 0; i < 5 && !info; i++) {
    try { info = await ch.start({ port: rnd + i, token: TEST_TOKEN }); }
    catch (e) { console.log(`  端口 ${rnd + i} 被占（${e.code || e.message}），换一个`); }
  }
  if (!info) { console.log('  ✗ 起服务失败（重试 5 次全被占）'); process.exit(1); }
  const port = info.port;
  ok('start 返回显式令牌且不标"生成"', info.token === TEST_TOKEN && info.tokenGenerated === false, `port=${port}`);

  console.log('\n=== 断言 4：端口绑定（只监听 127.0.0.1）===');
  const lo = await tryConnect('127.0.0.1', port);
  ok('127.0.0.1 可连（正向）', lo === 'connected', lo);
  const v6 = await tryConnect('::1', port);
  ok('::1（IPv6 环回）不可连 —— 绑定不含 IPv6', v6 !== 'connected', v6);
  const nicIps = [];
  for (const list of Object.values(os.networkInterfaces())) {
    for (const ni of list || []) {
      if (ni.family === 'IPv4' && !/^127\./.test(ni.address)) nicIps.push(ni.address);
    }
  }
  if (nicIps.length) {
    const r = await tryConnect(nicIps[0], port);
    ok(`非环回地址 ${nicIps[0]} 不可连 —— 绑定不含局域网`, r !== 'connected', r);
  } else {
    ok('本机无非环回 IPv4，局域网负向用例跳过（::1 已证明绑定面收紧）', true);
  }

  console.log('\n=== 断言 6a：无会话 → AddonUnavailableError（请求从未送出）===');
  const noSession = await withTimeout(ch.call('getValue', {}, 'excel').then(() => ({ settled: 'resolved' }), (e) => ({ settled: 'rejected', name: e.name })), 3000, 'no-session');
  ok('无 et 会话时 call 立即判未送出', noSession.settled === 'rejected' && noSession.name === 'AddonUnavailableError',
    JSON.stringify(noSession));

  console.log('\n=== 断言 2：握手失败（错 token → 401；错 type → 400）===');
  const badTok = await request('POST', port, '/handshake', { type: 'wps', token: 'wrong-token', addonVersion: 't' });
  ok('错 token → 401', badTok.status === 401, `status=${badTok.status}`);
  const badType = await request('POST', port, '/handshake', { type: 'foo', token: TEST_TOKEN, addonVersion: 't' });
  ok('错 type → 400', badType.status === 400, `status=${badType.status}`);
  const noTok = await request('POST', port, '/handshake', { type: 'wps', addonVersion: 't' });
  ok('缺 token → 401（undefined ≠ 令牌）', noTok.status === 401, `status=${noTok.status}`);

  console.log('\n=== 断言 1：握手成功（三型各登记一个会话）===');
  const sessionIds = {};
  for (const t of ['wps', 'et', 'wpp']) {
    const r = await request('POST', port, '/handshake', { type: t, token: TEST_TOKEN, addonVersion: 'test-1' });
    sessionIds[t] = r.body && r.body.sessionId;
    ok(`握手 ${t} → 200 + sessionId`, r.status === 200 && typeof sessionIds[t] === 'string' && sessionIds[t].length > 10,
      `pollIntervalMs=${r.body && r.body.pollIntervalMs}`);
  }
  const st0 = ch.status();
  ok('status() 三型都报已连接', st0.transport === 'long-poll' && st0.started &&
    st0.sessions.wps.connected && st0.sessions.et.connected && st0.sessions.wpp.connected,
    JSON.stringify({ wps: st0.sessions.wps.connected, et: st0.sessions.et.connected, wpp: st0.sessions.wpp.connected }));

  console.log('\n=== 断言 5：路由（word→wps、excel→et、ppt→wpp、common→wps）===');
  const seen = { wps: [], et: [], wpp: [] };
  const addons = {};
  for (const t of ['wps', 'et', 'wpp']) {
    addons[t] = await makeFakeAddon(port, t, (frame) => ({ ok: true, data: { handledBy: t, app: frame.app, action: frame.action } }));
    ok(`假加载项 ${t} 已握手，Node 侧会话一致`, addons[t].sessionId.length > 10 &&
      ch.status().sessions[t].sessionId === addons[t].sessionId,
      `sessionId=${addons[t].sessionId.slice(0, 8)}…`);
  }
  for (const [app, wantType] of [['word', 'wps'], ['excel', 'et'], ['ppt', 'wpp'], ['common', 'wps']]) {
    const r = await withTimeout(ch.call('routeProbe', { seq: app }, app), 3000, `route-${app}`);
    ok(`call(app=${app}) 由 ${wantType} 型会话处理`,
      r && r.handledBy === wantType && r.app === app,
      JSON.stringify(r && r.__hung ? { hung: true } : r));
  }

  console.log('\n=== 断言 6b：已送出的动作超时不重放（帧只送达一次）===');
  // wpp 换一个"收了帧不回结果"的假加载项（新握手覆盖旧会话，旧 poller 收 404 自停）
  let wppGotFrame = null;
  addons.wpp.running = false; // 停旧 poller
  const wppHung = await makeFakeAddon(port, 'wpp', (frame) => { wppGotFrame = frame; return null; });
  const hungCall = await withTimeout(
    ch.call('slowAction', {}, 'ppt').then(() => ({ settled: 'resolved' }), (e) => ({ settled: 'rejected', name: e.name, msg: e.message })),
    CALL_TIMEOUT + 2500,
    'inflight-timeout'
  );
  ok('已送达加载项、无回包 → 超时如实抛（普通 Error，不是 AddonUnavailableError）',
    hungCall.settled === 'rejected' && hungCall.name === 'Error' && /不.*重放|执行了一半/.test(hungCall.msg || ''),
    hungCall.settled === 'rejected' ? (hungCall.msg || '').slice(0, 70) : JSON.stringify(hungCall));
  ok('帧确实送达过一次', !!wppGotFrame && wppGotFrame.action === 'slowAction', `frame.action=${wppGotFrame && wppGotFrame.action}`);
  // 超时后任务位已清：用【这个挂住加载项自己的会话】再 poll，应拿不到同一帧（不重放）
  const repoll = await request('GET', port, '/poll?sessionId=' + encodeURIComponent(wppHung.sessionId));
  ok('超时后再 poll 不再下发同一帧（调用次数=1，绝不重放）', !!(repoll.body && repoll.body.idle),
    JSON.stringify(repoll.body && repoll.body.idle ? 'idle' : repoll.body));

  console.log('\n=== 断言 6c：加载项回 ok:false → 如实抛普通 Error ===');
  // 6b 的 poller 收帧后挂住不退（onFrame 返回 null 就停了）——这里重新挂一个"会回错误"的
  addons.wpp.running = false;
  await makeFakeAddon(port, 'wpp', (frame) => ({ ok: false, error: 'boom: ' + frame.action }));
  const errCall = await withTimeout(
    ch.call('boomAction', {}, 'ppt').then(() => ({ settled: 'resolved' }), (e) => ({ settled: 'rejected', msg: e.message })),
    3000, 'boom'
  );
  ok('ok:false → reject 且错误原文带回', errCall.settled === 'rejected' && /boomAction/.test(errCall.msg || ''),
    (errCall.msg || '').slice(0, 70));

  console.log('\n=== 断言 3：超时三态的另一支 —— 会话在、加载项不 poll → 未送出超时 ===');
  await request('POST', port, '/handshake', { type: 'et', token: TEST_TOKEN, addonVersion: 'test-2' }); // 覆盖 et，无 poller
  const orphan = await withTimeout(
    ch.call('nobodyPolls', {}, 'excel').then(() => ({ settled: 'resolved' }), (e) => ({ settled: 'rejected', name: e.name, msg: e.message })),
    CALL_TIMEOUT + 2500, 'orphan'
  );
  ok('任务无人取走 → 超时判"从未送出"（AddonUnavailableError），不是挂死',
    orphan.settled === 'rejected' && orphan.name === 'AddonUnavailableError',
    orphan.settled === 'rejected' ? (orphan.msg || '').slice(0, 70) : JSON.stringify(orphan));

  console.log('\n=== 断言 7：轮询闭环（Node → 帧 → 加载项 → result → Node resolve）===');
  const ping = await withTimeout(ch.call('ping', {}, 'word'), 3000, 'ping');
  ok('端到端闭环：call 拿到加载项回传数据（word 动作由 wps 会话处理并回传）',
    ping && !ping.__hung && ping.action === 'ping' && ping.handledBy === 'wps',
    JSON.stringify(ping && ping.__hung ? { hung: true } : ping).slice(0, 90));

  console.log('\n=== 附：加载项侧分发器（channel-client.js）骨架行为 ===');
  // B4：dispatch Promise 化（动作层异步 jsAsynCall），此段全部 await——28 断言数不变
  const d = clientJs.WpsAddonChannel.dispatch;
  const p1 = await d('ping', {}, 'word');
  ok("dispatch('ping') 真实现", p1.ok === true && p1.data.pong === true, JSON.stringify(p1.data).slice(0, 60));
  const p2 = await d('wireCheck', {}, 'word');
  ok("dispatch('wireCheck') 真实现", p2.ok === true && typeof p2.data.message === 'string', (p2.data.message || '').slice(0, 60));
  const p3 = await d('insertText', { text: 'x' }, 'word');
  ok('其余动作一律 unsupported:B3（显式拒绝，不静默成功）', p3.ok === false && p3.unsupported === 'B3', `action=insertText → ${JSON.stringify({ ok: p3.ok, unsupported: p3.unsupported })}`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exitCode = fail === 0 ? 0 : 1;
})().catch((e) => {
  console.error('门禁异常：', e && e.message ? e.message : e);
  process.exitCode = 2;
}).finally(async () => {
  // stop 等全部连接收尾（keep-alive 空闲连接/竞态重连可能拖住 close 回调）——兜底 3s 强制收场
  await Promise.race([ch.stop(), new Promise((r) => setTimeout(r, 3000))]);
  process.exit(process.exitCode || 0);
});
