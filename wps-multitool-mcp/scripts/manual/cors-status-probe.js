/**
 * Input: 独立桥管道（WPS_BRIDGE_PIPE 独立名 + EPHEMERAL）+ WPS 加载项面板 panel.html
 * Output: /status 三种 Origin 的响应头原文 + 面板页面能否读到状态 + P0/P1/正常 结论
 * Pos: CORS 收紧（'*' → 'null'）后的真机回归探测。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 背景：统筹者把桥 /status 的 ACAO 从 '*' 收紧成 'null'（src/bridge/index.ts:204），
 * 理由是"任何网页都能读进程与日志"。改完没验证过 WPS 任务窗格还能不能读状态。
 *
 * 这个脚本回答三个问题：
 *   1. 三种 Origin（null / https://evil.example / 不带）下，/status 的 ACAO 响应头分别是什么；
 *   2. 浏览器按 CORS 规则分别会放行还是拦截；
 *   3. 加载项面板页面（panel.html，file:// 页面 fetch http://127.0.0.1:<port>/status）
 *      在收紧后的策略下还能不能把状态读出来、渲染出来。
 *
 * 安全边界：
 *   - 只用独立管道名（...-agentC-cors）与 EPHEMERAL=1，用完自己 kill 桥，不碰用户的桥；
 *   - HTTP 只绑 127.0.0.1；面板测试用的是 panel.html 的【副本】（临时目录），不动已安装面板；
 *   - 不向 WPS 发任何写动作（只 ping 起桥），不改任何文档。
 *
 * 用法（需先 npm run build）：node scripts/manual/cors-status-probe.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const { execFile, spawn } = require('child_process');
const { pathToFileURL } = require('url');

// ---- 独立管道：与用户正在用的桥彻底分开（EPHEMERAL 让桥随本进程退出而退出）----
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-agentC-cors';
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '180000';
process.env.WPS_BRIDGE_START_TIMEOUT = process.env.WPS_BRIDGE_START_TIMEOUT || '60000';

const ROOT = path.resolve(__dirname, '..', '..');
const PORT = parseInt(process.env.WPS_BRIDGE_HTTP_PORT || '27731', 10);
const SPY_PORT = PORT + 1;
const PANEL_SRC = path.join(
  process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'),
  'kingsoft',
  'wps',
  'jsaddons',
  'wps-claude-addon_',
  'panel.html'
);
const MSEDGE = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
].find((p) => fs.existsSync(p));

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 端口是否空闲（没人监听） */
function portFree(port) {
  return new Promise((resolve) => {
    const s = http.request({ host: '127.0.0.1', port, path: '/', method: 'GET', timeout: 800 }, (res) => {
      res.resume();
      resolve(false); // 有响应 = 被占用
    });
    s.on('error', () => resolve(true));
    s.on('timeout', () => {
      s.destroy();
      resolve(true); // 没监听才会走到这里
    });
    s.end();
  });
}

/** 发一次原始请求，返回 { status, headers: Node 归一化后的小写头, body } */
function rawGet(port, urlPath, headers) {
  return new Promise((resolve, reject) => {
    const req = http.request({ host: '127.0.0.1', port, path: urlPath, method: 'GET', headers }, (res) => {
      let body = '';
      res.on('data', (c) => (body += c));
      // 用 res.headers（Node 已把小写归一）—— rawHeaders 保留服务端原始大小写，不能直接查小写键
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body }));
    });
    req.on('error', reject);
    req.end();
  });
}

function printHeaders(label, r, keys) {
  console.log(`  ${label}`);
  console.log(`    HTTP ${r.status}`);
  for (const k of keys) console.log(`    ${k}: ${r.headers[k] === undefined ? '(缺失)' : r.headers[k]}`);
}

/** 按浏览器 CORS 规则判定：返回 '放行' | '拦截' | '非CORS请求(直接可读)' */
function browserVerdict(origin, acao) {
  if (origin === undefined) return '非CORS请求（curl/服务端，直接可读）';
  if (acao === undefined) return '拦截（响应没有 ACAO）';
  if (acao === origin || acao === '*') return '放行';
  return '拦截';
}

/** 间谍服务：复刻桥的 CORS 策略（ACAO: null），记录每个请求的 Origin */
function startSpy() {
  const seen = [];
  const server = http.createServer((req, res) => {
    seen.push({ url: req.url, origin: req.headers.origin === undefined ? null : req.headers.origin });
    res.writeHead(200, {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': 'null', // 与 src/bridge/index.ts:204 完全一致
    });
    res.end(
      JSON.stringify({
        bridgePid: process.pid,
        pipe: 'spy',
        httpPort: SPY_PORT,
        startedAt: Date.now() - 5000,
        uptimeMs: 5000,
        clients: 1,
        inFlight: 0,
        queued: 0,
        queueMax: 3,
        queueTimeoutMs: 120000,
        appsRunning: ['common'],
        children: [{ app: 'common', pid: process.pid, ready: true }],
        recent: [{ t: Date.now() - 1000, app: 'common', action: 'ping', ok: true, summary: 'ping@spy' }],
      })
    );
  });
  return new Promise((resolve, reject) => {
    server.on('error', reject);
    server.listen(SPY_PORT, '127.0.0.1', () => resolve({ server, seen }));
  });
}

/** 把已安装的 panel.html 复制到临时目录，并注入 __PORT__ 指向目标端口 */
function makePanelCopy(port) {
  const html = fs.readFileSync(PANEL_SRC, 'utf8');
  const injected = html.replace(
    '<meta charset="utf-8">',
    `<meta charset="utf-8">\n<script>window.__PORT__=${port};</script>`
  );
  if (injected === html) throw new Error('panel.html 注入点没找到（<meta charset> 行变了？）');
  const out = path.join(os.tmpdir(), `wps-cors-panel-${port}-${Date.now()}.html`);
  fs.writeFileSync(out, injected, 'utf8');
  return out;
}

/** 用 headless Chromium 打开面板副本，dump DOM */
function dumpDom(filePath, timeoutMs = 60000) {
  return new Promise((resolve, reject) => {
    const args = [
      '--headless=new',
      '--disable-gpu',
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-extensions',
      `--user-data-dir=${path.join(os.tmpdir(), `wps-cors-msedge-${Date.now()}`)}`,
      '--virtual-time-budget=10000',
      '--dump-dom',
      pathToFileURL(filePath).href,
    ];
    const child = spawn(MSEDGE, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '';
    let err = '';
    const timer = setTimeout(() => child.kill(), timeoutMs);
    child.stdout.on('data', (c) => (out += c));
    child.stderr.on('data', (c) => (err += c));
    child.on('error', (e) => {
      clearTimeout(timer);
      reject(e);
    });
    child.on('close', () => {
      clearTimeout(timer);
      resolve({ out, err });
    });
  });
}

(async () => {
  console.log('\n=== CORS 回归探测：桥 /status + WPS 加载项面板 ===');
  console.log(`管道：${process.env.WPS_BRIDGE_PIPE}`);
  console.log(`HTTP 端口：${PORT}（默认 27731，可被 WPS_BRIDGE_HTTP_PORT 覆盖）`);

  if (!fs.existsSync(PANEL_SRC)) {
    console.log(`  ⚠ 找不到已安装的面板：${PANEL_SRC}（第 3 项只能跳过，不算通过也不算失败）`);
  }
  if (!MSEDGE) {
    console.log('  ⚠ 找不到 msedge.exe，面板复刻测试只能跳过（不算通过也不算失败）');
  }

  // ---- 起独立桥（只 ping，不碰任何文档）----
  const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));
  let bridgePid = null;
  try {
    const t0 = Date.now();
    const pong = await psDaemon.call('ping', {}, 'common');
    bridgePid = psDaemon.stats.bridgePid;
    console.log(`  ✓ 桥已起（pid=${bridgePid}，ping ${Date.now() - t0}ms，success=${pong.success}）`);
  } catch (e) {
    console.log(`  ⚠ 桥起不来：${String(e.message || e).slice(0, 80)}`);
    console.log('    （HTTP 端点无从测起。本轮跳过，不判失败 —— 但 /status 的 CORS 行为【未验证】）');
    psDaemon.dispose();
    process.exit(0);
  }

  // 等 HTTP 端点就绪
  let up = false;
  for (let i = 0; i < 40; i++) {
    try {
      const r = await rawGet(PORT, '/status');
      if (r.status === 200) {
        up = true;
        break;
      }
    } catch {
      /* 还没起来 */
    }
    await sleep(250);
  }
  if (!up) {
    console.log(`  ⚠ ${PORT} 端口 10 秒内没起来（可能被占用，见 src/bridge/index.ts:212 的 EADDRINUSE 静默退开）`);
    psDaemon.dispose();
    process.exit(0);
  }
  console.log(`  ✓ HTTP 端点已就绪：http://127.0.0.1:${PORT}/status`);

  const HEADER_KEYS = ['access-control-allow-origin', 'cache-control', 'content-type'];

  // ---- 1) 三种 Origin 的原始响应头 ----
  console.log('\n--- 1. 三种 Origin 的响应头原文 ---');
  const cases = [
    { label: 'Origin: null（file:// 页面就是这种，WPS 面板即 file://）', headers: { Origin: 'null' } },
    { label: 'Origin: https://evil.example（任意网页）', headers: { Origin: 'https://evil.example' } },
    { label: '不带 Origin（curl / 服务端 / 同源）', headers: {} },
  ];
  const results = cases.map((c) => ({ ...c, res: null }));
  for (const c of results) {
    c.res = await rawGet(PORT, '/status', c.headers);
    printHeaders(c.label, c.res, HEADER_KEYS);
    const acao = c.res.headers['access-control-allow-origin'];
    const origin = c.headers.Origin;
    console.log(`    浏览器判定：${browserVerdict(origin, acao)}`);
    console.log(`    响应体长度：${c.res.body.length} 字节（能读到就说明请求本身通了）`);
  }

  const acaoNull = results[0].res.headers['access-control-allow-origin'];
  const acaoEvil = results[1].res.headers['access-control-allow-origin'];
  const acaoNone = results[2].res.headers['access-control-allow-origin'];

  // ---- 2) 面板复刻：file:// 页面 fetch /status ----
  console.log('\n--- 2. 面板复刻（headless Chromium 加载同一份 panel.html）---');
  let panelOk = null;
  if (fs.existsSync(PANEL_SRC) && MSEDGE) {
    // 2a. 先打间谍服务：抓面板页面【真实发出】的 Origin 头
    try {
      const spy = await startSpy();
      const spyCopy = makePanelCopy(SPY_PORT);
      const spyDom = await dumpDom(spyCopy);
      fs.unlinkSync(spyCopy);
      const origins = [...new Set(spy.seen.map((s) => s.origin))];
      console.log(`  · 间谍服务收到 ${spy.seen.length} 次请求，Origin 取值：${JSON.stringify(origins)}`);
      spy.server.close();

      // 2b. 再打真桥：看收紧后的 ACAO 下页面能不能真的把状态渲染出来
      const realCopy = makePanelCopy(PORT);
      const dom = await dumpDom(realCopy);
      fs.unlinkSync(realCopy);
      // 注意：--dump-dom 输出含【内联脚本源码】，所以"连不上桥"字样一定会出现在源码里，
      // 不能拿它当失败信号 —— 只看渲染后的 span 内容。
      const spanText = (id) => {
        const m = new RegExp(`<span[^>]*id="${id}"[^>]*>([\\s\\S]*?)<\\/span>`).exec(dom.out);
        return m ? m[1].replace(/<[^>]+>/g, '').trim() : null;
      };
      const logText = (() => {
        const m = /<div class="log" id="log">([\s\S]*?)<\/div>/.exec(dom.out);
        return m ? m[1].replace(/<[^>]+>/g, ' ').trim() : null;
      })();
      const vBridge = spanText('v-bridge');
      const vQueue = spanText('v-queue');
      const vGroup = spanText('v-group');
      const connected = !!vBridge && /已连接/.test(vBridge);
      const hasQueue = !!vQueue && /\d+\s*\/\s*\d+/.test(vQueue);
      const hasGroup = !!vGroup && /基础/.test(vGroup) && /共/.test(vGroup);
      panelOk = connected && hasQueue && hasGroup;
      console.log(`  · 面板渲染「桥进程」= ${vBridge}`);
      console.log(`  · 面板渲染「队列」  = ${vQueue}`);
      console.log(`  · 面板渲染「工具组」= ${vGroup}`);
      console.log(`  · 面板渲染「最近改动」= ${logText ? logText.replace(/\s+/g, ' ').slice(0, 80) : '(空)'}`);
      console.log(`  · 判定：已连接=${connected} 队列=${hasQueue} 工具组=${hasGroup} → ${panelOk ? '面板正常' : '面板读不到状态'}`);
      if (dom.err.trim()) console.log(`  · msedge stderr（前 200 字）：${dom.err.trim().slice(0, 200)}`);
    } catch (e) {
      console.log(`  ⚠ 面板复刻测试自身出错：${String(e.message || e).slice(0, 80)}`);
    }
  }

  // ---- 收尾：杀掉自己起的桥，绝不碰用户的桥 ----
  try {
    psDaemon.dispose();
  } catch {
    /* ignore */
  }
  await sleep(500);
  const killed = await new Promise((resolve) => {
    execFile(
      'powershell',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        `Get-CimInstance Win32_Process -Filter "Name='node.exe'" -EA SilentlyContinue | Where-Object { $_.CommandLine -like '*agentC-cors*' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -EA SilentlyContinue; 1 } | Measure-Object | Select-Object -ExpandProperty Count`,
      ],
      (e, out) => resolve(e ? -1 : parseInt((out || '').trim(), 10) || 0)
    );
  });
  console.log(`\n收尾：EPHEMERAL 断开 + 补杀残留桥进程 ${killed} 个（只匹配 agentC-cors 管道名）`);

  // ---- 结论 ----
  console.log('\n=== 判定 ===');
  const corsTightened = acaoNull === 'null';
  const evilBlocked = browserVerdict('https://evil.example', acaoEvil) === '拦截';
  const nullAllowed = browserVerdict('null', acaoNull) === '放行';
  const noOriginReadable = browserVerdict(undefined, acaoNone).includes('可读');
  console.log(`  ACAO(null 来源)=${acaoNull} → ${nullAllowed ? '面板这类 file:// 页面放行' : '不放行'}`);
  console.log(`  ACAO(evil 来源)=${acaoEvil} → ${evilBlocked ? '恶意站点被拦' : '没拦住'}`);
  console.log(`  ACAO(无 Origin)=${acaoNone} → ${noOriginReadable ? '可读（非浏览器来源）' : '异常'}`);

  let verdict;
  if (panelOk === false) verdict = 'P0';
  else if (panelOk === null) verdict = '未验证（面板复刻没跑成）';
  else if (!corsTightened || !evilBlocked || !nullAllowed) verdict = 'P1';
  else verdict = '正常';
  console.log(`\n结论：${verdict}`);
  console.log(`PASSED=cors-status-probe:${verdict}`);
  process.exit(verdict === 'P0' ? 1 : 0);
})().catch((e) => {
  console.error('探测失败:', e && e.message ? e.message : String(e));
  process.exit(1);
});
