/**
 * 主线 · 引擎第一步探针驱动（走真通道：dist wpsClient → 桥 → ps1 __probeDeckkit，ppt 宿主）。
 * 输出：__probeDeckkit 的分段 JSON 原样落 %TEMP%\wps-round7\deckkit-probe-result.json。
 * 纪律：建锁由调用方负责（本脚本不建锁——真机探针统一由会话层持 WPS.lock）。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-deckkit-probe-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '300000';

const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

(async () => {
  console.log('[deckkit-probe] 调 __probeDeckkit（ppt 通道，冷启动可能 30s+）…');
  const t0 = Date.now();
  const r = await wpsClient.executeMethod('__probeDeckkit', {}, 'ppt');
  const dt = Date.now() - t0;
  console.log(`[deckkit-probe] 完成，耗时 ${dt}ms，success=${r.success}`);
  const out = path.join(process.env.TEMP || '.', 'wps-round7', 'deckkit-probe-result.json');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(r, null, 2));
  console.log('[deckkit-probe] 结果已写', out);
  console.log(JSON.stringify(r, null, 2).slice(0, 400));
  process.exit(r.success ? 0 : 1);
})().catch((e) => {
  console.error('[deckkit-probe] 异常：', e && e.message ? e.message : e);
  process.exit(2);
});
