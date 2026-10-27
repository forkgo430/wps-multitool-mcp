/**
 * 主线 · ps1 增补轮真机探针驱动：
 *   1) __probeCaptionSeq（word 宿主）：SEQ 域三形态 + TOF 收集条目数读回 + Bibliography（目标 A/B1）
 *   2) __probeDeckkit（ppt 宿主）：七坑段（含坑 1 圆角矩复测——目标 B2）
 * 内部自带收尾（临时文档/演示 Close）；锁由会话层持有。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-caption-seq-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '300000';

const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

(async () => {
  const out = {};
  console.log('[probe] __probeCaptionSeq（word 通道）…');
  const r1 = await wpsClient.executeMethod('__probeCaptionSeq', {}, 'word');
  out.captionSeq = r1;
  console.log('[probe] captionSeq success=' + r1.success);

  console.log('[probe] __probeDeckkit（ppt 通道，含坑 1 圆角矩复测）…');
  const r2 = await wpsClient.executeMethod('__probeDeckkit', {}, 'ppt');
  out.deckkit = r2;
  console.log('[probe] deckkit success=' + r2.success);

  const outPath = path.join(process.env.TEMP || '.', 'wps-round7', 'caption-seq-result.json');
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify(out, null, 2));
  console.log('[probe] 结果已写', outPath);
  process.exit(r1.success && r2.success ? 0 : 1);
})().catch((e) => {
  console.error('[probe] 异常：', e && e.message ? e.message : e);
  process.exit(2);
});
