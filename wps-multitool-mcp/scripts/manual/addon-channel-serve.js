/**
 * Input: 环境变量 WPS_ADDON_PORT / WPS_ADDON_TOKEN（可选），命令行端口（可选）
 * Output: 常驻的加载项通道服务（127.0.0.1，默认 27740）
 * Pos: Linux B3 真机联调的独立启动器。MCP server 进程内嵌通道是正式形态（B5），
 *      联调阶段用本脚本把通道单独拉起来，方便部署三型加载项与驱动 20 动作复测。
 *
 * 用法：
 *   WPS_ADDON_TOKEN=<你的令牌> node scripts/manual/addon-channel-serve.js [端口]
 * 未设 WPS_ADDON_TOKEN 时会随机生成并打到 stderr（找 [addon-channel] token=… 那行）。
 *
 * 部署加载项时的 URL 模板（jsplugins.xml / 浏览器打开均按此拼）：
 *   http://<本机>:3889/<型>/?server=http://127.0.0.1:27740&token=<令牌>
 * （channel-client.js 从 URL query 读 server 与 token；3889 是 S1 的静态分发服务器。）
 */
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const addonChannel = require(path.join(ROOT, 'dist', 'client', 'addon-channel.js'));

const port = Number(process.argv[2]) || Number(process.env.WPS_ADDON_PORT) || undefined;

addonChannel
  .start({ port })
  .then(() => {
    console.log('[serve] 加载项通道已启动');
    console.log('[serve] 端口: ' + (process.env.WPS_ADDON_PORT || '27740') + '（仅 127.0.0.1）');
    console.log('[serve] 令牌: 设了 WPS_ADDON_TOKEN 就用它；否则看 stderr 里 [addon-channel] token=… 那行');
    console.log('[serve] 加载项 URL 模板: http://<本机>:3889/<型>/?server=http://127.0.0.1:27740&token=<令牌>');
    console.log('[serve] Ctrl+C 停止');
  })
  .catch((e) => {
    console.error('[serve] 启动失败:', e.message);
    process.exit(1);
  });
