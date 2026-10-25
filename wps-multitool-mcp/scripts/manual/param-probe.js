/**
 * Input: 无
 * Output: 确认参数是否完整送达 PowerShell（用 __probeParams 回显）
 * Pos: 参数链路诊断脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-probe';
process.env.WPS_BRIDGE_IDLE_MS = '60000';

const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

(async () => {
  const cases = [
    ['纯 ASCII 多参数', { count: 2, spacing: 1.5, lineBetween: true }],
    ['含中文', { text: '第一章 绪论', level: 2 }],
    ['数组', { items: ['第一项', '第二项', '第三项'], type: 'number' }],
    ['混合', { eastAsia: '宋体', ascii: 'Times New Roman', size: 12 }],
  ];
  for (const [name, params] of cases) {
    try {
      const r = await psDaemon.call('__probeParams', params, 'common');
      const d = r.data || {};
      console.log(`--- ${name} ---`);
      console.log(`  发  : ${JSON.stringify(params)}`);
      console.log(`  收  : ${JSON.stringify(d.received)}`);
      console.log(`  键  : ${JSON.stringify(d.keys)}  类型=${d.ptype}`);
    } catch (e) {
      console.log(`--- ${name} --- 调用失败: ${e.message}`);
    }
  }
  psDaemon.dispose();
  process.exit(0);
})();
