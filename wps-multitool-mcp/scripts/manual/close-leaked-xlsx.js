// 一次性清理：关闭所有打开的 Excel 工作簿（仅用于收尾泄漏，运行前请确认都是临时产物）。
// 用法：node scripts/manual/close-leaked-xlsx.js
const path = require('path');
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentC-clean-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '60000';
process.env.WPS_DAEMON_TIMEOUT = '30000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';

(async () => {
  const list = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  let names = [];
  try { names = (JSON.parse(txt(list)).data.workbooks || []).map((w) => w.name); } catch {}
  console.log('关前共', names.length, '个:', names.join(','));
  let ok = 0, fail = 0;
  for (const n of names) {
    const cl = await call('wps_execute_method', { method: 'closeWorkbook', appType: 'excel', params: { name: n, saveChanges: false } });
    if (cl.success) ok++; else { fail++; console.log('  关失败', n, ':', txt(cl).slice(0, 50)); }
  }
  const after = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  let remain = 0;
  try { remain = (JSON.parse(txt(after)).data.workbooks || []).length; } catch {}
  console.log(`已关 ${ok}，失败 ${fail}，剩余 ${remain}`);
  process.exit(0);
})().catch((e) => { console.error('失败:', e && e.message ? e.message : String(e)); process.exit(1); });
