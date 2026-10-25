const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-clean2-${Date.now()}`;
process.env.WPS_BRIDGE_PIPE = PIPE;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '20000';
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => (r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '';
const names = (s) => { try { return (JSON.parse(txt(s)).data.workbooks || JSON.parse(txt(s)).data.documents || []).map((w) => w.name); } catch { return []; } };
(async () => {
  const x = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  const w = await call('wps_execute_method', { method: 'getOpenDocuments', appType: 'word', params: {} });
  const xb = names(x), wd = names(w);
  console.log('关前 Excel:', xb.join(',') || '(空)', '｜ Word:', wd.join(',') || '(空)');
  for (const n of xb) { try { await call('wps_excel_close_workbook', { name: n, saveChanges: false }); } catch {} }
  for (const n of wd) { try { await call('wps_word_close_document', { name: n, save: false }); } catch {} }
  const x2 = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  const w2 = await call('wps_execute_method', { method: 'getOpenDocuments', appType: 'word', params: {} });
  console.log('关后 Excel:', names(x2).length, '｜ Word:', names(w2).length);
  process.exit(0);
})();
