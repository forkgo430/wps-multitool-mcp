// 第 1 期第 2 章第 1 轮 · 验证者 补充核查：phantom 路由条目 + 现场残留
const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const raw = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8').replace(/^﻿/, '').split(/\r?\n/);
const switchIdx = raw.findIndex((l) => l.trim().startsWith('switch ($Action)'));
const top = new Set();
for (let i = switchIdx + 1; i < raw.length; i++) {
  const m = raw[i].match(/^\s*"([A-Za-z_]\w*)"\s*\{\s*$/);
  if (m && /^\s{4}"/.test(raw[i])) top.add(m[1]);
}
const RT = require(path.join(ROOT, 'dist', 'client', 'action-routes.generated.js'));
const R = RT.ACTION_ROUTES || (RT.default && RT.default.ACTION_ROUTES);
const phantom = Object.keys(R).filter((k) => !top.has(k));
console.log(`顶层可分发动作数: ${top.size} ｜ 路由表条目: ${Object.keys(R).length}`);
console.log(`phantom（路由表有、顶层 switch 没有）共 ${phantom.length} 个:`);
phantom.forEach((k) => console.log(`  "${k}" -> ${R[k]}`));

// 现场残留检查（独立短桥）
const PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-clean-${Date.now()}`;
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
(async () => {
  const x = await call('wps_execute_method', { method: 'getOpenWorkbooks', appType: 'excel', params: {} });
  const w = await call('wps_execute_method', { method: 'getOpenDocuments', appType: 'word', params: {} });
  const p = await call('wps_execute_method', { method: 'getOpenPresentations', appType: 'presentation', params: {} });
  console.log('残留  Excel:', (JSON.parse(txt(x)).data.workbooks || []).length,
    ' Word:', (JSON.parse(txt(w)).data.documents || []).length,
    ' PPT:', (JSON.parse(txt(p)).data.presentations || []).length);
  process.exit(0);
})();
