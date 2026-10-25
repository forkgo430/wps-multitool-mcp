/** 一次性诊断：WPS 里 MainSequence.AddEffect 到底有没有生效 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-anim-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => ((r.content && r.content[0] && r.content[0].text) || r.error || '').replace(/\s+/g, ' ');
(async () => {
  await call('wps_ppt_create_presentation', {});
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_shape', { slideIndex: 1, type: 'rectangle', left: 50, top: 50, width: 100, height: 60 });
  await call('wps_ppt_set_shape_text', { slideIndex: 1, shapeIndex: 1, text: 'A' });
  const before = await wpsClient.executeMethod('getAnimations', { slideIndex: 1 }, 'ppt');
  console.log('before:', JSON.stringify(before).slice(0, 160));
  const add = await wpsClient.executeMethod('addAnimation', { slideIndex: 1, shapeIndex: 1, effect: 'fadeIn' }, 'ppt');
  console.log('addAnimation:', JSON.stringify(add).slice(0, 240));
  const after = await wpsClient.executeMethod('getAnimations', { slideIndex: 1 }, 'ppt');
  console.log('after:', JSON.stringify(after).slice(0, 200));
  await call('wps_ppt_close_presentation', { save: false });
  process.exit(0);
})();
