/** 诊断：合并后的背景工具（规范名 + 两个旧别名）到底有没有真的改背景 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-bg2-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';
const path = require('path');
const fs = require('fs');
const os = require('os');
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
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'bg2-'));
const shot = async (n) => {
  const p = path.join(tmp, `${n}.png`);
  await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: p });
  return fs.existsSync(p) ? fs.readFileSync(p) : null;
};
(async () => {
  await call('wps_ppt_create_presentation', {});
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_shape', { slideIndex: 1, type: 'rectangle', left: 40, top: 40, width: 80, height: 50 });
  const a = await shot('a');
  const b = await shot('b');
  console.log('导出确定:', a && b && Buffer.compare(a, b) === 0);

  for (const [name, args] of [
    ['wps_ppt_set_background', { slideIndex: 1, type: 'solid', color: '#FF0000' }],
    ['wps_ppt_set_slide_background', { slideIndex: 1, background: { type: 'solid', color: '#00FF00' } }],
    ['wps_ppt_set_background_color', { slideIndex: 1, color: '#0000FF' }],
  ]) {
    const before = await shot('x');
    const r = await call(name, args);
    const after = await shot('y');
    const changed = before && after && Buffer.compare(before, after) !== 0;
    console.log(`${name.padEnd(30)} success=${r.success} 图有变化=${changed} ${r.success ? '' : '| ' + txt(r).slice(0, 60)}`);
  }
  await call('wps_ppt_close_presentation', { save: false });
  fs.rmSync(tmp, { recursive: true, force: true });
  process.exit(0);
})();
