/** 诊断样式族：到底是"参数没到 COM"还是"COM 不吃这个属性" */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-style-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';
const fs = require('fs');
const os = require('os');
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
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '').replace(/\s+/g, ' ');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'style-'));
const shot = async (tag) => {
  const p = path.join(tmp, `${tag}.png`);
  await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: p });
  return fs.existsSync(p) ? fs.readFileSync(p) : null;
};
const raw = (method, params) => wpsClient.executeMethod(method, params, 'ppt');

(async () => {
  await call('wps_ppt_create_presentation', {});
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_shape', { slideIndex: 1, type: 'rectangle', left: 80, top: 80, width: 200, height: 120 });
  const a = await shot('a');
  const b = await shot('b');
  console.log('导出确定:', !!a && !!b && Buffer.compare(a, b) === 0);

  // 1) 工具入口（嵌套 shadow 对象）
  let r = await call('wps_ppt_set_shape_shadow', { slideIndex: 1, shapeIndex: 1, shadow: { enabled: true, color: '#FF0000', blur: 20, distance: 10 } });
  let after = await shot('c');
  console.log('工具(nested shadow):', r.success, txt(r).slice(0, 50), '| 图变:', !!a && !!after && Buffer.compare(a, after) !== 0);

  // 2) 裸调扁平参数
  r = await raw('setShapeShadow', { slideIndex: 1, shapeIndex: 1, color: '#00FF00', blur: 20, distance: 10 });
  let after2 = await shot('d');
  console.log('裸调(扁平):', r.success, '| 图变:', !!after && !!after2 && Buffer.compare(after, after2) !== 0, JSON.stringify(r.data || r.error).slice(0, 80));

  // 3) 直接设边框与填充，验证"形状属性改得动"
  r = await raw('setShapeBorder', { slideIndex: 1, shapeIndex: 1, color: '#0000FF', width: 6 });
  let after3 = await shot('e');
  console.log('裸调边框:', r.success, '| 图变:', !!after2 && !!after3 && Buffer.compare(after2, after3) !== 0);

  r = await raw('setShapeGradient', { slideIndex: 1, shapeIndex: 1, color1: '#FF0000', color2: '#0000FF' });
  let after4 = await shot('f');
  console.log('裸调渐变:', r.success, '| 图变:', !!after3 && !!after4 && Buffer.compare(after3, after4) !== 0);

  await call('wps_ppt_close_presentation', { save: false });
  fs.rmSync(tmp, { recursive: true, force: true });
  process.exit(0);
})();
