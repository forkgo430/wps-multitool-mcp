/**
 * 一次性诊断：按冒烟矩阵的实测顺序，逐条调用，看演示文稿从哪一条开始"消失"。
 * 判据不是工具返回值，而是每次调用后还能不能读到活动文稿（get_open_presentations）。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-cascade-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '90000';
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => ((r.content && r.content[0] && r.content[0].text) || r.error || '').replace(/\s+/g, ' ');

const alive = async () => {
  const r = await call('wps_ppt_get_open_presentations', {});
  const t = txt(r);
  const ok = r.success && /演示文稿|pptx/i.test(t);
  return ok ? 'alive' : `DEAD(${t.slice(0, 50)})`;
};

const steps = [
  ['建靶', async () => {
    await call('wps_ppt_create_presentation', {});
    await call('wps_ppt_add_slide', { layout: 'blank' });
    await call('wps_ppt_add_slide', { layout: 'blank' });
    for (const p of [1, 2]) {
      await call('wps_ppt_add_shape', { slideIndex: p, type: 'rectangle', left: 60, top: 60, width: 120, height: 60 });
      await call('wps_ppt_add_shape', { slideIndex: p, type: 'oval', left: 240, top: 60, width: 120, height: 60 });
    }
    await call('wps_ppt_set_shape_text', { slideIndex: 1, shapeIndex: 1, text: '文字A' });
    return '';
  }],
  ['wps_ppt_beautify', () => call('wps_ppt_beautify', {})],
  ['wps_ppt_unify_font', () => call('wps_ppt_unify_font', { font_name: '宋体' })],
  ['wps_ppt_set_font_color', () => call('wps_ppt_set_font_color', { slideIndex: 1, shapeIndex: 1, color: '#FF0000' })],
  ['wps_ppt_set_slide_layout', () => call('wps_ppt_set_slide_layout', { slideIndex: 1, layout: 'title' })],
  ['wps_ppt_set_shape_style', () => call('wps_ppt_set_shape_style', { slideIndex: 1, shapeIndex: 1, fillColor: '#FF0000' })],
  ['wps_ppt_set_animation', () => call('wps_ppt_set_animation', { slideIndex: 1, shapeIndex: 1, animationType: 'fadeIn' })],
  ['wps_ppt_get_shapes', () => call('wps_ppt_get_shapes', { slideIndex: 1 })],
  ['wps_ppt_set_background', () => call('wps_ppt_set_background', { slideIndex: 1, type: 'solid', color: '#FF0000' })],
  ['wps_ppt_set_slide_size', () => call('wps_ppt_set_slide_size', { width: 960, height: 540 })],
];

(async () => {
  for (const [name, fn] of steps) {
    let r = null;
    try { r = await fn(); } catch (e) { console.log(`${name}: THROW ${e.message}`); continue; }
    const state = await alive();
    const detail = r && r.success === false ? `失败(${txt(r).slice(0, 60)})` : 'ok';
    console.log(`${name.padEnd(28)} ${detail.padEnd(70)} 文稿=${state}`);
    if (state.startsWith('DEAD')) { console.log('  ↑ 就是这一条把文稿弄没了，后面的失败都是它的连锁'); break; }
  }
  try { await call('wps_ppt_close_presentation', { save: false }); } catch { /* ignore */ }
  process.exit(0);
})();
