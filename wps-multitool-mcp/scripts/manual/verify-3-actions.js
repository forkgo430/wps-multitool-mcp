// 第 6 轮 验证者 任务 3：单独复验实现者补完的三个动作
//   wps_ppt_set_font_color / wps_ppt_set_slide_size / wps_ppt_set_slide_theme
// 用干净文稿（建靶→加页→加形状/文本框），逐个调用，确认不再 Unknown action 且真生效。
// 隔离：独立时间戳管道 + EPHEMERAL；只关自己建的。

const fs = require('fs');
const os = require('os');
const path = require('path');

const PIPE_SUFFIX = `agentC-3act-${Date.now()}`;
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-${PIPE_SUFFIX}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '120000';
process.env.WPS_DAEMON_TIMEOUT = '30000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => (r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '';

const results = [];
const check = (name, pass, detail) => {
  results.push({ name, pass, detail });
  console.log(`${pass ? '  ✓' : '  ✗'} ${name} — ${detail}`);
};

// 读回某页某形状的颜色（getShapes 的 JSON 里若有 fontColor 就用它）
async function readShapeColor(slideIndex, shapeIndex) {
  const r = await call('wps_ppt_get_shapes', { slideIndex });
  const m = txt(r).match(/"fontColor"\s*:\s*"([^"]*)"/);
  return m ? m[1] : null;
}
async function readSlideSize() {
  const r = await call('wps_execute_method', { method: 'getSlideInfo', appType: 'presentation', params: { slideIndex: 1 } });
  const w = txt(r).match(/"width"\s*:\s*(\d+)/);
  const h = txt(r).match(/"height"\s*:\s*(\d+)/);
  return { w: w ? +w[1] : null, h: h ? +h[1] : null };
}

(async () => {
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-3act-'));
  const mine = [];
  try {
    // 建干净文稿：新建是 0 页
    const cr = await call('wps_ppt_create_presentation', {});
    if (!cr.success) { check('建文稿', false, txt(cr).slice(0, 60)); return; }
    const pres = (txt(cr).match(/"name":"([^"]*)"/) || [])[1];
    if (pres) mine.push(pres);
    await call('wps_ppt_add_slide', { layout: 'blank' });   // 第 1 页
    await call('wps_ppt_add_slide', { layout: 'blank' });   // 第 2 页（确认多页不崩）
    // 第 1 页加一个带文字的形状（set_font_color 要改它的字色）
    const sh = await call('wps_ppt_add_shape', { slideIndex: 1, shapeType: 'rectangle', text: 'hello' });
    console.log(`建靶: pres=${pres} add_shape=${sh.success}`);

    // ---- 1. set_font_color ----
    const c0 = await readShapeColor(1, 1);
    const fc = await call('wps_ppt_set_font_color', { slideIndex: 1, shapeIndex: 1, color: '#FF0000' });
    const c1 = await readShapeColor(1, 1);
    check(
      'set_font_color 不再 Unknown action 且返回成功',
      fc.success === true && !/Unknown action/.test(txt(fc)),
      `success=${fc.success}｜${txt(fc).split('\n')[0].slice(0, 50)}`
    );
    check('set_font_color 真改了字色', c1 && c1 !== c0, `改前=${c0} 改后=${c1}`);

    // ---- 2. set_slide_size ----
    const s0 = await readSlideSize();
    const ss = await call('wps_ppt_set_slide_size', { width: 960, height: 540 });
    const s1 = await readSlideSize();
    check(
      'set_slide_size 不再 Unknown action 且返回成功',
      ss.success === true && !/Unknown action/.test(txt(ss)),
      `success=${ss.success}｜${txt(ss).split('\n')[0].slice(0, 50)}`
    );
    check('set_slide_size 真改了尺寸', s1.w === 960 && s1.h === 540, `改前=${s0.w}x${s0.h} 改后=${s1.w}x${s1.h}`);

    // ---- 3. set_slide_theme ----
    const st = await call('wps_ppt_set_slide_theme', { theme: 'Office' });
    check(
      'set_slide_theme 不再 Unknown action 且返回成功',
      st.success === true && !/Unknown action/.test(txt(st)) && !/Null 值表达式/.test(txt(st)),
      `success=${st.success}｜${txt(st).split('\n')[0].slice(0, 50)}`
    );
  } catch (e) {
    check('运行异常', false, e && e.message ? e.message : String(e));
  } finally {
    for (const n of mine) {
      try { await call('wps_ppt_close_presentation', { name: n, save: false }); } catch { /* 已关 */ }
    }
    try { fs.rmSync(tmpRoot, { recursive: true, force: true }); } catch { /* 占用 */ }
  }
  const pass = results.filter((r) => r.pass).length;
  console.log(`\n3 动作隔离复验：${pass}/${results.length} 通过`);
  process.exit(0);
})();
