/**
 * 建靶实测证据：跑一次 rebuild()，把页数 / 每页元素数 / 母版元素数 / 文本都打出来。
 * 统筹者用它确认"交付不是纸上的" —— 隔离验证跑的是不是一个真靶子，这里一眼能看出来。
 * 只关自己建的文稿（名字差集），不碰用户正在编辑的演示文稿。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-target-proof-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';

const fs = require('fs');
const os = require('os');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-target-'));
const png = path.join(tmpDir, 'one.png');
fs.writeFileSync(png, Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64'));

const call = (n, a = {}) => toolRegistry.callTool({ id: n, name: n, arguments: a });
const txt = (r) => String(((r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || ''));
const owned = new Set();
const names = async () => new Set([...txt(await call('wps_ppt_get_open_presentations', {}))
  .replace(/\s+/g, ' ').matchAll(/(\d+)\.\s*(\S+)\s*(?:\[当前活动\])?\s*路径:/g)].map((m) => m[2]));

(async () => {
  const before = await names();
  await call('wps_ppt_create_presentation', {});
  for (const n of await names()) if (!before.has(n)) owned.add(n);
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_slide', { layout: 'blank' });
  for (const p of [1, 2]) {
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'rectangle', left: 60, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'oval', left: 240, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_textbox', { slideIndex: p, left: 60, top: 160, width: 240, height: 60, text: `文字${p}` });
    await call('wps_ppt_insert_table', { slideIndex: p, rows: 2, cols: 2 });
    await call('wps_ppt_insert_image', { slideIndex: p, path: png });
    await call('wps_ppt_set_shape_text', { slideIndex: p, shapeIndex: 1, text: `文字${p}` });
  }

  console.log('===== 隔离验证器 rebuild() 的真机实测 =====');
  console.log('打开的演示文稿:', [...(await names())].join(' / '));
  console.log('页数:', txt(await call('wps_ppt_get_slide_count', {})).trim());
  for (const p of [1, 2]) {
    console.log(`\n--- 第 ${p} 页 ---`);
    console.log(txt(await call('wps_ppt_get_slide_info', { slideIndex: p })).trim());
    console.log('文本框读回:', txt(await call('wps_ppt_get_textboxes', { slideIndex: p })).trim().slice(0, 260));
    console.log('动画:', txt(await call('wps_ppt_get_animations', { slideIndex: p })).trim());
  }
  console.log('\n--- 表格单元格读回（第1页 1行1列）---');
  console.log(txt(await call('wps_ppt_get_table_cell', { slideIndex: 1, tableIndex: 1, row: 1, col: 1 })).trim());
  console.log('\n--- 母版 ---');
  console.log((txt(await call('wps_ppt_get_slide_master', {})).match(/母版元素: \d+ 个/) || [''])[0]);

  for (const n of await names()) {
    if (!owned.has(n)) continue;
    await call('wps_ppt_close_presentation', { name: n, save: false });
  }
  const left = [...(await names())].filter((n) => !before.has(n));
  console.log('\n收尾后本进程新建的文稿:', left.length ? left.join('/') : '(已全部关闭)');
  console.log('用户原有的文稿（未被触碰）:', [...before].join(' / ') || '(无)');
  fs.rmSync(tmpDir, { recursive: true, force: true });
  process.exit(0);
})();