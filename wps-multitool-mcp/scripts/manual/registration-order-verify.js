/**
 * Pos: 注册顺序对账门禁（代码技法卡 D 段，卫生轮附带发现③的看守落地）。
 * 一旦我被修改，请更新我的头部注释，以及所属目录的md。
 * 背景：shape-basic/shape-style 拆分后 ppt/index.ts 靠 slice 组合保序——无人看守；
 * 本门禁把"拆分块内工具的相对注册顺序"钉进快照，改组内顺序/增删工具即红。
 * 全 mock：直接 require dist 注册表，零 COM。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const results = [];
const check = (name, cond, info = '') => {
  results.push({ name, ok: !!cond });
  console.log(`  ${cond ? '✓' : '✗'} ${name}${info ? ` — ${info}` : ''}`);
};

new WpsMcpServer();
const flat = toolRegistry.listAllTools().map((t) => t.name);

// ---- 1. ppt 形状 16 件的相对顺序（拆分前 shapeBasicTools 的原始序，卫生轮逐一保持）
const SHAPE_ORDER = [
  'wps_ppt_add_shape', 'wps_ppt_add_textbox', 'wps_ppt_delete_shape', 'wps_ppt_get_shapes',
  'wps_ppt_set_shape_position', 'wps_ppt_set_shape_style', 'wps_ppt_set_shape_shadow',
  'wps_ppt_set_shape_gradient', 'wps_ppt_set_shape_border', 'wps_ppt_set_shape_transparency',
  'wps_ppt_set_shape_text', 'wps_ppt_set_shape_fill', 'wps_ppt_set_slide_title',
  'wps_ppt_align_shapes', 'wps_ppt_distribute_shapes', 'wps_ppt_group_shapes',
];
const idx = SHAPE_ORDER.map((n) => flat.indexOf(n));
check('ppt 形状 16 件全部已注册', idx.every((i) => i >= 0), idx.filter((i) => i < 0).map((i, k) => SHAPE_ORDER[k]).join(','));
const relOk = idx.slice(1).map((v, k) => v - idx[k]).every((d) => d > 0);
check('ppt 形状 16 件相对注册顺序 = 拆分前原序', relOk, relOk ? '' : `实际序: ${SHAPE_ORDER.map((n) => flat.indexOf(n)).join(',')}`);

// ---- 2. excel 工作表 16 件的相对顺序（sheet.ts + sheet-advanced.ts 拆分原序）
const SHEET_ORDER = [
  'wps_excel_create_sheet', 'wps_excel_delete_sheet', 'wps_excel_rename_sheet', 'wps_excel_copy_sheet',
  'wps_excel_get_sheet_list', 'wps_excel_switch_sheet', 'wps_excel_move_sheet', 'wps_excel_get_selection',
  'wps_excel_delete_row', 'wps_excel_freeze_panes', 'wps_excel_auto_fill', 'wps_excel_set_named_range',
  'wps_excel_hide_column', 'wps_excel_auto_sum', 'wps_excel_create_table', 'wps_excel_add_slicer',
];
const eidx = SHEET_ORDER.map((n) => flat.indexOf(n));
check('excel 工作表 16 件全部已注册', eidx.every((i) => i >= 0), eidx.filter((i) => i < 0).map((i, k) => SHEET_ORDER[k]).join(','));
const erelOk = eidx.slice(1).map((v, k) => v - eidx[k]).every((d) => d > 0);
check('excel 工作表 16 件相对注册顺序 = 拆分前原序', erelOk, erelOk ? '' : `实际序: ${SHEET_ORDER.map((n) => flat.indexOf(n)).join(',')}`);

// ---- 3. 总量与 professional 集合完整性（拆分不许丢工具）
// 287→284：3D 四件套 op-merge（host wps 优先卡 E 段数字连锁；卡面明定）
  check('注册总量 284（3D op-merge 后）', flat.length === 284, String(flat.length));
check('professional 工具集（allTools）全部在注册表', allTools.every((t) => flat.includes(t.definition.name)), '');

const passed = results.filter((x) => x.ok).length;
console.log(`汇总：${passed}/${results.length} 通过`);
if (passed !== results.length) process.exit(1);
