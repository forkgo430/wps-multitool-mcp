/**
 * B5-T3 主线 · 第一批：4 个新工具真机证据（走 callTool 真入口，逐条原始输出落文件）。
 * wps_word_review_mode / wps_word_undo / wps_excel_create_table / wps_excel_add_slicer
 * 纪律：临时文稿/工作簿只关自己建的；锁由调用方（会话层）持有。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-live4-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String(r && r.content && r.content[0] && r.content[0].text || '').replace(/\s+/g, ' ');
const EVID = {};
const rec = (k, r) => { EVID[k] = { ok: !!r.success, text: txt(r).slice(0, 300) }; console.log(`[${k}] ok=${r.success} → ${txt(r).slice(0, 150)}`); };

(async () => {
  // ---- 前置：建 Word 文档（review_mode 依赖活动文档）----
  rec('setup.wordDoc', await call('wps_word_create_document', {}));

  // ---- 1. wps_word_review_mode：开 → 查 → 关（三连）----
  rec('review_mode.on', await call('wps_word_review_mode', { enabled: true }));
  rec('review_mode.query', await call('wps_word_review_mode', {}));
  rec('review_mode.off', await call('wps_word_review_mode', { enabled: false }));

  // ---- 2. wps_word_undo：建文档 → 写入 → 撤销 1 步 → 读回验证 ----
  const cr = await call('wps_word_create_document', {});
  rec('undo.setup.create', cr);
  const ins = await call('wps_word_insert_text', { text: '撤销前的那句话' });
  rec('undo.setup.insert', ins);
  rec('undo.exec', await call('wps_word_undo', { steps: 1 }));
  const after = await call('wps_word_get_document_text', {});
  rec('undo.verify.textAfter', after);
  EVID['undo.verify.note'] = { ok: true, text: '撤销后文档文本不应再含「撤销前的那句话」——见 textAfter 原文' };
  // 清理：关掉自己建的文档（放弃保存）
  rec('undo.cleanup.close', await call('wps_word_close_document', { save: false }));

  // ---- 3. wps_excel_create_table：建簿 → 写表头+数据 → 建 Excel 表 ----
  rec('create_table.setup.wb', await call('wps_excel_create_workbook', {}));
  rec('create_table.setup.data', await call('wps_excel_write_range', { range: 'A1:B3', data: [['名称', '数量'], ['甲', 1], ['乙', 2]] }));
  const ct = await call('wps_excel_create_table', { range: 'A1:B3', tableName: '探针表' });
  rec('create_table.exec', ct);

  // ---- 4. wps_excel_add_slicer：对上面那张表按列加切片器 ----
  rec('add_slicer.exec', await call('wps_excel_add_slicer', { field: '名称', tableName: '探针表' }));

  const out = path.join(process.env.TEMP || '.', 'wps-round7', 'live4-evidence.json');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(EVID, null, 2));
  console.log('\n证据已写', out);
  process.exit(0);
})().catch((e) => {
  console.error('异常：', e && e.message ? e.message : e);
  const out = path.join(process.env.TEMP || '.', 'wps-round7', 'live4-evidence.json');
  fs.writeFileSync(out, JSON.stringify({ fatal: String(e && e.message ? e.message : e), partial: EVID }, null, 2));
  process.exit(2);
});
