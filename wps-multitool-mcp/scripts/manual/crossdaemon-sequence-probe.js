/**
 * Pos: 多文档顺序主锚探针（守护拓扑合并卡 · 主锚）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 病灶（本卡依据）：旧拓扑下 word 守护与 common 守护各持一个 wps.exe 岛，跨进程文档可见性
 * 靠 ROT 附着抽签——**多文档顺序场景（真实 AI 会话常态：建→写→存→再建……）恒在抽签**。
 * 合并后 common 动作落 word 通道（同进程同会话同缓存），此类抽签整类消失。
 *
 * 断言面（5 轮顺序场景，文档累积不关闭——最大抽签压力）：
 *   每轮：word create → word insert（轮次标记）→ common save_as（唯一路径）→ 文件落盘
 *         → word read 回读标记（同会话三对账）
 *   收尾：逐份关闭（未保存丢弃）+ 临时文件清理
 * 全 real-machine（需 WPS 可达）。不设 WPS_OFFICE——缺省态（auto → 双链 → wps 优先）。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-xseq-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '180000';
delete process.env.WPS_OFFICE;

const path = require('path');
const fs = require('fs');
const os = require('os');
const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);
const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '');

let pass = 0, fail = 0;
const ok = (label, cond, extra = '') => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

const ROUNDS = 5;

(async () => {
  console.log('== 多文档顺序主锚（跨守护抽签回归）==');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xseq-'));
  const paths = [];
  let allOk = true;

  for (let i = 1; i <= ROUNDS; i++) {
    const marker = `顺序轮次${i}标记`;
    const cr = await call('wps_word_create_document', {});
    if (cr.success === false) { ok(`轮${i} create`, false, txt(cr).slice(0, 60)); allOk = false; break; }
    await call('wps_word_insert_text', { text: marker, position: 'end' }).catch(() => {});

    const p = path.join(tmpDir, `seq-${i}.docx`);
    const sv = await call('wps_common_save_as', { filePath: p });
    const saveOk = sv.success === true && fs.existsSync(p);
    ok(`轮${i} common save_as（旧拓扑抽签点）`, saveOk, txt(sv).slice(0, 50));
    if (!saveOk) { allOk = false; break; }
    paths.push(p);

    const rd = await call('wps_word_get_document_text', {});
    ok(`轮${i} word read 回读标记（同会话三对账）`, rd.success !== false && txt(rd).includes(marker), '');
  }

  ok(`主锚：${ROUNDS}/${ROUNDS} 轮全过（多文档递增顺序场景）`, allOk && paths.length === ROUNDS, '');

  // 收尾：逐份关闭（未保存丢弃）+ 临时文件清理
  for (let i = 0; i < ROUNDS; i++) {
    await call('wps_word_close_document', { save: false }).catch(() => {});
  }
  for (const p of paths) { try { fs.unlinkSync(p); } catch { /* 已清 */ } }
  try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { /* 已清 */ }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch(async (e) => {
  console.error('PROBE FAIL:', e.message || e);
  for (let i = 0; i < 6; i++) { try { await call('wps_word_close_document', { save: false }); } catch { break; } }
  process.exit(1);
});
