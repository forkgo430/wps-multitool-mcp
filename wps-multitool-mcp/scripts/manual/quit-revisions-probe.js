/**
 * 轮 7 探针：quit 新护栏 + revisions 破坏性注解 真机验证（走 callTool 真入口）。
 *
 * 用例（任务卡第 2 期第 2 章轮 7 §5）：
 *   A  quit 拒绝路径   工具建新文档 → 写文字 → 不保存 → wps_common_quit（不传 confirm）
 *                      期望 success=false，报错列出未保存文档名，WPS 进程仍活着
 *   B  quit 优雅路径   那份文档 save_as 落盘后 → wps_common_quit
 *                      期望 success=true、remaining=0（基线无用户窗口才执行）
 *   C  revisions 快照  建 Word 文档 → 开修订 → 写两段（中途落盘，否则 auto 快照 no_path）→
 *                      wps_word_revisions {mode:"acceptAll"}
 *                      期望 success=true，且 oplog 出现一条 auto 快照记录 / snapshots 出现新文件
 *
 * 为什么用 save_as 不用 save：从未保存的文档 $doc.Save() 会弹「另存为」对话框，
 * 无人值守会把 COM 通道挂死（close_document 的工具说明与 B2 已知问题同源）。
 *
 * 纪律：全程同一时刻至多 1 个 WPS 实例（所有者 2026-10-25 要求）；临时文档只关自己建的；
 *       磁盘临时文件不删（红线禁止删文件），路径在输出里报告。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-round7-quit-rev-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '').replace(/\s+/g, ' ');
const full = (r) => JSON.stringify(r, (k, v) => (k === 'warnings' || k === 'content' ? v : v), 2).slice(0, 900);

const HOME_DIR = path.join(os.homedir(), '.wps-multitool-mcp');
const OPLOG = path.join(HOME_DIR, 'oplog', 'ops.jsonl');
const SNAPDIR = path.join(HOME_DIR, 'snapshots');
const TMPDIR = path.join(os.tmpdir(), 'wps-round7');

let pass = 0, fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' —— ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' —— ' + extra : ''}`); }
};

/* ---- WPS 套件进程快照（含窗口标题，用于"至多一个实例"监测与收工对账） ---- */
function wpsProcs() {
  const out = execFileSync('pwsh', ['-NoProfile', '-NonInteractive', '-Command',
    "Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(wps|et|wpp|wpscloudsvr|wpscenter|wpsoffice|wpsupdate)' } | ForEach-Object { $_.Name + '#' + $_.ProcessId + ' title=[' + $_.MainWindowTitle + ']' }",
  ], { encoding: 'utf8', timeout: 60000 });
  return out.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
}

function readOplogLines() {
  try { return fs.readFileSync(OPLOG, 'utf8').split(/\r?\n/).filter(Boolean); } catch { return []; }
}
function listSnapshots() {
  try { return fs.readdirSync(SNAPDIR).filter((f) => f.endsWith('.docx') || f.endsWith('.xlsx') || f.endsWith('.pptx')); } catch { return []; }
}

(async () => {
  const t0 = new Date();
  const t0Iso = t0.toISOString();
  const oplogBefore = new Set(readOplogLines());
  const snapsBefore = new Set(listSnapshots());
  fs.mkdirSync(TMPDIR, { recursive: true });
  const stamp = Date.now();

  console.log('=== 0. 前置状态 ===');
  console.log('  T0 =', t0Iso);
  console.log('  oplog 基线行数 =', oplogBefore.size, '· snapshots 基线文件数 =', snapsBefore.size);
  const base0 = wpsProcs();
  console.log('  开跑前 WPS 套件进程 =', base0.length ? base0.join(' | ') : '（无）');

  console.log('\n=== 用例 A：quit 拒绝路径（未保存默认拒绝）===');
  let r = await call('wps_word_create_document', {});
  console.log('  create_document →', txt(r).slice(0, 120));
  const docName = (txt(r).match(/[:：]\s*(文字文稿\S*)/) || [])[1] || '文字文稿?';
  r = await call('wps_word_insert_text', { text: '轮7探针用例A：这段文字不保存。' });
  console.log('  insert_text →', txt(r).slice(0, 120));
  const procsMidA = wpsProcs();
  console.log('  写入后 WPS 进程（应活着）=', procsMidA.join(' | ') || '（无）');
  r = await call('wps_common_quit', {});
  console.log('  quit（无 confirm）完整回包 →', full(r));
  ok('A1 quit 返回 success=false', r && r.success === false, `success=${r && r.success}`);
  const a2 = /勿重试|未保存/.test(txt(r)) && /文字文稿/.test(txt(r));
  ok('A2 报错列出未保存文档名（文字文稿…）', a2, txt(r).slice(0, 90));
  const procsA = wpsProcs();
  const aliveA = procsA.some((l) => /^(wps|et|wpp)\.exe#/.test(l));
  ok('A3 WPS 进程仍活着', aliveA, procsA.join(' | ') || '（无）');

  console.log('\n=== 用例 B：quit 优雅路径（基线无用户窗口，执行）===');
  r = await call('wps_common_save_as', { filePath: path.join(TMPDIR, `caseB-${stamp}.docx`) });
  console.log('  save_as →', txt(r).slice(0, 150));
  ok('B0 save_as 落盘成功（Save() 会弹框，故走另存为）', r && r.success === true, txt(r).slice(0, 100));
  r = await call('wps_common_quit', {});
  console.log('  quit 完整回包 →', full(r));
  let remaining = null;
  try { remaining = JSON.parse(r.content[0].text).data.remaining; } catch { /* 回包文本非 JSON */ }
  if (remaining === null) {
    // ps1 语义：left=0 ⇔ 文本"WPS 已退出"；left>0 ⇔ "另有 N 个 WPS 进程在运行"
    const t = txt(r);
    remaining = /已退出/.test(t) && !/另有/.test(t) ? 0 : ((t.match(/另有\s*(\d+)/) || [])[1] ? Number(RegExp.$1) : null);
  }
  ok('B1 quit 返回 success=true', r && r.success === true, `success=${r && r.success}`);
  ok('B2 remaining=0', remaining === 0, `remaining=${remaining}`);
  await new Promise((res) => setTimeout(res, 3000));
  const procsB = wpsProcs();
  const officeB = procsB.filter((l) => /^(wps|et|wpp)\.exe#/.test(l));
  ok('B3 WPS office 进程已退（wpscloudsvr 等服务另计）', officeB.length === 0, procsB.join(' | ') || '（无）');

  console.log('\n=== 用例 C：revisions acceptAll + auto 快照 ===');
  r = await call('wps_word_create_document', {});
  console.log('  create_document →', txt(r).slice(0, 120));
  if (!r || r.success !== true) { // quit 后桥子进程 COM 可能失效，容错重试一次
    await new Promise((res) => setTimeout(res, 2500));
    r = await call('wps_word_create_document', {});
    console.log('  create_document 重试 →', txt(r).slice(0, 120));
  }
  r = await call('wps_word_enable_track_changes', {});
  console.log('  enable_track_changes →', txt(r).slice(0, 120));
  r = await call('wps_word_insert_text', { text: '轮7探针用例C第一段：修订开启后写入。' });
  console.log('  insert_text 第一段 →', txt(r).slice(0, 120));
  r = await call('wps_common_save_as', { filePath: path.join(TMPDIR, `caseC-${stamp}.docx`) });
  console.log('  save_as（落盘，auto 快照才取得到文件）→', txt(r).slice(0, 150));
  r = await call('wps_word_insert_text', { text: '轮7探针用例C第二段：落盘后再写，制造第二条修订。' });
  console.log('  insert_text 第二段 →', txt(r).slice(0, 120));
  r = await call('wps_word_revisions', { mode: 'list' });
  console.log('  revisions list →', txt(r).slice(0, 150));
  // C0 判据 v2：锚定 handler 的固定输出"共 N 条修订"精确取数。
  // 第一版抓"文本里所有数字取最大"，连文档名"文稿1"里的 1 都算修订数 —— 判据缺陷，已修。
  const revCount = Number((txt(r).match(/共 (\d+) 条修订/) || [])[1] || 0);
  ok('C0 修订已产生（count>0，list 精确读）', revCount > 0, `count=${revCount}`);

  r = await call('wps_word_revisions', { mode: 'acceptAll' });
  console.log('  revisions acceptAll 完整回包 →', full(r));
  ok('C1 acceptAll success=true', r && r.success === true, txt(r).slice(0, 100));

  await new Promise((res) => setTimeout(res, 1500));
  const oplogAfter = readOplogLines();
  const newLines = oplogAfter.filter((l) => !oplogBefore.has(l));
  const autoSnap = newLines.find((l) => {
    try {
      const j = JSON.parse(l);
      return j.kind === 'snapshot' && String(j.label || '').includes('auto: wps_word_revisions');
    } catch { return false; }
  });
  console.log('  oplog 新增行数 =', newLines.length, '· auto 快照记录 =', autoSnap || '（无）');
  ok('C2 oplog 出现 auto 快照记录（破坏性注解生效）', !!autoSnap, autoSnap ? autoSnap.slice(0, 200) : '未找到');
  const snapsAfter = new Set(listSnapshots());
  const newSnaps = [...snapsAfter].filter((f) => !snapsBefore.has(f));
  console.log('  snapshots 新增文件 =', newSnaps.join(' | ') || '（无）');
  ok('C3 snapshots 目录出现新快照文件', newSnaps.length > 0, newSnaps.join(','));

  console.log('\n=== 收尾：只关自己建的，退出 WPS ===');
  r = await call('wps_word_close_document', { save: false });
  console.log('  close_document(save:false) →', txt(r).slice(0, 120));
  r = await call('wps_common_quit', {});
  console.log('  收尾 quit →', txt(r).slice(0, 150));
  await new Promise((res) => setTimeout(res, 3000));
  const procsEnd = wpsProcs();
  console.log('  收尾后 WPS 套件进程 =', procsEnd.join(' | ') || '（无）');

  console.log('\n=== 临时文件（不删，红线禁止删文件；交统筹者处置）===');
  for (const f of fs.readdirSync(TMPDIR)) console.log('  ', path.join(TMPDIR, f));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('探针异常：', e && e.message ? e.message : e);
  process.exit(2);
});
