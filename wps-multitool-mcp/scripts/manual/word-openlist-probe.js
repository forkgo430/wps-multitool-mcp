/**
 * A9 修复的机制验证（不依赖那场"窗口"会不会复现）：
 *   1) WPS 没在跑 → 取实例必然是"新建的" → 列表空 → 必须带 warning，且重试有界（不无限等）
 *   2) 有文档在跑 → 列表非空 → 立刻返回，不重试、不带 warning（不误伤正常路径）
 *
 * 这两条正是修复的两个分支：可疑才重试+警告，不可疑就当事实。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-a9-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

const path = require('path');
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

let pass = 0, fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

/** 把所有 wps/et/wpp 杀掉，制造"没有 WPS 在跑"的干净起点 */
function killAllWps() {
  const ps = `
$procs = @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue)
foreach ($p in $procs) { try { $p.Kill() } catch {} }
Start-Sleep -Milliseconds 800
Write-Output ('KILLED=' + @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue).Count)
`;
  const f = 'C:\\Users\\金东权\\AppData\\Local\\Temp\\opencode\\a9-kill.ps1';
  require('fs').writeFileSync(f, ps, 'utf8');
  try {
    return execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', f], { encoding: 'utf8', timeout: 60000 }).trim();
  } catch (e) { return 'ERR ' + String(e.message).slice(0, 60); }
}

(async () => {
  console.log('\n=== 1. 没有 WPS 在跑时：列表应为空 + 带 warning + 有界重试 ===');
  console.log('  清进程：', killAllWps());
  const t0 = Date.now();
  const r1 = await call('wps_word_get_open_documents', {});
  const ms = Date.now() - t0;
  console.log('  回包：', txt(r1).slice(0, 120));
  console.log('  耗时：', ms, 'ms');
  const w1 = (r1.warnings || []).map((w) => String(w.message || w)).join(' | ');
  console.log('  warnings：', w1 ? w1.slice(0, 160) : '(无)');
  ok('返回成功（不是报错）', r1.success, txt(r1).slice(0, 60));
  ok('带上了"别据此判断文档没了"的警告', /别据此判断文档没了/.test(w1), w1.slice(0, 60));
  ok('重试有界（不超过 ~4s）', ms < 6000, `${ms}ms`);

  console.log('\n=== 2. 有文档在跑时：列表非空 → 立刻返回，不重试、不带警告 ===');
  const created = await call('wps_word_create_document', {});
  console.log('  建文档：', txt(created).slice(0, 80));
  const t1 = Date.now();
  const r2 = await call('wps_word_get_open_documents', {});
  const ms2 = Date.now() - t1;
  console.log('  回包：', txt(r2).slice(0, 120));
  console.log('  耗时：', ms2, 'ms');
  const w2 = (r2.warnings || []).map((w) => String(w.message || w)).join(' | ');
  ok('列表里有文档', /文稿|文档/.test(txt(r2)), txt(r2).slice(0, 80));
  ok('正常路径【不带】那条警告（不误伤）', !/别据此判断文档没了/.test(w2), w2.slice(0, 60));
  ok('正常路径不重试（< 2s）', ms2 < 2000, `${ms2}ms`);

  console.log('\n=== 3. 收尾 ===');
  const names = ((await wpsClient.executeMethod('getOpenDocuments', {}, 'word')).data || {}).documents || [];
  for (const d of names) {
    await wpsClient.executeMethod('closeDocument', { name: d.name, saveChanges: false }, 'word');
  }
  console.log('  关掉我建的：', names.length, '个');
  console.log('  清进程：', killAllWps());

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('探针异常：', e && e.message ? e.message : e);
  process.exit(2);
});
