/**
 * 双装机绑定一致性 · 复现钉板 + 双场景验收（Office 收口合并卡 A 段）
 *
 * 复现（病灶）：多实例场（≥3 WINWORD）上 create→next-call 漂移 —— getter 取序
 *   "①ROT → ②缓存 → ③新建" 的 ROT 多条目抽签让同会话两次调用可能绑不同实例，
 *   症状 = createDocument 后下一次 getOpenDocuments 附着到另一个 0 文档实例（列表空）。
 *
 * 用法：
 *   node scripts/manual/word-binding-repro.js [--instances N] [--rounds N] [--keep]
 *   --instances 多实例场实例数（缺省 3）；--rounds 复现轮数（缺省 8）；--keep 保留造场实例（默认结束清场）
 *
 * 判据：任一轮 getOpenDocuments 返回列表不含刚建文档（或空）= 漂移发生。
 * 这不是 FAST：需要真机 + 真 Office。判定归决策者。
 */
process.env.WPS_OFFICE = process.env.WPS_OFFICE || 'on';
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_PIPE = process.env.WPS_BRIDGE_PIPE || `\\\\.\\pipe\\wps-multitool-mcp-bindrepro-${Date.now()}`;

const path = require('path');
const { execFileSync } = require('child_process');
const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

const args = process.argv.slice(2);
const argOf = (name, dflt) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? parseInt(args[i + 1], 10) : dflt;
};
const INSTANCES = argOf('--instances', 3);
const ROUNDS = argOf('--rounds', 8);
const KEEP = args.includes('--keep');

/** 用独立 powershell（STA 5.1）CoCreate 一个可见 WINWORD；PID 由调用方用进程差集取 */
function spawnWordInstance(i) {
  const script = `
    $w = New-Object -ComObject 'Word.Application'
    $w.Visible = $true
    $d = $w.Documents.Add()
    $d.Content.Text = "BINDREPRO-INSTANCE-${i}"
    Start-Sleep -Milliseconds 600
  `;
  const f = path.join(process.env.TEMP || 'C:/Temp', `bindrepro-spawn-${i}.ps1`);
  require('fs').writeFileSync(f, script, 'utf8');
  try {
    execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', f], { timeout: 60000, stdio: 'ignore' });
    return true;
  } catch (e) {
    console.log(`  ⚠ 造场实例 ${i} 失败: ${(e.message || '').slice(0, 100)}`);
    return false;
  }
}

/** 杀指定 PID（清场） */
function killPid(pid) {
  try { execFileSync('taskkill', ['/F', '/PID', String(pid), '/T'], { timeout: 15000, stdio: 'ignore' }); } catch { /* 已不在 */ }
}

/** 当前 WINWORD 进程列表 */
function wordPids() {
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-Command',
      "Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }"], { encoding: 'utf8', timeout: 20000 });
    return out.split(/\s+/).filter(Boolean).map(Number);
  } catch { return []; }
}

(async () => {
  console.log('=== 双装机绑定一致性 · 复现钉板 ===');
  console.log(`场景：多实例场（${INSTANCES} WINWORD）× ${ROUNDS} 轮 create→getOpenDocuments`);

  // 造场：先记录现有 WINWORD（那些不是本脚本造的，清场时不动）
  const preExisting = new Set(wordPids());
  console.log(`场内已有 WINWORD（非本脚本造）: ${[...preExisting].join(',') || '无'}`);
  const spawned = [];
  for (let i = 1; i <= INSTANCES; i++) {
    const before = new Set(wordPids());
    if (!spawnWordInstance(i)) continue;
    await new Promise((r) => setTimeout(r, 800));
    const newPids = wordPids().filter((p) => !before.has(p) && !spawned.includes(p));
    if (newPids.length) {
      spawned.push(...newPids);
      console.log(`  ✓ 造场实例 ${i}: pid=${newPids.join(',')}`);
    } else {
      console.log(`  ⚠ 造场实例 ${i}: CoCreate 成功但未观测到新 WINWORD 进程（可能复用既有实例——多实例场不成立）`);
    }
  }
  if (spawned.length < 2) {
    console.log('✗ 造场失败（<2 实例），钉板无法进行 —— 如实退出');
    for (const p of spawned) killPid(p);
    process.exit(2);
  }
  await new Promise((r) => setTimeout(r, 1500)); // 等 ROT 注册

  // 复现：create → next-call 读列表 ×ROUNDS
  let driftRounds = 0;
  const evidence = [];
  for (let round = 1; round <= ROUNDS; round++) {
    const cr = await wpsClient.executeMethod('createDocument', {}, 'word');
    if (!cr || cr.success !== true) {
      console.log(`  轮 ${round}: createDocument 失败: ${(cr && cr.error || '').slice(0, 80)}`);
      evidence.push({ round, create: 'fail' });
      driftRounds++;
      continue;
    }
    const docName = cr.data && (cr.data.name || '');
    const g = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
    const docs = g && g.success && g.data && (g.data.documents || g.data.list || []);
    const names = Array.isArray(docs) ? docs.map((d) => (typeof d === 'string' ? d : (d.name || ''))) : [];
    const found = names.some((n) => n && docName && n.indexOf(docName) !== -1);
    const status = found ? '绑定一致' : (names.length === 0 ? '漂移（附着 0 文档实例）' : '漂移（列表缺刚建文档）');
    if (!found) driftRounds++;
    console.log(`  轮 ${round}: create→"${docName}" → getOpenDocuments ${names.length} 份 [${names.slice(0, 3).join(',')}] → ${status}`);
    evidence.push({ round, docName, listCount: names.length, found });
    // 关掉本轮建的文档，避免列表膨胀掩盖漂移信号
    await wpsClient.executeMethod('closeDocument', { saveChanges: false }, 'word').catch(() => {});
  }

  console.log(`\n结果：${ROUNDS} 轮中漂移 ${driftRounds} 轮 —— ${driftRounds > 0 ? '病灶复现 ✓（钉板成立）' : '未复现（如实记录；按证据修——修复 A 的缓存优先在单实例场也无害）'}`);
  console.log(JSON.stringify(evidence));

  // 清场：只收本脚本造的实例
  if (!KEEP) {
    for (const p of spawned) killPid(p);
    console.log(`清场：已收本脚本造的 ${spawned.length} 个实例（PID 白名单，不动场内其他实例）`);
  } else {
    console.log(`--keep：保留 ${spawned.length} 个造场实例（pid: ${spawned.join(',')})`);
  }
  process.exit(0);
})().catch((e) => {
  console.error('钉板异常：', e && e.stack ? e.stack : e);
  process.exit(2);
});
