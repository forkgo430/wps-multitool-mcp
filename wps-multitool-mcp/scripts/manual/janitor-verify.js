/**
 * Janitor 看门狗 · FAST 门禁（第 5 期主线甲 D 项）
 *
 * 判据（决策函数注入 + HOME 重定向，零 WPS COM、零真进程被杀）：
 *   ① journal 落盘可见（Write-SpawnJournal 行格式 {pid,exe,startTime,spawnedBy,ts}）
 *   ② headless + 全命中 → kill 计划（sweep 的靶）
 *   ③ UI 实例（有窗口标题）→ skip:ui-window（最高红线：零误伤）
 *   ④ StartTime 不符 → skip:pid-reused（PID 复用免疫）
 *   ⑤ 死 PID → dead 销账（journal 卫生）
 *   ⑥ 非 host 家族 / 非 ps1- 前缀 → skip:not-ours
 *   ⑦ 滚动上限：>10MB journal 截断留尾
 *   ⑧ 执行壳真跑（HOME 重定向的假 journal，pid 全不存在 → 只销账不杀）
 *
 * 跑法：node scripts/manual/janitor-verify.js（verify-all FAST 档包含它）。
 * PowerShell 段全部用 pwsh7（用户令：多用 pwsh7）。
 */
const { spawnSync, execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const ROOT = path.resolve(__dirname, '..', '..');
const PWSH = 'pwsh';   // 用户令：多用 pwsh7
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

/** pwsh7 跑一段 ps1（HOME 重定向到临时目录），回 stdout */
function pwsh7(script, homeDir) {
  const env = { ...process.env, HOME: homeDir, USERPROFILE: homeDir };
  const r = spawnSync(PWSH, ['-NoProfile', '-NonInteractive', '-Command', script], {
    encoding: 'utf8', timeout: 120000, env, cwd: ROOT,
  });
  return { out: (r.stdout || '') + (r.stderr || ''), code: r.status };
}

/** dot-source 前置（daemon 模式只定义不执行） */
const PRELUDE = (home) => `
$ErrorActionPreference = 'Continue'
$env:WPS_COM_DAEMON = '1'
. '${path.join(ROOT, 'scripts', 'wps-com.ps1')}' -Instance user -App ''
`;

(async () => {
  console.log('=== Janitor 看门狗 FAST 门禁 ===');
  const tmpHome = fs.mkdtempSync(path.join(os.tmpdir(), 'janitor-home-'));
  const journal = path.join(tmpHome, '.wps-multitool-mcp', 'spawn-journal.jsonl');

  // ① journal 落盘可见（Write-SpawnJournal 真写——用真实存在的进程 pid=自身 pwsh）
  const r1 = pwsh7(PRELUDE(tmpHome) + `
    $pids = @($PID)
    $written = Write-SpawnJournal $pids 'ps1-user'
    Write-Output ("WRITTEN=" + $written)
    Write-Output ("JOURNAL=" + (Get-JournalPath))
    if (Test-Path (Get-JournalPath)) {
      $line = Get-Content (Get-JournalPath) | Select-Object -First 1
      Write-Output ("LINE=" + $line)
      $j = $line | ConvertFrom-Json
      Write-Output ("FIELDS=" + (($j.pid -gt 0) -and $j.exe -and $j.startTime -and $j.spawnedBy -eq 'ps1-user' -and $j.ts))
    }
  `, tmpHome);
  ok('① journal 落盘可见（写入 1 行）', r1.out.includes('WRITTEN=1'), '');
  ok('① journal 行格式 {pid,exe,startTime,spawnedBy,ts} 齐全', r1.out.includes('FIELDS=True'), '');
  // 路径分隔符归一化后比对（2026-11-21 CI 实测：Linux pwsh 下 ps1 产正斜杠路径，写死反斜杠必红）
  const normPath = (x) => x.replace(/\r/g, '').replace(/\\/g, '/');
  ok('① journal 路径在重定向 HOME 下（隔离测试）', normPath(r1.out).includes(normPath(journal)), journal);

  // ②③④⑥ 决策函数注入（fake LiveProcs，零真进程）
  const iso = (d) => d.toISOString();
  const now = new Date();
  const fakeEntries = [
    { pid: 101, exe: 'WINWORD', startTime: iso(now), spawnedBy: 'ps1-user', ts: iso(now) },      // headless 全命中
    { pid: 102, exe: 'WINWORD', startTime: iso(now), spawnedBy: 'ps1-user', ts: iso(now) },      // UI 实例
    { pid: 103, exe: 'WINWORD', startTime: iso(new Date(now - 86400000)), spawnedBy: 'ps1-user', ts: iso(now) }, // StartTime 不符
    { pid: 104, exe: 'chrome', startTime: iso(now), spawnedBy: 'ps1-user', ts: iso(now) },       // 非 host 家族
    { pid: 105, exe: 'EXCEL', startTime: iso(now), spawnedBy: 'coze', ts: iso(now) },            // 非 ps1- 前缀
    { pid: 106, exe: 'EXCEL', startTime: iso(now), spawnedBy: 'ps1-sandbox', ts: iso(now) },     // headless sandbox（也收）
  ];
  const fakeProcs = [
    { id: 101, name: 'WINWORD', startTime: iso(now), mainWindowTitle: '' },                      // headless
    { id: 102, name: 'WINWORD', startTime: iso(now), mainWindowTitle: '报告.docx - Microsoft Word' }, // UI
    { id: 103, name: 'WINWORD', startTime: iso(now), mainWindowTitle: '' },                      // 复用者（StartTime 不符）
    { id: 104, name: 'chrome', startTime: iso(now), mainWindowTitle: '' },
    { id: 105, name: 'EXCEL', startTime: iso(now), mainWindowTitle: '' },
    { id: 106, name: 'EXCEL', startTime: iso(now), mainWindowTitle: '' },
  ];
  const r2 = pwsh7(PRELUDE(tmpHome) + `
    $entries = @(
      @{ pid = 101; exe = 'WINWORD'; startTime = '${iso(now)}'; spawnedBy = 'ps1-user'; ts = '${iso(now)}' },
      @{ pid = 102; exe = 'WINWORD'; startTime = '${iso(now)}'; spawnedBy = 'ps1-user'; ts = '${iso(now)}' },
      @{ pid = 103; exe = 'WINWORD'; startTime = '${iso(new Date(now - 86400000))}'; spawnedBy = 'ps1-user'; ts = '${iso(now)}' },
      @{ pid = 104; exe = 'chrome'; startTime = '${iso(now)}'; spawnedBy = 'ps1-user'; ts = '${iso(now)}' },
      @{ pid = 105; exe = 'EXCEL'; startTime = '${iso(now)}'; spawnedBy = 'coze'; ts = '${iso(now)}' },
      @{ pid = 106; exe = 'EXCEL'; startTime = '${iso(now)}'; spawnedBy = 'ps1-sandbox'; ts = '${iso(now)}' }
    )
    $procs = @(
      @{ id = 101; name = 'WINWORD'; startTime = '${iso(now)}'; mainWindowTitle = '' },
      @{ id = 102; name = 'WINWORD'; startTime = '${iso(now)}'; mainWindowTitle = '报告.docx - Microsoft Word' },
      @{ id = 103; name = 'WINWORD'; startTime = '${iso(now)}'; mainWindowTitle = '' },
      @{ id = 104; name = 'chrome'; startTime = '${iso(now)}'; mainWindowTitle = '' },
      @{ id = 105; name = 'EXCEL'; startTime = '${iso(now)}'; mainWindowTitle = '' },
      @{ id = 106; name = 'EXCEL'; startTime = '${iso(now)}'; mainWindowTitle = '' }
    )
    $plan = Get-JournalSweepPlan $entries $procs
    $killPids = @($plan.kill | ForEach-Object { $_.pid }) -join ','
    $skipMap = @{}
    foreach ($s in @($plan.skip)) { $skipMap[[int]$s.pid] = [string]$s.reason }
    Write-Output ("KILL=" + $killPids)
    Write-Output ("SKIP102=" + $skipMap[102])
    Write-Output ("SKIP103=" + $skipMap[103])
    Write-Output ("SKIP104=" + $skipMap[104])
    Write-Output ("SKIP105=" + $skipMap[105])
    Write-Output ("K106=" + ($killPids -contains 106))
  `, tmpHome);
  ok('② headless 全命中 → kill 计划', r2.out.includes('KILL=101,106'), '');
  ok('③ UI 实例（有窗口标题）→ skip:ui-window（零误伤红线）', r2.out.includes('SKIP102=ui-window'), '');
  ok('④ StartTime 不符 → skip:pid-reused（PID 复用免疫）', r2.out.includes('SKIP103=pid-reused'), '');
  ok('⑥ 非 host 家族 → skip:not-ours', r2.out.includes('SKIP104=not-ours'), '');
  ok('⑥ 非 ps1- 前缀 → skip:not-ours', r2.out.includes('SKIP105=not-ours'), '');

  // ⑤⑧ 执行壳真跑（HOME 重定向的假 journal：pid 全不存在 → dead 销账，不杀任何真进程）
  const r5 = pwsh7(PRELUDE(tmpHome) + `
    $pids = @($PID); $null = Write-SpawnJournal $pids 'ps1-user'
    # 造两个必然不存在的 pid 行（销账对象）
    Add-Content (Get-JournalPath) ('{"pid":3999999,"exe":"WINWORD","startTime":"' + (Get-Date).ToString('o') + '","spawnedBy":"ps1-user","ts":"' + (Get-Date).ToString('o') + '"}')
    Add-Content (Get-JournalPath) ('{"pid":3999998,"exe":"WINWORD","startTime":"' + (Get-Date).ToString('o') + '","spawnedBy":"ps1-user","ts":"' + (Get-Date).ToString('o') + '"}')
    $before = (Get-Content (Get-JournalPath)).Count
    $jan = Invoke-JanitorSweep
    Write-Output ("DEAD=" + $jan.dead)
    Write-Output ("SWEPT=" + $jan.swept)
    Write-Output ("AFTER=" + (Get-Content (Get-JournalPath)).Count)
    Write-Output ("SELF-ALIVE=" + [bool](Get-Process -Id $PID -ErrorAction SilentlyContinue))
  `, tmpHome);
  ok('⑤ 死 PID → dead 销账（journal 卫生）', /DEAD=[2-9]/.test(r5.out), '');
  ok('⑧ 执行壳真跑不杀无辜（swept=0 且自身存活）', r5.out.includes('SWEPT=0') && r5.out.includes('SELF-ALIVE=True'), '');
  ok('⑧ 销账后 journal 行数回落', /AFTER=(\d+)/.test(r5.out) && parseInt((r5.out.match(/AFTER=(\d+)/) || [])[1], 10) < 5, '');

  // ⑦ 滚动上限：造 10MB+1 journal → 截断留尾（PS 反引号不用 [char]10 拼接，避免嵌套模板串）
  const r7 = pwsh7(PRELUDE(tmpHome) + `
    $jp = Get-JournalPath
    $lineJson = '{"pid":1,"exe":"WINWORD","startTime":"2026-01-01T00:00:00.000Z","spawnedBy":"ps1-user","ts":"2026-01-01T00:00:00.000Z"}'
    $nl = [string][char]10
    $big = ($lineJson + $nl) * 95000
    [System.IO.File]::WriteAllText($jp, $big)
    Write-Output ("BEFORE=" + (Get-Item $jp).Length)
    $jan = Invoke-JanitorSweep
    Write-Output ("TRUNCATED=" + $jan.truncated)
    Write-Output ("AFTER=" + (Get-Item $jp).Length)
  `, tmpHome);
  ok('⑦ >10MB journal 截断留尾（C）',
    r7.out.includes('TRUNCATED=True') && /AFTER=(\d+)/.test(r7.out) && parseInt((r7.out.match(/AFTER=(\d+)/) || [])[1], 10) < 10 * 1024 * 1024,
    '');

  // 清理临时 HOME
  try { fs.rmSync(tmpHome, { recursive: true, force: true }); } catch { /* ignore */ }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
