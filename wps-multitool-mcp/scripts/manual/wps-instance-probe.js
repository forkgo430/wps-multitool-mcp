/**
 * Input: 无（本机需有 WPS）。可选 --keep 保留新建的实例（默认退出）
 * Output: 一份 JSON + 结论：能否开第二实例 / 是否被复用 / 能否隐藏 / 文档集合是否互通
 * Pos: 多实例决策的探针。判据先于方案 —— 结果决定 DECISION-WPS-INSTANCE.md 走哪条路。
 *
 * 为什么必须先测，不能靠推测
 * --------------------------
 * 三条外部事实互相矛盾，谁都不能当依据：
 *   1) MS Office 的 CoCreateInstance 在已有实例时会【新建】一个进程；
 *   2) WPS 为兼容 Office 宏可能刻意做一致，也可能反过来做单实例；
 *   3) 现有代码 `wps-com.ps1:19` 先 GetActiveObject、失败才 New-Object —— 这条 fallback
 *      从未验证过"已有实例时走不通"，也就是说"能不能开第二实例"一直是未知数。
 *
 * 测法（关键：必须先造出一个"基线实例"再开第二个）
 * ----------------------------------------------
 *   只 CoCreate 一次，测的是"有没有实例"；要测"能不能并存"，必须先有基线再开第二个。
 *   所以本探针自己兜底造基线（用户已开着就借用，绝不退出用户的）。
 *
 * 安全约束（这是台用户机器，探针不许有副作用）
 * ------------------------------------------
 *   1) 只 Quit【本探针自己创建出来】的进程；借用用户已有实例时一律不 Quit，
 *      否则会关掉用户没保存的文档。
 *   2) 不调用 quitWps，不 taskkill。清理失败就如实报残留 PID，不强行杀。
 *   3) 只新建一份空白文档用于验证隔离，测完立即 Close(不保存)。
 */

const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const KEEP = process.argv.includes('--keep');

const PS_SCRIPT = String.raw`
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Get-Snap {
    try {
        return @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue | ForEach-Object {
            [pscustomobject]@{ id = $_.Id; name = $_.ProcessName; title = [string]$_.MainWindowTitle }
        })
    } catch { return @() }
}
function IdsOf($snap) { return @($snap | ForEach-Object { $_.id }) }

$r = [ordered]@{}
$before = @(Get-Snap)
$beforeIds = IdsOf $before
$r.before = @{ count = $before.Count; pids = $beforeIds }

# ---------- A: 造/借一个基线实例（模拟"用户开着的那个 WPS"）----------
$base = $null
$baseOwned = $false
try {
    $base = [Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    $r.a_baseline = @{ source = 'existing(借用用户实例)'; path = [string]$base.Path }
} catch {
    try {
        $base = New-Object -ComObject 'Kwps.Application'
        $baseOwned = $true
        $r.a_baseline = @{ source = 'created(探针自建基线)'; path = [string]$base.Path }
    } catch {
        $r.a_baseline = @{ source = 'none'; error = $_.Exception.Message }
    }
}
Start-Sleep -Milliseconds 1800
$afterBase = @(Get-Snap)
$baseIds = IdsOf $afterBase
$baseProcs = @($afterBase | Where-Object { $beforeIds -notcontains $_.id })
$r.a_baseline.newPids = @($baseProcs | ForEach-Object { $_.id })
$r.a_baseline.newNames = @($baseProcs | ForEach-Object { $_.name })
$r.a_baseline.totalProcs = $afterBase.Count
$r.a_baseline.docs = if ($null -ne $base) { try { [int]$base.Documents.Count } catch { $null } } else { $null }

# ---------- B: 在已有实例的前提下再开一个（这才是"第二进程"）----------
$second = $null
try {
    $second = New-Object -ComObject 'Kwps.Application'
    $r.b_second = @{ ok = $true; path = [string]$second.Path }
} catch {
    $r.b_second = @{ ok = $false; error = $_.Exception.Message }
}
Start-Sleep -Milliseconds 1800
$afterSecond = @(Get-Snap)
$secondIds = IdsOf $afterSecond
$newProcs = @($afterSecond | Where-Object { $baseIds -notcontains $_.id })
$newPids = @($newProcs | ForEach-Object { $_.id })
$r.b_second.newPids = $newPids
$r.b_second.newNames = @($newProcs | ForEach-Object { $_.name })
$r.b_second.reusedExisting = ($newPids.Count -eq 0)
$r.b_second.totalProcs = $afterSecond.Count

# ---------- C: 第二实例能否设成不可见（决定能不能做"无窗口沙箱"）----------
if ($null -ne $second) {
    try {
        $second.Visible = $false
        $rb = $null
        try { $rb = [bool]$second.Visible } catch {}
        $r.c_visibility = @{ setHidden = $true; readback = $rb }
    } catch {
        $r.c_visibility = @{ setHidden = $false; error = $_.Exception.Message }
    }
}

# ---------- D: 两个实例的文档集合是否互通（决定"沙箱"到底能不能隔离）----------
if ($null -ne $second -and $null -ne $base) {
    $doc = $null
    try {
        $doc = $second.Documents.Add()
        Start-Sleep -Milliseconds 800
        $r.d_isolation = @{
            secondDocs = [int]$second.Documents.Count
            baseDocs   = [int]$base.Documents.Count
            isolated   = ([int]$base.Documents.Count -eq [int]$r.a_baseline.docs)
        }
    } catch {
        $r.d_isolation = @{ error = $_.Exception.Message }
    } finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
    }
}

# ---------- 清理：只退出自己造的 ----------
$createdPids = @($newPids) + $(if ($baseOwned) { @($r.a_baseline.newPids) } else { @() })
if ($null -eq $second) {
    $r.cleanup = @{ quit = $false; reason = '没拿到第二个 COM 对象' }
} elseif ($newPids.Count -gt 0 -and -not $env:PROBE_KEEP) {
    $quitOk = $false
    try { $second.Quit(); $quitOk = $true } catch { $r.cleanupError = $_.Exception.Message }
    Start-Sleep -Milliseconds 1800
    $r.cleanup = @{ quit = $quitOk; scope = 'second' }
} elseif ($newPids.Count -gt 0 -and $env:PROBE_KEEP) {
    $r.cleanup = @{ quit = $false; reason = '--keep：保留第二实例'; pids = $newPids }
} else {
    # 复用了基线实例 —— 绝不能 Quit，那会关掉里面的文档
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($second) } catch {}
    $r.cleanup = @{ quit = $false; reason = '第二实例复用了基线，未退出（避免关掉文档）' }
}

if ($baseOwned) {
    try { $base.Quit() } catch {}
    Start-Sleep -Milliseconds 1200
    $r.cleanup.baseOwnedQuit = $true
}

$after = @(Get-Snap)
$afterIds = IdsOf $after
$r.cleanup.remainingPids = @($createdPids | Where-Object { $afterIds -contains $_ })
$r.cleanup.finalCount = $after.Count
$r | ConvertTo-Json -Depth 6 -Compress
`;

function main() {
  const script = `$env:PROBE_KEEP = '${KEEP ? '1' : ''}'\n` + PS_SCRIPT;
  const tmp = path.join(os.tmpdir(), `wps-instance-probe-${Date.now()}.ps1`);
  // 必须带 BOM：PowerShell 5.1 读无 BOM 的 .ps1 会按 ANSI 解释，中文变乱码，
  // 连字符串终止引号都被吃掉（实测报"字符串缺少终止符"）。wps-com.ps1 同样靠 BOM 活命。
  fs.writeFileSync(tmp, '\ufeff' + script, 'utf8');
  try {
    const out = execFileSync(
      'powershell',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', tmp],
      { encoding: 'utf8', timeout: 120000, windowsHide: true }
    );
    const line = out
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter(Boolean)
      .pop();
    if (!line || !line.startsWith('{')) {
      console.error('探针没有返回 JSON，原始输出：\n' + out);
      process.exit(1);
    }
    const r = JSON.parse(line);
    console.log(JSON.stringify(r, null, 2));
    verdict(r);
  } finally {
    try {
      fs.unlinkSync(tmp);
    } catch {}
  }
}

/** 把 JSON 翻成人话结论 —— 免得下一个人对着数字自己猜 */
function verdict(r) {
  const a = r.a_baseline || {};
  const b = r.b_second || {};
  console.log('\n=== 结论 ===');
  if (!b.ok) {
    console.log(`B 失败：第二个 COM 对象都拿不到（${b.error || '未知'}）→ 这条路直接否掉`);
    return;
  }
  if (b.reusedExisting) {
    console.log('B：已有实例时，New-Object 【复用】了同一个进程，没有新建。');
    console.log('   → 想开第二实例不能只靠 CoCreate，得另找入口（独立 COM 分区 / 命令行启动 + 独立管道）。');
  } else {
    console.log(`B：已有实例时，New-Object 新建了 ${b.newPids.length} 个进程（${b.newNames.join(',') || '-'}，PID ${b.newPids.join(',')}）。`);
    console.log('   → 第二实例技术上可达：同一 ProgID 可并存两个进程。');
  }
  const c = r.c_visibility || {};
  console.log(
    c.setHidden
      ? `C：可以设 Visible=false（读回 ${c.readback}）→ "无窗口沙箱实例"有戏`
      : `C：设 Visible=false 失败（${c.error || '未知'}）→ 第二实例会弹出可见窗口，产品形态要重想`
  );
  const d = r.d_isolation || {};
  if (d.error) {
    console.log(`D：隔离测试失败（${d.error}）`);
  } else {
    console.log(
      `D：第二实例建了文档后 —— 第二实例 Documents=${d.secondDocs}，基线实例 Documents=${d.baseDocs}；` +
        (d.isolated ? '基线没被动过 → 隔离成立' : '基线也变了 → 两个实例共享同一文档集合，隔离不成立')
    );
  }
  const cl = r.cleanup || {};
  if (cl.remainingPids && cl.remainingPids.length) {
    console.log(`清理：仍有残留 PID ${cl.remainingPids.join(',')} —— 探针不杀进程，请手动确认`);
  } else {
    console.log(`清理：无残留（最终 WPS 进程数 ${cl.finalCount}）`);
  }
}

try {
  main();
} catch (e) {
  console.error('探针失败：', e && e.message ? e.message : e);
  process.exit(1);
}
