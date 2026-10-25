# Input: Action 名称与 JSON 参数
# Output: WPS COM 调用结果 JSON
# Pos: Windows COM 桥接脚本。一旦我被修改，请更新我的头部注释（Updated: 2026-05-26 15:30:00 CST），以及所属文件夹的md。
# WPS COM Bridge - PowerShell script for WPS COM operations
# Full implementation for Excel, Word, PPT, and common conversions
# Usage: powershell -File wps-com.ps1 -Action <action> -Params <json>

param(
    [string]$Action,
    [string]$Params = "{}",
    # 第 1 期第 3 章轮 1：实例标识。user = 用户正在用的 WPS（GetActiveObject）；
    # sandbox-N = 本进程自己 New-Object 出来的隔离实例（句柄缓存在脚本作用域，常驻复用）。
    [string]$Instance = 'user',
    # 本进程负责哪个应用（word / excel / ppt / common）。word 才做文档寻址前置检查。
    [string]$App = ''
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==================== 实例上下文 ====================
# 这几个变量是【整个进程级】的，不是单次请求的：
#   WpsInstance       —— 决定了 Get-WpsWord 走 GetActiveObject 还是拿沙箱句柄
#   SandboxWordApp    —— 沙箱实例句柄缓存（常驻进程里复用，绝不每次 New-Object）
#   SandboxWordPids   —— 创建沙箱时新出现的 wps PID（quitWps 只回收这几个）
#   LastResolvedDoc    —— 本次请求最终落到哪份文档（由 Output-Json 带回给调用方）
$script:WpsInstance = if ($Instance) { $Instance.Trim() } else { 'user' }
$script:WpsApp = if ($App) { $App.Trim().ToLower() } else { '' }
$script:SandboxWordApp = $null
$script:SandboxWordPids = @()
$script:LastResolvedDoc = $null

<#
.SYNOPSIS  当前进程是不是沙箱实例
.DESCRIPTION
    判据只有一个：WpsInstance 不是 user。不要用"有没有 New-Object 过"来判 ——
    那样第一个请求还没创建实例时会被当成 user，于是第一个动作落到用户的文档上。
#>
function Test-SandboxInstance {
    return ([bool]$script:WpsInstance -and $script:WpsInstance -ne 'user')
}

# ==================== WPS 进程预算（硬上限，禁止超出）====================
# 为什么有这道闸：2026-10-21 真机跑 A1–A10 时，泄漏到 24 个 wps 进程。
# 口径（所有者 2026-10-21 定）：全机 wps/et/wpp 进程总数不超过 3 个。
# 注意：一个实例实测是 2 个进程（主进程 + 辅助），所以 3 个进程的预算 = 最多 1 个实例 + 1 个余量；
#       要同时开沙箱，必须先把上限显式调高（但硬顶仍是 3，不允许绕过）。
function Get-WpsProcessCount {
    try { return @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue).Count } catch { return 0 }
}

function Get-WpsProcBudget {
    $budget = 3
    if ($env:WPS_PROC_BUDGET) {
        try {
            $n = [int]$env:WPS_PROC_BUDGET
            if ($n -gt 0) { $budget = [Math]::Min($n, 3) }   # 硬顶 3，写 8 也只按 3 算
        } catch { }
    }
    return $budget
}

# 回收本进程自己创建过的沙箱 PID（不碰用户窗口），返回回收个数
function Remove-OwnSandboxProcs {
    $n = 0
    foreach ($procId in @($script:SandboxWordPids)) {
        try { Stop-Process -Id $procId -Force -ErrorAction Stop; $n++ } catch { }
    }
    $script:SandboxWordPids = @()
    if ($n -gt 0) { Start-Sleep -Milliseconds 300 }
    return $n
}

# ==================== ROT 附着（PowerShell 7 兼容 · 2026-10-25） ====================
# 恶例复盘：pwsh7(.NET Core) 没有 Marshal::GetActiveObject，6 处调用点全体必抛 →
# 每个动作都走 New-Object 回落 → 每动作新建一个 WPS，几分钟堆出 30+ 进程（执行者轮 7 铁证）。
# 这里用 ole32/oleaut32 的 P/Invoke 实现同语义附着：Marshal 内部就是这条路，5.1 与 7 通用。
if (-not ('WpsRot' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class WpsRot
{
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    // 与 Marshal.GetActiveObject 同语义：查不到 ProgID / 没有实例在跑时抛 COMException
    public static object GetActiveObject(string progId)
    {
        Guid clsid;
        int hr = CLSIDFromProgID(progId, out clsid);
        if (hr != 0) { throw new COMException("CLSIDFromProgID(" + progId + ") failed: 0x" + hr.ToString("X"), hr); }
        object obj;
        GetActiveObject(ref clsid, IntPtr.Zero, out obj);
        return obj;
    }
}
'@ -ErrorAction Stop | Out-Null
}

# 统一入口：全仓"附着已运行实例"只许走这里，不许再直接调 Marshal::GetActiveObject。
# 只取不建；取不到返回 $null（不抛，调用方按 $null 判断）。
function Get-RotObject([string]$ProgId) {
    try { return [WpsRot]::GetActiveObject($ProgId) } catch { return $null }
}

# 本进程（守护）自己创建过的 user 实例记账。缓存 + PID 两条腿：
# 缓存防"ROT 重注册窗口期反复新建"，PID 记账让 quit/收摊时能按 PID 精确回收自己的东西。
$script:UserWordApp = $null
$script:UserExcelApp = $null
$script:UserPptApp = $null
$script:UserCreatedPids = @()

# 新建一个 user 实例并记账（PID 差集 + 看门狗 Track）。返回 $null = 没建出来。
function New-TrackedWpsInstance([string]$ProgId) {
    $before = @()
    try { $before = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
    $app = $null
    try { $app = New-Object -ComObject $ProgId } catch { $app = $null }
    if ($null -eq $app) {
        # 旧路径兜底：个别环境 New-Object 失败但注册表能拿到 CLSID
        try {
            $clsid = (Get-ItemProperty -Path ("HKCU:\Software\Classes\" + $ProgId + "\CLSID") -ErrorAction Stop).'(default)'
            if ($clsid) {
                $type = [Type]::GetTypeFromCLSID($clsid)
                if ($type) { $app = [Activator]::CreateInstance($type) }
            }
        } catch { $app = $null }
    }
    if ($null -eq $app) { return $null }
    try { $app.Visible = $true } catch {}
    $after = @()
    try { $after = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
    $new = @($after | Where-Object { $before -notcontains $_ })
    if ($new.Count -gt 0) {
        $script:UserCreatedPids = @($script:UserCreatedPids + $new | Select-Object -Unique)
        foreach ($procId in $new) {
            try { if ('WpsParentWatchdog' -as [type]) { [WpsParentWatchdog]::Track([int]$procId) } } catch {}
        }
    }
    return $app
}

# 回收本进程创建过的 user 实例（Quit 挂钩与守护收摊共用）。只动自己的，绝不碰 ROT 附着的用户实例。
function Remove-OwnCreatedProcs {
    $n = 0
    foreach ($procId in @($script:UserCreatedPids)) {
        try { Stop-Process -Id $procId -Force -ErrorAction Stop; $n++ } catch { }
    }
    $script:UserCreatedPids = @()
    $script:UserWordApp = $null
    $script:UserExcelApp = $null
    $script:UserPptApp = $null
    if ($n -gt 0) { Start-Sleep -Milliseconds 300 }
    return $n
}

# ==================== COM Object Getters ====================

function Get-WpsExcel {
    # 第 1 期第 2 章第 4 轮（A9）：Fresh 标记"这次取到的实例是不是【新建】的"。
    # 新建的实例一定是空的 —— 用它读列表会得到"0 份"，而文档其实还在（ROT 重注册窗口）。
    $script:WpsExcelFresh = $false
    # ① 先附着已运行的实例（用户优先；pwsh7 兼容见 Get-RotObject）
    $rot = Get-RotObject 'Ket.Application'
    if ($null -ne $rot) { return $rot }
    # ② 再看自己的缓存（ROT 重注册窗口期防反复新建）
    if ($null -ne $script:UserExcelApp) {
        try { $null = $script:UserExcelApp.Workbooks.Count; return $script:UserExcelApp } catch { $script:UserExcelApp = $null }
    }
    # ③ 都没有才新建，并记账
    $app = New-TrackedWpsInstance 'Ket.Application'
    if ($app) { $script:WpsExcelFresh = $true; $script:UserExcelApp = $app; return $app }
    return $null
}

function Get-WpsWord {
    # 第 1 期第 3 章轮 1：沙箱实例走自己的句柄，绝不碰用户的 WPS。
    #
    # 为什么必须缓存：New-Object 每次都新开一个 WPS 进程（实测 +2 进程），
    # 一个常驻进程如果每请求都 New-Object，几分钟就能堆出几十个进程。
    # 为什么要探活：沙箱实例可能被用户手动关掉 / 被系统回收，
    # 缓存里那个 COM 对象就成了僵尸 —— 必须能发现并重建，否则后续全报 RPC 错。
    if (Test-SandboxInstance) {
        if ($null -ne $script:SandboxWordApp) {
            try {
                # 只探一个轻量属性：能取到就认为句柄还活着
                $null = $script:SandboxWordApp.Documents.Count
                return $script:SandboxWordApp
            } catch {
                $script:SandboxWordApp = $null
                $script:SandboxWordPids = @()
            }
        }
        # 新建之前先过预算闸：先回收自己造的无窗口残留，仍不够就【拒绝】——
        # 不新建、不排队、不重试。排队只会继续堆进程，正是要防的事。
        $recycled = Remove-OwnSandboxProcs
        $budget = Get-WpsProcBudget
        $now = Get-WpsProcessCount
        if (($now + 2) -gt $budget) {
            throw ("进程预算不足：当前 $now 个 WPS 进程（刚回收 $recycled 个），上限 $budget，新建沙箱实例还要占 2 个。" +
                   "请先关掉多余的 WPS（含无窗口残留）；上限本身不允许超过 3。")
        }
        $before = @()
        try { $before = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
        $app = $null
        try {
            $app = New-Object -ComObject 'Kwps.Application'
        } catch {
            $app = $null
        }
        if ($null -eq $app) { return $null }
        # 沙箱默认不可见：它不该在用户桌面上弹窗口（DECISION §2 实测 Visible=false 生效）
        try { $app.Visible = $false } catch {}
        $after = @()
        try { $after = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
        # 记下"创建前后新增的 PID" —— quitWps 只回收这批，绝不 taskkill /IM 全杀
        $script:SandboxWordPids = @($after | Where-Object { $before -notcontains $_ })
        foreach ($procId in $script:SandboxWordPids) {
            try { if ('WpsParentWatchdog' -as [type]) { [WpsParentWatchdog]::Track([int]$procId) } } catch {}
        }
        $script:SandboxWordApp = $app
        return $app
    }
    # 第 1 期第 2 章第 4 轮（A9）：Fresh 标记"这次取到的实例是不是【新建】的"。
    $script:WpsWordFresh = $false
    # ① 先附着已运行的实例（用户优先）
    $rot = Get-RotObject 'Kwps.Application'
    if ($null -ne $rot) { return $rot }
    # ② 自己的缓存（防 ROT 重注册窗口期反复新建）
    if ($null -ne $script:UserWordApp) {
        try { $null = $script:UserWordApp.Documents.Count; return $script:UserWordApp } catch { $script:UserWordApp = $null }
    }
    # ③ 都没有才新建，并记账
    $app = New-TrackedWpsInstance 'Kwps.Application'
    if ($app) { $script:WpsWordFresh = $true; $script:UserWordApp = $app; return $app }
    return $null
}

# 只从 ROT（Running Object Table）取正在跑的那个实例，【不新建】。
#
# 为什么需要它（第 1 期第 2 章第 4 轮 · A9）：
# Get-WpsWord 在 ROT 取不到时会新建一个【新的空实例】。
# 而 WPS 退出/重启之后要重新注册 ROT，那几秒里正好取不到 ——
# 于是"暂时看不见文档"被说成了"没有文档"。
# 列表读回必须能区分这两件事，所以给它一条"只取不建"的路。
function Get-WpsWordFromRot {
    return Get-RotObject 'Kwps.Application'
}
function Get-WpsExcelFromRot {
    return Get-RotObject 'Ket.Application'
}
function Get-WpsPpt {
    $rot = Get-RotObject 'Kwpp.Application'
    if ($null -ne $rot) { return $rot }
    if ($null -ne $script:UserPptApp) {
        try { $null = $script:UserPptApp.Presentations.Count; return $script:UserPptApp } catch { $script:UserPptApp = $null }
    }
    $app = New-TrackedWpsInstance 'Kwpp.Application'
    if ($app) { $script:UserPptApp = $app; return $app }
    return $null
}

# ==================== 窗口置前 / 进程退出 辅助 ====================

# 仅获取【已运行】的 WPS 实例（不创建新实例），供 showWps / quitWps 这类不应产生副作用的操作使用
function Get-ActiveWpsApp([string]$ProgId) {
    return Get-RotObject $ProgId
}

# 将已运行的 WPS 主窗口置于前台（按进程主窗口句柄激活）
function Bring-WpsToFront {
    try {
        $procs = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($procs.Count -eq 0) { return $false }
        try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch {}
        $ok = $false
        foreach ($proc in $procs) {
            try { if ([Microsoft.VisualBasic.Interaction]::AppActivate($proc.Id)) { $ok = $true } } catch {}
        }
        return $ok
    } catch { return $false }
}

# 定位 wps.exe 路径（优先取运行中进程的路径，其次查注册表 App Paths）
function Get-WpsExePath {
    try {
        $proc = Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.Path } | Select-Object -First 1
        if ($proc) { return $proc.Path }
    } catch {}
    foreach ($key in @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe'
    )) {
        try {
            $v = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
            if ($v) { return $v }
        } catch {}
    }
    return $null
}

# 解析目标演示文稿：优先按 presentationName 精确定位（避免多文稿打开时 ActivePresentation 漂移），
# 未提供则回退到当前活动文稿；提供了但找不到则返回 $null（让调用方报明确错误，而不是误改其它文稿）。
function Get-TargetPres($ppt, $p) {
    if ($null -eq $ppt) { return $null }
    if ($p.presentationName) {
        try { return $ppt.Presentations.Item([string]$p.presentationName) } catch { return $null }
    }
    return $ppt.ActivePresentation
}

# 取整型参数并带默认值。写成函数而不是内联 $(if ...)：内联 if 在 [int] 转换里
# 解析不稳（实测会让整个 switch 报"语句块中缺少右}"），而且散落在几十处很难统一改。
function Get-IntParam($p, [string]$name, [int]$default = 1) {
    if ($null -eq $p) { return $default }
    $v = $p.$name
    if ($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)) { return $default }
    try { return [int]$v } catch { return $default }
}

# ==================== 文档/工作簿定位（A：不再依赖 ActiveDocument）====================
# 为什么要换掉 ActiveDocument
# ---------------------------
# ActiveDocument 是【弱引用】：它在这些情况下会返回一个"非空但空壳"的对象 ——
#   .Name 为空、.Content/.Styles/.PageSetup 全为 null、.Paragraphs.Count = 0
# 而 Documents.Count 明明 > 0。实测踩过：Documents.Count=2 但 ActiveDocument 取不到内容，
# 后续调用要么诡异地失败，要么直接 RPC_E_CALL_REJECTED。
# 根因是"活动文档"这个概念本身依赖 UI 焦点，而 WPS 的焦点状态在自动化场景下不可靠。
#
# 现在的策略（三层，逐层兜底）
#   1) 显式指定 documentName / workbookName —— 最可靠，直接按名字/路径取
#   2) ActiveDocument/ActiveWorkbook，但必须【通过健康检查】
#   3) 都不行就遍历集合，找第一个健康的
# 找不到就返回 $null，由各动作原有的判空分支报 "No active document"（不新造错误文案）。

# 文档健康检查：Name 非空 且 Content 可取到。
# 为什么要探 Content：空壳对象的 Name 可能是空串，也可能非空但内容取不到（实测见过）。
function Test-DocHealthy($doc) {
    if ($null -eq $doc) { return $false }
    try {
        $n = [string]$doc.Name
        if ([string]::IsNullOrWhiteSpace($n)) { return $false }
    } catch { return $false }
    try { $null = $doc.Content } catch { return $false }
    return $true
}

# 工作簿健康检查：Name 非空 且 Worksheets（或 Sheets）能取到
function Test-BookHealthy($book) {
    if ($null -eq $book) { return $false }
    try {
        $n = [string]$book.Name
        if ([string]::IsNullOrWhiteSpace($n)) { return $false }
    } catch { return $false }
    try { $null = $book.Worksheets } catch { return $false }
    return $true
}

# 按名字在集合里找文档：先按 FullName 全等，再按 Name 全等，最后按"文件名（无扩展名）"全等。
function Find-DocByName($app, [string]$want) {
    if ($null -eq $app -or [string]::IsNullOrWhiteSpace($want)) { return $null }
    $items = @()
    try { $items = @($app.Documents) } catch { return $null }
    $leaf = $want
    if ($leaf.Contains('\')) { $leaf = $leaf.Substring($leaf.LastIndexOf('\') + 1) }
    $stem = $leaf
    if ($stem.Contains('.')) { $stem = $stem.Substring(0, $stem.LastIndexOf('.')) }

    $byLeaf = $null; $byStem = $null
    foreach ($d in $items) {
        if ($null -eq $d) { continue }
        $full = ''; $nm = ''
        try { $full = [string]$d.FullName } catch {}
        try { $nm = [string]$d.Name } catch {}
        if ($full -and $full.Equals($want, [StringComparison]::OrdinalIgnoreCase)) { return $d }
        if ($nm -and $nm.Equals($leaf, [StringComparison]::OrdinalIgnoreCase)) { $byLeaf = $d }
        if ($nm -and $stem -and $nm.Equals($stem + [IO.Path]::GetExtension($nm), [StringComparison]::OrdinalIgnoreCase)) { $byStem = $d }
    }
    if ($byLeaf) { return $byLeaf }
    return $byStem
}

function Find-BookByName($app, [string]$want) {
    if ($null -eq $app -or [string]::IsNullOrWhiteSpace($want)) { return $null }
    $items = @()
    try { $items = @($app.Workbooks) } catch { return $null }
    $leaf = $want
    if ($leaf.Contains('\')) { $leaf = $leaf.Substring($leaf.LastIndexOf('\') + 1) }
    foreach ($b in $items) {
        if ($null -eq $b) { continue }
        $full = ''; $nm = ''
        try { $full = [string]$b.FullName } catch {}
        try { $nm = [string]$b.Name } catch {}
        if ($full -and $full.Equals($want, [StringComparison]::OrdinalIgnoreCase)) { return $b }
        if ($nm -and $nm.Equals($leaf, [StringComparison]::OrdinalIgnoreCase)) { return $b }
    }
    return $null
}

# 目标文稿（Word）。所有文档类动作都应走它，而不是直接读 $word.ActiveDocument。
#
# 注意：本函数在同一动作里可能被调用两次（先判空、再使用），每次都会重新解析。
# 这是有意的 —— 引入跨动作的缓存会带来"用户刚关掉文档，我们还拿着旧对象"的新问题，
# 而重复解析只是几次 COM 属性访问，代价可忽略。
# 记下"这次到底落到哪一份"，由 Output-Json 带回给调用方（__target）。
# 为什么在 Resolve-Doc 里记而不是每个 case 里记：67 个 word 工具都走它，
# 逐个加必然漏；而且漏了也不报错，只表现为"返回里没有 __target"。
function Set-LastResolvedDoc($doc, [string]$by) {
    if ($null -eq $doc) { return }
    $nm = ''; $pth = ''
    try { $nm = [string]$doc.Name } catch {}
    try { $pth = [string]$doc.Path } catch {}
    $script:LastResolvedDoc = @{ name = $nm; path = $pth; by = $by }
}

function Resolve-Doc($app, $p) {
    if ($null -eq $app) { return $null }
    if ($p -and $p.documentName) {
        $d = Find-DocByName $app ([string]$p.documentName)
        if ($d) { Set-LastResolvedDoc $d 'documentName'; return $d }
        # 显式指定却找不到：明确返回 $null，不要静默回退到"活动文档" ——
        # 否则用户以为在改 A，结果改的是 B。
        # 具体报错在 Invoke-WpsAction 的入口前置检查里（那里能一次列出所有打开的文档名）。
        return $null
    }
    try { $act = $app.ActiveDocument } catch { $act = $null }
    if (Test-DocHealthy $act) { Set-LastResolvedDoc $act 'active'; return $act }
    # 空壳 → 遍历找一个健康的
    try {
        foreach ($d in @($app.Documents)) {
            if (Test-DocHealthy $d) { Set-LastResolvedDoc $d 'firstHealthy'; return $d }
        }
    } catch {}
    return $null
}

# 目标工作簿（Excel）
function Resolve-Book($app, $p) {
    if ($null -eq $app) { return $null }
    if ($p -and $p.workbookName) {
        $b = Find-BookByName $app ([string]$p.workbookName)
        if ($b) { return $b }
        return $null
    }
    try { $act = $app.ActiveWorkbook } catch { $act = $null }
    if (Test-BookHealthy $act) { return $act }
    try {
        foreach ($b in @($app.Workbooks)) {
            if (Test-BookHealthy $b) { return $b }
        }
    } catch {}
    return $null
}

# 第 3 轮加：把 $p.sheet 解析成工作表对象，没传就退回活动表。
# 为什么值得单开一个函数：审计（param-shape-audit.js）发现 30 多个 Excel 动作的 TS schema
# 都承诺了 sheet（"工作表名，不填则当前活动表"），PS 侧却直接读 $excel.ActiveSheet ——
# 于是"指定了表也没用，照样改活动表"，属于最坑的一类假成功（不报错、改错了地方）。
# 收敛到这一个函数后：没传 sheet 时行为与原来完全一致（就是 ActiveSheet）；
# 传了但表不存在时返回 $null，让调用方报错，而不是静默改错表。
function Resolve-Sheet($app, $p) {
    if ($null -eq $app) { return $null }
    $wb = $(Resolve-Book $app $p)
    if ($null -eq $wb) { return $null }
    if ($p -and $p.sheet) {
        $key = [string]$p.sheet
        $idx = 0
        if ([int]::TryParse($key, [ref]$idx)) {
            try { return $wb.Sheets.Item($idx) } catch {}
        }
        try { return $wb.Sheets.Item($key) } catch {}
        return $null
    }
    # 必须取【$wb 的】ActiveSheet，不能取 $app.ActiveSheet：
    # 解析到的工作簿未必是活动工作簿（例如探测里刚 Workbooks.Add 出来的临时簿），
    # 两者混用会把参数作用到别的工作簿上 —— 实测导致 __probeExcelCaps 全项报"不支持"。
    try { return $wb.ActiveSheet } catch { return $null }
}

# ==================== 模态对话框探测（C）====================
# 为什么需要
# ----------
# COM 报 RPC_E_CALL_REJECTED(0x80010001) 的最常见原因不是"忙"，而是
# WPS 弹了模态对话框（另存为的确认框、字体替换提示、文档恢复提示…）在等用户点。
# 这时任何 COM 调用都会被拒，而且【干等不会好】—— 用户不点就永远不通。
# 与其让调用方等到超时（120s）才发现，不如立刻告诉他"有个对话框挡着，请先关掉"。
#
# 判据（第一版踩过坑）：我用"可见 + 有标题 + 不占满屏"当对话框特征，
# 结果把 WPS 的【主窗口】也认成了对话框（类名 OpusApp，标题形如 "xxx.docx - WPS Office"）——
# 那意味着每次 COM 被拒都会附一句错误指令"请关闭这个窗口"，比不提示更糟。
#
# 现在的判据：只认【类名】。标准对话框是 #32770，WPS 自绘的一般带 Dialog 字样。
# 宁可漏报（少一句提示，无害），不可误报（给出错误指令，有害）。
if (-not ('WpsModalProbe' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class WpsModalProbe {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
    delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    public static string[] Find(int[] pids) {
        var set = new HashSet<int>(pids);
        var res = new List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            uint wp;
            GetWindowThreadProcessId(h, out wp);
            if (!set.Contains((int)wp)) return true;
            if (!IsWindowVisible(h)) return true;

            var cn = new StringBuilder(256);
            GetClassNameW(h, cn, 256);
            var cls = cn.ToString();
            // 只认对话框类名。主窗口类是 OpusApp / _WwG / bosa_sdm_* 等，一律不算。
            bool isDialog = string.Equals(cls, "#32770", StringComparison.OrdinalIgnoreCase)
                         || cls.IndexOf("Dialog", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!isDialog) return true;

            var sb = new StringBuilder(512);
            GetWindowTextW(h, sb, 512);
            var title = sb.ToString();
            if (string.IsNullOrWhiteSpace(title)) title = "(无标题)";
            res.Add(title + " [" + cls + "]");
            return true;
        }, IntPtr.Zero);
        return res.ToArray();
    }
}
'@ -ErrorAction SilentlyContinue
}

function Get-WpsModalDialogs {
    $pids = @()
    try { $pids = @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
    if ($pids.Count -eq 0) { return @() }
    try { return [WpsModalProbe]::Find([int[]]$pids) } catch { return @() }
}

# 返回被模态对话框挡住的提示语（没有则返回 $null）
function Get-ModalBlockHint {
    $dlgs = @(Get-WpsModalDialogs)
    if ($dlgs.Count -eq 0) { return $null }
    $t = (@($dlgs | Select-Object -First 3) -join '；')
    return "WPS 有对话框挡着（COM 调用被拒 RPC_E_CALL_REJECTED）。请先关掉这些窗口再试：$t"
}

# 统一的 COM 错误文案（C）：若是对话框挡路，把"该关哪个窗口"一并带出去。
#
# 为什么值得单独做：RPC_E_CALL_REJECTED 本身完全无法自解释 —— 看到它只知道
# "调用被拒"，不知道是被谁拒的。而实测里它的绝大多数来源就是模态对话框。
# 加一句"请先关闭 X 窗口"，调用方就不用去猜，也不用等 120s 超时。
# 非对话框类错误原样返回，不加任何噪音。
function Get-CallError($err) {
    $m = [string]$err.Exception.Message
    if (($m -match '0x80010001') -or ($m -match 'RPC_E_CALL_REJECTED') -or ($m -match 'Call was rejected by callee')) {
        $hint = Get-ModalBlockHint
        if ($hint) { return "$m —— $hint" }
    }
    return $m
}

# 只读动作的有限重试（C）
# ----------------------
# 规则：
#   - 只对【只读】动作重试（get*/list*/find* 开头，或在白名单里）。写动作一律不重试 ——
#     请求可能已执行一半，重放会变成执行两次。
#   - 有限次 + 退避（250/500/1000ms），最多 3 次。
#   - 每次失败先看有没有模态对话框：有就直接放弃并把提示带上（干等无意义）。
function Test-ReadOnlyAction([string]$action) {
    if ([string]::IsNullOrWhiteSpace($action)) { return $false }
    if ($action -match '^(get|list|find|read|query|count|search|ping|wpsStatus|__probe)') { return $true }
    return $action -in @('getAppInfo', 'showWps', 'wireCheck')
}

function Invoke-WithReadOnlyRetry($app, [string]$action, $p) {
    $max = 3
    for ($i = 1; $i -le $max; $i++) {
        try {
            return Invoke-WpsAction -Action $action -p $p
        } catch {
            $m = [string]$_.Exception.Message
            $isRejected = ($m -match '0x80010001') -or ($m -match 'RPC_E_CALL_REJECTED')
            if (-not $isRejected) { throw }
            # 有对话框就别耗着了：用户不点，永远不会通
            $hint = Get-ModalBlockHint
            if ($hint) { throw $hint }
            if ($i -ge $max) { throw }
            Start-Sleep -Milliseconds (250 * [math]::Pow(2, $i - 1))
        }
    }
    return $null
}

# 在一个 TextFrame 内查找替换：先精确匹配(保留原换行)；多行 find 未命中时，把 \r\n/\r/\v 与 \n 规范化后再匹配并写回。返回是否发生替换。
function Replace-FrameText($frame, [string]$findText, [string]$replaceText) {
    try {
        if ($null -eq $frame -or -not $frame.HasText) { return $false }
        $text = $frame.TextRange.Text
        if ($null -eq $text) { return $false }
        if ($text.Contains($findText)) {
            $frame.TextRange.Text = $text.Replace($findText, $replaceText)
            return $true
        }
        if ($findText.Contains("`n")) {
            $nf = ((($findText -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            $nt = ((($text -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            if ($nt.Contains($nf)) {
                $frame.TextRange.Text = $nt.Replace($nf, $replaceText)
                return $true
            }
            # 兜底1：去掉所有换行后，若 find 恰为整块文本(PDF抽取的"幻影换行")，整体替换该形状
            $sf = ($findText -replace "[`r`n`v]", "")
            $st = ($text -replace "[`r`n`v]", "")
            if ($sf.Length -gt 0 -and $st -eq $sf) {
                $frame.TextRange.Text = $replaceText
                return $true
            }
            # 兜底2：去换行做子串匹配(find 是大形状里的一段且含幻影换行)，命中后用 stripped→原文 索引映射在原位替换
            if ($sf.Length -gt 0) {
                $origChars = $text.ToCharArray()
                $sb = New-Object System.Text.StringBuilder
                $map = New-Object System.Collections.Generic.List[int]
                for ($k = 0; $k -lt $origChars.Length; $k++) {
                    $ch = $origChars[$k]
                    if ($ch -ne [char]13 -and $ch -ne [char]10 -and $ch -ne [char]11) {
                        [void]$sb.Append($ch); $map.Add($k)
                    }
                }
                $idx = $sb.ToString().IndexOf($sf)
                if ($idx -ge 0) {
                    $startOrig = $map[$idx]
                    $endOrig = $map[$idx + $sf.Length - 1]
                    $before = $text.Substring(0, $startOrig)
                    $after = $text.Substring($endOrig + 1)
                    $frame.TextRange.Text = $before + $replaceText + $after
                    return $true
                }
            }
        }
        return $false
    } catch { return $false }
}

# 递归处理一个形状(及其组合/表格内的所有子形状)的文本替换，返回命中次数。
# 修复：replacePptText 原来只扫顶层 Shapes，组合(Group)内文字(如封面/立项页的旧项目名)替换不到。
function Replace-InShapeTree($shape, [string]$findText, [string]$replaceText) {
    $cnt = 0
    try {
        $isGroup = $false
        try { if ($shape.Type -eq 6) { $isGroup = $true } } catch {}
        if ($isGroup) {
            for ($gi = 1; $gi -le $shape.GroupItems.Count; $gi++) {
                $cnt += Replace-InShapeTree $shape.GroupItems.Item($gi) $findText $replaceText
            }
        } else {
            try { if ($shape.HasTextFrame) { if (Replace-FrameText $shape.TextFrame $findText $replaceText) { $cnt++ } } } catch {}
            try {
                if ($shape.HasTable) {
                    $tbl = $shape.Table
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            try { if (Replace-FrameText $tbl.Cell($rr, $cc).Shape.TextFrame $findText $replaceText) { $cnt++ } } catch {}
                        }
                    }
                }
            } catch {}
        }
    } catch {}
    return $cnt
}

function Output-Json($obj) {
    # 常驻模式下把请求 id 带回响应，供 Node 侧按 id 匹配请求/响应。
    # 直跑模式 $script:WpsReqId 为空，不注入，输出与改造前逐字节一致。
    if ($script:WpsReqId -and ($obj -is [hashtable])) { $obj['id'] = $script:WpsReqId }

    # 第 1 期第 3 章轮 1：把"实际改的是哪一份"带回给调用方。
    # 为什么挂在统一出口：67 个 word 工具的 case 各自 Output-Json，逐个加必漏；
    # 而"改错了地方"恰恰最难自己发现 —— 只看 success: true 是看不出来的。
    # 只在成功响应上带：失败时文档可能根本没解析到，附一份空壳只会误导。
    if ($null -ne $script:LastResolvedDoc -and ($obj -is [hashtable]) -and $obj['success'] -and -not $obj.ContainsKey('__target')) {
        $obj['__target'] = @{
            name     = [string]$script:LastResolvedDoc.name
            path     = [string]$script:LastResolvedDoc.path
            by       = [string]$script:LastResolvedDoc.by
            instance = [string]$script:WpsInstance
        }
    }
    # 每条响应只带一次：不清掉的话，下一个请求（可能压根没解析文档）会继承上一份的名字，
    # 那比不带更坏 —— 它会让调用方确信自己改的是另一份。
    $script:LastResolvedDoc = $null

    $obj | ConvertTo-Json -Depth 10 -Compress
}

function Convert-ColumnLetterToNumber([string]$letters) {
    if (-not $letters) { return $null }
    $letters = $letters.ToUpper().Trim()
    $sum = 0
    foreach ($ch in $letters.ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Convert-ColumnNumberToLetter([int]$col) {
    if ($col -le 0) { return $null }
    $letter = ""
    while ($col -gt 0) {
        $mod = ($col - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $col = [Math]::Floor(($col - 1) / 26)
    }
    return $letter
}

function Convert-HexColorToRgbInt([string]$hex) {
    if (-not $hex) { return $null }
    $hex = $hex.Trim()
    if ($hex.StartsWith('#')) { $hex = $hex.Substring(1) }
    if ($hex.Length -eq 3) {
        $hex = ($hex[0] + $hex[0] + $hex[1] + $hex[1] + $hex[2] + $hex[2])
    }
    if ($hex.Length -ne 6) { return $null }
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

# 颜色 → RGB 整数：先按 #RRGGBB 解析，再退到常见颜色名。
# 为什么不复用 Convert-HexColorToRgbInt：那个只吃十六进制，而 wps_ppt_set_font_color 的
# schema 明确承诺"支持 #FF0000 或 red 这类颜色名"—— 承诺了就得做到，否则用户传 red 会静默不生效。
function Convert-ColorToRgbInt([string]$color) {
    if (-not $color) { return $null }
    # Convert-HexColorToRgbInt 内部 ToInt32 遇到非十六进制会【抛异常】而不是返回 $null
    # （实测 #ZZZZZZ → "未能找到任何可识别的数字"，用户看不懂）。这里接住，走统一的友好报错。
    $rgb = $null
    try { $rgb = Convert-HexColorToRgbInt $color } catch { $rgb = $null }
    if ($null -ne $rgb) { return $rgb }
    $names = @{
        black = '000000'; white = 'FFFFFF'; red = 'FF0000'; green = '00B050'; lime = '00FF00'
        blue = '0070C0'; yellow = 'FFFF00'; orange = 'FFA500'; purple = '800080'
        gray = '808080'; grey = '808080'; pink = 'FFC0CB'; brown = 'A52A2A'
        cyan = '00FFFF'; magenta = 'FF00FF'; navy = '000080'; teal = '008080'
        gold = 'FFD700'; silver = 'C0C0C0'
    }
    $key = $color.Trim().ToLower()
    if ($names.ContainsKey($key)) { return Convert-HexColorToRgbInt $names[$key] }
    return $null
}

function Get-RangeFromAddress($workbook, [string]$address) {
    if ($address -match "^(?<sheet>[^!]+)!(?<range>.+)$") {
        $sheetName = $matches.sheet.Trim("'")
        return $workbook.Sheets.Item($sheetName).Range($matches.range)
    }
    return $workbook.ActiveSheet.Range($address)
}

function Get-AppTypeByExtension([string]$filePath) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower().Trim('.')
    $wordExts = @('doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'txt', 'html', 'htm', 'xml', 'wps', 'wpt')
    $excelExts = @('xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'html', 'htm', 'et', 'ett')
    $pptExts = @('ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt')
    if ($wordExts -contains $ext) { return 'word' }
    if ($excelExts -contains $ext) { return 'excel' }
    if ($pptExts -contains $ext) { return 'ppt' }
    return $null
}

function Get-WordSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Word WdSaveFormat: doc=0, docx=16, pdf=17, rtf=6, xps=18, html=8, txt=2, xml=11
    $map = @{ doc = 0; docx = 16; pdf = 17; rtf = 6; xps = 18; html = 8; htm = 8; txt = 2; xml = 11 }
    return $map[$f]
}

function Get-ExcelFileFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Excel XlFileFormat: xls=-4143, xlsx=51, xlsm=52, xlsb=50, csv=6, html=44, xlt=17, xltx=54, xltm=53
    $map = @{ xls = -4143; xlsx = 51; xlsm = 52; xlsb = 50; csv = 6; html = 44; htm = 44; xlt = 17; xltx = 54; xltm = 53 }
    return $map[$f]
}

function Get-PptSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # PowerPoint PpSaveAsFileType: ppt=1, pptx=24, pptm=25, pdf=32, png=18, jpg=17, gif=16, bmp=19
    $map = @{ ppt = 1; pptx = 24; pptm = 25; pdf = 32; png = 18; jpg = 17; jpeg = 17; gif = 16; bmp = 19 }
    return $map[$f]
}

# ==================== 动作分发（常驻模式复用入口）====================
# 把原先裸的 switch 包进函数，是为了让 wps-comd.ps1 能常驻下来反复调用；
# 脚本直跑（-File wps-com.ps1 -Action X -Params Y）的用法完全不变（见文件末尾入口段）。
# 函数体内原先的 367 个 `exit` 已全部改成 `return` —— 常驻进程里 `exit` 会直接干掉守护进程，
# 而 switch 之后没有任何代码，两者语义等价。
# 注：函数体沿用原缩进（未整体右移），以便 diff 只体现必要改动。
function Invoke-WpsAction {
    param([string]$Action, $p)

# 注意：这里【绝不能】再写 `$p = $Params | ConvertFrom-Json`。
# ①A 常驻化时把原来裸 switch 包成函数，那句解析被一起带了进来，于是每次调用都用
# 脚本级 $Params（守护进程模式下就是默认值 "{}"）把传进来的 $p 覆盖成空对象 ——
# 表现为"所有带参数的动作都收不到参数"，而且不报错。实测于 ② 补参数化动作时发现。
# 直跑模式的解析在文件末尾的入口段里做，那里才是它该待的地方。

# ============================================================================
# 入参归一化（第 3 轮：消灭"TS 传了、PS 不读"的静默丢弃）
# ============================================================================
# 为什么放在 switch 之前
# ------------------
# param-shape-audit.js 扫出来的 90 处候选，本质是一类病：TS 的 schema 承诺了一个参数名，
# 而 PS 的 case 读的是另一个名字（或只读扁平键）。后果分两种：
#   静默丢弃 —— "返回 success 但什么都没做"（最坏，用户以为改完了）；
#   响亮失败 —— PS 缺必填键直接报错（用户至少知道自己没成功）。
# 本块把【静默丢弃】那一类在入口就抹平：改名的改名、嵌套的摊平，case 体一行都不用动。
# 值形态要转换的（比如 startRow/endRow → rows 数组）不在这里做，留在 case 里。
#
# 两条规则都是"只在对方没给时才补" —— 显式传扁平参数的调用方（wps_execute_method 裸调、
# 历史脚本）行为完全不变。
$ParamAlias = @{
    'insertPptChart'       = @{ chartType = 'type' }
    'addPptHyperlink'      = @{ url = 'address' }
    'alignShapes'          = @{ shapeIndices = 'names' }
    'distributeShapes'     = @{ shapeIndices = 'names' }
    'groupShapes'          = @{ shapeIndices = 'names' }
    'smartDistribute'      = @{ shapeIndices = 'shapes' }
    'addAnimation'         = @{ shapeIndex = 'shapeName' }
    'addAnimationPreset'   = @{ shapeIndex = 'shapeName' }
    'applyTransitionToAll' = @{ effect = 'transition' }
    'findReplace'          = @{ find = 'findText'; replace = 'replaceText' }
    'deleteSheet'          = @{ name = 'sheet' }
    'switchSheet'          = @{ name = 'sheet' }
    'renameSheet'          = @{ oldName = 'sheet' }
    'copySheet'            = @{ name = 'sheet' }
    'moveSheet'            = @{ name = 'sheet' }
    'deletePptImage'       = @{ imageIndex = 'shapeIndex' }
    'deleteTextBox'        = @{ textboxIndex = 'shapeIndex' }
    'setTextBoxText'       = @{ textboxIndex = 'shapeIndex' }
    'setTextBoxStyle'      = @{ textboxIndex = 'shapeIndex' }
    'setImageStyle'        = @{ imageIndex = 'shapeIndex' }
    'removeAnimation'      = @{ animationIndex = 'index' }
    'replacePptImage'      = @{ imagePath = 'filePath' }
    'setBackgroundImage'   = @{ imagePath = 'path'; filePath = 'path' }
    'insertPptImage'       = @{ filePath = 'path'; imagePath = 'path' }
    'insertExcelImage'     = @{ filePath = 'path'; imagePath = 'path' }
    'openPresentation'     = @{ filePath = 'path' }
    'setHyperlink'         = @{ url = 'address'; text = 'textToDisplay' }
    'setPageSetup'         = @{ marginTop = 'topMargin'; marginBottom = 'bottomMargin'; marginLeft = 'leftMargin'; marginRight = 'rightMargin' }
    'setShapeZOrder'       = @{ order = 'zOrder' }
    'setSlideNumber'       = @{ show = 'visible' }
    'setPptDateTime'       = @{ show = 'visible'; autoUpdate = 'useFixed' }
    'setBorder'            = @{ borderStyle = 'style' }
    'setSlideTransition'   = @{ transition = 'effect' }
    # 这条以前把 rotation:{x,y,z} 整个对象改名成 rotationX，于是 case 里
    # [float]$rx 收到一个 PSCustomObject → "无法将类型 psobject 转换为 float"。
    # case 已改成自己读 rotation.x/y/z，这里不能再改名。
    # 'set3DRotation'     = @{ rotation = 'rotationX' }
    'autoFilter'           = @{ column = 'field' }
    'createKpiCards'       = @{ data = 'cards' }
    'createMiniCharts'     = @{ data = 'items' }
    'createOrgChart'       = @{ data = 'nodes' }
    'createProgressBar'    = @{ value = 'progress' }
    # sortRange 的 column 现在由 case 自己换算成列字母（TS 发的是数字，不是列地址），
    # 不能再在这里改成 keyColumn —— 否则 case 会拿一个数字去 Range()。
    'subtotal'             = @{ columns = 'totalColumns' }
    'transpose'            = @{ source = 'sourceRange' }
    'setAnimationOrder'    = @{ animationIndex = 'from'; newOrder = 'to' }
    'insertImage'          = @{ imagePath = 'path' }
    'addDataValidation'    = @{ type = 'validationType'; formula = 'formula1' }
}

# 嵌套对象摊平：TS 传 style:{...}/format:{...}/shadow:{...}，PS 读扁平键。
# 把嵌套对象的每个属性提升成 $p 的同名键（已存在的绝不覆盖）。
$StyleFlatten = @(
    'setPptTableStyle', 'setPptTableCellStyle', 'setPptTableRowStyle', 'setPptChartStyle',
    'setCellStyle', 'setCellFormat', 'setShapeBorder', 'setShapeGradient', 'setShapeShadow',
    'setTextBoxStyle', 'setImageStyle', 'setMasterBackground', 'setBackgroundGradient',
    'setSlideBackground', 'addConditionalFormat', 'addMasterElement', 'addPageIndicator',
    'addTitleDecoration', 'protectSheet', 'protectWorkbook'
)
# 这些动作的嵌套键不叫 style，单独点名
$FlattenKeys = @{
    'setCellFormat'         = @('format')
    'setShapeBorder'        = @('border')
    'setShapeGradient'      = @('gradient')
    'setShapeShadow'        = @('shadow')
    'setBackgroundGradient' = @('gradient')
    'setMasterBackground'   = @('background')
    'setSlideBackground'    = @('background')
    'addConditionalFormat'  = @('condition', 'format')
    'addMasterElement'      = @('element')
    'addPageIndicator'      = @('position')
    'addTitleDecoration'    = @('style')
    'protectSheet'          = @('protect')
    'protectWorkbook'       = @('protect')
}


    # openAfterExport：TS 的 schema 承诺"导出后顺便打开"，以前 PS 压根不读这个键 ——
    # 用户勾了没反应，也没有任何提示。放在回包之前统一做，失败只记 warnings 不影响导出结果。
    function Open-Exported($path, $flag) {
        if (-not $flag) { return }
        try {
            if ($path -and (Test-Path -LiteralPath $path)) { Start-Process -FilePath $path | Out-Null }
        } catch { }
    }

if ($null -ne $p -and $p -is [psobject]) {
    $have = @($p.PSObject.Properties.Name)
    # 1) 改名：TS 的文档名 → PS 实际读的名
    if ($ParamAlias.ContainsKey($Action)) {
        foreach ($from in @($ParamAlias[$Action].Keys)) {
            $to = $ParamAlias[$Action][$from]
            if (($have -contains $from) -and -not ($have -contains $to)) {
                $p | Add-Member -NotePropertyName $to -NotePropertyValue $p.$from -Force
                $have = @($p.PSObject.Properties.Name)
            }
        }
    }
    # 2) 摊平：嵌套对象的属性提升成 $p 的扁平键
    if ($StyleFlatten -contains $Action) {
        $objKeys = if ($FlattenKeys.ContainsKey($Action)) { $FlattenKeys[$Action] } else { @('style') }
        foreach ($ok in $objKeys) {
            $obj = $p.$ok
            if ($null -eq $obj) { continue }
            foreach ($prop in @($obj.PSObject.Properties)) {
                if (-not ($have -contains $prop.Name)) {
                    $p | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
                    $have = @($p.PSObject.Properties.Name)
                }
            }
        }
    }
    # 3) 值形态转换：这几个不是"换个名字"能解决的 —— PS 要数组/枚举，TS 给的是起止行或布尔。
    #    同样放入口统一做，case 体一行不动。
    # （sortRange 的 ascending→order 已改到 TS 侧做：PS 的 order 收 'desc'/'asc' 字符串，
    #   在这里塞数字曾经让"降序"永远落成升序 —— 第 4 轮改回 TS 发字符串。）
    if (@('hideRows', 'showRows') -contains $Action -and ($have -contains 'startRow') -and -not ($have -contains 'rows')) {
        # TS 的 startRow/endRow（闭区间）→ PS 的 rows 数组。
        # 修之前 PS 只认 row/rows，于是 wps_excel_hide_rows / show_rows 必然报
        # "row/rows required" —— 工具等于坏的（响亮失败，不算假成功，但一样不可用）。
        $from = [int]$p.startRow
        $to = if (($have -contains 'endRow') -and $p.endRow) { [int]$p.endRow } else { $from }
        $list = @()
        for ($r = $from; $r -le $to; $r++) { $list += $r }
        $p | Add-Member -NotePropertyName rows -NotePropertyValue $list -Force
        $have = @($p.PSObject.Properties.Name)
    }
    if (@('hideColumns', 'showColumns') -contains $Action -and ($have -contains 'startColumn') -and -not ($have -contains 'columns')) {
        # 列号可能是数字也可能是字母；PS 只认数字，统一在这里转
        $fi = 0
        $ti = 0
        if (-not [int]::TryParse([string]$p.startColumn, [ref]$fi)) { $fi = Convert-ColumnLetterToNumber([string]$p.startColumn) }
        if (-not [int]::TryParse([string]$p.endColumn, [ref]$ti)) { $ti = Convert-ColumnLetterToNumber([string]$p.endColumn) }
        $list = @()
        for ($c = $fi; $c -le $ti; $c++) { $list += $c }
        $p | Add-Member -NotePropertyName columns -NotePropertyValue $list -Force
        $have = @($p.PSObject.Properties.Name)
    }
}

# ==================== 目标钉住的前置检查（第 1 期第 3 章轮 1）====================
# 传了 documentName 就【必须】能找到，找不到直接报错并列出所有打开的文档。
#
# 为什么放在这里而不是各 case 里：
#   Resolve-Doc 找不到时返回 $null，各 case 只会报 "No active document" ——
#   那句话有两个意思（"没开文档" 和 "没有你要的那份"），用户没法据此改对。
#   放在入口能一次性列出候选，而且只写一处、67 个工具全覆盖。
#
# 为什么只在 word 路（$script:WpsApp='word'）做：
#   excel/ppt 的 parameter 是 workbookName / presentationName，本轮不碰；
#   而直跑模式没有 App 信息，保持旧行为（不新增失败路径）。
if ($script:WpsApp -eq 'word' -and $null -ne $p -and $p.documentName) {
    $wantDoc = [string]$p.documentName
    if (-not [string]::IsNullOrWhiteSpace($wantDoc)) {
        $wordProbe = Get-WpsWord
        if ($null -eq $wordProbe) {
            Output-Json @{ success = $false; error = "WPS Word not running" }
            return
        }
        $hit = Find-DocByName $wordProbe $wantDoc
        if ($null -eq $hit) {
            $names = @()
            try {
                foreach ($d in @($wordProbe.Documents)) {
                    if (Test-DocHealthy $d) { $names += [string]$d.Name }
                }
            } catch {}
            $list = if ($names.Count -gt 0) { ($names -join '、') } else { '（当前没有打开的文档）' }
            Output-Json @{
                success = $false
                error   = "[勿重试·先改参数] 找不到文档「$wantDoc」（实例：$($script:WpsInstance)）。当前打开的文档共有 $($names.Count) 份：$list。要改其中一份请用完整文件名；要新建请用 openDocument/createDocument。"
            }
            return
        }
    }
}

switch ($Action) {

    # ==================== Common ====================
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
    }

    # 快照工具的前置查询：文档在哪 / 有没有未保存改动 / 只读没。
    # 这里【只读】，也不像 Get-WpsExcel 那样顺手拉起应用 —— 查个状态就把 WPS 拽起来是副作用，
    # 用户会在任务管理器里看到一堆自己没开的进程。所以只用 GetActiveWpsApp（只找已运行的实例）。
    # 真正的写文件放在 Node 侧：PS 一旦碰文件系统，路径写错就是把用户文档写坏。
    "getActiveDocumentState" {
        # 以前的实现是"Excel → Word → PPT 谁先有就报谁"，这在两个应用都开着时会报错人：
        # 实测过一次——用户正在 Word 里写，Excel 里还挂着一个探测残留的工作簿，
        # 结果快照把 Excel 那个工作簿当成"当前文档"，复制了错的文件。
        # 现在的口径：优先用调用方指定的 app；没指定就① 跳过取不到名字的空壳（WPS 的
        # ActiveDocument 是弱引用，Count>0 时也可能给空壳）② 优先"有未保存改动"的那个
        # （那才是需要保命的）③ 还分不出来就按 excel → word → ppt。
        $want = if ($p.app) { ([string]$p.app).ToLower() } else { '' }
        $cand = @()

        if (-not $want -or $want -eq 'excel') {
            $excel = Get-ActiveWpsApp 'Ket.Application'
            if ($null -ne $excel) {
                try {
                    $book = $excel.ActiveWorkbook
                    if ($null -ne $book -and -not [string]::IsNullOrEmpty([string]$book.Name)) {
                        $full = ""; try { $full = [string]$book.FullName } catch {}
                        # hasPath 必须用 Path 判，不能用 FullName —— 未保存的文档 FullName 也会
                        # 拼出一个"默认目录\文字编号1"（文件并不存在），拿它去 copy 只会 ENOENT。
                        $dir = ""; try { $dir = [string]$book.Path } catch {}
                        $saved = $true; try { $saved = [bool]$book.Saved } catch {}
                        $ro = $false; try { $ro = [bool]$book.ReadOnly } catch {}
                        $cand += @{ app = "excel"; name = [string]$book.Name; fullName = $full; dir = $dir; saved = $saved; readOnly = $ro; hasPath = (-not [string]::IsNullOrEmpty($dir)) }
                    }
                } catch {}
            }
        }

        if (-not $want -or $want -eq 'word') {
            $word = Get-ActiveWpsApp 'Kwps.Application'
            if ($null -ne $word) {
                try {
                    $doc = $word.ActiveDocument
                    if ($null -ne $doc -and -not [string]::IsNullOrEmpty([string]$doc.Name)) {
                        $full = ""; try { $full = [string]$doc.FullName } catch {}
                        $dir = ""; try { $dir = [string]$doc.Path } catch {}
                        $saved = $true; try { $saved = [bool]$doc.Saved } catch {}
                        $ro = $false; try { $ro = [bool]$doc.ReadOnly } catch {}
                        $cand += @{ app = "word"; name = [string]$doc.Name; fullName = $full; dir = $dir; saved = $saved; readOnly = $ro; hasPath = (-not [string]::IsNullOrEmpty($dir)) }
                    }
                } catch {}
            }
        }

        if (-not $want -or $want -eq 'ppt') {
            $ppt = Get-ActiveWpsApp 'Kwpp.Application'
            if ($null -ne $ppt) {
                try {
                    $pres = $ppt.ActivePresentation
                    if ($null -ne $pres -and -not [string]::IsNullOrEmpty([string]$pres.Name)) {
                        $full = ""; try { $full = [string]$pres.FullName } catch {}
                        $dir = ""; try { $dir = [string]$pres.Path } catch {}
                        $saved = $true; try { $saved = [bool]$pres.Saved } catch {}
                        $ro = $false; try { $ro = [bool]$pres.ReadOnly } catch {}
                        $cand += @{ app = "ppt"; name = [string]$pres.Name; fullName = $full; dir = $dir; saved = $saved; readOnly = $ro; hasPath = (-not [string]::IsNullOrEmpty($dir)) }
                    }
                } catch {}
            }
        }

        if ($cand.Count -eq 0) {
            Output-Json @{ success = $false; error = "No active document" }
            return
        }
        # 未保存优先（那才是需要保命的），否则保持 excel → word → ppt 的稳定顺序
        $chosen = $cand[0]
        foreach ($c in $cand) { if (-not $c.saved) { $chosen = $c; break } }
        Output-Json @{ success = $true; data = $chosen }
    }
    # 【临时诊断】列出挡路的对话框。用来验 C 项：Get-WpsModalDialogs 能不能真找到窗口。
    "__probeStyleAdd" {
        # 最小对比实验：探测里 styles.add 报 supported，实现里却 E_FAIL。
        # 逐项排除差异（字面量 vs 变量 / 参数类型 / 名字字符），找出到底是哪一处。
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $r = @{}

        # A: 与探测完全一致 —— 纯字面量
        try {
            $sA = $doc.Styles.Add("__probe_literal__", 1)
            $r["A_literal"] = if ($null -eq $sA) { "null" } else { "ok: " + [string]$sA.NameLocal }
            try { $doc.Styles.Item("__probe_literal__").Delete() } catch {}
        } catch { $r["A_literal"] = "FAIL: " + $_.Exception.Message }

        # B: 第一个参数走变量（与实现一致），第二个仍是字面量
        try {
            $nm = "__probe_var__"
            $sB = $doc.Styles.Add($nm, 1)
            $r["B_varName"] = if ($null -eq $sB) { "null" } else { "ok: " + [string]$sB.NameLocal }
            try { $doc.Styles.Item("__probe_var__").Delete() } catch {}
        } catch { $r["B_varName"] = "FAIL: " + $_.Exception.Message }

        # C: 两个参数都走变量（与实现完全一致）
        try {
            $nm2 = "__probe_bothvar__"
            $st = 1
            $sC = $doc.Styles.Add($nm2, $st)
            $r["C_bothVar"] = if ($null -eq $sC) { "null" } else { "ok: " + [string]$sC.NameLocal }
            try { $doc.Styles.Item("__probe_bothvar__").Delete() } catch {}
        } catch { $r["C_bothVar"] = "FAIL: " + $_.Exception.Message }

        # D: 中文样式名（实现里传的是中文名，探测用的是英文名）
        try {
            $sD = $doc.Styles.Add("__p01测试样式__", 1)
            $r["D_chineseName"] = if ($null -eq $sD) { "null" } else { "ok: " + [string]$sD.NameLocal }
            try { $doc.Styles.Item("__p01测试样式__").Delete() } catch {}
        } catch { $r["D_chineseName"] = "FAIL: " + $_.Exception.Message }

        # E: 从 PSCustomObject 取参数（实现是 $p.name 这种）
        try {
            $pE = @{ name = "__probe_pobj__"; styleType = $null }
            $nmE = [string]$pE.name
            $stE = 1
            if ($pE.styleType -eq "character") { $stE = 2 }
            $sE = $doc.Styles.Add($nmE, $stE)
            $r["E_pobjName"] = if ($null -eq $sE) { "null" } else { "ok: " + [string]$sE.NameLocal }
            try { $doc.Styles.Item("__probe_pobj__").Delete() } catch {}
        } catch { $r["E_pobjName"] = "FAIL: " + $_.Exception.Message }

        # F: 同一个文档上连续 Add 两次（看是不是第一次就污染了状态）
        try {
            $sF = $doc.Styles.Add("__probe_twice__", 1)
            $first = if ($null -eq $sF) { "null" } else { "ok" }
            try { $sF2 = $doc.Styles.Add("__probe_twice2__", 1) } catch { }
            $r["F_twice"] = $first + " / 第二次: " + $(if ($null -eq $sF2) { "null" } else { "ok" })
            try { $doc.Styles.Item("__probe_twice__").Delete() } catch {}
            try { $doc.Styles.Item("__probe_twice2__").Delete() } catch {}
        } catch { $r["F_twice"] = "FAIL: " + $_.Exception.Message }

        Output-Json @{ success = $true; data = $r }
    }

    "__probeClose" {
        # 试不同的关闭方式，找出哪种在"未命名未保存文档"上管用。
        # 背景：Close(0) 报 RPC_E_CALL_REJECTED，但读文档列表是成功的 ——
        # 说明不是 COM 整体坏掉，而是关闭这个特定动作被拒。
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $r = @{}
        $docs = @($word.Documents)
        $r.countBefore = $docs.Count
        $i = 0
        foreach ($d in $docs) {
            $i++
            $nm = ''
            try { $nm = [string]$d.Name } catch {}
            $full = ''
            try { $full = [string]$d.FullName } catch {}
            # 第 5 轮：判"存过盘没有"必须看 Path。未保存文档的 FullName 非空（等于文档名），
            # 用 [bool]$full 判会把每份未保存文档都标成 saved=true —— 这个探针的输出就是在骗人。
            $dirP = ''
            try { $dirP = [string]$d.Path } catch {}
            $tag = "$i(nm=$nm saved=$(-not [string]::IsNullOrEmpty($dirP)))"
            # 方式1：Close($false)
            try { $d.Close($false); $r["1_CloseFalse"] += "$tag:ok "; continue } catch { $r["1_CloseFalse"] += "$tag:FAIL($($_.Exception.Message.Substring(0,[Math]::Min(40,$_.Exception.Message.Length)))) " }
            # 方式2：先 SaveAs 到临时文件再关
            try {
                $tmp = [IO.Path]::GetTempFileName() + '.docx'
                $d.SaveAs([ref]$tmp, [ref]$missing)
                $d.Close($false)
                $r["2_SaveAsThenClose"] += "$tag:ok "
            } catch { $r["2_SaveAsThenClose"] += "$tag:FAIL " }
            # 方式3：Documents.Close 集合级
            try { $word.Documents.Close($false); $r["3_CollectionClose"] += "$tag:ok "; break } catch { $r["3_CollectionClose"] += "$tag:FAIL " }
        }
        $r.countAfter = @($word.Documents).Count
        Output-Json @{ success = $true; data = $r }
    }

    "__probeDialogs" {
        $dlgs = @(Get-WpsModalDialogs)
        Output-Json @{ success = $true; data = @{ dialogs = $dlgs; count = $dlgs.Count; hint = (Get-ModalBlockHint) } }
    }
    # ==================== 诊断与修复（D）====================
    "wpsRepair" {
        # 出问题时的"一键诊断+修复"。做两件事：
        #   1) 诊断：三个应用各自的实例、文档数、每个文档的健康度、挡路的对话框
        #   2) 修复：关掉【无名文档】（没路径没标题的，通常是脚本遗留的空壳）
        #      —— 只关无名文档，不动用户自己的文件；已命名的文档一律保留。
        $report = @{ apps = @{}; dialogs = @(); closed = @() }
        try { $report.dialogs = @(Get-WpsModalDialogs) } catch {}

        foreach ($a in @(
            @{ key = "excel"; progId = "Ket.Application";  coll = "Workbooks";     app = $null },
            @{ key = "word";  progId = "Kwps.Application"; coll = "Documents";     app = $null },
            @{ key = "ppt";   progId = "Kwpp.Application"; coll = "Presentations"; app = $null }
        )) {
            $inst = Get-ActiveWpsApp $a.progId
            $entry = @{ running = ($null -ne $inst); documents = @() }
            if ($null -ne $inst) {
                try { $entry.name = [string]$inst.Name } catch {}
                try { $entry.path = [string]$inst.Path } catch {}
                $docs = @()
                try { $docs = @($inst.($a.coll)) } catch {}
                $entry.count = $docs.Count
                foreach ($d in $docs) {
                    if ($null -eq $d) { continue }
                    $nm = ''; $full = ''; $dirP = ''; $healthy = $false
                    try { $nm = [string]$d.Name } catch {}
                    try { $full = [string]$d.FullName } catch {}
                    try { $dirP = [string]$d.Path } catch {}
                    if ($a.key -eq "word") { $healthy = Test-DocHealthy $d }
                    else { $healthy = Test-BookHealthy $d }
                    $entry.documents += @{
                        name = $nm
                        # 第 5 轮修：saved 必须看 Path。原来用 FullName 判 —— 未保存文档的
                        # FullName 非空（等于文档名），于是每份没存过的文档都报 saved=true。
                        saved = (-not [string]::IsNullOrWhiteSpace($dirP))
                        healthy = $healthy
                    }
                    # 只关"无名"文档：没有标题也没有路径 —— 那是自动化残留的空壳。
                    # ⚠️ 关闭用【集合级】Close，不要逐个 $d.Close()。
                    #   实测（__probeClose）：逐个关闭未命名文档会被 WPS 拒
                    #   （RPC_E_CALL_REJECTED），而 $app.Documents.Close($false) 一次成功
                    #   （countBefore=2 → countAfter=0）。这是 WPS 12.1 的实际行为。
                    #   真正执行在下面循环结束后统一做（先收集，再一次关）。
                    # 第 5 轮备注（刻意不改判据）：这里的 $full 仍是 FullName，而 Name 永远非空，
                    # 所以本分支实际上【不会触发】—— 也就是 repair 从不自动关文档。
                    # 这是偏安全的一侧：改成"没有 Path 就算空壳"会去关用户没存盘的活文档，
                    # 那是比"少关几个空壳"严重得多的事故。要真修，得先定义"空壳"的可靠特征。
                    if ([string]::IsNullOrWhiteSpace($nm) -and [string]::IsNullOrWhiteSpace($full)) {
                        $unnamed += $a.key
                    }
                }
                # 活动文档健康度单独报一份，便于判断"是不是 ActiveDocument 空壳"
                try {
                    $act = $null
                    if ($a.key -eq "word") { $act = $inst.ActiveDocument }
                    else { $act = $inst.ActiveWorkbook }
                    $entry.activeHealthy = if ($a.key -eq "word") { Test-DocHealthy $act } else { Test-BookHealthy $act }
                } catch { $entry.activeHealthy = $false }
            }
            $report.apps[$a.key] = $entry
        }

        $hint = $null
        if (@($report.dialogs).Count -gt 0) {
            $hint = "检测到挡路对话框，请先关闭：" + ((@($report.dialogs | Select-Object -First 3)) -join '；')
        }
        $report.hint = $hint
        # 收尾：对收集到的"有无名文档"的应用做一次【集合级】关闭
        # （逐个 $d.Close() 会被 WPS 拒，集合级一次成功 —— 见 __probeClose 实测）
        foreach ($k in ($unnamed | Select-Object -Unique)) {
            $app = $null
            $coll = 'Documents'
            if ($k -eq 'word') { $app = Get-ActiveWpsApp 'Kwps.Application' }
            elseif ($k -eq 'excel') { $app = Get-ActiveWpsApp 'Ket.Application'; $coll = 'Workbooks' }
            elseif ($k -eq 'ppt') { $app = Get-ActiveWpsApp 'Kwpp.Application'; $coll = 'Presentations' }
            if ($null -eq $app) { continue }
            try {
                $app.$coll.Close($false)
                $report.closed += "$k`: 已关闭无名文档（集合级 Close）"
            } catch {
                $report.closed += "$k`: 关闭失败 $($_.Exception.Message)"
            }
        }

        Output-Json @{ success = $true; data = $report }
    }

    # 只读状态探针：报告三件套分别有没有已附着的实例。
    # 关键区别：这里【只用 Get-ActiveWpsApp】——它只取已运行的实例，不会创建新的。
    # 对照 getAppInfo：那个走 Get-WpsExcel/Word/Ppt，拿不到就 New-Object 起一个，
    # 用它探状态会把 WPS 意外拉起来，而且只报第一个找到的应用。
    "wpsStatus" {
        # 身份判据说明（别再用 .Name 当权威）：
        # WPS 为了兼容 Excel/Word 宏，故意自报 Application.Name = "Microsoft Excel" /
        # "Microsoft Word"，Version 也报成 Office 的兼容版本号（如 12.0）。实测本机
        # Ket.Application 拿到的对象 Path 指向 WPS 安装目录、且机器上根本没装 EXCEL.EXE
        # —— 所以它是 WPS 表格本人，不是真 Excel。
        # 因此身份以【我们请求的 ProgID】+【Application.Path 落在哪个产品目录】为准，
        # 自报名单独放在 reportedName，只作参考，不作判据。
        $apps = @(
            @{ key = "excel"; progId = "Ket.Application";  app = "WPS 表格" },
            @{ key = "word";  progId = "Kwps.Application"; app = "WPS 文字" },
            @{ key = "ppt";   progId = "Kwpp.Application"; app = "WPS 演示" }
        )
        $states = @{}
        foreach ($a in $apps) {
            $entry = @{ running = $false; app = $a.app; progId = $a.progId }
            $inst = Get-ActiveWpsApp $a.progId
            if ($null -ne $inst) {
                $entry.running = $true
                try { $entry.reportedName = [string]$inst.Name } catch {}
                try { $entry.version = [string]$inst.Version } catch {}
                try {
                    $ip = [string]$inst.Path
                    if ($ip) {
                        $entry.path = $ip
                        # 由安装目录反推产品，作为身份交叉验证
                        $entry.product = if ($ip -match 'WPS\s*Office') { 'WPS Office' }
                                         elseif ($ip -match 'Microsoft\s+Office') { 'Microsoft Office' }
                                         else { 'unknown' }
                    }
                } catch {}
                # 文档数：三个应用各自的集合名不一样，按顺序试
                foreach ($coll in @("Workbooks", "Documents", "Presentations")) {
                    try {
                        $c = $inst.$coll
                        if ($null -ne $c) { $entry.documents = [int]$c.Count; break }
                    } catch {}
                }
            }
            $states[$a.key] = $entry
        }
        Output-Json @{ success = $true; data = @{
            apps = $states
            psPid = $PID
            psVersion = $PSVersionTable.PSVersion.ToString()
            daemon = ($env:WPS_COM_DAEMON -eq '1')
        } }
    }

    # 让 WPS 可见并置于前台：① 激活已有窗口 ② 唤起 COM 实例 ③ 兜底启动可见实例
    "showWps" {
        # 沙箱实例：只激活【自己那一个】。
        # 原来会遍历 ROT + 全部同名进程去找窗口，多实例时不知道该激活谁，
        # 而且会把用户那份也拉到前台（看着像"我明明只要看沙箱"）。
        if (Test-SandboxInstance) {
            $app = $script:SandboxWordApp
            if ($null -eq $app) {
                Output-Json @{
                    success = $false
                    error   = "沙箱实例「$($script:WpsInstance)」还没创建。先调一个带 target.instance 的动作，它会被自动拉起。"
                }
                return
            }
            try { $app.Visible = $true } catch {}
            try { $app.Activate() } catch {}
            Output-Json @{
                success = $true
                data    = @{
                    instance = $script:WpsInstance
                    visible  = $true
                    note     = '只激活本沙箱实例，未遍历其他 WPS 进程'
                }
            }
            return
        }

        $apps = @()
        $visible = 0
        $activated = 0
        foreach ($item in @(
            @{ name = 'Word';  progId = 'Kwps.Application' },
            @{ name = 'Excel'; progId = 'Ket.Application'  },
            @{ name = 'PPT';   progId = 'Kwpp.Application' }
        )) {
            $app = Get-ActiveWpsApp $item.progId
            if ($null -ne $app) {
                $apps += $item.name
                try { $app.Visible = $true; $visible++ } catch {}
                try { $app.Activate(); $activated++ } catch {}
            }
        }

        $front = Bring-WpsToFront
        $launched = $false

        # 兜底：没有可激活的窗口（COM 自动化实例本身无窗口）时，启动一个可见的 WPS 实例
        if (-not $front -and $p.launch -ne $false) {
            $exe = Get-WpsExePath
            if ($exe) {
                try {
                    Start-Process -FilePath $exe | Out-Null
                    $launched = $true
                    # 轮询等待可见窗口出现（最多约 14 秒）
                    for ($i = 0; $i -lt 14; $i++) {
                        Start-Sleep -Seconds 1
                        if (Bring-WpsToFront) { $front = $true; break }
                    }
                } catch {}
            }
        }

        if ($apps.Count -eq 0 -and -not $front -and -not $launched) {
            Output-Json @{ success = $false; error = "未能唤起 WPS 窗口（无可激活窗口，且启动可见实例失败）" }
            return
        }

        # 最终确认：只要系统里存在可见的 WPS 主窗口即视为成功
        if (-not $front) {
            $hasVisible = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }).Count -gt 0
            if ($hasVisible) { $front = $true }
        }

        $msg = if ($front) { "WPS 窗口已置于前台" } elseif ($launched) { "已启动 WPS 可见实例" } else { "已将 WPS 设为可见" }
        Output-Json @{ success = $true; data = @{ apps = $apps; visible = $visible; activated = $activated; foreground = $front; launched = $launched; message = $msg } }
    }

    # 结束 WPS（优雅退出；未保存内容默认拒绝，不做全量杀进程）
    "quitWps" {
        # 沙箱实例：只退【我们自己创建的那一个】，绝不 taskkill /IM 全杀。
        # 原来那条兜底 `taskkill /F /IM wps.exe /T` 是多实例的头号杀手 ——
        # 一退就把用户正开着的那份一起杀了，而且看起来"退出成功"。
        if (Test-SandboxInstance) {
            $app = $script:SandboxWordApp
            $killed = @()
            if ($null -ne $app) {
                try { $app.DisplayAlerts = 0 } catch {}
                # 第 2 期第 2 章轮 2 真机实测：沙箱里有未保存文档时，$app.Quit() 会卡在
                # 「是否保存」对话框上（无人值守 = 挂死，实测 300 秒超时，而且桥会因此
                # 杀掉整个 word 子进程）。沙箱里的文档本来就是可丢的，先逐个标 Saved 再关。
                try {
                    for ($di = $app.Documents.Count; $di -ge 1; $di--) {
                        try {
                            $dd = $app.Documents.Item($di)
                            $dd.Saved = $true
                            $dd.Close(0)   # wdDoNotSaveChanges
                        } catch {}
                    }
                } catch {}
                try { $app.Quit() } catch {}
            }
            Start-Sleep -Milliseconds 700
            # 按创建时记下的 PID 收尸：Quit 之后 WPS 常常还留着几个进程（DECISION §2 实测）
            foreach ($procId in @($script:SandboxWordPids)) {
                try {
                    $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
                    if ($proc) {
                        # 只杀这一个 PID，且不用 /T（避免连带子进程里别人的东西）
                        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
                        $killed += $procId
                    }
                } catch {}
            }
            $script:SandboxWordApp = $null
            $script:SandboxWordPids = @()
            Output-Json @{
                success = $true
                data    = @{
                    quit        = @('sandbox')
                    instance    = $script:WpsInstance
                    killedPids  = $killed
                    userUntouched = $true
                }
            }
            return
        }

        # 先数未保存的文档：直接 Quit 会因 DisplayAlerts=0 把未保存改动静默丢掉，
        # 而这里退出影响的是用户开着的全部 WPS —— 所以未保存时默认拒绝，confirm=true 才放行。
        $unsaved = @()
        foreach ($progId in @('Kwps.Application','Ket.Application','Kwpp.Application')) {
            $app = Get-ActiveWpsApp $progId
            if ($null -eq $app) { continue }
            $kind = switch ($progId) { 'Kwps.Application' { '文档' } 'Ket.Application' { '工作簿' } default { '演示文稿' } }
            $coll = $null
            try { $coll = $app.Documents } catch {}
            if (-not $coll) { try { $coll = $app.Workbooks } catch {} }
            if (-not $coll) { try { $coll = $app.Presentations } catch {} }
            if ($null -eq $coll) { continue }
            foreach ($item in $coll) {
                $saved = $true
                try { $saved = [bool]$item.Saved } catch {}
                if (-not $saved) {
                    $name = ''
                    try { $name = $item.Name } catch {}
                    $unsaved += "「$name」（$kind）"
                }
            }
        }
        if ($unsaved.Count -gt 0 -and $p.confirm -ne $true) {
            Output-Json @{ success = $false; error = "[勿重试·先保存] 退出会丢弃这些未保存内容：$($unsaved -join '、')。确认要放弃就传 confirm=true；想保住先 wps_common_save。" }
            return
        }

        $quit = @()
        foreach ($progId in @('Kwps.Application','Ket.Application','Kwpp.Application')) {
            $app = Get-ActiveWpsApp $progId
            if ($null -ne $app) {
                try { $app.DisplayAlerts = 0 } catch {}
                try { $app.Quit() } catch {}
                $quit += $progId
            }
        }
        Start-Sleep -Milliseconds 900
        # 本进程自己创建过的实例（user 记账 + 沙箱）按 PID 直接收掉 —— 这些是我们的，
        # 不受未保存护栏约束（护栏只保护 ROT 附着的用户实例）。任务结束/异常都能杀。
        $ownKilled = (Remove-OwnCreatedProcs) + (Remove-OwnSandboxProcs)
        # 不做全量 taskkill：Quit 之后的残留多半是用户自己开的其它 WPS 实例
        # （ROT 每个 ProgID 只能激活一个，COM 够不着的那几个 /IM 一杀就误杀）。
        # 列出剩余进程交给用户决定，wpscloudsvr 是 WPS 自己的云服务，从来不该杀。
        $leftovers = @(Get-Process -Name wps -ErrorAction SilentlyContinue | ForEach-Object {
            @{ pid = $_.Id; title = $_.MainWindowTitle }
        })
        $left = $leftovers.Count
        $msg = if ($left -eq 0) { "WPS 已退出" } else { "本会话连接的 WPS 已退出；另有 $left 个 WPS 进程在运行（多为用户自己打开的实例，未强行关闭）" }
        Output-Json @{ success = $true; data = @{ quit = $quit; forced = $false; ownKilled = $ownKilled; remaining = $left; leftovers = $leftovers; message = $msg } }
    }

    "wireCheck" {
        Output-Json @{ success = $true; data = @{ message = "WPS MCP Bridge 已连接" } }
    }

    "getAppInfo" {
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $excel.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "excel"; appName = $excel.Name; hasSelection = $hasSelection } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $word.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "word"; appName = $word.Name; hasSelection = $hasSelection } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $ppt.ActiveWindow.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "ppt"; appName = $ppt.Name; hasSelection = $hasSelection } }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "getSelectedText" {
        # 第 1 期第 3 章轮 1：选区是【全局焦点】，钉不住。
        # 带 target 或沙箱实例时显式拒绝 —— 宁可报错，也不许静默读到别人那份的选区。
        if (Test-SandboxInstance) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 选区（Selection）依赖全局焦点，沙箱实例里钉不住。要在沙箱里取文本，请用 getDocumentText / getParagraphs 这类按文档寻址的动作。" }
            return
        }
        if ($null -ne $p -and $p.target) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 带 target 的会话里不能用选区类动作：用户切一下焦点，读到的就是另一份文档的选区。请改用按文档寻址的动作（getDocumentText / getParagraphs）。" }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $text = $word.Selection.Text
                Output-Json @{ success = $true; data = @{ text = $text.Trim() } }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            return
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $text = if ($sel -is [string]) { $sel } else { $sel.Text }
                Output-Json @{ success = $true; data = @{ text = [string]$text } }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $text = $sel.TextRange.Text
                    Output-Json @{ success = $true; data = @{ text = [string]$text } }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "setSelectedText" {
        # 同 getSelectedText：写选区比读更危险（会改到用户当前光标处的文档）
        if (Test-SandboxInstance) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 写选区依赖全局焦点，沙箱实例里钉不住。要在沙箱里写文本，请用 insertText / setRangeText 这类按文档寻址的动作。" }
            return
        }
        if ($null -ne $p -and $p.target) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 带 target 的会话里不能用选区类动作：焦点一换就写到别的文档里去了。请改用 insertText（可带 documentName）。" }
            return
        }
        $text = if ($null -ne $p.text) { [string]$p.text } else { "" }
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $word.Selection.Text = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            return
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $sel.Value2 = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $sel.TextRange.Text = $text
                    Output-Json @{ success = $true }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "save" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $(Resolve-Book $excel $p).Save()
            Output-Json @{ success = $true; app = "excel" }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $(Resolve-Doc $word $p).Save()
            Output-Json @{ success = $true; app = "word" }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $ppt.ActivePresentation.Save()
            Output-Json @{ success = $true; app = "ppt" }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "saveAs" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; return }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }

        # 第 5 轮：保存之后【读回真实路径】，而不是把自己传进去的 $path 原样回显。
        # 为什么必须读回：SaveAs 可能被 WPS 改写落点（格式不符自动换扩展名、路径被规范化、
        # 同名已存在时弹选择框）。回显 $path 等于替 WPS 担保"你给的路径就是落点"。
        # 判据用 .Path（不是 .FullName —— 未保存文档的 FullName 就等于文档名，永远非空）。
        # WPS 保存后 .Path 有极短延迟，读到空就退回用户给的路径并标 verified=false，
        # 让调用方知道"这次没核实成功"，而不是假装核实了。
        $resolveSavedPath = {
            param($obj, $requested)
            $actual = ''
            try { $actual = [string]$obj.Path } catch {}
            if ([string]::IsNullOrEmpty($actual)) {
                return @{ path = $requested; requestedPath = $requested; verified = $false
                          note = "保存后没读到文档路径（可能是 WPS 有延迟），按你给的路径回报；落盘前请自行确认" }
            }
            $same = $false
            try { $same = ([IO.Path]::GetFullPath($actual) -ieq [IO.Path]::GetFullPath($requested)) } catch {}
            return @{ path = $actual; requestedPath = $requested; verified = $same }
        }

        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            $wb = $(Resolve-Book $excel $p)
            if ($null -eq $excel -or $null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
            $format = Get-ExcelFileFormat $p.format
            if ($null -ne $format) { $wb.SaveAs($path, $format) }
            else { $wb.SaveAs($path) }
            $r = & $resolveSavedPath $wb $path
            Output-Json @{ success = $true; data = @{ path = $r.path; requestedPath = $r.requestedPath; verified = $r.verified; note = $r.note; appType = "excel" } }
            return
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            $docItem = $(Resolve-Doc $word $p)
            if ($null -eq $word -or $null -eq $docItem) { Output-Json @{ success = $false; error = "No active document" }; return }
            $format = Get-WordSaveFormat $p.format
            if ($null -ne $format) { $docItem.SaveAs($path, $format) }
            else { $docItem.SaveAs($path) }
            $r = & $resolveSavedPath $docItem $path
            Output-Json @{ success = $true; data = @{ path = $r.path; requestedPath = $r.requestedPath; verified = $r.verified; note = $r.note; appType = "word" } }
            return
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt -or $null -eq $ppt.ActivePresentation) { Output-Json @{ success = $false; error = "No active presentation" }; return }
            $pres = $ppt.ActivePresentation
            $format = Get-PptSaveFormat $p.format
            if ($null -ne $format) { $pres.SaveAs($path, $format) }
            else { $pres.SaveAs($path) }
            $r = & $resolveSavedPath $pres $path
            Output-Json @{ success = $true; data = @{ path = $r.path; requestedPath = $r.requestedPath; verified = $r.verified; note = $r.note; appType = "ppt" } }
            return
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "openFile" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; return }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
            $excel.Workbooks.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            return
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
            $word.Documents.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            return
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
            $ppt.Presentations.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            return
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "createDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $n0 = 0
        try { $n0 = [int]$word.Documents.Count } catch {}
        $doc = $word.Documents.Add()

        # 等文档真正就绪（关键，之前漏了）
        # ----------------------------------
        # 原实现是 `Documents.Add() | Out-Null` 然后立刻返回 success，零等待。
        # 后果：新建的文档要过一小会儿才取得到 Content/Styles，紧接着的任何动作
        # 都会拿到"非空但空壳"的对象 —— 实测反复踩到（表现为莫名 "No active document"、
        # 以及样式/修订/交叉引用一连串失败）。
        #
        # 实现要点（踩过）：
        #   1) 每次都重新 `Documents.Item(Count)` 取最新文档，不要复用 Add() 的返回值 ——
        #      实测复用那个引用会让 Test-DocHealthy 永远 false（6 秒超时）。
        #   2) 等不到也【按成功返回】，只带 warning：Resolve-Doc 有遍历兜底，
        #      而且实测"不等 + 调用方自己等一会"是能工作的。抛异常会把后续全堵死。
        $warn = $null
        $ready = $false
        for ($i = 0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 150
            try {
                $cur = $word.Documents.Item($word.Documents.Count)
                if (Test-DocHealthy $cur) { $ready = $true; break }
            } catch {}
        }
        if (-not $ready) {
            $warn = "文档已创建但 6 秒内未通过健康检查（可能是 WPS 正在初始化）。" +
                    "后续操作会自动遍历集合找健康文档；若报错请重试或检查对话框。"
        }
        $name = ""
        try { $name = [string]$word.Documents.Item($word.Documents.Count).Name } catch {}
        $data = @{ count = [int]$word.Documents.Count; name = $name }
        if ($warn) { $data.warning = $warn }
        Output-Json @{ success = $true; data = $data }
    }

    "createPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = $ppt.Presentations.Add()
        # 必须回 data：Node 侧的 handler 判的是 success && data，
        # 只回 success 会被当成失败（"新建演示文稿失败: undefined" —— 实测踩过）
        $nm = ""
        try { $nm = [string]$pres.Name } catch {}
        Output-Json @{ success = $true; data = @{ name = $nm; count = $ppt.Presentations.Count } }
    }

    "convertToPDF" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $wb = $(Resolve-Book $excel $p)
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $wb.ExportAsFixedFormat(0, $outputPath)
            Open-Exported $outputPath $p.openAfterExport
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; opened = [bool]$p.openAfterExport } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $doc = $(Resolve-Doc $word $p)
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $doc.ExportAsFixedFormat($outputPath, 17)
            Open-Exported $outputPath $p.openAfterExport
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; opened = [bool]$p.openAfterExport } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $pres.SaveAs($outputPath, 32)
            Open-Exported $outputPath $p.openAfterExport
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; opened = [bool]$p.openAfterExport } }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "convertFormat" {
        $targetFormat = $p.targetFormat
        if (-not $targetFormat) { Output-Json @{ success = $false; error = "targetFormat required" }; return }
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $wb = $(Resolve-Book $excel $p)
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-ExcelFileFormat $targetFormat
            if ($null -ne $format) { $wb.SaveAs($outputPath, $format) } else { $wb.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; targetFormat = $targetFormat } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $doc = $(Resolve-Doc $word $p)
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-WordSaveFormat $targetFormat
            if ($null -ne $format) { $doc.SaveAs($outputPath, $format) } else { $doc.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; targetFormat = $targetFormat } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-PptSaveFormat $targetFormat
            if ($null -ne $format) { $pres.SaveAs($outputPath, $format) } else { $pres.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; targetFormat = $targetFormat } }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    # ==================== Excel Basic ====================
    "getActiveWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheetCount = $wb.Sheets.Count; sheets = $sheets } }
    }

    "getCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $cell = $sheet.Cells.Item($p.row, $p.col)
        Output-Json @{ success = $true; data = @{ value = $cell.Value2; text = $cell.Text; formula = $cell.Formula } }
    }

    "setCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        # 第 1 期第 2 章第 4 轮（B1）：与 setRangeData 同一个雷 —— 原样 `.Value2 = $p.value`
        # 在 daemon 里会抛 "无法将类型为 System.Int32 的对象强制转换为类型 System.String"。
        # 同一个分派：数字写数字、文本按文本格式写（保住 "007" 这类形状）、布尔写布尔、null 清空。
        $cell = $sheet.Cells.Item($p.row, $p.col)
        $v = $p.value
        if ($null -eq $v) {
            $cell.ClearContents()
        } elseif ($v -is [bool]) {
            $cell.NumberFormat = 'General'
            $cell.Value2 = [bool]$v
        } elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
            $cell.NumberFormat = 'General'
            $cell.Value2 = [double]$v
        } else {
            $cell.NumberFormat = '@'
            $cell.Value2 = [string]$v
        }
        Output-Json @{ success = $true }
    }

    "getRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $data = @()
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $range.Columns.Count; $c++) { $row += $range.Cells.Item($r, $c).Value2 }
            $data += ,@($row)
        }
        Output-Json @{ success = $true; data = @{ data = $data } }
    }

    "setRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        # 第 1 期第 2 章第 2 轮：两句分支完全相同 + 没给 sheet 时 Sheets.Item($null) 会抛。
        # wps_excel_evaluate_formula 的"写回单元格"要复用这个动作（它没有 sheet 参数），所以先修这层。
        $sheet = $(Resolve-Sheet $excel $p)
        if ($null -eq $sheet) { Output-Json @{ success = $false; error = "No active sheet" }; return }
        $range = $sheet.Range($p.range)
        for ($r = 0; $r -lt $p.data.Count; $r++) {
            for ($c = 0; $c -lt $p.data[$r].Count; $c++) {
                $cell = $range.Cells.Item($r + 1, $c + 1)
                $v = $p.data[$r][$c]
                # 第 1 期第 2 章第 4 轮（B1）：写数字必失败，报
                #   "无法将类型为 System.Int32 的对象强制转换为类型 System.String"
                # 根因不是"值是数字"，是【值的来源】：daemon 里 $p 来自 ConvertFrom-Json，
                # 取出来的元素是 PSObject 包着的，COM 绑定器接不住这种包装，
                # 于是退回去按 String 转换 → 数字转不过去就抛。
                #   实测（daemon 内，52 组对照）：原样赋值 Int32 过、Double 抛、String 抛；
                #   而【显式强转】过的值四种全过 —— 强转产出的是"裸" .NET 值，绑定器认。
                # 但直接 [string] 化会毁掉形状："007" 变 7、"2024-01-15" 变日期序列号 45306。
                # 所以按类型分派：
                #   数字 → 先把单元格格式设回常规再写数字（WPS 存成数值，=SUM() 能算）
                #   文本 → 先设成文本格式('@')再写，"007"/"1e5"/日期串原样保住
                #   布尔 → 写成真正的布尔单元格
                #   null → 清空
                if ($null -eq $v) { $cell.ClearContents(); continue }
                if ($v -is [bool]) {
                    $cell.NumberFormat = 'General'
                    $cell.Value2 = [bool]$v
                } elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                    $cell.NumberFormat = 'General'
                    $cell.Value2 = [double]$v
                } else {
                    $cell.NumberFormat = '@'
                    $cell.Value2 = [string]$v
                }
            }
        }
        Output-Json @{ success = $true }
    }

    "diagnoseFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = $cell.Formula
        $errorType = $null
        $diagnosis = ""
        $suggestion = ""
        $precedents = @()
        if ($value -is [string] -and $value.StartsWith('#')) {
            $errorType = $value
            switch ($errorType) {
                "#REF!" { $diagnosis = "引用了不存在的单元格或区域"; $suggestion = "检查引用区域是否被删除或移动" }
                "#N/A" { $diagnosis = "查找函数未找到匹配值"; $suggestion = "确认查找值存在，或检查匹配条件" }
                "#VALUE!" { $diagnosis = "参数类型不正确或运算类型不匹配"; $suggestion = "检查函数参数类型和引用单元格" }
                "#NAME?" { $diagnosis = "函数名或名称拼写错误"; $suggestion = "检查函数名是否正确" }
                "#DIV/0!" { $diagnosis = "除数为零"; $suggestion = "检查除数单元格，避免除以零" }
                "#NUM!" { $diagnosis = "数值无效或超出范围"; $suggestion = "检查函数参数范围" }
                "#NULL!" { $diagnosis = "交集为空"; $suggestion = "检查引用区域的交集是否存在" }
                default { $diagnosis = "未知错误"; $suggestion = "检查公式与引用" }
            }
        }
        try {
            $refs = $cell.DirectPrecedents
            if ($refs -ne $null) {
                foreach ($area in $refs.Areas) { $precedents += $area.Address() }
            }
        } catch {}
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; currentValue = $value; errorType = $errorType; diagnosis = $diagnosis; suggestion = $suggestion; precedents = $precedents } }
    }

    "cleanData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $opsResult = @()
        foreach ($op in $p.operations) {
            $success = $true
            $message = ""
            switch ($op) {
                "trim" {
                    foreach ($cell in $range) {
                        if ($cell.Value2 -is [string]) { $cell.Value2 = $cell.Value2.Trim() }
                    }
                    $message = "已去除前后空格"
                }
                "remove_duplicates" {
                    $colCount = $range.Columns.Count
                    $cols = @()
                    for ($i = 1; $i -le $colCount; $i++) { $cols += $i }
                    $range.RemoveDuplicates($cols, 1)
                    $message = "已删除重复行"
                }
                "unify_date" {
                    foreach ($cell in $range) {
                        try {
                            if ($cell.Value2) {
                                $dt = [DateTime]::FromOADate($cell.Value2)
                                $cell.Value2 = $dt.ToString("yyyy-MM-dd")
                            }
                        } catch {}
                    }
                    $message = "已统一日期格式"
                }
                "remove_empty_rows" {
                    for ($i = $range.Rows.Count; $i -ge 1; $i--) {
                        $row = $range.Rows.Item($i)
                        $isEmpty = $true
                        foreach ($cell in $row.Cells) { if ($cell.Value2) { $isEmpty = $false; break } }
                        if ($isEmpty) { $row.Delete() }
                    }
                    $message = "已删除空行"
                }
                default { $success = $false; $message = "不支持的操作" }
            }
            $opsResult += @{ operation = $op; success = $success; message = $message }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; operations = $opsResult; message = "cleanData completed" } }
    }

    "createPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sourceRange = Get-RangeFromAddress $wb $p.sourceRange
        $destSheet = if ($p.destinationSheet) { $wb.Sheets.Item($p.destinationSheet) } else { $excel.ActiveSheet }
        $destCell = $destSheet.Range($p.destinationCell)
        $pivotName = if ($p.tableName) { $p.tableName } else { "PivotTable" + [Guid]::NewGuid().ToString("N").Substring(0, 6) }
        $cache = $wb.PivotCaches().Create(1, $sourceRange)
        $table = $cache.CreatePivotTable($destCell, $pivotName)
        foreach ($fieldName in $p.rowFields) {
            $field = $table.PivotFields($fieldName)
            $field.Orientation = 1
            $field.Position = 1
        }
        if ($p.columnFields) {
            foreach ($fieldName in $p.columnFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 2
                $field.Position = 1
            }
        }
        if ($p.filterFields) {
            foreach ($fieldName in $p.filterFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 3
                $field.Position = 1
            }
        }
        foreach ($vf in $p.valueFields) {
            $field = $table.PivotFields($vf.field)
            $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
            $func = $funcMap[$vf.aggregation]
            if ($null -eq $func) { $func = -4157 }
            $table.AddDataField($field, $vf.field, $func) | Out-Null
        }
        Output-Json @{ success = $true; data = @{ pivotTableName = $pivotName; location = $destCell.Address(); rowCount = $table.RowRange.Rows.Count; columnCount = $table.TableRange1.Columns.Count } }
    }

    "updatePivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $table = $null
        if ($p.pivotTableName) {
            try { $table = $sheet.PivotTables($p.pivotTableName) } catch {}
        }
        if ($null -eq $table -and $p.pivotTableCell) {
            try { $table = $sheet.Range($p.pivotTableCell).PivotTable } catch {}
        }
        if ($null -eq $table) { Output-Json @{ success = $false; error = "PivotTable not found" }; return }
        $ops = @()
        if ($p.addRowFields) {
            foreach ($f in $p.addRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 1; $field.Position = 1
                $ops += @{ operation = "addRowFields"; success = $true; message = $f }
            }
        }
        if ($p.removeRowFields) {
            foreach ($f in $p.removeRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeRowFields"; success = $true; message = $f }
            }
        }
        if ($p.addColumnFields) {
            foreach ($f in $p.addColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 2; $field.Position = 1
                $ops += @{ operation = "addColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.removeColumnFields) {
            foreach ($f in $p.removeColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.addFilterFields) {
            foreach ($f in $p.addFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 3; $field.Position = 1
                $ops += @{ operation = "addFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.removeFilterFields) {
            foreach ($f in $p.removeFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.addValueFields) {
            foreach ($vf in $p.addValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $table.AddDataField($field, $vf.field, $func) | Out-Null
                $ops += @{ operation = "addValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.updateValueFields) {
            foreach ($vf in $p.updateValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $field.Function = $func
                $ops += @{ operation = "updateValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.removeValueFields) {
            foreach ($vf in $p.removeValueFields) {
                try { $table.DataFields($vf).Orientation = 0 } catch {}
                $ops += @{ operation = "removeValueFields"; success = $true; message = $vf }
            }
        }
        if ($p.refresh) { try { $table.RefreshTable() | Out-Null; $ops += @{ operation = "refresh"; success = $true; message = "refreshed" } } catch {} }
        Output-Json @{ success = $true; data = @{ pivotTableName = $table.Name; operations = $ops } }
    }

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        # 第 1 期第 2 章第 2 轮：原来两句分支完全相同（`if ($p.sheet -is [int])` 两边都写
        # `$wb.Sheets.Item($p.sheet)`），而且没给 sheet 时是 `Sheets.Item($null)` —— 会抛。
        # wps_excel_auto_sum 要映射到这个动作（它没有 sheet 参数），所以必须先修好这层解析。
        $sheet = $(Resolve-Sheet $excel $p)
        if ($null -eq $sheet) { Output-Json @{ success = $false; error = "No active sheet" }; return }
        $formula = [string]$p.formula
        if (-not $formula) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 要给 formula（如 =SUM(A1:A10)）" }
            return
        }
        $range = if ($p.range) { $sheet.Range([string]$p.range) } elseif ($p.row -and $p.col) { $sheet.Cells.Item($p.row, $p.col) } else { $null }
        if ($null -eq $range) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 要给 range（如 A11）或 row+col" }
            return
        }
        $range.Formula = $formula
        # 读回：公式写进去之后单元格的【值】是多少。写公式不等于算出了数
        # （引用空区域会得 0，引用坏区域会得 #REF!），把值一起回给调用方才算诚实。
        $value = $null
        try { $value = $range.Value2 } catch {}
        Output-Json @{ success = $true; data = @{ range = [string]$p.range; formula = $formula; value = $value } }
    }

    # 第 1 期第 2 章第 2 轮新增：wps_excel_evaluate_formula 一直调 evaluateFormula，
    # 而 ps1 里没有这个 case → 每次都是 Unknown action（这工具从上线起就没成功过）。
    # 能力就是 Excel 的 Application.Evaluate：只算不写；给了 cell 才把结果落格。
    # 与 setFormula 的分工：setFormula 是【把公式写进单元格（落盘）】，本动作是【算一个值出来（不落盘）】。
    "evaluateFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $formula = [string]$p.formula
        if (-not $formula) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 要给 formula（如 =SUM(A1:A10) 或 1+2）" }
            return
        }
        if (-not $formula.StartsWith('=')) { $formula = '=' + $formula }
        $result = $null
        try { $result = $excel.Evaluate($formula) } catch {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 公式无法计算：$($formula)。$($_.Exception.Message)" }
            return
        }
        if ($null -eq $result) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 公式 $formula 没有返回值（语法错或引用了不存在的区域）" }
            return
        }
        # 【只算不写】—— 这里刻意不落盘。
        # 为什么：cancellable.ts:67 把 evaluateFormula 归在 READ_ONLY_EXTRA 里（客户端会据此
        # 跳过写操作确认）。一个"被声明为只读"的动作如果偷偷写单元格，就是安全语义矛盾。
        # 要落盘由 TS 侧再发一次 setRangeData（写入动作，注解本身是写入型），职责分开。
        Output-Json @{ success = $true; data = @{ formula = $formula; result = $result } }
    }

    # 第 1 期第 2 章第 2 轮新增：wps_excel_set_zoom 一直调 setZoom，ps1 里没有 → Unknown action。
    # 这是【窗口缩放】（ActiveWindow.Zoom，10-400），不是【打印缩放】（PageSetup.Zoom）——
    # ps1 里 2706 行那个 $ps.Zoom 是 PageSetup 的，属于打印设置，两回事。
    "setZoom" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $pct = 0
        if (-not [int]::TryParse([string]$p.percent, [ref]$pct)) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] percent 必须是整数（收到「$($p.percent)」）" }
            return
        }
        if ($pct -lt 10 -or $pct -gt 400) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 缩放比例必须在 10-400 之间（收到 $pct）" }
            return
        }
        $win = $excel.ActiveWindow
        if ($null -eq $win) { Output-Json @{ success = $false; error = "没有活动窗口（工作簿可能未激活）" }; return }
        $win.Zoom = $pct
        # 读回真实值：WPS 可能夹取到自己的范围，回包要说"现在到底是多少"
        $actual = $null
        try { $actual = $win.Zoom } catch {}
        Output-Json @{ success = $true; data = @{ zoom = $actual; requested = $pct } }
    }

    "setCellFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        if ($p.numberFormat) { $range.NumberFormat = $p.numberFormat }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.numberFormat } }
    }

    "setCellStyle" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = [bool]$p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = [bool]$p.italic }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $range.Interior.Color = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $range.Font.Color = $fc }
        }
        if ($p.horizontalAlignment) {
            $hAlignMap = @{ left = -4131; center = -4108; right = -4152 }
            $hAlign = $hAlignMap[$p.horizontalAlignment]
            if ($null -ne $hAlign) { $range.HorizontalAlignment = $hAlign }
        }
        if ($p.verticalAlignment) {
            $vAlignMap = @{ top = -4160; center = -4108; bottom = -4107 }
            $vAlign = $vAlignMap[$p.verticalAlignment]
            if ($null -ne $vAlign) { $range.VerticalAlignment = $vAlign }
        }
        if ($null -ne $p.border -and $p.border) {
            $range.Borders.LineStyle = 1
            if ($p.borderColor) {
                $bc = Convert-HexColorToRgbInt([string]$p.borderColor)
                if ($null -ne $bc) { $range.Borders.Color = $bc }
            }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setBorder" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $styleMap = @{ thin = 1; medium = 2; thick = 4; double = 6; none = 0 }
        $style = $styleMap[$p.style]
        if ($null -eq $style) { $style = 1 }
        $position = if ($p.position) { $p.position } else { "all" }
        $borders = @()
        if ($position -eq "all" -or $position -eq "outside") { $borders += 7, 8, 9, 10 }
        if ($position -eq "all" -or $position -eq "inside") { $borders += 11, 12 }
        if ($position -eq "left") { $borders += 7 }
        if ($position -eq "top") { $borders += 8 }
        if ($position -eq "bottom") { $borders += 9 }
        if ($position -eq "right") { $borders += 10 }
        $colorValue = $null
        if ($p.color) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
        foreach ($b in $borders) {
            $border = $range.Borders.Item($b)
            if ($style -eq 0) {
                $border.LineStyle = -4142
            } else {
                $border.LineStyle = 1
                $border.Weight = $style
            }
            if ($null -ne $colorValue) { $border.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; style = $p.style; position = $position } }
    }

    "copyFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $sourceRange = $sheet.Range($p.source)
        $targetRange = $sheet.Range($p.target)
        $sourceRange.Copy()
        $targetRange.PasteSpecial(-4122)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.source; target = $p.target } }
    }

    "clearFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $range.ClearFormats()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "addConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $formatType = if ($p.type) { $p.type } else { "cellValue" }
        if ($formatType -eq "cellValue") {
            $operatorMap = @{ greater = 5; greaterThan = 5; less = 6; lessThan = 6; equal = 3; notEqual = 4; greaterEqual = 7; greaterThanOrEqual = 7; lessEqual = 8; lessThanOrEqual = 8; between = 1 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 3 }
            $val1 = if ($null -ne $p.value1) { $p.value1 } else { $p.value }
            $val2 = $p.value2
            $cf = $range.FormatConditions.Add(1, $op, $val1, $val2)
            if ($null -ne $cf -and $p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) { $cf.Interior.Color = $bg }
            }
            if ($null -ne $cf -and $p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cf.Font.Color = $fc }
            }
        } elseif ($formatType -eq "colorScale") {
            $scaleType = if ($p.colorScaleType) { $p.colorScaleType } else { 3 }
            $range.FormatConditions.AddColorScale($scaleType)
        } elseif ($formatType -eq "dataBar") {
            $range.FormatConditions.AddDatabar() | Out-Null
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $formatType } }
    }

    "removeConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        if ($p.index) {
            $range.FormatConditions.Item([int]$p.index).Delete()
        } else {
            $range.FormatConditions.Delete()
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getConditionalFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $formats = @()
        $count = $range.FormatConditions.Count
        for ($i = 1; $i -le $count; $i++) {
            $cf = $range.FormatConditions.Item($i)
            $formats += @{ index = $i; type = $cf.Type }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; formats = $formats; count = $count } }
    }

    "addDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $typeMap = @{ list = 3; whole = 1; decimal = 2; date = 4; time = 5; textLength = 6; custom = 7 }
        $validationType = $typeMap[$p.validationType]
        if ($null -eq $validationType) { $validationType = 3 }
        $range.Validation.Delete()
        if ($p.validationType -eq "list") {
            $listFormula = if ($p.formula1) { $p.formula1 } elseif ($p.list) { ($p.list -join ",") } else { "" }
            $range.Validation.Add($validationType, 1, 1, $listFormula)
            if ($null -ne $p.showDropdown -and -not [bool]$p.showDropdown) {
                $range.Validation.InCellDropdown = $false
            } else {
                $range.Validation.InCellDropdown = $true
            }
        } else {
            $operatorMap = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 1 }
            $range.Validation.Add($validationType, 1, $op, $p.formula1, $p.formula2)
        }
        if ($p.inputTitle -or $p.inputMessage) {
            $range.Validation.InputTitle = if ($p.inputTitle) { $p.inputTitle } else { "" }
            $range.Validation.InputMessage = if ($p.inputMessage) { $p.inputMessage } else { "" }
        }
        if ($p.errorTitle -or $p.errorMessage) {
            $range.Validation.ErrorTitle = if ($p.errorTitle) { $p.errorTitle } else { "" }
            $range.Validation.ErrorMessage = if ($p.errorMessage) { $p.errorMessage } else { "" }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.validationType } }
    }

    "removeDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $range.Validation.Delete()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getDataValidations" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $validation = $range.Validation
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $validation.Type; formula1 = $validation.Formula1; formula2 = $validation.Formula2; inputTitle = $validation.InputTitle; inputMessage = $validation.InputMessage } }
    }

    "getFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cell = $sheet.Range($p.cell)
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $formulaLocal = ""
        try { $formulaLocal = $cell.FormulaLocal } catch { $formulaLocal = "" }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; formulaLocal = $formulaLocal; hasFormula = ($formula -like "=*") } }
    }

    "setArrayFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $range.FormulaArray = $p.formula
        Output-Json @{ success = $true; data = @{ range = $p.range; formula = $p.formula } }
    }

    "getCellInfo" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $numberFormat = if ($cell.NumberFormat) { $cell.NumberFormat } else { "" }
        $fontName = if ($cell.Font.Name) { $cell.Font.Name } else { "" }
        $fontSize = if ($cell.Font.Size) { $cell.Font.Size } else { 0 }
        $bold = if ($null -ne $cell.Font.Bold) { [bool]$cell.Font.Bold } else { $false }
        $bgColor = if ($null -ne $cell.Interior.Color) { $cell.Interior.Color } else { 0 }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; value = $value; formula = $formula; numberFormat = $numberFormat; font = @{ name = $fontName; size = $fontSize; bold = $bold }; backgroundColor = $bgColor } }
    }

    "refreshLinks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $links = $wb.LinkSources(1)
        if ($links) {
            for ($i = 1; $i -le $links.Length; $i++) {
                $wb.UpdateLink($links[$i - 1], 1)
            }
            Output-Json @{ success = $true; data = @{ refreshed = $links.Length } }
        } else {
            Output-Json @{ success = $true; data = @{ refreshed = 0; message = "没有外部链接" } }
        }
    }

    "consolidate" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $destRange = $sheet.Range($p.destination)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $destRange.Consolidate($p.sources, $func, $p.topRow, $p.leftColumn, $p.createLinks)
        Output-Json @{ success = $true; data = @{ destination = $p.destination; sources = $p.sources } }
    }

    "calculateSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        if ($p.all) {
            $excel.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = "all" } }
        } else {
            $sheet = $excel.ActiveSheet
            $sheet.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = $sheet.Name } }
        }
    }

    # ==================== Excel Advanced ====================
    "getExcelContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRange = $usedRange.Address(); usedRangeAddress = $usedRange.Address()
        }}
    }

    "getContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRangeAddress = $usedRange.Address(); usedRange = $usedRange.Address()
        }}
    }

    "sortRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        # 第 4 轮修：TS 传的是 column（排序范围里的第几列，数字），PS 以前读 keyColumn 并把它
        # 直接塞进 Range() —— 传数字进去必然抛（Range 只收地址/名字）。现在两种都接：
        #   keyColumn（地址，如 "B:B"）直接用；column（数字）按排序范围的左边界换算成列字母。
        $keyCol = $null
        if ($p.keyColumn) {
            $keyCol = $sheet.Range([string]$p.keyColumn)
        } elseif ($p.column) {
            $baseCol = $range.Column
            $colLetter = Convert-ColumnNumberToLetter($baseCol + [int]$p.column - 1)
            $keyCol = $sheet.Range("${colLetter}1")
        } else {
            Output-Json @{ success = $false; error = "要给 column（排序依据的列号）或 keyColumn（列地址，如 B:B）" }
            return
        }
        # order 支持 'desc'/'asc' 字符串（TS 现在发这个）与 1/2 数字（历史裸调用）
        $order = 1
        if ([string]$p.order -eq 'desc') { $order = 2 }
        elseif ([string]$p.order -match '^\d+$') { $order = [int]$p.order }
        $range.Sort($keyCol, $order)
        Output-Json @{ success = $true; data = @{ range = $p.range; order = $order } }
    }

    "autoFilter" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        if ($p.criteria) {
            $range.AutoFilter($p.field, $p.criteria)
        } else {
            $range.AutoFilter()
        }
        Output-Json @{ success = $true }
    }

    "createChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.dataRange)
        $chartType = $p.chartType
        if ($null -eq $chartType) {
            $chartTypes = @{ column = 51; column_clustered = 51; column_stacked = 52; bar = 57; bar_clustered = 57; line = 4; line_markers = 65; pie = 5; doughnut = -4120; area = 1; scatter = -4169; radar = -4151 }
            $chartType = $chartTypes[$p.chartTypeName]
        }
        if ($null -eq $chartType) { $chartType = 51 }
        $chartTypeName = if ($p.chartTypeName) { $p.chartTypeName } else { $null }
        $left = if ($p.position -and $p.position.left) { $p.position.left } elseif ($p.left) { $p.left } else { $range.Left + $range.Width + 20 }
        $top = if ($p.position -and $p.position.top) { $p.position.top } elseif ($p.top) { $p.top } else { $range.Top }
        $width = if ($p.position -and $p.position.width) { $p.position.width } else { 400 }
        $height = if ($p.position -and $p.position.height) { $p.position.height } else { 300 }
        $chartObj = $sheet.ChartObjects().Add($left, $top, $width, $height)
        $chartObj.Chart.SetSourceData($range)
        $chartObj.Chart.ChartType = $chartType
        if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
        if ($null -ne $p.showLegend) { $chartObj.Chart.HasLegend = [bool]$p.showLegend }
        if ($p.showDataLabels) { $chartObj.Chart.ApplyDataLabels() }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; chartIndex = $chartObj.Index; dataRange = $p.dataRange; chartType = $chartTypeName; position = @{ left = $left; top = $top; width = $width; height = $height } } }
    }

    "updateChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $chartObj = $null
        if ($p.chartName) {
            try { $chartObj = $sheet.ChartObjects($p.chartName) } catch {}
        }
        if ($null -eq $chartObj -and $p.chartIndex) {
            try { $chartObj = $sheet.ChartObjects().Item([int]$p.chartIndex) } catch {}
        }
        if ($null -eq $chartObj) { Output-Json @{ success = $false; error = "Chart not found" }; return }
        $updated = @()
        if ($null -ne $p.title) {
            if ($p.title -eq "") { $chartObj.Chart.HasTitle = $false }
            else { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
            $updated += "title"
        }
        if ($null -ne $p.chartType) {
            $chartObj.Chart.ChartType = $p.chartType
            $updated += "chartType"
        }
        if ($null -ne $p.showLegend) {
            $chartObj.Chart.HasLegend = [bool]$p.showLegend
            $updated += "showLegend"
        }
        if ($null -ne $p.legendPosition) {
            $legendMap = @{ bottom = -4107; top = -4160; left = -4131; right = -4152 }
            $pos = $legendMap[$p.legendPosition]
            if ($pos) { $chartObj.Chart.Legend.Position = $pos; $updated += "legendPosition" }
        }
        if ($null -ne $p.showDataLabels) {
            if ($p.showDataLabels) {
                $chartObj.Chart.ApplyDataLabels()
            } else {
                try {
                    $series = $chartObj.Chart.SeriesCollection()
                    for ($i = 1; $i -le $series.Count; $i++) {
                        try { $series.Item($i).DataLabels().Delete() } catch {}
                    }
                } catch {}
            }
            $updated += "showDataLabels"
        }
        if ($p.dataRange) {
            $chartObj.Chart.SetSourceData($sheet.Range($p.dataRange))
            $updated += "dataRange"
        }
        if ($p.colors -and $p.colors.Count -gt 0) {
            $series = $chartObj.Chart.SeriesCollection()
            for ($i = 1; $i -le $series.Count -and $i -le $p.colors.Count; $i++) {
                $color = Convert-HexColorToRgbInt($p.colors[$i - 1])
                if ($null -ne $color) { $series.Item($i).Format.Fill.ForeColor.RGB = $color }
            }
            $updated += "colors"
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; updatedProperties = $updated } }
    }

    "exportChartAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        $chartName = $p.chartName
        if ([string]::IsNullOrEmpty($chartName)) { Output-Json @{ success = $false; error = "Missing chartName" }; return }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 Excel COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $chartObj = $sheet.ChartObjects($chartName)
        # 先删掉同名旧文件：否则下面"文件存在"的校验会被上一次的产物骗过
        if (Test-Path $outputPath) { try { Remove-Item -LiteralPath $outputPath -Force } catch {} }
        $chartObj.Chart.Export($outputPath, $filterName)
        # 关键校验：Chart.Export 会【静默写出 0 字节文件且不报错】（office-mcp 的 GOTCHAS 记过同一个坑）
        if (-not (Test-Path $outputPath)) {
            Output-Json @{ success = $false; error = "Chart.Export 未抛错但文件未生成: $outputPath" }; return
        }
        $bytes = (Get-Item $outputPath).Length
        if ($bytes -le 0) {
            Output-Json @{ success = $false; error = "Chart.Export 生成了 0 字节文件: $outputPath（filter/路径可能不被 WPS 接受）" }; return
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartName; outputPath = $outputPath; format = $filterName; bytes = $bytes } }
    }

    "exportRangeAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        if ([string]::IsNullOrEmpty($p.range)) { Output-Json @{ success = $false; error = "Missing range" }; return }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $range = $sheet.Range($p.range)
        # xlScreen=1 (Appearance), xlBitmap=2 (Format)
        $tempChart = $null
        try {
            # 先删同名旧文件，避免"文件存在"的校验被上一次产物骗过
            if (Test-Path $outputPath) { try { Remove-Item -LiteralPath $outputPath -Force } catch {} }
            $range.CopyPicture(1, 2)
            $tempChart = $sheet.ChartObjects().Add(0, 0, $range.Width, $range.Height)
            $tempChart.Activate()
            $tempChart.Chart.Paste()
            $tempChart.Chart.Export($outputPath, $filterName)
            $tempChart.Delete()
            $tempChart = $null
            # 关键校验：Chart.Export 会【静默写出 0 字节文件且不报错】，不查就等于假成功
            if (-not (Test-Path $outputPath)) {
                Output-Json @{ success = $false; error = "Chart.Export 未抛错但文件未生成: $outputPath" }; return
            }
            $bytes = (Get-Item $outputPath).Length
            if ($bytes -le 0) {
                Output-Json @{ success = $false; error = "Chart.Export 生成了 0 字节文件: $outputPath（filter/路径可能不被 WPS 接受）" }; return
            }
            Output-Json @{ success = $true; data = @{ range = $p.range; outputPath = $outputPath; format = $filterName; bytes = $bytes } }
        } catch {
            # 异常清理：剪贴板冲突时回滚临时图表
            if ($null -ne $tempChart) { try { $tempChart.Delete() } catch {} }
            Output-Json @{ success = $false; error = "导出区域为图片失败: $(Get-CallError $_)" }
        }
    }

    "removeDuplicates" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $originalCount = $range.Rows.Count
        $cols = @()
        if ($p.columns -and $p.columns.Count -gt 0) {
            foreach ($col in $p.columns) {
                if ($col -is [int]) { $cols += $col }
                else {
                    $num = Convert-ColumnLetterToNumber([string]$col)
                    if ($num) { $cols += $num }
                }
            }
        }
        if ($cols.Count -eq 0) { $cols = @(1) }
        $hasHeader = if ($null -ne $p.hasHeader) { [int]([bool]$p.hasHeader) } else { 1 }
        $range.RemoveDuplicates($cols, $hasHeader)
        $remainingCount = $range.Rows.Count
        $removedCount = $originalCount - $remainingCount
        Output-Json @{ success = $true; data = @{ originalCount = $originalCount; removedCount = $removedCount; remainingCount = $remainingCount } }
    }

    "createSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $wb.Sheets.Add()
        if ($p.name) { $sheet.Name = $p.name }
        Output-Json @{ success = $true; data = @{ sheetName = $sheet.Name; sheetIndex = $sheet.Index } }
    }

    "deleteSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $name = $sheet.Name
        $excel.DisplayAlerts = $false
        $sheet.Delete()
        $excel.DisplayAlerts = $true
        Output-Json @{ success = $true; data = @{ deletedSheet = $name } }
    }

    "renameSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $oldName = $sheet.Name
        $sheet.Name = $p.newName
        Output-Json @{ success = $true; data = @{ oldName = $oldName; newName = $p.newName } }
    }

    "copySheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        if ($p.before) {
            $sheet.Copy($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Copy($null, $wb.Sheets.Item($p.after))
        } else {
            $sheet.Copy($null, $wb.Sheets.Item($wb.Sheets.Count))
        }
        Output-Json @{ success = $true; data = @{ copiedFrom = $sheet.Name } }
    }

    "getSheetList" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sheet = $wb.Sheets.Item($i)
            $sheets += @{ name = $sheet.Name; index = $i; visible = $sheet.Visible }
        }
        Output-Json @{ success = $true; data = @{ sheets = $sheets; count = $sheets.Count; activeSheet = $excel.ActiveSheet.Name } }
    }

    # ===== 第 4 期补齐（常见高频缺口：表内查找 / 图表删除 / 透视表 / 页面设置）=====
    # 逐条说明"为什么缺它会真的卡住"，不是凑数。
    # 真机支持情况见 __probeCaps 的对应项；未验的一律走 warnings，不假装成功。

    # 表详情：以前只有 getSheetList（表名 + 可见性），用户问"这张表多大 / 用到哪 /
    # 哪些行被藏了"只能瞎猜。UsedRange 一次给全。
    "getSheetInfo" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $used = $sheet.UsedRange
        $hiddenRows = 0
        try {
            for ($r = 1; $r -le $used.Rows.Count; $r++) {
                if ($used.Rows.Item($r).Hidden) { $hiddenRows++ }
            }
        } catch { }
        $hiddenCols = 0
        try {
            for ($c = 1; $c -le $used.Columns.Count; $c++) {
                if ($used.Columns.Item($c).Hidden) { $hiddenCols++ }
            }
        } catch { }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name
            usedRange = [string]$used.Address($false, $false)
            rows = $used.Rows.Count
            columns = $used.Columns.Count
            firstRow = $used.Row
            firstColumn = $used.Column
            hiddenRows = $hiddenRows
            hiddenColumns = $hiddenCols
            charts = $(try { $sheet.ChartObjects().Count } catch { 0 })
            hasData = $(try { ($used.Cells.Item(1,1).Value2 -ne $null) } catch { $false })
        } }
    }

    # 列宽/行高自适应：以前只能 set_column_width 手填 px/字符数，
    # 中文表头一换就得逐列重调 —— 这是表格排版最高频的重复劳动。
    "autoFit" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $range = if ($p.range) { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }
        # dimension 不给就两个都做 —— 用户说"自适应"时的默认预期就是全都要
        $dim = if ($p.dimension) { [string]$p.dimension } else { "both" }
        $did = @()
        if ($dim -eq "column" -or $dim -eq "both") {
            $range.EntireColumn.AutoFit() | Out-Null
            $did += "column"
        }
        if ($dim -eq "row" -or $dim -eq "both") {
            $range.EntireRow.AutoFit() | Out-Null
            $did += "row"
        }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name
            dimension = $dim
            applied = $did
            # 行高自适应对合并单元格会抛错，属已知限制；说明白好过静默失败
            note = if ($did -contains "row") { "行高自适应遇到合并单元格时会跳过那些行（WPS 限制）" } else { $null }
        } }
    }

    # 自动换行：表格里最常碰的排版动作，之前完全没有
    "setWrapText" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $range = if ($p.range) { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }
        $want = if ($null -eq $p.enabled) { $true } else { [bool]$p.enabled }
        $range.WrapText = $want
        Output-Json @{ success = $true; data = @{ range = [string]$range.Address($false,$false); wrapText = $want } }
    }

    # 表内查找：以前只有 find_replace（改），没有"找"。用户问"这个数在哪"无从下手。
    # 只读，不改任何内容。
    "findInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $what = if ($p.value) { [string]$p.value } elseif ($p.find) { [string]$p.find } else { $null }
        if (-not $what) { Output-Json @{ success = $false; error = "value/find required" }; return }
        $scope = if ($p.range) { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }
        $matchCase = if ($null -eq $p.matchCase) { $false } else { [bool]$p.matchCase }
        $wholeCell = if ($null -eq $p.wholeCell) { $false } else { [bool]$p.wholeCell }
        $found = @()
        # 逐单元格扫，不用 Range.Find —— WPS 的 Find 参数签名与 Excel 有出入，
        # 而 FindNext 在空结果时会抛异常，边界比逐格扫难处理。
        try {
            $rows = $scope.Rows.Count
            $cols = $scope.Columns.Count
            if ($rows * $cols -gt 200000) {
                Output-Json @{ success = $false; error = "查找范围过大（$rows×$cols），请用 range 参数缩小" }; return
            }
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $cell = $scope.Cells.Item($r, $c)
                    $v = $cell.Value2
                    if ($null -eq $v) { continue }
                    $s = if ($v -is [double] -or $v -is [int]) { [string]$v } else { [string]$v }
                    $hit = if ($matchCase) { $s.Contains($what) } else { $s.IndexOf($what, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
                    if (-not $hit) { continue }
                    if ($wholeCell) {
                        $equal = if ($matchCase) { $s -ceq $what } else { $s -ieq $what }
                        if (-not $equal) { continue }
                    }
                    $found += @{
                        address = [string]$cell.Address($false, $false)
                        row = $cell.Row
                        column = $cell.Column
                        value = $s
                        formula = $(try { [string]$cell.Formula } catch { $null })
                    }
                    if ($p.maxResults -and $found.Count -ge [int]$p.maxResults) { break }
                }
                if ($p.maxResults -and $found.Count -ge [int]$p.maxResults) { break }
            }
        } catch {
            Output-Json @{ success = $false; error = "查找出错: $(Get-CallError $_)" }; return
        }
        Output-Json @{ success = $true; data = @{
            keyword = $what
            count = $found.Count
            matches = $found
            scope = [string]$scope.Address($false, $false)
        } }
    }

    # 删除图表：以前只有 create_chart / update_chart，建错了删不掉 —— COM 没有撤销
    "deleteChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $total = $sheet.ChartObjects().Count
        if ($total -eq 0) {
            Output-Json @{ success = $false; error = "该工作表没有图表（ChartObjects 数量为 0）" }; return
        }
        $deleted = @()
        $idx = $null
        if ($p.index) { $idx = [int]$p.index }
        elseif ($p.name) {
            for ($i = 1; $i -le $total; $i++) {
                if ($sheet.ChartObjects().Item($i).Name -eq [string]$p.name) { $idx = $i; break }
            }
            if ($null -eq $idx) {
                $names = @()
                for ($i = 1; $i -le $total; $i++) { $names += [string]$sheet.ChartObjects().Item($i).Name }
                Output-Json @{ success = $false; error = "找不到名为 $($p.name) 的图表。现有：$($names -join ', ')" }; return
            }
        } else {
            $idx = $total  # 不指定就删最后一个（等价 UI 里"撤销插入图表"）
        }
        $obj = $sheet.ChartObjects().Item($idx)
        $deleted += [string]$obj.Name
        $obj.Delete()
        $remaining = $sheet.ChartObjects().Count
        Output-Json @{ success = $true; data = @{
            deleted = $deleted
            remaining = $remaining
            note = if ($p.index -or $p.name) { $null } else { "没指定 index/name，删的是最后一个图表" }
            warnings = @()
        } }
    }

    # 透视表：以前只有 create/update。建完看不到有哪些、源数据变了刷不了、
    # 改错了清不掉 —— 三个出口全缺。
    "getPivotTables" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        # ⚠️ 实测：WPS 12.1 的【工作簿级】Workbook.PivotTables() 会抛异常，
        #    工作表级 Sheet.PivotTables() 正常。所以先走 sheet 级；
        #    只有调用方明确要"整个工作簿"时才试工作簿级，且失败要说清是版本限制。
        $list = @()
        $warnings = @()
        $scopeNames = @()
        if ($p.sheet) {
            $scopeNames += [string]$p.sheet
        } else {
            for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $scopeNames += [string]$wb.Sheets.Item($i).Name }
        }
        foreach ($sn in $scopeNames) {
            $sh = $wb.Sheets.Item($sn)
            $src = $null
            try { $src = $sh.PivotTables() } catch { $src = $null }
            if ($null -eq $src) { continue }
            for ($i = 1; $i -le $src.Count; $i++) {
                $pt = $src.Item($i)
                $list += @{
                    name = [string]$pt.Name
                    sheet = $sn
                    tableRange1 = $(try { [string]$pt.TableRange1.Address($false,$false) } catch { $null })
                    rowFields = $(try { $pt.RowFields.Count } catch { 0 })
                    columnFields = $(try { $pt.ColumnFields.Count } catch { 0 })
                    dataFields = $(try { $pt.DataFields.Count } catch { 0 })
                }
            }
        }
        if (-not $p.sheet) {
            # 补一句实测到的限制，别让调用方以为"工作簿级接口挂了是 bug"
            $warnings = @(@{
                step = "getPivotTables"
                message = "本机 WPS 12.1 的工作簿级 PivotTables() 不支持，已改为逐个工作表扫（结果等价）。"
            })
        }
        Output-Json @{ success = $true; data = @{
            count = $list.Count
            pivotTables = $list
            scope = $(if ($p.sheet) { $p.sheet } else { "整个工作簿（逐表扫描）" })
            warnings = $warnings
        } }
    }

    "refreshPivotTables" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        # WPS 12.1 的 Workbook.PivotTables() 不通（实测），逐个工作表扫
        $refreshed = @()
        $failed = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sh = $wb.Sheets.Item($i)
            $src = $null
            try { $src = $sh.PivotTables() } catch { continue }
            for ($j = 1; $j -le $src.Count; $j++) {
                $pt = $src.Item($j)
                if ($p.name -and ([string]$pt.Name) -ne [string]$p.name) { continue }
                try {
                    $pt.RefreshTable() | Out-Null
                    $refreshed += [string]$pt.Name
                } catch {
                    $failed += "$($pt.Name): $(Get-CallError $_)"
                }
            }
        }
        $warnings = @()
        if ($p.name -and $refreshed.Count -eq 0) {
            $warnings = @(@{ step = "refreshPivotTables"; message = "整个工作簿里没有名为「$($p.name)」的透视表（先查 wps_excel_get_pivot_tables）" })
        }
        if ($failed.Count -gt 0) {
            $warnings += @(@{ step = "refreshPivotTables"; message = "有透视表刷不了：$($failed -join '；')" })
        }
        Output-Json @{ success = $true; data = @{
            scope = $(if ($p.name) { "指定" } else { "全部" })
            refreshed = $refreshed
            failedCount = $failed.Count
            warnings = $warnings
        } }
    }

    # 清透视表（保留位置与格式、只清数据与字段布局）—— 比 Delete 安全，常用
    "clearPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required（要清哪一个透视表）" }; return }
        $found = $null
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $src = $null
            try { $src = $wb.Sheets.Item($i).PivotTables() } catch { continue }
            for ($j = 1; $j -le $src.Count; $j++) {
                if (([string]$src.Item($j).Name) -eq [string]$p.name) { $found = $src.Item($j); break }
            }
            if ($found) { break }
        }
        if ($null -eq $found) {
            Output-Json @{ success = $false; error = "整个工作簿里没有名为「$($p.name)」的透视表（先查 wps_excel_get_pivot_tables）" }
            return
        }
        try {
            $found.ClearTable()
            Output-Json @{ success = $true; data = @{ cleared = [string]$found.Name; note = "ClearTable 保留透视表位置与格式，只清掉数据与字段布局" } }
        } catch {
            Output-Json @{ success = $false; error = "清透视表失败: $(Get-CallError $_)" }
        }
    }

    # 页面设置：以前只有 set_print_area。多页表打印时"每页重复表头"、
    # 横向打印、纸张大小这些高频需求全做不了
    "setSheetPageSetup" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        $ps = $sheet.PageSetup
        $applied = @()
        $failed = @()
        # 逐属性设、逐属性容错：WPS 对纸张枚举的支持与 Excel 不完全一致，
        # 一个属性失败不该让整次调用失败（那是"部分成功"的典型场景）
        if ($p.orientation) {
            try {
                $o = ([string]$p.orientation).ToLower()
                if ($o -eq "landscape") { $ps.Orientation = 2 }
                elseif ($o -eq "portrait") { $ps.Orientation = 1 }
                else { $failed += "orientation(值不认识:$($p.orientation))" }
                if ($failed.Count -eq 0 -or $failed[-1] -notlike "orientation*") { $applied += "orientation" }
            } catch { $failed += "orientation" }
        }
        if ($null -ne $p.zoom) {
            try { $ps.Zoom = [int]$p.zoom; $applied += "zoom" } catch { $failed += "zoom" }
        }
        if ($p.paperSize) {
            try { $ps.PaperSize = [int]$p.paperSize; $applied += "paperSize" } catch { $failed += "paperSize" }
        }
        if ($null -ne $p.zoomFitWidth -or $null -ne $p.zoomFitHeight) {
            try {
                $ps.Zoom = $false   # 先关缩放才能设 FitTo
                $ps.FitToPagesWide = $(if ($null -ne $p.zoomFitWidth) { [int]$p.zoomFitWidth } else { $ps.FitToPagesWide })
                $ps.FitToPagesTall = $(if ($null -ne $p.zoomFitHeight) { [int]$p.zoomFitHeight } else { $ps.FitToPagesTall })
                $applied += "fitToPages"
            } catch { $failed += "fitToPages" }
        }
        $warnings = @()
        if ($failed.Count -gt 0) {
            $warnings = @(@{ step = "setSheetPageSetup"; message = "以下属性本机 WPS 不接受：$($failed -join '/')。其余已设上。" })
        }
        Output-Json @{ success = $true; data = @{
            sheet = $sheet.Name
            applied = $applied
            failed = $failed
            warnings = $warnings
        } }
    }

    # 每页重复表头行：多页表的刚需
    "setSheetPrintTitles" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item([string]$p.sheet) } else { $excel.ActiveSheet }
        if ($p.rows) {
            $sheet.PageSetup.PrintTitleRows = [string]$p.rows
            Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; printTitleRows = [string]$p.rows } }
        } elseif ($p.columns) {
            $sheet.PageSetup.PrintTitleColumns = [string]$p.columns
            Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; printTitleColumns = [string]$p.columns } }
        } elseif ($p.clear) {
            $sheet.PageSetup.PrintTitleRows = ""
            $sheet.PageSetup.PrintTitleColumns = ""
            Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; cleared = $true } }
        } else {
            # 单引号：$1:$1 里的 $ 在单引号里是字面量。用双引号会被当成变量引用
            # "变量引用无效，':' 后面的变量名称字符无效" —— 踩过。
            Output-Json @{ success = $false; error = 'rows / columns / clear 三选一必填（例：rows = $1:$1 表示第 1 行作表头）' }
        }
    }

    # Excel 侧能力探测。与 __probeCaps（Word 侧）分开：两者的 COM 对象模型
    # 完全不同，混在一起探测的结果没法归因。
    #
    # 自建工作簿而不是用用户当前打开的表 —— 探测里有"真的改"的操作
    # （AutoFit / WrapText / PageSetup），绝不能碰用户的文档。
    # 建完必删；删不掉会在 note 里报出来，不静默留垃圾。
    "__probeExcelCaps" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $null
        try {
            $wb = $excel.Workbooks.Add()
        } catch {
            Output-Json @{ success = $false; error = "建临时工作簿失败: $(Get-CallError $_)" }
            return
        }
        $caps = [ordered]@{}
        # 造一点数据：空表上很多探测项会被静默跳过（Count=0），测出来是假的支持
        try {
            $sh = $wb.Worksheets.Item(1)
            for ($r = 1; $r -le 3; $r++) {
                for ($c = 1; $c -le 3; $c++) {
                    $sh.Cells.Item($r, $c).Value2 = "v${r}_${c}"
                }
            }
        } catch { }

        # ---- 第 4 期新增接口 ----
        $caps["sheet.usedRange"]    = { $u = $wb.Worksheets.Item(1).UsedRange
                                          if ($u.Rows.Count -lt 1) { throw "UsedRange 为空，探测无意义" } }
        $caps["sheet.autoFitColumn"] = { $col = $wb.Worksheets.Item(1).Columns.Item(1); $w0 = $col.ColumnWidth
                                          $col.EntireColumn.AutoFit() | Out-Null
                                          try { $col.ColumnWidth = $w0 } catch { } }
        $caps["sheet.autoFitRow"]    = { $row = $wb.Worksheets.Item(1).Rows.Item(1); $h0 = $row.RowHeight
                                          $row.EntireRow.AutoFit() | Out-Null
                                          try { $row.RowHeight = $h0 } catch { } }
        $caps["sheet.wrapText"]      = { $c = $wb.Worksheets.Item(1).Range("A1"); $o = $c.WrapText
                                          try { $c.WrapText = $true } finally { $c.WrapText = $o } }
        $caps["sheet.findRead"]      = { $v = $wb.Worksheets.Item(1).Range("A1").Value2
                                          if ($null -eq $v) { throw "A1 没值，探测无意义" } }
        $caps["sheet.chartObjects"]  = { $n = $wb.Worksheets.Item(1).ChartObjects().Count
                                          if ($n -ne 0) { throw "新建工作簿不该有图表，实际 $n" } }
        $caps["sheet.pivotTables"]   = { $n = $wb.PivotTables().Count
                                          if ($n -ne 0) { throw "新建工作簿不该有透视表，实际 $n" } }
        # 工作簿级 PivotTables() 在 WPS 12.1 上抛异常（实测，见 probe 输出），
        # 所以真正常用的是【工作表级】Sheet.PivotTables()。两条路分开探，别混。
        $caps["sheet.pivotTablesOnSheet"] = { $n = $wb.Worksheets.Item(1).PivotTables().Count
                                          if ($n -ne 0) { throw "新建工作簿不该有透视表，实际 $n" } }
        $caps["sheet.pageSetup"]     = { $ps = $wb.Worksheets.Item(1).PageSetup; $o = $ps.Orientation
                                          try { $ps.Orientation = 2 } finally { $ps.Orientation = $o } }
        $caps["sheet.pageSetupZoom"] = { $ps = $wb.Worksheets.Item(1).PageSetup; $o = $ps.Zoom
                                          try { $ps.Zoom = 80 } finally { $ps.Zoom = $o } }
        $caps["sheet.printTitles"]   = { $ps = $wb.Worksheets.Item(1).PageSetup
                                          $ps.PrintTitleRows = '$1:$1'
                                          $got = $ps.PrintTitleRows
                                          if ($got -notmatch '1') { throw "设置后读回来不是第1行（实际=$got）" }
                                          $ps.PrintTitleRows = '' }

        $supported = 0; $unsupported = @(); $leaked = @()
        foreach ($k in $caps.Keys) {
            try {
                & $caps[$k] | Out-Null
                $supported++
            } catch {
                $unsupported += $k
            }
        }
        $note = $null
        try { $wb.Close($false) } catch { $leaked += "临时工作簿删不掉：$(Get-CallError $_)" }
        Output-Json @{ success = $true; data = @{
            total = $caps.Count
            supported = $supported
            unsupported = $unsupported
            leaked = $leaked
            note = $(if ($leaked.Count -gt 0) { $leaked -join '；' } else { $null })
            warning = "探测在临时工作簿上跑（改完改回），不会碰你的文档。"
        } }
    }

    "switchSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $wb.Sheets.Item($p.sheet)
        $sheet.Activate()
        Output-Json @{ success = $true; data = @{ activeSheet = $sheet.Name } }
    }

    "moveSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        # 第 4 轮修：TS 传的是 position（0 基目标位置），PS 只认 before/after ——
        # 两个都没给时【什么都不做还返回 success】，用户以为移动了。
        # 现在三条路都走通：position 优先，其次 before/after。
        if ($null -ne $p.position) {
            $target = [int]$p.position + 1
            if ($target -lt 1) { $target = 1 }
            if ($target -gt $wb.Sheets.Count) { $target = $wb.Sheets.Count }
            $sheet.Move($wb.Sheets.Item($target))
        } elseif ($p.before) {
            $sheet.Move($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Move($null, $wb.Sheets.Item($p.after))
        } else {
            Output-Json @{ success = $false; error = "要给 position（0 基目标位置），或 before/after（工作表名）。两个都没给就没法移动。" }
            return
        }
        Output-Json @{ success = $true; data = @{ movedSheet = $sheet.Name; position = $sheet.Index - 1 } }
    }

    "mergeCells" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $across = if ($null -ne $p.across) { [bool]$p.across } else { $false }
        $range.Merge($across)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "unmergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $range.UnMerge()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setColumnWidth" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $col = $p.column
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $sheet.Range("${col}:${col}").ColumnWidth = $p.width
        Output-Json @{ success = $true; data = @{ column = $col; width = $p.width } }
    }

    "setRowHeight" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $sheet.Range("$($p.row):$($p.row)").RowHeight = $p.height
        Output-Json @{ success = $true; data = @{ row = $p.row; height = $p.height } }
    }

    "autoFitColumn" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if ($p.range) {
            $sheet.Range($p.range).Columns.AutoFit()
        } elseif ($p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").AutoFit()
        } else {
            $sheet.UsedRange.Columns.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "列宽已自动调整" } }
    }

    "autoFitRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if ($p.range) {
            $sheet.Range($p.range).Rows.AutoFit()
        } elseif ($p.row) {
            $sheet.Range("$($p.row):$($p.row)").AutoFit()
        } else {
            $sheet.UsedRange.Rows.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "行高已自动调整" } }
    }

    "autoFitAll" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $range.Columns.AutoFit()
        $range.Rows.AutoFit()
        Output-Json @{ success = $true; data = @{ message = "列宽行高已自动调整" } }
    }

    "setNumberFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        if ($p.format) { $range.NumberFormat = $p.format }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.format } }
    }

    "wrapText" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $wrap = if ($null -ne $p.wrap) { [bool]$p.wrap } else { $true }
        $range.WrapText = $wrap
        Output-Json @{ success = $true; data = @{ range = $p.range; wrapText = $wrap } }
    }

    "setPrintArea" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if ($p.range) { $sheet.PageSetup.PrintArea = $p.range } else { $sheet.PageSetup.PrintArea = "" }
        Output-Json @{ success = $true; data = @{ printArea = if ($p.range) { $p.range } else { "cleared" } } }
    }

    "getSelection" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sel = $excel.Selection
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "No selection" }; return }
        $addr = $sel.Address()
        Output-Json @{ success = $true; data = @{ address = $addr; rows = $sel.Rows.Count; columns = $sel.Columns.Count } }
    }

    "clearRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $clearType = if ($p.type) { $p.type } else { "all" }
        if ($clearType -eq "contents") { $range.ClearContents() }
        elseif ($clearType -eq "formats") { $range.ClearFormats() }
        elseif ($clearType -eq "comments") { $range.ClearComments() }
        else { $range.Clear() }
        Output-Json @{ success = $true; data = @{ range = $p.range; clearType = $clearType } }
    }

    "insertRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; return }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Insert()
        Output-Json @{ success = $true; data = @{ insertedAt = $startRow; count = $count } }
    }

    "insertColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; return }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Insert()
        }
        Output-Json @{ success = $true; data = @{ insertedAt = $col; count = $count } }
    }

    "deleteRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; return }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Delete()
        Output-Json @{ success = $true; data = @{ deletedFrom = $startRow; count = $count } }
    }

    "deleteColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; return }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Delete()
        }
        Output-Json @{ success = $true; data = @{ deletedFrom = $col; count = $count } }
    }

    "hideRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; return }
        foreach ($r in $rows) { $sheet.Range("${r}:${r}").Hidden = $true }
        Output-Json @{ success = $true; data = @{ hiddenRows = $rows } }
    }

    "hideColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; return }
        $hidden = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").Hidden = $true
            $hidden += $col
        }
        Output-Json @{ success = $true; data = @{ hiddenColumns = $hidden } }
    }

    "showRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; return }
        foreach ($r in $rows) { $sheet.Range("${r}:${r}").Hidden = $false }
        Output-Json @{ success = $true; data = @{ shownRows = $rows } }
    }

    "showColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; return }
        $shown = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").Hidden = $false
            $shown += $col
        }
        Output-Json @{ success = $true; data = @{ shownColumns = $shown } }
    }

    "groupRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if (-not $p.startRow -or -not $p.endRow) { Output-Json @{ success = $false; error = "startRow/endRow required" }; return }
        $sheet.Range("$($p.startRow):$($p.endRow)").Group()
        Output-Json @{ success = $true; data = @{ grouped = "$($p.startRow):$($p.endRow)" } }
    }

    "groupColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if (-not $p.startColumn -or -not $p.endColumn) { Output-Json @{ success = $false; error = "startColumn/endColumn required" }; return }
        $startCol = $p.startColumn
        $endCol = $p.endColumn
        if ($startCol -is [int]) { $startCol = Convert-ColumnNumberToLetter([int]$startCol) }
        if ($endCol -is [int]) { $endCol = Convert-ColumnNumberToLetter([int]$endCol) }
        $sheet.Range("${startCol}:${endCol}").Group()
        Output-Json @{ success = $true; data = @{ grouped = "${startCol}:${endCol}" } }
    }

    "freezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        # 第 4 轮修三处：
        #   ① TS 传 freeze=false 表示"取消冻结"，以前照样冻上；
        #   ② 只给 row 或只给 column 时以前啥都不选（$p.row -and $p.column 才动），冻的还是上次的位置；
        #   ③ 三个都没给时 $sheet.Range($null) 直接抛 COM 错。
        $want = if ($null -ne $p.freeze) { [bool]$p.freeze } else { $true }
        if (-not $want) {
            $excel.ActiveWindow.FreezePanes = $false
            Output-Json @{ success = $true; data = @{ frozen = $false; message = "窗格冻结已取消" } }
            return
        }
        # 冻结 N 行选 A(N+1)、冻结 M 列选 (M+1)1、两个都要选 (M+1)(N+1) —— 与 Excel UI 的语义一致
        if ($p.cell) {
            $sheet.Range([string]$p.cell).Select()
        } elseif ($p.row -or $p.column) {
            $rowIdx = if ($p.row) { [int]$p.row + 1 } else { 1 }
            $colIdx = if ($p.column) { [int]$p.column + 1 } else { 1 }
            $colLetter = Convert-ColumnNumberToLetter($colIdx)
            $sheet.Range("$colLetter$rowIdx").Select()
        }
        # 三个都没给：不移动选区，就地冻结当前活动单元格的上方/左侧（Excel 的默认行为）
        $excel.ActiveWindow.FreezePanes = $true
        Output-Json @{ success = $true; data = @{ frozen = $true; message = "窗格已冻结" } }
    }

    "unfreezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $excel.ActiveWindow.FreezePanes = $false
        Output-Json @{ success = $true; data = @{ message = "窗格冻结已取消" } }
    }

    "replaceInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        $replaced = $searchRange.Replace($p.searchText, $p.replaceText, $lookAt)
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; replaceText = $p.replaceText; success = $replaced } }
    }

    "copyRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $range.Copy()
        Output-Json @{ success = $true; data = @{ range = $p.range; message = "已复制到剪贴板" } }
    }

    "pasteRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $destRange = $sheet.Range($p.destination)
        if ($p.pasteType -eq "values") {
            $destRange.PasteSpecial(-4163)
        } elseif ($p.pasteType -eq "formats") {
            $destRange.PasteSpecial(-4122)
        } elseif ($p.pasteType -eq "formulas") {
            $destRange.PasteSpecial(-4123)
        } else {
            $sheet.Paste($destRange)
        }
        Output-Json @{ success = $true; data = @{ destination = $p.destination } }
    }

    "fillSeries" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $startCell = $range.Cells.Item(1, 1)
        $startCell.Value2 = if ($null -ne $p.startValue) { $p.startValue } else { 1 }
        $typeMap = @{ linear = 0; growth = 1; date = 2; autoFill = 3 }
        $fillType = $typeMap[$p.type]
        if ($null -eq $fillType) { $fillType = 0 }
        $step = if ($null -ne $p.step) { $p.step } else { 1 }
        $range.DataSeries($null, -4132, $fillType, $step)
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.type } }
    }

    "transpose" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $sourceRange = $sheet.Range($p.sourceRange)
        $destCell = if ($p.destinationCell) { $p.destinationCell } elseif ($p.targetCell) { $p.targetCell } else { $null }
        if ($null -eq $destCell) { Output-Json @{ success = $false; error = "destinationCell/targetCell required" }; return }
        $destRange = $sheet.Range($destCell)
        $sourceRange.Copy()
        $destRange.PasteSpecial(-4163, -4142, $false, $true)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.sourceRange; destination = $destCell } }
    }

    "textToColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $delimiter = if ($p.delimiter) { $p.delimiter } else { "," }
        $tab = $false; $semicolon = $false; $comma = $false; $space = $false; $other = $false; $otherChar = $null
        if ($delimiter -eq "`t") { $tab = $true }
        elseif ($delimiter -eq ";") { $semicolon = $true }
        elseif ($delimiter -eq ",") { $comma = $true }
        elseif ($delimiter -eq " ") { $space = $true }
        else { $other = $true; $otherChar = $delimiter }
        $range.TextToColumns($null, 1, 1, $false, $tab, $semicolon, $comma, $space, $other, $otherChar)
        Output-Json @{ success = $true; data = @{ range = $p.range; delimiter = $delimiter } }
    }

    "subtotal" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $totalCols = $p.totalColumns
        if ($null -eq $totalCols) { $totalCols = $p.totalColumn }
        if ($totalCols -isnot [System.Array]) { $totalCols = @($totalCols) }
        $replace = if ($null -ne $p.replace) { [bool]$p.replace } else { $true }
        $range.Subtotal([int]$p.groupBy, $func, $totalCols, $replace, $false, $true)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "createNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $addr = $range.Address()
        $wb.Names.Add($p.name, "=" + $sheet.Name + "!" + $addr)
        Output-Json @{ success = $true; data = @{ name = $p.name; range = $p.range } }
    }

    "deleteNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $wb.Names.Item($p.name).Delete()
        Output-Json @{ success = $true; data = @{ deletedName = $p.name } }
    }

    "getNamedRanges" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $names = @()
        for ($i = 1; $i -le $wb.Names.Count; $i++) {
            $n = $wb.Names.Item($i)
            $names += @{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible }
        }
        Output-Json @{ success = $true; data = @{ names = $names; count = $names.Count } }
    }

    "addCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        if ($null -eq $sheet) { Output-Json @{ success = $false; error = "No sheet" }; return }
        $cellAddr = [string]$p.cell
        if (-not $cellAddr) { Output-Json @{ success = $false; error = "cell required" }; return }
        # 第 4 轮修：TS 的 schema 里这个参数叫 comment（历史 PS 只认 text）——
        # 曾经 $p.text 为 $null → AddComment($null) 造出一条空批注，还返回 success。
        # 两个键都认，comment 优先。
        $text = if ($p.comment) { $p.comment } else { $p.text }
        if (-not $text) { Output-Json @{ success = $false; error = "comment required（批注内容不能为空）" }; return }
        $target = $sheet.Range($cellAddr)
        # 同一格已有批注时 AddComment 会抛（WPS 与 Excel 都不允许一格两条）→ 先删后加，等价于覆盖
        try { if ($null -ne $target.Comment) { $target.Comment.Delete() } } catch {}
        $target.AddComment([string]$text) | Out-Null
        if ($p.visible) { try { $target.Comment.Visible = $true } catch {} }
        Output-Json @{ success = $true; data = @{ cell = $cellAddr; sheet = $sheet.Name; comment = [string]$text } }
    }

    "deleteCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        Output-Json @{ success = $true; data = @{ cell = $p.cell } }
    }

    "getCellComments" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $comments = @()
        for ($i = 1; $i -le $sheet.Comments.Count; $i++) {
            $c = $sheet.Comments.Item($i)
            $addr = $c.Parent.Address()
            $comments += @{ cell = ($addr -replace "\$", ""); text = $c.Text(); author = if ($c.Author) { $c.Author } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "protectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Protect($password, $p.drawingObjects, $p.contents, $p.scenarios)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $true } }
    }

    "unprotectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Unprotect($password)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $false } }
    }

    "protectWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $password = if ($p.password) { $p.password } else { "" }
        $structure = if ($null -ne $p.structure) { [bool]$p.structure } else { $true }
        $wb.Protect($password, $structure, $p.windows)
        Output-Json @{ success = $true; data = @{ workbook = $wb.Name; protected = $true } }
    }

    "insertExcelImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $left = if ($null -ne $p.left) { $p.left } else { 100 }
        $top = if ($null -ne $p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $sheet.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "setHyperlink" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.cell)
        $address = if ($p.address) { $p.address } else { "" }
        $subAddress = if ($p.subAddress) { $p.subAddress } else { "" }
        $screenTip = if ($p.screenTip) { $p.screenTip } else { "" }
        $textToDisplay = if ($p.textToDisplay) { $p.textToDisplay } else { "" }
        $sheet.Hyperlinks.Add($range, $address, $subAddress, $screenTip, $textToDisplay)
        Output-Json @{ success = $true; data = @{ cell = $p.cell; address = $address } }
    }

    "lockCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $(Resolve-Sheet $excel $p)
        $range = $sheet.Range($p.range)
        $locked = if ($null -ne $p.locked) { [bool]$p.locked } else { $true }
        $range.Locked = $locked
        Output-Json @{ success = $true; data = @{ range = $p.range; locked = $locked } }
    }

    "openWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Open($p.path, $p.updateLinks, $p.readOnly)
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    "getOpenWorkbooks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        # 第 1 期第 2 章第 4 轮（A9）：空列表【不能直接信】。
        # quit 之后 WPS 要重新注册到 ROT（Running Object Table），那几秒里
        # GetActiveObject 要么取不到（于是 Get-WpsWord 新建一个空的实例），
        # 要么取到一个刚重启、文档还没恢复的实例 —— 两种都表现为"列表 0 份"，
        # 可文档其实还在（真机 A9 实测：连续 4 次 × 1.5s 读到 0 份，同一时刻按名读内容 6/6 可读）。
        # 后果是模型据此判断"文档没了"，会误判。
        #
        # 所以：读到空列表时，重新【只从 ROT 取】再取一次，最多 3 次 × 400ms，取到非空就停。
        # 只从 ROT 取、不新建 —— 新建的实例必然是空的，重试没有意义。
        #
        # 沙箱实例不做这个重试：它本来就是 New-Object 出来的，"空"是常态；
        # 而且重试时从 ROT 取到的是【用户】实例，拿用户的列表回答沙箱的查询是错的。
        #
        # 顺带：每个工作簿的属性单独 try/catch —— 以前一个取属性抛错，整份列表就没了。
        $workbooks = @()
        $suspicious = $false
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $workbooks = @()
            $total = 0
            try { $total = $excel.Workbooks.Count } catch {}
            for ($i = 1; $i -le $total; $i++) {
                $wb = $null
                try { $wb = $excel.Workbooks.Item($i) } catch { continue }
                if ($null -eq $wb) { continue }
                $nm = ""; $fp = ""; $extra = 0; $act = $false
                try { $nm = [string]$wb.Name } catch {}
                try { $fp = [string]$wb.FullName } catch {}
                try { $extra = $wb.Sheets.Count } catch {}
                try { $act = ($nm -eq $(Resolve-Book $excel $p).Name) } catch {}
                $workbooks += @{ name = $nm; path = $fp; sheets = $extra; active = $act }
            }
            if ($workbooks.Count -gt 0) { break }
            $suspicious = (-not (Test-SandboxInstance))
            if (-not $suspicious -or $attempt -ge 3) { break }
            $rot = Get-WpsExcelFromRot
            if ($null -ne $rot) { $excel = $rot }
            Start-Sleep -Milliseconds 400
        }
        $data = @{ workbooks = $workbooks; count = $workbooks.Count }
        if ($suspicious -and $workbooks.Count -eq 0) {
            # 还是 0 份 —— 说清楚，别让模型当成"文档没了"。
            $data.warnings = @(@{
                step = "getOpenWorkbooks"
                message = "读到 0 份工作簿。如果刚执行过 quit/close，这可能是 WPS 重新注册到 ROT 的那几秒窗口（实测可持续数秒），工作簿可能还在 —— 别据此判断工作簿没了，隔一会儿再读一次，或按名字直接读内容核实。"
            })
        }
        Output-Json @{ success = $true; data = $data }
    }
    "switchWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Item($(if ($p.name) { $p.name } else { $p.index }))
        $wb.Activate()
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName } }
    }

    "closeWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = if ($p.name) { $excel.Workbooks.Item($p.name) } else { $(Resolve-Book $excel $p) }
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No workbook" }; return }
        $name = $wb.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        # 未保存工作簿的护栏（第 3 轮补，照 Word 侧 closeDocument 的同一套设计答案）：
        # 从未存过盘的工作簿 Close($true) 会弹「另存为」模态框，COM 是 STA ——
        # 弹框会把桥通道卡死（实测卡 120 秒，后续全部调用超时）。
        # 所以这里不猜、不硬关：明确报错让调用方先 save_as，或显式 saveChanges=false。
        if ($saveChanges) {
            $hasPath = $false
            try { $hasPath = -not [string]::IsNullOrEmpty([string]$wb.Path) } catch {}
            if (-not $hasPath) {
                Output-Json @{ success = $false; error = "[勿重试·先另存为] 工作簿「$name」从未保存过，直接关闭会弹出「另存为」对话框卡住通道（120 秒）。请先另存为（wps_excel_save_as），或显式传 save=false 放弃保存" }
                return
            }
        }
        $wb.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "createWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Add()
        if ($p.name) { $wb.SaveAs($p.name) }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    # ==================== Word ====================
    "getActiveDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        Output-Json @{ success = $true; data = @{
            name = $doc.Name; path = $doc.FullName
            paragraphCount = $doc.Paragraphs.Count; wordCount = $doc.Words.Count; characterCount = $doc.Characters.Count
        }}
    }

    "generateTOC" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $position = if ($p.position) { $p.position } else { "start" }
        $levels = if ($p.levels) { [int]$p.levels } else { 3 }
        $includePageNumbers = if ($null -ne $p.includePageNumbers) { [bool]$p.includePageNumbers } else { $true }
        $range = if ($position -eq "cursor") { $word.Selection.Range } else { $doc.Range(0, 0) }
        $doc.TablesOfContents.Add($range, $true, 1, $levels, $false, "", $true, $true, $includePageNumbers, $true)
        Output-Json @{ success = $true; data = @{ levels = $levels; position = $position; includePageNumbers = $includePageNumbers } }
    }

    "getDocumentText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = $doc.Content.Text
        if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) + "...(truncated)" }
        Output-Json @{ success = $true; data = @{ text = $text; length = $doc.Content.Text.Length } }
    }

    "getOpenDocuments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        # 第 1 期第 2 章第 4 轮（A9）：空列表【不能直接信】。
        # quit 之后 WPS 要重新注册到 ROT（Running Object Table），那几秒里
        # GetActiveObject 要么取不到（于是 Get-WpsWord 新建一个空的实例），
        # 要么取到一个刚重启、文档还没恢复的实例 —— 两种都表现为"列表 0 份"，
        # 可文档其实还在（真机 A9 实测：连续 4 次 × 1.5s 读到 0 份，同一时刻按名读内容 6/6 可读）。
        # 后果是模型据此判断"文档没了"，会误判。
        #
        # 所以：读到空列表时，重新【只从 ROT 取】再取一次，最多 3 次 × 400ms，取到非空就停。
        # 只从 ROT 取、不新建 —— 新建的实例必然是空的，重试没有意义。
        #
        # 沙箱实例不做这个重试：它本来就是 New-Object 出来的，"空"是常态；
        # 而且重试时从 ROT 取到的是【用户】实例，拿用户的列表回答沙箱的查询是错的。
        #
        # 顺带：每个文档的属性单独 try/catch —— 以前一个取属性抛错，整份列表就没了。
        $documents = @()
        $suspicious = $false
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $documents = @()
            $total = 0
            try { $total = $word.Documents.Count } catch {}
            for ($i = 1; $i -le $total; $i++) {
                $docItem = $null
                try { $docItem = $word.Documents.Item($i) } catch { continue }
                if ($null -eq $docItem) { continue }
                $nm = ""; $fp = ""; $extra = 0; $act = $false
                try { $nm = [string]$docItem.Name } catch {}
                try { $fp = [string]$docItem.FullName } catch {}
                try { $extra = $docItem.Paragraphs.Count } catch {}
                try { $act = ((Resolve-Doc $word $p).Name -eq $nm) } catch {}
                $documents += @{ name = $nm; path = $fp; paragraphs = $extra; active = $act }
            }
            if ($documents.Count -gt 0) { break }
            $suspicious = (-not (Test-SandboxInstance))
            if (-not $suspicious -or $attempt -ge 3) { break }
            $rot = Get-WpsWordFromRot
            if ($null -ne $rot) { $word = $rot }
            Start-Sleep -Milliseconds 400
        }
        $data = @{ documents = $documents; count = $documents.Count }
        if ($suspicious -and $documents.Count -eq 0) {
            # 还是 0 份 —— 说清楚，别让模型当成"文档没了"。
            $data.warnings = @(@{
                step = "getOpenDocuments"
                message = "读到 0 份文档。如果刚执行过 quit/close，这可能是 WPS 重新注册到 ROT 的那几秒窗口（实测可持续数秒），文档可能还在 —— 别据此判断文档没了，隔一会儿再读一次，或按名字直接读内容核实。"
            })
        }
        Output-Json @{ success = $true; data = $data }
    }
    "switchDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $docItem = $word.Documents.Item($(if ($p.name) { $p.name } else { $p.index }))
        $docItem.Activate()
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName } }
    }

    "openDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; return }
        $docItem = $word.Documents.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count } }
    }

    "closeDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $docItem = if ($p.name) { $word.Documents.Item($p.name) } else { $(Resolve-Doc $word $p) }
        if ($null -eq $docItem) { Output-Json @{ success = $false; error = "No document" }; return }
        $name = $docItem.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        # 未保存文档的【设计答案】：从未存过盘的文档 Close($true) 会弹「另存为」模态框，
        # 而 COM 是 STA —— 弹框会把桥通道卡死（后续全部调用超时）。
        # 所以这里不猜、不硬关：明确报错让调用方先 save_as，或显式 saveChanges=false。
        if ($saveChanges) {
            $hasPath = $false
            try { $hasPath = -not [string]::IsNullOrEmpty([string]$docItem.Path) } catch {}
            if (-not $hasPath) {
                # 第 1 期第 2 章第 4 轮（B3）：提示里原来写的是 saveChanges=false —— 那是【PS 动作】的
                # 参数名，不是【工具】的。wps_word_close_document 的公开 schema 参数叫 save；
                # 照提示传 saveChanges 的话工具收不到 → save 仍是 undefined → TS 兜成 saveChanges=true
                # → 护栏照样触发 → 用户卡死在同一个错误里出不来。
                # 实测（word-close-param-probe.js）：照提示传 saveChanges=false 关不掉，
                # 传 save=false 一次就关掉了。所以提示必须写【工具能收的那个名字】。
                Output-Json @{ success = $false; error = "文档「$name」从未保存过，直接关闭会弹出「另存为」对话框卡住通道。请先另存为（wps_word_save_as），或显式传 save=false 放弃保存" }
                return
            }
        }
        $docItem.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    # ==================== Word 深水区补齐（第 5 期） ====================
    # 补的是"读侧太弱、模型只能盲写"那批：统计/批注/书签/表格都只能写不能读，
    # 加上 create/close_document（PS 动作早有，一直没有 TS 工具暴露）。
    # 真机支持度由 __probeWordCaps 逐个验（在临时文档上探，探完删，不碰用户文档）。

    "getDocumentStats" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $warnings = @()
        $data = [ordered]@{ name = [string]$doc.Name; path = [string]$doc.FullName }
        $data.paragraphs = 0
        try { $data.paragraphs = [int]$doc.Paragraphs.Count } catch { $warnings += "段落数读取失败" }
        $data.words = 0
        try { $data.words = [int]$doc.Words.Count } catch { $warnings += "字数读取失败" }
        $data.characters = 0
        try { $data.characters = [int]$doc.Characters.Count } catch { $warnings += "字符数读取失败" }
        $data.sections = 0
        try { $data.sections = [int]$doc.Sections.Count } catch { $warnings += "节数读取失败" }
        $data.pages = 0
        # 页数两条路都试：ComputeStatistics 是正统接口，WPS 12.1 个别版本不认；
        # Range.Information(4)=wdActiveEndPageNumber 是备用口径。
        try { $data.pages = [int]$doc.ComputeStatistics(2) } catch {
            try { $data.pages = [int]$doc.Range.Information(4) } catch { $warnings += "页数读取失败（ComputeStatistics 与 Information 均不支持）" }
        }
        # 长文档常见的"结构规模"指标，顺手一并给出（拿不到就跳过，不算失败）
        foreach ($pair in @(@("lines", "Lines"), @("tables", "Tables"), @("inlineShapes", "InlineShapes"), @("footnotes", "Footnotes"), @("toc", "TablesOfContents"), @("tof", "TablesOfFigures"))) {
            try { $data[$pair[0]] = [int]$doc.($pair[1]).Count } catch {}
        }
        if ($warnings.Count) { $data.warnings = $warnings }
        Output-Json @{ success = $true; data = $data }
    }

    "getComments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $list = @()
        $n = 0
        try { $n = [int]$doc.Comments.Count } catch {}
        for ($i = 1; $i -le $n; $i++) {
            $c = $null
            try { $c = $doc.Comments.Item($i) } catch { continue }
            $e = [ordered]@{ index = $i }
            try { $e.author = [string]$c.Author } catch { $e.author = "" }
            try { $e.date = [string]$c.Date } catch { $e.date = "" }
            try { $e.text = ([string]$c.Range.Text) } catch { $e.text = "" }
            if ($e.text.Length -gt 200) { $e.text = $e.text.Substring(0, 200) + "…" }
            try { $e.reference = ([string]$c.Reference.Text) } catch { $e.reference = "" }
            if ($e.reference.Length -gt 100) { $e.reference = $e.reference.Substring(0, 100) + "…" }
            $list += $e
        }
        Output-Json @{ success = $true; data = @{ comments = $list; count = $list.Count } }
    }

    "deleteComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $n = 0
        try { $n = [int]$doc.Comments.Count } catch {}
        if ($n -le 0) { Output-Json @{ success = $false; error = "文档里没有批注" }; return }
        $target = $null
        if ($null -ne $p.index) {
            $idx = [int]$p.index
            if ($idx -lt 1 -or $idx -gt $n) { Output-Json @{ success = $false; error = "批注序号 $idx 不存在（共 $n 条，序号从 1 开始）" }; return }
            try { $target = $doc.Comments.Item($idx) } catch { Output-Json @{ success = $false; error = "取不到序号 $idx 的批注" }; return }
        } elseif ($p.text) {
            for ($i = 1; $i -le $n; $i++) {
                $c = $null
                try { $c = $doc.Comments.Item($i) } catch { continue }
                $t = ""
                try { $t = [string]$c.Range.Text } catch {}
                if ($t -like "*$($p.text)*") { $target = $c; break }
            }
            if ($null -eq $target) { Output-Json @{ success = $false; error = "没有找到包含「$($p.text)」的批注" }; return }
        } else {
            Output-Json @{ success = $false; error = "需要 index（序号）或 text（批注内容片段）之一" }; return
        }
        $before = ""
        try { $before = [string]$target.Range.Text } catch {}
        $target.Delete()
        $n1 = 0
        try { $n1 = [int]$doc.Comments.Count } catch {}
        Output-Json @{ success = $true; data = @{ deleted = $before; remaining = $n1 } }
    }

    "getBookmarks" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $list = @()
        foreach ($b in @($doc.Bookmarks)) {
            $e = [ordered]@{ name = [string]$b.Name }
            try { $e.start = [int]$b.Start } catch { $e.start = -1 }
            try { $e.text = [string]$b.Range.Text } catch { $e.text = "" }
            if ($e.text.Length -gt 100) { $e.text = $e.text.Substring(0, 100) + "…" }
            $list += $e
        }
        Output-Json @{ success = $true; data = @{ bookmarks = $list; count = $list.Count } }
    }

    "getTables" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $list = @()
        for ($i = 1; $i -le $doc.Tables.Count; $i++) {
            $t = $doc.Tables.Item($i)
            $e = [ordered]@{ index = $i; rows = [int]$t.Rows.Count; columns = [int]$t.Columns.Count }
            # 预览取第一行：够辨认是哪张表，又不把整表塞进上下文
            $firstRow = ""
            try {
                $parts = @()
                for ($c = 1; $c -le $t.Columns.Count; $c++) {
                    $cell = ""
                    try { $cell = ([string]$t.Cell(1, $c).Range.Text) -replace "[`r`n`a]", "" } catch {}
                    $parts += $cell
                }
                $firstRow = $parts -join " | "
            } catch {}
            if ($firstRow.Length -gt 120) { $firstRow = $firstRow.Substring(0, 120) + "…" }
            $e.firstRow = $firstRow
            $list += $e
        }
        Output-Json @{ success = $true; data = @{ tables = $list; count = $doc.Tables.Count } }
    }

    "getTableData" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        $idx = [int](Get-IntParam $p "tableIndex" 1)
        if ($idx -lt 1 -or $idx -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "表格序号 $idx 不存在（共 $($doc.Tables.Count) 张）" }; return }
        $t = $doc.Tables.Item($idx)
        $rows = [int]$t.Rows.Count
        $cols = [int]$t.Columns.Count
        # 上限保护：整表文本可能非常大（宽表 + 长段落），超了截断并在 warnings 里说清
        $maxRows = 100; $maxCols = 20
        $warnings = @()
        $outRows = $rows; $outCols = $cols
        if ($rows -gt $maxRows) { $outRows = $maxRows; $warnings += "表格共 $rows 行，只返回前 $maxRows 行" }
        if ($cols -gt $maxCols) { $outCols = $maxCols; $warnings += "表格共 $cols 列，只返回前 $maxCols 列" }
        $grid = @()
        for ($r = 1; $r -le $outRows; $r++) {
            $row = @()
            for ($c = 1; $c -le $outCols; $c++) {
                $cell = ""
                try { $cell = ([string]$t.Cell($r, $c).Range.Text) -replace "[`r`n`a]", "" } catch {}
                $row += $cell
            }
            $grid += ,$row
        }
        $data = [ordered]@{ tableIndex = $idx; rows = $rows; columns = $cols; returnedRows = $outRows; returnedColumns = $outCols; data = $grid }
        if ($warnings.Count) { $data.warnings = $warnings }
        Output-Json @{ success = $true; data = $data }
    }

    "insertHyperlink" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.url) { Output-Json @{ success = $false; error = "url required" }; return }
        $url = [string]$p.url
        $position = if ($p.position) { [string]$p.position } else { "cursor" }
        $range = $null
        switch ($position) {
            "start" { $range = $doc.Range(0, 0) }
            "end"   { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1) }
            default { $range = $word.Selection.Range }
        }
        # 显示文字不传就用 url 本身（Hyperlinks.Add 的 TextToDisplay 传空会插个空链）
        $display = if ($p.text) { [string]$p.text } else { $url }
        $h = $null
        try { $h = $doc.Hyperlinks.Add($range, $url, $null, $null, $display) } catch {}
        if ($null -eq $h) {
            # WPS 个别版本不认 5 参签名（报"值不在预期的范围内"，实测 12.1 命中）——
            # 退回 2 参调用，并把显示文字先写进锚点（老实现就是这么干的，真机验过）。
            try {
                if ($range.Text -and $range.Text.Trim() -ne "") {
                    $h = $doc.Hyperlinks.Add($range, $url)
                } else {
                    $range.Text = $display
                    $h = $doc.Hyperlinks.Add($doc.Range($range.Start, $range.Start + $display.Length), $url)
                }
            } catch {
                Output-Json @{ success = $false; error = "插入超链接失败: $(Get-CallError $_)" }; return
            }
        }
        if ($null -eq $h) { Output-Json @{ success = $false; error = "Hyperlinks.Add 返回 null（5 参与 2 参签名都调不通）" }; return }
        Output-Json @{ success = $true; data = @{ url = $url; text = $display; position = $position } }
    }

    "convertTableToText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        $idx = [int](Get-IntParam $p "tableIndex" 1)
        if ($idx -lt 1 -or $idx -gt $doc.Tables.Count) { Output-Json @{ success = $false; error = "表格序号 $idx 不存在（共 $($doc.Tables.Count) 张）" }; return }
        $sep = if ($null -ne $p.separator) { [string]$p.separator } else { "," }
        try { $doc.Tables.Item($idx).ConvertToText($sep) } catch {
            Output-Json @{ success = $false; error = "表格转文本失败: $(Get-CallError $_)" }; return
        }
        $left = 0
        try { $left = [int]$doc.Tables.Count } catch {}
        Output-Json @{ success = $true; data = @{ tableIndex = $idx; separator = $sep; remainingTables = $left } }
    }

    "__probeWordCaps" {
        # Word 侧真机能力探测（第 5 期新增接口）。
        # 与 __probeCaps 的区别：那个探【用户当前文档】上的既有能力；
        # 这个在【临时文档】上探新接口，探完原样关闭（不保存），不碰用户任何文档。
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $null
        try {
            $doc = $word.Documents.Add()
            # 等文档就绪（同 createDocument 的教训：零等待会拿到空壳）
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Milliseconds 150
                try { $cur = $word.Documents.Item($word.Documents.Count); if (Test-DocHealthy $cur) { $doc = $cur; break } } catch {}
            }
            # 铺探测素材：一段文字 + 一张 2x2 表 + 一个书签
            try { $null = $doc.Content.InsertAfter("__probe__ 内容") } catch {}
            try { $null = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 2) } catch {}
            try { $null = $doc.Bookmarks.Add("__probe_bm__", $doc.Range(0, 1)) } catch {}

            $caps = [ordered]@{}
            $caps["stats.words"]        = { $null = $doc.Words.Count }
            $caps["stats.pages"]        = { $null = $doc.ComputeStatistics(2) }
            $caps["stats.paragraphs"]   = { $null = $doc.Paragraphs.Count }
            $caps["comments.count"]     = { $null = $doc.Comments.Count }
            $caps["comments.addDelete"] = { $c = $doc.Comments.Add($doc.Range(0, 1), "__probe__"); if ($null -eq $c) { throw "Comments.Add 返回 null" }; $c.Delete() }
            $caps["bookmarks.count"]    = { $null = $doc.Bookmarks.Count }
            $caps["tables.count"]       = { $null = $doc.Tables.Count }
            $caps["tables.cellText"]    = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1, 1).Range.Text } }
            $caps["hyperlinks.add"]     = { $h = $null
                                          try { $h = $doc.Hyperlinks.Add($doc.Range(0, 1), "https://example.invalid", $null, $null, "__probe_link__") } catch {}
                                          if ($null -eq $h) { $r = $doc.Range(0, 1); $r.Text = "__probe_link__"; $h = $doc.Hyperlinks.Add($doc.Range($r.Start, $r.Start + 13), "https://example.invalid") }
                                          if ($null -eq $h) { throw "5 参与 2 参签名都调不通" } }
            $caps["tables.convertToText"] = { $t2 = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 1, 2); $t2.ConvertToText(",") }

            foreach ($k in @($caps.Keys)) {
                $status = "unknown"; $err = $null
                try {
                    & $caps[$k] | Out-Null
                    $status = "supported"
                } catch {
                    $status = "unsupported"
                    $err = [string]$_.Exception.Message
                    if ($err.Length -gt 160) { $err = $err.Substring(0, 160) }
                }
                $caps[$k] = @{ status = $status; error = $err }
            }
            $keys = @($caps.Keys)
            $data = [ordered]@{
                total = $keys.Count
                supported = @($keys | Where-Object { $caps[$_].status -eq "supported" })
                unsupported = @($keys | Where-Object { $caps[$_].status -ne "supported" })
                caps = $caps
                leaked = @()
            }
            Output-Json @{ success = $true; data = $data }
        } catch {
            Output-Json @{ success = $false; error = "探测失败: $(Get-CallError $_)" }
        } finally {
            # 探测不留垃圾：临时文档不保存直接关。
            # 第 5 轮更正：原注释说"新建文档 FullName 为空"是错的 —— 实测未保存文档的
            # FullName 非空（等于文档名），空的是 Path。判"存过盘没有"一律看 Path。
            # 这里不能用 closeDocument 的 saveChanges 守卫（它会为从未保存的文档报错），
            # 显式 Close($false) 正是"放弃保存"的语义。
            if ($null -ne $doc) { try { $doc.Close($false) } catch {} }
        }
    }

    "insertText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $position = if ($p.position) { $p.position } else { "cursor" }
        # 第 2 期第 2 章轮 2：钉住文档（documentName 有值）时**绝不能**走 Selection。
        # Selection 是【全局焦点】：target.doc=A、而 B 恰好是活动文档时，文字会写进 B，
        # 而且返回 success —— 这正是 A1/A2 要抓的"静默改错"。
        # 钉住时 cursor 没有意义（光标属于活动文档），退化成"追加到目标文档末尾"。
        $pinned = [bool]($p -and $p.documentName)
        $range = $null
        switch ($position) {
            "start" { $range = $doc.Range(0, 0); $range.InsertBefore([string]$p.text) }
            "end" { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $range.InsertAfter([string]$p.text) }
            default {
                if ($pinned) {
                    $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                    $range.InsertAfter([string]$p.text)
                } else {
                    $word.Selection.TypeText([string]$p.text)
                }
            }
        }
        if ($p.style) {
            try {
                if ($pinned -and $null -ne $range) { $range.Style = $p.style }
                else { $word.Selection.Range.Style = $p.style }
            } catch {}
        }
        $landed = ''
        try { $landed = [string]$doc.Name } catch {}
        # doc = 实际落到了哪一份（钉住时它必须等于 documentName，不等就是串档）
        Output-Json @{ success = $true; data = @{ position = $position; textLength = $p.text.Length; doc = $landed } }
    }

    "setFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        # 第 2 期第 2 章轮 2：钉住文档 + 非 all 范围 = 用的是【活动文档的选区】，会改错地方。
        # 与 getSelectedText/setSelectedText 同一条理由：选区钉不住，宁可报错。
        if ($p.documentName -and $p.range -ne "all") {
            Output-Json @{ success = $false; error = '[勿重试·先改参数] 钉住文档（documentName）时不能用选区范围改字体：选区属于活动文档，会改到别人身上。请改传 range=all（作用于整份目标文档）。' }
            return
        }
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = $p.italic }
        if ($null -ne $p.underline) { $range.Font.Underline = $p.underline }
        if ($p.color) {
            $colorMap = @{ red = 255; blue = 16711680; green = 65280; black = 0 }
            $colorValue = $colorMap[$p.color.ToLower()]
            if ($null -eq $colorValue) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
            if ($null -ne $colorValue) { $range.Font.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ settings = @{ fontName = $p.fontName; fontSize = $p.fontSize; bold = $p.bold; italic = $p.italic; underline = $p.underline; color = $p.color } } }
    }

    "findReplace" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $p.findText
        $find.Replacement.Text = $p.replaceText
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $replaceAll = if ($null -ne $p.replaceAll) { [bool]$p.replaceAll } else { $true }
        $replaceType = if ($replaceAll) { 2 } else { 1 }
        $result = $find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, $p.replaceText, $replaceType)
        Output-Json @{ success = $true; data = @{ replaced = $result } }
    }

    "insertTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $range = $word.Selection.Range
        $table = $doc.Tables.Add($range, $rows, $cols)
        if ($p.data) {
            for ($r = 0; $r -lt [Math]::Min($p.data.Count, $rows); $r++) {
                for ($c = 0; $c -lt [Math]::Min($p.data[$r].Count, $cols); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$p.data[$r][$c]
                }
            }
        }
        $table.Borders.Enable = $true
        Output-Json @{ success = $true }
    }

    "setParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        $para = $range.ParagraphFormat
        if ($null -ne $p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2; justify = 3 }
            $align = $alignMap[$p.alignment]
            if ($null -eq $align) { $align = 0 }
            $para.Alignment = $align
        }
        if ($p.lineSpacing) {
            $para.LineSpacingRule = 4
            $para.LineSpacing = [double]$p.lineSpacing * 12
        }
        if ($null -ne $p.spaceBefore) { $para.SpaceBefore = $p.spaceBefore }
        if ($null -ne $p.spaceAfter) { $para.SpaceAfter = $p.spaceAfter }
        if ($null -ne $p.firstLineIndent) { $para.FirstLineIndent = [double]$p.firstLineIndent * 28.35 }
        if ($null -ne $p.leftIndent) { $para.LeftIndent = [double]$p.leftIndent * 28.35 }
        if ($null -ne $p.rightIndent) { $para.RightIndent = [double]$p.rightIndent * 28.35 }
        # 分页控制：论文/报告里"标题不与正文分离""每章另起一页"常用
        if ($null -ne $p.pageBreakBefore) { $para.PageBreakBefore = [bool]$p.pageBreakBefore }
        if ($null -ne $p.keepWithNext) { $para.KeepWithNext = [bool]$p.keepWithNext }
        if ($null -ne $p.keepTogether) { $para.KeepTogether = [bool]$p.keepTogether }
        Output-Json @{ success = $true }
    }

    "setPageSetup" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $ps = $doc.PageSetup
        if ($null -ne $p.topMargin) { $ps.TopMargin = [double]$p.topMargin * 28.35 }
        if ($null -ne $p.bottomMargin) { $ps.BottomMargin = [double]$p.bottomMargin * 28.35 }
        if ($null -ne $p.leftMargin) { $ps.LeftMargin = [double]$p.leftMargin * 28.35 }
        if ($null -ne $p.rightMargin) { $ps.RightMargin = [double]$p.rightMargin * 28.35 }
        if ($null -ne $p.orientation) {
            $ps.Orientation = if ($p.orientation -eq "landscape") { 1 } else { 0 }
        }
        if ($null -ne $p.paperSize) {
            $sizeMap = @{ A4 = 7; A3 = 6; Letter = 1; Legal = 5 }
            $size = $sizeMap[$p.paperSize]
            if ($null -eq $size) { $size = 7 }
            $ps.PaperSize = $size
        }
        Output-Json @{ success = $true }
    }

    "insertPageBreak" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $breakType = if ($p.type) { $p.type } else { "page" }
        # 第 1 期第 2 章第 2 轮：补 evenPage/oddPage（wdSectionBreakEvenPage=4 / OddPage=5）。
        # 之前 wps_word_insert_section_break 走的是另一个不存在的动作，而要接住它的四种类型
        # 就得让这个映射表认识这四个值 —— 否则 evenPage/oddPage 会静默退化成【分页符】，
        # 用户以为分了节，其实只是换页（后面按节设页边距会全部落到同一节上）。
        $breakTypeMap = @{ page = 7; column = 8; section = 2; sectionContinuous = 3; evenPage = 4; oddPage = 5 }
        $bt = $breakTypeMap[$breakType]
        if ($null -eq $bt) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 不认识的分隔符类型「$breakType」。可用：page / column / section / sectionContinuous / evenPage / oddPage" }
            return
        }
        $word.Selection.InsertBreak($bt)
        Output-Json @{ success = $true; data = @{ type = $breakType; code = $bt } }
    }

    "insertHeader" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        # 第 4 轮修：原来硬编码 Sections.Item(1)，指定第 3 节完全无效（还返回 success）。
        $secIdx = if ($p.section) { [int]$p.section } else { 1 }
        if ($secIdx -lt 1 -or $secIdx -gt $doc.Sections.Count) {
            Output-Json @{ success = $false; error = "第 $secIdx 节不存在（本文档共 $($doc.Sections.Count) 节）" }
            return
        }
        $section = $doc.Sections.Item($secIdx)
        $header = $section.Headers.Item(1)
        $header.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $header.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true }
    }

    "insertFooter" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        # 同 insertHeader：section 参数以前被丢弃，永远改第 1 节
        $secIdx = if ($p.section) { [int]$p.section } else { 1 }
        if ($secIdx -lt 1 -or $secIdx -gt $doc.Sections.Count) {
            Output-Json @{ success = $false; error = "第 $secIdx 节不存在（本文档共 $($doc.Sections.Count) 节）" }
            return
        }
        $section = $doc.Sections.Item($secIdx)
        $footer = $section.Footers.Item(1)
        $footer.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $footer.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        if ($p.includePageNumber) {
            $footer.Range.InsertAfter(" - 第 ")
            $footer.Range.Fields.Add($footer.Range, -1, "PAGE", $false)
            $footer.Range.InsertAfter(" 页 ")
        }
        Output-Json @{ success = $true }
    }

    "insertBookmark" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; return }
        $range = $word.Selection.Range
        $doc.Bookmarks.Add($p.name, $range)
        Output-Json @{ success = $true; data = @{ name = $p.name } }
    }

    "getDocumentParagraphs" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $startIdx = if ($null -ne $p.startParagraph) { [int]$p.startParagraph } else { 1 }
        $endIdx = if ($null -ne $p.endParagraph) { [int]$p.endParagraph } else { [Math]::Min($doc.Paragraphs.Count, $startIdx + 49) }
        if ($endIdx -gt $doc.Paragraphs.Count) { $endIdx = $doc.Paragraphs.Count }
        $paragraphs = @()
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $para = $doc.Paragraphs.Item($i)
            $text = $para.Range.Text
            $text = $text.TrimEnd("`r`n", "`r", "`n")
            if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + "..." }
            $styleName = ""
            try { $styleName = $para.Range.Style.NameLocal } catch { $styleName = "" }
            $paragraphs += @{ index = $i; text = $text; style = $styleName; start = $para.Range.Start; end = $para.Range.End }
        }
        Output-Json @{ success = $true; data = @{ paragraphs = $paragraphs; totalCount = $doc.Paragraphs.Count; returnedCount = $paragraphs.Count } }
    }

    "findInDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.findText) { Output-Json @{ success = $false; error = "findText required" }; return }
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $maxResults = if ($null -ne $p.maxResults) { [int]$p.maxResults } else { 20 }
        $results = @()
        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.findText
        $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        while ($found -and $results.Count -lt $maxResults) {
            $matchStart = $searchRange.Start
            $matchEnd = $searchRange.End
            $matchText = $searchRange.Text
            $paraIdx = 0
            for ($pi = 1; $pi -le $doc.Paragraphs.Count; $pi++) {
                $pr = $doc.Paragraphs.Item($pi).Range
                if ($matchStart -ge $pr.Start -and $matchStart -le $pr.End) { $paraIdx = $pi; break }
            }
            $ctxStart = [Math]::Max(0, $matchStart - 50)
            $ctxEnd = [Math]::Min($doc.Content.End, $matchEnd + 50)
            $ctxRange = $doc.Range($ctxStart, $ctxEnd)
            $context = $ctxRange.Text
            $results += @{ text = $matchText; start = $matchStart; end = $matchEnd; paragraphIndex = $paraIdx; context = $context }
            $searchRange = $doc.Range($matchEnd, $doc.Content.End)
            $searchRange.Find.ClearFormatting()
            $searchRange.Find.Text = $p.findText
            $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        }
        Output-Json @{ success = $true; data = @{ results = $results; count = $results.Count; findText = $p.findText } }
    }

    "smartFillField" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.keyword) { Output-Json @{ success = $false; error = "keyword required" }; return }
        if ($null -eq $p.value) { Output-Json @{ success = $false; error = "value required（空字符串将清除该字段内容）" }; return }
        $fillMode = if ($p.fillMode) { $p.fillMode.ToLower() } else { "auto" }

        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.keyword
        $found = $searchRange.Find.Execute($p.keyword, $false, $false, $false, $false, $false, $true, 1, $false, "", 0)
        if (-not $found) { Output-Json @{ success = $false; error = "Keyword '$($p.keyword)' not found" }; return }

        $matchStart = $searchRange.Start
        $matchEnd = $searchRange.End
        $paraRange = $searchRange.Paragraphs.Item(1).Range
        $paraText = $paraRange.Text
        $detectedMode = $fillMode

        if ($fillMode -eq "auto") {
            if ($paraText -match '[\{\【\[]' + [regex]::Escape($p.keyword) + '[\}\】\]]') {
                $detectedMode = "placeholder"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]\s*_+') {
                $detectedMode = "underline"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]') {
                $detectedMode = "afterColon"
            }
            else {
                $afterKeyword = $paraText.Substring([Math]::Min($paraText.IndexOf($p.keyword) + $p.keyword.Length, $paraText.Length))
                $afterKeyword = $afterKeyword.TrimStart().TrimEnd("`r`n", "`r", "`n").TrimEnd()
                if ($afterKeyword.Length -eq 0 -or $afterKeyword -match '^[\s\r\n]*$') {
                    $detectedMode = "afterLabel"
                } else {
                    $detectedMode = "afterColon"
                }
            }
        }

        $fillResult = ""
        switch ($detectedMode) {
            "placeholder" {
                $fullPlaceholder = ""
                if ($paraText -match '([\{\【\[])' + [regex]::Escape($p.keyword) + '([\}\】\]])') {
                    $fullPlaceholder = $Matches[0]
                } else {
                    $fullPlaceholder = $p.keyword
                }
                $replaceRange = $doc.Content.Duplicate
                $replaceRange.Find.ClearFormatting()
                $replaceRange.Find.Text = $fullPlaceholder
                $found2 = $replaceRange.Find.Execute($fullPlaceholder, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                if ($found2) { $fillResult = "Replaced placeholder '$fullPlaceholder' with '$($p.value)'" }
                else { $fillResult = "Failed to replace placeholder" }
            }
            "underline" {
                $keywordEndPos = $matchEnd
                $fillDone = $false
                for ($ri = 1; $ri -le $paraRange.Words.Count; $ri++) {
                    $wordObj = $paraRange.Words.Item($ri)
                    if ($wordObj.Start -gt $keywordEndPos -and $wordObj.Font.Underline -ne 0 -and $wordObj.Text -match '_+') {
                        $wordObj.Text = $p.value
                        $wordObj.Font.Underline = 1
                        $fillDone = $true
                        $fillResult = "Filled underlined field after '$($p.keyword)' with '$($p.value)'"
                        break
                    }
                }
                if (-not $fillDone) {
                    $underscorePattern = "_+"
                    $afterKeyRange = $doc.Range($keywordEndPos, $paraRange.End)
                    $afterKeyRange.Find.ClearFormatting()
                    $afterKeyRange.Find.Text = $underscorePattern
                    $afterKeyRange.Find.MatchWildcards = $true
                    $foundUl = $afterKeyRange.Find.Execute($underscorePattern, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                    if ($foundUl) {
                        $fillResult = "Filled underline after '$($p.keyword)' with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($keywordEndPos, $keywordEndPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after keyword '$($p.keyword)' (no underline found)"
                    }
                }
            }
            "afterColon" {
                $afterKeyRange = $doc.Range($matchEnd, $paraRange.End)
                $afterKeyText = $afterKeyRange.Text
                $colonOffset = $afterKeyText.IndexOfAny([char[]]@('：', ':'))
                if ($colonOffset -ge 0) {
                    $insertPos = $matchEnd + $colonOffset + 1
                    $afterColonRange = $doc.Range($insertPos, $paraRange.End - 1)
                    $afterColonText = $afterColonRange.Text.TrimStart().TrimEnd("`r`n", "`r", "`n").TrimEnd()
                    if ($afterColonText.Length -gt 0 -and $afterColonText -notmatch '^[\s\r\n_　]+$') {
                        $afterColonRange.Text = $p.value
                        $fillResult = "Replaced content after colon with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($insertPos, $insertPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after colon following '$($p.keyword)'"
                    }
                } else {
                    $insertRange = $doc.Range($matchEnd, $matchEnd)
                    $insertRange.InsertAfter("：" + $p.value)
                    $fillResult = "No colon found, inserted '：$($p.value)' after '$($p.keyword)'"
                }
            }
            "afterLabel" {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after label '$($p.keyword)'"
            }
            default {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after '$($p.keyword)' (default mode)"
            }
        }
        Output-Json @{ success = $true; data = @{ keyword = $p.keyword; value = $p.value; fillMode = $detectedMode; result = $fillResult } }
    }

    "replaceBookmarkContent" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; return }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; return }
        try {
            $bm = $doc.Bookmarks.Item($p.name)
            $bmStart = $bm.Start
            $bmRange = $bm.Range
            $bmRange.Text = $p.text
            $newEnd = $bmStart + $p.text.Length
            $newRange = $doc.Range($bmStart, $newEnd)
            $doc.Bookmarks.Add($p.name, $newRange) | Out-Null
            Output-Json @{ success = $true; data = @{ name = $p.name; text = $p.text; start = $bmStart; end = $newEnd } }
        } catch {
            Output-Json @{ success = $false; error = "Bookmark '$($p.name)' not found or failed to replace: $(Get-CallError $_)" }
        }
    }

    "addComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $text = if ($p.text) { $p.text } else { $p.comment }
        if (-not $text) { Output-Json @{ success = $false; error = "text required" }; return }
        $range = $word.Selection.Range
        $doc.Comments.Add($range, $text)
        Output-Json @{ success = $true }
    }

    "insertImage" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $path = if ($p.path) { $p.path } else { $p.filePath }
        if (-not $path) { Output-Json @{ success = $false; error = "path required" }; return }
        $range = $word.Selection.Range
        $shape = $doc.InlineShapes.AddPicture($path, $false, $true, $range)
        if ($p.width) { $shape.Width = $p.width }
        if ($p.height) { $shape.Height = $p.height }
        if ($p.scale) { $shape.ScaleWidth = $p.scale; $shape.ScaleHeight = $p.scale }
        Output-Json @{ success = $true; data = @{ width = $shape.Width; height = $shape.Height } }
    }

    "applyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($p.range -and $null -ne $p.range.start -and $null -ne $p.range.end) {
            $range = $doc.Range([int]$p.range.start, [int]$p.range.end)
        } else {
            $range = $word.Selection.Range
        }
        $range.Style = $p.styleName
        Output-Json @{ success = $true; data = @{ affectedText = $range.Text } }
    }

    # ==================== Word: Proofreading ====================
    "enableTrackChanges" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $enable = if ($null -ne $p.enable) { [bool]$p.enable } else { $true }
        $doc.TrackRevisions = $enable
        Output-Json @{ success = $true; data = @{ trackChanges = $enable; active = $doc.TrackRevisions } }
    }

    "getTrackChangesStatus" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $revisionCount = 0
        try { $revisionCount = $doc.Revisions.Count } catch {}
        Output-Json @{ success = $true; data = @{ trackChanges = [bool]$doc.TrackRevisions; revisionCount = $revisionCount } }
    }

    "replaceRange" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($null -eq $p.startPos -or $null -eq $p.endPos) { Output-Json @{ success = $false; error = "startPos and endPos required" }; return }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; return }
        try {
            $range = $doc.Range([int]$p.startPos, [int]$p.endPos)
            $originalText = $range.Text
            $range.Text = $p.text
            $actualEndPos = $range.End
            Output-Json @{ success = $true; data = @{ startPos = [int]$p.startPos; originalEndPos = [int]$p.endPos; endPos = $actualEndPos; originalText = $originalText.TrimEnd("`r`n"); newText = $p.text } }
        } catch {
            Output-Json @{ success = $false; error = "Failed to replace range: $(Get-CallError $_)" }
        }
    }

    # ==================== PPT ====================
    "getActivePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $shapes = @()
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                $text = ""
                try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = $shape.TextFrame.TextRange.Text.Substring(0, [Math]::Min(50, $shape.TextFrame.TextRange.Text.Length)) } } catch {}
                $shapes += @{ name = $shape.Name; type = $shape.Type; text = $text }
            }
            $slides += @{ index = $i; shapeCount = $slide.Shapes.Count; shapes = $shapes }
        }
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count; slides = $slides } }
    }

    "openPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; return }
        $pres = $ppt.Presentations.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count } }
    }

    "closePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        # 按名字定位必须自己遍历：WPS 的 Presentations.Item() 只吃序号，传名字会抛
        # DISP_E_NOTACOLLECTION（"Does not support a collection"）—— 实测踩过，
        # 表现为"关不掉演示文稿"，而且报错完全看不出是参数类型问题。
        $pres = $null
        if ($p.name) {
            for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
                $cand = $ppt.Presentations.Item($i)
                if ([string]$cand.Name -eq [string]$p.name) { $pres = $cand; break }
            }
            if ($null -eq $pres) {
                Output-Json @{ success = $false; error = "找不到名为「$($p.name)」的演示文稿" }
                return
            }
        } else {
            $pres = $ppt.ActivePresentation
        }
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No presentation" }; return }
        $name = $pres.Name
        # 第 3 轮修：原来 $pres.Close() 不带参数，TS 传的 save 被整个丢掉 ——
        # 既"没保存就关了"，未保存文稿还会弹「另存为」卡死通道。与 closeDocument /
        # closeWorkbook 同一套设计答案：读 saveChanges，未保存 + 要求保存就明确报错。
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        if ($saveChanges) {
            $hasPath = $false
            try { $hasPath = -not [string]::IsNullOrEmpty([string]$pres.Path) } catch {}
            if (-not $hasPath) {
                Output-Json @{ success = $false; error = "[勿重试·先另存为] 演示文稿「$name」从未保存过，直接关闭会弹出「另存为」对话框卡住通道（120 秒）。请先另存为（wps_ppt_save_as），或显式传 save=false 放弃保存" }
                return
            }
        }
        # Presentation.Close() 在 WPS/PPT 的对象模型里【不接受参数】——
        # 传 $saveChanges 进去会抛 DISP_E_NOTACOLLECTION，看着像集合用法错，其实是方法签名不符。
        # 想丢弃改动又要抑制「另存为」弹框，标准做法是先把 Saved 置 true 再 Close()。
        if (-not $saveChanges) { try { $pres.Saved = $true } catch {} }
        $pres.Close()
        # 第 6 轮：回包补 name。TS 读的是 data.name（打印"名称: xxx"），原来只回 closed
        # → 每次关闭后用户看到"名称: undefined"（自检里就是这么发现的）。
        Output-Json @{ success = $true; data = @{ closed = $name; name = $name; saved = [bool]$saveChanges } }
    }

    "getOpenPresentations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $presentations = @()
        for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
            $presItem = $ppt.Presentations.Item($i)
            $presentations += @{ name = $presItem.Name; path = $presItem.FullName; slideCount = $presItem.Slides.Count }
        }
        Output-Json @{ success = $true; data = @{ presentations = $presentations; count = $presentations.Count } }
    }

    "switchPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $presItem = $ppt.Presentations.Item($(if ($p.name) { $p.name } else { $p.index }))
        $presItem.Windows.Item(1).Activate()
        Output-Json @{ success = $true; data = @{ name = $presItem.Name } }
    }

    "deleteSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $pres.Slides.Item($index).Delete()
        Output-Json @{ success = $true; data = @{ deleted = $index } }
    }

    "duplicateSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $newSlide = $slide.Duplicate()
        $newIndex = $newSlide.Item(1).SlideIndex
        Output-Json @{ success = $true; data = @{ sourceIndex = $index; newIndex = $newIndex } }
    }

    "insertSlidesFromFile" {
        # 从另一个 PPT 文件把整页幻灯片插入当前演示文稿，保留来源格式（跨PPT整合）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $src = if ($p.filePath) { $p.filePath } elseif ($p.path) { $p.path } else { $null }
        if (-not $src) { Output-Json @{ success = $false; error = "filePath required" }; return }
        if (-not (Test-Path $src)) { Output-Json @{ success = $false; error = "source file not found: $src" }; return }
        $before = $pres.Slides.Count
        $afterIndex = if ($null -ne $p.afterIndex) { [int]$p.afterIndex } else { $before }
        if ($afterIndex -lt 0) { $afterIndex = 0 }
        if ($afterIndex -gt $before) { $afterIndex = $before }
        if ($null -ne $p.slideStart -and $null -ne $p.slideEnd) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart, [int]$p.slideEnd)
        } elseif ($null -ne $p.slideStart) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart)
        } else {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex)
        }
        $after = $pres.Slides.Count
        Output-Json @{ success = $true; data = @{ inserted = ($after - $before); afterIndex = $afterIndex; totalSlides = $after; source = $src } }
    }

    "moveSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $fromIndex = if ($p.from) { $p.from } else { $p.fromIndex }
        $toIndex = if ($p.to) { $p.to } else { $p.toIndex }
        $pres.Slides.Item($fromIndex).MoveTo($toIndex)
        Output-Json @{ success = $true; data = @{ from = $fromIndex; to = $toIndex } }
    }

    "getSlideCount" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        Output-Json @{ success = $true; data = @{ count = $pres.Slides.Count } }
    }

    "getSlideInfo" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $txt = ""
            try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $txt = $shape.TextFrame.TextRange.Text } } catch {}
            try {
                if ($shape.HasTable -and [string]::IsNullOrEmpty($txt)) {
                    $tbl = $shape.Table; $rws = @()
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        $cs = @()
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            $cv = ""
                            try { $cv = $tbl.Cell($rr, $cc).Shape.TextFrame.TextRange.Text } catch {}
                            $cs += $cv
                        }
                        $rws += ($cs -join "|")
                    }
                    $txt = "[表格] " + ($rws -join " ;; ")
                }
            } catch {}
            $txt = ($txt -replace "[`r`n`v]", " ")
            if ($txt.Length -gt 90) { $txt = $txt.Substring(0, 90) + "…" }
            $hasTxt = $false
            if ($shape.HasTextFrame) { $hasTxt = $true }
            # placeholderType：Shape.Type 对所有占位符都是 14（msoPlaceholder），没法区分标题/副标题/正文。
            # 而 setSlideSubtitle / setSlideContent 正是按 PlaceholderFormat.Type 找占位符的 ——
            # 不报出来的话，那两个工具报"placeholder not found"时根本无从判断是哪种版式能写。
            # ppPlaceholder 枚举：1=Title 2=Body 3=CenterTitle 4=Subtitle 7=Object
            $phType = ""
            try { if ($shape.Type -eq 14) { $phType = "$($shape.PlaceholderFormat.Type)" } } catch {}
            $shapes += @{ name = $shape.Name; type = "$($shape.Type)"; placeholderType = $phType; hasText = $hasTxt; text = $txt }
        }
        Output-Json @{ success = $true; data = @{ index = $index; slideIndex = $index; shapeCount = $slide.Shapes.Count; shapesCount = $slide.Shapes.Count; shapes = $shapes; layout = "$($slide.Layout)" } }
    }

    "switchSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $ppt.ActiveWindow.View.GotoSlide($index)
        Output-Json @{ success = $true; data = @{ currentSlide = $index } }
    }

    "setSlideLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $layoutMap = @{ title = 1; titleContent = 2; blank = 12; twoColumn = 4; comparison = 5; titleOnly = 11 }
        $layout = $layoutMap[$p.layout]
        if ($null -eq $layout) { $layout = if ($p.layout) { $p.layout } else { 2 } }
        $pres.Slides.Item($index).Layout = $layout
        Output-Json @{ success = $true; data = @{ slideIndex = $index; layout = $layout } }
    }

    # 第 6 轮补：setSlideSize / setSlideTheme 两个动作 TS 一直在调，PS 里都没有 case
    # → 冒烟里稳定报 Unknown action。补上，并且把"做不到的"说清楚，不糊一个 success。
    "setSlideSize" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $w = [double]$p.width
        $h = [double]$p.height
        if ($w -le 0 -or $h -le 0) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 宽高必须是正数（收到 width=$($p.width) height=$($p.height)）" }
            return
        }
        # 单位换算：TS 的 schema 承诺的是【像素】，而 WPS 的 PageSetup 要的是【磅】。
        # 1 px = 0.75 pt（96 DPI）。不换算的话 960×720 会被当成 960×720 磅 = 13.3×10 英寸，
        # 页面上所有元素都跑到角落里去 —— 用户看到的是"设完尺寸版面全乱了"。
        $wPt = $w * 0.75
        $hPt = $h * 0.75
        $pres.PageSetup.SlideWidth = $wPt
        $pres.PageSetup.SlideHeight = $hPt
        Output-Json @{ success = $true; data = @{ width = $w; height = $h; widthPt = $wPt; heightPt = $hPt; slideCount = $pres.Slides.Count } }
    }

    "setSlideTheme" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $theme = [string]$p.theme
        if (-not $theme) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 要给 theme（模板文件的完整路径）" }
            return
        }
        # WPS 的 COM 只有 ApplyTemplate(文件路径)，没有"按主题名切换"的接口。
        # 传进来的是名字（如「商务主题」）时明确报错 —— 这一条最容易被糊成 success，
        # 因为 ApplyTemplate 对不存在的文件在某些版本上会静默什么都不做。
        if (-not (Test-Path -LiteralPath $theme)) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 主题需要【模板文件路径】，而「$theme」不是本机存在的文件。WPS 没有按主题名切换的接口，请改用 .thmx / .potx / .ppt 模板文件的完整路径" }
            return
        }
        $pres.ApplyTemplate($theme)
        Output-Json @{ success = $true; data = @{ theme = $theme; slideCount = $pres.Slides.Count } }
    }

    "getSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $notes = ""
        try { $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text } catch { $notes = "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index; notes = $notes } }
    }

    "setSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = if ($p.notes) { $p.notes } else { "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index } }
    }

    "slide.add" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 50 }
        $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $shape.TextFrame.TextRange.Text = $p.text
        if ($p.fontSize) { $shape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($p.fontName) { $shape.TextFrame.TextRange.Font.Name = $p.fontName }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "deleteTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getTextBoxes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textBoxes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            if ($shape.HasTextFrame) {
                $text = ""
                try { $text = $shape.TextFrame.TextRange.Text } catch {}
                $textBoxes += @{ name = $shape.Name; index = $i; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
            }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; textBoxes = $textBoxes; count = $textBoxes.Count } }
    }

    "setTextBoxText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setTextBoxStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $tr = $shape.TextFrame.TextRange
        if ($p.fontSize) { $tr.Font.Size = $p.fontSize }
        if ($p.fontName) { $tr.Font.Name = $p.fontName }
        if ($null -ne $p.bold) { $tr.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $tr.Font.Italic = $p.italic }
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $tr.Font.Color = $colorValue }
        }
        if ($p.alignment) {
            $alignMap = @{ left = 1; center = 2; right = 3 }
            $tr.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $p.title } }
    }

    "getSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $title = ""
        if ($slide.Shapes.HasTitle) { $title = $slide.Shapes.Title.TextFrame.TextRange.Text }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $title } }
    }

    "setSlideSubtitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            if ($placeholderType -eq 4) {
                $shape.TextFrame.TextRange.Text = if ($p.subtitle) { $p.subtitle } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; subtitle = $p.subtitle } }
                return
            }
        }
        # 这两个工具的占位符编号是按 PlaceholderFormat.Type 匹配的，编号是【实测】出来的，
        # 不是查表抄的（WPS 的版式与 PowerPoint 不完全一样）。见 getSlideInfo 的 placeholderType 字段。
        Output-Json @{ success = $false; error = "[勿重试·先换版式] 本页没有副标题占位符（需要 PlaceholderFormat.Type=4）。WPS 上只有 layout='title' 的页有；空白页/标题+正文页没有。用 wps_ppt_add_slide(layout:'title') 新建一页，或改用 wps_ppt_set_slide_content 写正文占位符" }
    }

    "setSlideContent" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            # 2 = ppPlaceholderBody，7 = ppPlaceholderObject。
            # 实测 WPS 的 title_content 版式给的是 2（不是 PowerPoint 常见的 7），
            # 而这里原来只认 7 —— 于是"标题+正文"版式的页永远报 Content placeholder not found。
            # two_column / comparison 版式还有多个正文占位符，取第一个。
            if ($placeholderType -eq 2 -or $placeholderType -eq 7) {
                $shape.TextFrame.TextRange.Text = if ($p.content) { $p.content } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; content = $p.content } }
                return
            }
        }
        Output-Json @{ success = $false; error = "[勿重试·先换版式] 本页没有正文占位符（需要 PlaceholderFormat.Type=2 或 7）。WPS 上 layout='title_content' / 'two_column' / 'comparison' 的页有；空白页没有。用 wps_ppt_add_slide(layout:'title_content') 新建一页，或改用 wps_ppt_add_textbox 自己放一个文本框" }
    }

    "addShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapeTypeMap = @{ rectangle = 1; oval = 9; triangle = 7; diamond = 4; pentagon = 51; hexagon = 52; arrow = 13; star = 12; heart = 21; cloud = 179 }
        $shapeType = $shapeTypeMap[$p.type]
        if ($null -eq $shapeType) { $shapeType = if ($p.type) { $p.type } else { 1 } }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 100 }
        $shape = $slide.Shapes.AddShape($shapeType, $left, $top, $width, $height)
        if ($p.text) { $shape.TextFrame.TextRange.Text = $p.text }
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; slideIndex = $slideIndex } }
    }

    "deleteShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapes += @{ name = $shape.Name; index = $i; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; shapes = $shapes; count = $shapes.Count } }
    }

    "setShapeStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        if ($p.lineColor) {
            $lineColor = Convert-HexColorToRgbInt([string]$p.lineColor)
            if ($null -ne $lineColor) { $shape.Line.ForeColor.RGB = $lineColor }
        }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    # 第 6 轮补：TS 侧 wps_ppt_set_font_color 一直在调 setFontColor，但 PS 里没有这个 case，
    # 于是永远报 Unknown action（冒烟里就是这条）。注意"空文本"必须报错而不是返回 success ——
    # 给一个没有文字的形状设字体颜色，改了也看不见，属于本轮要消灭的假成功。
    "setFontColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $rgb = Convert-ColorToRgbInt([string]$p.color)
        if ($null -eq $rgb) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 颜色「$($p.color)」无法识别。请用 #RRGGBB（如 #FF0000）或常见颜色名（red/green/blue/black/white/yellow/orange/purple/gray/pink/cyan/magenta）" }
            return
        }
        if (-not $shape.HasTextFrame) {
            Output-Json @{ success = $false; error = "第 $slideIndex 页第 $($p.shapeIndex) 个形状没有文本框，无法设置文字颜色（要改填充色用 wps_ppt_set_shape_fill）" }
            return
        }
        if (-not $shape.TextFrame.HasText) {
            Output-Json @{ success = $false; error = "第 $slideIndex 页第 $($p.shapeIndex) 个形状是空文本，设置字色不会有任何可见变化。请先 wps_ppt_set_shape_text 写入文字" }
            return
        }
        $shape.TextFrame.TextRange.Font.Color.RGB = $rgb
        Output-Json @{ success = $true; data = @{ name = $shape.Name; color = [string]$p.color } }
    }

    # 第 3 轮补：TS 侧 wps_ppt_set_shape_fill 早就在调这个动作，但 PS 里没有这个 case，
    # 于是每次都报 Unknown action（表现为"工具坏了"而不是"参数丢了"）。
    # 写法照 setShapeStyle 的 fillColor 分支，只是参数名叫 color（schema 承诺的就是 color）。
    "setShapeFill" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.color) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $fillColor) {
                $shape.Fill.Visible = $true
                $shape.Fill.Solid()
                $shape.Fill.ForeColor.RGB = $fillColor
            }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; color = $p.color } }
    }

    "setShapeText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setShapePosition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height } }
    }

    "insertPptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $slide.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "deletePptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "replacePptImage" {
        # 原位替换图片：保留原图位置/尺寸/旋转，删除旧图后在同一矩形插入新图（换图不动版式）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $sel = if ($p.name) { $p.name } else { $p.shapeIndex }
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "shapeIndex or name required" }; return }
        $newPath = if ($p.path) { $p.path } elseif ($p.filePath) { $p.filePath } else { $null }
        if (-not $newPath) { Output-Json @{ success = $false; error = "path required" }; return }
        if (-not (Test-Path $newPath)) { Output-Json @{ success = $false; error = "image not found: $newPath" }; return }
        $old = $slide.Shapes.Item($sel)
        $l = $old.Left; $t = $old.Top; $w = $old.Width; $h = $old.Height
        $rot = 0
        try { $rot = $old.Rotation } catch {}
        $old.Delete()
        $pic = $slide.Shapes.AddPicture($newPath, $false, $true, $l, $t, $w, $h)
        try { $pic.Rotation = $rot } catch {}
        Output-Json @{ success = $true; data = @{ name = $pic.Name; left = $l; top = $t; width = $w; height = $h; path = $newPath } }
    }

    "setImageStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        if ($null -ne $p.rotation) { $shape.Rotation = $p.rotation }
        # 样式：schema 承诺了 style{borderColor,borderWidth,...}，而这个 case 以前只摆位置，
        # 传了样式等于没传（返回 success，图上啥都没变）。这里真正落到 Line 上。
        if ($p.borderColor -or $p.borderWidth) {
            try {
                $shape.Line.Visible = $true
                if ($p.borderColor) {
                    $bc = Convert-HexColorToRgbInt([string]$p.borderColor)
                    if ($null -ne $bc) { $shape.Line.ForeColor.RGB = $bc }
                }
                if ($null -ne $p.borderWidth) { $shape.Line.Weight = [double]$p.borderWidth }
            } catch {}
        }
        if ($p.shadow -and $p.shadow -is [string]) { } # 历史形状，忽略
        Output-Json @{ success = $true; data = @{ name = $shape.Name; styled = [bool]($p.borderColor -or $p.borderWidth) } }
    }

    "exportSlideAsImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation (presentationName not found or no active presentation)" }; return }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        if ($slideIndex -lt 1 -or $slideIndex -gt $pres.Slides.Count) { Output-Json @{ success = $false; error = "slideIndex $slideIndex out of range (1..$($pres.Slides.Count)) in '$($pres.Name)'" }; return }
        $slide = $pres.Slides.Item($slideIndex)
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        # 确保父目录存在
        $dir = Split-Path -Parent $outputPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 PowerPoint COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $width = if ($p.width) { [int]$p.width } else { 1280 }
        $height = if ($p.height) { [int]$p.height } else { 720 }
        try {
            $slide.Export($outputPath, $filterName, $width, $height)
        } catch {
            Output-Json @{ success = $false; error = "Slide.Export failed: $(Get-CallError $_)" }; return
        }
        # 关键修复：校验文件是否真生成，避免 Export 静默失败却报成功
        if (Test-Path $outputPath) {
            $bytes = (Get-Item $outputPath).Length
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; outputPath = $outputPath; format = $filterName; width = $width; height = $height; bytes = $bytes; presentation = $pres.Name } }
        } else {
            Output-Json @{ success = $false; error = "Export 未抛错但文件未生成: $outputPath（WPS Slide.Export 可能不支持该 filter/尺寸参数或该路径）" }
        }
    }

    "insertPptTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $rows = if ($p.rows) { $p.rows } else { 3 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 200 }
        $table = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        # 第 3 轮补 tableIndex：原来只回 name，TS 侧读 tableIndex 于是永远打印
        # "表格索引: undefined"。而 set_table_cell / set_table_row_style / get_table_cell
        # 全都按"第 N 个表格"定位 —— 不给序号，调用方建完表没法接着操作。
        # 算法与 setPptTableCellStyle 的定位算法一致：数本页 HasTable 的形状，新表排第几就是几。
        $tableIndex = 0
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $s = $slide.Shapes.Item($i)
            if ($s.HasTable) {
                $tableIndex++
                if ($s.Name -eq $table.Name) { break }
            }
        }
        Output-Json @{ success = $true; data = @{ name = $table.Name; tableIndex = $tableIndex; rows = $rows; cols = $cols } }
    }

    "setPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $textValue = if ($p.text) { $p.text } elseif ($p.value) { $p.value } else { "" }
        $cell.Shape.TextFrame.TextRange.Text = $textValue
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $textValue } }
    }

    "getPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $table = $shape.Table
        $rowCount = $table.Rows.Count
        $colCount = $table.Columns.Count
        $value = $null
        if ($null -ne $p.row -and $null -ne $p.col) {
            try { $value = $table.Cell($p.row, $p.col).Shape.TextFrame.TextRange.Text } catch { $value = $null }
        }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; value = $value; rowCount = $rowCount; colCount = $colCount } }
    }

    "setPptTableStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        # tableIndex 是"第几个表格"，不是"第几个形状" —— 以前直接当形状序号用，
        # 拿到的往往是矩形文本框，于是设样式全落空、返回却是 success。
        $shape = $null
        if ($p.tableName) {
            try { $shape = $slide.Shapes.Item([string]$p.tableName) } catch {}
        }
        if ($null -eq $shape) {
            $want = if ($null -ne $p.tableIndex) { [int]$p.tableIndex } else { 1 }
            $seen = 0
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $cand = $slide.Shapes.Item($i)
                $isTable = $false
                try { $isTable = [bool]$cand.HasTable } catch {}
                if ($isTable) {
                    $seen++
                    if ($seen -eq $want) { $shape = $cand; break }
                }
            }
        }
        if ($null -eq $shape) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 本页找不到第 $($p.tableIndex) 个表格" }
            return
        }
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        # 样式：这个 case 以前只摆位置/尺寸，传了 headerColor/fontSize 也什么都不做。
        $styled = 0
        try {
            if ($shape.HasTable) {
                $tbl = $shape.Table
                if ($p.headerColor) {
                    $hc = Convert-HexColorToRgbInt([string]$p.headerColor)
                    if ($null -ne $hc) {
                        for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
                            $tbl.Cell(1, $c).Shape.Fill.ForeColor.RGB = $hc
                        }
                        $styled++
                    }
                }
                if ($null -ne $p.fontSize -or $p.fontColor -or $null -ne $p.bold) {
                    for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
                        for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
                            $tr = $tbl.Cell($r, $c).Shape.TextFrame.TextRange
                            if ($null -ne $p.fontSize) { $tr.Font.Size = [double]$p.fontSize }
                            if ($p.fontColor) {
                                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                                if ($null -ne $fc) { $tr.Font.Color.RGB = $fc }
                            }
                            if ($null -ne $p.bold) { $tr.Font.Bold = [bool]$p.bold }
                        }
                    }
                    $styled++
                }
            }
        } catch {}
        Output-Json @{ success = $true; data = @{ name = $shape.Name; styled = ($styled -gt 0) } }
    }

    "setPptTableCellStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $cellShape = $cell.Shape
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
        }
        if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col } }
    }

    "setPptTableRowStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $rowIndex = if ($p.row) { $p.row } else { 1 }
        $row = $shape.Table.Rows.Item($rowIndex)
        # 参数形状归一化（第 3 轮修"假成功"）：
        # TS 侧 schema 承诺 style:{backgroundColor,fontColor,fontSize,bold,italic,fontName,alignment,height}，
        # 而历史裸调用（wps_execute_method 直穿）用的是扁平 $p.backgroundColor。
        # 两个形状都接：扁平优先（显式传来的），嵌套补位。只认一个，另一半调用方就是
        # "返回 success 但一个格子都没变"—— 那正是本轮要消灭的假成功。
        $st = @{}
        if ($p.style) {
            foreach ($k in $p.style.PSObject.Properties.Name) { $st[$k] = $p.style.$k }
        }
        foreach ($k in @('backgroundColor', 'fontColor', 'fontSize', 'bold', 'italic', 'fontName', 'alignment', 'height')) {
            if ($p.PSObject.Properties.Name -contains $k -and -not $st.ContainsKey($k)) { $st[$k] = $p.$k }
        }
        if ($null -ne $st.height) { $row.Height = $st.height }
        for ($c = 1; $c -le $row.Cells.Count; $c++) {
            $cellShape = $row.Cells.Item($c).Shape
            if ($st.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$st.backgroundColor)
                if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
            }
            if ($st.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$st.fontColor)
                if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
            }
            if ($st.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $st.fontSize }
            if ($null -ne $st.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $st.bold }
            if ($null -ne $st.italic) { $cellShape.TextFrame.TextRange.Font.Italic = $st.italic }
            if ($st.fontName) { $cellShape.TextFrame.TextRange.Font.Name = $st.fontName }
            if ($st.alignment) {
                $alignMap = @{ left = 1; center = 2; right = 3 }
                $al = $alignMap[[string]$st.alignment]
                if ($null -ne $al) { $cellShape.TextFrame.TextRange.ParagraphFormat.Alignment = $al }
            }
        }
        Output-Json @{ success = $true; data = @{ row = $rowIndex } }
    }

    "setShapeShadow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Shadow.Visible = $true
        if ($p.color) {
            $shadowColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $shadowColor) { $shape.Shadow.ForeColor.RGB = $shadowColor }
        }
        if ($null -ne $p.transparency) { $shape.Shadow.Transparency = $p.transparency }
        if ($null -ne $p.blur) { $shape.Shadow.Blur = $p.blur }
        if ($null -ne $p.offsetX) { $shape.Shadow.OffsetX = $p.offsetX }
        if ($null -ne $p.offsetY) { $shape.Shadow.OffsetY = $p.offsetY }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setBackgroundGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $bg = $slide.Background.Fill
        $bg.Visible = $true
        # 同 setShapeGradient：schema 给 gradient:{colors:[...]}，摊平后是 $p.colors
        if ($null -eq $p.color1 -and $p.colors) {
            $arr = @($p.colors)
            if ($arr.Count -ge 1) { $p | Add-Member -NotePropertyName color1 -NotePropertyValue $arr[0] -Force }
            if ($arr.Count -ge 2) { $p | Add-Member -NotePropertyName color2 -NotePropertyValue $arr[1] -Force }
        }
        $bg.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $bg.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $bg.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "setShapeGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        # 文档承诺的是 gradient:{type,colors:["#a","#b"]}，入口摊平后变成 $p.colors；
        # 而这里只认 $p.color1/$p.color2 —— 于是"传了渐变却什么都没变"。两种都接。
        if ($null -eq $p.color1 -and $p.colors) {
            $arr = @($p.colors)
            if ($arr.Count -ge 1) { $p | Add-Member -NotePropertyName color1 -NotePropertyValue $arr[0] -Force }
            if ($arr.Count -ge 2) { $p | Add-Member -NotePropertyName color2 -NotePropertyValue $arr[1] -Force }
        }
        $shape.Fill.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $shape.Fill.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $shape.Fill.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeBorder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.color) {
            $borderColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $borderColor) { $shape.Line.ForeColor.RGB = $borderColor }
        }
        if ($p.width) { $shape.Line.Weight = $p.width }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeTransparency" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.transparency) {
            # 工具/文档说的是"百分比 0~100"，COM 要的是 0~1 —— 直接塞 50 会报"值不在预期的范围内"。
            $tv = [double]$p.transparency
            if ($tv -gt 1) { $tv = $tv / 100 }
            if ($tv -lt 0) { $tv = 0 }
            if ($tv -gt 1) { $tv = 1 }
            $shape.Fill.Transparency = $tv
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeRoundness" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.roundness) { $shape.Adjustments.Item(1) = $p.roundness }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeFullStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) { $shape.Fill.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.fillColor) }
        if ($p.lineColor) { $shape.Line.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.lineColor) }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = $p.transparency }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "alignShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $alignMap = @{ left = 1; center = 2; right = 3; top = 4; middle = 5; bottom = 6 }
        $align = $alignMap[$p.alignment]
        if ($null -eq $align) { $align = 1 }
        # 没给形状清单就【不要】替用户选。
        # Shapes.Range($null) 在 WPS 上不是报错，是"整页所有形状"——
        # 于是一个只想对齐两个形状的调用，会把表格、图片、母版继承的占位符一起搬走，
        # 而且工具回 success。宁可响亮失败（[勿重试·先改参数]），也不要静默改整页。
        if (-not $p.names -or @($p.names).Count -eq 0) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] alignShapes 需要 shapeIndices（要参与对齐的形状序号数组）。不给就整页对齐会误伤表格/图片，请传 shapeIndices（1 起）" }
            return
        }
        $range = $slide.Shapes.Range($p.names)
        $range.Align($align, 1)
        Output-Json @{ success = $true; data = @{ alignment = $p.alignment } }
    }

    "distributeShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $distMap = @{ horizontal = 0; vertical = 1 }
        $dist = $distMap[$p.direction]
        if ($null -eq $dist) { $dist = 0 }
        $range = $slide.Shapes.Range($p.names)
        $range.Distribute($dist, 1)
        Output-Json @{ success = $true; data = @{ direction = $p.direction } }
    }

    "groupShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $range = $slide.Shapes.Range($p.names)
        $group = $range.Group()
        Output-Json @{ success = $true; data = @{ name = $group.Name } }
    }

    "duplicateShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $dup = $shape.Duplicate()
        Output-Json @{ success = $true; data = @{ name = $dup.Name } }
    }

    "setShapeZOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $zMap = @{ front = 0; forward = 2; backward = 3; back = 1 }
        $z = $zMap[$p.zOrder]
        if ($null -eq $z) { $z = 0 }
        $shape.ZOrder($z)
        Output-Json @{ success = $true; data = @{ name = $shape.Name; zOrder = $p.zOrder } }
    }

    "insertPptChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 300 }
        # AddChart 的第一个参数是 xlChartType 枚举（数字），不是类型名字符串。
        # 之前这里直接把 $p.type 丢进去：TS 传 chartType 时它是 $null（靠 COM 兜默认值蒙混过去），
        # 入口归一化把 chartType 映射成 type 之后，"bar" 这个字符串就撞上来了 ——
        # 报"无法将类型 string 的 bar 值转换为类型"。根因是这张枚举表从来没写过。
        # 现在显式翻译：名字 → 数字；已经是数字的照用；认不出来的退回柱状图（Excel 侧同一套表，见 addChart）。
        $chartTypeEnum = @{
            column = 51; column_clustered = 51; column_stacked = 52
            bar = 57; bar_clustered = 57; bar_stacked = 58
            line = 4; line_markers = 65; line_stacked = 63
            pie = 5; doughnut = -4120
            area = 1; area_stacked = 76
            scatter = -4169; bubble = 15; radar = -4151
        }
        $rawType = if ($null -ne $p.type) { $p.type } elseif ($null -ne $p.chartType) { $p.chartType } else { $null }
        $chartType = if ($rawType -is [int] -or $rawType -is [long] -or $rawType -is [double]) {
            [int]$rawType
        } elseif ($null -ne $rawType -and $chartTypeEnum.ContainsKey([string]$rawType)) {
            $chartTypeEnum[[string]$rawType]
        } else {
            51
        }
        $before = $slide.Shapes.Count
        $shape = $slide.Shapes.AddChart($chartType, $left, $top, $width, $height)
        # 回 chartIndex（= 插入后的形状总数，因为 AddChart 总是追加在最后）。
        # TS 的 insertPptChartHandler 打印的是 data.chartIndex，原来不返回 → 用户永远看到
        # "图表索引: undefined"，而后续 set_ppt_chart_data / set_ppt_chart_style 恰恰靠它定位图表。
        Output-Json @{ success = $true; data = @{ name = $shape.Name; chartType = $chartType; chartIndex = $slide.Shapes.Count; previousShapeCount = $before } }
    }

    "addAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        # 两个致命点，都是实测踩出来的（不是推测）：
        #   1) 老实现只读 shapeName，而工具传的是 shapeIndex → Shapes.Item($null) 拿到坏对象，
        #      AddEffect 之后【整页】的 Shapes/Background 全部报"不能对 Null 值表达式调用方法"，
        #      而且文稿列表还正常 —— 表现像"69 个工具一起坏了"，实际是这一条把页状态弄坏了。
        #   2) effect 直接透传字符串（"fadeIn"），而 AddEffect 要的是 MsoAnimEffect 枚举数字。
        #      这里补一张小表，认不出的明确报错，不再把坏值喂给 COM。
        $shape = $null
        if ($p.shapeName) { try { $shape = $slide.Shapes.Item([string]$p.shapeName) } catch {} }
        if ($null -eq $shape -and $p.shapeIndex) { try { $shape = $slide.Shapes.Item([int]$p.shapeIndex) } catch {} }
        if ($null -eq $shape) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 找不到要加动画的形状：shapeIndex（1 起）/shapeName 都没给或不存在" }
            return
        }
        # animationType 是已下线工具 wps_ppt_set_animation 的旧键名，仍认（否则老调用一律失败）
        $rawEffect = if ($p.effect) { $p.effect } elseif ($p.animationType) { $p.animationType } else { $null }
        if (-not $rawEffect) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 缺少 effect（可用：appear/fadeIn/flyIn/wipeIn/zoomIn/pulse/spin/grow/teeter，或直接给 MsoAnimEffect 数字）" }
            return
        }
        # 第 8 轮动画族合并后，这张表要同时接住三类：入场（entrance）、强调（emphasis）、退场（exit）。
        # 强调那几个值沿用 addEmphasisAnimation 一直在用的；退场在 WPS 上没有稳定的枚举名，
        # 所以【不猜】—— 认不出的名字会明确报错并提示改传 MsoAnimEffect 数字（AddEffect 是吃数字的）。
        $effectMap = @{
            # 入场
            appear = 1; fadeIn = 10; flyIn = 2; wipeIn = 22; zoomIn = 53; bounceIn = 26; spinIn = 15
            dissolveIn = 9; boxIn = 4; circleIn = 6; diamondIn = 8; wheelIn = 21; plusIn = 13
            wedgeIn = 20; randomBarsIn = 14; stripsIn = 18; blindsIn = 3; checkerboardIn = 5
            crawlIn = 7; stretchIn = 17; swivelIn = 19; floatUp = 30; riseUp = 34; thinLine = 36
            unfold = 37; whip = 38; expand = 50; shrink = 46
            # 强调（与 addEmphasisAnimation 用的是同一套）
            pulse = 63; spin = 15; grow = 59; teeter = 28; colorPulse = 62
        }
        $effectId = if ($rawEffect -is [int] -or $rawEffect -is [long] -or $rawEffect -is [double]) { [int]$rawEffect } elseif ($effectMap.ContainsKey([string]$rawEffect)) { $effectMap[[string]$rawEffect] } else { $null }
        if ($null -eq $effectId) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 不认识的 effect「$($p.effect)」；常用：appear/fadeIn/flyIn/wipeIn/zoomIn（入场）、pulse/spin/grow/teeter（强调）。退场效果请直接给 MsoAnimEffect 数字（AddEffect 吃的是数字）" }
            return
        }
        try {
            # 第三个参数是 Level（文本级别），不是"触发时机"。对没有文本的形状必须传
            # msoAnimateLevelNone(0)：传 1 会让 WPS 建出【两条】一模一样的效果
            # （实测 count=2，getAnimations 里两条同名同效果），删一条还剩一条，看起来像删不掉。
            $slide.TimeLine.MainSequence.AddEffect($shape, $effectId, 0) | Out-Null
        } catch {
            Output-Json @{ success = $false; error = "添加动画失败：$(Get-CallError $_)" }
            return
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; effectId = $effectId; count = $slide.TimeLine.MainSequence.Count } }
    }

    "setSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        # 合并后的 wps_ppt_set_transition 传的是 effect；旧的直接调用方还在传 transition，两个都认。
        if ($null -eq $p.effect -and $p.transition) { $p | Add-Member -NotePropertyName effect -NotePropertyValue $p.transition -Force }
        $effectName = if ($p.effect) { ([string]$p.effect).ToLower() } else { "fade" }
        # EntryEffect 是 PpEntryEffect 枚举：直接赋字符串 WPS 会抛"指定的转换无效"，
        # 而以前这里就是把 $p.effect 原样赋过去 —— 取消切换（none）更是必然炸。
        $effectMap = @{ none = 0; fade = 1; push = 2; wipe = 3; split = 4; reveal = 5; cover = 6; dissolve = 7; curtains = 8; random = 6; blinds = 3 }
        $entry = -1
        if ($effectMap.ContainsKey($effectName)) { $entry = $effectMap[$effectName] }
        elseif ($effectName -match '^\d+$') { $entry = [int]$effectName }
        if ($entry -lt 0) {
            Output-Json @{ success = $false; error = "不认识的切换效果：$effectName（可用：none/fade/push/wipe/split/reveal/cover/dissolve/curtains）" }
            return
        }
        $tr = $slide.SlideShowTransition
        $tr.EntryEffect = $entry
        $dur = 0.0
        if ($p.duration) { try { $dur = [double]$p.duration } catch { $dur = 0.0 } }
        if ($dur -gt 0) { try { $tr.Duration = $dur } catch {} }
        # WPS COM 没有"切换时播放声音"的接口（那是 PowerPoint 的 AddMediaObject2），传了如实回报而不是假装生效
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; effect = $effectName; entryEffect = $entry; duration = $dur; soundApplied = $false } }
    }

    "setSlideBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
        }
        if ($p.imagePath) {
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.UserPicture($p.imagePath)
        }
        # 渐变：合并后的 wps_ppt_set_background(target=slide,type=gradient) 走的就是这条，
        # 而这里以前只认 color/imagePath → 传渐变没反应。colors 数组与 color1/color2 都接。
        if ($null -eq $p.color1 -and $p.colors) {
            $arr = @($p.colors)
            if ($arr.Count -ge 1) { $p | Add-Member -NotePropertyName color1 -NotePropertyValue $arr[0] -Force }
            if ($arr.Count -ge 2) { $p | Add-Member -NotePropertyName color2 -NotePropertyValue $arr[1] -Force }
        }
        if ($p.color1) {
            try {
                $slide.FollowMasterBackground = $false
                $bgf = $slide.Background.Fill
                $bgf.Visible = $true
                $bgf.TwoColorGradient(1, 1)
                $c1 = Convert-HexColorToRgbInt([string]$p.color1)
                if ($null -ne $c1) { $bgf.ForeColor.RGB = $c1 }
                if ($p.color2) {
                    $c2 = Convert-HexColorToRgbInt([string]$p.color2)
                    if ($null -ne $c2) { $bgf.BackColor.RGB = $c2 }
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "slide.unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        $includeTitle = if ($null -ne $p.includeTitle) { [bool]$p.includeTitle } else { $true }
        $includeBody = if ($null -ne $p.includeBody) { [bool]$p.includeBody } else { $true }
        $start = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $end = if ($p.slideIndex) { [int]$p.slideIndex } else { $pres.Slides.Count }
        for ($i = $start; $i -le $end; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $isTitle = $false
                        try { $isTitle = $shape.PlaceholderFormat.Type -eq 1 -or $shape.PlaceholderFormat.Type -eq 3 } catch {}
                        if (($isTitle -and $includeTitle) -or (-not $isTitle -and $includeBody)) {
                            $shape.TextFrame.TextRange.Font.Name = $fontName
                            $count++
                        }
                    }
                } catch {}
            }
        }
        $slideCount = if ($p.slideIndex) { 1 } else { $pres.Slides.Count }
        Output-Json @{ success = $true; data = @{ font = $fontName; count = $count; slideCount = $slideCount } }
    }

    "addConnector" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 0 }
        $connector = $slide.Shapes.AddConnector(1, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $connector.Name } }
    }

    "addArrow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 120 }
        $height = if ($p.height) { $p.height } else { 60 }
        $shape = $slide.Shapes.AddShape(13, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "applyColorScheme" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemeKey = if ($p.scheme) { $p.scheme } elseif ($p.colorScheme) { $p.colorScheme } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ scheme = $schemeKey; count = $count } }
    }

    "autoBeautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch {}
        }
        # 第 6 轮：把"改了几个文本"和"处理了哪页"分开回，TS 才不会把 count 当页数打印。
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count; slideCount = 1; slideIndex = $slideIndex } }
    }

    "createKpiCards" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        # `@(...)` 不是多余的：写成 `$cards = if ($p.cards) { $p.cards }` 时，
        # PowerShell 会把【单元素】数组在管道里摊平，$cards 变成裸 PSCustomObject，
        # 它的 .Count 是空的 → 下面的 for 一次都不跑 → 工具返回 success 却一个形状都没建。
        # （createMiniCharts / smartDistribute 是同一个写法，同一个坑。）
        $cards = @(if ($p.cards) { $p.cards })
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 350 }
        $cardWidth = if ($p.cardWidth) { $p.cardWidth } else { 150 }
        $cardHeight = if ($p.cardHeight) { $p.cardHeight } else { 80 }
        $gap = if ($p.gap) { $p.gap } else { 20 }

        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1", "#fd7e14")
        $created = @()

        for ($i = 0; $i -lt $cards.Count; $i++) {
            $card = $cards[$i]
            $x = $startX + ($i * ($cardWidth + $gap))
            $color = if ($card.color) { $card.color } else { $colors[$i % $colors.Count] }
            $shape = $slide.Shapes.AddShape(5, $x, $startY, $cardWidth, $cardHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$color)
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.Shadow.Visible = $true
            $shape.Shadow.Blur = 10
            $shape.Shadow.OffsetX = 3
            $shape.Shadow.OffsetY = 3
            $shape.Shadow.Transparency = 0.6
            $title = if ($card.title) { $card.title } else { "" }
            $value = if ($card.value) { $card.value } else { "" }
            $shape.TextFrame.TextRange.Text = ($title + "`n" + $value)
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 14
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $created += $shape.Name
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; cards = $created } }
    }

    "createStyledTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 5 }
        $cols = if ($p.cols) { $p.cols } else { 4 }
        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 150 }
        $width = if ($p.width) { $p.width } else { 620 }
        $height = if ($p.height) { $p.height } else { 200 }

        $tableShape = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        $table = $tableShape.Table

        $styleKey = if ($p.style) { $p.style } else { "business" }
        $styles = @{
            business = @{ header = "#1a365d"; headerText = "#ffffff"; oddRow = "#f8f9fa"; evenRow = "#ffffff" }
            tech = @{ header = "#0d47a1"; headerText = "#ffffff"; oddRow = "#e3f2fd"; evenRow = "#ffffff" }
            elegant = @{ header = "#37474f"; headerText = "#ffffff"; oddRow = "#eceff1"; evenRow = "#ffffff" }
        }
        $style = $styles[$styleKey]
        if ($null -eq $style) { $style = $styles["business"] }

        for ($c = 1; $c -le $cols; $c++) {
            $headerCell = $table.Cell(1, $c)
            $headerColor = Convert-HexColorToRgbInt([string]$style.header)
            if ($null -ne $headerColor) { $headerCell.Shape.Fill.Solid(); $headerCell.Shape.Fill.ForeColor.RGB = $headerColor }
            $headerTextColor = Convert-HexColorToRgbInt([string]$style.headerText)
            if ($null -ne $headerTextColor) { $headerCell.Shape.TextFrame.TextRange.Font.Color.RGB = $headerTextColor }
            $headerCell.Shape.TextFrame.TextRange.Font.Bold = $true
            $headerCell.Shape.TextFrame.TextRange.Font.Size = 12
        }

        for ($r = 2; $r -le $rows; $r++) {
            $rowColor = if ($r % 2 -eq 0) { $style.evenRow } else { $style.oddRow }
            $rowColorValue = Convert-HexColorToRgbInt([string]$rowColor)
            for ($c = 1; $c -le $cols; $c++) {
                $cell = $table.Cell($r, $c)
                if ($null -ne $rowColorValue) { $cell.Shape.Fill.Solid(); $cell.Shape.Fill.ForeColor.RGB = $rowColorValue }
                $cell.Shape.TextFrame.TextRange.Font.Size = 11
            }
        }

        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    if (($r + 1) -le $rows -and ($c + 1) -le $cols) {
                        $table.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = [string]$rowData[$c]
                    }
                }
            }
        }

        Output-Json @{ success = $true; data = @{ name = $tableShape.Name; rows = $rows; cols = $cols } }
    }

    "addTitleDecoration" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 130 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 5 }
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        $bar = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
        if ($null -ne $colorValue) { $bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $colorValue }
        $bar.Line.Visible = $false
        Output-Json @{ success = $true; data = @{ name = $bar.Name } }
    }

    "addPageIndicator" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $total = $pres.Slides.Count
        $current = $slideIndex
        $text = "$current / $total"
        $left = if ($p.left) { $p.left } else { 880 }
        $top = if ($p.top) { $p.top } else { 500 }
        $box = $slide.Shapes.AddTextbox(1, $left, $top, 80, 30)
        $box.TextFrame.TextRange.Text = $text
        $box.TextFrame.TextRange.Font.Size = 12
        $box.TextFrame.TextRange.Font.Color.RGB = if ($p.dark) { 0 } else { 16777215 }
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        Output-Json @{ success = $true; data = @{ name = $box.Name; page = $text } }
    }

    "beautifyAllSlides" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        # 第 6 轮：TS 读 data.slideCount（"处理页数: N 页"），原来只回 count → 永远 undefined
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count; slideCount = $pres.Slides.Count } }
    }

    "createProgressBar" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 300 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 30 }
        $progress = if ($null -ne $p.progress) { [double]$p.progress } else { 0.5 }
        if ($progress -lt 0) { $progress = 0 }
        if ($progress -gt 1) { $progress = 1 }
        $label = if ($p.label) { $p.label } else { "" }

        $bgBar = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
        $bgColor = Convert-HexColorToRgbInt([string]$(if ($p.bgColor) { $p.bgColor } else { "#e0e0e0" }))
        if ($null -ne $bgColor) { $bgBar.Fill.Solid(); $bgBar.Fill.ForeColor.RGB = $bgColor }
        $bgBar.Line.Visible = $false

        $progressWidth = $width * $progress
        $progressBar = $slide.Shapes.AddShape(5, $left, $top, $progressWidth, $height)
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#28a745" }))
        if ($null -ne $colorValue) { $progressBar.Fill.Solid(); $progressBar.Fill.ForeColor.RGB = $colorValue }
        $progressBar.Line.Visible = $false

        $percentText = $slide.Shapes.AddTextbox(1, $left + $width + 10, $top, 60, $height)
        $percentText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($progress * 100))
        $percentText.TextFrame.TextRange.Font.Size = 14
        $percentText.TextFrame.TextRange.Font.Bold = $true

        if ($label) {
            $labelBox = $slide.Shapes.AddTextbox(1, $left, $top - 25, $width, 20)
            $labelBox.TextFrame.TextRange.Text = $label
            $labelBox.TextFrame.TextRange.Font.Size = 12
        }

        Output-Json @{ success = $true; data = @{ bgBar = $bgBar.Name; progressBar = $progressBar.Name; progress = $progress } }
    }

    "createGauge" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 300 }
        $radius = if ($p.radius) { $p.radius } else { 100 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.5 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]"#e0e0e0")
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.7
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $valueText = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 20, 80, 40)
        $valueText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($value * 100))
        $valueText.TextFrame.TextRange.Font.Size = 24
        $valueText.TextFrame.TextRange.Font.Bold = $true
        $valueText.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 10, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 14
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ gauge = $outerCircle.Name; value = $value } }
    }

    "createMiniCharts" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $items = @(if ($p.items) { $p.items })
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 400 }
        $itemWidth = if ($p.itemWidth) { $p.itemWidth } else { 100 }
        $gap = if ($p.gap) { $p.gap } else { 30 }
        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1")
        $created = @()

        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            $x = $startX + ($i * ($itemWidth + $gap))
            $color = if ($item.color) { $item.color } else { $colors[$i % $colors.Count] }
            $colorValue = Convert-HexColorToRgbInt([string]$color)

            $valueBox = $slide.Shapes.AddTextbox(1, $x, $startY, $itemWidth, 30)
            # [string] 不是多余的：data[].value 从 JSON 过来是 Int32，
            # 直接赋给 TextRange.Text 会抛"指定的转换无效"（同 setPptChartData 里 Value2 的坑）。
            $valueBox.TextFrame.TextRange.Text = [string]$(if ($item.value) { $item.value } else { "0" })
            $valueBox.TextFrame.TextRange.Font.Size = 20
            $valueBox.TextFrame.TextRange.Font.Bold = $true
            if ($null -ne $colorValue) { $valueBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }

            $labelBox = $slide.Shapes.AddTextbox(1, $x, $startY + 30, $itemWidth, 20)
            $labelBox.TextFrame.TextRange.Text = if ($item.label) { $item.label } else { "" }
            $labelBox.TextFrame.TextRange.Font.Size = 11
            $labelBox.TextFrame.TextRange.Font.Color.RGB = 6710886

            if ($item.trend) {
                $trendBox = $slide.Shapes.AddTextbox(1, $x + $itemWidth - 30, $startY, 30, 20)
                $trendBox.TextFrame.TextRange.Text = if ($item.trend -eq "up") { "↑" } else { "↓" }
                $trendBox.TextFrame.TextRange.Font.Size = 14
                $trendColor = if ($item.trend -eq "up") { "#28a745" } else { "#dc3545" }
                $trendColorValue = Convert-HexColorToRgbInt([string]$trendColor)
                if ($null -ne $trendColorValue) { $trendBox.TextFrame.TextRange.Font.Color.RGB = $trendColorValue }
            }

            $created += @{ label = $item.label; value = $item.value }
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; items = $created } }
    }

    "createDonutChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 280 }
        $radius = if ($p.radius) { $p.radius } else { 80 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.75 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }
        $centerText = if ($p.centerText) { $p.centerText } else { "{0}%" -f ([Math]::Round($value * 100)) }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#0d47a1" }))
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.6
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $textBox = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 15, 80, 30)
        $textBox.TextFrame.TextRange.Text = $centerText
        $textBox.TextFrame.TextRange.Font.Size = 18
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 15, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 12
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ donut = $outerCircle.Name; value = $value } }
    }

    "autoLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $layout = if ($p.layout) { $p.layout } else { "grid" }
        $margin = if ($p.margin) { [double]$p.margin } else { 50.0 }
        $gap = if ($p.gap) { [double]$p.gap } else { 20.0 }
        $slideWidth = 720.0
        $slideHeight = 540.0

        # 逐个形状取值：属性读取会抛。实测在整轮冒烟里，前面的用例给形状加过 3D 旋转/材质之后，
        # 再读 Width/Height 就抛"指定的转换无效" —— 一个形状读失败不该让整页排版全废。
        # 这里把可排的形状收成 { shape, w, h }，取不到尺寸的直接跳过。
        $shapes = @()
        $skipped = 0
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $s = $null
            try { $s = $slide.Shapes.Item($i) } catch { $skipped++; continue }
            $t = -1
            try { $t = [int]$s.Type } catch { $t = -1 }
            if ($t -ne 1 -and $t -ne 6 -and $t -ne 17) { continue }
            $w = 0.0; $h = 0.0
            try { $w = [double]$s.Width } catch { $skipped++; continue }
            try { $h = [double]$s.Height } catch { $skipped++; continue }
            $shapes += , @($s, $w, $h)
        }

        if ($shapes.Count -eq 0) {
            Output-Json @{ success = $true; data = @{ message = "没有可排列的形状"; count = 0; skipped = $skipped } }
            return
        }

        if ($layout -eq "horizontal") {
            $totalWidth = 0.0
            foreach ($row in $shapes) { $totalWidth += [double]$row[1] }
            $availWidth = $slideWidth - 2 * $margin
            $spacing = ($availWidth - $totalWidth) / ($shapes.Count + 1)
            $currentX = [double]$margin + $spacing
            foreach ($row in $shapes) {
                $shape = $row[0]; $w = [double]$row[1]; $h = [double]$row[2]
                try { $shape.Left = [double]$currentX } catch { $skipped++ }
                try { $shape.Top = [double](($slideHeight - $h) / 2) } catch { $skipped++ }
                $currentX = $currentX + $w + $spacing
            }
        } elseif ($layout -eq "vertical") {
            # 全部走 double 本地变量再赋值：直接拿表达式结果写 COM 的 Single 属性
            # 会抛"指定的转换无效"（实测只有 vertical 分支炸）。
            $currentY = [double]$margin
            foreach ($row in $shapes) {
                $shape = $row[0]; $w = [double]$row[1]; $h = [double]$row[2]
                try { $shape.Left = [double](($slideWidth - $w) / 2) } catch { $skipped++ }
                try { $shape.Top = [double]$currentY } catch { $skipped++ }
                $currentY = $currentY + $h + [double]$gap
            }
        } else {
            # 显式 int：Math.Ceiling/Floor 返回 double，拿去做取模/赋值时 COM 的 Single 属性会抛"指定的转换无效"
            $cols = if ($p.cols) { [int]$p.cols } else { [int][Math]::Ceiling([Math]::Sqrt($shapes.Count)) }
            if ($cols -lt 1) { $cols = 1 }
            $rows = [int][Math]::Ceiling($shapes.Count / $cols)
            if ($rows -lt 1) { $rows = 1 }
            $cellWidth = ($slideWidth - 2 * $margin) / $cols
            $cellHeight = ($slideHeight - 2 * $margin) / $rows
            for ($idx = 0; $idx -lt $shapes.Count; $idx++) {
                $col = $idx % $cols
                $rowIdx = [int][Math]::Floor($idx / $cols)
                $row = $shapes[$idx]
                $shape = $row[0]; $w = [double]$row[1]; $h = [double]$row[2]
                $newLeft = [double]$margin + [double]$col * [double]$cellWidth + ([double]$cellWidth - $w) / 2
                $newTop = [double]$margin + [double]$rowIdx * [double]$cellHeight + ([double]$cellHeight - $h) / 2
                try { $shape.Left = [double]$newLeft } catch { $skipped++ }
                try { $shape.Top = [double]$newTop } catch { $skipped++ }
            }
        }

        Output-Json @{ success = $true; data = @{ layout = $layout; count = $shapes.Count; skipped = $skipped } }
    }

    "smartDistribute" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $shapeNames = @(if ($p.shapes) { $p.shapes })
        $shapes = @()

        if ($shapeNames.Count -gt 0) {
            foreach ($name in $shapeNames) {
                try { $shapes += $slide.Shapes.Item($name) } catch {}
            }
        } else {
            # 同 autoLayout：Shape.Type 与尺寸都可能抛（页面里有 3D/表格/图片时），
            # 收集阶段就逐个 try/catch，别让一个形状把整次分布拖垮。
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $null
                try { $s = $slide.Shapes.Item($i) } catch { continue }
                $t = -1
                try { $t = [int]$s.Type } catch { $t = -1 }
                if ($t -eq 1 -or $t -eq 6 -or $t -eq 17) {
                    $okSize = $true
                    try { $null = [double]$s.Width } catch { $okSize = $false }
                    if ($okSize) { try { $null = [double]$s.Height } catch { $okSize = $false } }
                    if ($okSize) { $shapes += $s }
                }
            }
        }

        if ($shapes.Count -lt 2) {
            Output-Json @{ success = $true; data = @{ message = "需要至少2个形状"; count = $shapes.Count } }
            return
        }

        # 不用 Sort-Object：COM 形状对象参与比较会抛"指定的转换无效"（只取到 1 个时返回的还不是数组）。
        # 自己按 [double] 键插入排序，键先落本地变量。
        $keys = @()
        foreach ($shp in $shapes) {
            $v = 0.0
            try { $v = if ($direction -eq "horizontal") { [double]$shp.Left } else { [double]$shp.Top } } catch { $v = 0.0 }
            $keys += , @($v, $shp)
        }
        for ($a = 1; $a -lt $keys.Count; $a++) {
            $cur = $keys[$a]; $b = $a - 1
            while ($b -ge 0 -and $keys[$b][0] -gt $cur[0]) { $keys[$b + 1] = $keys[$b]; $b-- }
            $keys[$b + 1] = $cur
        }
        $shapes = @()
        foreach ($kv in $keys) { $shapes += $kv[1] }

        # 与 autoLayout 同一个坑：COM 属性参与算术会抛"指定的转换无效"，统一先落 [double]
        if ($direction -eq "horizontal") {
            $minX = 0.0; try { $minX = [double]$shapes[0].Left } catch {}
            $maxX = $minX
            try { $last = $shapes[$shapes.Count - 1]; $maxX = [double]$last.Left + [double]$last.Width } catch {}
            $totalWidth = 0.0
            foreach ($shape in $shapes) { try { $totalWidth += [double]$shape.Width } catch {} }
            $spacing = ($maxX - $minX - $totalWidth) / [Math]::Max(1, $shapes.Count - 1)
            $currentX = $minX
            foreach ($shape in $shapes) {
                $w = 0.0; try { $w = [double]$shape.Width } catch {}
                try { $shape.Left = [double]$currentX } catch {}
                $currentX = $currentX + $w + $spacing
            }
        } else {
            $minY = 0.0; try { $minY = [double]$shapes[0].Top } catch {}
            $maxY = $minY
            try { $last = $shapes[$shapes.Count - 1]; $maxY = [double]$last.Top + [double]$last.Height } catch {}
            $totalHeight = 0.0
            foreach ($shape in $shapes) { try { $totalHeight += [double]$shape.Height } catch {} }
            $spacing = ($maxY - $minY - $totalHeight) / [Math]::Max(1, $shapes.Count - 1)
            $currentY = $minY
            foreach ($shape in $shapes) {
                $h = 0.0; try { $h = [double]$shape.Height } catch {}
                try { $shape.Top = [double]$currentY } catch {}
                $currentY = $currentY + $h + $spacing
            }
        }

        Output-Json @{ success = $true; data = @{ direction = $direction; count = $shapes.Count } }
    }

    "createGrid" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 2 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $margin = if ($p.margin) { $p.margin } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 15 }
        $slideWidth = 720
        $slideHeight = 540

        $cellWidth = ($slideWidth - 2 * $margin - ($cols - 1) * $gap) / $cols
        $cellHeight = ($slideHeight - 2 * $margin - ($rows - 1) * $gap) / $rows
        $cells = @()

        for ($r = 0; $r -lt $rows; $r++) {
            for ($c = 0; $c -lt $cols; $c++) {
                $left = $margin + $c * ($cellWidth + $gap)
                $top = $margin + $r * ($cellHeight + $gap)
                $cell = $slide.Shapes.AddShape(5, $left, $top, $cellWidth, $cellHeight)
                $cellColor = Convert-HexColorToRgbInt([string]$(if ($p.cellColor) { $p.cellColor } else { "#f8f9fa" }))
                $borderColor = Convert-HexColorToRgbInt([string]$(if ($p.borderColor) { $p.borderColor } else { "#dee2e6" }))
                if ($null -ne $cellColor) { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = $cellColor }
                if ($null -ne $borderColor) { $cell.Line.ForeColor.RGB = $borderColor }
                $cell.Line.Weight = 1
                $cells += $cell.Name
            }
        }

        Output-Json @{ success = $true; data = @{ rows = $rows; cols = $cols; cells = $cells } }
    }

    "addAnimationPreset" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $preset = if ($p.preset) { $p.preset } else { "fadeIn" }
        $timeline = $slide.TimeLine
        $delay = 0
        $delayIncrement = if ($p.delayIncrement) { $p.delayIncrement } else { 0.3 }
        $presets = @{ fadeIn = @{ effect = 10; duration = 0.5 }; flyIn = @{ effect = 2; duration = 0.5 }; zoomIn = @{ effect = 53; duration = 0.4 }; wipeIn = @{ effect = 22; duration = 0.5 }; appear = @{ effect = 1; duration = 0 } }
        $config = $presets[$preset]
        if ($null -eq $config) { $config = $presets["fadeIn"] }
        $animatedCount = 0

        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            try {
                $effect = $timeline.MainSequence.AddEffect($shape, $config.effect, 0, 1)
                $effect.Timing.Duration = $config.duration
                $effect.Timing.TriggerDelayTime = $delay
                $delay += $delayIncrement
                $animatedCount++
            } catch {}
        }

        Output-Json @{ success = $true; data = @{ preset = $preset; animatedShapes = $animatedCount } }
    }

    "addEmphasisAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $effectType = if ($p.effect) { $p.effect } else { "pulse" }
        $effectMap = @{ pulse = 63; spin = 15; grow = 53; teeter = 28 }
        $effectId = if ($effectMap[$effectType]) { $effectMap[$effectType] } else { 63 }

        $timeline = $slide.TimeLine
        $effect = $timeline.MainSequence.AddEffect($shape, $effectId, 0, 2)
        $effect.Timing.Duration = if ($p.duration) { $p.duration } else { 0.5 }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; effect = $effectType } }
    }

    "createFlowChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $steps = if ($p.steps) { $p.steps } else { @("开始", "步骤1", "步骤2", "结束") }
        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 250 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 60 }
        $colors = @{ start = "#28a745"; process = "#0d47a1"; end = "#dc3545" }
        $shapes = @()

        for ($i = 0; $i -lt $steps.Count; $i++) {
            if ($direction -eq "horizontal") {
                $x = $startX + $i * ($boxWidth + $gap)
                $y = $startY
            } else {
                $x = $startX
                $y = $startY + $i * ($boxHeight + $gap)
            }

            $shapeType = 5
            $colorKey = "process"
            if ($i -eq 0) { $shapeType = 9; $colorKey = "start" }
            elseif ($i -eq $steps.Count - 1) { $shapeType = 9; $colorKey = "end" }

            $shape = $slide.Shapes.AddShape($shapeType, $x, $y, $boxWidth, $boxHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$colors[$colorKey])
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.TextFrame.TextRange.Text = $steps[$i]
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 12
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $shapes += $shape.Name

            if ($i -lt $steps.Count - 1) {
                if ($direction -eq "horizontal") {
                    $arrowX1 = $x + $boxWidth
                    $arrowY1 = $y + $boxHeight / 2
                    $arrowX2 = $x + $boxWidth + $gap
                    $arrowY2 = $arrowY1
                } else {
                    $arrowX1 = $x + $boxWidth / 2
                    $arrowY1 = $y + $boxHeight
                    $arrowX2 = $arrowX1
                    $arrowY2 = $y + $boxHeight + $gap
                }
                $arrow = $slide.Shapes.AddLine($arrowX1, $arrowY1, $arrowX2, $arrowY2)
                $arrow.Line.EndArrowheadStyle = 2
                $arrow.Line.Weight = 2
                $arrowColor = Convert-HexColorToRgbInt([string]"#666666")
                if ($null -ne $arrowColor) { $arrow.Line.ForeColor.RGB = $arrowColor }
            }
        }

        Output-Json @{ success = $true; data = @{ steps = $steps.Count; shapes = $shapes } }
    }

    "createOrgChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $nodes = if ($p.nodes) { $p.nodes } else { @(@{ name = "CEO"; level = 0 }, @{ name = "技术总监"; level = 1 }, @{ name = "市场总监"; level = 1 }, @{ name = "财务总监"; level = 1 }) }
        $centerX = 360
        $startY = if ($p.startY) { $p.startY } else { 80 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 40 }
        $levelGap = if ($p.levelGap) { $p.levelGap } else { 80 }
        $levelColors = @("#1a365d", "#2d5a87", "#4a7ab0", "#6b9bd1")

        $levels = @{}
        foreach ($node in $nodes) {
            $level = if ($node.level -ne $null) { $node.level } else { 0 }
            if (-not $levels.ContainsKey($level)) { $levels[$level] = @() }
            $levels[$level] += $node
        }

        $levelKeys = $levels.Keys | Sort-Object
        $shapes = @()
        $levelIndex = 0
        foreach ($levelKey in $levelKeys) {
            $levelNodes = $levels[$levelKey]
            $levelY = $startY + $levelIndex * $levelGap
            $totalWidth = $levelNodes.Count * $boxWidth + ($levelNodes.Count - 1) * 30
            $startX = $centerX - $totalWidth / 2

            for ($i = 0; $i -lt $levelNodes.Count; $i++) {
                $x = $startX + $i * ($boxWidth + 30)
                $shape = $slide.Shapes.AddShape(5, $x, $levelY, $boxWidth, $boxHeight)
                $colorValue = Convert-HexColorToRgbInt([string]$levelColors[$levelIndex % $levelColors.Count])
                if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
                $shape.Line.Visible = $false
                $shape.TextFrame.TextRange.Text = $levelNodes[$i].name
                $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
                $shape.TextFrame.TextRange.Font.Size = 11
                $shape.TextFrame.TextRange.Font.Bold = $true
                $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $shape.Shadow.Visible = $true
                $shape.Shadow.Blur = 5
                $shape.Shadow.OffsetX = 2
                $shape.Shadow.OffsetY = 2
                $shape.Shadow.Transparency = 0.7
                $shapes += $shape.Name
            }

            $levelIndex++
        }

        Output-Json @{ success = $true; data = @{ levels = $levelKeys.Count; totalNodes = $nodes.Count; shapes = $shapes } }
    }

    "createTimeline" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $events = if ($p.events) { $p.events } else { @(@{ date = "2020"; title = "成立" }, @{ date = "2021"; title = "融资" }, @{ date = "2022"; title = "扩张" }, @{ date = "2023"; title = "上市" }) }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 280 }
        $lineLength = if ($p.lineLength) { $p.lineLength } else { 560 }

        $mainLine = $slide.Shapes.AddLine($startX, $startY, $startX + $lineLength, $startY)
        $lineColor = Convert-HexColorToRgbInt([string]$(if ($p.lineColor) { $p.lineColor } else { "#1a365d" }))
        if ($null -ne $lineColor) { $mainLine.Line.ForeColor.RGB = $lineColor }
        $mainLine.Line.Weight = 3

        $gap = if ($events.Count -gt 1) { $lineLength / ($events.Count - 1) } else { 0 }
        $shapes = @()

        for ($i = 0; $i -lt $events.Count; $i++) {
            $x = $startX + ($i * $gap)
            $above = ($i % 2 -eq 0)

            $dot = $slide.Shapes.AddShape(9, $x - 8, $startY - 8, 16, 16)
            $dotColor = Convert-HexColorToRgbInt([string]$(if ($p.dotColor) { $p.dotColor } else { "#0d47a1" }))
            if ($null -ne $dotColor) { $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = $dotColor }
            $dot.Line.Visible = $false

            $lineY = if ($above) { $startY - 50 } else { $startY + 50 }
            $connector = $slide.Shapes.AddLine($x, $startY, $x, $lineY)
            $connector.Line.Weight = 2
            $connectorColor = Convert-HexColorToRgbInt([string]"#cccccc")
            if ($null -ne $connectorColor) { $connector.Line.ForeColor.RGB = $connectorColor }

            $dateBox = $slide.Shapes.AddTextbox(1, $x - 30, $(if ($above) { $lineY - 45 } else { $lineY + 5 }), 60, 20)
            $dateBox.TextFrame.TextRange.Text = $events[$i].date
            $dateBox.TextFrame.TextRange.Font.Size = 12
            $dateBox.TextFrame.TextRange.Font.Bold = $true
            $dateBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $titleBox = $slide.Shapes.AddTextbox(1, $x - 40, $(if ($above) { $lineY - 25 } else { $lineY + 25 }), 80, 20)
            $titleBox.TextFrame.TextRange.Text = $events[$i].title
            $titleBox.TextFrame.TextRange.Font.Size = 10
            $titleColor = Convert-HexColorToRgbInt([string]"#666666")
            if ($null -ne $titleColor) { $titleBox.TextFrame.TextRange.Font.Color.RGB = $titleColor }
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $shapes += $dot.Name
        }

        Output-Json @{ success = $true; data = @{ events = $events.Count; shapes = $shapes } }
    }

    "getSlideMaster" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster
        $masters = @()
        for ($i = 1; $i -le $master.Shapes.Count; $i++) {
            $shape = $master.Shapes.Item($i)
            $masters += @{ name = $shape.Name; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ shapeCount = $masters.Count; shapes = $masters } }
    }

    "setMasterBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster
        if ($p.gradient) {
            $master.Background.Fill.TwoColorGradient(1, 1)
            $color1 = Convert-HexColorToRgbInt([string]$(if ($p.color1) { $p.color1 } else { "#1a365d" }))
            $color2 = Convert-HexColorToRgbInt([string]$(if ($p.color2) { $p.color2 } else { "#2d5a87" }))
            if ($null -ne $color1) { $master.Background.Fill.GradientStops.Item(1).Color.RGB = $color1 }
            if ($null -ne $color2) { $master.Background.Fill.GradientStops.Item(2).Color.RGB = $color2 }
        } else {
            $master.Background.Fill.Solid()
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#ffffff" }))
            if ($null -ne $colorValue) { $master.Background.Fill.ForeColor.RGB = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ message = "母版背景已更新" } }
    }

    "addMasterElement" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster

        $type = if ($p.type) { $p.type } else { "textbox" }
        $left = if ($p.left) { $p.left } else { 600 }
        $top = if ($p.top) { $p.top } else { 20 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 40 }

        if ($type -eq "textbox") {
            $shape = $master.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "Logo" }
            $shape.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 14 }
            $shape.TextFrame.TextRange.Font.Bold = $true
        } elseif ($type -eq "shape") {
            $shape = $master.Shapes.AddShape($(if ($p.shapeType) { $p.shapeType } else { 5 }), $left, $top, $width, $height)
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
        } else {
            Output-Json @{ success = $false; error = "Unsupported master element type" }
            break
        }

        Output-Json @{ success = $true; data = @{ name = $shape.Name; type = $type } }
    }

    "set3DRotation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        # 工具传的是 rotation:{x,y,z} 对象，而这里读的是扁平 rotationX/Y/Z ——
        # 三种形状都认，否则整包 psobject 会被塞进 float 属性（实测报"无法将 psobject 转换为类型 float"）。
        $rot = $p.rotation
        $rx = if ($null -ne $p.rotationX) { $p.rotationX } elseif ($null -ne $rot -and $null -ne $rot.x) { $rot.x } else { 0 }
        $ry = if ($null -ne $p.rotationY) { $p.rotationY } elseif ($null -ne $rot -and $null -ne $rot.y) { $rot.y } else { 0 }
        $rz = if ($null -ne $p.rotationZ) { $p.rotationZ } elseif ($null -ne $rot -and $null -ne $rot.z) { $rot.z } else { 0 }
        $shape.ThreeD.RotationX = [float]$rx
        $shape.ThreeD.RotationY = [float]$ry
        $shape.ThreeD.RotationZ = [float]$rz

        if ($p.preset) {
            $presets = @{ isometric = @{ x = 45; y = 45; z = 0 }; perspective = @{ x = 30; y = 30; z = 0 }; oblique = @{ x = 20; y = 60; z = 0 }; tiltLeft = @{ x = 0; y = -30; z = 0 }; tiltRight = @{ x = 0; y = 30; z = 0 } }
            $preset = $presets[$p.preset]
            if ($preset) {
                $shape.ThreeD.RotationX = $preset.x
                $shape.ThreeD.RotationY = $preset.y
                $shape.ThreeD.RotationZ = $preset.z
            }
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; rotationX = $shape.ThreeD.RotationX; rotationY = $shape.ThreeD.RotationY } }
    }

    "set3DDepth" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $depth = if ($p.depth) { $p.depth } else { 20 }
        $shape.ThreeD.Depth = $depth
        if ($p.depthColor) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.depthColor)
            if ($null -ne $colorValue) { $shape.ThreeD.ExtrusionColor.RGB = $colorValue }
        }
        if ($p.lighting) {
            $lightingMap = @{ bright = 1; normal = 2; dim = 3; flat = 4 }
            $lightingValue = if ($lightingMap[$p.lighting]) { $lightingMap[$p.lighting] } else { 2 }
            $shape.ThreeD.PresetLighting = $lightingValue
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; depth = $depth } }
    }

    "set3DMaterial" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $material = if ($p.material) { $p.material } else { "plastic" }
        $materialMap = @{ matte = 1; plastic = 2; metal = 3; wireFrame = 4; glass = 5 }
        $materialValue = if ($materialMap[$material]) { $materialMap[$material] } else { 2 }
        $shape.ThreeD.PresetMaterial = $materialValue

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; material = $material } }
    }

    "create3DText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $text = if ($p.text) { $p.text } else { "3D文字" }
        $left = if ($p.left) { $p.left } else { 200 }
        $top = if ($p.top) { $p.top } else { 200 }
        $width = if ($p.width) { $p.width } else { 300 }
        $height = if ($p.height) { $p.height } else { 100 }
        $textBox = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $textBox.TextFrame.TextRange.Text = $text
        $textBox.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 48 }
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        if ($null -ne $colorValue) { $textBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        $textBox.ThreeD.RotationX = if ($p.rotationX -ne $null) { $p.rotationX } else { 15 }
        $textBox.ThreeD.RotationY = if ($p.rotationY -ne $null) { $p.rotationY } else { 30 }
        $textBox.ThreeD.Depth = if ($p.depth) { $p.depth } else { 10 }

        $textBox.Shadow.Visible = $true
        $textBox.Shadow.Blur = 8
        $textBox.Shadow.OffsetX = 5
        $textBox.Shadow.OffsetY = 5
        $textBox.Shadow.Transparency = 0.5

        Output-Json @{ success = $true; data = @{ name = $textBox.Name; text = $text } }
    }

    "setPptChartData" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        $chart.ChartData.Activate()
        $dataSheet = $chart.ChartData.Workbook.Worksheets.Item(1)
        if ($p.data) {
            # 参数形状归一化（第 3 轮修"假成功"）：
            # TS 侧 schema 承诺 {categories:[...], series:[{name, values:[...]}]}（文档是给用户看的承诺），
            # 历史裸调用（wps_execute_method 直穿）传的是二维数组。两个形状都接。
            # 第 4 个坑：数字从 JSON 过来是 Int32，直接塞 Value2 会报 Int32→String 类型错 ——
            # 数值统一转 [double]，字符串才按字符串写。
            $rows = @()
            if ($p.data -is [System.Array]) {
                # 二维数组（历史形状）：原样逐格写
                foreach ($r in $p.data) { $rows += ,@($r) }
            } elseif ($p.data.PSObject.Properties.Name -contains 'categories' -or $p.data.PSObject.Properties.Name -contains 'series') {
                # 文档承诺的形状：摊成"表头行 + 每类别一行"
                $cats = @()
                if ($p.data.categories) { $cats = @($p.data.categories) }
                $seriesList = @()
                if ($p.data.series) { $seriesList = @($p.data.series) }
                $header = @('类别')
                foreach ($s in $seriesList) { $header += [string]$s.name }
                $rows += ,@($header)
                for ($i = 0; $i -lt $cats.Count; $i++) {
                    $line = @([string]$cats[$i])
                    foreach ($s in $seriesList) {
                        $vals = @($s.values)
                        $v = if ($i -lt $vals.Count) { $vals[$i] } else { $null }
                        $line += $v
                    }
                    $rows += ,$line
                }
            }
            for ($r = 0; $r -lt $rows.Count; $r++) {
                $rowData = $rows[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    $cell = $dataSheet.Cells.Item($r + 1, $c + 1)
                    $v = $rowData[$c]
                    if ($null -eq $v) { continue }
                    if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                        $cell.Value2 = [double]$v
                    } else {
                        $cell.Value2 = [string]$v
                    }
                }
            }
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "setPptChartStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        # 参数形状归一化（第 3 轮修"假成功"）：
        # TS 侧 schema 承诺 style:{title,showLegend,legendPosition,showDataLabels,fontSize}，
        # 历史裸调用用扁平 $p.title / $p.hasLegend。两个都接，扁平优先。
        # 注意 schema 写的是 showLegend，PS 老代码读的是 hasLegend —— 两个键都认。
        $st = @{}
        if ($p.style) {
            foreach ($k in $p.style.PSObject.Properties.Name) { $st[$k] = $p.style.$k }
        }
        foreach ($k in @('title', 'hasLegend', 'showLegend', 'legendPosition', 'showDataLabels', 'fontSize', 'colorScheme')) {
            if ($p.PSObject.Properties.Name -contains $k -and -not $st.ContainsKey($k)) { $st[$k] = $p.$k }
        }
        $applied = @()
        if ($st.title) {
            try {
                $chart.HasTitle = $true
                $chart.ChartTitle.Text = [string]$st.title
                $applied += 'title'
            } catch { }
        }
        $legendFlag = $null
        if ($null -ne $st.showLegend) { $legendFlag = $st.showLegend }
        elseif ($null -ne $st.hasLegend) { $legendFlag = $st.hasLegend }
        if ($null -ne $legendFlag) {
            try {
                $chart.HasLegend = [bool]$legendFlag
                $applied += 'legend'
            } catch { }
        }
        if ($st.legendPosition) {
            # 图例位置枚举与 Excel 侧 updateChart（wps-com.ps1 的 updateChart case）保持一致
            $legMap = @{ bottom = -4107; top = -4160; left = -4131; right = -4152 }
            $lp = $legMap[[string]$st.legendPosition]
            if ($null -ne $lp) {
                try {
                    $chart.Legend.Position = $lp
                    $applied += 'legendPosition'
                } catch { }
            }
        }
        if ($null -ne $st.showDataLabels) {
            # 只开不关：关标签要逐系列 Delete，失败率高；调用方要关就别传这个键
            if ([bool]$st.showDataLabels) {
                try {
                    $chart.ApplyDataLabels()
                    $applied += 'dataLabels'
                } catch { }
            }
        }
        if ($st.fontSize) {
            try {
                $chart.ChartArea.Font.Size = $st.fontSize
                $applied += 'fontSize'
            } catch { }
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name; applied = $applied } }
    }

    "removeAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        # 工具传的是 animationIndex，这里以前只读 index → 落进 else 分支把【所有动画都删了】，
        # 计数变化看起来"成功"，但删多删少都不受控。两个键都认，并做越界检查。
        $idx = if ($null -ne $p.animationIndex) { [int]$p.animationIndex } elseif ($null -ne $p.index) { [int]$p.index } else { 0 }
        if ($seq.Count -eq 0) {
            Output-Json @{ success = $true; data = @{ removed = 0; note = "本页本来就没有动画" } }
            return
        }
        if ($idx -ge 1) {
            if ($idx -gt $seq.Count) {
                Output-Json @{ success = $false; error = "[勿重试·先改参数] 动画序号 $idx 越界：本页共 $($seq.Count) 个（1 起）" }
                return
            }
            $seq.Item($idx).Delete()
        } else {
            while ($seq.Count -gt 0) { $seq.Item(1).Delete() }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "getAnimations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $animations = @()
        for ($i = 1; $i -le $seq.Count; $i++) {
            $effect = $seq.Item($i)
            $animations += @{ index = $i; shapeName = $effect.Shape.Name; effectType = $effect.EffectType; duration = $effect.Timing.Duration }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; animations = $animations; count = $animations.Count } }
    }

    "setAnimationOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        # 工具传的是 animationIndex/newOrder，这里以前只读 from/to → Item($null) 必报"值不在预期的范围内"
        $fromIdx = if ($null -ne $p.animationIndex) { [int]$p.animationIndex } elseif ($null -ne $p.from) { [int]$p.from } else { 0 }
        $toIdx = if ($null -ne $p.newOrder) { [int]$p.newOrder } elseif ($null -ne $p.to) { [int]$p.to } else { 0 }
        $cnt = $seq.Count
        if ($cnt -eq 0) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 本页还没有动画，谈不上排序；先 add_animation" }
            return
        }
        if ($fromIdx -lt 1 -or $fromIdx -gt $cnt -or $toIdx -lt 1 -or $toIdx -gt $cnt) {
            Output-Json @{ success = $false; error = "[勿重试·先改参数] 动画序号越界：本页共 $cnt 个（1 起），收到 from=$fromIdx to=$toIdx" }
            return
        }
        $effect = $seq.Item($fromIdx)
        $effect.MoveTo($toIdx)
        Output-Json @{ success = $true; data = @{ from = $fromIdx; to = $toIdx } }
    }

    "removeSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = 0
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "applyTransitionToAll" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        # 参数名是 effect（合并后的 set_transition target='all' 就是它）；
        # 以前这里只读 $p.transition，工具传的一直是 effect → 映射表查不到 → 静默全设成 fade。
        if ($null -eq $p.effect -and $p.transition) { $p | Add-Member -NotePropertyName effect -NotePropertyValue $p.transition -Force }
        $effectName = if ($p.effect) { ([string]$p.effect).ToLower() } else { "fade" }
        $transitionMap = @{ none = 0; fade = 1; push = 2; wipe = 3; split = 4; reveal = 5; cover = 6; dissolve = 7; curtains = 8; random = 6; blinds = 3 }
        $transition = -1
        if ($transitionMap.ContainsKey($effectName)) { $transition = $transitionMap[$effectName] }
        elseif ($effectName -match '^\d+$') { $transition = [int]$effectName }
        if ($transition -lt 0) {
            Output-Json @{ success = $false; error = "不认识的切换效果：$effectName（可用：none/fade/push/wipe/split/reveal/cover/dissolve/curtains）" }
            return
        }
        $dur = 0.0
        if ($p.duration) { try { $dur = [double]$p.duration } catch { $dur = 0.0 } }
        $applied = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $tr = $slide.SlideShowTransition
            $tr.EntryEffect = $transition
            if ($dur -gt 0) { try { $tr.Duration = $dur } catch {} }
            $applied++
        }
        # 回读字段名与 setSlideTransition 对齐（slideCount）；appliedTo 留着兼容老调用方
        Output-Json @{ success = $true; data = @{ effect = $effectName; entryEffect = $transition; duration = $dur; slideCount = $applied; appliedTo = $applied } }
    }

    "setBackgroundColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $colorValue = Convert-HexColorToRgbInt([string]$p.color)
        if ($p.applyToAll) {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $slide.FollowMasterBackground = $false
                $slide.Background.Fill.Solid()
                if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            }
            Output-Json @{ success = $true; data = @{ color = $p.color; appliedTo = $pres.Slides.Count } }
        } else {
            $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
            $slide = $pres.Slides.Item($slideIndex)
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.Solid()
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; color = $p.color } }
        }
    }

    "setBackgroundImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.FollowMasterBackground = $false
        $slide.Background.Fill.UserPicture($p.path)
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; path = $p.path } }
    }

    "addPptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $actionSettings = $shape.ActionSettings.Item(1)
        $actionSettings.Hyperlink.Address = if ($p.address) { $p.address } else { "" }
        if ($p.subAddress) { $actionSettings.Hyperlink.SubAddress = $p.subAddress }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name; address = $p.address } }
    }

    "removePptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $shape.ActionSettings.Item(1).Hyperlink.Address = ""
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "setSlideNumber" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -ne $p.visible) { $pres.SlideMaster.HeadersFooters.SlideNumber.Visible = $p.visible }
        Output-Json @{ success = $true; data = @{ visible = $p.visible } }
    }

    "setPptFooter" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideMaster.HeadersFooters.Footer.Visible = $true
        $pres.SlideMaster.HeadersFooters.Footer.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ text = $p.text } }
    }

    "setPptDateTime" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideMaster.HeadersFooters.DateAndTime.Visible = if ($p.visible -ne $null) { $p.visible } else { $true }
        if ($p.useFixed) {
            $pres.SlideMaster.HeadersFooters.DateAndTime.UseFormat = $false
            $pres.SlideMaster.HeadersFooters.DateAndTime.Text = if ($p.text) { $p.text } else { "" }
        }
        Output-Json @{ success = $true; data = @{ visible = if ($p.visible -ne $null) { $p.visible } else { $true } } }
    }

    "findPptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $results = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $text = $shape.TextFrame.TextRange.Text
                        if ($text -and $text.Contains($p.text)) {
                            $results += @{ slideIndex = $i; shapeName = $shape.Name; text = $text }
                        }
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.text; results = $results; count = $results.Count } }
    }

    "replacePptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation" }; return }
        # 兼容两种参数名：TS 工具发送 find/replace，旧实现读 findText/replaceText（此前不匹配导致 0 替换）
        $findText = if ($null -ne $p.find) { [string]$p.find } else { [string]$p.findText }
        $replaceText = if ($null -ne $p.replace) { [string]$p.replace } elseif ($null -ne $p.replaceText) { [string]$p.replaceText } else { "" }
        if ([string]::IsNullOrEmpty($findText)) { Output-Json @{ success = $false; error = "find text required" }; return }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                # 递归处理：文本框 + 表格单元格 + 组合(Group)内所有子形状
                $count += Replace-InShapeTree $shape $findText $replaceText
            }
        }
        Output-Json @{ success = $true; data = @{ findText = $findText; replaceText = $replaceText; count = $count } }
    }

    "startSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideShowSettings.Run()
        Output-Json @{ success = $true }
    }

    "endSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        if ($ppt.SlideShowWindows.Count -gt 0) { $ppt.SlideShowWindows.Item(1).View.Exit() }
        Output-Json @{ success = $true }
    }

    "unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $shape.TextFrame.TextRange.Font.Name = $fontName
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ fontName = $fontName; count = $count } }
    }

    "slide.beautify" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $schemeKey = if ($p.style -and $p.style.colorScheme) { $p.style.colorScheme } elseif ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $fontName = if ($p.style -and $p.style.font) { $p.style.font } else { $null }
        $slideIndex = $p.slideIndex
        $slides = @()
        if ($slideIndex -eq "all") {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) { $slides += $pres.Slides.Item($i) }
        } else {
            $target = if ($slideIndex) { [int]$slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
            $slides += $pres.Slides.Item($target)
        }
        $ops = @(
            @{ operation = "unify_font"; count = 0 },
            @{ operation = "apply_color_scheme"; count = 0 },
            @{ operation = "align"; count = 0 },
            @{ operation = "optimize_spacing"; count = 0 }
        )
        foreach ($slide in $slides) {
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                        else { $textRange.Font.Color.RGB = $scheme.body }
                        $ops[1].count++
                        if ($fontName) { $textRange.Font.Name = $fontName; $ops[0].count++ }
                    }
                } catch {}
            }
        }
        $slideIndexOut = if ($slideIndex -eq "all") { "all" } else { ($slides | Select-Object -First 1).SlideIndex }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndexOut; operations = $ops } }
    }

    # 第 6 轮重写。原来的三处病：
    #   ① 回包只有 {style,count}，而 TS 读 data.operations 并 .forEach → 直接崩
    #      "Cannot read properties of undefined (reading 'forEach')"（冒烟里 79 条失败的第一条）；
    #   ② TS 传的是嵌套 style:{colorScheme,font}，PS 却把 $p.style 当字符串查配色表
    #      → 永远落到 business，配色方案参数形同摆设；
    #   ③ slideIndex 传 'all'（美化整份）时 $pres.Slides.Item('all') 必炸。
    # 现在：配色/字体两种形状都认、支持 all、把【真做了哪几件事】按 operations 数组回给 TS。
    # 说明：描述里承诺的"对齐元素/优化间距"没有实现，所以不往 operations 里塞 —— 假报一条
    # 比少报一条更坏（用户会以为版面被整理过）。
    "beautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }

        # 参数归一：嵌套 style{} 与扁平两套都收（TS 发的是嵌套）
        $schemeKey = if ($p.colorScheme) { [string]$p.colorScheme } elseif ($p.style -and ($p.style.PSObject.Properties.Name -contains 'colorScheme')) { [string]$p.style.colorScheme } elseif ($p.style -is [string]) { [string]$p.style } else { 'business' }
        $fontName = if ($p.font) { [string]$p.font } elseif ($p.style -and ($p.style.PSObject.Properties.Name -contains 'font')) { [string]$p.style.font } else { '微软雅黑' }

        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"]; $schemeKey = 'business' }

        # 范围：'all' 或没给（以当前页为准）
        $allSlides = ([string]$p.slideIndex -eq 'all')
        $startIdx = 1
        $endIdx = $pres.Slides.Count
        if (-not $allSlides) {
            if ($p.slideIndex) { $startIdx = [int]$p.slideIndex; $endIdx = $startIdx }
            else {
                try { $cur = $ppt.ActiveWindow.Selection.SlideRange.SlideIndex; if ($cur) { $startIdx = $cur; $endIdx = $cur } } catch {}
            }
        }

        $fontCount = 0
        $colorCount = 0
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        $textRange.Font.Name = $fontName
                        $fontCount++
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                        else { $textRange.Font.Color.RGB = $scheme.body }
                        $colorCount++
                    }
                } catch {}
            }
        }

        $operations = @(
            @{ operation = 'unify_font'; count = $fontCount },
            @{ operation = 'apply_color_scheme'; count = $colorCount }
        )
        Output-Json @{ success = $true; data = @{
            slideIndex = $(if ($allSlides) { 'all' } else { $startIdx })
            slideCount = $endIdx - $startIdx + 1
            style = $schemeKey
            font = $fontName
            count = $colorCount
            operations = $operations
        } }
    }

    # ==================== Word 长文档 / 学术排版（补 office-mcp 那条缺失链路）====================
    # 这一组是为了补齐"认真写一篇论文/报告"所需的能力：标题层级、题注、脚注、图表目录、
    # 章节自动编号、分栏、列表、表格样式、刷新域、大纲抽取。
    # 口径：每个动作都先拿 $doc，拿不到就如实报错返回，不做半截操作。

    # 按级别插入标题：用 Word 内置 Heading 样式（wdStyleHeadingN = -1-N），保证能进目录
    "addHeading" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $level = if ($p.level) { [int]$p.level } else { 1 }
        if ($level -lt 1) { $level = 1 }
        if ($level -gt 9) { $level = 9 }
        $text = [string]$p.text
        try {
            $word.Selection.Style = $doc.Styles.Item(-1 - $level)
            $word.Selection.TypeText($text)
            $word.Selection.TypeParagraph()
            Output-Json @{ success = $true; data = @{ level = $level; text = $text } }
        } catch {
            Output-Json @{ success = $false; error = "插入标题失败: $(Get-CallError $_)" }
        }
    }

    # 插入段落：给了 index 就在该段之前插，没给就追加到文末
    "insertParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = [string]$p.text
        if ($p.index) {
            $idx = [int]$p.index
            if ($idx -lt 1) { $idx = 1 }
            if ($idx -gt $doc.Paragraphs.Count) { $idx = $doc.Paragraphs.Count }
            $doc.Paragraphs.Item($idx).Range.InsertBefore($text + "`r")
        } else {
            $doc.Content.InsertAfter($text + "`r")
        }
        Output-Json @{ success = $true; data = @{ paragraphs = $doc.Paragraphs.Count } }
    }

    # 删除段落：从 index 起删 count 段（1 基，越界自动收敛）
    "deleteParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $idx = if ($p.index) { [int]$p.index } else { $doc.Paragraphs.Count }
        $cnt = if ($p.count) { [int]$p.count } else { 1 }
        if ($idx -lt 1) { $idx = 1 }
        $deleted = 0
        for ($k = 0; $k -lt $cnt; $k++) {
            if ($idx -gt $doc.Paragraphs.Count) { break }
            try { $doc.Paragraphs.Item($idx).Range.Delete(); $deleted++ } catch { break }
        }
        Output-Json @{ success = $true; data = @{ deleted = $deleted; paragraphs = $doc.Paragraphs.Count } }
    }

    # 抽取大纲：按段落的 OutlineLevel（1..9 是标题级，10 是正文）列出标题树
    "getOutline" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $limit = if ($p.limit) { [int]$p.limit } else { 500 }
        $outline = @()
        $n = [int]$doc.Paragraphs.Count
        for ($i = 1; $i -le $n; $i++) {
            if ($outline.Count -ge $limit) { break }
            try {
                $para = $doc.Paragraphs.Item($i)
                $lvl = [int]$para.OutlineLevel
                if ($lvl -ge 1 -and $lvl -le 9) {
                    $t = [string]$para.Range.Text
                    $t = $t -replace "[`r`n`a]", ""
                    if ($t.Trim()) {
                        $outline += @{ index = $i; level = $lvl; text = $t.Trim() }
                    }
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ headings = $outline; count = $outline.Count } }
    }

    # 默认字体：改 Normal 样式的中文字体/西文字体/字号，正文与新建段落都跟着变
    "setDefaultFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $normal = $doc.Styles.Item(-1)   # wdStyleNormal
            if ($p.eastAsia) { $normal.Font.NameFarEast = [string]$p.eastAsia }
            if ($p.ascii) { $normal.Font.NameAscii = [string]$p.ascii }
            if ($p.name) { $normal.Font.Name = [string]$p.name }
            if ($p.size) { $normal.Font.Size = [double]$p.size }
            Output-Json @{ success = $true; data = @{
                name = [string]$normal.Font.Name
                eastAsia = [string]$normal.Font.NameFarEast
                size = [double]$normal.Font.Size
            } }
        } catch {
            Output-Json @{ success = $false; error = "设置默认字体失败: $(Get-CallError $_)" }
        }
    }

    # 题注：给图表/表格加自动编号标签（Word 的 Caption 会自动续号）
    "addCaption" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $label = if ($p.label) { [string]$p.label } else { "图" }
        # wdCaptionPositionAbove = 0, wdCaptionPositionBelow = 1
        $pos = if ($p.position -eq "above") { 0 } else { 1 }
        $title = [string]$p.text
        $mode = $null
        $captionText = $null
        $nativeErr = $null

        # ① 先试原生题注：它会插入可见编号域，之后"图表目录"能自动收集到
        try {
            $sel = $word.Selection
            $sel.Collapse(0)     # wdCollapseEnd
            $sel.InsertCaption($label, $title, $pos, $false) | Out-Null
            $mode = "field"
        } catch {
            $nativeErr = $_.Exception.Message
        }

        # ② 原生不支持就手工兜底：纯文本题注 + 自增序号。
        #    实测本机 WPS 12.1 的 Range/Selection.InsertCaption 都会抛"值不在预期的范围内"
        #    （CaptionLabels 集合也不可用），所以必须有这条兜底，否则这个功能直接废掉。
        #    代价：序号是写死的文本，不是活编号域 —— 图表目录收集不到，已在返回值里如实标注。
        if ($null -eq $mode) {
            try {
                $n = 0
                foreach ($para in $doc.Paragraphs) {
                    $t = [string]$para.Range.Text
                    if ($t -match ("^\s*" + [regex]::Escape($label) + "\s+\d+")) { $n++ }
                }
                $n = $n + 1
                $captionText = "$label $n $title"
                $rng = $word.Selection.Range
                $rng.Collapse(0)
                $rng.InsertAfter($captionText + "`r")
                $mode = "manual"
            } catch {
                Output-Json @{ success = $false; error = "插入题注失败：原生接口报「$nativeErr」；手工兜底也失败：$(Get-CallError $_)" }
                return
            }
        }

        # 降级插题注属于"主操作做完了但有保留" —— 走 warnings 约定，
        # 让上层渲染成 [部分成功|...]，模型不会当成原生域插成功了。
        $capWarnings = @()
        if ($mode -eq "manual") {
            $capWarnings = @(@{
                step = "insertCaption.原生接口"
                message = "原生题注接口在本机 WPS 不可用，已按标签自增序号手工插入；序号是文本不是域，图表目录无法自动收集"
            })
        }
        Output-Json @{ success = $true; data = @{
            label = $label
            text = $title
            mode = $mode
            caption = $captionText
            note = if ($mode -eq "manual") { $capWarnings[0].message } else { $null }
            warnings = $capWarnings
            nativeError = $nativeErr
        } }
    }

    # 脚注：在当前选区（或文末）插入一条脚注
    "addFootnote" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = [string]$p.text
        try {
            $rng = if ($p.atEnd) { $doc.Content } else { $word.Selection.Range }
            $fn = $doc.Footnotes.Add($rng, $text)
            Output-Json @{ success = $true; data = @{ total = [int]$doc.Footnotes.Count; index = [int]$fn.Index } }
        } catch {
            Output-Json @{ success = $false; error = "插入脚注失败: $(Get-CallError $_)" }
        }
    }

    # 图表目录：基于题注标签生成（要先把题注加上才有内容）
    "insertTableOfFigures" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $label = if ($p.label) { [string]$p.label } else { "图" }
        try {
            $rng = $doc.Content
            $rng.Collapse(0)   # wdCollapseEnd → 放文末
            $tof = $doc.TablesOfFigures.Add($rng, $label)
            try { $tof.Update() } catch {}
            Output-Json @{ success = $true; data = @{ label = $label; total = [int]$doc.TablesOfFigures.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入图表目录失败: $(Get-CallError $_)" }
        }
    }

    # 刷新域：目录、图表目录、交叉引用、页码等一次性更新
    "updateFields" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $report = @{ fields = 0; toc = 0; tof = 0; errors = @() }
        try { $doc.Fields.Update(); $report.fields = [int]$doc.Fields.Count } catch { $report.errors += "fields: $(Get-CallError $_)" }
        try { foreach ($t in $doc.TablesOfContents) { $t.Update() }; $report.toc = [int]$doc.TablesOfContents.Count } catch { $report.errors += "toc: $(Get-CallError $_)" }
        try { foreach ($f in $doc.TablesOfFigures) { $f.Update() }; $report.tof = [int]$doc.TablesOfFigures.Count } catch { $report.errors += "tof: $(Get-CallError $_)" }
        Output-Json @{ success = $true; data = $report }
    }

    # 分栏：整节等宽多栏
    "setColumns" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $count = if ($p.count) { [int]$p.count } else { 2 }
        if ($count -lt 1) { $count = 1 }
        if ($count -gt 6) { $count = 6 }
        try {
            $doc.PageSetup.TextColumns.SetCount($count)
            if ($p.spacing) {
                for ($i = 1; $i -le $count; $i++) {
                    try { $doc.PageSetup.TextColumns.Item($i).Spacing = [double]$p.spacing * 28.35 } catch {}
                }
            }
            if ($null -ne $p.lineBetween) {
                try { $doc.PageSetup.TextColumns.LineBetween = [bool]$p.lineBetween } catch {}
            }
            Output-Json @{ success = $true; data = @{ columns = [int]$doc.PageSetup.TextColumns.Count } }
        } catch {
            # 分栏在不同 WPS 版本上接口差异较大 —— 失败就明说，别伪装成功
            Output-Json @{ success = $false; error = "设置分栏失败: $(Get-CallError $_)" }
        }
    }

    # 章节自动编号：给 Heading1~N 挂多级编号（近似实现，版本差异可能导致效果不同）
    "setHeadingNumbering" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $maxLevel = if ($p.levels) { [int]$p.levels } else { 3 }
        if ($maxLevel -lt 1) { $maxLevel = 1 }
        if ($maxLevel -gt 9) { $maxLevel = 9 }
        try {
            # 实测：`ListTemplates.Add(...)` 在本机 WPS 12.1 返回 null（"不能对 Null 值表达式调用方法"），
            # 所以不用"新建列表模板"，改成直接借用【大纲编号库】里现成的模板来挂 Heading 样式 ——
            # 这是 Word 里的经典做法，兼容性明显更好。
            # wdOutlineNumberGallery = 2
            $tpl = $word.ListGalleries.Item(2).ListTemplates.Item(1)
            for ($i = 1; $i -le $maxLevel; $i++) {
                $doc.Styles.Item(-1 - $i).LinkToListTemplate($tpl, $i)
            }
            Output-Json @{ success = $true; data = @{ levels = $maxLevel; mode = "outline-gallery" } }
        } catch {
            Output-Json @{ success = $false; error = "设置章节编号失败（本机 WPS 可能不支持该接口；可在 WPS 里用「开始 → 多级列表」手动设置）: $(Get-CallError $_)" }
        }
    }

    # 列表：项目符号或编号，一次写多条
    "addList" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $items = @($p.items)
        if ($items.Count -eq 0) { Output-Json @{ success = $false; error = "items 为空" }; return }
        $type = if ($p.type -eq "number") { "number" } else { "bullet" }
        try {
            $sel = $word.Selection
            if ($type -eq "number") { $sel.Range.ListFormat.ApplyNumberDefault() }
            else { $sel.Range.ListFormat.ApplyBulletDefault() }
            foreach ($it in $items) {
                $sel.TypeText([string]$it)
                $sel.TypeParagraph()
            }
            # 收尾：去掉列表末那个空项目符号，回到正文
            try { $sel.Range.ListFormat.RemoveNumbers() } catch {}
            Output-Json @{ success = $true; data = @{ type = $type; count = $items.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入列表失败: $(Get-CallError $_)" }
        }
    }

    # 表格样式：套用样式名 + 表头加粗 + 边框
    "formatTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $idx = if ($p.tableIndex) { [int]$p.tableIndex } else { 1 }
        if ($idx -lt 1 -or $idx -gt $doc.Tables.Count) {
            Output-Json @{ success = $false; error = "表格序号越界（共 $($doc.Tables.Count) 个）" }; return
        }
        $applied = @{ style = $false; headerBold = $false; borders = $false }
        try {
            $tbl = $doc.Tables.Item($idx)
            if ($p.style) {
                try { $tbl.Style = [string]$p.style; $applied.style = $true } catch {}
            }
            if ($p.headerBold) {
                try { $tbl.Rows.Item(1).Range.Font.Bold = $true; $applied.headerBold = $true } catch {}
            }
            if ($p.borders) {
                try {
                    $tbl.Borders.InsideLineStyle = 1     # wdLineStyleSingle
                    $tbl.Borders.OutsideLineStyle = 1
                    $applied.borders = $true
                } catch {}
            }
            Output-Json @{ success = $true; data = @{ tableIndex = $idx; applied = $applied } }
        } catch {
            Output-Json @{ success = $false; error = "设置表格样式失败: $(Get-CallError $_)" }
        }
    }

    # 【诊断】故意占住这一路一段时间 —— 用来实测排队、超时、取消这些时序行为。
    # 没有它，"队列上限有没有生效""取消后 inFlight 会不会提前变 false"这类问题
    # 只能靠猜，因为本机动作都太快，观察不到中间状态。
    "__sleep" {
        $ms = 1000
        try { if ($p.ms) { $ms = [int]$p.ms } } catch {}
        if ($ms -lt 0) { $ms = 0 }
        if ($ms -gt 600000) { $ms = 600000 }
        Start-Sleep -Milliseconds $ms
        Output-Json @{ success = $true; data = @{ slept = $ms } }
    }

    # 同上，但被 TS 侧判为【只读】（见 bridge/cancellable.ts 的白名单）。
    # 之所以要两个：取消机制对"已在途"的请求是按只读/写入分岔的，
    # 两条路都得能实测，只留一个就有一条永远验不到。
    "__sleepRead" {
        $ms = 1000
        try { if ($p.ms) { $ms = [int]$p.ms } } catch {}
        if ($ms -lt 0) { $ms = 0 }
        if ($ms -gt 600000) { $ms = 600000 }
        Start-Sleep -Milliseconds $ms
        Output-Json @{ success = $true; data = @{ slept = $ms } }
    }

    # 【临时诊断】原样回显收到的参数，用来判断"参数到底有没有送到"
    "__probeParams" {
        $keys = @()
        try { $keys = @($p.PSObject.Properties | ForEach-Object { $_.Name }) } catch {}
        Output-Json @{ success = $true; data = @{ received = $p; keys = $keys; ptype = $p.GetType().FullName } }
    }

    # 【临时诊断】查 WPS 文字的 COM 状态（进程/文档数/ActiveDocument 是否可用）
    "__probeDoc" {
        $info = @{}
        try { $word = Get-WpsWord } catch { $info.getWordErr = $_.Exception.Message; $word = $null }
        $info.wordNull = ($null -eq $word)
        if ($null -ne $word) {
            try { $info.appName = [string]$word.Name } catch { $info.appNameErr = $_.Exception.Message }
            try { $info.appPath = [string]$word.Path } catch { $info.appPathErr = $_.Exception.Message }
            try { $info.docCount = [int]$word.Documents.Count } catch { $info.docCountErr = $_.Exception.Message }
            try {
                $d = $(Resolve-Doc $word $p)
                $info.activeDocNull = ($null -eq $d)
                if ($null -ne $d) {
                    try { $info.activeDocName = [string]$d.Name } catch { $info.adnErr = $_.Exception.Message }
                    try { $info.contentNull = ($null -eq $d.Content) } catch { $info.contentErr = $_.Exception.Message }
                    try { $info.paraCount = [int]$d.Paragraphs.Count } catch { $info.paraErr = $_.Exception.Message }
                    try { $info.stylesNull = ($null -eq $d.Styles) } catch { $info.stylesErr = $_.Exception.Message }
                    try { $info.pageSetupNull = ($null -eq $d.PageSetup) } catch { $info.psErr = $_.Exception.Message }
                }
            } catch { $info.activeDocErr = $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = $info }
    }

    # ==================== 能力探测（阶段 6 的第一步）====================
    # 为什么先探测再实现
    # ------------------
    # 前面已经在本机 WPS 12.1 上踩了三次接口不支持（InsertCaption、CaptionLabels、
    # ListTemplates.Add）。如果按"Word 官方 API 应该都有"直接写，实现完才发现
    # 一半跑不通，返工成本很高。
    # 所以先把候选接口在【真机】上各试一次，输出一张"支持 / 不支持"的表，
    # 再据此定实现清单 —— 这一步能省掉大量返工。
    #
    # ⚠️ 探针绝大多数是只读的；两处例外（styles.modifyFont / table.merge / table.addRow）
    #    会短暂改动后复原或依赖文档现状，所以【必须在副本上跑】。
    "__probeCaps" {
        $word = Get-ActiveWpsApp 'Kwps.Application'
        $doc = if ($word) { Resolve-Doc $word $p } else { $null }
        if (-not $doc) {
            Output-Json @{ success = $false; error = "没有可用的文稿（能力探测需要打开一个文档）" }
            return
        }
        $caps = [ordered]@{}

        # ---- P1：样式 CUD ----
        $caps["styles.list"]          = { $null = $doc.Styles.Count; $null = $doc.Styles.Item("正文") }
        $caps["styles.add"]           = { $s = $doc.Styles.Add("__probe_style__", 1); if ($null -eq $s) { throw "Styles.Add 返回 null" }
                                          $back = $doc.Styles.Item("__probe_style__"); if ($null -eq $back) { throw "样式建了但取不回来" }
                                          try { $back.Delete() } catch {} }
        $caps["styles.delete"]        = { $s = $doc.Styles.Add("__probe_del__", 1); $s.Delete() }
        $caps["styles.modifyFont"]    = { $s = $doc.Styles.Item("正文"); $old = $s.Font.NameFarEast
                                          try { $s.Font.NameFarEast = "__probe__" } finally { $s.Font.NameFarEast = $old } }
        $caps["styles.modifyParagraph"] = { $s = $doc.Styles.Item("正文"); $old = $s.ParagraphFormat.SpaceAfter
                                          try { $s.ParagraphFormat.SpaceAfter = 99 } finally { $s.ParagraphFormat.SpaceAfter = $old } }
        $caps["styles.byIndex"]       = { $null = $doc.Styles.Item(-1) }   # wdStyleHeading1

        # ---- P1：交叉引用 ----
        $caps["crossRef.items"]       = { $null = $doc.GetCrossReferenceItems(1) }   # wdRefTypeHeading
        $caps["crossRef.fieldAdd"]    = { $rng = $doc.Range(0, [Math]::Min(1, $doc.Content.End))
                                          $f = $doc.Fields.Add($rng, -1, "STYLEREF 1", $false)
                                          if ($null -eq $f) { throw "Fields.Add 返回 null" }
                                          $t = ""; try { $t = [string]$f.Code.Text } catch {}
                                          if ($t -notmatch "STYLEREF") { throw "域已插入但代码不是 STYLEREF（实际=$t）" }
                                          try { $f.Delete() } catch {} }

        # ---- P1：索引 ----
        $caps["indexes.count"]        = { $null = $doc.Indexes.Count }
        $caps["indexes.markEntry"]    = { $r = $doc.Indexes.MarkEntry($doc.Range(0,1)); if ($null -eq $r) { throw "MarkEntry 返回 null" } }
        $caps["indexes.add"]          = { $ix = $doc.Indexes.Add($doc.Range(0, [Math]::Min(1, $doc.Content.End))); if ($null -eq $ix) { throw "Indexes.Add 返回 null" }; try { $ix.Delete() } catch {} }

        # ---- P0：修订接受/拒绝 ----
        $caps["revisions.count"]      = { $null = $doc.Revisions.Count }
        $caps["revisions.item"]       = { if ($doc.Revisions.Count -gt 0) { $null = $doc.Revisions.Item(1).Range.Text } }
        $caps["revisions.acceptAll"]  = { $doc.Revisions.AcceptAll() }
        $caps["revisions.rejectAll"]  = { $doc.Revisions.RejectAll() }

        # ---- P0：表格单元格级 ----
        $caps["table.cell"]           = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Range.Text } }
        $caps["table.rowsAdd"]        = { if ($doc.Tables.Count -le 0) { throw "文档里没有表格，跳过" }
                                          $t = $doc.Tables.Item(1); $n0 = $t.Rows.Count; $null = $t.Rows.Add()
                                          if ($t.Rows.Count -ne $n0 + 1) { throw "Rows.Add 返回了但行数没变（$n0→$($t.Rows.Count)）" } }
        $caps["table.colsAdd"]        = { if ($doc.Tables.Count -le 0) { throw "文档里没有表格，跳过" }
                                          $t = $doc.Tables.Item(1); $n0 = $t.Columns.Count; $null = $t.Columns.Add()
                                          if ($t.Columns.Count -ne $n0 + 1) { throw "Columns.Add 返回了但列数没变（$n0→$($t.Columns.Count)）" } }
        $caps["table.cellMerge"]      = { if ($doc.Tables.Count -gt 0) { $t=$doc.Tables.Item(1); $t.Cell(1,1).Merge($t.Cell(1, [Math]::Max(2,[Math]::Min(2,$t.Columns.Count)))) } }
        $caps["table.cellSplit"]      = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Split(1,2) } }
        $caps["table.cellBorders"]    = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Borders } }
        $caps["table.style"]          = { if ($doc.Tables.Count -gt 0) { $doc.Tables.Item(1).Style = "网格型" } }
        $caps["table.rowsDelete"]     = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Rows.Count } }

        # ---- P1：页码/奇偶页/首页不同 ----
        $caps["pageNumbers.add"]      = { $f = $doc.Sections.Item(1).Footers.Item(1); $n0 = $f.PageNumbers.Count
                                          $null = $f.PageNumbers.Add(4, $true)
                                          if ($f.PageNumbers.Count -le $n0) { throw "PageNumbers.Add 后页码数没增加（$n0→$($f.PageNumbers.Count)）" } }
        $caps["pageSetup.differentFirst"] = { $v = $doc.PageSetup.DifferentFirstPageHeaderFooter; $doc.PageSetup.DifferentFirstPageHeaderFooter = $v }
        $caps["pageSetup.oddEven"]    = { $v = $doc.PageSetup.OddAndEvenPagesHeaderFooter; $doc.PageSetup.OddAndEvenPagesHeaderFooter = $v }

        # ---- P2：水印/内容控件/形状/保护/查找/批注/比较/邮件合并 ----
        $caps["watermark"]            = { $null = $doc.Sections.Item(1).Headers.Item(1).Shapes }
        $caps["contentControl"]       = { $null = $doc.ContentControls.Count }
        $caps["shape.addTextbox"]     = { $null = $doc.Shapes.AddTextbox(1, 0, 0, 100, 50) }
        $caps["docProtect"]           = { $null = $doc.ProtectionType }
        $caps["find.wildcards"]       = { $f = $doc.Content.Find; $f.MatchWildcards = $true; $f.ClearFormatting(); $null = $f.Execute("__probe__") }
        $caps["comments.count"]       = { $null = $doc.Comments.Count }
        $caps["comments.add"]         = { $c = $doc.Comments.Add($doc.Range(0,1), "__probe__"); try { $c.Delete() } catch {} }
        $caps["docCompare"]           = { $null = $doc.Compare }
        $caps["mailMerge"]            = { $null = $doc.MailMerge.State }
        $caps["exportPdf"]            = { $null = $doc.ExportAsFixedFormat }

        # ---- 对照锚点：这三个是我【曾经】判定为"不支持"的接口。
        #      放进同一张表是为了让探针自己回答"当初的归因对不对"。
        #
        #      ⚠️ 重要教训（第二轮探测才暴露的）：判据不能只看"有没有抛异常"。
        #      第一轮这三个里有两个报了 supported，说明"接口不可用"的结论下早了 ——
        #      真实原因是【调用方式】不对（参数/返回值），不是接口缺失。
        #      所以这里的探针都带"有效性断言"：能调用 ≠ 能用。
        $caps["anchor.insertCaption"] = {
            $null = $doc.CaptionLabels.Count
            $c = $word.Selection.Range.InsertCaption("图", "probe", 1, $false)
            if ($null -eq $c -or [string]::IsNullOrWhiteSpace([string]$c.Text)) {
                throw "InsertCaption 可调用但没产生题注（返回空）"
            }
        }
        $caps["anchor.captionLabels"] = {
            $n = $doc.CaptionLabels.Count
            if ($n -le 0) { throw "CaptionLabels 接口可用，但当前文档没有任何题注标签（Count=$n）" }
        }
        $caps["anchor.listTemplatesAdd"] = {
            $lt = $doc.ListTemplates.Add($true)
            if ($null -eq $lt) { throw "ListTemplates.Add 返回 null —— 接口可调用但没有产出列表模板" }
        }

        foreach ($k in @($caps.Keys)) {
            $status = "unknown"; $err = $null
            try {
                & $caps[$k] | Out-Null
                $status = "supported"
            } catch {
                $status = "unsupported"
                $err = [string]$_.Exception.Message
                if ($err.Length -gt 160) { $err = $err.Substring(0, 160) }
            }
            $caps[$k] = @{ status = $status; error = $err }
        }

        # 对照组自检：三个锚点必须【都】不是 supported —— 否则说明探针判据太松
        # （比如只 catch 异常、把"返回空"当成成功）。
        $anchorNames = @($caps.Keys | Where-Object { $_ -like "anchor.*" })
        $anchorLeaked = @($anchorNames | Where-Object { $caps[$_].status -eq "supported" })

        $keys = @($caps.Keys)
        Output-Json @{ success = $true; data = @{
            docName = [string]$doc.Name
            total = $keys.Count
            supported = @($keys | Where-Object { $caps[$_].status -eq "supported" })
            unsupported = @($keys | Where-Object { $caps[$_].status -ne "supported" })
            caps = $caps
            anchorSelfCheck = @{
                total = $anchorNames.Count
                leaked = $anchorLeaked
                passed = ($anchorLeaked.Count -eq 0)
                note = "锚点全部报 unsupported 才算探针有效；有 leaked 说明判据太松，不能拿这张表做决策。"
            }
            note = "探针多数只读；styles.modify* 会改完改回，table.rowsAdd/colsAdd 会真的加行加列 —— 请在【副本】上跑。表格类探针需要文档里已有表格，否则会因为 Count=0 静默跳过而误报 supported。"
        } }
    }

    # ==================== P0：修订的接受 / 拒绝 ====================
    # 之前只能开/查询修订，没有"处理"这个出口 —— 半个功能。
    # mode: acceptAll | rejectAll | acceptOne | rejectOne
    "handleRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $mode = [string]$p.mode
        $before = 0
        try { $before = [int]$doc.Revisions.Count } catch {}
        if ($before -eq 0) {
            Output-Json @{ success = $true; data = @{ mode = $mode; count = 0; handled = 0; note = "文档里没有修订" } }
            return
        }
        try {
            switch ($mode) {
                "acceptAll" { $doc.Revisions.AcceptAll() }
                "rejectAll" { $doc.Revisions.RejectAll() }
                "acceptOne" {
                    if ($null -eq $p.index) { Output-Json @{ success = $false; error = "acceptOne 需要 index（从 1 开始）" }; return }
                    $doc.Revisions.Item([int]$p.index).Accept()
                }
                "rejectOne" {
                    if ($null -eq $p.index) { Output-Json @{ success = $false; error = "rejectOne 需要 index（从 1 开始）" }; return }
                    $doc.Revisions.Item([int]$p.index).Reject()
                }
                default { Output-Json @{ success = $false; error = "mode 只能是 acceptAll/rejectAll/acceptOne/rejectOne，收到: $mode" }; return }
            }
            $after = [int]$doc.Revisions.Count
            Output-Json @{ success = $true; data = @{ mode = $mode; before = $before; after = $after; handled = ($before - $after) } }
        } catch {
            Output-Json @{ success = $false; error = "处理修订失败: $(Get-CallError $_)" }
        }
    }

    # 列出修订（谁改的、改成什么）—— 处理之前得先看得见
    "listRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $items = @()
            $n = [int]$doc.Revisions.Count
            for ($i = 1; $i -le $n; $i++) {
                $r = $doc.Revisions.Item($i)
                $e = @{ index = $i }
                try { $e.type = [int]$r.Type } catch {}            # 1=插入 2=删除 3=属性 4=段落标记...
                try { $e.author = [string]$r.Author } catch {}
                try { $e.date = [string]$r.Date } catch {}
                try { $e.text = [string]$r.Range.Text } catch {}
                try { $e.page = [int]$r.Range.Information(3) } catch {}   # wdActiveEndPageNumber
                $items += $e
            }
            Output-Json @{ success = $true; data = @{ count = $n; revisions = $items } }
        } catch {
            Output-Json @{ success = $false; error = "读取修订失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P0：表格单元格级 ====================
    # tableIndex 默认 1（第一个表格）。坐标一律 1-based，与 Word UI 一致。
    "mergeTableCells" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $a = $t.Cell([int]$p.row1, [int]$p.col1)
            $b = $t.Cell([int]$p.row2, [int]$p.col2)
            $a.Merge($b)
            Output-Json @{ success = $true; data = @{ merged = "$($p.row1),$($p.col1)-$($p.row2),$($p.col2)" } }
        } catch {
            Output-Json @{ success = $false; error = "合并 Word 表格单元格失败（注意：不能跨越已有合并区）: $(Get-CallError $_)" }
        }
    }

    "splitCell" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            # 两个参数必须先落到变量再传：写成 $c.Split((..), (..)) 时，
            # 那个逗号会被 PowerShell 当成数组运算符，整个块解析失败（实测踩过）。
            $nr = Get-IntParam $p "numRows" 1
            $nc = Get-IntParam $p "numCols" 1
            $c.Split($nr, $nc)
            Output-Json @{ success = $true; data = @{ split = "$($p.row),$($p.col)"; numRows = $nr; numCols = $nc } }
        } catch {
            Output-Json @{ success = $false; error = "拆分单元格失败（该单元格可能未合并过）: $(Get-CallError $_)" }
        }
    }

    # 插入/删除行列。dir: rows | cols；where: before | after；count 默认 1
    "insertTableRowCol" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $cnt = (Get-IntParam $p "count" 1)
            $where = if ($p.where -eq "before") { 0 } else { 1 }   # wdRowHeightAuto=0? 实测 0=前 1=后
            $info = @{ table = $t.Index; before = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count } }
            if ($p.dir -eq "cols") {
                $col = $t.Columns.Item([int]$p.index)
                if ($where -eq 0) { for ($k = 0; $k -lt $cnt; $k++) { $null = $col.Select(); $null = $t.Columns.Add($null) } }
                else { $null = $col.Select(); $null = $t.Columns.Add() }
            } else {
                $row = $t.Rows.Item([int]$p.index)
                if ($where -eq 0) { for ($k = 0; $k -lt $cnt; $k++) { $null = $row.Select(); $null = $t.Rows.Add($null) } }
                else { $null = $row.Select(); $null = $t.Rows.Add() }
            }
            $info.after = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            Output-Json @{ success = $true; data = $info }
        } catch {
            Output-Json @{ success = $false; error = "插入行列失败: $(Get-CallError $_)" }
        }
    }

    "deleteTableRowCol" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $before = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            if ($p.dir -eq "cols") { $t.Columns.Item([int]$p.index).Delete() } else { $t.Rows.Item([int]$p.index).Delete() }
            $after = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            Output-Json @{ success = $true; data = @{ before = $before; after = $after } }
        } catch {
            Output-Json @{ success = $false; error = "删除行列失败: $(Get-CallError $_)" }
        }
    }

    # 单元格边框。side: all | top | bottom | left | right；style 可选（1=单线 3=粗线 0=无）
    "setCellBorders" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            $style = if ($null -ne $p.style) { [int]$p.style } else { 1 }
            # wdBorderTop=1 Bottom=3 Left=2 Right=4 Horizontal=5 Vertical=6
            $map = @{ top = 1; left = 2; bottom = 3; right = 4; horizontal = 5; vertical = 6 }
            $applied = @()
            if ($p.side -eq "all") {
                foreach ($k in @("top", "left", "bottom", "right")) {
                    $c.Borders.Item($map[$k]).LineStyle = $style
                    $c.Borders.Item($map[$k]).LineWidth = 4      # wdLineWidth050pt
                    $applied += $k
                }
            } else {
                $key = [string]$p.side
                if (-not $map.ContainsKey($key)) { Output-Json @{ success = $false; error = "side 只能是 top/bottom/left/right/horizontal/vertical/all" }; return }
                $c.Borders.Item($map[$key]).LineStyle = $style
                if ($style -ne 0) { $c.Borders.Item($map[$key]).LineWidth = 4 }
                $applied += $key
            }
            Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; applied = $applied; style = $style } }
        } catch {
            Output-Json @{ success = $false; error = "设置单元格边框失败: $(Get-CallError $_)" }
        }
    }

    # 读单元格文本（表格类操作的前置确认）
    "getCellText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            $txt = ""
            try { $txt = [string]$c.Range.Text } catch {}
            $txt = $txt -replace "[`r`n`a]", ""      # 去掉单元格结尾标记
            $text = ""
            try { $text = [string]$c.Range.Text } catch {}
            Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $txt; raw = $text } }
        } catch {
            Output-Json @{ success = $false; error = "读取单元格失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：样式 新建 / 修改 / 删除 ====================
    "listStyles" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $onlyUser = $false
            if ($null -ne $p.onlyUserDefined) { $onlyUser = [bool]$p.onlyUserDefined }
            $items = @()
            foreach ($s in @($doc.Styles)) {
                $builtin = $true
                try { $builtin = [bool]$s.BuiltIn } catch {}
                if ($onlyUser -and $builtin) { continue }
                $e = @{ name = [string]$s.NameLocal; builtIn = $builtin }
                try { $e.type = [int]$s.Type } catch {}
                try { $e.inUse = [bool]$s.InUse } catch {}
                $items += $e
            }
            Output-Json @{ success = $true; data = @{ count = $items.Count; onlyUserDefined = $onlyUser; styles = $items } }
        } catch {
            Output-Json @{ success = $false; error = "列出样式失败: $(Get-CallError $_)" }
        }
    }

    # baseOn 可选（基于哪个样式派生）。其余属性可选，不传就不改。
    "addStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        # "样式是否已存在"必须用 try/catch 判，不能看返回值
        # --------------------------------------------------
        # $doc.Styles.Item(不存在的名字) 在 WPS 里是【抛异常】，不是返回 null。
        # 第一版写成 `if ($null -ne $doc.Styles.Item($p.name))`，异常被外层 catch
        # 接住，于是"想建个新样式"永远报 E_FAIL —— 而同样的 Styles.Add 调用
        # 单独跑（见 __probeStyleAdd 的 A~F 六组对照）是完全成功的。
        # 坑就在这个已存在检查上。
        $alreadyExists = $false
        try { $null = $doc.Styles.Item([string]$p.name); $alreadyExists = $true } catch { $alreadyExists = $false }
        if ($alreadyExists) {
            Output-Json @{ success = $false; error = "样式 '$($p.name)' 已存在；改样式请用 modifyStyle" }; return
        }
        try {
            # wdStyleTypeParagraph=1
            # 样式类型用【字符串】参数 styleType，而不是数字 type ——
            # 数字 1/2 直接传太脆（调用方要查表），而 schema 里的 type 字段本身就是字符串，
            # 用同一个名字会让 TS 侧的类型定义自相矛盾。
            $stType = 1   # wdStyleTypeParagraph
            if ($p.styleType -eq "character") { $stType = 2 }

            # 建样式与设属性必须【分开】处理：实测把两者写在一个 try 里时，
            # 只要某个属性改不了就整体报 E_FAIL，样式根本建不出来。
            # 现在建样式单独 try —— 建成功就是建成功，属性设不上的逐项如实报告。
            $s = $null
            try {
                $s = $doc.Styles.Add([string]$p.name, $stType)
            } catch {
                Output-Json @{ success = $false; error = "Styles.Add 失败（样式名冲突？或本机 WPS 不支持新建样式）: $(Get-CallError $_)" }
                return
            }
            if ($null -eq $s) { Output-Json @{ success = $false; error = "Styles.Add 返回 null" }; return }
            $failed = @()
            if ($p.baseOn) { try { $s.BaseStyle = [string]$p.baseOn } catch { $failed += "baseOn" } }
            $applied = @()
            # 逐属性设置：设上了记 applied，设不上记 failed —— 不静默吞掉，也不整体失败
            if ($p.fontSize) { try { $s.Font.Size = [double]$p.fontSize; $applied += "fontSize" } catch { $failed += "fontSize" } }
            if ($p.fontName) {
                try { $s.Font.NameFarEast = [string]$p.fontName; $applied += "fontName" } catch { $failed += "fontName" }
                try { $s.Font.NameAscii = [string]$p.fontName } catch { }
            }
            if ($null -ne $p.bold) { try { $s.Font.Bold = [bool]$p.bold; $applied += "bold" } catch { $failed += "bold" } }
            if ($null -ne $p.italic) { try { $s.Font.Italic = [bool]$p.italic; $applied += "italic" } catch { $failed += "italic" } }
            if ($null -ne $p.color) {
                try {
                    $bgr = [int]$p.color
                    $s.Font.Color = ($bgr -band 0xFF) * 65536 + ($bgr -shr 8 -band 0xFF) * 256 + ($bgr -shr 16 -band 0xFF)
                    $applied += "color"
                } catch { $failed += "color" }
            }
            if ($null -ne $p.spaceBefore) { try { $s.ParagraphFormat.SpaceBefore = [double]$p.spaceBefore; $applied += "spaceBefore" } catch { $failed += "spaceBefore" } }
            if ($null -ne $p.spaceAfter) { try { $s.ParagraphFormat.SpaceAfter = [double]$p.spaceAfter; $applied += "spaceAfter" } catch { $failed += "spaceAfter" } }
            if ($null -ne $p.lineSpacing) { try { $s.ParagraphFormat.LineSpacingRule = 5; $s.ParagraphFormat.LineSpacing = [double]$p.lineSpacing; $applied += "lineSpacing" } catch { $failed += "lineSpacing" } }
            if ($null -ne $p.alignment) {
                $map = @{ left = 0; center = 1; right = 2; justify = 3 }
                $k = [string]$p.alignment
                if ($map.ContainsKey($k)) { try { $s.ParagraphFormat.Alignment = $map[$k]; $applied += "alignment" } catch { $failed += "alignment" } }
            }
            # ⚠️ warnings 先算好再放进 hashtable。不要写成 `warnings = (if ...)`：
            #    hashtable 字面里带括号的 if 会在运行期报
            #    "无法将 if 项识别为 cmdlet"（而文件级 PARSE 却是通的，很迷惙）。
            #    同文件 insertCaption 那处不带括号能跗，但别依赖这种微妙差异。
            $warnings = @()
            if ($failed.Count -gt 0) {
                $warnings = @(@{
                    step = "addStyle.属性"
                    message = "以下属性本机 WPS 不接受：$($failed -join '/')。样式已建成，其余属性已设上。"
                })
            }
            Output-Json @{ success = $true; data = @{
                name = [string]$p.name
                styleName = [string]$s.NameLocal
                applied = $applied
                failed = $failed
                note = if ($warnings.Count -gt 0) { $warnings[0].message } else { $null }
                warnings = $warnings
            } }
        } catch {
            Output-Json @{ success = $false; error = "新建样式失败: $(Get-CallError $_)" }
        }
    }

    "modifyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        try {
            $s = $doc.Styles.Item([string]$p.name)
            if ($null -eq $s) { Output-Json @{ success = $false; error = "样式不存在: $($p.name)" }; return }
            if ($s.BuiltIn) {
                # 内置样式不允许整体改名/删除，但改属性是允许的 —— 这里只提示不拦
            }
            $applied = @()
            if ($p.fontSize) { $s.Font.Size = [double]$p.fontSize; $applied += "fontSize" }
            if ($p.fontName) { $s.Font.NameFarEast = [string]$p.fontName; $s.Font.NameAscii = [string]$p.fontName; $applied += "fontName" }
            if ($null -ne $p.bold) { $s.Font.Bold = [bool]$p.bold; $applied += "bold" }
            if ($null -ne $p.italic) { $s.Font.Italic = [bool]$p.italic; $applied += "italic" }
            if ($null -ne $p.color) {
                $bgr = [int]$p.color
                $s.Font.Color = ($bgr -band 0xFF) * 65536 + ($bgr -shr 8 -band 0xFF) * 256 + ($bgr -shr 16 -band 0xFF)
                $applied += "color"
            }
            if ($null -ne $p.spaceBefore) { $s.ParagraphFormat.SpaceBefore = [double]$p.spaceBefore; $applied += "spaceBefore" }
            if ($null -ne $p.spaceAfter) { $s.ParagraphFormat.SpaceAfter = [double]$p.spaceAfter; $applied += "spaceAfter" }
            if ($null -ne $p.lineSpacing) { $s.ParagraphFormat.LineSpacingRule = 5; $s.ParagraphFormat.LineSpacing = [double]$p.lineSpacing; $applied += "lineSpacing" }
            if ($null -ne $p.alignment) {
                $map = @{ left = 0; center = 1; right = 2; justify = 3 }
                $k = [string]$p.alignment
                if ($map.ContainsKey($k)) { $s.ParagraphFormat.Alignment = $map[$k]; $applied += "alignment" }
            }
            if ($null -ne $p.newName) {
                try { $s.NameLocal = [string]$p.newName; $applied += "newName" } catch {
                    $applied += "newName(内置样式不可改名，已跳过)"
                }
            }
            if ($applied.Count -eq 0) {
                Output-Json @{ success = $false; error = "没有任何可改的属性（fontSize/fontName/bold/italic/color/spaceBefore/spaceAfter/lineSpacing/alignment/newName）" }; return
            }
            Output-Json @{ success = $true; data = @{ name = [string]$p.name; applied = $applied } }
        } catch {
            Output-Json @{ success = $false; error = "修改样式失败: $(Get-CallError $_)" }
        }
    }

    "deleteStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        try {
            $s = $doc.Styles.Item([string]$p.name)
            if ($null -eq $s) { Output-Json @{ success = $false; error = "样式不存在: $($p.name)" }; return }
            if ($s.BuiltIn) {
                Output-Json @{ success = $false; error = "'$($p.name)' 是内置样式，不能删除（可以 modifyStyle 改它的属性）" }; return
            }
            $s.Delete()
            Output-Json @{ success = $true; data = @{ deleted = [string]$p.name } }
        } catch {
            Output-Json @{ success = $false; error = "删除样式失败（样式正被段落使用时可能失败）: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：交叉引用 ====================
    # 在【当前光标位置】插入对某个标题的引用。
    # refType: heading | bookmark | numberedItem | footnote
    "insertCrossReference" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            # wdRefTypeHeading=1 wdRefTypeBookmark=2 wdRefTypeNumberedItem=3 wdRefTypeFootnote=4
            $rtMap = @{ heading = 1; bookmark = 2; numberedItem = 3; footnote = 4 }
            $k = if ($p.refType) { [string]$p.refType } else { "heading" }
            if (-not $rtMap.ContainsKey($k)) { Output-Json @{ success = $false; error = "refType 只能是 heading/bookmark/numberedItem/footnote" }; return }
            $rt = $rtMap[$k]
            $items = @($doc.GetCrossReferenceItems($rt))
            if ($items.Count -eq 0) {
                Output-Json @{ success = $false; error = "文档里没有可引用的$($k)（GetCrossReferenceItems 返回空）。heading 需要先有标题样式段落" }; return
            }
            $idx = (Get-IntParam $p "itemIndex" 1)
            if ($idx -lt 1 -or $idx -gt $items.Count) {
                Output-Json @{ success = $false; error = "itemIndex 越界：文档里共 $($items.Count) 个可引用项，收到 $idx" }; return
            }
            $target = [string]$items[$idx - 1]
            $rng = $word.Selection.Range
            $rng.Collapse(0)   # wdCollapseEnd
            # insertAsHyperlink / includePosition / separateNumbers / separatorString
            $null = $doc.Fields.Add($rng, -1, "REF $target `h", $true)
            Output-Json @{ success = $true; data = @{ refType = $k; itemIndex = $idx; target = $target; total = $items.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入交叉引用失败: $(Get-CallError $_)" }
        }
    }

    "listCrossReferenceTargets" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $rtMap = @{ heading = 1; bookmark = 2; numberedItem = 3; footnote = 4 }
            $out = @{}
            foreach ($kv in $rtMap.GetEnumerator()) {
                try { $out[$kv.Key] = @($doc.GetCrossReferenceItems($kv.Value)) } catch { $out[$kv.Key] = @() }
            }
            Output-Json @{ success = $true; data = @{ targets = $out } }
        } catch {
            Output-Json @{ success = $false; error = "列出可引用项失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：索引 ====================
    # MarkEntry 在探测里报了 E_FAIL（当时是在【新建的空文档】上试的），
    # 所以这里实现时多给了两个参数并保留详细报错，便于真机再判一次。
    "markIndexEntry" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $rng = if ($null -ne $p.start -and $null -ne $p.end) { $doc.Range([int]$p.start, [int]$p.end) } else { $word.Selection.Range }
            if ($rng -eq $null -or [string]::IsNullOrWhiteSpace([string]$rng.Text)) {
                Output-Json @{ success = $false; error = "选区为空：请先选中要编入索引的文字，或传 start/end" }; return
            }
            $entry = if ($p.entry) { [string]$p.entry } else { ([string]$rng.Text).Trim() }
            $r = $doc.Indexes.MarkEntry($rng, [string]$entry)
            if ($null -eq $r) {
                Output-Json @{ success = $false; error = "Indexes.MarkEntry 返回 null（本机 WPS 12.1 探测时报 E_FAIL，可能不支持该接口）" }; return
            }
            Output-Json @{ success = $true; data = @{ entry = $entry } }
        } catch {
            Output-Json @{ success = $false; error = "标记索引项失败: $(Get-CallError $_)" }
        }
    }

    "insertIndex" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $marked = 0
            try { $marked = [int]$doc.Indexes.MarkEntries.Count } catch {}
            $rng = $word.Selection.Range
            $rng.Collapse(0)
            $ix = $doc.Indexes.Add($rng)
            if ($null -eq $ix) { Output-Json @{ success = $false; error = "Indexes.Add 返回 null" }; return }
            $cols = 1; $heading = $null
            if ($null -ne $p.columns) { try { $ix.NumberOfColumns = [int]$p.columns; $cols = [int]$p.columns } catch {} }
            if ($p.heading) { try { $ix.HeadingSeparator = [string]$p.heading; $heading = [string]$p.heading } catch {} }
            Output-Json @{ success = $true; data = @{ markedEntries = $marked; columns = $cols; heading = $heading } }
        } catch {
            Output-Json @{ success = $false; error = "插入索引失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：页码域 / 奇偶页 / 首页不同 ====================
    # sectionIndex 默认 1（第一个节）
    "addPageNumber" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $sec = $doc.Sections.Item((Get-IntParam $p "sectionIndex" 1))
            # wdHeaderFooterPrimary=1 FirstPage=2 EvenPages=3
            $hfMap = @{ primary = 1; first = 2; even = 3 }
            $k = if ($p.which) { [string]$p.which } else { "primary" }
            if (-not $hfMap.ContainsKey($k)) { Output-Json @{ success = $false; error = "which 只能是 primary/first/even" }; return }
            $hf = $sec.Headers.Item($hfMap[$k])
            # wdAlignPageNumberRight=2
            $align = if ($p.alignment -eq "center") { 1 } elseif ($p.alignment -eq "left") { 0 } else { 2 }
            $n0 = [int]$hf.PageNumbers.Count
            $null = $hf.PageNumbers.Add($align, $true)
            if ($hf.PageNumbers.Count -le $n0) {
                Output-Json @{ success = $false; error = "PageNumbers.Add 调用了但页码数没增加（$n0→$($hf.PageNumbers.Count)）" }; return
            }
            Output-Json @{ success = $true; data = @{ section = $sec.Index; which = $k; alignment = $p.alignment; count = [int]$hf.PageNumbers.Count } }
        } catch {
            Output-Json @{ success = $false; error = "添加页码失败: $(Get-CallError $_)" }
        }
    }

    # 奇偶页不同 / 首页不同 —— PageSetup 级的开关
    "setHeaderFooterLayout" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $applied = @()
            if ($null -ne $p.differentFirstPage) {
                $doc.PageSetup.DifferentFirstPageHeaderFooter = [bool]$p.differentFirstPage
                $applied += "differentFirstPage"
            }
            if ($null -ne $p.oddAndEvenPages) {
                $doc.PageSetup.OddAndEvenPagesHeaderFooter = [bool]$p.oddAndEvenPages
                $applied += "oddAndEvenPages"
            }
            if ($applied.Count -eq 0) {
                Output-Json @{ success = $false; error = "至少要传 differentFirstPage 或 oddAndEvenPages 之一" }; return
            }
            Output-Json @{ success = $true; data = @{
                applied = $applied
                differentFirstPage = [bool]$doc.PageSetup.DifferentFirstPageHeaderFooter
                oddAndEvenPages = [bool]$doc.PageSetup.OddAndEvenPagesHeaderFooter
            } }
        } catch {
            Output-Json @{ success = $false; error = "设置页眉页脚布局失败: $(Get-CallError $_)" }
        }
    }

    default {
        Output-Json @{ success = $false; error = "Unknown action: $Action" }
    }
}

}   # function Invoke-WpsAction（上一行的 } 是 switch 的收尾）

# ==================== 入口 ====================
# 常驻模式：wps-comd.ps1 先设 $env:WPS_COM_DAEMON=1 再 dot-source 本文件，
#           此时只需拿到 Invoke-WpsAction 的定义，不执行任何动作。
# 直跑模式（默认）：行为与改造前完全一致。
if ($env:WPS_COM_DAEMON -ne '1') {
    try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }
    Invoke-WpsAction -Action $Action -p $p
}
