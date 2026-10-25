# Input: 用户执行安装命令（本地 -File 或 irm | iex 一键）
# Output: 依赖安装 + TypeScript 构建 + WPS 加载项部署 + MCP 配置示例 + skill 安装
# Pos: Windows 安装入口脚本。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
#
# 两种用法：
#   一键（需能访问 GitHub）：
#     irm https://cdn.jsdelivr.net/gh/forkgo430/wps-multitool-mcp@main/scripts/install.ps1 | iex
#   本地（已克隆）：
#     powershell -ExecutionPolicy Bypass -File scripts/install.ps1
#
# 参数：
#   -InstallDir <dir>  一键模式下装到哪，默认 %USERPROFILE%\wps-multitool-mcp
#   -SkipNodeCheck     跳过 Node.js 检查
#   -SkipWpsCheck      跳过 WPS 安装检查
#   -Dev               加载项用符号链接部署（改源码即时生效，需管理员权限）
#   -DryRun            只检查 + 打印计划，不做任何改动
#   -Force             目标目录已存在时删除重建
#
# 普通权限即可运行（只有 -Dev 的目录符号链接需要管理员）。
# 不用 #Requires：一键模式是 `irm | iex` 执行的字符串，#Requires 在该上下文会报错，
# 所以改成运行时检查（作用相同，但两种模式都能跑）。

param(
    [string]$InstallDir = (Join-Path $env:USERPROFILE 'wps-multitool-mcp'),
    [switch]$SkipNodeCheck,
    [switch]$SkipWpsCheck,
    [switch]$Dev,
    [switch]$DryRun,
    [switch]$Force
)

$ErrorActionPreference = 'Continue'

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "需要 PowerShell 5.1 或更高版本（当前 $($PSVersionTable.PSVersion)）" -ForegroundColor Red
    return
}

$RepoGit = 'https://github.com/forkgo430/wps-multitool-mcp.git'
$RepoTar = 'https://codeload.github.com/forkgo430/wps-multitool-mcp/tar.gz/refs/heads/main'

# ============================================================================
# 输出辅助
# ============================================================================
function Write-Step { param([string]$m) Write-Host "`n>>> $m" -ForegroundColor Cyan }
function Write-Ok { param([string]$m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn2 { param([string]$m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Err2 { param([string]$m) Write-Host "  [X]  $m" -ForegroundColor Red }
function Write-Hint { param([string]$m) Write-Host "       $m" -ForegroundColor DarkGray }

function Show-Banner {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   WPS Office 多面手 MCP  ·  wps-multitool-mcp" -ForegroundColor White
    Write-Host "   让 AI 直接操控你正开着的 WPS 文字 / 表格 / 演示" -ForegroundColor DarkGray
    Write-Host "  ============================================================" -ForegroundColor Cyan
    if ($DryRun) { Write-Warn2 "DryRun 模式：只检查与打印计划，不做任何改动" }
    Write-Host ""
}

# ============================================================================
# 定位项目根目录
#   本地模式：脚本位于 <repo>/scripts/install.ps1
#   一键模式：$PSScriptRoot 为空 —— 先看 InstallDir 是否已有项目，没有就拉取
# ============================================================================
function Resolve-ProjectRoot {
    if ($PSScriptRoot) {
        $parent = (Get-Item $PSScriptRoot).Parent.FullName
        if (Test-Path (Join-Path $parent 'wps-multitool-mcp')) {
            Write-Ok "使用本地项目: $parent"
            return $parent
        }
    }

    $hasProject = Test-Path (Join-Path $InstallDir 'wps-multitool-mcp')

    if ($hasProject -and -not $Force) {
        Write-Ok "使用已存在项目: $InstallDir（加 -Force 可删除重建）"
        return $InstallDir
    }

    if ($DryRun) {
        Write-Warn2 "DryRun：将把项目拉取到 $InstallDir"
        return $null
    }

    if ($hasProject -and $Force) {
        Write-Warn2 "删除已存在目录: $InstallDir"
        Remove-Item -Path $InstallDir -Recurse -Force
    }

    Write-Step "获取项目文件 -> $InstallDir"
    $hasGit = [bool](Get-Command git -ErrorAction SilentlyContinue)
    if ($hasGit) {
        Write-Hint "git clone --depth 1"
        git clone --depth 1 $RepoGit $InstallDir
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path (Join-Path $InstallDir 'wps-multitool-mcp'))) {
            throw "git clone 失败。常见原因：访问不了 github.com。可先用代理，或手动下载仓库后运行 scripts/install.ps1"
        }
    } else {
        Write-Hint "未检测到 git，改用 tarball 下载"
        $tmp = Join-Path $env:TEMP ("wps-multitool-" + [guid]::NewGuid().ToString('N') + ".tar.gz")
        Invoke-WebRequest -Uri $RepoTar -OutFile $tmp -UseBasicParsing
        New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
        tar -xzf $tmp -C $InstallDir --strip-components=1
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path (Join-Path $InstallDir 'wps-multitool-mcp'))) {
            throw "解压失败。常见原因：访问不了 github.com。可手动下载仓库后运行 scripts/install.ps1"
        }
    }
    Write-Ok "项目已就绪: $InstallDir"
    return $InstallDir
}

function Get-McpVersion {
    param([string]$ProjectRoot)
    $pkg = Join-Path $ProjectRoot 'wps-multitool-mcp\package.json'
    if (Test-Path $pkg) {
        try {
            # 用正则而不是 ConvertFrom-Json：老版本 PowerShell 对带 BOM/特殊字符的
            # json 偶尔会抛异常，版本号这种一行就能取的东西不值得冒险。
            $raw = Get-Content $pkg -Raw
            if ($raw -match '"version"\s*:\s*"([^"]+)"') { return $Matches[1] }
        } catch {}
    }
    return 'dev'
}

# ============================================================================
# 环境检查
# ============================================================================
function Test-NodeVersion {
    Write-Step "检查 Node.js"
    if ($SkipNodeCheck) { Write-Warn2 "已跳过"; return $true }
    try {
        $v = node --version 2>$null
        if (-not $v) { throw "未安装" }
        $major = [int]($v.TrimStart('v').Split('.')[0])
        if ($major -lt 18) {
            Write-Err2 "Node.js 版本过低: $v（需要 >= 18）"
            Write-Hint "下载: https://nodejs.org"
            return $false
        }
        Write-Ok "Node.js $v"
        return $true
    } catch {
        Write-Err2 "未检测到 Node.js"
        Write-Hint "请安装 Node.js 18+：https://nodejs.org"
        return $false
    }
}

function Test-NpmAvailable {
    try {
        $v = npm --version 2>$null
        if (-not $v) { throw "未安装" }
        Write-Ok "npm $v"
        return $true
    } catch {
        Write-Err2 "npm 不可用（Node.js 安装可能不完整）"
        return $false
    }
}

function Test-WpsInstallation {
    Write-Step "检查 WPS Office"
    if ($SkipWpsCheck) { Write-Warn2 "已跳过"; return $true }

    # 检测顺序：运行中的进程 → 卸载项 → 已知安装路径 → 注册表 6.0 子键。
    # 为什么不能只查 Program Files：实测有用户装在 D:\APPdate\WPS Office\...
    # 为什么不需要安装路径：加载项目录固定在 %APPDATA%\kingsoft\wps\jsaddons，
    # 只需要确认"装过 WPS"这件事。

    # 1) 运行中的进程（最可靠的正面信号）
    foreach ($n in @('wps', 'et', 'wpp', 'Ket', 'Kwps', 'Kwpp')) {
        $p = Get-Process -Name $n -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($p -and $p.Path) {
            Write-Ok "WPS Office 正在运行: $($p.Path)"
            return $true
        }
    }

    # 2) 卸载项（装过但没在运行）
    $uninstallRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($root in $uninstallRoots) {
        $hit = Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
            $item = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($item -and $item.DisplayName -like '*WPS Office*') { $item }
        } | Select-Object -First 1
        if ($hit) {
            Write-Ok "WPS Office 已安装: $($hit.DisplayName)"
            return $true
        }
    }

    # 3) 常见安装路径
    $candidates = @(
        "${env:ProgramFiles}\Kingsoft\WPS Office\ksolaunch.exe",
        "${env:ProgramFiles(x86)}\Kingsoft\WPS Office\ksolaunch.exe",
        "${env:LOCALAPPDATA}\Kingsoft\WPS Office\ksolaunch.exe",
        "${env:ProgramFiles}\Kingsoft\WPS Office\office6\wps.exe",
        "${env:ProgramFiles(x86)}\Kingsoft\WPS Office\office6\wps.exe"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) { Write-Ok "WPS Office: $p"; return $true }
    }

    # 4) 注册表子键（存在 Office\6.0 即说明装过）
    foreach ($reg in @('HKLM:\SOFTWARE\Kingsoft\Office\6.0', 'HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\6.0', 'HKCU:\SOFTWARE\Kingsoft\Office\6.0')) {
        if (Test-Path $reg) { Write-Ok "注册表存在 WPS Office 安装痕迹: $reg"; return $true }
    }

    Write-Warn2 "未检测到 WPS Office"
    Write-Hint "请先安装 WPS Office：https://www.wps.cn（加载项必须有它才有意义）"
    return $false
}

# ============================================================================
# 依赖安装 + 构建
# ============================================================================
function Install-McpServer {
    param([string]$ProjectRoot)
    $mcpPath = Join-Path $ProjectRoot 'wps-multitool-mcp'
    if (-not (Test-Path $mcpPath)) {
        Write-Err2 "找不到 wps-multitool-mcp 目录: $mcpPath"
        return $false
    }
    if ($DryRun) {
        Write-Step "DryRun：将执行 npm ci && npm run build（在 $mcpPath）"
        return $true
    }

    Push-Location $mcpPath
    try {
        Write-Step "安装依赖"
        if (Test-Path 'package-lock.json') { npm ci --loglevel=error } else { npm install --loglevel=error }
        if ($LASTEXITCODE -ne 0) { Write-Err2 "依赖安装失败"; return $false }
        Write-Ok "依赖就绪"

        Write-Step "编译 TypeScript"
        npm run build
        if ($LASTEXITCODE -ne 0) { Write-Err2 "编译失败"; return $false }
        if (-not (Test-Path 'dist\index.js')) { Write-Err2 "编译产物 dist/index.js 不存在"; return $false }
        Write-Ok "构建完成: dist/index.js"
        return $true
    } finally {
        Pop-Location
    }
}

# ============================================================================
# 合并 publish.xml 里的加载项条目（绝不整文件覆盖）
# ============================================================================
# 为什么必须是合并，不能 Out-File 一把写进去
# ----------------------------------------
# %APPDATA%\kingsoft\wps\jsaddons\publish.xml 是【全机所有 WPS 加载项】的共享清单：
# 我们的、别人厂商的、WPS 自带的，条目全在这一个文件里。整文件覆盖会把别人家的
# <jsplugin> 一起清掉 —— 症状是用户装完我们的插件，别的插件就从功能区消失了，
# 而他根本想不到是我们干的（我们这边还显示"安装成功"）。
#
# 行为：文件不存在 → 新建，只含我们这一条；
#       已存在 → 先删掉我们自己上次那条（重复执行不会堆条目），别人的行一个字符都不碰，
#                再把我们这条插到 </jsplugins> 前；
#       已存在但不是加载项清单（没有 <jsplugins>）→ 不动它，返回 $false 让调用方知道。
#
# 这个函数在 deploy-addon.ps1 里有一份同样的拷贝 —— 因为 install.ps1 要支持
# `irm | iex` 单文件执行，不能依赖旁边有没有 lib 文件。改这边记得同步那边。
function Merge-PublishXmlEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Entry
    )
    $line = "  $Entry"

    if (-not (Test-Path $Path)) {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        @"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
$line
</jsplugins>
"@ -replace '\s+$', '' | Out-File -FilePath $Path -Encoding utf8
        return $true
    }

    $raw = Get-Content -Raw -Path $Path
    if ($raw -notmatch '<jsplugins') {
        Write-Warn2 "publish.xml 已存在但不是加载项清单（没有 <jsplugins>），不动它：$Path"
        return $false
    }

    # 删掉自己上次那条（幂等）；行尾换行可选，兼容"我们那条恰好在文件末尾"
    $cleaned = $raw -replace "(?m)^[ \t]*<jsplugin\s+name=`"$([regex]::Escape($Name))`"[^>]*/?>[ \t]*\r?\n?", ''
    $idx = $cleaned.LastIndexOf('</jsplugins>')
    if ($idx -lt 0) {
        Write-Warn2 "publish.xml 里找不到 </jsplugins> 收尾，不动它：$Path"
        return $false
    }
    # 结尾的换行必须去掉：Out-File 自己会补一个行尾，带着原来的换行写会每次多一个空行
    # （实测重复部署三次，文件尾部堆出三个空行 —— 幂等就是这么悄悄坏掉的）。
    ($cleaned.Substring(0, $idx) + $line + "`r`n" + $cleaned.Substring($idx)) -replace '\s+$', '' |
        Out-File -FilePath $Path -Encoding utf8
    return $true
}

# ============================================================================
# 部署 WPS 加载项（%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_）
# ============================================================================
function Install-WpsAddon {
    param([string]$ProjectRoot)
    $src = Join-Path $ProjectRoot 'wps-multitool-addon'
    if (-not (Test-Path $src)) { Write-Err2 "找不到加载项目录: $src"; return $false }

    $addonsDir = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
    $target = Join-Path $addonsDir 'wps-claude-addon_'

    if ($DryRun) {
        Write-Step "DryRun：将部署加载项 -> $target"
        return $true
    }

    Write-Step "部署 WPS 加载项"
    New-Item -ItemType Directory -Force -Path $addonsDir | Out-Null
    if (Test-Path $target) {
        Write-Hint "清理旧版本"
        Remove-Item -Path $target -Recurse -Force
    }

    if ($Dev) {
        $null = cmd /c mklink /D "`"$target`"" "`"$src`"" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Warn2 "符号链接失败（需要管理员权限），降级为复制"
            Copy-Item -Path $src -Destination $target -Recurse -Force
        } else {
            Write-Ok "已链接: $target -> $src"
        }
    } else {
        Copy-Item -Path $src -Destination $target -Recurse -Force
        Write-Ok "已复制到: $target"
    }

    # publish.xml：WPS 靠它扫描加载项（目录名尾部 _ 是必须的）。
    # 这里【合并】条目而不是整文件覆盖 —— 该文件是全机所有加载项的共享清单，
    # 覆盖会清掉别人厂商的条目（详见 Merge-PublishXmlEntry 的注释）。
    $publishXml = Join-Path $addonsDir 'publish.xml'
    $null = Merge-PublishXmlEntry -Path $publishXml `
        -Name 'wps-claude-addon' `
        -Entry '<jsplugin name="wps-claude-addon" type="wps,et,wpp" url="wps-claude-addon_/" enable="enable_dev"/>'
    Write-Ok "publish.xml: $publishXml"
    return $true
}

# ============================================================================
# 安装 skill（存在对应客户端的 skills 目录才装）
# ============================================================================
function Install-Skill {
    param([string]$ProjectRoot)
    $src = Join-Path $ProjectRoot 'wps-multitool-mcp\skills\wps-multitool'
    if (-not (Test-Path $src)) { Write-Warn2 "找不到 skill 源目录，跳过"; return $true }

    $targets = @(
        (Join-Path $env:USERPROFILE '.claude\skills'),
        (Join-Path $env:USERPROFILE '.workbuddy\skills')
    )
    $installed = @()
    foreach ($base in $targets) {
        if (-not (Test-Path $base)) { continue }
        $dst = Join-Path $base 'wps-multitool'
        if ($DryRun) { Write-Hint "DryRun：将安装 skill -> $dst"; $installed += $dst; continue }
        if (Test-Path $dst) { Remove-Item -Path $dst -Recurse -Force }
        Copy-Item -Path $src -Destination $dst -Recurse -Force
        $installed += $dst
    }
    if ($installed.Count) {
        $verb = if ($DryRun) { 'skill 将安装到' } else { 'skill 已安装' }
        Write-Ok "${verb}: $($installed -join '; ')"
    } else {
        Write-Warn2 "未发现 .claude / .workbuddy skills 目录，跳过 skill 安装（用到再手动复制）"
    }
    return $true
}

# ============================================================================
# MCP 配置示例
# ============================================================================
function Show-McpConfig {
    param([string]$ProjectRoot)
    $dist = (Join-Path $ProjectRoot 'wps-multitool-mcp\dist\index.js').Replace('\', '/')
    # 用 here-string 而不是 ConvertTo-Json：PS 5.1 输出的一层层缩进很难看，配置要能直接复制
    $cfg = @"
{
  "mcpServers": {
    "wps-multitool-mcp": {
      "command": "node",
      "args": ["$dist"]
    }
  }
}
"@

    Write-Step "下一步：自动写客户端配置（推荐）"
    Write-Host "  node `"$dist`" install" -ForegroundColor Yellow
    Write-Hint "会自动识别 CodeBuddy / WorkBuddy / Codex / Claude Code / Cursor，写前备份、写后校验"
    Write-Hint "只想看会改什么：加 --dry-run；想自己粘：加 --print"

    Write-Step "或者手动粘（不推荐，路径容易漏一层）"
    Write-Host $cfg -ForegroundColor Yellow
    Write-Hint "注意：路径里出现两次 wps-multitool-mcp 是正常的（外层仓库目录、内层项目目录）"
}

# ============================================================================
# 完成信息
# ============================================================================
function Show-Completion {
    param([string]$ProjectRoot)
    $v = Get-McpVersion $ProjectRoot
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   安装完成（wps-multitool-mcp v$v）" -ForegroundColor Green
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  接下来：" -ForegroundColor White
    Write-Host "   1. 重启 WPS Office（加载项只在启动时加载）" -ForegroundColor White
    Write-Host "   2. 把上面的 MCP 配置写进你的 AI 客户端，重启客户端" -ForegroundColor White
    Write-Host "   3. 验证：" -ForegroundColor White
    Write-Host "        cd $(Join-Path $ProjectRoot 'wps-multitool-mcp')" -ForegroundColor DarkGray
    Write-Host "        npm run verify" -ForegroundColor DarkGray
    Write-Host "   4. 用法：在客户端里直接说任务，例如" -ForegroundColor White
    Write-Host "        「在这份文档末尾加一段总结」" -ForegroundColor DarkGray
    Write-Host "        「把这个表格按销售额降序并画柱状图」" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  遇到问题：" -ForegroundColor White
    Write-Host "   - MCP 连不上：确认 dist/index.js 存在、客户端配置路径正确" -ForegroundColor DarkGray
    Write-Host "   - WPS 面板不显示：确认 %APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_ 存在并重启 WPS" -ForegroundColor DarkGray
    Write-Host "   - 其他：https://github.com/forkgo430/wps-multitool-mcp/issues" -ForegroundColor DarkGray
    Write-Host ""
}

# ============================================================================
# 主流程
# ============================================================================
function Start-Installation {
    Show-Banner

    $ok = $true
    if (-not (Test-NodeVersion)) { $ok = $false }
    if ($ok -and -not (Test-NpmAvailable)) { $ok = $false }
    if ($ok -and -not (Test-WpsInstallation)) { $ok = $false }
    if (-not $ok) {
        Write-Host ""
        Write-Err2 "环境检查未通过，已停止。修好上面的问题后重跑。"
        return 1
    }

    try {
        $projectRoot = Resolve-ProjectRoot
    } catch {
        Write-Err2 $_.Exception.Message
        return 1
    }
    if (-not $projectRoot) {
        Write-Host ""
        Write-Warn2 "DryRun 结束（未实际下载/构建/部署）"
        return 0
    }

    if (-not (Install-McpServer $projectRoot)) { Write-Err2 "MCP Server 安装失败"; return 1 }
    if (-not (Install-WpsAddon $projectRoot)) { Write-Err2 "加载项部署失败"; return 1 }
    if (-not (Install-Skill $projectRoot)) { Write-Err2 "skill 安装失败"; return 1 }

    Show-McpConfig $projectRoot
    Show-Completion $projectRoot
    return 0
}

$code = Start-Installation
if ($PSScriptRoot) { exit $code }
