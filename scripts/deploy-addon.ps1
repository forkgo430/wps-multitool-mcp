# 把加载项从源目录部署到 WPS 真正读取的位置。
#
# 为什么需要这一步（踩过的坑）
# --------------------------
# install.ps1 用的是 Copy-Item（符号链接需要管理员权限，失败时回退复制），
# 所以 %APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_ 是一份【副本】。
# 只改源目录不重新部署，WPS 重启多少次都还是旧内容 —— 赢容反馈
# "wps上方还是之前的MCP工具没变，重启也一样"就是这个原因。
#
# 用法：powershell -ExecutionPolicy Bypass -File scripts\deploy-addon.ps1
#   -Check  只检查差异，不复制

param(
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$addonSource = Join-Path $projectRoot 'wps-claude-addon'
$wpsAddonsPath = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$addonTarget = Join-Path $wpsAddonsPath 'wps-claude-addon_'
$publishXmlPath = Join-Path $wpsAddonsPath 'publish.xml'

Write-Host ''
Write-Host '=== 部署 WPS 加载项 ===' -ForegroundColor Cyan
Write-Host "  源 : $addonSource"
Write-Host "  目标: $addonTarget"
Write-Host ''

if (-not (Test-Path $addonSource)) {
    Write-Host "✗ 源目录不存在: $addonSource" -ForegroundColor Red
    exit 1
}

# 源文件清单（只部署这些，别把 README 之类带进去）
$files = @('ribbon.xml', 'manifest.xml', 'js\main.js', 'panel.html')

if ($Check) {
    $diff = 0
    foreach ($f in $files) {
        $src = Join-Path $addonSource $f
        $dst = Join-Path $addonTarget ($f -replace '\\', '\')
        if (-not (Test-Path $dst)) {
            Write-Host "  [缺]   $f" -ForegroundColor Yellow
            $diff++
            continue
        }
        $a = Get-FileHash $src -Algorithm MD5
        $b = Get-FileHash $dst -Algorithm MD5
        if ($a.Hash -ne $b.Hash) {
            Write-Host "  [变更] $f" -ForegroundColor Yellow
            $diff++
        } else {
            Write-Host "  [一致] $f" -ForegroundColor DarkGray
        }
    }
    Write-Host ''
    if ($diff -eq 0) {
        Write-Host '✓ 部署位置与源一致，无需操作。' -ForegroundColor Green
    } else {
        Write-Host "⚠ 有 $diff 个文件不一致。去掉 -Check 参数即可部署。" -ForegroundColor Yellow
        Write-Host '  部署后必须【完全退出 WPS 再重开】才会生效。' -ForegroundColor Yellow
    }
    exit 0
}

# 确保 WPS 没在跑（否则文件被占用，复制失败或复制一半）
$wpsProcs = @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue)
if ($wpsProcs.Count -gt 0) {
    Write-Host "⚠ 检测到 WPS 正在运行（$($wpsProcs.Count) 个进程）。" -ForegroundColor Yellow
    Write-Host '  正在运行的 WPS 可能占用文件导致复制失败。' -ForegroundColor Yellow
    # 非交互：不该卡在 Read-Host 上（脚本串联 / 自动化调用会直接挂住）。
    # 默认直接关 —— 部署必须独占文件，且调用方跑了部署就等于同意重启 WPS。
    # 想保留 WPS 就先自己退出，再跑本脚本。
    foreach ($p in $wpsProcs) {
        try { Stop-Process -Id $p.Id -Force } catch {}
    }
    Start-Sleep -Seconds 2
    Write-Host '  已关闭 WPS（非交互模式默认关闭）。' -ForegroundColor Green
}

try {
    if (-not (Test-Path $addonTarget)) {
        New-Item -ItemType Directory -Path $addonTarget -Force | Out-Null
        Write-Host "✓ 已创建目标目录" -ForegroundColor Green
    }

    $copied = 0
    foreach ($f in $files) {
        $src = Join-Path $addonSource $f
        $dst = Join-Path $addonTarget $f
        $dstDir = Split-Path -Parent $dst
        if (-not (Test-Path $dstDir)) {
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        }
        Copy-Item -Path $src -Destination $dst -Force
        $hash = (Get-FileHash $dst -Algorithm MD5).Hash.Substring(0, 8)
        Write-Host "  ✓ $f  (md5 $hash)" -ForegroundColor Green
        $copied++
    }

    # publish.xml：WPS 的 JS 插件加载器靠它扫描。
    # 这里【合并】条目而不是整文件覆盖 —— 该文件是全机所有加载项的共享清单，
    # 覆盖会清掉别人厂商的 <jsplugin>（他们的功能区条目直接消失）。
    # 逻辑与 scripts/install.ps1 里的 Merge-PublishXmlEntry 相同（那份是权威拷贝，
    # 因为 install.ps1 要支持 irm|iex 单文件执行，不能依赖旁边的 lib 文件）。
    if (-not (Test-Path $publishXmlPath)) {
        $newContent = @'
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="wps-claude-addon" type="wps,et,wpp" url="wps-claude-addon_/" enable="enable_dev"/>
</jsplugins>
'@
        $newContent | Out-File -FilePath $publishXmlPath -Encoding utf8
        Write-Host '  ✓ publish.xml（新建）' -ForegroundColor Green
    } else {
        $raw = Get-Content -Raw -Path $publishXmlPath
        if ($raw -notmatch '<jsplugins') {
            Write-Host '  ⚠ publish.xml 已存在但不是加载项清单（没有 <jsplugins>），不动它' -ForegroundColor Yellow
        } else {
            # 删掉自己上次那条（幂等，重复部署不会堆条目），别人的行一个字符都不碰
            $cleaned = $raw -replace '(?m)^[ \t]*<jsplugin\s+name="wps-claude-addon"[^>]*/?>[ \t]*\r?\n?', ''
            $idx = $cleaned.LastIndexOf('</jsplugins>')
            if ($idx -lt 0) {
                Write-Host '  ⚠ publish.xml 里找不到 </jsplugins> 收尾，不动它' -ForegroundColor Yellow
            } else {
                $line = '  <jsplugin name="wps-claude-addon" type="wps,et,wpp" url="wps-claude-addon_/" enable="enable_dev"/>'
                # 结尾换行必须去掉：Out-File 自己会补一个行尾，带着原来的换行写会每次多一个空行
                (($cleaned.Substring(0, $idx) + $line + "`r`n" + $cleaned.Substring($idx)) -replace '\s+$', '') |
                    Out-File -FilePath $publishXmlPath -Encoding utf8
                Write-Host '  ✓ publish.xml（已合并，其他加载项条目保留）' -ForegroundColor Green
            }
        }
    }

    Write-Host ''
    Write-Host "✓ 部署完成，共 $copied 个文件。" -ForegroundColor Green
    Write-Host ''
    Write-Host '下一步：' -ForegroundColor Cyan
    Write-Host '  1) 完全退出 WPS（任务栏右键退出，不是关窗口）'
    Write-Host '  2) 重新打开 WPS，看功能区「多面手 MCP」'
    Write-Host '  3) 若仍是旧的：检查 WPS 是否用了别的用户配置目录'
} catch {
    Write-Host "✗ 部署失败: $_" -ForegroundColor Red
    exit 1
}
