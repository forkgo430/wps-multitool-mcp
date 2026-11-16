# Input: renderTemplatePpt 参数 JSON（-Params '{"templatePath":"...","outputPath":"...","data":{...}}'）
# Output: 与 renderTemplatePpt 回包 data 同形的 JSON
# Pos: 模板三格式补全卡 · renderTemplatePpt 的 5.1 专用渲染辅助——pwsh7 daemon 的 STA 会话态
#      COM 绑定在"Close 后再 Open"场景返回 null/调用炸（B1 findInSheet 同款家族），5.1 子进程
#      附着 ROT PPT 宿主后全链执行（5.1 绑定干净）。零新依赖；路径由调用方归一。
#      一旦我被修改，请更新我的头部注释。
# 编码：必须存成 UTF-8 with BOM。
param([string]$Params = '{}')

# 参数支持 base64（调用方 base64 过命令行防引号/反斜杠折叠）；
# 以 '{' 开头视为裸 JSON（老调用兼容）。
$ParamsDecoded = $Params
try {
    if ($Params -and -not $Params.StartsWith('{')) {
        $ParamsDecoded = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Params))
    }
} catch { $ParamsDecoded = $Params }

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try { $p = $ParamsDecoded | ConvertFrom-Json } catch { $p = @{} }

try {
    # 自包含 CoCreate（不附着 ROT——WPS 宿主无文档时不上 ROT，MK_E_UNAVAILABLE 实测）
    $progId = if ($p.progId) { [string]$p.progId } else { 'PowerPoint.Application' }
    $ppt = New-Object -ComObject $progId
    if ($null -eq $ppt) { throw "CoCreate  failed" }
    try { $ppt.Visible = $false } catch {}
    if ($null -eq $p.data -or $p.data -isnot [PSCustomObject]) { throw 'data required（JSON 对象）' }
    $dataProps = @($p.data.PSObject.Properties)
    if ($dataProps.Count -eq 0) { throw 'data 为空对象' }

    # 空集合撑起（实测：Presentations 集合为空时 ReadOnly 无窗 Open 必返 null，先撑起集合再 Open）
    if ([int]$ppt.Presentations.Count -eq 0) {
        try { $null = $ppt.Presentations.Add(1) } catch { }
        try { $ppt.Presentations.Item(1).Close() } catch { }
    }

    $pres = $ppt.Presentations.Open([string]$p.templatePath, $true, $false, $false)   # ReadOnly/无窗
    if ($null -eq $pres) { throw "模板打开失败: $($p.templatePath)" }

    # 递归收集文本形状（组合/嵌套 GroupItems 递归）
    function Collect-TextShapes([object]$shapes, [object]$acc) {
        for ($i = 1; $i -le $shapes.Count; $i++) {
            $sh = $shapes.Item($i)
            if ([int]$sh.Type -eq 6) {
                try { Collect-TextShapes $sh.GroupItems $acc } catch { $acc.skipped += $i }
                continue
            }
            $has = $false
            try { $has = $sh.HasTextFrame -and $sh.TextFrame.HasText } catch {}
            if ($has) { [void]$acc.shapes.Add($sh) }
        }
    }
    $acc = @{ shapes = (New-Object System.Collections.ArrayList); skipped = 0 }
    for ($si = 1; $si -le $pres.Slides.Count; $si++) {
        Collect-TextShapes $pres.Slides.Item($si).Shapes $acc
    }

    # 先扫 {{key}} 集合
    $allText = ""
    foreach ($sh in $acc.shapes) { try { $allText += [string]$sh.TextFrame.TextRange.Text + [string][char]10 } catch {} }
    $phKeys = @()
    foreach ($m in @([regex]::Matches($allText, '\{\{([A-Za-z0-9_\-]+)\}\}'))) {
        if ($phKeys -notcontains $m.Groups[1].Value) { $phKeys += $m.Groups[1].Value }
    }

    # 替换：TextRange.Replace(FindWhat, ReplaceWhat) 标量两参
    $filled = @()
    $missing = @()
    foreach ($prop in $dataProps) {
        $needle = '{{' + $prop.Name + '}}'
        if ($phKeys -notcontains $prop.Name) { continue }
        $replaced = 0
        foreach ($sh in $acc.shapes) {
            try {
                $t = [string]$sh.TextFrame.TextRange.Text
                if ($t.Contains($needle)) {
                    $null = $sh.TextFrame.TextRange.Replace($needle, [string]$prop.Value)
                    $replaced++
                }
            } catch { }
        }
        if ($replaced -gt 0) { $filled += @{ anchor = $prop.Name; via = 'placeholder'; shapes = $replaced } }
        else { $missing += $prop.Name }
    }
    foreach ($k in $phKeys) {
        $given = $dataProps | Where-Object { $_.Name -eq $k } | Select-Object -First 1
        if ($null -eq $given -and -not ($filled | Where-Object { $_.anchor -eq $k })) { $missing += $k }
    }

    # SaveAs（DisplayAlerts 护栏）
    $pptApp = $ppt
    $da0 = $null
    try { $da0 = $pptApp.DisplayAlerts } catch {}
    try { $pptApp.DisplayAlerts = 0 } catch {}
    try { $pres.SaveAs([string]$p.outputPath) } finally { try { if ($null -ne $da0) { $pptApp.DisplayAlerts = $da0 } } catch {} }
    $pres.Close()
    try { $ppt.Quit() } catch {}
    $verified = $false; $size = 0
    try { if (Test-Path ([string]$p.outputPath)) { $size = (Get-Item ([string]$p.outputPath)).Length; $verified = ($size -gt 0) } } catch {}
    if (-not $verified) { throw "渲染完成但落盘校验失败: $($p.outputPath)" }

    @{ success = $true; data = @{ outputPath = [string]$p.outputPath; size = $size; filled = $filled; skipped = @($missing | Select-Object -Unique); filledCount = @($filled).Count } } | ConvertTo-Json -Depth 5 -Compress
} catch {
    @{ success = $false; error = "render-ppt51: $($_.Exception.Message)" } | ConvertTo-Json -Compress
}
