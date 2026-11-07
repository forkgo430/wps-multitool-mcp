# Input: findInSheet 参数 JSON（-Params '{"what":"x","sheet":"S1","range":"A1:C9",...}'）
# Output: 与 findInSheet 回包 data 同形的 JSON（keyword/count/matches/scope）
# Pos: Office 收口合并卡 B1 · findInSheet 的 5.1 专用查找辅助——pwsh7 daemon 的 COM binder
#      对循环/带参调用会话态敏感（四轮四错在案），本脚本由 PowerShell 5.1 子进程执行
#      （5.1 binder 干净），附着 ROT Excel 后在内存层完成查找。
#      ⚠ B1 实锤两坑：①COM SAFEARRAY 非零基——5.1 的 $raw[$r,$c] 索引器炸，必须
#      Array.GetValue(r,c)；②Address/Find/FindNext 等"缺参"COM 方法调用对本机 Excel
#      晚绑定不稳定——全部绕开（Address 用列字母拼接）。一旦我被修改，请更新我的头部注释。
# 编码：必须存成 UTF-8 with BOM（5.1 对无 BOM 按 ANSI 解码，中文参数/输出会乱）。
param([string]$Params = '{}')

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }
$what = [string]$p.what
$matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
$wholeCell = if ($null -ne $p.wholeCell) { [bool]$p.wholeCell } else { $false }
$maxResults = 0
if ($p.maxResults) { try { $maxResults = [int]$p.maxResults } catch {} }

try {

    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')

    $wb = $excel.ActiveWorkbook

    if ($null -eq $wb) { throw 'No active workbook' }
    if ($p.sheet) {
        $sheet = $null
        try { $sheet = $wb.Sheets.Item([string]$p.sheet) } catch { throw "sheet not found: $($p.sheet)" }
    } else {
        $sheet = $wb.ActiveSheet
        if ($null -eq $sheet) { throw 'No active sheet' }
    }

    $scope = if ($p.range) { $sheet.Range([string]$p.range) } else { $sheet.UsedRange }

    # ⚠⚠ 二期 B1 终极破案（2026-11-07）：跨区域格式/批注操作会把 UsedRange 变成
    # 【多区域 Range】（Areas.Count>1）——多区域 Range 的 Row/Column/Value2 语义
    # 完全不同：Row 返回各区域行号的【数组】（[int] 转换即炸"Object[]→Int32"），
    # Value2 返回扁平一维数组。二期抽验"四轮四错"全部是这一个根因的变体。
    # 修法 = 按 Areas 枚举（IEnumVARIANT 集合枚举——Areas.Item(ai) 是带参调用，
    # 同属"缺参/带参 COM 调用不稳定"家族且实测会挂，必须用 foreach 枚举），
    # 逐【单区域】读 Row/Column/Value2（单区域语义标准）。
    $areaCount = 1
    try { $areaCount = [int]$scope.Areas.Count } catch { $areaCount = 1 }

    function Convert-ToLetter([int]$num) {
        $letter = ''
        $n = $num
        while ($n -gt 0) { $m = ($n - 1) % 26; $letter = [char](65 + $m) + $letter; $n = [int](($n - $m - 1) / 26) }
        return $letter
    }

    $found = @()
    $scopeAddrParts = @()
    $totalCells = 0
    $areaIndex = 0
    foreach ($area in @($scope.Areas)) {
        $areaIndex++
        if ($areaCount -eq 1) { $area = $scope }
        $r0 = [int]$area.Row
        $c0 = [int]$area.Column
        $rows = [int]$area.Rows.Count
        $cols = [int]$area.Columns.Count

        $totalCells += $rows * $cols
        if ($totalCells -gt 200000) { throw "查找范围过大（$totalCells 格），请用 range 参数缩小" }
        $scopeAddrParts += ("{0}{1}:{2}{3}" -f (Convert-ToLetter $c0), $r0, (Convert-ToLetter ($c0 + $cols - 1)), ($r0 + $rows - 1))

        $raw = $area.Value2
        # Value2 归一（单区域：标量 / 一维 / 二维）。
        # ⚠ COM SAFEARRAY 非零基——5.1 必须用 Array.GetValue（$raw[$r,$c] 索引器炸）。
        $grid = @()
        if ($null -eq $raw) { $grid = ,@($null) }
        elseif ($raw -is [object[,]]) {
            $lb0 = $raw.GetLowerBound(0); $lb1 = $raw.GetLowerBound(1)
            for ($r = 1; $r -le $rows; $r++) {
                $rowArr = @()
                for ($c = 1; $c -le $cols; $c++) { $rowArr += $raw.GetValue($lb0 + $r - 1, $lb1 + $c - 1) }
                $grid += ,$rowArr
            }
        } elseif ($raw -is [array]) {
            if ($rows -le 1) { $grid = ,@($raw) } else { foreach ($v in $raw) { $grid += ,@($v) } }
        } else { $grid = ,@($raw) }

        for ($r = 0; $r -lt $grid.Count; $r++) {
            for ($c = 0; $c -lt $grid[$r].Count; $c++) {
                $v = $grid[$r][$c]
                if ($null -eq $v) { continue }
                $s = [string]$v
                $hit = if ($matchCase) { $s.Contains($what) } else { $s.IndexOf($what, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
                if (-not $hit) { continue }
                if ($wholeCell) {
                    $equal = if ($matchCase) { $s -ceq $what } else { $s -ieq $what }
                    if (-not $equal) { continue }
                }
                $ar = $r0 + $r
                $ac = $c0 + $c
                $found += @{ address = ("{0}{1}" -f (Convert-ToLetter $ac), $ar); row = $ar; column = $ac; value = $s; formula = $null }
                if ($maxResults -gt 0 -and $found.Count -ge $maxResults) { break }
            }
            if ($maxResults -gt 0 -and $found.Count -ge $maxResults) { break }
        }
        if ($maxResults -gt 0 -and $found.Count -ge $maxResults) { break }
    }

    $scopeAddr = $scopeAddrParts -join ','

    @{ success = $true; data = @{ keyword = $what; count = $found.Count; matches = $found; scope = $scopeAddr } } | ConvertTo-Json -Depth 5 -Compress
} catch {
    Write-Output "[ERR] $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    @{ success = $false; error = "find51: $($_.Exception.Message)" } | ConvertTo-Json -Compress
}
