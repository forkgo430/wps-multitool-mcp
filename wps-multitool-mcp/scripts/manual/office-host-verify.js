/**
 * Office 支持一期（host 维度）· 结构与试点断言（全 mock/静态，不起 WPS/Office）
 * 覆盖：
 *   A. host 解析纯函数（parseLocalServerBinary / resolveProgidHost）——含顶名劫持场景
 *   B. ps1 结构 gate：host 函数五件套 + 三 getter 接线 + Excel 记账链 + quirk 修复位点
 *   C. 试点 32 case 名单全部在路由表（word 10 / excel 12 / ppt 10）
 *   D. WPS 专用位点保持不变（showWps 等——WPS 专用标注一致性）
 * 跑法：node scripts/manual/office-host-verify.js（verify-all FAST 档包含它）
 */

const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

let pass = 0;
let fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
const routesSrc = fs.readFileSync(path.join(ROOT, 'src', 'client', 'action-routes.generated.ts'), 'utf8');
const { parseLocalServerBinary, resolveProgidHost } = require(
  path.join(ROOT, 'dist', 'server', 'backend-selector.js')
);

console.log('\n=== A. host 探测纯函数（顶名劫持防线）===');
ok('Office C2R 无引号形态 → office',
  parseLocalServerBinary('C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE /Automation').host === 'office');
ok('WPS 顶名劫持（官方 CLSID 指向 wps.exe）→ wps',
  parseLocalServerBinary('D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation').host === 'wps');
ok('未知二进制 → unknown（不猜）',
  parseLocalServerBinary('C:\\X\\Weird.exe /x').host === 'unknown');
// WPS 独装机劫持场景（注入 reg 读取器）：ProgID 在、CLSID 只在 WOW6432Node → 必须=wps
const regHijack = (key) => {
  if (key === 'HKCU\\SOFTWARE\\Classes\\Word.Application\\CLSID') return '{000209FF-0000-0000-C000-000000000046}';
  if (key === 'HKCU\\SOFTWARE\\Classes\\WOW6432Node\\CLSID\\{000209FF-0000-0000-C000-000000000046}\\LocalServer32') return 'D:\\APPdate\\WPS Office\\12.1.0.29680\\office6\\wps.exe /Automation';
  return null;
};
ok('WPS 独装劫持：resolveProgidHost → wps（不是 office）', resolveProgidHost('Word.Application', regHijack) === 'wps');

console.log('\n=== B. ps1 结构 gate ===');
// B1 host 函数五件套
for (const fn of ['function Get-HostKind', 'function Get-ProgidForApp', 'function Get-HostProcNames',
  'function Test-AppIdentity', 'function Resolve-ProgidHost', 'function Test-ServerBinaryHost']) {
  ok(`host 函数在位：${fn.replace('function ', '')}`, ps1.includes(fn));
}
// host wps 优先（2026-11-17 翻转）：Get-HostKind 双链探测结构断言
{
  const ghk = ps1.slice(ps1.indexOf('function Get-HostKind'), ps1.indexOf('function Get-ProgidForApp'));
  ok('Get-HostKind 双链：Office 三链在位', ['Word.Application', 'Excel.Application', 'PowerPoint.Application'].every((x) => ghk.includes(x)));
  ok('Get-HostKind 双链：WPS 三链在位', ['Kwps.Application', 'Ket.Application', 'Kwpp.Application'].every((x) => ghk.includes(x)));
  ok('Get-HostKind 双链：WPS_HOST_PREFERENCE 取舍在位', ghk.includes('WPS_HOST_PREFERENCE'));
}
// B2 三 getter 接线（host 感知 ProgID + office 身份校验）
ok('Get-WpsExcel 走 Get-ProgidForApp', /function Get-WpsExcel[\s\S]{0,900}Get-ProgidForApp 'excel'/.test(ps1));
ok('Get-WpsWord 走 Get-ProgidForApp', /function Get-WpsWord[\s\S]{0,2400}Get-ProgidForApp 'word'/.test(ps1));
ok('Get-WpsPpt 走 Get-ProgidForApp', /function Get-WpsPpt[\s\S]{0,700}Get-ProgidForApp 'ppt'/.test(ps1));
const identityCount = (ps1.match(/Test-AppIdentity \$rot/g) || []).length;
ok('三 getter 的 ROT 附着带身份校验（office 防顶名混入）', identityCount === 3, `count=${identityCount}`);
// B3 Excel 记账链（F3' 范式平移）
ok('Set-LastResolvedBook 在位', ps1.includes('function Set-LastResolvedBook'));
ok('Resolve-Book 记账层在位（显式名 > 记账 > Active > 遍历）',
  /function Resolve-Book[\s\S]{0,1200}LastResolvedBook/.test(ps1));
for (const anchor of ["Set-LastResolvedBook $wb 'create'", "Set-LastResolvedBook $wb 'open'", "Set-LastResolvedBook $wb 'switch'"]) {
  ok(`记账点在位：${anchor}`, ps1.includes(anchor));
}
// B4 quirk 修复位点
const intSlides = (ps1.match(/\.Slides\.Item\(\[int\]\$/g) || []).length;
ok('Slides.Item [int] 强转（Office 拒 VT_I8/pwsh7 JSON Int64）', intSlides >= 95, `count=${intSlides}`);
const intCells = (ps1.match(/\.Cells\.Item\(\[int\]\$p\.row, \[int\]\$p\.col\)/g) || []).length;
ok('Cells.Item row/col [int] 强转', intCells === 3, `count=${intCells}`);
for (const c of ['"setCellValue"', '"getRangeData"', '"getCellValue"']) {
  ok(`${c} 走 Resolve-Sheet（Sheets.Item(null) quirk）`,
    new RegExp(c.replace(/"/g, '"') + ' \\{[\\s\\S]{0,600}Resolve-Sheet').test(ps1));
}
const nfOffice = (ps1.match(/if \(\(Get-HostKind\) -ne 'office'\) \{ \$cell\.NumberFormat = 'General' \}/g) || []).length;
ok("NumberFormat='General' office 分支（Office 拒 General 重置）", nfOffice === 4, `count=${nfOffice}`);
ok("createChart 字符串 chartType 过枚举表（Office 严格绑定）",
  /if \(\$chartType -is \[string\]\) \{\s*\$chartType = \$chartTypes\[\$chartType\]/.test(ps1));
// B5 试点教训位点：Test-AppIdentity 瞬态不定态放行
ok('Test-AppIdentity 瞬态忙 ≠ wrong host（不定态放行）',
  /catch \{ return \$true \}[\s\S]{0,60}读不到值：不定态/.test(ps1.replace(/\r\n/g, '\n')) || /瞬态忙：不定态，放行/.test(ps1));

console.log('\n=== C. 试点 32 case 名单全部在路由表 ===');
const routes = {};
for (const m of routesSrc.matchAll(/"([A-Za-z0-9_]+)":\s*'(\w+)'/g)) routes[m[1]] = m[2];
const pilot = {
  word: ['createDocument', 'insertText', 'addHeading', 'setFont', 'insertTable', 'getTableData',
    'getDocumentText', 'getDocumentStats', 'addPageNumber'],
  excel: ['createWorkbook', 'setCellValue', 'getCellValue', 'setRangeData', 'getRangeData', 'setFormula',
    'setCellFormat', 'createChart', 'createExcelTable', 'autoFilter', 'sortRange'],
  ppt: ['createPresentation', 'addSlide', 'addTextBox', 'setTextBoxStyle', 'insertPptTable',
    'insertPptChart', 'getShapes', 'getSlideInfo', 'exportSlideAsImage'],
};
// convertToPDF 是 common 组 case（三应用共用导出），试点在三个 app 下各用一次
ok('convertToPDF 在 common 组（word/excel/ppt 共用导出）', routes['convertToPDF'] === 'common');
let pilotTotal = 0;
for (const app of Object.keys(pilot)) {
  const missing = pilot[app].filter((a) => routes[a] !== app);
  pilotTotal += pilot[app].length;
  ok(`试点 ${app} ${pilot[app].length} case 全在路由表（app 归属正确）`, missing.length === 0,
    missing.length ? '缺: ' + missing.join(',') : `${pilot[app].length}/${pilot[app].length}`);
}
ok('试点口径 = 32 run（29 app 专属 + convertToPDF×3 复用）', pilotTotal === 29 && routes['convertToPDF'] === 'common', `app专属=${pilotTotal}`);

console.log('\n=== D. WPS 专用位点一致性 ===');
ok('showWps 保持 WPS ProgID（矩阵标注 WPS 专用，不做 host 分支）',
  /"showWps" \{[\s\S]{0,1200}Kwps\.Application/.test(ps1));

// ==================== 二期（Office 支持二期 2026-11-07）====================
console.log('\n=== E. 二期 quirk 落地断言 ===');
// E1: Resolve-Sheet 化 11 处（cleanData/switchSheet + 9 excel case）
const rsCount = (ps1.match(/# 统一 Resolve-Sheet 化（与 setCellValue 同款）。/g) || []).length;
ok('Resolve-Sheet 化 11 处（二期 quirk ⑦）', rsCount === 11, `count=${rsCount}`);
// E2: try Shapes.Title（setSlideTitle/getSlideTitle/addSlide）
const ttCount = (ps1.match(/try \{ \$null = \$slide\.Shapes\.Title \} catch \{ \$titled = \$false \}/g) || []).length;
ok('try Shapes.Title 三方可靠判定（HasTitle 两形态皆不可靠）', ttCount >= 3, `count=${ttCount}`);
ok('HasTitle 判断已清除（静默假成功根除）', !/\$slide\.HasTitle\)/.test(ps1.replace(/#[^\n]*/g, '')));
// E3: copySheet [Type]::Missing
const missCount = (ps1.match(/\$sheet\.Copy\(\[Type\]::Missing,/g) || []).length;
ok('copySheet Copy $null 参数 → [Type]::Missing（Office 拒 VT_NULL）', missCount === 2, `count=${missCount}`);
// E4: setPageSetup 单位
ok('setPageSetup 边距 office host 不换行（×28.35 是 WPS 语义）',
  /\$marginScale = if \(\(Get-HostKind\) -eq 'office'\) \{ 1\.0 \} else \{ 28\.35 \}/.test(ps1));
// E5: setShapeRoundness 手柄防护
ok('setShapeRoundness 手柄防护（Adjustments.Count=0 明确报错，不静默假成功）',
  /Adjustments\.Count=0[^\n]*圆角请对圆角矩形/.test(ps1) && /if \(\$adjCount -lt 1\) \{/.test(ps1));
// E6: freezePanes 补窗口
ok('freezePanes headless 补窗口等效实现（ActiveWindow=null 明确报错）',
  /Excel 无活动窗口（headless 实例），冻结窗格需要活动窗口/.test(ps1) && /\$win\.FreezePanes = \$true/.test(ps1));
// E7: saveAs DisplayAlerts 护栏
const daCount = (ps1.match(/try \{ \$excel\.DisplayAlerts = \$false \} catch \{\}/g) || []).length;
ok('saveAs DisplayAlerts 护栏（同名覆盖确认框卡死通道；renderTemplateXlsx 同款护栏加入后总数增）', daCount >= 1, `count=${daCount}`);
// E8: convertToPDF app 短路
ok('convertToPDF app 参数短路（excel/word/ppt/auto，缺省 auto）',
  /\$want -notin @\('excel', 'word', 'ppt', 'auto'\)/.test(ps1) &&
  /\(\$want -eq 'auto' -or \$want -eq 'excel'\)/.test(ps1) &&
  /\(\$want -eq 'auto' -or \$want -eq 'word'\)/.test(ps1) &&
  /\(\$want -eq 'auto' -or \$want -eq 'ppt'\)/.test(ps1));
// E9: Resolve-Doc 记账对称化
ok('Resolve-Doc 记账消费层（显式名 > 记账 > Active > 遍历）',
  /Set-LastResolvedDoc \$cand '记账'/.test(ps1) &&
  /function Resolve-Doc[\s\S]{0,1600}script:LastResolvedDoc/.test(ps1));
ok('createDocument 记账点（created，收口轮 A3 健康门控版）',
  /Set-LastResolvedDoc \$curDoc 'created'/.test(ps1) &&
  /else \{ \$script:LastResolvedDoc = \$null \}/.test(ps1));
ok('openDocument 记账点（open）', /Set-LastResolvedDoc \$docItem 'open'/.test(ps1));
// E10: WPS 专用守卫
ok('WPS 专用 11 office host 守卫（明确报错文案）',
  /wpsOnlyActions = @\('showWps'[\s\S]{0,200}'__probeDeckkit'\)/.test(ps1) &&
  /WPS 专用，Office host 不支持/.test(ps1));
// E11: findInSheet office 分支（收口轮 B1：5.1 子进程——daemon STA 会话态防线）
ok('findInSheet office 分支走 5.1 子进程（E11 收口版）',
  /office host 的查找交由 PowerShell 5.1 子进程/.test(ps1) && /wps-find51.ps1/.test(ps1));

console.log('\n=== F. 对账表一致性门禁（分母 314 不变（3D op-merge 后））===');
const matrixPath = path.join(ROOT, '..', 'docs', 'agents', 'office-compat-matrix.md');
if (fs.existsSync(matrixPath)) {
  const matrix = fs.readFileSync(matrixPath, 'utf8');
  const rowMatches = matrix.match(/^\| `[A-Za-z0-9_]+` \|/gm) || [];
  ok('对账表行数 = 314（与 gen-action-routes 分母一致）', rowMatches.length === 314, `rows=${rowMatches.length}`);
  ok('对账表四类统计行齐备（直通/已落地/未落地/WPS 专用）',
    matrix.includes('需 quirk 分支(已落地)') && matrix.includes('需 quirk 分支(未落地)') &&
    matrix.includes('WPS 专用') && matrix.includes('| **合计** | **314** |'));
  // 表内 case 与路由表一一对应（不许漂）
  const matrixCases = new Set((matrix.match(/^\| `([A-Za-z0-9_]+)` \|/gm) || []).map((r) => r.match(/`([A-Za-z0-9_]+)`/)[1]));
  const routeCases = [...routesSrc.matchAll(/"([A-Za-z0-9_]+)":\s*'(\w+)'/g)].map((m) => m[1]);
  const missingInMatrix = routeCases.filter((c) => !matrixCases.has(c));
  ok('路由表 314 case 全部在对账表（无漂移）', missingInMatrix.length === 0 && routeCases.length === 314,
    missingInMatrix.length ? '缺: ' + missingInMatrix.slice(0, 5).join(',') : `314/314`);
} else {
  ok('对账表存在', false, matrixPath + ' 不存在');
}
// 台账 officeHost 维度
const ledger = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'manual', 'tool-verification.json'), 'utf8'));
ok('台账 officeHost 维度存在', !!ledger.officeHost && !!ledger.officeHost.live);
const ohCount = Object.keys(ledger.officeHost?.live || {}).length;
ok('台账 officeHost.live ≥ 60（一期 31 + 二期扩条）', ohCount >= 60, `count=${ohCount}`);

// ==================== 收口合并轮（A 绑定一致性 / B findInSheet / C 孤儿桥）====================
console.log('\n=== G. 收口轮断言（绑定一致性 / findInSheet 5.1 子进程 / 孤儿桥监测）===');
// G1: A2 取序——三 getter 缓存探活优先于 ROT
const cacheFirst = (ps1.match(/① 会话缓存探活通过即优先复用（Office 收口合并卡 A 段/g) || []).length;
ok('三 getter 缓存探活优先于 ROT（同会话绑定一致性）', cacheFirst === 3, `count=${cacheFirst}`);
ok('取序后 ROT 层保留身份校验（Test-AppIdentity 一行不动）',
  /② ROT 附着（缓存空\/失效时）[\s\S]{0,400}Test-AppIdentity/.test(ps1));
// G2: A3 记账健康门控
ok('createDocument 记账健康门控（未就绪清账，不记空壳）',
  /if \(\$null -ne \$curDoc -and \(Test-DocHealthy \$curDoc\)\) \{ Set-LastResolvedDoc \$curDoc 'created' \}/.test(ps1));
// G3: B1 find51 子进程
const fs51 = fs.existsSync(path.join(ROOT, 'scripts', 'wps-find51.ps1'));
ok('wps-find51.ps1 存在（5.1 子进程查找辅助）', fs51);
if (fs51) {
  const f51 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-find51.ps1'), 'utf8');
  ok('find51 用 Areas 枚举（多区域 Range 终极破案——Item 带参调用会挂）',
    /foreach \(\$area in @\(\$scope\.Areas\)\)/.test(f51));
  ok('find51 用 Array.GetValue（COM SAFEARRAY 非零基——5.1 索引器会炸）',
    /\$raw\.GetValue\(\$lb0 \+ \$r - 1, \$lb1 \+ \$c - 1\)/.test(f51));
  ok('find51 无 Address 两参调用（缺参 COM 调用不稳定家族）',
    !/\.Address\(\$false, \$false\)/.test(f51));
}
ok('findInSheet office 分支走 5.1 子进程（wps-find51.ps1）',
  /office host 的查找交由 PowerShell 5\.1 子进程[\s\S]{0,200}wps-find51\.ps1/.test(ps1));
// G4: C 段 EPHEMERAL 父存活监测（bridge/index.ts 源码）
const bridgeSrc = fs.readFileSync(path.join(ROOT, 'src', 'bridge', 'index.ts'), 'utf8');
ok('桥 EPHEMERAL 父存活监测在位（process.ppid + kill 0 探测）',
  /const EPHEMERAL = process\.env\.WPS_BRIDGE_EPHEMERAL === '1';/.test(bridgeSrc) &&
  /process\.kill\(parentPid, 0\)/.test(bridgeSrc) &&
  /ESRCH/.test(bridgeSrc));
ok('默认桥不启用监测（活过 server 重启是设计意图）',
  /【默认（非 EPHEMERAL）桥不启用本监测】/.test(bridgeSrc) ||
  /default.*不启用|非 EPHEMERAL.*不启用/.test(bridgeSrc.replace(/\r\n/g, '\n')));
// G5: 验收脚本入库
ok('word-binding-repro.js 钉板脚本入库', fs.existsSync(path.join(ROOT, 'scripts', 'manual', 'word-binding-repro.js')));
ok('binding-accept.js 双场景验收入库', fs.existsSync(path.join(ROOT, 'scripts', 'manual', 'binding-accept.js')));
ok('orphan-bridge-verify.js 验收入库', fs.existsSync(path.join(ROOT, 'scripts', 'manual', 'orphan-bridge-verify.js')));

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);
