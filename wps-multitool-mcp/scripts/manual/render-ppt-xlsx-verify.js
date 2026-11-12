/**
 * 模板三格式补全卡 · FAST 门禁（ppt/xlsx 渲染路径，静态+mock 形态）
 *
 * 判据：
 *   1. ps1 两 case 在位（renderTemplatePpt/renderTemplateXlsx），路径归一 + DisplayAlerts 护栏 + 缺 key 列全
 *   2. ppt：组合递归（GroupItems）+ TextRange.Replace 标量两参 + 5.1 子进程收口（daemon STA 会话态防线）
 *   3. xlsx：多区域 Value2 归一 + GetValue（SAFEARRAY 非零基）+ Range.Replace + ReadOnly 打开
 *   4. TS format 分发：auto 缺省按扩展名三路（docx 行为不变）；慢档四动作 120s
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== 模板三格式补全卡 FAST 门禁 ===');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
  const pptCase = ps1.slice(ps1.indexOf('"renderTemplatePpt"'), ps1.indexOf('"renderTemplateXlsx"'));
  const xlsxCase = ps1.slice(ps1.indexOf('"renderTemplateXlsx"'), ps1.indexOf('"getTextBoxes"'));

  // 1) 两 case 在位 + 共同纪律
  ok('ps1 case renderTemplatePpt 在位', ps1.includes('"renderTemplatePpt"'));
  ok('ps1 case renderTemplateXlsx 在位', ps1.includes('"renderTemplateXlsx"'));
  ok('R2 路径归一（Normalize-WinPath 入口）',
    (pptCase.match(/Normalize-WinPath/g) || []).length >= 2 && (xlsxCase.match(/Normalize-WinPath/g) || []).length >= 2);
  ok('outputPath 已存在拒绝（两 case）',
    pptCase.includes('不静默覆盖') && xlsxCase.includes('不静默覆盖'));
  ok('缺 key 列全（phKeys 收集 + missing 清单）',
    pptCase.includes('$phKeys') && xlsxCase.includes('$phKeys') &&
    pptCase.includes('$missing') && xlsxCase.includes('$missing'));

  // 2) ppt 特有
  ok('ppt 组合递归（GroupItems msoGroup=6）', pptCase.includes('msoGroup') && pptCase.includes('GroupItems'));
  ok('ppt TextRange.Replace 标量两参', pptCase.includes('$sh.TextFrame.TextRange.Replace($needle, [string]$prop.Value)'));
  ok('ppt 5.1 子进程收口（daemon STA 会话态防线）', ps1.includes('wps-render-ppt51.ps1'));
  const f51 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-render-ppt51.ps1'), 'utf8');
  ok('wps-render-ppt51.ps1 在位（BOM + 自包含 CoCreate）', f51.length > 500 && f51.includes('New-Object -ComObject $progId') && f51.includes('$ppt.Quit()'));
  ok('ppt51 按 progId 附着（WPS 宿主 Kwpp / Office 宿主 PowerPoint）',
    f51.includes("$progId = if ($p.progId)") && f51.includes('GetActiveObject($progId)') === false || f51.includes("GetActiveObject($progId)"));

  // 3) xlsx 特有
  ok('xlsx Value2 归一（SAFEARRAY 非零基 GetValue）',
    xlsxCase.includes('$raw.GetValue($lb0 + $r - 1, $lb1 + $c - 1)'));
  ok('xlsx ReadOnly 打开模板（不写坏源文件）', xlsxCase.includes('$wb = $excel.Workbooks.Open($templatePath, $null, $true)'));
  ok('xlsx Range.Replace 标量两参', xlsxCase.includes('$scope.Replace($needle, [string]$prop.Value)'));
  ok('xlsx DisplayAlerts 护栏', xlsxCase.includes('$excel.DisplayAlerts = $false'));

  // 4) TS format 分发
  const templateSrc = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'common', 'template.ts'), 'utf8');
  ok('TS format 参数（auto/docx/pptx/xlsx）', templateSrc.includes("format?: 'auto' | 'docx' | 'pptx' | 'xlsx'"));
  ok('TS auto 按扩展名分发三路', templateSrc.includes("ext === 'pptx' ? 'pptx' : ext === 'xlsx' ? 'xlsx' : 'docx'"));
  ok('docx 路径行为不变（renderTemplate 主调用保留）', templateSrc.includes("}>(" + "'renderTemplate'"));
  ok('ppt 分发走 renderTemplatePpt', templateSrc.includes("action = fmt === 'pptx' ? 'renderTemplatePpt'"));
  ok('xlsx 分发走 renderTemplateXlsx', templateSrc.includes("'renderTemplatePpt' : 'renderTemplateXlsx'"));

  // 5) 慢档四动作
  const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));
  ok('renderTemplate/renderTemplatePpt/renderTemplateXlsx 慢档 120s',
    callTimeoutMsFor('renderTemplate') === 120000 && callTimeoutMsFor('renderTemplatePpt') === 120000 && callTimeoutMsFor('renderTemplateXlsx') === 120000);
  ok('ping 快档 30s 不放宽', callTimeoutMsFor('ping') === 30000);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
