/**
 * 模板层大卡 B 段 · wps_extract_template FAST 断言（静态 + mock 形态，零 WPS COM）
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
  console.log('=== wps_extract_template 断言 ===');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
  const caseBody = ps1.slice(ps1.indexOf('"extractTemplate"'), ps1.indexOf('"renderTemplate"') === ps1.indexOf('"extractTemplate"') - 1 ? undefined : ps1.indexOf('"extractTemplate"') + 1);
  const body = ps1.slice(ps1.indexOf('"extractTemplate"'), ps1.indexOf('"extractTemplate"') + 3400);

  // 边界如实：v1 只做 word（Get-WpsWord 宿主）
  ok('B 边界：v1 word 宿主（Get-WpsWord）', body.includes('Get-WpsWord'));
  // 复制保样式（对副本逐段替换，样式随段落保留）
  ok('实现方式：Copy-Item 复制源文档为模板底稿（样式/版式保留）', body.includes('Copy-Item -LiteralPath $srcPath -Destination $outputPath'));
  ok('正文段 → {{占位}}（Range.Text 赋值保样式）', body.includes('$rng.Text = "{{$key}}"'));
  // 标题结构保留
  ok('标题层级保序（heading 记锚点清单，不转占位）', body.includes("kind = 'heading'") && body.includes('$isHeading'));
  // 占位上限
  ok('maxSlots 占位上限（默认 200 ≤500）', body.includes('$maxSlots = 200') && body.includes('[Math]::Min(500'));
  // 路径归一 + 已存在拒绝
  ok('路径归一（Normalize-WinPath 入口）', (body.match(/Normalize-WinPath/g) || []).length >= 2);
  ok('outputPath 已存在拒绝（不静默覆盖）', body.includes('不静默覆盖'));
  // 回包：锚点清单 + 占位计数
  ok('回包 anchors 清单 + placeholderCount/headingCount', body.includes('placeholderCount') && body.includes('headingCount'));
  // TS 壳注册
  const templateSrc = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'common', 'template.ts'), 'utf8');
  ok('TS 壳 wps_extract_template 注册（common 组）', templateSrc.includes("name: 'wps_extract_template'"));
  // 慢档
  const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));
  ok('extractTemplate 慢档 120s（R4 防复发）', callTimeoutMsFor('extractTemplate') === 120000);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
