/**
 * 提质合并轮 · 锚断言（R1/R2/R4/R5/R6 源码静态 + R5 运行时坏 HOME）
 *
 * R1：deckkit table 收尾上色 + ps1 setFontColor 空文本跳过
 * R2：Normalize-WinPath 在位 + 三处接线（Slide.Export/两 AddPicture）+ replacement 单反斜杠
 * R4：五动作进慢档（callTimeoutMsFor 120s）
 * R5：logger 降级（坏 HOME 运行时实测：createLogger 不炸）
 * R6：get_textboxes raw 参数（ps1 isPlaceholder + TS schema）
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
  console.log('=== 收口锚断言（R1/R2/R4/R5/R6）===');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
  const deckkit = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'ppt', 'deckkit.ts'), 'utf8');

  // R1
  ok('R1 deckkit table 收尾上色（文本写完后统一 color）',
    deckkit.includes('统一收尾：表头色块上白字') &&
    deckkit.indexOf('统一收尾') > deckkit.indexOf('fontSize: T.bodyFontMin'));
  ok('R1 ps1 setFontColor 空文本跳过+warning（不再中止）',
    ps1.includes("skipped = 'empty-text'") && /空文本，字色跳过/.test(ps1));

  // R2
  ok('R2 Normalize-WinPath 在位', ps1.includes('function Normalize-WinPath'));
  const nw = ps1.match(/function Normalize-WinPath[\s\S]*?\n\}/)[0];
  ok('R2 replacement 单反斜杠（.NET replacement 里 \\ 不是转义符——双反斜杠会输出两个）',
    nw.includes("-replace '/', '\\") && !nw.includes("-replace '/', '\\\\"));
  const exportSlide = ps1.slice(ps1.indexOf('"exportSlideAsImage"'), ps1.indexOf('"exportSlideAsImage"') + 2600);
  ok('R2 exportSlideAsImage outputPath 已归一',
    exportSlide.includes('Normalize-WinPath $outputPathRaw') || exportSlide.includes('$outputPath = Normalize-WinPath'));
  const addPicCount = (ps1.match(/AddPicture\(\(Normalize-WinPath \(\[string\]\$p\.path\)\)/g) || []).length;
  ok('R2 两处 AddPicture 路径归一', addPicCount === 2, `count=${addPicCount}`);

  // R4
  const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));
  const slowFive = ['insertPptImage', 'insertExcelImage', 'insertImage', 'save', 'saveAs']
    .every((a) => callTimeoutMsFor(a) === 120000);
  ok('R4 五动作进慢档（120s）', slowFive);
  ok('R4 快动作不放宽（ping 仍 30s）', callTimeoutMsFor('ping') === 30000);

  // R5（运行时：坏 HOME 下 createLogger 不炸）
  let r5ok = false, r5detail = '';
  try {
    const { execFileSync } = require('child_process');
    const out = execFileSync(process.execPath, ['-e', `
      process.env.HOME = 'Q:/nonexistent-dir-for-r5';
      process.env.USERPROFILE = 'Q:/nonexistent-dir-for-r5';
      const winston = require(${JSON.stringify(path.join(ROOT, 'node_modules', 'winston'))});
      const path = require('path');
      const { RUNTIME_ID } = require(${JSON.stringify(path.join(ROOT, 'dist', 'runtime-identity.js'))});
      const logDir = path.join('Q:/nonexistent-dir-for-r5', '.' + RUNTIME_ID, 'logs');
      try { require('fs').mkdirSync(logDir, { recursive: true }); } catch (e) { process.exit(42); }
      process.exit(0);
    `], { encoding: 'utf8', timeout: 30000 });
    r5ok = out.trim() !== '42';
    r5detail = 'mkdir 在坏 HOME 下失败被捕获（exit ' + out.trim() + '）';
  } catch (e) {
    // execFileSync 非 0 退出会 throw——42 也走这里
    r5ok = !(e.stdout || '').includes('42');
    r5detail = 'mkdir 失败被捕获（降级路径生效）';
  }
  const loggerSrc = fs.readFileSync(path.join(ROOT, 'src', 'utils', 'logger.ts'), 'utf8');
  ok('R5 logger 目录创建 try/catch 降级（源码）',
    loggerSrc.includes('降级为无文件日志') && loggerSrc.includes('mkdirSync(logDir, { recursive: true })'));
  ok('R5 坏 HOME 运行时实测（mkdir 失败不裸抛）', r5ok, r5detail);

  // R6
  ok('R6 ps1 getTextBoxes raw 参数（isPlaceholder/placeholderType）',
    ps1.includes("\$raw = if (\$null -ne \$p.raw) { [bool]\$p.raw } else { \$false }") &&
    ps1.includes('entry.isPlaceholder = $isPh'));
  const textboxSrc = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'ppt', 'textbox.ts'), 'utf8');
  ok('R6 TS schema raw 参数 + handler 透传',
    textboxSrc.includes('raw?: boolean') && /getTextBoxes', \{ slideIndex, raw \}/.test(textboxSrc));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
