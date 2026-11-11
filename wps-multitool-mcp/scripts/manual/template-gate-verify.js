/**
 * 模板层大卡 · 防复发 gate（实地六缺陷教训断言）+ 双宿主台账标记
 *
 * 判据：
 *   1. R2 教训：模板两 case 的路径参数入口归一（Normalize-WinPath）
 *   2. R4 教训：render/extract 慢档 120s
 *   3. R1 教训：setFontColor 空文本跳过（deckkit 收尾上色的兜底）
 *   4. 提质小卡先例：text 主载体（render/extract 回包是人类可读行文，非 JSON 文本）
 *   5. schema 短描述纪律：两新工具 property description 均为"名+类型+约束"（无多行长句）
 *   6. 双宿主台账：officeHost.live 含 wps_render_template/wps_extract_template（双维度标记结构在）
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
  console.log('=== 模板层 gate（教训防复发 + 双宿主台账）===');
  const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
  const rt = ps1.slice(ps1.indexOf('"renderTemplate"'), ps1.indexOf('"extractTemplate"'));
  const et = ps1.slice(ps1.indexOf('"extractTemplate"'), ps1.indexOf('"extractTemplate"') + 3600);

  ok('① R2：renderTemplate 路径归一（templatePath+outputPath）', (rt.match(/Normalize-WinPath/g) || []).length >= 2);
  ok('① R2：extractTemplate 路径归一（path+outputPath）', (et.match(/Normalize-WinPath/g) || []).length >= 2);
  const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));
  ok('② R4：render/extract 慢档 120s', callTimeoutMsFor('renderTemplate') === 120000 && callTimeoutMsFor('extractTemplate') === 120000);
  ok('③ R1：setFontColor 空文本跳过兜底在位', ps1.includes("skipped = 'empty-text'"));
  const templateSrc = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'common', 'template.ts'), 'utf8');
  ok('④ text 主载体：render/extract 回包为人类可读行文（非 JSON 文本）',
    !templateSrc.includes("content: [{ type: 'text', text: JSON.stringify"));
  // ⑤ schema 短描述：两新工具 schema 里无多行长句（换行符的 description = 长句标志）
  const badDesc = (templateSrc.match(/description: `[^\n]*\n/g) || []).filter((d) => d.includes('inputSchema') === false);
  const schemaSection = templateSrc.slice(templateSrc.indexOf('inputSchema'), templateSrc.lastIndexOf('templateTools'));
  ok('⑤ schema 短描述纪律（property description 无多行长句）',
    !(schemaSection.match(/description: `[^\n]*\n/) || []).length);
  // ⑥ 双宿主台账结构
  const ledger = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'manual', 'tool-verification.json'), 'utf8'));
  ok('⑥ officeHost.live 含 wps_render_template（双宿主标记结构在）',
    !!ledger.officeHost && !!ledger.officeHost.live['wps_render_template']);
  ok('⑥ officeHost.live 含 wps_extract_template', !!ledger.officeHost.live['wps_extract_template']);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
