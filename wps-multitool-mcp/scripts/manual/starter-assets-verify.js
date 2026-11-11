/**
 * 模板层大卡 D 段 · starter 资产三件断言（files 对账 + 锚点预置内容）
 *
 * 判据：
 *   1. templates/ 三件在仓（weekly-report.docx / briefing.pptx / ledger.xlsx，均 >0 字节）
 *   2. package.json files 含 templates/**（实地"打包漏文件"教训防线）
 *   3. 资产内容锚点预置（docx/pptx/xlsx 解包后含 {{key}} 占位标记——zip 级字节搜索，不碰 COM）
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== D 段 starter 资产断言 ===');
  const tdir = path.join(ROOT, 'templates');
  const three = ['weekly-report.docx', 'briefing.pptx', 'ledger.xlsx'];
  for (const f of three) {
    const p = path.join(tdir, f);
    const exists = fs.existsSync(p) && fs.statSync(p).size > 0;
    ok(`资产 ${f} 在仓且 >0B`, exists, exists ? fs.statSync(p).size + 'B' : '缺失');
  }
  // files 字段
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  ok('package.json files 含 templates/**', (pkg.files || []).some((f) => String(f).includes('templates')));
  // 锚点预置内容（zip 内文件名/内容字节搜索：OOXML zip 的 document.xml 是 deflate——直接搜压缩流不可靠，
  // 改搜 zip 本地文件头里的锚点副产物不可行。退而求其次：用 PowerShell Expand-Archive 解包临时目录查 {{key}}
  const tmp = fs.mkdtempSync(path.join(require('os').tmpdir(), 'starter-check-'));
  try {
    const zipChecks = [
      ['weekly-report.docx', 'word/document.xml', ['{{week_summary}}', '{{task_1}}', '{{next_plan}}']],
      ['briefing.pptx', 'ppt/slides/slide1.xml', ['{{brief_title}}', '{{presenter}}']],
      ['ledger.xlsx', 'xl/sharedStrings.xml', ['{{ledger_title}}', '{{ledger_item_1}}']],
    ];
    for (const [asset, member, anchors] of zipChecks) {
      const dest = path.join(tmp, asset.replace(/\./g, '_'));
      fs.mkdirSync(dest, { recursive: true });
      try {
        // Expand-Archive 要求 .zip 后缀——OOXML 资产先复制为 zip 再解包
        const zipCopy = path.join(tmp, asset + '.zip');
        fs.copyFileSync(path.join(tdir, asset), zipCopy);
        execFileSync('powershell', ['-NoProfile', '-Command',
          `Expand-Archive -LiteralPath '${zipCopy}' -DestinationPath '${dest}' -Force`],
          { timeout: 60000, stdio: 'ignore' });
        const memberPath = path.join(dest, ...member.split('/'));
        if (fs.existsSync(memberPath)) {
          const content = fs.readFileSync(memberPath, 'utf8');
          ok(`${asset} 锚点预置（${anchors.join('/')}）`, anchors.every((a) => content.includes(a)));
        } else {
          ok(`${asset} 成员 ${member} 存在`, false, '解包后未见成员');
        }
      } catch (e) {
        ok(`${asset} 解包`, false, (e.message || '').slice(0, 80));
      }
    }
  } finally {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* ignore */ }
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
