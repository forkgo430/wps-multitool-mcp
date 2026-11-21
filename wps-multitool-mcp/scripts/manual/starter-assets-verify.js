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
  // 纯 Node 解包：中央目录解析 + zlib.inflateRawSync，不走 PowerShell）
  // zip 成员读取器（EOCD → 中央目录 → 本地头定位 → inflateRaw），零依赖、跨平台确定性
  function zipReadMember(buf, member) {
    const eocd = buf.lastIndexOf(Buffer.from('PK\x05\x06'));
    if (eocd < 0) return null;
    const count = buf.readUInt16LE(eocd + 10);
    let p = buf.readUInt32LE(eocd + 16);
    for (let i = 0; i < count; i++) {
      if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) return null;
      const method = buf.readUInt16LE(p + 10);
      const csize = buf.readUInt32LE(p + 20);
      const nameLen = buf.readUInt16LE(p + 28);
      const extraLen = buf.readUInt16LE(p + 30);
      const commentLen = buf.readUInt16LE(p + 32);
      const lho = buf.readUInt32LE(p + 42);
      const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
      if (name === member) {
        const nlen2 = buf.readUInt16LE(lho + 26);
        const elen2 = buf.readUInt16LE(lho + 28);
        const data = buf.slice(lho + 30 + nlen2 + elen2, lho + 30 + nlen2 + elen2 + csize);
        return method === 0 ? data : require('zlib').inflateRawSync(data);
      }
      p += 46 + nameLen + extraLen + commentLen;
    }
    return null;
  }
  const tmp = fs.mkdtempSync(path.join(require('os').tmpdir(), 'starter-check-'));
  try {
    const zipChecks = [
      ['weekly-report.docx', 'word/document.xml', ['{{week_summary}}', '{{task_1}}', '{{next_plan}}']],
      ['briefing.pptx', 'ppt/slides/slide1.xml', ['{{brief_title}}', '{{presenter}}']],
      ['ledger.xlsx', 'xl/sharedStrings.xml', ['{{ledger_title}}', '{{ledger_item_1}}']],
    ];
    for (const [asset, member, anchors] of zipChecks) {
      try {
        // 纯 Node 解包（2026-11-21）：Expand-Archive 在 CI runner 上解出成员缺失（假红），
        // 改用 zip 中央目录解析 + zlib.inflateRawSync——零依赖、跨平台确定性
        const buf = fs.readFileSync(path.join(tdir, asset));
        const memberData = zipReadMember(buf, member);
        if (memberData === null) {
          ok(`${asset} 成员 ${member} 存在`, false, 'zip 内未见成员');
        } else {
          const content = memberData.toString('utf8');
          ok(`${asset} 锚点预置（${anchors.join('/')}）`, anchors.every((a) => content.includes(a)));
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
