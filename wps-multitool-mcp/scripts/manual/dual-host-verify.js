/**
 * 模板层大卡 · 双宿主真机标记断言（live 探针 VERDICT=PASS 的固化锚）
 *
 * 判据：
 *   1. 双宿主渲染产物在案（WPS 宿主 off + Office 宿主 on 各 rendered-*.docx >0B）
 *   2. 台账 officeHost.live 两工具 note 含双宿主证据字样
 *   3. 提取产物在案（extracted-*.docx）
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const TMP = 'C:/Users/金东权/AppData/Local/Temp/template-dual';
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== 双宿主真机标记 ===');
  // 1) 渲染产物在案
  for (const mode of ['off', 'on']) {
    const p = path.join(TMP, `rendered-${mode}.docx`);
    ok(`渲染产物 rendered-${mode}.docx 在案（>0B）`, fs.existsSync(p) && fs.statSync(p).size > 0,
      fs.existsSync(p) ? fs.statSync(p).size + 'B' : '缺失');
    const e = path.join(TMP, `extracted-${mode}.docx`);
    ok(`提取产物 extracted-${mode}.docx 在案`, fs.existsSync(e) && fs.statSync(e).size > 0);
  }
  // 2) 台账双宿主证据
  const ledger = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'manual', 'tool-verification.json'), 'utf8'));
  const note = (ledger.officeHost && ledger.officeHost.live['wps_render_template']) || '';
  ok('台账 wps_render_template note 含双宿主证据（off/on 各落盘）',
    note.includes('off') && note.includes('on') && note.includes('PASS'), note.slice(0, 80));
  const note2 = (ledger.officeHost && ledger.officeHost.live['wps_extract_template']) || '';
  ok('台账 wps_extract_template note 含双宿主证据', note2.includes('双宿主'), note2.slice(0, 80));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
