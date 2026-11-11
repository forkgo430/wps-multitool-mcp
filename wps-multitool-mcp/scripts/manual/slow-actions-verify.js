/**
 * R4 · 超时分级表断言（提质合并卡）
 *
 * 判据：五动作（insertPptImage/insertExcelImage/insertImage/save/saveAs）进慢档 120s；
 * 快动作不放宽（ping/hasTool 类仍 30s）。
 */
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));

let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

(async () => {
  console.log('=== R4 超时分级表 ===');
  const slowFive = ['insertPptImage', 'insertExcelImage', 'insertImage', 'save', 'saveAs'];
  for (const a of slowFive) {
    ok(`${a} → 120s 慢档`, callTimeoutMsFor(a) === 120000, `${callTimeoutMsFor(a)}ms`);
  }
  ok('ping 仍 30s 快档（不放宽）', callTimeoutMsFor('ping') === 30000, `${callTimeoutMsFor('ping')}ms`);
  ok('hasTool 仍 30s 快档（不放宽）', callTimeoutMsFor('hasTool') === 30000, `${callTimeoutMsFor('hasTool')}ms`);
  // 兜底是既有设计："宁可多等，不可误杀"——未传 action 按慢档 120s（非 R4 改动）
  ok('未传 action 兜底慢档 120s（宁多等不误杀，既有语义）', callTimeoutMsFor(undefined) === 120000, `${callTimeoutMsFor(undefined)}ms`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
