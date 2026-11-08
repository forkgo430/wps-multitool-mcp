/**
 * lint 落地（提质小卡 2026-11-08）
 *
 * 判据：npm run lint 真跑退出码 0；eslint 以【精确版本】锁在 devDependencies；
 * 最小 config 在位（.eslintrc.json，忽略 dist/生成物/manual）。
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
  console.log('=== lint 落地 ===');
  // 1) 真跑
  let exitCode = 0, out = '';
  try { out = require('child_process').execSync('npm run lint', { encoding: 'utf8', timeout: 180000, cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe'] }); }
  catch (e) { exitCode = e.status || 1; out = e.stdout || e.message; }
  ok('npm run lint 退出码 0', exitCode === 0, (out.match(/✖?\s*\d+ problem/g) || ['0 problems']).join(','));

  // 2) 精确锁版本
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  const dd = pkg.devDependencies || {};
  ok('eslint 精确锁版本（无 ^/~）', /^\d+\.\d+\.\d+$/.test(dd.eslint || ''), dd.eslint || '缺失');
  ok('@typescript-eslint/parser 精确锁版本', /^\d+\.\d+\.\d+$/.test(dd['@typescript-eslint/parser'] || ''), dd['@typescript-eslint/parser'] || '缺失');
  ok('@typescript-eslint/eslint-plugin 精确锁版本', /^\d+\.\d+\.\d+$/.test(dd['@typescript-eslint/eslint-plugin'] || ''), dd['@typescript-eslint/eslint-plugin'] || '缺失');
  ok('lint script 在位', (pkg.scripts.lint || '').includes('eslint'));

  // 3) 最小 config
  const cfgPath = path.join(ROOT, '.eslintrc.json');
  ok('.eslintrc.json 在位', fs.existsSync(cfgPath));
  if (fs.existsSync(cfgPath)) {
    const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
    ok('config 忽略 dist 与生成物', (cfg.ignorePatterns || []).some((p) => p.includes('dist')) && (cfg.ignorePatterns || []).some((p) => p.includes('generated')));
    ok('config 核心规则在位（no-unused-vars/prefer-const/no-var）',
      cfg.rules['@typescript-eslint/no-unused-vars'] && cfg.rules['prefer-const'] && cfg.rules['no-var']);
  }

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
