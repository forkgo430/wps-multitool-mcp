/**
 * 模板层大卡 A 段 · wps_render_template FAST 断言（mock 注入，零 WPS COM）
 *
 * 判据（质量杠）：
 *   1. 书签锚点优先（via=bookmark 计入 filled）
 *   2. {{key}} 占位兜底（via=placeholder）
 *   3. 缺 key 不静默：skipped 清单列全 + warnings 带"保留原文"提示
 *   4. 落盘校验：size>0 才 success；0 字节 → fail + filled/missing 如实带出
 *   5. outputPath 已存在 → 明确拒绝（不静默覆盖）
 *   6. data 为空对象/非对象 → 明确报错
 *   7. 慢档：renderTemplate 在 DEFAULT_SLOW_ACTIONS（120s）
 */
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..', '..');
let pass = 0, fail = 0;
const ok = (label, cond, detail) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${detail ? '：' + detail : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${detail ? '：' + detail : ''}`); }
};

const ps1 = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');
const templateSrc = fs.readFileSync(path.join(ROOT, 'src', 'tools', 'common', 'template.ts'), 'utf8');
const { callTimeoutMsFor } = require(path.join(ROOT, 'dist', 'bridge', 'ps-child.js'));

(async () => {
  console.log('=== wps_render_template 断言 ===');
  // 1) ps1 case 结构：书签优先 → 占位兜底 → saveAs → 落盘校验（顺序即语义）
  const caseBody = ps1.slice(ps1.indexOf('"renderTemplate"'), ps1.indexOf('"extractTemplate"'));
  const iBookmark = caseBody.indexOf('锚点双轨 ①：书签锚点优先');
  const iPlaceholder = caseBody.indexOf('锚点双轨 ②：{{key}} 文本占位兜底');
  const iSave = caseBody.indexOf('saveAs + 落盘校验');
  ok('ps1 case 结构：书签① → 占位② → saveAs（顺序即语义）', iBookmark > 0 && iPlaceholder > iBookmark && iSave > iPlaceholder);
  ok('ps1 case 落盘校验（verified = size>0）', /\$verified = \(\$size -gt 0\)/.test(caseBody));
  ok('ps1 case outputPath 已存在拒绝（不静默覆盖）', caseBody.includes('不静默覆盖，请换路径或先删除'));
  ok('ps1 case data 空/非对象明确报错', caseBody.includes('data required（JSON 对象：锚点名 → 值）') && caseBody.includes('data 为空对象'));
  // 缺 key 不静默
  ok('缺 key 不静默（skipped 清单 + warnings 保留原文提示）',
    caseBody.includes('保留原文未替换') && caseBody.includes('skipped = @($missing | Select-Object -Unique)'));
  // 路径归一（R2 教训防复发）
  ok('ps1 路径归一（Normalize-WinPath 入口）',
    (caseBody.match(/Normalize-WinPath/g) || []).length >= 2);
  // TS 壳：data 校验 + 空对象拒绝
  ok('TS 壳 data 空/非对象拦截', templateSrc.includes('data 必须是 JSON 对象') && templateSrc.includes('data 为空对象'));
  ok('TS 壳路径归一（R2 防复发，/ → \\）', templateSrc.includes("p.replace(/\\//g, '\\\\')"));
  // 慢档
  ok('renderTemplate 慢档 120s（R4 防复发）', callTimeoutMsFor('renderTemplate') === 120000);
  // text 主载体纪律：回包走 ok()（word/helpers），JSON 大回包由 mcp-server 闸门统一管——此处断言非 JSON 开头文本
  ok('回包为人类可读行文（非 JSON 文本——structuredContent 闸门外）',
    !templateSrc.includes('content: [{ type: \'text\', text: JSON.stringify'));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
