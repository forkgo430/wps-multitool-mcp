/**
 * 模板层大卡 C 段 · 配方追加断言（属地授权例外核验）
 *
 * 判据：
 *   1. task-packages.json 合法 JSON，packages = 19（17 + 追加 2）
 *   2. 追加的 resume_fill / weekly_report_template 结构合法（task/match/steps）
 *   3. 既有 17 包一字未动（对账键集合）
 *   4. SKILL.md 两轨分工段在位
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
  console.log('=== C 段配方追加断言 ===');
  const tp = JSON.parse(fs.readFileSync(path.join(ROOT, 'src', 'server', 'task-packages.json'), 'utf8'));
  ok('task-packages.json 合法 JSON', !!tp.packages && Array.isArray(tp.packages));
  ok('packages = 19（17 既有 + 追加 2）', tp.packages.length === 19, `count=${tp.packages.length}`);
  const tasks = tp.packages.map((p) => p.task);
  ok('resume_fill 已追加', tasks.includes('resume_fill'));
  ok('weekly_report_template 已追加', tasks.includes('weekly_report_template'));
  const rf = tp.packages.find((p) => p.task === 'resume_fill');
  ok('resume_fill 结构合法（execute=server + renderTemplate 步）',
    rf.execute === 'server' && rf.steps.some((s) => s.action === 'renderTemplate'));
  const wr = tp.packages.find((p) => p.task === 'weekly_report_template');
  ok('weekly_report_template 结构合法（extractTemplate + renderTemplate 双步）',
    wr.steps.some((s) => s.action === 'extractTemplate') && wr.steps.some((s) => s.action === 'renderTemplate'));
  // 既有 17 包一字未动（锚集合）
  const preExisting = ['weekly_report', 'official_document', 'contract_review', 'paper_format', 'data_entry',
    'demo_beautify', 'data_charts', 'format_unify', 'meeting_notes', 'project_tracker', 'compare_analysis',
    'template_fill', 'thesis_skeleton', 'literature_matrix', 'defense_deck', 'journal_format', 'group_meeting_report'];
  ok('既有 17 包全在（键集合对账）', preExisting.every((t) => tasks.includes(t)));

  // SKILL 两轨分工
  const skill = fs.readFileSync(path.join(ROOT, 'skills', 'wps-multitool', 'SKILL.md'), 'utf8');
  ok('SKILL 两轨分工段在位（数据模板 vs 配方）', skill.includes('模板两轨分工') && skill.includes('wps_render_template') && skill.includes('wps_extract_template'));
  ok('SKILL 默认档计数 42（common 组 +2）', skill.includes('默认只列出 42 个基础工具'));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => { console.error('异常：', e && e.stack ? e.stack : e); process.exit(2); });
