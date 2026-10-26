/**
 * 一次性诊断：找出"TS 传了参数、PS 却根本不读"的工具（静默丢弃 = 假成功）。
 *
 * 起因：验证者发现 wps_ppt_set_table_row_style 传嵌套 style，PS 读扁平 $p.backgroundColor，
 * 结果"返回成功但什么都没做"。这类 bug 最危险 —— 不报错，用户以为改完了。
 *
 * 第 3 轮起，本脚本只做【打印】：判定规则与门禁都放在 param-shape-core.js 里，
 * 由 param-shape-verify.js（进 verify-all 快测档）负责"红不红"。
 * 两处各写一套规则必然分裂，而分裂的结局一定是门禁被当成误报 bypass 掉 —— 所以规则只有一份。
 *
 * 看候选被分在哪一类、用什么理由接受的：
 *   node scripts/manual/param-shape-verify.js
 *
 * 已知局限（报出来的不一定都是 bug）：
 *   - 动态拼参数的调用点抓不到（不报 = 漏报，不是安全）
 *   - 入口归一化（改名/摊平/值转换）由扫描器按表认定，看不到转换本身
 *
 * 用法：node scripts/manual/param-shape-audit.js
 */
const { scanCandidates } = require('./param-shape-core');

const { rows, caseCount, aliasCount, flattenCount, helperCount } = scanCandidates();

console.log(`扫描器：case 体 ${caseCount} 个（同名只算第一个）、助手 ${helperCount} 个、` +
  `入口改名 ${aliasCount} 条、入口摊平 ${flattenCount} 条`);
console.log(`候选（TS 传了、PS 不读）共 ${rows.length} 个调用点：\n`);
for (const r of rows) console.log(`  ${r.act.padEnd(24)} ${r.missing.padEnd(28)} ${r.file}`);
console.log('\n下一步：逐个人工确认 —— 真丢参数的进 P1 修，PS 侧另有解析的写进 param-shape-verify.js 的白名单。');