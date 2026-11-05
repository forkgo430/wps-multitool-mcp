#!/usr/bin/env node
/**
 * Input: 无（全程 mock COM，不碰 WPS、不起桥）
 * Output: 协作安全网的断言组 —— 首次写入不提醒 / 人工编辑后提醒+建议重读 /
 *         无新编辑不重复提醒 / word 侧修订模式建议只出现一次 / 探测失败静默
 * Pos: 2026-10-29 协作安全网（决策者实现，所有者批准的三件协作增强）的验收脚本。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 怎么 mock：与 auto-snapshot-verify 同一套——替换 wpsClient.invokeAction，
 * getActiveDocumentState 回一份带可控 editProbe 的假文档；指纹用 editProbe 变量
 * 手动推进（P1→P2 模拟"用户在两次 AI 操作之间改了文档"）。
 *
 * 用法：node scripts/manual/collab-guard-verify.js
 */
'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-collab-verify-'));
process.env.HOME = TMP;
process.env.USERPROFILE = TMP;
delete process.env.WPS_NO_AUTO_SNAPSHOT;

const FAKE = path.join(TMP, 'probe-workbook.xlsx');
fs.writeFileSync(FAKE, Buffer.from('fake workbook bytes'));

const ROOT = path.resolve(__dirname, '..', '..');
const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const guard = require(path.join(ROOT, 'dist/server/collab-guard.js'));

let editProbe = 'P1';
let docName = 'probe-workbook.xlsx';
let docFullName = FAKE;
wpsClient.invokeAction = async (action) => {
  if (action === 'getActiveDocumentState') {
    return {
      success: true,
      data: {
        app: 'excel', name: docName, fullName: docFullName,
        saved: true, readOnly: false, hasPath: true, editProbe,
      },
    };
  }
  return { success: true, data: { message: 'mocked' } };
};

new WpsMcpServer();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};
const call = (name, args = {}) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));
const warningsOf = (r) => (r.warnings || []).map((w) => w.message || '');

(async () => {
  guard.resetCollabGuardForTest();

  console.log('\n=== 1. 首次写入：没有"上次"可比 → 不提醒 ===');
  editProbe = 'P1';
  const r1 = await call('wps_excel_delete_rows', { startRow: 1, count: 1 });
  ok('调用 success', r1.success === true, String(r1.error || ''));
  ok('无协作提醒', !warningsOf(r1).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(r1)));

  console.log('\n=== 2. 人工编辑后写入：提醒"可能是人工编辑"+ 建议重读 ===');
  editProbe = 'P2';
  const r2 = await call('wps_excel_delete_rows', { startRow: 2, count: 1 });
  ok('提醒"可能是人工编辑"', warningsOf(r2).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(r2)));
  ok('提醒里建议重新读取', warningsOf(r2).some((w) => w.includes('重新读取')));
  ok('提醒带上下文（上次存档→本次，误报可自查）', warningsOf(r2).some((w) => w.includes('上次存档') && w.includes('本次')));
  ok('提醒 severity=info（体验修复轮 B：不挂"部分成功"头）', (r2.warnings || []).every((w) => w.severity === 'info'), JSON.stringify((r2.warnings || []).map((w) => w.severity)));
  ok('写入没有被阻断（success=true）', r2.success === true);

  console.log('\n=== 3. 紧接着再写（无新的人工编辑）：不重复提醒 ===');
  const r3 = await call('wps_excel_delete_rows', { startRow: 3, count: 1 });
  ok('无协作提醒（AI 自己的写入已重新存档）', !warningsOf(r3).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(r3)));

  console.log('\n=== 4. word 侧：修订模式建议只出现一次 ===');
  guard.resetCollabGuardForTest();
  editProbe = 'W1';
  const w0 = await guard.collabGuardCheckBefore('word');
  ok('word 首次检查无存档 → 不提醒', w0.length === 0, JSON.stringify(w0));
  await guard.collabGuardNoteAfter('word');
  editProbe = 'W2';
  const w1 = await guard.collabGuardCheckBefore('word');
  ok('指纹变化 → 提醒', w1.some((w) => w.includes('人工编辑')), JSON.stringify(w1));
  ok('带修订模式建议', w1.some((w) => w.includes('修订模式')), JSON.stringify(w1));
  await guard.collabGuardNoteAfter('word');
  editProbe = 'W3';
  const w2 = await guard.collabGuardCheckBefore('word');
  ok('再次变化 → 仍有提醒', w2.some((w) => w.includes('人工编辑')), JSON.stringify(w2));
  ok('修订模式建议不重复（每会话一次）', !w2.some((w) => w.includes('修订模式')), JSON.stringify(w2));

  console.log('\n=== 5. 探测失败（无 WPS 形态的异常）→ 静默降级，不外抛 ===');
  guard.resetCollabGuardForTest();
  const origInvoke = wpsClient.invokeAction;
  wpsClient.invokeAction = async () => { throw new Error('no wps (CI linux 模拟)'); };
  let silenced = true;
  try {
    const w5 = await guard.collabGuardCheckBefore('word');
    silenced = Array.isArray(w5) && w5.length === 0;
  } catch (e) {
    silenced = false;
    console.log('  抛出：' + e.message);
  }
  wpsClient.invokeAction = origInvoke;
  ok('异常被吞掉、返回空提醒', silenced);

  console.log('\n=== 6. 换文档不误报（体验修复轮 B：save_as/切换文档器场景）===');
  guard.resetCollabGuardForTest();
  editProbe = 'D1'; docName = '报告A.xlsx'; docFullName = path.join(TMP, '报告A.xlsx');
  await call('wps_excel_delete_rows', { startRow: 1, count: 1 }); // 文档A：首次写入建档
  editProbe = 'D9'; docName = '报告B.xlsx'; docFullName = path.join(TMP, '报告B.xlsx');
  const c2 = await call('wps_excel_delete_rows', { startRow: 1, count: 1 }); // 另存/切到文档B
  ok('切到新文档 → 不误报"人工编辑"（静默重建档）',
    !warningsOf(c2).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(c2)));
  editProbe = 'D10';
  const c3 = await call('wps_excel_delete_rows', { startRow: 2, count: 1 }); // B 内真有人工改动
  ok('新文档内真有人工改动 → 照常提醒（键是文档级，B 有了自己的基线）',
    warningsOf(c3).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(c3)));
  editProbe = 'D1'; docName = '报告A.xlsx'; docFullName = path.join(TMP, '报告A.xlsx');
  const c4 = await call('wps_excel_delete_rows', { startRow: 1, count: 1 }); // 切回A，期间无人动
  ok('切回旧文档且指纹未变 → 不误报（A 的基线还在）',
    !warningsOf(c4).some((w) => w.includes('人工编辑')), JSON.stringify(warningsOf(c4)));

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch { /* 清不掉不影响结论 */ }
  process.exit(fail === 0 ? 0 : 1);
})();
