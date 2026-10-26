#!/usr/bin/env node
/**
 * Input: 无（全程 mock COM，不碰 WPS、不起桥）
 * Output: 单工具自动快照的四条断言 —— 破坏性 1 份 / 批量 3 步仍 1 份 / 只读 0 份 / 关开关 0 份
 * Pos: 第 4 轮"单工具写入也自动快照"的验收脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么必须能离开真机验
 * ------------------
 * 这层保险的价值全在"它到底建了几份"上：建少了等于没保险，建多了会把用户磁盘塞满、
 * 还会把保留策略（每文档 20 份）里真正有用的那几份挤掉。份数这种东西靠肉眼看不出来，
 * 只能数。而数它不需要真 COM —— 快照的落盘路径是纯 Node 逻辑，把 COM 出口换掉就行。
 *
 * 怎么 mock：只替换 wpsClient.invokeAction（executeMethod 内部就走它），
 *   - getActiveDocumentState 返回一份"已存盘的假文档"
 *   - 其余动作一律返回 success，让 handler 能跑到头
 * 于是快照会真的写进临时 HOME 下，数目录里的文件就知道建了几份。
 *
 * 用法：node scripts/manual/auto-snapshot-verify.js
 */

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');

// ---- 1. 先把 HOME 指到临时目录：快照/oplog 全部落在那里，绝不碰真实用户目录 --------
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'wps-snap-verify-'));
process.env.HOME = TMP;
process.env.USERPROFILE = TMP;
delete process.env.WPS_NO_AUTO_SNAPSHOT;

const SNAPSHOT_DIR = path.join(TMP, '.wps-multitool-mcp', 'snapshots');

// 假文档：createSnapshot 要求 hasPath=true 且文件真的存在
const FAKE_DOC = path.join(TMP, 'probe-workbook.xlsx');
fs.writeFileSync(FAKE_DOC, Buffer.from('fake workbook bytes'));

const { wpsClient } = require(path.join(ROOT, 'dist/client/wps-client.js'));
const { toolRegistry, ToolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist/server/tool-annotations.js'));
const { snapshotRiskOf } = require(path.join(ROOT, 'dist/server/auto-snapshot.js'));

// ---- 2. 装 mock：只换 invokeAction -----------------------------------------
let docStateOk = true;
const calls = [];
wpsClient.invokeAction = async (action) => {
  calls.push(action);
  if (action === 'getActiveDocumentState') {
    if (!docStateOk) return { success: false, error: 'mock: 拿不到活动文档' };
    return {
      success: true,
      data: {
        app: 'excel',
        name: 'probe-workbook.xlsx',
        fullName: FAKE_DOC,
        saved: true,
        readOnly: false,
        hasPath: true,
      },
    };
  }
  return { success: true, data: { message: 'mocked' } };
};

// ---- 3. 走真实注册路径 ------------------------------------------------------
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

/** 数快照目录里现在有几份 */
const countSnaps = () => {
  try {
    return fs.readdirSync(SNAPSHOT_DIR).length;
  } catch {
    return 0;
  }
};

const call = (name, args = {}) => toolRegistry.callTool(ToolRegistry.createRequest(name, args));

(async () => {
  console.log('\n=== 0. 前提：挑中的三个工具注解符合预期 ===');
  const DESTRUCTIVE = 'wps_excel_delete_rows';
  const READONLY = 'wps_excel_get_cell_value';
  const dAnn = getToolAnnotations(DESTRUCTIVE);
  const rAnn = getToolAnnotations(READONLY);
  ok(`${DESTRUCTIVE} 被判为需要快照`, snapshotRiskOf(DESTRUCTIVE) !== '', snapshotRiskOf(DESTRUCTIVE));
  ok(`${READONLY} 被判为不需要快照`, snapshotRiskOf(READONLY) === '', JSON.stringify(rAnn));
  ok(`${DESTRUCTIVE} 注解 destructiveHint=true`, dAnn && dAnn.destructiveHint === true, JSON.stringify(dAnn));

  console.log('\n=== 1. 单工具（破坏性）→ 正好 1 份 ===');
  const before1 = countSnaps();
  const r1 = await call(DESTRUCTIVE, { startRow: 1, count: 1 });
  const after1 = countSnaps();
  ok('调用返回 success', r1.success === true, String(r1.error || ''));
  ok('快照 +1', after1 - before1 === 1, `${before1} → ${after1}`);

  console.log('\n=== 2. 批量 3 步（每步都破坏性）→ 仍只 1 份 ===');
  const before2 = countSnaps();
  const r2 = await call('wps_common_batch', {
    steps: [
      { tool: DESTRUCTIVE, params: { startRow: 1, count: 1 } },
      { tool: DESTRUCTIVE, params: { startRow: 2, count: 1 } },
      { tool: DESTRUCTIVE, params: { startRow: 3, count: 1 } },
    ],
  });
  const after2 = countSnaps();
  ok('批量返回 success', r2.success === true, String(r2.error || ''));
  ok('整批只建 1 份（不是 1+3=4 份）', after2 - before2 === 1, `${before2} → ${after2}`);

  console.log('\n=== 3. 只读工具 → 0 份 ===');
  const before3 = countSnaps();
  const r3 = await call(READONLY, { sheet: 'Sheet1', row: 1, col: 1 });
  const after3 = countSnaps();
  ok('调用返回 success', r3.success === true, String(r3.error || ''));
  ok('快照 +0', after3 - before3 === 0, `${before3} → ${after3}`);

  console.log('\n=== 4. WPS_NO_AUTO_SNAPSHOT=1 → 破坏性工具也 0 份 ===');
  process.env.WPS_NO_AUTO_SNAPSHOT = '1';
  const before4 = countSnaps();
  await call(DESTRUCTIVE, { startRow: 5, count: 1 });
  const after4 = countSnaps();
  delete process.env.WPS_NO_AUTO_SNAPSHOT;
  ok('开关生效：快照 +0', after4 - before4 === 0, `${before4} → ${after4}`);

  console.log('\n=== 5. 建不了快照时不阻断，但必须进 warnings ===');
  docStateOk = false;
  const r5 = await call(DESTRUCTIVE, { startRow: 9, count: 1 });
  docStateOk = true;
  const warned = (r5.warnings || []).some((w) => /没有回滚保险/.test(w.message));
  ok('执行没被拦住（success=true）', r5.success === true, String(r5.error || ''));
  ok('warnings 里明说这次没有回滚保险', warned, JSON.stringify((r5.warnings || []).map((w) => w.step)));

  console.log('\n=== 6. 快照摘要 ===');
  const names = (() => {
    try {
      return fs.readdirSync(SNAPSHOT_DIR);
    } catch {
      return [];
    }
  })();
  console.log(`  目录 ${SNAPSHOT_DIR}`);
  names.slice(0, 5).forEach((n) => console.log(`    ${n}`));
  ok('快照文件名带原文档名（可追溯到哪个文件）', names.some((n) => n.includes('probe-workbook')), names[0] || '(空)');
  ok('每份快照都进了 oplog（可恢复）', (() => {
    try {
      const log = fs.readFileSync(path.join(TMP, '.wps-multitool-mcp', 'oplog', 'ops.jsonl'), 'utf8');
      return (log.match(/"kind":"snapshot"/g) || []).length >= 2;
    } catch {
      return false;
    }
  })());

  console.log(`\n汇总：${pass}/${pass + fail} 通过`);
  // 清理临时目录（只清自己建的）
  try {
    fs.rmSync(TMP, { recursive: true, force: true });
  } catch {
    /* 清不掉不影响结论 */
  }
  process.exit(fail === 0 ? 0 : 1);
})();
