/**
 * 快照 / 恢复的真机验证（需要 WPS 文字在跑）。
 *
 * 为什么必须真机跑
 * --------------
 * 快照的核心不是"复制文件"（那部分单元测试能覆盖），而是【和正在编辑的文档怎么相处】：
 * 文档开着的时候覆盖文件到底会怎样、未保存的文档有没有路径、关闭后再覆盖能不能成。
 * 这些只有真机答得出来。
 *
 * 全程在临时目录里造文档，结束时关掉并删干净，不碰用户任何文档。
 * 无 WPS 时按既有约定"跳过不判失败"，但会明说"这批接口在真机上未验证"。
 *
 * 用法：node scripts/manual/snapshot-verify.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

// 用独立管道 + 短空闲：bridge 是常驻的，它手里的 PS 子进程拿着【旧版】wps-com.ps1。
// 共享管道会测到上一份脚本，改完动作库后必须换管道才能测到新的（与 p01-smoke 同一招）。
// 这比"去 kill 掉用户的桥"温和：只多起一个临时桥，用完自己退出。
// 每次跑用全新管道名（固定名会连到上一轮还活着的桥，桥里是旧版 wps-com.ps1）
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-snapshot-verify-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '30000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { SNAPSHOT_ROOT, readOpLog } = require(path.join(ROOT, 'dist', 'tools', 'common', 'snapshot.js'));

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

const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';

(async () => {
  console.log('\n=== A. 注册与注解（不需要 WPS）===');
  const TOOLS = ['wps_common_snapshot_create', 'wps_common_snapshot_list', 'wps_common_snapshot_restore'];
  ok('三个快照工具都已注册', TOOLS.every((n) => toolRegistry.hasTool(n)),
    TOOLS.filter((n) => !toolRegistry.hasTool(n)).join(',') || '无');

  const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));
  ok('create 标为非只读（它写磁盘）', getToolAnnotations('wps_common_snapshot_create')?.readOnlyHint === false);
  ok('restore 标为破坏性（overwrite 会盖原文件）', getToolAnnotations('wps_common_snapshot_restore')?.destructiveHint === true);
  ok('list 标为只读', getToolAnnotations('wps_common_snapshot_list')?.readOnlyHint === true);

  console.log('\n=== B. 真机验证（需要 WPS 文字可访问）===');
  const tmpDir = path.join(os.tmpdir(), `wps-snapshot-verify-${Date.now()}`);
  const docPath = path.join(tmpDir, 'snapshot-verify.docx');
  fs.mkdirSync(tmpDir, { recursive: true });
  let createdName = null;
  let snapRecord = null;

  const cleanupSnapshotFiles = () => {
    for (const r of readOpLog()) {
      if (r.snapshotPath && r.originalPath && path.resolve(r.originalPath) === path.resolve(docPath)) {
        try { fs.unlinkSync(r.snapshotPath); } catch { /* 已删 */ }
      }
    }
  };

  try {
    // 建一份临时文档并落盘，否则没有 FullName 可选
    const created = await call('wps_word_create_document', {});
    if (!created.success) throw new Error(`建文档失败：${txt(created)}`);
    createdName = (txt(created).match(/「([^」]+)」/) || [])[1] || null;

    await call('wps_word_insert_text', { text: 'snapshot-verify v1', position: 'end' });
    const saved = await call('wps_common_save_as', { filePath: docPath });
    ok('临时文档已另存到磁盘', saved.success && fs.existsSync(docPath), txt(saved).slice(0, 80));
    // 另存为之后文档名就变成文件名了 —— 后面关闭/定位都按这个名字来
    createdName = path.basename(docPath);

    // --- 建快照 ---
    const snap = await call('wps_common_snapshot_create', { label: 'verify' });
    ok('建快照成功', snap.success, txt(snap).slice(0, 80));
    snapRecord = readOpLog().filter((r) => r.kind === 'snapshot' && r.label === 'verify').pop();
    ok('快照文件真的落在磁盘上', !!snapRecord && fs.existsSync(snapRecord.snapshotPath));
    const v1 = fs.readFileSync(docPath);
    ok('快照内容与磁盘版本一致', !!snapRecord && Buffer.compare(fs.readFileSync(snapRecord.snapshotPath), v1) === 0);

    const list = await call('wps_common_snapshot_list', { originalPath: docPath });
    ok('snapshot_list 能列到刚建的那份', list.success && txt(list).includes('snapshot-verify.docx'));

    // --- 改内容并落盘，制造"改坏了"的场景 ---
    const ins = await call('wps_word_insert_text', { text: 'BROKEN EDIT', position: 'end' });
    ok('插入破坏性改动', ins.success, txt(ins).slice(0, 60));
    const sv = await call('wps_common_save', {});
    ok('保存改动', sv.success, txt(sv).slice(0, 60));
    const v2 = fs.readFileSync(docPath);
    ok('磁盘版本已经变了（证明后面恢复不是空转）', Buffer.compare(v1, v2) !== 0,
      `v1=${v1.length}B v2=${v2.length}B`);

    // --- 默认恢复：写副本，不碰原文件 ---
    const restoreCopy = await call('wps_common_snapshot_restore', { path: snapRecord.snapshotPath });
    ok('默认恢复到副本成功', restoreCopy.success, txt(restoreCopy).split('\n')[0]);
    const copyPath = (txt(restoreCopy).match(/：\n(.+\.restored-[^\n]+)/) || [])[1];
    ok('副本内容 = 快照内容（不是当前的坏版本）',
      !!copyPath && fs.existsSync(copyPath) && Buffer.compare(fs.readFileSync(copyPath), v1) === 0,
      copyPath || '没解析出副本路径');
    ok('原文件未被默认恢复改动', Buffer.compare(fs.readFileSync(docPath), v2) === 0);
    if (copyPath && fs.existsSync(copyPath)) fs.unlinkSync(copyPath);

    // --- 原地覆盖：文档开着时必须拒绝 ---
    const openRefuse = await call('wps_common_snapshot_restore', {
      path: snapRecord.snapshotPath, overwrite: true, confirm: true,
    });
    ok('文档开着时拒绝原地覆盖', openRefuse.success === false && /先关文档/.test(txt(openRefuse)));

    // --- 缺 confirm 也必须拒绝 ---
    const noConfirm = await call('wps_common_snapshot_restore', { path: snapRecord.snapshotPath, overwrite: true });
    ok('overwrite 少了 confirm 被拒', noConfirm.success === false && /confirm/.test(txt(noConfirm)));

    // --- 关掉文档后覆盖 ---
    const closed = await call('wps_word_close_document', { name: createdName || undefined, save: false });
    ok('关掉临时文档', closed.success, txt(closed).slice(0, 60));
    const restoreOver = await call('wps_common_snapshot_restore', {
      path: snapRecord.snapshotPath, overwrite: true, confirm: true,
    });
    ok('关文档后原地覆盖成功', restoreOver.success, txt(restoreOver).split('\n')[0]);
    ok('原文件已回到 v1', fs.existsSync(docPath) && Buffer.compare(fs.readFileSync(docPath), v1) === 0);
    const pre = readOpLog().filter(
      (r) => r.label === 'pre-restore' && path.resolve(String(r.originalPath)) === path.resolve(docPath)
    );
    ok(
      '覆盖前留了 pre-restore 备份（且放在快照根下，会被保留策略统计）',
      pre.length >= 1 && String(pre[pre.length - 1].snapshotPath).startsWith(SNAPSHOT_ROOT),
      `${pre.length} 份`
    );

    // --- 保留策略：超过 20 份要自动清理 ---
    const { createSnapshot } = require(path.join(ROOT, 'dist', 'tools', 'common', 'snapshot.js'));
    const reopened = await call('wps_word_open_document', { filePath: docPath });
    if (reopened.success) {
      for (let i = 0; i < 21; i++) await createSnapshot(`prune-${i}`);
      const mine = readOpLog().filter(
        (r) => r.kind === 'snapshot' && path.resolve(r.originalPath) === path.resolve(docPath)
          && fs.existsSync(r.snapshotPath)
      );
      const labels = mine.map((r) => `${r.label || 'verify'}@${r.ts.slice(11, 23)}`).sort();
      ok('每个文档最多留 20 份快照', mine.length === 20, `实际 ${mine.length} 份：${labels.join(' ')}`);
      await call('wps_word_close_document', { name: createdName || undefined, save: false });
    } else {
      ok('重新打开文档（保留策略用例前置）', false, txt(reopened).slice(0, 80));
    }

    // --- 未保存过的文档：必须确定性报错，不许糊弄 ---
    const blank = await call('wps_word_create_document', {});
    if (blank.success) {
      const noPath = await call('wps_common_snapshot_create', {});
      ok('未保存文档拒绝建快照并给出下一步', noPath.success === false && /另存为/.test(txt(noPath)),
        txt(noPath).slice(0, 120));
      await call('wps_word_close_document', { save: false });
    } else {
      ok('新建空白文档（未保存用例前置）', false, txt(blank).slice(0, 60));
    }
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log(`  ⚠ 真机探测未完成：${msg}`);
    console.log('    （需要 WPS 文字可访问。这一项本轮不判失败，但也不算数 ——');
    console.log('     意味着快照系列在真机上【未验证】。）');
  } finally {
    try {
      if (createdName) await call('wps_word_close_document', { name: createdName, saveChanges: false });
    } catch { /* 已经关了 */ }
    cleanupSnapshotFiles();
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { /* Windows 偶发占用 */ }
    console.log(`\n快照根目录：${SNAPSHOT_ROOT}`);
    console.log(`汇总：${pass} 通过 / ${fail} 失败`);
    process.exit(fail === 0 ? 0 : 1);
  }
})();
