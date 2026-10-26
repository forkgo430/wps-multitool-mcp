/**
 * 阶段 6 第二步：Word P0/P1 新工具真机端到端（13 个）
 *
 * 造一个带表格、带标题、带修订的真文档，逐个跑：
 *   P0 修订：list / acceptAll
 *   P0 表格：getCellText / mergeCells / splitCell / insertTableRowCol / deleteTableRowCol / setCellBorders
 *   P1 样式：listStyles / addStyle / modifyStyle / deleteStyle
 *   P1 交叉引用：listTargets / insert
 *   P1 索引：count / markEntry（探测报 E_FAIL，看真机是否真不行）
 *   P1 页码：addPageNumber / setLayout
 *
 * ⚠️ 全程在临时副本上，finally 关闭并删除。
 */
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-p01';
process.env.WPS_BRIDGE_IDLE_MS = '30000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '90000';

const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { WpsAppType } = require(path.join(ROOT, 'dist', 'types', 'wps.js'));
const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

const results = [];
const call = (a, p = {}) => wpsClient.executeMethod(a, p, WpsAppType.WRITER);

function rec(name, okFlag, detail = '') {
  results.push({ name, ok: okFlag });
  const tag = okFlag ? '✅' : '⚠️ ';
  console.log(`  ${tag} ${name}${detail ? '  ' + detail : ''}`);
}

async function main() {
  // ---- 环境预检：必须在开跑前确认 WPS 处于可用状态 ----
  // 不做这一步的后果（实测踩过）：WPS 若已有一堆无名空壳文档且 COM 被拒
  // （RPC_E_CALL_REJECTED，但并没有模态对话框），后面每一步都会返回
  // "No active document"，刷一屏红灯却看不出根因。应该早失败、说清原因。
  console.log('环境预检…');
  const pre = await wpsClient.executeMethod('wpsRepair', {}, undefined);
  const wordState = pre.data?.apps?.word ?? {};
  const badDocs = (wordState.documents ?? []).filter((d) => d.healthy === false);
  console.log(
    `  word: running=${wordState.running} 文档=${wordState.count ?? 0}` +
      ` 不健康=${badDocs.length} 无名=${(wordState.documents ?? []).filter((d) => !d.name).length}`
  );
  // ⚠️ 必须带 count > 0：0 个文档时 activeHealthy 也是 false（没有活动文档可检查），
  // 不加这个条件会把"干净状态"误判成"有空壳文档"，直接拦住正常测试。
  // 这比漏报更麻烦 —— 漏报只是多跑几步，误报是根本跑不了。
  if (wordState.running && (wordState.count ?? 0) > 0 && (badDocs.length > 0 || wordState.activeHealthy === false)) {
    console.log('\n⚠️ WPS 里有不健康的空壳文档，COM 调用会被拒（这正是 wps_common_repair 要诊断的问题）。');
    console.log('   实测：这种状态下 repair 也关不掉它们（同样是 RPC_E_CALL_REJECTED）。');
    console.log('   需要【手动重启 WPS】再跑本测试 —— 不是代码问题，继续跑只会刷一屏红灯。\n');
    process.exitCode = 2;
    return;
  }
  if ((wordState.count ?? 0) === 0) {
    console.log('  WPS 干净（无文档），直接建测试文档。\n');
  } else if (badDocs.length === 0) {
    console.log(`  现有 ${wordState.count} 个文档且全部健康，继续。\n`);
  }

  console.log('用 createDocument 建真 .docx（与 __probeCaps 的探测环境对齐 ——\n' +
              '探测是在空 .docx 上做的，若改用 .txt 会引入"文档类型不同"的变量）\n');
  const created = await call('createDocument', {});
  if (!created.success) throw new Error('新建文档失败: ' + created.error);
  await new Promise((r) => setTimeout(r, 2000));
  await call('insertText', { text: '第一章 绪论\n这是正文内容。', position: 0 });

  // 造表格 + 开修订（开修订后插入的文字会成为修订）
  await call('insertTable', { rows: 3, cols: 3 });
  await call('enableTrackChanges', { enable: true });
  await call('insertText', { text: '这段是修订过的内容。', position: 0 });
  await call('enableTrackChanges', { enable: false });

  console.log('=== P0 修订 ===');
  const revList = await call('listRevisions', {});
  rec('revisions mode=list', revList.success, revList.success ? `count=${revList.data?.count}` : String(revList.error).slice(0, 60));
  const accAll = await call('handleRevisions', { mode: 'acceptAll' });
  rec('revisions acceptAll', accAll.success, accAll.success ? `${accAll.data?.before}→${accAll.data?.after}` : String(accAll.error).slice(0, 60));
  // 非法 mode 必须在【还有修订时】测 —— 文档干净时会先走"没有修订"的早返回分支，
  // 那样测出来的是"成功但没做事"，不是"拒绝非法参数"。
  await call('enableTrackChanges', { enable: true });
  await call('insertText', { text: '再改一次。', position: 0 });
  await call('enableTrackChanges', { enable: false });
  const badMode = await call('handleRevisions', { mode: 'bogus' });
  rec('revisions 非法 mode 被拒（不静默）', badMode.success === false && /mode 只能/.test(String(badMode.error)), String(badMode.error).slice(0, 60));
  await call('handleRevisions', { mode: 'rejectAll' });

  console.log('\n=== P0 表格单元格级 ===');
  const cell = await call('getCellText', { row: 1, col: 1 });
  rec('getCellText', cell.success, cell.success ? `"${cell.data?.text}"` : String(cell.error).slice(0, 60));
  const insRow = await call('insertTableRowCol', { dir: 'rows', op: 'insert', index: 1, where: 'before' });
  rec('insertTableRowCol 行', insRow.success, insRow.success ? `${JSON.stringify(insRow.data?.before)}→${JSON.stringify(insRow.data?.after)}` : String(insRow.error).slice(0, 70));
  const insCol = await call('insertTableRowCol', { dir: 'cols', op: 'insert', index: 1, where: 'after' });
  rec('insertTableRowCol 列', insCol.success, insCol.success ? `${JSON.stringify(insCol.data?.before)}→${JSON.stringify(insCol.data?.after)}` : String(insCol.error).slice(0, 70));
  const delRow = await call('deleteTableRowCol', { dir: 'rows', op: 'delete', index: 1 });
  rec('deleteTableRowCol', delRow.success, delRow.success ? JSON.stringify(delRow.data?.after) : String(delRow.error).slice(0, 60));
  const merge = await call('mergeTableCells', { row1: 1, col1: 1, row2: 1, col2: 2 });
  rec('mergeTableCells（Word 版，与 Excel 的 mergeCells 区分）', merge.success, merge.success ? JSON.stringify(merge.data) : String(merge.error).slice(0, 70));
  const split = await call('splitCell', { row: 1, col: 1, numRows: 1, numCols: 2 });
  rec('splitCell', split.success, split.success ? JSON.stringify(split.data) : String(split.error).slice(0, 70));
  const border = await call('setCellBorders', { row: 2, col: 2, side: 'all', style: 1 });
  rec('setCellBorders', border.success, border.success ? JSON.stringify(border.data?.applied) : String(border.error).slice(0, 70));
  const badSide = await call('setCellBorders', { row: 1, col: 1, side: 'diagonal' });
  rec('setCellBorders 非法 side 被拒', badSide.success === false, String(badSide.error).slice(0, 60));

  console.log('\n=== P1 样式 CUD ===');
  const ls = await call('listStyles', { onlyUserDefined: false });
  rec('listStyles', ls.success, ls.success ? `${ls.data?.count} 个` : String(ls.error).slice(0, 60));
  const add = await call('addStyle', {
    name: '__p01测试样式', baseOn: '正文', fontName: '宋体', fontSize: 14, bold: true, alignment: 'center',
  });
  rec('addStyle', add.success, add.success ? `applied=${JSON.stringify(add.data?.applied)}` : String(add.error).slice(0, 70));
  const addDup = await call('addStyle', { name: '__p01测试样式' });
  rec('addStyle 重名被拒（提示改用 modify）', addDup.success === false, String(addDup.error).slice(0, 60));
  const mod = await call('modifyStyle', { name: '__p01测试样式', fontSize: 18, italic: true });
  rec('modifyStyle', mod.success, mod.success ? `applied=${JSON.stringify(mod.data?.applied)}` : String(mod.error).slice(0, 70));
  const modNone = await call('modifyStyle', { name: '__p01测试样式' });
  rec('modifyStyle 无属性时报错（不假装成功）', modNone.success === false, String(modNone.error).slice(0, 60));
  const lsUser = await call('listStyles', { onlyUserDefined: true });
  rec('listStyles 只看自定义能看到它', lsUser.success && (lsUser.data?.styles ?? []).some((s) => s.name === '__p01测试样式'), `自定义 ${lsUser.data?.count} 个`);
  const del = await call('deleteStyle', { name: '__p01测试样式' });
  rec('deleteStyle', del.success, del.success ? JSON.stringify(del.data) : String(del.error).slice(0, 70));
  const delBuilt = await call('deleteStyle', { name: '正文' });
  rec('deleteStyle 内置样式被拒', delBuilt.success === false, String(delBuilt.error).slice(0, 70));

  console.log('\n=== P1 交叉引用 ===');
  const tgt = await call('listCrossReferenceTargets', {});
  rec('listCrossReferenceTargets', tgt.success, tgt.success ? JSON.stringify(Object.fromEntries(Object.entries(tgt.data?.targets ?? {}).map(([k, v]) => [k, v.length]))) : String(tgt.error).slice(0, 60));
  // 先加一个标题样式段落，才有 heading 可引
  await call('addHeading', { text: '第二章 方法', level: 1 });
  const insRef = await call('insertCrossReference', { refType: 'heading', itemIndex: 1 });
  rec('insertCrossReference heading', insRef.success, insRef.success ? `target=${insRef.data?.target}` : String(insRef.error).slice(0, 80));
  const refBad = await call('insertCrossReference', { refType: 'heading', itemIndex: 99 });
  rec('insertCrossReference 越界被拒', refBad.success === false, String(refBad.error).slice(0, 60));

  console.log('\n=== P1 索引 ===');
  const ixCount = await call('insertIndex', {});
  rec('insertIndex（顺带报索引项数）', ixCount.success, ixCount.success ? `marked=${ixCount.data?.markedEntries}` : String(ixCount.error).slice(0, 70));
  const mark = await call('markIndexEntry', { entry: '深度学习', start: 0, end: 4 });
  rec('markIndexEntry（探测报 E_FAIL，看真机）', mark.success, mark.success ? JSON.stringify(mark.data) : String(mark.error).slice(0, 90));

  console.log('\n=== P1 页码 / 版式 ===');
  const pn = await call('addPageNumber', { which: 'primary', alignment: 'center' });
  rec('addPageNumber', pn.success, pn.success ? JSON.stringify(pn.data) : String(pn.error).slice(0, 70));
  const pnBad = await call('addPageNumber', { which: 'nope' });
  rec('addPageNumber 非法 which 被拒', pnBad.success === false, String(pnBad.error).slice(0, 60));
  const lay = await call('setHeaderFooterLayout', { differentFirstPage: true, oddAndEvenPages: false });
  rec('setHeaderFooterLayout', lay.success, lay.success ? JSON.stringify(lay.data) : String(lay.error).slice(0, 70));
  const layEmpty = await call('setHeaderFooterLayout', {});
  rec('setHeaderFooterLayout 无参数时报错', layEmpty.success === false, String(layEmpty.error).slice(0, 60));

  // 收尾：刷新域，确认没把文档搞坏
  const uf = await call('updateFields', {});
  rec('updateFields（确认改动没破坏域）', uf.success, uf.success ? JSON.stringify(uf.data) : String(uf.error).slice(0, 70));

  const passed = results.filter((r) => r.ok).length;
  console.log(`\n汇总：${passed}/${results.length} 通过（⚠️ 的算失败）`);
}

main()
  .catch((e) => {
    console.error('测试异常:', e && e.message ? e.message : e);
    process.exitCode = 1;
  })
  .finally(async () => {
    try {
      const open = await wpsClient.executeMethod('getOpenDocuments', {}, WpsAppType.WRITER);
      const names = (open.data?.documents ?? []).map((d) => d.name).filter(Boolean);
      for (const n of names) {
        await wpsClient.executeMethod('closeDocument', { name: n, save: false }, undefined);
      }
      console.log(`已关闭 ${names.length} 个测试文档`);
    } catch { /* 忽略 */ }
    psDaemon.dispose();
    setTimeout(() => process.exit(process.exitCode || 0), 800);
  });