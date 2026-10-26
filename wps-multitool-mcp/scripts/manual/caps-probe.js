/**
 * 阶段 6 第一步：能力探测（真机）
 *
 * 目的：在实现 P0/P1 新能力之前，先把候选 COM 接口在本机 WPS 上各试一次，
 * 产出"支持 / 不支持"表。前面对接过三次接口不支持（InsertCaption / CaptionLabels /
 * ListTemplates.Add），不先探测就会写完才发现跑不通。
 *
 * 造一个带表格的文档，否则 table.* 探针会因 Count=0 静默跳过而误报 supported。
 *
 * ⚠️ 全程在临时副本上操作，finally 里关闭并删除，不碰用户文件。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-caps';
process.env.WPS_BRIDGE_IDLE_MS = '30000';
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '120000';

const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
const { WpsAppType } = require(path.join(ROOT, 'dist', 'types', 'wps.js'));
const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

const TMP = path.join(os.tmpdir(), `wps_caps_${process.pid}.txt`);
const call = (a, p = {}) => wpsClient.executeMethod(a, p, WpsAppType.WRITER);

// 按优先级分组的期望（只是标注，不是断言 —— 支持与否以真机结果为准）
const GROUPS = {
  'P1 样式 CUD': ['styles.list', 'styles.add', 'styles.delete', 'styles.modifyFont', 'styles.modifyParagraph', 'styles.byIndex'],
  'P1 交叉引用': ['crossRef.items', 'crossRef.fieldAdd'],
  'P1 索引': ['indexes.count', 'indexes.markEntry', 'indexes.add'],
  'P0 修订': ['revisions.count', 'revisions.item', 'revisions.acceptAll', 'revisions.rejectAll'],
  'P0 表格单元格级': ['table.cell', 'table.rowsAdd', 'table.colsAdd', 'table.cellMerge', 'table.cellSplit', 'table.cellBorders', 'table.style', 'table.rowsDelete'],
  'P1 页码/分节': ['pageNumbers.add', 'pageSetup.differentFirst', 'pageSetup.oddEven'],
  'P2 其他': ['watermark', 'contentControl', 'shape.addTextbox', 'docProtect', 'find.wildcards', 'comments.count', 'comments.add', 'docCompare', 'mailMerge', 'exportPdf'],
  '对照锚点（应当都不支持）': ['anchor.insertCaption', 'anchor.captionLabels', 'anchor.listTemplatesAdd'],
};

async function main() {
  fs.writeFileSync(TMP, '能力探测用临时文档。\n第二行。\n', 'utf8');
  console.log(`临时文档：${TMP}\n`);

  const open = await call('openDocument', { path: TMP });
  if (!open.success) throw new Error('打开文档失败: ' + open.error);
  await new Promise((r) => setTimeout(r, 3500));

  // 插一个表格，让 table.* 探针有意义
  const ti = await call('insertTable', { rows: 3, cols: 3 });
  console.log('插入测试表格:', ti.success ? 'OK' : `失败(${ti.error})`);

  const probe = await call('__probeCaps', {});
  if (!probe.success) throw new Error('探测失败: ' + probe.error);
  const d = probe.data;

  console.log(`\n探测文档：${d.docName}   共 ${d.total} 项\n`);
  for (const [grp, names] of Object.entries(GROUPS)) {
    console.log(`--- ${grp} ---`);
    for (const n of names) {
      const c = d.caps[n];
      if (!c) { console.log(`  ?  ${n}  （探针缺失）`); continue; }
      const mark = c.status === 'supported' ? '✅' : '❌';
      const err = c.error ? '  ' + c.error.slice(0, 78) : '';
      console.log(`  ${mark} ${n.padEnd(28)}${err}`);
    }
    console.log('');
  }

  // 有效自检：三个已知不支持项必须都报 unsupported，否则说明探针被 try 吞掉了
  const anchors = ['knownBad.insertCaption', 'knownBad.captionLabels', 'knownBad.listTemplatesAdd'];
  const anchorOk = anchors.every((n) => d.caps[n] && d.caps[n].status === 'unsupported');
  console.log(`对照组自检（三个已知不支持项都应报 unsupported）: ${anchorOk ? '✅ 探针有效' : '❌ 探针可能失效'}`);

  console.log(`\n汇总：支持 ${d.supported.length} / ${d.total}`);
  console.log('\n可实施（据此决定先做哪些）：');
  const impl = [];
  for (const [grp, names] of Object.entries(GROUPS)) {
    if (grp.startsWith('对照')) continue;
    const okNames = names.filter((n) => d.caps[n]?.status === 'supported');
    if (okNames.length) {
      impl.push(`${grp}: ${okNames.join(', ')}`);
      console.log(`  ${grp} → ${okNames.join(', ')}`);
    }
  }
  console.log(`\n（共 ${impl.length} 组可做）`);
}

main()
  .catch((e) => {
    console.error('探测异常:', e && e.message ? e.message : e);
    process.exitCode = 1;
  })
  .finally(async () => {
    try {
      await wpsClient.executeMethod('closeDocument', { name: path.basename(TMP), save: false }, undefined);
      console.log('已关闭测试文档');
    } catch { /* 忽略 */ }
    try { if (fs.existsSync(TMP)) fs.unlinkSync(TMP); console.log('已删除临时文件'); } catch { /* 忽略 */ }
    psDaemon.dispose();
    setTimeout(() => process.exit(process.exitCode || 0), 800);
  });