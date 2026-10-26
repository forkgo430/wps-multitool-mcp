/**
 * Input: 无
 * Output: 对新增 Word 长文档动作的端到端实测结果
 * Pos: ② 的手工验证脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 流程：新建文档 → 依次调用新增动作 → 关闭文档【不保存】。
 * 会真的把 WPS 文字拉起来（正常行为：Get-WpsWord 附不上就 New-Object），
 * 结束时用 closeDocument(saveChanges=false) 静默关掉，不留未保存文档、不弹保存对话框。
 *
 * 用法：node scripts/manual/word-academic-smoke.js
 */

const path = require('path');

const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-academic-smoke';
process.env.WPS_BRIDGE_IDLE_MS = '300000';
process.env.WPS_DAEMON_TIMEOUT = '60000';

const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

const results = [];
async function step(name, action, params) {
  const t = Date.now();
  try {
    const r = await psDaemon.call(action, params || {}, 'word');
    const ms = Date.now() - t;
    const ok = r.success === true;
    results.push({ name, ok, ms, detail: ok ? JSON.stringify(r.data) : String(r.error) });
    console.log(`  ${ok ? '✅' : '❌'} ${name.padEnd(22)} ${String(ms).padStart(6)}ms  ${ok ? String(JSON.stringify(r.data)).slice(0, 110) : String(r.error).slice(0, 110)}`);
    return r;
  } catch (e) {
    const ms = Date.now() - t;
    results.push({ name, ok: false, ms, detail: e.message });
    console.log(`  ❌ ${name.padEnd(22)} ${String(ms).padStart(6)}ms  ${e.message.slice(0, 110)}`);
    return { success: false, error: e.message };
  }
}

(async () => {
  // 顺手核准注册的工具数
  const { academicTools } = require(path.join(ROOT, 'dist', 'tools', 'word', 'academic.js'));
  console.log(`academic 工具注册数 = ${academicTools.length}`);
  console.log(`  ${academicTools.map((t) => t.definition.name).join('\n  ')}`);
  console.log('');

  console.log('=== ② Word 长文档动作端到端 ===');
  await step('createDocument', 'createDocument', {});
  await step('setDefaultFont', 'setDefaultFont', { eastAsia: '宋体', ascii: 'Times New Roman', size: 12 });
  await step('addHeading L1', 'addHeading', { text: '第一章 绪论', level: 1 });
  await step('addHeading L2', 'addHeading', { text: '1.1 研究背景', level: 2 });
  await step('insertParagraph', 'insertParagraph', { text: '这是正文第一段。' });
  await step('addList number', 'addList', { items: ['第一项', '第二项', '第三项'], type: 'number' });
  await step('addCaption', 'addCaption', { text: '系统架构图', label: '图' });
  await step('addFootnote', 'addFootnote', { text: '这是一条脚注。', atEnd: true });
  await step('setColumns 2', 'setColumns', { count: 2 });
  const outline = await step('getOutline', 'getOutline', {});
  await step('setHeadingNumbering', 'setHeadingNumbering', { levels: 3 });
  await step('insertTableOfFigures', 'insertTableOfFigures', { label: '图' });
  await step('updateFields', 'updateFields', {});
  const stats = await step('getDocumentStats', 'getDocumentStats', {});
  await step('closeDocument', 'closeDocument', { saveChanges: false });

  if (outline && outline.success) {
    console.log(`\n  大纲实测：${JSON.stringify(outline.data.headings)}`);
  }
  if (stats && stats.success) {
    console.log(`  统计实测：页=${stats.data.pages} 字=${stats.data.words} 段=${stats.data.paragraphs} 表=${stats.data.tables} 脚注=${stats.data.footnotes}`);
  }

  const okCount = results.filter((r) => r.ok).length;
  console.log(`\n汇总：${okCount}/${results.length} 成功`);
  const failed = results.filter((r) => !r.ok);
  if (failed.length) {
    console.log('失败项：');
    for (const f of failed) console.log(`  - ${f.name}: ${f.detail}`);
  }
  psDaemon.dispose();
  process.exit(failed.length ? 1 : 0);
})().catch((e) => {
  console.error('冒烟测试异常:', e.message);
  process.exit(1);
});
