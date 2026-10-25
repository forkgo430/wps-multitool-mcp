/**
 * Word 侧能力探测（第 5 期新增接口的真机验证）
 * ------------------------------------------------
 * 探什么：__probeWordCaps 在【临时文档】上逐个试 WPS 12.1 到底认不认这些接口，
 * 铺探测素材（文字 + 2x2 表 + 书签）、探完原样关闭（不保存）——
 * 不碰用户任何文档，也不留临时文档垃圾。
 *
 * 与 excel-caps-probe 的分工：那个探 Excel 接口；这个探 Word 深水区新接口
 * （统计/批注/书签/表格/超链接/表转文本）。__probeCaps（Word 老探针）探的是
 * 【用户当前文档】上的既有能力，两者互补，别合并。
 *
 * 判据要说清楚：supported 只代表"调用没抛异常"，不代表语义完全对。
 *
 * 用法：node scripts/manual/word-caps-probe.js
 */

const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');

const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { getToolAnnotations } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

// 第 5 期新增的 10 个工具
const NEW_TOOLS = [
  'wps_word_get_document_stats',
  'wps_word_get_comments',
  'wps_word_delete_comment',
  'wps_word_get_bookmarks',
  'wps_word_get_tables',
  'wps_word_get_table_data',
  'wps_word_create_document',
  'wps_word_close_document',
  'wps_word_insert_hyperlink',
  'wps_word_convert_table_to_text',
];

// 新建文档必须单独探：它有病史（"零等待返回空壳文档"），
// 判据是 Documents.Count 真的 +1、且新文档能立刻操作（健康检查通过）。
const MUST_VERIFY = ['wps_word_create_document'];

(async () => {
  let pass = 0;
  let fail = 0;
  const ok = (label, cond, extra) => {
    if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
    else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
  };

  console.log('\n=== A. 10 个新工具的注册与元数据（不需要 WPS）===');
  ok('新工具全部注册', NEW_TOOLS.every((n) => toolRegistry.hasTool(n)),
    `缺：${NEW_TOOLS.filter((n) => !toolRegistry.hasTool(n)).join(',') || '无'}`);
  const noAnn = NEW_TOOLS.filter((n) => !getToolAnnotations(n));
  ok('新工具都有注解', noAnn.length === 0, noAnn.join(','));

  // 只读/写入的 Annotation 必须符合预期：读侧 5 个只读，写侧 5 个非只读
  const READ_ONLY = NEW_TOOLS.filter((n) => n.startsWith('wps_word_get_'));
  const WRITERS = NEW_TOOLS.filter((n) => !n.startsWith('wps_word_get_'));
  const badRead = READ_ONLY.filter((n) => getToolAnnotations(n)?.readOnlyHint !== true);
  ok('5 个读侧工具标为只读', badRead.length === 0, badRead.join(','));
  const badWrite = WRITERS.filter((n) => getToolAnnotations(n)?.readOnlyHint !== false);
  ok('5 个写侧工具标为非只读', badWrite.length === 0, badWrite.join(','));

  console.log('\n=== B. 真机能力探测（需要 WPS 文字在跑）===');
  const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
  let probe;
  try {
    const r = await wpsClient.executeMethod('__probeWordCaps', {}, 'word');
    probe = r.data;
    if (!r.success) throw new Error(r.error);
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    console.log(`  ⚠ 探测未执行：${msg}`);
    console.log('    （需要 WPS 文字可访问。这一项本轮不判失败，但也不算数 ——');
    console.log('     意味着这批新接口在真机上【未验证】。）');
    console.log(`\n汇总：${pass} 通过 / ${fail} 失败（真机探测跳过）`);
    process.exit(fail === 0 ? 0 : 1);
  }

  const supported = probe.supported || [];
  const unsupported = probe.unsupported || [];
  console.log(`  探测项 ${probe.total} 个：支持 ${supported.length} / 不支持 ${unsupported.length}`);
  if (unsupported.length) console.log(`  不支持：${unsupported.join(', ')}`);
  ok('临时文档删干净了（探测不留垃圾）', (probe.leaked || []).length === 0,
    (probe.leaked || []).join('；'));

  // 核心接口必须支持（不支持 = 对应工具在本机不可用，要在文档里写明）
  for (const k of ['stats.words', 'stats.pages', 'comments.count', 'bookmarks.count', 'tables.count', 'tables.cellText']) {
    ok(`核心接口 ${k} 受支持`, !unsupported.includes(k));
  }

  console.log('\n=== C. create_document 病史专项（Count 真变 + 能立刻操作）===');
  let n0 = 0;
  let n1 = 0;
  let created = null;
  try {
    const { wpsClient: wc } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));
    const before = await wc.executeMethod('getOpenDocuments', {}, 'word');
    n0 = (before.data && before.data.count) || 0;
    const cr = await wc.executeMethod('createDocument', {}, 'word');
    if (!cr.success) throw new Error(cr.error);
    created = cr.data;
    const after = await wc.executeMethod('getOpenDocuments', {}, 'word');
    n1 = (after.data && after.data.count) || 0;
    ok('createDocument 返回成功', true, `name=${created && created.name}`);
    ok(`Documents.Count 真的变了（${n0} → ${n1}）`, n1 === n0 + 1, `${n0} → ${n1}`);
    ok('createDocument 报告了文档名', !!(created && created.name), created && created.name);
    // 能立刻操作：对新文档直接读统计（不等用户干预）
    const st = await wc.executeMethod('getDocumentStats', {}, 'word');
    ok('建完能立刻操作（对新文档直接读统计成功）', st.success && st.data && st.data.words >= 0,
      st.success ? `words=${st.data && st.data.words}` : st.error);
  } catch (e) {
    ok('create_document 专项', false, e && e.message ? e.message : String(e));
  } finally {
    // 清理：把刚建的文档关掉（不保存）。新建文档 FullName 为空，
    // 走 closeDocument 的 saveChanges=false 分支，不会弹框。
    if (created && created.name) {
      try {
        await require(path.join(ROOT, 'dist', 'client', 'wps-client.js')).wpsClient.executeMethod(
          'closeDocument', { name: created.name, saveChanges: false }, 'word'
        );
        const afterClose = await require(path.join(ROOT, 'dist', 'client', 'wps-client.js')).wpsClient.executeMethod(
          'getOpenDocuments', {}, 'word'
        );
        const n2 = (afterClose.data && afterClose.data.count) || 0;
        ok('专项文档已清理（Count 回到 ' + n0 + '）', n2 === n0, `${n2}`);
      } catch (e) {
        console.log(`  ⚠ 专项文档清理失败: ${e && e.message ? e.message : e}`);
      }
    }
  }

  // 每个写侧工具都在探测里有对应项（防止 PS 加了动作但忘了探）
  const probeKeys = Object.keys(probe.caps || {});
  ok('写侧动作都有探测项', ['hyperlinks.add', 'tables.convertToText', 'comments.addDelete'].every((k) => probeKeys.includes(k)),
    probeKeys.join(', '));

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})();
