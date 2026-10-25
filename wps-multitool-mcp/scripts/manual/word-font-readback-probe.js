/**
 * T1 真机读回探针（第 1 期 · 第 2 章 · 第 1 轮）
 * ------------------------------------------------------------------
 * 验的是两个"从来没生效过"的 Word 工具：
 *   wps_word_set_text_color   原先调 'setTextColor'   （ps1 里零命中）
 *   wps_word_set_line_spacing 原先调 'setLineSpacing' （ps1 里零命中）
 *
 * 为什么判据必须是"读回值"
 * ------------------------
 * 这两个工具的历史问题正是"回 success 但什么都没发生"。所以判据不能是"回包说成功"，
 * 必须是【读回来的值真的变了】。
 *
 * 读回这一侧为什么必须独立于 MCP 通道
 * ----------------------------------
 * wps-com.ps1 里【没有】读 Font.Color / ParagraphFormat.LineSpacing 的动作
 * （逐个动作核过：getDocumentText / getDocumentParagraphs / getSelection 都不带格式）。
 * 所以读回只能自己挂 WPS Word 的 COM 读：
 *   [Marshal]::GetActiveObject('Kwps.Application') 挂到同一个正在跑的 WPS 实例，
 *   读 ActiveDocument.Content 的 Font.Color 与 ParagraphFormat.LineSpacing。
 * 这一侧不走 MCP 通道 —— 否则就是"自己写自己读"，证明不了任何事。
 *
 * 顺带证明"旧动作名确实是坏的"：直接裸调 'setTextColor' / 'setLineSpacing'，
 * 看它们是不是真的落进分发器的 Unknown action 分支。
 *
 * 文档纪律：临时文稿 save:false；只关自己创建的那一个（用打开列表差集认领）。
 *
 * 用法：node scripts/manual/word-font-readback-probe.js
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';

const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { wpsClient } = require(path.join(ROOT, 'dist', 'client', 'wps-client.js'));

const s = new WpsMcpServer();
s.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const call = (n, a = {}) => toolRegistry.getTool(n).handler(a);
const txt = (r) => String((r && r.content && r.content[0] && r.content[0].text) || '').replace(/\s+/g, ' ');

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

/* ------------------------------------------------------------------ */
/* 读回侧：独立挂 COM，只读，不改任何东西                              */
/* ------------------------------------------------------------------ */
const PS_READ = `
$ErrorActionPreference = 'Stop'
try { $w = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application') }
catch { try { $w = New-Object -ComObject 'Kwps.Application' } catch { $w = $null } }
if ($null -eq $w) { Write-Output '{"ok":false,"error":"no Kwps.Application"}'; exit 0 }
$doc = $null
try { $doc = $w.ActiveDocument } catch {}
if ($null -eq $doc) { Write-Output '{"ok":false,"error":"no active document"}'; exit 0 }
try {
  $r = $doc.Content
  $o = [ordered]@{
    ok          = $true
    name        = $doc.Name
    text        = ($r.Text -replace "[\\r\\n]", ' ')
    color       = $r.Font.Color
    lineSpacing = $r.ParagraphFormat.LineSpacing
    lineRule    = $r.ParagraphFormat.LineSpacingRule
    paraCount   = $doc.Paragraphs.Count
  }
  Write-Output ($o | ConvertTo-Json -Compress)
} catch {
  $msg = $_.Exception.Message -replace '"', ''
  Write-Output ('{"ok":false,"error":"' + $msg + '"}')
}
`;

function readState() {
  const out = execFileSync(
    'powershell',
    ['-NoProfile', '-NonInteractive', '-Command', PS_READ],
    { encoding: 'utf8', timeout: 60000 }
  );
  const line = String(out).trim().split(/\r?\n/).filter(Boolean).pop();
  if (!line) throw new Error('读回侧没有任何输出');
  return JSON.parse(line);
}

/* 选中范围语义：默认走 $word.Selection.Range，必须先有选区。
   选中这一步用直接 COM 做（它是【布置】，不是【断言】）。 */
const PS_SELECT = `
$ErrorActionPreference = 'Stop'
try { $w = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application') }
catch { $w = New-Object -ComObject 'Kwps.Application' }
$doc = $w.ActiveDocument
$doc.Content.Select()
Write-Output ('{"ok":true,"selected":"' + $w.Selection.Text.Length + '"}')
`;

function selectAll() {
  const out = execFileSync(
    'powershell',
    ['-NoProfile', '-NonInteractive', '-Command', PS_SELECT],
    { encoding: 'utf8', timeout: 60000 }
  );
  return String(out).trim().split(/\r?\n/).filter(Boolean).pop();
}

(async () => {
  console.log('\n=== 0. 先证明"旧动作名"确实是坏的（裸调 PS 层，不经工具）===');
  for (const act of ['setTextColor', 'setLineSpacing']) {
    const r = await wpsClient.executeMethod(act, act === 'setTextColor' ? { color: '#FF0000' } : { lineSpacing: 2 }, 'word');
    console.log(`  裸调 '${act}' → success=${r.success}  error=${r.error || '(无)'}`);
    ok(`'${act}' 落进 Unknown action 分支（证实工具从未生效）`,
      r.success === false && /Unknown action/i.test(String(r.error || '')),
      String(r.error || '').slice(0, 60));
  }

  console.log('\n=== 1. 认领一份自己的临时文稿（打开列表差集）===');
  const before = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const beforeNames = new Set(((before.data && before.data.documents) || []).map((d) => d.name));
  console.log(`  开工前打开的文稿：${beforeNames.size} 个`);

  const created = await call('wps_word_create_document', {});
  ok('wps_word_create_document 建出新文稿', created.success, txt(created).slice(0, 60));

  const after = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const afterDocs = (after.data && after.data.documents) || [];
  const mine = afterDocs.filter((d) => !beforeNames.has(d.name));
  ok('打开列表差集 = 恰好 1 个新文稿（只认领自己的）', mine.length === 1,
    `差集 ${mine.length} 个: ${mine.map((d) => d.name).join(', ')}`);
  const myName = mine.length ? mine[0].name : null;
  if (!myName) { console.log('  ✗ 认领失败，中止（不关任何文档）'); process.exit(1); }

  const inserted = await call('wps_word_insert_text', { text: 'R11读回测试文本' });
  ok('wps_word_insert_text 写入测试文本', inserted.success, txt(inserted).slice(0, 50));

  console.log('\n=== 2. 改之前的读回值（基线）===');
  const st0 = readState();
  console.log('  ', JSON.stringify(st0));
  ok('读回侧挂上了同一个文稿', st0.ok && st0.name === myName, `读到 "${st0.name}"，认领 "${myName}"`);
  ok('文本确实写进去了', String(st0.text || '').includes('R11读回测试文本'), `text=${JSON.stringify(st0.text)}`);

  console.log('\n=== 3. wps_word_set_text_color：颜色真的变了吗 ===');
  const r1 = await call('wps_word_set_text_color', { color: '#FF0000', range: 'all' });
  console.log('  工具回包：', txt(r1).slice(0, 90));
  ok('工具返回成功', r1.success, txt(r1).slice(0, 60));
  const st1 = readState();
  console.log('  ', JSON.stringify(st1));
  ok('读回的颜色真的变了（不是回 success 而已）', st1.color !== st0.color,
    `color ${st0.color} → ${st1.color}`);
  ok('读回的颜色就是设定的红色（BGR 255）', st1.color === 255, `期望 255，实际 ${st1.color}`);

  console.log('\n=== 4. wps_word_set_line_spacing：行距真的变了吗 ===');
  const r2 = await call('wps_word_set_line_spacing', { lineSpacing: 2, range: 'all' });
  console.log('  工具回包：', txt(r2).slice(0, 90));
  ok('工具返回成功', r2.success, txt(r2).slice(0, 60));
  const st2 = readState();
  console.log('  ', JSON.stringify(st2));
  ok('读回的行距真的变了（不是回 success 而已）', st2.lineSpacing !== st1.lineSpacing,
    `lineSpacing ${st1.lineSpacing} → ${st2.lineSpacing}`);
  ok('读回的行距就是 2 倍（PS 侧 lineSpacing×12 = 24 磅）', st2.lineSpacing === 24,
    `期望 24，实际 ${st2.lineSpacing}`);
  ok('行距规则变成多倍行距（LineSpacingRule=4）', st2.lineRule === 4, `实际 ${st2.lineRule}`);

  console.log('\n=== 5. 默认路径（不传 range → 走 $word.Selection.Range）也要真生效 ===');
  const sel = selectAll();
  console.log('  已选中字符数：', sel);
  const st3 = readState();
  const r3 = await call('wps_word_set_line_spacing', { lineSpacing: 1.5 });
  console.log('  工具回包：', txt(r3).slice(0, 90));
  ok('不传 range 时工具返回成功', r3.success, txt(r3).slice(0, 60));
  const st4 = readState();
  console.log('  ', JSON.stringify(st4));
  ok('默认（选中）路径下行距真的变了', st4.lineSpacing !== st3.lineSpacing,
    `lineSpacing ${st3.lineSpacing} → ${st4.lineSpacing}`);
  ok('默认（选中）路径下行距 = 1.5 倍 × 12 = 18 磅', st4.lineSpacing === 18,
    `期望 18，实际 ${st4.lineSpacing}`);

  selectAll();
  const st5 = readState();
  const r4 = await call('wps_word_set_text_color', { color: '#0000FF' });
  console.log('  工具回包：', txt(r4).slice(0, 90));
  ok('不传 range 时颜色工具返回成功', r4.success, txt(r4).slice(0, 60));
  const st6 = readState();
  console.log('  ', JSON.stringify(st6));
  ok('默认（选中）路径下颜色真的变了', st6.color !== st5.color,
    `color ${st5.color} → ${st6.color}`);
  ok('默认（选中）路径下颜色 = 蓝色（BGR 16711680）', st6.color === 16711680,
    `期望 16711680，实际 ${st6.color}`);

  console.log('\n=== 6. 收尾：只关自己创建的那一个，save:false ===');
  const closed = await call('wps_word_close_document', { name: myName, save: false });
  console.log('  ', txt(closed).slice(0, 90));
  const afterClose = await wpsClient.executeMethod('getOpenDocuments', {}, 'word');
  const left = ((afterClose.data && afterClose.data.documents) || []).map((d) => d.name);
  ok('自己的文稿已关闭', !left.includes(myName), `剩余 ${left.length} 个`);
  ok('没动别人的文稿', beforeNames.size === left.length || left.every((n) => beforeNames.has(n)),
    `开工前 ${beforeNames.size} 个 / 现在 ${left.length} 个`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('探针异常：', e && e.message ? e.message : e);
  process.exit(2);
});
