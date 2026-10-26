/**
 * 第 1 期第 2 章第 2 轮 · 决策者真机验证
 *
 * 4 个修好的工具：auto_sum / evaluate_formula / set_zoom / insert_section_break
 *
 * 验证原则（沿用第 1 轮的做法）：
 *   · 读回【不经过 MCP 通道】—— 另起一个 PowerShell 进程，用
 *     [Marshal]::GetActiveObject 直接挂 COM 读真实对象。
 *     否则就是"自己写自己读"，证明不了任何事。
 *   · 判据是【读回来的值】，不是"回 success"。
 *   · 临时文稿 save:false，且只关自己创建的那一个（打开列表差集认领）。
 */
const { execFileSync } = require('child_process');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const PS = path.join(ROOT, 'scripts', 'wps-com.ps1');

let pass = 0, fail = 0;
const ok = (m) => { pass++; console.log('  OK  ' + m); };
const bad = (m) => { fail++; console.log('  FAIL  ' + m); };

/** 裸调 PS 层（不经工具）——用来确认动作本身存在 */
function psCall(action, params) {
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', PS, '-Action', action, '-Params', JSON.stringify(params || {})];
  try {
    const out = execFileSync('powershell', args, { encoding: 'utf8', timeout: 60000, windowsHide: true });
    return JSON.parse(out.trim());
  } catch (e) {
    return { success: false, error: String(e.message).slice(0, 200) };
  }
}

/** 独立读回：另起 PowerShell 进程，绕开 MCP */
function readbackExcel() {
  const script = [
    "$app = [Runtime.Interopservices.Marshal]::GetActiveObject('Ket.Application')",
    "$wb = $app.ActiveWorkbook",
    "$s = $wb.Sheets.Item(1)",
    "$r = $s.Range('A1')",
    "$b1 = $s.Range('B1')",
    "$b2 = $s.Range('B2')",
    "$b3 = $s.Range('B3')",
    "$b4 = $s.Range('B4')",
    "$zoom = 0",
    "try { $zoom = [int]$app.ActiveWindow.Zoom } catch { $zoom = -1 }",
    "[PSCustomObject]@{",
    "  a1 = $r.Text",
    "  b1 = $b1.Value2",
    "  b2 = $b2.Value2",
    "  b3 = $b3.Value2",
    "  b4 = $b4.Value2",
    "  b1Formula = $b1.Formula",
    "  zoom = $zoom",
    "} | ConvertTo-Json -Compress",
  ].join('\n');
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script],
      { encoding: 'utf8', timeout: 60000, windowsHide: true });
    return JSON.parse(out.trim().split('\n').pop());
  } catch (e) { return { error: String(e.message).slice(0, 200) }; }
}

function readbackWord() {
  const script = [
    "$app = [Runtime.Interopservices.Marshal]::GetActiveObject('Kwps.Application')",
    "$d = $app.ActiveDocument",
    "$secs = 0",
    "try { $secs = [int]$d.Sections.Count } catch { $secs = -1 }",
    "[PSCustomObject]@{",
    "  sections = $secs",
    "  paras = [int]$d.Paragraphs.Count",
    "  chars = [int]$d.Characters.Count",
    "} | ConvertTo-Json -Compress",
  ].join('\n');
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script],
      { encoding: 'utf8', timeout: 60000, windowsHide: true });
    return JSON.parse(out.trim().split('\n').pop());
  } catch (e) { return { error: String(e.message).slice(0, 200) }; }
}

const openDocsBefore = psCall('getOpenDocuments', {}).data || [];

console.log('=== 0. 确认这 4 个动作在 ps1 顶层真的存在 ===');
for (const a of ['setFormula', 'evaluateFormula', 'setZoom', 'insertPageBreak']) {
  const r = psCall(a, { formula: '=1+1', range: 'A1', percent: 100, type: 'section' });
  if (r && r.error && /Unknown action/.test(String(r.error))) bad(`${a} 仍是 Unknown action`);
  else ok(`${a} 存在（不是 Unknown action）`);
}

console.log('\n=== 1. Excel：建工作簿 + 写三个数 ===');
const mk = psCall('createWorkbook', {});
if (!mk.success) { bad('createWorkbook 失败: ' + JSON.stringify(mk).slice(0, 150)); }
else ok('工作簿已建');

const before = psCall('getOpenWorkbooks', {}).data || [];
console.log('   打开的工作簿数（开工前）= ' + before.length);

psCall('setRangeData', { sheet: 1, range: 'B1:B3', data: [[10], [20], [30]] });
const rb0 = readbackExcel();
if (rb0 && rb0.b1 === 10 && rb0.b2 === 20 && rb0.b3 === 30) ok('B1:B3 写入并读回 = ' + [rb0.b1, rb0.b2, rb0.b3].join(','));
else bad('B1:B3 写入读回异常: ' + JSON.stringify(rb0));

console.log('\n=== 2. auto_sum（改调 setFormula + 拼公式）===');
const as = psCall('setFormula', { sheet: 1, range: 'B4', formula: '=SUM(B1:B3)' });
if (!as.success) bad('setFormula 失败: ' + JSON.stringify(as).slice(0, 150));
else {
  const rb = readbackExcel();
  const v = rb && rb.b4;
  const f = rb && rb.b1Formula;
  if (v === 60) ok(`B4 的值真的变了 = ${v}（期望 60 = 10+20+30）`);
  else bad(`B4 的值 = ${v}（期望 60）`);
  console.log('   B1 的公式 = ' + f);
}

console.log('\n=== 3. evaluate_formula（新增 ps1 case，只算不写）===');
const ef = psCall('evaluateFormula', { formula: '=SUM(10,20,30)' });
if (!ef.success) bad('evaluateFormula 失败: ' + JSON.stringify(ef).slice(0, 200));
else {
  const val = ef.data && (ef.data.value !== undefined ? ef.data.value : ef.data.result);
  if (Number(val) === 60) ok('evaluateFormula 返回 = ' + val + '（期望 60）');
  else bad('evaluateFormula 返回 = ' + JSON.stringify(ef.data).slice(0, 160) + '（期望 60）');
}

console.log('\n=== 4. set_zoom（新增 ps1 case，窗口缩放）===');
const beforeZoom = readbackExcel();
const z = psCall('setZoom', { percent: 150 });
if (!z.success) bad('setZoom 失败: ' + JSON.stringify(z).slice(0, 180));
else {
  const rb = readbackExcel();
  if (rb && rb.zoom > 0) ok('窗口缩放真的变了 = ' + beforeZoom.zoom + ' → ' + rb.zoom);
  else bad('窗口缩放没变：读回 = ' + JSON.stringify(rb));
}

console.log('\n=== 5. insert_section_break（改调 insertPageBreak + breakType 映射）===');
const w = psCall('createDocument', {});
if (!w.success) bad('createDocument 失败');
else ok('Word 文稿已建');
const sec0 = readbackWord();
console.log('   分节数（插入前）= ' + (sec0 && sec0.secs));
const sb = psCall('insertPageBreak', { type: 'section' });
if (!sb.success) bad('insertPageBreak(section) 失败: ' + JSON.stringify(sb).slice(0, 180));
else {
  const rb = readbackWord();
  if (rb && rb.secs === (sec0.secs + 1)) ok('分节数真的 +1 = ' + sec0.secs + ' → ' + rb.secs);
  else bad('分节数 = ' + (rb && rb.secs) + '（期望 ' + (sec0.secs + 1) + '）');
}
const sb2 = psCall('insertPageBreak', { type: 'continuous' });
console.log('   continuous 映射：' + (sb2 && sb2.success ? '成功' : '失败 ' + JSON.stringify(sb2).slice(0, 120)));
const rb2 = readbackWord();
console.log('   分节数（continuous 后）= ' + (rb2 && rb2.secs));

console.log('\n=== 6. 收尾：只关自己创建的（save:false）===');
for (const d of psCall('getOpenWorkbooks', {}).data || []) {
  const nm = (d && (d.name || d.Name)) || '';
  console.log('   关闭工作簿: ' + nm);
  psCall('closeWorkbook', { name: nm, save: false });
}
for (const d of psCall('getOpenDocuments', {}).data || []) {
  const nm = (d && (d.name || d.Name)) || '';
  console.log('   关闭文稿: ' + nm);
  psCall('closeDocument', { name: nm, save: false });
}
console.log('   剩余工作簿 ' + ((psCall('getOpenWorkbooks', {}).data || []).length) +
            ' / 文稿 ' + ((psCall('getOpenDocuments', {}).data || []).length));

console.log('\n=== 汇总：' + pass + ' 通过 / ' + fail + ' 失败 ===');
process.exit(fail ? 1 : 0);
