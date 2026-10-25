const { execFileSync } = require('child_process');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const PS = path.join(ROOT, 'scripts', 'wps-com.ps1');
function ps(action, params) {
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', PS, '-Action', action, '-Params', JSON.stringify(params || {})];
  try { return JSON.parse(execFileSync('powershell', args, { encoding: 'utf8', timeout: 90000, windowsHide: true }).trim()); }
  catch (e) { return { success: false, error: String(e.message).slice(0, 300) }; }
}
function readback() {
  const script = [
    "$app = [Runtime.Interopservices.Marshal]::GetActiveObject('Kwps.Application')",
    "$d = $app.ActiveDocument",
    "$sel = 'null'",
    "try { $sel = [string]$app.Selection.Text } catch { $sel = 'ERR:' + $_.Exception.Message }",
    "$paras = 0; try { $paras = [int]$d.Paragraphs.Count } catch {}",
    "$secs = 0; try { $secs = [int]$d.Sections.Count } catch {}",
    "[PSCustomObject]@{ sel=$sel; paras=$paras; sections=$secs } | ConvertTo-Json -Compress",
  ].join('\n');
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script],
      { encoding: 'utf8', timeout: 60000, windowsHide: true });
    return JSON.parse(out.trim().split('\n').pop());
  } catch (e) { return { error: String(e.message).slice(0, 200) }; }
}
console.log('=== 建 Word 文稿 ===');
console.log(JSON.stringify(ps('createDocument', {})));
console.log('  读回:', JSON.stringify(readback()));

console.log('\n=== A. insertPageBreak 不传 type（默认 page）===');
console.log(JSON.stringify(ps('insertPageBreak', {})));
console.log('  读回:', JSON.stringify(readback()));

console.log('\n=== B. insertPageBreak type=section ===');
console.log(JSON.stringify(ps('insertPageBreak', { type: 'section' })));
console.log('  读回:', JSON.stringify(readback()));

console.log('\n=== C. insertPageBreak type=sectionContinuous ===');
console.log(JSON.stringify(ps('insertPageBreak', { type: 'sectionContinuous' })));
console.log('  读回:', JSON.stringify(readback()));

console.log('\n=== D. 先写点文字（让文档有内容/焦点）再插 ===');
console.log(JSON.stringify(ps('insertText', { text: 'R2 探针文字' })));
console.log('  读回:', JSON.stringify(readback()));
console.log(JSON.stringify(ps('insertPageBreak', { type: 'section' })));
console.log('  读回:', JSON.stringify(readback()));

console.log('\n=== 收尾 ===');
const dn = ((ps('getOpenDocuments', {}).data) || {}).documents || [];
for (const d of dn) { console.log('  关闭', d.name, JSON.stringify(ps('closeDocument', { name: d.name, saveChanges: false }))); }
