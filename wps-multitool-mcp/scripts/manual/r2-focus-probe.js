const { execFileSync } = require('child_process');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const PS = path.join(ROOT, 'scripts', 'wps-com.ps1');
function ps(action, params) {
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', PS, '-Action', action, '-Params', JSON.stringify(params || {})];
  try { return JSON.parse(execFileSync('powershell', args, { encoding: 'utf8', timeout: 90000, windowsHide: true }).trim()); }
  catch (e) { return { success: false, error: String(e.message).slice(0, 220) }; }
}
console.log('=== 场景 1：只有 Word（Word 是前台）===');
console.log('  createDocument:', JSON.stringify(ps('createDocument', {})));
console.log('  insertPageBreak(section):', JSON.stringify(ps('insertPageBreak', { type: 'section' })));

console.log('\n=== 场景 2：Excel 在前台，Word 文稿已存在 ===');
console.log('  createWorkbook:', JSON.stringify(ps('createWorkbook', {})));
console.log('  insertPageBreak(section):', JSON.stringify(ps('insertPageBreak', { type: 'section' })));

console.log('\n=== 场景 3：把 Excel 关掉，再试 ===');
for (const w of (((ps('getOpenWorkbooks', {}).data) || {}).workbooks || []))
  ps('closeWorkbook', { name: w.name, saveChanges: false });
console.log('  insertPageBreak(section):', JSON.stringify(ps('insertPageBreak', { type: 'section' })));

console.log('\n=== 收尾 ===');
for (const d of (((ps('getOpenDocuments', {}).data) || {}).documents || []))
  console.log('  关闭', d.name, JSON.stringify(ps('closeDocument', { name: d.name, saveChanges: false })));
for (const w of (((ps('getOpenWorkbooks', {}).data) || {}).workbooks || []))
  console.log('  关闭', w.name, JSON.stringify(ps('closeWorkbook', { name: w.name, saveChanges: false })));
