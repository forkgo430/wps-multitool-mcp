const { execFileSync } = require('child_process');
const path = require('path');
const ROOT = path.resolve(__dirname, '..', '..');
const PS = path.join(ROOT, 'scripts', 'wps-com.ps1');
function ps(action, params) {
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', PS, '-Action', action, '-Params', JSON.stringify(params || {})];
  try { return JSON.parse(execFileSync('powershell', args, { encoding: 'utf8', timeout: 90000, windowsHide: true }).trim()); }
  catch (e) { return { success: false, error: String(e.message).slice(0, 160) }; }
}
for (const [act, key, flag] of [['getOpenWorkbooks', 'workbooks', 'saveChanges'], ['getOpenDocuments', 'documents', 'saveChanges']]) {
  const d = (ps(act, {}).data) || {};
  const list = d[key] || [];
  console.log(`${act}: ${list.length} 个`);
  for (const it of list) {
    const nm = it.name || it.Name;
    console.log(`  关闭「${nm}」 (${flag}=false)`);
    const r = ps(act.startsWith('getOpenWork') ? 'closeWorkbook' : 'closeDocument', { name: nm, [flag]: false });
    console.log(`    → ${r.success ? '已关闭' : '失败: ' + String(r.error).slice(0, 110)}`);
  }
}
setTimeout(() => {
  console.log('复查：表格 ' + (((ps('getOpenWorkbooks', {}).data) || {}).count) +
              ' / 文稿 ' + (((ps('getOpenDocuments', {}).data) || {}).count));
}, 1200);
