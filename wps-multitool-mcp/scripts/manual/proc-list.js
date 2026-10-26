/**
 * Input: 无（可用参数过滤关键字）
 * Output: 列出 node.exe / powershell.exe 的 PID 与命令行（用于核验桥与子进程）
 * Pos: 进程核查小工具。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 用法：node scripts/manual/proc-list.js [keyword]
 */

const { execFileSync } = require('child_process');

const kw = process.argv[2] || '';

function run(name) {
  const script =
    `Get-CimInstance Win32_Process -Filter "Name='${name}'" -EA SilentlyContinue |` +
    ` Select-Object ProcessId, ParentProcessId, CommandLine | ConvertTo-Json -Compress -Depth 3`;
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script], {
      encoding: 'utf8',
    }).trim();
    if (!out) return [];
    const parsed = JSON.parse(out);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    console.log(`[${name}] 查询失败: ${e.message.slice(0, 300)}`);
    return [];
  }
}

for (const name of ['node.exe', 'powershell.exe']) {
  const rows = run(name).filter((r) => !kw || String(r.CommandLine || '').includes(kw));
  console.log(`=== ${name}（${rows.length} 条${kw ? `，含 "${kw}"` : ''}）===`);
  for (const r of rows) {
    const cl = String(r.CommandLine || '');
    const short = cl.length > 190 ? cl.slice(0, 190) + '…' : cl;
    console.log(`  PID=${r.ProcessId} PPID=${r.ParentProcessId}  ${short}`);
  }
}
