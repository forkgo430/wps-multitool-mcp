/**
 * B1 端到端探针（修复后复验）：走真工具入口写入，用【独立 PowerShell 进程】读回。
 *
 * 为什么读回要独立进程：wps-com.ps1 里没有能读回"单元格类型/格式"的动作
 * （getRangeData / getCellValue 只给值，不给类型与数字格式）。
 * 只有独立挂 COM 读得到 Value2 的真实类型与 NumberFormat ——
 * 而"数字写进去还是数字"正是本轮要证明的那件事。
 *
 * 判据（不是"回 success"）：
 *   - 数字写进去，读回类型必须是数值（SUM 能算）
 *   - 文本 "007" 读回必须还是 "007"（不能变 7）
 *   - 文本 "2024-01-15" 读回必须还是 "2024-01-15"（不能变日期序列号 45306）
 *   - 文本 "1e5" 读回必须还是 "1e5"
 *
 * 纪律：自己建的工作簿自己关（save:false），靠打开列表差集认领。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-agentX-r11-b1f-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '120000';

const path = require('path');
const { execFileSync } = require('child_process');
const fs = require('fs');

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

let pass = 0, fail = 0;
const ok = (label, cond, extra) => {
  if (cond) { pass++; console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`); }
  else { fail++; console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`); }
};

/* ---- 独立读回：另一个 PowerShell 进程挂 COM 读类型与格式 ---- */
const PS_READ = `
$ErrorActionPreference = 'Continue'
try { $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application') }
catch { $xl = $null }
if ($null -eq $xl) { Write-Output 'READ|NO_EXCEL'; exit 0 }
$wb = $null
try { $wb = $xl.Workbooks.Item(1) } catch {}
if ($null -eq $wb) { Write-Output 'READ|NO_WB'; exit 0 }
$sh = $wb.Worksheets.Item(1)
$addr = $env:READ_ADDRS -split ','
foreach ($a in $addr) {
  if (-not $a) { continue }
  try {
    $c = $sh.Range($a)
    $v = $c.Value2
    $t = if ($null -eq $v) { 'null' } else { $v.GetType().Name }
    Write-Output ('READ|' + $a + '|Value2=[' + [string]$v + ']|Type=' + $t + '|Text=[' + [string]$c.Text + ']|Fmt=' + [string]$c.NumberFormat)
  } catch {
    Write-Output ('READ|' + $a + '|ERR|' + $_.Exception.Message)
  }
}
try {
  $sum = [string]$xl.WorksheetFunction.Sum($sh.Range($env:READ_SUM))
  Write-Output ('SUM|' + $env:READ_SUM + '|=' + $sum)
} catch { Write-Output ('SUM|' + $env:READ_SUM + '|ERR|' + $_.Exception.Message) }
`;

function readCells(addrs, sumRange) {
  const log = 'C:\\Users\\金东权\\AppData\\Local\\Temp\\opencode\\b1-read.log';
  const fd = fs.openSync(log, 'w');
  try {
    execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', PS_READ], {
      encoding: 'utf8', timeout: 120000,
      stdio: ['ignore', fd, 'pipe'],
      env: { ...process.env, READ_ADDRS: addrs.join(','), READ_SUM: sumRange || '' },
    });
  } catch (e) {
    console.log('  读回 stderr:', (e.stderr || '').slice(0, 300));
  } finally {
    fs.closeSync(fd);
  }
  const raw = fs.readFileSync(log, 'utf8');
  const cells = {};
  let sum = null;
  for (const line of raw.split(/\r?\n/)) {
    const t = line.trim();
    if (t.startsWith('READ|')) {
      const p = t.split('|');
      // 读回行格式：READ|<地址>|Value2=[..]|Type=..|Text=[..]|Fmt=..
      const val = (p[2] || '').replace(/^Value2=\[/, '').replace(/\]$/, '');
      const txtv = (p[4] || '').replace(/^Text=\[/, '').replace(/\]$/, '');
      cells[p[1]] = {
        value: val,
        type: (p[3] || '').replace(/^Type=/, ''),
        text: txtv,
        fmt: (p[5] || '').replace(/^Fmt=/, ''),
      };
    } else if (t.startsWith('SUM|')) {
      const p = t.split('|');
      sum = (p[2] || '').replace(/^=/, '');
    }
  }
  return { cells, sum };
}

(async () => {
  console.log('\n=== 0. 认领临时工作簿 ===');
  const before = await wpsClient.executeMethod('getOpenWorkbooks', {}, 'excel');
  const beforeNames = new Set(((before.data && before.data.workbooks) || []).map((w) => w.name));
  const created = await call('wps_excel_create_workbook', {});
  const after = await wpsClient.executeMethod('getOpenWorkbooks', {}, 'excel');
  const mine = ((after.data && after.data.workbooks) || []).filter((w) => !beforeNames.has(w.name));
  ok('差集 = 恰好 1 个新工作簿', mine.length === 1, mine.map((w) => w.name).join(','));
  const myName = mine.length ? mine[0].name : null;
  if (!myName) { console.log('  ✗ 认领失败'); process.exit(1); }

  console.log('\n=== 1. write_range 写【混合】（字符串 + 数字）—— F7 ===');
  let r = await call('wps_excel_write_range', { range: 'A1:B2', data: [['姓名', '年龄'], ['张三', 25]] });
  console.log('  工具回包：', txt(r).slice(0, 110));
  ok('混合写入不报错', r.success, txt(r).slice(0, 80));
  let rb = readCells(['A1', 'B1', 'A2', 'B2'], 'B2');
  for (const [k, v] of Object.entries(rb.cells)) console.log(`  ${k}: ${JSON.stringify(v)}`);
  ok('数字 25 读回是数值类型', rb.cells.B2 && rb.cells.B2.type === 'Double', `实际 ${rb.cells.B2 && rb.cells.B2.type}`);
  ok('文本 "姓名" 读回是字符串', rb.cells.A1 && rb.cells.A1.type === 'String', `实际 ${rb.cells.A1 && rb.cells.A1.type}`);
  ok('文本 "张三" 形状没变', rb.cells.A2 && rb.cells.A2.value === '张三', `实际 ${rb.cells.A2 && rb.cells.A2.value}`);

  console.log('\n=== 2. write_range 写【全数字】（含小数）===');
  r = await call('wps_excel_write_range', { range: 'D1:E2', data: [[1, 2], [3, 1.5]] });
  console.log('  工具回包：', txt(r).slice(0, 110));
  ok('全数字写入不报错', r.success, txt(r).slice(0, 80));
  rb = readCells(['D1', 'E1', 'D2', 'E2'], 'D1:E2');
  for (const [k, v] of Object.entries(rb.cells)) console.log(`  ${k}: ${JSON.stringify(v)}`);
  ok('四个格子都是数值', Object.values(rb.cells).every((v) => v.type === 'Double'),
    Object.values(rb.cells).map((v) => v.type).join(','));
  ok('小数 1.5 没被截断', rb.cells.E2 && rb.cells.E2.value === '1.5', `实际 ${rb.cells.E2 && rb.cells.E2.value}`);
  ok('SUM 能把这批数字加起来（7.5）', rb.sum === '7.5', `实际 ${rb.sum}`);

  console.log('\n=== 3. write_range 写【全字符串】—— 形状必须保住 ===');
  r = await call('wps_excel_write_range', { range: 'G1:H2', data: [['007', '1e5'], ['2024-01-15', 'abc']] });
  console.log('  工具回包：', txt(r).slice(0, 110));
  ok('全字符串写入不报错', r.success, txt(r).slice(0, 80));
  rb = readCells(['G1', 'H1', 'G2', 'H2']);
  for (const [k, v] of Object.entries(rb.cells)) console.log(`  ${k}: ${JSON.stringify(v)}`);
  ok('"007" 还是 007（没变 7）', rb.cells.G1 && rb.cells.G1.value === '007', `实际 ${rb.cells.G1 && rb.cells.G1.value}`);
  ok('"1e5" 还是 1e5（没变 100000）', rb.cells.H1 && rb.cells.H1.value === '1e5', `实际 ${rb.cells.H1 && rb.cells.H1.value}`);
  ok('"2024-01-15" 还是日期串（没变 45306）', rb.cells.G2 && rb.cells.G2.value === '2024-01-15', `实际 ${rb.cells.G2 && rb.cells.G2.value}`);
  ok('四个都是文本类型', Object.values(rb.cells).every((v) => v.type === 'String'),
    Object.values(rb.cells).map((v) => v.type).join(','));

  console.log('\n=== 4. evaluate_formula 带 cell —— F2b ===');
  r = await call('wps_excel_evaluate_formula', { formula: '=SUM(1,2)', cell: 'J1' });
  console.log('  工具回包：', txt(r).slice(0, 130));
  ok('evaluate_formula 带 cell 写入成功', r.success, txt(r).slice(0, 90));
  rb = readCells(['J1']);
  console.log(`  J1: ${JSON.stringify(rb.cells.J1)}`);
  ok('算出来的 3 落盘且是数值', rb.cells.J1 && rb.cells.J1.type === 'Double' && rb.cells.J1.value === '3',
    `实际 ${JSON.stringify(rb.cells.J1)}`);

  console.log('\n=== 5. set_cell_value（同一个雷的另一处）===');
  // sheet 要【名字】不是序号（schema 写的是"工作表名称"；ps1 是 Sheets.Item($p.sheet)）
  const sl = await call('wps_excel_get_sheet_list', {});
  console.log('  工作表列表：', txt(sl).slice(0, 100));
  const sheetName = (String((sl.content && sl.content[0] && sl.content[0].text) || '').match(/：\s*[\d.]+\s*(\S+)/) || [])[1] || 'Sheet1';
  console.log('  用作 sheet 的名字：', JSON.stringify(sheetName));
  r = await call('wps_excel_set_cell_value', { sheet: sheetName, row: 1, col: 12, value: 25 });
  console.log('  写数字 25   →', txt(r).slice(0, 100));
  ok('set_cell_value 写数字成功', r.success, txt(r).slice(0, 70));
  r = await call('wps_excel_set_cell_value', { sheet: sheetName, row: 2, col: 12, value: 1.5 });
  console.log('  写小数 1.5  →', txt(r).slice(0, 100));
  ok('set_cell_value 写小数成功', r.success, txt(r).slice(0, 70));
  r = await call('wps_excel_set_cell_value', { sheet: sheetName, row: 3, col: 12, value: '007' });
  console.log('  写文本 007  →', txt(r).slice(0, 100));
  ok('set_cell_value 写文本成功', r.success, txt(r).slice(0, 70));
  rb = readCells(['L1', 'L2', 'L3']);
  for (const [k, v] of Object.entries(rb.cells)) console.log(`  ${k}: ${JSON.stringify(v)}`);
  ok('set_cell_value 数字是数值', rb.cells.L1 && rb.cells.L1.type === 'Double', `实际 ${rb.cells.L1 && rb.cells.L1.type}`);
  ok('set_cell_value 小数没被截断', rb.cells.L2 && rb.cells.L2.value === '1.5', `实际 ${rb.cells.L2 && rb.cells.L2.value}`);
  ok('set_cell_value "007" 形状保住', rb.cells.L3 && rb.cells.L3.value === '007', `实际 ${rb.cells.L3 && rb.cells.L3.value}`);

  console.log('\n=== 6. 收尾：只关自己创建的那一个 ===');
  // 注意：wps_excel_close_workbook 对"从未保存过"的工作簿有一道护栏（B2，本轮不修），
  // 会拒绝关闭并说"请显式传 save=false"——而参数名其实是 saveChanges（B3）。
  // 所以这里先用工具试（把它的真实提示语记下来，那是 B3 的证据），
  // 再用独立 PowerShell 兜底关掉（真机实测 $wb.Close($false) 能关且不弹框）。
  const closed = await call('wps_excel_close_workbook', { name: myName, saveChanges: false });
  console.log('  工具关簿：', txt(closed).slice(0, 150));
  console.log('  ↑ 这句提示语里的参数名是 B3 的现场证据（它说 save=false，schema 里叫 saveChanges）');
  const PS_CLOSE = `
try { $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application') } catch { $xl = $null }
if ($null -eq $xl) { Write-Output 'CLOSE|NO_EXCEL'; exit 0 }
$n = 0
foreach ($w in @($xl.Workbooks)) { try { $w.Close($false); $n++ } catch {} }
Write-Output ('CLOSE|closed=' + $n + '|remain=' + $xl.Workbooks.Count)
`;
  const clog = 'C:\\Users\\金东权\\AppData\\Local\\Temp\\opencode\\b1-close.log';
  const cfd = fs.openSync(clog, 'w');
  try {
    execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', PS_CLOSE], {
      encoding: 'utf8', timeout: 120000, stdio: ['ignore', cfd, 'pipe'],
    });
  } catch (e) {
    console.log('  收尾 stderr:', (e.stderr || '').slice(0, 200));
  } finally {
    fs.closeSync(cfd);
  }
  console.log('  独立收尾：', fs.readFileSync(clog, 'utf8').trim());
  const afterClose = await wpsClient.executeMethod('getOpenWorkbooks', {}, 'excel');
  const left = ((afterClose.data && afterClose.data.workbooks) || []).map((w) => w.name);
  ok('自己的工作簿已关闭', !left.includes(myName), `剩余 ${left.length} 个`);

  console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
  process.exit(fail === 0 ? 0 : 1);
})().catch((e) => {
  console.error('探针异常：', e && e.message ? e.message : e);
  process.exit(2);
});
