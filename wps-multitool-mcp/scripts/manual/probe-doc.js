/**
 * Input: 无
 * Output: WPS 文字的 COM 状态 + wps/et/excel 进程数（排查 ActiveDocument 为何为空）
 * Pos: 诊断脚本。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 */

const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '../..');
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-probedoc';
process.env.WPS_BRIDGE_IDLE_MS = '60000';

const { psDaemon } = require(path.join(ROOT, 'dist', 'client', 'ps-daemon.js'));

function psOut(script) {
  try {
    return execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script], {
      encoding: 'utf8',
    }).trim();
  } catch (e) {
    return 'ERR: ' + e.message.slice(0, 200);
  }
}

(async () => {
  for (let i = 0; i < 2; i++) {
    const r = await psDaemon.call('__probeDoc', {}, 'word').catch((e) => ({ success: false, error: e.message }));
    console.log(`probe ${i + 1}: ${JSON.stringify(r)}`);
  }

  console.log('进程数:', psOut("(Get-Process wps,et,EXCEL -EA SilentlyContinue | Group-Object ProcessName | ForEach-Object { $_.Name + '=' + $_.Count }) -join ' '"));

  // 有没有模态对话框：看 wps 进程的窗口标题
  console.log(
    'WPS 窗口:',
    psOut("(Get-Process wps,et -EA SilentlyContinue | Where-Object { $_.MainWindowTitle } | ForEach-Object { \"$($_.ProcessName)[$($_.Id)]=$($_.MainWindowTitle)\" }) -join ' | '")
  );

  psDaemon.dispose();
  process.exit(0);
})();
