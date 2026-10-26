/**
 * 孤儿实验：分别验证「Node 正常退出」和「Node 被强杀」后，守护进程会不会自己走。
 * 用法：node orphan-test.js normal | killed
 */
const path = require('path');
const dir = require('path').resolve(__dirname, '../../..');
const isKilled = process.argv.includes('killed');

const { psDaemon } = require(path.join(dir, 'dist/client/ps-daemon.js'));

(async () => {
  await psDaemon.call('ping', {});
  const pid = psDaemon.proc ? psDaemon.proc.pid : null;
  console.log(`PARENT_PID=${process.pid}`);
  console.log(`DAEMON_PID=${pid}`);
  console.log(`MODE=${isKilled ? 'killed' : 'normal'}`);
  console.log('READY');
  if (isKilled) {
    // 模拟 MCP server 被强杀：不 dispose、不清理，直接 TerminateProcess 自己
    // （Windows 上 process.kill(pid,'SIGKILL') 就是 TerminateProcess，等价 taskkill /F）
    setTimeout(() => {
      process.kill(process.pid, 'SIGKILL');
    }, 800);
  } else {
    // 模拟正常退出：什么都不做，直接让事件循环结束
    process.exit(0);
  }
})();
