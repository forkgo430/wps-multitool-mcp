/**
 * gen-install-hash.js —— 安装脚本防篡改物料（2026-10-25，外部评审 P1-5）
 *
 * 背景：README 的"一行命令安装"（irm https://cdn.jsdelivr.net/...@main/scripts/install.ps1 | iex）
 * 是从网上拉脚本直接执行，中间任何一环被动手脚用户都察觉不到。
 *
 * 本脚本算出 scripts/install.ps1 的 SHA256，生成一段可直接粘贴的校验指引，
 * 写到 docs/install-hash.txt（docs/ 不进 GitHub，作为发版时的底稿）并打印到控制台。
 * 发版时的动作：把"期望值"贴到 GitHub Release 页 —— 用户下载后用 Get-FileHash 比对。
 *
 * 注意：@main 的内容会随提交变化，所以校验值必须【随发版】更新；
 * 锚定不变更的版本请把安装命令里的 @main 换成 @<tag>。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.resolve(__dirname, '..', '..'); // → wps-multitool-mcp/
const REPO = path.resolve(ROOT, '..');            // → 仓库根
const target = path.join(REPO, 'scripts', 'install.ps1');

if (!fs.existsSync(target)) {
  console.error(`✗ 找不到安装脚本：${target}`);
  process.exit(1);
}

const hash = crypto.createHash('sha256').update(fs.readFileSync(target)).digest('hex');

const guidance = [
  '# install.ps1 校验物料（由 scripts/manual/gen-install-hash.js 生成）',
  `# 生成时间：${new Date().toISOString()}`,
  `# 对应文件：scripts/install.ps1（${fs.statSync(target).size} 字节）`,
  '',
  `SHA256 = ${hash}`,
  '',
  '# 用户侧校验命令（Windows PowerShell）：',
  '#   Invoke-WebRequest https://cdn.jsdelivr.net/gh/forkgo430/wps-multitool-mcp@main/scripts/install.ps1 -OutFile install.ps1',
  '#   Get-FileHash .\\install.ps1 -Algorithm SHA256   ← 结果应等于上面的期望值',
  '# 发版时把"SHA256 = …"这一行贴到 GitHub Release 页的说明里。',
].join('\n');

const outPath = path.join(REPO, 'docs', 'install-hash.txt');
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, guidance + '\n');
console.log(guidance);
console.log(`\n已写入 ${outPath}`);
