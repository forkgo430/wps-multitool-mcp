/**
 * Input: GITHUB_TOKEN、GITHUB_REPOSITORY、package.json 的 version
 * Output: 对应 tag 的 GitHub Release 正文追加/更新"安装脚本校验"节（install.ps1 的 SHA256）
 * Pos: 发布流程的收尾步（release.yml 在 changesets publish 成功后调用）。
 *      评审 P1-5：README 让用户"与 Release 页公布的校验值比对"，这一步保证那个值真的存在。
 *
 * 幂等：正文里已有"安装脚本校验"节则整节替换，不会重复追加。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const https = require('https');

const ROOT = path.resolve(__dirname, '..', '..');
const version = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
const tag = `v${version}`;
const repo = process.env.GITHUB_REPOSITORY;
if (!repo || !process.env.GITHUB_TOKEN) {
  console.error('✗ 需要 GITHUB_REPOSITORY 与 GITHUB_TOKEN（仅在 release workflow 里运行）');
  process.exit(1);
}

const hash = crypto
  .createHash('sha256')
  .update(fs.readFileSync(path.join(ROOT, '..', 'scripts', 'install.ps1')))
  .digest('hex');

const H = {
  'User-Agent': 'attach-install-hash',
  Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
  'Content-Type': 'application/json',
};

function req(method, p, body) {
  return new Promise((res) => {
    const rq = https.request(
      { host: 'api.github.com', path: p, method, headers: { ...H, ...(body ? { 'Content-Length': Buffer.byteLength(body) } : {}) } },
      (r) => {
        let b = '';
        r.on('data', (d) => (b += d));
        r.on('end', () => res({ code: r.statusCode, json: JSON.parse(b || '{}') }));
      }
    );
    if (body) rq.write(body);
    rq.end();
  });
}

(async () => {
  const got = await req('GET', `/repos/${repo}/releases/tags/${tag}`);
  if (got.code !== 200) {
    console.log(`Release ${tag} 不存在，跳过（本步骤只跟随已发布的版本）`);
    process.exit(0);
  }
  const rel = got.json;
  const section = `## 安装脚本校验\n\ninstall.ps1 的 SHA256：\n\n\`${hash}\``;
  const base = (rel.body || '').replace(/\n*## 安装脚本校验\n*[\s\S]*$/, '').trimEnd();
  const patch = await req('PATCH', `/repos/${repo}/releases/${rel.id}`, JSON.stringify({ body: base + '\n\n' + section + '\n' }));
  console.log(`✓ 安装脚本校验值已写入 ${tag} 的 Release 正文（HTTP ${patch.code}，sha ${hash.slice(0, 16)}…）`);
})();
