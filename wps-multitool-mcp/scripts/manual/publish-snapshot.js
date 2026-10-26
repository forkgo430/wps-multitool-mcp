#!/usr/bin/env node
/**
 * Input: 本地 HEAD 的树 + ~/.codebuddy/mcp.json 里的 GitHub Token
 * Output: 远端 main 上出现一个内容与本地 HEAD 完全一致的无父级提交
 * Pos: 发布快照工具（维护者用）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 什么时候用它（以及为什么不用 git push）
 * --------------------------------------
 * 本仓库是【浅克隆】出来的：本地 main 的历史边界之上还挂着一个不存在的祖先对象。
 * 直接 `git push` 会报 `did not receive expected object …`（GitHub 校验包时要求祖先齐全）。
 * 同时，这个项目也不该把上游祖先历史带进新仓库。
 * 所以发布 = 用当前树建一个【无父级根提交】再强推 —— 远端永远只有干净的快照历史。
 *
 * 代价（要知情）：远端历史不保留本地提交的父链，`git log` 在 GitHub 上只看到快照。
 * 本地历史（含每个改动的来龙去脉）不受影响。若哪天换成完整克隆，普通 git push 更好。
 *
 * 用法：node scripts/manual/publish-snapshot.js "提交信息"
 * 环境：需要 git 能读 origin，且 Token 在 ~/.codebuddy/mcp.json 的 GitHub 连接器里。
 */

const { execSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..'); // wps-multitool-mcp/
const REPO_ROOT = path.resolve(ROOT, '..');
const MCP_CFG = path.join(os.homedir(), '.codebuddy', 'mcp.json');

const run = (cmd, cwd = REPO_ROOT) =>
  execSync(cmd, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

function readToken() {
  const cfg = JSON.parse(fs.readFileSync(MCP_CFG, 'utf8'));
  const t = cfg?.mcpServers?.GitHub?.env?.GITHUB_PERSONAL_ACCESS_TOKEN;
  if (!t) {
    console.error('✗ 没在 ~/.codebuddy/mcp.json 里找到 GitHub Token');
    process.exit(1);
  }
  return t;
}

function originUrl() {
  const url = run('git remote get-url origin').trim();
  const m = url.match(/github\.com[/:]([^/]+)\/([^/.]+)(?:\.git)?$/);
  if (!m) {
    console.error(`✗ 解析不了 origin: ${url}`);
    process.exit(1);
  }
  return { owner: m[1], repo: m[2], clean: `https://github.com/${m[1]}/${m[2]}.git` };
}

(async () => {
  const msg = process.argv[2] || `snapshot ${new Date().toISOString()}`;
  const { owner, repo, clean } = originUrl();
  const token = readToken();

  // 用 %T 取树哈希：cmd 会把 ^ 当转义符（HEAD^{tree} 会被吃成 HEAD{tree}）
  const tree = run('git show -s --format=%T HEAD').trim();
  const commit = run(`git commit-tree ${tree} -m "${msg.replace(/"/g, '\\"')}"`).trim();
  console.log(`快照提交 ${commit.slice(0, 7)}（树 ${tree.slice(0, 7)}）→ ${owner}/${repo}`);

  const authed = `https://x-access-token:${token}@github.com/${owner}/${repo}.git`;
  run(`git remote set-url origin ${authed}`);
  let ok = false;
  try {
    for (let i = 1; i <= 8; i++) {
      try {
        run(`git push -f origin ${commit}:refs/heads/main`);
        ok = true;
        console.log(`push 成功（第 ${i} 次尝试）`);
        break;
      } catch (e) {
        const tail = `${e.stderr || e.message || ''}`.split('\n').filter(Boolean).slice(-2).join(' | ');
        console.log(`第 ${i} 次失败: ${tail.slice(0, 140)}（等 ${i * 5}s 重试）`);
        if (i < 8) execSync(`node -e "setTimeout(()=>{},${i * 5000})"`);
      }
    }
  } finally {
    run(`git remote set-url origin ${clean}`);
  }

  if (!ok) {
    console.error('✗ 8 次均失败，网络不可用');
    process.exit(1);
  }

  // 自检：远端 main 必须等于刚推的快照
  const refs = run(`git ls-remote ${clean} main`).trim();
  if (!refs.startsWith(commit)) {
    console.error(`✗ 远端 main 不是刚推的提交: ${refs}`);
    process.exit(1);
  }
  console.log(`✓ 远端 main = ${commit.slice(0, 7)}（内容与本地 HEAD 一致）`);
  process.exit(0);
})();
