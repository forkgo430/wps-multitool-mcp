#!/usr/bin/env node
/**
 * Input: 本地 HEAD 的 tracked 文件树 + ~/.codebuddy/mcp.json 里的 GitHub Token
 * Output: 远端 refs/heads/main 指向一个内容与本地 HEAD 完全一致的快照提交
 * Pos: 发布通道 B（维护者用，走 api.github.com 的 Git Data API）。一旦我被修改，请更新我的头部注释，以及所属目录的 md。
 *
 * 什么时候用它（以及为什么还要 publish-snapshot.js）
 * ------------------------------------------------
 * 通道 A = scripts/manual/publish-snapshot.js，用 git 协议推（要求 github.com:443 通）。
 * 本脚本是同一个结果的第二条通道：全程只用 HTTPS 443 上的 REST API，
 * 适合 git 协议被墙或被代理掐掉的环境。
 *
 * 流程与 publish-snapshot.js 保持同一约定：建【无父级】的根提交再强推，
 * 远端永远是干净的快照历史；本地历史不受影响。
 *
 * 步骤：git ls-tree -r HEAD → 逐个 POST /git/blobs → POST /git/trees → POST /git/commits → PATCH /git/refs/heads/分支
 * 每步自检：blob sha 必须回填一致、生成的 tree sha 必须等于本地 %T、推完后远端 tree 必须等于本地 %T。
 *
 * 用法：
 *   node scripts/manual/push-api.js "提交信息"           # 与远端同一棵树时自动跳过
 *   node scripts/manual/push-api.js "提交信息" -f        # 强制重推（换提交信息时用）
 *   node scripts/manual/push-api.js "提交信息" --dry     # 只建对象，不动 ref（验证连通性用）
 *   node scripts/manual/push-api.js "提交信息" --branch gh-pages
 * 环境：Token 优先取 GITHUB_TOKEN，否则读 ~/.codebuddy/mcp.json 的 GitHub 连接器。
 *       Token 需要 Contents 读写权限；推荐另开免 2FA 的 Granular Token（见 HANDOFF 待办）。
 */

const { execSync, execFileSync } = require('child_process');
const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..'); // wps-multitool-mcp/
const REPO_ROOT = path.resolve(ROOT, '..');
const MCP_CFG = path.join(os.homedir(), '.codebuddy', 'mcp.json');
const MAX_BLOB = 95 * 1024 * 1024; // GitHub 单 blob 上限 100MB，留点余量

const run = (cmd, cwd = REPO_ROOT) =>
  execSync(cmd, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 1 << 26 });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------- 入参 ----------
const argv = process.argv.slice(2);
const dry = argv.includes('--dry');
const force = argv.includes('-f') || argv.includes('--force');
const historyMode = argv.includes('--history');
const bIdx = argv.indexOf('--branch');
const branch = bIdx >= 0 ? argv[bIdx + 1] : 'main';
const taken = new Set([...(bIdx >= 0 ? [argv[bIdx], argv[bIdx + 1]] : [])]);
const msg = argv.find((a) => !a.startsWith('-') && !taken.has(a)) || `snapshot ${new Date().toISOString()}`;

// ---------- 凭据 / 仓库 ----------
function readToken() {
  if (process.env.GITHUB_TOKEN) return process.env.GITHUB_TOKEN;
  try {
    const cfg = JSON.parse(fs.readFileSync(MCP_CFG, 'utf8'));
    const t = cfg && cfg.mcpServers && cfg.mcpServers.GitHub
      && cfg.mcpServers.GitHub.env && cfg.mcpServers.GitHub.env.GITHUB_PERSONAL_ACCESS_TOKEN;
    if (t) return t;
  } catch (e) { /* 落到下面的报错 */ }
  console.error('没找到 Token：请设 GITHUB_TOKEN，或在 ~/.codebuddy/mcp.json 的 GitHub 连接器里配 PAT');
  process.exit(1);
}

function originUrl() {
  const url = run('git remote get-url origin').trim();
  const m = url.match(/github\.com[/:]([^/]+)\/([^/.]+)(?:\.git)?$/);
  if (!m) {
    console.error(`解析不了 origin: ${url}`);
    process.exit(1);
  }
  return { owner: m[1], repo: m[2] };
}

// ---------- HTTPS ----------
function api(method, urlPath, token, body) {
  return new Promise((resolve, reject) => {
    const payload = body === undefined ? null : Buffer.from(JSON.stringify(body), 'utf8');
    const headers = {
      'User-Agent': 'wps-multitool-mcp-push-api',
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${token}`,
      'X-GitHub-Api-Version': '2022-11-28',
    };
    if (payload) {
      headers['Content-Type'] = 'application/json';
      headers['Content-Length'] = payload.length;
    }
    const r = https.request({ host: 'api.github.com', path: urlPath, method, timeout: 120000, headers }, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (c) => { data += c; });
      res.on('end', () => {
        let json = null;
        try { json = data ? JSON.parse(data) : null; } catch (e) { /* 非 JSON 走 status 分支 */ }
        if (res.statusCode >= 200 && res.statusCode < 300) return resolve({ json, headers: res.headers });
        const err = new Error(`${method} ${urlPath} -> HTTP ${res.statusCode} ${(json && json.message) || data.slice(0, 160)}`);
        err.status = res.statusCode;
        // 403 可能是限流也可能是真没权限；只有配额明确告罄时才等
        const remaining = res.headers['x-ratelimit-remaining'];
        err.waitMs = remaining === '0'
          ? Math.max(0, Number(res.headers['x-ratelimit-reset']) * 1000 - Date.now() + 1000)
          : 0;
        reject(err);
      });
    });
    r.on('timeout', () => r.destroy(new Error(`${method} ${urlPath} 超时（120s）`)));
    r.on('error', reject);
    if (payload) r.write(payload);
    r.end();
  });
}

async function withRetry(label, fn, tries = 6) {
  for (let i = 1; i <= tries; i++) {
    try {
      return await fn();
    } catch (e) {
      const retryable = !e.status || e.status >= 500 || e.status === 429 || e.waitMs > 0;
      if (i === tries || !retryable) throw e;
      const wait = e.waitMs || Math.min(30000, i * i * 2000);
      console.log(`  retry ${label} 第 ${i} 次失败（${e.message.slice(0, 100)}），等 ${Math.round(wait / 1000)}s`);
      await sleep(wait);
    }
  }
}

// ---------- 读本地树 ----------
function localEntries() {
  // -z 避免中文/空格文件名被拆开：<mode> SP <type> SP <sha> TAB <path> NUL
  const raw = execFileSync('git', ['ls-tree', '-r', '-z', 'HEAD'], {
    cwd: REPO_ROOT, maxBuffer: 1 << 26, stdio: ['ignore', 'pipe', 'pipe'],
  }).toString('utf8');
  return raw.split('\0').filter(Boolean).map((line) => {
    const parts = line.split('\t');
    const meta = parts[0].split(/\s+/);
    return { mode: meta[0], type: meta[1], sha: meta[2], path: parts.slice(1).join('\t') };
  });
}

function blobContent(sha) {
  return execFileSync('git', ['cat-file', 'blob', sha], {
    cwd: REPO_ROOT, maxBuffer: 1 << 28, stdio: ['ignore', 'pipe', 'pipe'],
  });
}

// ---------- 模式 B：按真实历史推送（--history） ----------
// 快照模式会把历史压成一个无父级提交；--history 则逐个重放本地提交，
// 让远端拥有可追溯的提交链。每一步都建在【已创建的上一个提交】之上，
// 所以新仓库的父链是完整的 —— GitHub 不会因为缺祖先而拒绝。

function parseCommit(sha) {
  const raw = execFileSync('git', ['cat-file', 'commit', sha], {
    cwd: REPO_ROOT, maxBuffer: 1 << 28,
  }).toString('utf8');
  const cut = raw.indexOf('\n\n');
  const header = cut >= 0 ? raw.slice(0, cut) : raw;
  const message = cut >= 0 ? raw.slice(cut + 2) : '';
  const info = { tree: '', parents: [], author: '', committer: '' };
  for (const line of header.split('\n')) {
    if (line.startsWith('tree ')) info.tree = line.slice(5).trim();
    else if (line.startsWith('parent ')) info.parents.push(line.slice(7).trim());
    else if (line.startsWith('author ')) info.author = line.slice(7);
    else if (line.startsWith('committer ')) info.committer = line.slice(10);
  }
  const ident = (s) => {
    const m = s.match(/^(.*) <(.*)> (\d+) ([+-]\d{4})$/);
    if (!m) return { name: 'unknown', email: 'unknown@example.com', date: new Date().toISOString() };
    return { name: m[1], email: m[2], date: new Date(Number(m[3]) * 1000).toISOString() };
  };
  return { ...info, message, authorObj: ident(info.author), committerObj: ident(info.committer) };
}

/** 该提交相对父提交改了哪些文件（根提交用 --root 全量列出） */
function changedFiles(sha) {
  const out = execFileSync(
    'git',
    ['-c', 'core.quotepath=false', 'diff-tree', '-r', '--no-commit-id', '--name-status', '--root', sha],
    { cwd: REPO_ROOT, encoding: 'utf8', maxBuffer: 1 << 26 }
  );
  const rows = [];
  for (const line of out.split('\n')) {
    if (!line.trim()) continue;
    const tab = line.indexOf('\t');
    const status = line.slice(0, tab).trim().charAt(0); // A/M/D/R...
    const p = line.slice(tab + 1);
    if (status === 'D') {
      rows.push({ status: 'D', path: p });
      continue;
    }
    const meta = execFileSync('git', ['-c', 'core.quotepath=false', 'ls-tree', sha, '--', p], {
      cwd: REPO_ROOT, encoding: 'utf8', maxBuffer: 1 << 20,
    }).trim();
    const [mode, type, blob] = meta.split('\t')[0].split(/\s+/);
    if (type !== 'blob') continue; // 子模块之类，Git Data API 不收
    rows.push({ status: 'A', path: p, mode, blob });
  }
  return rows;
}

async function pushHistory({ owner, repo, token }) {
  const base = `/repos/${owner}/${repo}`;
  const localTree = run('git show -s --format=%T HEAD').trim();
  const list = run('git rev-list --reverse HEAD').trim().split('\n').filter(Boolean);
  console.log(`重放 ${list.length} 个提交 → ${owner}/${repo}@${branch}（目标树 ${localTree.slice(0, 7)}）`);

  const mapped = new Map(); // 旧 sha → 新 sha
  const treeOf = new Map(); // 旧 sha → 新 tree sha
  const uploaded = new Set(); // 已上传过的 blob sha（同一内容只传一次）
  let uploadedCount = 0;
  let reused = 0;

  // 远端已有的提交按 tree 建索引：内容一模一样就复用。
  // 没有这一步，每次推送都得把整条链重传一遍（54 个提交 / 400 个 blob / 8 分钟），
  // 而绝大多数提交其实上次已经推过了 —— 只该补新的那几条。
  const remoteByTree = new Map();
  try {
    const res = await withRetry('读远端提交', () =>
      api('GET', `${base}/commits?sha=${branch}&per_page=100`, token));
    for (const c of res.json || []) {
      remoteByTree.set(c.commit.tree.sha, { sha: c.sha, parents: c.parents.map((p) => p.sha) });
    }
    console.log(`远端已有 ${remoteByTree.size} 个提交可按 tree 复用`);
  } catch (e) {
    console.log(`读远端提交失败（${e.message.slice(0, 60)}），将全量重建`);
  }

  for (let i = 0; i < list.length; i++) {
    const sha = list[i];
    const c = parseCommit(sha);
    const parentNew = c.parents.map((p) => mapped.get(p)).filter(Boolean);

    // 树上完全一致、父链也对得上 → 直接复用远端那个提交，不再上传任何东西
    const hit = remoteByTree.get(c.tree);
    if (
      hit &&
      hit.parents.length === parentNew.length &&
      hit.parents.every((p, idx) => p === parentNew[idx])
    ) {
      mapped.set(sha, hit.sha);
      treeOf.set(sha, c.tree);
      reused++;
      process.stdout.write(`\r  复用 ${reused}/${list.length}   `);
      continue;
    }

    const baseTree = c.parents[0] ? treeOf.get(c.parents[0]) : undefined;
    const files = changedFiles(sha);

    // 先并行把缺的 blob 传上去（串行会慢到看不下去：几百个 blob × 每次几百毫秒）
    const missing = files.filter((f) => f.status !== 'D' && !uploaded.has(f.blob));
    for (const f of missing) uploaded.add(f.blob); // 先占位，避免同一提交内重复
    let cursor = 0;
    await Promise.all(
      Array.from({ length: Math.min(6, missing.length || 1) }, async () => {
        while (cursor < missing.length) {
          const f = missing[cursor++];
          const content = execFileSync('git', ['cat-file', 'blob', f.blob], {
            cwd: REPO_ROOT, maxBuffer: 1 << 28,
          }).toString('base64');
          const res = await withRetry(`blob ${f.path}`, () =>
            api('POST', `${base}/git/blobs`, token, { content, encoding: 'base64' }));
          if (res.json.sha !== f.blob) throw new Error(`blob sha 不符：${f.path}`);
          uploadedCount++;
          process.stdout.write(`\r  提交 ${i + 1}/${list.length}  blob ${uploadedCount}   `);
        }
      })
    );

    const items = [];
    for (const f of files) {
      if (f.status === 'D') {
        items.push({ path: f.path, mode: '100644', type: 'blob', sha: null });
        continue;
      }
      items.push({ path: f.path, mode: f.mode, type: 'blob', sha: f.blob });
    }

    const treeRes = await withRetry(`tree ${sha.slice(0, 7)}`, () =>
      api('POST', `${base}/git/trees`, token, baseTree ? { base_tree: baseTree, tree: items } : { tree: items }));
    const newTree = treeRes.json.sha;

    const commitRes = await withRetry(`commit ${sha.slice(0, 7)}`, () =>
      api('POST', `${base}/git/commits`, token, {
        message: c.message,
        tree: newTree,
        parents: parentNew,
        author: c.authorObj,
        committer: c.committerObj,
      }));

    mapped.set(sha, commitRes.json.sha);
    treeOf.set(sha, newTree);
    process.stdout.write(`\r  提交 ${i + 1}/${list.length}（blob ${uploadedCount}）   `);
  }
  console.log('');

  const newTip = mapped.get(list[list.length - 1]);
  const tipTree = treeOf.get(list[list.length - 1]);
  if (tipTree !== localTree) {
    console.error(`✗ 远端树 ${tipTree} ≠ 本地 ${localTree}，不移动 ref`);
    process.exit(1);
  }
  console.log(
    `✓ 重建完成：${list.length} 个提交（复用 ${reused}，新建 ${list.length - reused}），` +
      `tip ${newTip.slice(0, 7)}，树与本地一致`
  );

  if (dry) {
    console.log('--dry：ref 未动');
    return;
  }

  try {
    await withRetry('update ref', () =>
      api('PATCH', `${base}/git/refs/heads/${branch}`, token, { sha: newTip, force: true }));
  } catch (e) {
    if (!`${e.message}`.includes('404')) throw e;
    await withRetry('create ref', () =>
      api('POST', `${base}/git/refs`, token, { ref: `refs/heads/${branch}`, sha: newTip }));
  }

  const check = await withRetry('自检', () => api('GET', `${base}/commits/heads/${branch}`, token));
  const head = check.json;
  const remoteTree = head && (head.commit ? head.commit.tree.sha : head.tree && head.tree.sha);
  if (remoteTree !== localTree) {
    console.error(`✗ 推完仍不一致：远端 ${remoteTree} vs 本地 ${localTree}`);
    process.exit(1);
  }
  console.log(`✓ 远端 ${branch} = ${newTip.slice(0, 7)}，${list.length} 个真实提交，树 ${remoteTree.slice(0, 7)}`);
}

// ---------- 主流程 ----------
(async () => {
  const { owner, repo } = originUrl();
  const token = readToken();
  const localTree = run('git show -s --format=%T HEAD').trim(); // %T 而非 HEAD^{tree}：cmd 会吃掉 ^
  const base = `/repos/${owner}/${repo}`;
  console.log(`本地树 ${localTree.slice(0, 7)} -> ${owner}/${repo}@${branch}${dry ? '（--dry，不动 ref）' : ''}`);

  if (historyMode) {
    await pushHistory({ owner, repo, token });
    process.exit(0);
  }

  // 0. 远端已经是同一棵树就别白跑
  if (!force) {
    let remoteTree = null;
    try {
      const res = await withRetry('读远端 HEAD', () => api('GET', `${base}/commits/heads/${branch}`, token));
      remoteTree = res.json && (res.json.commit ? res.json.commit.tree.sha : res.json.tree && res.json.tree.sha);
    } catch (e) {
      if (!`${e.message}`.includes('404')) throw e;
      console.log(`远端还没有 ${branch}（首次推送）`);
    }
    if (remoteTree === localTree) {
      console.log(`远端 ${branch} 已是同一棵树（${remoteTree.slice(0, 7)}），无事可做。换提交信息重推请加 -f`);
      process.exit(0);
    }
    if (remoteTree) console.log(`远端树 ${remoteTree.slice(0, 7)} != 本地，继续`);
  }

  // 1. 过滤不支持的对象
  const all = localEntries();
  const entries = [];
  for (const e of all) {
    if (e.type === 'commit') { console.log(`  跳过子模块 ${e.path}`); continue; }
    if (e.mode === '040000') continue; // ls-tree -r 不会给纯树，兜底
    const size = Number(run(`git cat-file -s ${e.sha}`).trim());
    if (size > MAX_BLOB) {
      console.error(`${e.path} 超过 95MB，Git Data API 收不下`);
      process.exit(1);
    }
    entries.push(e);
  }
  console.log(`${entries.length} 个文件待上传（本树共 ${all.length} 个对象）`);

  // 2. 逐个建 blob（并发 6，幂等：同内容重复 POST 返回同一个 sha）
  const blobs = new Map();
  let done = 0;
  const queue = entries.slice();
  await Promise.all(Array.from({ length: 6 }, async () => {
    while (queue.length) {
      const e = queue.shift();
      const content = blobContent(e.sha).toString('base64');
      const res = await withRetry(`blob ${e.path}`, () =>
        api('POST', `${base}/git/blobs`, token, { content, encoding: 'base64' }));
      if (res.json.sha !== e.sha) {
        throw new Error(`blob sha 不符：${e.path} 本地 ${e.sha.slice(0, 7)} vs 远端 ${res.json.sha.slice(0, 7)}`);
      }
      blobs.set(e.sha, res.json.sha);
      done += 1;
      if (done % 25 === 0 || done === entries.length) console.log(`  blob ${done}/${entries.length}`);
    }
  }));

  // 3. 一次性建树（路径带 / 时 GitHub 自动补中间子树）
  const treeItems = entries.map((e) => ({ path: e.path, mode: e.mode, type: 'blob', sha: blobs.get(e.sha) }));
  const treeRes = await withRetry('create tree', () => api('POST', `${base}/git/trees`, token, { tree: treeItems }));
  const tree = treeRes.json;
  if (tree.sha !== localTree) {
    console.error(`tree sha 不符：本地 ${localTree.slice(0, 7)} vs 远端 ${tree.sha.slice(0, 7)}（可能有文件被跳过）`);
    process.exit(1);
  }
  console.log(`tree ${tree.sha.slice(0, 7)}（与本地 %T 一致）`);

  // 4. 无父级根提交 —— 与 publish-snapshot.js 的快照约定一致
  const commitRes = await withRetry('create commit', () =>
    api('POST', `${base}/git/commits`, token, { message: msg, tree: tree.sha, parents: [] }));
  const commit = commitRes.json;
  console.log(`commit ${commit.sha.slice(0, 7)}`);

  if (dry) {
    console.log('--dry：对象都建好了，ref 未动（悬空对象 GitHub 会自行 GC）');
    process.exit(0);
  }

  // 5. 移动 ref（force=true：这一步本来就是快照覆盖）
  try {
    await withRetry('update ref', () =>
      api('PATCH', `${base}/git/refs/heads/${branch}`, token, { sha: commit.sha, force: true }));
  } catch (e) {
    if (!`${e.message}`.includes('404')) throw e;
    await withRetry('create ref', () =>
      api('POST', `${base}/git/refs`, token, { ref: `refs/heads/${branch}`, sha: commit.sha }));
  }

  // 6. 自检：远端这棵树必须等于本地 %T
  const check = await withRetry('自检远端 HEAD', () => api('GET', `${base}/commits/heads/${branch}`, token));
  const head = check.json;
  const remote = head && (head.commit ? head.commit.tree.sha : head.tree && head.tree.sha);
  if (remote !== localTree) {
    console.error(`推完仍不一致：远端树 ${remote} vs 本地 ${localTree}`);
    process.exit(1);
  }
  console.log(`远端 ${branch} = ${head.sha.slice(0, 7)}，树 ${remote.slice(0, 7)} 与本地 HEAD 一致`);
  process.exit(0);
})().catch((e) => {
  console.error(`失败：${e.message}`);
  process.exit(1);
});
