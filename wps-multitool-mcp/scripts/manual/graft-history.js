#!/usr/bin/env node
/**
 * Input: 本地仓库（浅克隆）+ .git/shallow 边界
 * Output: 把本地历史重接成"从根开始的完整历史"，行数不变、内容（tree）不变
 * Pos: 一次性维修工具（维护者用）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 要解决什么
 * ----------
 * 本仓库是浅克隆出来的：最早的提交对象里记着一个【本地不存在】的父提交。
 * 平时 git log 看着没问题（shallow 会假装那条边不存在），但一旦 push：
 *   error: did not receive expected object ...
 * GitHub 校验包里要求祖先齐全，所以这个仓库永远推不上去，只能靠快照提交覆盖。
 *
 * 做法
 * ----
 * 逐个提交重放：同样的 tree、同样的作者/时间/信息，父链重接；边界提交的父级直接去掉。
 * 于是整条链变成"从根开始"的真实历史 —— 能被普通 git push 接受。
 *
 * 代价（要知情）：所有 SHA 都会变。远端已有历史与本地无共同祖先，推的时候必须强推。
 * 所以本脚本只重写【本地】，推送是另一件事（先建远端备份分支再推）。
 *
 * 安全检查
 * --------
 *   - 跑之前先 bundle 备份整库，备份失败就不动
 *   - 重写后逐条校验：提交数一致、新 tip 的 tree == 旧 tip 的 tree
 *   - 任何一条不满足就报错退出，不 update-ref
 *
 * 用法：node scripts/manual/graft-history.js [--apply]
 *   不带 --apply 只做演练（打印新旧映射，不改 ref）
 */

const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const REPO_ROOT = path.resolve(ROOT, '..');
const ADVANCE = process.argv.includes('--apply');

const git = (args, opts = {}) =>
  execFileSync('git', args, {
    cwd: REPO_ROOT,
    encoding: opts.raw ? 'buffer' : 'utf8',
    maxBuffer: 1 << 28,
    stdio: ['pipe', 'pipe', 'pipe'],
    ...opts,
  });

function gitOut(args) {
  return String(git(args)).trim();
}

// ---- 1. 前置检查 -----------------------------------------------------------
const shallowFile = path.join(REPO_ROOT, '.git', 'shallow');
const isShallow = gitOut(['rev-parse', '--is-shallow-repository']) === 'true';
const headRef = gitOut(['symbolic-ref', '--short', 'HEAD']);
const oldTip = gitOut(['rev-parse', 'HEAD']);
const oldTree = gitOut(['rev-parse', 'HEAD^{tree}']);

console.log(`仓库：${REPO_ROOT}`);
console.log(`分支：${headRef}  HEAD：${oldTip.slice(0, 7)}  tree：${oldTree.slice(0, 7)}`);
console.log(`浅克隆：${isShallow ? '是（需要重接）' : '否（无需重接）'}`);

if (!isShallow) {
  console.log('\n本来就是完整历史，什么都不用做。');
  process.exit(0);
}

const boundaries = fs
  .readFileSync(shallowFile, 'utf8')
  .split('\n')
  .map((s) => s.trim())
  .filter(Boolean);
console.log(`shallow 边界 ${boundaries.length} 个：${boundaries.map((b) => b.slice(0, 7)).join(', ')}`);

const commits = gitOut(['rev-list', '--reverse', 'HEAD']).split('\n').filter(Boolean);
console.log(`本地提交 ${commits.length} 个（重写后数量必须一致）`);

const hasStaged = gitOut(['diff', '--cached', '--name-only']) !== '';
if (hasStaged) {
  console.error('\n✗ 暂存区有未提交的改动。先提交或撤下，避免把工作区状态和重写搅在一起。');
  process.exit(1);
}

// ---- 2. 备份 ---------------------------------------------------------------
const bundle = path.join(
  os.homedir(),
  `wps-multitool-mcp-pre-graft-${Date.now()}.bundle`
);
git(['bundle', 'create', bundle, '--all']);
const bundleSize = fs.statSync(bundle).size;
console.log(`\n已备份整库：${bundle}（${Math.round(bundleSize / 1024)} KB）`);
console.log('回滚方式：git clone ' + JSON.stringify(bundle) + ' <目录>');

// ---- 3. 逐条重放 -----------------------------------------------------------
/** 解析原始 commit 对象：tree / parents / author / committer / message */
function parseCommit(sha) {
  const raw = git(['cat-file', 'commit', sha], { raw: true }).toString('utf8');
  const split = raw.indexOf('\n\n');
  const header = split >= 0 ? raw.slice(0, split) : raw;
  const message = split >= 0 ? raw.slice(split + 2) : '';
  const parents = [];
  const info = { tree: '', author: '', committer: '', encoding: '' };
  for (const line of header.split('\n')) {
    if (line.startsWith('tree ')) info.tree = line.slice(5).trim();
    else if (line.startsWith('parent ')) parents.push(line.slice(7).trim());
    else if (line.startsWith('author ')) info.author = line.slice(7);
    else if (line.startsWith('committer ')) info.committer = line.slice(10);
    else if (line.startsWith('encoding ')) info.encoding = line.slice(9);
  }
  return { ...info, parents, message };
}

/** 把 "名字 <邮箱> 时间戳 时区" 拆成 GIT_*_NAME/EMAIL/DATE */
function identityEnv(committer) {
  const m = committer.match(/^(.*) <(.*)> (\d+ [+-]\d{4})$/);
  if (!m) return {};
  return {
    name: m[1],
    email: m[2],
    date: `${m[3].replace(' ', ' ')}`, // git 认 "<unix> <tz>" 形式
  };
}

const inHistory = new Set(commits);
const map = new Map();
let skippedParentEdges = 0;

for (const sha of commits) {
  const c = parseCommit(sha);
  // 父提交不在本地历史里 = 浅克隆边界那条"看不见的边"，直接去掉
  const parents = [];
  for (const p of c.parents) {
    if (inHistory.has(p)) parents.push(map.get(p));
    else skippedParentEdges++;
  }

  const authorEnv = identityEnv(c.author);
  const committerEnv = identityEnv(c.committer);
  const args = ['commit-tree', c.tree];
  for (const p of parents) args.push('-p', p);

  const env = {
    ...process.env,
    GIT_AUTHOR_NAME: authorEnv.name,
    GIT_AUTHOR_EMAIL: authorEnv.email,
    GIT_AUTHOR_DATE: authorEnv.date,
    GIT_COMMITTER_NAME: committerEnv.name,
    GIT_COMMITTER_EMAIL: committerEnv.email,
    GIT_COMMITTER_DATE: committerEnv.date,
  };
  const newSha = execFileSync('git', args, {
    cwd: REPO_ROOT,
    input: c.message,
    env,
    encoding: 'utf8',
    stdio: ['pipe', 'pipe', 'pipe'],
  }).trim();
  map.set(sha, newSha);
}

const newTip = map.get(oldTip);
const newTree = gitOut(['rev-parse', `${newTip}^{tree}`]);

// ---- 4. 自检 ---------------------------------------------------------------
console.log(`\n重写完：新 tip ${newTip.slice(0, 7)}，tree ${newTree.slice(0, 7)}`);
console.log(`去掉的边界父边 ${skippedParentEdges} 条（正常应为 ${boundaries.length}）`);

if (newTree !== oldTree) {
  console.error(`✗ tree 不一致：${newTree} ≠ ${oldTree}。内容被改动了，拒绝 update-ref。`);
  process.exit(1);
}
const newCount = gitOut(['rev-list', '--count', newTip]);
if (Number(newCount) !== commits.length) {
  console.error(`✗ 提交数变了：${newCount} ≠ ${commits.length}`);
  process.exit(1);
}

const sample = commits.slice(-3).map((s) => `${s.slice(0, 7)} → ${map.get(s).slice(0, 7)}`);
console.log('抽样映射（最后 3 条）：');
for (const s of sample) console.log('  ' + s);

if (!ADVANCE) {
  console.log('\n演练模式：没有改动任何 ref。确认无误后加 --apply 真正重接。');
  process.exit(0);
}

// ---- 5. 落盘 ---------------------------------------------------------------
git(['update-ref', `refs/heads/${headRef}`, newTip, oldTip]);
console.log(`\n✓ refs/heads/${headRef} 已指向 ${newTip.slice(0, 7)}`);

// shallow 文件里如果全是已重接的边界，就该把它删掉，否则 git 仍认为仓库是浅的
const remaining = boundaries.filter((b) => inHistory.has(b));
if (remaining.length === boundaries.length) {
  fs.unlinkSync(shallowFile);
  console.log('✓ 已移除 .git/shallow（现在是一个完整仓库，普通 git push 可用）');
}
console.log(`\n下一步：建远端备份分支后强推（见 docs/HANDOFF.md 的「历史 graft」一节）。`);
