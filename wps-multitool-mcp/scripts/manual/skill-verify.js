/**
 * skill 包校验
 * -----------
 * 按 Agent Skills 规范（agentskills.io/specification）逐条查，不是随便看看：
 *   1) name 必须等于父目录名，且符合命名规则（规范硬要求，错���客户端不认）
 *   2) description ≤1024 字符 —— 上限是仓库里一条提交踩出来的（1176 → 992）
 *   3) SKILL.md <500 行 —— 超了模型一次读不完，渐进披露就失效了
 *   4) 引用文件必须存在，且【只一层深】（规范明确要求）
 *   5) 目录名用 references/（复数），别用 reference/
 *   6) 【第 10 轮】内容真伪：SKILL.md 说的分组验证结论必须与 tool-verification.json 对得上
 *   7) 【第 10 轮】已装副本同步（只警告不 FAIL，可用 --no-sync-check 跳过）
 *
 * 为什么第 6 条必须有：SKILL.md 是 **Agent 读来决定怎么用这个 MCP 的**，
 * 它说错话就是 Agent 对用户说错话。第 9 轮 PPT 隔离验证做到 98/98 之后，
 * SKILL.md 里"PPT 没真机覆盖验证过……整个 ppt 组都是 static"已经是假的，
 * 而当时 20 条断言全绿 —— 说明原来的门禁一个字的内容真伪都不验。
 *
 * 用法：
 *   node scripts/manual/skill-verify.js [skill目录] [--no-sync-check]
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const truth = require('./skill-content-truth.js');

const ROOT = path.resolve(__dirname, '..', '..');
const argv = process.argv.slice(2).filter((a) => a !== '--no-sync-check');
const NO_SYNC_CHECK = process.argv.includes('--no-sync-check');
const skillDir = argv[0] ? path.resolve(argv[0]) : path.join(ROOT, 'skills', 'wps-multitool');

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

console.log(`\n校验 ${path.relative(ROOT, skillDir) || skillDir}`);

const skillPath = path.join(skillDir, 'SKILL.md');
if (!fs.existsSync(skillPath)) {
  console.log(`  ✗ 找不到 SKILL.md`);
  process.exit(1);
}
const raw = fs.readFileSync(skillPath, 'utf8');

// ---- frontmatter 解析（够用即可，不引 yaml 依赖） ----
const fmMatch = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n/);
ok('有 YAML frontmatter', !!fmMatch);
if (!fmMatch) process.exit(1);
const fm = fmMatch[1];

const getField = (key) => {
  const m = fm.match(new RegExp(`^${key}:\\s*([\\s\\S]*?)(?=\\n[a-z_]+:|$)`, 'm'));
  return m ? m[1].trim() : '';
};

const name = getField('name');
const description = getField('description');
const compatibility = getField('compatibility');

// ---- 1. name ----
const dirName = path.basename(skillDir);
ok('name 等于父目录名', name === dirName, `name=${name} dir=${dirName}`);
ok('name 长度 ≤64', name.length <= 64, `${name.length}`);
ok('name 只含小写字母/数字/连字符', /^[a-z0-9-]+$/.test(name));
ok('name 不以连字符开头结尾', !name.startsWith('-') && !name.endsWith('-'));
ok('name 无连续连字符', !name.includes('--'));
ok('name 不含保留词', !/claude|anthropic/i.test(name));

// ---- 2. description ----
ok('description 非空', description.length > 0);
ok('description ≤1024 字符', description.length <= 1024, `${description.length}`);
ok('description 写了"什么时候用"', /当|使用|适合|不要使用|需要/.test(description));

// ---- 3. compatibility ----
if (compatibility) {
  ok('compatibility ≤500 字符', compatibility.length <= 500, `${compatibility.length}`);
}

// ---- 4. 体量 ----
const lines = raw.split(/\r?\n/).length;
ok('SKILL.md <500 行', lines < 500, `${lines} 行`);
ok('SKILL.md 未超 5000 token 量级', raw.length < 20000, `${raw.length} 字符`);

// ---- 5. 目录约定 ----
ok('用 references/（复数）', fs.existsSync(path.join(skillDir, 'references')));
ok('没有 reference/（单数，规范名）', !fs.existsSync(path.join(skillDir, 'reference')));

// ---- 6. 引用完整性 + 只有一层深 ----
// 只有 references/（规范明确要求的那一层）算"技能包内引用"。
// 正文里也可能出现【仓库路径】——第 10 轮起 SKILL.md 要指明真值来源
// （scripts/manual/tool-verification.json），那是给人看的出处、不是包内引用。
// 以前一律当引用查，结果把 scripts/manual 报成"引用不存在"，是误报。
const body = raw.slice(fmMatch[0].length);
const allPaths = [...new Set([...body.matchAll(/(?:references|scripts|assets)\/[\w./-]+/g)].map((m) => m[0]))];
const uniqRefs = allPaths.filter((p) => p.startsWith('references/'));
ok('SKILL.md 里引用了参考文档', uniqRefs.length > 0, uniqRefs.join(', '));
for (const r of uniqRefs) {
  ok(`引用存在：${r}`, fs.existsSync(path.join(skillDir, r)));
}
const repoCitations = allPaths.filter((p) => !p.startsWith('references/'));
if (repoCitations.length) {
  console.log(`  – 另引用了 ${repoCitations.length} 处仓库路径（出处说明，不计入引用检查）：${repoCitations.join(', ')}`);
}

// 参考文档内部不允许再互相引用（规范：只一层深）
let nested = 0;
const refDir = path.join(skillDir, 'references');
if (fs.existsSync(refDir)) {
  for (const f of fs.readdirSync(refDir)) {
    const txt = fs.readFileSync(path.join(refDir, f), 'utf8');
    const inner = [...txt.matchAll(/references\/[\w.-]+/g)];
    if (inner.length) {
      nested += inner.length;
      console.log(`  ! ${f} 里还有 ${inner.length} 处 references/ 引用`);
    }
  }
}
ok('参考文档没有嵌套引用（只一层深）', nested === 0);

const refFiles = fs.existsSync(refDir) ? fs.readdirSync(refDir) : [];
console.log(`\n  参考文档 ${refFiles.length} 份：${refFiles.join(', ')}`);

// ==================== 7. 内容真伪（第 10 轮 T1）====================
// SKILL.md 是 Agent 的行为依据。它说「PPT 没验证过」，Agent 就会对用户说这句话。
// 所以这里不是查格式，是拿验证台账的真值去对账。
console.log(`\n内容真伪（对账 scripts/manual/tool-verification.json）：`);
const ledger = truth.loadLedger();
if (ledger.error) {
  ok('验证台账可读', false, ledger.error);
} else {
  ok('验证台账可读', true);

  // 7a. 已知失效断言黑名单
  const docsToScan = [
    { name: 'SKILL.md', text: raw },
    ...(fs.existsSync(refDir)
      ? fs.readdirSync(refDir).map((f) => ({
          name: `references/${f}`,
          text: fs.readFileSync(path.join(refDir, f), 'utf8'),
        }))
      : []),
  ];
  for (const c of truth.RETIRED_CLAIMS) {
    const hit = docsToScan.filter((d) => c.re.test(d.text));
    ok(
      `无失效断言：${c.label}`,
      hit.length === 0,
      hit.length ? `命中 ${hit.map((h) => h.name).join(', ')}（已被实测证伪：${c.why}）` : ''
    );
  }

  // 7b. 分组验证结论逐句对账（主力口径：不靠黑名单）
  const groups = Object.keys(ledger.groups || {});
  const claims = truth.extractGroupClaims(
    docsToScan.map((d) => d.text).join('\n'),
    groups
  );
  const contradictions = [];
  for (const c of claims) {
    const bad = truth.contradictionOf(c, ledger);
    if (bad) contradictions.push(bad);
  }
  ok(
    '分组验证结论与台账一致（逐句）',
    contradictions.length === 0,
    contradictions.length ? contradictions.slice(0, 3).join(' | ') : `对账 ${claims.length} 句`
  );

  // 7c. 文档里引用的数字必须与台账 evidence 同源（防止数字各说各话）
  const ev = ledger.evidence && ledger.evidence.ppt;
  if (ev) {
    const n = truth.numbersFromEvidence(ev);
    const has = (v) => new RegExp(`\\b${v}\\b`).test(raw);
    const livePptCount = Object.entries(ledger.tools || {}).filter(
      ([name, st]) => st === 'live' && name.includes('_ppt_')
    ).length;
    ok('SKILL.md 引用的「矩阵行数」与台账一致', n.matrixRows !== undefined && has(n.matrixRows), `台账=${n.matrixRows}`);
    ok('SKILL.md 引用的「执行条数」与台账一致', n.executed !== undefined && has(n.executed), `台账=${n.executed}`);
    ok('SKILL.md 引用的「跳过条数」与台账一致', n.skipped !== undefined && has(n.skipped), `台账=${n.skipped}`);
    ok('SKILL.md 引用的「PPT 组工具数」与台账一致', n.groupTools !== undefined && has(n.groupTools), `台账=${n.groupTools}`);
    ok(
      'SKILL.md 引用的「逐条 live 数」与台账 tools 一致',
      n.liveTools !== undefined && has(n.liveTools) && livePptCount === n.liveTools,
      `台账 tools 实算=${livePptCount} evidence=${n.liveTools}`
    );
  } else {
    ok('台账里有 ppt 证据条目', false, 'evidence.ppt 缺失，无法对账数字');
  }
}

// ==================== 8. 已装副本同步（第 10 轮 T2）====================
// 只警告不 FAIL：副本可能被用户自己改过，或机器上根本没装。
// 但不同步意味着 Agent 读的是旧的（这次 PPT 那句话就是从副本流出去的）。
console.log(`\n已装副本同步检查：`);
if (NO_SYNC_CHECK) {
  console.log('  – 已跳过（--no-sync-check）');
} else {
  const repoMd5 = crypto.createHash('md5').update(raw).digest('hex').toUpperCase();
  console.log(`  仓库版 md5=${repoMd5}`);
  const homes = [process.env.USERPROFILE, process.env.HOME].filter(Boolean);
  const candidates = [];
  for (const h of homes) {
    candidates.push(path.join(h, '.workbuddy', 'skills', 'wps-multitool', 'SKILL.md'));
    candidates.push(path.join(h, '.codebuddy', 'skills', 'wps-multitool', 'SKILL.md'));
  }
  let found = 0;
  for (const p of [...new Set(candidates)]) {
    if (!fs.existsSync(p)) continue;
    found++;
    const installed = crypto.createHash('md5').update(fs.readFileSync(p)).digest('hex').toUpperCase();
    const pretty = path.dirname(p).replace(homes[0] || '', '~');
    if (installed === repoMd5) {
      console.log(`  ✓ ${pretty} — md5 一致`);
    } else {
      console.log(`  ! ${pretty} — 仓库版与已装副本不一致，请同步（仓库=${repoMd5} 副本=${installed}）`);
      console.log('      同步命令：node scripts/manual/install-verify.js --print  或手动复制');
    }
  }
  if (!found) console.log('  – 本机没有已装副本（CI 上属正常）');
}

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);