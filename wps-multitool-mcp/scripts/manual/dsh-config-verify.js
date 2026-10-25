/**
 * Input: 本仓库的工具定义（src/tools/**）与 dsh 配置片段（scripts/manual/fixtures/cordis.wps.patch.yml）
 * Output: 一份静态校验报告；有 FAIL 则 exit 1
 * Pos: dsh 适配的门禁。改工具名 / 改 dsh 片段 / 改桥超时之前先跑它。
 *
 * 为什么需要这道门禁
 * ------------------
 * dsh 对工具名是【合成名】：mcp__<serverName>__<rawName>，字符集 [A-Za-z0-9_-]、总长 ≤64。
 * 这几个约束在服务端完全看不出来 —— 我们的工具名多一个点、长一个字符，
 * 在 CodeBuddy/Claude 上照跑，在 dsh 上整份列表被判无效（"沿用上一份"，即静默丢工具）。
 * 这正是本项目最怕的那类失败：不报错、只是工具没了。
 *
 * 校验的是【静态定义】，不是运行时 tools/list。运行时的权威名单靠真机 dsh 看，
 * 这里只保证"不会因为命名/超时这种硬约束翻车"。
 *
 * 片段原先放在 docs/（gitignored）——2026-10-25 Linux 云端测试抓出：任何干净检出
 * （CI、Linux 源码测试）这门禁必红。已迁入 fixtures/（进 git），docs/ 里的旧副本仅作本地参考。
 *
 * 用法：node scripts/manual/dsh-config-verify.js
 */

const fs = require('fs');
const path = require('path');

/** 包根：src/ 与 dist/ 所在处 */
const ROOT = path.resolve(__dirname, '..', '..');
/** 仓库根：dsh-bundle/ 在这里（包根是仓库的子目录，别算错一层） */
const REPO_ROOT = path.resolve(ROOT, '..');
const TOOLS_DIR = path.join(ROOT, 'src', 'tools');
/** dsh 配置样例：迁入 fixtures/（进 git）——原 docs/ 路径被 gitignore，干净检出必红（Linux 云端测试发现） */
const PATCH = path.join(__dirname, 'fixtures', 'cordis.wps.patch.yml');

const DSH_NAME_MAX = 64;
const SERVER_NAME_RE = /^[A-Za-z0-9_-]{1,32}$/;
const TOOL_NAME_RE = /^[A-Za-z0-9_-]+$/;
/** dsh 会把这些环境变量从子进程环境里剔掉 */
const ENV_STRIP_RE = /KEY|PASSWORD|SECRET|TOKEN/i;
/** 桥的默认单动作超时（ps-child.ts / ps-daemon.ts 都是 120000） */
const BRIDGE_TIMEOUT_DEFAULT = 120000;

let failed = 0;
function ok(label, pass, detail) {
  console.log(`${pass ? '[OK]  ' : '[FAIL]'} ${label}${detail ? ' — ' + detail : ''}`);
  if (!pass) failed++;
}

function walk(dir, out = []) {
  for (const f of fs.readdirSync(dir)) {
    const p = path.join(dir, f);
    const st = fs.statSync(p);
    if (st.isDirectory()) walk(p, out);
    else if (f.endsWith('.ts') && !f.endsWith('.d.ts')) out.push(p);
  }
  return out;
}

// ---------- A. 工具名硬约束 ----------
const files = fs.existsSync(TOOLS_DIR) ? walk(TOOLS_DIR) : [];
const names = new Map();
for (const f of files) {
  const src = fs.readFileSync(f, 'utf8');
  for (const m of src.matchAll(/\bname:\s*'([A-Za-z0-9_.:-]+)'/g)) {
    if (!names.has(m[1])) names.set(m[1], path.relative(ROOT, f));
  }
}
ok('工具定义能被扫到', names.size > 0, `${names.size} 个（静态扫 name: 字段，运行时为准）`);

const bad = [...names.keys()].filter((n) => !TOOL_NAME_RE.test(n));
ok('工具名只用 [A-Za-z0-9_-]', bad.length === 0, bad.slice(0, 5).join(', '));

const longest = [...names.keys()].sort((a, b) => b.length - a.length).slice(0, 3);

// ---------- B. dsh 片段约束 ----------
if (!fs.existsSync(PATCH)) {
  ok('dsh 配置片段存在', false, PATCH);
} else {
  const yml = fs.readFileSync(PATCH, 'utf8');
  const grab = (re) => (yml.match(re) || [])[1];

  const serverName = grab(/serverName:\s*([A-Za-z0-9_-]+)/);
  ok('serverName 合法（[A-Za-z0-9_-]{1,32}）', !!serverName && SERVER_NAME_RE.test(serverName), serverName);

  const prefix = `mcp__${serverName || '?'}__`;
  const overLong = [...names.keys()].filter((n) => prefix.length + n.length > DSH_NAME_MAX);
  ok(
    `合成名 mcp__<server>__<tool> ≤ ${DSH_NAME_MAX} 字符`,
    overLong.length === 0,
    overLong.length ? overLong.slice(0, 5).join(', ') : `最长 ${prefix.length + (longest[0] || '').length}（${longest[0] || '-'}）`
  );

  const cmd = grab(/command:\s*(.+)/);
  ok(
    '启动命令不是 npx / npx.cmd',
    !!cmd && !/npx/i.test(cmd),
    `${(cmd || '').trim()}（dsh 透传参数不过 shell，Windows 上 npx 是 .cmd，会 ENOENT）`
  );

  ok('transport 是 stdio', /transport:\s*stdio/.test(yml), '本项目只有 stdio');

  const timeout = parseInt(grab(/toolCallTimeoutMs:\s*(\d+)/) || '0', 10);
  ok(
    `toolCallTimeoutMs ≥ 桥默认超时 ${BRIDGE_TIMEOUT_DEFAULT}`,
    timeout >= BRIDGE_TIMEOUT_DEFAULT,
    `当前 ${timeout || '未设（dsh 默认 60000，会先掐断）'}`
  );

  ok('failOnStartupError: true', /failOnStartupError:\s*true/.test(yml), '默认 false 会静默降级成零工具');

  const envBlock = (yml.match(/env:\s*\n((?:\s{10,}.*\n?)*)/) || [])[1] || '';
  const envKeys = [...envBlock.matchAll(/^\s+([A-Z_][A-Z0-9_]*):/gm)].map((m) => m[1]);
  const stripped = envKeys.filter((k) => ENV_STRIP_RE.test(k));
  ok(
    'env 里没有会被 dsh 剔除的键',
    stripped.length === 0,
    stripped.length ? stripped.join(', ') : envKeys.join(', ') || '（无）'
  );
}

// ---------- C. 生成器实测（install --client dsh）----------
// 光校验静态片段不够：真正会被用户用的是 install 生成出来的那一份。
// 这里直接跑一遍生成器，确认"生成物"而不是"样例"是合规的。
const { spawnSync } = require('child_process');
const distEntry = path.join(ROOT, 'dist', 'index.js');
if (fs.existsSync(distEntry)) {
  const r = spawnSync(process.execPath, [distEntry, 'install', '--client', 'dsh', '--print'], {
    encoding: 'utf8',
  });
  const out = r.stdout || '';
  ok('install --client dsh 能生成片段', r.status === 0 && out.includes('serverName: wps'));
  ok('生成物里超时 ≥ 桥默认值', /toolCallTimeoutMs:\s*(\d+)/.test(out) && parseInt(out.match(/toolCallTimeoutMs:\s*(\d+)/)[1], 10) >= BRIDGE_TIMEOUT_DEFAULT);
  const argsLine = (out.match(/^\s+- '([^']*index\.js)'/m) || [])[1];
  ok('生成物里指向当前构建产物', !!argsLine && fs.existsSync(argsLine), argsLine || '没解析到 args');
  // 只看 command 那一行：注释里出现 "不能用 npx" 是解释，不是违规
  const genCmd = (out.match(/^\s+command:\s*(.+)$/m) || [])[1] || '';
  ok('生成物的 command 不是 npx', !!genCmd && !/npx/i.test(genCmd), genCmd.trim());
} else {
  ok('dist 已构建（生成器实测需要它）', false, '先 npm run build');
}

// ---------- C2. bundle 结构（dsh-bundle/）----------
// 这一层防的是"包结构写错，装上去不被识别成 bundle"——那种失败 dsh 只打一行警告，
// 用户看到的是"装了没用"。
//
// v0.4 起 bundle **没有 index.js**：挂载改由 dsh 的 loader 按包名解析
// （我们自己的代码 import 不到 app.asar 里的官方客户端，v0.3 真机失败过）。
// 所以这里改为检查"两个文件 + assets 技能"的结构。
const BUNDLE = path.join(REPO_ROOT, 'dsh-bundle');
if (!fs.existsSync(BUNDLE)) {
  console.log('[SKIP] 没有 dsh-bundle/（bundle 尚未启用）');
} else {
  const pkgPath = path.join(BUNDLE, 'package.json');
  let pkg = null;
  try {
    pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
  } catch (e) {
    ok('dsh-bundle/package.json 能解析', false, e.message);
  }
  if (pkg) {
    ok('bundle 声明了 dsh.bundle.patch', !!pkg.dsh?.bundle?.patch, pkg.dsh?.bundle?.patch || '缺');
    const patchFile = path.join(BUNDLE, String(pkg.dsh?.bundle?.patch || 'cordis.patch.yml'));
    ok('patch 文件存在', fs.existsSync(patchFile), path.relative(REPO_ROOT, patchFile));
    ok('声明了 engines.dsh 版本范围', !!pkg.engines?.dsh, pkg.engines?.dsh || '缺');
    ok('不再依赖 index.js（v0.4 起由 loader 解析）', !fs.existsSync(path.join(BUNDLE, 'index.js')));
    ok('files 里含 assets（技能随包分发）', (pkg.files || []).includes('assets'), JSON.stringify(pkg.files || []));

    if (fs.existsSync(patchFile)) {
      const yml = fs.readFileSync(patchFile, 'utf8');
      // 关键：patch 里不许出现【字面】绝对路径（发布时不可能知道用户机器）。
      // 路径一律由 !!js 在求值时算 —— 所以先剥掉 !!js 表达式再查。
      const withoutJs = yml
        .replace(/^#.*$/gm, '')
        .replace(/!!js\s+>-[\s\S]*?(?=\n\s*[a-zA-Z#-]|\n\s*$)/g, '!!js <expr>');
      const absPath = /([A-Za-z]:\\|\/Users\/|\/home\/)/.test(withoutJs);
      ok('patch 里没有字面绝对路径', !absPath, absPath ? '路径要写成 !!js 表达式' : '路径都由 !!js 算');

      // 三条关键行必须都在
      ok('挂了官方 MCP 客户端', yml.includes("name: '@deepseek-ai/dsh-mcp-client'"));
      ok('启用了 skill-filesystem（带 customSkillDirs）', /- id: skill-filesystem/.test(yml) && /customSkillDirs/.test(yml));
      ok('启用了 tool-skill', /- id: tool-skill/.test(yml));

      const cfgTimeout = (yml.match(/toolCallTimeoutMs:\s*(\d+)/) || [])[1];
      ok(
        `bundle 传给官方插件的超时 ≥ ${BRIDGE_TIMEOUT_DEFAULT}`,
        !!cfgTimeout && parseInt(cfgTimeout, 10) >= BRIDGE_TIMEOUT_DEFAULT,
        cfgTimeout
      );
      ok('failOnStartupError: true（不静默降级成零工具）', /failOnStartupError:\s*true/.test(yml));
      ok('command 是 node 不是 npx', /command:\s*node\b/.test(yml) && !/command:\s*npx/.test(yml));

      // serverName 决定工具名前缀 mcp__<serverName>__<工具名>。
      // npm 上已有 dsh-wps@0.1.2（WPS 云文档集成），它的前缀是 mcp__wps__*。
      // 我们必须用别的名字，否则两个包装在一起时 dsh 会报 serverName 已被占用。
      const sn = (yml.match(/serverName:\s*([A-Za-z0-9_-]+)/) || [])[1];
      ok('serverName 存在且合 dsh 规则', !!sn && /^[A-Za-z0-9_-]{1,32}$/.test(sn), sn || '缺');
      ok(
        'serverName 不是 wps（避开 npm 上 dsh-wps 的 mcp__wps__* 前缀）',
        sn !== 'wps',
        sn === 'wps' ? '会和别人的 dsh-wps 撞名' : `用的是 ${sn}`
      );
    }

    // 技能资产
    const skillDir = path.join(BUNDLE, 'assets', 'wps-multitool');
    const skillMd = path.join(skillDir, 'SKILL.md');
    ok('assets/wps-multitool/SKILL.md 存在', fs.existsSync(skillMd));
    if (fs.existsSync(skillMd)) {
      const src = fs.readFileSync(skillMd, 'utf8');
      const fm = /^---\r?\n([\s\S]*?)\r?\n---/.exec(src);
      const nameM = fm ? /^name:\s*(.+)$/m.exec(fm[1]) : null;
      const descM = fm ? /^description:\s*(.+)$/m.exec(fm[1]) : null;
      const sname = nameM?.[1]?.trim();
      ok('SKILL.md 的 name 合 dsh 的 kebab-case 规则', !!sname && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(sname), sname || '缺');
      ok('SKILL.md 有 description（dsh 必填）', !!descM?.[1]?.trim());
    }
    ok('自检脚本存在（node dsh-bundle/selftest.js）', fs.existsSync(path.join(BUNDLE, 'selftest.js')));
  }
}

// ---------- D. 交叉一致性 ----------
ok(
  '桥超时与 dsh 超时的口径已绑在一起',
  /toolCallTimeoutMs/.test(fs.existsSync(PATCH) ? fs.readFileSync(PATCH, 'utf8') : '') &&
    /WPS_DAEMON_TIMEOUT/.test(fs.readFileSync(path.join(ROOT, 'src/client/ps-daemon.ts'), 'utf8')),
  '改任一处都要重跑本脚本（另一处在 src/client/ps-daemon.ts:47）'
);

console.log('');
if (failed) {
  console.log(`${failed} 项 FAIL —— dsh 侧会出现"工具静默消失"，先修再发`);
  process.exit(1);
}
console.log('全部通过。运行时仍需真机确认：mcp__wpslocal__* 出现在 dsh 工具列表里。');
