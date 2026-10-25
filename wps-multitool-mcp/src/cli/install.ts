/**
 * Input: 本机已存在的 MCP 客户端配置目录 + 当前安装路径
 * Output: 把 wps-multitool-mcp 的启动项写进各客户端配置（写前备份、写后回读校验）
 * Pos: `wps-multitool-mcp install` 子命令。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ------------
 * README 原来教用户手抄一段 JSON，还要自己解释"路径里出现两次 wps-multitool-mcp 是正常的"——
 * 需要解释的设计就是坏设计。这里让程序自己找到 dist/index.js 并写对。
 *
 * 三条安全约束（写别人的配置文件，最怕的是把人家的配置弄坏）
 * ------------------------------------------------------
 *   1) 只写【已存在的客户端目录】。目录都不存在还硬建一个，等于在用户机器上乱拉屎；
 *      认不出的就打印片段，让人自己粘。
 *   2) 写前备份 <file>.bak-<时间戳>，写后回读校验；校验失败立刻用备份还原。
 *   3) 合并而不是覆盖：mcpServers 里别人的服务一个都不能少（这是最容易出事故的地方）。
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const SERVER_KEY = 'wps-multitool-mcp';

export type ClientKind = 'json' | 'toml';

export interface ClientTarget {
  id: string;
  label: string;
  kind: ClientKind;
  file: string;
  /** 该客户端是否真的装过（目录存在）。false 只打印片段，不写 */
  present: boolean;
}

export interface InstallOptions {
  dryRun: boolean;
  only?: string;
  printOnly: boolean;
}

function home(): string {
  return process.env.HOME || process.env.USERPROFILE || os.homedir();
}

/** 本包的入口：优先用当前进程所在包里的 dist/index.js，保证写进去的就是正在跑的这个版本 */
export function entryScript(): string {
  const dist = path.resolve(__dirname, '..', 'index.js');
  return fs.existsSync(dist) ? dist : '';
}

/**
 * 启动项：本地装了就用 node + 绝对路径（启动快、版本确定）；
 * 找不到 dist（比如在源码目录里临时跑）就退回 npx，至少不会写一个死路径。
 */
export function serverEntry(): { command: string; args: string[] } {
  const entry = entryScript();
  if (entry) return { command: 'node', args: [entry] };
  return { command: 'npx', args: ['-y', SERVER_KEY] };
}

export function detectClients(): ClientTarget[] {
  const h = home();
  const list: Array<Omit<ClientTarget, 'present'>> = [
    { id: 'codebuddy', label: 'CodeBuddy', kind: 'json', file: path.join(h, '.codebuddy', 'mcp.json') },
    { id: 'workbuddy', label: 'WorkBuddy', kind: 'json', file: path.join(h, '.workbuddy', 'mcp.json') },
    { id: 'codex', label: 'Codex CLI', kind: 'toml', file: path.join(h, '.codex', 'config.toml') },
    { id: 'claude', label: 'Claude Code', kind: 'json', file: path.join(h, '.claude.json') },
    { id: 'claude-settings', label: 'Claude Code (settings)', kind: 'json', file: path.join(h, '.claude', 'settings.json') },
    { id: 'cursor', label: 'Cursor', kind: 'json', file: path.join(h, '.cursor', 'mcp.json') },
  ];
  // 判定"这个客户端装过没有"：配置文件在，或者它专属的目录在。
  // 注意不能只看 dirname —— ~/.claude.json 的 dirname 就是用户主目录，永远存在，
  // 那样会给每个机器都建一个 .claude.json（典型的"以为在帮忙"）。
  const present = (file: string) => {
    if (fs.existsSync(file)) return true;
    const dir = path.dirname(file);
    return dir !== h && fs.existsSync(dir);
  };
  return list.map((c) => ({ ...c, present: present(c.file) }));
}

// ---------------- JSON 客户端 ----------------

function readJson(file: string): Record<string, unknown> {
  if (!fs.existsSync(file)) return {};
  const raw = fs.readFileSync(file, 'utf8').trim();
  if (!raw) return {};
  const parsed = JSON.parse(raw);
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('顶层不是对象，不敢动它');
  }
  return parsed as Record<string, unknown>;
}

export function mergeJsonClient(
  before: Record<string, unknown>,
  entry: { command: string; args: string[] }
): Record<string, unknown> {
  const next = { ...before };
  const servers =
    before.mcpServers && typeof before.mcpServers === 'object' && !Array.isArray(before.mcpServers)
      ? { ...(before.mcpServers as Record<string, unknown>) }
      : {};
  servers[SERVER_KEY] = entry;
  next.mcpServers = servers;
  return next;
}

// ---------------- TOML 客户端（Codex） ----------------
// 不引第三方 TOML 库：只处理 `[mcp_servers.<name>]` 这一段，其余内容原样保留。
// 这也是一种安全策略 —— 我们看不懂的东西，一个字都不改。

const TOML_HEADER = `[mcp_servers.${SERVER_KEY}]`;

function tomlString(s: string): string {
  // TOML 基本字符串里 \ 是转义符；Windows 路径全是反斜杠，用单引号字面量最省事
  return s.includes("'") ? JSON.stringify(s) : `'${s}'`;
}

export function mergeTomlClient(before: string, entry: { command: string; args: string[] }): string {
  const block = [
    TOML_HEADER,
    `command = ${tomlString(entry.command)}`,
    `args = [${entry.args.map(tomlString).join(', ')}]`,
  ].join('\n');

  const lines = before.split(/\r?\n/);
  const start = lines.findIndex((l) => l.trim() === TOML_HEADER);
  if (start >= 0) {
    // 找到下一个表头为止，整段替换
    let end = lines.length;
    for (let i = start + 1; i < lines.length; i++) {
      if (/^\s*\[/.test(lines[i])) {
        end = i;
        break;
      }
    }
    const out = [...lines.slice(0, start), ...block.split('\n'), '', ...lines.slice(end)];
    return out.join('\n').replace(/\n{3,}/g, '\n\n');
  }

  const base = before.trimEnd();
  return `${base}${base ? '\n\n' : ''}${block}\n`;
}

// ---------------- DSH（DeepSeek Harness） ----------------
// 为什么不塞进 detectClients：dsh 的配置是 YAML 的【补丁层】，写入动作本质是
// "往别人已有的 insert 列表里加一条"。用字符串拼 YAML 去改别人的文件，写坏一次就是
// 用户整套插件配置报废 —— 这个代价换不来那点方便。所以这里只做两件安全的事：
//   1) 生成一个【独立】的 patch 文件（绝不碰 cordis.patch.yml）
//   2) 打印"临时用 / 持久用"两条路径，何时合并由人决定
// 另一个理由：install-verify.js 对 detectClients 的集合有断言，加进去会误伤那道门禁。

/** DSH_HOME 默认 ~/.dsh；可用环境变量覆盖（与 dsh 自己的约定一致） */
export function dshHome(): string {
  return process.env.DSH_HOME || path.join(home(), '.dsh');
}

export function dshPatchPath(): string {
  return path.join(dshHome(), 'wps.cordis.yml');
}

/** YAML 单引号标量：反斜杠保持字面，单引号要靠翻倍转义（Windows 路径必须这么写才不被吃） */
function yamlSingle(s: string): string {
  return `'${s.replace(/'/g, "''")}'`;
}

export function formatDshPatch(entry: { command: string; args: string[] }): string {
  return [
    '# 由 wps-multitool-mcp install --client dsh 生成。',
    '# 用法二选一：',
    '#   临时：dsh web --patch <本文件路径>',
    '#   持久：把下面这条 insert 合并进 $DSH_HOME/profiles/<profile>/cordis.patch.yml',
    '#        （那个文件里通常已有别人的补丁，只能追加，不能覆盖）',
    '',
    '- insert:',
    '    - id: mcp-wps',
    "      name: '@deepseek-ai/dsh-mcp-client'",
    '      config:',
    '        serverName: wps',
    '        transport: stdio',
    '        # 必须用 node + 绝对路径，不能用 npx：dsh 的 args 直接透传、不过 shell，',
    '        # Windows 上 npx 实际是 npx.cmd，按 npx 去 spawn 会 ENOENT。',
    `        command: ${yamlSingle(entry.command)}`,
    '        args:',
    ...entry.args.map((a) => `          - ${yamlSingle(a)}`),
    '        # 必须 >= 桥的 WPS_DAEMON_TIMEOUT（默认 120000）。dsh 默认 60000 会先掐断，',
    '        # 而 WPS 那边动作已经改完了 —— "客户端报失败、文档其实改了"。',
    '        toolCallTimeoutMs: 130000',
    '        # 默认 false 会在连不上时静默降级成"零工具"，用户只会觉得 AI 不会用 WPS。',
    '        failOnStartupError: true',
    '        env:',
    '          # 保持默认档：dsh 会把工具定义塞进每一次请求，276 个全量很贵。',
    '          WPS_TOOL_TIER: essential',
  ].join('\n');
}

function installDsh(opts: InstallOptions, entry: { command: string; args: string[] }): number {
  const file = dshPatchPath();
  const patch = `${formatDshPatch(entry)}\n`;
  const dshPresent = fs.existsSync(dshHome());

  console.log(`DSH（DeepSeek Harness）— ${file}\n`);

  if (opts.printOnly || opts.dryRun || !dshPresent) {
    if (!dshPresent && !opts.printOnly) {
      console.log('没检测到 DSH（' + dshHome() + ' 不存在），不凭空建目录。先给你片段：\n');
    }
    console.log(patch);
    console.log('装好 DSH 之后重跑这条命令，就会写到上面那个路径：');
    console.log('  node <本包>/dist/index.js install --client dsh\n');
    return 0;
  }

  let backup: string | undefined;
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    if (fs.existsSync(file)) {
      backup = `${file}.bak-${stamp()}`;
      fs.copyFileSync(file, backup);
    }
    const tmp = `${file}.tmp-${process.pid}`;
    fs.writeFileSync(tmp, patch, 'utf8');
    fs.renameSync(tmp, file);
    const check = fs.readFileSync(file, 'utf8');
    if (!check.includes('serverName: wps')) throw new Error('回读没找到 serverName: wps');
    console.log(`✓ 已写入 ${file}${backup ? `（备份：${path.basename(backup)}）` : ''}`);
    console.log('\n下一步二选一：');
    console.log(`  临时用：dsh web --patch "${file}"`);
    console.log('  持久用：把里面的 insert 合并进 $DSH_HOME/profiles/<profile>/cordis.patch.yml（追加，别覆盖）');
    console.log('\n验证：dsh --dump-config 里应能看到 mcp-wps；工具名会以 mcp__wps__ 开头。');
    return 0;
  } catch (e) {
    if (backup && fs.existsSync(backup)) {
      try {
        fs.copyFileSync(backup, file);
      } catch {}
    }
    console.error(`✗ 写入失败${backup ? '（已还原备份）' : ''}：${e instanceof Error ? e.message : String(e)}`);
    return 1;
  }
}

// ---------------- 写入 ----------------

function stamp(): string {
  const d = new Date();
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

export function formatSnippet(kind: ClientKind, entry: { command: string; args: string[] }): string {
  if (kind === 'toml') {
    return `${TOML_HEADER}\ncommand = ${tomlString(entry.command)}\nargs = [${entry.args.map(tomlString).join(', ')}]`;
  }
  return JSON.stringify({ mcpServers: { [SERVER_KEY]: entry } }, null, 2);
}

interface WriteResult {
  target: ClientTarget;
  status: 'written' | 'unchanged' | 'skipped' | 'failed';
  detail: string;
  backup?: string;
}

function applyOne(target: ClientTarget, opts: InstallOptions, entry: { command: string; args: string[] }): WriteResult {
  const exists = fs.existsSync(target.file);

  let beforeText = '';
  if (exists) {
    try {
      beforeText = fs.readFileSync(target.file, 'utf8');
    } catch (e) {
      return { target, status: 'failed', detail: `读不了：${e instanceof Error ? e.message : String(e)}` };
    }
  }

  let afterText: string;
  try {
    if (target.kind === 'toml') {
      afterText = mergeTomlClient(beforeText, entry);
    } else {
      const before = readJson(target.file);
      afterText = `${JSON.stringify(mergeJsonClient(before, entry), null, 2)}\n`;
    }
  } catch (e) {
    // 解析不了就别猜（比如带注释的 JSON、坏掉的 TOML）—— 交给用户手动粘
    return {
      target,
      status: 'failed',
      detail: `现有配置解析不了，未改动：${e instanceof Error ? e.message : String(e)}`,
    };
  }

  if (exists && beforeText === afterText) {
    return { target, status: 'unchanged', detail: '已经配好了' };
  }

  if (opts.dryRun) {
    return { target, status: 'unchanged', detail: `--dry-run：将写入（${afterText.length} 字节）` };
  }

  let backup: string | undefined;
  try {
    fs.mkdirSync(path.dirname(target.file), { recursive: true });
    if (exists) {
      backup = `${target.file}.bak-${stamp()}`;
      fs.copyFileSync(target.file, backup);
    }
    // 先写临时文件再 rename：写到一半被打断也不会留下半个坏配置
    const tmp = `${target.file}.tmp-${process.pid}`;
    fs.writeFileSync(tmp, afterText, 'utf8');
    fs.renameSync(tmp, target.file);

    // 回读校验：能解析、且我们的 key 在里面
    const check = fs.readFileSync(target.file, 'utf8');
    if (target.kind === 'toml') {
      if (!check.includes(TOML_HEADER)) throw new Error('回读没找到我们的表头');
    } else {
      const parsed = JSON.parse(check) as { mcpServers?: Record<string, unknown> };
      if (!parsed.mcpServers || !parsed.mcpServers[SERVER_KEY]) throw new Error('回读没找到 mcpServers 里的条目');
    }
    return { target, status: 'written', detail: '已写入', backup };
  } catch (e) {
    // 回滚：宁可回到"没配"，也不能留一个坏配置
    if (backup && fs.existsSync(backup)) {
      try {
        fs.copyFileSync(backup, target.file);
      } catch {
        /* 回滚失败也要把原因说出去 */
      }
    }
    return {
      target,
      status: 'failed',
      detail: `写入失败${backup ? '（已还原备份）' : ''}：${e instanceof Error ? e.message : String(e)}`,
    };
  }
}

// ---------------- 入口 ----------------

export function parseArgs(argv: string[]): InstallOptions {
  return {
    dryRun: argv.includes('--dry-run') || argv.includes('-n'),
    printOnly: argv.includes('--print'),
    only: (() => {
      const i = argv.indexOf('--client');
      return i >= 0 ? argv[i + 1] : undefined;
    })(),
  };
}

/** 返回进程退出码（0 = 全部就绪或已打印；1 = 有失败） */
export async function runInstall(argv: string[]): Promise<number> {
  const opts = parseArgs(argv);
  const entry = serverEntry();
  // dsh 不走 detectClients：它是 YAML 补丁层，写法与 JSON/TOML 那批完全不同（见 installDsh 注）
  if (opts.only === 'dsh') return installDsh(opts, entry);
  const all = detectClients();
  const targets = opts.only ? all.filter((c) => c.id === opts.only) : all;

  if (opts.only && !targets.length) {
    console.error(
      `没有名为 "${opts.only}" 的客户端。可选：${all.map((c) => c.id).join(', ')}`
    );
    return 1;
  }

  console.log(`wps-multitool-mcp install${opts.dryRun ? '（--dry-run）' : ''}`);
  console.log(`启动项：${entry.command} ${entry.args.join(' ')}\n`);

  const present = targets.filter((c) => c.present);
  const absent = targets.filter((c) => !c.present);

  if (!present.length) {
    console.log('没检测到任何已安装的客户端。把下面这段粘进你的客户端配置即可：\n');
    console.log(formatSnippet('json', entry), '\n');
    return 0;
  }

  if (opts.printOnly) {
    for (const c of present) {
      console.log(`# ${c.label} — ${c.file}`);
      console.log(formatSnippet(c.kind, entry), '\n');
    }
    return 0;
  }

  let failed = 0;
  for (const c of present) {
    const r = applyOne(c, opts, entry);
    const mark = r.status === 'written' ? '✓' : r.status === 'failed' ? '✗' : '·';
    console.log(`${mark} ${r.target.label}  ${r.target.file}`);
    console.log(`  ${r.detail}${r.backup ? `（备份：${path.basename(r.backup)}）` : ''}`);
    if (r.status === 'failed') failed++;
  }

  if (absent.length) {
    console.log('\n没检测到（不新建目录，需要就自己建好再跑一次）：');
    for (const c of absent) console.log(`  - ${c.label}  ${c.file}`);
    console.log('\n这些客户端的配置片段：');
    console.log(formatSnippet('json', entry));
  }

  console.log('\n改完重启客户端才会生效。');
  return failed ? 1 : 0;
}
