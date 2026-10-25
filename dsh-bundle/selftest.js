/**
 * Input: 无（可选 WPS_MCP_ENTRY）
 * Output: 结构自检 + 连接层自检；失败 exit 1
 * Pos: dsh-bundle 的自检脚本。**不需要 dsh**，只验两件 bundle 最容易坏的事：
 *      ① 包结构符合 dsh 的 bundle 规范；② wps-multitool-mcp 起得来、工具列得出来。
 *
 * 为什么不再 import './index.js'
 * -----------------------------
 * v0.4 起 bundle **没有 index.js 了** —— 挂载改由 dsh 的 loader 按包名解析
 * （我们自己的代码 import 不到 app.asar 里的官方客户端）。所以这里自己算入口，
 * 不再依赖已被删除的模块。
 *
 * 用法：node dsh-bundle/selftest.js
 */

import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
let failed = 0;
const ok = (msg) => console.log(`  ✓ ${msg}`);
const bad = (msg) => {
  failed++;
  console.error(`  ✗ ${msg}`);
};

// ---------- 1. 结构检查 ----------
console.log('结构：');

const pkgPath = join(HERE, 'package.json');
const pkg = JSON.parse(readFileSync(pkgPath, 'utf8'));

if (pkg.dsh?.bundle?.patch === './cordis.patch.yml') ok('package.json 声明了 dsh.bundle.patch');
else bad(`dsh.bundle.patch 不对：${JSON.stringify(pkg.dsh?.bundle?.patch)}`);

if (pkg.engines?.dsh) ok(`engines.dsh = ${pkg.engines.dsh}`);
else bad('缺 engines.dsh，dsh 会拒绝安装');

const patchPath = join(HERE, 'cordis.patch.yml');
if (existsSync(patchPath)) ok('cordis.patch.yml 存在');
else bad('cordis.patch.yml 不存在');

const patchText = existsSync(patchPath) ? readFileSync(patchPath, 'utf8') : '';

// 三条关键行必须都在
for (const [needle, label] of [
  ["name: '@deepseek-ai/dsh-mcp-client'", '挂了官方 MCP 客户端'],
  ['- id: skill-filesystem', '启用了 skill-filesystem'],
  ['- id: tool-skill', '启用了 tool-skill'],
]) {
  if (patchText.includes(needle)) ok(label);
  else bad(`patch 里缺：${label}`);
}

// patch 里不许出现相对路径的 args（发布后必然是错的）
if (/args:\s*\n\s*-\s*['"]?\.\.?\//.test(patchText)) bad('patch 的 args 里出现了相对路径');
else ok('args 里没有相对路径');

// ---------- 2. 技能资产检查 ----------
console.log('技能：');

const skillDir = join(HERE, 'assets', 'wps-multitool');
const skillMd = join(skillDir, 'SKILL.md');
if (existsSync(skillMd)) {
  const text = readFileSync(skillMd, 'utf8');
  const fm = /^---\r?\n([\s\S]*?)\r?\n---/.exec(text);
  if (!fm) {
    bad('SKILL.md 没有 YAML frontmatter');
  } else {
    const nameM = /^name:\s*(.+)$/m.exec(fm[1]);
    const descM = /^description:\s*(.+)$/m.exec(fm[1]);
    const name = nameM?.[1]?.trim();
    if (name && /^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name)) ok(`SKILL.md name = ${name}（kebab-case 合规）`);
    else bad(`SKILL.md 的 name 不合规：${name ?? '(缺失)'}`);
    if (descM?.[1]?.trim()) ok('SKILL.md 有 description');
    else bad('SKILL.md 缺 description（dsh 必填）');
  }
} else {
  bad('assets/wps-multitool/SKILL.md 不存在');
}

const refDir = join(skillDir, 'references');
if (existsSync(refDir)) {
  const refs = readdirSync(refDir).filter((f) => f.endsWith('.md'));
  if (refs.length) ok(`references/ 有 ${refs.length} 个文件`);
  else bad('references/ 是空的');
} else {
  bad('references/ 不存在');
}

// ---------- 3. 连接层检查 ----------
console.log('连接层：');

function resolveEntry() {
  if (process.env.WPS_MCP_ENTRY && existsSync(process.env.WPS_MCP_ENTRY)) return process.env.WPS_MCP_ENTRY;
  try {
    const got = createRequire(import.meta.url).resolve('wps-multitool-mcp/dist/index.js');
    if (existsSync(got)) return got;
  } catch {
    /* 落到下一档 */
  }
  const sibling = join(HERE, '..', 'wps-multitool-mcp', 'dist', 'index.js');
  if (existsSync(sibling)) return sibling;
  return '';
}

const entry = resolveEntry();
if (!entry) {
  bad('找不到 wps-multitool-mcp 入口。先 npm run build，或设 WPS_MCP_ENTRY');
} else {
  ok(`入口：${entry}`);
  try {
    const req = createRequire(entry);
    const clientUrl = pathToFileURL(req.resolve('@modelcontextprotocol/sdk/client/index.js')).href;
    const stdioUrl = pathToFileURL(req.resolve('@modelcontextprotocol/sdk/client/stdio.js')).href;
    const { Client } = await import(clientUrl);
    const { StdioClientTransport } = await import(stdioUrl);

    const started = Date.now();
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [entry],
      env: { ...process.env, WPS_TOOL_TIER: process.env.WPS_TOOL_TIER || 'essential' },
      stderr: 'inherit',
    });
    const client = new Client({ name: 'dsh-multitool-selftest', version: '0.4.0' }, { capabilities: {} });
    await client.connect(transport);
    const { tools } = await client.listTools();
    const ms = Date.now() - started;

    if (tools.length === 0) bad('工具数为 0：server 起来了但没列工具');
    else ok(`工具数：${tools.length}（握手+发现 ${ms}ms）`);

    const names = tools.map((t) => t.name);
    const illegal = names.filter((n) => !/^[A-Za-z0-9_-]+$/.test(n));
    if (illegal.length) bad(`有 ${illegal.length} 个工具名不合 dsh 字符集：${illegal.slice(0, 3).join(', ')}`);
    else ok('工具名全部合 dsh 的字符集 [A-Za-z0-9_-]');

    // 前缀是 mcp__wpslocal__（不是 mcp__wps__ —— 那是 npm 上别人的 dsh-wps 用的）
    const tooLong = names.filter((n) => `mcp__wpslocal__${n}`.length > 64);
    if (tooLong.length) bad(`有 ${tooLong.length} 个工具名加前缀后超 64 字符：${tooLong.slice(0, 3).join(', ')}`);
    else ok('加 mcp__wpslocal__ 前缀后全部 ≤ 64 字符');

    await client.close();
  } catch (e) {
    bad(`连接失败：${e.message}`);
  }
}

console.log(failed === 0 ? '\n全部通过' : `\n${failed} 项失败`);
process.exit(failed === 0 ? 0 : 1);
