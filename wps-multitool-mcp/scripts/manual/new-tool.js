#!/usr/bin/env node
/**
 * Input: 一个工具名 + --group + --action（可选 --param / --required / --file / --dry）
 * Output: 可直接编译的工具骨架（新 .ts 文件 + 组 index.ts 注册 + 注解 TODO 占位）+ 剩余人工步骤清单
 * Pos: 新增工具的脚手架，把"机械那几步"固定下来，避免每次凭记忆拼错文件名/漏注册。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 它为什么存在
 * ------------
 * 加一个工具真正费时间的不是 handler，而是那条链路上的五处互相牵制的东西：
 *   definition/handler（src/tools/<组>/）→ 组 index.ts 里 spread 进去 → 注解（推断不出来要手写，
 *   且 tool-annotations.ts 与 handwritten-annotations.js 必须两处同步）→ PS 动作 + 路由表 →
 *   验证状态登记 → 文档生成器。漏任何一处都不是"少个功能"，而是【静默错】：
 *   漏注册 → 工具根本不存在；漏路由 → 桥不知道走哪路；漏验证登记 → gen-tool-verification.js 非 0 退出；
 *   漏 gen:docs → CI 的 diff 门禁红。
 *   所以这里只自动化"确定能自动化"的部分，剩下的四件事【打印出来并说清漏了会怎样】，
 *   而不是假装能替人做完（PS 动作的语义、注解的真值、验证状态的真机证据，都不是脚本能猜的）。
 *
 * 刻意不做的事
 * ------------
 * - 不生成 PS 动作：动作体要读 WPS 对象模型，还得决定失败时返回什么，猜出来的代码比空着更危险。
 * - 不写 tool-verification.json：验证状态是"有没有真机证据"的声明，脚手架只能写 static，
 *   让脚本替人声明会变成"没验过也敢写 live"的滑坡 —— 所以只打印该加的那一行。
 * - 不自动跑 build / gen：那几条命令的副作用（dist 变更、生成物刷新）必须由人决定何时发生。
 *
 * 已存在就退出（不覆盖）
 * --------------------
 * 工具名或动作名撞车一律硬失败：PowerShell 的 switch 对同名 case 只认第一个，后加的那个是死代码，
 * 而且【不会报错】。所以这里在写盘之前就把重名查出来，连"先写着试试"的机会都不给。
 *
 * 用法
 * ----
 *   node scripts/manual/new-tool.js wps_excel_set_row_height --group excel --action setRowHeight
 *   node scripts/manual/new-tool.js wps_word_add_watermark --group word --action addWatermark \
 *        --param text:string --param opacity:number --required text
 *   node scripts/manual/new-tool.js wps_ppt_xxx --group ppt --action xxxAction --dry   # 只打印不改盘
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const SRC = path.join(ROOT, 'src');
const PS1 = path.join(ROOT, 'scripts', 'wps-com.ps1');
const ROUTES = path.join(SRC, 'client', 'action-routes.generated.ts');
const ANNOTATIONS = path.join(SRC, 'server', 'tool-annotations.ts');

/** 组的元信息：写进骨架里的三样东西（category / appType / index.ts 的数组名都是 <组>Tools） */
const GROUPS = {
  excel: {
    zh: 'Excel',
    category: 'SPREADSHEET',
    app: 'SPREADSHEET',
    ps: 'Get-WpsExcel',
    psErr: 'WPS Excel not running',
    obj: '$excel',
  },
  word: {
    zh: 'Word',
    category: 'DOCUMENT',
    app: 'WRITER',
    ps: 'Get-WpsWord',
    psErr: 'WPS Writer not running',
    obj: '$word',
  },
  ppt: {
    zh: 'PPT',
    category: 'PRESENTATION',
    app: 'PRESENTATION',
    ps: 'Get-WpsPpt',
    psErr: 'WPS PPT not running',
    obj: '$ppt',
  },
  common: {
    zh: 'Common',
    category: 'COMMON',
    // common 是跨应用工具，wpsClient 不收 appType（传了反而会把动作钉死在某一路）
    app: null,
    ps: null,
    psErr: null,
    obj: null,
  },
};

// ---------------------------------------------------------------------------
// 参数解析（不引入 commander/yargs：这个脚本只有 6 个选项，加依赖不值）
// ---------------------------------------------------------------------------

function usage(code) {
  const msg = [
    '用法: node scripts/manual/new-tool.js <工具名> --group <excel|word|ppt|common> --action <PS动作名> [选项]',
    '',
    '必填:',
    '  <工具名>           形如 wps_excel_set_row_height（前缀必须与 --group 一致）',
    '  --group <g>        excel / word / ppt / common',
    '  --action <name>    wps-com.ps1 里 switch ($Action) 的 case 字面量，驼峰，如 setRowHeight',
    '',
    '可选:',
    '  --param <名:类型>  可重复。类型限 string / number / boolean，如 --param range:string',
    '  --required <名,名> 标记必填参数（必须是 --param 里出现过的）',
    '  --file <文件名>    落地文件名（不含 .ts），默认由动作名转 kebab-case',
    '  --dry              只打印将要生成/改动的文件，不写盘',
  ].join('\n');
  (code === 0 ? console.log : console.error)(msg);
  process.exit(code);
}

function parseArgv(argv) {
  const out = { positional: [], params: [], required: [], dry: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-h') usage(0);
    else if (a === '--dry') out.dry = true;
    else if (a === '--group') out.group = argv[++i];
    else if (a === '--action') out.action = argv[++i];
    else if (a === '--file') out.file = argv[++i];
    else if (a === '--param') out.params.push(argv[++i]);
    else if (a === '--required') out.required.push(...String(argv[++i] || '').split(','));
    else if (a.startsWith('--')) {
      console.error(`未知选项: ${a}`);
      usage(1);
    } else out.positional.push(a);
  }
  out.tool = out.positional[0];
  out.required = out.required.map((s) => s.trim()).filter(Boolean);
  return out;
}

// ---------------------------------------------------------------------------
// 命名工具
// ---------------------------------------------------------------------------

const kebab = (s) => s.replace(/([a-z0-9])([A-Z])/g, '$1-$2').replace(/_/g, '-').toLowerCase();
const camel = (s) =>
  s
    .replace(/[-_]+(.)?/g, (_, c) => (c ? c.toUpperCase() : ''))
    .replace(/^(.)/, (c) => c.toLowerCase());

/** 行尾风格跟随文件本身，避免整文件 diff（仓库里 CRLF 混用过） */
const eolOf = (text) => (text.includes('\r\n') ? '\r\n' : '\n');

// ---------------------------------------------------------------------------
// 重名检查（写盘之前做完，任何一条命中都不写）
// ---------------------------------------------------------------------------

function walk(dir, acc = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, acc);
    else if (e.isFile() && p.endsWith('.ts')) acc.push(p);
  }
  return acc;
}

/**
 * 工具名是否已存在于 src。
 * 生成物也一起查（生成物里有、源码里没有 = 有人删了工具却忘了重跑生成器，仍算撞名），
 * 但报错优先指向【非生成物】的真身 —— 让人去改生成器没有任何意义。
 */
function findToolNameOwner(toolName) {
  const needle = new RegExp(`['"]${toolName}['"]`);
  let fallback = null;
  for (const file of walk(SRC)) {
    const lines = fs.readFileSync(file, 'utf8').split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      if (!needle.test(lines[i])) continue;
      const hit = { file, line: i + 1, text: lines[i].trim() };
      if (file.endsWith('.generated.ts')) {
        fallback = fallback || hit;
        continue;
      }
      return hit;
    }
  }
  return fallback;
}

/** PS 的 switch 对字符串是大小写不敏感的：撞名判断也必须不敏感，否则小写写法会漏 */
function findExistingAction(action) {
  const raw = fs.readFileSync(PS1, 'utf8').replace(/^\uFEFF/, '');
  const lines = raw.split(/\r?\n/);
  const exact = new RegExp(`^\\s*"([A-Za-z_][A-Za-z0-9_]*)"\\s*\\{\\s*$`);
  const want = action.toLowerCase();
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(exact);
    if (m && m[1].toLowerCase() === want) return { name: m[1], line: i + 1 };
  }
  return null;
}

/** 路由表里已有同名（多半是 ps1 改了没重跑 gen-action-routes.js，属于既有不一致，提示但不拦） */
function findExistingRoute(action) {
  if (!fs.existsSync(ROUTES)) return null;
  const lines = fs.readFileSync(ROUTES, 'utf8').split(/\r?\n/);
  const re = new RegExp(`^\\s*['"]?${action}['"]?\\s*:` , 'i');
  for (let i = 0; i < lines.length; i++) if (re.test(lines[i])) return { line: i + 1, text: lines[i].trim() };
  return null;
}

function die(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}

// ---------------------------------------------------------------------------
// 骨架内容
// ---------------------------------------------------------------------------

function buildToolFile(opts) {
  const { tool, action, group, fileBase, params } = opts;
  const g = GROUPS[group];
  const id = camel(action);
  const L = [];

  const paramText = params.length
    ? params.map((p) => `${p.name}:${p.type}`).join(', ')
    : '无参数';

  L.push('/**');
  L.push(` * Input: ${tool} 的参数（${paramText}）`);
  L.push(` * Output: ${action} 的执行结果`);
  L.push(` * Pos: ${g.zh} 第 N 期缺口 —— 【待填：一句话说清现有工具为什么做不到这件事】。`);
  L.push(' *      一旦我被修改，请更新我的头部注释，以及所属目录的md。');
  L.push(' *');
  L.push(' * 本文件由 scripts/manual/new-tool.js 生成，落地前必须补完 4 件事（脚本跑完会打印）：');
  L.push(` *   1) scripts/wps-com.ps1 里实现动作 "${action}"`);
  L.push(' *   2) node scripts/gen-action-routes.js');
  L.push(` *   3) scripts/manual/tool-verification.json 里登记 "${tool}": "static"`);
  L.push(' *   4) node scripts/manual/gen-tool-annotations.js && npm run gen:docs');
  L.push(' */');
  L.push('');
  L.push("import { v4 as uuidv4 } from 'uuid';");
  L.push('import {');
  L.push('  ToolDefinition,');
  L.push('  ToolHandler,');
  L.push('  ToolCallResult,');
  L.push('  ToolCategory,');
  L.push('  RegisteredTool,');
  L.push('  ToolWarning,');
  L.push("} from '../../types/tools';");
  L.push("import { wpsClient } from '../../client/wps-client';");
  if (g.app) L.push("import { WpsAppType } from '../../types/wps';");
  L.push('');
  L.push('type Args = Record<string, unknown>;');
  L.push('');
  L.push('/** PS 侧动作名 —— 必须与 wps-com.ps1 的 case 字面量完全一致（大小写也照抄） */');
  L.push(`const ACTION = '${action}';`);
  L.push('');
  L.push('/**');
  L.push(' * 错误尾注：按动作语义三选一，别照抄这行占位。');
  L.push(" *   读操作   → '[可重试·请求没执行]'");
  L.push(" *   写操作   → '[勿重试·可能已执行一半] 先读回来确认再决定'");
  L.push(" *   参数不对 → '[勿重试·先改参数]'");
  L.push(' * COM 不是事务，写入失败【绝不】自动重试 —— 提示语要替调用方守住这条。');
  L.push(' */');
  L.push("const ERROR_HINT = '【待填：读=可重试 / 写=勿重试·可能已执行一半 / 参数=勿重试·先改参数】';");
  L.push('');
  L.push('/** 成功：部分成功也走这里，把 PS 侧提到顶层的 warnings 带回去（第三态） */');
  L.push('const ok = (text: string, warnings?: ToolWarning[]): ToolCallResult => ({');
  L.push('  id: uuidv4(),');
  L.push('  success: true,');
  L.push("  content: [{ type: 'text', text }],");
  L.push('  ...(warnings && warnings.length ? { warnings } : {}),');
  L.push('});');
  L.push('');
  L.push('const fail = (text: string, error?: string): ToolCallResult => ({');
  L.push('  id: uuidv4(),');
  L.push('  success: false,');
  L.push("  content: [{ type: 'text', text }],");
  L.push('  error: error ?? text,');
  L.push('});');
  L.push('');
  L.push(`export const ${id}Definition: ToolDefinition = {`);
  L.push(`  name: '${tool}',`);
  L.push('  description:');
  L.push("    '【待填：一句话说清它做什么、改的是哪一部分】。\\n\\n' +");
  L.push("    '什么时候用：【待填：用户说了什么/文档到了什么状态就说明该用它 —— 写场景，别写\"需要时用\"】。\\n' +");
  L.push("    '与相近工具的区别：【待填：名字相近的那个管什么，没有就删掉这一行】。',");
  L.push(`  category: ToolCategory.${g.category},`);
  L.push('  inputSchema: {');
  L.push("    type: 'object',");
  if (params.length) {
    L.push('    properties: {');
    for (const p of params) {
      L.push(`      ${p.name}: { type: '${p.type}', description: '【待填】${p.name}' },`);
    }
    L.push('    },');
    if (opts.required.length) L.push(`    required: [${opts.required.map((r) => `'${r}'`).join(', ')}],`);
  } else {
    L.push('    properties: {},');
  }
  L.push('  },');
  L.push('};');
  L.push('');
  if (params.length) {
    L.push(`export const ${id}Handler: ToolHandler = async (args: Args): Promise<ToolCallResult> => {`);
    const typed = params
      .map((p) => `${p.name}?: ${p.type === 'number' ? 'number' : p.type === 'boolean' ? 'boolean' : 'string'}`)
      .join('; ');
    L.push(`  const { ${params.map((p) => p.name).join(', ')} } = args as { ${typed} };`);
    L.push('  const params: Args = {};');
    for (const p of params) L.push(`  if (${p.name} !== undefined) params.${p.name} = ${p.name};`);
  } else {
    L.push(`export const ${id}Handler: ToolHandler = async (): Promise<ToolCallResult> => {`);
    L.push('  const params: Args = {};');
  }
  L.push('  const r = await wpsClient.executeMethod<Record<string, unknown>>(');
  L.push('    ACTION,');
  L.push('    params' + (g.app ? ',' : ''));
  if (g.app) L.push(`    WpsAppType.${g.app}`);
  L.push('  );');
  L.push('  if (!r.success) {');
  L.push(
    "    return fail(`【待填：中文动作名】执行失败：${r.error ?? '未知原因'}。${ERROR_HINT}`, r.error);"
  );
  L.push('  }');
  L.push("  return ok('【待填：把 r.data 说成人能读的一句话，别回显整个对象】', r.warnings);");
  L.push('};');
  L.push('');
  L.push(`export const ${id}Tools: RegisteredTool[] = [`);
  L.push(`  { definition: ${id}Definition, handler: ${id}Handler },`);
  L.push('];');
  L.push('');
  L.push(`export default ${id}Tools;`);
  L.push('');

  return {
    id,
    fileBase,
    content: L.join('\n'),
    relPath: `src/tools/${group}/${fileBase}.ts`,
  };
}

/** 在组 index.ts 里加 import + spread；两处锚点都找不到就硬失败（宁可不动，也不写半截） */
function patchGroupIndex(group, id, fileBase, dry) {
  const file = path.join(SRC, 'tools', group, 'index.ts');
  const before = fs.readFileSync(file, 'utf8');
  const eol = eolOf(before);

  const importLine = `import { ${id}Tools } from './${fileBase}';`;
  if (before.includes(importLine) || before.includes(`'./${fileBase}'`)) {
    die(`${path.relative(ROOT, file)} 里已经引用了 './${fileBase}' —— 不重复注册`);
  }

  // 锚点 1：最后一个相对 import（组 index.ts 的 import 块是连续的，紧跟其后的就是注释/数组）
  const lines = before.split(/\r?\n/);
  let lastImport = -1;
  for (let i = 0; i < lines.length; i++) {
    if (/^import\s+.*from\s+'\.\/[^']+';?\s*$/.test(lines[i])) lastImport = i;
  }
  if (lastImport < 0) die(`${path.relative(ROOT, file)} 里找不到相对 import 块，拒绝盲写`);

  // 锚点 2：<组>Tools 数组的收尾 ];（从数组声明处往后找第一个，别撞上文件里其它数组）
  const declRe = new RegExp(`^export const ${group}Tools: RegisteredTool\\[\\] = \\[`, 'm');
  const decl = before.match(declRe);
  if (!decl) die(`${path.relative(ROOT, file)} 里找不到 "export const ${group}Tools: RegisteredTool[] = ["`);
  const afterDecl = before.slice(decl.index).split(/\r?\n/);
  let closeOffset = -1;
  for (let i = 1; i < afterDecl.length; i++) {
    if (/^\];\s*$/.test(afterDecl[i])) {
      closeOffset = i;
      break;
    }
  }
  if (closeOffset < 0) die(`${path.relative(ROOT, file)} 里 ${group}Tools 数组没有找到收尾 ];`);

  const declLine = before.slice(0, decl.index).split(/\r?\n/).length - 1;
  const spreadAt = declLine + closeOffset;

  // 先插靠后的行，避免下标位移
  lines.splice(spreadAt, 0, `  ...${id}Tools,`);
  lines.splice(lastImport + 1, 0, importLine);

  const after = lines.join(eol);
  if (!dry) fs.writeFileSync(file, after, 'utf8');
  return { file, before, after };
}

/**
 * 注解占位：只写一段 TODO 注释，不写假的注解值。
 * 理由是硬性的 —— 注解会被客户端拿去做"要不要自动执行"的安全决策，
 * 猜一个 destructiveHint=false 上去，比没有注解更糟（tool-annotations.ts 顶部已写死这条）。
 */
function patchAnnotations(tool, action, dry) {
  const before = fs.readFileSync(ANNOTATIONS, 'utf8');
  const eol = eolOf(before);
  const anchor = 'const EXTRA_ANNOTATIONS: Record<string, ToolAnnotations> = {';
  const idx = before.indexOf(anchor);
  if (idx < 0) die(`${path.relative(ROOT, ANNOTATIONS)} 里找不到 EXTRA_ANNOTATIONS，拒绝盲写`);

  const tail = before.slice(idx).split(/\r?\n/);
  let closeOffset = -1;
  for (let i = 1; i < tail.length; i++) {
    if (/^};\s*$/.test(tail[i])) {
      closeOffset = i;
      break;
    }
  }
  if (closeOffset < 0) die(`${path.relative(ROOT, ANNOTATIONS)} 里 EXTRA_ANNOTATIONS 没有收尾 };`);

  const at = before.slice(0, idx).split(/\r?\n/).length - 1 + closeOffset;
  const lines = before.split(/\r?\n/);
  const todo = [
    `  // TODO(脚手架 ${tool})：动作名 ${action}`,
    '  // 若 gen-tool-annotations.js 观测得到动作名、推断结果也对 → 删掉本段，什么都不用加；',
    '  // 若观测不到（编排型/运行期参数）或推断会骗人 → 在这里补真值，',
    '  // 并【同步】scripts/manual/handwritten-annotations.js 的 EXTRA_ANNOTATIONS_KEYS。',
    '  // 两处漏一处，生成器的同步断言就非 0 退出（这是故意的，静默漂移会让工具悄悄退回按名字猜）。',
  ];
  lines.splice(at, 0, ...todo);

  const after = lines.join(eol);
  if (!dry) fs.writeFileSync(ANNOTATIONS, after, 'utf8');
  return { file: ANNOTATIONS, before, after };
}

// ---------------------------------------------------------------------------
// 收尾清单
// ---------------------------------------------------------------------------

function printRemaining(opts, created) {
  const { tool, action, group } = opts;
  const g = GROUPS[group];
  const rel = (p) => path.relative(ROOT, p).replace(/\\/g, '/');

  const psSnippet = g.ps
    ? [
        `    "${action}" {`,
        `        ${g.obj} = ${g.ps}`,
        `        if ($null -eq ${g.obj}) { Output-Json @{ success = $false; error = "${g.psErr}" }; return }`,
        '        # TODO: 取对象（Resolve-* / ActiveSheet / ActivePresentation），做动作，把结果塞进 data',
        '        Output-Json @{ success = $true; data = @{ /* TODO */ } }',
        '    }',
      ]
    : [
        `    "${action}" {`,
        '        # TODO: common 动作（跨应用或纯基础设施），按需要取三件套之一',
        '        Output-Json @{ success = $true; data = @{ /* TODO */ } }',
        '    }',
      ];

  console.log('');
  console.log('════════════════════════════════════════════════════════════');
  console.log(' 还剩 4 件事必须人工做（脚本不替你做，因为都不是机械步骤）');
  console.log('════════════════════════════════════════════════════════════');
  console.log('');
  console.log(`[1/4] 在 scripts/wps-com.ps1 的 switch ($Action) 里实现动作 "${action}"`);
  console.log('      重名检查已通过（PS 的 switch 只认第一个同名 case，后加的是死代码且不报错）');
  console.log('      建议插在同类动作附近，别贴文件末尾（末尾是 default 与收尾逻辑）：');
  console.log('');
  for (const l of psSnippet) console.log(l);
  console.log('');
  console.log('[2/4] 重跑路由表 + 编译');
  console.log('      node scripts/gen-action-routes.js && npm run build');
  console.log('      漏了 → action-routes.generated.ts 里没有这个动作，桥不知道派给哪个 app 子进程');
  console.log('');
  console.log('[3/4] 登记验证状态（默认 static，交验证者跑真机）');
  console.log('      在 scripts/manual/tool-verification.json 的 "tools" 里加一行：');
  console.log(`          "${tool}": "static",`);
  console.log('      node scripts/manual/gen-tool-verification.js');
  console.log('      漏了 → 生成器直接非 0 退出（"注册表里有、表里没有"是硬门禁，不允许默认当验过）');
  console.log('');
  console.log('[4/4] 重生成注解与文档');
  console.log('      node scripts/manual/gen-tool-annotations.js && npm run gen:docs');
  console.log('      漏了 → annotations/README/TOOLS.md 停在旧数字，CI 的 diff 门禁红');
  console.log('      注解 TODO 占位在 src/server/tool-annotations.ts 的 EXTRA_ANNOTATIONS 里，按注释决定删或补');
  console.log('');
  console.log('── 落地前自检（都不起桥，可以和真机验证并行跑）────────────────');
  console.log('  npm run build && node scripts/manual/annotation-verify.js \\');
  console.log('    && node scripts/manual/verification-verify.js && node scripts/manual/basic-tier-verify.js');
  console.log('');
  if (created.length) {
    console.log('── 反悔了怎么撤 ────────────────────────────────────────────');
    console.log(`  del "${created[0].replace(/\//g, '\\')}"   # 本次新建的骨架文件`);
    for (const f of created.slice(1)) console.log(`  git checkout -- "${f}"`);
    console.log('');
  }
}

// ---------------------------------------------------------------------------
// main
// ---------------------------------------------------------------------------

function main() {
  const argv = process.argv.slice(2);
  if (!argv.length) usage(1);
  const o = parseArgv(argv);

  if (!o.tool) die('缺少工具名。用法见 --help');
  if (!o.group) die('缺少 --group（excel / word / ppt / common）');
  if (!o.action) die('缺少 --action（wps-com.ps1 里的 case 名，驼峰）');
  if (!GROUPS[o.group]) die(`--group 只能是 excel / word / ppt / common，收到 "${o.group}"`);

  const nameRe = new RegExp(`^wps_${o.group}_[a-z0-9]+(_[a-z0-9]+)*$`);
  if (!nameRe.test(o.tool)) {
    die(
      `工具名 "${o.tool}" 不合规：必须以 wps_${o.group}_ 开头，其余全小写下划线（如 wps_${o.group}_set_xxx）`
    );
  }
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(o.action)) {
    die(`--action "${o.action}" 不是合法标识符（PS case 字面量与 gen-action-routes 的解析都只认这个形状）`);
  }
  if (o.tool.length - `wps_${o.group}_`.length < 2) die('工具名后半段太短，起个能被搜到的名字');

  const params = o.params.map((raw) => {
    const [name, type = 'string'] = String(raw).split(':');
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name || '')) die(`--param 名字 "${name}" 不是合法标识符`);
    if (!['string', 'number', 'boolean'].includes(type)) die(`--param 类型只能是 string/number/boolean，收到 "${type}"`);
    return { name, type };
  });
  const paramNames = new Set(params.map((p) => p.name));
  for (const r of o.required) {
    if (!paramNames.has(r)) die(`--required "${r}" 不在 --param 列表里`);
  }

  // ---- 重名：三处都查，任何一处命中都不写盘 ----
  const owner = findToolNameOwner(o.tool);
  if (owner) {
    die(`工具名 "${o.tool}" 已存在：${path.relative(ROOT, owner.file)}:${owner.line}  →  ${owner.text.slice(0, 80)}`);
  }
  const dupAction = findExistingAction(o.action);
  if (dupAction) {
    die(
      `动作名 "${o.action}" 在 scripts/wps-com.ps1:${dupAction.line} 已经有了（写成 "${dupAction.name}"）。\n` +
        '  PS 的 switch 只执行第一个同名 case，再加一个就是死代码 —— 换个动作名，或直接复用现有动作。'
    );
  }
  const staleRoute = findExistingRoute(o.action);
  if (staleRoute) {
    console.warn(
      `! 路由表 src/client/action-routes.generated.ts:${staleRoute.line} 里已经有 ${o.action}，` +
        '但 ps1 里没这个 case —— 说明生成物与真相源已不一致，落地时先跑一次 gen-action-routes.js'
    );
  }

  const fileBase = o.file ? String(o.file).replace(/\.ts$/, '') : kebab(o.action);
  const toolFileAbs = path.join(SRC, 'tools', o.group, `${fileBase}.ts`);
  if (fs.existsSync(toolFileAbs)) {
    die(`${path.relative(ROOT, toolFileAbs)} 已存在，拒绝覆盖（换 --file 或先确认要改的是不是它）`);
  }

  const built = buildToolFile({ ...o, fileBase, params });

  // 先把两个补丁在内存里算完，全部成功再落盘 —— 免得留下"骨架写了、注册没写"的半截状态
  const indexPatch = patchGroupIndex(o.group, built.id, fileBase, true);
  const annoPatch = patchAnnotations(o.tool, o.action, true);

  const touched = [built.relPath, path.relative(ROOT, indexPatch.file).replace(/\\/g, '/'), path.relative(ROOT, annoPatch.file).replace(/\\/g, '/')];

  if (o.dry) {
    console.log(`[dry] 将新建 ${touched[0]}`);
    console.log(`[dry] 将改动 ${touched[1]}`);
    console.log(`[dry] 将改动 ${touched[2]}`);
    console.log('');
    console.log('---- ' + touched[0] + ' ----');
    console.log(built.content);
    printRemaining(o, []);
    return;
  }

  fs.writeFileSync(toolFileAbs, built.content, 'utf8');
  fs.writeFileSync(indexPatch.file, indexPatch.after, 'utf8');
  fs.writeFileSync(annoPatch.file, annoPatch.after, 'utf8');

  console.log(`✓ 新建 ${touched[0]}`);
  console.log(`✓ 注册 ${touched[1]}（import + ...${built.id}Tools）`);
  console.log(`✓ 注解占位 ${touched[2]}（EXTRA_ANNOTATIONS 里一段 TODO 注释，不是注解值）`);
  printRemaining(o, touched);

  console.log('下一步先跑 npm run build 把类型错挡住，再做 PS 动作。');
}

main();
