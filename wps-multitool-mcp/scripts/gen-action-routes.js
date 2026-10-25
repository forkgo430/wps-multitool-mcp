#!/usr/bin/env node
/**
 * Input: scripts/wps-com.ps1（**只读顶层 switch ($Action)** 的直接 case）
 * Output: src/client/action-routes.generated.ts（action -> 应用 的路由表）
 * Pos: 路由表生成器 + 顶层 case 解析器（gen-tool-annotations.js 复用同一个解析器）。
 *      一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要生成而不是手写
 * ----------------------
 * 桥是按应用拆子进程的，所以每个 action 必须知道该走哪一路。手写必错且会腐坏。
 * 这里直接从 wps-com.ps1 的 case 体里推断：该 case 里用到了哪个应用的对象
 *   Get-WpsExcel / $excel  -> excel
 *   Get-WpsWord  / $word   -> word
 *   Get-WpsPpt   / $ppt    -> ppt
 * 命中多个（或一个都没有，比如 ping / save / wpsStatus 这类跨应用或纯基础设施动作）-> common。
 *
 * 第 1 期第 2 章第 2 轮修的口径（重要）
 * ------------------------------------
 * 原来这里用 `lines.findIndex(switch ($Action))` 之后**扫到文件末尾**，把所有形如
 * `  "x" {` 的行都当 case —— 于是嵌套 switch 的分支标签也被算了进来：
 *   cleanData 的 switch ($op)：trim / remove_duplicates / unify_date / remove_empty_rows
 *   smartFillField 的分支：placeholder / underline / afterColon / afterLabel
 *   handleRevisions 的分支：acceptOne / rejectOne
 * 它们不是动作，运行时也没人拿它们当动作名调（所以一直没炸），但让"动作数"虚高 10。
 *
 * 现在用**括号深度**判"顶层"：主 switch 的直接 case 深度是 1，嵌套 switch 里的标签深度 >= 2。
 * 判深度前先剥注释（注释里的花括号会把计数带偏），剥的时候保留换行，行号才对得上。
 *
 * 三个数别混（N = 顶层 case 数）
 *   MCP 工具数        ← 注册表，与本脚本无关
 *   可分发动作数 N    ← 本脚本统计的（含 __ 开头的内部探针）
 *   面向用户的动作数   ← N 减去 __ 探针
 *
 * 用法：node scripts/gen-action-routes.js
 * 改完 wps-com.ps1 就该重跑一次，然后 `npm run build`。
 */

const fs = require('fs');
const path = require('path');

const PS = path.join(__dirname, 'wps-com.ps1');
const OUT = path.join(__dirname, '..', 'src', 'client', 'action-routes.generated.ts');

/** 形如 `    "setFormula" {` —— case 必须独占一行、花括号后无其它内容 */
const CASE_RE = /^\s*"([A-Za-z_][A-Za-z0-9_]*)"\s*\{\s*$/;

/**
 * 剥掉 PowerShell 注释（# 行注释 与 <# #> 块注释），保留字符串字面量里的 #。
 * 只用于结构分析，不求 100% 语法精确 —— 但必须做到两件事：
 *   ① 注释里的花括号不参与深度计数；② 行数不变（下游按行号取 case 体）。
 */
function stripComments(src) {
  const out = [];
  let i = 0;
  let inSingle = false;
  let inDouble = false;
  let inBlock = false;
  while (i < src.length) {
    const c = src[i];
    const next = src[i + 1];
    if (inBlock) {
      if (c === '#' && next === '>') {
        inBlock = false;
        i += 2;
        continue;
      }
      if (c === '\n') out.push('\n');
      i += 1;
      continue;
    }
    if (!inSingle && !inDouble && c === '<' && next === '#') {
      inBlock = true;
      i += 2;
      continue;
    }
    if (c === "'" && !inDouble) inSingle = !inSingle;
    else if (c === '"' && !inSingle) inDouble = !inDouble;
    else if (c === '#' && !inSingle && !inDouble) {
      // 行注释：跳到行尾（不吞掉 \n，行数必须保持）
      while (i < src.length && src[i] !== '\n') i += 1;
      continue;
    }
    out.push(c);
    i += 1;
  }
  return out.join('');
}

/** 一行里净增多少个花括号（用于深度跟踪） */
function braceDelta(line) {
  let d = 0;
  for (const ch of line) {
    if (ch === '{') d += 1;
    else if (ch === '}') d -= 1;
  }
  return d;
}

/**
 * 解析 wps-com.ps1 的**顶层** switch ($Action) 直接 case。
 *
 * 返回 { cases: [{name, line}], bodyOf(index), total, dups, psPath }：
 *   cases  —— 按出现顺序、**含重复**（重复的那份是死代码，由调用方决定要不要报）
 *   bodyOf —— 取第 i 个 case 的正文（已剥注释，含 case 标签行到下一个 case 之前）
 *
 * 为什么导出：gen-tool-annotations.js 要用同一份"ps1 里到底有哪些动作"做全量交叉校验。
 * 两处各写一套解析必然分裂，而分裂的结局是"路由表说 304、校验说 314"。
 */
function parseTopLevelCases(psPath = PS) {
  const raw = fs.readFileSync(psPath, 'utf8').replace(/^\uFEFF/, '');
  const clean = stripComments(raw);
  const lines = clean.split(/\r?\n/);

  const switchIdx = lines.findIndex((l) => l.trim().startsWith('switch ($Action)'));
  if (switchIdx < 0) {
    throw new Error('未在 wps-com.ps1 里找到 switch ($Action)');
  }

  // 主 switch 的花括号在同一行上，处理完这一行深度就是 1
  let depth = braceDelta(lines[switchIdx]);

  const cases = [];
  /**
   * 深度 1 上的所有 `"x" {` 标签（**含**花括号后还带内容的那些）。
   * 用途：verify-action-routing.js 要回答"有没有真动作因为格式不纯而被生成器漏掉"。
   * pure=true 才进 cases（生成器的 CASE_RE 只认独占一行的那种）。
   */
  const labels = [];
  const LOOSE_CASE_RE = /^\s*"([A-Za-z_][A-Za-z0-9_]*)"\s*\{/;
  let switchEndLine = lines.length;
  for (let i = switchIdx + 1; i < lines.length; i++) {
    // 先认 case（此时 depth===1 才是顶层），再更新深度 —— 顺序反了会把 case 自己的 { 算进去
    if (depth === 1) {
      const pure = CASE_RE.test(lines[i]);
      const loose = lines[i].match(LOOSE_CASE_RE);
      if (loose) {
        labels.push({ name: loose[1], line: i, pure });
        if (pure) cases.push({ name: loose[1], line: i });
      }
    }
    depth += braceDelta(lines[i]);
    if (depth <= 0) {
      switchEndLine = i; // 主 switch 收尾
      break;
    }
  }

  if (cases.length === 0) {
    throw new Error('未解析到任何顶层 case（括号深度跟踪可能失效了）');
  }

  const bodyOf = (i) => {
    const start = cases[i].line;
    const end = i + 1 < cases.length ? cases[i + 1].line : switchEndLine + 1;
    return lines.slice(start, end).join('\n');
  };

  const seen = new Map();
  const dups = [];
  for (const c of cases) {
    if (seen.has(c.name)) dups.push(c.name);
    else seen.set(c.name, c.line);
  }

  return { cases, labels, bodyOf, total: cases.length, dups, psPath };
}

function main() {
  const { cases, bodyOf, total, dups } = parseTopLevelCases();

  const table = {};
  const dupReported = [];
  for (let i = 0; i < cases.length; i++) {
    const body = bodyOf(i); // 已剥注释：注释里提到的 Get-WpsExcel 不会被误判成本 case 的应用

    const hits = new Set();
    for (const m of body.matchAll(/Get-Wps(Excel|Word|Ppt)\b/g)) hits.add(m[1].toLowerCase());
    for (const m of body.matchAll(/\$(excel|word|ppt)\b/g)) hits.add(m[1].toLowerCase());
    for (const m of body.matchAll(/ComObject\s+'?(Ket|Kwps|Kwpp)\.Application/g)) {
      hits.add(m[1] === 'Ket' ? 'excel' : m[1] === 'Kwps' ? 'word' : 'ppt');
    }

    // PowerShell 的 switch 是【第一个匹配的分支生效】，同名分支后面的都是死代码。
    // 同名时保留第一次出现的判定，并把重复项报出来。
    const name = cases[i].name;
    if (name in table) {
      dupReported.push(name);
      continue;
    }
    table[name] = hits.size === 1 ? [...hits][0] : 'common';
  }

  const counts = {};
  for (const v of Object.values(table)) counts[v] = (counts[v] || 0) + 1;
  const userFacing = Object.keys(table).filter((n) => !n.startsWith('__')).length;

  const header = `/**
 * 本文件由 scripts/gen-action-routes.js 自动生成，请勿手改。
 * 源：scripts/wps-com.ps1 的【顶层】switch ($Action) 分支（嵌套 switch 的分支标签不算动作）。
 * 重新生成：node scripts/gen-action-routes.js
 *
 * 统计：${Object.entries(counts).map(([k, v]) => `${k}=${v}`).join('  ')}  合计 ${Object.keys(table).length}
 * 口径：可分发动作 ${Object.keys(table).length} 个（含 __ 内部探针 ${Object.keys(table).length - userFacing} 个）→ 面向用户 ${userFacing} 个
 */
`;

  const body = Object.keys(table)
    .sort()
    .map((k) => `  ${JSON.stringify(k)}: '${table[k]}',`)
    .join('\n');

  fs.writeFileSync(OUT, header + '\nimport type { AppKey } from \'./action-routes\';\n\nexport const ACTION_ROUTES: Record<string, AppKey> = {\n' + body + '\n};\n', 'utf8');

  console.log(`已生成 ${path.relative(path.join(__dirname, '..'), OUT)}`);
  console.log(
    `顶层 case ${total} 个，去重后 ${Object.keys(table).length}：${Object.entries(counts)
      .map(([k, v]) => `${k}=${v}`)
      .join('  ')}`
  );
  console.log(`口径：可分发动作 ${Object.keys(table).length}（含 __ 探针 ${Object.keys(table).length - userFacing}）→ 面向用户 ${userFacing}`);
  if (dupReported.length) {
    console.warn(`注意：有 ${dupReported.length} 个重复 case 名（switch 里后者是死代码，已按第一个取值）：${[...new Set(dupReported)].join(', ')}`);
  } else {
    console.log('重复 case 检查：无同名 case（每个动作只有一个实现，不存在死代码）');
  }
}

module.exports = { parseTopLevelCases, stripComments, CASE_RE };

if (require.main === module) {
  main();
}
