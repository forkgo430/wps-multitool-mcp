/**
 * Input: scripts/wps-com.ps1（case 体 + 助手函数体）+ dist/tools 的编译产物
 * Output: 候选列表 [{act, missing, file}] —— TS 传了、PS 侧读不到的参数键
 * Pos: param-shape-audit.js 与 param-shape-verify.js 共用的扫描核心。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么从 audit 里抽出来
 * ----------------------
 * audit（诊断，给人看）与 verify（门禁，给 CI 看）必须是同一套规则。两处各写一份，
 * 迟早出现"审计说不算问题、门禁说算"的分裂 —— 而分裂的结局一定是门禁被当成误报 bypass 掉。
 *
 * 第 3 轮细化过的两条规则（原来会误报）：
 *   1) 助手函数感知：case 体里出现 Resolve-Book / Resolve-Doc / Get-TargetPres / Resolve-Sheet
 *      等把整个 $p 传进去的助手时，该助手内部读的键算这个 case 也读。
 *      不认这条，光 Excel 侧就会因为 sheet 误报 30 多处 —— 而 sheet 恰恰是已修的那一类。
 *   2) 同一动作多个 case 体（同名 case 后者是死代码）只算第一个，与 PowerShell 实际行为一致。
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');

/** 读 $p.<键> 与 $p["<键>"] 两种写法 */
function collectKeys(src) {
  const set = new Set();
  const dot = /\$p\.([A-Za-z_][A-Za-z0-9_]*)/g;
  let m;
  while ((m = dot.exec(src))) set.add(m[1]);
  const br = /\$p\[["']([A-Za-z0-9_]+)["']\]/g;
  while ((m = br.exec(src))) set.add(m[1]);
  return set;
}

/** 把 ps1 按 case 切开：同名 case 只保留第一个（PS 的 switch 也只执行第一个） */
function splitCases(ps) {
  const lines = ps.split(/\r?\n/);
  const switchIdx = lines.findIndex((l) => l.trim().startsWith('switch ($Action)'));
  if (switchIdx < 0) return { helpers: {}, cases: [] };

  // 1) switch 之前的顶层 function：收集"把 $p 传进去的助手"读了哪些键
  // 函数名是 PowerShell 的 Verb-Noun 短横线风格（Resolve-Book / Get-TargetPres），
  // 字符类里必须带 '-' —— 漏了它会让本正则一条都匹配不上，然后"助手感知"静默失效。
  const helpers = {};
  const head = lines.slice(0, switchIdx).join('\n');
  const fnRe = /^function\s+([A-Za-z_][A-Za-z0-9_-]*)\s*\(([^)]*)\)\s*\{/gm;
  let fm;
  while ((fm = fnRe.exec(head))) {
    const name = fm[1];
    if (!fm[2].includes('$p')) continue; // 不收 $p 的助手与参数形状无关
    // 函数体：从签名后到下一个顶层 function 或文件尾
    const rest = head.slice(fm.index + fm[0].length);
    const nextFn = rest.search(/^function\s+[A-Za-z_]/m);
    const body = nextFn < 0 ? rest : rest.slice(0, nextFn);
    helpers[name] = collectKeys(body);
  }

  // 2) switch 里的 case 体
  const caseRe = /^\s{4}"([A-Za-z_][A-Za-z0-9_]*)"\s*\{\s*$/;
  const cases = [];
  const seen = new Set();
  for (let i = switchIdx; i < lines.length; i++) {
    const m = lines[i].match(caseRe);
    if (!m) continue;
    if (seen.has(m[1])) continue; // 同名 case 后者是死代码，不参与判定
    seen.add(m[1]);
    let end = lines.length;
    for (let j = i + 1; j < lines.length; j++) {
      if (caseRe.test(lines[j])) {
        end = j;
        break;
      }
    }
    cases.push({ name: m[1], line: i + 1, body: lines.slice(i, end).join('\n') });
  }
  return { helpers, cases };
}

/** 一个 case 实际读到的键 = 自己读的 + 它调用的助手读的 */
function effectiveKeys(caseBody, helpers) {
  const keys = collectKeys(caseBody);
  for (const [fn, fnKeys] of Object.entries(helpers)) {
    if (!caseBody.includes(fn)) continue;
    for (const k of fnKeys) keys.add(k);
  }
  return keys;
}

/** TS 侧每个 executeMethod('<action>', {字面量}) 传了哪些键 */
function tsCallSites() {
  const files = [];
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) {
        if (e.name !== 'node_modules') walk(p);
      } else if (e.name.endsWith('.js')) files.push(p);
    }
  };
  walk(path.join(ROOT, 'dist', 'tools'));

  const sites = [];
  for (const f of files) {
    const src = fs.readFileSync(f, 'utf8');
    const er = /executeMethod\(\s*'([A-Za-z0-9_]+)'\s*,\s*\{([^{}]*)\}/g;
    let x;
    while ((x = er.exec(src))) {
      const act = x[1];
      const keys = x[2]
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean)
        .map((s) => (s.indexOf(':') > 0 ? s.slice(0, s.indexOf(':')) : s).trim())
        .filter((s) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(s));
      sites.push({ act, keys, file: path.relative(ROOT, f).replace(/\\/g, '/') });
    }
  }
  return sites;
}

/**
 * 解析入口归一化块（wps-com.ps1 的 `$ParamAlias` / `$StyleFlatten` / `$FlattenKeys`）。
 *
 * 第 3 轮加：静默丢弃的那一类是在 switch 之前就抹平的（改名 + 嵌套摊平），
 * 规则表本身就在 ps1 里。既然表在那儿，扫描就该认它 —— 否则明明修好的键还会被报成候选，
 * 白名单就得堆 40 条，而白名单一大就等于没有门禁。
 */
function parseEntryNormalization(ps) {
  const start = ps.indexOf('# 入参归一化');
  const end = ps.indexOf('switch ($Action)');
  const empty = { alias: {}, flatten: {}, flattenKeys: {} };
  if (start < 0 || end < 0 || end <= start) return empty;

  const block = ps.slice(start, end);
  const alias = {};
  const aliasSec = block.match(/\$ParamAlias\s*=\s*@\{([\s\S]*?)\n\}/);
  if (aliasSec) {
    const entryRe = /'([A-Za-z0-9_]+)'\s*=\s*@\{([^}]*)\}/g;
    let m;
    while ((m = entryRe.exec(aliasSec[1]))) {
      const map = {};
      const pairRe = /(\w+)\s*=\s*'(\w+)'/g;
      let pm;
      while ((pm = pairRe.exec(m[2]))) map[pm[1]] = pm[2];
      alias[m[1]] = map;
    }
  }

  const flatten = [];
  const flatSec = block.match(/\$StyleFlatten\s*=\s*@\(([\s\S]*?)\n\)/);
  if (flatSec) {
    const re = /'([A-Za-z0-9_]+)'/g;
    let m;
    while ((m = re.exec(flatSec[1]))) flatten.push(m[1]);
  }

  const flattenKeys = {};
  const fkSec = block.match(/\$FlattenKeys\s*=\s*@\{([\s\S]*?)\n\}/);
  if (fkSec) {
    const e2 = /'([A-Za-z0-9_]+)'\s*=\s*@\(([^)]*)\)/g;
    let m2;
    while ((m2 = e2.exec(fkSec[1]))) {
      const keys = [];
      const re = /'(\w+)'/g;
      let m3;
      while ((m3 = re.exec(m2[2]))) keys.push(m3[1]);
      flattenKeys[m2[1]] = keys;
    }
  }
  return { alias, flatten, flattenKeys };
}

/** 入口归一化表让哪些键"算被读过" */
function entryNormalizedKeys(act, norm) {
  const keys = new Set();
  const alias = norm.alias[act];
  if (alias) {
    for (const [from, to] of Object.entries(alias)) {
      keys.add(from);
      keys.add(to);
    }
  }
  if (norm.flatten.includes(act)) {
    const objKeys = norm.flattenKeys[act] || ['style'];
    for (const k of objKeys) keys.add(k);
  }
  return keys;
}

/**
 * 扫全部候选。返回 { rows, total }。
 * rows: {act, missing, file} —— missing 是 TS 传了但该 action 读不到的键
 * （已含：助手函数里读的键 + 入口归一化表里改过名/摊平过的键）
 */
function scanCandidates() {
  const ps = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8').replace(/^\uFEFF/, '');
  const { helpers, cases } = splitCases(ps);
  const norm = parseEntryNormalization(ps);
  const keyByAction = new Map();
  for (const c of cases) {
    const ks = effectiveKeys(c.body, helpers);
    for (const k of entryNormalizedKeys(c.name, norm)) ks.add(k);
    keyByAction.set(c.name, ks);
  }

  const rows = [];
  for (const site of tsCallSites()) {
    const known = keyByAction.get(site.act);
    if (!known) continue; // PS 里没这个 case：另一种病（Unknown action），不归本门禁管
    const missing = site.keys.filter((k) => !known.has(k));
    if (missing.length) rows.push({ act: site.act, missing: missing.join(','), file: site.file });
  }
  rows.sort((a, b) => a.act.localeCompare(b.act) || a.file.localeCompare(b.file));
  return {
    rows,
    helperCount: Object.keys(helpers).length,
    caseCount: cases.length,
    aliasCount: Object.keys(norm.alias).length,
    flattenCount: norm.flatten.length,
  };
}

module.exports = { scanCandidates, collectKeys, splitCases, parseEntryNormalization };
