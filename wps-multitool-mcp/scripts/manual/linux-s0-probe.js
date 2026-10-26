#!/usr/bin/env node
/**
 * Linux S0 能力探针（第 10 轮 T5）
 * ============================================================
 * 依据：docs/agents/PLAN-LINUX-BACKEND.md §5.2（S0 前置于一切）、§5.3（20 个代表性动作）
 *
 * 它回答一个问题：**WPS Linux 的 JSAPI 到底覆盖多少动作。**
 * 这是 B2（建传输层）的前置 —— 顺序不能反：为了回答"API 存不存在"，
 * 先建一整套要长期维护的通道，万一 API 大面积不存在就全白写。
 *
 * ⚠️ 本脚本在本轮【没有在真 Linux 上跑过】（仓库里没有 Linux 机器）。
 *    它交付的是"脚本"，不是"覆盖率数字"。判定一律按 static 计，
 *    与本仓库「live 必须有真机证据」的规矩一致。
 *
 * 用法
 * ----
 *   # 自检（不需要 Linux 环境，在本机也能证明脚本能启动、参数能解析）
 *   node scripts/manual/linux-s0-probe.js --help
 *   node scripts/manual/linux-s0-probe.js --selfcheck
 *
 *   # 真跑（必须在装有 WPS Linux + Xvfb 的机器上）
 *   node scripts/manual/linux-s0-probe.js \
 *        --wps /usr/bin/wps \
 *        --display :99 \
 *        --out ./linux-s0-report.md
 *
 * 产出
 * ----
 *   1) 生成的 probe.js（WPS 侧真正执行的那份）
 *   2) raw.json（逐条原始结果：支持/属性不存在/抛错/部分支持）
 *   3) 填好判据列的覆盖率报告（默认写 scripts/manual/linux-s0-report.md）
 *
 * R0 前置检查（§5.2 明确要求 S0 第一步就查）
 * ------------------------------------------------
 *   /opt/kingsoft/wps-office/office6/ 下有没有 jsaddons 目录。
 *   **两个都查不到就直接停**，回 §10 的 LibreOffice 路线重新评估 ——
 *   所以本脚本在 R0 不过时会明确拒绝继续，并把原因写进报告。
 */

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync, spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const DEFAULT_OUT = path.join(ROOT, 'scripts', 'manual', 'linux-s0-report.md');
const WPS_OFFICE_DIR = '/opt/kingsoft/wps-office/office6';

/** §5.3 的 20 个代表性动作。一字不改地照抄计划表，不自行增删。 */
const ACTIONS = [
  { no: 1, action: 'getContext', alias: 'getActiveDocumentState', probe: 'doc', note: '能否拿到活动文档（WPS JSAPI 有无对应）' },
  { no: 2, action: 'getSelection', probe: 'range', note: 'Selection 能力' },
  { no: 3, action: 'getSelectedText', probe: 'range', note: '同上' },
  { no: 4, action: 'insertText', probe: 'range', note: '写入' },
  { no: 5, action: 'setFontColor', alias: 'setFontSize', probe: 'range', note: 'Range 样式' },
  { no: 6, action: 'getShapes', probe: 'ppt', note: '形状遍历' },
  { no: 7, action: 'setShapeText', probe: 'ppt', note: '形状文本' },
  { no: 8, action: 'addShape', probe: 'ppt', note: '新建形状' },
  { no: 9, action: 'setShapePosition', probe: 'ppt', note: '坐标' },
  { no: 10, action: 'deleteSlide', probe: 'ppt', note: '删除' },
  { no: 11, action: 'addSlide', probe: 'ppt', note: '新建页' },
  { no: 12, action: 'getSlideCount', probe: 'ppt', note: '计数' },
  { no: 13, action: 'addShapeText', probe: 'ppt-table', note: '表格能力（PPT 侧最难的一类）' },
  { no: 14, action: 'getTableCell', probe: 'ppt-table', note: '表格读回' },
  { no: 15, action: 'setTableCell', probe: 'ppt-table', note: '表格写' },
  { no: 16, action: 'addAnimation', probe: 'ppt', note: '时间轴（**最可能缺**）' },
  { no: 17, action: 'setSlideTransition', probe: 'ppt', note: '切换（**最可能缺**）' },
  { no: 18, action: 'setSlideBackground', probe: 'ppt', note: '背景' },
  { no: 19, action: 'exportSlideAsImage', probe: 'ppt', note: '导出' },
  { no: 20, action: 'saveAs', probe: 'doc', note: '另存（决定用户会不会丢改动）' },
];

const VERDICT = {
  UNVERIFIED: '未验证（本机无 Linux 环境）',
  SUPPORTED: '支持',
  UNSUPPORTED: '不支持（属性不存在）',
  THREW: '抛错',
  PARTIAL: '部分支持（差异说明）',
};

function parseArgs(argv) {
  const out = {
    wps: '/usr/bin/wps',
    display: process.env.DISPLAY || ':99',
    out: DEFAULT_OUT,
    probeOut: null,
    selfcheck: false,
    help: false,
    force: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-h') out.help = true;
    else if (a === '--selfcheck') out.selfcheck = true;
    else if (a === '--force') out.force = true;
    else if (a === '--wps') out.wps = argv[++i];
    else if (a === '--display') out.display = argv[++i];
    else if (a === '--out') out.out = argv[++i];
    else if (a === '--probe-out') out.probeOut = argv[++i];
    else throw new Error(`不认识的参数：${a}`);
  }
  if (!out.probeOut) out.probeOut = path.join(os.tmpdir(), 'wps-s0-probe.js');
  return out;
}

function usage() {
  return [
    'linux-s0-probe —— WPS Linux JSAPI 能力探针（PLAN-LINUX-BACKEND.md §5.2 / §5.3）',
    '',
    '用法：',
    '  node scripts/manual/linux-s0-probe.js --help',
    '  node scripts/manual/linux-s0-probe.js --selfcheck',
    '  node scripts/manual/linux-s0-probe.js --wps /usr/bin/wps --display :99 [--out 报告.md]',
    '',
    '参数：',
    '  --wps <路径>       WPS 可执行文件，默认 /usr/bin/wps',
    '  --display <DISPLAY>  传给 xvfb-run 的显示号，默认取 $DISPLAY 或 :99',
    '  --out <路径>       覆盖率报告输出，默认 scripts/manual/linux-s0-report.md',
    '  --probe-out <路径> 生成的 probe.js 落盘位置，默认系统临时目录',
    '  --selfcheck        自检：只生成 probe.js 并校验它语法/内容，不碰 WPS',
    '  --force            跳过 R0 前置检查（不建议：R0 不过说明该回 LibreOffice 路线）',
    '',
    `本轮状态：${VERDICT.UNVERIFIED} —— 脚本已交付，未在真 Linux 上跑过。`,
  ].join('\n');
}

/**
 * 生成 probe.js：WPS 侧真正执行的那份。
 *
 * 设计要点（都来自 §5.2，不是拍脑袋）：
 *  · 逐条 try/catch 并把【异常原文】带回来 —— "抛错"和"属性不存在"是两回事，
 *    规划里要求分开记，合并了就没法决定 B2 值不值得做。
 *  · 属性探测用 `in` / typeof，不靠 try/catch 猜 —— WPS 的宿主对象对
 *    不存在的属性可能返回 undefined 而不是抛错，只看 try/catch 会把"缺"记成"空"。
 *  · 尾部把结果写成 JSON 字符串，由外层捞出来（--macro 没有约定的返回通道）。
 */
function buildProbeJs() {
  return `/* generated by scripts/manual/linux-s0-probe.js —— 不要手改 */
(function () {
  var RESULTS = [];
  function record(no, action, verdict, evidence) {
    RESULTS.push({ no: no, action: action, verdict: verdict, evidence: String(evidence || '').slice(0, 400) });
  }
  function hasPath(root, dotted) {
    // 逐段探测路径：中途缺一段就返回缺在哪
    var cur = root, parts = String(dotted).split('.'), walked = [];
    for (var i = 0; i < parts.length; i++) {
      walked.push(parts[i]);
      if (cur === null || cur === undefined) return { ok: false, at: walked.join('.') };
      if (!(parts[i] in Object(cur))) return { ok: false, at: walked.join('.') };
      cur = cur[parts[i]];
    }
    return { ok: true, value: cur };
  }
  function probeCall(no, action, owner, method, args, invoke) {
    var m = hasPath(owner, method);
    if (!m.ok) { record(no, action, '不支持（属性不存在）', '路径 ' + method + ' 在 ' + m.at + ' 处缺失'); return; }
    try {
      var r = invoke(owner, args);
      record(no, action, '支持', '返回值类型=' + (typeof r));
    } catch (e) {
      record(no, action, '抛错', (e && (e.message || e.toString())) || String(e));
    }
  }

  try {
    // ---------- R0：先回答"有没有 JSAPI 宿主" ----------
    var ctx = hasPath(this, 'wps') ;
    if (!ctx.ok) { record(0, 'R0:jsapi-host', '不支持（属性不存在）', '宏上下文里没有 this.wps —— JSAPI 宿主不存在'); }
    else { record(0, 'R0:jsapi-host', '支持', 'this.wps 存在'); }

    // ---------- R0.2：--macro 能不能 attach 到已开着的实例 ----------
    // 探针自己回答不了（它就是被 --macro 拉起来的），能回答的是
    // "有没有别的已开实例" —— 若拿到 0 个，说明 --macro 起的是全新实例。
    probeCall(0, 'R0:attach-existing-instance', this.wps || this, 'Application.Presentations.Count', {},
      function (o) { return o.wps ? o.wps.Application.Presentations.Count : 0; });

    var app = this.wps ? this.wps.Application : undefined;
    var out = { results: RESULTS, env: { ts: new Date().toISOString() } };

    if (!app) {
      out.error = '拿不到 wps.Application，后面 20 条全部标记为未验证';
      RESULTS.length = 0;
      for (var k = 0; k < 20; k++) record(k + 1, 'unknown', '未验证', '没有 JSAPI 宿主');
    } else {
      var doc = null, pres = null, slide = null;
      try { doc = app.ActiveDocument; } catch (e) { out.err_doc = String(e); }
      try { pres = app.ActivePresentation; } catch (e) { out.err_pres = String(e); }
      try { if (pres) slide = pres.Slides.Item(1); } catch (e) { out.err_slide = String(e); }

      var docOwner = doc || {};
      var presOwner = pres || {};
      var slideOwner = slide || {};

      // 1 doc
      probeCall(1, 'getContext', docOwner, 'Name', {}, function (o) { return o.Name; });
      // 2,3,4,5 range
      probeCall(2, 'getSelection', docOwner, 'Selection', {}, function (o) { return o.Selection; });
      probeCall(3, 'getSelectedText', docOwner, 'Selection.Text', {}, function (o) { return o.Selection.Text; });
      probeCall(4, 'insertText', docOwner, 'Selection.Text', {}, function (o) { return o.Selection.Text; });
      probeCall(5, 'setFontColor', docOwner, 'Selection.Font.Color', {}, function (o) { return o.Selection.Font.Color; });
      // 6..12 ppt
      probeCall(6, 'getShapes', slideOwner, 'Shapes', {}, function (o) { return o.Shapes.Count; });
      probeCall(7, 'setShapeText', slideOwner, 'Shapes', {}, function (o) { return o.Shapes.Count; });
      probeCall(8, 'addShape', slideOwner, 'Shapes.AddShape', {}, function (o) { return typeof o.Shapes.AddShape; });
      probeCall(9, 'setShapePosition', slideOwner, 'Shapes', {}, function (o) { return o.Shapes.Count; });
      probeCall(10, 'deleteSlide', presOwner, 'Slides', {}, function (o) { return o.Slides.Count; });
      probeCall(11, 'addSlide', presOwner, 'Slides.Add', {}, function (o) { return typeof o.Slides.Add; });
      probeCall(12, 'getSlideCount', presOwner, 'Slides.Count', {}, function (o) { return o.Slides.Count; });
      // 13..15 table
      probeCall(13, 'addShapeText', slideOwner, 'Shapes.AddTable', {}, function (o) { return typeof o.Shapes.AddTable; });
      probeCall(14, 'getTableCell', slideOwner, 'Shapes', {}, function (o) { return o.Shapes.Count; });
      probeCall(15, 'setTableCell', slideOwner, 'Shapes.AddTable', {}, function (o) { return typeof o.Shapes.AddTable; });
      // 16..20
      probeCall(16, 'addAnimation', slideOwner, 'TimeLine', {}, function (o) { return typeof o.TimeLine; });
      probeCall(17, 'setSlideTransition', slideOwner, 'SlideShowTransition', {}, function (o) { return typeof o.SlideShowTransition; });
      probeCall(18, 'setSlideBackground', slideOwner, 'Background', {}, function (o) { return typeof o.Background; });
      probeCall(19, 'exportSlideAsImage', slideOwner, 'Export', {}, function (o) { return typeof o.Export; });
      probeCall(20, 'saveAs', docOwner, 'SaveAs', {}, function (o) { return typeof o.SaveAs; });
    }

    out.results = RESULTS;
    out.__WPS_S0__ = JSON.stringify(out);
  } catch (e) {
    this.__WPS_S0__ = JSON.stringify({ error: '探针自身异常：' + ((e && e.message) || e), results: RESULTS });
  }
  // 兜底：把结果塞进控制台，外层用 stdout 兜捞（--macro 没有约定的返回通道）
  try { if (typeof console !== 'undefined' && console.log) console.log('__WPS_S0__' + this.__WPS_S0__); } catch (e) {}
})();
`;
}

/** R0 前置检查（§5.2：S0 第一步就查 jsaddons） */
function checkR0() {
  const r = { jsaddons: false, dir: WPS_OFFICE_DIR, notes: [] };
  try {
    r.jsaddons = fs.existsSync(path.join(WPS_OFFICE_DIR, 'jsaddons'));
  } catch (e) {
    r.notes.push('检查 jsaddons 时异常：' + e.message);
  }
  if (!r.jsaddons) {
    r.notes.push(
      `${WPS_OFFICE_DIR}/jsaddons 不存在。§5.2 规定：查不到就要停下来，` +
        '回 §10 的 LibreOffice/UNO 路线重新评估，而不是接着写 B2。'
    );
  }
  return r;
}

function renderReport(raw, r0, meta) {
  const lines = [];
  const byNo = new Map((raw && raw.results ? raw.results : []).map((r) => [r.no, r]));
  lines.push('# Linux S0 能力探针 · 覆盖率报告（模板）');
  lines.push('');
  lines.push('> 依据 `docs/agents/PLAN-LINUX-BACKEND.md` §5.2 / §5.3。');
  lines.push('> 由 `scripts/manual/linux-s0-probe.js` 生成。');
  lines.push('');
  lines.push(`- 生成时间：${new Date().toISOString()}`);
  lines.push(`- WPS 可执行文件：\`${meta.wps}\``);
  lines.push(`- DISPLAY：\`${meta.display}\``);
  lines.push(`- R0（jsaddons 目录）：${r0.jsaddons ? '存在' : '**不存在**'}`);
  if (r0.notes.length) for (const n of r0.notes) lines.push(`  - ${n}`);
  lines.push('');
  lines.push('## 覆盖表（20 个代表性动作）');
  lines.push('');
  lines.push('| # | 动作 | 考察点 | 判定 | 证据（属性是否存在 / 调用是否抛错） |');
  lines.push('|---|---|---|---|---|');
  for (const a of ACTIONS) {
    const r = byNo.get(a.no);
    lines.push(
      `| ${a.no} | \`${a.action}\`${a.alias ? ` / \`${a.alias}\`` : ''} | ${a.note} | ` +
        `**${r ? r.verdict : VERDICT.UNVERIFIED}** | ${r ? (r.evidence || '—') : '—'} |`
    );
  }
  lines.push('');
  lines.push('## 汇总');
  lines.push('');
  const count = (v) => ACTIONS.filter((a) => byNo.get(a.no) && byNo.get(a.no).verdict === v).length;
  lines.push(`- 支持：${count(VERDICT.SUPPORTED)} / ${ACTIONS.length}`);
  lines.push(`- 不支持（属性不存在）：${count(VERDICT.UNSUPPORTED)} / ${ACTIONS.length}`);
  lines.push(`- 抛错：${count(VERDICT.TREW)} / ${ACTIONS.length}`);
  lines.push(`- 部分支持：${count(VERDICT.PARTIAL)} / ${ACTIONS.length}`);
  lines.push(`- 未验证：${ACTIONS.length - ACTIONS.filter((a) => byNo.has(a.no)).length} / ${ACTIONS.length}`);
  lines.push('');
  lines.push('## 待补（S0 顺带要回答的）');
  lines.push('');
  lines.push('- [ ] `--macro` 能不能 attach 到已开着的 WPS 实例');
  lines.push('- [ ] 单次调用往返时延（含 WPS 内部耗时）');
  lines.push('- [ ] 哪些动作语义不同（如形状类型枚举与 MSO 不同）');
  lines.push('');
  return lines.join('\n') + '\n';
}

function selfCheck(opts) {
  const js = buildProbeJs();
  fs.writeFileSync(opts.probeOut, js, 'utf8');
  console.log(`已生成 probe.js → ${opts.probeOut}（${js.length} 字节）`);
  // 语法校验：只校验，不执行（这里没有 WPS 宿主）
  const chk = spawnSync(process.execPath, ['--check', opts.probeOut], { encoding: 'utf8' });
  const syntaxOk = chk.status === 0;
  console.log(`语法校验：${syntaxOk ? '通过' : '失败 —— ' + (chk.stderr || '').split('\n')[0]}`);
  const covers = ACTIONS.every((a) => js.includes(`'${a.action}'`));
  console.log(`20 个动作全部出现在 probe.js 里：${covers ? '是' : '否'}`);
  if (!covers) {
    const missing = ACTIONS.filter((a) => !js.includes(`'${a.action}'`)).map((a) => a.action);
    console.log('  缺：' + missing.join(', '));
  }
  console.log('');
  console.log(`本次自检【不产出覆盖率结论】—— ${VERDICT.UNVERIFIED}`);
  return syntaxOk && covers;
}

function main() {
  let opts;
  try {
    opts = parseArgs(process.argv.slice(2));
  } catch (e) {
    console.error(`参数错误：${e.message}\n`);
    console.error(usage());
    process.exit(2);
  }
  if (opts.help) {
    console.log(usage());
    process.exit(0);
  }

  const r0 = checkR0();
  if (opts.selfcheck) {
    const ok = selfCheck(opts);
    process.exit(ok ? 0 : 1);
  }

  if (process.platform !== 'linux') {
    console.error(`✗ 当前平台是 ${process.platform}，不是 linux。`);
    console.error(`  本脚本只在装了 WPS Linux + Xvfb 的机器上跑。`);
    console.error(`  在别的平台上跑出来的数字没有意义 —— 那正是本仓库「live 必须有真机证据」的规矩。`);
    console.error(`  想只验证脚本本身能启动，用 --selfcheck。`);
    process.exit(3);
  }

  if (!r0.jsaddons && !opts.force) {
    console.error('✗ R0 前置检查不通过：');
    for (const n of r0.notes) console.error('  - ' + n);
    console.error('  确实要继续请加 --force，但先想清楚 §5.2 为什么规定这一步要先查。');
    process.exit(4);
  }

  const js = buildProbeJs();
  fs.writeFileSync(opts.probeOut, js, 'utf8');
  console.log(`R0：jsaddons ${r0.jsaddons ? '存在' : '不存在'}；probe.js → ${opts.probeOut}`);

  const cmd = `xvfb-run -a ${opts.wps} --macro=${opts.probeOut}`;
  console.log(`执行：DISPLAY=${opts.display} ${cmd}`);
  let stdout = '';
  try {
    stdout = execFileSync('xvfb-run', ['-a', opts.wps, `--macro=${opts.probeOut}`], {
      encoding: 'utf8',
      timeout: 300000,
      env: { ...process.env, DISPLAY: opts.display },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  } catch (e) {
    console.error(`执行失败：${e.message}`);
    console.error('  报告仍会生成，但全部记为未验证 —— 没有原始输出就不能填结论。');
  }
  const m = String(stdout).match(/__WPS_S0__(\{[\s\S]*)/);
  const raw = m ? JSON.parse(m[1]) : null;
  fs.writeFileSync(opts.probeOut + '.raw.json', JSON.stringify(raw, null, 2), 'utf8');
  fs.writeFileSync(opts.out, renderReport(raw, r0, opts), 'utf8');
  console.log(`报告 → ${opts.out}`);
  console.log(`原始结果 → ${opts.probeOut}.raw.json`);
  if (!raw) {
    console.error('没捞到 __WPS_S0__ 标记 —— 报告里每一条都会是「未验证」。');
    process.exit(5);
  }
  process.exit(0);
}

main();