/**
 * Input: wps-multitool-addon-linux/ 落仓后的 actions-sync 五件 + 三型接线
 * Output: 0（字节指纹/规格计数/双模型分支/接线结构全过）| 1（逐条列出）
 * Pos: Linux 落仓轮门禁（B5-T4 成果收进主仓，2026-11-08）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 三道闸（对应落仓卡"逐字节等价"红线）：
 *   A. 字节指纹钉 —— 落仓五件的 SHA-256 必须 == 云端 B5-T4 真机验证版指纹（2003.md §六）。
 *      指纹写死在本文件：任何人改动这些文件（哪怕一个字节），FAST 即红——
 *      「云端 123/123 是唯一真机证据，改一行就作废」由本闸机械化。
 *   B. 规格与双模型分支 —— vm 沙箱加载 runtime+三型规格：动作计数 word13/ppt15/excel95；
 *      model-detect 无门面 → async（不覆盖注册表）/ 有门面 → sync（覆盖注册表）。
 *   C. 接线结构 —— 三型 index.html 的三行 script（顺序+位置）与 boot.js 的 apply 行。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const vm = require('vm');

const ADDON = path.resolve(__dirname, '..', '..', '..', 'wps-multitool-addon-linux');
let pass = 0, fail = 0;
const problems = [];
const ok = (v, msg) => { if (v) pass++; else { fail++; problems.push(msg); } };

/** 云端 B5-T4 真机验证版指纹（2003.md §六；改这些文件 = 作废真机证据） */
const FINGERPRINTS = {
  'js/actions-sync/runtime.js': '2a01a443cfb37b10348feafce6922091043eaa6cacac57ae41423ee475852b6e',
  'js/actions-sync/word.js':    '5148ec9bfcb1fb4490b6ff4ed6e549183e4aae7f4e1ee83c1a9cdb8148522e9b',
  'js/actions-sync/ppt.js':     '364f07e2a7c653c617bd38ed92a192ab4e3da93851662d807183f1cf7d2fc4e1',
  'js/actions-sync/excel.js':   '22db319e5c9c772be497f883049be3feef4789a0e7825609a570d22f5c18aae7',
  'js/model-detect.js':         'ba5f1c13dbf8919c22091a62a5dc2fdbf511e674b1b9f1e76f6ecec326e6213a',
};

// ── A. 字节指纹钉 ──
const sources = {};
for (const [rel, sha] of Object.entries(FINGERPRINTS)) {
  const p = path.join(ADDON, rel);
  ok(fs.existsSync(p), 'A 缺文件: ' + rel);
  if (!fs.existsSync(p)) continue;
  const buf = fs.readFileSync(p);
  const real = crypto.createHash('sha256').update(buf).digest('hex');
  ok(real === sha, 'A 指纹不符: ' + rel + ' 现=' + real.slice(0, 12) + '… 期望=' + sha.slice(0, 12) + '…（落仓后被人改动？）');
  sources[rel] = buf.toString('utf8');
}

// ── B. 规格计数与双模型分支（vm 沙箱，window 共享） ──
function loadAddon(sandboxWindow, files) {
  const ctx = vm.createContext({ window: sandboxWindow, Math, JSON, Object, Array, String, Number, Boolean, Date, Error, TypeError, Promise, parseInt, parseFloat, isNaN, console: { log: () => {} }, XMLHttpRequest: function () { this.open = () => {}; this.setRequestHeader = () => {}; this.send = () => {}; } });
  for (const rel of files) vm.runInContext(sources[rel], ctx, { filename: rel });
  return sandboxWindow;
}

try {
  const win = { };
  loadAddon(win, ['js/actions-sync/runtime.js', 'js/actions-sync/word.js', 'js/actions-sync/ppt.js', 'js/actions-sync/excel.js', 'js/model-detect.js']);
  const counts = {
    word: win.WpsAddonActionsSyncWord ? Object.keys(win.WpsAddonActionsSyncWord).length : -1,
    ppt: win.WpsAddonActionsSyncPpt ? Object.keys(win.WpsAddonActionsSyncPpt).length : -1,
    excel: win.WpsAddonActionsSyncExcel ? Object.keys(win.WpsAddonActionsSyncExcel).length : -1,
  };
  ok(counts.word === 13, 'B Word 同步动作数应 13，实得 ' + counts.word);
  ok(counts.ppt === 15, 'B PPT 同步动作数应 15，实得 ' + counts.ppt);
  ok(counts.excel === 95, 'B Excel 同步动作数应 95，实得 ' + counts.excel);
  ok(Object.values(win.WpsAddonActionsSyncWord).every(f => typeof f === 'function'), 'B Word 动作应全为函数');

  // 双模型分支①：无门面 → async，不覆盖注册表
  const winAsync = {};
  loadAddon(winAsync, ['js/actions-sync/runtime.js', 'js/actions-sync/excel.js', 'js/model-detect.js']);
  winAsync.WpsAddonActionsExcel = { keepMe: () => {} };   // 模拟 11.x 异步注册表
  const modeAsync = winAsync.WpsAddonModelDetect.apply('et');
  ok(modeAsync === 'async', 'B 无门面应判 async，实得 ' + modeAsync);
  ok(winAsync.WpsAddonActionsExcel && typeof winAsync.WpsAddonActionsExcel.keepMe === 'function', 'B async 分支不得覆盖 11.x 注册表');

  // 双模型分支②：有门面+可读活动对象 → sync，覆盖注册表
  const winSync = {};
  winSync.Application = { ActiveWorkbook: { ActiveSheet: {} }, Workbooks: { Count: 1 } };
  loadAddon(winSync, ['js/actions-sync/runtime.js', 'js/actions-sync/excel.js', 'js/model-detect.js']);
  winSync.WpsAddonActionsExcel = { keepMe: () => {} };
  const modeSync = winSync.WpsAddonModelDetect.apply('et');
  ok(modeSync === 'sync', 'B 有门面应判 sync，实得 ' + modeSync);
  ok(winSync.WpsAddonActionsExcel && typeof winSync.WpsAddonActionsExcel.getActiveWorkbook === 'function', 'B sync 分支应把注册表替换为同步实现');

  // unsupported 分支：S1 判"不支持"的规格显式回 unsupported，不假装成功
  const winU = {};
  winSync.Application && (winU.Application = winSync.Application);
  loadAddon(winU, ['js/actions-sync/runtime.js']);
  const RT = winU.WpsAddonSyncRuntime;
  const unspec = RT.makeAction({ v: '不支持（属性不存在）', jsapi: 'x', evidence: 'S1 证据', code: 'return 1;' });
  unspec({}).then(r => {
    ok(r && r.unsupported === 'S1:不支持（属性不存在）', 'B unsupported 分支应显式回 S1:unsupported');
    finish();
  }).catch(e => { ok(false, 'B unsupported 分支异常: ' + e.message); finish(); });
} catch (e) {
  ok(false, 'B 沙箱加载异常: ' + e.message);
  finish();
}

function finish() {
  // ── C. 接线结构 ──
  for (const [type, app, syncFile] of [['wps', 'word', 'word.js'], ['et', 'excel', 'excel.js'], ['wpp', 'ppt', 'ppt.js']]) {
    const ih = fs.readFileSync(path.join(ADDON, 'addon-' + type, 'index.html'), 'utf8');
    const iActions = ih.indexOf(`../js/actions/${app}.js`);
    const iMd = ih.indexOf('../js/model-detect.js');
    const iRt = ih.indexOf('../js/actions-sync/runtime.js');
    const iSync = ih.indexOf(`../js/actions-sync/${syncFile}`);
    const iCh = ih.indexOf('../js/channel-client.js');
    ok(iActions >= 0 && iMd > iActions && iRt > iMd && iSync > iRt && iCh > iSync,
      'C addon-' + type + '/index.html 三行 script 顺序/位置不对');
    const boot = fs.readFileSync(path.join(ADDON, 'addon-' + type, 'js', 'boot.js'), 'utf8');
    const iBoot = boot.indexOf('function boot() {');
    const iApply = boot.indexOf(`WpsAddonModelDetect.apply('${type}')`);
    ok(iBoot >= 0 && iApply > iBoot, 'C addon-' + type + '/boot.js apply 行缺失或位置不对');
  }

  console.log(`=== actions-sync 落仓门禁: ${pass} 通过 / ${fail} 失败 ===`);
  if (problems.length) for (const p of problems) console.error('  - ' + p);
  process.exit(fail ? 1 : 0);
}
