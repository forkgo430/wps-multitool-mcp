/**
 * Input: 动作名
 * Output: 该动作能否被安全取消
 * Pos: 取消机制的安全红线。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 取消的两种时机（差别很大，别混为一谈）
 * --------------------------------------
 *   A. 【排队中】请求还没写进 PowerShell —— 任何动作都能取消，删掉队列项即可。
 *      请求从未执行，既不产生副作用，也不需要重放。这是最有用的一档。
 *   B. 【已在途】请求已经写进去，WPS 正在跑 —— **取消不了**。
 *      COM 是 STA，PowerShell 那边正在等一个 COM 调用返回，我们没有安全的中断手段
 *      （杀进程会让 WPS 处于半完成状态，那正是 wps_common_repair 要收拾的烂摊子）。
 *      所以这一档能做的只有【放弃等待】：调用方立刻返回，子进程继续跑完，
 *      回话到了直接丢掉。
 *
 * B 档为什么还要分只读 / 写入
 * ---------------------------
 *   只读动作放弃等待：没有副作用，丢掉回话零风险。
 *   写入动作放弃等待：调用方会以为"没做成"，实际它【可能已经做成了一半】
 *   —— 于是要么重放（变成执行两次），要么以为失败却已经生效。两种都比等着更糟。
 *   所以写入动作【拒绝放弃】，让它老实跑完、如实返回结果。
 *
 * 一句话：能取消的只有"还没开始"的；已经开始的一律不打断，只读的可以不等，写入的不许不等。
 */

/** 取消动作时抛这个。与业务失败区分开 —— 它不是 WPS 出错，是我们主动不干了。 */
export class CancelError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CancelError';
  }
}

/** 一眼就是只读的动作前缀 */
const READ_ONLY_PREFIXES = ['get', 'list', 'read', 'find', 'query', 'probe', 'ping', 'check', 'caps', 'diagnose'];

/**
 * 名字带 find/探测味、但其实会改文档的动作 —— 必须从只读里单独挖出来。
 * find_replace 是真例子：它以 find 开头（看起来像查找），实际是替换内容。
 * 误判成只读的后果：允许"放弃等待"，调用方以为没改成、内容其实改了。
 */
const WRITE_OVERRIDE = new Set([
  'findReplace',
  'find_replace',
]);

/** 名字看不出来但实际只读的（探测类内建动作 + 少数历史命名） */
const READ_ONLY_EXTRA = new Set([
  '__probeParams',
  '__probeDoc',
  '__probeCaps',
  '__probeDialogs',
  '__probeStyleAdd',
  '__probeClose',
  // 时序诊断动作：只是占住通道，无任何副作用（配合 __sleep 对照，见 scripts/wps-com.ps1）
  '__sleepRead',
  '__bridgeStatus',
  'wpsStatus',
  // 精确项而不是前缀：statusBar（设置状态栏文字）是写入，加前缀会误伤它
  'status',
  'getAppInfo',
  'getOpenDocuments',
  'getOpenWorkbooks',
  'getOpenPresentations',
  'proofreadBasic',
  'proofread_basic',
  'evaluateFormula',
  'evaluate_formula',
  'diagnoseFormula',
  'diagnose_formula',
]);

/** 这些前缀【一定】是写入，优先级高于上面的猜测（防止 getXxx 里混进写操作） */
const WRITE_PREFIXES = ['set', 'add', 'insert', 'delete', 'remove', 'update', 'modify', 'save', 'create', 'close', 'quit', 'convert', 'apply', 'merge', 'split', 'move', 'copy', 'paste', 'replace', 'clear', 'hide', 'show', 'protect', 'unprotect', 'sort', 'fill', 'group', 'mark', 'accept', 'reject', 'enable', 'disable', 'start', 'switch', 'open', 'export', 'import', 'rename', 'rotate', 'duplicate', 'wrap', 'lock'];

/**
 * 下划线 ↔ 驼峰互转。
 * 为什么需要：同一个动作在两处是两种拼法 —— PS 侧的 case 标签是驼峰
 * （`evaluateFormula`），MCP 工具名是下划线（`wps_excel_evaluate_formula`）。
 * 两边各判一次、任一命中即生效，否则同一个动作会在两处得出相反结论。
 */
function camel(s: string): string {
  return s.replace(/_([a-z0-9])/g, (_m, c: string) => c.toUpperCase());
}
function snake(s: string): string {
  return s.replace(/([a-z0-9])([A-Z])/g, (_m, a: string, b: string) => `${a}_${b.toLowerCase()}`);
}

/** 该动作是不是只读。
 * 判定顺序：显式写入覆盖 → 显式只读白名单 → 写入前缀（一票否决）→ 只读前缀 → 默认【否】。
 * 默认否是刻意的：判不出来就按"不能安全放弃等待"处理，宁可让调用方多等一会儿。
 */
export function isReadOnlyAction(action: string): boolean {
  if (!action) return false;
  const forms = [action, camel(action), snake(action)];
  if (forms.some((f) => WRITE_OVERRIDE.has(f))) return false;
  if (forms.some((f) => READ_ONLY_EXTRA.has(f))) return true;

  const a = action.charAt(0).toLowerCase() + action.slice(1);
  for (const p of WRITE_PREFIXES) {
    if (a.startsWith(p)) return false;
  }
  for (const p of READ_ONLY_PREFIXES) {
    if (a.startsWith(p)) return true;
  }
  return false;
}

/**
 * 会造成不可逆后果（删内容、丢未保存改动、原地改格式）的动作。
 * 给 MCP 的 destructiveHint 用 —— 判据是"用户会不会想先看一眼再确认"，
 * 所以不包括普通写操作（设个字号、改个公式不算破坏）。
 */
const DESTRUCTIVE_PREFIXES = ['delete', 'remove', 'clear', 'replace', 'convert', 'drop', 'kill', 'wipe', 'purge'];
const DESTRUCTIVE_EXACT = new Set([
  'closeDocument', 'closeWorkbook', 'closePresentation',
  'closeDocumentWithoutSave', 'closeWithoutSave',
  'quitApplication', 'quitApp',
  'saveAs', 'saveAsDocument', 'saveAsWorkbook', 'saveAsPresentation',
  'close', 'quit',
  // quitWps 退出的是用户开着的全部 WPS，且历史上伴随全量 taskkill —— 影响面远超动作名字面
  'quitWps',
  // handleRevisions：接受/拒绝修订会整块增删内容，修订记录处理掉就没了 —— 用户会想先看一眼
  'handleRevisions',
  // undoSteps：撤销 = 把最近的编辑丢掉，连调几次会连退几步（非幂等）—— 撤销前必须留快照
  'undoSteps',
]);

export function isDestructiveAction(action: string): boolean {
  if (!action) return false;
  const forms = [action, camel(action), snake(action)];
  if (forms.some((f) => DESTRUCTIVE_EXACT.has(f))) return true;
  const a = action.charAt(0).toLowerCase() + action.slice(1);
  return DESTRUCTIVE_PREFIXES.some((p) => a.startsWith(p));
}

/**
 * 重复执行结果是否一致（幂等）。
 * 只读天然幂等；写入里"设/改/存"也是（重复设同一个值无变化），
 * 但"插入/追加/合并"每跑一次就多一份，不能重放。
 */
const NON_IDEMPOTENT_PREFIXES = ['insert', 'add', 'append', 'copy', 'paste', 'merge', 'split', 'duplicate', 'import', 'generate', 'create', 'start', 'move'];

export function isIdempotentAction(action: string): boolean {
  if (!action) return false;
  if (isReadOnlyAction(action)) return true;
  if (isDestructiveAction(action)) return false;
  const forms = [action, camel(action), snake(action)];
  if (forms.some((f) => WRITE_OVERRIDE.has(f))) return false;
  const a = action.charAt(0).toLowerCase() + action.slice(1);
  if (NON_IDEMPOTENT_PREFIXES.some((p) => a.startsWith(p))) return false;
  return true; // set / update / modify / save / format 之类
}
