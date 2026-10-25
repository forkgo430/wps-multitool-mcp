/**
 * Input: 活动文档状态（app / 路径 / 是否已保存）+ 快照文件路径
 * Output: 快照的创建 / 列出 / 恢复，以及一条可复核的操作日志
 * Pos: 写入的"保命层"。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ------------
 * COM 不是事务：一个 batch 三步里第二步失败，第一步的改动已经落地了，没有回滚。
 * 以前的对策写在 skill 里 ——"动文档前让用户去面板存一份快照"。这是把安全责任
 * 推给用户，实际执行时没人会做。
 *
 * 这里做的是【补偿】而不是【原子】：不假装能把 COM 变成事务，只保证改坏了有一份
 * 可以退回去的副本。所以有三条规矩：
 *   1) 快照是文件级复制，不碰文档内容、不触发 WPS 弹窗（COM 是 STA，一弹框就卡通道）
 *   2) 恢复必须显式 confirm；默认写到新文件，覆盖原文件要再加 overwrite
 *   3) 未保存过的文档（没有 Path）直接报错让调用方先另存为 ——
 *      给一个"临时路径"存快照是在自欺欺人：那份文件根本不是用户正在编辑的东西
 *      注意判据是 Path 不是 FullName：未保存文档的 FullName 非空（等于文档名），
 *      拿 FullName 判会把"没存过盘"误判成"有文件"，然后去复制一个不存在的路径。
 *
 * 未保存改动（dirty）
 * ------------------
 * 磁盘上的文件落后于用户屏幕上的内容时，复制出来的快照【不含】未保存改动。
 * 这种情况不静默通过：记录里带 dirty:true，工具结果里明说"快照是磁盘版本"。
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { RUNTIME_ID } from '../../runtime-identity';

const HOME = process.env.HOME || process.env.USERPROFILE || os.homedir();
/** 与日志同根：~/.wps-multitool-mcp/{snapshots,oplog} */
export const SNAPSHOT_ROOT = path.join(HOME, `.${RUNTIME_ID}`, 'snapshots');
const OPLOG_DIR = path.join(HOME, `.${RUNTIME_ID}`, 'oplog');
const OPLOG_FILE = path.join(OPLOG_DIR, 'ops.jsonl');
/** 每个原文件最多留几份。20 份足够覆盖"改错了想退两步"，再多就是占盘 */
const KEEP_PER_DOC = 20;

export interface DocState {
  app: string;
  name: string;
  fullName: string;
  /** 文档所在目录。未保存的文档这里是空串 —— 判"有没有盘上文件"只能看它 */
  dir?: string;
  saved: boolean;
  readOnly: boolean;
  hasPath: boolean;
}

export interface SnapshotRecord {
  ts: string;
  kind: 'snapshot' | 'restore' | 'batch' | 'prune';
  snapshotPath?: string;
  originalPath?: string;
  app?: string;
  name?: string;
  /** 原文件在磁盘上落后于屏幕内容（有未保存改动） */
  dirty?: boolean;
  size?: number;
  label?: string;
  tool?: string;
  [k: string]: unknown;
}

const text = (t: string, ok = true, extra: Partial<ToolCallResult> = {}): ToolCallResult => ({
  id: uuidv4(),
  success: ok,
  content: [{ type: 'text', text: t }],
  ...extra,
});

const fail = (t: string, error?: string): ToolCallResult =>
  text(t, false, { error: error ?? t });

/**
 * oplog 超过这个大小就轮转：把当前的 ops.jsonl 改名成 ops.1.jsonl，
 * 只留一份旧的，更老的直接删。
 *
 * 为什么必须有上限：这是追加写的操作日志，COM 没有事务，快照又是"改前保命"的手段 ——
 * 用户用得越勤它越长，而它只被 readOpLog 读来列举快照，留 20 年的记录没有意义。
 * 10 MB 按每条 ~200 字节算约 5 万条，远超"退两步"的需要。
 */
const OPLOG_MAX_BYTES = 10 * 1024 * 1024;

/** 轮转：ops.jsonl → ops.1.jsonl（更老的直接删）。失败就放弃轮转，不阻断业务 */
function rotateOpLogIfNeeded(): void {
  try {
    if (!fs.existsSync(OPLOG_FILE)) return;
    if (fs.statSync(OPLOG_FILE).size < OPLOG_MAX_BYTES) return;
    const old = path.join(OPLOG_DIR, 'ops.1.jsonl');
    if (fs.existsSync(old)) fs.rmSync(old, { force: true });
    fs.renameSync(OPLOG_FILE, old);
  } catch {
    /* 轮转失败不该影响业务：接着往原文件里写就行 */
  }
}

/** 追加一条操作日志。日志写不了不该把业务打断 —— 它是证据，不是前置条件 */
export function appendOpLog(entry: SnapshotRecord): void {
  try {
    fs.mkdirSync(OPLOG_DIR, { recursive: true });
    rotateOpLogIfNeeded();
    fs.appendFileSync(OPLOG_FILE, `${JSON.stringify(entry)}\n`, 'utf8');
  } catch {
    /* 磁盘满 / 权限不足时静默降级 */
  }
}

/** 读操作日志（坏行直接跳过：日志是追加写的，进程被杀可能留半行） */
export function readOpLog(): SnapshotRecord[] {
  try {
    if (!fs.existsSync(OPLOG_FILE)) return [];
    const out: SnapshotRecord[] = [];
    for (const line of fs.readFileSync(OPLOG_FILE, 'utf8').split('\n')) {
      if (!line.trim()) continue;
      try {
        out.push(JSON.parse(line) as SnapshotRecord);
      } catch {
        /* 半行不致命 */
      }
    }
    return out;
  } catch {
    return [];
  }
}

/**
 * 取活动文档状态。只读动作，不会拉起没运行的 WPS（PS 侧用 GetActiveWpsApp）。
 *
 * app 提示不是可有可无的：三个应用可能同时开着，而 ActiveDocument 是弱引用。
 * 不指定时 PS 只能猜（跳过空壳、优先未保存的那个），猜错就会把别人的文件当快照源 ——
 * 实测踩过一次：用户正在 Word 里写，快照却抓了 Excel 里一个残留工作簿。
 * 调用方知道自己在哪个应用时，务必传进来。
 */
async function getActiveState(
  app?: 'word' | 'excel' | 'ppt'
): Promise<{ ok: true; state: DocState } | { ok: false; error: string }> {
  const r = await wpsClient.executeMethod<DocState>(
    'getActiveDocumentState',
    app ? { app } : {}
  );
  if (!r.success) return { ok: false, error: r.error || '拿不到活动文档状态' };
  const data = (r.data ?? undefined) as DocState | undefined;
  if (!data) return { ok: false, error: '拿不到活动文档状态（返回空）' };
  return { ok: true, state: data };
}

/** 时间戳：yyyyMMdd-HHmmss-SSS，排序友好且不会撞 */
function stamp(): string {
  const d = new Date();
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  return (
    `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-` +
    `${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}-${p(d.getMilliseconds(), 3)}`
  );
}

/** 原文件名塞进快照名时去掉非法字符，同时加路径哈希防同名不同目录相撞 */
function snapshotFileName(state: DocState): string {
  const base = path.basename(state.fullName);
  const safe = base.replace(/[\\/:*?"<>|]/g, '_');
  const hash = Buffer.from(state.fullName.toLowerCase(), 'utf8').toString('hex').slice(0, 8);
  return `${stamp()}_${state.app}_${hash}_${safe}`;
}

/**
 * 按原文件清理旧快照，只保留最近 KEEP_PER_DOC 份。
 * 时间排序而不是依赖日志追加顺序：日志可能被并行进程交错写入，顺序不可靠。
 * 只删 SNAPSHOT_ROOT 里的文件 —— 一旦解析出的路径跑出这个目录，宁可不删。
 */
/**
 * 文档的规范键：先解析真实路径（把 8.3 短名、不同大小写统一掉），再小写比较。
 *
 * 为什么不能直接比字符串：同一个文件在"另存为当时"和"重新打开之后"，WPS 给出的
 * FullName 可能不是一个拼法（实测短名/大小写都会变）。按字符串比会把同一份文档拆成
 * 两个键 —— 表现为快照数量越攒越多，保留策略形同虚设。
 */
function docKey(p: string): string {
  try {
    return fs.realpathSync.native(p).toLowerCase();
  } catch {
    return path.resolve(p).toLowerCase();
  }
}

function prune(originalPath: string): number {
  const root = path.resolve(SNAPSHOT_ROOT) + path.sep;
  const target = docKey(originalPath);
  let removed = 0;

  // 每轮重新读日志再数一遍，而不是"一次算好该删几条"：
  // 单次计算时只要有一条 existsSync 短暂为 false（杀软扫描、刚写入还没可见），
  // 该删的条数就会算少。这里删到数够为止。
  for (let round = 0; round < 5; round++) {
    const live = readOpLog()
      .filter(
        (r) =>
          r.kind === 'snapshot' &&
          r.snapshotPath &&
          r.originalPath &&
          docKey(r.originalPath) === target &&
          fs.existsSync(r.snapshotPath)
      )
      .sort((a, b) => String(b.ts).localeCompare(String(a.ts)));

    if (live.length <= KEEP_PER_DOC) break;

    let progressed = false;
    for (const r of live.slice(KEEP_PER_DOC)) {
      const p = path.resolve(r.snapshotPath as string);
      if (!p.startsWith(root)) continue; // 防手滑删到用户文件
      try {
        fs.unlinkSync(p);
        removed++;
        progressed = true;
      } catch {
        /* 删不掉：下一轮再看，超过 5 轮就放弃（说明是真的被占用） */
      }
    }
    if (!progressed) break;
  }
  return removed;
}

/**
 * 建一份快照。给工具用，也给 batch 用（batch 直接调它，不走工具层，避免递归）。
 * 返回失败时一定带可操作的原因 —— "快照失败"这四个字对调用方没有任何用处。
 */
export async function createSnapshot(label?: string, app?: 'word' | 'excel' | 'ppt'): Promise<
  | { ok: true; record: SnapshotRecord; dirty: boolean }
  | { ok: false; error: string; code: 'no_document' | 'no_path' | 'io' }
> {
  const st = await getActiveState(app);
  if (!st.ok) return { ok: false, error: st.error, code: 'no_document' };
  const { state } = st;

  // hasPath 来自 COM 的 Path 属性；再加一次 fs 存在性检查做兜底 ——
  // 见过未保存文档的 FullName 拼出"默认目录\文字文稿1"（文件根本不存在），
  // 直接 copy 只会给一句 ENOENT，对调用方毫无指导意义。
  if (!state.hasPath || !state.fullName || !fs.existsSync(state.fullName)) {
    return {
      ok: false,
      code: 'no_path',
      error:
        `[勿重试·先另存为] 文档「${state.name}」还没有磁盘文件（从未保存过），没法建快照。` +
        '先 wps_common_save_as 存一次，再重试。',
    };
  }

  try {
    fs.mkdirSync(SNAPSHOT_ROOT, { recursive: true });
    const dest = path.join(SNAPSHOT_ROOT, snapshotFileName(state));
    fs.copyFileSync(state.fullName, dest);
    const size = fs.statSync(dest).size;
    const record: SnapshotRecord = {
      ts: new Date().toISOString(),
      kind: 'snapshot',
      snapshotPath: dest,
      originalPath: state.fullName,
      app: state.app,
      name: state.name,
      dirty: !state.saved,
      size,
      label,
      tool: 'wps_common_snapshot_create',
    };
    appendOpLog(record);
    prune(state.fullName);
    return { ok: true, record, dirty: !state.saved };
  } catch (e) {
    return {
      ok: false,
      code: 'io',
      error: `复制失败：${e instanceof Error ? e.message : String(e)}（源文件可能被占用或路径不可读）`,
    };
  }
}

/**
 * 覆盖文件时的重试：WPS 关掉文档后，进程仍可能握着句柄一小会儿（实测 EBUSY）。
 * 退避 8 次（约 7 秒）足够它松手；再不行就是真占用了，交给调用方去彻底退出 WPS。
 */
async function copyWithRetry(src: string, dest: string, tries = 8): Promise<void> {
  for (let i = 1; i <= tries; i++) {
    try {
      fs.copyFileSync(src, dest);
      return;
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code || '';
      const retryable = code === 'EBUSY' || code === 'EPERM' || code === 'EACCES';
      if (!retryable || i === tries) throw e;
      await new Promise((r) => setTimeout(r, 200 * i));
    }
  }
}

// ==================== 工具定义 ====================

export const snapshotCreateDefinition: ToolDefinition = {
  name: 'wps_common_snapshot_create',
  description: `给当前文档存一份快照（复制磁盘文件到 ~/.${RUNTIME_ID}/snapshots）。

**什么时候用**：任何"要连续改好几处、改错了不好还原"的写入之前。
wps_common_batch 在检测到破坏性步骤时会自动调它，单独调用一般是为了打一个"锚点"。

**它会如实告诉你的两件事**：
- 文档从未保存过 → 报 [勿重试·先另存为]，不会拿临时路径糊弄你
- 文档有未保存改动 → 快照是【磁盘版本】，不含屏幕上还没存的内容（结果里 dirty=true）

不做 COM 写入，不触发任何 WPS 弹窗。`,
  category: ToolCategory.COMMON,
  group: 'common',
  inputSchema: {
    type: 'object',
    properties: {
      label: { type: 'string', description: '给这份快照起个名字，例如 before-restyle（可省略）' },
      app: {
        type: 'string',
        enum: ['word', 'excel', 'ppt'],
        description:
          '要快照哪个应用里当前打开的文档。三个应用都开着时强烈建议给，' +
          '否则只能按"有未保存改动的优先"猜，可能抓错文件（可省略）',
      },
    },
  },
};

export const snapshotCreateHandler: ToolHandler = async (args) => {
  const label = typeof args.label === 'string' ? args.label : undefined;
  const appArg = args.app === 'word' || args.app === 'excel' || args.app === 'ppt' ? args.app : undefined;
  const r = await createSnapshot(label, appArg);
  if (!r.ok) return fail(r.error, r.code);

  const lines = [
    `已建快照：${r.record.snapshotPath}`,
    `原文档：${r.record.originalPath}（${Math.round((r.record.size ?? 0) / 1024)} KB）`,
  ];
  const warnings = [];
  if (r.dirty) {
    lines.push('注意：文档有未保存改动，这份快照是磁盘上的版本，不含还没保存的内容。');
    warnings.push({
      step: 'snapshot.dirty',
      message: '快照取自磁盘，不包含未保存改动；要连未保存内容一起保命，先 wps_common_save 再建快照',
    });
  }
  return text(lines.join('\n'), true, warnings.length ? { warnings } : {});
};

export const snapshotListDefinition: ToolDefinition = {
  name: 'wps_common_snapshot_list',
  description: `列出本机快照（只读），按时间倒序。可按原文件路径过滤。

用途：恢复之前先看清有哪些版本，别拿一份更旧的盖掉现在的。
日志文件：~/.${RUNTIME_ID}/oplog/ops.jsonl（只记路径与结果，不记文档内容）。`,
  category: ToolCategory.COMMON,
  group: 'common',
  inputSchema: {
    type: 'object',
    properties: {
      originalPath: { type: 'string', description: '只看某个文档的快照（完整路径）' },
      limit: { type: 'number', description: '最多返回几条，默认 20' },
    },
  },
};

export const snapshotListHandler: ToolHandler = async (args) => {
  const limit = typeof args.limit === 'number' && args.limit > 0 ? args.limit : 20;
  const filter = typeof args.originalPath === 'string' ? args.originalPath.toLowerCase() : '';
  const records = readOpLog()
    .filter((r) => r.kind === 'snapshot' && r.snapshotPath && fs.existsSync(r.snapshotPath))
    .filter((r) => !filter || String(r.originalPath).toLowerCase() === filter)
    .sort((a, b) => String(b.ts).localeCompare(String(a.ts)))
    .slice(0, limit);

  if (!records.length) {
    return text('没有快照。写入前先调 wps_common_snapshot_create，或在 batch 里让它自动建。');
  }
  const lines = records.map((r, i) => {
    const size = Math.round((r.size ?? 0) / 1024);
    const dirty = r.dirty ? ' · 磁盘版本' : '';
    return (
      `${i + 1}. ${r.ts}  ${r.app}  ${r.name}（${size} KB）${dirty}\n` +
      `   ${r.snapshotPath}\n   原文件 ${r.originalPath}`
    );
  });
  return text(`共 ${records.length} 份（最新的在最前）：\n\n${lines.join('\n')}`);
};

export const snapshotRestoreDefinition: ToolDefinition = {
  name: 'wps_common_snapshot_restore',
  description: `把快照恢复回去。

**默认不覆盖原文件**：写到 <原名>.restored-<时间戳>.<扩展名>，零风险，你确认没问题再自己替换。
要原地覆盖必须同时给 overwrite=true 与 confirm=true，并且**先关掉那个文档**——
WPS 开着的时候覆盖文件，要么失败，要么被 WPS 下次保存覆盖回去。

恢复动作不可撤销（覆盖前会把当前文件另存一份 pre-restore 快照）。`,
  category: ToolCategory.COMMON,
  group: 'common',
  inputSchema: {
    type: 'object',
    properties: {
      path: { type: 'string', description: '快照文件的完整路径（来自 snapshot_list）' },
      latest: { type: 'boolean', description: '不给 path 时用这个：恢复该文档最新的一份快照' },
      originalPath: { type: 'string', description: '配合 latest 使用：原文档路径；不给则取全局最新一份' },
      overwrite: { type: 'boolean', description: '原地覆盖原文件（危险），默认 false' },
      confirm: { type: 'boolean', description: '覆盖时必须为 true，否则直接拒绝' },
    },
  },
};

export const snapshotRestoreHandler: ToolHandler = async (args) => {
  const snapshotPath = typeof args.path === 'string' ? args.path : '';
  const useLatest = args.latest === true;
  const overwrite = args.overwrite === true;
  const confirmed = args.confirm === true;

  const records = readOpLog().filter((r) => r.kind === 'snapshot' && r.snapshotPath);
  let record: SnapshotRecord | undefined;
  if (snapshotPath) {
    record = records.find((r) => path.resolve(String(r.snapshotPath)) === path.resolve(snapshotPath));
    if (!record) return fail(`没找到这份快照：${snapshotPath}（先 wps_common_snapshot_list 看看）`, 'not_found');
  } else if (useLatest) {
    const filter = typeof args.originalPath === 'string' ? args.originalPath.toLowerCase() : '';
    record = records
      .filter((r) => !filter || String(r.originalPath).toLowerCase() === filter)
      .sort((a, b) => String(b.ts).localeCompare(String(a.ts)))[0];
    if (!record) return fail('没有可恢复的快照。', 'not_found');
  } else {
    return fail('要给 path，或者给 latest=true。', 'invalid_params');
  }

  const src = String(record.snapshotPath);
  const original = String(record.originalPath);
  if (!fs.existsSync(src)) return fail(`快照文件已不在磁盘上：${src}（被清理或删掉了）`, 'not_found');

  if (!overwrite) {
    const ext = path.extname(original);
    const dest = `${original.slice(0, original.length - ext.length)}.restored-${stamp()}${ext}`;
    try {
      fs.copyFileSync(src, dest);
    } catch (e) {
      return fail(`写出恢复副本失败：${e instanceof Error ? e.message : String(e)}`, 'io');
    }
    appendOpLog({ ts: new Date().toISOString(), kind: 'restore', snapshotPath: src, originalPath: original, tool: 'wps_common_snapshot_restore' });
    return text(
      `已恢复出一份副本（原文件未动）：\n${dest}\n\n` +
        '确认内容没问题后自己替换原文件；要直接覆盖原文件，得先关掉 WPS 里的这个文档，' +
        '再调一次并带 overwrite=true、confirm=true。'
    );
  }

  // ---- 原地覆盖：破坏性，先过三道闸 ----
  if (!confirmed) {
    return fail(
      `[未执行] 原地覆盖 ${original} 会丢掉当前内容。确认无误请重发一次并带 confirm=true。`,
      'confirm_required'
    );
  }
  const st = await getActiveState();
  if (st.ok && st.state.fullName && path.resolve(st.state.fullName) === path.resolve(original)) {
    return fail(
      `[勿重试·先关文档] 「${st.state.name}」正开着，覆盖它要么失败、要么被 WPS 下次保存冲掉。` +
        '先在 WPS 里关掉这个文档，再重试。',
      'document_open'
    );
  }
  try {
    // 覆盖前先把当前文件留一份，否则"恢复"本身就成了不可挽回的操作。
    // 备份放进 SNAPSHOT_ROOT（不是原文件旁边）：用户的文档目录不该被我们塞垃圾，
    // 而且放在同一根下才会被保留策略统计到 —— 否则 pre-restore 会无限增长。
    if (fs.existsSync(original)) {
      fs.mkdirSync(SNAPSHOT_ROOT, { recursive: true });
      const pre = path.join(
        SNAPSHOT_ROOT,
        `${stamp()}_${record.app || 'doc'}_pre-restore_${path.basename(original)}`
      );
      fs.copyFileSync(original, pre);
      appendOpLog({
        ts: new Date().toISOString(),
        kind: 'snapshot',
        snapshotPath: pre,
        originalPath: original,
        app: record.app,
        name: record.name,
        label: 'pre-restore',
        tool: 'wps_common_snapshot_restore',
      });
    }
    // WPS 关掉文档后不一定立刻释放文件句柄（实测 EBUSY）—— 重试几次比直接报错有用
    await copyWithRetry(src, original);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (msg.includes('EBUSY') || msg.includes('EPERM')) {
      return fail(
        `[勿重试·先彻底关闭] 文件仍被 WPS 占用，没能覆盖：${original}\n` +
          '文档已从 WPS 关掉但进程还握着句柄。把 WPS 完全退出后重试，或改用默认模式（写出副本）再手动替换。',
        'file_locked'
      );
    }
    return fail(`覆盖失败：${msg}`, 'io');
  }
  appendOpLog({ ts: new Date().toISOString(), kind: 'restore', snapshotPath: src, originalPath: original, tool: 'wps_common_snapshot_restore' });
  return text(
    `已覆盖回：${original}\n快照来源：${src}\n` +
      `覆盖前的版本已另存一份 pre-restore 快照（见 wps_common_snapshot_list）。` +
      '现在去 WPS 里重新打开这个文档即可看到恢复后的内容。'
  );
};

export const snapshotTools: RegisteredTool[] = [
  { definition: snapshotCreateDefinition, handler: snapshotCreateHandler },
  { definition: snapshotListDefinition, handler: snapshotListHandler },
  { definition: snapshotRestoreDefinition, handler: snapshotRestoreHandler },
];
