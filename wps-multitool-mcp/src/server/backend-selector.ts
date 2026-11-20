/**
 * Input: 环境变量（WPS_BACKEND / WPS_OFFICE）/ process.platform / WSL 探测 / ProgID→二进制解析
 * Output: 后端识别（com | addon | off）+ host 识别（wps | office）+ 判断依据
 * Pos: Linux 适配的地基——五层决策链（计划 §4.1）；Office 支持一期（2026-11-07）加 host 维度。
 * 原则：探测，不猜。platform 只是线索，能不能干活看通道通不通。
 *   1. WPS_BACKEND 环境变量强制覆盖（com / addon / off；调试与 CI 用，写错直接抛）
 *      （不特判的话 WSL 用户会被错误甩去 addon 通道，平白多一层延迟且原因不可见）
 *   4. Linux/macOS 能力探测（端口上有加载项握手吗）——加载项通道未实现前到不了这层
 *   5. 兜底：明确抛错 + 说清该等什么/设什么，绝不静默降级
 * host 维度（Office 支持一期 · 与 WPS_BACKEND 正交；host wps 优先 = 2026-11-17 所有者裁决）：
 *   WPS_OFFICE=on | off | auto（缺省 auto）。
 *   - on/off：强制指定 host，用户显式选择不探测（显式锁定优先于一切，控制面切换也被拒）。
 *   - auto：双链探测——Office 链（Word/Excel/PowerPoint.Application）与 WPS 链
 *     （Kwps/Ket/Kwpp.Application）各自解析 ProgID 到【实际服务器二进制】真身：
 *     ProgID 存在性、CoCreate 成败、Application.Name/Version 自报统统不可信：
 *     COM 激活会被 WPS 动态类工厂顶名应答（Name 自报 'Microsoft Word'、Version 12.0）。
 *     唯一可靠判据 = LocalServer32 指向的 exe 真身（wps.exe → wps；WINWORD.EXE → office）。
 *     两链都成立 → WPS_HOST_PREFERENCE（缺省 wps）取舍；单链成立 → 该链。
 *   - 探测不可用 / 结论不一致 → 保守按 wps（现任主通道），不猜、不静默假装 office 可用。
 *   - backend=off 时 host 维度不适用，固定 'wps'（不探测）。
 * （Windows 延迟不变）：识别结果模块内缓存一次；Windows 上每次调用只是读缓存，
 * COM 分支一个字节没动（wps-client.ts 只有顶部一个转发点）。
 */

export type Backend = 'com' | 'addon' | 'off';
export type Host = 'wps' | 'office';

export interface BackendDecision {
  backend: Backend;
  /** 实际的 process.platform（原样报告，不做翻译） */
  platform: string;
  /** 是否判为 WSL（platform=linux 但探到 Windows PowerShell） */
  wsl: boolean;
  /** host 维度：这次会话驱动哪套办公套件（Office 支持一期） */
  host: Host;
  /** host 判断依据（人话，与 basis 分开——两个维度两条依据） */
  hostBasis: string;
  /** 判断依据（人话，给 wps_backend_status / 排障日志用） */
  basis: string;
}

/** WSL 标记：WSL 里能看到 Windows 侧的 PowerShell */
export const WSL_MARKER = '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe';

/** Linux/macOS 通道尚未实现时的显式失败（区别于"连不上"的传输层错误） */
export class BackendUnavailableError extends Error {
  constructor(basis: string) {
    super(basis);
    this.name = 'BackendUnavailableError';
  }
}

// ==================== host 探测：ProgID 链 → 服务器二进制 ====================

/** Office 家族二进制名（LocalServer32 指向这些 exe → host=office） */
export const OFFICE_SERVER_EXES = ['winword.exe', 'excel.exe', 'powerpnt.exe'];
/** WPS 家族二进制名（LocalServer32 指向这些 exe → host=wps——顶名劫持的实锤形态） */
export const WPS_SERVER_EXES = ['wps.exe', 'et.exe', 'wpp.exe', 'kwps.exe', 'ket.exe', 'kwpp.exe'];

export interface ParsedServerBinary {
  exe: string;
  host: Host | 'unknown';
}

/**
 * 从 LocalServer32 注册表值解析服务器 exe 并归类（纯函数，可单测）。
 * 兼容两种写法：带引号（"C:\...\WINWORD.EXE" /Automation）与不带引号（C:\...\WINWORD.EXE /Automation）。
 * 解析不出 / 认不出的二进制 → host:'unknown'（调用方按 null 处理，绝不猜）。
 */
export function parseLocalServerBinary(value: string): ParsedServerBinary | null {
  if (!value || !value.trim()) return null;
  const p = value.trim();
  let exe = '';
  if (p.startsWith('"')) {
    const end = p.indexOf('"', 1);
    if (end === -1) return null;
    exe = p.slice(1, end);
  } else {
    // 不带引号：截到第一个 " /" 参数分隔（路径含空格但不含 " /" 是 Office/WPS 注册的实际形态）
    const m = p.match(/^(.*?)\s+\//);
    exe = m ? m[1] : p;
  }
  exe = exe.trim();
  if (!exe) return null;
  const base = exe.split('\\').pop()?.toLowerCase() ?? '';
  if (OFFICE_SERVER_EXES.includes(base)) return { exe, host: 'office' };
  if (WPS_SERVER_EXES.includes(base)) return { exe, host: 'wps' };
  return { exe, host: 'unknown' };
}

/** 注册表读取器：给定键路径，返回默认值（读不到返回 null）。测试注入用同一形状。 */
export type RegRunner = (key: string) => string | null;

/**
 * 解析一个 ProgID 到实际服务器二进制的 host 归类（可注入 reg 读取器，纯逻辑）。
 * 链路：ProgID → CLSID（HKCU 优先于 HKLM，模拟合并视图优先级）
 *       → LocalServer32（64 位视图 HKCU→HKLM，再扫 WOW6432Node——WPS 32 位劫持的历史注册位）。
 * 任一环断了返回 null（不是失败也不是 office，交给调用方保守处理）。
 */
export function resolveProgidHost(
  progid: string,
  runReg: RegRunner
): Host | null {
  const clsid =
    runReg(`HKCU\\SOFTWARE\\Classes\\${progid}\\CLSID`) ??
    runReg(`HKLM\\SOFTWARE\\Classes\\${progid}\\CLSID`);
  if (!clsid) return null;
  const server =
    runReg(`HKCU\\SOFTWARE\\Classes\\CLSID\\${clsid}\\LocalServer32`) ??
    runReg(`HKLM\\SOFTWARE\\Classes\\CLSID\\${clsid}\\LocalServer32`) ??
    runReg(`HKLM\\SOFTWARE\\Classes\\WOW6432Node\\CLSID\\${clsid}\\LocalServer32`) ??
    runReg(`HKCU\\SOFTWARE\\Classes\\WOW6432Node\\CLSID\\${clsid}\\LocalServer32`);
  if (!server) return null;
  const parsed = parseLocalServerBinary(server);
  if (!parsed || parsed.host === 'unknown') return null;
  return parsed.host;
}

// 真机 reg 查询（probeHostLive 专用）：reg.exe query /ve 的输出形如
// "    LocalServer32    REG_SZ    C:\Program Files\...\WINWORD.EXE /Automation"
function runRegQuery(key: string): string | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { execFileSync } = require('child_process') as typeof import('child_process');
    const out = execFileSync('reg', ['query', key, '/ve'], {
      encoding: 'utf8',
      timeout: 5000,
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'ignore'], // reg.exe 对缺失键会往 stderr 喊话，探针不该带噪音
    });
    const m = out.match(/REG_SZ\s+(.*)/);
    if (!m) return null;
    const v = m[1].trim();
    return v || null;
  } catch {
    return null;
  }
}

let hostProbeCache: Host | null | undefined;

/** Office 家族 ProgID（Office 链：三件全部解析到 office 二进制才算成立） */
export const OFFICE_PROGIDS = ['Word.Application', 'Excel.Application', 'PowerPoint.Application'];
/** WPS 家族 ProgID（WPS 链：与 Office 三链对称，Kwps/Ket/Kwpp） */
export const WPS_PROGIDS = ['Kwps.Application', 'Ket.Application', 'Kwpp.Application'];

export interface ChainResult {
  /** 链结论（null = 链不成立：断链/残缺/不一致/被顶名） */
  host: Host | null;
  /** Office ProgID 被解析到 WPS 二进制 = 顶名劫持实锤（WPS 独装机的典型形态） */
  hijacked: boolean;
}

/**
 * 单链解析：一条链的三件 ProgID 全部解析成功且结论一致才成立；
 * 结论必须是期望 host（Office 链全部解析到 wps 二进制 = WPS 顶名劫持，链不成立而非结论）。
 * 纯逻辑，注入 reg 读取器可单测。
 */
export function resolveChain(progids: string[], expected: Host, runReg: RegRunner): ChainResult {
  const results: Host[] = [];
  let hijacked = false;
  for (const pid of progids) {
    const h = resolveProgidHost(pid, runReg);
    if (h) {
      results.push(h);
      if (h !== expected) hijacked = true;
    }
  }
  if (results.length !== progids.length) return { host: null, hijacked };
  if (!results.every((h) => h === results[0])) return { host: null, hijacked };
  if (results[0] !== expected) return { host: null, hijacked: true };
  return { host: expected, hijacked: false };
}

/**
 * 双链取舍（纯函数）：两链都成立 → preference 取舍；单链 → 该链；双链都不成立 → null
 * （detectHost 保守 wps）。host wps 优先 = 缺省 preference 'wps'（所有者 11-17 裁决）。
 */
export function decideDualChainHost(
  officeChain: Host | null,
  wpsChain: Host | null,
  preference: 'wps' | 'office'
): Host | null {
  if (officeChain && wpsChain) return preference;
  if (wpsChain) return 'wps';
  if (officeChain) return 'office';
  return null;
}

/** 优先策略读值：WPS_HOST_PREFERENCE=wps|office（缺省 wps；非法值按 wps，不抛） */
export function hostPreference(env: Record<string, string | undefined> = process.env): 'wps' | 'office' {
  return (env.WPS_HOST_PREFERENCE ?? '').trim().toLowerCase() === 'office' ? 'office' : 'wps';
}

/**
 * 真机 host 探测 v2（双链，结果缓存）：
 *   Office 链 = Word/Excel/PowerPoint.Application → 期望 office
 *   WPS 链   = Kwps/Ket/Kwpp.Application → 期望 wps
 * 两链都成立 → WPS_HOST_PREFERENCE（缺省 wps）取舍；单链成立 → 该链；都不成立 → null
 * （detectHost 保守 wps）。探测失败不是错误——是"不知道"，绝不猜。
 */
export function probeHostLive(): Host | null {
  if (hostProbeCache !== undefined) return hostProbeCache;
  if (process.platform !== 'win32') {
    hostProbeCache = null;
    return null;
  }
  const office = resolveChain(OFFICE_PROGIDS, 'office', runRegQuery);
  const wps = resolveChain(WPS_PROGIDS, 'wps', runRegQuery);
  hostProbeCache = decideDualChainHost(office.host, wps.host, hostPreference());
  return hostProbeCache;
}

// ==================== 会话级 host 切换（控制面 override 层 · 变更二）====================
// POST /host 把切换写进桥进程的这里；优先级：env 显式锁定 > override > auto 探测。
// 探测缓存不被 override 污染——auto 的结论仍是基准，override 只盖最终结果。

let hostOverrideState: Host | null = null;

/** 控制面写入口：null = 清除切换（回到 env/探测语义） */
export function setHostOverride(h: Host | null): void {
  hostOverrideState = h;
}

export function getHostOverride(): Host | null {
  return hostOverrideState;
}

// ==================== host 维度识别（与 backend 正交）====================

export interface DetectHostInput {
  env?: Record<string, string | undefined>;
  probe?: () => Host | null;
}

export function detectHost(input?: DetectHostInput): { host: Host; hostBasis: string } {
  const env = input?.env ?? process.env;
  const officeEnv = (env.WPS_OFFICE ?? '').trim().toLowerCase();
  if (officeEnv === 'on') return { host: 'office', hostBasis: 'WPS_OFFICE=on 强制指定（Office host，不探测）' };
  if (officeEnv === 'off') return { host: 'wps', hostBasis: 'WPS_OFFICE=off 强制指定（WPS host，不探测）' };
  const override = getHostOverride();
  if (override) {
    return { host: override, hostBasis: '控制面 hostOverride：会话级切换生效（env 未显式锁定；auto 探测结论仍可作为基准对照）' };
  }
  if (officeEnv === '' || officeEnv === 'auto') {
    const probed = input?.probe ? input.probe() : null;
    if (probed) {
      const pref = hostPreference(env);
      const dualNote =
        pref === 'wps'
          ? '双链探测（Office+Kwps/Ket/Kwpp），双装缺省 wps 优先（WPS_HOST_PREFERENCE 可改）'
          : '双链探测（Office+Kwps/Ket/Kwpp），WPS_HOST_PREFERENCE=office';
      return { host: probed, hostBasis: `WPS_OFFICE=auto（缺省）：${dualNote} → ${probed}` };
    }
    return {
      host: 'wps',
      hostBasis: 'WPS_OFFICE=auto（缺省）：探测不可用或结论不一致（可能 WPS 顶名劫持/残缺安装），保守按 wps（现任主通道）',
    };
  }
  throw new BackendUnavailableError(
    `WPS_OFFICE=${env.WPS_OFFICE} 不是合法值（可用：on / off / auto）`
  );
}

export function detectBackend(input?: {
  env?: Record<string, string | undefined>;
  platform?: string;
  exists?: (p: string) => boolean;
  probe?: () => Host | null;
}): BackendDecision {
  const env = input?.env ?? process.env;
  const platform = input?.platform ?? process.platform;
  const exists =
    input?.exists ??
    ((p: string) => {
      try {
        // 延迟 require，避免测试注入时加载 fs
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        return require('fs').existsSync(p);
      } catch {
        return false;
      }
    });

  // 第 1 层：环境变量强制覆盖（backend）。off 时 host 维度不适用（不探测，固定 wps）。
  const forced = (env.WPS_BACKEND || '').trim().toLowerCase();
  if (forced) {
    if (forced === 'com' || forced === 'addon') {
      const h = detectHost({ env, probe: input?.probe });
      const basis = forced === 'com'
        ? 'WPS_BACKEND=com 强制指定'
        : 'WPS_BACKEND=addon 强制指定';
      return {
        backend: forced,
        platform,
        wsl: false,
        host: h.host,
        hostBasis: h.hostBasis,
        basis,
      };
    }
    if (forced === 'off') {
      return {
        backend: 'off',
        platform,
        wsl: false,
        host: 'wps',
        hostBasis: 'backend=off：host 维度不适用，固定 wps（不探测）',
        basis: 'WPS_BACKEND=off 强制指定（全部调用显式禁用）',
      };
    }
    throw new BackendUnavailableError(
      `WPS_BACKEND=${env.WPS_BACKEND} 不是合法值（可用：com / addon / off）`
    );
  }

  // host 维度（与 backend 正交，auto 缺省探测）——win32/WSL 两条 COM 路共用
  const h = detectHost({ env, probe: input?.probe });

  // 第 2 层：原生 Windows → COM
  if (platform === 'win32') {
    return { backend: 'com', platform, wsl: false, host: h.host, hostBasis: h.hostBasis, basis: '原生 Windows（win32），走 COM 命名管道' };
  }

  // 第 3 层：WSL 特判
  if (platform === 'linux' && exists(WSL_MARKER)) {
    return {
      backend: 'com',
      platform,
      wsl: true,
      host: h.host,
      hostBasis: h.hostBasis,
      basis: 'WSL 环境（探到 Windows PowerShell），能力识别走 COM；管道在 WSL 的可达性由通道层实测',
    };
  }

  // 第 4/5 层：Linux/macOS——能力探测（加载项握手），加载项通道没实现前到不了那层。
  // 现状 = 明确报错 + 说清选项，绝不静默降级去撞一个必然失败的通道。
  const os = platform === 'darwin' ? 'macOS' : 'Linux';
  throw new BackendUnavailableError(
    `${os} 通道尚未实现（addon-channel 属 B2；加载项按宿主分三型：wps/et/wpp）。` +
      '可选：① 在 Windows / WSL 上使用（COM）；② 设 WPS_BACKEND=off 明确禁用；③ 等 B2 交付。'
  );
}

/** 识别缓存：模块内只算一次。失败不缓存——每次调用都拿到最新的一次明确报错 */
let cached: BackendDecision | null = null;

export function backendDecision(): BackendDecision {
  if (!cached) cached = detectBackend({ probe: probeHostLive });
  return cached;
}

/**
 * 便捷常量（计划 §4.2 的形状）：addon 通道实现后，Linux + 加载项在线时为 true。
 * Windows 上恒 false。探测失败（Linux 无通道）按 false 处理——真正的报错由
 * backendDecision() 在调用时抛，别在模块加载期把整个 server 带崩。
 */
export const USE_ADDON = (() => {
  try {
    return backendDecision().backend === 'addon';
  } catch {
    return false;
  }
})();

/** 后端状态摘要（wps_backend_status 工具的数据源） */
export function backendStatus(): BackendDecision & { transport: string } {
  const d = backendDecision();
  return {
    ...d,
    transport: d.backend === 'com' ? '命名管道 JSON 行（Windows）' : '待定（WS / 轮询，B4 裁决）',
  };
}
