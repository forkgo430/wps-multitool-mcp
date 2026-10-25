/**
 * Input: 环境变量（WPS_BACKEND）/ process.platform / WSL 探测结果
 * Output: 后端判定（com | addon | off）+ 判定依据
 * Pos: Linux 适配的地基——五层决策链（计划 §4.1）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 原则：探测，不猜。platform 只是线索，能不能干活看通道通不通。
 *
 * 五层（自上而下，命中即返回）：
 *   1. WPS_BACKEND 环境变量强制覆盖（com / addon / off；调试与 CI 用，写错直接抛）
 *   2. process.platform === 'win32' → COM（现状主通道）
 *   3. WSL 特判：platform 报 linux 但存在 /mnt/c/.../powershell.exe → 仍走 COM
 *      （不特判的话 WSL 用户会被错误甩去 addon 通道，平白多一层延迟且原因不可见）
 *   4. Linux/macOS 能力探测（端口上有加载项握手吗）——归 B2，通道未实现前到不了这层
 *   5. 兜底：明确抛错 + 说清该等什么/设什么，绝不静默降级
 *
 * C4（Windows 延迟不变）：判定结果模块内缓存一次；Windows 上每次调用只是读缓存，
 * COM 分支一个字节没动（wps-client.ts 只有顶部一个转发点，计划 §4.2 的"唯一改动"）。
 */

export type Backend = 'com' | 'addon' | 'off';

export interface BackendDecision {
  backend: Backend;
  /** 实际的 process.platform（原样报告，不做翻译） */
  platform: string;
  /** 是否判为 WSL（platform=linux 但探到 Windows PowerShell） */
  wsl: boolean;
  /** 判定依据（人话，给 wps_backend_status / 排障日志用） */
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

export function detectBackend(input?: {
  env?: Record<string, string | undefined>;
  platform?: string;
  exists?: (p: string) => boolean;
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

  // 第 1 层：环境变量强制覆盖
  const forced = (env.WPS_BACKEND || '').trim().toLowerCase();
  if (forced) {
    if (forced === 'com') return { backend: 'com', platform, wsl: false, basis: 'WPS_BACKEND=com 强制指定' };
    if (forced === 'addon') return { backend: 'addon', platform, wsl: false, basis: 'WPS_BACKEND=addon 强制指定' };
    if (forced === 'off') return { backend: 'off', platform, wsl: false, basis: 'WPS_BACKEND=off 强制指定（全部调用显式禁用）' };
    throw new BackendUnavailableError(
      `WPS_BACKEND=${env.WPS_BACKEND} 不是合法值（可用：com / addon / off）`
    );
  }

  // 第 2 层：原生 Windows → COM
  if (platform === 'win32') {
    return { backend: 'com', platform, wsl: false, basis: '原生 Windows（win32），走 COM 命名管道' };
  }

  // 第 3 层：WSL 特判
  if (platform === 'linux' && exists(WSL_MARKER)) {
    return {
      backend: 'com',
      platform,
      wsl: true,
      basis: 'WSL 环境（探到 Windows PowerShell），能力判定走 COM；管道在 WSL 的可达性由通道层实测（B2/B4）',
    };
  }

  // 第 4/5 层：Linux/macOS——能力探测（加载项握手）归 B2，通道没实现前到不了那层。
  // 现状 = 明确报错 + 说清选项，绝不静默降级去撞一个必然失败的通道。
  const os = platform === 'darwin' ? 'macOS' : 'Linux';
  throw new BackendUnavailableError(
    `${os} 通道尚未实现（addon-channel 属 B2；加载项按宿主分三型：wps/et/wpp）。` +
      '可选：① 在 Windows / WSL 上使用（COM）；② 设 WPS_BACKEND=off 明确禁用；③ 等 B2 交付。'
  );
}

/** 判定缓存：模块内只算一次（C4）。失败不缓存——每次调用都拿到最新的一次明确报错 */
let cached: BackendDecision | null = null;

export function backendDecision(): BackendDecision {
  if (!cached) cached = detectBackend();
  return cached;
}

/**
 * 便捷常量（计划 §4.2 的形状）：addon 通道实现后，Linux + 加载项在线时为 true。
 * Windows 上恒 false。判定失败（Linux 无通道）按 false 处理——真正的报错由
 * backendDecision() 在调用时抛，别在模块加载期把整个 server 带崩。
 */
export const USE_ADDON = (() => {
  try {
    return backendDecision().backend === 'addon';
  } catch {
    return false;
  }
})();

/** 后端状态摘要（wps_backend_status 工具的数据源，B2 注册工具时直接用） */
export function backendStatus(): BackendDecision & { transport: string } {
  const d = backendDecision();
  return { ...d, transport: d.backend === 'com' ? '命名管道 JSON 行（Windows）' : '待定（WS / 轮询，B4 裁决）' };
}
