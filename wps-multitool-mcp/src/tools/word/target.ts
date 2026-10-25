/**
 * Input: 调用参数里的 target（可选）
 * Output: 规范化后的 instance 路由键 + 已合流 documentName 的参数
 * Pos: word 类工具的"改哪一份"显式化。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ------------
 * 现在"改哪份文档"靠活动文档推断（wps-com.ps1 的 Resolve-Doc 三层兜底）。一次只改一份时够用，
 * 连续操作两份就会**静默改错**：改完返回 success: true，改的却是别人那份。这是本项目最贵的一类 bug。
 *
 * 口径（三条，缺一条就白做）
 * ------------------------
 *   1) 不传 target = 旧行为一字不变（否则 67 个 word 工具全要回归）
 *   2) `doc` **合流到现成的 documentName**，不新造第二个名字 —— 两个名字并存必然漂
 *   3) `instance` 只做路由，不改语义：user = 用户正在用的那个；sandbox* = 我们自己的隔离实例
 *
 * 为什么 instance 不放进 schema 的 required
 * ----------------------------------------
 * 不传就是 user（旧行为）。要求必填等于逼所有调用方加字段，模型每步都要多写一个它不关心的值。
 */

/** 池容量上限：再多人肉也管不过来，且每个实例 ~400MB（见 DECISION §3.3） */
export const MAX_SANDBOX_INSTANCES = 3;

/**
 * 每个 word 工具 schema 里复用的 target 片段。
 *
 * 为什么做成常量而不是复制 67 份：复制出来的第 2 份起就会各自漂（改一处忘一处），
 * 而这类漂移没有任何门禁会红 —— 模型只会"偶尔发现某个工具的 target 不认 doc"。
 */
export const WORD_TARGET_PROPERTY = {
  type: 'object' as const,
  description:
    '可选：钉住"改哪一份"。不传 = 用当前活动文档（旧行为）。\n' +
    '- doc：文档名或完整路径。传了就**只**改它，找不到直接报错（不会回退到别的文档）。\n' +
    '- instance：user（默认，你正在用的 WPS）/ sandbox（隔离实例，不碰你打开的文档）/ sandbox-2（指定第几个沙箱）。',
  properties: {
    instance: {
      type: 'string' as const,
      description: 'user（默认）/ sandbox（取第一个空闲沙箱）/ sandbox-<n>（指定第 n 个，1 起）',
    },
    doc: {
      type: 'string' as const,
      description: '文档名或完整路径，如 "报告.docx" 或 "D:\\docs\\报告.docx"',
    },
  },
  required: [] as string[],
};

/** 路由键：'user' 或 'sandbox-N'（N 是 1 起的位置） */
export type WpsInstance = string;

export interface NormalizedTarget {
  /** 规范化后的 instance（默认 'user'） */
  instance: WpsInstance;
  /** 合流进 documentName 的文档名（没传就是 undefined） */
  doc?: string;
  /** 是否显式传了 target（选区类动作据此拒绝） */
  explicit: boolean;
}

/**
 * 解析并校验 target。
 *
 * 校验只做"形状"，不做"池子有没有这个位置"：容量是运行期配置（WPS_SANDBOX_MAX），
 * 只有桥知道。这里越权判容量，会出现"NODE 说没这个沙箱、桥说可以有"的两套口径。
 *
 * @throws Error 带 [勿重试·先改参数]，因为这种错重试一百次也一样
 */
export function normalizeTarget(raw: unknown): NormalizedTarget {
  if (raw === undefined || raw === null) {
    return { instance: 'user', explicit: false };
  }
  if (typeof raw !== 'object' || Array.isArray(raw)) {
    throw new Error('[勿重试·先改参数] target 必须是对象，如 {"doc":"报告.docx"} 或 {"instance":"sandbox"}');
  }
  const obj = raw as Record<string, unknown>;

  const instanceRaw = obj.instance;
  let instance: WpsInstance = 'user';
  if (instanceRaw !== undefined && instanceRaw !== null && instanceRaw !== '') {
    if (typeof instanceRaw !== 'string') {
      throw new Error('[勿重试·先改参数] target.instance 必须是字符串（user / sandbox / sandbox-<n>）');
    }
    const s = instanceRaw.trim();
    if (s === 'user') {
      instance = 'user';
    } else if (s === 'sandbox') {
      // 无后缀 = 池里第一个空闲，由桥决定具体是几号
      instance = 'sandbox';
    } else {
      const m = s.match(/^sandbox-(\d+)$/);
      if (!m) {
        throw new Error(
          `[勿重试·先改参数] 不认识的 target.instance「${s}」。可用：user / sandbox / sandbox-1 … sandbox-${MAX_SANDBOX_INSTANCES}`
        );
      }
      const n = Number(m[1]);
      if (n < 1 || n > MAX_SANDBOX_INSTANCES) {
        throw new Error(
          `[勿重试·先改参数] sandbox-${n} 超出范围。沙箱最多 ${MAX_SANDBOX_INSTANCES} 个（sandbox-1 … sandbox-${MAX_SANDBOX_INSTANCES}）`
        );
      }
      instance = `sandbox-${n}`;
    }
  }

  const docRaw = obj.doc;
  let doc: string | undefined;
  if (docRaw !== undefined && docRaw !== null && docRaw !== '') {
    if (typeof docRaw !== 'string') {
      throw new Error('[勿重试·先改参数] target.doc 必须是字符串（文档名或完整路径）');
    }
    const d = docRaw.trim();
    if (d) doc = d;
  }

  return { instance, doc, explicit: true };
}

/**
 * 把 target.doc 合流进 documentName，并把 target / instance 一并留给下游。
 *
 * 为什么把 instance 塞进 params（而不是只走路由）：PS 侧要靠它判断
 * "这个请求是不是指向沙箱"（选区类动作要拒绝、quitWps 要只退自己）。
 * 路由键和参数各一份是故意的 —— 前者决定"送到哪个进程"，后者决定"进程里怎么处理"。
 */
export function applyWordTarget(
  params: Record<string, unknown> | undefined
): { params: Record<string, unknown>; instance: WpsInstance } {
  const base = params ? { ...params } : {};
  const t = normalizeTarget(base.target);
  if (!t.explicit) return { params: base, instance: 'user' };

  // doc 合流：**每次调用的 target.doc 优先级最高**。
  //
  // 为什么不是"已有 documentName 就不覆盖"：executeMethod 里有【会话级】target
  // （switch_document / create_document 会把当前文档记进 wordTargetName，再自动注入
  // documentName）。如果这里让位，就会出现：用户明明每个调用都写了 target.doc，
  // 实际却一直被会话文档接管 —— 而且完全不报错。
  // 第 2 期第 2 章轮 2 的真机 A1/A3 就是这么"看起来跑通、其实全写进同一份"的。
  if (t.doc) {
    base.documentName = t.doc;
  }
  base.target = { instance: t.instance, doc: t.doc };
  return { params: base, instance: t.instance };
}
