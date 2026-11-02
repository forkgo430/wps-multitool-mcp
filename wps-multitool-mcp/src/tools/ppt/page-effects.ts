/**
 * Input: 页面效果与备注操作参数（动画、背景、切换、演讲者备注）
 * Output: 页面效果操作结果
 * Pos: PPT 页面效果工具实现（风格统一+拆分轮新建，见 ppt-audit §三之二——
 *      这 4 个自 slide-ops.ts 迁入；背景族其余成员已在第 5/8 轮并入这里的两个合并工具）。
 *      一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 *
 * 包含（4 个）：
 * - wps_ppt_set_animation: 设置元素动画（第 8 轮合并了动画族 5 个工具）
 * - wps_ppt_set_background: 设置背景（第 5 轮合并了背景族 6 个工具）
 * - wps_ppt_set_transition: 设置幻灯片切换效果（第 8 轮合并了切换族 3 个工具）
 * - wps_ppt_add_speaker_notes: 添加演讲者备注
 */

import { v4 as uuidv4 } from 'uuid';
import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
  ToolWarning,
} from '../../types/tools';
import { wpsClient } from '../../client/wps-client';
import { WpsAppType } from '../../types/wps';
import { call, ok, fail, errText } from './helpers';

// ============================================================
// 1. wps_ppt_set_animation - 设置元素动画
// ============================================================

export const setAnimationDefinition: ToolDefinition = {
  name: 'wps_ppt_set_animation',
  description: `给形状加动画，或调整已有动画的播放顺序（第 8 轮合并了动画族 5 个工具）。

**两个分支用 \`op\` 选**：
- \`op:'add'\`（默认）：给形状加动画，要 \`shapeIndex\` + \`category\` + \`effect\`
- \`op:'reorder'\`：把已有动画挪位置，要 \`animationIndex\` + \`order\`

\`category\` 决定动作落在哪个动画窗格：
- entrance 入场（appear / fadeIn / flyIn / wipeIn / zoomIn / bounceIn / spinIn）
- emphasis 强调（pulse / grow / teeter / colorPulse）
- exit 退场（fadeOut / flyOut / wipeOut / zoomOut / disappear）

什么时候用：
- "给第1页第2个形状加个淡入" → { op:'add', slideIndex:1, shapeIndex:2, category:'entrance', effect:'fadeIn' }
- "让这个元素闪一下强调" → { op:'add', slideIndex:1, shapeIndex:2, category:'emphasis', effect:'pulse' }
- "把第3个动画移到最前面" → { op:'reorder', slideIndex:1, animationIndex:3, order:1 }

限制：\`.add\` 是【累加】的 —— 同一形状重复加会叠出多条动画，不会替换。
想看现在有哪些动画用 wps_ppt_get_animations，删掉某一条用 wps_ppt_remove_animation。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      op: {
        type: 'string',
        enum: ['add', 'reorder'],
        description: '操作：add 加动画（默认）/ reorder 调顺序',
      },
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）',
      },
      shapeIndex: {
        type: 'number',
        description: '形状索引（从1开始）；op=add 时必填',
      },
      category: {
        type: 'string',
        enum: ['entrance', 'emphasis', 'exit'],
        description: '动画类别；op=add 时必填',
      },
      effect: {
        type: 'string',
        description: '具体效果名。entrance: fadeIn/flyIn/wipeIn/zoomIn/bounceIn/spinIn；emphasis: pulse/spin/grow/teeter/colorPulse；exit: fadeOut/flyOut/wipeOut/zoomOut/disappear',
      },
      trigger: {
        type: 'string',
        enum: ['onClick', 'withPrevious', 'afterPrevious'],
        description: '触发方式，默认 onClick（单击时）',
      },
      animationIndex: {
        type: 'number',
        description: '已有动画的序号（从1开始）；op=reorder 时必填',
      },
      order: {
        type: 'number',
        description: '新的播放位置（从1开始）；op=reorder 时必填',
      },
    },
    required: ['slideIndex'],
  },
};

/** 旧 animationType → { category, effect }。旧名 wps_ppt_set_animation 只收一个字符串，
 *  收尾先按 Out/Exit 结尾判退场，其余按入场；未知名默认入场（宁可多加一个也别不加）。 */
function splitLegacyAnimationType(animationType: string): {
  category: 'entrance' | 'exit';
  effect: string;
} {
  return {
    category: /(out|exit)$/i.test(animationType) ? 'exit' : 'entrance',
    effect: animationType,
  };
}

const ANIM_NAME_MAP: Record<string, string> = {
  fadeIn: '淡入', flyIn: '飞入', wipeIn: '擦除', zoomIn: '缩放进入',
  bounceIn: '弹跳进入', spinIn: '旋转进入', fadeOut: '淡出', flyOut: '飞出',
  wipeOut: '擦除退场', zoomOut: '缩小退场', disappear: '消失',
  pulse: '脉冲', grow: '放大缩小', teeter: '跷跷板', colorPulse: '颜色脉冲',
};

const TRIGGER_NAME_MAP: Record<string, string> = {
  onClick: '单击时',
  withPrevious: '与上一动画同时',
  afterPrevious: '上一动画之后',
};

const CATEGORY_NAME_MAP: Record<string, string> = {
  entrance: '入场', emphasis: '强调', exit: '退场',
};

/**
 * category + effect → PS 真正认的效果名。
 * 为什么要在 TS 侧合成：WPS COM 的 AddEffect 只吃一个效果 ID，不认"入场/强调/退场"这个维度。
 * 不做这层合成的话，category 就是一个装饰参数（PS 收不到、进不了动画窗格对应格），
 * 那才是真的骗人 —— 用户以为加了退场动画，实际加的还是入场。
 */
function resolveEffectId(category: string, effect: string): string {
  const normalized = String(effect || '').trim();
  const lower = normalized.toLowerCase();
  if (category === 'exit') {
    const outMap: Record<string, string> = {
      fade: 'fadeOut', fly: 'flyOut', wipe: 'wipeOut', zoom: 'zoomOut',
      vanish: 'vanish', disappear: 'disappear', dissolve: 'dissolveOut',
    };
    // 已经带 Out 尾名的（fadeOut）原样送下去，不重复拼
    if (lower.endsWith('out') || lower === 'disappear' || lower === 'vanish') return normalized;
    return outMap[lower] || normalized;
  }
  const inMap: Record<string, string> = {
    fade: 'fadeIn', fly: 'flyIn', wipe: 'wipeIn', zoom: 'zoomIn',
    dissolve: 'dissolveIn', appear: 'appear',
  };
  if (lower.endsWith('in') || lower === 'appear') return normalized;
  return inMap[lower] || normalized;
}

export const setAnimationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const op = (args.op as string) || 'add';

  if (op === 'reorder') {
    const { slideIndex, animationIndex, order } = args as {
      slideIndex: number;
      animationIndex: number;
      order: number;
    };
    if (!animationIndex || !order) {
      return fail(
        `op=reorder 时必须同时给 animationIndex 和 order`,
        'Missing required parameter: animationIndex / order'
      );
    }
    const r = await call<{ success: boolean }>(
      'setAnimationOrder',
      { slideIndex, animationIndex, newOrder: order }
    );
    if (!r.ok) return fail(errText('调整动画顺序失败', r), r.error);
    return ok(
      `动画顺序调整成功！\n幻灯片: 第 ${slideIndex} 页\n动画 ${animationIndex} → 新位置 ${order}`,
      r.warnings
    );
  }

  // op=add：effect 缺失时兜住旧的 animationType（第 2 轮起这个参数就叫 animationType）
  const raw = args as {
    slideIndex: number;
    shapeIndex?: number;
    category?: string;
    effect?: string;
    trigger?: string;
    animationType?: string;
  };
  const legacy = !raw.effect && raw.animationType
    ? splitLegacyAnimationType(String(raw.animationType))
    : null;
  const category = raw.category || (legacy ? legacy.category : undefined);
  const effect = raw.effect || (legacy ? legacy.effect : undefined);

  if (!raw.shapeIndex || !category || !effect) {
    return fail(
      `op=add 时必须同时给 shapeIndex、category（entrance/emphasis/exit）和 effect`,
      'Missing required parameter: shapeIndex / category / effect'
    );
  }

  // category 在这里落地成真正的效果名（exit+fade → fadeOut），PS 侧只认效果 ID
  const effectId = resolveEffectId(category as string, effect as string);

  const r = await call<{ animationIndex?: number }>(
    'addAnimation',
    {
      slideIndex: raw.slideIndex,
      shapeIndex: raw.shapeIndex,
      effect: effectId,
      trigger: raw.trigger || 'onClick',
    }
  );
  if (!r.ok) return fail(errText('设置动画失败', r), r.error);
  const lines = [
    `动画设置成功！`,
    `幻灯片: 第 ${raw.slideIndex} 页`,
    `形状: 第 ${raw.shapeIndex} 个`,
    `类别: ${CATEGORY_NAME_MAP[category as string] || category}`,
    `动画: ${ANIM_NAME_MAP[effect as string] || effect}（${effectId}）`,
    `触发: ${TRIGGER_NAME_MAP[raw.trigger || 'onClick'] || raw.trigger || 'onClick'}`,
  ];
  if (r.data && r.data.animationIndex) {
    lines.push(`动画序号: ${r.data.animationIndex}`);
  }
  return ok(lines.join('\n'), r.warnings);
};

// ============================================================
// 2. wps_ppt_set_background - 设置背景
// ============================================================

export const setBackgroundDefinition: ToolDefinition = {
  name: 'wps_ppt_set_background',
  description: `设置背景：某一页、整份文稿、或幻灯片母版。

**三个维度一次说清**（合并了原来 6 个背景工具，别再找 set_background_color / _image / _gradient / set_slide_background / set_master_background）：
- \`target\`: slide（默认，单页）/ all（整份文稿）/ master（母版）
- \`type\`: solid（纯色）/ image（图片）/ gradient（渐变）/ pattern
- 具体的填充物：\`color\` / \`imagePath\` / \`gradient\`

什么时候用：
- "把这页背景改成深蓝" → { target:'slide', slideIndex:2, type:'solid', color:'#003366' }
- "整份文稿统一白底" → { target:'all', type:'solid', color:'#FFFFFF' }
- "用这张图做背景" → { target:'slide', slideIndex:1, type:'image', imagePath:'C:\\\\pics\\\\bg.png' }
- "母版加渐变" → { target:'master', type:'gradient', gradient:{ type:'linear', colors:['#FF0000','#00FF00'] } }

不填 \`type\` 也能用：按你给的 color / imagePath / gradient 自动推断（老名字走别名过来时就是这条路）。

限制（真限制，不是没实现）：
- \`target:'master'\` 只支持纯色与渐变 —— 母版背景换图片 WPS 不认，会明确报错而不是假装成功。
- \`type:'pattern'\`（图案填充）WPS COM 没有对应接口，明确报错。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      target: {
        type: 'string',
        enum: ['slide', 'all', 'master'],
        description: '作用范围：slide 单页（默认）/ all 整份文稿 / master 母版',
      },
      slideIndex: {
        type: 'number',
        description: '页码（从1开始）；target=all/master 时忽略',
      },
      type: {
        type: 'string',
        enum: ['solid', 'gradient', 'image', 'pattern'],
        description: '填充类型。不填则按 color→solid / imagePath→image / gradient→gradient 推断',
      },
      color: { type: 'string', description: 'type=solid 时的颜色，十六进制如 #FF0000' },
      imagePath: { type: 'string', description: 'type=image 时的图片路径' },
      gradient: {
        type: 'object',
        description:
          'type=gradient 时的渐变配置：{ type:"linear"|"radial", colors:["#FF0000","#00FF00"] }。' +
          'WPS 只吃前两个颜色，多余的颜色会被忽略（会在提示里说明）',
      },
      pattern: { type: 'string', description: 'type=pattern 时的图案名（WPS 不支持，传了会明确报错）' },
      // 老名字（别名）可能会带过来的历史参数，接受但不推荐
      background: {
        type: 'object',
        description: '（历史形状）{type,color,colors,imagePath,pattern}，等价于把里面几项平铺到顶层',
      },
    },
    // 刻意不 required['type']：草案里是必填，但那样 5 个旧名就没法走别名（别名不能注入参数），
    // 而合并的目标恰恰是"旧名无损映射"。改为按参数推断，并在描述里写明。
    required: [],
  },
};

/** 把老形状 background{} 摊平到顶层（只补缺失的键） */
function flattenLegacyBackground(args: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = { ...args };
  const bg = args.background as Record<string, unknown> | undefined;
  if (bg && typeof bg === 'object') {
    for (const [k, v] of Object.entries(bg)) {
      if (out[k] === undefined) out[k] = v;
    }
    // 老形状里 "colors" 是渐变色数组
    if (out.gradient === undefined && Array.isArray(bg.colors)) {
      out.gradient = { type: 'linear', colors: bg.colors };
    }
  }
  return out;
}

export const setBackgroundHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const a = flattenLegacyBackground(args);
  const target = (a.target as string) || 'slide';
  const slideIndex = a.slideIndex as number | undefined;
  const color = a.color as string | undefined;
  const imagePath = a.imagePath as string | undefined;
  const pattern = a.pattern as string | undefined;
  const gradient = a.gradient as { type?: string; colors?: string[]; angle?: number } | undefined;

  // type 推断：草案要求必填，但旧名走别名时给不出 type（别名不能注入参数）。
  // 按"给了什么填什么"推断 —— 唯一有歧义的情况是同时给了多个，那时按 image > gradient > color 取先。
  const type =
    (a.type as string) ||
    (imagePath ? 'image' : gradient ? 'gradient' : pattern ? 'pattern' : color ? 'solid' : '');

  const failLocal = (msg: string, err?: string): ToolCallResult => fail(msg, err ?? msg);

  if (!type) {
    return failLocal(
      '要说清怎么填：给 color（纯色）/ imagePath（图片）/ gradient（渐变）之一，或显式传 type。' +
        '示例：{ "target":"slide", "slideIndex":1, "type":"solid", "color":"#003366" }'
    );
  }
  if (type === 'pattern') {
    return failLocal('WPS COM 没有图案填充的接口，type:"pattern" 无法执行（这不是参数问题，换 solid/gradient/image）。');
  }
  if (type === 'solid' && !color) return failLocal('type=solid 需要 color，例如 "#003366"。');
  if (type === 'image' && !imagePath) return failLocal('type=image 需要 imagePath（图片在本机的绝对路径）。');
  if (type === 'gradient') {
    const colors = (gradient && gradient.colors) || [];
    if (!Array.isArray(colors) || colors.length < 2) {
      return failLocal('type=gradient 需要 gradient.colors 至少两个颜色，例如 ["#FF0000","#00FF00"]。');
    }
  }
  if (target === 'master' && type === 'image') {
    return failLocal('母版背景不支持图片（WPS 的母版 Fill 只吃纯色/渐变）。要整份文稿用图片背景，用 target:"all"。');
  }
  if (target === 'slide' && !slideIndex) {
    return failLocal('target=slide 需要 slideIndex（从 1 开始）。要改整份文稿用 target:"all"。');
  }

  // 一个内部小包装。注意：动作名走变量 → param-shape-audit（扫 executeMethod 的字面量）看不到这 4 个
  // 调用点，属于那个脚本自己写明的"动态拼参数抓不到 = 漏报"。所以这里的参数是【手工对着 PS 的读法核过的】：
  //   setMasterBackground 读 color / color1 / color2
  //   setBackgroundColor  读 color / applyToAll
  //   setSlideBackground  读 slideIndex / color / imagePath
  //   getSlideCount       无参数
  // 改这里的参数名时，请同步改上面这份清单（它就是这几个动作的"参数契约"）。
  const exec = <T>(action: string, params: Record<string, unknown>) =>
    wpsClient.executeMethod<T>(action, params, WpsAppType.PRESENTATION);

  const warnings: ToolWarning[] = [];
  const where = target === 'all' ? '整份文稿' : target === 'master' ? '母版' : `第 ${slideIndex} 页`;

  // ---- 母版 ----
  if (target === 'master') {
    const params: Record<string, unknown> =
      type === 'gradient'
        ? { color1: gradient!.colors![0], color2: gradient!.colors![1] }
        : { color };
    const r = await exec('setMasterBackground', params);
    if (!r.success) return fail(`设置母版背景失败: ${r.error}`, r.error);
    return ok(`母版背景已设置：${type === 'gradient' ? `渐变 ${gradient!.colors!.slice(0, 2).join(' → ')}` : color}`, r.warnings);
  }

  // ---- 整份文稿 ----
  if (target === 'all') {
    if (type === 'solid') {
      const r = await exec('setBackgroundColor', { color, applyToAll: true });
      if (!r.success) return fail(`设置整份文稿背景失败: ${r.error}`, r.error);
      return ok(`整份文稿背景已设为 ${color}（applyToAll）`, r.warnings);
    }
    // 渐变/图片没有"一次全设"的动作 → 逐页设。先问页数，别拿失败当终止条件。
    const cnt = await exec<{ count?: number }>('getSlideCount', {});
    const total = Number((cnt.data && cnt.data.count) || 0);
    if (!total) return failLocal('拿不到页数，无法逐页设置背景。');
    let done = 0;
    const failedPages: number[] = [];
    for (let i = 1; i <= total; i++) {
      const r =
        type === 'image'
          ? await exec('setSlideBackground', { slideIndex: i, imagePath })
          : await exec('setSlideBackground', {
              slideIndex: i,
              color1: gradient!.colors![0],
              color2: gradient!.colors![1],
            });
      if (r.success) done++;
      else failedPages.push(i);
    }
    if (done === 0) return failLocal(`整份文稿逐页设置背景全部失败（共 ${total} 页）。`);
    if (failedPages.length) {
      warnings.push({
        step: 'set_background.all',
        message: `第 ${failedPages.join('、')} 页没设置成功（共 ${total} 页，成功 ${done} 页）`,
      });
    }
    return {
      id: uuidv4(),
      success: true,
      content: [{ type: 'text', text: `整份文稿背景已设置：${done}/${total} 页` }],
      ...(warnings.length ? { warnings } : {}),
    };
  }

  // ---- 单页 ----
  const params: Record<string, unknown> = { slideIndex };
  if (type === 'solid') params.color = color;
  else if (type === 'image') params.imagePath = imagePath;
  else {
    params.color1 = gradient!.colors![0];
    params.color2 = gradient!.colors![1];
  }
  const r = await exec('setSlideBackground', params);
  if (!r.success) return fail(`设置背景失败: ${r.error}`, r.error);
  const detail =
    type === 'solid' ? `颜色 ${color}` : type === 'image' ? `图片 ${imagePath}` : `渐变 ${gradient!.colors!.slice(0, 2).join(' → ')}`;
  if (type === 'gradient' && gradient!.colors!.length > 2) {
    warnings.push({ step: 'gradient', message: 'WPS 的渐变只支持两个色标，多余的颜色已忽略' });
  }
  return {
    id: uuidv4(),
    success: true,
    content: [{ type: 'text', text: `背景已设置：${where}，${detail}` }],
    ...(warnings.length ? { warnings } : {}),
  };
};

// ============================================================
// 3. wps_ppt_set_transition - 设置幻灯片切换效果
// ============================================================

export const setTransitionDefinition: ToolDefinition = {
  name: 'wps_ppt_set_transition',
  description: `设置幻灯片切换效果：单页 or 整份文稿（第 8 轮合并了切换族 3 个工具）。

- \`target\`: slide（默认，只改 \`slideIndex\` 那一页）/ all（整份文稿统一）
- \`effect\`（**必填**）: fade 淡出 / push 推入 / wipe 擦除 / split 拆分 / reveal 揭开 / cover 覆盖 / dissolve 溶解 / curtains 帷幕 / **none 取消切换**

什么时候用：
- "给第1页加个淡出切换" → { target:'slide', slideIndex:1, effect:'fade', duration:1 }
- "所有页统一用推入" → { target:'all', effect:'push' }
- "第2页不要切换效果" → { target:'slide', slideIndex:2, effect:'none' }

取消切换用 \`effect:'none'\`，不要去找"删除切换"的工具 —— 切换本来就是"切走时播放的东西"，
把它设成没有就等价于删掉，而且这样不会因为名字里有"删/删掉"而弹破坏性确认框。

真限制：没有"切换音效"参数。WPS 的 COM 接口里没有"切走时播一段音频"这个能力
（那是 PowerPoint 的 AddMediaObject2），所以本工具不提供该参数 —— 要音效得先在 WPS 里录进切换设置。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      target: {
        type: 'string',
        enum: ['slide', 'all'],
        description: '作用范围：slide 单页（默认）/ all 整份文稿',
      },
      slideIndex: {
        type: 'number',
        description: '幻灯片索引（从1开始）；target=all 时忽略',
      },
      effect: {
        type: 'string',
        description: '切换效果：fade / push / wipe / split / reveal / cover / dissolve / curtains；none = 取消切换',
      },
      duration: {
        type: 'number',
        description: '切换持续时间（秒），默认1秒',
      },
    },
    required: [],
  },
};

const TRANS_NAME_MAP: Record<string, string> = {
  fade: '淡出', push: '推入', wipe: '擦除', split: '拆分',
  reveal: '揭开', cover: '覆盖', dissolve: '溶解', curtains: '帷幕',
  none: '取消切换',
};

export const setTransitionHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, duration } = args as {
    slideIndex?: number;
    duration?: number;
  };
  const target = ((args.target as string) || 'slide') as 'slide' | 'all';
  // 第 2 轮起的旧参数名是 transition；旧名调用经别名进来时带的是它
  const effect = (args.effect as string) || (args.transition as string);

  if (!effect) {
    return fail(
      `必须给 effect（切换效果名，或 none 表示取消切换）`,
      'Missing required parameter: effect'
    );
  }
  if (target === 'slide' && !slideIndex) {
    return fail(
      `target='slide' 时必须给 slideIndex`,
      'Missing required parameter: slideIndex'
    );
  }

  // 整份文稿走 PS 侧的 applyTransitionToAll：它是独立的动作，
  // 而 setSlideTransition 的循环版不在这个 case 里，硬塞 all 进去等于让 PS 认一个它不认的键。
  const r =
    target === 'all'
      ? await call<{ slideCount?: number }>(
          'applyTransitionToAll',
          { effect, duration: duration || 1 }
        )
      : await call<{ slideCount?: number }>(
          'setSlideTransition',
          { slideIndex, effect, duration: duration || 1 }
        );

  if (!r.ok) return fail(errText('设置切换效果失败', r), r.error);
  const lines = [
    `切换效果设置成功！`,
    target === 'all' ? `范围: 全部幻灯片` : `幻灯片: 第 ${slideIndex} 页`,
    `切换效果: ${TRANS_NAME_MAP[effect] || effect}`,
    `持续时间: ${duration || 1} 秒`,
  ];
  if (r.data && r.data.slideCount) {
    lines.push(`已应用: ${r.data.slideCount} 页`);
  }
  return ok(lines.join('\n'), r.warnings);
};

// ============================================================
// 4. wps_ppt_add_speaker_notes - 添加演讲者备注
// ============================================================

export const addSpeakerNotesDefinition: ToolDefinition = {
  name: 'wps_ppt_add_speaker_notes',
  description: `设置（覆盖）指定幻灯片的演讲者备注。

注意：是**整段替换**，不是追加——ps1 侧直接把备注页文本整体赋值（审计三.4 与行为对齐）。
要在原备注后补内容，先 wps_ppt_get_slide_notes 读回来拼接再写。

使用场景：
- "给第1页添加演讲者备注"
- "在备注中写上提示词"
- "把第3页的备注换成新稿"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number', description: '幻灯片索引（从1开始）' },
      notes: { type: 'string', description: '演讲者备注内容' },
    },
    required: ['slideIndex', 'notes'],
  },
};

export const addSpeakerNotesHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, notes } = args as {
    slideIndex: number;
    notes: string;
  };

  const r = await call('setSlideNotes', { slideIndex, notes });
  if (!r.ok) return fail(errText('添加演讲者备注失败', r), r.error);
  return ok(
    `演讲者备注添加成功！\n幻灯片: 第 ${slideIndex} 页\n备注内容: "${notes.length > 50 ? notes.substring(0, 50) + '...' : notes}"`,
    r.warnings
  );
};

/**
 * 导出所有页面效果相关的Tools
 */
export const pageEffectsTools: RegisteredTool[] = [
  { definition: setAnimationDefinition, handler: setAnimationHandler },
  { definition: setBackgroundDefinition, handler: setBackgroundHandler },
  { definition: setTransitionDefinition, handler: setTransitionHandler },
  { definition: addSpeakerNotesDefinition, handler: addSpeakerNotesHandler },
];

export default pageEffectsTools;
