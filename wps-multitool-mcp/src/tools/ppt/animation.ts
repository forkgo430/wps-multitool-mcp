/**
 * Input: PPT 动画列表与动画删除操作参数
 * Output: 动画查询/删除结果
 * Pos: PPT 动画工具实现。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * PPT动画Tools - 动画查询与删除模块
 *
 * 只留两个不能合并的工具：
 * - wps_ppt_get_animations: 查询动画列表（只读，没东西可并入别的工具）
 * - wps_ppt_remove_animation: 删除单个动画（破坏性，混进"设置"工具会污染注解）
 *
 * 第 8 轮合并掉了本文件原先的 7 个工具，它们都是别名，能力没丢：
 * - wps_ppt_add_animation / add_animation_preset / add_emphasis_animation / set_animation_order
 *   → wps_ppt_set_animation（op: 'add' | 'reorder'，现居 page-effects.ts）
 * - wps_ppt_set_slide_transition / apply_transition_to_all / remove_slide_transition
 *   → wps_ppt_set_transition（target: 'slide' | 'all'，effect: 'none' 即取消切换，现居 page-effects.ts）
 * 旧名走 scripts/manual/tool-aliases.json 的别名表调用，参数映射也在那张表里。
 */

import {
  ToolDefinition,
  ToolHandler,
  ToolCallResult,
  ToolCategory,
  RegisteredTool,
} from '../../types/tools';
import { call, ok, fail, errText } from './helpers';

/**
 * MsoAnimEffect 编号 → 中文名。表抄自 wps-com.ps1 里 addAnimation 的 $effectMap（同一份真相），
 * 补上入口/强调里常用的几个。getAnimations 回的是编号，不是名字。
 * 认不出的就显示"效果 (#编号)"，不猜。
 */
const EFFECT_ID_NAME_MAP: Record<number, string> = {
  1: '出现', 2: '飞入', 3: '百叶窗', 4: '框', 5: '棋盘', 6: '圆形', 7: '爬行',
  8: '菱形', 9: '溶解', 10: '淡入', 12: '切入', 13: '加号', 14: '随机条',
  15: '旋转', 22: '擦除', 23: '百叶窗', 26: '弹跳', 28: '摆动', 30: '上浮',
  34: '升起', 36: '细线', 37: '展开', 38: '鞭状', 46: '收缩', 50: '伸展',
  53: '缩放', 59: '放大', 62: '强调颜色', 63: '脉冲',
};

// ==================== 1. 移除动画效果 ====================

/**
 * 移除动画效果
 * 从指定幻灯片中移除指定的动画
 */
export const removeAnimationDefinition: ToolDefinition = {
  name: 'wps_ppt_remove_animation',
  description: `移除幻灯片中指定的动画效果。

使用场景：
- "删除第1页的第2个动画"
- "移除这个动画效果"

先想清楚删的是哪一个：wps_ppt_get_animations 能列出序号，确认序号再删。
序号是本页动画窗格里的顺序（1 开始），加完动画会变。`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
      animationIndex: {
        type: 'number',
        description: '动画索引（从1开始）',
      },
    },
    required: ['slideIndex', 'animationIndex'],
  },
};

export const removeAnimationHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex, animationIndex } = args as {
    slideIndex: number;
    animationIndex: number;
  };

  const r = await call('removeAnimation', { slideIndex, animationIndex });
  if (!r.ok) return fail(errText('移除动画失败', r), r.error);
  return ok(`动画移除成功！\n幻灯片: 第 ${slideIndex} 页\n已移除第 ${animationIndex} 个动画`, r.warnings);
};

// ==================== 2. 获取动画列表 ====================

/**
 * 获取动画列表
 * 查看指定幻灯片上所有动画效果
 */
export const getAnimationsDefinition: ToolDefinition = {
  name: 'wps_ppt_get_animations',
  description: `获取幻灯片上所有动画效果的列表。

使用场景：
- "查看第1页有哪些动画"
- "列出这页的所有动画"`,
  category: ToolCategory.PRESENTATION,
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: {
        type: 'number',
        description: '幻灯片页码（从1开始）',
      },
    },
    required: ['slideIndex'],
  },
};

export const getAnimationsHandler: ToolHandler = async (
  args: Record<string, unknown>
): Promise<ToolCallResult> => {
  const { slideIndex } = args as { slideIndex: number };

  const r = await call<{
    animations?: Array<{
      index: number;
      shapeName: string;
      effectType?: number;
      effect?: string | number;
      duration?: number;
      trigger?: string;
    }>;
  }>('getAnimations', { slideIndex });
  if (!r.ok) return fail(errText('获取动画列表失败', r), r.error);
  const animations = r.data?.animations;

  if (!animations || animations.length === 0) {
    return ok(`第 ${slideIndex} 页没有动画效果。`, r.warnings);
  }

  let output = `第 ${slideIndex} 页的动画效果（共 ${animations.length} 个）：\n`;
  animations.forEach((anim) => {
    const rawEffect = anim.effectType ?? anim.effect;
    const effectText =
      typeof rawEffect === 'number'
        ? `${EFFECT_ID_NAME_MAP[rawEffect] ?? '效果'} (#${rawEffect})`
        : String(rawEffect ?? '未知效果');
    const dur = anim.duration === undefined || anim.duration === null ? '' : `，时长 ${anim.duration}s`;
    output += `  ${anim.index}. [${anim.shapeName}] ${effectText}${dur}\n`;
  });
  output += `要调顺序用 wps_ppt_set_animation(op='reorder')，要删某一条用 wps_ppt_remove_animation。`;

  return ok(output, r.warnings);
};

/**
 * 导出动画查询/删除相关的Tools
 */
export const animationTools: RegisteredTool[] = [
  { definition: removeAnimationDefinition, handler: removeAnimationHandler },
  { definition: getAnimationsDefinition, handler: getAnimationsHandler },
];

export default animationTools;
