/**
 * Input: 工具 inputSchema 与调用入参
 * Output: 校验结果（null = 通过；否则是给调用方的完整错误消息）
 * Pos: 入参真校验（JSON Schema 子集：type / enum / required / 数值范围 / minItems / maxItems /
 *      数组元素类型一层）。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么自写子集校验器而不用 zod：zod 在依赖树里是 @modelcontextprotocol/sdk 的传递依赖，
 * 不是本包直接依赖——import 它就得把它声明进 dependencies（新增依赖，须决策者批）。
 * 子集覆盖本仓 287 个工具 schema 实际用到的全部特性（type/enum/required/items/minItems/maxItems），
 * 数值范围字段（minimum/maximum）一并支持，当前 schema 暂未用到、留作批量上限类工具的护栏。
 *
 * 纪律（代码技法卡 A 段）：
 * - 不做隐式类型转换：传 "3" 给 number 参数 = 明确报错，不静默转。
 * - 只校验 schema 里【声明过的】参数；未声明键（路由层的 target 透传等）不拦。
 * - 数组元素类型只查一层（元素本身的类型）；二维数组的第二层不查——
 *   既有 schema 有"声明 string、实际收数字"的宽松声明（write_range 的 data），
 *   严格查第二层会误杀合法入参，违反"合法入参零变化"红线。
 * - 枚举只对标量值比对（数组/对象不做枚举比对）。
 * - 值为 undefined/null 的已声明参数不查类型（required 检查已单独管"缺失"；
 *   可选参数传 null 视为"没传"，维持既有语义）。
 */

import { ToolDefinition, ToolParameterSchema } from '../types/tools';

/** 值的简短描述（进错误文案，帮助调用方一眼看到传了什么） */
function describeValue(value: unknown): string {
  if (value === null) return 'null';
  if (Array.isArray(value)) return `array(长度 ${value.length})`;
  const t = typeof value;
  if (t === 'string') {
    const v = value as string;
    const shown = v.length > 30 ? `${v.slice(0, 30)}…` : v;
    return `string("${shown}")`;
  }
  if (t === 'number' || t === 'boolean' || t === 'bigint') return `${t}(${String(value)})`;
  if (t === 'object') return `object(键: ${Object.keys(value as object).slice(0, 5).join(', ') || '无'})`;
  return `${t}(${String(value)})`;
}

/** 单值校验：返回问题描述（null = 通过） */
function checkValue(value: unknown, prop: ToolParameterSchema): string | null {
  const t = prop.type;
  const actual = Array.isArray(value) ? 'array' : typeof value;

  if (t === 'number') {
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      return `应为 number，实为 ${describeValue(value)}`;
    }
    if (prop.minimum !== undefined && value < prop.minimum) {
      return `应 ≥ ${prop.minimum}，实为 ${value}`;
    }
    if (prop.maximum !== undefined && value > prop.maximum) {
      return `应 ≤ ${prop.maximum}，实为 ${value}`;
    }
  } else if (actual !== t) {
    return `应为 ${t}，实为 ${describeValue(value)}`;
  }

  // 枚举：仅标量值比对
  if (prop.enum && (typeof value === 'string' || typeof value === 'number')) {
    if (!(prop.enum as readonly unknown[]).includes(value)) {
      return `应为枚举值之一 [${prop.enum.join(' | ')}]，实为 ${describeValue(value)}`;
    }
  }

  if (Array.isArray(value)) {
    if (prop.minItems !== undefined && value.length < prop.minItems) {
      return `至少需要 ${prop.minItems} 个元素，实为 ${value.length}`;
    }
    if (prop.maxItems !== undefined && value.length > prop.maxItems) {
      return `至多 ${prop.maxItems} 个元素，实为 ${value.length}`;
    }
    // 数组元素类型（一层嵌套：只查元素类型本身，不再深入元素内部）
    if (prop.items && value.length > 0) {
      for (let i = 0; i < value.length; i++) {
        const el = value[i];
        const elActual = Array.isArray(el) ? 'array' : typeof el;
        if (elActual !== prop.items.type) {
          return `第 ${i + 1} 个元素应为 ${prop.items.type}，实为 ${describeValue(el)}`;
        }
      }
    }
  }

  return null;
}

/**
 * 校验入参。返回 null = 通过；否则返回完整错误消息（调用方原样回给模型）。
 * 注册表在 handler 执行前调用——校验失败请求不进 COM。
 */
export function validateToolArgs(
  definition: ToolDefinition,
  args: Record<string, unknown>
): string | null {
  const { inputSchema } = definition;
  const props = inputSchema.properties || {};
  const required = inputSchema.required || [];

  // 1) 必填缺失
  for (const param of required) {
    if (!(param in args) || args[param] === undefined || args[param] === null) {
      const expect = props[param];
      return `[勿重试·先改参数] 缺少必填参数: ${param}${expect ? `（类型应为 ${expect.type}）` : ''}`;
    }
  }

  // 2) 已声明参数的类型/枚举/范围校验（未声明键不拦；null 视为没传）
  for (const [key, value] of Object.entries(args)) {
    const prop = props[key];
    if (!prop || value === undefined || value === null) continue;
    const problem = checkValue(value, prop);
    if (problem) {
      return `[勿重试·先改参数] 参数 ${key} ${problem}`;
    }
  }

  return null;
}
