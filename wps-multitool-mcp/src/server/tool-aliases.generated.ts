/**
 * Input: 无（纯常量）
 * Output: 旧工具名 → 规范工具名的映射与解析函数
 * Pos: 别名表【生成物】。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 别手改这个文件：真相源是 scripts/manual/tool-aliases.json，
 * 生成器是 scripts/manual/gen-tool-aliases.js，共 24 条。
 *
 * 别名不注册成工具：它不出现在 tools/list、不计入工具总数，只在 registry 的入口被解析成规范名。
 * 为什么只做一跳解析：链式别名（a→b→c）会让"到底谁是真的"变成谜，一份表里就该一步到位。
 */

/** 旧名 → 规范名。键是已下线的旧工具名，值是仍在注册表里的规范名 */
export const TOOL_ALIASES: Readonly<Record<string, string>> = {
  "wps_common_ping": "wps_check_connection",
  "wps_excel_delete_column": "wps_excel_delete_columns",
  "wps_excel_insert_column": "wps_excel_insert_columns",
  "wps_excel_insert_row": "wps_excel_insert_rows",
  "wps_excel_zoom": "wps_excel_set_zoom",
  "wps_get_active_document": "wps_word_get_active_document",
  "wps_get_cell_value": "wps_excel_get_cell_value",
  "wps_insert_text": "wps_word_insert_text",
  "wps_ppt_add_animation": "wps_ppt_set_animation",
  "wps_ppt_add_animation_preset": "wps_ppt_set_animation",
  "wps_ppt_add_chart": "wps_ppt_insert_ppt_chart",
  "wps_ppt_add_emphasis_animation": "wps_ppt_set_animation",
  "wps_ppt_apply_transition_to_all": "wps_ppt_set_transition",
  "wps_ppt_duplicate_slide": "wps_ppt_copy_slide",
  "wps_ppt_remove_slide_transition": "wps_ppt_set_transition",
  "wps_ppt_set_animation_order": "wps_ppt_set_animation",
  "wps_ppt_set_background_color": "wps_ppt_set_background",
  "wps_ppt_set_background_gradient": "wps_ppt_set_background",
  "wps_ppt_set_background_image": "wps_ppt_set_background",
  "wps_ppt_set_slide_background": "wps_ppt_set_background",
  "wps_ppt_set_slide_notes": "wps_ppt_add_speaker_notes",
  "wps_ppt_set_slide_transition": "wps_ppt_set_transition",
  "wps_set_cell_value": "wps_excel_set_cell_value",
  "wps_word_generate_doc_toc": "wps_word_generate_toc",
};

/**
 * 旧名 → 该旧名的参数映射。只收带映射的那些（纯改名的旧名不需要任何处理）。
 *
 * 族合并的失效方式有一半在参数上：光把名字指向合并工具，旧调用会 Missing required parameter，
 * 或者更糟 —— newOrder 这类改名被静默忽略，工具按默认值把动画排到一个错的位置。
 * rename = 旧参数名 → 新参数名（调用方没给新键时才搬，不覆盖规范名自己收的参数）；
 * set = 【缺了才补】的固定值（调用方显式给了就听调用方的）。
 */
export interface ToolAliasParamMap {
  rename?: Readonly<Record<string, string>>;
  set?: Readonly<Record<string, unknown>>;
}

export const TOOL_ALIAS_PARAMS: Readonly<Record<string, ToolAliasParamMap>> = {
  "wps_ppt_add_animation": { set: {"op":"add","category":"entrance"} },
  "wps_ppt_add_animation_preset": { rename: {"preset":"effect"}, set: {"op":"add","category":"entrance"} },
  "wps_ppt_add_emphasis_animation": { set: {"op":"add","category":"emphasis"} },
  "wps_ppt_apply_transition_to_all": { set: {"target":"all"} },
  "wps_ppt_remove_slide_transition": { set: {"target":"slide","effect":"none"} },
  "wps_ppt_set_animation_order": { rename: {"newOrder":"order"}, set: {"op":"reorder"} },
  "wps_ppt_set_slide_transition": { set: {"target":"slide"} },
};

/** 全部别名（旧名）。用于"这个请求是不是旧名"的判断与门禁遍历 */
export const ALIAS_NAMES: readonly string[] = Object.keys(TOOL_ALIASES);

/**
 * 把（可能是旧名的）工具名解析成规范名。
 * 不是别名就原样返回 —— 调用方不用先判断再解析，少一个分支就少一处漏判。
 */
export function resolveToolAlias(name: string): string {
  return TOOL_ALIASES[name] ?? name;
}

/** 这个名字是不是已下线的旧名（用于在结果里告诉调用方"你用的是旧名"） */
export function isToolAlias(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(TOOL_ALIASES, name);
}

/**
 * 把旧名的参数搬成规范名的参数。非别名、或没有参数映射时原样返回同一个对象
 * （不复制：绝大多数调用走的是这条路，复制纯属白花 CPU）。
 */
export function applyAliasParams(
  name: string,
  args: Record<string, unknown>
): Record<string, unknown> {
  const map = TOOL_ALIAS_PARAMS[name];
  if (!map) return args;

  const out: Record<string, unknown> = { ...(args || {}) };
  for (const [oldKey, newKey] of Object.entries(map.rename || {})) {
    if (oldKey in out && !(newKey in out)) {
      out[newKey] = out[oldKey];
      delete out[oldKey];
    }
  }
  for (const [key, value] of Object.entries(map.set || {})) {
    if (out[key] === undefined) out[key] = value;
  }
  return out;
}
