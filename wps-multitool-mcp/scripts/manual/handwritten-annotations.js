/**
 * 单一真相源：哪些工具的注解是【手写】的、而不是从动作名推断出来的。
 *
 * 为什么单开一个文件
 * ------------------
 * 这份名单有两处要用：tool-annotations.ts（写注解本体）与
 * gen-tool-annotations.js（生成时跳过它们，不让名字推断去猜）。
 * 两处各写一份必然漂 —— 而漂了的后果是【静默的】：
 * 一边加了手写注解、另一边没加，那个工具就悄悄退回"按名字推断"，
 * 而推断对编排类工具（wps_common_batch）是猜错的（会判成非破坏性、幂等）。
 *
 * 生成器里有断言卡这两处同步：不同步就非 0 退出。
 *
 * 进这份名单的判断标准：它的行为【不由某个固定动作决定】。
 *   1) 动作名是运行期参数（wps_execute_method 能执行任意 WPS 动作）
 *   2) 它是编排别人的容器，本身不发指令（wps_common_batch 转手调注册表）
 *   3) 副作用不在 COM 里（快照系列：只调只读的 getActiveDocumentState，
 *      真正写盘的是 Node 侧的 fs 复制 —— 按动作名推断会得出"只读"，是错的）
 * 这三类都不该靠工具名猜。
 *
 * 第 3 轮新增一条（第 4 类）：动作名【带点】，观测器的合法标识符判定直接把它归成 unresolved。
 *   wps_ppt_unify_font → slide.unifyFont。
 * 注解值是人工判定的，写在 tool-annotations.ts 的 EXTRA_ANNOTATIONS 里（改那边记得同步这里）。
 */
module.exports = {
  EXTRA_ANNOTATIONS_KEYS: [
    'wps_execute_method',
    'wps_common_batch',
    'wps_common_snapshot_create',
    'wps_common_snapshot_restore',
    'wps_common_snapshot_list',
    'wps_ppt_unify_font',
    // 第 5 轮补：这 13 个不发 COM 动作（内置/本地）或动作名查不到，注解只能手写。
    // 少了它们，生成器会把注解记成空数组而没人负责 —— 覆盖从 284/284 掉到 266/284。
    'wps_cache_data',
    'wps_get_cached_data',
    'wps_list_cache',
    'wps_clear_cache',
    'wps_word_proofread_basic',
    'wps_list_tool_groups',
    'wps_load_tool_group',
    'wps_unload_tool_group',
    'wps_search_tools',
    'wps_route_intent',
    // 第 1 期第 2 章第 1 轮移出：wps_word_set_text_color / wps_word_set_line_spacing
    // 原先在这里，是因为它们调的动作名在 ps1 里查不到（setTextColor / setLineSpacing，零命中），
    // 动作名推不出来，只能挂手写注解表。动作名修对之后（setFont / setParagraph），
    // 它们能观测到真实动作名了，就不再需要"手写注解兜底"这份特例。
    'wps_ppt_set_active_target',
  ],
};
