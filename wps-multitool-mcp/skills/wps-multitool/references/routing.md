# 路由与选型

## 四个组

| 组 | 装什么 | 什么时候加载 |
|---|---|---|
| `core` | 工具组管理、搜索、意图识别 | 永远加载，不用管 |
| `word` | 文字：样式、修订、表格、交叉引用、索引、页眉页脚 | 动文档时 |
| `excel` | 表格：公式、图表、单元格、区域、工作表 | 动表格时 |
| `ppt` | 演示：版式、动画、母版、图表、配色 | 动演示时 |
| `common` | 跨应用：保存、建文档、状态、修复、格式转换 | 任何写操作前后 |

默认只列 40 个基础工具。要用别的，先 `wps_load_tool_group`。

## 什么时候别信意图识别

`wps_route_intent` 是关键词打分，组间语义差别大时它很准；下面这些情况**别硬猜**，
直接 `wps_search_tools` 搜中文关键词（"插入表格""改背景""做动画"都能命中）：

- 一句话里同时提到两个应用（"把这张表的数据做成 PPT"）
- 诉求是排版/版式类（"这篇论文排得像正式期刊"）
- 搜出来的结果里 `notFor` 说"这个不干 X"—— 说明还有更合适的

## 近名工具分辨表

本 MCP 里有大量近名工具。选错**不会报错**，只会把事情改到别处，用户还看不出来。

### 目录 / 索引 / 交叉引用（最常混）

| 需求 | 用 | 不是什么 |
|---|---|---|
| 按标题级别生成目录 | `wps_word_generate_toc` | 不是重建整篇目录（会覆盖） |
| 整篇重建目录 | `wps_word_generate_doc_toc` | 不是插段落 |
| 插"见图 3"这类引用 | `wps_word_cross_reference` | 不是参考文献、不是索引 |
| 插关键词索引表 | `wps_word_index` | 不是目录 |

### 样式：套用 vs 定义

| 需求 | 用 |
|---|---|
| 给已有段落套"标题 1"样式 | `wps_word_apply_style` |
| 新建一个自定义样式 | `wps_word_add_style` |
| 改样式定义本身 | `wps_word_modify_style` |
| 删样式定义 | `wps_word_delete_style` |
| 插入一级标题段落 | `wps_word_add_heading` |

### 表格：插入 vs 改造

| 需求 | 用 |
|---|---|
| 新建一张表 | `wps_word_insert_table` |
| 改已有表的样式/边框 | `wps_word_format_table` / `wps_word_set_cell_borders` |
| 合并单元格 | `wps_word_merge_table_cells` |
| 拆单元格 | `wps_word_split_cell` |

### Excel：单条 vs 批量

| 需求 | 用 |
|---|---|
| 写一个格子 | `wps_excel_set_cell_value` |
| 写一大片区域 | `wps_excel_write_range` |
| 写公式 | `wps_excel_set_formula` |
| 插一行 / 多行 | `wps_excel_insert_row` / `wps_excel_insert_rows` |
| 缩放（`zoom` 与 `set_zoom` 重复） | 二选一，优先 `set_zoom` |

### 修订 vs 批注

| 需求 | 用 |
|---|---|
| 查看/接受/拒绝已有修订 | `wps_word_revisions` |
| 插入一条新批注 | `wps_word_insert_comment` |

### PPT 的背景与动画

`background` 有 6 个、`animation` 有 5 个，全部靠名字分不出来。
**先搜再读 `notFor`**，不要凭名字挑：不同工具管的是幻灯片背景 / 母版背景 /
图表区背景，纯色 / 图片 / 渐变又是三回事；动画分普通 / 预设 / 强调 / 顺序 / 清除。

## 搜索返回什么

```
name        工具名
group       所属组（未加载时 listed: false）
listed      当前是否已列出
aliases     中文别名
examples    什么诉求用它
notFor      它不干什么、该用谁   ← 必读
```