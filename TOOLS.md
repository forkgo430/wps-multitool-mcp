# 工具清单

> 本文件由 `wps-multitool-mcp/scripts/manual/gen-tools-doc.js` 从**运行时注册表**生成，请勿手改。
> 改完工具跑一次 `node scripts/manual/gen-tools-doc.js`；CI 会重跑并 `git diff --exit-code`，对不上就红。

**287 个工具**（7 内置 + 280 专业/组工具）· 默认列出 40 个 · 注解覆盖 287/287 · 真机验证 200/287

| 标记 | 含义 |
|---|---|
| 默认档 `是` | 不加载组、直接出现在 `tools/list` 里 |
| 注解 `R` / `D` / `N` | 只读 / 破坏性 / 非幂等（`—` = 无法从动作名判定） |
| 验证 `live` | 该工具在真机上跑通过 |

## core 组 · 常驻管理（5）

| 工具 | 默认档 | 注解 | 验证 | 说明 |
|---|---|---|---|---|
| `wps_list_tool_groups` | 是 | R | live | 列出所有工具组及当前是否已在列。<br><sub>别名：列出工具工具组 / 列举工具工具组 / 工具</sub> |
| `wps_load_tool_group` | 是 | — | live | 加载一个工具组，加载后该组工具才会出现在工具列表里。<br><sub>别名：加载组 / 拉起组 / 把PPT工具弄出来</sub> |
| `wps_route_intent` | 是 | R | live | 意图路由（三层）：一句话进来，返回【能直接行动】的结果，而不是只给一个组名。<br><sub>别名：意图识别 / 该用哪个组 / 这句话是干嘛</sub> |
| `wps_search_tools` | 是 | R | live | 在所有工具里搜索（包括未加载组的工具），返回名字 + 所属组 + 是否已列出 + 中文别名 + 验证状态。<br><sub>别名：找工具 / 搜工具 / 有没有什么工具</sub> |
| `wps_unload_tool_group` | 是 | — | live | 卸载一个工具组，卸载后该组工具不再出现在工具列表里（仍可直接调用，但不推荐）。<br><sub>别名：卸载工具组合 / 工具 / 组合</sub> |

## 通用（21）

| 工具 | 默认档 | 注解 | 验证 | 说明 |
|---|---|---|---|---|
| `wps_common_save` | 是 | — | live | 保存当前文档。<br><sub>别名：保存 / 存盘</sub> |
| `wps_common_save_as` | 是 | DN | live | 将当前文档另存为指定路径和格式。<br><sub>别名：另存为 / 存到 / 保存as</sub> |
| `wps_common_show_window` | 是 | — | live | 把 WPS 窗口显示出来并置于前台。<br><sub>别名：显示窗口 / 展示窗口 / 窗口</sub> |
| `wps_common_status` | 是 | R | live | 查看 WPS MCP 的运行状态：桥（守护进程）这一层 + WPS 三件套这一层。<br><sub>别名：状态 / 桥状态 / 看看还活着吗</sub> |
| `wps_convert_format` | 是 | DN | live | 将当前文档转换为其他格式。<br><sub>别名：格式转换 / 转格式 / docx转xlsx</sub> |
| `wps_convert_to_pdf` | 是 | DN | live | 将当前文档转换为PDF格式。支持Word、Excel、PPT文档。<br><sub>别名：转PDF / 导出PDF / 转成PDF</sub> |
| `wps_execute_method` | 是 | DN | live | 执行自定义WPS API方法<br><sub>别名：直接调COM / 万能方法 / 没有对应工具</sub> |
| `wps_cache_data` |  | — | live | 缓存数据到MCP Server，用于跨应用数据传递。例如：从Excel读取数据后缓存，然后在PPT中使用。<br><sub>别名：暂存数据 / 缓存数据 / 存中间结果</sub> |
| `wps_check_connection` |  | R | live | 检查WPS Office连接状态<br><sub>别名：检查连接 / 连接状态 / 通不通</sub> |
| `wps_clear_cache` |  | D | live | 清除MCP Server中的缓存数据<br><sub>别名：清缓存 / 清空暂存 / 删掉暂存数据</sub> |
| `wps_common_batch` |  | DN | live | 一次执行 1~3 个操作，省掉逐个调用的往返。 |
| `wps_common_get_app_info` |  | R | live | 获取WPS应用的基本信息。<br><sub>别名：获取应用信息 / 读取应用信息 / 查看应用信息</sub> |
| `wps_common_get_selected_text` |  | R | live | 获取当前文档中选中的文本内容。<br><sub>别名：获取选中文本 / 读取选中文本 / 查看选中文本</sub> |
| `wps_common_quit` |  | DN | live | 结束 WPS（Word / Excel / PPT 全部）。<br><sub>别名：退出</sub> |
| `wps_common_repair` |  | — | live | 诊断并在【有限范围内】修复 WPS 自动化常见问题。<br><sub>别名：修复 / 卡死了 / 自救</sub> |
| `wps_common_set_selected_text` |  | — | live | 替换当前文档中选中的文本内容。<br><sub>别名：设置选中文本 / 设定选中文本 / 选中</sub> |
| `wps_common_snapshot_create` |  | N | live | 给当前文档存一份快照（复制磁盘文件到 ~/.wps-multitool-mcp/snapshots）。<br><sub>别名：snapshot创建 / 创建 / 新建</sub> |
| `wps_common_snapshot_list` |  | R | live | 列出本机快照（只读），按时间倒序。可按原文件路径过滤。<br><sub>别名：snapshot列出 / 列出 / 列举</sub> |
| `wps_common_snapshot_restore` |  | DN | live | 把快照恢复回去。 |
| `wps_get_cached_data` |  | R | live | 从MCP Server获取缓存的数据，用于跨应用数据传递。<br><sub>别名：取暂存数据 / 读缓存 / 取暂存结果</sub> |
| `wps_list_cache` |  | R | live | 列出MCP Server中所有缓存的数据键名<br><sub>别名：看缓存列表 / 暂存了哪些 / 列出cache</sub> |

## Word（69）

| 工具 | 默认档 | 注解 | 验证 | 说明 |
|---|---|---|---|---|
| `wps_word_add_heading` | 是 | N | live | 按级别插入一个标题段落（使用 Word 内置 Heading 样式，因此能自动进目录）。<br><sub>别名：加标题 / 插入标题 / 几级标题</sub> |
| `wps_word_find_replace` | 是 | N | live | 在Word文档中查找并替换文本。<br><sub>别名：查找替换 / 搜索替换 / 替换</sub> |
| `wps_word_get_document_text` | 是 | R | live | 获取当前Word文档的文本内容。<br><sub>别名：获取文档文本 / 读取文档文本 / 查看文档文本</sub> |
| `wps_word_get_open_documents` | 是 | R | live | 获取当前WPS Writer中所有已打开的文档列表。<br><sub>别名：获取打开文档 / 读取打开文档 / 查看打开文档</sub> |
| `wps_word_get_outline` | 是 | R | live | 抽取文档的标题大纲（按段落的大纲级别 1-9），返回层级 + 段落序号。<br><sub>别名：获取大纲 / 读取大纲 / 查看大纲</sub> |
| `wps_word_insert_table` | 是 | N | live | 在Word文档光标位置插入表格<br><sub>别名：插入表格 / 加表格 / 建表</sub> |
| `wps_word_insert_text` | 是 | N | live | 在Word文档中插入文本。<br><sub>别名：插入文本 / 新增文本 / 加入文本</sub> |
| `wps_word_open_document` | 是 | — | live | 打开指定路径的Word文档。<br><sub>别名：打开文档 / 文档</sub> |
| `wps_word_set_default_font` | 是 | — | live | 设置文档默认字体（改 Normal 样式，正文与后续新建段落都会跟着变）。<br><sub>别名：设置默认字体 / 设定默认字体 / 默认</sub> |
| `wps_word_switch_document` | 是 | — | live | 切换到指定名称的文档。<br><sub>别名：切换文档 / 文档</sub> |
| `wps_word_add_caption` |  | N | live | 在当前位置插入题注（自动编号，如"图 1""表 2"）。<br><sub>别名：添加题注 / 新增题注 / 加题注</sub> |
| `wps_word_add_footnote` |  | N | live | 插入脚注（页脚注释，自动编号）。<br><sub>别名：添加脚注 / 新增脚注 / 加脚注</sub> |
| `wps_word_add_list` |  | N | live | 在当前位置插入项目符号或编号列表（一次写多条）。<br><sub>别名：添加列出 / 新增列出 / 加列出</sub> |
| `wps_word_add_style` |  | N | live | 新建一个样式（不直接应用到任何段落）。<br><sub>别名：添加样式 / 新增样式 / 加样式</sub> |
| `wps_word_apply_style` |  | — | live | 应用Word样式到当前选中区域或指定范围。<br><sub>别名：套样式 / 应用样式 / 改成标题样式</sub> |
| `wps_word_close_document` |  | DN | live | 关闭指定的 WPS 文字文档（不填关当前目标文档），可选是否保存。<br><sub>别名：关闭文档 / 关掉这个Word / 文档</sub> |
| `wps_word_convert_table_to_text` |  | DN | live | 把文档里的表格转换成纯文本（按分隔符拼接），表格本身消失。<br><sub>别名：表格转文本 / 表格拆成文字 / 去表格线</sub> |
| `wps_word_create_document` |  | N | live | 新建一个空白 WPS 文字文档，并把它设为后续 word 操作的目标。<br><sub>别名：新建文档 / 建个Word / 创建文档</sub> |
| `wps_word_cross_reference` |  | N | live | 在【当前光标位置】插入交叉引用（指向某个标题/书签），或列出可引用的目标。<br><sub>别名：交叉引用 / 引用图表 / 见图</sub> |
| `wps_word_delete_comment` |  | DN | live | 删除一条批注：按序号（get_comments 返回的 #n）或按内容片段定位。<br><sub>别名：删批注 / 去掉批注 / 清理批注</sub> |
| `wps_word_delete_paragraph` |  | DN | live | 删除段落。从 index 起删 count 段（1 基，越界自动停下并如实报告删了几段）。<br><sub>别名：删除段落 / 删掉段落 / 去掉段落</sub> |
| `wps_word_delete_style` |  | DN | live | 删除一个自定义样式。<br><sub>别名：删除样式 / 删掉样式 / 去掉样式</sub> |
| `wps_word_enable_track_changes` |  | — | live | 开启或关闭Word文档的修订模式（Track Changes）。<br><sub>别名：开启跟踪更改 / 启用跟踪更改 / 跟踪</sub> |
| `wps_word_find_in_document` |  | R | live | 在Word文档中查找文本并返回位置信息，不执行替换操作。<br><sub>别名：查找in文档 / 搜索in文档 / 文档</sub> |
| `wps_word_format_table` |  | — | live | 设置文档里某个表格的样式 / 表头加粗 / 边框。<br><sub>别名：格式表格 / 表格</sub> |
| `wps_word_generate_toc` |  | N | live | 根据文档中的标题样式自动生成目录。<br><sub>别名：生成目录 / 插入目录 / 自动目录</sub> |
| `wps_word_get_active_document` |  | R | live | 获取当前WPS Writer活动文档的基本信息<br><sub>别名：获取当前文档 / 读取当前文档 / 查看当前文档</sub> |
| `wps_word_get_bookmarks` |  | R | live | 列出文档里所有书签：名称、位置、锚定文本。<br><sub>别名：书签列表 / 有哪些书签 / 查看书签</sub> |
| `wps_word_get_cell_text` |  | R | live | 读取表格里某个单元格的文本。表格类操作前先读一下，确认坐标对不对。<br><sub>别名：获取单元格文本 / 读取单元格文本 / 查看单元格文本</sub> |
| `wps_word_get_comments` |  | R | live | 列出文档里所有批注：序号、作者、日期、内容、所批注的原文片段。<br><sub>别名：看批注 / 批注列表 / 有哪些批注</sub> |
| `wps_word_get_document_stats` |  | R | live | 读文档规模：字数 / 页数 / 段落数 / 字符数 / 节数。<br><sub>别名：文档统计 / 多少字 / 多少页</sub> |
| `wps_word_get_paragraphs` |  | R | live | 获取Word文档的段落结构信息，返回每段的文本、样式和字符位置。<br><sub>别名：获取段落 / 读取段落 / 查看段落</sub> |
| `wps_word_get_table_data` |  | R | live | 读整张表格的内容（二维文本，按序号定位）。<br><sub>别名：读表格 / 表格内容 / 表里有什么</sub> |
| `wps_word_get_tables` |  | R | live | 列出文档里所有表格：序号、行数、列数、第一行预览。<br><sub>别名：有几个表 / 表格列表 / 文档里的表格</sub> |
| `wps_word_get_track_changes_status` |  | R | live | 获取当前文档的修订模式状态。<br><sub>别名：获取跟踪更改状态 / 读取跟踪更改状态 / 查看跟踪更改状态</sub> |
| `wps_word_index` |  | N | live | 给文字编索引条目，或在光标处插入索引目录。<br><sub>别名：索引 / 关键词索引 / 索引表</sub> |
| `wps_word_insert_bookmark` |  | N | live | 在当前光标位置或选中区域插入书签。<br><sub>别名：插入书签 / 新增书签 / 加入书签</sub> |
| `wps_word_insert_comment` |  | N | live | 在Word文档选中内容处插入批注<br><sub>别名：加批注 / 插入批注 / 评论</sub> |
| `wps_word_insert_footer` |  | N | live | 设置页脚内容。<br><sub>别名：插入页脚 / 新增页脚 / 加入页脚</sub> |
| `wps_word_insert_header` |  | N | live | 设置页眉内容。<br><sub>别名：插入页眉 / 新增页眉 / 加入页眉</sub> |
| `wps_word_insert_hyperlink` |  | N | live | 在文档中插入超链接（可指定显示文字与插入位置）。<br><sub>别名：超链接 / 加链接 / 插入链接</sub> |
| `wps_word_insert_image` |  | N | live | 在Word文档中插入图片。<br><sub>别名：插入图片 / 新增图片 / 加入图片</sub> |
| `wps_word_insert_page_break` |  | N | live | 在文档光标位置插入分页符<br><sub>别名：插入页面分隔符 / 新增页面分隔符 / 加入页面分隔符</sub> |
| `wps_word_insert_paragraph` |  | N | live | 插入一个正文段落。给了 index 就插在该段之前，不给就追加到文末。<br><sub>别名：插入段落 / 新增段落 / 加入段落</sub> |
| `wps_word_insert_section_break` |  | N | live | 在光标处插入分节符，把文档分成不同的节（各节可以有不同的页边距/纸张/页眉页脚）。<br><sub>别名：插入节分隔符 / 新增节分隔符 / 加入节分隔符</sub> |
| `wps_word_insert_table_of_figures` |  | N | live | 在文末生成"图表目录"（基于题注标签；所以要先有 add_caption 加的题注）。<br><sub>别名：插入表格offigures / 新增表格offigures / 加入表格offigures</sub> |
| `wps_word_list_styles` |  | R | live | 列出文档里的样式。<br><sub>别名：列出样式 / 列举样式 / 样式</sub> |
| `wps_word_merge_table_cells` |  | N | live | 合并表格里的一块矩形区域。<br><sub>别名：合并单元格 / 合并表格单元格 / 表格</sub> |
| `wps_word_modify_style` |  | — | live | 修改已有样式的属性。<br><sub>别名：修改样式 / 改样式 / 样式</sub> |
| `wps_word_page_number` |  | N | live | 在页眉里插入页码，或设置"首页不同/奇偶页不同"。<br><sub>别名：页面数字 / 页数字 / 数字</sub> |
| `wps_word_proofread_basic` |  | R | live | 对中文文本进行基础校对，检测常见问题。<br><sub>别名：校对基础 / 基础</sub> |
| `wps_word_replace_bookmark_content` |  | DN | live | 替换Word文档中书签的内容。通过书签名定位，替换书签范围内的文本，保持原有格式。<br><sub>别名：替换书签内容 / 书签 / 内容</sub> |
| `wps_word_replace_range` |  | DN | live | 按字符范围精确替换Word文档中的文本。<br><sub>别名：替换区域 / 区域 / 范围</sub> |
| `wps_word_review_mode` |  | — | static | 开启/关闭/查询这份文档的修订模式。 |
| `wps_word_revisions` |  | DN | live | 处理文档的修订（tracked changes）：列出 / 全部接受 / 全部拒绝 / 单条接受或拒绝。<br><sub>别名：修订 / 批注 / 审阅</sub> |
| `wps_word_set_cell_borders` |  | — | live | 设置表格单元格的边框。<br><sub>别名：设置单元格borders / 设定单元格borders / 单元格</sub> |
| `wps_word_set_columns` |  | — | live | 设置分栏（整节等宽多栏）。<br><sub>别名：设置列 / 设定列 / 列</sub> |
| `wps_word_set_font` |  | — | live | 设置字体格式，包括字体名称、字号、加粗、斜体、颜色等。<br><sub>别名：设置字体 / 设定字体 / 字体</sub> |
| `wps_word_set_font_style` |  | — | live | [DEPRECATED·改用 wps_word_set_font] 设置选中文字的字体样式属性<br><sub>别名：设置字体样式 / 设定字体样式 / 字体</sub> |
| `wps_word_set_heading_numbering` |  | — | live | 给标题加章节自动编号（1. / 1.1. / 1.1.1. 这种）。<br><sub>别名：设置标题编号 / 设定标题编号 / 标题</sub> |
| `wps_word_set_line_spacing` |  | — | live | 设置段落行距。<br><sub>别名：设置行距间距 / 设定行距间距 / 行距</sub> |
| `wps_word_set_page_setup` |  | — | live | 设置文档页面布局，包括页面方向和边距。<br><sub>别名：设置页面设置 / 设定页面设置 / 页面</sub> |
| `wps_word_set_paragraph` |  | — | live | 设置当前段落格式（对齐方式、行间距等）<br><sub>别名：设置段落 / 设定段落 / 段落</sub> |
| `wps_word_set_text_color` |  | — | live | 设置选中文字的颜色。颜色可以写颜色名（red / blue / green / black）或十六进制（#FF0000）。<br><sub>别名：设置文本颜色 / 设定文本颜色 / 文本</sub> |
| `wps_word_smart_fill_field` |  | — | live | 智能填写Word模板中的字段。自动识别关键字附近的填写模式，在正确位置插入内容并保持原有格式。<br><sub>别名：智能填充域 / 填充 / 域</sub> |
| `wps_word_split_cell` |  | N | live | 拆分一个已合并的单元格。<br><sub>别名：拆分单元格 / 分割单元格 / 单元格</sub> |
| `wps_word_table_row_col` |  | N | live | 在表格里插入或删除行/列。<br><sub>别名：表格行col / 行</sub> |
| `wps_word_undo` |  | DN | static | 按步撤销这份文档里最近的编辑。 |
| `wps_word_update_fields` |  | — | live | 刷新文档里的所有域：目录、图表目录、交叉引用、页码等。<br><sub>别名：更新域 / 刷新域 / 域</sub> |

## Excel（91）

| 工具 | 默认档 | 注解 | 验证 | 说明 |
|---|---|---|---|---|
| `wps_excel_create_chart` | 是 | N | live | 在Excel中创建图表。支持柱状图、折线图、饼图、散点图等多种类型。<br><sub>别名：创建图表 / 新建图表 / 生成图表</sub> |
| `wps_excel_get_cell_value` | 是 | R | live | 获取Excel指定单元格的值。<br><sub>别名：获取单元格值 / 读取单元格值 / 查看单元格值</sub> |
| `wps_excel_get_open_workbooks` | 是 | R | live | 获取当前所有已打开的Excel工作簿列表。<br><sub>别名：获取打开工作簿 / 读取打开工作簿 / 查看打开工作簿</sub> |
| `wps_excel_get_sheet_list` | 是 | R | live | 获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。<br><sub>别名：获取工作表列出 / 读取工作表列出 / 查看工作表列出</sub> |
| `wps_excel_open_workbook` | 是 | — | live | 打开指定路径的Excel工作簿文件。<br><sub>别名：打开工作簿 / 工作簿 / 表格文件</sub> |
| `wps_excel_read_range` | 是 | R | live | 读取Excel指定范围的单元格数据，返回二维数组格式的数据。<br><sub>别名：读取区域 / 区域 / 范围</sub> |
| `wps_excel_set_cell_value` | 是 | — | live | 设置Excel指定单元格的值。<br><sub>别名：写单元格 / 填值 / 设置单元格内容</sub> |
| `wps_excel_set_formula` | 是 | — | live | 在指定单元格设置Excel公式。公式必须以=开头，支持所有Excel内置函数。<br><sub>别名：写公式 / 设置公式 / 填公式</sub> |
| `wps_excel_switch_workbook` | 是 | — | live | 切换到指定名称的Excel工作簿。<br><sub>别名：切换工作簿 / 工作簿 / 表格文件</sub> |
| `wps_excel_write_range` | 是 | — | live | 向Excel指定范围写入数据。数据格式为二维数组，从指定单元格开始向右下方填充。<br><sub>别名：批量写 / 写区域 / 一次写一片</sub> |
| `wps_excel_add_comment` |  | N | live | 给单元格添加批注（同一格已有批注会被覆盖）。<br><sub>别名：添加批注 / 新增批注 / 加批注</sub> |
| `wps_excel_add_slicer` |  | N | static | 给已创建的 Excel 表格添加切片器（按某列筛选的按钮面板）。<br><sub>别名：添加slicer / 新增slicer / 加slicer</sub> |
| `wps_excel_auto_fill` |  | — | live | [DEPRECATED·改用 wps_excel_fill_series] 自动填充单元格区域。根据源区域的数据模式自动填充到…<br><sub>别名：自动填充 / 填充</sub> |
| `wps_excel_auto_filter` |  | — | live | 对Excel指定范围应用自动筛选。<br><sub>别名：自动筛选 / 筛选 / 过滤</sub> |
| `wps_excel_auto_fit` |  | — | live | 按内容自动调整行高/列宽（等同 Excel 里双击列边界）。<br><sub>别名：自适应 / 自动调整列宽 / 自动列宽</sub> |
| `wps_excel_auto_sum` |  | — | live | 把一段区域求和的结果写进一个单元格（等价于在该格写 =SUM(区域)）。<br><sub>别名：自动求和 / 求和</sub> |
| `wps_excel_clean_data` |  | — | live | 数据清洗工具，支持多种清洗操作的组合。<br><sub>别名：清洗数据 / 清理数据 / 数据</sub> |
| `wps_excel_clear_pivot_table` |  | DN | live | 清空指定透视表的数据与字段布局，保留透视表本身的位置与格式。<br><sub>别名：清空透视表 / 透视表重置 / 把透视表内容清掉</sub> |
| `wps_excel_clear_range` |  | DN | live | 清除指定范围的内容、格式或全部。<br><sub>别名：清除区域 / 清空区域 / 区域</sub> |
| `wps_excel_close_workbook` |  | DN | live | 关闭指定的Excel工作簿，可选是否保存。<br><sub>别名：关闭工作簿 / 工作簿 / 表格文件</sub> |
| `wps_excel_copy_range` |  | N | live | 复制Excel指定范围到目标位置。<br><sub>别名：复制区域 / 拷贝区域 / 区域</sub> |
| `wps_excel_copy_sheet` |  | N | live | 复制当前工作簿中的指定工作表。可指定新名称和插入位置。<br><sub>别名：复制工作表 / 拷贝工作表 / 工作表</sub> |
| `wps_excel_create_pivot_table` |  | N | live | 创建Excel透视表，用于数据汇总和分析。<br><sub>别名：创建数据透视表表格 / 新建数据透视表表格 / 生成数据透视表表格</sub> |
| `wps_excel_create_sheet` |  | N | live | 在当前工作簿中创建新的工作表。可指定名称和插入位置。<br><sub>别名：创建工作表 / 新建工作表 / 生成工作表</sub> |
| `wps_excel_create_table` |  | N | static | 把一块区域创建为"Excel 表格"（结构化引用区）。<br><sub>别名：创建表格 / 新建表格 / 生成表格</sub> |
| `wps_excel_create_workbook` |  | N | live | 新建一个空白Excel工作簿。<br><sub>别名：创建工作簿 / 新建工作簿 / 生成工作簿</sub> |
| `wps_excel_delete_cell_comment` |  | DN | live | 删除Excel单元格上的批注。<br><sub>别名：删除单元格批注 / 删掉单元格批注 / 去掉单元格批注</sub> |
| `wps_excel_delete_chart` |  | DN | live | 删除工作表上的图表。<br><sub>别名：删图表 / 删除图表 / 把图删了</sub> |
| `wps_excel_delete_columns` |  | DN | live | 删除Excel中指定位置的一列或多列。<br><sub>别名：删除列 / 删掉列 / 去掉列</sub> |
| `wps_excel_delete_row` |  | DN | live | [DEPRECATED·改用 wps_excel_delete_rows] 删除指定行。可指定起始行号和删除行数。<br><sub>别名：删除行 / 删一行 / 删掉行</sub> |
| `wps_excel_delete_rows` |  | DN | live | 删除Excel中指定位置的一行或多行。<br><sub>别名：删除多行 / 批量删行 / 删除行</sub> |
| `wps_excel_delete_sheet` |  | DN | live | 删除当前工作簿中的指定工作表。注意：此操作不可撤销。<br><sub>别名：删除工作表 / 删掉工作表 / 去掉工作表</sub> |
| `wps_excel_diagnose_formula` |  | R | live | 诊断公式错误，分析原因并提供修复建议。<br><sub>别名：诊断公式 / 排查公式 / 公式</sub> |
| `wps_excel_evaluate_formula` |  | — | live | 算一个公式的【值】并返回，不改单元格（除非你给了 cell）。<br><sub>别名：求值公式 / 计算公式 / 公式</sub> |
| `wps_excel_export_chart_as_image` |  | — | live | 将工作表中指定的图表导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。<br><sub>别名：导出图表as图片 / 输出图表as图片 / 图表</sub> |
| `wps_excel_export_range_as_image` |  | — | live | 将工作表中指定区域导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。<br><sub>别名：区域转图片 / 导出图片 / 导出区域as图片</sub> |
| `wps_excel_fill_series` |  | — | live | 自动填充序列数据。<br><sub>别名：填充序列 / 序列</sub> |
| `wps_excel_find_in_sheet` |  | R | live | 在工作表里找值，返回所有命中的单元格地址。只读，不改任何内容。<br><sub>别名：表内查找 / 在哪 / 这个数在哪</sub> |
| `wps_excel_find_replace` |  | N | live | 在Excel中查找并替换内容。<br><sub>别名：查找替换 / 搜索替换 / 替换</sub> |
| `wps_excel_freeze_panes` |  | — | live | 冻结/取消冻结窗格。可指定冻结的行和列位置。<br><sub>别名：冻结panes</sub> |
| `wps_excel_generate_formula` |  | R | live | 根据自然语言描述生成Excel公式。这是解决用户"公式不会写"痛点的核心工具。<br><sub>别名：生成公式 / 创建公式 / 公式</sub> |
| `wps_excel_get_cell_comments` |  | R | live | 获取Excel指定范围内的所有批注。不指定范围则获取当前工作表所有批注。<br><sub>别名：获取单元格批注 / 读取单元格批注 / 查看单元格批注</sub> |
| `wps_excel_get_cell_info` |  | R | live | 获取单元格的详细信息（值、公式、格式等）。<br><sub>别名：获取单元格信息 / 读取单元格信息 / 查看单元格信息</sub> |
| `wps_excel_get_formula` |  | R | live | 获取Excel指定单元格的公式。<br><sub>别名：获取公式 / 读取公式 / 查看公式</sub> |
| `wps_excel_get_pivot_tables` |  | R | live | 列出工作簿/工作表里的所有透视表：名字、所在表、占的区域、各有几行/列/值字段。<br><sub>别名：透视表列表 / 有哪些透视表 / 透视表现在什么样</sub> |
| `wps_excel_get_selection` |  | R | live | 获取当前Excel中选中区域的信息，包括范围地址、行列数等。<br><sub>别名：获取选区 / 读取选区 / 查看选区</sub> |
| `wps_excel_get_sheet_info` |  | R | live | 查工作表的实际使用范围与规模：用到哪一块、多少行多少列、有没有隐藏行列、有几个图表。<br><sub>别名：表详情 / 表多大 / 用到哪</sub> |
| `wps_excel_group_rows` |  | — | live | 对Excel中指定范围的行进行分组，便于折叠/展开管理。<br><sub>别名：组合行 / 分组行 / 行</sub> |
| `wps_excel_hide_column` |  | — | live | 隐藏或显示指定列。可指定起始列号、列数和隐藏/显示状态。<br><sub>别名：隐藏列 / 列</sub> |
| `wps_excel_hide_row` |  | — | live | [DEPRECATED·改用 wps_excel_hide_rows] 隐藏或显示Excel指定行。可一次操作连续多行。<br><sub>别名：隐藏行 / 行</sub> |
| `wps_excel_hide_rows` |  | — | live | 隐藏Excel中指定范围的行。<br><sub>别名：隐藏行 / 行</sub> |
| `wps_excel_insert_columns` |  | N | live | 在Excel中指定位置插入一列或多列。<br><sub>别名：插入列 / 新增列 / 加入列</sub> |
| `wps_excel_insert_excel_image` |  | N | live | 在Excel中插入图片到指定位置。<br><sub>别名：插入图片 / 新增图片 / 加入图片</sub> |
| `wps_excel_insert_rows` |  | N | live | 在Excel中指定位置插入一行或多行。<br><sub>别名：插入多行 / 批量插入行 / 插入行</sub> |
| `wps_excel_lock_cells` |  | — | live | 锁定或解锁Excel指定范围的单元格（需配合工作表保护使用）。<br><sub>别名：锁定单元格 / 单元格</sub> |
| `wps_excel_merge_cells` |  | N | live | 合并Excel指定范围的单元格。合并后保留左上角单元格的值。<br><sub>别名：合并单元格 / 单元格</sub> |
| `wps_excel_move_sheet` |  | N | live | 移动指定工作表到新的位置。<br><sub>别名：移动工作表 / 工作表</sub> |
| `wps_excel_paste_range` |  | N | live | 粘贴已复制的内容到指定位置。<br><sub>别名：粘贴区域 / 区域 / 范围</sub> |
| `wps_excel_protect_sheet` |  | — | live | 保护或取消保护工作表。<br><sub>别名：保护工作表 / 工作表</sub> |
| `wps_excel_protect_workbook` |  | — | live | 保护或取消保护工作簿，防止结构被修改（如添加/删除工作表）。<br><sub>别名：保护工作簿 / 工作簿 / 表格文件</sub> |
| `wps_excel_refresh_pivot_tables` |  | — | live | 刷新透视表（源数据改了之后重新统计）。<br><sub>别名：刷新透视表 / 透视表更新 / 数据没变</sub> |
| `wps_excel_remove_duplicates` |  | DN | live | 删除指定范围内的重复行。可以指定根据哪些列判断重复。<br><sub>别名：删除duplicates / 移除duplicates / 去掉duplicates</sub> |
| `wps_excel_rename_sheet` |  | — | live | 重命名当前工作簿中的指定工作表。<br><sub>别名：重命名工作表 / 改名工作表 / 工作表</sub> |
| `wps_excel_set_array_formula` |  | — | live | 为Excel指定范围设置数组公式（CSE数组公式）。<br><sub>别名：设置数组公式 / 设定数组公式 / 数组</sub> |
| `wps_excel_set_border` |  | — | live | 设置Excel单元格边框样式，支持不同粗细、位置和颜色。<br><sub>别名：设置边框 / 设定边框 / 边框</sub> |
| `wps_excel_set_cell_format` |  | — | live | 设置Excel单元格格式，包括字体、颜色、背景色、粗体、斜体、字号等。<br><sub>别名：设置单元格格式 / 设定单元格格式 / 单元格</sub> |
| `wps_excel_set_cell_style` |  | — | live | 应用预定义样式到Excel单元格，如标题、强调、输入、输出等内置样式。<br><sub>别名：设置单元格样式 / 设定单元格样式 / 单元格</sub> |
| `wps_excel_set_column_width` |  | — | live | 设置Excel指定列的列宽。支持单列或连续多列。<br><sub>别名：设置列宽度 / 设定列宽度 / 列</sub> |
| `wps_excel_set_conditional_format` |  | N | live | 设置条件格式。<br><sub>别名：设置条件格式 / 设定条件格式 / 条件</sub> |
| `wps_excel_set_data_validation` |  | N | live | 设置Excel单元格的数据验证规则，如下拉列表、数值范围、日期范围等。<br><sub>别名：设置数据数据验证 / 设定数据数据验证 / 数据</sub> |
| `wps_excel_set_hyperlink` |  | — | live | 为Excel单元格设置超链接。<br><sub>别名：设置超链接 / 设定超链接 / 超链接</sub> |
| `wps_excel_set_named_range` |  | N | live | 设置命名范围。为指定单元格区域创建或更新命名范围。<br><sub>别名：设置命名区域 / 设定命名区域 / 命名</sub> |
| `wps_excel_set_number_format` |  | — | live | 设置Excel单元格的数字格式。<br><sub>别名：设置数字格式 / 设定数字格式 / 数字</sub> |
| `wps_excel_set_print_area` |  | — | live | 设置打印区域<br><sub>别名：设置打印区域 / 设定打印区域 / 打印</sub> |
| `wps_excel_set_row_height` |  | — | live | 设置Excel指定行的行高。<br><sub>别名：设置行高度 / 设定行高度 / 行</sub> |
| `wps_excel_set_sheet_page_setup` |  | — | live | 设置打印页面：方向、缩放比例、纸张大小、按宽高缩放到几页。<br><sub>别名：页面设置 / 打印设置 / 横向打印</sub> |
| `wps_excel_set_sheet_print_titles` |  | — | live | 设置每页重复打印的行/列（等同"打印标题行"），或清除。<br><sub>别名：打印标题 / 每页重复表头 / 重复表头行</sub> |
| `wps_excel_set_wrap_text` |  | — | live | 设置/取消单元格自动换行。表格里最常碰的排版动作。<br><sub>别名：自动换行 / 换行显示 / 文字折行</sub> |
| `wps_excel_set_zoom` |  | — | live | 设置【当前窗口】的显示缩放（10-400%），只影响你看到的比例，不改变打印。<br><sub>别名：设置缩放 / 显示比例 / 设定缩放</sub> |
| `wps_excel_show_columns` |  | — | live | 显示Excel中已隐藏的列。<br><sub>别名：显示列 / 展示列 / 列</sub> |
| `wps_excel_show_rows` |  | — | live | 显示Excel中已隐藏的行。<br><sub>别名：显示行 / 展示行 / 行</sub> |
| `wps_excel_sort_range` |  | — | live | 对Excel选定区域按指定列排序。<br><sub>别名：排序区域 / 区域 / 范围</sub> |
| `wps_excel_subtotal` |  | — | live | 创建分类汇总。<br><sub>别名：分类汇总</sub> |
| `wps_excel_switch_sheet` |  | — | live | 切换到指定的工作表，使其成为活动工作表。<br><sub>别名：切换工作表 / 工作表</sub> |
| `wps_excel_text_to_columns` |  | — | live | 将文本按分隔符拆分到多列。<br><sub>别名：文本to列 / 文字to列 / 列</sub> |
| `wps_excel_transpose` |  | — | live | 转置数据（行列互换）。<br><sub>别名：转置</sub> |
| `wps_excel_unmerge_cells` |  | — | live | 拆分Excel中已合并的单元格，恢复为独立的单元格。<br><sub>别名：取消合并单元格 / 单元格</sub> |
| `wps_excel_unprotect_sheet` |  | — | live | 取消保护当前工作表。<br><sub>别名：取消保护工作表 / 工作表</sub> |
| `wps_excel_update_chart` |  | — | live | 更新Excel图表的属性，包括标题、颜色、图例、数据标签等。<br><sub>别名：更新图表 / 刷新图表 / 图表</sub> |
| `wps_excel_update_pivot_table` |  | — | live | 更新已有透视表的配置，包括添加/移除字段、修改聚合方式等。<br><sub>别名：更新数据透视表表格 / 刷新数据透视表表格 / 数据透视表</sub> |
| `wps_get_active_workbook` |  | R | live | 获取当前打开的WPS表格工作簿信息<br><sub>别名：当前工作簿 / 活动工作簿 / 打开的是哪个表格</sub> |

## PPT（101）

| 工具 | 默认档 | 注解 | 验证 | 说明 |
|---|---|---|---|---|
| `wps_ppt_add_slide` | 是 | N | live | 添加新幻灯片到演示文稿。<br><sub>别名：添加幻灯片 / 新增幻灯片 / 加幻灯片</sub> |
| `wps_ppt_add_speaker_notes` | 是 | — | live | 添加或追加演讲者备注到指定幻灯片。<br><sub>别名：添加演讲者备注 / 新增演讲者备注 / 加演讲者备注</sub> |
| `wps_ppt_create_presentation` | 是 | N | live | 新建空白演示文稿。<br><sub>别名：创建演示文稿 / 新建演示文稿 / 生成演示文稿</sub> |
| `wps_ppt_export_slide_as_image` | 是 | — | live | 将指定幻灯片导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。<br><sub>别名：导出幻灯片as图片 / 输出幻灯片as图片 / 幻灯片</sub> |
| `wps_ppt_get_open_presentations` | 是 | R | live | 获取当前所有已打开的演示文稿列表。<br><sub>别名：获取打开演示文稿 / 读取打开演示文稿 / 查看打开演示文稿</sub> |
| `wps_ppt_get_slide_count` | 是 | R | live | 获取演示文稿中的幻灯片总数。<br><sub>别名：获取幻灯片count / 读取幻灯片count / 查看幻灯片count</sub> |
| `wps_ppt_insert_image` | 是 | N | live | 在幻灯片中插入图片。<br><sub>别名：插入图片 / 加图片 / 新增图片</sub> |
| `wps_ppt_set_slide_title` | 是 | — | live | 设置幻灯片的标题文本。<br><sub>别名：设置标题 / 改标题 / 设置幻灯片标题</sub> |
| `wps_get_active_presentation` |  | R | static | 获取当前打开的WPS演示文稿信息<br><sub>别名：当前演示文稿 / 活动演示 / 打开的是哪个PPT</sub> |
| `wps_ppt_add_master_element` |  | N | static | 向母版中添加新元素。<br><sub>别名：添加母版element / 新增母版element / 加母版element</sub> |
| `wps_ppt_add_page_indicator` |  | N | static | 为演示文稿添加页码指示器。<br><sub>别名：添加页面指示器 / 新增页面指示器 / 加页面指示器</sub> |
| `wps_ppt_add_ppt_hyperlink` |  | N | static | 为幻灯片中的形状添加超链接。<br><sub>别名：添加超链接 / 新增超链接 / 加超链接</sub> |
| `wps_ppt_add_shape` |  | N | live | 在幻灯片中添加形状。<br><sub>别名：添加形状 / 新增形状 / 加形状</sub> |
| `wps_ppt_add_textbox` |  | N | static | 在幻灯片中添加文本框。<br><sub>别名：加文本框 / 插入文本框 / 添加文本框</sub> |
| `wps_ppt_add_title_decoration` |  | N | static | 为幻灯片标题添加装饰性元素。<br><sub>别名：添加标题装饰 / 新增标题装饰 / 加标题装饰</sub> |
| `wps_ppt_align_objects` |  | — | live | [DEPRECATED·改用 wps_ppt_align_shapes] 对齐幻灯片中的对象。<br><sub>别名：对齐objects</sub> |
| `wps_ppt_align_shapes` |  | — | static | 对齐幻灯片中的多个形状。<br><sub>别名：对齐形状 / 形状</sub> |
| `wps_ppt_apply_color_scheme` |  | — | static | 为幻灯片应用统一配色方案。<br><sub>别名：应用颜色scheme / 套用颜色scheme / 颜色</sub> |
| `wps_ppt_auto_beautify_slide` |  | — | static | 自动美化指定的单页幻灯片。<br><sub>别名：自动美化幻灯片 / 美化 / 幻灯片</sub> |
| `wps_ppt_auto_layout` |  | — | static | 自动调整幻灯片中所有元素的布局。<br><sub>别名：自动版式 / 版式 / 布局</sub> |
| `wps_ppt_beautify` |  | — | static | 一键美化幻灯片，优化排版、配色、字体和间距。<br><sub>别名：美化 / 美化单页</sub> |
| `wps_ppt_beautify_all_slides` |  | — | static | 批量美化演示文稿中的所有幻灯片。<br><sub>别名：美化全部 / 整篇美化 / 美化all幻灯片</sub> |
| `wps_ppt_close_presentation` |  | DN | static | 关闭演示文稿。可指定文稿名称，不指定则关闭当前活动文稿。<br><sub>别名：关闭演示文稿 / 演示文稿 / PPT</sub> |
| `wps_ppt_copy_slide` |  | N | static | 复制幻灯片到指定位置。<br><sub>别名：复制幻灯片 / 拷贝幻灯片 / 幻灯片</sub> |
| `wps_ppt_create_3d_text` |  | N | static | 在幻灯片中创建带有3D效果的文字。<br><sub>别名：创建3d文本 / 新建3d文本 / 生成3d文本</sub> |
| `wps_ppt_create_donut_chart` |  | N | static | 在幻灯片中创建环形图。<br><sub>别名：创建环形图图表 / 新建环形图图表 / 生成环形图图表</sub> |
| `wps_ppt_create_flow_chart` |  | N | static | 在幻灯片中创建流程图。<br><sub>别名：创建流程图表 / 新建流程图表 / 生成流程图表</sub> |
| `wps_ppt_create_gauge` |  | N | static | 在幻灯片中创建仪表盘图表。<br><sub>别名：创建仪表盘 / 新建仪表盘 / 生成仪表盘</sub> |
| `wps_ppt_create_grid` |  | N | static | 在幻灯片中创建网格布局。<br><sub>别名：创建网格 / 新建网格 / 生成网格</sub> |
| `wps_ppt_create_kpi_cards` |  | N | static | 在幻灯片上创建KPI指标卡片。<br><sub>别名：创建KPI卡片 / 新建KPI卡片 / 生成KPI卡片</sub> |
| `wps_ppt_create_mini_charts` |  | N | static | 在幻灯片中创建迷你图表。<br><sub>别名：创建迷你图表 / 新建迷你图表 / 生成迷你图表</sub> |
| `wps_ppt_create_org_chart` |  | N | static | 在幻灯片中创建组织架构图。<br><sub>别名：创建组织图表 / 新建组织图表 / 生成组织图表</sub> |
| `wps_ppt_create_progress_bar` |  | N | static | 在幻灯片中创建进度条。<br><sub>别名：创建进度条 / 新建进度条 / 生成进度条</sub> |
| `wps_ppt_create_styled_table` |  | N | static | 在幻灯片中创建带预设样式的精美表格。<br><sub>别名：创建styled表格 / 新建styled表格 / 生成styled表格</sub> |
| `wps_ppt_create_timeline` |  | N | static | 在幻灯片中创建时间线。<br><sub>别名：创建时间轴 / 新建时间轴 / 生成时间轴</sub> |
| `wps_ppt_delete_ppt_image` |  | DN | static | 删除幻灯片中指定的图片。<br><sub>别名：删除图片 / 删掉图片 / 去掉图片</sub> |
| `wps_ppt_delete_shape` |  | DN | static | 删除幻灯片中指定的形状。<br><sub>别名：删除形状 / 删掉形状 / 去掉形状</sub> |
| `wps_ppt_delete_slide` |  | DN | static | 删除指定的幻灯片。<br><sub>别名：删除幻灯片 / 删掉幻灯片 / 去掉幻灯片</sub> |
| `wps_ppt_delete_textbox` |  | DN | static | 删除幻灯片上指定的文本框。<br><sub>别名：删除文本框 / 删掉文本框 / 去掉文本框</sub> |
| `wps_ppt_distribute_shapes` |  | — | static | 等距分布幻灯片中的多个形状。<br><sub>别名：分布形状 / 等距分布形状 / 形状</sub> |
| `wps_ppt_duplicate_shape` |  | N | static | 复制幻灯片中的指定形状。<br><sub>别名：复制形状 / 副本形状 / 形状</sub> |
| `wps_ppt_find_ppt_text` |  | R | static | 在演示文稿中搜索指定文本。<br><sub>别名：查找文本 / 搜索文本 / 文本</sub> |
| `wps_ppt_get_animations` |  | R | static | 获取幻灯片上所有动画效果的列表。<br><sub>别名：获取动画 / 读取动画 / 查看动画</sub> |
| `wps_ppt_get_shapes` |  | R | static | 获取幻灯片中所有形状的列表信息。<br><sub>别名：获取形状 / 读取形状 / 查看形状</sub> |
| `wps_ppt_get_slide_info` |  | R | static | 获取指定幻灯片的详细信息，包括布局、元素列表等。<br><sub>别名：获取幻灯片信息 / 读取幻灯片信息 / 查看幻灯片信息</sub> |
| `wps_ppt_get_slide_master` |  | R | static | 获取当前演示文稿的母版信息。<br><sub>别名：获取幻灯片母版 / 读取幻灯片母版 / 查看幻灯片母版</sub> |
| `wps_ppt_get_slide_notes` |  | R | static | 获取指定幻灯片的备注内容。<br><sub>别名：获取幻灯片备注 / 读取幻灯片备注 / 查看幻灯片备注</sub> |
| `wps_ppt_get_slide_title` |  | R | static | 获取指定幻灯片的标题。<br><sub>别名：获取幻灯片标题 / 读取幻灯片标题 / 查看幻灯片标题</sub> |
| `wps_ppt_get_table_cell` |  | R | static | 获取PPT表格中指定单元格的文本内容。<br><sub>别名：获取表格单元格 / 读取表格单元格 / 查看表格单元格</sub> |
| `wps_ppt_get_textboxes` |  | R | static | 获取幻灯片上所有文本框的列表。<br><sub>别名：获取文本框 / 读取文本框 / 查看文本框</sub> |
| `wps_ppt_group_shapes` |  | — | static | 将幻灯片中的多个形状组合为一个组。<br><sub>别名：组合形状 / 分组形状 / 形状</sub> |
| `wps_ppt_insert_ppt_chart` |  | N | live | 在幻灯片中插入数据图表。<br><sub>别名：插入图表 / 新增图表 / 加入图表</sub> |
| `wps_ppt_insert_ppt_image` |  | N | static | [DEPRECATED·改用 wps_ppt_insert_image] 插入图片到幻灯片中。<br><sub>别名：插入图片 / 新增图片 / 加入图片</sub> |
| `wps_ppt_insert_slide_image` |  | N | static | 在幻灯片中插入图片。<br><sub>别名：插入幻灯片图片 / 新增幻灯片图片 / 加入幻灯片图片</sub> |
| `wps_ppt_insert_slides_from_file` |  | N | static | 从另一个 PPT 文件把整页幻灯片插入到【当前活动演示文稿】，并保留来源幻灯片的原始格式（字体/配色/版式/图片）。用于把多个…<br><sub>别名：插入幻灯片fromfile / 新增幻灯片fromfile / 加入幻灯片fromfile</sub> |
| `wps_ppt_insert_table` |  | N | live | 在幻灯片中插入表格。<br><sub>别名：插入表格 / 加表格 / 新增表格</sub> |
| `wps_ppt_move_slide` |  | N | static | 移动幻灯片到指定位置。<br><sub>别名：移动幻灯片 / 幻灯片 / 页面</sub> |
| `wps_ppt_open_presentation` |  | — | static | 打开指定路径的演示文稿文件。<br><sub>别名：打开演示文稿 / 演示文稿 / PPT</sub> |
| `wps_ppt_remove_animation` |  | DN | static | 移除幻灯片中指定的动画效果。<br><sub>别名：删除动画 / 移除动画 / 去掉动画</sub> |
| `wps_ppt_remove_ppt_hyperlink` |  | DN | static | 移除幻灯片中形状的超链接。<br><sub>别名：删除超链接 / 移除超链接 / 去掉超链接</sub> |
| `wps_ppt_replace_ppt_image` |  | DN | static | 原位替换幻灯片中的某张图片：保留原图的位置、尺寸、旋转角度，删除旧图后在同一矩形内插入新图。用于在保持版式不变的前提下把模板里…<br><sub>别名：替换图片 / 换图 / 图片</sub> |
| `wps_ppt_replace_ppt_text` |  | DN | static | 在演示文稿中查找并替换文本。<br><sub>别名：替换文本 / 文本 / 文字</sub> |
| `wps_ppt_set_3d_depth` |  | — | static | 设置幻灯片中形状的3D挤出深度。<br><sub>别名：设置3d深度 / 设定3d深度 / 深度</sub> |
| `wps_ppt_set_3d_material` |  | — | static | 设置幻灯片中形状的3D材质效果。<br><sub>别名：设置3d材质 / 设定3d材质 / 材质</sub> |
| `wps_ppt_set_3d_rotation` |  | — | static | 设置幻灯片中形状的3D旋转效果。<br><sub>别名：设置3d旋转 / 设定3d旋转 / 旋转</sub> |
| `wps_ppt_set_active_target` |  | — | static | 锁定后续所有 PPT 操作的【目标演示文稿】（按文件名）。锁定后，本服务器会给所有演示类调用自动注入该文稿名，底层精确定位，彻…<br><sub>别名：设置当前目标 / 设定当前目标 / 当前</sub> |
| `wps_ppt_set_animation` |  | N | static | 给形状加动画，或调整已有动画的播放顺序（第 8 轮合并了动画族 5 个工具）。<br><sub>别名：设置动画 / 添加动画 / 加动画</sub> |
| `wps_ppt_set_background` |  | — | static | 设置背景：某一页、整份文稿、或幻灯片母版。<br><sub>别名：设置背景 / 换背景 / 背景色</sub> |
| `wps_ppt_set_font_color` |  | — | static | 设置幻灯片中指定形状的文字颜色。<br><sub>别名：设置字体颜色 / 设定字体颜色 / 字体</sub> |
| `wps_ppt_set_image_style` |  | — | static | 设置幻灯片中指定图片的样式。<br><sub>别名：设置图片样式 / 设定图片样式 / 图片</sub> |
| `wps_ppt_set_ppt_chart_data` |  | — | live | 更新幻灯片中已有图表的数据。<br><sub>别名：改图表数据 / 设置图表数据 / 设定图表数据</sub> |
| `wps_ppt_set_ppt_chart_style` |  | — | live | 设置幻灯片中图表的样式属性。<br><sub>别名：设置图表样式 / 设定图表样式 / 图表</sub> |
| `wps_ppt_set_ppt_date_time` |  | — | static | 设置演示文稿日期时间显示。<br><sub>别名：设置日期时间 / 设定日期时间 / 日期</sub> |
| `wps_ppt_set_ppt_footer` |  | — | static | 设置演示文稿页脚文本。<br><sub>别名：设置页脚 / 设定页脚 / 页脚</sub> |
| `wps_ppt_set_shape_border` |  | — | static | 设置幻灯片中指定形状的边框样式。<br><sub>别名：设置形状边框 / 设定形状边框 / 形状</sub> |
| `wps_ppt_set_shape_fill` |  | — | live | 设置幻灯片中指定形状的填充颜色。<br><sub>别名：设置形状填充 / 设定形状填充 / 形状</sub> |
| `wps_ppt_set_shape_gradient` |  | — | static | 设置幻灯片中指定形状的渐变填充效果。<br><sub>别名：设置形状渐变 / 设定形状渐变 / 形状</sub> |
| `wps_ppt_set_shape_position` |  | — | static | 设置幻灯片中指定形状的位置和大小。<br><sub>别名：设置形状位置 / 设定形状位置 / 形状</sub> |
| `wps_ppt_set_shape_shadow` |  | — | static | 设置幻灯片中指定形状的阴影效果。<br><sub>别名：设置形状阴影 / 设定形状阴影 / 形状</sub> |
| `wps_ppt_set_shape_style` |  | — | static | 设置幻灯片中形状的样式，包括填充颜色、边框颜色和边框粗细。<br><sub>别名：设置形状样式 / 设定形状样式 / 形状</sub> |
| `wps_ppt_set_shape_text` |  | — | live | 设置幻灯片中指定形状的文字内容。<br><sub>别名：设置形状文本 / 设定形状文本 / 形状</sub> |
| `wps_ppt_set_shape_transparency` |  | — | static | 设置幻灯片中指定形状的透明度。<br><sub>别名：设置形状透明度 / 设定形状透明度 / 形状</sub> |
| `wps_ppt_set_shape_z_order` |  | — | static | 设置形状在幻灯片中的层级顺序（Z轴排列）。<br><sub>别名：设置形状z顺序 / 设定形状z顺序 / 形状</sub> |
| `wps_ppt_set_slide_content` |  | — | static | 设置幻灯片的正文内容区域。<br><sub>别名：设置幻灯片内容 / 设定幻灯片内容 / 幻灯片</sub> |
| `wps_ppt_set_slide_layout` |  | — | static | 设置幻灯片的版式布局。<br><sub>别名：设置幻灯片版式 / 设定幻灯片版式 / 幻灯片</sub> |
| `wps_ppt_set_slide_number` |  | — | static | 设置幻灯片页码的显示状态和起始编号。<br><sub>别名：设置幻灯片数字 / 设定幻灯片数字 / 幻灯片</sub> |
| `wps_ppt_set_slide_size` |  | — | static | 设置演示文稿的幻灯片尺寸（按像素给，内部换算成 WPS 的磅）。<br><sub>别名：设置幻灯片size / 设定幻灯片size / 幻灯片</sub> |
| `wps_ppt_set_slide_subtitle` |  | — | static | 设置幻灯片的副标题。<br><sub>别名：设置幻灯片副标题 / 设定幻灯片副标题 / 幻灯片</sub> |
| `wps_ppt_set_slide_theme` |  | — | static | 应用主题模板（换整套配色/字体/版式）。<br><sub>别名：设置幻灯片主题 / 设定幻灯片主题 / 幻灯片</sub> |
| `wps_ppt_set_table_cell` |  | — | live | 设置PPT表格中指定单元格的文本内容。<br><sub>别名：设置表格单元格 / 设定表格单元格 / 表格</sub> |
| `wps_ppt_set_table_cell_style` |  | — | static | 设置PPT表格中指定单元格的样式。<br><sub>别名：设置表格单元格样式 / 设定表格单元格样式 / 表格</sub> |
| `wps_ppt_set_table_row_style` |  | — | live | 设置PPT表格中指定行的样式。<br><sub>别名：设置表格行样式 / 设定表格行样式 / 表格</sub> |
| `wps_ppt_set_table_style` |  | — | static | 设置PPT表格的整体样式。<br><sub>别名：设置表格样式 / 设定表格样式 / 表格</sub> |
| `wps_ppt_set_textbox_style` |  | — | static | 设置文本框样式（字体大小、颜色、粗体等）。<br><sub>别名：设置文本框样式 / 设定文本框样式 / 文本框</sub> |
| `wps_ppt_set_textbox_text` |  | — | static | 设置指定文本框的文本内容。<br><sub>别名：设置文本框文本 / 设定文本框文本 / 文本框</sub> |
| `wps_ppt_set_transition` |  | — | static | 设置幻灯片切换效果：单页 or 整份文稿（第 8 轮合并了切换族 3 个工具）。<br><sub>别名：切换效果 / 切换动画 / 翻页效果</sub> |
| `wps_ppt_smart_distribute` |  | — | static | 将指定形状进行等距分布排列。<br><sub>别名：智能分布 / 分布 / 等距分布</sub> |
| `wps_ppt_start_slide_show` |  | N | static | 开始幻灯片放映。<br><sub>别名：开始幻灯片显示 / 启动幻灯片显示 / 幻灯片</sub> |
| `wps_ppt_switch_presentation` |  | — | static | 切换到指定名称的演示文稿。<br><sub>别名：切换演示文稿 / 演示文稿 / PPT</sub> |
| `wps_ppt_switch_slide` |  | — | static | 切换到指定的幻灯片页面。<br><sub>别名：切换幻灯片 / 幻灯片 / 页面</sub> |
| `wps_ppt_unify_font` |  | — | static | 统一演示文稿中所有幻灯片的字体。<br><sub>别名：统一字体 / 字体</sub> |

---

想按中文关键词找工具，别翻这份表：直接用 `wps_search_tools`（支持"插入表格""批注""改背景"这类说法，
返回里带 `notFor` 反例，能分辨近名工具）。这份表是给人看的，也是给搜索引擎看的。
