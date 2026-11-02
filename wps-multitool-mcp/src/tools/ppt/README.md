index.ts, presentation.ts, slide-ops.ts, slide.ts, table.ts, textbox.ts, shape-basic.ts, image.ts, page-effects.ts, page-info.ts, animation.ts, chart-flow.ts, data-viz.ts, beautify-advanced.ts, misc.ts, deckkit.ts, helpers.ts
PPT 工具定义与处理器目录，含幻灯片页操作/形状/文本框/演示文稿管理/页面信息/页面效果/表格/图表/美化/杂项模块。
一旦这里的结构发生变化，请务必更新我。

B5-T3（E2）：新增 deckkit.ts——deckkit 引擎骨架（build_ppt_deck：JSON 页面描述 → 整套演示文稿，10 种页型，排版纪律常量 DECK_TYPO 内置；不进默认档）。

E3：deckkit.ts 扩——①三套主题 DECK_THEMES（学术蓝/政务红/商务灰，自创数值）+ build_ppt_deck 的 theme 参数（缺省学术蓝）；②生成后溢出检查（每页铺完经新只读工具 wps_ppt_get_textbox_metrics 读回 BoundHeight vs 帧高，超框写 warnings 并缩字一档复测一次，仍超只警告；量不了如实降级不算失败）；③textbox.ts 三个按索引定位的工具 wire 错位顺修（桥读 shapeIndex，工具此前发 textboxIndex，恒败）。

质量修复轮（2026-10-30）：审计三.1-三.7 落地——0 坐标 ?? 修复、deckkit 发 type 键、get_slide_info 补占位符字段、speaker_notes 描述对齐行为、chart data 可选、定位参数前置校验、set_active_target 补 appType、头注清漂。

风格统一+拆分轮（2026-11-03）：①新建 helpers.ts（call/ok/fail/errText/errMsg，warnings 一路透传，范式照搬 word/helpers）；②slide-ops.ts 2087 行拆三群——页操作 8 个留原地、形状与文本 6 个迁 shape-basic.ts（该文件 10→16 工具）、动画/背景/切换/备注 4 个迁新 page-effects.ts、insert_image 迁 image.ts（5→6）；③background.ts 改名 page-info.ts（内容本就是页码/页脚/日期时间/复制/层级，文件名漂移决策者预批）；④全组收尾风格统一到助手、uuid 死重量清除；工具名/动作名/参数键/成功文案逐字保留（纯行为保持重构），10644→约 7700 行。
