index.ts, presentation.ts, slide-ops.ts, slide.ts, table.ts, textbox.ts
PPT 工具定义与处理器目录，含幻灯片/幻灯片操作/演示文稿管理/文本框/表格模块。
一旦这里的结构发生变化，请务必更新我。

B5-T3（E2）：新增 deckkit.ts——deckkit 引擎骨架（build_ppt_deck：JSON 页面描述 → 整套演示文稿，10 种页型，排版纪律常量 DECK_TYPO 内置；不进默认档）。

E3：deckkit.ts 扩——①三套主题 DECK_THEMES（学术蓝/政务红/商务灰，自创数值）+ build_ppt_deck 的 theme 参数（缺省学术蓝）；②生成后溢出检查（每页铺完经新只读工具 wps_ppt_get_textbox_metrics 读回 BoundHeight vs 帧高，超框写 warnings 并缩字一档复测一次，仍超只警告；量不了如实降级不算失败）；③textbox.ts 三个按索引定位的工具 wire 错位顺修（桥读 shapeIndex，工具此前发 textboxIndex，恒败）。
