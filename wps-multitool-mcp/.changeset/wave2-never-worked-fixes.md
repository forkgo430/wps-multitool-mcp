---
"wps-multitool-mcp": patch
---

修 6 个「从上线起从未成功过」的工具，并堵住让它们漏网的口子

这 6 个工具的动作名在 `scripts/wps-com.ps1` 的顶层 `switch ($Action)` 里**根本不存在**，
每次调用都落到默认分支回 `Unknown action`。它们都带着一句 `// NOTE: 仅Windows支持` ——
那句注释是假的：Windows 上也不支持，ps1 里压根没有这些 case。读代码的人会以为"Windows 能用"。

已修（动作名改正或补 case，逐个真机读过值）：

- `wps_word_set_text_color` → 改调 `setFont`（`setFont` 才吃 color）
- `wps_word_set_line_spacing` → 改调 `setParagraph`（`setFont` 的参数表里没有 lineSpacing）
- `wps_excel_auto_sum` → 改调 `setFormula` 并自己拼 `=SUM(区域)`；写进去的是**活公式**，源区域改了会跟着变
- `wps_excel_evaluate_formula` → 补 `evaluateFormula` case（`Application.Evaluate`，**只算不写**；
  给了 `cell` 才由 TS 另发一次写入动作，避免"声明只读却动手改"）
- `wps_excel_set_zoom` → 补 `setZoom` case（`ActiveWindow.Zoom`）。
  它与 `PageSetup.Zoom`（打印缩放）是两回事，别混用
- `wps_word_insert_section_break` → 改调 `insertPageBreak`，并给它补 `evenPage`/`oddPage`
  两个映射值；不认识的分隔符类型现在**明确报错**，不再静默退化成分页符

顺带（同一条命根子上的两个洞）：

- 注解生成器新增**全量交叉校验**：`TOOL_ACTIONS` 里每个动作名都必须能在 ps1 顶层 switch 里查到，
  查不到就非 0 退出。此前它只校验手写的 `STATIC_ACTIONS`，自动推断出来的动作名不校验 —— 这 6 个就是这么漏的
- 路由表生成器改成**只扫主 `switch ($Action)`**（括号深度判顶层）。此前它扫整个文件，
  把 10 个嵌套 switch 的 mode 标签（`trim` / `placeholder` / `acceptOne` …）当成了动作，
  让"动作数"虚高 10。运行时无害，但那个数字被写进了文档与断言

另外修了 `npm run verify` 的档位开关：它此前被硬编码成全量，与 `verify:all` 完全一样，
文件头和 `package.json` 承诺的"快测档"实际不存在。
