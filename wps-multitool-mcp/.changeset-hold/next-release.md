---
"wps-multitool-mcp": minor
---

写入可恢复、验证状态、一键接入、自动发版

- 新增 `wps_common_snapshot_create` / `snapshot_list` / `snapshot_restore`：写入前存一份磁盘快照，改坏了能退回去。`wps_common_batch` 检测到破坏性步骤时自动快照，并把快照路径写进结果。
- 每个工具标注真机验证状态（`live` / `static`）。默认档改为"基础档 ∩ 真机验证"，PPT 整组（101 个）在真机验证前不占默认名额；`wps_search_tools` 与 `wps_list_tool_groups` 会显示状态。
- 新增 `wps-multitool-mcp install` 子命令：自动识别 CodeBuddy / WorkBuddy / Codex / Claude Code / Cursor，写前备份、写后回读校验、失败自动还原。支持 `--dry-run` / `--print`。
- 新增生成的 [TOOLS.md](TOOLS.md)：284 个工具的可浏览清单，由注册表产出，CI 做 diff 门禁。
- 新增 [SECURITY.md](SECURITY.md)：威胁模型、隐私边界、高风险工具清单；桥的只读状态端点 CORS 由 `*` 收紧为仅 `file://`。
- 引入 changesets + GitHub Actions 自动发版（npm trusted publishing，无需每次输 OTP）。
- PPT 基础档 8 个工具完成真机验证（新建 / 加页 / 标题 / 备注 / 插图 / 导出图片 / 列文稿），默认档恢复到 40 个。
- 修复 `wps_ppt_create_presentation` 永远报"失败: undefined"（PS 侧只回 success 不回 data，handler 却要求两者都有）。
- 清理：移除未使用的 `axios` 依赖，以及指向不存在目录的 jest 配置与相关 devDependencies。
