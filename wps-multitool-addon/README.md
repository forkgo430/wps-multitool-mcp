manifest.xml, ribbon.xml, js/, panel.html
Windows WPS 加载项配置、脚本入口与任务窗格页面目录。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

## 结构

| 文件 | 作用 |
|---|---|
| `manifest.xml` | 加载项清单（版本/作者/ hosts / ribbon 入口） |
| `ribbon.xml` | 功能区定义 —— 「多面手 MCP」9 个按钮，分 4 组（v2.4） |
| `js/main.js` | ribbon 回调：快照/撤销/导出/面板/关于等 |
| `panel.html` | 任务窗格页面（桥状态 / 操作目标 / 快捷动作 / 最近改动） |

## 部署（重要，踩过坑）

安装脚本用的是 **Copy-Item 而不是符号链接**（符号链接要管理员权限，失败时回退复制），
所以 `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_` 是一份**副本**。
只改本目录不重新部署，WPS 重启多少次都还是旧内容。

```powershell
powershell -ExecutionPolicy Bypass -File ..\..\scripts\deploy-addon.ps1        # 部署
powershell -ExecutionPolicy Bypass -File ..\..\scripts\deploy-addon.ps1 -Check # 只查差异
```

部署后必须**完全退出 WPS 再重开**（任务栏右键退出，不是关窗口）。

## 面板的数据来源

任务窗格跑在内嵌浏览器里，连不到命名管道（内核对象，仅限本进程）。
所以桥额外开了一个只读的 localhost HTTP 端点：`http://127.0.0.1:27731/status`，
`panel.html` 每 2 秒轮询。端口可用 `WPS_BRIDGE_HTTP_PORT` 改。

「最近改动」由桥在管道上记录（它看得见每一个请求），只记摘要、不记文档正文。

## 已知限制

- WPS JS 加载项**没有**应用内通知/toast，提示只能走 `alert` 或任务窗格内文字。
- 任务窗格的 `CreateTaskPane` API 在本机 WPS 12.1 上按官方社区示例实现，**未经真机验证**。
- 「工具组 基础 40」在面板里是写死的说明文字（桥拿不到 MCP server 的注册表）。
