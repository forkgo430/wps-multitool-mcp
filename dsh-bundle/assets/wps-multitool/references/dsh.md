# 在 DeepSeek Harness（DSH）里用本 MCP

> 什么时候读：**你正在 DSH 里跑**，或者用户说"用 DSH / Harness 调 WPS"。
> 在别的客户端（Claude Code / Cursor / CodeBuddy / WorkBuddy / Codex）里**不用读这篇**。

---

## 1. 怎么挂上去

**挂官方插件，不要挂我们的 bundle。**

bundle（`dsh-bundle/`）那份代码在真宿主里必然失败：它要 `import('@deepseek-ai/dsh-mcp-client')`，
而那个包只在 DSH 自己的 `app.asar` 内部、只对 DSH 的代码可见。
**从 bundle 目录永远解析不到它** —— 这不是路径没配对，是解析边界。

正确做法是在 patch 层直接挂官方插件：

```yaml
- insert:
    - id: mcp-wps
      name: '@deepseek-ai/dsh-mcp-client'
      config:
        serverName: wps
        transport: stdio
        command: node
        args:
          - C:\绝对路径\wps-multitool-mcp\dist\index.js
        toolCallTimeoutMs: 130000
        failOnStartupError: true
```

生成这段配置：`node dist/index.js install --client dsh`（会写 `~/.dsh/wps.cordis.yml`）。

**三个字段每一个都有理由，别随手改**：

| 字段 | 为什么必须是它 |
|---|---|
| `command: node` + **绝对路径** | DSH 的 args 直接透传、不过 shell。Windows 上 `npx` 实际是 `npx.cmd`，按 `npx` 去 spawn 会 ENOENT |
| `toolCallTimeoutMs: 130000` | 必须 **≥ 桥的 `WPS_DAEMON_TIMEOUT`（默认 120000）**。DSH 默认 60000 会先掐断 —— 而 WPS 那边动作**已经改完了**，症状是"客户端报失败、文档其实改了" |
| `failOnStartupError: true` | 默认 `false` 会在连不上时**静默降级成"零工具"**，用户只会觉得"AI 不会用 WPS" |

---

## 2. DSH 特有的两个坑

### 2.1 `desktop` profile 被 Electron 独占

CLI 只能用 `--profile web`，或者自己新建一个 profile。
用 `desktop` 跑 CLI 会起不来（Electron 已经占着它）。

### 2.2 装完必须清干净

测完记得 `dsh plugin --profile web remove dsh-wps`，
并确认 `dsh --profile web --dump-config` 里已经没有 `dsh-wps` 那一层。

---

## 3. 平台限制（这几条对 DSH 用户同样成立）

| 限制 | 说明 |
|---|---|
| **Windows only** | WPS 的 COM 自动化只在 Windows 可用。DSH 跑在别的系统上就没有这条通道 |
| **默认档 40 个工具** | DSH 会把工具定义塞进**每一次请求**，全量 284 个很贵。保持默认档，别在 DSH 里改成全量 |
| **超时上限 120 秒** | 那是**单次动作的上限**，不是常态耗时。正常一次调用 2~3 毫秒 |
| **Linux 通道还没有** | 本项目的跨平台路线尚未落地，DSH 在 Linux 上同样没有 |

---

## 4. 一条要提醒的读回限制

`wps_word_get_open_documents` / `wps_excel_get_open_workbooks` 在
**刚执行过 `quit` / `close` 之后**可能短暂读到 0 份 —— 那是 WPS 重新注册到 ROT 的那几秒窗口，
**文档其实还在**。所以：

- 读到 0 份时，**别据此告诉用户"文档没了"**；
- 隔一会儿再读一次，或**按文档名直接读内容**核实（那个更硬）；
- 工具返回里带 warnings 时，**把 warnings 念给用户听**。

---

## 5. 出问题先看哪

| 症状 | 先查 |
|---|---|
| DSH 里一个 WPS 工具都没有 | `--dump-config` 里有没有 `dsh-wps` 那层；`failOnStartupError` 是不是 true |
| 报"文档其实改了但客户端说失败" | `toolCallTimeoutMs` 是不是小于 130000 |
| 工具列表为空但 CLI 手动跑正常 | 是不是在 `desktop` profile 上跑的 CLI |
| 桥进程正常但调用超时 | 单次动作 120 秒上限；看是不是发了个导出类大动作 |
