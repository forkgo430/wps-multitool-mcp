# dsh-multitool

给 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（dsh）用的**组合包（bundle）**。

一条命令，同时得到两样东西：

- **40 个 WPS 工具** —— 读写**本机正开着的** WPS 文字/表格/演示
- **一份技能** —— 教 AI 什么时候该用这些工具、怎么用才不会选错

```sh
dsh plugin --profile <你的profile> add dsh-multitool
```

---

## 它和 `dsh-wps` 不是一回事（**装之前先看这段**）

npm 上已经有一个 `dsh-wps`，是别人做的 **WPS / 金山文档「云文档」集成**（走官方 SkillHub MCP 端点）。

本包操作的是**你电脑上正开着的 WPS 窗口**，不是云文档。两者能力完全不同：

| | 本包 `dsh-multitool` | 别人的 `dsh-wps` |
|---|---|---|
| 操作对象 | **本机正开着的** WPS 文档 | WPS 云文档 |
| 需要联网 | 不需要 | 需要 |
| 需要授权 | 不需要 | 要 OAuth |
| 工具前缀 | `mcp__wpslocal__*` | `mcp__wps__*` |
| 能做什么 | 改格式、插图表、排版、批量处理 | 云端文件读写 |

**工具前缀刻意不同**（`wpslocal` vs `wps`），所以两个包可以同时装，不会撞车。

---

## 它做了什么

```
dsh-multitool/
├── package.json        # 声明 dsh.bundle.patch + engines.dsh
├── cordis.patch.yml    # 配置层：挂 MCP + 启用技能
└── assets/
    └── wps-multitool/
        ├── SKILL.md
        └── references/{routing,recipes,errors,dsh}.md
```

`cordis.patch.yml` 做三件事：

1. **挂官方 MCP 客户端**（`@deepseek-ai/dsh-mcp-client`），指向本机的 `wps-multitool-mcp`
2. **启用 `skill-filesystem`**，把本包的 `assets/` 加进技能扫描根
3. **启用 `tool-skill`**，让模型能看到并加载技能

---

## 为什么这样写（四条硬约束，改之前先读）

### 1. 插件行按【包名】引用，不写相对源码路径

dsh 的 loader 从 profile 的 `node_modules` 解析包名。我们**不自己 `import`** 官方 MCP 客户端 ——
那个包只在 dsh 的 `app.asar` 内部可见，从 profile 目录**永远**解析不到。

> 这条是踩过坑的：v0.3 让 `index.js` 里 `import('@deepseek-ai/dsh-mcp-client')`，
> 真机报 `Cannot find package`。改成"由 loader 解析"一次就通。

### 2. 路径必须绝对，用 `!!js` 在求值时算

发布时不可能知道用户把包装在哪，所以 `args` 和 `customSkillDirs` 都用 `!!js`：

```yaml
customSkillDirs:
  - !!js >-
      process.getBuiltinModule('node:path').join(
        process.getBuiltinModule('node:path').dirname(
          process.getBuiltinModule('node:module').createRequire(baseUrl).resolve('dsh-multitool/package.json')
        ),
        'assets'
      )
```

**求值上下文提供 `baseUrl` 和 `process.getBuiltinModule`。** 这是从官方 `dsh-base` 的
`skill-filesystem` 配置里学的写法（它自己就这么算 preset skills 目录）。

### 3. `skill-filesystem` / `tool-skill` 在模板里默认禁用

web 模板里这两个是 `disabled: true`。本包的 patch 按 `id` 覆盖它们、**且不写 `disabled`**，
等于重新启用。

> 官方语义：后层按 `id` **整行覆盖**，`config` 是**整体替换不是深合并** —— 所以 `name` 必须写全。

### 4. serverName 用 `wpslocal`，不是 `wps`

见上方"和 `dsh-wps` 不是一回事"。**这条别改回去** —— 改成 `wps` 会和别人的包撞名，
两个都装时 dsh 会报 `serverName "wps" is already in use`。

---

## 两个参数为什么是这两个值

| 参数 | 值 | 不这么写会怎样 |
|---|---|---|
| `toolCallTimeoutMs` | `130000` | dsh 默认 60000 < 桥的 120000 → **客户端报超时、文档其实已经改完**，用户以为失败又重做一遍 |
| `failOnStartupError` | `true` | 默认 `false` 会在连不上时**静默降级成零工具**，用户只觉得"AI 不会用 WPS"，看不到任何错误 |
| `command` | `node`（不是 `npx`） | dsh 的 args 直接透传、不过 shell，Windows 上 `npx` 实际是 `npx.cmd`，按 `npx` spawn 会 **ENOENT** |

---

## 依赖

`wps-multitool-mcp` 作为 `dependencies` 声明。入口路径按两档解析：

1. 先按依赖解析（`createRequire(baseUrl).resolve('wps-multitool-mcp/dist/index.js')`）
2. 解析不到就退回**本包同级目录**（开发者本地 checkout 的情形）

---

## 验证状态（2026-10-22）

| 项 | 结果 |
|---|---|
| `dsh plugin add ./dsh-bundle` | ✅ exit 0 |
| `--dump-config` 出现本包层 | ✅ 且显示 `patched by ... dsh-multitool`（覆盖了 skill-filesystem / tool-skill） |
| 真启动 `dsh --profile web` | ✅ **MCP 进程起来了**（`node ...\wps-multitool-mcp\dist\index.js`） |
| 技能插件是否报错 | ✅ **stderr 全空**（路径错会抛 ENOENT） |
| 卸载干净 | ✅ `dsh plugin remove` 后配置树无残留 |
| 自检 `node dsh-bundle/selftest.js` | ✅ 16 项全过，工具 40 个 |
| `npm run verify:all` | ✅ 23/23 |

**未验证**：工具名在模型侧是否真以 `mcp__wpslocal__*` 出现、技能是否真的进了会话目录 ——
这两条要在有 API key 的对话回合里才能看到。

---

## 平台限制

**Windows only。** 本包背后是 WPS Office 的 COM 自动化，没有 Windows 跑不了。

默认档 40 个工具（不是全部 284）—— dsh 会把工具定义塞进每一次请求，全量很贵。
需要更多就用 `WPS_TOOL_TIER=full`，或按组加载。

---

## 许可

MIT。本包只做配置与技能分发，实际能力在 [`wps-multitool-mcp`](https://github.com/forkgo430/wps-multitool-mcp)。
