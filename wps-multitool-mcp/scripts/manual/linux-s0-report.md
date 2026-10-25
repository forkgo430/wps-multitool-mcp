# Linux S0 能力探针 · 覆盖率报告（模板）

> 依据 `docs/agents/PLAN-LINUX-BACKEND.md` §5.2 / §5.3。
> 由 `scripts/manual/linux-s0-probe.js` 生成；本文件是**模板**，20 行判据先写好，等真 Linux 机器上跑完自动填。

## 本轮状态

**每一行都是 `未验证（本机无 Linux 环境）`。**

这是仓库自己的规矩：`live` 必须有真机证据。本轮交付的是**脚本**，不是覆盖率数字。
仓库里没有 Linux 机器，脚本一行都没在 WPS Linux 上跑过 —— 所以判定一律按 `static` 计。
**任何人都不许在没有真机的情况下往这张表里填数。**

要真跑：

```bash
# 装了 WPS Linux + Xvfb 的机器上
node scripts/manual/linux-s0-probe.js --wps /usr/bin/wps --display :99 --out ./linux-s0-report.md
```

脚本会：R0 先查 `/opt/kingsoft/wps-office/office6/jsaddons` → 生成 `probe.js` →
`xvfb-run -a wps --macro=probe.js` → 逐条记录「支持 / 不支持（属性不存在）/ 抛错 / 部分支持」→
把本文件覆盖成填好的版本。

## R0 前置检查

| 项 | 结果 | 说明 |
|---|---|---|
| `/opt/kingsoft/wps-office/office6/jsaddons` | **未验证（本机无 Linux 环境）** | §5.2 要求 S0 第一步就查。查不到就直接停，回 §10 的 LibreOffice/UNO 路线重新评估 —— 不要接着写 B2 |
| 官方有没有 Linux 版加载项启用说明 | **未验证（本机无 Linux 环境）** | 同上 |
| `--macro` 能否 attach 到已开着的 WPS 实例 | **未验证（本机无 Linux 环境）** | §5.2 的第二个问题。能 → 核心卖点保住；不能 → 它就只是个探针，覆盖率数字照样拿得到 |

## 覆盖表（§5.3 的 20 个代表性动作）

| # | 动作 | 考察点 | 判定 | 证据（属性是否存在 / 调用是否抛错） |
|---|---|---|---|---|
| 1 | `getContext` / `getActiveDocumentState` | 能否拿到活动文档（WPS JSAPI 有无对应） | **未验证（本机无 Linux 环境）** | — |
| 2 | `getSelection` | Selection 能力 | **未验证（本机无 Linux 环境）** | — |
| 3 | `getSelectedText` | 同上 | **未验证（本机无 Linux 环境）** | — |
| 4 | `insertText` | 写入 | **未验证（本机无 Linux 环境）** | — |
| 5 | `setFontColor` / `setFontSize` | Range 样式 | **未验证（本机无 Linux 环境）** | — |
| 6 | `getShapes` | 形状遍历 | **未验证（本机无 Linux 环境）** | — |
| 7 | `setShapeText` | 形状文本 | **未验证（本机无 Linux 环境）** | — |
| 8 | `addShape` | 新建形状 | **未验证（本机无 Linux 环境）** | — |
| 9 | `setShapePosition` | 坐标 | **未验证（本机无 Linux 环境）** | — |
| 10 | `deleteSlide` | 删除 | **未验证（本机无 Linux 环境）** | — |
| 11 | `addSlide` | 新建页 | **未验证（本机无 Linux 环境）** | — |
| 12 | `getSlideCount` | 计数 | **未验证（本机无 Linux 环境）** | — |
| 13 | `addShapeText` → 表格 | 表格能力（PPT 侧最难的一类） | **未验证（本机无 Linux 环境）** | — |
| 14 | `getTableCell` | 表格读回 | **未验证（本机无 Linux 环境）** | — |
| 15 | `setTableCell` | 表格写 | **未验证（本机无 Linux 环境）** | — |
| 16 | `addAnimation` | 时间轴（**最可能缺**） | **未验证（本机无 Linux 环境）** | — |
| 17 | `setSlideTransition` | 切换（**最可能缺**） | **未验证（本机无 Linux 环境）** | — |
| 18 | `setSlideBackground` | 背景 | **未验证（本机无 Linux 环境）** | — |
| 19 | `exportSlideAsImage` | 导出 | **未验证（本机无 Linux 环境）** | — |
| 20 | `saveAs` | 另存（决定用户会不会丢改动） | **未验证（本机无 Linux 环境）** | — |

## 汇总

- 支持：**未验证（本机无 Linux 环境）**
- 不支持（属性不存在）：**未验证（本机无 Linux 环境）**
- 抛错：**未验证（本机无 Linux 环境）**
- 部分支持：**未验证（本机无 Linux 环境）**
- 未验证：20 / 20

## 待补（S0 顺带要回答的）

- [ ] 单次调用往返时延（含 WPS 内部耗时）
      —— §2.4/§4.6：50ms 轮询平均往返用户无感，"为了低延迟必须上长连"不成立；
      轮询的决定性劣势是**定时器节流**，这个得实测确认影响面。
- [ ] 哪些动作语义不同（例如 WPS JSAPI 的形状类型枚举与 MSO 不同）
      —— §6 的语义差异清单要靠这里填。
- [ ] `--macro` 能否 attach 到已开着的 WPS 实例