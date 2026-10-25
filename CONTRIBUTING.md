# 贡献指南

## 环境

- Windows 10/11 + WPS Office（真机验证绕不开，很多接口只有 WPS 见过世面）
- Node.js ≥ 18
- 首次：`cd wps-multitool-mcp && npm ci && npm run build`

## 常用命令

```powershell
cd wps-multitool-mcp
npm run build          # TypeScript 构建
npm run verify         # 快测门禁（FAST 14 项，秒级，不起 WPS）
npm run verify:all     # 全量门禁（21 项 = FAST 14 + HEAVY 7，含真机探测，约 3 分钟）
npm run gen:annotations  # 重新生成「工具 → PS 动作」映射表
node scripts/manual/gen-doc-stats.js   # 重新生成 README / 面板里的统计数字
```

> 门禁清单的唯一真源是 `scripts/manual/verify-all.js:29-44`（FAST）与 `:47-55`（HEAVY）。
> **本文档里的项数可能过期，以那里为准** —— 这个仓库吃过"两处维护"漂移的亏。
> 其中 6 项 HEAVY 需要真 WPS：excel/word/ppt 三个 caps-probe、snapshot-verify、bridge-test、cancel-verify。
> 它们跑起来会占满本机的 COM 通道（见 `docs/agents/README.md` 的 WPS 锁）。

## 三条硬规矩

1. **生成物不许手改**
   - `src/server/tool-annotations.generated.ts`（`npm run gen:annotations` 产出）
   - `src/client/action-routes.generated.ts`（`node scripts/gen-action-routes.js` 产出）
   - README / README.en.md 的 `<!-- AUTO:… -->` 区间、panel.html 的数字（`gen-doc-stats.js` 产出）
   CI 会跑生成器再 `git diff --exit-code`，手改过或者忘了重跑都会红。

2. **统计数字只许来自注册表**
   总数、按组分布、注解分布，一律走真实注册路径（`registerBuiltinTools()` + `registerAll(allTools)`）。
   少做第一步就少算 12 个内置工具——这种"两处维护"的漂移吃过亏。

3. **PS 动作改动要重跑路由表**
   改完 `scripts/wps-com.ps1` 就跑 `node scripts/gen-action-routes.js`（它按 case 体推断动作属于哪个应用）。
   同一个动作名出现两个 case 时 PowerShell 只认第一个，后者是死代码——加动作前先 `findstr` 确认没重名。
   ⚠️ 漏跑**没有任何门禁会告诉你**：`src/client/action-routes.ts:53` 会静默回落成 `'common'`。
   ⚠️ case 格式必须是"独占一行 + 花括号后无内容"，否则 `gen-action-routes.js:35` 的正则不认，路由表静默缺项。

4. **生成器跑在 build 之后**
   `gen-tool-annotations.js` / `gen-doc-stats.js` / `gen-tools-doc.js` 都 `require dist/`，
   读的是编译产物。忘了先 `npm run build`，它们会对着上次的 `dist/` 跑——**不报错，但内容全是旧的**。
   正确顺序见下面「新增一个工具的流程」第 9 步。

5. **推送只有一条通道**
   ```powershell
   node scripts/manual/push-api.js --history
   ```
   `git push` 到 github.com:443 在本机一直被 `Connection was reset`，必须走 Git Data API。
   禁用 `push-api.js -f`（无父级快照模式会把全部提交压成 1 个根提交）。
   本地与远端 SHA 恒不同（API 重建的 SHA 由服务端生成），**校验同步要看 tree，不看 commit 数**。

5. **真机验证要持锁**
   全机 COM 通道只有一个。真机脚本（`*caps-probe*` / `snapshot-verify` / `bridge-test` / `cancel-verify` /
   `verify:all` / `ppt-isolated-verify`）开跑前建 `docs/agents/WPS.lock`，跑完删除。
   细节见 `docs/agents/README.md`。

6. **没跑过就写"未验证"**
   验证状态只有两档：`live`（有真机证据）和 `static`（只有静态证据）。**没有第三档。**
   把没验过的说成验过，是在拿用户的文档赌。

---

## 一些会咬人的坑（都不是我编的，是这个仓库真实交过学费的）

| 现象 | 真因 |
|---|---|
| 改了坐标，导出图字节完全一致 | **顺序跑的冒烟不可信**——前一条用例的残留状态让后一条分不清是工具坏还是靶子已经变了。用 `ppt-isolated-verify.js`，每条重建全新靶子 |
| 导出图差分结果毫无意义 | 差分前**必须连导两次确认字节一致**，否则比的是噪声 |
| "回 success，内容全是 `undefined`" | 最阴的一类假绿。读回类用例要断言**具体值**（调完再调另一个工具读回来），不能只断言回包长度 > 8 |
| 调了工具，对象报"指定的转换无效" | `[Math]::Ceiling/Floor` 返回 `double`，赋给 COM 的 `Single` 属性会抛错；`Sort-Object` 不能用来比较 COM 对象 |
| 数组参数传了没反应 | PowerShell 管道会把**单元素数组摊平**，`.Count` 为空、`for` 一次都不跑。要 `@(...)` 包裹 |
| 给整页加效果，结果整页都变了 | `$slide.Shapes.Range($null)` 在 WPS 上**不是报错，是"整页"**。清单为空时必须响亮报错，不能替用户选 |
| 加一个效果生成两条 | `AddEffect(shape, id, level)` 第三参是 Level（文本级别）不是触发时机。无文本形状要传 `msoAnimateLevelNone(0)` |
| 枚举直接赋字符串必抛错 | COM 枚举不能赋字符串，`EntryEffect = "fade"` 会抛 |
| 中文工具名搜不到 | `searchTools` 的同义词扩展只在**原词没命中**时上场 |
| 改完"传了没反应" | TS→PS 要走三步：改名（`$ParamAlias`）→ 摊平嵌套（`$StyleFlatten`）→ 值形态转换。少一步就是静默失效，`param-shape-verify.js` 天天扫这个 |

> 大表在 `docs/agents/PPT_SMOKE_MATRIX.md` 与 `docs/HANDOFF.md`。上表只挑了最容易重复踩的。

## 新增一个工具的流程

> 完整 15 步（含内部行号与验收细节）在 `docs/agents/ROLES.md` §6。这里只列**顺序**和**会静默失败的那几条**。

### 顺序

1. **建文件/加文件** —— `src/tools/<app>/<file>.ts` 里写 definition（`name` / `category` / `inputSchema`）+ handler
   （参考 `word/gaps.ts` 与 `excel/sheet-gaps.ts`）
2. **登记** —— 同文件底部的 `xxxTools` 数组（样板：`src/tools/word/format.ts:528`）
3. **挂载** —— 只有**新建子文件**才需要改 `src/tools/<app>/index.ts` 两处：组数组 + 导出块
4. **补 README** —— 只有新建子文件才改 `src/tools/<app>/README.md` 的文件清单（有门禁）
5. **接动作** —— handler 里调 `wpsClient.executeMethod('<action>', …, WpsAppType.WRITER)`。
   动作名写错的话 PS 侧会回 `Unknown action`

### ⚠️ 三条漏了不报错、只会静默出错的

6. **加 `case` 前先查重名。** 同一个动作名出现两个 `case` 时 PowerShell 只认第一个，
   后者是**死代码**；而 `gen-action-routes.js` 遇到重名**只 warn 不 exit**，
   你会拿到一个"生成成功"的假象。

   ```powershell
   findstr /n "^\s*""<动作名>""\s*\{" scripts\wps-com.ps1
   ```

7. **`case` 必须独占一行、花括号后无内容** —— `gen-action-routes.js:35` 的正则只认这个格式。
   写成 `"name") { ... }` 路由表会静默缺项。

8. **漏跑 `node scripts/gen-action-routes.js` 没有任何门禁会告诉你** ——
   `src/client/action-routes.ts:53` 会把动作静默回落成 `'common'`。

   ```powershell
   node scripts/gen-action-routes.js
   ```

### ⚠️ 一条顺序反了会整场作废的

9. **`npm run build` 必须排在所有生成器之前。** `gen-tool-annotations.js` / `gen-doc-stats.js` /
   `gen-tools-doc.js` 都 `require dist/` —— 读的是**编译产物**。忘了先 build，生成器会对着上一次的
   `dist/` 跑，**跑完不报错，但内容全是旧的**。

   ```powershell
   npm run build                        # ← 先
   node scripts/gen-action-routes.js    # 再跑生成器
   node scripts/manual/gen-tool-annotations.js
   npm run gen:docs
   ```

### 收尾

10. **验证状态** —— 要让单个工具标成 `live`，得在 `scripts/manual/tool-verification.json` 里单条写 `tools`
    （组级默认已经替它表态了）。没真机跑过就留 `static`，**不要写 live**。
11. **进默认档** —— 要 `tool-tiers.ts` 的白名单**和**真机 `live` **两个条件都满足**，缺一会被 `basic-tier-verify` 拦。
12. **近名工具补 `notFor`** —— 在 `src/server/tool-metadata.ts` 的 `CURATED` 里写清"这个不是用来干 X 的"。
    漏了 `search-meta-verify` **不会红**，但检索会静默退化成猜。
13. `npm run verify`（快测）→ `npm run verify:all`（全量，含真机）。


## 提交信息

沿用 `type(scope): 摘要`，例如：

```
feat(word): 补 get_tables / get_table_data
fix(ci): 修正文档统计门禁的 pathspec
docs: 双语 README 门面化
```

**说清"为什么"比说清"改了什么"更重要**——这个仓库的注释和提交信息都以"踩过的坑"为主，
便于下一个人不重踩。

## 关于 LICENSE

根目录 MIT 许可与版权声明必须保留（协议要求）。文档与代码里不再引用其它项目的署名，
但许可文本本身不可删除或改写。
