# wps-multitool-mcp

## 1.0.2

### Patch Changes

- e6a104d: 工具可发现性：`tools/list` 新增 `examples` 字段 + 同分排序改四元。
  
  ## 为什么是 patch（统筹者裁定，执行者只照做）
  
  `tools/list` 新增了可选字段，严格按 semver 属于 minor（返回结构变了）。
  
  **统筹者裁定降为 patch**，理由：只新增可选字段、没有删改任何既有字段，调用方不改代码也不会坏；
  且这一版的目标是 1.0.2 —— 只装本轮内容，不含积压的那两条 minor（`next-release` / `ppt-merge`
  已暂存到 `.changeset-hold/`，下一版放出）。
  
  > 定级归统筹者。这里只记录裁定与理由，执行者不做二次判断。
  
  ## tools/list 新增 examples
  
  - `ToolDefinition` 新增 `examples?: string[]`（用户真会说的话，口语，不是功能描述）。
  - `tools/list` 的每个工具现在都可能带 `examples` 数组。
  - **默认档 40 个工具的 `examples` 覆盖率从 2/40 补到 40/40。**
  
  ### 为什么是独立字段，不是拼进 description
  
  两条原因，都踩过：
  
  1. 搜索路径会把 `description` 截成**首行 120 字**，examples 拼在描述尾部会被整段吃掉。
  2. 拼进描述会污染模型的工具选择——描述回答的是"这工具能干啥"，`examples` 回答的是"用户会怎么说"。
     两者混在一起会让模型把示例当成能力描述。
  
  ### 读回判据：调完再调另一个工具把结果读回来
  
  这一条比"新增了 examples"本身更重要，是判据口径的修正：
  
  原来的判据只断言"回包不是空话"（长度 > 8 且开头不含 `undefined`/`null`）。这条判据会让
  **回 `success` 但内容全是 `undefined`** 的工具照样通过 —— 第 9 轮挖出的四个"读回形同虚设"
  （`get_textboxes` / `get_animations` / `open_presentation` / `create_kpi_cards`）
  在旧判据下全是绿的。
  
  现在的"读回值"判据是：**调用完目标工具之后，再调另一个工具把结果读回来，断言目标字符串出现在它的输出里。**
  表格单元格、形状文字、母版元素数、动画顺序、文稿锁定都由这条判据覆盖。
  
  导出图差分类同时加了一条前置：**连导两次确认字节一致**，否则比的是噪声。
  
  ## 同分排序：从纯字母序改为四元
  
  同分意味着"跟这句诉求的相关度差不多"，这时候真正该拉开差距的是**能不能直接用**和**验证过没有**：
  
  - 之前：`(score)` + `name` 字母序 → `static`、未列入的工具会和 `listed`+`live` 的混排，
    模型看到的前几条反而是不该首选的那批
  - 现在：`(score, isListed, isLive, name)`
  
  口径复用现成判定函数（`isListed()` 与 `isLiveVerified()`），不另写一套。
  排序键不随工具增删变化，所以补工具不会改变既有工具的相对位置。
  
  `tool-groups-test.js` 的一条断言随之改了判据：**性质没变，变的是探针**。
  原来要求"搜索 `ppt` 的前 5 条里必须有未列出的项"，四元排序后前 5 条全是已列出的基础档工具。
  现在放大 limit 验证"搜索能覆盖未列出工具"这个性质，另外单开一条断言把四元排序钉住。
  （原因写在 `scripts/manual/tool-groups-test.js` 里。）
- e6a104d: 内部重构：把 `execWpsAction` 拆成「选通道」与「跑通道」两层。**对外零感知。**
  
  ## 改了什么
  
  `src/client/wps-client.ts` 里原本一个函数同时管两件事：**决定走哪条通道**、**把请求送出去**。
  现在拆成两个：
  
  | 函数 | 只负责 |
  |---|---|
  | `pickTransport()` | **选一次**，把结果缓存起来 |
  | `runTransport(kind, ...)` | **只管跑**，不做任何选择/降级判断 |
  
  拆分的目的是给将来的第三条通道留一个**插入点**。
  
  ## 对外零变化
  
  - `execWpsAction(action, params, app)` 的**签名一个字没改**。
  - **所有调用方一个字没改** —— 全仓只有一个调用点，内部调用，没有对外接口变化。
  - 工具数量、`tools/list` 返回结构、参数、返回结构：**全部不变**。
  
  ## Windows 行为逐条不变
  
  拆分只动结构，不动语义。四件容易在重构里走样的事，逐条对过：
  
  - **超时常量**：`PS_TIMEOUT` 等原样，未调整。
  - **降级时机**：**只有"请求压根没送出去"（桥起不来，`DaemonUnavailableError`）才退回直跑**；
    其他错误一律如实抛出。
  - **超时分类**：超时那类错误**不走降级**。原因是请求可能已经在 WPS 里执行了一半，
    自动重放会变成**执行两次**。这条判断必须留在原处，不能被"顺手加个重试"改掉。
  - **`daemonBroken` 闩锁语义**：一旦落下，本进程后续一律直跑，**再也不会回 false**。
    闩锁落下时会把**已缓存的通道选择一起翻成 `direct`**，否则缓存会和闩锁打架。
  
  > 闩锁的"缓存"为什么等价于原来的每次重算：原逻辑每次调用都重跑一遍
  > `USE_DAEMON && !daemonBroken`；现在缓存，是因为 `daemonBroken` 一旦为 `true`
  > 就再也不会变回 `false`，所以这一次判断的结果对后续调用恒定。
  
  ## 降级口径仍全仓只有一份
  
  拆分前后的关键约束：**"什么时候允许退回直跑"这条判断，全仓只有一处**，
  拆完之后仍然只有一处。改的时候核过 —— `instanceof DaemonUnavailableError`
  在整个 `src/` 里**仅出现 1 次**。
  
  这条不是洁癖：降级判断一旦出现第二份，两份迟早会走偏，
  而症状是"有时候超时会执行两次"，极难排查。
  
  ## `addon` 通道：已留口，故意不实现
  
  `TransportKind` 里已经有第三个取值 `addon`，但**本轮没有任何地方会返回它**，
  `runTransport` 里也没有对应的 `case`（落到 `default` 抛错）。
  
  **为什么不写这个分支**：本机没有 Linux 环境，写出来的分支没有任何证据支撑，
  现在看它"跑通了"只是因为根本没被调用过。等真的要做 Linux 后端时，
  往 `pickTransport` 加一个判断、往 `runTransport` 加一个 `case` 即可 ——
  Windows 那两个分支一个字都不用动。
  
  > TypeScript 的穷尽检查会**逼着**补一个跑不通的分支，而那比没有更危险 ——
  > 所以这里选择留空并抛错，而不是写一个假装能用的实现。
  
  ## 这条为什么要记
  
  本轮的主题是"改了没入账正是要消灭的那类债"。第 10 轮的代码改动里，
  **功能性的部分都已写进 changeset**，只剩这条纯重构漏了。
  内部重构对调用方零影响，CHANGELOG 里没有它升级后行为也完全一致 ——
  但"改了没入账"这件事本身就是那类债，所以补上。
- e6a104d: PPT 工具族合并（见 `ppt-merge-2026-10-18.md`）的两个副作用点名。
  
  那条 changeset 的定级是 `minor` 且**保持不动**（破坏面窄、且旧参数名已被参数映射兜住）。
  但副作用必须在这里写明，否则升级的人看不到。
  
  ## 副作用一：7 个工具名不再出现在 `tools/list`
  
  PPT 动画族与切换族合并后，这 7 个工具名降为**别名**：
  
  ```
  wps_ppt_add_animation        wps_ppt_add_animation_preset
  wps_ppt_add_emphasis_animation  wps_ppt_set_animation_order
  wps_ppt_set_slide_transition wps_ppt_apply_transition_to_all
  wps_ppt_remove_slide_transition
  ```
  
  - **按名字直接调用仍然可用** —— 别名在入口处解析到规范名，旧名 + 旧参数都能跑。
  - **它们不再出现在 `tools/list` 里。** 任何按 `tools/list` 枚举、或按清单硬编码工具名的脚本，
    升级后会"找不到"这 7 个。
  - 工具总数因此从 291 → 284。
  
  > 如果你的代码枚举 `tools/list` 来生成 UI 或做校验，请把工具总数与工具名清单改成动态获取，
  > 别硬编码。
  
  ## 副作用二：`set_transition` 的 `sound` 参数被移除
  
  - WPS 的 COM 接口**根本没有"切换音效"这个能力**（那是 PowerPoint 的 `AddMediaObject2`）。
  - 该参数此前**永远不生效** —— 传了它和没传一样。
  - 现在从 schema 里删掉。留着它只会让人以为这个功能存在。
  - 传过 `sound` 的调用不会报错，但该参数不再有意义。
  
  ## 一句话
  
  以上两点都不是"功能变少了"，是**把两个一直存在的假象拿掉**：一个假清单、一个假参数。
- e6a104d: 第 4~9 轮的真 bug 修复（10 个）。这批此前一个字都没进 changeset，本次补记。
  
  ## 四个"从来就没成功过"的工具
  
  以下四个工具**在此之前每一次调用都没有真正生效**，但都返回 `success`。如果你以前用过它们，
  看到的结果不是"参数写错了"，而是"这功能实际上一直没工作"。
  
  - **`wps_ppt_insert_slide_image`** —— 它调用的动作名是 `insertImage`，而该动作属于 **Word** 段
    （`Get-WpsWord` + `InlineShapes.AddPicture`）。PPT 工具去调它必然拿不到文档对象，
    于是每次调用都报"不能对 Null 值表达式调用方法"。现在改走 PPT 段的 `insertPptImage`。
  - **`wps_ppt_create_kpi_cards`** —— 回 `success`，但**一张卡片都没建**。
    PowerShell 把单元素数组在管道里摊平，`$cards.Count` 变成空值，`for` 循环一次都不跑。
  - **`wps_ppt_create_mini_charts`** —— 同一个数组摊平问题导致 0 个图表；
    且多项时 `data[].value` 从 JSON 过来是 `Int32`，直接赋给 `TextRange.Text` 抛"指定的转换无效"。
  - **`wps_ppt_align_objects`** —— 回"对象对齐完成"，实际把**整页所有形状**一起左对齐
    （同页的表格、图片、占位符都被搬走）。原因是没给形状清单时，`Shapes.Range($null)`
    在 WPS 上不报错，含义是"整页"。现在 `shapeIndices` 改为必填，底层对空清单**响亮报错**而不是替用户选。
  
  ## 四个"读回形同虚设"的工具
  
  以下工具返回成功，但回包里的关键字段恒为 `undefined`，等于读回功能不存在：
  
  - `wps_ppt_get_textboxes` —— 读 `data.textboxes`，底层回的是 `textBoxes`（大写 B），
    永远打印 `第N页文本框: []`
  - `wps_ppt_get_animations` —— 按 `effect` / `trigger`（字符串）读，底层回的是
    `effectType`（MsoAnimEffect 枚举数字）与 `duration`，列表一直是 `1. [矩形 1] undefined - undefined`
  - `wps_ppt_open_presentation` —— 读 `data.filePath`，底层回的是 `path`，永远打印 `路径: undefined`
  - `wps_ppt_create_kpi_cards` / `wps_ppt_create_mini_charts` —— 读 `cardCount` / `chartCount`，
    底层回的是 `count`，永远打印"数量: undefined 个"
  
  ## 两个占位符编号错的
  
  - `wps_ppt_set_slide_subtitle` 与 `wps_ppt_set_slide_content` **必然**报
    `placeholder not found`。原因是占位符编号是查表抄的而不是实测的：
    实测 WPS 的 `layout='title'` 页里副标题是 `PlaceholderFormat.Type=4`（居中标题=3），
    `layout='title_content'` 页里正文是 **2**；而代码里副标题找 2、正文找 7，两个都不对。
    现在按实测编号匹配，报错信息也改成说明"哪种版式能写"。
  
  ## 一个加一次生成两条的
  
  - `addAnimation` 加**一个**动画实际生成**两条**一模一样的效果，删一条还剩一条，看起来像删不掉。
    原因是 `AddEffect(shape, id, level)` 的第三个参数是 **Level（文本级别）**而不是触发时机，
    对没有文本的形状必须传 `msoAnimateLevelNone(0)`。
  
  ## 最狠的一条：切换效果被静默改写
  
  - `applyTransitionToAll` 读的是 `$p.transition`，而工具一直传的是 `effect` →
    **整份演示文稿的切换效果被静默设成 fade**。你要 push，得到的是 fade。
    参数名不匹配不会报错，门禁和用户都没发现。同批还修了 `setSlideTransition`
    （把效果名字符串直接赋给 `PpEntryEffect` 枚举会抛错、且不认 `none`，导致"取消切换"必炸）。
  
  ## 表现最唬人的一条：像 69 个工具一起坏了
  
  - `addAnimation` 曾用 `Shapes.Item($p.shapeName)` 取形状，而工具传的是 `shapeIndex` →
    拿到空对象后调 `AddEffect`，于是该页的 `Shapes` / `Background` 全部报 Null，
    而文稿列表仍然正常。**症状是"大量 PPT 工具同时失效"，根因只有一个。**
    现在同时认 `shapeName` 与 `shapeIndex`。
  
  ## 排版类
  
  - `auto_layout` / `smart_distribute` 的形状收集与赋值改为逐个 `try/catch`：
    页面里有 3D / 表格 / 图片时读 `Width` / `Height` 会抛"指定的转换无效"，此前会让整页排版失败；
    现在跳过读不到尺寸的形状，并在结果里回报 `skipped` 计数。
  
  ## 测试口径同步修正
  
  - PPT 冒烟从顺序执行改为**每条重建全新靶子**（`ppt-isolated-verify.js`）。顺序执行时前一条留下的
    状态会让"导出图无变化"分不清是工具坏了还是靶子已经变了 —— 实测过同一批工具顺序跑大面积失败、
    单跑全绿。
  - 21 条"需外部素材 / 前置不足"的用例改判据后救回，其中 `remove_animation`、`set_animation_order`
    这类所谓"靶子没有可操作对象"的，本质是**用例设计没跟上**，不是工具问题。
  - 真机覆盖矩阵 100 行 → 103 行，执行 74 条 → **98 条**，全部通过；跳过项由 21 条收敛到 **5 条**
    （5 条跳过各有实测原因，不是漏跑）。
- e6a104d: 停止让 AI 对用户说假话：修正 SKILL.md 里已经失效的 PPT 验证状态陈述，并给这类陈述加门禁。
  
  ## 这不是错别字
  
  `skills/wps-multitool/SKILL.md` 与 `references/recipes.md` 里长期写着：
  
  > PPT 相关的工具没有真机覆盖验证过。……目前整个 ppt 组都是 static。
  
  而 PPT 的真机隔离验证早已做到 **98/98**。问题的严重性在于：**SKILL.md 是 AI 读来决定
  "怎么用这个 MCP、以及该怎么向你汇报"的**。这句话会让 AI
  
  1. 主动对用户说"这部分未在真机验证"——**说的不是自己的判断，是文档里抄来的假话**；
  2. 低估整组 PPT 工具，把本来能用的能力报成不可用。
  
  所以这次改的不是措辞，是**行为依据**。
  
  ## 改了什么
  
  - `SKILL.md` 第三节改为当前真值：组级口径仍记 `static`（矩阵里 5 条有据可查地跳过），
    但真机覆盖已上量（PPT 组 100 个工具，隔离验证矩阵 103 行，实际执行 98 条全部跑通）。
    刻意**没有**改成"PPT 全部已验证"——那是不准确的另一种。
  - `references/recipes.md` 里同一句假话出现在配方三的"坑"里，一并修正。
  - 用户机器上已安装的副本（`~/.workbuddy/skills/wps-multitool/`）同步更新。
  
  ## 新增门禁（这次真正值钱的部分）
  
  - **`scripts/manual/skill-content-truth.js`**（新文件）：维护"已知失效断言"清单，
    并提供把文档里每个「分组 + 验证结论」与 `tool-verification.json` **逐句对账**的能力。
    真值来自台账，不是硬编码一句"PPT 已验证"——台账变了而文档没跟着变，门禁就红。
  - **`skill-verify.js`** 断言数 20 → 30：
    - 扫 SKILL.md 与全部 `references/`，命中已知失效断言 → FAIL
    - 分组验证结论与台账矛盾 → FAIL（例：台账里某组有 `live` 工具，文档却说该组"没有真机覆盖"）
    - 文档里引用的工具数 / 执行条数 / 跳过条数 / `live` 条数，必须与台账 evidence 里抠出来的
      同一批数字一致 → 对不上 FAIL
  - **`skill-verify.js` 新增副本同步检查**：比对仓库版与已安装副本的 md5，不一致时打印警告
    （不 FAIL——副本可能被用户自己改过）。可用 `--no-sync-check` 跳过（CI 上没装副本）。
  
  门禁在加入后被验证过**确实会红**：把原句写回 SKILL.md 后，3 条断言失败、进程退出码 1；改回后恢复全绿。

---

## 1.0.2 的补记（2026-10-20）

> 本节是**事后补账**，不改版本号。发布 1.0.2 时清空了 `.changeset/`，导致下面这些
> 已经随本版发出的内容**没有进 CHANGELOG**；而它们对应的 changeset（`next-release` /
> `ppt-merge-2026-10-18`）还留在待发区，会让下一版把它们**再记一次**。
> 已按 npm 上 `wps-multitool-mcp@1.0.2` 的实际包内容逐项核对（文件实存，不是推断），
> 两条 changeset 已归档到 `.changeset-hold/`。
>
> 核对方式：`npm pack wps-multitool-mcp@1.0.2` → 看包内是否存在对应文件；
> 同一份清单对照 `wps-multitool-mcp@1.0.1`：以下四项在 1.0.1 里**都不存在**，在 1.0.2 里**都存在**。

- **写入可恢复**：`dist/tools/common/snapshot.js`（`wps_common_snapshot_create` / `_list` / `_restore`；
  `wps_common_batch` 检测到破坏性步骤时自动快照）
- **验证状态**：`dist/server/tool-verification.js` + `tool-verification.generated.js`（每工具 `live` / `static`，
  默认档只收 live）
- **一键接入**：`dist/cli/install.js`（`wps-multitool-mcp install`，写前备份、写后回读、失败自动还原）
- **别名参数映射**：`dist/server/tool-aliases.generated.js` 的 `applyAliasParams`
  （族合并时旧参数名要能搬到新参数上；没这条的话旧调用会 `Missing required parameter`，
  或被静默忽略后按默认值乱改）
