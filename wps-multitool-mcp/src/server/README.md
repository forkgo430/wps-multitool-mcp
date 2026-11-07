backend-selector.ts, mcp-server.ts, tool-registry.ts, tool-group-events.ts, tool-tiers.ts, tool-annotations.ts, tool-annotations.generated.ts, tool-metadata.ts, tool-verification.ts, tool-verification.generated.ts
常驻桥的客户端侧与服务端注册核心。一旦我被修改，请更新我的头部注释，以及所属目录的md。
新增文件请同时更新本清单。
真机验证状态的唯一真相源在 scripts/manual/tool-verification.json，生成物是两个 *.generated.ts，别手改。
MCP Server 服务与工具注册实现目录。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

backend-selector.ts（Linux B0 五层决策链 + Office 支持一期 host 维度）：
后端判定 com/addon/off 与 host 判定 wps/office 双维度正交。host 维度（WPS_OFFICE=on|off|auto，
缺省 auto）探测必须解析 ProgID 链到实际服务器二进制（LocalServer32 → exe 真身）——
WPS 独装机会顶名应答 Office ProgID（CoCreate 成功、Name 自报 Microsoft Word），ProgID
存在性/CoCreate/Name 全不可信。探测不可用/结论不一致 → 保守按 wps。
单测：scripts/manual/backend-selector-test.js（含 WPS_OFFICE × WPS_BACKEND 优先级矩阵 + 顶名劫持场景）。
