wps-com.ps1, wps-comd.ps1, gen-action-routes.js, manual/
Windows PowerShell COM 桥接脚本目录（含图表/去重/上下文等扩展实现）。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

发布走 `manual/` 下两条通道，结果等价（都产出内容与本地 HEAD 相同的无父级快照提交）：
常规 `publish-snapshot.js`（git 推，要求 github.com:443 通）；那条不通时用
`push-api.js`（纯 HTTPS 的 Git Data API，支持 `-f` / `--dry` / `--branch`）。