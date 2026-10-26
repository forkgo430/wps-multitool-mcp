install.ts
命令行子命令目录。目前只有 `wps-multitool-mcp install`（写客户端 MCP 配置，写前备份、写后校验）。
入口在 src/index.ts：argv[2] 是 install 时走这里，无参数时仍然启动 stdio MCP server（不能破坏握手）。
一旦这里的结构发生变化，请务必更新我。
