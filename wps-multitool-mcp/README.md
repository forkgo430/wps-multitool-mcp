# wps-multitool-mcp

让 AI 直接操控你正开着的办公套件——WPS 文字、表格、演示，以及 Microsoft Office 365。

同一套 MCP 工具面，host 自动探测选路；两个套件都装时默认驱动 WPS，随时可切换。

## 能做什么

- **操控正开着的文档**：AI 操作的就是你眼前的活文档，所见即所得——不用关文档、不用上传
- **从零生成**：空白页起一份带目录和图表的文稿；把表格里的数字变成图表与透视表，再贴进演示
- **听懂位置**：支持"第 N 章 / 第 N 段 / 第 N 句 / 锚点文本"级别的落笔，交叉引用一次批量插完
- **套模板**：docx、pptx、xlsx 三种格式都能套模板生成，也能从现有文档反提模板；三份现成模板随包附带
- **改坏了能退**：破坏性操作前自动快照、支持逐步撤销、面板一键恢复
- **省上下文**：默认 40 个工具按需加载，其余 244 个用搜索或按组调出

## 快速开始

```powershell
npx -y wps-multitool-mcp
```

或加入 MCP 客户端配置：

```json
{
  "mcpServers": {
    "wps-multitool-mcp": {
      "command": "npx",
      "args": ["-y", "wps-multitool-mcp"]
    }
  }
}
```

重启客户端，打开一个 WPS 或 Microsoft Office 365 文档，然后直接说：

> 把这份合同的金额段落加粗，并在末尾补一段风险提示

## 环境

- Windows + 已安装 WPS Office 或 Microsoft Office 365（两个都装时默认驱动 WPS，`WPS_OFFICE=on` 切 Office）
- 支持 MCP 协议的客户端（Claude Code / Cursor / CodeBuddy / WorkBuddy 等）

## 完整文档

工具清单、配置项、双套件适配对账、故障排除：见 [GitHub 仓库 README](https://github.com/forkgo430/wps-multitool-mcp#readme)。

## License

MIT
