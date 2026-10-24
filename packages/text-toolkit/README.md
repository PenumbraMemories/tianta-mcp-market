# 文本工具箱

一个用于演示 Tianta MCP 市场安装流程的示例包。它不访问网络，也不读写文件。

## 工具

| 工具 | 作用 | 设置 |
| --- | --- | --- |
| `text_stats` | 统计字符、单词和行数 | `includePreview`：是否返回原文预览 |
| `slugify` | 生成适合文件名或网址的 slug | `separator`：`-` 或 `_` |

## 本地运行

```powershell
python server.py
```

脚本通过标准输入输出使用 MCP JSON-RPC 协议，不要直接向 stdout 打印日志。

## 打包

zip 根目录必须直接包含 `mcp.json`、`server.py`、`mcp_runtime.py` 和本文件，不能多包一层目录。

```powershell
Compress-Archive -Path .\* -DestinationPath ..\text-toolkit-0.1.0.zip
```
