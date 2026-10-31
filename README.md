# Tianta MCP Market

Tianta 的本地 MCP 市场索引仓库。市场索引只有一个 `registry.json`，每个包发布在 GitHub Releases 中。

## 目录结构

```text
registry.json                 市场索引
packages/<package-id>/        包的源码（便于审阅）
.github/workflows/            可选：自动打包与发布
```

Release 里的 zip 必须把 `mcp.json`、入口脚本和依赖文件放在根目录，不能多套一层文件夹。Tianta 下载后会解压到本地 `data/mcps/<package-id>/`。

## 添加一个包

1. 在 `packages/<package-id>/` 放源码。
2. 确认 `mcp.json` 的 `id` 与包目录名一致。
3. 打包 zip，计算 SHA256 和文件大小。
4. 在 GitHub 创建 Release，上传 zip。
5. 在 `registry.json` 的 `packages` 数组中加入对应条目：

```json
{
  "id": "package-id",
  "name": "显示名称",
  "summary": "一句话说明",
  "description": "完整说明",
  "version": "0.1.0",
  "author": "作者",
  "tags": ["标签"],
  "homepage": "https://github.com/...",
  "permissions": ["无文件访问"],
  "workspaceAccess": "none",
  "minTiantaVersion": "1.0.0",
  "download": {
    "url": "https://github.com/.../releases/download/v0.1.0/package-id-0.1.0.zip",
    "sha256": "小写十六进制哈希",
    "size": 4204
  }
}
```

`workspaceAccess` 只能是 `none`、`read` 或 `read-write`。市场安装器只接受 HTTPS 下载、校验 SHA256，并在安装时拒绝路径穿越和符号链接。

## 安全说明

MCP 子进程以当前用户权限运行，市场页会展示包声明的权限，但权限文本不是操作系统级沙箱。安装前请审阅源码。
