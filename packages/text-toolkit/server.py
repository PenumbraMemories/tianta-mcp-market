"""文本工具箱：MCP 市场示例包的入口脚本。

设计目标：
1. 只使用 Python 标准库，不访问网络和文件系统；
2. 与 mcp_runtime.py 一起打包，安装到任何目录都能独立运行；
3. 工具参数与 mcp.json 中声明的设置保持一致。
"""

import re

from mcp_runtime import run_server


TOOLS = [
    {
        "name": "text_stats",
        "description": "统计一段文本的字符数、单词数和行数。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "要统计的文本。",
                },
                "includePreview": {
                    "type": "boolean",
                    "description": "是否在结果中附带原文预览。",
                    "default": True,
                },
            },
            "required": ["text"],
            "additionalProperties": False,
        },
    },
    {
        "name": "slugify",
        "description": "把一段文本转换成适合做文件名或网址的 slug。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "要转换的文本。",
                },
                "separator": {
                    "type": "string",
                    "description": "单词之间的连接字符。",
                    "enum": ["-", "_"],
                    "default": "-",
                },
            },
            "required": ["text"],
            "additionalProperties": False,
        },
    },
]


def _require_text(arguments):
    text = arguments.get("text", "")
    if not isinstance(text, str) or not text.strip():
        raise ValueError("text 不能为空。")
    if len(text) > 50000:
        raise ValueError("text 不能超过 50000 个字符。")
    return text


def text_stats(arguments):
    """返回字符、单词、行数，以及可选的原文预览。"""

    text = _require_text(arguments)
    include_preview = bool(arguments.get("includePreview", True))

    words = re.findall(r"\w+", text, flags=re.UNICODE)
    lines = text.splitlines() or [""]

    result = {
        "characters": len(text),
        "words": len(words),
        "lines": len(lines),
    }

    if include_preview:
        compact = " ".join(text.split())
        result["preview"] = compact[:120]

    return result


def slugify(arguments):
    """生成只包含字母、数字、中文与连接符的 slug。"""

    text = _require_text(arguments)
    separator = arguments.get("separator", "-")
    if separator not in ("-", "_"):
        separator = "-"

    cleaned = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", separator, text)
    cleaned = cleaned.strip(separator).lower()

    if not cleaned:
        raise ValueError("这段文本里没有可用于 slug 的字符。")

    return {"slug": cleaned}


def call_tool(name, arguments):
    """按工具名分发。未知工具会返回明确的错误，方便市场联调。"""

    if name == "text_stats":
        return text_stats(arguments)
    if name == "slugify":
        return slugify(arguments)
    raise ValueError("未知工具：" + str(name))


if __name__ == "__main__":
    run_server("text-toolkit", "0.1.0", TOOLS, call_tool)
