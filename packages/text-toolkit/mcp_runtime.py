"""最小 MCP stdio 运行时，随示例包一起分发。

这个文件不依赖 Tianta 主仓库，复制到任何 Python MCP 包里都能直接用。
标准输出只写 JSON-RPC 消息；调试信息必须写 stderr。
"""

import json
import sys


PROTOCOL_VERSION = "2024-11-05"


def _send(message):
    sys.stdout.write(json.dumps(message, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _send_result(request_id, result):
    _send({"jsonrpc": "2.0", "id": request_id, "result": result})


def _send_error(request_id, message):
    _send({"jsonrpc": "2.0", "id": request_id, "error": {"code": -32000, "message": message}})


def run_server(server_name, version, tools, call_tool):
    """启动一个只实现 tools/list 与 tools/call 的 MCP 服务。"""

    try:
        sys.stdout.reconfigure(encoding="utf-8", newline="\n")
        sys.stdin.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass

    for line in sys.stdin:
        text = line.strip()
        if not text:
            continue

        try:
            request = json.loads(text)
        except ValueError:
            continue

        method = request.get("method")
        request_id = request.get("id")

        if method == "initialize":
            _send_result(
                request_id,
                {
                    "protocolVersion": PROTOCOL_VERSION,
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": server_name, "version": version},
                },
            )
            continue

        if method == "notifications/initialized":
            continue

        if method == "ping":
            _send_result(request_id, {})
            continue

        if method == "tools/list":
            _send_result(request_id, {"tools": tools})
            continue

        if method == "tools/call":
            params = request.get("params", {})
            if not isinstance(params, dict):
                params = {}

            name = params.get("name", "")
            arguments = params.get("arguments", {})
            if not isinstance(arguments, dict):
                arguments = {}

            try:
                value = call_tool(name, arguments)
                _send_result(
                    request_id,
                    {
                        "content": [{"type": "text", "text": json.dumps(value, ensure_ascii=False)}],
                        "isError": False,
                    },
                )
            except Exception as error:
                _send_result(
                    request_id,
                    {
                        "content": [{"type": "text", "text": str(error)}],
                        "isError": True,
                    },
                )
            continue

        if request_id is not None:
            _send_error(request_id, "不支持的方法：" + str(method))
