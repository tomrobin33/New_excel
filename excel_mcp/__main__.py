import asyncio
import sys

from .server import run_sse, run_stdio, run_streamable_http

# 如需命令行启动，可用 python -m excel_mcp.server 方式

if __name__ == "__main__":
    # 默认启动 stdio
    run_stdio() 