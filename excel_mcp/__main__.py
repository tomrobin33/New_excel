import asyncio
import typer
import sys

deps = [
    ("typer", "typer"),
    ("openpyxl", "openpyxl"),
    ("mcp", "mcp"),
    ("click", "click"),
    ("rich", "rich"),
    ("pygments", "pygments"),
    ("markdown_it_py", "markdown_it_py"),
    ("mdurl", "mdurl"),
    ("shellingham", "shellingham"),
    ("typing_extensions", "typing_extensions"),
    ("colorama", "colorama"),
]

for name, mod in deps:
    try:
        __import__(mod)
        print(f"[LOG] {name} 导入成功", file=sys.stderr)
    except ImportError:
        print(f"[ERROR] {name} 导入失败", file=sys.stderr)

from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(help="Excel MCP Server")

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    print("Excel MCP Server - SSE mode")
    print("----------------------")
    print("Press Ctrl+C to exit")
    try:
        asyncio.run(run_sse())
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    print("Excel MCP Server - Streamable HTTP mode")
    print("---------------------------------------")
    print("Press Ctrl+C to exit")
    try:
        asyncio.run(run_streamable_http())
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    print("Excel MCP Server - Stdio mode")
    print("-----------------------------")
    print("Press Ctrl+C to exit")
    try:
        run_stdio()
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

if __name__ == "__main__":
    app() 