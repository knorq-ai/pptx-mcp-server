from __future__ import annotations


def main():
    try:
        from .server import main as _server_main
    except ImportError as e:
        if "mcp" in str(e).lower():
            raise ImportError(
                "pptx-mcp-server CLI requires the 'mcp' extra. "
                "Install with: pip install 'pptx-mcp-server[mcp]'"
            ) from e
        raise
    return _server_main()


__all__ = ["main"]
