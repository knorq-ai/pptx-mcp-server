# pptx-mcp-server

MCP server for creating, reading, and editing PowerPoint (.pptx) presentations.
Provides 25 tools for slide management, shape/text manipulation, table operations,
composite layouts, and slide rendering -- all accessible via the Model Context Protocol.

## Installation

`pptx-mcp-server` ships with two install paths: a minimal library install for
driving the engine programmatically, and an `[mcp]` extra for running the MCP
server CLI.

```bash
# Library usage (no MCP runtime — only python-pptx + lxml):
pip install pptx-mcp-server

# MCP server (includes the mcp SDK):
pip install 'pptx-mcp-server[mcp]'
```

The MCP SDK lives behind the `[mcp]` extra so that pure-library consumers do
not pay for the MCP SDK + anyio transitive dependencies at install time. The
`pptx-mcp-server` CLI and `pptx_mcp_server.server` module require the `[mcp]`
extra; importing them without it raises a clear ImportError pointing back here.

## Claude Desktop Configuration

Add the following to your Claude Desktop MCP config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "pptx-editor": {
      "command": "pptx-mcp-server"
    }
  }
}
```

Or, if running from source:

```json
{
  "mcpServers": {
    "pptx-editor": {
      "command": "python",
      "args": ["-m", "pptx_mcp_server"]
    }
  }
}
```

## Use as a Python library

`pptx-mcp-server` also ships a pure-Python engine you can drive directly,
without starting an MCP server. The `engine` and `theme` modules have no
dependency on the MCP SDK at import time, so you can build decks from scripts,
notebooks, or batch jobs.

Install it the same way you would any other package — the bare install does
not pull the MCP SDK:

```bash
# From PyPI (once published) — library only:
pip install pptx-mcp-server

# From a local checkout (editable install during development):
pip install -e /path/to/pptx-mcp-server
```

Then compose a deck by calling the engine functions directly:

```python
from pptx_mcp_server.engine.pptx_io import create_presentation, open_pptx
from pptx_mcp_server.engine.shapes import add_textbox
from pptx_mcp_server.engine.slides import add_slide
from pptx_mcp_server.theme import MCKINSEY

out = "deck.pptx"
create_presentation(out, width_inches=13.333, height_inches=7.5)
add_slide(out, layout_index=6)
add_textbox(
    out,
    slide_index=0,
    left=1.0, top=1.0, width=10.0, height=1.0,
    text="Hello from pptx_mcp_server",
    font_name=MCKINSEY.fonts.get("body", "Arial"),
    font_size=24,
    bold=True,
)

prs = open_pptx(out)
print(f"Slides: {len(prs.slides)}")
```

The `mcp` SDK is an **optional** extra (`pip install 'pptx-mcp-server[mcp]'`)
and is only required when launching the `pptx-mcp-server` CLI / importing
`pptx_mcp_server.server`. Nothing in `pptx_mcp_server.engine` or
`pptx_mcp_server.theme` imports it — an AST-level CI guardrail in
`tests/test_library_usage.py` enforces that, and
`tests/test_packaging_extras.py` enforces the packaging side of the split.

## Tools

### Presentation

| Tool | Description |
|------|-------------|
| `pptx_create` | Create a new blank PPTX file (default 16:9 widescreen) |
| `pptx_get_info` | Get presentation overview: slide count, dimensions, shape summaries |
| `pptx_read_slide` | Read detailed content of a slide -- all shapes, text, tables |
| `pptx_list_shapes` | List all shapes on a slide with indices, types, positions, text preview |

### Slides

| Tool | Description |
|------|-------------|
| `pptx_add_slide` | Add a new slide with a specified layout |
| `pptx_delete_slide` | Delete a slide by 0-based index |
| `pptx_duplicate_slide` | Duplicate a slide (appended at end) |
| `pptx_set_slide_background` | Set solid background color for a slide |
| `pptx_set_dimensions` | Set presentation slide dimensions in inches |

### Text & Shapes

| Tool | Description |
|------|-------------|
| `pptx_add_textbox` | Add a text box with full formatting options |
| `pptx_edit_text` | Edit text content and formatting in an existing shape |
| `pptx_add_paragraph` | Append a new paragraph to an existing shape |
| `pptx_add_shape` | Add an auto shape (rectangle, oval, arrow, chevron, etc.) |
| `pptx_add_image` | Add an image (PNG, JPG, SVG) to a slide with optional sizing |
| `pptx_delete_shape` | Delete a shape from a slide by index |
| `pptx_format_shape` | Reposition, resize, or restyle an existing shape |

### Tables

| Tool | Description |
|------|-------------|
| `pptx_add_table` | Add a professionally formatted table with headers and alternating rows |
| `pptx_edit_table_cell` | Edit a single table cell's text and formatting |
| `pptx_edit_table_cells` | Batch edit multiple table cells |
| `pptx_format_table` | Apply bulk formatting to an entire table |

### Composites

| Tool | Description |
|------|-------------|
| `pptx_add_content_slide` | Add a content slide with action title, divider, footnote, page number |
| `pptx_add_section_divider` | Add a section divider slide with dark background and accent stripes |
| `pptx_add_kpi_row` | Add a row of KPI callout boxes |
| `pptx_add_bullet_block` | Add a bulleted text block with multiple items |

### Rendering

| Tool | Description |
|------|-------------|
| `pptx_render_slide` | Render slide(s) to PNG via LibreOffice for visual verification |

## Dependencies

Required (installed by `pip install pptx-mcp-server`):

- [python-pptx](https://python-pptx.readthedocs.io/) -- PPTX file manipulation
- [lxml](https://lxml.de/) -- XML processing

Optional extras:

- `[mcp]` -- `pip install 'pptx-mcp-server[mcp]'` pulls the
  [mcp](https://modelcontextprotocol.io/) SDK, required for the
  `pptx-mcp-server` CLI / `pptx_mcp_server.server`.
- `[validation]` -- `pip install 'pptx-mcp-server[validation]'` pulls
  [fontTools](https://fonttools.readthedocs.io/) and enables the real-font
  overflow validation path. See "Layout validation" below.

### Layout validation

`check_deck_extended` / `check_text_overflow` ship two paths for text overflow
detection:

- `font_source="heuristic"` (default) -- zero-deps; uses the in-tree width
  heuristic (shared with `add_auto_fit_textbox`).
- `font_source="real"` (opt-in, needs the `[validation]` extra) -- reads
  real advance widths from TTF/TTC via fontTools. This gives an
  **independent source of truth** against the heuristic, so that drift
  between the heuristic and PowerPoint's actual rendering cannot hide an
  overflow behind the auto-fit primitive.

```python
from pptx import Presentation
from pptx_mcp_server.engine.font_metrics import discover_system_fonts
from pptx_mcp_server.engine.validation import check_deck_extended

prs = Presentation("deck.pptx")
report = check_deck_extended(
    prs,
    font_source="real",
    font_paths=discover_system_fonts(),  # or {"Arial": "/path/to/Arial.ttf"}
)
print(report["summary"])
```

Fonts that can't be resolved fall back to the heuristic per-paragraph and
emit a `font_not_measured` warning so partial coverage is still useful.

### System tools

- **LibreOffice** -- required for `pptx_render_slide` (PPTX to PDF conversion).
  Install with `brew install --cask libreoffice` (macOS) or your system package manager.
- **pdftoppm** (poppler-utils) -- required for `pptx_render_slide` (PDF to PNG conversion).
  Install with `brew install poppler` (macOS) or `apt install poppler-utils` (Debian/Ubuntu).

## Development

```bash
# Install in editable mode
pip install -e .

# Run tests
python -m pytest tests/ -v
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for development conventions.

## License

MIT -- see [LICENSE](LICENSE) for details.
