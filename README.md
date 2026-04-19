# pptx-mcp-server

MCP server for creating, reading, and editing PowerPoint (.pptx) presentations.
Provides 25 tools for slide management, shape/text manipulation, table operations,
composite layouts, and slide rendering -- all accessible via the Model Context Protocol.

## Installation

```bash
pip install pptx-mcp-server
```

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

- [python-pptx](https://python-pptx.readthedocs.io/) -- PPTX file manipulation
- [lxml](https://lxml.de/) -- XML processing
- [mcp](https://modelcontextprotocol.io/) -- Model Context Protocol SDK

### Optional

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

## License

MIT -- see [LICENSE](LICENSE) for details.
