# Changelog

All notable changes to `pptx-mcp-server` are documented in this file.

## v0.3.0 — BREAKING: structured MCP surface

Flagged by OpenAI/Codex 5.4 as the remaining blockers to LOGO-READY status;
closes #97, #98, #99, #100.

### BREAKING CHANGES

#### #97 — Structured MCP parameters (`*_json: str` removed)

Every JSON-string-wrapped parameter was replaced with a native Python type.
FastMCP generates JSON Schema from Python typing, so agents send native
objects — no client-side string-encoding, no server-side `json.loads` layer.

| Tool | Before | After |
|------|--------|-------|
| `pptx_add_flex_container` | `items_json: str` | `items: list[dict]` |
| `pptx_add_responsive_card_row` | `cards_json: str` | `cards: list[dict]` |
| `pptx_build_slide` | `spec_json: str` | `spec: dict` |
| `pptx_build_deck` | `slides_json: str` | `slides: list[dict]` |
| `pptx_add_chart` | `chart_json: str` | `chart: dict` |
| `pptx_add_kpi_row` | `kpis_json: str` | `kpis: list[dict]` |
| `pptx_add_table` | `rows_json: str` + `col_widths_json: str` | `rows: list[list[Any]]` + `col_widths: list[float] \| None` |
| `pptx_add_bullet_block` | `items_json: str` | `bullets: list[str]` |
| `pptx_edit_table_cells` | `edits_json: str` | `edits: list[dict]` |

Dict-content validation (e.g. `CardSpec` unknown-key rejection, flex item
sizing keys) continues to run at the tool boundary and returns
`INVALID_PARAMETER` with `parameter` + `message` pointing at the bad card/item.

#### #98 — `result` is always a dict

Previously `result` was sometimes a string (default path), sometimes a dict
(auto-render path), sometimes a JSON-encoded string (`check_layout` detailed
path). Agents could not write a single parser.

v0.3.0 normalizes: **`result` is always a dict**, minimally
`{"message": "..."}`. Composite tools with auto-render integration add
`preview_path` or `render_warning` as additional keys on the same dict
(previously the payload was re-wrapped under a `value` key).

```python
# Before (v0.2.x)
{"ok": true, "result": "Added content slide [1]: ..."}
{"ok": true, "result": {"value": "Added ...", "preview_path": "/tmp/s.png"}}

# After (v0.3.0)
{"ok": true, "result": {"message": "Added content slide [1]: ..."}}
{"ok": true, "result": {"message": "Added ...", "preview_path": "/tmp/s.png"}}
```

`_success()` now requires a dict argument; passing a non-dict raises
`TypeError` defensively (catches regressions during development).

#### #99 — Flatten double-decode in `pptx_check_layout(detailed=True)`

The detailed path previously returned `_success(json.dumps(result))`,
forcing consumers to `json.loads(response)["result"]` → still a string →
`json.loads(that)`. Fixed to pass the dict inline:

```python
# Before (v0.2.x)
json.loads(tool_response)["result"]   # → '{"slides":[...]}'  (still a string)

# After (v0.3.0)
json.loads(tool_response)["result"]   # → {"slides": [...]}   (single decode)
```

Non-detailed `pptx_check_layout` keeps the human-readable summary, wrapped
per #98 under `result["message"]`.

#### #100 — Consistent malformed-JSON handling

With #97 in place, all raw `json.loads()` calls inside tool bodies are gone
— FastMCP's type validation fires before the tool function runs, returning a
structured error if the caller sent the wrong shape. The server layer now
contains **zero** `json.loads()` / `JSONDecodeError` references in tool
bodies (only the `json.dumps()` used to build the outer envelope remains).

Malformed input no longer surfaces as `INTERNAL_ERROR: JSONDecodeError`;
instead FastMCP returns a schema-validation error with parameter info.
Explicit `isinstance` guards at tool entry catch the remaining edge cases
(e.g. `rows=None`) with a dedicated `INVALID_PARAMETER` response carrying
`parameter`, `hint`, and `issue` fields.

### Non-breaking

- README and in-process `INSTRUCTIONS` prompt updated to document the
  v0.3.0 response shape and structured-param convention.
- Test suite updated from 638 → 652 passing tests (1 skipped).
  New coverage: `rows=None` rejection, `detailed=True` single-decode
  regression, result-always-dict assertions across all primitive tools,
  structured-param type-rejection tests for every converted tool.

### Refactor (no behaviour change; closes #101, #102)

- `INSTRUCTIONS` trimmed from 178 → 50 lines. Removed McKinsey-specific
  layout rules, data density guidelines, color palette prose, and slide /
  chart / icon / callout worked examples. Kept only parameter conventions,
  workflow guidance, theme enumeration, response-shape docs, and
  auto-render env vars. UX / styling belongs in the calling agent's
  system prompt.
- Structured-response helpers (`_success`, `_error`, `_err`,
  `_success_with_render`) extracted into
  `pptx_mcp_server._envelope` (module-private).
- Auto-render gate (`_auto_render_enabled`, `_auto_render_timeout`,
  timeout-enforced runner) extracted into `pptx_mcp_server._render`.
  `server.py` keeps a thin `_auto_render` shim that injects its own
  `render_slide` binding so `monkeypatch.setattr(server, "render_slide",
  …)` tests stay green.
- README tool count corrected (25 → 37); README tables now match
  `@mcp.tool()` registrations 1:1.

## v0.2.0 — Beta, structured responses, opt-in auto-render

See prior issues #86 (auto-render opt-in) and #88 (structured `{ok,
result|error}` envelope) for the earlier breaking changes.
