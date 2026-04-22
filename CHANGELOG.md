# Changelog

All notable changes to `pptx-mcp-server` are documented in this file.

## Unreleased / v0.6.0 — shadcn-for-PPTX block-component layer

The v0.6.0 foundation: a container-scoped block-component layer that
composes atomic primitives inside bounded rectangles and validates
containment. Closes #130, #131, #132, #133, #134, #135, #136.

### Added

- **Container primitive (#130)** —
  `engine.components.container.begin_container` context manager that
  declares a bounded rectangle for a block component on a slide. Shapes
  added inside the `with` block via the atomic primitives
  (`_add_shape` / `_add_textbox` / `_add_image` / `add_auto_fit_textbox`)
  are auto-tagged against the innermost active container via a
  thread-local stack. Supports nested containers and per-container
  `padding`. New `ContainerBounds` dataclass exposes `inner_bounds()`
  after padding. New `check_containment(presentation, *, tolerance=0.01)`
  validator flags any tagged child whose bbox exits its container's
  padded bounds (`category="shape_outside_container"`,
  `severity="error"`). `check_deck_extended` now emits a `containment`
  key on each slide and counts findings toward `summary.errors`.
- **Theme-aware atomic primitives (#131)** — `add_auto_fit_textbox` and
  `_add_shape` (plus their MCP wrappers) accept an optional
  `theme: str | None` kwarg. Color fields resolve theme tokens
  (`"primary"`, `"text_secondary"`, `"rule_subtle"`, etc.) through the
  central `theme.resolve_theme_color` helper. Raw hex still works and
  `theme=None` is fully backwards-compatible.
- **MetricCard block component (#132)** —
  `pptx_add_metric_card_row` MCP tool + `engine.components.metric_card`
  (`MetricCardSpec`, `MetricEntry`, `add_metric_card`,
  `add_metric_card_row`). N bounded cards side-by-side, each stacking
  label / title / chart slot (image or placeholder) / optional
  metrics row. Theme-aware colors, configurable padding, and
  auto-tagging into the active container.
- **KPIRow block component (#133)** — block-component
  `pptx_add_kpi_row` MCP tool with the new label / big-value /
  optional-detail layout. Each cell stacks a 9pt label, a 26pt bold
  single-line auto-fit value, and an optional 8pt detail line. Cells
  are evenly distributed across `width` with `gap` inches between them.
  Supports optional `card_fill` / `card_border` frames (e.g.
  `card_fill="highlight_row"`, `card_border="rule_subtle"`) and a
  `theme` kwarg for theme-token resolution. Auto-tags child shapes into
  the active container. New `engine.components.kpi_row.KPISpec`
  dataclass + `add_kpi_row_block` engine export (aliased to avoid a
  name collision with the legacy `engine.composites.add_kpi_row`).
- **NumberedList block component (#134)** —
  `pptx_add_numbered_list` MCP tool +
  `engine.components.numbered_list` (`NumberedItem`,
  `add_numbered_list`). N numbered items stacked vertically inside a
  bounded rectangle, each rendering number + caption row, bold title,
  and wrapped body paragraph. Optional horizontal rules between items
  (`rule_between=True` by default).
- **PageMarker + SlideFooter block components (#135)** —
  `pptx_add_page_marker` (top-right "section / P.XX" marker) and
  `pptx_add_slide_footer` (bottom-edge left/right footer text) MCP
  tools. Both use fixed conventional position and accept a `theme`
  kwarg for default `text_secondary` color-token resolution. Engine
  specs `PageMarkerSpec` / `SlideFooterSpec` live in
  `engine.components.markers`.
- **SectionHeader block component (#136)** —
  `pptx_add_section_header` MCP tool +
  `engine.components.section_header` (`SectionHeaderSpec`,
  `add_section_header`). Title + optional subtitle + full-width
  divider rule. Title/subtitle use single-line auto-fit with ellipsis
  truncation. Returns `title_bounds`, `subtitle_bounds`,
  `divider_bounds`, and `consumed_height` for downstream placement.

### Changed (breaking — MCP surface)

- **Legacy `pptx_add_kpi_row` MCP tool renamed to
  `pptx_add_kpi_row_legacy`** — callers of the 4-arg signature
  `(file_path, slide_index, kpis, y)` will fail against the new 7+ arg
  schema that now owns the canonical name. **The legacy tool is
  deprecated and will be removed in v0.7.0.** New callers should
  migrate to the block-component `pptx_add_kpi_row`.
- **MCP tool tiering (#137)** — the MCP registry is now split into a
  focused **default surface** (~20 tools: block components, high-level
  primitives, batch build, validation / rendering, setup / inspection)
  and an **advanced tier** (~25 tools: raw shape primitives, low-level
  edit ops, slide-level editing, table editing primitives,
  connectors / callouts / icons, composite helpers,
  `pptx_add_kpi_row_legacy`). Advanced-tier tools are **hidden from
  agents by default**; opt in via `PPTX_MCP_INCLUDE_ADVANCED=1`
  (accepts `1` / `true` / `yes`). The advanced functions remain
  importable for library-mode use (`from pptx_mcp_server.server
  import pptx_add_textbox`) — only the FastMCP registration is
  gated. Implementation lives in
  `pptx_mcp_server._tool_registry.mcp_tool(mcp, advanced=...)`.

### Migration

```python
# Before (v0.5.x legacy):
pptx_add_kpi_row(file_path, slide_index, kpis=[...], y=2.5)

# After (v0.6.0, block component):
pptx_add_kpi_row(file_path, slide_index, kpis=[...], left=0.5, top=2.5, width=12.0)

# Or, to keep the old behavior unchanged for now (removed in v0.7.0):
pptx_add_kpi_row_legacy(file_path, slide_index, kpis=[...], y=2.5)
```

### Non-breaking

- `engine.composites.add_kpi_row` (library-level) is unchanged; only
  the MCP tool name moved.
- The legacy tool's runtime behavior is identical — the rename is
  surface-only so migration is a one-line rename, not a refactor.
- Atomic primitive signatures are unchanged; container tagging is
  purely additive (no-op when no container is active). Registry is
  validator-time metadata only — never serialized into the PPTX XML.

## v0.5.1 — timeline theme default

Flagged by Codex gpt-5.4 v0.5.0 review: the timeline tool
(`pptx_add_milestone_timeline` + engine `add_milestone_timeline`) still
defaulted to `theme="mckinsey"` instead of `None`. Omitting `theme` was
not neutral — passing an unknown token without explicit theme silently
resolved via mckinsey.

### Fix

- `add_milestone_timeline`: `theme: str = "mckinsey"` → `Optional[str] = None`
- `pptx_add_milestone_timeline`: same
- Both now consistent with `add_responsive_card_row` and `add_data_table`
- Regression test pins the contract

713 → 714 tests passing.

## v0.5.0 — theme-aware primitives

Closes the ergonomic gaps Codex gpt-5.4 flagged in the v0.4.0 review:
primitives now resolve theme tokens (`"rule_subtle"`, `"primary"`, etc.)
instead of only accepting raw hex. Closes #123, #124, #125.

### New

- **`theme.resolve_theme_color(token_or_hex, theme_name)`** — central
  helper that resolves theme tokens to 6-hex (without `#` prefix).
  Handles theme registry lookup, unknown-token passthrough, and empty-string
  disable signal.

### Changed

- **`add_responsive_card_row`** now accepts a `theme: str | None = None`
  kwarg. All `CardSpec` color fields (`fill_color`, `accent_color`,
  `border_color`, `title_color`, `body_color`, `label_color`) resolve
  theme tokens. Cards are copied via `dataclasses.replace` — caller
  input is not mutated.
- **`add_data_table`** accepts `theme` kwarg. `alt_row_color`,
  `highlight_color`, `rule_color`, and per-column `value_color` /
  `header_color` all resolve theme tokens.
- **`add_milestone_timeline`** now actually uses its `theme` arg.
  `phase_rule_color`, `milestone_rule_color`, and milestone style colors
  (primary/secondary) all resolve through the theme.

### Non-breaking

- Raw hex still works as before (passthrough with `#` stripped)
- `theme=None` is the default — no behavior change for existing callers
- MCP tools `pptx_add_responsive_card_row`, `pptx_add_data_table`,
  `pptx_add_milestone_timeline` each gain an optional `theme` param

### Tests

713 passing (up from 703). +10 regression tests across
`test_theme.py`, `test_cards.py`, `test_timeline.py`.

## v0.4.1 — IR primitives polish

Flagged by Codex gpt-5.4 v0.4.0 review. Four surgical fixes to the IR
primitives shipped in v0.4.0.

### Fixes

- **#119** — `add_milestone_timeline` now anchors rules at the declared
  `chart_top`, not `top + phase_band_height`. The param was being validated
  but not used. Contract now matches documentation.
- **#120** — `pptx_add_data_table` and `pptx_add_milestone_timeline` now
  include a `message` key in their result, matching the v0.3.0 envelope
  invariant that `result` is always a dict with at least `message`.
- **#121** — `add_data_table` rejects `rule_thickness < 0` with
  `INVALID_PARAMETER` instead of silently producing negative-height rule
  rectangles.
- **#122** — `INSTRUCTIONS` themes enumeration now lists all 4 themes
  (`mckinsey`, `deloitte`, `neutral`, `ir`).

700 → 703 tests passing.

## v0.3.1 — close OpenAI/Codex 5.4 remaining blockers

Flagged by the OpenAI/Codex 5.4 v0.3.0 review as the remaining partials that
prevent LOGO-READY. Closes #105, #106, #107, #108.

### Fixes

#### #105 — Strict nested-key validation + legacy string-JSON fallback removal

The v0.3.0 surface change (#97) converted `*_json: str` parameters to
structured types but left several dict-shaped specs without unknown-key
rejection. Adversarial input like `_validate_chart_data('column', ['A'],
['oops'])` surfaced as `AttributeError` → `INTERNAL_ERROR` rather than a
structured `INVALID_PARAMETER`.

v0.3.1 adds `frozenset`-backed allowed-keys validation for every
remaining dict-shaped spec:

- `pptx_add_chart` → `_validate_chart_spec` + per-series key validation
- `pptx_build_slide` / `pptx_build_deck` → `_validate_slide_spec`
- `pptx_add_kpi_row` → `_validate_kpi_spec`
- `pptx_edit_table_cells` → `_validate_edit_cells_spec`

Unknown keys raise `EngineError(INVALID_PARAMETER, "<spec>: unknown keys
[...]; allowed: [...]")`. Wrong types on critical fields (e.g. `series`
not a list, `row` not int, `value` not str/number) are rejected cleanly
without leaking `AttributeError` / `TypeError`.

**BREAKING — engine layer only** (MCP layer unaffected): the legacy
`isinstance(x, str)` → `json.loads` fallbacks in `build_slide`,
`build_deck`, `add_kpi_row`, `add_bullet_block` engine wrappers were
removed. Library consumers must now pass structured values. The MCP
boundary already passed dicts since v0.3.0, so MCP clients are unaffected.

#### #106 — `message` field in detailed layout result

`pptx_check_layout(detailed=True)` now includes a human-readable
`result["message"]` alongside `slides` / `summary`, so generic agents
that read `result.message` work for every tool:

```json
{"ok": true, "result": {
    "message": "Found 3 errors, 0 warnings, 12 info findings across 22 slides",
    "slides": [...],
    "summary": {...}
}}
```

#### #107 — `FILE_NOT_FOUND` taxonomy fix for `pptx_check_layout`

`pptx_check_layout` called `Presentation(file_path)` directly, so a
missing file surfaced as `INTERNAL_ERROR: PackageNotFoundError`. It now
routes through `engine.pptx_io.open_pptx`, which maps missing files to
`EngineError(FILE_NOT_FOUND)` — consistent with every other tool.

Audit confirmed no other direct `Presentation(path)` calls remain in
`server.py`.

#### #108 — Drift cleanup

- Dropped the stale `"All 25 tools"` assertion in `tests/test_server.py`
  (actual count is 37+). The test now asserts the core tools are
  registered without pinning a hardcoded total.
- CHANGELOG already recorded `INSTRUCTIONS` at 50 lines (accurate — the
  v0.3.0 claim was never 41, contrary to some review drafts); no further
  line-count edit needed.

### Tests

Coverage extended for each new validator path:

- Chart / slide / KPI / edit-cell unknown-key rejection.
- `build_slide(prs, '{"title":"x"}')` string input now raises
  `EngineError(INVALID_PARAMETER)` — the legacy silent-parse is gone.
- `pptx_check_layout("/nonexistent.pptx")` and
  `pptx_check_layout("/nonexistent.pptx", detailed=True)` both return
  `error.code == "FILE_NOT_FOUND"`.
- `pptx_check_layout(path, detailed=True)` result contains `message`.

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
