# Contributing to pptx-mcp-server

Thanks for considering a contribution. This document codifies the conventions
that keep the engine/MCP split clean and the test suite meaningful.

## Dual-API pattern

Every new primitive exposes two layered forms:

- **In-memory** — `verb_noun(slide, ...)` operates on an already-open
  `python-pptx` `slide` object. This is the canonical form: composable,
  side-effect-free w.r.t. the filesystem, and fast to test.
- **File-based** — `verb_noun_file(path, slide_index, ...)` is a thin
  open → call in-memory → save wrapper around the in-memory primitive.
  It exists so MCP tools (`pptx_verb_noun`) can dispatch without having
  to know about `Presentation` lifetimes.

Rule of thumb: write the in-memory form first, get it tested, then add the
file-based wrapper. Never put behavior in the file-based wrapper that is not
also exercised through the in-memory form — the in-memory form must be the
source of truth.

## Naming conventions

| Layer | Pattern | Example |
|------|---------|---------|
| MCP tool | `pptx_verb_noun` | `pptx_add_textbox` |
| File-based wrapper | `verb_noun_file` | `add_textbox_file` |
| In-memory primitive | `verb_noun` (no suffix) | `add_textbox` |
| JSON input via MCP | `_json` suffix on the arg | `items_json` |
| Font size argument | `_pt` suffix | `font_size_pt` |
| Colors | 6-hex, no `#` | `"051C2C"` |
| Coordinates / dimensions | inches (float) | `left=1.0, width=10.0` |

Stick to these so the dispatch layer stays mechanical. A new primitive that
breaks naming forces every downstream agent using the MCP surface to special-
case it.

## Testing conventions

- Prefer in-memory tests using the `slide` fixture in `tests/conftest.py`.
  These don't touch the filesystem and run in milliseconds.
- File-based tests are for when the I/O itself is under test (e.g., round-
  tripping through `open_pptx` + `save`). Don't use them to test behavior
  that the in-memory primitive already covers.
- Test file naming mirrors the engine: `engine/shapes.py` → `tests/test_shapes.py`,
  `engine/text_metrics.py` → `tests/test_text_metrics.py`.
- Run the full suite (`python -m pytest`) before committing. Baseline is 440+
  tests and is kept green on `main`.

## Calibration

The `text_metrics` width heuristic is calibrated for Arial (ASCII) and CJK
full-width. `tests/test_calibration.py` measures real font advance widths via
fontTools and guards against order-of-magnitude calibration errors. These
tests skip when no compatible TTF is available locally. The `calibration`
CI job runs them against Liberation Sans + Noto Sans CJK on Ubuntu. To run
locally on macOS, system Arial + Hiragino Sans are auto-detected. On Ubuntu:

```bash
sudo apt install fonts-liberation fonts-noto-cjk
python -m pytest tests/test_calibration.py -v
```

Tolerance bands are deliberately loose (±35% per ASCII char, ±25% ASCII
mixed string, ±20% CJK char, ±15% CJK mixed string): the heuristic is a
±10-15% model, and the calibration suite exists to catch order-of-magnitude
miscalibrations, not hide them behind tight bands. If you change a width
constant, re-run locally and update bands only if they reflect a genuine
improvement; do NOT loosen bands to paper over a new miscalibration.

## Engine / MCP boundary

`src/pptx_mcp_server/engine/*.py` MUST NOT import `mcp`. Library consumers
can install without the `[mcp]` extra (see issue #36 / PR #53), and the
engine must remain fully functional in that configuration.

The guardrail is enforced at the AST level by
`tests/test_library_usage.py::test_engine_has_no_mcp_imports`. The
packaging side is enforced by `tests/test_packaging_extras.py`. Both are
part of the default test run and also run in the library-only CI job
(`.github/workflows/ci.yml`) without the `[mcp]` extra installed.

If you need MCP-specific behavior, put it in `src/pptx_mcp_server/server.py`
or a sibling module that is only imported when the CLI starts.

## File safety

`save_pptx` is atomic (temp-file-then-rename) but does NOT guard against
concurrent writes from multiple processes. The single-writer contract is:
at most one process mutates a given `.pptx` path at a time. Consumers that
need multi-writer coordination should layer their own locking on top.

## Pull request flow

1. Branch off `main` with a short, topic-style name
   (`fix-docs-wave4-bundle`, `feat-flex-container`, etc.).
2. Run `python -m pytest` locally — all tests must pass before you push.
3. Reference issues in the commit message. Preferred shape:

   ```
   fix(module): short summary (closes #N)
   ```

   For multi-issue PRs, list each `closes #N`. Use a HEREDOC to pass the
   message so newlines land correctly:

   ```bash
   git commit -m "$(cat <<'EOF'
   fix(module): short summary (closes #N)

   Longer body explaining the why, not the what.
   EOF
   )"
   ```

4. Avoid trailing "what I did" summaries. The diff speaks; the body should
   cover motivation and trade-offs.
5. Do not open PRs that skip tests (no `--no-verify`, no disabled pytest
   markers without a linked follow-up issue).

## Code style

- `from __future__ import annotations` at the top of every module.
- Type hints on all public functions.
- Docstrings in Japanese, `だ・である調` for statements. `です`・`ます` は
  使わない (the Japanese style-guide rule is: use plain-form declaratives,
  not polite-form).
- No emojis in code or docs.
- No "fixes" or "I did X" trailers in commit messages — let the diff speak.

## Sub-issue / wave structure

Larger changes are occasionally split into numbered waves
(W1, W2, …) with per-issue sub-tasks. When that happens, each sub-issue
should be landable in isolation; a PR that requires a sister PR to pass
tests is a sign the split is wrong.
