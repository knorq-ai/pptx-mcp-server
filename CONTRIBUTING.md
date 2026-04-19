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

### Adding a new script

Only ASCII/Latin-1 + CJK are currently calibrated. Hangul, Arabic, Thai,
Devanagari, Hebrew, etc. fall back to ASCII-normal width and will
silently mis-layout (see README "Supported scripts" for the matrix). If
you need one of these, follow this process so the new script is
permanently covered by the calibration suite rather than bolted on ad hoc:

1. **Add the Unicode range(s)** — extend `_CJK_RANGES` in
   `src/pptx_mcp_server/engine/text_metrics.py`, or introduce a new script
   set alongside it if the script does not share CJK width characteristics.
   For non-em-width scripts, also add a predicate like `is_half_width_kana`.
2. **Add or calibrate a width constant** — e.g.
   `_HANGUL_WIDTH_PER_PT: float = ...`. Derive the value by running
   `scripts/calibrate_ascii.py` (or a script modelled on it) against a real
   system font for that script and taking the harmonic-mean representative
   across the bucket, not a hand-picked value.
3. **Extend `tests/test_calibration.py`** — add parametrize entries with
   sentinel characters from the new script and a tolerance band
   consistent with the rest of the suite.
4. **Add font path candidates** — `tests/calibration_helpers.py` enumerates
   system-font probe paths. Add the new script's font (e.g. Noto Sans KR,
   Noto Naskh Arabic) to both the macOS and Linux probe lists so the
   test skips cleanly when the font is missing and runs when it is present.
5. **Document in the README matrix** — move the script row from
   "Unsupported" to "Supported" and state the accuracy band. Update the
   module docstring of `text_metrics.py` accordingly.

PRs adding a new script without all five steps will be asked to split:
the calibration constant is meaningless without the calibration test, and
the test is incomplete without matching font-probe plumbing.

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

## Validation

`check_deck_extended` / `check_text_overflow` have two paths for overflow
detection:

- **`font_source="heuristic"`** (default) — zero-deps, uses the
  `text_metrics` width heuristic. Fast, good for iteration, but shares its
  model with `add_auto_fit_textbox`. If the heuristic drifts from real
  PowerPoint font metrics (e.g. Yu Gothic advance tweaks, font substitution
  policy changes), **both the auto-fit primitive and the validator drift
  together** — a silent echo chamber.
- **`font_source="real"`** (opt-in, `[validation]` extra) — measures
  advance widths directly from TTF/TTC via fontTools, independent of the
  heuristic. Requires `pip install pptx-mcp-server[validation]` and
  `font_paths={font_name: ttf_path, ...}`. For zero-config probing use
  `discover_system_fonts()`:

  ```python
  from pptx_mcp_server.engine.font_metrics import discover_system_fonts
  from pptx_mcp_server.engine.validation import check_deck_extended

  report = check_deck_extended(
      prs, font_source="real", font_paths=discover_system_fonts()
  )
  ```

When to use real vs heuristic:

- CI / pre-commit quick checks → heuristic (fast, no font install required).
- Before shipping a deck to production → real (catches drift the heuristic
  would miss).
- Debugging "PowerPoint clips but our validator says fine" → real path
  likely disagrees with heuristic; that disagreement *is* the signal.

Fonts that can't be resolved fall back to the heuristic per-paragraph and
emit a `font_not_measured` warning so partial coverage is still useful.

## File safety

`save_pptx` performs an atomic temp-file-then-rename (via `os.replace`), so
a crash during save preserves the original file. This is atomicity, not
durability: the new file's contents may still be in page cache when the call
returns. For crash-durable saves, pass `fsync=True` — this adds an I/O
barrier on the temp file and the containing directory before returning.

The default is `fsync=False` because most deck-generation workflows don't
need power-loss durability (the process can re-run). Enable fsync when:
- Saving a long-running batch result that would be expensive to regenerate
- Writing to a mount that buffers aggressively
- Required by a compliance/archival policy

Concurrency contract: at most one process mutates a given .pptx path at a
time. Cross-process locking is not provided.

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
