"""Container primitive と check_containment バリデータのテスト (Issue #130).

- ``begin_container`` 内部で追加された shape が境界内なら findings 0
- 0.3" 下はみ出しで shape_outside_container を 1 件検出
- ネスト時、inner を出るが outer には収まる shape は inner 側でのみ検出
- ``padding`` で内側を絞った場合に overflow が出る
- ``tolerance`` で微小 float drift は無視される
- ``check_deck_extended`` の ``slides[i].containment`` へ発火 + summary.errors に集計

test の独立性を保つため、各テストの先頭で ``clear_container_registry`` を呼ぶ
(モジュール level dict のため、同一プロセス内テスト間で干渉しうる)。
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.engine.components.container import (
    ContainerBounds,
    begin_container,
    clear_container_registry,
)
from pptx_mcp_server.engine.shapes import (
    _add_shape,
    _add_textbox,
    add_auto_fit_textbox,
)
from pptx_mcp_server.engine.validation import (
    check_containment,
    check_deck_extended,
)


@pytest.fixture(autouse=True)
def _isolated_registry():
    """Ensure each test starts with a clean container registry."""
    clear_container_registry()
    yield
    clear_container_registry()


# ---------------------------------------------------------------------------
# 単一 container: 内側 / 外側
# ---------------------------------------------------------------------------


def test_container_child_fully_inside_has_no_finding(one_slide_prs):
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide,
        name="metric_card_0",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide, 1.2, 1.2, 2.6, 0.4, text="Title")
        _add_textbox(slide, 1.2, 1.8, 2.6, 0.8, text="Body copy within bounds")

    findings = check_containment(one_slide_prs)
    assert findings == []


def test_container_child_overflow_bottom_flagged(one_slide_prs):
    """Issue 原因となった pattern: metric_y が card bottom から 0.3" 下."""
    slide = one_slide_prs.slides[0]
    # Card: top=1.0, height=2.0 → bottom = 3.0"
    # Child textbox: top=2.8, height=0.5 → bottom = 3.3" (0.3" 超過)
    with begin_container(
        slide,
        name="metric_card_0",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide, 1.2, 2.8, 2.6, 0.5, text="42%")

    findings = check_containment(one_slide_prs)
    assert len(findings) == 1
    f = findings[0]
    assert f.category == "shape_outside_container"
    assert f.severity == "error"
    assert f.slide_index == 0
    # メッセージに overshoot 量と container name が含まれる
    assert "metric_card_0" in f.message
    assert "0.300" in f.message or "0.3" in f.message
    assert "bottom" in f.message


def test_container_child_overflow_right_flagged(one_slide_prs):
    slide = one_slide_prs.slides[0]
    # Card: left=1.0, width=3.0 → right = 4.0"
    # Child: left=3.5, width=1.0 → right = 4.5" (0.5" 超過)
    with begin_container(
        slide,
        name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_shape(slide, "rectangle", 3.5, 1.2, 1.0, 0.5)

    findings = check_containment(one_slide_prs)
    assert len(findings) == 1
    assert "right" in findings[0].message


# ---------------------------------------------------------------------------
# Nested containers
# ---------------------------------------------------------------------------


def test_nested_inner_overflow_flagged_on_inner_only(one_slide_prs):
    """inner を出るが outer には収まる child: inner 側だけで 1 件."""
    slide = one_slide_prs.slides[0]
    # Outer: [0.5, 0.5, 10.0, 6.0]  → right=10.5, bottom=6.5
    # Inner: [1.0, 1.0,  3.0, 2.0]  → right=4.0,  bottom=3.0
    # Child: left=3.8, top=1.2, w=1.0, h=0.5 → right=4.8, bottom=1.7
    #        inner right=4.0 を 0.8" 超過。outer right=10.5 の内側。
    with begin_container(
        slide, name="outer",
        left=0.5, top=0.5, width=10.0, height=6.0,
    ):
        with begin_container(
            slide, name="inner",
            left=1.0, top=1.0, width=3.0, height=2.0,
        ):
            _add_textbox(slide, 3.8, 1.2, 1.0, 0.5, text="overflow")

    findings = check_containment(one_slide_prs)
    # inner のみ flag される (outer からは child 未登録)。
    assert len(findings) == 1
    assert "inner" in findings[0].message
    assert "outer" not in findings[0].message


def test_nested_child_inside_both_has_no_finding(one_slide_prs):
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="outer",
        left=0.5, top=0.5, width=10.0, height=6.0,
    ):
        with begin_container(
            slide, name="inner",
            left=1.0, top=1.0, width=3.0, height=2.0,
        ):
            _add_textbox(slide, 1.2, 1.2, 2.6, 0.5, text="inside both")

    assert check_containment(one_slide_prs) == []


# ---------------------------------------------------------------------------
# Padding
# ---------------------------------------------------------------------------


def test_padding_shrinks_inner_bounds(one_slide_prs):
    """padding=0.1 で inner 境界が内側に寄ると、境界線上の child が flag される."""
    slide = one_slide_prs.slides[0]
    # Container: [1.0, 1.0, 3.0, 2.0], padding=0.1
    # padding 適用後 inner: [1.1, 1.1, 3.9, 2.9]
    # Child: [1.0, 1.1, 0.5, 0.5] → left=1.0 は inner_left=1.1 から 0.1" はみ出す。
    with begin_container(
        slide, name="padded",
        left=1.0, top=1.0, width=3.0, height=2.0,
        padding=0.1,
    ):
        _add_shape(slide, "rectangle", 1.0, 1.1, 0.5, 0.5)

    findings = check_containment(one_slide_prs)
    assert len(findings) == 1
    assert "left" in findings[0].message


def test_padding_zero_allows_boundary_touch(one_slide_prs):
    """padding=0 では境界線上の child は flag されない (0 exit だから)."""
    slide = one_slide_prs.slides[0]
    # Container: [1.0, 1.0, 3.0, 2.0]  → inner = outer
    # Child flush to the top-left corner: [1.0, 1.0, 0.5, 0.5]
    with begin_container(
        slide, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_shape(slide, "rectangle", 1.0, 1.0, 0.5, 0.5)

    assert check_containment(one_slide_prs) == []


# ---------------------------------------------------------------------------
# Tolerance
# ---------------------------------------------------------------------------


def test_tolerance_absorbs_small_float_drift(one_slide_prs):
    """tolerance (既定 0.01") 未満のはみ出しは flag されない."""
    slide = one_slide_prs.slides[0]
    # Container right = 4.0
    # Child right = 4.005 (0.005" 超過; tolerance=0.01 > 0.005)
    with begin_container(
        slide, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_shape(slide, "rectangle", 3.5, 1.2, 0.505, 0.5)

    assert check_containment(one_slide_prs) == []


def test_tolerance_custom_strict(one_slide_prs):
    """tolerance=0 で同じ child が flag される."""
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_shape(slide, "rectangle", 3.5, 1.2, 0.505, 0.5)

    findings = check_containment(one_slide_prs, tolerance=0.0)
    assert len(findings) == 1


# ---------------------------------------------------------------------------
# Integration with check_deck_extended
# ---------------------------------------------------------------------------


def test_check_deck_extended_includes_containment_key(one_slide_prs):
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="metric_card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide, 1.2, 2.8, 2.6, 0.5, text="42%")  # bottom overflow

    result = check_deck_extended(one_slide_prs)
    assert "slides" in result
    assert "containment" in result["slides"][0]
    # 1 件検出
    assert len(result["slides"][0]["containment"]) == 1
    # summary.errors に加算されている
    assert result["summary"]["errors"] >= 1
    # finding の structure (ValidationFinding.to_dict 由来)
    finding_dict = result["slides"][0]["containment"][0]
    assert finding_dict["category"] == "shape_outside_container"
    assert finding_dict["severity"] == "error"
    assert finding_dict["slide_index"] == 0


def test_check_deck_extended_clean_deck_has_empty_containment(one_slide_prs):
    """何も declare しなくても ``containment`` key は存在し空リストとなる."""
    result = check_deck_extended(one_slide_prs)
    assert result["slides"][0]["containment"] == []


def test_auto_fit_textbox_is_auto_tagged(one_slide_prs):
    """add_auto_fit_textbox 経由の shape も containment チェックの対象."""
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        # top=2.9, height=0.3 → bottom=3.2 (card bottom=3.0 を 0.2" 超過)
        add_auto_fit_textbox(
            slide, "overflow!", 1.2, 2.9, 2.6, 0.3,
            font_size_pt=10, min_size_pt=8,
        )

    findings = check_containment(one_slide_prs)
    assert len(findings) == 1
    assert "bottom" in findings[0].message


# ---------------------------------------------------------------------------
# API guarantees
# ---------------------------------------------------------------------------


def test_container_bounds_dataclass_fields():
    b = ContainerBounds(name="x", left=0.0, top=0.0, width=1.0, height=1.0)
    assert b.padding == 0.0
    assert b.children == []
    assert b.inner_bounds() == (0.0, 0.0, 1.0, 1.0)


def test_begin_container_yields_bounds(one_slide_prs):
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0, padding=0.05,
    ) as bounds:
        assert isinstance(bounds, ContainerBounds)
        assert bounds.name == "card"
        assert bounds.padding == 0.05
        inner = bounds.inner_bounds()
        assert inner == pytest.approx((1.05, 1.05, 3.95, 2.95))


def test_no_container_active_is_noop(one_slide_prs):
    """Shapes added outside any begin_container are not tagged."""
    slide = one_slide_prs.slides[0]
    _add_textbox(slide, 1.0, 1.0, 2.0, 0.5, text="naked")
    # No container declared → check_containment sees nothing.
    assert check_containment(one_slide_prs) == []


def test_check_containment_consumes_registry_no_phantom_findings(
    blank_prs,
):
    """Second build of the deck must not inherit stale container bounds.

    Regression for the id(slide)-reuse + memory-leak hazard: check_containment
    should clear its per-slide registry entries as it goes, so a later deck
    built in the same process cannot receive phantom shape_outside_container
    findings from a previous run.
    """
    from pptx_mcp_server.engine.components.container import (
        _SLIDE_REGISTRY,
        iter_slide_containers,
    )

    # --- First deck: overflow on purpose, validate, confirm cleanup. ---
    layout = blank_prs.slide_layouts[6]
    slide1 = blank_prs.slides.add_slide(layout)
    with begin_container(
        slide1, name="card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide1, 1.2, 2.8, 2.6, 0.5, text="overflow!")

    findings1 = check_containment(blank_prs)
    assert len(findings1) == 1
    # After validation, the registry must not retain entries for this slide.
    assert list(iter_slide_containers(slide1)) == []
    assert id(slide1) not in _SLIDE_REGISTRY

    # --- Second pass (no new declarations): zero phantom findings. ---
    findings2 = check_containment(blank_prs)
    assert findings2 == []

    # --- Second deck built in the same process: declare a CLEAN container.
    # If registry cleanup regressed, a stale list for id(slide1) (or a reused
    # id) could leak overflow children into this new pass.
    slide2 = blank_prs.slides.add_slide(layout)
    with begin_container(
        slide2, name="card2",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide2, 1.2, 1.2, 2.0, 0.5, text="inside")

    findings3 = check_containment(blank_prs)
    assert findings3 == []


def test_check_deck_extended_summary_errors_match_placed_containment(
    one_slide_prs,
):
    """summary.errors must count only containment findings that were placed.

    Regression for the double-count hazard where a finding with an
    out-of-range slide_index was dropped by _place but still counted in
    summary.errors — producing a summary that didn't match the per-slide
    lists.
    """
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="metric_card",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide, 1.2, 2.8, 2.6, 0.5, text="overflow!")  # bottom overflow

    result = check_deck_extended(one_slide_prs)

    placed_containment_errors = sum(
        1
        for slide_data in result["slides"]
        for f in slide_data["containment"]
        if f.get("severity") == "error"
    )
    placed_overlaps = sum(len(s["overlaps"]) for s in result["slides"])
    placed_oob = sum(len(s["out_of_bounds"]) for s in result["slides"])

    # Per-slide containment list should have exactly 1 entry (normal path).
    assert placed_containment_errors == 1
    # summary.errors should equal the sum of all placed error-severity items
    # — i.e. never larger than what the caller can see in the per-slide view.
    assert result["summary"]["errors"] == (
        placed_containment_errors + placed_overlaps + placed_oob
    )


def test_stack_cleaned_after_context_exit(one_slide_prs):
    """begin_container exit 後に innermost 判定が外れる (stack pop 確認)."""
    slide = one_slide_prs.slides[0]
    with begin_container(
        slide, name="c1",
        left=1.0, top=1.0, width=3.0, height=2.0,
    ):
        _add_textbox(slide, 1.2, 1.2, 2.0, 0.5, text="in")
    # context 終了後の add は tag されない
    _add_textbox(slide, 0.0, 0.0, 0.1, 0.1, text="out")

    # c1 には child が 1 つだけ登録され、"out" は含まれない
    from pptx_mcp_server.engine.components.container import iter_slide_containers

    containers = list(iter_slide_containers(slide))
    assert len(containers) == 1
    assert len(containers[0].children) == 1
