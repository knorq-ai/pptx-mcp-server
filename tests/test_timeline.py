"""``add_milestone_timeline`` プリミティブのテスト.

IR 成長ストーリースライド向けに、フェーズ帯 + 垂直ルール + マイルストーン
注記を既存チャート領域に重ねるプリミティブを検証する。shape 追加数・
位置・入力検証の各挙動を網羅する。
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.engine.pptx_io import EngineError, ErrorCode
from pptx_mcp_server.engine.timeline import (
    TimelineMilestone,
    TimelinePhase,
    add_milestone_timeline,
)


# EMU → inches 換算 (1 inch = 914400 EMU)
def _in(emu: int) -> float:
    return emu / 914400


def test_three_phases_four_milestones_shape_counts(slide):
    """3 フェーズ + 4 マイルストーンで各 shape 数が仕様通りか検証する.

    期待値:
      - phase_shapes: 3 × 3 (index/label/year) = 9
      - rule_shapes:  (3-1) phase rules + 4 milestone rules = 6
      - milestone_shapes: 4
    """
    phases = [
        TimelinePhase(label="事業基盤の確立", index_label="01", year_range="2012 — 2016"),
        TimelinePhase(label="事業拡大", index_label="02", year_range="2017 — 2020"),
        TimelinePhase(label="プラットフォーム化", index_label="03", year_range="2021 — 現在"),
    ]
    milestones = [
        TimelineMilestone(x_pos=0.20, year="2016", label="IPO\n東証マザーズ上場"),
        TimelineMilestone(x_pos=0.45, year="2019", label="BillFlow 提供開始", style="secondary"),
        TimelineMilestone(x_pos=0.70, year="2022", label="海外展開"),
        TimelineMilestone(x_pos=0.90, year="2024", label="東証プライム上場"),
    ]
    result = add_milestone_timeline(
        slide,
        phases,
        milestones,
        left=0.9, top=1.15, width=11.533,
        phase_band_height=0.9,
        chart_top=2.05, chart_bottom=6.5,
    )

    assert len(result["phase_shapes"]) == 9
    assert len(result["rule_shapes"]) == 6
    assert len(result["milestone_shapes"]) == 4
    assert result["consumed_height"] == pytest.approx(0.9)


def test_empty_phases_with_milestones(slide):
    """フェーズなし・マイルストーン 2 つのとき、フェーズ関連 shape は 0 のまま.

    ``chart_top``/``chart_bottom`` とマイルストーン・ルールのみ描画される。
    """
    milestones = [
        TimelineMilestone(x_pos=0.3, year="2016", label="A"),
        TimelineMilestone(x_pos=0.7, year="2020", label="B"),
    ]
    result = add_milestone_timeline(
        slide,
        [],
        milestones,
        left=1.0, top=1.0, width=10.0,
        phase_band_height=0.9,
        chart_top=2.0, chart_bottom=6.5,
    )

    assert result["phase_shapes"] == []
    # phase rules 0 (境界なし) + milestone rules 2
    assert len(result["rule_shapes"]) == 2
    for rs in result["rule_shapes"]:
        assert rs["kind"] == "milestone_rule"
    assert len(result["milestone_shapes"]) == 2


def test_phases_only_no_milestones(slide):
    """3 フェーズ・マイルストーン 0 件のとき、phase rules のみ描画される.

    期待値:
      - phase_shapes: 9
      - rule_shapes:  2 (inter-phase)
      - milestone_shapes: 0
    """
    phases = [
        TimelinePhase(label="P1", index_label="01", year_range="Y1"),
        TimelinePhase(label="P2", index_label="02", year_range="Y2"),
        TimelinePhase(label="P3", index_label="03", year_range="Y3"),
    ]
    result = add_milestone_timeline(
        slide,
        phases,
        [],
        left=0.9, top=1.15, width=11.0,
        chart_top=2.05, chart_bottom=6.5,
    )

    assert len(result["phase_shapes"]) == 9
    assert len(result["rule_shapes"]) == 2
    for rs in result["rule_shapes"]:
        assert rs["kind"] == "phase_rule"
    assert result["milestone_shapes"] == []


def test_invalid_x_pos_raises(slide):
    """``x_pos`` が ``[0.0, 1.0]`` を外れたら ``INVALID_PARAMETER``."""
    with pytest.raises(EngineError) as excinfo:
        add_milestone_timeline(
            slide,
            [],
            [TimelineMilestone(x_pos=-0.1, year="2016", label="X")],
            left=1.0, top=1.0, width=10.0,
            chart_top=2.0, chart_bottom=6.5,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER
    assert "x_pos" in str(excinfo.value)


def test_invalid_chart_range_raises(slide):
    """``chart_top >= chart_bottom`` なら ``INVALID_PARAMETER``."""
    with pytest.raises(EngineError) as excinfo:
        add_milestone_timeline(
            slide,
            [],
            [],
            left=1.0, top=1.0, width=10.0,
            chart_top=5.0, chart_bottom=3.0,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER
    msg = str(excinfo.value)
    assert "chart_top" in msg and "chart_bottom" in msg


def test_single_phase_no_inter_phase_rule(slide):
    """フェーズが 1 つのとき inter-phase 境界はなく、phase rules は 0 本."""
    phases = [TimelinePhase(label="Only", index_label="01", year_range="Y")]
    result = add_milestone_timeline(
        slide,
        phases,
        [],
        left=1.0, top=1.0, width=10.0,
        chart_top=2.0, chart_bottom=6.5,
    )

    # 1 フェーズ分 3 shape
    assert len(result["phase_shapes"]) == 3
    # inter-phase rule は 0
    assert result["rule_shapes"] == []


def test_negative_width_raises(slide):
    """幅が負のとき ``INVALID_PARAMETER``."""
    with pytest.raises(EngineError) as excinfo:
        add_milestone_timeline(
            slide,
            [],
            [],
            left=0.0, top=1.0, width=-1.0,
            chart_top=2.0, chart_bottom=6.5,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER


def test_milestone_x_pos_positioning(slide):
    """マイルストーン・ルールの x 座標が ``left + x_pos * width`` と一致する."""
    milestone = TimelineMilestone(x_pos=0.5, year="2020", label="Mid")
    result = add_milestone_timeline(
        slide,
        [],
        [milestone],
        left=2.0, top=1.0, width=10.0,
        chart_top=2.0, chart_bottom=6.5,
    )

    rule = result["rule_shapes"][0]
    shape = slide.shapes[rule["shape_index"]]
    # left + 0.5 * width = 7.0; rule は 0.01" 幅で中央揃え → shape.left = 7.0 - 0.005
    assert _in(shape.left) == pytest.approx(7.0 - 0.005, abs=1e-4)


def test_secondary_style_uses_gray(slide):
    """``style='secondary'`` のマイルストーンラベルはグレー文字色で描画される."""
    milestone = TimelineMilestone(
        x_pos=0.5, year="2020", label="Minor event", style="secondary",
    )
    result = add_milestone_timeline(
        slide,
        [],
        [milestone],
        left=1.0, top=1.0, width=10.0,
        chart_top=2.0, chart_bottom=6.5,
    )

    label = result["milestone_shapes"][0]
    shape = slide.shapes[label["shape_index"]]
    # 文字色は最初の run の rgb で確認する。
    run = shape.text_frame.paragraphs[0].runs[0]
    rgb = run.font.color.rgb
    assert str(rgb) == "666666"


# ---------------------------------------------------------------------------
# Issue #119 — chart_top regression
# ---------------------------------------------------------------------------


def test_milestone_rules_anchor_at_chart_top(one_slide_prs: Presentation) -> None:
    """chart_top (not top+phase_band_height) must be the rule anchor (#119).

    Before the fix: milestone rules always started at `top + phase_band_height`,
    ignoring chart_top. This test pins the corrected contract.
    """
    slide = one_slide_prs.slides[0]

    result = add_milestone_timeline(
        slide,
        phases=[],
        milestones=[
            TimelineMilestone(x_pos=0.5, year="2020", label="Test"),
        ],
        left=1.0,
        top=1.0,
        width=10.0,
        phase_band_height=0.5,   # top + phase_band_height = 1.5
        chart_top=3.0,           # chart explicitly starts at y=3.0
        chart_bottom=6.0,
    )

    emu_per_in = 914400
    milestone_rules = [r for r in result["rule_shapes"] if r.get("kind") == "milestone_rule"]
    assert milestone_rules, "expected at least one milestone rule shape"
    for rule in milestone_rules:
        shape = slide.shapes[rule["shape_index"]]
        y_in = shape.top / emu_per_in
        assert abs(y_in - 3.0) < 0.01, (
            f"Milestone rule top={y_in:.3f}, expected 3.0 (chart_top)"
        )


# ---------------------------------------------------------------------------
# #125: theme-aware rendering
# ---------------------------------------------------------------------------


def test_phase_rule_color_resolves_ir_theme_token(one_slide_prs):
    """phase_rule_color="rule_subtle" with theme="ir" → resolves to theme's value (#125)."""
    from pptx_mcp_server.engine.timeline import add_milestone_timeline, TimelinePhase
    slide = one_slide_prs.slides[0]
    result = add_milestone_timeline(
        slide,
        phases=[
            TimelinePhase(label="A", index_label="01", year_range="2012-2016"),
            TimelinePhase(label="B", index_label="02", year_range="2017-2021"),
        ],
        milestones=[],
        left=1.0, top=1.0, width=10.0,
        phase_band_height=0.5,
        chart_top=2.0, chart_bottom=5.0,
        phase_rule_color="rule_subtle",  # theme token, not hex
        theme="ir",
    )
    # One phase boundary rule between A and B
    phase_rules = [r for r in result["rule_shapes"] if r.get("kind") == "phase_rule"]
    assert phase_rules, "expected phase boundary rule"
    shape = slide.shapes[phase_rules[0]["shape_index"]]
    rgb = shape.fill.fore_color.rgb
    # IR theme's rule_subtle = #E0E0E0
    assert str(rgb) == "E0E0E0"


def test_milestone_primary_style_uses_theme_primary_color_on_label(one_slide_prs):
    """style='primary' with theme='ir' → label text color = ir theme primary (#125)."""
    from pptx_mcp_server.engine.timeline import add_milestone_timeline, TimelineMilestone
    slide = one_slide_prs.slides[0]
    result = add_milestone_timeline(
        slide,
        phases=[],
        milestones=[TimelineMilestone(x_pos=0.5, year="2020", label="Test", style="primary")],
        left=1.0, top=1.0, width=10.0,
        phase_band_height=0.5,
        chart_top=2.0, chart_bottom=5.0,
        theme="ir",
    )
    label_info = result["milestone_shapes"][0]
    label = slide.shapes[label_info["shape_index"]]
    run = label.text_frame.paragraphs[0].runs[0]
    # IR theme primary = #0A2540
    assert str(run.font.color.rgb) == "0A2540"
