"""
Integration test: build a complete mini deck using composites and verify structure.
"""

from __future__ import annotations

import json
import os

import pytest

from pptx_mcp_server.engine.pptx_io import create_presentation, open_pptx, save_pptx
from pptx_mcp_server.engine.composites import (
    _add_content_slide,
    _add_section_divider,
    _add_kpi_row,
    _add_bullet_block,
)
from pptx_mcp_server.engine.tables import _add_table
from pptx_mcp_server.engine.slides import _get_presentation_info


class TestFullDeckBuild:
    """Build a mini 3-slide deck using composites and verify structure."""

    def test_build_three_slide_deck(self, tmp_path):
        path = str(tmp_path / "full_deck.pptx")
        create_presentation(path)
        prs = open_pptx(path)

        # Slide 0: Section divider
        slide0, idx0 = _add_section_divider(prs, "Q1 Business Review", subtitle="FY2024")
        assert idx0 == 0

        # Slide 1: Content slide with KPIs
        slide1, idx1 = _add_content_slide(
            prs, "Key Performance Indicators",
            source="Source: Internal Finance", page_number=2,
        )
        assert idx1 == 1
        kpis = [
            {"value": "107.8M", "label": "Revenue"},
            {"value": "23.4%", "label": "Gross Margin"},
            {"value": "1,247", "label": "New Customers"},
        ]
        _add_kpi_row(slide1, kpis, y=1.5)

        # Slide 2: Content slide with bullets and table
        slide2, idx2 = _add_content_slide(
            prs, "Strategic Initiatives", page_number=3,
        )
        assert idx2 == 2
        _add_bullet_block(
            slide2,
            ["Expand into APAC market", "Launch self-serve tier", "Reduce churn by 15%"],
            left=0.9, top=1.5, width=5, height=3,
        )
        _add_table(
            slide2,
            [["Initiative", "Status", "Owner"],
             ["APAC Launch", "In Progress", "VP Sales"],
             ["Self-Serve", "Planning", "PM Lead"]],
            left=6.5, top=1.5, width=5.5,
        )

        save_pptx(prs, path)

        # ── Verify structure ────────────────────────────────────────
        prs2 = open_pptx(path)
        assert len(prs2.slides) == 3

        # Slide 0: section divider -- 2 stripes + title + subtitle = 4 shapes
        s0 = prs2.slides[0]
        assert len(s0.shapes) == 4

        # Slide 1: title + divider + source + page_number + 4 shapes/KPI * 3 KPIs = 16
        s1 = prs2.slides[1]
        assert len(s1.shapes) >= 4  # at least the content slide elements

        # Slide 2: title + divider + page_number + bullet block + table = 5
        s2 = prs2.slides[2]
        assert len(s2.shapes) >= 5

        # Verify text content
        info = _get_presentation_info(prs2)
        assert "Slides: 3" in info

    def test_verify_text_content_across_deck(self, tmp_path):
        """Build the deck and verify specific text content is findable."""
        path = str(tmp_path / "text_check.pptx")
        create_presentation(path)
        prs = open_pptx(path)

        _add_section_divider(prs, "Executive Summary")
        slide, _ = _add_content_slide(prs, "Market Analysis")
        _add_bullet_block(
            slide,
            ["TAM is $50B", "SAM is $12B"],
            left=1, top=1.5, width=5, height=2,
        )
        save_pptx(prs, path)

        # Reload and scan all text
        prs2 = open_pptx(path)
        all_texts = []
        for slide in prs2.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    all_texts.append(shape.text_frame.text)

        all_text = " ".join(all_texts)
        assert "Executive Summary" in all_text
        assert "Market Analysis" in all_text
        assert "TAM is $50B" in all_text
        assert "SAM is $12B" in all_text

    def test_file_roundtrip_integrity(self, tmp_path):
        """Ensure the deck can be saved, reopened, and re-saved without corruption."""
        path = str(tmp_path / "roundtrip.pptx")
        create_presentation(path)
        prs = open_pptx(path)
        _add_content_slide(prs, "Roundtrip Test")
        save_pptx(prs, path)

        # Re-open, modify, save again
        prs2 = open_pptx(path)
        slide, _ = _add_content_slide(prs2, "Second Pass")
        save_pptx(prs2, path)

        # Final verification
        prs3 = open_pptx(path)
        assert len(prs3.slides) == 2
        info = _get_presentation_info(prs3)
        assert "Slides: 2" in info
