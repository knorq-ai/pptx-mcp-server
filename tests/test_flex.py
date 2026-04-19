"""Flex コンテナのレイアウトテスト.

``add_flex_container`` による main 軸・cross 軸サイズ配分と位置計算の挙動を
検証する。shape 生成は伴わず、呼び出し履歴を記録するスタブ render を用いる。
"""

from __future__ import annotations

import math

import pytest

from pptx_mcp_server.engine.flex import FlexItem, add_flex_container
from pptx_mcp_server.engine.pptx_io import EngineError, ErrorCode


EPS = 1e-6


def _recorder():
    """呼び出しを記録する render コールバックを生成する."""
    calls: list[tuple[float, float, float, float]] = []

    def render(x: float, y: float, w: float, h: float) -> None:
        calls.append((x, y, w, h))

    return render, calls


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def test_three_fixed_items_row_with_gap():
    """3 つの固定 2" 項目を 10" コンテナに gap 0.15 で配置する."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=2.0, render=renders[0][0]),
        FlexItem(sizing="fixed", size=2.0, render=renders[1][0]),
        FlexItem(sizing="fixed", size=2.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.15, padding=0.0, align="stretch",
    )
    assert len(allocs) == 3
    # 幅は 2.0 ずつ
    for (x, y, w, h) in allocs:
        assert _approx(w, 2.0)
        assert _approx(h, 1.0)
        assert _approx(y, 0.0)
    # x 位置: 0, 2.15, 4.30
    assert _approx(allocs[0][0], 0.0)
    assert _approx(allocs[1][0], 2.15)
    assert _approx(allocs[2][0], 4.30)
    # render が各項目で 1 回ずつ呼ばれている
    for _, calls in renders:
        assert len(calls) == 1


def test_fixed_plus_two_grow_equal_split():
    """fixed(2") + grow(1) + grow(1) in 10", gap 0 → grow 各 4.0."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=2.0, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 2.0)
    assert _approx(allocs[1][2], 4.0)
    assert _approx(allocs[2][2], 4.0)
    # x 位置の連続性
    assert _approx(allocs[1][0], 2.0)
    assert _approx(allocs[2][0], 6.0)


def test_fixed_grow_content_mixed():
    """fixed(2") + grow(1) + content(3") in 10", gap 0 → grow が 5."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=2.0, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
        FlexItem(sizing="content", content_size=3.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 2.0)
    assert _approx(allocs[1][2], 5.0)
    assert _approx(allocs[2][2], 3.0)


def test_grow_ratio_one_to_three():
    """fixed(2") + grow(1) + grow(3) → 残 8 を 1:3 で分割し 2.0 と 6.0."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=2.0, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
        FlexItem(sizing="grow", grow=3.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[1][2], 2.0)
    assert _approx(allocs[2][2], 6.0)


def test_grow_max_size_clamped_redistributes():
    """2 つの grow 項目で 1 つが max_size=3 にクランプされると残余が他へ."""
    # 残余 = 10 (fixed なし、gap=0)、2 つ grow(1:1) なら各 5.0 になるが、
    # 1 つ目に max_size=3.0 を設定すると 3.0 にクランプされ、2 つ目が 7.0 を受ける。
    renders = [_recorder() for _ in range(2)]
    items = [
        FlexItem(sizing="grow", grow=1.0, max_size=3.0, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 3.0)
    assert _approx(allocs[1][2], 7.0)


def test_column_direction_stack_vertically():
    """direction='column' で 3 つの fixed 項目が縦に積み上がる."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=1.0, render=renders[0][0]),
        FlexItem(sizing="fixed", size=1.5, render=renders[1][0]),
        FlexItem(sizing="fixed", size=2.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=1.0, top=2.0, width=5.0, height=10.0,
        direction="column", gap=0.2, padding=0.0,
    )
    # main 軸 = 縦、cross 軸 = 横
    assert _approx(allocs[0][1], 2.0)
    assert _approx(allocs[0][3], 1.0)
    assert _approx(allocs[1][1], 2.0 + 1.0 + 0.2)
    assert _approx(allocs[1][3], 1.5)
    assert _approx(allocs[2][1], 2.0 + 1.0 + 0.2 + 1.5 + 0.2)
    assert _approx(allocs[2][3], 2.0)
    # cross 軸 (x, width) は全項目で同じ
    for (x, y, w, h) in allocs:
        assert _approx(x, 1.0)
        assert _approx(w, 5.0)


def test_align_center_cross_axis_size():
    """align='center' の cross サイズは usable_cross と一致する (MVP 仕様)."""
    renders = [_recorder() for _ in range(1)]
    items = [
        FlexItem(sizing="fixed", size=3.0, render=renders[0][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=2.0,
        direction="row", gap=0.0, padding=0.0, align="center",
    )
    # cross サイズは usable_cross = height - 2*padding = 2.0
    assert _approx(allocs[0][3], 2.0)
    # cross 位置は padding から
    assert _approx(allocs[0][1], 0.0)


def test_padding_insets_all_items():
    """padding=0.2 がコンテナ全周に適用され、各項目が内側に配置される."""
    renders = [_recorder() for _ in range(2)]
    items = [
        FlexItem(sizing="fixed", size=1.0, render=renders[0][0]),
        FlexItem(sizing="fixed", size=1.0, render=renders[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=1.0, top=2.0, width=5.0, height=3.0,
        direction="row", gap=0.0, padding=0.2,
    )
    # 1 つ目の x は left + padding
    assert _approx(allocs[0][0], 1.2)
    # y は top + padding
    assert _approx(allocs[0][1], 2.2)
    # cross 高さ = height - 2*padding = 2.6
    assert _approx(allocs[0][3], 2.6)
    assert _approx(allocs[1][3], 2.6)
    # 2 つ目の x は 1 つ目の右端から (gap=0)
    assert _approx(allocs[1][0], 1.2 + 1.0)


def test_grow_only_equal_share():
    """全項目 grow=1 のみで等分される."""
    renders = [_recorder() for _ in range(4)]
    items = [
        FlexItem(sizing="grow", grow=1.0, render=renders[i][0])
        for i in range(4)
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=8.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    for (x, y, w, h) in allocs:
        assert _approx(w, 2.0)


def test_empty_items_returns_empty_list():
    """items=[] は空リストを返し、例外を投げない."""
    allocs = add_flex_container(
        slide=None,
        items=[],
        left=0.0, top=0.0, width=10.0, height=1.0,
    )
    assert allocs == []


def test_grow_min_size_clamps_up():
    """min_size でクランプされた項目の不足分は他 grow に割当てられない
    (通常の flex 動作: min はサイズを増やす方向のクランプ)."""
    # 残余 = 4.0、grow 項目 2 つ (1:1) なら各 2.0、しかし 1 つ目に min_size=3.5。
    # クランプで 3.5 を確定、remain=0.5 が 2 つ目へ。
    renders = [_recorder() for _ in range(2)]
    items = [
        FlexItem(sizing="grow", grow=1.0, min_size=3.5, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=4.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 3.5)
    assert _approx(allocs[1][2], 0.5)


def test_zero_grow_gets_no_share():
    """grow=0 は分配を受けない (0 として扱う)."""
    renders = [_recorder() for _ in range(2)]
    items = [
        FlexItem(sizing="grow", grow=0.0, render=renders[0][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 0.0)
    assert _approx(allocs[1][2], 10.0)


def test_over_budget_three_fixed_raises():
    """3 × fixed(5) in 10" コンテナ (padding=0, gap=0) は INVALID_PARAMETER で失敗する (#25)."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=5.0, render=renders[0][0]),
        FlexItem(sizing="fixed", size=5.0, render=renders[1][0]),
        FlexItem(sizing="fixed", size=5.0, render=renders[2][0]),
    ]
    with pytest.raises(EngineError) as excinfo:
        add_flex_container(
            slide=None,
            items=items,
            left=0.0, top=0.0, width=10.0, height=1.0,
            direction="row", gap=0.0, padding=0.0,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER
    # エラーメッセージには過剰な合計値と usable_main が含まれる (デバッグ用)
    msg = str(excinfo.value)
    assert "15.00" in msg
    assert "10.00" in msg
    # render は一度も呼ばれていない
    for _, calls in renders:
        assert calls == []


def test_fixed_plus_grow_with_gap_just_fits():
    """fixed(4) + fixed(4) + grow(1) in 10" (gap=0.5): fixed 8 + gap 1 = 9 ≤ 10, grow は 1."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=4.0, render=renders[0][0]),
        FlexItem(sizing="fixed", size=4.0, render=renders[1][0]),
        FlexItem(sizing="grow", grow=1.0, render=renders[2][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.5, padding=0.0,
    )
    assert _approx(allocs[0][2], 4.0)
    assert _approx(allocs[1][2], 4.0)
    assert _approx(allocs[2][2], 1.0)


def test_fixed_plus_content_exactly_fits():
    """fixed(4) + content(5) in 10" (gap=0): 9 ≤ 10、余り 1 は grow がないので未使用."""
    renders = [_recorder() for _ in range(2)]
    items = [
        FlexItem(sizing="fixed", size=4.0, render=renders[0][0]),
        FlexItem(sizing="content", content_size=5.0, render=renders[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert _approx(allocs[0][2], 4.0)
    assert _approx(allocs[1][2], 5.0)


def test_fixed_content_fixed_over_budget_raises():
    """fixed(4) + content(5) + fixed(2) in 10" (gap=0): 11 > 10 で raises."""
    renders = [_recorder() for _ in range(3)]
    items = [
        FlexItem(sizing="fixed", size=4.0, render=renders[0][0]),
        FlexItem(sizing="content", content_size=5.0, render=renders[1][0]),
        FlexItem(sizing="fixed", size=2.0, render=renders[2][0]),
    ]
    with pytest.raises(EngineError) as excinfo:
        add_flex_container(
            slide=None,
            items=items,
            left=0.0, top=0.0, width=10.0, height=1.0,
            direction="row", gap=0.0, padding=0.0,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER
    msg = str(excinfo.value)
    assert "11.00" in msg
    assert "10.00" in msg


def test_padding_tight_fit_ok_then_over_budget():
    """padding=0.5 エッジケース: usable = 10 - 1 = 9.

    - fixed(4)+fixed(4) なら fixed_total=8 ≤ 9 で OK
    - fixed(4)+fixed(4)+fixed(1.5) なら 9.5 > 9 で raises
    """
    # 境界内ケース
    renders_ok = [_recorder() for _ in range(2)]
    items_ok = [
        FlexItem(sizing="fixed", size=4.0, render=renders_ok[0][0]),
        FlexItem(sizing="fixed", size=4.0, render=renders_ok[1][0]),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items_ok,
        left=0.0, top=0.0, width=10.0, height=1.0,
        direction="row", gap=0.0, padding=0.5,
    )
    assert _approx(allocs[0][2], 4.0)
    assert _approx(allocs[1][2], 4.0)

    # 境界外ケース
    renders_bad = [_recorder() for _ in range(3)]
    items_bad = [
        FlexItem(sizing="fixed", size=4.0, render=renders_bad[0][0]),
        FlexItem(sizing="fixed", size=4.0, render=renders_bad[1][0]),
        FlexItem(sizing="fixed", size=1.5, render=renders_bad[2][0]),
    ]
    with pytest.raises(EngineError) as excinfo:
        add_flex_container(
            slide=None,
            items=items_bad,
            left=0.0, top=0.0, width=10.0, height=1.0,
            direction="row", gap=0.0, padding=0.5,
        )
    assert excinfo.value.code == ErrorCode.INVALID_PARAMETER
    msg = str(excinfo.value)
    assert "9.50" in msg
    assert "9.00" in msg


def test_render_receives_correct_args():
    """render コールバックが allocations と一致する引数で呼ばれる."""
    render, calls = _recorder()
    items = [
        FlexItem(sizing="fixed", size=3.0, render=render),
    ]
    allocs = add_flex_container(
        slide=None,
        items=items,
        left=0.5, top=0.75, width=5.0, height=1.0,
        direction="row", gap=0.0, padding=0.0,
    )
    assert len(calls) == 1
    assert calls[0] == allocs[0]
