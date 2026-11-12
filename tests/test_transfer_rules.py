# -*- coding: utf-8 -*-
"""契約測試：轉貨接收店集中規範（TRANSFER_RULES_SPEC.md）。

三層驗證：
1. 黃金向量（transfer_rules_golden.json）：期望值由轉貨引擎歷史行為產生，
   ``transfer_rules.py`` 必須逐條一致——包括 lines 的次序。
2. 引擎委派：``modes.transfer_p90d`` 的私有函數必須與共用規範同值
   （防止有人日後在引擎內另寫一份實作）。
3. 不變量與邊界：R1–R9 的性質以隨機輸入驗證，另外檢查規範文件的分檔表
   與程式實際行為一致（文件漂移會令測試失敗）。
"""
import json
import random
from pathlib import Path

import pytest

import transfer_rules as rules
from modes import transfer_p90d as engine

ROOT = Path(__file__).resolve().parent.parent
GOLDEN = json.loads((ROOT / "transfer_rules_golden.json").read_text(encoding="utf-8"))
SPEC = (ROOT / "TRANSFER_RULES_SPEC.md").read_text(encoding="utf-8")
POLICY = rules.DEFAULT_POLICY


def _ids(cases, *keys):
    out = []
    for case in cases:
        parts = [str(case[k]) if isinstance(k, str) else str(case[k]) for k in keys]
        out.append("_".join(parts).replace(" ", "")[:60])
    return out


# --------------------------------------------------------------------------------------
# 0. 版本與指紋
# --------------------------------------------------------------------------------------
def test_rules_version_matches_golden():
    assert GOLDEN["rules_version"] == rules.RULES_VERSION


def test_policy_fingerprint_matches_golden():
    """參數一改指紋即變；指紋不符代表文件／向量與程式已不同步。"""
    assert GOLDEN["policy"] == POLICY.as_dict()
    assert GOLDEN["policy_fingerprint"] == POLICY.fingerprint()
    # 規範文件頁首也記錄了指紋
    assert POLICY.fingerprint() in SPEC


# --------------------------------------------------------------------------------------
# 1. 件數 → 接收店數（R2–R5）
# --------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "total,category,expected",
    [tuple(case) for case in GOLDEN["receiver_count"]],
    ids=_ids(GOLDEN["receiver_count"], 0, 1),
)
def test_receiver_count_golden(total, category, expected):
    assert POLICY.receiver_count(total, category=category) == expected


def test_receiver_count_boundary_bands():
    """分檔表本身即規格：一般分類與迷你裝各有自己的目標件數。"""
    assert [POLICY.receiver_count(n) for n in (1, 3, 4, 6, 7, 9, 10, 12, 13, 30, 31, 60, 61, 90)] == [
        1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7
    ]
    assert [POLICY.receiver_count(n, category="0302 - MINIATURE") for n in (4, 5, 8, 9, 12, 13, 16, 17, 30)] == [
        1, 2, 2, 3, 3, 4, 4, 5, 5
    ]
    # 0303 不命中覆寫，沿用預設 3 件
    assert POLICY.receiver_count(6, category="0303 - MINIATURE SET") == 2
    assert POLICY.receiver_count(7, category="0303 - MINIATURE SET") == 3


def test_receiver_count_is_monotonic():
    """I7：件數增加，店數不會減少。"""
    for category in ("", "0302 - MINIATURE"):
        values = [POLICY.receiver_count(n, category=category) for n in range(0, 400)]
        assert all(b >= a for a, b in zip(values, values[1:]))


# --------------------------------------------------------------------------------------
# 2. 分配演算法（R7）
# --------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "total,weights,expected",
    [tuple(case) for case in GOLDEN["allocate_proportional"]],
    ids=_ids(GOLDEN["allocate_proportional"], 0),
)
def test_allocate_proportional_golden(total, weights, expected):
    assert rules.allocate_proportional(total, weights) == expected


@pytest.mark.parametrize(
    "total,weights,min_per,expected",
    [tuple(case) for case in GOLDEN["allocate_concentrated"]],
    ids=_ids(GOLDEN["allocate_concentrated"], 0),
)
def test_allocate_concentrated_golden(total, weights, min_per, expected):
    assert rules.allocate_concentrated(total, weights, min_per) == expected


@pytest.mark.parametrize(
    "total,weights,caps,alloc,unalloc",
    [tuple(case) for case in GOLDEN["allocate_with_caps"]],
    ids=_ids(GOLDEN["allocate_with_caps"], 0),
)
def test_allocate_with_caps_golden(total, weights, caps, alloc, unalloc):
    got_alloc, got_unalloc = rules.allocate_with_caps(total, weights, caps)
    assert got_alloc == alloc
    assert got_unalloc == unalloc


def test_concentrated_keeps_two_per_shop_when_possible():
    """I4：件數足夠時不出現只收 1 件的碎片。"""
    rng = random.Random(20261112)
    for _ in range(300):
        total = rng.randint(1, 300)
        n = rng.randint(1, min(8, max(1, total)))
        weights = [float(rng.randint(1, 80)) for _ in range(n)]
        alloc = rules.allocate_concentrated(total, weights)
        assert sum(alloc) == total
        assert min(alloc) >= min(2, total // n)


def test_with_caps_invariants():
    """I1／I3：不超上限，且已分配＋未分配＝總量。"""
    rng = random.Random(7)
    for _ in range(300):
        total = rng.randint(0, 200)
        n = rng.randint(1, 6)
        weights = [float(rng.randint(0, 50)) for _ in range(n)]
        caps = [rng.randint(0, 30) for _ in range(n)]
        alloc, unalloc = rules.allocate_with_caps(total, weights, caps)
        assert alloc == [min(a, c) for a, c in zip(alloc, caps)]
        assert sum(alloc) + unalloc == total


def test_proportional_always_fully_allocated():
    """I1／I2：權重總和為 0 時退化為等權，仍然全數分完。"""
    rng = random.Random(11)
    for _ in range(200):
        total = rng.randint(1, 200)
        n = rng.randint(1, 6)
        weights = [0.0] * n
        alloc = rules.allocate_proportional(total, weights)
        assert sum(alloc) == total


# --------------------------------------------------------------------------------------
# 3. 選店（R1／R6）
# --------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "case",
    GOLDEN["select_pool"],
    ids=[f"t{c['total']}_m{c['mode']}_c{(c['category'] or 'none')[:8]}" for c in GOLDEN["select_pool"]],
)
def test_select_pool_golden(case):
    pool, dropped = rules.select_pool(
        case["candidates"], case["total"], mode=case["mode"], category=case["category"]
    )
    assert [r[0] for r in pool] == case["pool"]
    assert dropped == case["dropped"]


def test_select_pool_tie_break_by_site_code():
    """同銷量以 Site Code 升序——規範要求可重現，不能靠輸入次序。"""
    pool, _ = rules.select_pool(
        [("S03", 9, 0), ("S01", 9, 0), ("S02", 9, 0)], 3, mode=2
    )
    assert [r[0] for r in pool] == ["S01", "S02"]


def test_select_pool_unlimited_keeps_site_order():
    """不限模式：全數候選店，按 Site Code 升序（與未啟用集中時一致）。"""
    pool, dropped = rules.select_pool(
        [("S03", 1, 0), ("S01", 9, 0), ("S02", 5, 0)], 30, mode=None
    )
    assert [r[0] for r in pool] == ["S01", "S02", "S03"]
    assert dropped == 0


# --------------------------------------------------------------------------------------
# 4. 端到端（R9）
# --------------------------------------------------------------------------------------
@pytest.mark.parametrize(
    "case",
    GOLDEN["plans"],
    ids=[f"t{c['total']}_m{c['mode']}_cap{c['cap_days']}_{i}" for i, c in enumerate(GOLDEN["plans"])],
)
def test_plan_allocation_golden(case):
    result = rules.plan_allocation(
        case["total"], case["candidates"], mode=case["mode"],
        category=case["category"], cap_days=case["cap_days"],
    )
    assert [[site, qty] for site, qty in result.lines] == case["lines"]
    assert result.allocated == case["allocated"]
    assert result.unallocated == case["unallocated"]
    assert result.dropped == case["dropped"]
    assert result.allocated + result.unallocated == case["total"]


def test_plan_allocation_helpers():
    result = rules.plan_allocation(
        13, [("S02", 9, 0), ("S03", 7, 0), ("S04", 5, 0)], mode="auto"
    )
    assert result.mode == "auto"
    assert result.quantities == {"S02": 5, "S03": 4, "S04": 4}
    assert result.receivers == ("S02", "S03", "S04")
    assert POLICY.fingerprint() in result.summary()
    assert "13 件" in result.summary()

    unlimited = rules.plan_allocation(10, [("S02", 9, 0), ("S03", 7, 0)], mode=None)
    assert unlimited.mode == rules.MODE_UNLIMITED


def test_plan_allocation_cap_and_concentration_combined():
    """封頂優先於保底：30 件對一間只能補 14 日的店，剩 16 件保留原店。"""
    result = rules.plan_allocation(30, [("S02", 90, 0)], mode="auto", cap_days=14)
    assert result.lines == (("S02", 14),)
    assert result.unallocated == 16


# --------------------------------------------------------------------------------------
# 5. 邊界與錯誤處理
# --------------------------------------------------------------------------------------
def test_boundaries():
    empty = rules.plan_allocation(0, [("S02", 9, 0)])
    assert empty.lines == () and empty.allocated == 0 and empty.unallocated == 0

    no_receiver = rules.plan_allocation(10, [], mode="auto")
    assert no_receiver.lines == () and no_receiver.unallocated == 10

    single = rules.plan_allocation(7, [("S02", 3, 0)])
    assert single.lines == (("S02", 7),)

    small = rules.plan_allocation(3, [("S0%d" % i, 1, 0) for i in range(1, 6)])
    assert small.allocated == 3 and small.allocated + small.unallocated == 3


def test_candidates_accept_mapping_and_pairs():
    by_pair = rules.plan_allocation(5, [("S02", 9), ("S03", 7)])
    by_dict = rules.plan_allocation(
        5, [{"site": "S02", "sales": 9}, {"Site Code": "S03", "Sold Qty (P90D)": 7}]
    )
    assert by_pair.lines == by_dict.lines


@pytest.mark.parametrize("bad", ["auto2", "concentrate", 0, -1, True, 3.5])
def test_invalid_mode_rejected(bad):
    with pytest.raises(ValueError) as exc:
        rules.normalize_mode(bad)
    assert "集中上限" in str(exc.value)


@pytest.mark.parametrize(
    "value,expected",
    [("auto", "auto"), (" AUTO ", "auto"), (None, None), ("unlimited", None),
     ("none", None), ("", None), ("4", 4), (2, 2)],
)
def test_normalize_mode_accepts(value, expected):
    assert rules.normalize_mode(value) == expected


# --------------------------------------------------------------------------------------
# 6. 引擎必須用同一份實作（防止日後在引擎內另寫一份）
# --------------------------------------------------------------------------------------
def test_engine_delegates_to_shared_rules():
    rng = random.Random(2026)
    for total in list(range(0, 100)) + [rng.randint(1, 500) for _ in range(50)]:
        assert engine._auto_receiver_count(total) == POLICY.receiver_count(total)
        assert engine._auto_receiver_cap(total) == POLICY.cap(total)
        for category in ("", "0302 - MINIATURE", "0303 - MINIATURE SET"):
            assert engine._auto_receiver_target(category) == POLICY.target_for(category)
    for _ in range(50):
        total = rng.randint(0, 200)
        n = rng.randint(0, 6)
        weights = [float(rng.randint(0, 60)) for _ in range(n)]
        caps = [rng.randint(0, 20) for _ in range(n)]
        assert engine._allocate_proportional(total, weights) == rules.allocate_proportional(total, weights)
        assert engine._allocate_concentrated(total, weights) == rules.allocate_concentrated(total, weights)
        assert engine._allocate_with_caps(total, weights, caps) == rules.allocate_with_caps(total, weights, caps)


def test_engine_policy_matches_shared_default_policy():
    """本專案的參數值必須與規範預設一致，否則黃金向量不再代表引擎行為。"""
    assert engine._TRANSFER_POLICY.as_dict() == POLICY.as_dict()


def test_engine_receiver_cap_uses_14_days():
    assert engine._receiver_cap(90, 0) == rules.daily_demand(90, 0, 14)
    assert engine._receiver_cap(90, 0) == 14
    assert engine._receiver_cap(30, 10) == 0


# --------------------------------------------------------------------------------------
# 7. 參數調整
# --------------------------------------------------------------------------------------
def test_policy_overrides_do_not_mutate_default():
    looser = POLICY.with_overrides(base_cap=4)
    assert POLICY.base_cap == 5
    assert looser.base_cap == 4
    assert looser.receiver_count(30) == 4
    assert POLICY.receiver_count(30) == 5


def test_policy_from_dict_roundtrip():
    assert rules.policy_from_dict(POLICY.as_dict()) == POLICY
    assert rules.policy_from_dict({"policy": POLICY.as_dict()}) == POLICY
    assert rules.policy_from_dict(None) == POLICY
    assert rules.policy_from_dict({"unknown": 1}) == POLICY


def test_policy_describe_matches_table():
    text = POLICY.describe()
    assert "3 件或以下 → 1 間" in text
    assert "13–30 件 → 5 間" in text
    assert "保底 2 件" in text


# --------------------------------------------------------------------------------------
# 8. 文件不得漂移：SPEC 的分檔表須與程式一致
# --------------------------------------------------------------------------------------
def _spec_rows(category=""):
    rows = []
    for line in rules.format_rule_table(category=category):
        left, right = line.split(" → ")
        # 末列在程式輸出會帶「（目前 N 間）」，文件表格不寫，比對時統一去掉
        right = right.split("（")[0]
        cell = right if right.startswith("每") else right.replace(" 間", "")
        rows.append(f"| {left} | {cell} |")
    return rows


@pytest.mark.parametrize("category", ["", "0302 - MINIATURE"])
def test_spec_table_matches_code(category):
    rows = _spec_rows(category)
    assert rows, "分檔表不應為空"
    missing = [row for row in rows if row not in SPEC]
    assert not missing, f"規範文件分檔表與程式不一致，缺少：{missing}"
