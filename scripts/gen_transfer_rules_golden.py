# -*- coding: utf-8 -*-
"""產生轉貨集中規則的驗收向量（transfer_rules_golden.json）。

期望值一律由**現行轉貨引擎**（``modes/transfer_p90d.py``）算出，再由
``transfer_rules.py`` 比對；兩邊一致才寫檔。因此這份 JSON 是「歷史行為的快照」，
不是新模組的自我複述——引擎日後改動若令輸出偏離，測試會立刻失敗。

用法：
    py -3.14 scripts/gen_transfer_rules_golden.py            # 比對並覆寫 JSON
    py -3.14 scripts/gen_transfer_rules_golden.py --check     # 只比對，不寫檔
"""
from __future__ import annotations

import json
import random
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import transfer_rules as rules  # noqa: E402
from modes.transfer_p90d import (  # noqa: E402
    _allocate_concentrated,
    _allocate_proportional,
    _allocate_with_caps,
    _auto_receiver_count,
    _auto_receiver_target,
    _receiver_cap,
)

OUT = ROOT / "transfer_rules_golden.json"

CATEGORIES = ["", "0302 - MINIATURE", "0302", "0303 - MINIATURE SET", "01 - SKIN CARE"]


def engine_pool(total: int, candidates, mode, category: str):
    """現行引擎的選店邏輯（與 transform_p90d_transfer 內的分支一致）。"""
    rows = [(str(s), float(sa), float(st)) for s, sa, st in candidates]
    if mode is None:
        return sorted(rows, key=lambda r: r[0])
    take = (
        _auto_receiver_count(total, _auto_receiver_target(category))
        if mode == "auto"
        else int(mode)
    )
    return sorted(rows, key=lambda r: (-r[1], r[0]))[: max(take, 0)]


def engine_alloc(total: int, pool, mode, cap_days):
    weights = [r[1] for r in pool]
    if not pool:
        return [0] * len(pool), int(total)
    if cap_days is not None:
        caps = [_receiver_cap(r[1], r[2]) for r in pool]
        return _allocate_with_caps(total, weights, caps)
    if mode == "auto":
        alloc = _allocate_concentrated(total, weights)
        return alloc, int(total) - sum(alloc)
    alloc = _allocate_proportional(total, weights)
    return alloc, int(total) - sum(alloc)


def build() -> dict:
    rng = random.Random(20261112)
    data: dict = {
        "rules_version": rules.RULES_VERSION,
        "generated_from": "modes/transfer_p90d.py (v1.95.3) 的歷史行為",
        "note": "期望值為整數件；lines 的次序即分配次序，實作須完全一致才算通過。",
        "policy": rules.DEFAULT_POLICY.as_dict(),
        "policy_fingerprint": rules.DEFAULT_POLICY.fingerprint(),
    }

    # A. 件數 → 接收店數
    counts = []
    for category in CATEGORIES:
        for total in range(0, 121):
            expected = (
                0
                if total <= 0
                else _auto_receiver_count(total, _auto_receiver_target(category))
            )
            counts.append([total, category, int(expected)])
    data["receiver_count"] = counts

    # B. 純比例分配
    prop_cases = [
        (10, [3, 1]),
        (7, [1, 1, 1]),
        (0, [1, 2]),
        (5, []),
        (13, [9.0, 7.0, 5.0, 3.0, 1.0]),
        (1, [0.0, 0.0, 0.0]),
        (9, [0.0, 5.0]),
        (100, [1.0]),
        (30, [2.0, 3.0, 5.0, 7.0, 11.0]),
    ]
    for _ in range(40):
        total = rng.randint(1, 300)
        n = rng.randint(1, 8)
        prop_cases.append((total, [float(rng.randint(0, 60)) for _ in range(n)]))
    data["allocate_proportional"] = [
        [total, list(weights), _allocate_proportional(total, list(weights))]
        for total, weights in prop_cases
    ]

    # C. 集中分配（先保底）
    conc_cases = [
        (10, [9.0, 7.0, 5.0, 3.0]),
        (7, [9.0, 7.0, 5.0]),
        (6, [9.0, 7.0]),
        (3, [9.0]),
        (1, [9.0]),
        (0, [1.0, 2.0]),
        (10, [1000.0, 1.0, 1.0, 1.0]),
        (4, [9.0, 7.0]),
        (5, [0.0, 0.0]),
        (13, [9.0, 7.0, 5.0]),
    ]
    for _ in range(40):
        total = rng.randint(1, 200)
        n = rng.randint(1, min(8, max(1, total)))
        conc_cases.append((total, [float(rng.randint(1, 80)) for _ in range(n)]))
    data["allocate_concentrated"] = [
        [total, list(weights), 2, _allocate_concentrated(total, list(weights))]
        for total, weights in conc_cases
    ]

    # D. 封頂分配
    cap_cases = [
        (60, [1.0, 100.0], [28, 2]),
        (30, [90.0], [14]),
        (30, [90.0], [0]),
        (10, [1.0, 1.0, 1.0], [2, 2, 2]),
        (7, [5.0, 5.0], [3, 100]),
        (0, [1.0], [5]),
        (12, [1.0, 2.0, 3.0], [0, 0, 0]),
    ]
    for _ in range(30):
        total = rng.randint(1, 120)
        n = rng.randint(1, 5)
        weights = [float(rng.randint(1, 40)) for _ in range(n)]
        caps = [rng.randint(0, 20) for _ in range(n)]
        cap_cases.append((total, weights, caps))
    data["allocate_with_caps"] = []
    for total, weights, caps in cap_cases:
        alloc, unalloc = _allocate_with_caps(total, list(weights), list(caps))
        data["allocate_with_caps"].append(
            [total, list(weights), list(caps), alloc, int(unalloc)]
        )

    # E. 選店次序
    select_cases = [
        (13, [["S04", 5, 0], ["S02", 9, 0], ["S03", 7, 0]], "auto", ""),
        (13, [["S04", 5, 0], ["S02", 9, 0], ["S03", 7, 0]], 2, ""),
        (13, [["S04", 5, 0], ["S02", 9, 0], ["S03", 7, 0]], 4, ""),
        (13, [["S04", 5, 0], ["S02", 9, 0], ["S03", 7, 0]], None, ""),
        (4, [["S04", 5, 0], ["S02", 9, 0], ["S03", 7, 0]], "auto", "0302 - MINIATURE"),
        (6, [["S02", 9, 0], ["S03", 9, 0], ["S01", 9, 0]], "auto", ""),
        (30, [["S0%d" % i, 20 - i, 0] for i in range(1, 9)], "auto", ""),
    ]
    data["select_pool"] = []
    for total, candidates, mode, category in select_cases:
        pool = engine_pool(total, candidates, mode, category)
        data["select_pool"].append({
            "total": total,
            "mode": mode,
            "category": category,
            "candidates": candidates,
            "pool": [r[0] for r in pool],
            "dropped": len(candidates) - len(pool),
        })

    # F. 端到端分配
    plans = []
    plan_cases = [
        (13, [["S02", 9, 0], ["S03", 7, 0], ["S04", 5, 0]], "auto", "", None),
        (3, [["S02", 9, 0], ["S03", 7, 0]], "auto", "", None),
        (4, [["S02", 9, 0], ["S03", 7, 0]], "auto", "0302 - MINIATURE", None),
        (30, [["S02", 90, 0]], "auto", "", 14),
        (10, [["S02", 9, 0], ["S03", 7, 0], ["S04", 5, 0]], None, "", None),
        (10, [["S02", 9, 0], ["S03", 7, 0], ["S04", 5, 0]], 2, "", None),
        (7, [["S02", 9, 3], ["S03", 7, 0]], "auto", "", 14),
        (5, [], "auto", "", None),
    ]
    for _ in range(25):
        total = rng.randint(1, 150)
        n = rng.randint(1, 6)
        candidates = [
            [f"R{i:02d}", float(rng.randint(1, 60)), float(rng.randint(0, 6))]
            for i in range(1, n + 1)
        ]
        plan_cases.append((
            total, candidates, rng.choice(["auto", "auto", None, 2, 4]), "",
            rng.choice([None, None, 14]),
        ))
    for total, candidates, mode, category, cap_days in plan_cases:
        pool = engine_pool(total, candidates, mode, category)
        alloc, unalloc = engine_alloc(total, pool, mode, cap_days)
        plans.append({
            "total": total,
            "mode": mode,
            "category": category,
            "cap_days": cap_days,
            "candidates": candidates,
            "lines": [[r[0], int(a)] for r, a in zip(pool, alloc)],
            "allocated": int(sum(alloc)),
            "unallocated": int(unalloc),
            "dropped": len(candidates) - len(pool),
        })
    data["plans"] = plans
    return data


def verify(data: dict) -> list[str]:
    """用 transfer_rules 重跑同一批向量，回傳不一致的描述（空＝全數通過）。"""
    errors: list[str] = []
    policy = rules.DEFAULT_POLICY

    for total, category, expected in data["receiver_count"]:
        got = policy.receiver_count(total, category=category)
        if got != expected:
            errors.append(f"receiver_count({total}, {category!r}) = {got}，期望 {expected}")

    for total, weights, expected in data["allocate_proportional"]:
        got = rules.allocate_proportional(total, weights)
        if got != expected:
            errors.append(f"allocate_proportional({total}, {weights}) = {got}，期望 {expected}")

    for total, weights, min_per, expected in data["allocate_concentrated"]:
        got = rules.allocate_concentrated(total, weights, min_per)
        if got != expected:
            errors.append(f"allocate_concentrated({total}, {weights}) = {got}，期望 {expected}")

    for total, weights, caps, alloc, unalloc in data["allocate_with_caps"]:
        got_alloc, got_unalloc = rules.allocate_with_caps(total, weights, caps)
        if got_alloc != alloc or got_unalloc != unalloc:
            errors.append(
                f"allocate_with_caps({total}, {weights}, {caps}) = {got_alloc}/{got_unalloc}"
                f"，期望 {alloc}/{unalloc}"
            )

    for case in data["select_pool"]:
        pool, dropped = rules.select_pool(
            case["candidates"], case["total"], mode=case["mode"], category=case["category"]
        )
        sites = [r[0] for r in pool]
        if sites != case["pool"] or dropped != case["dropped"]:
            errors.append(
                f"select_pool(total={case['total']}, mode={case['mode']!r}) = {sites}/{dropped}"
                f"，期望 {case['pool']}/{case['dropped']}"
            )

    for case in data["plans"]:
        result = rules.plan_allocation(
            case["total"], case["candidates"], mode=case["mode"],
            category=case["category"], cap_days=case["cap_days"],
        )
        lines = [[site, qty] for site, qty in result.lines]
        if (
            lines != case["lines"]
            or result.allocated != case["allocated"]
            or result.unallocated != case["unallocated"]
            or result.dropped != case["dropped"]
        ):
            errors.append(
                f"plan(total={case['total']}, mode={case['mode']!r}, cap_days={case['cap_days']})"
                f" = {lines}/{result.allocated}/{result.unallocated}/{result.dropped}"
                f"，期望 {case['lines']}/{case['allocated']}/{case['unallocated']}/{case['dropped']}"
            )
    return errors


def main() -> int:
    check_only = "--check" in sys.argv
    data = build()
    errors = verify(data)
    if errors:
        print(f"✗ 與現行引擎不一致 {len(errors)} 項：")
        for line in errors[:20]:
            print("  -", line)
        return 1
    if check_only:
        print("✓ 驗收向量與引擎一致（未寫檔）")
        return 0
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    total_cases = sum(
        len(data[k]) for k in (
            "receiver_count", "allocate_proportional", "allocate_concentrated",
            "allocate_with_caps", "select_pool", "plans",
        )
    )
    print(f"✓ 已寫入 {OUT.name}：共 {total_cases} 個向量，與引擎完全一致")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
