"""通用模式搬遷器：將指定模式的函數／常數由 transformer.py 移至 modes/<mod>.py。

用法: python move_mode.py <mode_key>
每次只搬一個模式；搬完需跑測試，通過才繼續下一個。
狀態: C:/Users/BestO/AppData/Local/Temp/opencode/moved.json
"""
import ast
import json
import pathlib
import sys

REPO = pathlib.Path(__file__).resolve().parents[2]
STATE = pathlib.Path(__file__).resolve().parent / "moved.json"

SPECS = {
    "oos": {
        "module": "oos", "label": "OOS Review",
        "seeds": ["transform_oos", "generate_excel_oos"],
        "transform": "transform_oos", "generate_excel": "generate_excel_oos",
    },
    "nd": {
        "module": "nd", "label": "SKU 精簡 Review",
        "seeds": ["transform_nd", "generate_excel_nd"],
        "transform": "transform_nd", "generate_excel": "generate_excel_nd",
    },
    "overstock": {
        "module": "overstock", "label": "Overstock Review",
        "seeds": ["transform", "generate_excel"],
        "transform": "transform", "generate_excel": "generate_excel",
    },
    "carry": {
        "module": "carry", "label": "Carry SKU Review",
        "seeds": ["transform_carry", "generate_excel_carry"],
        "transform": "transform_carry", "generate_excel": "generate_excel_carry",
    },
    "ndrf": {
        "module": "ndrf", "label": "NDRF 月份對比",
        "seeds": ["transform_ndrf_comparison", "generate_excel_ndrf_comparison"],
        "transform": "transform_ndrf_comparison", "generate_excel": "generate_excel_ndrf_comparison",
    },
    "stock_checking": {
        "module": "stock_checking", "label": "Stock Checking",
        "seeds": ["transform_stock_checking", "generate_excel_stock_checking"],
        "transform": "transform_stock_checking", "generate_excel": "generate_excel_stock_checking",
    },
    "safety_review": {
        "module": "safety_review", "label": "Safety Stock Review",
        "seeds": ["transform_combined", "generate_excel_combined"],
        "transform": "transform_combined", "generate_excel": "generate_excel_combined",
    },
    "stock_impact": {
        "module": "stock_impact", "label": "Stock Impact Analysis",
        "seeds": ["transform_stock_impact", "generate_excel_stock_impact",
                  "generate_excel_stock_impact_simplified"],
        "transform": "transform_stock_impact", "generate_excel": "generate_excel_stock_impact",
    },
    "hybrid": {
        "module": "hybrid", "label": "Safety Stock x NetStock Review",
        "seeds": ["transform_hybrid", "generate_excel_hybrid"],
        "transform": "transform_hybrid", "generate_excel": "generate_excel_hybrid",
    },
    "forecast": {
        "module": "forecast", "label": "Forecast Review",
        "seeds": ["transform_forecast", "generate_excel_forecast"],
        "transform": "transform_forecast", "generate_excel": "generate_excel_forecast",
    },
}

# 共用核心已抽出（modes/common／issues／framework），此處不再保留任何前綴；
# 各模式專屬 helper 由閉包自動歸位，跨模式引用走 cross-import。
KEEP_PREFIXES = ()

HEADER = '''# -*- coding: utf-8 -*-
"""{label}模式：由 transformer.py 拆分而來。
共用 helper 由 transformer（最終為 modes.common／issues／framework）引用；
transformer 以墊片 re-export 本模組，舊有 import 路徑不變。
"""
import datetime
import gc
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

'''


def node_names(node):
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
        return [node.name]
    if isinstance(node, ast.Assign):
        out = []
        for t in node.targets:
            for x in ast.walk(t):
                if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store):
                    out.append(x.id)
        return out
    return []


def main():
    key = sys.argv[1]
    spec = SPECS[key]
    src_path = REPO / "transformer.py"
    text = src_path.read_text(encoding="utf-8-sig")
    srclines = text.splitlines(keepends=True)
    mod = ast.parse(text)

    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
    already_moved = set(state)

    # 所有頂層 def/const 節點
    nodes = [n for n in mod.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Assign))]
    name2node = {}
    for n in nodes:
        for nm in node_names(n):
            name2node[nm] = n

    # 模組級已載入名稱集合（供參照解析）
    mod_names = set(name2node)

    # 每個節點引用的模組級名稱
    refs = {}
    for n in nodes:
        hit = set()
        for node in ast.walk(n):
            if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
                if node.id in mod_names:
                    hit.add(node.id)
        refs[n] = hit - set(node_names(n))

    # 反向：名稱 -> 引用它的節點集合
    referrers = {nm: set() for nm in mod_names}
    for n in nodes:
        for r in refs[n]:
            referrers[r].add(n)

    def node_key(n):
        return tuple(sorted(node_names(n)))

    # 閉包：種子 + 僅被閉包內引用的名稱
    closure_nodes = set()
    for s in spec["seeds"]:
        closure_nodes.add(name2node[s])
    changed = True
    while changed:
        changed = False
        closure_names = set()
        for n in closure_nodes:
            closure_names.update(node_names(n))
        for n in list(closure_nodes):
            for r in refs[n]:
                if r in closure_names or r in already_moved:
                    continue
                if any(node_names(rn)[0].startswith(p) for p in KEEP_PREFIXES for rn in [name2node[r]]):
                    continue
                rn = name2node[r]
                rn_refs = set()
                for q in referrers[r]:
                    rn_refs.update(node_names(q))
                if rn_refs <= closure_names:
                    closure_nodes.add(rn)
                    changed = True

    closure_names = set()
    for n in closure_nodes:
        closure_names.update(node_names(n))

    # 共用引用（仍在 transformer）與跨模式引用（已搬遷）
    need_shared, need_cross = set(), {}
    for n in closure_nodes:
        for r in refs[n]:
            if r in closure_names:
                continue
            if r in already_moved:
                need_cross.setdefault(state[r], set()).add(r)
            elif r in mod_names:
                need_shared.add(r)
            # 非模組級名稱（builtins/pd/nq…）忽略，由檔頭統一引入

    print(f"[{key}] move {len(closure_names)} names")
    for nm in sorted(closure_names):
        print(f"  move: {nm}")
    print(f"  shared from transformer: {sorted(need_shared)}")
    for m, ns in need_cross.items():
        print(f"  cross from modes.{m}: {sorted(ns)}")

    # 產生 modes/<mod>.py（依原檔順序）
    ordered = sorted(closure_nodes, key=lambda n: (n.lineno, n.col_offset))
    parts = []
    for n in ordered:
        start = n.lineno - 1
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and n.decorator_list:
            start = min(d.lineno for d in n.decorator_list) - 1
        # 回收緊貼上方的註解行
        while start > 0 and srclines[start - 1].strip().startswith("#"):
            start -= 1
        parts.append("".join(srclines[start:n.end_lineno]))
    imports = []
    if need_shared:
        imports.append("from transformer import (\n" + "".join(f"    {x},\n" for x in sorted(need_shared)) + ")")
    for m in sorted(need_cross):
        imports.append(f"from modes.{m} import (\n" + "".join(f"    {x},\n" for x in sorted(need_cross[m])) + ")")
    body = HEADER.format(label=spec["label"])
    if imports:
        body += "\n\n\n".join(imports) + "\n"
    body += "\n\n__all__ = [\n" + "".join(f'    "{x}",\n' for x in sorted(closure_names)) + "]\n"
    body += "\n\n" + "\n\n\n".join(parts) + "\n"
    (REPO / "modes" / f"{spec['module']}.py").write_text(body, encoding="utf-8-sig")

    # transformer.py：刪除已搬節點，尾部加墊片
    remove = set()
    for n in closure_nodes:
        remove.update(range(n.lineno - 1, n.end_lineno))
        # 連同緊貼上方的註解行一併移除（與 modes 檔回收邏輯一致）
        s = n.lineno - 1
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and n.decorator_list:
            s = min(d.lineno for d in n.decorator_list) - 1
        while s > 0 and srclines[s - 1].strip().startswith("#"):
            s -= 1
        remove.update(range(s, n.lineno - 1))
    kept = [ln for i, ln in enumerate(srclines) if i not in remove]
    shim = (
        f"\n\n# ─── {spec['label']}：實作已遷至 modes/{spec['module']}.py（墊片 re-export）───\n"
        f"from modes.{spec['module']} import *  # noqa: F401,F403\n"
    )
    src_path.write_text("".join(kept).rstrip("\n") + "\n\n\n" + shim.lstrip("\n"), encoding="utf-8-sig")

    # 註冊表
    init = REPO / "modes" / "__init__.py"
    reg = init.read_text(encoding="utf-8-sig").rstrip("\n") + "\n"
    reg += (
        f'\nMODE_REGISTRY["{spec["label"]}"] = {{\n'
        f'    "module": "modes.{spec["module"]}",\n'
        f'    "transform": "{spec["transform"]}",\n'
        f'    "generate_excel": "{spec["generate_excel"]}",\n'
        f"}}\n"
    )
    init.write_text(reg, encoding="utf-8-sig")

    state.update({nm: spec["module"] for nm in closure_names})
    STATE.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")

    for f in [src_path, REPO / "modes" / f"{spec['module']}.py", init]:
        ast.parse(f.read_text(encoding="utf-8-sig"))

    # 定義期求值檢查：殘留模組級 Assign 若引用本次搬走名稱，一併遷入同模組
    # （def 預設值／裝飾器命中則中止，需人手處理）
    moved_here = set(closure_names)
    tmod = ast.parse(src_path.read_text(encoding="utf-8-sig"))
    eager_assigns, eager_defs = [], []
    for n in tmod.body:
        if isinstance(n, ast.Assign):
            refs = {x.id for x in ast.walk(n.value) if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load)}
            if refs & moved_here:
                eager_assigns.append(n)
        elif isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            refs = set()
            for b in list(n.args.defaults) + [d for d in n.args.kw_defaults if d is not None] + list(n.decorator_list):
                refs |= {x.id for x in ast.walk(b) if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load)}
            if refs & moved_here:
                eager_defs.append(n)
    assert not eager_defs, f"def-time refs to moved names: {[n.name for n in eager_defs]}"
    if eager_assigns:
        mp = REPO / "modes" / f"{spec['module']}.py"
        mt = mp.read_text(encoding="utf-8-sig")
        cur = src_path.read_text(encoding="utf-8-sig").splitlines(keepends=True)
        follow_names = []
        for n in eager_assigns:
            for t_ in n.targets:
                for x in ast.walk(t_):
                    if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store):
                        follow_names.append(x.id)
            mt += "\n\n\n" + "".join(cur[n.lineno - 1:n.end_lineno])
            print(f"  eager-follow: {follow_names[-1]}")
        m = re.search(r"__all__ = \[(.*?)\]", mt, re.S)
        assert m
        mt = mt[:m.end() - 1] + "".join(f'    "{x}",\n' for x in sorted(set(follow_names))) + mt[m.end() - 1:]
        mp.write_text(mt, encoding="utf-8-sig")
        # transformer 刪除跟隨節點
        tlines = src_path.read_text(encoding="utf-8-sig").splitlines(keepends=True)
        tmod2 = ast.parse("".join(tlines))
        rm = set()
        for n in tmod2.body:
            if isinstance(n, ast.Assign):
                tn = [t_.id for t_ in n.targets if isinstance(t_, ast.Name)]
                if any(x in follow_names for x in tn):
                    rm.update(range(n.lineno - 1, n.end_lineno))
        src_path.write_text("".join(ln for i, ln in enumerate(tlines) if i not in rm), encoding="utf-8-sig")
        state.update({x: spec["module"] for x in follow_names})
        STATE.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")
        ast.parse(src_path.read_text(encoding="utf-8-sig"))
        ast.parse(mp.read_text(encoding="utf-8-sig"))
        print("parse OK (eager-follow)")
    print("parse OK")


if __name__ == "__main__":
    main()
