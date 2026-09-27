# -*- coding: utf-8 -*-
"""check_modes.py：modes/ 架構健康檢查（唯讀，不修改任何檔案）。

檢查項目：
  A. 引用缺口 — 各模組參照但未定義、未引入的名稱（ERROR）
  B. 循環引用 — modes 模組間 import 有向圖必須無環（ERROR）
  C. 對外契約 — views／tests／scripts 經 transformer 引入的名稱皆可解析（ERROR）
  D. 定義期求值 — transformer 殘留的模組級常數／預設值不可引用墊片名稱（ERROR）
  E. __all__／moved.json 一致性（WARNING）

用法：python scripts/modtools/check_modes.py
結束碼 0 為通過，1 為有 ERROR。
"""
import ast
import builtins
import json
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parents[2]
STATE = pathlib.Path(__file__).resolve().parent / "moved.json"
sys.path.insert(0, str(REPO))

THIRD = {
    "pd", "np", "gc", "re", "unicodedata", "datetime", "BytesIO",
    "ThreadPoolExecutor", "Workbook", "Font", "Alignment", "Border",
    "Side", "PatternFill", "get_column_letter", "DataValidation", "st",
} | set(dir(builtins))

errors, warnings = [], []


def mod_names(path):
    """模組頂層定義名稱。"""
    m = ast.parse(path.read_text(encoding="utf-8-sig"))
    out = set()
    for n in m.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            out.add(n.name)
        elif isinstance(n, ast.Assign):
            for tg in n.targets:
                for x in ast.walk(tg):
                    if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store):
                        out.add(x.id)
    return out


def local_names(tree):
    out = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            out.add(node.name)
        elif isinstance(node, ast.ClassDef):
            out.add(node.name)
        elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            out.add(node.id)
        elif isinstance(node, ast.Import):
            for a in node.names:
                out.add((a.asname or a.name).split(".")[0])
    for node in ast.walk(tree):  # 參數（含 lambda）
        args = getattr(node, "args", None)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)) and args:
            for x in list(args.posonlyargs) + list(args.args) + list(args.kwonlyargs):
                out.add(x.arg)
            if args.vararg:
                out.add(args.vararg.arg)
            if args.kwarg:
                out.add(args.kwarg.arg)
    return out


def main():
    origin = json.loads(STATE.read_text(encoding="utf-8"))
    t_text = (REPO / "transformer.py").read_text(encoding="utf-8-sig")
    tbody = ast.parse(t_text)
    body_names = set()
    for n in tbody.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            body_names.add(n.name)
        elif isinstance(n, ast.Assign):
            for tg in n.targets:
                for x in ast.walk(tg):
                    if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store):
                        body_names.add(x.id)

    # A. 引用缺口＋B. 邊收集
    edges = {}
    for modfile in sorted((REPO / "modes").glob("*.py")):
        if modfile.name == "__init__.py":
            continue
        m = ast.parse(modfile.read_text(encoding="utf-8-sig"))
        local = local_names(m)
        loads = {x.id for x in ast.walk(m) if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load)}
        for x in sorted(loads - local - THIRD):
            if x.startswith("__") or x in origin or x in body_names:
                continue
            errors.append(f"A. {modfile.name} 引用缺口：{x}")
        edge = set()
        for n in m.body:
            if isinstance(n, ast.ImportFrom) and n.module and n.module.startswith("modes."):
                edge.add(n.module.split(".")[1])
        edges[modfile.stem] = edge

    # B. 環檢測（DFS）
    WHITE, GRAY, BLACK = 0, 1, 2
    color = {k: WHITE for k in edges}
    stack = []

    def dfs(u):
        color[u] = GRAY
        stack.append(u)
        for v in edges.get(u, ()):
            if v not in color:
                continue
            if color[v] == GRAY:
                errors.append(f"B. 循環引用：{' -> '.join(stack + [v])}")
            elif color[v] == WHITE:
                dfs(v)
        stack.pop()
        color[u] = BLACK

    for u in sorted(edges):
        if color[u] == WHITE:
            dfs(u)

    # C. 對外契約
    import transformer  # noqa: E402
    names = set()
    for d in ("views", "tests", "scripts"):
        for f in (REPO / d).glob("*.py"):
            t = f.read_text(encoding="utf-8-sig")
            for mm in re.finditer(r"from transformer import \((.*?)\)", t, re.S):
                names.update(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", mm.group(1)))
    missing = sorted(x for x in names if not hasattr(transformer, x))
    for x in missing:
        errors.append(f"C. transformer 缺少對外名稱：{x}")

    # D. 定義期求值
    for n in tbody.body:
        targets = []
        if isinstance(n, ast.Assign):
            bodies, targets = [n.value], [t.id for t in n.targets if isinstance(t, ast.Name)]
        elif isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            bodies = list(n.args.defaults) + [d for d in n.args.kw_defaults if d is not None] + list(n.decorator_list)
        else:
            continue
        refs = set()
        for b in bodies:
            refs |= {x.id for x in ast.walk(b) if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load)}
        bad = sorted(r for r in refs if r not in body_names and r not in THIRD and not r.startswith("__"))
        # 僅限 transformer 自身頂層名稱才算數（排除 pd 等第三方誤判已由 THIRD 覆蓋）
        bad = [r for r in bad if r in origin]
        if bad:
            errors.append(f"D. transformer L{n.lineno}（{targets or n.name}）定義期引用墊片名稱：{bad}")

    # E. 一致性（WARNING）
    fresh = {}
    for modfile in sorted((REPO / "modes").glob("*.py")):
        if modfile.name == "__init__.py":
            continue
        for x in mod_names(modfile) - {"__all__"}:
            fresh[x] = modfile.stem
    for x, mod in sorted(fresh.items()):
        if x not in origin:
            warnings.append(f"E. {x}（{mod}）未登記於 moved.json")
    for x, mod in sorted(origin.items()):
        if x not in fresh:
            warnings.append(f"E. moved.json 有 {x}（{mod}）但模組內無此定義")

    for w in warnings:
        print("WARNING", w)
    for e in errors:
        print("ERROR", e)
    print(f"check done: {len(errors)} errors, {len(warnings)} warnings")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
