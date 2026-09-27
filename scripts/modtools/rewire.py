"""rewire.py：重算全部 modes/*.py 的引用缺口，按來源分流重建 import 塊。

needed = 載入 − 本地定義（含參數／賦值／推導變數） − 第三方／內建
來源：origin map → modes.X；仍在 transformer 主體 → transformer；其他 → 報錯。
用法: python rewire.py [--check]  (--check 只報告不寫入)
"""
import ast
import builtins
import json
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parents[2]
STATE = pathlib.Path(__file__).resolve().parent / "moved.json"
CHECK = "--check" in sys.argv

THIRD = {
    "pd", "np", "gc", "re", "unicodedata", "datetime", "BytesIO",
    "ThreadPoolExecutor", "Workbook", "Font", "Alignment", "Border",
    "Side", "PatternFill", "get_column_letter", "DataValidation", "st",
} | set(dir(builtins))

SKIP = {"__init__.py"}


def local_names(tree):
    out = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            out.add(node.name)
            args = node.args
        elif isinstance(node, ast.Lambda):
            args = node.args
        else:
            args = None
        if args is not None:
            for x in list(args.posonlyargs) + list(args.args) + list(args.kwonlyargs):
                out.add(x.arg)
            if args.vararg:
                out.add(args.vararg.arg)
            if args.kwarg:
                out.add(args.kwarg.arg)
            continue
        elif isinstance(node, ast.ClassDef):
            out.add(node.name)
        elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            out.add(node.id)
        elif isinstance(node, (ast.Import,)):
            for al in node.names:
                out.add((al.asname or al.name).split(".")[0])
        elif isinstance(node, ast.ExceptHandler) and node.name:
            out.add(node.name)
    return out


def main():
    origin = json.loads(STATE.read_text(encoding="utf-8"))
    tbody = ast.parse((REPO / "transformer.py").read_text(encoding="utf-8-sig"))
    body_names = set()
    for n in tbody.body:
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            body_names.add(n.name)
        elif isinstance(n, ast.Assign):
            for tg in n.targets:
                for x in ast.walk(tg):
                    if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store):
                        body_names.add(x.id)
    errors = []
    for modfile in sorted((REPO / "modes").glob("*.py")):
        if modfile.name in SKIP:
            continue
        t = modfile.read_text(encoding="utf-8-sig")
        # 剝除舊引用塊後再分析
        t2 = re.sub(r"from transformer import \(.*?\)", "", t, flags=re.S)
        t2 = re.sub(r"from modes\.\w+ import \(.*?\)", "", t2, flags=re.S)
        m = ast.parse(t2)
        local = local_names(m)
        loads = {x.id for x in ast.walk(m) if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Load)}
        needed = sorted(x for x in loads - local - THIRD if not x.startswith("__"))
        groups, unknown = {}, []
        for x in needed:
            if x in origin:
                groups.setdefault(origin[x], []).append(x)
            elif x in body_names:
                groups.setdefault("transformer", []).append(x)
            else:
                unknown.append(x)
        if unknown:
            errors.append((modfile.name, unknown))
        print(f"{modfile.name}: " + (", ".join(f"{g}({len(v)})" for g, v in sorted(groups.items())) or "(none)"))
        if unknown:
            print(f"  UNRESOLVED: {unknown}")
        if CHECK:
            continue
        blocks = []
        for g in sorted(groups, key=lambda g: (g != "transformer", g)):
            if g == "transformer":
                blocks.append("from transformer import (\n" + "".join(f"    {x},\n" for x in sorted(groups[g])) + ")")
            else:
                blocks.append(f"from modes.{g} import (\n" + "".join(f"    {x},\n" for x in sorted(groups[g])) + ")")
        mt2 = ast.parse(t2)
        ins = 0
        for n in mt2.body:
            if isinstance(n, (ast.Import, ast.ImportFrom)):
                ins = max(ins, n.end_lineno)
        tlines = t2.splitlines(keepends=True)
        if blocks:
            ins_text = "\n\n" + "\n\n\n".join(blocks) + "\n"
            t2 = "".join(tlines[:ins]).rstrip("\n") + "\n" + ins_text + "".join(tlines[ins:]).lstrip("\n")
        modfile.write_text(t2, encoding="utf-8-sig")
        ast.parse(t2)
    if errors:
        print("ERRORS:", errors)
        sys.exit(1)
    print("rewire OK")


if __name__ == "__main__":
    main()
