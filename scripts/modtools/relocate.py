# -*- coding: utf-8 -*-
"""relocate.py：將名稱由一個 modes 模組搬至另一個（切斷循環引用或修正歸屬）。

用法：
    python scripts/modtools/relocate.py <名稱> <來源模組> <目標模組> [<名稱> <來源模組> <目標模組> ...]
    python scripts/modtools/relocate.py _write_guide_sheet safety_review common

搬運後自動更新雙方 __all__ 與 moved.json；完成後請執行 rewire.py 與測試。
"""
import ast
import json
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parents[2]
STATE = pathlib.Path(__file__).resolve().parent / "moved.json"


def main():
    args = sys.argv[1:]
    assert len(args) >= 3 and len(args) % 3 == 0, "用法：relocate.py <名稱> <來源> <目標> [...]"
    moves = [(args[i], args[i + 1], args[i + 2]) for i in range(0, len(args), 3)]
    state = json.loads(STATE.read_text(encoding="utf-8"))
    for nm, src_mod, dst_mod in moves:
        sp = REPO / "modes" / f"{src_mod}.py"
        dp = REPO / "modes" / f"{dst_mod}.py"
        assert sp.exists() and dp.exists(), f"模組不存在：{src_mod}／{dst_mod}"
        st = sp.read_text(encoding="utf-8-sig")
        sm = ast.parse(st)
        found = [n for n in sm.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Assign))
                 and nm in ([n.name] if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
                            else [x.id for t in n.targets for x in ast.walk(t)
                                  if isinstance(x, ast.Name) and isinstance(x.ctx, ast.Store)])]
        assert found, f"{nm} 不在 {src_mod}"
        node = found[0]
        lines = st.splitlines(keepends=True)
        s = node.lineno - 1
        while s > 0 and lines[s - 1].strip().startswith("#"):
            s -= 1
        seg = "".join(lines[s:node.end_lineno])
        rest = "".join(lines[:s] + lines[node.end_lineno:])
        rest = re.sub(rf'^    "{nm}",\n', "", rest, flags=re.M)
        sp.write_text(rest, encoding="utf-8-sig")
        dt = dp.read_text(encoding="utf-8-sig") + "\n\n\n" + seg
        m = re.search(r"__all__ = \[(.*?)\]", dt, re.S)
        assert m, f"__all__ 不在 {dst_mod}"
        dt = dt[:m.end() - 1] + f'    "{nm}",\n' + dt[m.end() - 1:]
        dp.write_text(dt, encoding="utf-8-sig")
        state[nm] = dst_mod
        print(f"relocate {nm}: {src_mod} -> {dst_mod}")
    STATE.write_text(json.dumps(state, indent=1, ensure_ascii=False), encoding="utf-8")
    for _, src_mod, dst_mod in moves:
        for mod in {src_mod, dst_mod}:
            ast.parse((REPO / "modes" / f"{mod}.py").read_text(encoding="utf-8-sig"))
    print("parse OK；請接著執行 rewire.py 與完整測試")


if __name__ == "__main__":
    main()
