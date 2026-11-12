# -*- coding: utf-8 -*-
"""轉貨接收店集中規範（Transfer Receiver Concentration Rules）— 零依賴參考實作。

本檔是「調貨接收店集中」規則的唯一可執行來源，只用標準庫，不依賴 pandas／
Streamlit／Excel，可整檔複製到其他調貨程式使用。

三層結構：
1. 政策 ``ReceiverPolicy``：可調參數——每間目標件數、件數上限階梯、個別分類覆寫、
   每間保底件數、封頂天數。改參數不必改演算法。
2. 演算法：純函數，輸入件數與權重，輸出整數件，無副作用。
   ``allocate_proportional``（最大餘數法）／``allocate_concentrated``（先保底再按比例）／
   ``allocate_with_caps``（按比例但受每店上限約束）。
3. 入口 ``plan_allocation``：一次完成「選店 → 分配 → 封頂」，回傳 ``Allocation``。

規則的完整文字說明（含 pseudocode、不變量、邊界情況、跨專案接入方式）見
``TRANSFER_RULES_SPEC.md``；驗收向量見 ``transfer_rules_golden.json``，
任何語言都可拿它驗證自己的實作是否與本規範一致。

最小用法::

    from transfer_rules import plan_allocation

    result = plan_allocation(
        13,                                    # 同 SKU 轉出總量（件）
        [("S02", 9, 0), ("S03", 7, 0), ("S04", 5, 0)],   # (Site Code, 銷售, 調前庫存)
    )
    result.lines        # (("S02", 3), ("S03", 3), ("S04", 3), ("S02", 4)) 之類的分配結果
    result.allocated    # 13
    result.unallocated  # 0

只想要數字、自己處理資料管線時::

    from transfer_rules import DEFAULT_POLICY, select_pool, allocate

    pool, dropped = select_pool(candidates, total=13, mode="auto", category="0302 - MINIATURE")
    alloc, unalloc = allocate(total=13, weights=[r[1] for r in pool], mode="auto")
"""
from __future__ import annotations

import hashlib
import json
import math
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from types import MappingProxyType

__all__ = [
    "RULES_VERSION",
    "MODE_AUTO",
    "MODE_UNLIMITED",
    "ReceiverPolicy",
    "DEFAULT_POLICY",
    "Allocation",
    "allocate",
    "allocate_concentrated",
    "allocate_proportional",
    "allocate_with_caps",
    "daily_demand",
    "normalize_mode",
    "plan_allocation",
    "policy_from_dict",
    "receiver_rule_table",
    "select_pool",
]

# 規範版本：規則語意有變（會令同一批資料得出不同分配）時必須遞增，
# 並同步更新 TRANSFER_RULES_SPEC.md 與 transfer_rules_golden.json。
RULES_VERSION = "1.0"

MODE_AUTO = "auto"
MODE_UNLIMITED = "unlimited"

DEFAULT_BASE_DAYS = 90  # 銷售量口徑：Sold Qty (P90D) 為 90 日銷量
DEFAULT_CAP_DAYS = 14  # 封頂口徑：每店最多補到 14 日需求


# --------------------------------------------------------------------------------------
# 政策：所有可調參數集中一處
# --------------------------------------------------------------------------------------
@dataclass(frozen=True)
class ReceiverPolicy:
    """集中規則的參數集合。所有欄位皆為純量或不可變映射，可安全共用。

    ``per_shop_target``
        每間接收店的目標件數（預設 3）。接收店數＝ceil(轉出總量 ÷ 目標件數)。
    ``target_by_category``
        個別分類的目標件數覆寫，鍵為分類代碼前綴（預設 ``{"0302": 4}``，
        即 ``0302 - MINIATURE`` 迷你裝每間 4 件）。以 ``str.startswith`` 比對，
        依插入順序取第一個命中；無命中則用 ``per_shop_target``。
    ``cap_from`` / ``base_cap`` / ``cap_step``
        件數上限階梯：``cap_from`` 件以下不封頂（由每間目標件數決定）；其後到
        ``cap_step`` 件為止最多 ``base_cap`` 間；超過 ``cap_step`` 每 ``cap_step``
        件加 1 間。預設 10／5／30 → 10–30 件最多 5 間、31–60 件 6 間、61–90 件 7 間。
    ``min_per_receiver``
        集中分配的每間保底件數（預設 2），用於消除「只收 1 件」的碎片；
        件數不足時退化為均分，不會超額。
    """

    per_shop_target: int = 3
    target_by_category: Mapping[str, int] = field(
        default_factory=lambda: MappingProxyType({"0302": 4})
    )
    cap_from: int = 10
    base_cap: int = 5
    cap_step: int = 30
    min_per_receiver: int = 2

    # -- 參數查詢 ------------------------------------------------------------------
    def target_for(self, category: str = "") -> int:
        """回傳該分類的每間目標件數；分類為空或無覆寫時用預設值。"""
        text = str(category or "").strip()
        for prefix, target in self.target_by_category.items():
            if text.startswith(str(prefix)):
                return int(target) or int(self.per_shop_target)
        return int(self.per_shop_target)

    def cap(self, total: int) -> int:
        """回傳該轉出總量下的接收店數上限（件數太少時至少 1，不會回 0）。"""
        total = int(total)
        if total < int(self.cap_from):
            return max(total, 1)
        if total <= int(self.cap_step):
            return int(self.base_cap)
        return int(self.base_cap) + math.ceil(
            (total - int(self.cap_step)) / int(self.cap_step)
        )

    def receiver_count(self, total: int, category: str = "", target: int | None = None) -> int:
        """依轉出總量決定接收店數：ceil(總量 ÷ 目標件數)，再受件數上限封頂。

        ``target`` 有值時優先於分類覆寫（呼叫端已有既定目標件數的情況）。
        """
        total = int(total)
        if total <= 0:
            return 0
        per_shop = int(target or 0) or self.target_for(category)
        if per_shop <= 0:
            per_shop = int(self.per_shop_target)
        wanted = max(1, math.ceil(total / per_shop))
        return min(wanted, self.cap(total))

    # -- 可追溯性 ------------------------------------------------------------------
    def as_dict(self) -> dict:
        """參數的 JSON 友好表示（鍵排序，供存檔／比對）。"""
        return {
            "per_shop_target": int(self.per_shop_target),
            "target_by_category": {str(k): int(v) for k, v in self.target_by_category.items()},
            "cap_from": int(self.cap_from),
            "base_cap": int(self.base_cap),
            "cap_step": int(self.cap_step),
            "min_per_receiver": int(self.min_per_receiver),
        }

    def fingerprint(self) -> str:
        """參數指紋（8 位十六進位）。

        把指紋寫進轉貨單／log，日後可確認某張單是用哪一組參數跑出來的。
        參數一改指紋即變，不用比對整份設定檔。
        """
        payload = json.dumps(self.as_dict(), ensure_ascii=False, sort_keys=True,
                             separators=(",", ":"))
        return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:8]

    def describe(self, category: str = "", upto: int = 90) -> str:
        """一句可讀的規則摘要（適合放進 UI 說明或報表頁首）。

        由 ``format_rule_table`` 產生，數字與實際行為同源，不會人手抄錯。
        """
        bands = format_rule_table(self, category=category, upto=upto)
        return "；".join(bands) + f"。每間保底 {self.min_per_receiver} 件。"

    # -- 衍生政策 ------------------------------------------------------------------
    def with_overrides(self, **changes) -> "ReceiverPolicy":
        """回傳套用個別參數後的新政策（原物件不變）。"""
        return replace(self, **changes)


DEFAULT_POLICY = ReceiverPolicy()


def policy_from_dict(data: Mapping | None) -> ReceiverPolicy:
    """由 dict（例如 ``params.json``）建立政策；缺欄用預設值，未知鍵忽略。"""
    if not data:
        return DEFAULT_POLICY
    src = data.get("policy", data) if isinstance(data, Mapping) else {}
    allowed = {
        "per_shop_target", "target_by_category", "cap_from", "base_cap",
        "cap_step", "min_per_receiver",
    }
    kwargs = {k: v for k, v in src.items() if k in allowed and v is not None}
    if "target_by_category" in kwargs:
        kwargs["target_by_category"] = MappingProxyType(
            {str(k): int(v) for k, v in dict(kwargs["target_by_category"]).items()}
        )
    return ReceiverPolicy(**kwargs)


# --------------------------------------------------------------------------------------
# 模式
# --------------------------------------------------------------------------------------
def normalize_mode(mode) -> str | int | None:
    """把各種寫法統一成 ``"auto"``／``None``（不限）／正整數（固定上限）。

    接受 ``"auto"``（大小寫、前後空白不拘）、``None``／``"unlimited"``／``"none"``／
    ``""``（不限）、正整數字串（固定上限）。其他值一律 ``ValueError``。
    """
    if mode is None:
        return None
    if isinstance(mode, str):
        text = mode.strip()
        if text.lower() == MODE_AUTO:
            return MODE_AUTO
        if text.lower() in (MODE_UNLIMITED, "none", "no_limit", ""):
            return None
        try:
            value = int(text)
        except (TypeError, ValueError):
            raise ValueError(
                '接收店舖集中上限必須為正整數、"auto" 或「不限」。'
            ) from None
    elif isinstance(mode, bool):
        raise ValueError('接收店舖集中上限必須為正整數、"auto" 或「不限」。')
    else:
        # 非整數的浮點數（如 3.5）一律視為錯誤，避免悄悄截斷成 3 間
        if isinstance(mode, float) and not float(mode).is_integer():
            raise ValueError('接收店舖集中上限必須為正整數、"auto" 或「不限」。')
        try:
            value = int(mode)
        except (TypeError, ValueError):
            raise ValueError(
                '接收店舖集中上限必須為正整數、"auto" 或「不限」。'
            ) from None
    if value <= 0:
        raise ValueError('接收店舖集中上限必須為正整數、"auto" 或「不限」。')
    return value


def _mode_label(mode: str | int | None) -> str:
    if mode == MODE_AUTO:
        return "auto"
    if mode is None:
        return MODE_UNLIMITED
    return f"fixed:{int(mode)}"


# --------------------------------------------------------------------------------------
# 候選接收店
# --------------------------------------------------------------------------------------
Candidate = tuple[str, float, float]  # (Site Code, 銷售量, 調前庫存)


def _normalize_candidates(candidates: Iterable) -> list[Candidate]:
    """接受 ``(site, sales)``／``(site, sales, stock)`` 或含同名鍵的 mapping。

    缺庫存視為 0；庫存只影響封頂，不影響選店排序。
    """
    rows: list[Candidate] = []
    for item in candidates or ():
        if isinstance(item, Mapping):
            site = str(item.get("site", item.get("Site Code", ""))).strip()
            sales = float(item.get("sales", item.get("Sold Qty (P90D)", 0)) or 0)
            stock = float(item.get("stock", item.get("Shop Stock Qty (CM)", 0)) or 0)
        else:
            seq = tuple(item)
            if len(seq) < 2:
                raise ValueError("候選接收店至少需要 (Site Code, 銷售量) 兩項。")
            site = str(seq[0]).strip()
            sales = float(seq[1] or 0)
            stock = float(seq[2]) if len(seq) > 2 and seq[2] is not None else 0.0
        if site:
            rows.append((site, sales, stock))
    return rows


def _rank_by_sales(rows: Sequence[Candidate]) -> list[Candidate]:
    """選店次序：銷售量由高至低，同銷量以 Site Code 升序（確保結果可重現）。"""
    return sorted(rows, key=lambda r: (-r[1], r[0]))


def select_pool(
    candidates: Iterable,
    total: int,
    *,
    mode="auto",
    category: str = "",
    policy: ReceiverPolicy | None = None,
) -> tuple[list[Candidate], int]:
    """依模式挑出實際參與分配的接收店，回傳 ``(pool, dropped)``。

    ``dropped`` 為候選店中因集中而未納入的數目，供報表列「未納入店舖數」。

    - ``mode`` 為「不限」：回傳全部候選店，按 Site Code 升序（與未啟用集中時一致）。
    - ``mode`` 為 ``"auto"``：店數＝``policy.receiver_count(total, category)``。
    - ``mode`` 為正整數：店數＝該上限。
    兩種集中模式都按銷售量高至低取前 N 間（同銷量以 Site Code 升序）。
    """
    rows = _normalize_candidates(candidates)
    resolved = normalize_mode(mode)
    if not rows:
        return [], 0
    if resolved is None:
        return sorted(rows, key=lambda r: r[0]), 0
    resolved_policy = policy or DEFAULT_POLICY
    take = (
        resolved_policy.receiver_count(int(total), category=category)
        if resolved == MODE_AUTO
        else int(resolved)
    )
    pool = _rank_by_sales(rows)[: max(take, 0)]
    return pool, len(rows) - len(pool)


# --------------------------------------------------------------------------------------
# 分配演算法（純函數）
# --------------------------------------------------------------------------------------
def allocate_proportional(total: int, weights: Sequence[float]) -> list[int]:
    """按權重比例分配整數件（最大餘數法）；權重總和 > 0 時加總恆等於 ``total``。

    同餘數時取權重較大者，再相同則取索引較小者——純決定性，同一輸入必得同一輸出。
    ``weights`` 總和為 0 或負值時退化為均分。
    """
    n = len(weights)
    if n == 0 or total <= 0:
        return [0] * n
    wsum = float(sum(weights))
    if wsum <= 0:
        weights = [1.0] * n
        wsum = float(n)
    raw = [total * w / wsum for w in weights]
    alloc = [math.floor(x) for x in raw]
    rest = max(total - sum(alloc), 0)
    order = sorted(range(n), key=lambda i: (-(raw[i] - alloc[i]), -weights[i], i))
    for i in order[:rest]:
        alloc[i] += 1
    return alloc


def allocate_concentrated(
    total: int,
    weights: Sequence[float],
    min_per_receiver: int = 2,
) -> list[int]:
    """集中模式分配：每間先保底 ``min_per_receiver`` 件（件數不足時均分），
    餘量再按權重以最大餘數法分配。

    與純比例分配的分別：高銷量店仍分得較多，但不會出現某店只分到 1 件的碎片。
    ``sum(結果) == total`` 恆成立。
    """
    n = len(weights)
    if n == 0 or total <= 0:
        return [0] * n
    base = min(int(min_per_receiver), int(total) // n)
    alloc = [base] * n
    rest = int(total) - base * n
    if rest > 0:
        extra = allocate_proportional(rest, weights)
        alloc = [a + e for a, e in zip(alloc, extra)]
    return alloc


def allocate_with_caps(
    total: int,
    weights: Sequence[float],
    caps: Sequence[int],
) -> tuple[list[int], int]:
    """按權重比例分配，惟每個接收店不得超出上限 ``caps``；超出上限的餘量在
    尚有需求的店舖之間重新按權重分配，直至派完或全部到頂。
    回傳 ``(alloc, unallocated)``。

    ``alloc[i] <= caps[i]`` 與 ``sum(alloc) + unallocated == total`` 恆成立。
    """
    n = len(weights)
    alloc = [0] * n
    if n == 0 or total <= 0:
        return alloc, max(int(total), 0)
    remaining = int(total)
    active = [i for i in range(n) if caps[i] > 0]
    while remaining > 0 and active:
        share = allocate_proportional(remaining, [weights[i] for i in active])
        moved = 0
        nxt: list[int] = []
        for i, s in zip(active, share):
            room = caps[i] - alloc[i]
            take = min(s, room)
            if take > 0:
                alloc[i] += take
                moved += take
            if caps[i] > alloc[i]:
                nxt.append(i)
        remaining -= moved
        if moved == 0:
            break
        active = nxt
    return alloc, remaining


def allocate(
    total: int,
    weights: Sequence[float],
    *,
    mode="auto",
    caps: Sequence[int] | None = None,
    min_per_receiver: int = 2,
) -> tuple[list[int], int]:
    """依模式分配並回傳 ``(alloc, unallocated)``。

    - ``caps`` 有值（封頂開啟）：走 ``allocate_with_caps``，與模式無關；保底讓位於上限。
    - ``mode == "auto"``：走 ``allocate_concentrated``（每間保底）。
    - 其他（不限／固定上限）：走 ``allocate_proportional``。
    """
    n = len(weights)
    if n == 0 or total <= 0:
        return [0] * n, max(int(total), 0)
    if caps is not None:
        return allocate_with_caps(total, weights, caps)
    resolved = normalize_mode(mode)
    if resolved == MODE_AUTO:
        alloc = allocate_concentrated(total, weights, min_per_receiver)
        return alloc, max(int(total) - sum(alloc), 0)
    alloc = allocate_proportional(total, weights)
    return alloc, max(int(total) - sum(alloc), 0)


def daily_demand(
    sales: float,
    stock: float,
    days: int | None = DEFAULT_CAP_DAYS,
    base_days: int = DEFAULT_BASE_DAYS,
) -> int:
    """每店需求上限：``max(0, ceil(銷售 ÷ 90 × days) − 調前庫存)``。

    與調貨段「調入需求量（補到 14 日）」同一公式：先換算日均銷量，乘目標日數後
    無條件進位，再扣調前庫存。``days`` 為 ``None`` 時回 0（不封頂，由呼叫端忽略）。
    """
    if days is None:
        return 0
    need = math.ceil(float(sales) / float(base_days) * int(days)) - int(round(float(stock)))
    return max(need, 0)


# --------------------------------------------------------------------------------------
# 入口：一次完成選店、分配、封頂
# --------------------------------------------------------------------------------------
@dataclass(frozen=True)
class Allocation:
    """一次分配的結果，足夠讓呼叫端直接產生轉貨單。"""

    mode: str  # "auto" / "unlimited" / "fixed:N"
    lines: tuple[tuple[str, int], ...]  # 入選店舖與分配量（依分配序，含 0 件者）
    allocated: int
    unallocated: int
    dropped: int  # 候選店中因集中而未納入的數目
    rules_version: str = RULES_VERSION
    policy_fingerprint: str = ""

    @property
    def quantities(self) -> dict[str, int]:
        """``{Site Code: 分配量}``；只含分配量大於 0 的店舖。"""
        return {site: qty for site, qty in self.lines if qty > 0}

    @property
    def receivers(self) -> tuple[str, ...]:
        """實際收到貨的店舖（分配量大於 0）。"""
        return tuple(site for site, qty in self.lines if qty > 0)

    def summary(self) -> str:
        """一句可讀摘要，適合寫入 log 或報表備註。"""
        return (
            f"{self.mode}：{len(self.receivers)} 間接收店，已分配 {self.allocated} 件"
            f"，未分配 {self.unallocated} 件，集中後未納入 {self.dropped} 間"
            f"（規則 v{self.rules_version}／{self.policy_fingerprint or '-'}）"
        )


def plan_allocation(
    total: int,
    candidates: Iterable,
    *,
    mode="auto",
    category: str = "",
    policy: ReceiverPolicy | None = None,
    cap_days: int | None = None,
    base_days: int = DEFAULT_BASE_DAYS,
) -> Allocation:
    """給定同 SKU 轉出總量與候選接收店，回傳完整分配結果。

    參數
    ----
    total
        同 SKU 的轉出總量（件）。呼叫端負責先做好前置過濾——同 SKU、非轉出店、
        有銷售記錄（``Sold Qty (P90D) > 0``）、必要時排除 ND 店。
    candidates
        ``(Site Code, 銷售量, 調前庫存)`` 或 ``(Site Code, 銷售量)``，亦可傳
        ``{"site":…, "sales":…, "stock":…}``。庫存只在 ``cap_days`` 有值時使用。
    mode
        ``"auto"``（按件數自動集中）、``None``／``"unlimited"``（全數按比例）、
        正整數（固定上限）。
    category
        分類代碼（例如 MCH2 的 ``"0302 - MINIATURE"``），用於套用分類覆寫的目標件數。
    cap_days
        ``None``＝不封頂；給 14 即「每店最多補到 14 日需求」。封頂與集中可併用，
        此時保底讓位於上限，派不完的部分列 ``unallocated``。
    """
    resolved_policy = policy or DEFAULT_POLICY
    total = int(total)
    rows = _normalize_candidates(candidates)
    if total <= 0 or not rows:
        return Allocation(
            mode=_mode_label(normalize_mode(mode)),
            lines=(),
            allocated=0,
            unallocated=max(total, 0),
            dropped=0,
            policy_fingerprint=resolved_policy.fingerprint(),
        )

    pool, dropped = select_pool(
        rows, total, mode=mode, category=category, policy=resolved_policy
    )
    weights = [r[1] for r in pool]
    caps = (
        [daily_demand(r[1], r[2], cap_days, base_days) for r in pool]
        if cap_days is not None
        else None
    )
    alloc, unalloc = allocate(
        total,
        weights,
        mode=mode,
        caps=caps,
        min_per_receiver=resolved_policy.min_per_receiver,
    )
    return Allocation(
        mode=_mode_label(normalize_mode(mode)),
        lines=tuple((r[0], int(a)) for r, a in zip(pool, alloc)),
        allocated=int(sum(alloc)),
        unallocated=int(unalloc),
        dropped=int(dropped),
        policy_fingerprint=resolved_policy.fingerprint(),
    )


# --------------------------------------------------------------------------------------
# 文件／UI 用：由參數自動產生件數分檔表
# --------------------------------------------------------------------------------------
def receiver_rule_table(
    policy: ReceiverPolicy | None = None,
    category: str = "",
    upto: int = 120,
) -> list[tuple[int, int | None, int]]:
    """產生「件數帶 → 接收店數」分檔表，回傳 ``[(低, 高或 None, 店數), …]``。

    由 ``receiver_count`` 逐件掃描歸納，因此說明頁、文件、UI 的數字不會與程式
    實際行為脫節；``高`` 為 ``None`` 代表 ``upto`` 以上持續（每階段加 1 間）。

    只應在 ``upto`` 之後仍是固定遞增時才把最後一列視為開放區間。
    """
    resolved = policy or DEFAULT_POLICY
    upto = max(int(upto), 1)
    bands: list[tuple[int, int | None, int]] = []
    start = 1
    prev = resolved.receiver_count(1, category=category)
    for value in range(2, upto + 1):
        current = resolved.receiver_count(value, category=category)
        if current != prev:
            bands.append((start, value - 1, prev))
            start, prev = value, current
    step = int(resolved.cap_step)
    grows = resolved.receiver_count(upto + step, category=category) > prev
    bands.append((start, None if grows else upto, prev))
    return bands


def format_rule_table(
    policy: ReceiverPolicy | None = None,
    category: str = "",
    upto: int = 120,
) -> list[str]:
    """把 ``receiver_rule_table`` 格式化為 ``["3 件或以下 → 1 間", "4–6 件 → 2 間", …]``。"""
    out: list[str] = []
    for low, high, count in receiver_rule_table(policy, category, upto):
        if high is None:
            out.append(f"{low} 件以上 → 每 {int((policy or DEFAULT_POLICY).cap_step)} 件加 1 間（目前 {count} 間）")
        elif low == high:
            out.append(f"{low} 件 → {count} 間")
        elif low == 1:
            out.append(f"{high} 件或以下 → {count} 間")
        else:
            out.append(f"{low}–{high} 件 → {count} 間")
    return out
