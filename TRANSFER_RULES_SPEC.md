# 轉貨接收店集中規範 v1.0

> Transfer Receiver Concentration Spec — 調貨程式共用規則
> 版本：1.0　｜　規則指紋：`0bb5cf16`　｜　日期：2026-11-12
> 來源專案：I14 report review v1.95.4（P90D 0 Sales Review／ND SKU with Stock Review 轉貨模式）

---

## 0. 這份文件連同兩個檔案，就是「規則本身」

規則一旦只存在於對話、UI 文案或某個函數裡，換一個專案就會走樣。所以規則用三件套表達：

| 檔案 | 角色 | 給誰看 |
|---|---|---|
| `TRANSFER_RULES_SPEC.md`（本文件） | 規範：定義、公式、邊界、驗收條件 | 人（PM／店舖營運／工程師） |
| `transfer_rules.py` | 可執行參考實作：零依賴，純標準庫 | Python 專案直接 import |
| `transfer_rules_golden.json` | 781 個驗收向量（輸入 → 期望輸出） | 任何語言；用測試釘住行為 |

三者關係：**文件說規則，程式跑規則，向量證明兩者一致**。任何一方改動，另外兩方必須同步——
`scripts/gen_transfer_rules_golden.py` 會把向量與引擎重跑一次，不一致就報錯。

**為甚麼不用「寫一段文字描述」就算**：這套規則有 6 個互相牽動的參數（每間目標件數、分類覆寫、
上限起點、上限基數、上限階距、保底件數），文字版本一定會在某次微調後與程式脫節。
參數化的政策物件 + 自動產生的分檔表（`format_rule_table()`），才不會出現「文件寫 3 件、程式跑 4 件」。

---

## 1. 適用範圍

**適用**：任何「指定轉出店舖 → 把同 SKU 的貨分配給其他有銷售記錄店舖」的調貨程式。
本規範只管**接收店要開幾間、每間分幾件**。

**不適用**（屬呼叫端職責，各專案可不同）：

| 項目 | 由誰決定 |
|---|---|
| 哪些店舖可轉出（0 Sales／RP Type＝ND／庫存門檻） | 呼叫端 |
| 候選接收店的過濾（同 SKU、非轉出店、`Sold Qty (P90D) > 0`、排除 ND 店） | 呼叫端 |
| 轉出總量（同 SKU 各轉出店庫存合計） | 呼叫端 |
| 轉出店 ↔ 接收店的配對次序 | 呼叫端（建議見 §9.4） |
| 是否啟用「補到 14 日」封頂 | 呼叫端傳 `cap_days` |

換句話說：**你只要交出「總量 + 一份合格候選店清單」，本規範保證給你「哪幾間、各幾件」**。

---

## 2. 名詞

| 名詞 | 定義 |
|---|---|
| 轉出總量 `total` | 同一個 SKU 在所有已選轉出店舖的可轉出件數合計（整數） |
| 候選接收店 `candidates` | `(Site Code, 銷售量, 調前庫存)`；銷售量＝`Sold Qty (P90D)`，只作權重與排序 |
| 目標件數 `target` | 每間接收店期望收幾件（預設 3；迷你裝 4） |
| 店數上限 `cap` | 該轉出總量下允許的最大接收店數 |
| 保底 `min_per_receiver` | 集中分配時每間至少幾件（預設 2） |
| 封頂 `cap_days` | 每店最多補到幾天需求（`14` 即「補到 14 日」；`None`＝不封頂） |

---

## 3. 規範（依序執行）

### R1　模式

| `mode` | 語意 | 選店規則 |
|---|---|---|
| `"auto"` | 按件數自動集中（**建議預設**） | R2–R6 |
| `None`／`"unlimited"` | 不限，全數候選店按比例分配 | 全部候選店，按 Site Code 升序 |
| 正整數 `N` | 固定上限 | 銷售量最高的 N 間 |

### R2　每間目標件數

```
target = target_by_category 中第一個命中的前綴（如 "0302" → 4）
         否則 per_shop_target（預設 3）
```
分類比對用 `startswith`，依設定順序取第一個命中；分類字串為空或無命中即用預設。
`0302 - MINIATURE`（迷你裝）＝ 4 件；`0303 - MINIATURE SET` 不命中，仍用 3 件。

### R3　理想店數

```
wanted = max(1, ceil(total / target))
```

### R4　店數上限

```
if total <  cap_from:            cap = max(total, 1)
elif total <= cap_step:          cap = base_cap
else:                            cap = base_cap + ceil((total - cap_step) / cap_step)
```
預設 `cap_from=10, base_cap=5, cap_step=30` → 10–30 件最多 5 間、31–60 件 6 間、61–90 件 7 間。

### R5　實際接收店數

```
take = min(wanted, cap)        # total <= 0 時 take = 0
```
件數上升時 `wanted` 線性增長，但 `cap` 壓住它——這是「10 件以上不再越分越多」的關鍵。

### R6　選店

```
pool = 候選店按 (銷售量 desc, Site Code asc) 排序後取前 take 間
```
排序必須穩定且可重現；**同銷量以 Site Code 升序**是規範的一部分，不可用任意次序。
候選店少於 `take` 時全數取用（不補、不重複）。

### R7　分配（未封頂）

- 模式 `"auto"`：**先保底，餘量按比例**
  ```
  base  = min(min_per_receiver, total // n)     # n = 入選店數
  alloc = [base] * n
  rest  = total - base * n
  alloc += 最大餘數法(rest, weights=銷售量)
  ```
- 其他模式：直接 `最大餘數法(total, weights)`

最大餘數法（`allocate_proportional`）：
```
raw[i]  = total * w[i] / sum(w)          # sum(w) <= 0 時視為等權
alloc[i]= floor(raw[i])
剩餘 rest 件，按 (raw 小數部分 desc, w desc, 索引 asc) 逐件補 1
```

### R8　分配（封頂，`cap_days` 有值）

```
cap[i] = max(0, ceil(銷售量[i] / 90 × cap_days) − round(調前庫存[i]))
```
按比例分配後，超出 `cap[i]` 的餘量在**仍有空間**的店舖之間按權重重新分配，
直至派完或全部到頂；仍未派完的列為**未分配**（保留原店）。
**封頂優先於保底**：此模式下不走 R7 的保底邏輯。

### R9　輸出

```
lines       = [(Site Code, 分配量), …]     # 次序＝入選次序，含 0 件者
allocated   = sum(分配量)
unallocated = total − allocated
dropped     = 候選店數 − 入選店數            # 因集中而未納入的店數
```

---

## 4. 不變量（可直接寫成斷言）

| # | 不變量 |
|---|---|
| I1 | `allocated + unallocated == total` |
| I2 | 未封頂時 `unallocated == 0`（候選店權重總和 > 0 或全為 0 皆成立） |
| I3 | 封頂時 `0 <= alloc[i] <= cap[i]` |
| I4 | `auto` 模式下，若 `total >= n × min_per_receiver`，則每間 `alloc[i] >= min_per_receiver`（不出現 1 件碎片） |
| I5 | `take <= cap(total)`，且 `len(lines) <= len(candidates)` |
| I6 | 決定性：同一輸入（含候選店次序）必得同一輸出；同銷量以 Site Code 升序 |
| I7 | `receiver_count(total, category)` 對 `total` 單調不減（件數增加，店數不會減少） |

---

## 5. 參數表

| 參數 | 預設 | 作用 | 調大 → | 調小 → |
|---|---|---|---|---|
| `per_shop_target` | 3 | 每間目標件數 | 店數更少、每店更多 | 店數更多、每店更少 |
| `target_by_category` | `{"0302": 4}` | 分類覆寫 | 該類更集中 | 該類更分散 |
| `cap_from` | 10 | 上限開始生效的件數 | 更多件數不受上限約束 | 更早收緊 |
| `base_cap` | 5 | 第一段上限店數 | 10–30 件可分更多店 | 更集中 |
| `cap_step` | 30 | 每加一店的件數階距 | 大件數時擴張更慢 | 擴張更快 |
| `min_per_receiver` | 2 | 每間保底件數 | 碎片更少、小單更集中 | 更貼近純比例 |

**調參不用改程式**：

```python
from transfer_rules import DEFAULT_POLICY, plan_allocation

policy = DEFAULT_POLICY.with_overrides(base_cap=4, per_shop_target=4)   # 更集中
result = plan_allocation(total, candidates, mode="auto", policy=policy)
result.policy_fingerprint   # 這批結果用了哪組參數，寫進 log／報表頁首
```

指紋＝參數 JSON 的 SHA-1 前 8 位（本版 `0bb5cf16`）。**指紋變了，就代表分配結果可能不同**——
排查「為甚麼上星期的單跟今天不一樣」時，先比指紋。

參數也可存成 JSON 再載入（`transfer_rules_golden.json` 的 `policy` 區塊本身就是一份可用範本）：

```python
import json
from transfer_rules import policy_from_dict

policy = policy_from_dict(json.load(open("transfer_rules_golden.json", encoding="utf-8")))
```

---

## 6. 件數分檔表（由 `format_rule_table()` 產生，勿手抄）

**一般分類**（`per_shop_target=3`）

| 轉出件數 | 接收店數 |
|---|---|
| 3 件或以下 | 1 |
| 4–6 件 | 2 |
| 7–9 件 | 3 |
| 10–12 件 | 4 |
| 13–30 件 | 5 |
| 31–60 件 | 6 |
| 61–90 件 | 7 |
| 91 件以上 | 每 30 件加 1 間 |

**`0302 - MINIATURE`**（`target_by_category={"0302": 4}`）

| 轉出件數 | 接收店數 |
|---|---|
| 4 件或以下 | 1 |
| 5–8 件 | 2 |
| 9–12 件 | 3 |
| 13–16 件 | 4 |
| 17–30 件 | 5 |
| 31–60 件 | 6 |
| 61–90 件 | 7 |
| 91 件以上 | 每 30 件加 1 間 |

說明頁、UI 提示、Excel 文案都應呼叫 `format_rule_table()` 生成，不要各自抄一份。

---

## 7. 邊界情況

| 情況 | 行為 |
|---|---|
| `total <= 0` | 回傳空分配，`allocated=0`、`unallocated=0`、`dropped=0` |
| 候選店為空 | 空分配，`unallocated=total`（呼叫端應標示「無接收店，保留原店」） |
| 候選店少於 `take` | 全數取用，`dropped=0` |
| 權重全為 0 | 視為等權分配（不會除零） |
| 只有 1 間候選店 | 全數給該店 |
| `total` 少於店數（如 3 件 5 間） | 保底退化為 `total // n`（可為 0），餘量按比例 |
| 封頂全部為 0 | `allocated=0`，`unallocated=total` |
| 重複 Site Code | 屬呼叫端資料問題，須先去重（重複會令庫存重複計算） |
| `mode` 非法值（負數、`0`、亂字串） | `ValueError`，訊息：`接收店舖集中上限必須為正整數、"auto" 或「不限」。` |

---

## 8. 驗收

### 8.1 Python 專案

```bash
py -3.14 -m pytest tests/test_transfer_rules.py -q
```

測試內容：載入 `transfer_rules_golden.json`，逐條比對 `receiver_count`、
`allocate_proportional`、`allocate_concentrated`、`allocate_with_caps`、`select_pool`、
`plan_allocation`，另加 §4 的不變量隨機測試。

### 8.2 非 Python 專案

把 `transfer_rules_golden.json` 讀進你的測試框架，逐條比對即可——**不需要看懂我的程式**：

```jsonc
// 片段
"receiver_count": [
  [0, "", 0],
  [3, "", 1],
  [4, "", 2],
  [13, "", 5],
  [4, "0302 - MINIATURE", 1],
  [9, "0302 - MINIATURE", 3]
],
"plans": [
  {
    "total": 13, "mode": "auto", "category": "", "cap_days": null,
    "candidates": [["S02", 9, 0], ["S03", 7, 0], ["S04", 5, 0]],
    "lines": [["S02", 5], ["S03", 4], ["S04", 4]],
    "allocated": 13, "unallocated": 0, "dropped": 0
  }
]
```

`lines` 的**次序**也是規範的一部分（分配序），不是無關的排序結果。

---

## 9. 接到其他調貨專案

### 9.1 三種方式，按情況選

| 方式 | 做法 | 適合 | 代價 |
|---|---|---|---|
| **A. 直接引入（推薦）** | 複製 `transfer_rules.py` + `transfer_rules_golden.json` + 測試到新專案 | 同為 Python（你現有專案都是） | 每次改規則要手動同步多份 |
| **B. 共用套件** | 放成內部 package／git submodule，各專案 pip 安裝或 submodule 引用 | 專案多、要同步升級 | 一次設定，長期省事 |
| **C. 按規範重寫** | 只帶本文件 + golden JSON，用該專案語言重寫（Java／Node／SQL） | 非 Python、或不能引入外部檔 | 要自行實作並跑通 781 個向量 |

**A 與 B 的差別只在同步方式**：A 是複製，B 是引用。以你目前「幾個調貨程式」的規模，
建議 **A 起步**；當同一條規則要改時發覺要改三個地方，就升級成 B。

### 9.2 五分鐘接入（pandas 專案）

```python
import pandas as pd
from transfer_rules import plan_allocation, DEFAULT_POLICY

def allocate_for_sku(df_sku: pd.DataFrame, total: int, mch2: str = "") -> pd.DataFrame:
    """df_sku 只含已過濾的候選接收店（同 SKU、非轉出店、P90D > 0）。"""
    candidates = df_sku[["Site Code", "Sold Qty (P90D)", "Shop Stock Qty (CM)"]].itertuples(
        index=False, name=None
    )
    result = plan_allocation(total, candidates, mode="auto", category=mch2)
    return pd.DataFrame(result.lines, columns=["Site Code", "轉貨數量 (件)"])


# 想封頂就加 cap_days=14；想對照舊行為就 mode=None
result = plan_allocation(
    total=13,
    candidates=[("S02", 9, 0), ("S03", 7, 0), ("S04", 5, 0)],
    mode="auto",
    cap_days=None,
)
print(result.summary())   # auto：3 間接收店，已分配 13 件，未分配 0 件…
```

### 9.3 只想要店數／只要數字

```python
from transfer_rules import DEFAULT_POLICY, select_pool, allocate

count = DEFAULT_POLICY.receiver_count(13)                    # 5
count = DEFAULT_POLICY.receiver_count(9, category="0302 - MINIATURE")   # 3

pool, dropped = select_pool(candidates, 13, mode="auto")      # 已排序的入選店
alloc, unalloc = allocate(13, [r[1] for r in pool], mode="auto")
```

### 9.4 呼叫端仍需自己做的三件事

1. **候選池過濾**：同 SKU、排除轉出店、`Sold Qty (P90D) > 0`；ND 轉貨模式另須 `RP Type＝RF`。
2. **配對**：轉出店依 Site Code 升序、接收店依分配量由大至小，依序配對成「轉出店 → 接收店」行。
   本規範不規定配對，但同一份轉貨單內應維持此慣例，否則同一批貨會出現無謂的跨店搬運。
3. **例外標示**：`unallocated > 0` 時要寫明原因（無接收店／封頂已滿），並保留原店；
   `dropped` 要能回答「為甚麼這間店沒收到貨」。

### 9.5 升級流程（改規則時照做）

1. 只調數值（例如迷你裝改 5 件）→ 改 `ReceiverPolicy` 參數，`RULES_VERSION` 不動，記在 §11。
2. 改公式／新增分類邏輯 → `RULES_VERSION` +0.1（如 1.1），更新 §3、§6，重跑
   `scripts/gen_transfer_rules_golden.py`（它會先驗證與引擎一致才寫檔），把新 JSON 一併發佈。
3. **不要**為了讓測試通過而去改 golden 的期望值——那等於把 bug 寫成規範。
   golden 只能由「已確認正確的行為」產生。

---

## 10. 常見改動對照

| 想要 | 改哪裡 | 不用碰 |
|---|---|---|
| 每間由 3 件改 4 件 | `per_shop_target=4` | 任何演算法 |
| 新增分類覆寫（如 `0303` 每間 5 件） | `target_by_category={"0302": 4, "0303": 5}` | 任何演算法 |
| 10 件以上最多 4 間 | `base_cap=4` | 任何演算法 |
| 取消上限（回到純 ceil） | `cap_from=10**9` | 任何演算法 |
| 不出現 1 件碎片 | `min_per_receiver=2`（預設已開） | 任何演算法 |
| 加「補到 14 日」封頂 | 傳 `cap_days=14` | 任何演算法 |
| 改封頂日數（如 21 日） | 傳 `cap_days=21` | 任何演算法 |

---

## 11. 變更記錄

| 版本 | 日期 | 變更 | 影響 |
|---|---|---|---|
| 1.0 | 2026-11-12 | 首版：`auto` 模式、件數上限階梯（10–30→5、之後每 30 加 1）、`0302` 每間 4 件、保底 2 件、14 日封頂可併用 | 收錄自 I14 report review v1.95.4 |

**待觀察（未納入規範）**：目前 `cap` 以「件數」為唯一依據，未考慮單店陳列容量或店舖等級；
若日後要按店舖分級（A 店可收較多、C 店封頂較低），建議以「每店權重上限」擴充 R8，
而不是再加一層 if——那會令分檔表無法自動生成。
