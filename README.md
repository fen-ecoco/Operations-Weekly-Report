# ecoco 每週客訴週報 自動產出系統

**版本：v6　｜　維護：行銷客服專員　｜　技術：pptxgenjs + Python + PowerShell**

---

## 📁 專案結構

```
D:\info\0507_Weekly-Report\
├── 每週客訴內容.csv              ← 每週更新（週一 11:00 前放入）
├── 收瓶量分析報告.csv            ← 站點區別對照表（A/B/C 代碼）
└── Operations-Weekly-Report\
    ├── README.md                 ← 本文件
    ├── weekly-ppt\               ← 每週產出的 PPTX（自動存入）
    └── automation\
        ├── analyze.py            ← Step 1+2：CSV 分析 → data.json
        ├── generate_ppt_auto.js  ← Step 3：data.json → PPTX
        ├── run_weekly.ps1        ← 主流程腳本（Steps 1-4 + push）
        ├── setup_scheduler.ps1   ← 首次安裝工作排程器
        ├── config.json           ← 本機設定（含路徑與 PAT）
        ├── config.json.example   ← 設定範本
        ├── history.json          ← 週次歷史資料（自動維護）
        └── run_log.txt           ← 執行紀錄
```

---

## 🚀 每週自動流程（每週一 11:00 自動觸發）

```
Task Scheduler 自動啟動 run_weekly.ps1
  │
  ├─ Step 1+2  python analyze.py
  │    ├─ 讀取 每週客訴內容.csv
  │    ├─ 讀取 收瓶量分析報告.csv（站點 → 區別代碼對照）
  │    └─ 輸出 data.json
  │
  ├─ Step 3    node generate_ppt_auto.js
  │    └─ 產出 ecoco_weekly_第N週_MMDD-MMDD.pptx
  │         ├─ 本機：D:\info\0507_Weekly-Report\
  │         └─ GitHub：weekly-ppt\
  │
  ├─ Step 4    自動開啟 PPTX 目視確認
  │
  └─ git commit + push → GitHub ✅
```

---

## 📋 兩頁 PPT 結構

### PAGE 1 — 客訴問題分析

| 區塊 | 內容 |
|------|------|
| 趨勢表 | 近 4 週客訴數量趨勢（含 ▲▼ delta） |
| 當週客訴佔比 | 圓餅圖（6 類）＋圖例（件數 + %） |
| 非機台問題 Top3 | 3 張卡片（問題名稱、件數・%大字、說明文字） |

### PAGE 2 — 機台問題佔比 與 熱門站點

| 區塊 | 內容 |
|------|------|
| 客訴詳情 | 機台問題圓餅 ＋ 收瓶機 % / 電池機 % 大卡 |
| 收瓶機 Top3 | 問題排名、大 % 數字、mini bar |
| 電池機 Top2 | 同上 |
| 本週熱點站點 | 前 3 城市各前 3 站點，含區別代碼 |
| 警示框 | 本週異常提醒 |

---

## 🏷️ 熱點站點區別代碼說明

本週熱點站點顯示格式：
```
臺南                          17
  A  億進寢具安南店站           6
  B  全聯福利中心柳營中山店站    2
```

- **區別代碼**（A/B/C）來源：`收瓶量分析報告.csv` 的「區別」欄位
- 字體：Arial Bold，深黑色 `#1A1A1A`
- 對應邏輯：以「站點名稱」欄位精確比對，找到後取對應「區別」代碼
- 若無對應代碼：僅顯示站點名稱，不顯示代碼

---

## ⚙️ config.json 設定說明

```json
{
  "csv_path":           "D:\\info\\0507_Weekly-Report\\每週客訴內容.csv",
  "output_dir":         "D:\\info\\0507_Weekly-Report",
  "repo_dir":           "D:\\info\\0507_Weekly-Report\\Operations-Weekly-Report",
  "bottle_report_path": "D:\\info\\0507_Weekly-Report\\收瓶量分析報告.csv",
  "python_path":        "python",
  "node_path":          "node",
  "git_path":           "git",
  "github_repo":        "https://github.com/fen-ecoco/Operations-Weekly-Report.git",
  "github_user":        "fen-ecoco",
  "github_pat":         "ghp_你的PAT"
}
```

> ⚠️ `bottle_report_path` 為新增欄位（v6），請確認 config.json 已加入此設定。

---

## 🎨 ecoco VI 色系

| 色名 | 色碼 | 用途 |
|------|------|------|
| Orange | `#FF5000` | Section bar、加總欄、熱點城市件數 |
| Blue | `#060E9F` | 表頭、收瓶機側欄、section 文字 |
| Yellow | `#FFCE00` | 警示框邊框、圓餅第 4 色 |
| LightBlue | `#8EB8C9` | 圓餅第 5 色、電池機大卡邊框 |
| Beige | `#FAE0B8` | 最新週底色、熱點城市列、警示框底 |
| DarkBlue | `#0076A9` | 圓餅第 3 色、▼ 下降 delta |
| DarkGray | `#333333` | 圖例文字、站點名稱 |
| NearBlack | `#1A1A1A` | 熱點站點區別代碼（Bold） |

---

## ⚠️ 異常偵測規則（自動產生警示文字）

| 條件 | 等級 |
|------|------|
| 非機台問題佔比 > 10% | ⚠️ 自動加入警示框 |
| 收瓶機 Top1 佔比 ≥ 25% | ⚠️ 提醒確認清空頻率 |
| 站點區域填入異常值（站點編號、`-`） | 🔧 自動過濾，不顯示 |

---

## 🔧 手動執行

```powershell
cd "D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation"

# 完整流程（含 push）
powershell -ExecutionPolicy Bypass -File ".\run_weekly.ps1"

# 僅測試（不 push）
powershell -ExecutionPolicy Bypass -File ".\run_weekly.ps1" -DryRun
```

---

## 🔑 PAT 更新方式

GitHub Classic PAT 若過期：
1. GitHub → Settings → Developer settings → Tokens (classic) → Generate new token（勾選 `repo`）
2. 貼入 `automation\config.json` 的 `github_pat` 欄位

---

## 📝 版本紀錄

| 版本 | 日期 | 變更說明 |
|------|------|----------|
| v6 | 2026-06 | 熱點站點加入區別代碼（A/B/C）顯示；過濾站點編號異常值 |
| v5 | 2026-06 | 像素精準版面，符合 ecoco VI 色系規範 |
| v4 | 2026-06 | 移除 Header/Footer 色塊，改純文字標題 |
| v3 | 2026-06 | 圓餅圖改用 pptxgenjs 原生圖表 |
| v2 | 2026-06 | 初版 PPT 自動產出 |

---

## 🔗 相關資源

- GitHub Repo：[fen-ecoco/Operations-Weekly-Report](https://github.com/fen-ecoco/Operations-Weekly-Report)
- 客訴分析系統：[ecoco-complaint-analyzer.onrender.com](https://ecoco-complaint-analyzer.onrender.com)
