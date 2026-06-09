# ecoco 週報自動化系統

每週一 11:00 前自動完成：分析 CSV → 產出 PPT → Push GitHub → 本機同步

---

## 📁 檔案說明

| 檔案 | 說明 |
|------|------|
| `analyze.py` | Step 1+2：讀取 CSV，計算統計，輸出 `data.json` |
| `generate_ppt_auto.js` | Step 3：讀取 `data.json`，產出 PPTX（v5 版面） |
| `run_weekly.ps1` | 主流程：Steps 1–4 + Git push + 本機同步 |
| `setup_scheduler.ps1` | **一次性安裝**：套件安裝 + 工作排程器設定 |
| `config.json.example` | 設定檔範本（複製為 `config.json` 後填入） |
| `history.json` | 累積各週加總（自動維護，勿手動修改） |
| `run_log.txt` | 執行紀錄（每次自動追加） |

---

## 🚀 首次安裝（5 分鐘）

### Step A　Clone repo 並建立設定檔

以系統管理員開啟 PowerShell：

```powershell
# Clone 到本機
git clone https://github.com/fen-ecoco/Operations-Weekly-Report.git ^
  "D:\info\0507_Weekly-Report\Operations-Weekly-Report"

cd "D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation"

# 建立設定檔
Copy-Item config.json.example config.json
notepad config.json   # 填入 github_pat
```

`config.json` 需填入：
```json
{
  "csv_path":    "D:\\info\\0507_Weekly-Report\\每週客訴內容.csv",
  "output_dir":  "D:\\info\\0507_Weekly-Report",
  "repo_dir":    "D:\\info\\0507_Weekly-Report\\Operations-Weekly-Report",
  "python_path": "python",
  "node_path":   "node",
  "git_path":    "git",
  "github_repo": "https://github.com/fen-ecoco/Operations-Weekly-Report.git",
  "github_user": "fen-ecoco",
  "github_pat":  "ghp_你的新PAT"
}
```

### Step B　一鍵安裝（套件 + 工作排程）

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
.\setup_scheduler.ps1
```

這會自動：
- 安裝 `pandas`（Python）
- 安裝 `pptxgenjs`（Node.js）
- 建立 Windows 工作排程（每週一 11:00）

---

## ⚙️ 每週自動流程

```
每週一 11:00（Task Scheduler 自動觸發）
  └─ run_weekly.ps1
      ├─ Step 1+2  python analyze.py
      │    └─ D:\info\0507_Weekly-Report\每週客訴內容.csv
      │         → automation\data.json（DATA 物件）
      ├─ Step 3    node generate_ppt_auto.js
      │    └─ data.json → ecoco_週報_第N週_MMDD-MMDD.pptx
      │         → D:\info\0507_Weekly-Report\（本機）
      │         → weekly-ppt\（GitHub repo）
      ├─ Step 4    開啟 PPTX 目視確認版面
      ├─ git add + commit + push → GitHub
      └─ ✅ 完成：本機 + GitHub 同步
```

---

## 🧪 手動執行

```powershell
cd "D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation"

# 僅測試（不 push）
.\run_weekly.ps1 -DryRun

# 完整執行（含 push）
.\run_weekly.ps1
```

---

## 🔑 PAT 更新方式

PAT 過期後只需更新 `config.json`：
1. GitHub → Settings → Developer settings → Tokens (classic) → Generate new token（勾選 `repo`）
2. 複製 token → 貼入 `config.json` 的 `github_pat`

---

## 📝 注意事項

- 每週客訴內容.csv 須在週一 11:00 前放置於 `D:\info\0507_Weekly-Report\`
- 未填入問題類型的記錄不計入分析總件數
- 趨勢表自動取 `history.json` 最近 4 週資料
- 執行紀錄查看：`automation\run_log.txt`
