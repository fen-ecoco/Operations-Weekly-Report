# ecoco 週報自動化系統

每週一中午前自動完成：分析 CSV → 產出 PPT → Push GitHub → 本機同步

---

## 📁 檔案說明

| 檔案 | 說明 |
|------|------|
| `analyze.py` | Step 1+2：讀取 CSV，計算統計，輸出 `data.json` |
| `generate_ppt_auto.js` | Step 3：讀取 `data.json`，產出 PPTX（v5 版面） |
| `run_weekly.ps1` | 主流程：Steps 1–4 + Git push + 本機同步 |
| `setup_scheduler.ps1` | 一次性設定 Windows 工作排程器 |
| `config.json.example` | 設定檔範本（複製為 `config.json` 後填入） |
| `history.json` | 累積各週加總（自動維護，勿手動修改） |
| `run_log.txt` | 執行紀錄（每次追加） |

---

## 🚀 首次安裝（10分鐘）

### 1. 安裝必要工具

確認已安裝：
```powershell
python --version    # 需 3.9+
node --version      # 需 18+
git --version
pip install pandas  # Python 套件
npm install pptxgenjs  # Node 套件（在 automation 目錄下執行）
```

### 2. 建立設定檔

```powershell
cd C:\path\to\Operations-Weekly-Report\automation
Copy-Item config.json.example config.json
```

編輯 `config.json`，填入：
- `github_pat`：Classic PAT（需有 `repo` 權限）

### 3. 設定工作排程（以系統管理員身分執行）

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
.\setup_scheduler.ps1
```

---

## ⚙️ 每週自動流程

```
每週一 11:00（自動）
  └─ run_weekly.ps1
      ├─ [Step 1+2] python analyze.py
      │    └─ 讀取 每週客訴內容.csv → data.json
      ├─ [Step 3]  node generate_ppt_auto.js
      │    └─ data.json → ecoco_週報_第N週_MMDD-MMDD.pptx
      ├─ [Step 4]  開啟 PPTX 供目視確認
      ├─ git commit + push → GitHub
      └─ 複製 PPTX 至本機資料夾
```

---

## 🔑 PAT 過期處理

GitHub Classic PAT 無預設過期時間，但若手動撤銷需重新產生。

更新方式：
1. 到 GitHub → Settings → Developer settings → Tokens (classic)
2. 產生新 token（勾選 `repo`）
3. 更新 `config.json` 中的 `github_pat`

---

## 🧪 手動執行測試

```powershell
# 完整流程（含 push）
.\run_weekly.ps1

# 僅測試（不 push）
.\run_weekly.ps1 -DryRun
```

---

## ⚠️ 注意事項

- `每週客訴內容.csv` 須在週一 11:00 前更新完畢
- 107件「未分類」記錄不計入總件數（僅統計已填問題類型的資料）
- 趨勢表自動從 `history.json` 取最近4週資料
- `run_log.txt` 可查看每次執行結果
