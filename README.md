# ecoco 每週客訴週報 自動產出系統

版本：v6　|　維護：行銷客服專員　|　技術：pptxgenjs + Python + PowerShell

---

## 📂 專案結構

```
D:\info\0507_Weekly-Report\
├── 客訴內容分析.csv          ← 每週更新（週一 11:00 前放入）
├── 月回收量等級.csv          ← 每月更新（全站等級／排名／回收量主檔，525+ 站）
├── 資料來源.md               ← 等級計算規則說明（S～F 門檻＋ARK加分，異動時更新）
└── Operations-Weekly-Report\
    ├── README.md              ← 本文件
    ├── weekly-ppt\            ← 每週產出的 PPTX 自動存入這裡
    └── automation\
        ├── analyze.py                       ← Step 1+2：讀取 CSV → data.json
        ├── generate_ppt_auto.js             ← Step 3：data.json → PPTX
        ├── make_icons.js                    ← 產生簡報用圖示（icons/，通常不需重跑）
        ├── icons/                            ← 簡報圖示素材
        ├── run_weekly.ps1                    ← 主流程腳本（Steps 1-5 + git push）
        ├── config.json                       ← 本機設定（含各資料檔路徑，UTF-8）
        ├── volume_history.json               ← 站點回收量歷史紀錄（自動累積，需 git 追蹤）
        ├── monthly_low_volume_history.json   ← 每月低回收量Top10名單累積紀錄（第四頁資料，自動累積，需 git 追蹤）
        └── README.md                         ← 自動化系統技術文件（欄位格式、config 說明等）
```

## 🗓 執行方式

由 Windows Task Scheduler 於**每週一 11:00** 自動觸發 `run_weekly.ps1`，已啟用
`StartWhenAvailable`（機器未開機時，開機後自動補跑）。

需要手動重跑時：
```powershell
cd D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation
powershell -ExecutionPolicy Bypass -File .\run_weekly.ps1
```

## 📊 週報內容（四頁）

1. **客訴趨勢分析** — 近4週趨勢表、當週四格數據卡（總件數／最高主題／最高站點／非機台問題佔比）、非機台問題 Top3
2. **客訴機台類型與高頻站點分析** — 收瓶機/方舟、二代電池機 Top3，最高回報站點附站點等級（例：`A｜桃園/全聯福利中心大溪員林店站`）
3. **月低回收量站點改善清單 Top10** — 依「近30日合計(瓶)」由低到高排序，含等級／Hive排名／機型／租賃買斷徽章／月增減率
4. **客訴趨勢洞察（MoM／QoQ）** — 一句話重點結論橫幅、月度／季度KPI總覽卡、總客訴量與客訴密度雙軸趨勢圖、客訴主題結構比變化（QoQ堆疊橫條圖）、本季結構變化重點說明

## 📥 三份來源資料

| 檔案 | 用途 | 更新頻率 |
|---|---|---|
| `客訴內容分析.csv` | 客訴原始紀錄，程式自動判斷「上一個完整週」；第四頁 MoM/QoQ 亦直接由此檔動態彙總（本身保留全年歷史，免額外維護） | 每週 |
| `月回收量等級.csv` | 全站等級／排名／回收量主檔，第二、三頁共用 | 每月 |
| `資料來源.md` | 等級門檻與計算規則，第三頁自動解析顯示 | 規則異動時 |

> ⚠️ **`收瓶量分析報告.csv`、獨立的「站點區別對照表 (A/B/C 代碼)」已不再使用**，
> 站點等級自 v6 起統一改由 `月回收量等級.csv` 提供（等級制度也已從 A/B/C 三級
> 擴充為 **S/A/B/C/D/E/F 七級 + ARK 特殊類別**）。

詳細欄位格式、`config.json` 設定、月增減率運作邏輯、**第四頁 MoM/QoQ 資料來源與限制**，請見
[`automation/README.md`](./automation/README.md)。

## 🔧 常見問題

- **產出失敗、找不到欄位**：先確認三份來源 CSV/MD 是否為最新版、欄位名稱是否被異動
- **想重新產出某一週的簡報**：更新好 `客訴內容分析.csv` 後，直接重跑 `run_weekly.ps1` 即可，會覆蓋當天已產出的檔案
- **`.ps1` 出現中文編碼錯誤**：所有中文路徑一律走 `config.json`（UTF-8），不要把中文字元直接寫進 `.ps1`
- **第四頁「客訴密度指標」是什麼**：目前資料源缺乏站點覆蓋用戶數，暫以「每站平均客訴量」代替萬人客訴率，待有用戶數資料源後可直接替換公式
