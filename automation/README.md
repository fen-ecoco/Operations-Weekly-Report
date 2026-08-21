# ecoco 客服週報自動化系統 (v6)

## 這是什麼
每週一 11:00 由 Windows Task Scheduler 觸發，自動讀取客訴與回收量資料，產出三頁 PPTX 週報：
1. 客訴趨勢分析（近4週趨勢表＋當週四格數據卡＋非機台問題 Top3）
2. 客訴機台類型與高頻站點分析（收瓶機/方舟、二代電池機 Top3，含站點等級 A/B/C）
3. 月低回收量站點改善清單 Top10（含等級／Hive排名／MOM排名趨勢）

## 資料夾結構
```
automation/
├── analyze.py            # 讀取 CSV，輸出 data.json
├── generate_ppt_auto.js  # 讀取 data.json，產出 PPTX
├── make_icons.js         # 產生第二頁圖示 PNG（icons/ 資料夾，通常不需重跑）
├── icons/                 # 第二頁使用的圖示素材
├── config.json            # 所有檔案路徑設定（UTF-8，勿在 .ps1 內寫中文路徑）
├── run_weekly.ps1         # 主控腳本，Task Scheduler 排程的進入點
├── rank_history.json      # 站點排名歷史紀錄（首次執行後自動產生，需 git 追蹤）
└── README.md              # 本檔案
```

## 執行前必備的三份資料檔（放在 `D:\info\0507_Weekly-Report\`）
| 檔案 | 用途 | 更新頻率 |
|---|---|---|
| `客訴內容分析.csv` | 全部客訴原始紀錄，程式自動抓「上一個完整週」 | 建議每週一早上、排程執行前更新 |
| `月回收量等級.csv` | 月低回收量站點改善清單（第三頁六欄資料來源，見下方格式說明） | 每月更新 |
| `收瓶量分析報告.csv` | 全台站點等級 A/B/C（僅供第二頁站點等級標示使用） | 每月更新 |

### `月回收量等級.csv` 欄位格式（固定，「530」每月手動更新）
```
等級 / Hive排名(總排名530) / 城市 / 站點名稱 / 總回收量(瓶) / MOM排名趨勢 / 週別
```
- **「Hive排名(總排名530)」欄名裡的數字**：程式會自動解析這個數字作為第三頁標題與表頭的「總排名」母體數，每月改欄名時該數字就會自動套用，不需要另外改程式或 `config.json`
- **「MOM排名趨勢」欄位**：來源檔本身留空即可，程式不會讀取此欄的內容，而是用 `rank_history.json` 自動累積比對算出（見下方說明）
- **「週別」欄位**：不使用，可略過

## config.json 說明
```json
{
  "complaint_csv_path": "D:\\info\\0507_Weekly-Report\\客訴內容分析.csv",
  "volume_csv_path": "D:\\info\\0507_Weekly-Report\\月回收量等級.csv",
  "grade_report_path": "D:\\info\\0507_Weekly-Report\\收瓶量分析報告.csv",
  "total_network_stations": 530,
  "output_dir": "D:\\info\\0507_Weekly-Report\\Operations-Weekly-Report\\weekly-ppt"
}
```
`total_network_stations` 僅作為備援值（若 `月回收量等級.csv` 的「Hive排名」欄名解析失敗時才會用到），正常情況下總排名數字由 CSV 欄名自動帶入，不需手動同步兩邊。

## 手動測試執行方式
```powershell
cd D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation
python analyze.py
node generate_ppt_auto.js
```
或直接執行主控腳本（等同 Task Scheduler 觸發的內容）：
```powershell
powershell -File run_weekly.ps1
```
執行紀錄會寫入 `run_weekly.log`。

## MOM 排名趨勢運作方式
`analyze.py` 每次執行都會把當月（依系統日期，格式 `YYYY-MM`）10 站的 Hive 排名寫入 `rank_history.json`。
- **首次執行**：無比較基準，第三頁會顯示「－（首次記錄）」
- **第二個月起**：自動比對上個月同站點排名，計算「↑進步 N 名／↓退步 N 名」
- `rank_history.json` 必須隨每次執行 `git push` 一併留存，否則下個月會遺失比較基準（`run_weekly.ps1` 的 `git add .` 已包含此檔案）

## Task Scheduler
- 觸發時間：每週一 11:00
- 已啟用 `StartWhenAvailable`（機器未開機時，開機後補跑）
- 若排程建立失敗，通常是 `.ps1` 內含中文字元導致編碼錯誤 → 本版已改為純 ASCII 腳本，中文路徑一律經由 `config.json`（UTF-8）讀取

## 版本重點（本次交付）
- 第三頁改為僅讀取 `月回收量等級.csv`（不再比對 `收瓶量分析報告.csv`）
- 第三頁加入 MOM 排名趨勢歷史累積機制
- 第一、二頁版面比例、字級、留白已依最新指示調整
