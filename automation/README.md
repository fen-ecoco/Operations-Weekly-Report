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
├── make_icons.js         # 產生第二、三頁圖示 PNG（icons/ 資料夾，通常不需重跑）
├── icons/                 # 簡報使用的圖示素材
├── config.json            # 所有檔案路徑設定（UTF-8，勿在 .ps1 內寫中文路徑）
├── run_weekly.ps1         # 主控腳本，Task Scheduler 排程的進入點
├── volume_history.json    # 站點「近30日合計(瓶)」歷史紀錄（首次執行後自動產生，需 git 追蹤）
└── README.md              # 本檔案
```

## 執行前必備的資料檔（放在 `D:\info\0507_Weekly-Report\`）
| 檔案 | 用途 | 更新頻率 |
|---|---|---|
| `客訴內容分析.csv` | 全部客訴原始紀錄，程式自動抓「上一個完整週」 | 建議每週一早上、排程執行前更新 |
| `月回收量等級.csv` | 全站等級與回收量主檔（第二頁站點等級標示＋第三頁改善清單，共用同一份） | 每月更新 |
| `資料來源.md` | 等級計算方式說明，第三頁下方會自動讀取顯示 | 計算規則變動時更新 |

### `月回收量等級.csv` 使用說明（整份文件，欄位可能超過30欄，程式只取用需要的欄位）
第二、三頁共用這一份主檔，比對／取值方式：
- **第二頁站點等級**：依「站點名稱」對應「等級」欄位，顯示在「最高回報站點：」文字前方（例：`A｜桃園/全聯福利中心大溪員林店站`）
- **第三頁改善清單**：整份資料依「近30日合計(瓶)」由低到高排序，取最低10筆
- **「Hive排名(總NNN)」欄名裡的數字**：程式自動解析當作第三頁表頭的「總排名」母體數，每月改欄名時自動套用
- **「機型」「租賃/買斷」欄位**：轉成第三頁的彩色膠囊徽章
- **「月增減率」欄位**：來源檔的值不採用，改由 `volume_history.json` 自動累積比對「近30日合計(瓶)」計算月增減率（見下方）
- **「活動期間」「週別」欄位**：不使用，不會顯示在簡報上

## config.json 說明
```json
{
  "complaint_csv_path": "D:\\info\\0507_Weekly-Report\\客訴內容分析.csv",
  "volume_csv_path": "D:\\info\\0507_Weekly-Report\\月回收量等級.csv",
  "data_source_md_path": "D:\\info\\0507_Weekly-Report\\資料來源.md",
  "total_network_stations": 542,
  "output_dir": "D:\\info\\0507_Weekly-Report\\Operations-Weekly-Report\\weekly-ppt"
}
```
`total_network_stations` 僅作為備援值（若「Hive排名」欄名解析失敗時才會用到），正常情況下總排名數字由 CSV 欄名自動帶入。

## 資料來源.md 說明
第三頁下方會自動讀取這份檔案並顯示：
- 第一行文字：直接當作資料來源說明句
- 「級別/瓶量/加分」段落：自動解析成 S～F 等級門檻徽章列
- 「級別/特殊加分」段落：自動解析成 ARK 特殊加分說明

未來若門檻數字調整，直接改這份 `.md` 檔案即可，簡報會自動套用最新內容，不需要改程式。

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

## 月增減率運作方式
`analyze.py` 每次執行都會把當月（依系統日期，格式 `YYYY-MM`）10 站的「近30日合計(瓶)」寫入 `volume_history.json`。
- **首次執行**：無比較基準，第三頁會顯示「－（首次記錄）」
- **第二個月起**：自動比對上個月同站點的回收量，計算「▲上升N%／▼下降N%」
- `volume_history.json` 必須隨每次執行 `git push` 一併留存，否則下個月會遺失比較基準（`run_weekly.ps1` 的 `git add .` 已包含此檔案）

## Task Scheduler
- 觸發時間：每週一 11:00
- 已啟用 `StartWhenAvailable`（機器未開機時，開機後補跑）
- 若排程建立失敗，通常是 `.ps1` 內含中文字元導致編碼錯誤 → 本版已改為純 ASCII 腳本，中文路徑一律經由 `config.json`（UTF-8）讀取

## 版本重點（本次交付）
- 第二頁站點等級、第三頁改善清單改為共用同一份 `月回收量等級.csv`（不再使用 `收瓶量分析報告.csv`）
- 第三頁改為「近30日合計(瓶)」由低到高排序，並新增機型／租賃買斷徽章
- 第三頁下方資料來源說明改為自動讀取 `資料來源.md`
