# ecoco 每週客訴分析 + 例會PPT 自動產出系統

**版本：v5　｜　維護：行銷客服專員　｜　技術：pptxgenjs + Python**

---

## 📁 檔案結構

```
ecoco-weekly-analysis/
├── SKILL.md              # Claude AI 技能定義（觸發規則 + 完整規格）
├── README.md             # 本文件
├── generate_ppt.js       # PPT 產出主程式（每週更新 DATA 物件即可）
└── 輸出範例/
    └── ecoco_週報_第22週_0525-0531_v5.pptx
```

---

## 🚀 每週操作步驟

### Step 1　統計本週資料

上傳 `每週客訴內容.csv` 到 Claude，說：

> 「請分析本週，同上週的解析」

Claude 會自動輸出：問題類型佔比、非機台 Top3、收瓶機/電池機 Top3、熱點站點、趨勢 delta 數值。

---

### Step 2　更新 DATA 物件

將統計結果填入 `generate_ppt.js` 的 `const D = { ... }` 區塊：

| 欄位 | 說明 |
|------|------|
| `week` | 週次，如 `"第23週"` |
| `range` | 日期區間，如 `"06/01 ~ 06/07"` |
| `total` | 總件數 |
| `cats` | 6 種問題類型 count/pct（固定順序） |
| `trend` | 近 4 週趨勢，每格 `[值, delta]` |
| `nonMachine` | 非機台問題 Top3，含 note 說明文字 |
| `machIssues` | 機台問題前5 + 其他合計 |
| `bottleTop3` | 收瓶機 Top3 |
| `batTop` | 電池機 Top2 |
| `hotAreas` | 熱點城市前3（各含前3站點） |
| `alertText` | 警示說明文字（無異常填 `""`） |

---

### Step 3　產出 PPT

```bash
node generate_ppt.js
```

輸出至 `/mnt/user-data/outputs/ecoco_週報_第N週_MMDD-MMDD.pptx`

---

### Step 4　視覺 QA

```bash
python3 /mnt/skills/public/pptx/scripts/office/soffice.py \
  --headless --convert-to pdf 輸出檔案.pptx

pdftoppm -jpeg -r 150 輸出檔案.pdf /tmp/slide
```

開啟 `/tmp/slide-1.jpg` 和 `/tmp/slide-2.jpg` 與參考圖對照。

---

## 🎨 ecoco VI 色系

| 色名 | 色碼 | 主要用途 |
|------|------|----------|
| Orange | `#FF5000` | Section bar、加總欄、熱點件數 |
| Blue | `#060E9F` | 表頭、收瓶機側欄、section 文字 |
| Yellow | `#FFCE00` | 警示框邊框、圓餅第4色 |
| LightBlue | `#8EB8C9` | 圓餅第5色、電池機大卡邊框 |
| Beige | `#FAE0B8` | 最新週底色、熱點城市列、警示框底 |
| DarkBlue | `#0076A9` | 圓餅第3色、▼下降 delta |
| DarkGray | `#333333` | 圖例文字（類別名稱、件數與%） |

---

## 📋 兩頁結構說明

### PAGE 1 — 客訴問題分析

| 區塊 | 內容 |
|------|------|
| 標題 | 客服課　日期　客訴問題分析 |
| 客訴趨勢分析 | 近4週趨勢表（含 ▲▼ delta） |
| 當週客訴佔比 | 圓餅圖（6類）＋圖例 |
| 非機台問題 Top3 | 3張卡片（件數・%大字＋說明） |

### PAGE 2 — 機台問題佔比 與 熱門站點

| 區塊 | 內容 |
|------|------|
| 標題 | 客服課　日期　機台問題佔比 與 熱門站點 |
| 客訴詳情 | 機台問題圓餅 ＋ 收瓶機93%/電池機7%大卡 |
| 收瓶機/電池機 Top3 | 列表（大%字＋mini bar） |
| 本週熱點站點 | 前3城市各前3站點 |
| 警示框 | 本週異常提醒（beige底） |

---

## ⚠️ 異常偵測規則

| 條件 | 等級 |
|------|------|
| 單一非機台問題 > 30% | 🚨 嚴重 |
| 同一問題單日 > 10件 | ⚠️ 注意 |
| 問題類型排名翻轉 | ⚠️ 注意 |
| 電池機件數 = 0 | ℹ️ 觀察 |
| 同問題連兩週上升 | ⚠️ 注意 |

---

## 🔗 相關資源

- 週報資料來源：Google Sheets（每週匯出 CSV）
- PPT 參考範本：`D:\info\0507_test\每週營運例會報告ppt\`
- 部署系統：Render（`ecoco-complaint-analyzer.onrender.com`）
- GitHub Repo：`fen-ecoco/fen-ecoco-complaint_webapp2`
