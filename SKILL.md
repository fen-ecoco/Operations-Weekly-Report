---
name: ecoco-weekly-analysis
description: |
  ecoco 宜可可循環經濟｜每週客訴回報分析 + 例會PPT自動產出技能（v5）。
  觸發詞：「分析本週」「同上週解析」「產出PPT」「例會報告」「兩頁簡報」。
  同時上傳 ecoco_complaint_report.html + 每週客訴內容.csv 時，直接產出PPT兩頁。
---

# ecoco 每週客訴分析 + 例會PPT 技能 v5

---

## 一、ecoco VI 色系

| 色名 | 色碼 | 用途 |
|------|------|------|
| Orange   | `#FF5000` | Section bar rect、電池機側欄、Top3加總欄、熱點件數、section2文字 |
| Blue     | `#060E9F` | 表頭fill、收瓶機側欄、section1文字、hot area城市名 |
| Yellow   | `#FFCE00` | 警示框border、圓餅第4色 |
| LightBlue| `#8EB8C9` | 圓餅第5色、電池機大卡border |
| Beige    | `#FAE0B8` | 趨勢表最新週底色、熱點城市行底色、警示框底色 |
| DarkBlue | `#0076A9` | 圓餅第3色、▼下降delta、Non-machine card2色 |
| DarkGray | `#333333` | 圖例件數/% **全部使用此色**、卡片名稱、Top3問題名 |
| TextGray | `#888888` | 卡片說明注文、件數小字 |

### 圓餅固定色序
```js
const CAT_C  = ["060E9F","FF5000","0076A9","FFCE00","8EB8C9","FAE0B8"]; // 客訴主分類
const MACH_C = ["FF5000","060E9F","0076A9","FFCE00","8EB8C9","B8BEC8"]; // 機台問題
```

### 字型
```js
const FC = "Noto Sans TC";   // 全頁中文統一
// 純數字/英數大字（%、加總）使用 "Arial"
```

### Delta 箭頭色
| 方向 | 色碼 |
|------|------|
| ▲ 上升 | `#FF5000` |
| ▼ 下降 | `#0076A9` |

---

## 二、PAGE 1 版面規格（pixel-measured from reference）

```
y:0.06  Title（bold #333333 fontSize:16，無背景色塊）
y:0.36  ▌Section bar "客訴趨勢分析"（bar:#FF5000 text:#060E9F）
y:0.60  Trend table
          x:0.18 w:9.64 h:1.96 rowH:0.37
          colW:[1.05,1.15,1.1,1.1,1.1,0.88,0.88,0.88,0.84]
          表頭 fill:#060E9F；加總 fill:#FF5000；最新週底色:#FAE0B8
          delta: ▲#FF5000 ▼#0076A9 fontSize:6.5
y:2.63  ▌Section bar "當週客訴佔比＋主分類佔比（日期）"（text:#FF5000）
y:2.89  圓餅圖（x:0.07 w:3.86 h:2.56）dataLabel: white 9pt bold "0%"
y:2.89  圖例（x:4.05 LH:0.425 per item）
          ├ 色塊 w:0.20 h:0.20
          ├ 類別名 fontSize:10 bold #333333
          └ 件數件　X.X% fontSize:8 #333333
y:2.89  Top3 卡片（x:6.33 CW:3.57 CH:0.875 gap:0.025）
          ├ 左色條 w:0.05 fill:問題色
          ├ 排名 fontSize:9 bold 問題色
          ├ 問題名稱 fontSize:10 bold #333333
          ├ 件數・% fontSize:18 bold 問題色（w:2.42 防換行）
          ├ 說明注文 fontSize:7 #888888 右側（x+2.80 w:0.73）
          └ mini bar h:0.05 底部
```

### Top3 卡片色對應
| 排名 | 邊框/件數色 | 背景 |
|------|------------|------|
| 1 | `#FF5000` | `#FFF4EF` |
| 2 | `#0076A9` | `#EEF6FB` |
| 3 | `#060E9F` | `#F5F7FF` |

---

## 三、PAGE 2 版面規格（pixel-measured from reference）

```
y:0.06  Title（bold #333333 fontSize:16）
y:0.36  ▌Section bar "客訴詳情與分類佔比"（text:#060E9F）
y:0.52  機台圓餅（x:0.07 w:3.05 h:1.70）
          dataLabel: white 8pt bold "0%"
y:0.54  機台圖例（x:3.18 MLH:0.245）
          色塊 w:0.15 h:0.15 + 文字 fontSize:8.5 #333333
y:0.52  大%卡片（各 w:1.10 h:1.70）
          收瓶機 x:6.12 bg:#FFF3EC border:#FF5000 %色:#FF5000
          電池機 x:7.32 bg:#EDF5FB border:#8EB8C9 %色:#0076A9
          %數字 fontSize:32 Arial
          件數  fontSize:8 #888888
y:2.32  ▌Section bar "收瓶機、電池機客訴 Top 3"（text:#060E9F）
y:2.32  ▌Section bar "本週熱點站點（第X週：日期）"（x:6.20 text:#FF5000）
y:2.52  收瓶機 sidebar（x:0.18 w:0.72 h:1.18 fill:#060E9F）
          items：x:0.92 w:5.18 h:0.365 rowGap:0.385
          排名 fontSize:10 bold 問題色
          名稱 fontSize:9.5 #333333
          % fontSize:20 Arial bold 問題色（x:5.26 w:0.78 align:right）
          mini bar h:0.038
y:3.715 電池機 sidebar（w:0.72 h:0.80 fill:#FF5000）
          items：h:0.358 rowGap:0.378
y:4.54  警示框（x:0.18 w:5.88 h:0.60）
          fill:#FAE0B8 border:#FFCE00
          text fontSize:7.5 #6B4800
y:2.52  熱點站點（x:6.20）
          城市行 h:0.26 fill:#FAE0B8 名 fontSize:12 bold #060E9F 件數 fontSize:13 #FF5000
          站點行 h:0.22 fontSize:8.5 #333333，件數 #888888
          城市間距 aY+:0.27；站點間距 aY+:0.225；城市gap:0.02
```

---

## 四、DATA 物件範本

```js
const D = {
  week:"第N週", range:"MM/DD ~ MM/DD", total:0,
  cats:[
    // 固定順序（對應 CAT_C 色序）
    {label:"APP帳號設定問題",count:0,pct:0},
    {label:"機台問題",        count:0,pct:0},
    {label:"回收點數問題",    count:0,pct:0},
    {label:"顧客關係",        count:0,pct:0},
    {label:"優惠券問題",      count:0,pct:0},
    {label:"APP使用問題",     count:0,pct:0},
  ],
  trend:[
    // 近4週，格式：[值, delta]，0=無變化
    {w:"NN週",d:"MM/DD-MM/DD",
     bt:[0,0],b2:[0,0],rg:[0,0],pt:[0,0],cp:[0,0],ap:[0,0],cr:[0,0],tot:[0,0]},
  ],
  nonMachine:[
    // 非機台問題Top3
    {rank:1,name:"",count:0,pct:0,col:"FF5000",note:""},
    {rank:2,name:"",count:0,pct:0,col:"0076A9",note:""},
    {rank:3,name:"",count:0,pct:0,col:"060E9F",note:""},
  ],
  machIssues:[
    // 機台問題前5項 + 其他合計
    {label:"",count:0,pct:0},
    ...
    {label:"其他機台問題",count:0,pct:0},
  ],
  bottleTop3:[
    {name:"",count:0,pct:0,col:"FF5000"},
    {name:"",count:0,pct:0,col:"060E9F"},
    {name:"",count:0,pct:0,col:"0076A9"},
  ],
  batTop:[
    {name:"",count:0,pct:0,col:"0076A9"},
    {name:"",count:0,pct:0,col:"8EB8C9"},
  ],
  bottleTotal:0,batTotal:0,bottlePct:0,batPct:0,
  hotAreas:[
    {area:"",total:0,spots:[{name:"",count:0}]},
  ],
  alertText:"【注意】...",
};
```

---

## 五、統計流程

```python
import pandas as pd

df = pd.read_csv('每週客訴內容.csv')

# 標籤統一
label_map = {
    'app畫面顯示與機台狀態不符':'APP畫面顯示與機台狀態不符',
    'app無法登入':'APP無法登入',
    'app多重異常狀況':'APP多重異常狀況',
}
df['簡報使用'] = df['簡報使用'].replace(label_map)
total = len(df)

# 問題類型佔比
qt = df['問題類型'].value_counts()

# 非機台問題 Top3（PAGE1 右卡）
nm_top3 = df[df['問題類型']!='機台問題類型']['簡報使用'].value_counts().head(3)

# 機台問題詳細（PAGE2 圓餅，前5+其他）
mach = df[df['問題類型']=='機台問題類型']['簡報使用'].value_counts()

# 收瓶機/電池機
bottle = df[df['機台類型'].isin(['收瓶機','方舟站收瓶機','方舟站'])]
bat    = df[df['機台類型']=='電池機']

# 熱點站點（前3區域各前3站）
area_top = df[df['站點區域'].notna()]['站點區域'].value_counts().head(3)

# 趨勢 delta：[當週值, 與上週差值]，正=▲，負=▼，0=無
# 從 HTML 解析或手動計算
```

---

## 六、執行流程

```bash
# 1. 統計 CSV → 填入 DATA
# 2. node generate_ppt.js
# 3. 視覺 QA（必做）
python3 /mnt/skills/public/pptx/scripts/office/soffice.py \
  --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf /tmp/slide
# view /tmp/slide-1.jpg /tmp/slide-2.jpg
# 對照參考圖確認
```

---

## 七、品質檢查清單

- [ ] 無 Header/Footer 色塊（標題改純文字 #333333）
- [ ] 無 ecoco 橘色 Logo 文字
- [ ] 字型全頁 `Noto Sans TC`，大數字/% 用 `Arial`
- [ ] 圖例件數與% 使用 `#333333`（不用問題色）
- [ ] 趨勢 delta ▲橘 ▼藍，最新週 Beige 底
- [ ] 卡片件數・% 單行（fontSize:18 w:2.42）
- [ ] PAGE2 大%卡片 fontSize:32（確認不換行）
- [ ] PAGE2 Top3 % 文字寬 0.78"（確認完整顯示）
- [ ] 熱點站點3城市完整顯示（aY動態檢查 < 5.55）
- [ ] 警示框 Beige底 + Yellow border
- [ ] 視覺 QA 對照參考圖確認無截斷/重疊
- [ ] 檔名含週次與日期區間
