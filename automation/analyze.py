# -*- coding: utf-8 -*-
"""
ecoco 客服週報 - 資料分析腳本 (v6)
資料來源：客訴內容分析.csv（取代原本的 每週客訴內容.csv）
輸出：data.json，供 generate_ppt_auto.js 讀取產出三頁 PPT
"""
import json
import os
import re
from datetime import datetime, timedelta
import pandas as pd

# ---------- 讀取設定檔（路徑一律放在 UTF-8 的 config.json，避免中文路徑寫在 .ps1 造成編碼問題）----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, "config.json"), "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

COMPLAINT_CSV = CONFIG["complaint_csv_path"]
VOLUME_CSV = CONFIG["volume_csv_path"]
RANK_HISTORY_PATH = os.path.join(SCRIPT_DIR, "volume_history.json")
DATA_SOURCE_MD = CONFIG.get("data_source_md_path")  # 資料來源.md：等級計算說明，顯示於第三頁下方

# ---------- 週次設定：預設自動抓取「上一個完整週」(週一~週日) ----------
# 若要手動指定其他週次，將下方兩行取消註解並填入日期即可覆蓋自動判斷
# WEEK_START_OVERRIDE = "2026-08-10"
WEEK_START_OVERRIDE = None

if WEEK_START_OVERRIDE:
    WEEK_START = WEEK_START_OVERRIDE
else:
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)  # 上上週一 -> 這是上週一
    WEEK_START = last_monday.strftime("%Y-%m-%d")

_week_num = datetime.strptime(WEEK_START, "%Y-%m-%d").isocalendar()[1]
WEEK_LABEL = f"第{_week_num}週"

# ================= 1. 讀取與清理 =================
df = pd.read_csv(COMPLAINT_CSV)
df.columns = df.columns.str.strip()
df["機台類型"] = df["機台類型"].astype(str).str.strip().replace({"nan": None})
df["問題類型"] = df["問題類型"].astype(str).str.strip()

label_map = {
    "app畫面顯示與機台狀態不符": "APP畫面顯示與機台狀態不符",
    "app無法登入": "APP無法登入",
    "app多重異常狀況": "APP多重異常狀況",
}
df["問題細項"] = df["問題細項"].replace(label_map)
df["dt"] = pd.to_datetime(df["進件日期"], errors="coerce")

# 讀取「月回收量等級.csv」：全台站點等級（ARK/S~F）＋排名，供第二頁站點等級標示、第三頁改善清單共用
# 這份檔案的欄位在合作過程中改過幾次，為避免格式再變動時整份週報中斷，以下全程防呆處理。
vol_master_df = None
if VOLUME_CSV and os.path.exists(VOLUME_CSV):
    try:
        _tmp = pd.read_csv(VOLUME_CSV)
        _tmp.columns = _tmp.columns.str.strip()
        if "站點名稱" in _tmp.columns and "等級" in _tmp.columns:
            vol_master_df = _tmp
        else:
            print(f"警告：{VOLUME_CSV} 缺少「站點名稱」或「等級」欄位（目前欄位：{list(_tmp.columns)}），本次第二頁將不顯示站點等級。")
    except Exception as e:
        print(f"警告：讀取 {VOLUME_CSV} 失敗（{e}），本次第二頁將不顯示站點等級。")


def lookup_grade(station_name):
    if vol_master_df is None:
        return None
    match = vol_master_df[vol_master_df["站點名稱"] == station_name]
    return str(match.iloc[0]["等級"]) if len(match) else None

CATS_ORDER = [
    ("APP帳號設定問題", "APP帳號設定問題類型"),
    ("機台問題", "機台問題類型"),
    ("回收點數問題", "回收點數問題類型"),
    ("顧客關係", "顧客關係類型"),
    ("優惠券問題", "優惠券問題類型"),
    ("APP使用問題", "APP使用問題類型"),
]
CAT_C = ["060E9F", "FF5000", "0076A9", "FFCE00", "8EB8C9", "FAE0B8"]


def week_slice(monday_str):
    s = pd.Timestamp(monday_str)
    e = s + pd.Timedelta(days=7)
    return df[(df["dt"] >= s) & (df["dt"] < e)].copy(), s, e


def top_station(sub):
    s = sub["站點名稱"].dropna()
    if len(s) == 0:
        return "", 0
    vc = s.value_counts()
    top = vc.index[0]
    cnt = int(vc.iloc[0])
    area = sub[sub["站點名稱"] == top]["站點區域"].dropna()
    area = area.iloc[0] if len(area) else ""
    grade = lookup_grade(top)
    grade_prefix = f"{grade}｜" if grade else ""
    return f"{grade_prefix}{area}/{top}", cnt


wk, wstart, wend = week_slice(WEEK_START)
total = len(wk)
range_label = f"{wstart.strftime('%m/%d')} ~ {(wend - pd.Timedelta(days=1)).strftime('%m/%d')}"

# ================= 2. 客訴主分類佔比 =================
vc_type = wk["問題類型"].value_counts()
cats = []
for label, key in CATS_ORDER:
    c = int(vc_type.get(key, 0))
    cats.append({"label": label, "count": c, "pct": round(c / total * 100, 1) if total else 0})

# ================= 3. 近4週趨勢 =================
# 動態計算「近4週」= 本週往前推 3、2、1 週 + 本週，避免寫死日期導致跨週執行時對不上（例如自動判斷週次已推進，但趨勢表仍卡在舊的固定週次）
def week_label(monday_ts):
    wn = monday_ts.isocalendar()[1]
    sunday_ts = monday_ts + pd.Timedelta(days=6)
    rng = f"{monday_ts.strftime('%m/%d')}-{sunday_ts.strftime('%m/%d')}"
    return f"第{wn}週", rng

WEEKS_TREND = []
for offset in [3, 2, 1, 0]:
    monday_ts = wstart - pd.Timedelta(days=7 * offset)
    wl, rng = week_label(monday_ts)
    WEEKS_TREND.append((wl, rng, monday_ts.strftime("%Y-%m-%d")))

trend_raw = []
for wl, rng, monday in WEEKS_TREND:
    twk, _, _ = week_slice(monday)
    ttot = len(twk)
    vc = twk["問題類型"].value_counts()
    row = {
        "w": wl, "d": rng,
        "app_acc": int(vc.get("APP帳號設定問題類型", 0)),
        "machine": int(vc.get("機台問題類型", 0)),
        "points": int(vc.get("回收點數問題類型", 0)),
        "customer": int(vc.get("顧客關係類型", 0)),
        "coupon": int(vc.get("優惠券問題類型", 0)),
        "app_use": int(vc.get("APP使用問題類型", 0)),
        "tot": ttot,
    }
    trend_raw.append(row)

trend = []
for i, row in enumerate(trend_raw):
    entry = {"w": row["w"], "d": row["d"]}
    for key in ["app_acc", "machine", "points", "customer", "coupon", "app_use", "tot"]:
        cur = row[key]
        delta = cur - trend_raw[i - 1][key] if i > 0 else 0
        entry[key] = [cur, delta]
    trend.append(entry)

# ================= 4. 非機台問題 Top3 =================
nm = wk[wk["問題類型"] != "機台問題類型"]
vc_nm = nm["問題細項"].value_counts()
nm_colors = ["FF5000", "0076A9", "060E9F"]
nonMachine = []
for i, (name, cnt) in enumerate(vc_nm.head(3).items()):
    nonMachine.append({
        "rank": i + 1, "name": name, "count": int(cnt),
        "pct": round(cnt / total * 100, 1), "col": nm_colors[i], "note": "",
    })

# ================= 5. 機台問題細項（前5 + 其他）=================
mach = wk[wk["問題類型"] == "機台問題類型"]
mach_total = len(mach)
vc_mach = mach["問題細項"].value_counts()
machIssues = []
top5 = vc_mach.head(5)
for name, cnt in top5.items():
    machIssues.append({"label": name, "count": int(cnt), "pct": round(cnt / mach_total * 100, 1)})
other = int(vc_mach.iloc[5:].sum())
if other > 0:
    machIssues.append({"label": "其他機台問題", "count": other, "pct": round(other / mach_total * 100, 1)})

# ================= 6. 收瓶機 / 電池機 Top3（圖示版第二頁用）=================
bottle = wk[wk["機台類型"].isin(["收瓶機", "方舟站", "方舟站收瓶機"])].copy()
bat = wk[wk["機台類型"] == "電池機"].copy()
bottle_total = len(bottle)
bat_total = len(bat)
machine_total = bottle_total + bat_total

bottle["merged"] = bottle["問題細項"].replace({"回收箱滿艙": "滿艙問題（回收箱／寶特瓶）", "寶特瓶滿艙": "滿艙問題（回收箱／寶特瓶）"})
vc_bottle = bottle["merged"].value_counts()
bottle_colors = ["FF5000", "060E9F", "0076A9"]
bottleTop3 = []
for i, (name, cnt) in enumerate(vc_bottle.head(3).items()):
    sub = bottle[bottle["merged"] == name]
    station, scount = top_station(sub)
    bottleTop3.append({
        "rank": i + 1, "name": name, "count": int(cnt),
        "pct": round(cnt / bottle_total * 100, 1) if bottle_total else 0,
        "col": bottle_colors[i], "station": station, "stationCount": scount,
    })

vc_bat = bat["問題細項"].value_counts()
bat_colors = ["FF5000", "060E9F", "0076A9"]
batTop3 = []
for i, (name, cnt) in enumerate(vc_bat.head(3).items()):
    sub = bat[bat["問題細項"] == name]
    station, scount = top_station(sub)
    batTop3.append({
        "rank": i + 1, "name": name, "count": int(cnt),
        "pct": round(cnt / bat_total * 100, 1) if bat_total else 0,
        "col": bat_colors[i], "station": station, "stationCount": scount,
    })

# ================= 7. 熱點站點（前3區域，各前3站）=================
top_areas = wk["站點區域"].value_counts().head(3)
hotAreas = []
for area, cnt in top_areas.items():
    sub = wk[wk["站點區域"] == area]
    spots = []
    for name, c in sub["站點名稱"].value_counts().head(3).items():
        spots.append({"name": name, "count": int(c)})
    hotAreas.append({"area": area, "total": int(cnt), "spots": spots})

# ================= 8. 融合洞察 / 警示文字（依當週實際 Top1 動態生成，避免寫死特定問題名稱）=================
if bottleTop3:
    b1 = bottleTop3[0]
    if b1["station"]:
        bottle_insight = (
            f"收瓶機客訴集中於「{b1['name']}」（{b1['pct']}%，{b1['count']}件），"
            f"最高回報站點為 {b1['station']}（{b1['stationCount']}件），建議優先派工檢修"
        )
    else:
        bottle_insight = f"收瓶機客訴集中於「{b1['name']}」（{b1['pct']}%，{b1['count']}件），建議優先派工檢修"
else:
    bottle_insight = "本週收瓶機無客訴紀錄"

if batTop3:
    b2 = batTop3[0]
    battery_insight = f"電池機客訴則集中於「{b2['name']}」（{b2['pct']}%，{b2['count']}件），建議持續追蹤"
else:
    battery_insight = "本週電池機無客訴紀錄"

alertText = f"【本週融合洞察】{bottle_insight}；{battery_insight}。"

# ================= 9. 月低回收量站點（第三頁）=================
# 資料來源僅使用 月回收量等級.csv 整份文件（525+ 站全量），欄位包含：
#   等級 / Hive排名(總NNN) / 縣市 / 站點名稱 / 機型 / 租賃/買斷 / 近30日合計(瓶) / 月增減率 / 活動期間 / 週別
# 「活動期間」「週別」不顯示。「總NNN」母體數自動從欄名解析。
# 「月增減率」欄位本身不採用來源檔的值，改由 volume_history.json 每月自動累積比對「近30日合計(瓶)」計算（見下方）。
if vol_master_df is None:
    raise KeyError(f"無法讀取 {VOLUME_CSV}，請確認檔案存在且含「站點名稱」「等級」欄位。")
vol = vol_master_df.copy()
vol = vol.dropna(subset=["站點名稱"]).copy()

hive_rank_col = next((c for c in vol.columns if c.startswith("Hive排名")), None)
volume_col = next((c for c in vol.columns if "近30日合計" in c or c.replace(" ", "").startswith("總回收量")), None)
city_col = "縣市" if "縣市" in vol.columns else ("城市" if "城市" in vol.columns else None)
if hive_rank_col is None or volume_col is None or city_col is None:
    raise KeyError(
        f"找不到「Hive排名(...)」「近30日合計(瓶)」或「縣市/城市」欄位，請確認 {VOLUME_CSV} 的欄位名稱。"
        f"目前讀到的欄位為：{list(vol.columns)}"
    )

rank_match = re.search(r"(\d+)", hive_rank_col)
total_network_stations = int(rank_match.group(1)) if rank_match else CONFIG.get("total_network_stations")

vol["總量_num"] = vol[volume_col].astype(str).str.replace(",", "").str.strip().astype(float)
vol = vol.sort_values("總量_num", ascending=True).head(10)

# 讀取／更新「近30日合計(瓶)」歷史紀錄（供月增減率使用；首次執行無比較基準，之後每月自動累積）
current_month_key = datetime.now().strftime("%Y-%m")
volume_history = {}
if os.path.exists(RANK_HISTORY_PATH):
    with open(RANK_HISTORY_PATH, "r", encoding="utf-8") as f:
        volume_history = json.load(f)

lowVolumeStations = []
for _, r in vol.iterrows():
    name = str(r["站點名稱"])
    grade = str(r["等級"]) if pd.notna(r.get("等級")) else None
    hive_rank = int(r[hive_rank_col]) if pd.notna(r.get(hive_rank_col)) else None
    machine_type = str(r["機型"]).strip() if pd.notna(r.get("機型")) else None
    lease_type = str(r["租賃/買斷"]).strip() if pd.notna(r.get("租賃/買斷")) else None
    cur_volume = r["總量_num"]

    station_hist = volume_history.get(name, {})
    prev_months = sorted([m for m in station_hist.keys() if m != current_month_key], reverse=True)
    mom_change = None
    if prev_months:
        prev_volume = station_hist[prev_months[0]]
        if prev_volume:
            pct = round((cur_volume - prev_volume) / prev_volume * 100, 1)
            mom_change = {"prev_volume": prev_volume, "pct": pct}

    volume_history.setdefault(name, {})[current_month_key] = cur_volume

    lowVolumeStations.append({
        "name": name, "city": str(r[city_col]), "contribution": int(cur_volume),
        "hiveRank": hive_rank, "grade": grade, "momChange": mom_change,
        "machineType": machine_type, "leaseType": lease_type,
    })

with open(RANK_HISTORY_PATH, "w", encoding="utf-8") as f:
    json.dump(volume_history, f, ensure_ascii=False, indent=2)

# 第三頁下方「資料來源」說明：取用 資料來源.md 內容（若找不到或格式不符則用內建預設文字，確保不中斷）
DEFAULT_DATA_SOURCE_TEXT = (
    "資料來源：Hive，一般站依近30個台灣日曆日實際清運的塑膠＋鋁罐瓶數分級（S～F）；"
    "三台以上有效收瓶機獨立列為 ARK。"
)
DEFAULT_GRADE_TIERS = [
    {"grade": "S", "range": "17,000瓶以上"},
    {"grade": "A", "range": "13,300～16,999瓶"},
    {"grade": "B", "range": "11,100～13,299瓶"},
    {"grade": "C", "range": "9,000～11,099瓶"},
    {"grade": "D", "range": "6,900～8,999瓶"},
    {"grade": "E", "range": "4,100～6,899瓶"},
    {"grade": "F", "range": "4,100瓶以下"},
]
dataSourceNote = DEFAULT_DATA_SOURCE_TEXT
gradeTiers = DEFAULT_GRADE_TIERS
arkNote = "三台以上有效收瓶機獨立列為 ARK（額外+8分）"

if DATA_SOURCE_MD and os.path.exists(DATA_SOURCE_MD):
    try:
        with open(DATA_SOURCE_MD, "r", encoding="utf-8") as f:
            md_text = f.read()
        first_line = next((ln.strip() for ln in md_text.splitlines() if ln.strip()), None)
        if first_line:
            dataSourceNote = first_line

        # 動態解析「等級/瓶量/加分」門檻，例如： S /17000~/+18 。A/13300~16999/+12。...
        tier_pattern = r"([A-Z])\s*/\s*([\d,]*)~([\d,]*)/\+(\d+)"
        matches = re.findall(tier_pattern, md_text)
        if matches:
            parsed_tiers = []
            for g, lo, hi, bonus in matches:
                if lo and hi:
                    rng = f"{int(lo):,}～{int(hi):,}瓶"
                elif lo and not hi:
                    rng = f"{int(lo):,}瓶以上"
                elif hi and not lo:
                    rng = f"{int(hi):,}瓶以下"
                else:
                    rng = ""
                parsed_tiers.append({"grade": g, "range": rng})
            gradeTiers = parsed_tiers

        ark_match = re.search(r"級別/特殊加分[：:]\s*(.+)", md_text)
        if ark_match:
            arkNote = ark_match.group(1).strip()
    except Exception as e:
        print(f"警告：讀取 {DATA_SOURCE_MD} 失敗（{e}），第三頁資料來源說明改用內建預設文字。")


# ================= 10. 頂部四格數據卡（本週 vs 上週 WoW；近28日 vs 前28日 MoM）=================
prev_monday = (wstart - pd.Timedelta(days=7)).strftime("%Y-%m-%d")
wk_prev, _, _ = week_slice(prev_monday)
prev_total = len(wk_prev)

wow_total_pct = round((total - prev_total) / prev_total * 100, 1) if prev_total else 0

cur28 = df[(df["dt"] >= wend - pd.Timedelta(days=28)) & (df["dt"] < wend)]
prev28 = df[(df["dt"] >= wend - pd.Timedelta(days=56)) & (df["dt"] < wend - pd.Timedelta(days=28))]
mom_total_pct = round((len(cur28) - len(prev28)) / len(prev28) * 100, 1) if len(prev28) else 0

# 卡1：本週客訴總件數
card1 = {
    "value": total, "prev": prev_total,
    "diff": total - prev_total,
    "wow_pct": wow_total_pct,
}

# 卡2：回報最高主題
vc_issue = wk["問題細項"].value_counts()
top_issue_name = vc_issue.index[0]
top_issue_cnt = int(vc_issue.iloc[0])
top_issue_share = round(top_issue_cnt / total * 100, 1) if total else 0
vc_issue_prev = wk_prev["問題細項"].value_counts()
prev_issue_cnt = int(vc_issue_prev.get(top_issue_name, 0))
prev_issue_share = round(prev_issue_cnt / prev_total * 100, 1) if prev_total else 0
card2 = {
    "name": top_issue_name, "count": top_issue_cnt, "share": top_issue_share,
    "share_delta": round(top_issue_share - prev_issue_share, 1),
}

# 卡3：客訴量最高站點
vc_station = wk["站點名稱"].value_counts()
top_station_name = vc_station.index[0]
top_station_cnt = int(vc_station.iloc[0])
top_station_area = wk[wk["站點名稱"] == top_station_name]["站點區域"].dropna()
top_station_area = top_station_area.iloc[0] if len(top_station_area) else ""
prev_station_cnt = int((wk_prev["站點名稱"] == top_station_name).sum())
card3 = {
    "name": top_station_name, "area": top_station_area, "count": top_station_cnt,
    "diff": top_station_cnt - prev_station_cnt,
}

# 卡4：非機台問題佔比
nm_cur = wk[wk["問題類型"] != "機台問題類型"]
nm_pct = round(len(nm_cur) / total * 100, 1) if total else 0
nm_prev = wk_prev[wk_prev["問題類型"] != "機台問題類型"]
nm_prev_pct = round(len(nm_prev) / prev_total * 100, 1) if prev_total else 0
sample_row = nm_cur.dropna(subset=["用戶內容"]).iloc[0] if len(nm_cur.dropna(subset=["用戶內容"])) else None
sample_type = sample_row["問題細項"] if sample_row is not None else ""
card4 = {
    "pct": nm_pct, "pct_delta": round(nm_pct - nm_prev_pct, 1),
    "sample_type": sample_type,
}

statCards = {
    "card1": card1, "card2": card2, "card3": card3, "card4": card4,
    "wow_total_pct": wow_total_pct, "mom_total_pct": mom_total_pct,
}

# ================= 輸出 =================
data = {
    "week": WEEK_LABEL,
    "range": range_label,
    "total": total,
    "cats": cats,
    "trend": trend,
    "nonMachine": nonMachine,
    "machIssues": machIssues,
    "bottleTop3": bottleTop3,
    "batTop3": batTop3,
    "bottleTotal": bottle_total,
    "batTotal": bat_total,
    "machineTotal": machine_total,
    "bottlePct": round(bottle_total / machine_total * 100, 1) if machine_total else 0,
    "batPct": round(bat_total / machine_total * 100, 1) if machine_total else 0,
    "hotAreas": hotAreas,
    "alertText": alertText,
    "lowVolumeStations": lowVolumeStations,
    "totalNetworkStations": total_network_stations,
    "dataSourceNote": dataSourceNote,
    "gradeTiers": gradeTiers,
    "arkNote": arkNote,
    "statCards": statCards,
    "reportGeneratedDate": datetime.now().strftime("%Y/%m/%d"),
    "dataPeriodLabel": f"{(datetime.now() - timedelta(days=30)).strftime('%Y/%m/%d')}-{datetime.now().strftime('%m/%d')}",
}

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("完成，總件數：", total)
print("收瓶機：", bottle_total, "電池機：", bat_total)
print("低回收量站點筆數：", len(lowVolumeStations))
