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

# ---------- 資料新鮮度防呆檢查 ----------
# 目的：避免忘記更新「客訴內容分析.csv」時，程式仍靜默用舊資料重複產出週報而沒人發現。
# 邏輯：比對資料中最新一筆日期與「今天」的天數差，超過門檻(預設10天，可在 config.json 用
#       "data_staleness_warn_days" 覆蓋)就視為過期，並把警告寫進 data.json 供第一頁顯示，
#       同時印出醒目的 console 警告(會進 run_weekly.log)，但不中斷產出流程。
STALENESS_WARN_DAYS = CONFIG.get("data_staleness_warn_days", 10)
latest_data_date = df["dt"].max()
data_staleness = {"isStale": False, "latestDataDate": None, "daysStale": None, "message": None}
if pd.notna(latest_data_date):
    days_stale = (pd.Timestamp(datetime.now().date()) - latest_data_date.normalize()).days
    data_staleness["latestDataDate"] = latest_data_date.strftime("%Y-%m-%d")
    data_staleness["daysStale"] = int(days_stale)
    if days_stale > STALENESS_WARN_DAYS:
        data_staleness["isStale"] = True
        data_staleness["message"] = (
            f"客訴內容分析.csv 最新資料日期為 {latest_data_date.strftime('%Y-%m-%d')}，"
            f"距今已 {days_stale} 天沒有更新，本次可能是用舊資料產出，請確認來源檔是否忘記更新。"
        )
        print(f"\n{'!' * 60}")
        print(f"警告：{data_staleness['message']}")
        print(f"{'!' * 60}\n")
else:
    data_staleness["message"] = "客訴內容分析.csv 的「進件日期」欄位無法解析出任何有效日期，請確認來源檔格式。"
    data_staleness["isStale"] = True
    print(f"警告：{data_staleness['message']}")

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
    {"grade": "S", "bonus": 18, "badge": "S+18", "range": "17,000瓶以上"},
    {"grade": "A", "bonus": 12, "badge": "A+12", "range": "13,300～16,999瓶"},
    {"grade": "B", "bonus": 8, "badge": "B+8", "range": "11,100～13,299瓶"},
    {"grade": "C", "bonus": 4, "badge": "C+4", "range": "9,000～11,099瓶"},
    {"grade": "D", "bonus": 2, "badge": "D+2", "range": "6,900～8,999瓶"},
    {"grade": "E", "bonus": 1, "badge": "E+1", "range": "4,100～6,899瓶"},
    {"grade": "F", "bonus": 0, "badge": "F+0", "range": "4,100瓶以下"},
    {"grade": "ARK", "bonus": 8, "badge": "ARK+8", "range": "3台有效收瓶機"},
]
dataSourceNote = DEFAULT_DATA_SOURCE_TEXT
gradeTiers = DEFAULT_GRADE_TIERS
tierSectionLabel = "等級+加分/瓶量："

if DATA_SOURCE_MD and os.path.exists(DATA_SOURCE_MD):
    try:
        with open(DATA_SOURCE_MD, "r", encoding="utf-8") as f:
            md_text = f.read()
        first_line = next((ln.strip() for ln in md_text.splitlines() if ln.strip()), None)
        if first_line:
            dataSourceNote = first_line

        # 動態解析「等級/瓶量/加分」門檻，例如： S /17000~/+18 。A/13300~16999/+12。...
        tier_pattern = r"([A-Z])\s*/\s*([\d,]*)\\?~([\d,]*)/\+(\d+)"
        matches = re.findall(tier_pattern, md_text)
        parsed_tiers = []
        if matches:
            for g, lo, hi, bonus in matches:
                if lo and hi:
                    rng = f"{int(lo):,}～{int(hi):,}瓶"
                elif lo and not hi:
                    rng = f"{int(lo):,}瓶以上"
                elif hi and not lo:
                    rng = f"{int(hi):,}瓶以下"
                else:
                    rng = ""
                parsed_tiers.append({"grade": g, "bonus": int(bonus), "badge": f"{g}+{bonus}", "range": rng})

        # 解析 ARK 特殊加分，例如：ARK/3台有效收瓶機/+8 -> 比照 S~F 同樣的「等級+分數／描述」呈現方式
        ark_match = re.search(r"ARK\s*/\s*([^/]+?)\s*/\s*\+(\d+)", md_text)
        if ark_match:
            ark_desc = ark_match.group(1).strip()
            ark_bonus = int(ark_match.group(2))
            parsed_tiers.append({"grade": "ARK", "bonus": ark_bonus, "badge": f"ARK+{ark_bonus}", "range": ark_desc})

        if parsed_tiers:
            gradeTiers = parsed_tiers
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

# ================= 11. 第四頁：客訴趨勢洞察（MoM／QoQ）=================
# 「最近一個完整月」＝目前月份的前一個月（當月尚未走完，用當月資料會失真，故一律往前推一個月）
today_ts = pd.Timestamp(datetime.now().date())
last_complete_month = (today_ts.replace(day=1) - pd.Timedelta(days=1)).to_period("M")
prev_month = last_complete_month - 1
earliest_period = df["dt"].dropna().dt.to_period("M").min()


def month_slice_p4(period):
    return df[df["dt"].dt.to_period("M") == period]


def month_stats(period):
    sub = month_slice_p4(period)
    tot = len(sub)
    vc = sub["問題類型"].value_counts()
    cats_m = {label: int(vc.get(key, 0)) for label, key in CATS_ORDER}
    return {"period": str(period), "total": tot, "cats": cats_m}


cur_month_stats = month_stats(last_complete_month)
prev_month_stats = month_stats(prev_month) if prev_month >= earliest_period else None
mom_pct = (
    round((cur_month_stats["total"] - prev_month_stats["total"]) / prev_month_stats["total"] * 100, 1)
    if prev_month_stats and prev_month_stats["total"] else None
)

# 「每站平均客訴量」：萬人客訴率的暫代指標（目前資料源無站點覆蓋用戶數，待有來源後可直接換算公式，介面不變）
active_station_count = len(vol_master_df) if vol_master_df is not None else None


def per_station(total_count):
    return round(total_count / active_station_count, 2) if active_station_count else None


cur_per_station = per_station(cur_month_stats["total"])
prev_per_station = per_station(prev_month_stats["total"]) if prev_month_stats else None
per_station_delta_pct = (
    round((cur_per_station - prev_per_station) / prev_per_station * 100, 1)
    if cur_per_station is not None and prev_per_station else None
)

# 近6個完整月趨勢（動態抓最近6個月，若資料起點不足6個月則有多少畫多少，不補0避免誤導）
monthly_trend = []
for i in range(5, -1, -1):
    p = last_complete_month - i
    if p < earliest_period:
        continue
    stats = month_stats(p)
    monthly_trend.append({
        "period": str(p), "label": f"{p.month}月",
        "total": stats["total"], "perStation": per_station(stats["total"]),
    })

# 季度判斷：找出「最近一個已完整走完的季度」(季末月份為3/6/9/12月)
m = last_complete_month
while m.month % 3 != 0:
    m -= 1
last_complete_quarter_end = m
prev_quarter_end = last_complete_quarter_end - 3


def quarter_label(end_period):
    q_num = end_period.month // 3
    return f"{end_period.year % 100}年Q{q_num}"


def quarter_stats(end_period):
    if end_period < earliest_period:
        return None
    start_period = max(end_period - 2, earliest_period)
    mask = (df["dt"].dt.to_period("M") >= start_period) & (df["dt"].dt.to_period("M") <= end_period)
    sub = df[mask]
    tot = len(sub)
    vc = sub["問題類型"].value_counts()
    cats_q = []
    for label, key in CATS_ORDER:
        c = int(vc.get(key, 0))
        cats_q.append({"label": label, "count": c, "pct": round(c / tot * 100, 1) if tot else 0})
    return {"total": tot, "cats": cats_q}


cur_quarter_stats = quarter_stats(last_complete_quarter_end)
prev_quarter_stats = quarter_stats(prev_quarter_end)
qoq_pct = (
    round((cur_quarter_stats["total"] - prev_quarter_stats["total"]) / prev_quarter_stats["total"] * 100, 1)
    if cur_quarter_stats and prev_quarter_stats and prev_quarter_stats["total"] else None
)

machine_pct_cur = next((c["pct"] for c in cur_quarter_stats["cats"] if c["label"] == "機台問題"), 0) if cur_quarter_stats else 0
machine_pct_prev = next((c["pct"] for c in prev_quarter_stats["cats"] if c["label"] == "機台問題"), 0) if prev_quarter_stats else 0
machine_pct_delta = round(machine_pct_cur - machine_pct_prev, 1) if prev_quarter_stats else None

# 找出 QoQ 佔比變化最大的問題類型（動態計算，不寫死類別名稱）
biggest_change = None
if cur_quarter_stats and prev_quarter_stats:
    for cc, pc in zip(cur_quarter_stats["cats"], prev_quarter_stats["cats"]):
        delta = round(cc["pct"] - pc["pct"], 1)
        if biggest_change is None or abs(delta) > abs(biggest_change["delta"]):
            biggest_change = {"label": cc["label"], "delta": delta, "cur_pct": cc["pct"], "prev_pct": pc["pct"]}

# 本季結構變化重點文字（依實際變化最大項動態生成，不假設成因，避免誤導）
# 註：判斷「季節性」須有至少一年（同季 YoY）以上資料比對，目前資料起點為2026年，尚不足以下此結論，
#     故此處僅如實描述「本季 vs 上季」的結構變化事實，不冠上「季節性」等因果推論字眼
quarters_available = df["dt"].dropna().dt.to_period("Q").nunique()
if biggest_change and biggest_change["delta"] > 0:
    structural_insight = (
        f"本季「{biggest_change['label']}」佔比較上季上升 {biggest_change['delta']} 個百分點"
        f"（{biggest_change['prev_pct']}% → {biggest_change['cur_pct']}%），為本季結構變化最大的問題類型，建議追蹤後續走勢並評估對應改善措施。"
    )
elif biggest_change and biggest_change["delta"] < 0:
    structural_insight = (
        f"本季「{biggest_change['label']}」佔比較上季下降 {abs(biggest_change['delta'])} 個百分點"
        f"（{biggest_change['prev_pct']}% → {biggest_change['cur_pct']}%），為本季結構變化最大的問題類型，改善成效可留意是否延續。"
    )
else:
    structural_insight = "本季各問題類型佔比與上季相近，未見明顯結構性變化。"
if quarters_available < 5:
    structural_insight += "（資料累積中，須滿一年以上同季比較才能判斷是否為季節性規律）"

# 執行摘要（一句話重點，給主管/跨部門一眼看懂本月+本季整體方向）
def _dir_word(pct):
    if pct is None:
        return None, None
    if pct < 0:
        return "下降", abs(pct)
    if pct > 0:
        return "上升", pct
    return "持平", 0

mom_dir, mom_abs = _dir_word(mom_pct)
qoq_dir, qoq_abs = _dir_word(qoq_pct)

if mom_pct is not None and qoq_pct is not None:
    if mom_pct <= 0 and qoq_pct <= 0:
        overall_judgement = "整體呈現改善趨勢"
    elif mom_pct >= 0 and qoq_pct >= 0:
        overall_judgement = "整體呈現上升趨勢，需留意"
    else:
        overall_judgement = "月度與季度方向不一致，建議留意近期變化"
else:
    overall_judgement = "資料累積中"

headline_parts = []
if mom_dir:
    headline_parts.append(f"本月客訴量較上月{mom_dir} {mom_abs}%")
if qoq_dir:
    headline_parts.append(f"本季較上季{qoq_dir} {qoq_abs}%")
headline = "，".join(headline_parts)
headline = f"{headline}，{overall_judgement}。" if headline else "資料累積中，尚無法產出完整趨勢摘要。"
if biggest_change and biggest_change["delta"] != 0:
    bc_dir = "上升" if biggest_change["delta"] > 0 else "下降"
    headline += f" 本季變化最大：「{biggest_change['label']}」{bc_dir} {abs(biggest_change['delta'])} 個百分點，建議優先關注。"


# 系統／活動影響評估：需 activity_calendar_path（行銷活動／系統更新時間表），目前無此資料源則整段略過，不中斷報告
ACTIVITY_CSV = CONFIG.get("activity_calendar_path")
activity_insight = None
if ACTIVITY_CSV and os.path.exists(ACTIVITY_CSV):
    try:
        act_df = pd.read_csv(ACTIVITY_CSV)
        act_df.columns = act_df.columns.str.strip()
        act_df["date"] = pd.to_datetime(act_df["日期"], errors="coerce")
        month_start = last_complete_month.start_time
        month_end = last_complete_month.end_time
        acts = act_df[(act_df["date"] >= month_start) & (act_df["date"] <= month_end)]
        if len(acts):
            names = "、".join(acts["活動名稱"].astype(str).tolist())
            activity_insight = f"本月活動／系統異動：{names}，建議比對客訴日期分布確認是否有關聯。"
        else:
            activity_insight = "本月無登記活動或系統異動紀錄。"
    except Exception as e:
        print(f"警告：讀取 {ACTIVITY_CSV} 失敗（{e}），系統/活動影響評估段落略過。")

# ================= 12. 高風險站點長效追蹤（連續2個月以上進入月低回收量清單）=================
LOW_VOLUME_HISTORY_PATH = os.path.join(SCRIPT_DIR, "monthly_low_volume_history.json")
low_vol_history = {}
if os.path.exists(LOW_VOLUME_HISTORY_PATH):
    with open(LOW_VOLUME_HISTORY_PATH, "r", encoding="utf-8") as f:
        low_vol_history = json.load(f)

current_month_key2 = datetime.now().strftime("%Y-%m")
current_low_names = [s["name"] for s in lowVolumeStations]
low_vol_history[current_month_key2] = current_low_names
keep_months = sorted(low_vol_history.keys())[-12:]  # 只保留近12個月，避免檔案無限長大
low_vol_history = {k: low_vol_history[k] for k in keep_months}

with open(LOW_VOLUME_HISTORY_PATH, "w", encoding="utf-8") as f:
    json.dump(low_vol_history, f, ensure_ascii=False, indent=2)

sorted_months_desc = sorted(low_vol_history.keys(), reverse=True)
highRiskStations = []
for name in current_low_names:
    streak = 0
    for m_key in sorted_months_desc:
        if name in low_vol_history.get(m_key, []):
            streak += 1
        else:
            break
    if streak >= 2:
        grade = next((s["grade"] for s in lowVolumeStations if s["name"] == name), None)
        highRiskStations.append({"name": name, "streak": streak, "grade": grade})
highRiskStations.sort(key=lambda s: -s["streak"])

page4 = {
    "kpi": {
        "monthTotal": {
            "value": cur_month_stats["total"], "period": f"{last_complete_month.year % 100}年{last_complete_month.month}月",
            "deltaPct": mom_pct, "prevValue": prev_month_stats["total"] if prev_month_stats else None,
        },
        "perStation": {
            "value": cur_per_station, "deltaPct": per_station_delta_pct, "prevValue": prev_per_station,
        },
        "quarterTotal": {
            "value": cur_quarter_stats["total"] if cur_quarter_stats else None,
            "quarterLabel": quarter_label(last_complete_quarter_end),
            "deltaPct": qoq_pct,
            "prevValue": prev_quarter_stats["total"] if prev_quarter_stats else None,
        },
        "machinePct": {"value": machine_pct_cur, "deltaPct": machine_pct_delta},
    },
    "monthlyTrend": monthly_trend,
    "quarterCompare": {
        "curLabel": quarter_label(last_complete_quarter_end),
        "prevLabel": quarter_label(prev_quarter_end) if prev_quarter_stats else None,
        "cur": cur_quarter_stats["cats"] if cur_quarter_stats else [],
        "prev": prev_quarter_stats["cats"] if prev_quarter_stats else [],
        "biggestChange": biggest_change,
    },
    "structuralInsight": structural_insight,
    "headline": headline,
    "activityInsight": activity_insight,
    "highRiskStations": highRiskStations,
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
    "tierSectionLabel": tierSectionLabel,
    "statCards": statCards,
    "page4": page4,
    "dataStaleness": data_staleness,
    "reportGeneratedDate": datetime.now().strftime("%Y/%m/%d"),
    "dataPeriodLabel": f"{(datetime.now() - timedelta(days=30)).strftime('%Y/%m/%d')}-{datetime.now().strftime('%m/%d')}",
}

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("完成，總件數：", total)
print("收瓶機：", bottle_total, "電池機：", bat_total)
print("低回收量站點筆數：", len(lowVolumeStations))
