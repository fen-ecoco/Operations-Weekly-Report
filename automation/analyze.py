import re
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ecoco 每週客訴分析腳本
每週一中午前自動執行 — Step 1 + Step 2
讀取 CSV → 計算統計 → 輸出 data.json
"""

import pandas as pd
import json
import os
import sys
from datetime import datetime

# ── 設定路徑 ──
BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE, "config.json")
HISTORY_PATH = os.path.join(BASE, "history.json")
DATA_OUT = os.path.join(BASE, "data.json")

def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)

def load_history():
    if os.path.exists(HISTORY_PATH):
        with open(HISTORY_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"weeks": []}

def save_history(history):
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def clean_df(df):
    """清理欄位、統一標籤"""
    df["機台類型"] = df["機台類型"].astype(str).str.strip().replace({"nan": "", "": None})
    df["問題類型"] = df["問題類型"].astype(str).str.strip().replace("nan", None)
    df["簡報使用"] = df["簡報使用"].astype(str).str.strip().replace("nan", "")
    df["站點區域"] = df["站點區域"].astype(str).str.strip().replace("nan", None)
    df["站點名稱"] = df["站點名稱"].astype(str).str.strip().replace("nan", None)

    label_map = {
        "機台需維護-故障提醒":       "機台需維護/故障提醒",
        "操作流程異常-無法正常操作":  "操作流程異常/無法正常操作",
        "忘記密碼-無法重設密碼":      "忘記密碼/無法重設密碼",
        "機台當機-無回應":            "機台當機/無回應",
        "螢幕異常顯示-畫面異常":      "螢幕異常顯示/畫面異常",
        "機台關閉-無法啟動":          "機台關閉/無法啟動",
        "兌換失敗-顯示錯誤":          "兌換失敗/顯示錯誤",
        "許願新增站點-設站建議":       "許願新增站點/設站建議",
        "使用規則-限制條件說明":       "使用規則/限制條件說明",
        "app畫面顯示與機台狀態不符":  "APP畫面顯示與機台狀態不符",
        "app無法登入":                "APP無法登入",
        "app多重異常狀況":            "APP多重異常狀況",
    }
    df["簡報使用"] = df["簡報使用"].replace(label_map)
    return df

def get_week_info(df):
    """從資料日期推算週次與區間"""
    # 取進件日期
    dates = pd.to_datetime(df["進件日期"], errors="coerce")
    min_d, max_d = dates.min(), dates.max()
    week_num = min_d.isocalendar()[1]
    range_str = f"{min_d.strftime('%m/%d')} ~ {max_d.strftime('%m/%d')}"
    return f"第{week_num}週", range_str, week_num, min_d, max_d

def analyze(df):
    """主要統計運算"""
    # 只計有問題類型的記錄
    df_cat = df[df["問題類型"].notna()]
    total = len(df_cat)

    # ── 問題類型佔比（固定順序對應 VI 色序）──
    cat_order = [
        "APP帳號設定問題類型",
        "機台問題類型",
        "回收點數問題類型",
        "顧客關係類型",
        "優惠券問題類型",
        "APP使用問題類型",
    ]
    cat_labels = [
        "APP帳號設定問題", "機台問題", "回收點數問題",
        "顧客關係", "優惠券問題", "APP使用問題"
    ]
    qt = df_cat["問題類型"].value_counts()
    cats = []
    for raw, label in zip(cat_order, cat_labels):
        cnt = int(qt.get(raw, 0))
        cats.append({"label": label, "count": cnt,
                     "pct": round(cnt / total * 100, 1) if total else 0})

    # ── 機台 ──
    bottle = df[df["機台類型"].isin(["收瓶機", "方舟站收瓶機", "方舟站"])]
    bat    = df[df["機台類型"] == "電池機"]
    b_cnt  = len(bottle); bat_cnt = len(bat)
    m_tot  = b_cnt + bat_cnt or 1

    # ── 機台問題詳細 (for pie) ──
    mach = df_cat[df_cat["問題類型"] == "機台問題類型"]
    mach_vc = mach["簡報使用"].value_counts()
    top5 = mach_vc.head(5)
    other_cnt = int(mach_vc[5:].sum())
    mach_total = len(mach) or 1
    machIssues = [{"label": k, "count": int(v),
                   "pct": round(v / mach_total * 100, 1)}
                  for k, v in top5.items()]
    machIssues.append({"label": "其他機台問題", "count": other_cnt,
                       "pct": round(other_cnt / mach_total * 100, 1)})

    # ── 非機台 Top3 ──
    nm = df_cat[df_cat["問題類型"] != "機台問題類型"]
    nm_vc = nm["簡報使用"].value_counts()
    nm_colors = ["FF5000", "0076A9", "060E9F"]
    nm_notes  = ["", "", ""]  # 需人工補充說明文字
    nonMachine = []
    for i, (k, v) in enumerate(nm_vc.head(3).items()):
        nonMachine.append({
            "rank": i + 1, "name": k,
            "count": int(v),
            "pct": round(v / total * 100, 1),
            "col": nm_colors[i],
            "note": nm_notes[i]
        })

    # ── 收瓶機 Top3 ──
    b_colors = ["FF5000", "060E9F", "0076A9"]
    b_vc = bottle["簡報使用"].value_counts()
    bottleTop3 = [{"name": k, "count": int(v),
                   "pct": round(v / b_cnt * 100, 0) if b_cnt else 0,
                   "col": b_colors[i]}
                  for i, (k, v) in enumerate(b_vc.head(3).items())]

    # ── 電池機 Top ──
    bat_colors = ["0076A9", "8EB8C9"]
    bat_vc = bat["簡報使用"].value_counts()
    batTop = [{"name": k, "count": int(v),
               "pct": round(v / bat_cnt * 100, 0) if bat_cnt else 0,
               "col": bat_colors[i]}
              for i, (k, v) in enumerate(bat_vc.head(2).items())]

    # ── 熱點站點（過濾無效區域值：空值、破折號、站點編號如 es0984）──
    def is_valid_area(v):
        if not v or str(v).strip() in ["-", "--", "nan", "", "None"]:
            return False
        if re.match(r"^[a-zA-Z0-9_\-]+$", str(v).strip()):  # 純英數字 = 站點編號
            return False
        if len(str(v).strip()) < 2:
            return False
        return True

    valid_mask = df["站點區域"].apply(
        lambda x: is_valid_area(x) if pd.notna(x) else False
    )
    area_df  = df[valid_mask]
    area_top = area_df["站點區域"].value_counts().head(3)
    hotAreas = []
    for area, area_cnt in area_top.items():
        sub = area_df[area_df["站點區域"] == area]["站點名稱"].value_counts().head(3)
        spots = [
            {"name": s, "count": int(c)}
            for s, c in sub.items()
            if s and str(s).strip() not in ["-", "nan", "", "None"]
        ]
        if spots:
            hotAreas.append({
                "area": area,
                "total": int(area_cnt),
                "spots": spots
            })

    # ── 加總 ──
    summary = {
        "bt": b_cnt, "b2": bat_cnt,
        "rg": int(qt.get("APP帳號設定問題類型", 0)),
        "pt": int(qt.get("回收點數問題類型", 0)),
        "cp": int(qt.get("優惠券問題類型", 0)),
        "ap": int(qt.get("APP使用問題類型", 0)),
        "cr": int(qt.get("顧客關係類型", 0)),
        "tot": total
    }

    return {
        "total": total,
        "cats": cats,
        "machIssues": machIssues,
        "nonMachine": nonMachine,
        "bottleTop3": bottleTop3,
        "batTop": batTop,
        "bottleTotal": b_cnt, "batTotal": bat_cnt,
        "bottlePct": round(b_cnt / m_tot * 100),
        "batPct": round(bat_cnt / m_tot * 100),
        "hotAreas": hotAreas,
        "summary": summary,
        "alertText": ""   # 由人工填寫或後續邏輯自動產生
    }

def build_trend(history, week_num, summary):
    """取最近4週趨勢（含本週）"""
    weeks = history.get("weeks", [])
    # 更新或插入本週
    existing = next((w for w in weeks if w["week_num"] == week_num), None)
    if not existing:
        weeks.append({"week_num": week_num, **summary})
    else:
        existing.update(summary)
    # 排序並取最近4週
    weeks.sort(key=lambda w: w["week_num"])
    history["weeks"] = weeks
    recent4 = weeks[-4:]

    trend = []
    for i, w in enumerate(recent4):
        prev = recent4[i - 1] if i > 0 else None
        def delta(key):
            cur = w.get(key, 0)
            p   = prev.get(key, 0) if prev else cur
            return [int(cur), int(cur - p) if prev else 0]

        wn = w["week_num"]
        # 推算日期（ISO週）
        d_start = datetime.fromisocalendar(2026, wn, 1)
        d_end   = datetime.fromisocalendar(2026, wn, 7)
        trend.append({
            "w": f"{wn}週",
            "d": f"{d_start.strftime('%m/%d')}-{d_end.strftime('%m/%d')}",
            "bt": delta("bt"), "b2": delta("b2"), "rg": delta("rg"),
            "pt": delta("pt"), "cp": delta("cp"), "ap": delta("ap"),
            "cr": delta("cr"), "tot": delta("tot"),
        })
    return trend

def auto_alert(stats):
    """自動產生警示文字"""
    alerts = []
    for item in stats["nonMachine"][:2]:
        if item["pct"] >= 10:
            alerts.append(f"【注意】{item['name']} 持續高位｜本週{item['count']}件（{item['pct']}%）")
    if stats["bottleTop3"]:
        top = stats["bottleTop3"][0]
        if top["pct"] >= 25:
            alerts.append(f"收瓶機 {top['name']} {top['count']}件（{top['pct']}%）居首，請確認清空頻率")
    return "。".join(alerts) + "。" if alerts else ""

def main():
    cfg = load_config()
    csv_path = cfg["csv_path"]
    print(f"[1/4] 載入 CSV：{csv_path}")
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df = clean_df(df)

    week_label, date_range, week_num, _, _ = get_week_info(df)
    print(f"[2/4] 分析週次：{week_label}（{date_range}）")

    stats = analyze(df)
    history = load_history()
    trend = build_trend(history, week_num, stats["summary"])
    save_history(history)

    stats["alertText"] = auto_alert(stats)

    data = {
        "week":       week_label,
        "range":      date_range,
        "week_num":   week_num,
        "total":      stats["total"],
        "cats":       stats["cats"],
        "trend":      trend,
        "nonMachine": stats["nonMachine"],
        "machIssues": stats["machIssues"],
        "bottleTop3": stats["bottleTop3"],
        "batTop":     stats["batTop"],
        "bottleTotal":stats["bottleTotal"],
        "batTotal":   stats["batTotal"],
        "bottlePct":  stats["bottlePct"],
        "batPct":     stats["batPct"],
        "hotAreas":   stats["hotAreas"],
        "alertText":  stats["alertText"],
    }

    with open(DATA_OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"[3/4] data.json 輸出完成（{stats['total']}件已分類）")
    print(f"[4/4] 警示：{stats['alertText'] or '無'}")
    return week_label, date_range

if __name__ == "__main__":
    main()
