# Changelog

## v5　2026-06-05　pixel-accurate layout

### 變更
- 版面座標依參考圖像素精確量測（1399×798px → 10"×5.625" 換算）
- Title: `fontSize:16 bold #333333`，無背景色塊
- 趨勢表 `rowH:0.37`，共 5 列 × 0.37" = 1.85"
- PAGE1 Top3 卡片件數・% `fontSize:18 w:2.42`（解決換行問題）
- PAGE2 大%卡片 `fontSize:32`（解決 93% 換行問題）
- PAGE2 Top3 %文字 `x:5.26 w:0.78`（解決截斷問題）
- 熱點站點僅顯示前3城市，完整呈現不截斷

## v4　2026-06-04　ecoco VI 色系 + Noto Sans TC

### 變更
- 全頁色彩切換至 ecoco VI 規範色系（Orange/Blue/Yellow/LightBlue/Beige/DarkBlue）
- 字型統一 `Noto Sans TC`
- 圖例件數與% 改用 `#333333`（深灰，非問題色）
- 移除 Header/Footer 藍色色塊與 ecoco logo

## v3　2026-06-03　版面重構

### 變更
- 移除 PNG banner，改用純色 Header
- 移除所有水平分隔線
- PAGE2 圓餅圖使用 pres.addChart 原生圖表

## v2　2026-06-02　初版 PPT 自動產出

### 新增
- pptxgenjs 兩頁 PPT 自動產出
- PAGE1：趨勢表 + 圓餅 + 非機台 Top3
- PAGE2：機台圓餅 + 收瓶/電池 Top3 + 熱點站點
