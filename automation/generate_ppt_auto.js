const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'data.json'), 'utf-8'));
const ICON = (n) => path.join(__dirname, 'icons', `${n}.png`);

// ---------- ecoco VI 色系 ----------
const C = {
  orange: 'FF5000',
  blue: '060E9F',
  yellow: 'FFCE00',
  beige: 'FAE0B8',
  lightBlue: '8EB9C9',
  darkBlue: '0076A9',
  darkGray: '333333',
  textGray: '888888',
  white: 'FFFFFF',
};

const F_BLACK = 'Noto Sans TC Black';
const F_BOLD = 'Noto Sans TC Bold';
const F_MED = 'Noto Sans TC Medium';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.3 x 7.5
const PW = 13.3, PH = 7.5;

function sectionBar(slide, text, x, y, w, opts = {}) {
  const barColor = opts.barColor || C.orange;
  const textColor = opts.textColor || C.blue;
  slide.addShape('rect', { x, y, w: 0.07, h: 0.28, fill: { color: barColor }, line: { type: 'none' } });
  slide.addText(text, {
    x: x + 0.15, y: y - 0.03, w: w - 0.15, h: 0.34,
    fontFace: F_BLACK, fontSize: 13, bold: true, color: textColor,
    align: 'left', valign: 'middle', margin: 0,
  });
}

// ============================================================
// SLIDE 1 — 客訴趨勢分析
// ============================================================
function buildSlide1() {
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // ---- 標題 + 週次橢圓徽章（右上角） ----
  slide.addText('ecoco 客服週報', {
    x: 0.4, y: 0.16, w: 5, h: 0.4,
    fontFace: F_BLACK, fontSize: 18, bold: true, color: C.darkGray, margin: 0,
  });
  const pillW = 2.7, pillH = 0.42;
  slide.addShape('roundRect', {
    x: PW - 0.4 - pillW, y: 0.18, w: pillW, h: pillH, rectRadius: pillH / 2,
    fill: { color: C.blue }, line: { type: 'none' },
  });
  slide.addText(`${D.week}（${D.range}）`, {
    x: PW - 0.4 - pillW, y: 0.18, w: pillW, h: pillH,
    fontFace: F_BOLD, fontSize: 12.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0,
  });

  // ---- 資料新鮮度警告（僅在偵測到來源CSV過期時顯示，放在標題與週次徽章之間的空白處，不影響其餘版面）----
  const DS = D.dataStaleness;
  if (DS && DS.isStale) {
    const warnX = 5.55, warnW = PW - 0.4 - pillW - 0.15 - warnX;
    slide.addShape('roundRect', {
      x: warnX, y: 0.18, w: warnW, h: pillH, rectRadius: 0.08,
      fill: { color: 'FDECEC' }, line: { color: 'D9534F', width: 1 },
    });
    slide.addText(
      DS.latestDataDate
        ? `⚠ 資料可能過期：來源CSV最新日期 ${DS.latestDataDate}，已 ${DS.daysStale} 天未更新，請確認`
        : `⚠ ${DS.message}`,
      {
        x: warnX + 0.12, y: 0.18, w: warnW - 0.24, h: pillH,
        fontFace: F_BOLD, fontSize: 9.5, bold: true, color: 'C0392B', valign: 'middle', margin: 0,
      }
    );
  }

  // ---- 頂部四格數據卡 ----
  const cardsY = 0.75, cardsH = 1.52, cardGap = 0.2;
  const cardW = (PW - 0.8 - cardGap * 3) / 4;
  const SC = D.statCards;

  function trendArrow(val, goodWhenDown = true) {
    // 回傳 {text, color} ； 數值上升 -> ▲ 橘；下降 -> ▼ 深藍（與趨勢表配色一致）
    if (val > 0) return { text: `▲${val}%`, color: C.orange };
    if (val < 0) return { text: `▼${Math.abs(val)}%`, color: C.darkBlue };
    return { text: '—', color: C.textGray };
  }

  function statCard(x, accent, icon, title, mainNode, subText, badge) {
    slide.addShape('roundRect', {
      x, y: cardsY, w: cardW, h: cardsH, rectRadius: 0.09,
      fill: { color: 'FFFFFF' }, line: { color: 'E5E5E5', width: 1 },
      shadow: { type: 'outer', color: '999999', opacity: 0.2, blur: 5, offset: 1.5, angle: 90 },
    });
    slide.addShape('roundRect', {
      x, y: cardsY, w: cardW, h: 0.07, rectRadius: 0.03,
      fill: { color: accent }, line: { type: 'none' },
    });
    slide.addText(title, {
      x: x + 0.2, y: cardsY + 0.16, w: cardW - 0.7, h: 0.34,
      fontFace: F_BOLD, fontSize: 12, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
    });
    slide.addImage({ path: ICON(icon), x: x + cardW - 0.48, y: cardsY + 0.18, w: 0.28, h: 0.28 });

    slide.addText(mainNode, {
      x: x + 0.2, y: cardsY + 0.54, w: cardW - 0.4, h: 0.52, valign: 'middle', margin: 0,
    });

    slide.addText(subText, {
      x: x + 0.2, y: cardsY + 1.10, w: cardW - 1.1, h: 0.55,
      fontFace: F_MED, fontSize: 9.2, color: C.textGray, valign: 'top', margin: 0, lineSpacingMultiple: 1.08,
    });
    if (badge) {
      slide.addShape('roundRect', {
        x: x + cardW - 1.02, y: cardsY + 1.10, w: 0.84, h: 0.32, rectRadius: 0.16,
        fill: { color: badge.color === C.orange ? 'FFE9DC' : 'DDEAF5' }, line: { type: 'none' },
      });
      slide.addText(badge.text, {
        x: x + cardW - 1.02, y: cardsY + 1.10, w: 0.84, h: 0.32,
        fontFace: F_BOLD, fontSize: 10.2, bold: true, color: badge.color, align: 'center', valign: 'middle', margin: 0,
      });
    }
  }

  // 卡1：本週客訴總件數
  {
    const x = 0.4;
    const d = SC.card1;
    const diffText = d.diff >= 0 ? `增加 ${d.diff} 件` : `減少 ${Math.abs(d.diff)} 件`;
    statCard(x, C.blue, 'headset_blue', `本週客訴總件數 (${D.week})`,
      [{ text: `${d.value} `, options: { fontSize: 27, bold: true, color: C.blue, fontFace: F_BLACK } },
       { text: '件', options: { fontSize: 13.5, bold: true, color: C.blue, fontFace: F_BOLD } }],
      `較上週 (${d.prev}件) ${diffText}`,
      trendArrow(d.wow_pct));
  }
  // 卡2：回報最高主題
  {
    const x = 0.4 + (cardW + cardGap);
    const d = SC.card2;
    statCard(x, C.orange, 'warn_orange', '回報最高主題',
      [{ text: d.name, options: { fontSize: 17, bold: true, color: C.orange, fontFace: F_BLACK } }],
      `佔比 ${d.share}%（共 ${d.count} 件）`,
      trendArrow(d.share_delta));
  }
  // 卡3：客訴量最高站點
  {
    const x = 0.4 + 2 * (cardW + cardGap);
    const d = SC.card3;
    statCard(x, C.darkBlue, 'pin_darkblue', '客訴量最高站點',
      [{ text: d.name.replace(/站$/, ''), options: { fontSize: 15.5, bold: true, color: C.darkBlue, fontFace: F_BLACK } }],
      `${d.area}市（本週 ${d.count} 件）`,
      { text: (d.diff >= 0 ? `▲${d.diff}件` : `▼${Math.abs(d.diff)}件`), color: d.diff >= 0 ? C.orange : C.darkBlue });
  }
  // 卡4：非機台問題佔比
  {
    const x = 0.4 + 3 * (cardW + cardGap);
    const d = SC.card4;
    statCard(x, C.yellow, 'usercog_gold', '非機台問題佔比',
      [{ text: `${d.pct}`, options: { fontSize: 27, bold: true, color: '9C7A00', fontFace: F_BLACK } },
       { text: '%', options: { fontSize: 13.5, bold: true, color: '9C7A00', fontFace: F_BOLD } }],
      `代表案例類型：${d.sample_type}`,
      trendArrow(d.pct_delta));
  }

  // ---- 左下：客訴趨勢分析（近4週）單位：件數 ----
  const bottomY = cardsY + cardsH + 0.42;
  const leftW = 7.8, colGap = 0.3, rightX = 0.4 + leftW + colGap, rightW = PW - 0.4 - rightX;
  sectionBar(slide, '客訴趨勢分析（近4週）　單位：件數', 0.4, bottomY, leftW);

  const catKeysShort = ['machine', 'app_acc', 'points', 'customer', 'coupon', 'app_use'];
  const catLabelsShort = ['機台問題', 'APP帳號設定問題', '回收點數問題', '顧客關係', '優惠券問題', 'APP使用問題'];
  const cw2 = [1.12, 0.95, 0.95, 0.95, 0.95, 0.95, 0.95, 0.99];
  const rows2 = [];
  rows2.push(['週次', ...catLabelsShort, '加總'].map(t => ({
    text: t, options: { fill: { color: C.blue }, color: C.white, bold: true, fontFace: F_BOLD, fontSize: 11, align: 'center', valign: 'middle' },
  })));
  D.trend.forEach((row, idx) => {
    const isLast = idx === D.trend.length - 1;
    const rowFill = isLast ? C.beige : (idx % 2 === 0 ? 'FFFFFF' : 'F7F7F9');
    const cells = [{
      text: row.w,
      options: { fill: { color: rowFill }, color: C.darkGray, bold: isLast, fontFace: isLast ? F_BOLD : F_MED, fontSize: 10.5, align: 'center', valign: 'middle' },
    }];
    catKeysShort.forEach((k) => {
      const [val, delta] = row[k];
      let deltaStr = '', deltaColor = C.textGray;
      if (delta > 0) { deltaStr = `▲${delta}`; deltaColor = C.orange; }
      else if (delta < 0) { deltaStr = `▼${Math.abs(delta)}`; deltaColor = C.darkBlue; }
      cells.push({
        text: [
          { text: `${val}`, options: { fontSize: 13.5, bold: true, color: C.darkGray, fontFace: F_BOLD } },
          { text: deltaStr ? `\n${deltaStr}` : '', options: { fontSize: 9, color: deltaColor, fontFace: F_MED } },
        ],
        options: { fill: { color: rowFill }, align: 'center', valign: 'middle' },
      });
    });
    const [tval, tdelta] = row.tot;
    const tDeltaStr = tdelta > 0 ? `▲${tdelta}` : (tdelta < 0 ? `▼${Math.abs(tdelta)}` : '');
    cells.push({
      text: [
        { text: `${tval}`, options: { fontSize: 13.5, bold: true, color: C.white, fontFace: F_BOLD } },
        { text: tDeltaStr ? `\n${tDeltaStr}` : '', options: { fontSize: 9, color: C.white, fontFace: F_MED } },
      ],
      options: { fill: { color: C.orange }, align: 'center', valign: 'middle' },
    });
    rows2.push(cells);
  });
  slide.addTable(rows2, {
    x: 0.4, y: bottomY + 0.38, w: leftW, h: 3.3, rowH: 0.66,
    colW: cw2, border: { type: 'solid', color: 'E5E5E5', pt: 0.5 }, autoPage: false,
  });

  // ---- 右下：非機台問題 Top3 ----
  sectionBar(slide, '非機台問題 Top 3', rightX, bottomY, rightW, { textColor: C.orange });
  const cardH3 = 1.3, cardGap3 = 0.17;
  const cardTopY = bottomY + 0.38;
  const cardBg3 = ['FFF4EF', 'EEF6FB', 'F5F7FF'];
  D.nonMachine.forEach((item, i) => {
    const cy = cardTopY + i * (cardH3 + cardGap3);
    slide.addShape('roundRect', {
      x: rightX, y: cy, w: rightW, h: cardH3, rectRadius: 0.08,
      fill: { color: cardBg3[i] }, line: { color: item.col, width: 1 },
    });
    slide.addShape('rect', { x: rightX, y: cy, w: 0.06, h: cardH3, fill: { color: item.col }, line: { type: 'none' } });
    slide.addText(`${item.rank}`, {
      x: rightX + 0.16, y: cy + 0.12, w: 0.5, h: 0.46,
      fontFace: F_BLACK, fontSize: 17, bold: true, color: item.col, margin: 0,
    });
    slide.addText(item.name, {
      x: rightX + 0.68, y: cy + 0.1, w: rightW - 0.9, h: 0.4,
      fontFace: F_BOLD, fontSize: 12.5, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
    });
    slide.addText([
      { text: `${item.count}件`, options: { fontSize: 17, bold: true, color: item.col, fontFace: F_BLACK } },
      { text: `　${item.pct}%`, options: { fontSize: 11.5, bold: true, color: item.col, fontFace: F_BOLD } },
    ], { x: rightX + 0.68, y: cy + 0.52, w: rightW - 1.0, h: 0.4, valign: 'middle', margin: 0 });
    const maxCount = D.nonMachine[0].count;
    const barW = (rightW - 0.9) * (item.count / maxCount);
    slide.addShape('rect', { x: rightX + 0.68, y: cy + cardH3 - 0.22, w: rightW - 0.9, h: 0.07, fill: { color: 'E8E8E8' }, line: { type: 'none' } });
    slide.addShape('rect', { x: rightX + 0.68, y: cy + cardH3 - 0.22, w: barW, h: 0.07, fill: { color: item.col }, line: { type: 'none' } });
  });
}

// ============================================================
// SLIDE 2 — 機台類型客訴與高頻站點融合分析（圖示版）
// ============================================================
function buildSlide2() {
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  slide.addShape('rect', { x: 0.4, y: 0.22, w: 0.08, h: 0.36, fill: { color: C.darkBlue }, line: { type: 'none' } });
  slide.addText('客訴機台類型與高頻站點分析', {
    x: 0.58, y: 0.16, w: 9, h: 0.48,
    fontFace: F_BLACK, fontSize: 19, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
  });
  const pillW2 = 2.7, pillH2 = 0.42;
  slide.addShape('roundRect', {
    x: PW - 0.4 - pillW2, y: 0.18, w: pillW2, h: pillH2, rectRadius: pillH2 / 2,
    fill: { color: C.blue }, line: { type: 'none' },
  });
  slide.addText(`${D.week}（${D.range}）`, {
    x: PW - 0.4 - pillW2, y: 0.18, w: pillW2, h: pillH2,
    fontFace: F_BOLD, fontSize: 12.5, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0,
  });

  const cardY = 0.95, cardH = 4.55, cardGap = 0.35;
  const cardW = (PW - 0.8 - cardGap) / 2;
  const leftX = 0.4, rightX = leftX + cardW + cardGap;

  function machineCard(x, title, icon, badgeText, badgeColor, top3, cardBorder, headColor) {
    const badgeTextColor = badgeColor === C.yellow ? C.blue : C.white;
    slide.addShape('roundRect', {
      x, y: cardY, w: cardW, h: cardH, rectRadius: 0.1,
      fill: { color: 'FFFFFF' }, line: { color: 'E2E2E2', width: 1 },
      shadow: { type: 'outer', color: '999999', opacity: 0.25, blur: 6, offset: 2, angle: 90 },
    });
    // header
    slide.addImage({ path: ICON(icon), x: x + 0.28, y: cardY + 0.22, w: 0.38, h: 0.38 });
    slide.addText(title, {
      x: x + 0.76, y: cardY + 0.17, w: 3.0, h: 0.46,
      fontFace: F_BLACK, fontSize: 18, bold: true, color: headColor, valign: 'middle', margin: 0,
    });
    slide.addShape('roundRect', {
      x: x + cardW - 2.0, y: cardY + 0.22, w: 1.7, h: 0.4, rectRadius: 0.2,
      fill: { color: badgeColor }, line: { type: 'none' },
    });
    slide.addText(badgeText, {
      x: x + cardW - 2.0, y: cardY + 0.22, w: 1.7, h: 0.4,
      fontFace: F_BOLD, fontSize: 11.5, bold: true, color: badgeTextColor, align: 'center', valign: 'middle', margin: 0,
    });

    // top3 list（縮減內襯與間距，加大字級）
    const headerOffset = 0.75, blockGap = 0.11;
    let iy = cardY + headerOffset;
    const itemH = (cardH - headerOffset) / 3;
    top3.forEach((item) => {
      slide.addShape('roundRect', {
        x: x + 0.24, y: iy, w: cardW - 0.48, h: itemH - blockGap, rectRadius: 0.06,
        fill: { color: 'F8F9FB' }, line: { type: 'none' },
      });
      slide.addText(`${item.rank}`, {
        x: x + 0.36, y: iy + 0.11, w: 0.42, h: 0.42,
        fontFace: F_BLACK, fontSize: 18, bold: true, color: item.col, margin: 0,
      });
      slide.addText([
        { text: item.name, options: { fontSize: 16, bold: true, color: C.darkGray, fontFace: F_BOLD } },
        { text: `　合計 ${item.pct}% (${item.count}件)`, options: { fontSize: 14, bold: true, color: item.col, fontFace: F_BOLD } },
      ], {
        x: x + 0.85, y: iy + 0.06, w: cardW - 1.15, h: 0.5, valign: 'top', margin: 0,
      });
      slide.addImage({ path: ICON('pin_gray'), x: x + 0.85, y: iy + 0.60, w: 0.15, h: 0.15 });
      const stationTxt = item.station ? `最高回報站點：${item.station}（${item.stationCount}件）` : '無站點資料';
      slide.addText(stationTxt, {
        x: x + 1.05, y: iy + 0.55, w: cardW - 1.35, h: 0.26,
        fontFace: F_MED, fontSize: 10.5, color: C.textGray, valign: 'middle', margin: 0,
      });
      iy += itemH;
    });
  }

  machineCard(
    leftX, '收瓶機 ／ 方舟', 'recycle_blue',
    `${D.bottleTotal}件 (${D.bottlePct}%)`, C.orange,
    D.bottleTop3, 'E2E2E2', C.blue
  );
  machineCard(
    rightX, '二代電池機', 'battery_orange',
    `${D.batTotal}件 (${D.batPct}%)`, C.yellow,
    D.batTop3, 'E2E2E2', C.blue
  );

  // alert box
  const alertY = cardY + cardH + 0.25;
  slide.addShape('roundRect', {
    x: 0.4, y: alertY, w: PW - 0.8, h: 0.85, rectRadius: 0.08,
    fill: { color: C.beige }, line: { color: C.yellow, width: 1.25 },
  });
  slide.addImage({ path: ICON('warn'), x: 0.62, y: alertY + 0.24, w: 0.34, h: 0.34 });
  slide.addText(D.alertText, {
    x: 1.1, y: alertY + 0.1, w: PW - 1.6, h: 0.65,
    fontFace: F_MED, fontSize: 10.5, color: '6B4800', valign: 'middle', margin: 0, lineSpacingMultiple: 1.15,
  });
}

// ============================================================
// SLIDE 3 — 月低回收量站點分析
// ============================================================
function buildSlide3() {
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  slide.addShape('rect', { x: 0.4, y: 0.22, w: 0.08, h: 0.36, fill: { color: C.orange }, line: { type: 'none' } });
  slide.addText('月低回收量站點改善清單 Top 10', {
    x: 0.58, y: 0.16, w: 9, h: 0.48,
    fontFace: F_BLACK, fontSize: 19, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
  });
  slide.addText(`資料更新日期：${D.reportGeneratedDate}（區間${D.dataPeriodLabel}）`, {
    x: 7.3, y: 0.28, w: 5.6, h: 0.32,
    fontFace: F_BOLD, fontSize: 11, bold: true, color: C.darkBlue, align: 'right', margin: 0,
  });

  // table：等級／Hive排名／縣市／站點名稱／機型／租賃買斷／近30日合計(瓶)／月增減率
  const tX = 0.4, tY = 0.85, tW = PW - 0.8;
  const cw = [0.75, 1.5, 1.0, 2.65, 1.5, 1.7, 1.8, 1.6];
  const headerH = 0.5, dataRowH = 0.42;
  const gradeColor = { S: C.orange, A: C.darkBlue, B: C.blue, C: '4F86A0', D: C.textGray, E: C.textGray, F: C.textGray, ARK: '9C7A00' };
  const headerCells = ['等級', `Hive排名\n（總${D.totalNetworkStations ?? '－'}）`, '縣市', '站點名稱', '機型', '租賃/買斷', '近30日合計（瓶）', '月增減率'].map(t => ({
    text: t, options: { fill: { color: C.blue }, color: C.white, bold: true, fontFace: F_BOLD, fontSize: 10.5, align: 'center', valign: 'middle' },
  }));
  const rows = [headerCells];
  D.lowVolumeStations.forEach((s, i) => {
    const rowFill = i % 2 === 0 ? 'FFFFFF' : 'F7F7F9';
    const g = s.grade;
    let momCell;
    if (!s.momChange) {
      momCell = { text: '－（首次記錄）', options: { fill: { color: rowFill }, color: C.textGray, fontFace: F_MED, fontSize: 9.5, align: 'center', valign: 'middle' } };
    } else {
      const pct = s.momChange.pct;
      if (pct > 0) momCell = { text: `▲${pct}%`, options: { fill: { color: rowFill }, color: C.darkBlue, bold: true, fontFace: F_BOLD, fontSize: 11, align: 'center', valign: 'middle' } };
      else if (pct < 0) momCell = { text: `▼${Math.abs(pct)}%`, options: { fill: { color: rowFill }, color: C.orange, bold: true, fontFace: F_BOLD, fontSize: 11, align: 'center', valign: 'middle' } };
      else momCell = { text: '持平', options: { fill: { color: rowFill }, color: C.textGray, fontFace: F_MED, fontSize: 10.5, align: 'center', valign: 'middle' } };
    }
    rows.push([
      { text: g || '－', options: { fill: { color: rowFill }, color: g ? (gradeColor[g] || C.textGray) : C.textGray, bold: true, fontFace: F_BLACK, fontSize: 13, align: 'center', valign: 'middle' } },
      { text: s.hiveRank ? `第 ${s.hiveRank} 名` : '－', options: { fill: { color: rowFill }, color: C.textGray, fontFace: F_MED, fontSize: 10.5, align: 'center', valign: 'middle' } },
      { text: s.city, options: { fill: { color: rowFill }, color: C.darkGray, fontFace: F_MED, fontSize: 11, align: 'center', valign: 'middle' } },
      { text: s.name, options: { fill: { color: rowFill }, color: C.darkGray, bold: true, fontFace: F_BOLD, fontSize: 11, align: 'left', valign: 'middle' } },
      { text: '', options: { fill: { color: rowFill } } }, // 機型：改由徽章 (badge) 覆蓋繪製
      { text: '', options: { fill: { color: rowFill } } }, // 租賃/買斷：改由徽章 (badge) 覆蓋繪製
      { text: `${s.contribution.toLocaleString()}`, options: { fill: { color: rowFill }, color: C.darkBlue, bold: true, fontFace: F_BOLD, fontSize: 12, align: 'center', valign: 'middle' } },
      momCell,
    ]);
  });

  slide.addTable(rows, {
    x: tX, y: tY, w: tW, colW: cw, rowH: [headerH, ...Array(D.lowVolumeStations.length).fill(dataRowH)],
    border: { type: 'solid', color: 'E5E5E5', pt: 0.5 },
    autoPage: false,
  });

  // ---- 機型／租賃買斷 徽章 (badge) 覆蓋繪製 ----
  const machineColors = { 'H30': C.orange, 'AI-4': C.blue, 'AI-HD': C.darkBlue, 'AI': '4F86A0' };
  const colX_machine = tX + cw[0] + cw[1] + cw[2] + cw[3];
  const colX_lease = colX_machine + cw[4];

  D.lowVolumeStations.forEach((s, i) => {
    const rowY = tY + headerH + i * dataRowH;
    const rowCenterY = rowY + dataRowH / 2;

    // 機型 badge：實心淺色底（帶透明度）＋同色細邊框＋全圓角膠囊
    if (s.machineType) {
      const mColor = machineColors[s.machineType] || C.textGray;
      const bw = 1.05, bh = 0.28;
      const bx = colX_machine + (cw[4] - bw) / 2, by = rowCenterY - bh / 2;
      slide.addShape('roundRect', {
        x: bx, y: by, w: bw, h: bh, rectRadius: bh / 2,
        fill: { color: mColor, transparency: 85 }, line: { color: mColor, width: 1 },
      });
      slide.addText(s.machineType, {
        x: bx, y: by, w: bw, h: bh,
        fontFace: F_BOLD, fontSize: 10, bold: true, color: mColor, align: 'center', valign: 'middle', margin: 0,
      });
    }

    // 租賃/買斷 badge：白底＋細邊框膠囊＋圖示（🔑租賃／💰買斷）
    if (s.leaseType) {
      const isLease = s.leaseType.includes('租賃');
      const bw = 1.3, bh = 0.28;
      const bx = colX_lease + (cw[5] - bw) / 2, by = rowCenterY - bh / 2;
      slide.addShape('roundRect', {
        x: bx, y: by, w: bw, h: bh, rectRadius: bh / 2,
        fill: { color: 'FFFFFF' }, line: { color: C.blue, width: 1 },
      });
      slide.addImage({ path: ICON(isLease ? 'key_blue' : 'coins_blue'), x: bx + 0.14, y: by + 0.06, w: 0.16, h: 0.16 });
      slide.addText(s.leaseType, {
        x: bx + 0.32, y: by, w: bw - 0.4, h: bh,
        fontFace: F_BOLD, fontSize: 10, bold: true, color: C.blue, align: 'left', valign: 'middle', margin: 0,
      });
    }
  });

  // ---- 資料來源說明（取自 資料來源.md，以徽章列呈現等級門檻，方便一目瞭然）----
  const noteY2 = tY + headerH + dataRowH * D.lowVolumeStations.length + 0.22;
  const noteH = 1.05;
  slide.addShape('roundRect', {
    x: 0.4, y: noteY2, w: tW, h: noteH, rectRadius: 0.06,
    fill: { color: 'F5F7FF' }, line: { color: C.lightBlue, width: 0.75 },
  });
  slide.addText(D.dataSourceNote, {
    x: 0.65, y: noteY2 + 0.08, w: tW - 0.5, h: 0.26,
    fontFace: F_MED, fontSize: 9.5, color: C.darkGray, valign: 'middle', margin: 0,
  });
  slide.addText(D.tierSectionLabel || '等級+加分/瓶量：', {
    x: 0.65, y: noteY2 + 0.35, w: 3, h: 0.22,
    fontFace: F_BOLD, fontSize: 9, bold: true, color: C.blue, valign: 'middle', margin: 0,
  });

  // 等級＋加分／瓶量 徽章列（S～F＋ARK 共8格，同一種呈現方式，平均分配寬度避免溢出）
  const tierGradeColor = { S: C.orange, A: C.darkBlue, B: C.blue, C: '4F86A0', D: C.textGray, E: C.textGray, F: C.textGray, ARK: '9C7A00' };
  const tiers = D.gradeTiers || [];
  const legendAreaW = tW - 0.5;
  const slotW = tiers.length ? legendAreaW / tiers.length : legendAreaW;
  const legendY = noteY2 + 0.60;
  const chipW = 0.55, chipH = 0.24;
  tiers.forEach((t, i) => {
    const tc = tierGradeColor[t.grade] || C.textGray;
    const slotX = 0.65 + i * slotW;
    slide.addShape('roundRect', {
      x: slotX, y: legendY, w: chipW, h: chipH, rectRadius: chipH / 2,
      fill: { color: tc, transparency: 85 }, line: { color: tc, width: 0.75 },
    });
    slide.addText(t.badge || t.grade, {
      x: slotX, y: legendY, w: chipW, h: chipH,
      fontFace: F_BOLD, fontSize: 8.5, bold: true, color: tc, align: 'center', valign: 'middle', margin: 0,
    });
    slide.addText(t.range, {
      x: slotX + chipW + 0.04, y: legendY, w: slotW - chipW - 0.06, h: chipH,
      fontFace: F_MED, fontSize: 7, color: C.textGray, valign: 'middle', margin: 0,
    });
  });
}

// ============================================================
// SLIDE 4 — 客訴趨勢洞察（MoM／QoQ）
// ============================================================
function buildSlide4() {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  const P4 = D.page4;
  if (!P4) return; // 防呆：舊版 data.json 若無 page4 欄位，略過此頁避免整份簡報中斷

  slide.addShape('rect', { x: 0.4, y: 0.2, w: 0.08, h: 0.34, fill: { color: C.blue }, line: { type: 'none' } });
  slide.addText('客訴趨勢洞察（MoM／QoQ）', {
    x: 0.58, y: 0.14, w: 7, h: 0.44,
    fontFace: F_BLACK, fontSize: 18, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
  });
  const pillW4 = 2.7, pillH4 = 0.4;
  slide.addShape('roundRect', {
    x: PW - 0.4 - pillW4, y: 0.16, w: pillW4, h: pillH4, rectRadius: pillH4 / 2,
    fill: { color: C.blue }, line: { type: 'none' },
  });
  slide.addText(`${D.week}（${D.range}）`, {
    x: PW - 0.4 - pillW4, y: 0.16, w: pillW4, h: pillH4,
    fontFace: F_BOLD, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0,
  });

  // ---- KPI 總覽卡片 (4欄) ----
  const cardsY = 0.72, cardsH = 1.25, cardGap = 0.2;
  const cardW = (PW - 0.8 - cardGap * 3) / 4;

  function trendArrow4(val, positiveIsGood = false) {
    // 客訴量／機台佔比：數值下降代表改善，這裡用語意色（綠＝好／橙＝需留意），不套用其餘頁面「▲橙／▼深藍」的中性配色
    if (val === null || val === undefined) return { text: '－', color: C.textGray };
    if (val > 0) return { text: `▲${val}%`, color: positiveIsGood ? '1D9E75' : C.orange };
    if (val < 0) return { text: `▼${Math.abs(val)}%`, color: positiveIsGood ? C.orange : '1D9E75' };
    return { text: '持平', color: C.textGray };
  }

  function kpiCard(x, accent, icon, title, mainNode, subText, badge) {
    slide.addShape('roundRect', {
      x, y: cardsY, w: cardW, h: cardsH, rectRadius: 0.09,
      fill: { color: 'FFFFFF' }, line: { color: 'E5E5E5', width: 1 },
      shadow: { type: 'outer', color: '999999', opacity: 0.2, blur: 5, offset: 1.5, angle: 90 },
    });
    slide.addShape('roundRect', {
      x, y: cardsY, w: cardW, h: 0.07, rectRadius: 0.03,
      fill: { color: accent }, line: { type: 'none' },
    });
    slide.addText(title, {
      x: x + 0.2, y: cardsY + 0.12, w: cardW - 0.7, h: 0.28,
      fontFace: F_BOLD, fontSize: 11, bold: true, color: C.darkGray, valign: 'middle', margin: 0,
    });
    slide.addImage({ path: ICON(icon), x: x + cardW - 0.44, y: cardsY + 0.13, w: 0.24, h: 0.24 });
    slide.addText(mainNode, {
      x: x + 0.2, y: cardsY + 0.40, w: cardW - 0.4, h: 0.42, valign: 'middle', margin: 0,
    });
    slide.addText(subText, {
      x: x + 0.2, y: cardsY + 0.80, w: cardW - 1.0, h: 0.3,
      fontFace: F_MED, fontSize: 8.5, color: C.textGray, valign: 'top', margin: 0, lineSpacingMultiple: 1.05,
    });
    if (badge) {
      slide.addShape('roundRect', {
        x: x + cardW - 0.98, y: cardsY + 0.78, w: 0.8, h: 0.28, rectRadius: 0.14,
        fill: { color: badge.color === '1D9E75' ? 'DFF3EA' : (badge.color === C.orange ? 'FFE9DC' : 'F0F0F0') }, line: { type: 'none' },
      });
      slide.addText(badge.text, {
        x: x + cardW - 0.98, y: cardsY + 0.78, w: 0.8, h: 0.28,
        fontFace: F_BOLD, fontSize: 9.5, bold: true, color: badge.color, align: 'center', valign: 'middle', margin: 0,
      });
    }
  }

  // 卡1：本月客訴總量 (MoM)
  {
    const x = 0.4;
    const k = P4.kpi.monthTotal;
    const sub = k.prevValue != null ? `較上月（${k.prevValue}件）` : '無上月比較基準';
    kpiCard(x, C.blue, 'headset_blue', `本月客訴總量（${k.period}）`,
      [{ text: `${k.value} `, options: { fontSize: 25, bold: true, color: C.blue, fontFace: F_BLACK } },
       { text: '件', options: { fontSize: 12.5, bold: true, color: C.blue, fontFace: F_BOLD } }],
      sub, trendArrow4(k.deltaPct));
  }
  // 卡2：客訴密度指標（每站平均客訴量，暫代萬人客訴率；待有用戶數資料源後可直接換算，介面不變）
  {
    const x = 0.4 + (cardW + cardGap);
    const k = P4.kpi.perStation;
    kpiCard(x, C.darkBlue, 'pin_darkblue', '客訴密度指標（件/站）',
      [{ text: `${k.value ?? '－'} `, options: { fontSize: 25, bold: true, color: C.darkBlue, fontFace: F_BLACK } },
       { text: '件/站', options: { fontSize: 11, bold: true, color: C.darkBlue, fontFace: F_BOLD } }],
      '每站平均客訴量，暫代萬人客訴率', trendArrow4(k.deltaPct));
  }
  // 卡3：本季客訴總量 (QoQ)
  {
    const x = 0.4 + 2 * (cardW + cardGap);
    const k = P4.kpi.quarterTotal;
    const sub = k.prevValue != null ? `較上季（${k.prevValue}件）` : '無上季比較基準';
    kpiCard(x, C.orange, 'warn_orange', `本季客訴總量（${k.quarterLabel}）`,
      [{ text: `${k.value ?? '－'} `, options: { fontSize: 25, bold: true, color: C.orange, fontFace: F_BLACK } },
       { text: '件', options: { fontSize: 12.5, bold: true, color: C.orange, fontFace: F_BOLD } }],
      sub, trendArrow4(k.deltaPct));
  }
  // 卡4：機台問題佔比 (QoQ)
  {
    const x = 0.4 + 3 * (cardW + cardGap);
    const k = P4.kpi.machinePct;
    kpiCard(x, C.yellow, 'recycle_blue', `機台問題佔比（${P4.quarterCompare.curLabel}）`,
      [{ text: `${k.value}`, options: { fontSize: 25, bold: true, color: '9C7A00', fontFace: F_BLACK } },
       { text: '%', options: { fontSize: 12.5, bold: true, color: '9C7A00', fontFace: F_BOLD } }],
      '較上季佔比變化', trendArrow4(k.deltaPct, true));
  }

  // ---- 圖表區：圖A（MoM 雙軸趨勢）／ 圖B（QoQ 結構比，加寬）----
  const chartY = cardsY + cardsH + 0.3;
  const chartH = 2.5;
  const leftW4 = 5.7, colGap4 = 0.3, rightX4 = 0.4 + leftW4 + colGap4, rightW4 = PW - 0.4 - rightX4;

  sectionBar(slide, `總客訴量與客訴密度趨勢圖（MoM，近${P4.monthlyTrend.length}個月）`, 0.4, chartY, leftW4);
  const barData4 = [{ name: '客訴總量（件）', labels: P4.monthlyTrend.map(m => m.label), values: P4.monthlyTrend.map(m => m.total) }];
  const lineData4 = [{ name: '客訴密度（件/站）', labels: P4.monthlyTrend.map(m => m.label), values: P4.monthlyTrend.map(m => m.perStation) }];
  slide.addChart(
    [
      { type: pres.ChartType.bar, data: barData4, options: { chartColors: [C.blue] } },
      { type: pres.ChartType.line, data: lineData4, options: { chartColors: [C.orange], secondaryValAxis: true, secondaryCatAxis: true, lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 6 } },
    ],
    {
      x: 0.4, y: chartY + 0.38, w: leftW4, h: chartH - 0.38,
      showLegend: true, legendPos: 'b', legendFontSize: 9, legendColor: C.textGray,
      // 雙軸圖表：pptxgenjs 需在頂層明確給出 valAxes / catAxes 兩組設定，
      // 光靠各 series 的 secondaryValAxis:true 並不會真的產生第二條座標軸(僅是參照，若沒有對應定義會導致折線塌縮貼底)
      valAxes: [
        { valAxisLabelFontSize: 8.5, valAxisLabelColor: C.textGray, showValAxisTitle: false },
        { valAxisLabelFontSize: 8.5, valAxisLabelColor: C.textGray, showValAxisTitle: false, valAxisMinVal: 0 },
      ],
      catAxes: [
        { catAxisLabelFontSize: 9.5, catAxisLabelColor: C.textGray },
        { catAxisHidden: true },
      ],
      barGapWidthPct: 40,
    }
  );

  sectionBar(slide, '客訴主題結構比變化（QoQ）', rightX4, chartY, rightW4, { textColor: C.orange });
  if (P4.quarterCompare.prev.length) {
    const qLabels = [P4.quarterCompare.prevLabel, P4.quarterCompare.curLabel];
    const stackColors = { '機台問題': C.blue, '回收點數問題': C.darkBlue, '顧客關係': '7F77DD', '優惠券問題': 'E24B4A', 'APP帳號設定問題': 'BA7517', 'APP使用問題': 'D4537E' };
    const stackData4 = P4.quarterCompare.cur.map((c, i) => ({
      name: c.label,
      labels: qLabels,
      values: [P4.quarterCompare.prev[i].pct, c.pct],
    }));

    slide.addChart(pres.ChartType.bar, stackData4, {
      x: rightX4, y: chartY + 0.38, w: rightW4, h: chartH - 0.38,
      barDir: 'bar', barGrouping: 'percentStacked',
      chartColors: P4.quarterCompare.cur.map(c => stackColors[c.label] || C.textGray),
      showLegend: true, legendPos: 'b', legendFontSize: 8, legendColor: C.textGray,
      catAxisLabelFontSize: 9.5, catAxisLabelColor: C.textGray,
      valAxisLabelFontSize: 8.5, valAxisLabelColor: C.textGray,
      valAxisLabelFormatCode: '0%',
      showValAxisTitle: true, valAxisTitle: '客訴佔比', valAxisTitleFontSize: 9, valAxisTitleColor: C.textGray,
      showValue: true, dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF',
      dataLabelFontSize: 6.5, dataLabelFontFace: F_BOLD, dataLabelFormatCode: '0.0"%"',
    });
  } else {
    slide.addText('資料累積中，尚無上一季度可供比較', {
      x: rightX4, y: chartY + 0.9, w: rightW4, h: 0.5,
      fontFace: F_MED, fontSize: 11, color: C.textGray, align: 'center', margin: 0,
    });
  }

  // ---- 客服觀點與改善行動建議 ----
  const insightY = chartY + chartH + 0.28;
  sectionBar(slide, '客服觀點與改善行動建議', 0.4, insightY, PW - 0.8);

  function insightBox(y, h, barColor, bgColor, label, text, textColor, fontSize) {
    slide.addShape('roundRect', {
      x: 0.4, y, w: PW - 0.8, h, rectRadius: 0.06,
      fill: { color: bgColor }, line: { color: barColor, width: 0.75 },
    });
    slide.addShape('rect', { x: 0.4, y, w: 0.06, h, fill: { color: barColor }, line: { type: 'none' } });
    slide.addText(label, {
      x: 0.62, y: y + 0.06, w: 1.9, h: h - 0.12,
      fontFace: F_BOLD, fontSize: 10.5, bold: true, color: textColor || barColor, valign: 'middle', margin: 0,
    });
    slide.addText(text, {
      x: 2.4, y: y + 0.06, w: PW - 0.8 - 2.0, h: h - 0.12,
      fontFace: F_MED, fontSize: fontSize || 10, color: textColor || C.darkGray, valign: 'middle', margin: 0, lineSpacingMultiple: 1.12, wrap: true,
    });
  }

  let iy = insightY + 0.36;
  const boxH = 0.58, boxGap = 0.13;

  // 重點結論放在最上面：一句話講完本月＋本季方向，主管/跨部門優先看這一格
  insightBox(iy, 0.62, C.blue, C.blue, '重點結論', P4.headline, C.white, 9.5);
  iy += 0.62 + boxGap;

  insightBox(iy, boxH, C.orange, 'FFF4EF', '本季結構變化重點', P4.structuralInsight);
  iy += boxH + boxGap;

  if (P4.activityInsight) {
    insightBox(iy, boxH, C.darkBlue, 'EEF6FB', '系統／活動影響評估', P4.activityInsight);
    iy += boxH + boxGap;
  }
}

buildSlide1();
buildSlide2();
buildSlide3();
buildSlide4();

const outFile = path.join(__dirname, `ecoco_客服週報_${D.week}_${D.range.replace(/[\s\/]/g, '')}.pptx`);
pres.writeFile({ fileName: outFile }).then(() => {
  console.log('PPT 產出完成：', outFile);
});
