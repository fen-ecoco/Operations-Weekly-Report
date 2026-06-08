const pptxgen = require("pptxgenjs");

const VI = {
  orange:"FF5000", blue:"060E9F", yellow:"FFCE00",
  ltBlue:"8EB8C9", beige:"FAE0B8", dkBlue:"0076A9",
  white:"FFFFFF", darkGray:"333333", textGray:"888888",
  border:"E0E6F0", rowAlt:"F5F7FC",
  trendUp:"FF5000", trendDn:"0076A9",
};
const CAT_C  = ["060E9F","FF5000","0076A9","FFCE00","8EB8C9","FAE0B8"];
const MACH_C = ["FF5000","060E9F","0076A9","FFCE00","8EB8C9","B8BEC8"];
const FC = "Noto Sans TC";

const D = {
  week:"第23週", range:"06/01 ~ 06/07", total:168,
  cats:[
    {label:"APP帳號設定問題",count: 44,pct:26.2},
    {label:"機台問題",        count: 78,pct:46.4},
    {label:"回收點數問題",     count: 19,pct:11.3},
    {label:"顧客關係",         count: 11,pct: 6.5},
    {label:"優惠券問題",       count: 13,pct: 7.7},
    {label:"APP使用問題",      count:  3,pct: 1.8},
  ],
  trend:[
    {w:"20週",d:"05/11-05/17",bt:[68,0], b2:[0,0],  rg:[8,0],    pt:[12,0],  cp:[6,0],  ap:[4,0], cr:[15,0], tot:[113,0]},
    {w:"21週",d:"05/18-05/24",bt:[65,-3],b2:[12,12],rg:[232,224],pt:[15,3],  cp:[9,3],  ap:[11,7],cr:[13,-2],tot:[357,244]},
    {w:"22週",d:"05/25-05/31",bt:[95,30],b2:[5,-7], rg:[123,-109],pt:[22,7], cp:[6,-3], ap:[8,-3],cr:[16,3], tot:[275,-82]},
    {w:"23週",d:"06/01-06/07",bt:[95,0], b2:[6,1],  rg:[44,-79], pt:[19,-3], cp:[13,7], ap:[3,-5],cr:[11,-5],tot:[168,-107]},
  ],
  nonMachine:[
    {rank:1,name:"無法接收簡訊驗證碼",count:36,pct:13.1,col:"FF5000",
     note:"連三週持續高位，建議追蹤SMS服務商改善進度"},
    {rank:2,name:"點數未入帳號",       count:19,pct: 6.9,col:"0076A9",
     note:"收瓶機18件、方舟站1件"},
    {rank:3,name:"忘記密碼/無法重設",  count: 5,pct: 1.8,col:"060E9F",
     note:"與簡訊驗證碼問題高度相關"},
  ],
  machIssues:[
    {label:"瓶蓋桶已滿",           count:29,pct:37.2},
    {label:"機台需維護/故障提醒",   count: 9,pct:11.5},
    {label:"履帶未作動或異常抖動",  count: 8,pct:10.3},
    {label:"操作流程異常/無法正常操作",count:5,pct:6.4},
    {label:"機台操作畫面無法登入",  count: 5,pct: 6.4},
    {label:"其他機台問題",          count:22,pct:28.2},
  ],
  bottleTop3:[
    {name:"瓶蓋桶已滿",          count:29,pct:31,col:"FF5000"},
    {name:"點數未入帳號",         count:19,pct:20,col:"060E9F"},
    {name:"履帶未作動或異常抖動", count: 8,pct: 8,col:"0076A9"},
  ],
  batTop:[
    {name:"操作流程異常/無法正常操作",count:2,pct:33,col:"0076A9"},
    {name:"機台操作畫面無法登入",     count:1,pct:17,col:"8EB8C9"},
  ],
  bottleTotal:95,batTotal:6,bottlePct:94,batPct:6,
  hotAreas:[
    {area:"高雄",total:26,spots:[
      {name:"全聯苓雅三多店站",      count:8},
      {name:"小北百貨高雄自強店站",  count:4},
      {name:"小北百貨高雄大昌店站",  count:2},
    ]},
    {area:"新北",total:18,spots:[
      {name:"特力屋新莊店站",        count:7},
      {name:"全聯蘆洲重陽店站",      count:4},
      {name:"全聯中和新生店站",      count:1},
    ]},
    {area:"臺南",total:11,spots:[
      {name:"全家西港南海店站",      count:1},
      {name:"全家仁德新德店站",      count:1},
      {name:"台南總圖方舟站",        count:1},
    ]},
  ],
  alertText:"【注意】瓶蓋桶已滿 連兩週居首｜本週29件（30.5%），較上週明顯上升。無法接收簡訊驗證碼持續高位36件，已連三週，建議追蹤SMS服務商改善進度。",
};

// ── 工具 ──
function dCell(arr, isTotal, isLast) {
  const [v, d] = arr;
  const bg  = isLast ? "FAE0B8" : "FFFFFF";
  const vc  = isTotal ? "FF5000" : (isLast ? "333333" : "333333");
  const parts = [{text:String(v), options:{bold:isLast||isTotal, color:vc, fontSize:isTotal?11:10}}];
  if (d!==0) parts.push({text:` ${d>0?"▲":"▼"}${Math.abs(d)}`,
    options:{fontSize:6.5, color:d>0?"FF5000":"0076A9", bold:false}});
  return {text:parts, options:{align:"center", fill:bg, fontFace:FC, valign:"middle"}};
}

function secBar(sl, x, y, label, textColor="060E9F") {
  sl.addShape("rect",{x, y, w:0.05, h:0.22, fill:{color:"FF5000"}});
  sl.addText(label, {x:x+0.09, y:y+0.01, w:8.8, h:0.22,
    fontSize:10.5, bold:true, color:textColor, fontFace:FC, valign:"middle"});
}

// ═══════════════════════════════════════════════════
//  PAGE 1
// ═══════════════════════════════════════════════════
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

let s1 = pres.addSlide();
s1.background = {color:VI.white};

// Title（pixel-measured: y≈0.06, fontSize≈16）
s1.addText(`客服課　${D.range}　客訴問題分析`, {
  x:0.18, y:0.06, w:9.6, h:0.26,
  fontSize:16, bold:true, color:VI.darkGray, fontFace:FC});

// Section bar 1
secBar(s1, 0.18, 0.36);

// Trend table（measured: y:0.578, h:1.995, rowH:0.399≈0.38）
const TH = [
  {text:"週次/期間",  options:{bold:true,fontSize:8.5,fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"收瓶機/方舟",options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"二代電池機", options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"註冊帳號",   options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"回收點數",   options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"優惠券",     options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"APP使用",    options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"顧客關係",   options:{bold:true,fontSize:8,  fill:"060E9F",color:"FFFFFF",align:"center",fontFace:FC}},
  {text:"加總",       options:{bold:true,fontSize:8,  fill:"FF5000",color:"FFFFFF",align:"center",fontFace:FC}},
];

const TB = D.trend.map((r,ri)=>{
  const last = ri===D.trend.length-1;
  const bg   = last ? "FAE0B8" : (ri%2===0 ? "FFFFFF" : "F5F7FC");
  const wk   = {text:`${r.w}\n${r.d}`, options:{fontSize:8,align:"center",fill:bg,color:last?"333333":"555555",bold:last,fontFace:FC,valign:"middle"}};
  // reg cell special
  const [rv,rd] = r.rg;
  const rgParts = [{text:String(rv), options:{bold:last||rv>100,color:rv>100?"FF5000":"333333",fontSize:rv>100?11:10}}];
  if (rd!==0) rgParts.push({text:` ${rd>0?"▲":"▼"}${Math.abs(rd)}`, options:{fontSize:6.5,color:rd>0?"FF5000":"0076A9"}});
  const rgCell = {text:rgParts, options:{align:"center",fill:last?"FAE0B8":"FFFFFF",fontFace:FC,valign:"middle"}};
  return [wk, dCell(r.bt,false,last), dCell(r.b2,false,last),
    rgCell, dCell(r.pt,false,last), dCell(r.cp,false,last),
    dCell(r.ap,false,last), dCell(r.cr,false,last), dCell(r.tot,true,last)];
});

s1.addTable([TH,...TB],{
  x:0.18, y:0.60, w:9.64, h:1.96,
  border:{pt:0.5,color:"D1DCF0"},
  rowH:0.37,
  colW:[1.05,1.15,1.1,1.1,1.1,0.88,0.88,0.88,0.84],
});

// Section bar 2（orange text）
secBar(s1, 0.18, 2.63, `當週客訴佔比 + 主分類佔比（${D.range}）`, "FF5000");

// ── 圓餅圖（measured: x:0.071, w:3.86, h:2.62）──
s1.addChart(pres.ChartType.pie,
  [{name:"佔比", labels:D.cats.map(c=>c.label), values:D.cats.map(c=>c.count)}],
  {
    x:0.07, y:2.89, w:3.86, h:2.56,
    showLegend:false, showTitle:false,
    showPercent:true, showValue:false,
    dataLabelFormatCode:"0%",
    dataLabelFontSize:9, dataLabelFontBold:true,
    dataLabelColor:"FFFFFF",
    chartColors:CAT_C,
    shadow:{type:"none"},
  });

// ── 圖例（measured x:4.00, items 6×）──
const LX=4.05, LY0=2.89, LH=0.425;
D.cats.forEach((c,i)=>{
  const y = LY0 + i*LH;
  s1.addShape("rect",{x:LX, y:y+0.07, w:0.20, h:0.20, fill:{color:CAT_C[i]}});
  s1.addText(c.label, {x:LX+0.26, y:y+0.03, w:2.15, h:0.22,
    fontSize:10, bold:true, color:"333333", fontFace:FC});
  s1.addText(`${c.count}件　${c.pct}%`, {x:LX+0.26, y:y+0.24, w:2.15, h:0.18,
    fontSize:8, color:"333333", fontFace:FC});
});

// ── Top3 Cards（measured x:6.326, w:3.574, per-card h:0.893）──
const CX=6.33, CY0=2.89, CH=0.875, CW=3.57;
D.nonMachine.forEach((item,i)=>{
  const ry  = CY0 + i*(CH+0.025);
  const bgs = ["FFF4EF","EEF6FB","F5F7FF"];
  // 卡片
  s1.addShape("rect",{x:CX, y:ry, w:CW, h:CH,
    fill:{color:bgs[i]}, line:{color:"E0E6F0",pt:0.5}});
  // 左色條
  s1.addShape("rect",{x:CX, y:ry, w:0.05, h:CH, fill:{color:item.col}});
  // 排名
  s1.addText(String(item.rank), {x:CX+0.10, y:ry+0.06, w:0.20, h:0.22,
    fontSize:9, bold:true, color:item.col, fontFace:FC});
  // 問題名稱（row1）
  s1.addText(item.name, {x:CX+0.33, y:ry+0.06, w:2.0, h:0.24,
    fontSize:10, bold:true, color:"333333", fontFace:FC});
  // 大字件數・%（row2, measured fontSize≈20）
  s1.addText(`${item.count}件・${item.pct}%`, {x:CX+0.33, y:ry+0.30, w:2.42, h:0.36,
    fontSize:18, bold:true, color:item.col, fontFace:FC});
  // 說明文字（右側）
  s1.addText(item.note, {x:CX+2.80, y:ry+0.08, w:0.73, h:0.62,
    fontSize:7, color:"888888", fontFace:FC, wrap:true, valign:"top"});
  // 底部 mini bar
  s1.addShape("rect",{x:CX+0.05, y:ry+CH-0.06, w:CW-0.05, h:0.05, fill:{color:"EBEBEB"}});
  s1.addShape("rect",{x:CX+0.05, y:ry+CH-0.06, w:(CW-0.05)*(item.pct/100), h:0.05, fill:{color:item.col}});
});

// ═══════════════════════════════════════════════════
//  PAGE 2
// ═══════════════════════════════════════════════════
let s2 = pres.addSlide();
s2.background = {color:VI.white};

// Title
s2.addText(`客服課　${D.range}　機台問題佔比 與 熱門站點`, {
  x:0.18, y:0.06, w:9.6, h:0.26,
  fontSize:16, bold:true, color:VI.darkGray, fontFace:FC});

// Section bar "客訴詳情與分類佔比"
secBar(s2, 0.18, 0.36);

// 機台圓餅（measured: x:0.073, y:0.521, w:2.988, h:1.723）
s2.addChart(pres.ChartType.pie,
  [{name:"機台", labels:D.machIssues.map(m=>m.label), values:D.machIssues.map(m=>m.count)}],
  {
    x:0.07, y:0.52, w:3.05, h:1.70,
    showLegend:false, showTitle:false,
    showPercent:true, showValue:false,
    dataLabelFormatCode:"0%",
    dataLabelFontSize:8, dataLabelFontBold:true,
    dataLabelColor:"FFFFFF",
    chartColors:MACH_C,
    shadow:{type:"none"},
  });

// 機台圓餅圖例
const ML=3.18, MLY0=0.54, MLH=0.245;
D.machIssues.forEach((m,i)=>{
  const y = MLY0 + i*MLH;
  s2.addShape("rect",{x:ML, y:y+0.04, w:0.15, h:0.15, fill:{color:MACH_C[i]}});
  s2.addText(`${m.label}　${m.count}件（${m.pct}%）`, {x:ML+0.20, y, w:3.05, h:0.22,
    fontSize:8.5, color:"333333", fontFace:FC});
});

// 大數字卡片（measured: card1 x:6.122, w:1.094, h:1.723, gap:0.109）
const bigCards = [
  {x:6.12, label:"收瓶機/方舟", pct:D.bottlePct, count:D.bottleTotal, bg:"FFF3EC", border:"FF5000", txt:"FF5000"},
  {x:7.32, label:"電池機",      pct:D.batPct,    count:D.batTotal,    bg:"EDF5FB", border:"8EB8C9", txt:"0076A9"},
];
bigCards.forEach(c=>{
  s2.addShape("rect",{x:c.x, y:0.52, w:1.10, h:1.70,
    fill:{color:c.bg}, line:{color:c.border, pt:1.2}});
  s2.addText(c.label, {x:c.x, y:0.58, w:1.10, h:0.20,
    fontSize:8, color:"888888", align:"center", fontFace:FC});
  s2.addText(`${c.pct}%`, {x:c.x, y:0.80, w:1.10, h:0.56,
    fontSize:32, bold:true, color:c.txt, align:"center", fontFace:"Arial"});
  s2.addText(`（${c.count}件）`, {x:c.x, y:1.76, w:1.10, h:0.22,
    fontSize:8, color:"888888", align:"center", fontFace:FC});
});

// Section bars（bottom half: y:2.317）
secBar(s2, 0.18, 2.32, "收瓶機、電池機客訴 Top 3");
secBar(s2, 6.20, 2.32, `本週熱點站點（${D.week}：${D.range}）`, "FF5000");

// 收瓶機（measured: y bottom content:2.498, end:3.692 → h:1.194, 3 items each 0.398）
const BSTART = 2.52;
s2.addShape("rect",{x:0.18, y:BSTART, w:0.72, h:1.18, fill:{color:"060E9F"}});
s2.addText("收\n瓶\n機",{x:0.18, y:BSTART, w:0.72, h:1.18,
  fontSize:10.5, bold:true, color:"FFFFFF", align:"center", valign:"middle", fontFace:FC});

D.bottleTop3.forEach((item,i)=>{
  const ry = BSTART+0.02 + i*0.385;
  const bg = i===0?"EDF3FF": i===1?"F4F8FF":"F9FBFF";
  s2.addShape("rect",{x:0.92, y:ry, w:5.18, h:0.365,
    fill:{color:bg}, line:{color:"E0E6F0",pt:0.4}});
  s2.addText(`${i+1}`,{x:0.98, y:ry+0.07, w:0.20, h:0.22,
    fontSize:10, bold:true, color:item.col, fontFace:FC});
  s2.addText(item.name,{x:1.24, y:ry+0.07, w:3.3, h:0.22,
    fontSize:9.5, color:"333333", fontFace:FC});
  s2.addText(`${item.pct}%`,{x:5.26, y:ry+0.04, w:0.78, h:0.28,
    fontSize:20, bold:true, color:item.col, align:"right", valign:"middle", fontFace:"Arial"});
  s2.addShape("rect",{x:1.24, y:ry+0.315, w:3.55,              h:0.038, fill:{color:"EBEBEB"}});
  s2.addShape("rect",{x:1.24, y:ry+0.315, w:3.55*(item.pct/100),h:0.038, fill:{color:item.col}});
});

// 電池機（measured: end:4.488, sidebar h:0.796）
const BATSTART = BSTART+1.195;
s2.addShape("rect",{x:0.18, y:BATSTART, w:0.72, h:0.80, fill:{color:"FF5000"}});
s2.addText("電\n池\n機",{x:0.18, y:BATSTART, w:0.72, h:0.80,
  fontSize:10.5, bold:true, color:"FFFFFF", align:"center", valign:"middle", fontFace:FC});

D.batTop.forEach((item,i)=>{
  const ry = BATSTART+0.02 + i*0.378;
  const bg = i===0?"F0FDF7":"F7FAFE";
  s2.addShape("rect",{x:0.92, y:ry, w:5.18, h:0.358,
    fill:{color:bg}, line:{color:"E0E6F0",pt:0.4}});
  s2.addText(`${i+1}`,{x:0.98, y:ry+0.07, w:0.20, h:0.22,
    fontSize:10, bold:true, color:item.col, fontFace:FC});
  s2.addText(item.name,{x:1.24, y:ry+0.07, w:3.3, h:0.22,
    fontSize:9.5, color:"333333", fontFace:FC});
  s2.addText(`${item.pct}%`,{x:5.26, y:ry+0.04, w:0.78, h:0.28,
    fontSize:20, bold:true, color:item.col, align:"right", valign:"middle", fontFace:"Arial"});
  s2.addShape("rect",{x:1.24, y:ry+0.305, w:3.55,              h:0.038, fill:{color:"EBEBEB"}});
  s2.addShape("rect",{x:1.24, y:ry+0.305, w:3.55*(item.pct/100),h:0.038, fill:{color:item.col}});
});

// 警示框（measured: y:4.525, h:0.615）
s2.addShape("rect",{x:0.18, y:4.54, w:5.88, h:0.60,
  fill:{color:"FAE0B8"}, line:{color:"FFCE00",pt:1}});
s2.addText(D.alertText,{x:0.28, y:4.56, w:5.72, h:0.56,
  fontSize:7.5, color:"6B4800", fontFace:FC, wrap:true, valign:"top"});

// 熱點站點（measured: x:6.195, y:2.498, total h:2.642）
// 3 cities × (0.22city + 3×0.20spot + 0.02gap) = 3×0.88 = 2.64 ✓
let aY = 2.52;
D.hotAreas.forEach(area=>{
  // 城市列
  s2.addShape("rect",{x:6.20, y:aY, w:3.65, h:0.26, fill:{color:"FAE0B8"}});
  s2.addText(area.area,{x:6.26, y:aY, w:1.8, h:0.26,
    fontSize:12, bold:true, color:"060E9F", valign:"middle", fontFace:FC});
  s2.addText(String(area.total),{x:9.2, y:aY, w:0.60, h:0.26,
    fontSize:13, bold:true, color:"FF5000", align:"right", valign:"middle", fontFace:"Arial"});
  aY += 0.27;
  // 站點列
  area.spots.forEach(spot=>{
    s2.addText(`・${spot.name}`,{x:6.30, y:aY, w:2.80, h:0.22,
      fontSize:8.5, color:"333333", fontFace:FC, valign:"middle"});
    s2.addText(String(spot.count),{x:9.2, y:aY, w:0.60, h:0.22,
      fontSize:9, bold:true, color:"888888", align:"right", fontFace:"Arial"});
    aY += 0.225;
  });
  aY += 0.02;
});

const OUT="/mnt/user-data/outputs/ecoco_週報_第23週_0601-0607.pptx";
pres.writeFile({fileName:OUT}).then(()=>console.log("✅ v5："+OUT));
