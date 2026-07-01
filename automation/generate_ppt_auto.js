/**
 * ecoco 每週例會PPT 自動產出腳本 — Step 3
 * 讀取 data.json → 產出 PPTX（v5 pixel-accurate layout）
 */
const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const path    = require("path");

// ── 路徑設定 ──
const BASE     = __dirname;
const cfg      = JSON.parse(fs.readFileSync(path.join(BASE,"config.json"),"utf8"));
const D        = JSON.parse(fs.readFileSync(path.join(BASE,"data.json"),"utf8"));
const OUT_DIR  = cfg.output_dir;
const REPO_DIR = path.join(BASE,"..","weekly-ppt");
const FNAME    = `ecoco_週報_${D.week}_${D.range.replace(/ /g,"").replace(/\//g,"").replace(/~/g,"-")}.pptx`;
const OUT_LOCAL = path.join(OUT_DIR, FNAME);
const OUT_REPO  = path.join(REPO_DIR, FNAME);

// ── VI 色系 ──
const VI = {
  orange:"FF5000", blue:"060E9F", yellow:"FFCE00",
  ltBlue:"8EB8C9", beige:"FAE0B8", dkBlue:"0076A9",
  white:"FFFFFF",  darkGray:"333333", textGray:"888888",
  border:"E0E6F0", rowAlt:"F5F7FC",
  trendUp:"FF5000", trendDn:"0076A9",
};
const CAT_C  = ["060E9F","FF5000","0076A9","FFCE00","8EB8C9","FAE0B8"];
const MACH_C = ["FF5000","060E9F","0076A9","FFCE00","8EB8C9","B8BEC8"];
const FC = "Noto Sans TC";

// ── 工具 ──
function dCell([v, d], isTotal, isLast) {
  const bg = isLast ? "FAE0B8" : "FFFFFF";
  const vc = isTotal ? "FF5000" : "333333";
  const parts = [{text:String(v), options:{bold:isLast||isTotal, color:vc, fontSize:isTotal?11:10}}];
  if (d!==0) parts.push({text:` ${d>0?"▲":"▼"}${Math.abs(d)}`,
    options:{fontSize:6.5, color:d>0?"FF5000":"0076A9", bold:false}});
  return {text:parts, options:{align:"center", fill:bg, fontFace:FC, valign:"middle"}};
}
function secBar(sl, x, y, label, textColor="060E9F") {
  sl.addShape("rect",{x, y, w:0.05, h:0.22, fill:{color:"FF5000"}});
  sl.addText(label,{x:x+0.09, y:y+0.01, w:8.8, h:0.22,
    fontSize:10.5, bold:true, color:textColor, fontFace:FC, valign:"middle"});
}

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// ══════════════════════════════════
//  PAGE 1
// ══════════════════════════════════
let s1 = pres.addSlide();
s1.background = {color:VI.white};
s1.addText(`客服課　${D.range}　客訴問題分析`,{
  x:0.18,y:0.06,w:9.6,h:0.26, fontSize:16,bold:true,color:VI.darkGray,fontFace:FC});
secBar(s1,0.18,0.36,"客訴趨勢分析");

// Trend table
const TH = ["週次/期間","收瓶機/方舟","二代電池機","註冊帳號","回收點數","優惠券","APP使用","顧客關係","加總"]
  .map((t,i)=>({text:t, options:{bold:true,fontSize:i===0?8.5:8,
    fill:i===8?"FF5000":"060E9F",color:"FFFFFF",align:"center",fontFace:FC}}));

const TB = D.trend.map((r,ri)=>{
  const last=ri===D.trend.length-1;
  const bg=last?"FAE0B8":(ri%2===0?"FFFFFF":"F5F7FC");
  const wk={text:`${r.w}\n${r.d}`,options:{fontSize:8,align:"center",fill:bg,color:last?"333333":"555555",bold:last,fontFace:FC,valign:"middle"}};
  const [rv,rd]=r.rg;
  const rgParts=[{text:String(rv),options:{bold:last||rv>100,color:rv>100?"FF5000":"333333",fontSize:rv>100?11:10}}];
  if(rd!==0) rgParts.push({text:` ${rd>0?"▲":"▼"}${Math.abs(rd)}`,options:{fontSize:6.5,color:rd>0?"FF5000":"0076A9"}});
  const rgCell={text:rgParts,options:{align:"center",fill:last?"FAE0B8":"FFFFFF",fontFace:FC,valign:"middle"}};
  return [wk,dCell(r.bt,false,last),dCell(r.b2,false,last),rgCell,
    dCell(r.pt,false,last),dCell(r.cp,false,last),dCell(r.ap,false,last),dCell(r.cr,false,last),dCell(r.tot,true,last)];
});
s1.addTable([TH,...TB],{x:0.18,y:0.60,w:9.64,h:1.96,
  border:{pt:0.5,color:"D1DCF0"},rowH:0.37,
  colW:[1.05,1.15,1.1,1.1,1.1,0.88,0.88,0.88,0.84]});

secBar(s1,0.18,2.63,`當週客訴佔比 + 主分類佔比（${D.range}）`,"FF5000");
s1.addChart(pres.ChartType.pie,
  [{name:"佔比",labels:D.cats.map(c=>c.label),values:D.cats.map(c=>c.count)}],
  {x:0.07,y:2.89,w:3.86,h:2.56,showLegend:false,showTitle:false,
   showPercent:true,dataLabelFormatCode:"0%",dataLabelFontSize:9,
   dataLabelFontBold:true,dataLabelColor:"FFFFFF",chartColors:CAT_C,shadow:{type:"none"}});

const LX=4.05,LY0=2.89,LH=0.425;
D.cats.forEach((c,i)=>{
  const y=LY0+i*LH;
  s1.addShape("rect",{x:LX,y:y+0.07,w:0.20,h:0.20,fill:{color:CAT_C[i]}});
  s1.addText(c.label,{x:LX+0.26,y:y+0.03,w:2.15,h:0.22,fontSize:10,bold:true,color:"333333",fontFace:FC});
  s1.addText(`${c.count}件　${c.pct}%`,{x:LX+0.26,y:y+0.24,w:2.15,h:0.18,fontSize:8,color:"333333",fontFace:FC});
});

const CX=6.33,CY0=2.89,CH=0.875,CW=3.57;
D.nonMachine.forEach((item,i)=>{
  const ry=CY0+i*(CH+0.025);
  const bgs=["FFF4EF","EEF6FB","F5F7FF"];
  s1.addShape("rect",{x:CX,y:ry,w:CW,h:CH,fill:{color:bgs[i]},line:{color:"E0E6F0",pt:0.5}});
  s1.addShape("rect",{x:CX,y:ry,w:0.05,h:CH,fill:{color:item.col}});
  s1.addText(String(item.rank),{x:CX+0.10,y:ry+0.06,w:0.20,h:0.22,fontSize:9,bold:true,color:item.col,fontFace:FC});
  s1.addText(item.name,{x:CX+0.33,y:ry+0.06,w:2.0,h:0.24,fontSize:10,bold:true,color:"333333",fontFace:FC});
  s1.addText(`${item.count}件・${item.pct}%`,{x:CX+0.33,y:ry+0.30,w:2.42,h:0.36,fontSize:18,bold:true,color:item.col,fontFace:FC});
  s1.addText(item.note||"",{x:CX+2.80,y:ry+0.08,w:0.73,h:0.62,fontSize:6,color:"888888",fontFace:FC,wrap:true,valign:"top"});
  s1.addShape("rect",{x:CX+0.05,y:ry+CH-0.06,w:CW-0.05,h:0.05,fill:{color:"EBEBEB"}});
  s1.addShape("rect",{x:CX+0.05,y:ry+CH-0.06,w:(CW-0.05)*(item.pct/100),h:0.05,fill:{color:item.col}});
});

// ══════════════════════════════════
//  PAGE 2
// ══════════════════════════════════
let s2 = pres.addSlide();
s2.background = {color:VI.white};
s2.addText(`客服課　${D.range}　機台問題佔比 與 熱門站點`,{
  x:0.18,y:0.06,w:9.6,h:0.26,fontSize:16,bold:true,color:VI.darkGray,fontFace:FC});
secBar(s2,0.18,0.36,"客訴詳情與分類佔比");
s2.addChart(pres.ChartType.pie,
  [{name:"機台",labels:D.machIssues.map(m=>m.label),values:D.machIssues.map(m=>m.count)}],
  {x:0.07,y:0.52,w:3.05,h:1.70,showLegend:false,showTitle:false,
   showPercent:true,dataLabelFormatCode:"0%",dataLabelFontSize:8,
   dataLabelFontBold:true,dataLabelColor:"FFFFFF",chartColors:MACH_C,shadow:{type:"none"}});

const ML=3.18,MLY0=0.54,MLH=0.245;
D.machIssues.forEach((m,i)=>{
  const y=MLY0+i*MLH;
  s2.addShape("rect",{x:ML,y:y+0.04,w:0.15,h:0.15,fill:{color:MACH_C[i]}});
  s2.addText(`${m.label}　${m.count}件（${m.pct}%）`,{x:ML+0.20,y,w:3.05,h:0.22,fontSize:8.5,color:"333333",fontFace:FC});
});

const bigCards=[
  {x:6.12,label:"收瓶機/方舟",pct:D.bottlePct,count:D.bottleTotal,bg:"FFF3EC",border:"FF5000",txt:"FF5000"},
  {x:7.32,label:"電池機",      pct:D.batPct,  count:D.batTotal,  bg:"EDF5FB",border:"8EB8C9",txt:"0076A9"},
];
bigCards.forEach(c=>{
  s2.addShape("rect",{x:c.x,y:0.52,w:1.10,h:1.70,fill:{color:c.bg},line:{color:c.border,pt:1.2}});
  s2.addText(c.label,{x:c.x,y:0.58,w:1.10,h:0.20,fontSize:8,color:"888888",align:"center",fontFace:FC});
  s2.addText(`${c.pct}%`,{x:c.x,y:0.80,w:1.10,h:0.56,fontSize:32,bold:true,color:c.txt,align:"center",fontFace:"Arial"});
  s2.addText(`（${c.count}件）`,{x:c.x,y:1.76,w:1.10,h:0.22,fontSize:8,color:"888888",align:"center",fontFace:FC});
});

secBar(s2,0.18,2.32,"收瓶機、電池機客訴 Top 3");
secBar(s2,6.20,2.32,`本週熱點站點（${D.week}：${D.range}）`,"FF5000");

const BSTART=2.52;
s2.addShape("rect",{x:0.18,y:BSTART,w:0.72,h:1.18,fill:{color:"060E9F"}});
s2.addText("收\n瓶\n機",{x:0.18,y:BSTART,w:0.72,h:1.18,fontSize:10.5,bold:true,color:"FFFFFF",align:"center",valign:"middle",fontFace:FC});
D.bottleTop3.forEach((item,i)=>{
  const ry=BSTART+0.02+i*0.385;
  const bg=i===0?"EDF3FF":i===1?"F4F8FF":"F9FBFF";
  s2.addShape("rect",{x:0.92,y:ry,w:5.18,h:0.365,fill:{color:bg},line:{color:"E0E6F0",pt:0.4}});
  s2.addText(`${i+1}`,{x:0.98,y:ry+0.07,w:0.20,h:0.22,fontSize:10,bold:true,color:item.col,fontFace:FC});
  s2.addText(item.name,{x:1.24,y:ry+0.07,w:3.3,h:0.22,fontSize:9.5,color:"333333",fontFace:FC});
  s2.addText(`${item.pct}%`,{x:5.26,y:ry+0.04,w:0.78,h:0.28,fontSize:20,bold:true,color:item.col,align:"right",valign:"middle",fontFace:"Arial"});
  s2.addShape("rect",{x:1.24,y:ry+0.315,w:3.55,h:0.038,fill:{color:"EBEBEB"}});
  s2.addShape("rect",{x:1.24,y:ry+0.315,w:3.55*(item.pct/100),h:0.038,fill:{color:item.col}});
});

const BATSTART=BSTART+1.195;
s2.addShape("rect",{x:0.18,y:BATSTART,w:0.72,h:0.80,fill:{color:"FF5000"}});
s2.addText("電\n池\n機",{x:0.18,y:BATSTART,w:0.72,h:0.80,fontSize:10.5,bold:true,color:"FFFFFF",align:"center",valign:"middle",fontFace:FC});
D.batTop.forEach((item,i)=>{
  const ry=BATSTART+0.02+i*0.378;
  s2.addShape("rect",{x:0.92,y:ry,w:5.18,h:0.358,fill:{color:i===0?"F0FDF7":"F7FAFE"},line:{color:"E0E6F0",pt:0.4}});
  s2.addText(`${i+1}`,{x:0.98,y:ry+0.07,w:0.20,h:0.22,fontSize:10,bold:true,color:item.col,fontFace:FC});
  s2.addText(item.name,{x:1.24,y:ry+0.07,w:3.3,h:0.22,fontSize:9.5,color:"333333",fontFace:FC});
  s2.addText(`${item.pct}%`,{x:5.26,y:ry+0.04,w:0.78,h:0.28,fontSize:20,bold:true,color:item.col,align:"right",valign:"middle",fontFace:"Arial"});
  s2.addShape("rect",{x:1.24,y:ry+0.305,w:3.55,h:0.038,fill:{color:"EBEBEB"}});
  s2.addShape("rect",{x:1.24,y:ry+0.305,w:3.55*(item.pct/100),h:0.038,fill:{color:item.col}});
});

s2.addShape("rect",{x:0.18,y:4.54,w:5.88,h:0.60,fill:{color:"FAE0B8"},line:{color:"FFCE00",pt:1}});
s2.addText(D.alertText,{x:0.28,y:4.56,w:5.72,h:0.56,fontSize:7.5,color:"6B4800",fontFace:FC,wrap:true,valign:"top"});

let aY=2.52;
D.hotAreas.forEach(area=>{
  if(aY>5.50) return;
  s2.addShape("rect",{x:6.20,y:aY,w:3.65,h:0.26,fill:{color:"FAE0B8"}});
  s2.addText(area.area,{x:6.26,y:aY,w:1.8,h:0.26,fontSize:12,bold:true,color:"060E9F",valign:"middle",fontFace:FC});
  s2.addText(String(area.total),{x:9.2,y:aY,w:0.60,h:0.26,fontSize:13,bold:true,color:"FF5000",align:"right",valign:"middle",fontFace:"Arial"});
  aY+=0.27;
  area.spots.forEach(spot=>{
    if(aY>5.50) return;
    if(spot.zone && spot.zone !== "" && spot.zone !== "nan") {
      // ・  [區別深黑粗體]  [站點名稱深灰]
      s2.addText([
        {text:`・  `,          options:{color:"333333", fontSize:8.5}},
        {text:`${spot.zone}  `,options:{color:"1A1A1A", fontSize:8.5, bold:true}},
        {text:spot.name,       options:{color:"333333", fontSize:8.5}},
      ], {x:6.28, y:aY, w:2.85, h:0.22, fontFace:FC, valign:"middle"});
    } else {
      s2.addText(`・${spot.name}`,{x:6.30,y:aY,w:2.80,h:0.22,fontSize:8.5,color:"333333",fontFace:FC,valign:"middle"});
    }
    s2.addText(String(spot.count),{x:9.2,y:aY,w:0.60,h:0.22,fontSize:9,bold:true,color:"888888",align:"right",fontFace:"Arial"});
    aY+=0.225;
  });
  aY+=0.02;
});

// ── 輸出 ──
pres.writeFile({fileName: path.join(BASE, FNAME)}).then(()=>{
  // 複製到 repo
  if (!fs.existsSync(REPO_DIR)) fs.mkdirSync(REPO_DIR, {recursive:true});
  fs.copyFileSync(path.join(BASE,FNAME), OUT_REPO);
  // 複製到本機輸出目錄（Windows路徑）
  try {
    if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, {recursive:true});
    fs.copyFileSync(path.join(BASE,FNAME), OUT_LOCAL);
    console.log(`✅ PPT 已存至本機：${OUT_LOCAL}`);
  } catch(e) {
    console.warn(`⚠ 本機路徑寫入失敗（${OUT_DIR}）：${e.message}`);
  }
  console.log(`✅ PPT 完成：${FNAME}`);
});
