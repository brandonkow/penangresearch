"use client"

import { useState, useRef, useCallback, useEffect } from "react"
import * as XLSX from "xlsx"
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  ComposedChart, RadarChart, Radar, PolarGrid, PolarAngleAxis,
  PolarRadiusAxis, ReferenceLine,
} from "recharts"
import {
  Upload, TrendingUp, Home, Building2, Factory, Hotel, Users, BarChart3,
  Brain, Send, Loader2, ChevronRight, RefreshCw, Sparkles, DollarSign,
  ArrowUpRight, ArrowDownRight, Minus, FileText, CheckCircle, Target,
  Search, Download, Globe, MapPin, Presentation, SlidersHorizontal,
} from "lucide-react"

// ─── DATA ─────────────────────────────────────────────────────────────────────
var D = {
  tx: [
    {y:"1998",v:12476,val:2578,rv:8651,rv2:1190,cv:1124,cv2:385,iv:499,iv2:370},
    {y:"2000",v:16339,val:3276,rv:12380,rv2:1857,cv:1424,cv2:473,iv:572,iv2:399},
    {y:"2002",v:15835,val:3433,rv:12194,rv2:1842,cv:1350,cv2:558,iv:407,iv2:245},
    {y:"2004",v:24292,val:5653,rv:19205,rv2:3238,cv:2101,cv2:695,iv:568,iv2:454},
    {y:"2006",v:20877,val:5490,rv:15439,rv2:2740,cv:2056,cv2:796,iv:518,iv2:433},
    {y:"2008",v:22323,val:7277,rv:16284,rv2:3718,cv:2683,cv2:1041,iv:566,iv2:673},
    {y:"2010",v:25986,val:9372,rv:18233,rv2:4834,cv:3082,cv2:1463,iv:561,iv2:657},
    {y:"2011",v:39415,val:13073,rv:30674,rv2:7723,cv:3320,cv2:1617,iv:779,iv2:974},
    {y:"2012",v:30977,val:12905,rv:23266,rv2:7093,cv:2826,cv2:1829,iv:724,iv2:832},
    {y:"2013",v:24346,val:13491,rv:17700,rv2:7093,cv:2240,cv2:1803,iv:549,iv2:804},
    {y:"2014",v:25555,val:13770,rv:18410,rv2:7579,cv:2397,cv2:1739,iv:504,iv2:700},
    {y:"2015",v:21555,val:11702,rv:15291,rv2:6173,cv:2140,cv2:2070,iv:632,iv2:735},
    {y:"2016",v:18352,val:9562,rv:13244,rv2:5365,cv:1696,cv2:1192,iv:443,iv2:909},
    {y:"2017",v:16592,val:10387,rv:12090,rv2:5413,cv:1344,cv2:1050,iv:452,iv2:678},
    {y:"2018",v:17087,val:10190,rv:12551,rv2:5465,cv:1304,cv2:1112,iv:458,iv2:1168},
    {y:"2019",v:17118,val:8633,rv:12718,rv2:4960,cv:1448,cv2:1270,iv:392,iv2:865},
    {y:"2020",v:15772,val:7815,rv:11736,rv2:4583,cv:1324,cv2:1302,iv:362,iv2:816},
    {y:"2021",v:17526,val:10448,rv:13648,rv2:5966,cv:1269,cv2:1317,iv:464,iv2:1613},
    {y:"2022",v:23481,val:13376,rv:17892,rv2:7959,cv:1973,cv2:1520,iv:522,iv2:1623},
    {y:"2023",v:24683,val:17069,rv:18663,rv2:8207,cv:2522,cv2:4437,iv:633,iv2:2055},
    {y:"2024",v:24428,val:15925,rv:18122,rv2:8374,cv:2416,cv2:2896,iv:500,iv2:1758},
  ],
  hpi: [
    {p:"2009",allH:92.7,terr:89.8,det:99.7,semi:101.6,hr:90.8},
    {p:"2011",allH:105.1,terr:102.6,det:93.7,semi:109.3,hr:108.2},
    {p:"2013",allH:132.2,terr:129.6,det:117.3,semi:131.1,hr:137.8},
    {p:"2015",allH:165.6,terr:154.6,det:168.2,semi:164.7,hr:179.2},
    {p:"2017",allH:186.2,terr:169.0,det:177.4,semi:183.2,hr:212.1},
    {p:"2019",allH:203.8,terr:185.1,det:195.2,semi:199.4,hr:242.1},
    {p:"2021",allH:212.4,terr:192.8,det:204.5,semi:207.6,hr:258.9},
    {p:"2023",allH:224.1,terr:203.6,det:215.8,semi:219.3,hr:278.2},
  ],
  hotels: [
    {y:"2009",rooms:9170,iR:8213,mR:957},{y:"2013",rooms:9438,iR:8155,mR:1283},
    {y:"2017",rooms:12155,iR:10287,mR:1868},{y:"2019",rooms:13210,iR:11342,mR:1868},
    {y:"2021",rooms:11696,iR:9690,mR:2006},{y:"2022",rooms:12807,iR:10801,mR:2006},
    {y:"2023",rooms:14337,iR:11968,mR:2369},{y:"2024",rooms:15661,iR:12949,mR:2712},
    {y:"2025",rooms:16770,iR:13726,mR:3044},
  ],
  overhang: [
    {p:"2016",res:533,com:0,ind:0},{p:"2018",res:889,com:0,ind:0},
    {p:"2019",res:1415,com:10,ind:0},{p:"2020",res:2131,com:77,ind:0},
    {p:"2021",res:2570,com:123,ind:0},{p:"2022",res:3012,com:123,ind:0},
    {p:"3Q23",res:3341,com:105,ind:7},{p:"3Q24",res:3150,com:81,ind:26},
  ],
  pop: [
    {y:"2014",p:1678100},{y:"2015",p:1698100},{y:"2016",p:1717700},
    {y:"2017",p:1744100},{y:"2018",p:1762800},{y:"2019",p:1768800},{y:"2020",p:1770400},
  ],
  pie: [
    {name:"Residential",val:8374,c:"#C0272D"},{name:"Commercial",val:2896,c:"#E05555"},
    {name:"Industrial",val:1758,c:"#F0AAAA"},{name:"Dev Land",val:1548,c:"#F7D0D0"},
    {name:"Agricultural",val:1347,c:"#FAEAEA"},
  ],
  indStock: [
    {y:"2009",isl:1312,main:6018},{y:"2013",isl:1296,main:6364},
    {y:"2017",isl:1296,main:6690},{y:"2021",isl:1296,main:7035},{y:"2024",isl:1296,main:7380},
  ],
  distTx: [
    {y:"2015",iV:9484,mV:12071,iA:6147,mA:5555},{y:"2016",iV:8075,mV:10277,iA:5023,mA:4539},
    {y:"2017",iV:7300,mV:9292,iA:5453,mA:4934},{y:"2018",iV:7518,mV:9569,iA:5350,mA:4840},
    {y:"2019",iV:7532,mV:9586,iA:4532,mA:4101},{y:"2020",iV:6940,mV:8832,iA:4104,mA:3711},
    {y:"2021",iV:7711,mV:9815,iA:5487,mA:4961},{y:"2022",iV:10332,mV:13149,iA:7027,mA:6349},
    {y:"2023",iV:10860,mV:13823,iA:8966,mA:8103},{y:"2024",iV:10748,mV:13680,iA:8361,mA:7564},
  ],
  afford: [
    {y:"2012",income:4039,price:304858,ratio:6.3},{y:"2016",income:4937,price:403816,ratio:6.8},
    {y:"2019",income:5650,price:390025,ratio:5.8},{y:"2022",income:6800,price:444813,ratio:5.4},
    {y:"2024",income:7200,price:462109,ratio:5.3},
  ],
}

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
var G = "#C0272D", G2 = "#E05555", G3 = "#F0AAAA", AM = "#F59E0B", RD = "#EF4444", PU = "#8B5CF6"
var CS = {fontFamily:"var(--font-sans)",fontSize:11,fill:"var(--color-text-secondary)"}
var CTX = "Senior property market analyst for Gamuda Land. Penang property market data:\n2024: 24,428 tx, RM15.9B, RM462K avg. Res:18,122/RM8.37B. Com:2,416/RM2.9B. Ind:500/RM1.76B.\nIsland ~44% transactions, higher prices. Mainland ~56%.\nHPI(2000=100): AllHouse 224, HiRise 278, Terr 204. Overhang 3,257(3Q24). Afford ratio 5.3x.\nIndustrial: 8,676 units, near-zero overhang, 15-25% Msia FDI. E&E hub.\nHotels: 71/16,770 rooms(2025). Airport 7.78M pax(2024). Pop 1.77M(2020). Mfg 40-45% GDP.\nGamuda Land is a leading Malaysian property developer with projects across Penang including townships, integrated developments, and affordable housing."

var TABS = [
  {id:"ov",  label:"Overview",           icon:BarChart3},
  {id:"tx",  label:"Transactions",       icon:TrendingUp},
  {id:"res", label:"Residential",        icon:Home},
  {id:"hpi", label:"Price Index",        icon:DollarSign},
  {id:"com", label:"Commercial",         icon:Building2},
  {id:"ind", label:"Industrial",         icon:Factory},
  {id:"hot", label:"Hotel & Tourism",    icon:Hotel},
  {id:"dem", label:"Demographics",       icon:Users},
  {id:"div"},
  {id:"fc",  label:"Forecasting",        icon:TrendingUp, badge:"AI"},
  {id:"scr", label:"Investment Screener",icon:Target,     badge:"AI"},
  {id:"rfr", label:"Data Refresh",       icon:RefreshCw,  badge:"AI"},
  {id:"ai",  label:"AI Analysis",        icon:Brain,      badge:"AI"},
]

var STIT = {
  tx:"Transaction Analysis", res:"Residential Market", hpi:"House Price Index",
  com:"Commercial Market",   ind:"Industrial Sector",  hot:"Hotel & Tourism", dem:"Demographics & Population",
}

// ─── UTILS ────────────────────────────────────────────────────────────────────
function fmt(n, d) {
  if (n == null) return ""
  if (d == null) d = 0
  return Number(n).toLocaleString("en-MY", {minimumFractionDigits:d, maximumFractionDigits:d})
}
function fmtRM(n)  { return n == null ? "" : "RM " + Number(n).toLocaleString("en-MY", {minimumFractionDigits:0, maximumFractionDigits:0}) }
function fmtRMM(n) { return n == null ? "" : "RM " + Number(n).toLocaleString("en-MY", {minimumFractionDigits:1, maximumFractionDigits:1}) + "M" }
function fmtPop(n) { return Number(n).toLocaleString("en-MY") }
function fmtSqm(n) { return Number(n).toLocaleString("en-MY") + " sqm" }
function fmtPax(n) { return Number(n).toLocaleString("en-MY") + "K pax" }

// All Claude API calls go through the secure server-side /api/claude route
async function callAI(sys, msgs, tok) {
  if (!tok) tok = 1800
  const res = await fetch("/api/claude", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({system: sys, messages: msgs, max_tokens: tok}),
  })
  const d = await res.json()
  if (d.content && d.content[0]) return d.content[0].text || ""
  if (d.error) throw new Error(d.error)
  return ""
}

function safeJSON(t) {
  try { return JSON.parse(t.replace(/```json|```/g, "").trim()) }
  catch (e) { return null }
}

function linReg(ys) {
  var n = ys.length
  var xs = ys.map((_, k) => k)
  var sx = 0, sy = 0, sxy = 0, sx2 = 0
  for (var i = 0; i < n; i++) { sx += xs[i]; sy += ys[i]; sxy += xs[i]*ys[i]; sx2 += xs[i]*xs[i] }
  var denom = n*sx2 - sx*sx
  var slope = denom !== 0 ? (n*sxy - sx*sy) / denom : 0
  var intercept = (sy - slope*sx) / n
  return function(x) { return Math.max(0, slope*x + intercept) }
}

function downloadHTML(html, fn) {
  var b = new Blob([html], {type:"text/html;charset=utf-8"})
  var u = URL.createObjectURL(b)
  var a = document.createElement("a")
  a.href = u; a.download = fn; a.click()
  URL.revokeObjectURL(u)
}

// ─── TOOLTIP ─────────────────────────────────────────────────────────────────
var TOOLTIP_STYLE = {
  contentStyle: {background:"#1f2937", border:"1px solid #374151", borderRadius:10, padding:"10px 14px", boxShadow:"0 6px 20px rgba(0,0,0,0.4)"},
  labelStyle:   {color:"#f9fafb", fontWeight:700, marginBottom:6, fontSize:12},
  itemStyle:    {color:"#d1d5db", fontSize:12},
}
function tipFormatterUnits(v) { return [fmt(v) + " units", "Volume"] }
function tipFormatterRMM(v)   { return fmtRMM(v) }
function tipFormatterPop(v)   { return [fmtPop(v), "Population"] }
function tipFormatterSqm(v)   { return [fmtSqm(v), "NLA"] }
function tipFormatterPax(v)   { return [fmtPax(v), "Passengers"] }
function tipFormatterRMVal(v, name) {
  if (name === "Ratio (x)") return v + "x"
  return fmtRM(v)
}

// ─── MARKDOWN RENDERER ────────────────────────────────────────────────────────
function inlineFmt(text) {
  return text.split(/(\*\*[^*]+\*\*|\*[^*]+\*)/g).map(function(p, i) {
    if (p.length > 4 && p.slice(0,2) === "**" && p.slice(-2) === "**")
      return <strong key={i}>{p.slice(2,-2)}</strong>
    if (p.length > 2 && p.slice(0,1) === "*" && p.slice(-1) === "*")
      return <em key={i}>{p.slice(1,-1)}</em>
    return p
  })
}

function MD({text}) {
  if (!text) return null
  var lines = text.split("\n")
  var out = [], i = 0
  while (i < lines.length) {
    var l = lines[i]
    if (!l.trim()) { out.push(<div key={"g"+i} style={{height:8}}/>); i++; continue }
    if (l.slice(0,3) === "## ") {
      out.push(<p key={i} style={{fontSize:14,fontWeight:600,color:G,margin:"14px 0 4px",paddingBottom:3,borderBottom:"1px solid var(--color-border-tertiary)"}}>{inlineFmt(l.slice(3))}</p>)
      i++; continue
    }
    if (l.slice(0,4) === "### ") {
      out.push(<p key={i} style={{fontSize:13,fontWeight:600,margin:"10px 0 3px"}}>{inlineFmt(l.slice(4))}</p>)
      i++; continue
    }
    if (l.slice(0,2) === "# ") {
      out.push(<p key={i} style={{fontSize:15,fontWeight:700,margin:"14px 0 5px"}}>{inlineFmt(l.slice(2))}</p>)
      i++; continue
    }
    if (l.match(/^[-*] /)) {
      var items = []
      while (i < lines.length && lines[i].match(/^[-*] /)) { items.push(lines[i].slice(2)); i++ }
      out.push(
        <ul key={"ul"+i} style={{margin:"5px 0",paddingLeft:18,display:"flex",flexDirection:"column",gap:3}}>
          {items.map((it,j) => <li key={j} style={{fontSize:12,lineHeight:1.65}}>{inlineFmt(it)}</li>)}
        </ul>
      )
      continue
    }
    if (l.match(/^\d+\. /)) {
      var oitems = []
      while (i < lines.length && lines[i].match(/^\d+\. /)) { oitems.push(lines[i].replace(/^\d+\. /,"")); i++ }
      out.push(
        <ol key={"ol"+i} style={{margin:"5px 0",paddingLeft:20,display:"flex",flexDirection:"column",gap:3}}>
          {oitems.map((it,j) => <li key={j} style={{fontSize:12,lineHeight:1.65}}>{inlineFmt(it)}</li>)}
        </ol>
      )
      continue
    }
    out.push(<p key={i} style={{fontSize:12,lineHeight:1.75,margin:"2px 0"}}>{inlineFmt(l)}</p>)
    i++
  }
  return <div>{out}</div>
}

// ─── SVG CHARTS (for PDF export) ─────────────────────────────────────────────
function nf(v) {
  if (v >= 1e6) return (v/1e6).toFixed(1) + "M"
  if (v >= 1e3) return (v/1e3).toFixed(0) + "K"
  return Number(v||0).toFixed(0)
}

function svgBar(rows, xk, yk, fill) {
  var W=560, H=200, Pl=54, Pt=12, Pr=8, Pb=42
  var cW=W-Pl-Pr, cH=H-Pt-Pb
  var vals = rows.map(r => Math.max(Number(r[yk])||0, 0))
  var mx = Math.max(...vals, 1) * 1.1
  var bw = Math.min(cW/rows.length*0.6, 32), gap = cW/rows.length
  var step = Math.max(1, Math.ceil(rows.length/9))
  var g = "", b = ""
  for (var gi = 0; gi < 5; gi++) {
    var gv = mx*gi/4, gy = (Pt + cH*(1-gi/4)).toFixed(1)
    g += `<line x1="${Pl}" x2="${Pl+cW}" y1="${gy}" y2="${gy}" stroke="#e5e7eb" stroke-width="0.7"/>`
    g += `<text x="${Pl-5}" y="${(parseFloat(gy)+3).toFixed(1)}" text-anchor="end" font-size="9" fill="#9ca3af">${nf(gv)}</text>`
  }
  for (var bi = 0; bi < rows.length; bi++) {
    var rv = Number(rows[bi][yk])||0, bH = cH*rv/mx
    var bx = (Pl + bi*gap + (gap-bw)/2).toFixed(1), by = (Pt+cH-bH).toFixed(1)
    b += `<rect x="${bx}" y="${by}" width="${bw.toFixed(1)}" height="${Math.max(bH,0).toFixed(1)}" fill="${fill}" rx="2"/>`
    if (bi%step===0) b += `<text x="${(Pl+bi*gap+gap/2).toFixed(1)}" y="${(Pt+cH+13).toFixed(1)}" text-anchor="middle" font-size="8" fill="#9ca3af">${rows[bi][xk]}</text>`
  }
  return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">${g}${b}<line x1="${Pl}" x2="${Pl}" y1="${Pt}" y2="${Pt+cH}" stroke="#d1d5db"/><line x1="${Pl}" x2="${Pl+cW}" y1="${Pt+cH}" y2="${Pt+cH}" stroke="#d1d5db"/></svg>`
}

function svgSBar(rows, xk, keys, fills) {
  var W=560, H=200, Pl=54, Pt=12, Pr=8, Pb=42
  var cW=W-Pl-Pr, cH=H-Pt-Pb
  var tots = rows.map(r => keys.reduce((s,k) => s+(Number(r[k])||0), 0))
  var mx = Math.max(...tots, 1) * 1.1
  var bw = Math.min(cW/rows.length*0.6, 32), gap = cW/rows.length
  var step = Math.max(1, Math.ceil(rows.length/9))
  var g = "", b = ""
  for (var gi = 0; gi < 5; gi++) {
    var gv = mx*gi/4, gy = (Pt+cH*(1-gi/4)).toFixed(1)
    g += `<line x1="${Pl}" x2="${Pl+cW}" y1="${gy}" y2="${gy}" stroke="#e5e7eb" stroke-width="0.7"/>`
    g += `<text x="${Pl-5}" y="${(parseFloat(gy)+3).toFixed(1)}" text-anchor="end" font-size="9" fill="#9ca3af">${nf(gv)}</text>`
  }
  for (var bi = 0; bi < rows.length; bi++) {
    var yo = 0, bx = (Pl+bi*gap+(gap-bw)/2).toFixed(1)
    for (var ki = 0; ki < keys.length; ki++) {
      var rv = Number(rows[bi][keys[ki]])||0, bH = cH*rv/mx
      var by = (Pt+cH-yo-bH).toFixed(1)
      if (bH > 0) b += `<rect x="${bx}" y="${by}" width="${bw.toFixed(1)}" height="${bH.toFixed(1)}" fill="${fills[ki]}" rx="${ki===keys.length-1?2:0}"/>`
      yo += bH
    }
    if (bi%step===0) b += `<text x="${(Pl+bi*gap+gap/2).toFixed(1)}" y="${(Pt+cH+13).toFixed(1)}" text-anchor="middle" font-size="8" fill="#9ca3af">${rows[bi][xk]}</text>`
  }
  return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">${g}${b}<line x1="${Pl}" x2="${Pl}" y1="${Pt}" y2="${Pt+cH}" stroke="#d1d5db"/><line x1="${Pl}" x2="${Pl+cW}" y1="${Pt+cH}" y2="${Pt+cH}" stroke="#d1d5db"/></svg>`
}

function svgLine(rows, xk, yks, strokes) {
  var W=560, H=200, Pl=54, Pt=12, Pr=8, Pb=42
  var cW=W-Pl-Pr, cH=H-Pt-Pb
  var allV = rows.flatMap(r => yks.map(k => Number(r[k])||0))
  var mx = Math.max(...allV, 1) * 1.08
  var n = rows.length, step = Math.max(1, Math.ceil(n/9))
  var xi = i => Pl + i*cW/Math.max(n-1,1)
  var yi = v => Pt + cH*(1-v/mx)
  var g = ""
  for (var gi = 0; gi < 5; gi++) {
    var gv = mx*gi/4, gy = yi(gv).toFixed(1)
    g += `<line x1="${Pl}" x2="${Pl+cW}" y1="${gy}" y2="${gy}" stroke="#e5e7eb" stroke-width="0.7"/>`
    g += `<text x="${Pl-5}" y="${(parseFloat(gy)+3).toFixed(1)}" text-anchor="end" font-size="9" fill="#9ca3af">${nf(gv)}</text>`
  }
  var xl = rows.map((r,li) => li%step===0 ? `<text x="${xi(li).toFixed(1)}" y="${(Pt+cH+13).toFixed(1)}" text-anchor="middle" font-size="8" fill="#9ca3af">${r[xk]}</text>` : "").join("")
  var paths = yks.map((k,pi) => {
    var d = rows.map((r,di) => (di===0?"M":"L") + xi(di).toFixed(1) + "," + yi(Number(r[k])||0).toFixed(1)).join("")
    return `<path d="${d}" stroke="${strokes[pi]}" stroke-width="2" fill="none" stroke-linejoin="round"/>`
  }).join("")
  return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto">${g}${xl}${paths}<line x1="${Pl}" x2="${Pl}" y1="${Pt}" y2="${Pt+cH}" stroke="#d1d5db"/><line x1="${Pl}" x2="${Pl+cW}" y1="${Pt+cH}" y2="${Pt+cH}" stroke="#d1d5db"/></svg>`
}

function getChart(s) {
  if (s === "tx")  return svgSBar(D.tx, "y", ["rv","cv","iv"], [G,G2,G3])
  if (s === "res") return svgLine(D.tx, "y", ["rv","rv2"], [G,AM])
  if (s === "hpi") return svgLine(D.hpi, "p", ["allH","terr","hr"], [G,G2,RD])
  if (s === "com") return svgBar(D.tx.slice(-12), "y", "cv2", G2)
  if (s === "ind") return svgSBar(D.indStock, "y", ["isl","main"], [G,G2])
  if (s === "hot") return svgSBar(D.hotels, "y", ["iR","mR"], [G,G2])
  if (s === "dem") return svgLine(D.pop, "y", ["p"], [G])
  return svgBar(D.tx.slice(-10), "y", "v", G)
}

var CHART_LEGENDS = {
  tx:  '<span class="leg"><span class="dot" style="background:#C0272D"></span>Residential</span><span class="leg"><span class="dot" style="background:#E05555"></span>Commercial</span><span class="leg"><span class="dot" style="background:#F0AAAA"></span>Industrial</span>',
  res: '<span class="leg"><span class="dot" style="background:#C0272D"></span>Volume</span><span class="leg"><span class="dot" style="background:#F59E0B"></span>Value (RM M)</span>',
  hpi: '<span class="leg"><span class="dot" style="background:#C0272D"></span>All House</span><span class="leg"><span class="dot" style="background:#E05555"></span>Terraced</span><span class="leg"><span class="dot" style="background:#8B5CF6"></span>High-Rise</span>',
  ind: '<span class="leg"><span class="dot" style="background:#C0272D"></span>Island</span><span class="leg"><span class="dot" style="background:#E05555"></span>Mainland</span>',
  hot: '<span class="leg"><span class="dot" style="background:#C0272D"></span>Island Rooms</span><span class="leg"><span class="dot" style="background:#E05555"></span>Mainland Rooms</span>',
}

// ─── PDF BUILDER ──────────────────────────────────────────────────────────────
function buildSectionPDF(secId, insight) {
  var title = STIT[secId] || secId
  var date = new Date().toLocaleDateString("en-MY", {year:"numeric",month:"long",day:"numeric"})
  var chart = getChart(secId)
  var legend = CHART_LEGENDS[secId] || ""
  var insightHTML = insight.split("\n").filter(p => p.trim()).map(p => `<p>${p}</p>`).join("")
  var css = "@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Inter',sans-serif;color:#111827;background:#fff;font-size:13px}.cov{background:linear-gradient(140deg,#1A0508,#C0272D 60%,#E04040);color:#fff;padding:52px 64px 40px}.badge{display:inline-block;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 14px;font-size:10px;letter-spacing:.1em;text-transform:uppercase;margin-bottom:22px}.title{font-size:36px;font-weight:700;line-height:1.2;margin-bottom:10px}.sub{font-size:13px;color:rgba(255,255,255,.7);line-height:1.6}.rule{border:none;border-top:1px solid rgba(255,255,255,.2);margin:22px 0}.meta{display:flex;gap:40px}.ml{font-size:9px;text-transform:uppercase;letter-spacing:.1em;color:rgba(255,255,255,.4);margin-bottom:4px}.mv{font-size:12px;font-weight:500}.body{padding:40px 64px}.st{font-size:20px;font-weight:700;color:#111827;margin-bottom:4px}.rule2{width:40px;height:3px;background:#C0272D;border-radius:2px;margin-bottom:24px}.legend{display:flex;gap:14px;margin-bottom:14px;align-items:center}.leg{display:flex;align-items:center;gap:5px;font-size:10.5px;color:#6b7280}.dot{display:inline-block;width:10px;height:10px;border-radius:2px;flex-shrink:0}.cw{background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:18px 20px;margin-bottom:8px}.ib{margin-top:28px;border-left:3px solid #C0272D;padding:18px 22px;background:#fdf2f2;border-radius:0 10px 10px 0}.ilbl{font-size:9.5px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#C0272D;margin-bottom:12px}.ib p{font-size:12.5px;line-height:1.85;color:#374151;margin-bottom:10px}.footer{padding:14px 64px;background:#f9fafb;border-top:1px solid #e5e7eb;display:flex;justify-content:space-between}.footer span{font-size:10px;color:#9ca3af}@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}"
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${css}</style></head><body><div class="cov"><div class="badge">Gamuda Intelligence - Property Market Data</div><div class="title">${title}</div><div class="sub">Gamuda Property Market Intelligence - Property market data, 1995-2024</div><hr class="rule"/><div class="meta"><div><div class="ml">Date</div><div class="mv">${date}</div></div><div><div class="ml">Source</div><div class="mv">Market Data</div></div><div><div class="ml">AI</div><div class="mv">Claude Sonnet</div></div></div></div><div class="body"><div class="st">${title}</div><div class="rule2"></div>${legend ? `<div class="legend">${legend}</div>` : ""}<div class="cw">${chart}</div><div class="ib"><div class="ilbl">AI Analyst Insights</div>${insightHTML}</div></div><div class="footer"><span>Gamuda Intelligence Dashboard</span><span>${date} - Download HTML then Print to save as PDF</span></div><script>window.onload=function(){setTimeout(function(){window.print()},600)}<\/script></body></html>`
}

// ─── SLIDES GENERATOR ────────────────────────────────────────────────────────
function buildSlidesDeck(bullets) {
  var date = new Date().toLocaleDateString("en-MY", {year:"numeric",month:"long",day:"numeric"})
  var SECTIONS = [
    {title:"Market Overview",    chart:getChart("ov"),  key:"overview"},
    {title:"Transactions",       chart:getChart("tx"),  key:"transactions"},
    {title:"Residential Market", chart:getChart("res"), key:"residential"},
    {title:"House Price Index",  chart:getChart("hpi"), key:"priceIndex"},
    {title:"Commercial Market",  chart:getChart("com"), key:"commercial"},
    {title:"Industrial Sector",  chart:getChart("ind"), key:"industrial"},
    {title:"Hotel & Tourism",    chart:getChart("hot"), key:"hotel"},
    {title:"Demographics",       chart:getChart("dem"), key:"demographics"},
  ]
  var css = `@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Inter',sans-serif;background:#111}@media print{body{background:#fff}.slide{margin:0;box-shadow:none;page-break-after:always}}.slide{width:1280px;height:720px;position:relative;overflow:hidden;margin:0 auto 20px;background:#fff;box-shadow:0 4px 24px rgba(0,0,0,.35)}.cover{background:linear-gradient(135deg,#1A0508 0%,#C0272D 55%,#E04040 100%)}.cover-inner{padding:88px;display:flex;flex-direction:column;justify-content:center;height:100%}.badge{display:inline-block;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:20px;padding:6px 18px;font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.8);margin-bottom:36px;width:fit-content}.cover-title{font-size:54px;font-weight:700;color:#fff;line-height:1.15;margin-bottom:18px}.cover-sub{font-size:19px;color:rgba(255,255,255,.65);margin-bottom:52px}.cover-meta{display:flex;gap:52px;padding-top:32px;border-top:1px solid rgba(255,255,255,.18)}.ml{font-size:10px;text-transform:uppercase;letter-spacing:.1em;color:rgba(255,255,255,.4);margin-bottom:5px}.mv{font-size:13px;color:#fff;font-weight:500}.slide-header{height:88px;background:#C0272D;display:flex;align-items:center;padding:0 52px;justify-content:space-between}.slide-title{font-size:28px;font-weight:700;color:#fff}.slide-logo{display:flex;align-items:center;gap:10px}.mark{width:36px;height:36px;background:rgba(255,255,255,.2);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:17px;font-weight:700;color:#fff}.logo-name{font-size:12px;color:rgba(255,255,255,.8);font-weight:500}.slide-body{display:flex;height:590px}.chart-area{flex:1.45;padding:24px 20px 24px 44px;display:flex;flex-direction:column;justify-content:center;border-right:1px solid #f0f0f0}.chart-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#C0272D;margin-bottom:14px}.insights{flex:1;padding:32px 40px;display:flex;flex-direction:column;justify-content:center;background:#fafafa}.ins-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#C0272D;margin-bottom:18px;padding-bottom:10px;border-bottom:2px solid #C0272D}.bullet{display:flex;gap:12px;margin-bottom:16px;align-items:flex-start}.dot{width:7px;height:7px;background:#C0272D;border-radius:50%;margin-top:6px;flex-shrink:0}.bullet p{font-size:14px;line-height:1.65;color:#1f2937}.slide-footer{position:absolute;bottom:0;left:0;right:0;height:40px;background:#f9fafb;border-top:1px solid #e5e7eb;display:flex;align-items:center;justify-content:space-between;padding:0 44px}.slide-footer span{font-size:10px;color:#9ca3af}`
  var cover = `<div class="slide cover"><div class="cover-inner"><div class="badge">Gamuda Property Intelligence</div><div class="cover-title">Penang Property<br/>Market Intelligence</div><div class="cover-sub">Market Overview · Data Analysis · Strategic Insights</div><div class="cover-meta"><div><div class="ml">Date</div><div class="mv">${date}</div></div><div><div class="ml">Market</div><div class="mv">Penang, Malaysia</div></div><div><div class="ml">Data Period</div><div class="mv">1995 – 2024</div></div><div><div class="ml">Powered by</div><div class="mv">Claude AI</div></div></div></div></div>`
  var slides = SECTIONS.map((s,i) => {
    var pts = (bullets && bullets[s.key]) || ["Market data for this section is available in the dashboard.","Review charts for detailed trend analysis.","AI insights available on demand."]
    var bHTML = pts.map(p => `<div class="bullet"><div class="dot"></div><p>${p}</p></div>`).join("")
    return `<div class="slide"><div class="slide-header"><div class="slide-title">${s.title}</div><div class="slide-logo"><div class="mark">G</div><div class="logo-name">Gamuda Intelligence</div></div></div><div class="slide-body"><div class="chart-area"><div class="chart-label">Market Data · 1995–2024</div>${s.chart}</div><div class="insights"><div class="ins-label">Key Insights</div>${bHTML}</div></div><div class="slide-footer"><span>Gamuda Property Intelligence · Penang</span><span>Slide ${i+2} of ${SECTIONS.length+1}</span></div></div>`
  }).join("")
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Gamuda Property Intelligence</title><style>${css}</style></head><body>${cover}${slides}</body></html>`
}

function SlidesButton() {
  var [loading, setLoading] = useState(false)
  function run() {
    setLoading(true)
    var prompt = `Generate 3 concise insight bullet points for each Penang property market section for a presentation. Return valid JSON only:\n{"overview":["p1","p2","p3"],"transactions":["p1","p2","p3"],"residential":["p1","p2","p3"],"priceIndex":["p1","p2","p3"],"commercial":["p1","p2","p3"],"industrial":["p1","p2","p3"],"hotel":["p1","p2","p3"],"demographics":["p1","p2","p3"]}\nRules: max 18 words each, use specific numbers, insight-focused not descriptive.`
    callAI(CTX, [{role:"user",content:prompt}], 2000)
      .then(r => {
        var bullets = safeJSON(r)
        var d = new Date().toLocaleDateString("en-MY").replace(/\//g,"-")
        downloadHTML(buildSlidesDeck(bullets), `Gamuda_Penang_Slides_${d}.html`)
        setLoading(false)
      })
      .catch(() => {
        var d = new Date().toLocaleDateString("en-MY").replace(/\//g,"-")
        downloadHTML(buildSlidesDeck(null), `Gamuda_Penang_Slides_${d}.html`)
        setLoading(false)
      })
  }
  return (
    <button onClick={run} disabled={loading} style={{padding:"7px 16px",borderRadius:8,background:loading?"var(--color-background-secondary)":"#1A0508",border:`1px solid ${loading?"var(--color-border-secondary)":"#C0272D"}`,cursor:loading?"default":"pointer",color:loading?"var(--color-text-secondary)":"#fff",fontWeight:500,fontSize:12,display:"flex",alignItems:"center",gap:7,whiteSpace:"nowrap"}}>
      {loading ? <Loader2 size={13} style={{animation:"spin 1s linear infinite"}}/> : <Presentation size={13} color={loading?"var(--color-text-secondary)":"#E05555"}/>}
      {loading ? "Generating..." : "Generate Slides"}
    </button>
  )
}

// ─── SECTION REPORT BUTTON ────────────────────────────────────────────────────
function SectionReport({secId}) {
  var [loading, setLoading] = useState(false)
  function run() {
    setLoading(true)
    var title = STIT[secId] || secId
    var prompt = `Write a detailed professional analyst report for the Penang ${title} section using Gamuda property market data.\n\n## Market Overview\n2 paragraphs with key figures.\n\n## Key Trends\n2 paragraphs on trends and drivers.\n\n## Risks and Opportunities\n1 paragraph on risks. 1 paragraph on opportunities.\n\n## Outlook\n1 paragraph on the 12-24 month outlook.\n\nBe direct. Use specific figures.`
    callAI(CTX, [{role:"user",content:prompt}], 1500)
      .then(insight => {
        var date = new Date().toLocaleDateString("en-MY").replace(/\//g,"-")
        downloadHTML(buildSectionPDF(secId, insight), `Penang_${title.replace(/ /g,"_")}_${date}.html`)
        setLoading(false)
      })
      .catch(() => {
        downloadHTML(buildSectionPDF(secId, "AI insights could not be generated."), `Penang_${title.replace(/ /g,"_")}.html`)
        setLoading(false)
      })
  }
  return (
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:16}}>
      <button onClick={run} disabled={loading} style={{padding:"8px 18px",borderRadius:8,background:loading?"var(--color-background-secondary)":G,border:"none",cursor:loading?"default":"pointer",color:loading?"var(--color-text-secondary)":"#fff",fontWeight:500,fontSize:12,display:"flex",alignItems:"center",gap:7,boxShadow:loading?"none":"0 1px 6px rgba(0,106,77,0.25)"}}>
        {loading ? <Loader2 size={13} style={{animation:"spin 1s linear infinite"}}/> : <Download size={13}/>}
        {loading ? "Generating PDF..." : "Download PDF Report"}
      </button>
    </div>
  )
}

// ─── SHARED UI ────────────────────────────────────────────────────────────────
function KPI({label, value, sub, trend, accent}) {
  var up = trend > 0, dn = trend < 0
  return (
    <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1rem 1.25rem",display:"flex",flexDirection:"column",gap:5,borderLeft:accent?`3px solid ${accent}`:undefined}}>
      <span style={{fontSize:11,color:"var(--color-text-secondary)",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.05em"}}>{label}</span>
      <span style={{fontSize:20,fontWeight:500,color:"var(--color-text-primary)",lineHeight:1.2}}>{value}</span>
      {sub && <span style={{fontSize:11,color:"var(--color-text-secondary)"}}>{sub}</span>}
      {trend != null && (
        <div style={{display:"flex",alignItems:"center",gap:4}}>
          {up ? <ArrowUpRight size={13} color={G}/> : dn ? <ArrowDownRight size={13} color="#dc2626"/> : <Minus size={13} color="#9ca3af"/>}
          <span style={{fontSize:11,color:up?G:dn?"#dc2626":"#9ca3af",fontWeight:500}}>{Math.abs(trend).toFixed(1)}% YoY</span>
        </div>
      )}
    </div>
  )
}

function Box({title, h, children}) {
  return (
    <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1rem 1.25rem",height:h}}>
      {title && <p style={{fontWeight:500,marginBottom:12,fontSize:13,margin:"0 0 12px"}}>{title}</p>}
      {children}
    </div>
  )
}

function DistrictFilter({dist, setDist, yearRange, setYearRange, propType, setPropType}) {
  var distOpts = [
    {id:"all",      label:"All Penang",     icon:Globe},
    {id:"island",   label:"Penang Island",  icon:MapPin},
    {id:"mainland", label:"Seberang Perai", icon:MapPin},
  ]
  var propOpts = [
    {id:"all", label:"All Types"},
    {id:"res", label:"Residential"},
    {id:"com", label:"Commercial"},
    {id:"ind", label:"Industrial"},
  ]
  var years = D.tx.map(r => r.y)
  var selStyle = {padding:"4px 8px",borderRadius:6,fontSize:11,border:"0.5px solid var(--color-border-secondary)",background:"var(--color-background-primary)",color:"var(--color-text-primary)",cursor:"pointer"}
  return (
    <div style={{background:"var(--color-background-secondary)",borderRadius:10,marginBottom:16,border:"0.5px solid var(--color-border-tertiary)",overflow:"hidden"}}>
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",flexWrap:"wrap"}}>
        <SlidersHorizontal size={12} color="var(--color-text-secondary)"/>
        <span style={{fontSize:11,fontWeight:500,color:"var(--color-text-secondary)",textTransform:"uppercase",letterSpacing:"0.06em"}}>District:</span>
        {distOpts.map(o => {
          var active = dist === o.id
          var Icon = o.icon
          return (
            <button key={o.id} onClick={() => setDist(o.id)} style={{padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:active?600:400,cursor:"pointer",background:active?G:"transparent",color:active?"#fff":"var(--color-text-secondary)",border:`0.5px solid ${active?G:"var(--color-border-secondary)"}`,display:"flex",alignItems:"center",gap:4}}>
              <Icon size={10}/>{o.label}
            </button>
          )
        })}
        <div style={{display:"flex",alignItems:"center",gap:6,marginLeft:"auto"}}>
          <span style={{fontSize:11,fontWeight:500,color:"var(--color-text-secondary)"}}>Year:</span>
          <select value={yearRange.from} onChange={e => setYearRange(p => ({...p, from:e.target.value}))} style={selStyle}>
            {years.filter(y => y <= yearRange.to).map(y => <option key={y} value={y}>{y}</option>)}
          </select>
          <span style={{fontSize:11,color:"var(--color-text-secondary)"}}>–</span>
          <select value={yearRange.to} onChange={e => setYearRange(p => ({...p, to:e.target.value}))} style={selStyle}>
            {years.filter(y => y >= yearRange.from).map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6,padding:"6px 12px 8px",borderTop:"0.5px solid var(--color-border-tertiary)"}}>
        <span style={{fontSize:11,fontWeight:500,color:"var(--color-text-secondary)",textTransform:"uppercase",letterSpacing:"0.06em"}}>Type:</span>
        {propOpts.map(o => {
          var active = propType === o.id
          return (
            <button key={o.id} onClick={() => setPropType(o.id)} style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:active?600:400,cursor:"pointer",background:active?G:"transparent",color:active?"#fff":"var(--color-text-secondary)",border:`0.5px solid ${active?G:"var(--color-border-secondary)"}`}}>
              {o.label}
            </button>
          )
        })}
      </div>
    </div>
  )
}

// ─── STANDARD TABS ────────────────────────────────────────────────────────────
function Overview({dist, yearRange}) {
  var last = D.tx[D.tx.length-1], prev = D.tx[D.tx.length-2]
  var dt = D.distTx[D.distTx.length-1], dt2 = D.distTx[D.distTx.length-2]
  var vol = dist==="island"?dt.iV : dist==="mainland"?dt.mV : last.v
  var val = dist==="island"?dt.iA : dist==="mainland"?dt.mA : last.val
  var pvol = dist==="island"?dt2.iV : dist==="mainland"?dt2.mV : prev.v
  var pval = dist==="island"?dt2.iA : dist==="mainland"?dt2.mA : prev.val
  var yr = yearRange || {from:"1998", to:"2024"}
  var recent = (dist==="island" ? D.distTx.map(r=>({y:r.y,v:r.iV})) : dist==="mainland" ? D.distTx.map(r=>({y:r.y,v:r.mV})) : D.tx.map(r=>({y:r.y,v:r.v}))).filter(r => r.y >= yr.from && r.y <= yr.to)
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(155px,1fr))",gap:12}}>
        <KPI label="Transactions (2024)" value={fmt(vol)+" units"} trend={(vol-pvol)/pvol*100}/>
        <KPI label="Total Value (2024)" value={fmtRMM(val)} trend={(val-pval)/pval*100}/>
        <KPI label="Avg Value/Tx" value={fmtRM(Math.round(val*1e6/vol))}/>
        <KPI label="Hotel Rooms (2025)" value={fmt(dist==="mainland"?3044:dist==="island"?13726:16770)}/>
        <KPI label="Res Overhang" value={fmt(D.overhang[D.overhang.length-1].res)+" units"}/>
        <KPI label="HPI All House" value={D.hpi[D.hpi.length-1].allH.toFixed(1)} sub="Base 2000=100"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1.6fr 1fr",gap:16}}>
        <Box title="Annual Transaction Volume">
          <ResponsiveContainer width="100%" height={200}>
            <AreaChart data={recent} margin={{top:4,right:8,bottom:0,left:0}}>
              <defs><linearGradient id="vg1" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={G} stopOpacity={0.15}/><stop offset="95%" stopColor={G} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterUnits}/>
              <Area type="monotone" dataKey="v" name="Volume" stroke={G} fill="url(#vg1)" strokeWidth={2} dot={false}/>
            </AreaChart>
          </ResponsiveContainer>
        </Box>
        <Box title="Value by Sector (2024)">
          <ResponsiveContainer width="100%" height={150}>
            <PieChart>
              <Pie data={D.pie} cx="50%" cy="50%" innerRadius={40} outerRadius={70} paddingAngle={2} dataKey="val">
                {D.pie.map((e,i) => <Cell key={i} fill={e.c}/>)}
              </Pie>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterRMM}/>
            </PieChart>
          </ResponsiveContainer>
          {D.pie.map((e,i) => (
            <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:11,marginTop:3}}>
              <span style={{width:8,height:8,borderRadius:2,background:e.c,flexShrink:0}}/>
              <span style={{color:"var(--color-text-secondary)",flex:1}}>{e.name}</span>
              <span style={{fontWeight:500}}>{fmtRMM(e.val)}</span>
            </div>
          ))}
        </Box>
      </div>
    </div>
  )
}

function Transactions({dist, yearRange, propType}) {
  var [view, setView] = useState("vol")
  var yr = yearRange || {from:"1998", to:"2024"}
  var pt = propType || "all"
  var data = (dist==="all" ? D.tx : D.distTx).filter(r => r.y >= yr.from && r.y <= yr.to)
  var vk = dist==="island"?"iV" : dist==="mainland"?"mV" : (pt==="res"?"rv" : pt==="com"?"cv" : pt==="ind"?"iv" : "v")
  var ak = dist==="island"?"iA" : dist==="mainland"?"mA" : (pt==="res"?"rv2" : pt==="com"?"cv2" : pt==="ind"?"iv2" : "val")
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="tx"/>
      <div style={{display:"flex",gap:8}}>
        {["vol","val","stack"].map(id => (
          <button key={id} onClick={() => setView(id)} style={{padding:"6px 14px",borderRadius:6,fontSize:12,fontWeight:500,cursor:"pointer",background:view===id?G:"transparent",color:view===id?"#fff":"var(--color-text-secondary)",border:`0.5px solid ${view===id?G:"var(--color-border-secondary)"}`}}>
            {id==="vol"?"Volume":id==="val"?"Value":"Sector Split"}
          </button>
        ))}
      </div>
      <Box>
        <ResponsiveContainer width="100%" height={300}>
          {view==="stack" && dist==="all" ? (
            <BarChart data={D.tx.slice(-12)} margin={{top:4,right:8,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Bar dataKey="rv" name="Residential" fill={G} stackId="s"/>
              <Bar dataKey="cv" name="Commercial" fill={G2} stackId="s"/>
              <Bar dataKey="iv" name="Industrial" fill={G3} stackId="s" radius={[3,3,0,0]}/>
              <Legend iconType="square" iconSize={10} wrapperStyle={{fontSize:11}}/>
            </BarChart>
          ) : view==="stack" ? (
            <BarChart data={D.distTx} margin={{top:4,right:8,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Bar dataKey="iV" name="Island" fill={G} stackId="s"/>
              <Bar dataKey="mV" name="Mainland" fill={G2} stackId="s" radius={[3,3,0,0]}/>
              <Legend iconType="square" iconSize={10} wrapperStyle={{fontSize:11}}/>
            </BarChart>
          ) : view==="val" ? (
            <AreaChart data={data} margin={{top:4,right:8,bottom:0,left:0}}>
              <defs><linearGradient id="vg2" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={G} stopOpacity={0.15}/><stop offset="95%" stopColor={G} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => "RM "+fmt(v)+"M"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterRMM}/>
              <Area type="monotone" dataKey={ak} name="Value" stroke={G} fill="url(#vg2)" strokeWidth={2} dot={false}/>
            </AreaChart>
          ) : (
            <BarChart data={data} margin={{top:4,right:8,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Bar dataKey={vk} name="Volume" fill={G} radius={[3,3,0,0]}/>
            </BarChart>
          )}
        </ResponsiveContainer>
      </Box>
    </div>
  )
}

function Residential({dist}) {
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="res"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <KPI label="Res Volume (2024)" value={fmt(dist==="island"?7974:dist==="mainland"?10148:18122)+" units"} trend={-2.9}/>
        <KPI label="Res Value (2024)" value={fmtRMM(dist==="island"?4697:dist==="mainland"?3677:8374)} trend={2.0}/>
        <KPI label="Avg Price (2024)" value={fmtRM(dist==="island"?589000:dist==="mainland"?362000:462000)}/>
        <KPI label="Overhang 3Q 2024" value={fmt(3257)+" units"} trend={-5.7}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Box title="Overhang Trend">
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={D.overhang} margin={{top:4,right:4,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="p" tick={{fontFamily:"var(--font-sans)",fontSize:9,fill:"var(--color-text-secondary)"}} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Bar dataKey="res" name="Residential" fill={G} stackId="s"/>
              <Bar dataKey="com" name="Commercial" fill={AM} stackId="s"/>
              <Bar dataKey="ind" name="Industrial" fill={RD} stackId="s" radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Box>
        <Box title="Affordability Ratio (x)">
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={D.afford} margin={{top:4,right:30,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis yAxisId="l" tick={CS} tickLine={false} axisLine={false} tickFormatter={v => "RM "+fmt(v/1000)+"K"}/>
              <YAxis yAxisId="r" orientation="right" tick={CS} tickLine={false} axisLine={false} tickFormatter={v => v+"x"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterRMVal}/>
              <Bar yAxisId="l" dataKey="price" name="Avg Price" fill={G} radius={[3,3,0,0]} opacity={0.8}/>
              <Line yAxisId="r" type="monotone" dataKey="ratio" name="Ratio (x)" stroke={AM} strokeWidth={2.5} dot={{r:4,fill:AM}}/>
            </ComposedChart>
          </ResponsiveContainer>
        </Box>
      </div>
    </div>
  )
}

function HPI() {
  var [sel, setSel] = useState(["allH","hr"])
  var LINES = [
    {k:"allH",label:"All House",c:G},
    {k:"terr",label:"Terraced",c:G2},
    {k:"det", label:"Detached",c:AM},
    {k:"semi",label:"Semi-D",  c:PU},
    {k:"hr",  label:"High-Rise",c:RD},
  ]
  var last = D.hpi[D.hpi.length-1]
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="hpi"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12}}>
        <KPI label="All House HPI" value={last.allH.toFixed(1)} sub="Base 2000=100"/>
        <KPI label="High-Rise HPI" value={last.hr.toFixed(1)} accent={RD}/>
        <KPI label="Terraced HPI" value={last.terr.toFixed(1)}/>
        <KPI label="Semi-D HPI" value={last.semi.toFixed(1)}/>
      </div>
      <Box>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <p style={{fontWeight:500,fontSize:13,margin:0}}>House Price Index (2009-2023)</p>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {LINES.map(l => {
              var on = sel.includes(l.k)
              return (
                <button key={l.k} onClick={() => setSel(p => on ? p.filter(x=>x!==l.k) : [...p,l.k])} style={{padding:"3px 10px",borderRadius:20,fontSize:11,cursor:"pointer",background:on?l.c+"22":"transparent",color:on?l.c:"var(--color-text-secondary)",border:`1px solid ${on?l.c:"var(--color-border-tertiary)"}`}}>
                  {l.label}
                </button>
              )
            })}
          </div>
        </div>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={D.hpi} margin={{top:4,right:8,bottom:0,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
            <XAxis dataKey="p" tick={{fontFamily:"var(--font-sans)",fontSize:9,fill:"var(--color-text-secondary)"}} tickLine={false} axisLine={false}/>
            <YAxis tick={CS} tickLine={false} axisLine={false} domain={[85,290]}/>
            <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle}/>
            {LINES.filter(l => sel.includes(l.k)).map(l => (
              <Line key={l.k} type="monotone" dataKey={l.k} name={l.label} stroke={l.c} strokeWidth={2} dot={false}/>
            ))}
          </LineChart>
        </ResponsiveContainer>
      </Box>
    </div>
  )
}

function Commercial() {
  var ret = [
    {loc:"George Town",sqm:438184},{loc:"BM Town",sqm:324657},
    {loc:"Seberang Jaya",sqm:117416},{loc:"Sungai Nibong",sqm:105492},
    {loc:"Bayan Baru",sqm:95821},{loc:"Pulau Tikus",sqm:52131},
  ]
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="com"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <KPI label="Commercial Vol (2024)" value={fmt(2416)+" units"} trend={-4.2}/>
        <KPI label="Commercial Val (2024)" value={fmtRMM(2896)} trend={-34.7}/>
        <KPI label="PBO Rental Index" value="135.8" sub="2019 Q4"/>
        <KPI label="GT Retail NLA" value="438,184 sqm" sub="81% Occupancy"/>
      </div>
      <Box title="Retail NLA by Location (sqm)">
        <ResponsiveContainer width="100%" height={210}>
          <BarChart data={ret} layout="vertical" margin={{top:0,right:30,bottom:0,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" horizontal={false}/>
            <XAxis type="number" tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
            <YAxis dataKey="loc" type="category" tick={{fontFamily:"var(--font-sans)",fontSize:10,fill:"var(--color-text-secondary)"}} tickLine={false} axisLine={false} width={90}/>
            <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterSqm}/>
            <Bar dataKey="sqm" name="NLA" fill={G} radius={[0,3,3,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </Box>
    </div>
  )
}

function Industrial({dist}) {
  var stockData = dist==="island" ? D.indStock.map(r=>({y:r.y,v:r.isl})) : dist==="mainland" ? D.indStock.map(r=>({y:r.y,v:r.main})) : null
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="ind"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <KPI label="Industrial Stock" value={fmt(dist==="island"?1296:dist==="mainland"?7380:8676)+" units"} sub={dist==="island"?"Island":dist==="mainland"?"Mainland":"All Penang"}/>
        <KPI label="Industrial Vol (2024)" value={fmt(500)+" units"} trend={-21.0}/>
        <KPI label="Industrial Val (2024)" value={fmtRMM(1758)} trend={-14.5}/>
        <KPI label="Overhang (2025)" value="7 units" sub="Near zero"/>
      </div>
      <Box title="Industrial Stock - Island vs Mainland">
        <ResponsiveContainer width="100%" height={220}>
          {dist==="all" ? (
            <BarChart data={D.indStock} margin={{top:4,right:4,bottom:0,left:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Bar dataKey="isl" name="Island" fill={G} stackId="s"/>
              <Bar dataKey="main" name="Mainland" fill={G2} stackId="s" radius={[3,3,0,0]}/>
              <Legend iconType="square" iconSize={10} wrapperStyle={{fontSize:11}}/>
            </BarChart>
          ) : (
            <AreaChart data={stockData} margin={{top:4,right:4,bottom:0,left:0}}>
              <defs><linearGradient id="igs" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={G} stopOpacity={0.15}/><stop offset="95%" stopColor={G} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" units"}/>
              <Area type="monotone" dataKey="v" name="Units" stroke={G} fill="url(#igs)" strokeWidth={2}/>
            </AreaChart>
          )}
        </ResponsiveContainer>
      </Box>
    </div>
  )
}

function HotelTab({dist}) {
  var ap = [
    {y:"2016",p:6777},{y:"2018",p:7735},{y:"2019",p:7510},{y:"2020",p:1942},
    {y:"2021",p:935},{y:"2022",p:4832},{y:"2023",p:7145},{y:"2024",p:7780},
  ]
  var rk = dist==="island"?"iR" : dist==="mainland"?"mR" : "rooms"
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="hot"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <KPI label="Hotels (2025)" value={fmt(dist==="island"?57:dist==="mainland"?14:71)} trend={9.2}/>
        <KPI label="Rooms (2025)" value={fmt(dist==="island"?13726:dist==="mainland"?3044:16770)} trend={7.1}/>
        <KPI label="Island Share" value="82% of rooms" sub="2025"/>
        <KPI label="Airport Pax" value="7.78M" sub="2024" trend={8.9}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <Box title="Hotel Room Supply">
          <ResponsiveContainer width="100%" height={200}>
            {dist==="all" ? (
              <BarChart data={D.hotels} margin={{top:4,right:4,bottom:0,left:0}}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
                <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
                <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
                <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" rooms"}/>
                <Bar dataKey="iR" name="Island" fill={G} stackId="s"/>
                <Bar dataKey="mR" name="Mainland" fill={G2} stackId="s" radius={[3,3,0,0]}/>
                <Legend iconType="square" iconSize={10} wrapperStyle={{fontSize:11}}/>
              </BarChart>
            ) : (
              <BarChart data={D.hotels} margin={{top:4,right:4,bottom:0,left:0}}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
                <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
                <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => fmt(v/1000)+"K"}/>
                <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={v => fmt(v)+" rooms"}/>
                <Bar dataKey={rk} name="Rooms" fill={G} radius={[3,3,0,0]}/>
              </BarChart>
            )}
          </ResponsiveContainer>
        </Box>
        <Box title="Airport Passengers ('000)">
          <ResponsiveContainer width="100%" height={200}>
            <AreaChart data={ap} margin={{top:4,right:4,bottom:0,left:0}}>
              <defs><linearGradient id="apg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={G} stopOpacity={0.15}/><stop offset="95%" stopColor={G} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => (v/1000).toFixed(0)+"M"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterPax}/>
              <Area type="monotone" dataKey="p" name="Passengers" stroke={G} fill="url(#apg)" strokeWidth={2} dot={false}/>
            </AreaChart>
          </ResponsiveContainer>
        </Box>
      </div>
    </div>
  )
}

function Demographics() {
  var eth = [
    {n:"Chinese",       v:619712, c:G2},
    {n:"Bumiputera",    v:473260, c:G},
    {n:"Non-Malaysian", v:314311, c:PU},
    {n:"Indian",        v:181937, c:AM},
    {n:"Others",        v:12330,  c:G3},
  ]
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <SectionReport secId="dem"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <KPI label="Population (2020)" value="1,770,405"/>
        <KPI label="Penang Island" value="794,313" sub="44.9%"/>
        <KPI label="Seberang Perai" value="976,092" sub="55.1%"/>
        <KPI label="Pop CAGR" value="0.8%" sub="2014-2020"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1.2fr 1fr",gap:16}}>
        <Box title="Population Trend">
          <ResponsiveContainer width="100%" height={210}>
            <AreaChart data={D.pop} margin={{top:4,right:4,bottom:0,left:0}}>
              <defs><linearGradient id="pg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={G} stopOpacity={0.15}/><stop offset="95%" stopColor={G} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
              <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
              <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={v => (v/1e6).toFixed(2)+"M"}/>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterPop}/>
              <Area type="monotone" dataKey="p" name="Population" stroke={G} fill="url(#pg)" strokeWidth={2} dot={{r:3,fill:G}}/>
            </AreaChart>
          </ResponsiveContainer>
        </Box>
        <Box title="Ethnic Composition (2020)">
          <ResponsiveContainer width="100%" height={130}>
            <PieChart>
              <Pie data={eth} cx="50%" cy="50%" innerRadius={32} outerRadius={58} paddingAngle={2} dataKey="v">
                {eth.map((e,i) => <Cell key={i} fill={e.c}/>)}
              </Pie>
              <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle} formatter={tipFormatterPop}/>
            </PieChart>
          </ResponsiveContainer>
          {eth.map((e,i) => (
            <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:11,marginTop:3}}>
              <span style={{width:8,height:8,borderRadius:2,background:e.c,flexShrink:0}}/>
              <span style={{flex:1,color:"var(--color-text-secondary)"}}>{e.n}</span>
              <span style={{fontWeight:500}}>{fmtPop(e.v)}</span>
            </div>
          ))}
        </Box>
      </div>
    </div>
  )
}

// ─── FORECAST CHART ───────────────────────────────────────────────────────────
function ForecastChart({title, chartData, hk="hist", pk="proj", hc=G, ry="2024", yf=v=>String(v)}) {
  return (
    <Box title={title}>
      <ResponsiveContainer width="100%" height={190}>
        <ComposedChart data={chartData} margin={{top:4,right:8,bottom:0,left:0}}>
          <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border-tertiary)" vertical={false}/>
          <XAxis dataKey="y" tick={CS} tickLine={false} axisLine={false}/>
          <YAxis tick={CS} tickLine={false} axisLine={false} tickFormatter={yf}/>
          <Tooltip contentStyle={TOOLTIP_STYLE.contentStyle} labelStyle={TOOLTIP_STYLE.labelStyle} itemStyle={TOOLTIP_STYLE.itemStyle}/>
          <ReferenceLine x={ry} stroke="#9ca3af" strokeDasharray="4 3"/>
          <Line type="monotone" dataKey={hk} name="Historical" stroke={hc} strokeWidth={2} dot={false} connectNulls/>
          <Line type="monotone" dataKey={pk} name="Forecast" stroke={AM} strokeWidth={2.5} strokeDasharray="6 3" dot={{r:5,fill:AM}} connectNulls/>
        </ComposedChart>
      </ResponsiveContainer>
      <p style={{fontSize:10,color:"var(--color-text-secondary)",marginTop:5}}>Solid = historical · Dashed amber = regression projection</p>
    </Box>
  )
}

// ─── FORECASTING AGENT ────────────────────────────────────────────────────────
function ForecastingAgent() {
  var [loading, setLoading] = useState(false)
  var [narrative, setNarrative] = useState(null)
  var recent = D.tx.slice(-10)
  var vols   = recent.map(r => r.v)
  var vals   = recent.map(r => r.val)
  var prices = recent.map(r => Math.round(r.rv2*1e6/r.rv))
  var hpiVals = D.hpi.slice(-8).map(r => r.allH)
  var fV = linReg(vols), fA = linReg(vals), fP = linReg(prices), fH = linReg(hpiVals)
  var n = vols.length, nh = hpiVals.length
  var p5v = Math.round(fV(n)), p7v = Math.round(fV(n+2))
  var p5p = Math.round(fP(n)), p7p = Math.round(fP(n+2))
  var vCagr = (((p7v/vols[vols.length-1])-1)/3*100).toFixed(1)
  var pCagr = (((p7p/prices[prices.length-1])-1)/3*100).toFixed(1)
  function makeChart(getData, getProj) {
    var hist = recent.slice(-7).map(r => ({y:r.y, hist:getData(r), proj:null}))
    return [...hist, {y:"2025",hist:null,proj:getProj(0)}, {y:"2026",hist:null,proj:getProj(1)}, {y:"2027",hist:null,proj:getProj(2)}]
  }
  var volChart   = makeChart(r => r.v, i => Math.round(fV(n+i)))
  var priceChart = makeChart(r => Math.round(r.rv2*1e6/r.rv), i => Math.round(fP(n+i)))
  var valChart   = makeChart(r => r.val, i => Math.round(fA(n+i)*10)/10)
  var hpiChart   = [...D.hpi.slice(-6).map(r=>({y:r.p,hist:r.allH,proj:null})), {y:"2025",hist:null,proj:Math.round(fH(nh)*10)/10},{y:"2026",hist:null,proj:Math.round(fH(nh+1)*10)/10},{y:"2027",hist:null,proj:Math.round(fH(nh+2)*10)/10}]
  function run() {
    setLoading(true)
    var prompt = `Write a professional 3-year Penang property market forecast.\nProjections: 2025 Volume ${fmt(p5v)} units, Avg Price ${fmtRM(p5p)}. 2027 Volume ${fmt(p7v)} units, Avg Price ${fmtRM(p7p)}.\n\n## Forecast Overview\n## Transaction Volume Outlook\n## Price Appreciation Forecast\n## Sectoral Outlook\n## Key Upside Risks\n## Key Downside Risks\n## Conclusion\n\nBe direct. Use specific numbers.`
    callAI(CTX, [{role:"user",content:prompt}], 2000)
      .then(r => { setNarrative(r); setLoading(false) })
      .catch(() => { setNarrative("Connection error."); setLoading(false) })
  }
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{background:"linear-gradient(135deg,#1a1a3a,#0f0f24)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",alignItems:"center",justifyContent:"space-between",gap:14}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <div style={{background:"rgba(245,158,11,0.2)",borderRadius:10,padding:10,display:"flex"}}><TrendingUp size={22} color={AM}/></div>
          <div>
            <p style={{fontWeight:500,color:"#fff",fontSize:15,margin:0}}>Forecasting Agent</p>
            <p style={{fontSize:12,color:"rgba(255,255,255,0.6)",margin:0}}>Linear regression · 10-year market data · 2025-2027 projection</p>
          </div>
        </div>
        <button onClick={run} disabled={loading} style={{padding:"9px 20px",borderRadius:8,background:loading?"rgba(255,255,255,0.1)":AM,border:"none",cursor:loading?"default":"pointer",color:loading?"rgba(255,255,255,0.4)":"#fff",fontWeight:500,fontSize:12,display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
          {loading ? <Loader2 size={14} style={{animation:"spin 1s linear infinite"}}/> : <Sparkles size={14}/>}
          {loading ? "Generating..." : narrative ? "Regenerate" : "Generate AI Forecast"}
        </button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        <KPI label="Proj Volume 2025" value={fmt(p5v)+" units"} accent={AM}/>
        <KPI label="Proj Volume 2027" value={fmt(p7v)+" units"} sub={"CAGR "+vCagr+"%"} accent={AM}/>
        <KPI label="Proj Avg Price 2025" value={fmtRM(p5p)} accent={G}/>
        <KPI label="Proj Avg Price 2027" value={fmtRM(p7p)} sub={"CAGR "+pCagr+"%"} accent={G}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <ForecastChart title="Transaction Volume Forecast" chartData={volChart} yf={v => fmt(v/1000)+"K"}/>
        <ForecastChart title="Avg Residential Price Forecast" chartData={priceChart} yf={v => "RM "+fmt(v/1000)+"K"}/>
        <ForecastChart title="Transaction Value Forecast (RM M)" chartData={valChart} hk="hist" pk="proj" hc={G2} yf={v => "RM "+fmt(v)+"M"}/>
        <ForecastChart title="House Price Index Forecast" chartData={hpiChart} hc={PU} ry="2023" yf={v => String(v)}/>
      </div>
      {(loading || narrative) && (
        <Box title="">
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,paddingBottom:10,borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
            <div style={{background:AM+"22",borderRadius:8,padding:7,display:"flex"}}><Sparkles size={15} color={AM}/></div>
            <p style={{fontWeight:500,fontSize:13,margin:0}}>AI Market Forecast Narrative</p>
          </div>
          {loading ? (
            <div style={{display:"flex",alignItems:"center",gap:10,padding:"16px 0"}}>
              <Loader2 size={16} color={AM} style={{animation:"spin 1s linear infinite"}}/>
              <span style={{fontSize:13,color:"var(--color-text-secondary)"}}>Generating 3-year forecast...</span>
            </div>
          ) : <MD text={narrative}/>}
        </Box>
      )}
    </div>
  )
}

// ─── AI CHAT ──────────────────────────────────────────────────────────────────
function AIAnalysis() {
  var [msgs, setMsgs] = useState([{role:"assistant",content:"Hello! I am your Gamuda AI analyst. Ask me anything about market trends, investment analysis, or sector comparisons."}])
  var [input, setInput] = useState("")
  var [loading, setLoading] = useState(false)
  var endRef = useRef(null)
  useEffect(() => { if (endRef.current) endRef.current.scrollIntoView({behavior:"smooth"}) }, [msgs, loading])
  var send = useCallback(function(text) {
    var msg = text || input
    if (!msg || !msg.trim() || loading) return
    msg = msg.trim()
    setInput("")
    setLoading(true)
    var history = [...msgs, {role:"user",content:msg}]
    setMsgs(history)
    callAI(CTX, history.map(m => ({role:m.role,content:m.content})))
      .then(r => { setMsgs(p => [...p, {role:"assistant",content:r}]); setLoading(false) })
      .catch(() => { setMsgs(p => [...p, {role:"assistant",content:"Connection error. Please try again."}]); setLoading(false) })
  }, [input, loading, msgs])
  var QUICK = ["What drove the 2023 commercial spike?","Is Penang affordable for M40?","Best investment opportunity now?","Industrial FDI outlook","Island vs Mainland comparison"]
  return (
    <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 200px)",minHeight:500}}>
      <div style={{background:"linear-gradient(135deg,#C0272D,#8B0000)",borderRadius:"var(--border-radius-lg)",padding:"1rem 1.25rem",marginBottom:16,display:"flex",alignItems:"center",gap:12}}>
        <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:8}}><Sparkles size={20} color="#fff"/></div>
        <div>
          <p style={{fontWeight:500,color:"#fff",fontSize:14,margin:0}}>Gamuda AI Analyst</p>
          <p style={{fontSize:12,color:"rgba(255,255,255,0.75)",margin:0}}>Powered by Claude Sonnet</p>
        </div>
      </div>
      <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:12,paddingRight:4,marginBottom:12}}>
        {msgs.map((m,i) => (
          <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start",flexDirection:m.role==="user"?"row-reverse":"row"}}>
            <div style={{width:30,height:30,borderRadius:50,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",background:m.role==="user"?G:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-secondary)"}}>
              {m.role==="user" ? <span style={{fontSize:12,color:"#fff",fontWeight:500}}>U</span> : <Sparkles size={14} color={G}/>}
            </div>
            <div style={{maxWidth:"82%",background:m.role==="user"?G:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,borderTopRightRadius:m.role==="user"?4:12,borderTopLeftRadius:m.role==="user"?12:4,padding:"10px 14px"}}>
              {m.role==="user" ? <p style={{fontSize:13,lineHeight:1.65,margin:0,color:"#fff"}}>{m.content}</p> : <MD text={m.content}/>}
            </div>
          </div>
        ))}
        {loading && (
          <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
            <div style={{width:30,height:30,borderRadius:50,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-secondary)",display:"flex",alignItems:"center",justifyContent:"center"}}><Sparkles size={14} color={G}/></div>
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:12,borderTopLeftRadius:4,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}>
              <Loader2 size={14} color={G} style={{animation:"spin 1s linear infinite"}}/><span style={{fontSize:13,color:"var(--color-text-secondary)"}}>Analysing...</span>
            </div>
          </div>
        )}
        <div ref={endRef}/>
      </div>
      {msgs.length <= 1 && (
        <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12}}>
          {QUICK.map(q => <button key={q} onClick={() => send(q)} style={{padding:"6px 12px",borderRadius:20,fontSize:11,cursor:"pointer",background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-secondary)",color:"var(--color-text-secondary)"}}>{q}</button>)}
        </div>
      )}
      <div style={{display:"flex",gap:8,background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-secondary)",borderRadius:12,padding:"8px 8px 8px 14px",alignItems:"flex-end"}}>
        <textarea value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => { if (e.key==="Enter" && !e.shiftKey) { e.preventDefault(); send() }}} placeholder="Ask about Penang real estate..." rows={2} style={{flex:1,border:"none",outline:"none",resize:"none",fontSize:13,lineHeight:1.5,background:"transparent",color:"var(--color-text-primary)",fontFamily:"var(--font-sans)"}}/>
        <button onClick={() => send()} disabled={loading || !input.trim()} style={{width:36,height:36,borderRadius:8,background:input.trim()&&!loading?G:"var(--color-background-secondary)",border:"none",cursor:input.trim()&&!loading?"pointer":"default",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
          <Send size={16} color={input.trim()&&!loading?"#fff":"var(--color-text-secondary)"}/>
        </button>
      </div>
    </div>
  )
}

// ─── DATA REFRESH ─────────────────────────────────────────────────────────────
function DataRefresh() {
  var [checking, setChecking] = useState(false)
  var [intel, setIntel] = useState(null)
  var [upResult, setUpResult] = useState(null)
  var [uploading, setUploading] = useState(false)
  var [mode, setMode] = useState("int")
  var fileRef = useRef(null)
  function check() {
    setChecking(true); setIntel(null)
    var p = "Write a Penang property market intelligence update for 2024-2025 in clean markdown.\n## Market Headline\n## Latest Developments\n3-4 paragraphs tagged Positive/Neutral/Negative\n## Key Metrics to Watch\nBullet list.\n## 2025 Outlook\n## Data Gaps"
    callAI(CTX, [{role:"user",content:p}], 1800)
      .then(r => { setIntel(r); setChecking(false) })
      .catch(() => { setIntel("Connection error."); setChecking(false) })
  }
  function handleUp(e) {
    var f = e.target.files && e.target.files[0]
    if (!f) return
    setUploading(true); setUpResult(null)
    f.arrayBuffer().then(buf => {
      var wb = XLSX.read(buf, {type:"array"})
      var sum = `File: ${f.name}, ${wb.SheetNames.length} sheets: ${wb.SheetNames.join(", ")}`
      var p = `Compare uploaded property data file to dashboard (2024: 24,428 tx, RM15.9B, RM462K avg, 3,257 overhang).\nFile: ${sum}\n## File Assessment\n## Sheets Found\n## Potential Updates\nBullet: Field - current to detected - action\n## Recommendation`
      return callAI(CTX, [{role:"user",content:p}], 1200)
    }).then(r => { setUpResult(r); setUploading(false) })
      .catch(() => { setUpResult("Error."); setUploading(false) })
  }
  var modes = [{id:"int",label:"Market Intelligence",icon:Search},{id:"up",label:"Smart Upload",icon:Upload}]
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{background:"linear-gradient(135deg,#1a2a3a,#0f1a24)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",alignItems:"center",gap:14}}>
        <div style={{background:"rgba(59,130,246,0.2)",borderRadius:10,padding:10,display:"flex"}}><RefreshCw size={22} color="#60a5fa"/></div>
        <div>
          <p style={{fontWeight:500,color:"#fff",fontSize:15,margin:0}}>Data Refresh Agent</p>
          <p style={{fontSize:12,color:"rgba(255,255,255,0.6)",margin:0}}>Market intelligence or upload new property data file for AI change detection</p>
        </div>
      </div>
      <div style={{display:"flex",gap:8}}>
        {modes.map(m => {
          var Icon = m.icon
          return (
            <button key={m.id} onClick={() => setMode(m.id)} style={{padding:"8px 16px",borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",background:mode===m.id?G:"transparent",color:mode===m.id?"#fff":"var(--color-text-secondary)",border:`0.5px solid ${mode===m.id?G:"var(--color-border-secondary)"}`,display:"flex",alignItems:"center",gap:6}}>
              <Icon size={13}/>{m.label}
            </button>
          )
        })}
      </div>
      {mode === "int" && (
        <div style={{display:"grid",gridTemplateColumns:"260px 1fr",gap:16}}>
          <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",flexDirection:"column",gap:12}}>
            <p style={{fontWeight:500,fontSize:13,margin:0}}>Market Intelligence Check</p>
            <p style={{fontSize:12,color:"var(--color-text-secondary)"}}>Query Claude for latest Penang property developments.</p>
            <button onClick={check} disabled={checking} style={{padding:"10px",borderRadius:8,background:checking?"var(--color-background-secondary)":G,border:"none",cursor:checking?"default":"pointer",color:checking?"var(--color-text-secondary)":"#fff",fontWeight:500,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
              {checking ? <Loader2 size={14} style={{animation:"spin 1s linear infinite"}}/> : <Search size={14}/>}
              {checking ? "Fetching..." : "Check for Updates"}
            </button>
          </div>
          {intel ? (
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",overflowY:"auto",maxHeight:440}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><CheckCircle size={16} color="#C0272D"/><p style={{fontWeight:500,fontSize:13,margin:0}}>Intelligence Report</p></div>
              <MD text={intel}/>
            </div>
          ) : (
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",alignItems:"center",justifyContent:"center"}}>
              <p style={{fontSize:13,color:"var(--color-text-secondary)"}}>Click Check for Updates to fetch market intelligence</p>
            </div>
          )}
        </div>
      )}
      {mode === "up" && (
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleUp} style={{display:"none"}}/>
          <div style={{background:"var(--color-background-primary)",border:"2px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-lg)",padding:"2rem",display:"flex",flexDirection:"column",alignItems:"center",gap:12,cursor:"pointer"}} onClick={() => fileRef.current && fileRef.current.click()}>
            {uploading ? <Loader2 size={28} color={G} style={{animation:"spin 1s linear infinite"}}/> : <Upload size={28} color="var(--color-text-secondary)"/>}
            <p style={{fontWeight:500,fontSize:14,margin:0}}>{uploading ? "Analysing file..." : "Upload New Property Data File"}</p>
          </div>
          {upResult && (
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem"}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><CheckCircle size={16} color="#C0272D"/><p style={{fontWeight:500,fontSize:13,margin:0}}>File Analysis Complete</p></div>
              <MD text={upResult}/>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

// ─── INVESTMENT SCREENER ──────────────────────────────────────────────────────
function Screener() {
  var DISTS = ["Penang Island - Timur Laut (NE)","Penang Island - Barat Daya (SW)","Seberang Perai Utara","Seberang Perai Tengah","Seberang Perai Selatan"]
  var TYPES = ["High-Rise Residential","Terraced House","Semi-Detached","Detached House","Shop Office","Industrial Factory","Serviced Apartment"]
  var [form, setForm] = useState({dist:DISTS[0],type:TYPES[0],budget:"500000",horizon:"5",purpose:"Capital Appreciation"})
  var [screening, setScreening] = useState(false)
  var [result, setResult] = useState(null)
  function setField(k, v) { setForm(p => ({...p,[k]:v})) }
  function screen() {
    setScreening(true); setResult(null)
    var p = `Screen Penang property. Return ONLY valid JSON no fences:\n{"overall_score":7.5,"verdict":"RECOMMENDED","verdict_reason":"one sentence","scores":{"price_growth":{"score":8,"label":"Price Growth Potential","comment":"2 sentences"},"liquidity":{"score":7,"label":"Market Liquidity","comment":"2 sentences"},"affordability":{"score":5,"label":"Entry Affordability","comment":"2 sentences"},"supply_demand":{"score":8,"label":"Supply-Demand Balance","comment":"2 sentences"},"rental_yield":{"score":6,"label":"Rental Yield Est.","comment":"yield range"},"macro_fundamentals":{"score":9,"label":"Economic Fundamentals","comment":"2 sentences"}},"key_metrics":{"est_price_range":"RM X-Y","hpi_5yr_cagr":"X%","overhang_risk":"Low/Med/High","rental_yield_est":"X-Y%","capital_gain_5yr":"approx X%"},"pros":["p1","p2","p3"],"cons":["c1","c2"],"recommendation":"3 sentences actionable prose."}\nDistrict=${form.dist}, Type=${form.type}, Budget=${fmtRM(parseInt(form.budget))}, Horizon=${form.horizon}yr, Purpose=${form.purpose}`
    callAI(CTX, [{role:"user",content:p}], 1800)
      .then(r => { setResult(safeJSON(r) || {verdict:"ERROR",verdict_reason:r,scores:{},pros:[],cons:[],overall_score:0,key_metrics:{}}); setScreening(false) })
      .catch(() => { setResult({verdict:"ERROR",verdict_reason:"Connection failed.",scores:{},pros:[],cons:[],overall_score:0,key_metrics:{}}); setScreening(false) })
  }
  function vs(v) {
    if (!v) return {bg:"#f3f4f6",color:"#6b7280",border:"#e5e7eb"}
    if (v==="RECOMMENDED") return {bg:"#fff5f5",color:"#9b1c1c",border:"#fca5a5"}
    if (v==="AVOID") return {bg:"#fef2f2",color:"#991b1b",border:"#fca5a5"}
    return {bg:"#fefce8",color:"#854d0e",border:"#fde68a"}
  }
  var vst = vs(result && result.verdict)
  var rd = result && result.scores ? Object.values(result.scores).map(s => ({subject:s.label.split(" ").slice(0,2).join(" "),A:s.score,fullMark:10})) : []
  var fields = [{label:"District",key:"dist",opts:DISTS},{label:"Property Type",key:"type",opts:TYPES},{label:"Purpose",key:"purpose",opts:["Capital Appreciation","Rental Income","Both","Owner Occupation"]}]
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{background:"linear-gradient(135deg,#2a1a3a,#1a0f24)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",alignItems:"center",gap:14}}>
        <div style={{background:"rgba(139,92,246,0.2)",borderRadius:10,padding:10,display:"flex"}}><Target size={22} color="#a78bfa"/></div>
        <div>
          <p style={{fontWeight:500,color:"#fff",fontSize:15,margin:0}}>Investment Screener Agent</p>
          <p style={{fontSize:12,color:"rgba(255,255,255,0.6)",margin:0}}>AI scores any Penang property across 6 investment dimensions</p>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"300px 1fr",gap:16}}>
        <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"1.25rem",display:"flex",flexDirection:"column",gap:14}}>
          <p style={{fontWeight:500,fontSize:13,margin:0}}>Screen an Investment</p>
          {fields.map(f => (
            <div key={f.key}>
              <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:"0 0 5px",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.04em"}}>{f.label}</p>
              <select value={form[f.key]} onChange={e => setField(f.key, e.target.value)} style={{width:"100%",padding:"7px 10px",borderRadius:6,fontSize:12,border:"0.5px solid var(--color-border-secondary)",background:"var(--color-background-secondary)",color:"var(--color-text-primary)"}}>
                {f.opts.map(o => <option key={o}>{o}</option>)}
              </select>
            </div>
          ))}
          <div>
            <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:"0 0 5px",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.04em"}}>Budget (RM)</p>
            <input type="number" value={form.budget} onChange={e => setField("budget", e.target.value)} style={{width:"100%",padding:"7px 10px",borderRadius:6,fontSize:12,border:"0.5px solid var(--color-border-secondary)",background:"var(--color-background-secondary)",color:"var(--color-text-primary)"}}/>
            <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:"4px 0 0"}}>{form.budget ? fmtRM(parseInt(form.budget)) : ""}</p>
          </div>
          <div>
            <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:"0 0 5px",fontWeight:500,textTransform:"uppercase",letterSpacing:"0.04em"}}>Horizon: {form.horizon} years</p>
            <input type="range" min="1" max="15" value={form.horizon} onChange={e => setField("horizon", e.target.value)} style={{width:"100%",accentColor:G}}/>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:"var(--color-text-secondary)"}}><span>1yr</span><span>5yr</span><span>10yr</span><span>15yr</span></div>
          </div>
          <button onClick={screen} disabled={screening} style={{padding:"11px",borderRadius:8,background:screening?"var(--color-background-secondary)":PU,border:"none",cursor:screening?"default":"pointer",color:screening?"var(--color-text-secondary)":"#fff",fontWeight:500,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
            {screening ? <Loader2 size={15} style={{animation:"spin 1s linear infinite"}}/> : <Target size={15}/>}
            {screening ? "Screening..." : "Run AI Screening"}
          </button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          {!result && !screening && (
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"2rem",display:"flex",flexDirection:"column",alignItems:"center",gap:12,flex:1}}>
              <Target size={40} color="var(--color-text-secondary)" strokeWidth={1}/>
              <p style={{fontSize:14,color:"var(--color-text-secondary)"}}>Configure parameters and click Run AI Screening</p>
            </div>
          )}
          {screening && (
            <div style={{background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-lg)",padding:"2rem",display:"flex",flexDirection:"column",alignItems:"center",gap:14,flex:1}}>
              <Loader2 size={34} color={PU} style={{animation:"spin 1s linear infinite"}}/>
              <p style={{fontSize:14,fontWeight:500}}>Analysing 30 years of market data...</p>
            </div>
          )}
          {result && (
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <div style={{background:vst.bg,border:`1px solid ${vst.border}`,borderRadius:"var(--border-radius-lg)",padding:"1rem 1.25rem",display:"flex",alignItems:"center",gap:16}}>
                <div style={{textAlign:"center",flexShrink:0}}>
                  <div style={{fontSize:26,fontWeight:700,color:vst.color,lineHeight:1}}>{(result.overall_score||0).toFixed(1)}</div>
                  <div style={{fontSize:10,color:vst.color,fontWeight:500}}>/ 10</div>
                </div>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                    <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,background:vst.color,color:"#fff"}}>{result.verdict}</span>
                    <span style={{fontSize:13,fontWeight:500,color:vst.color}}>{result.verdict_reason}</span>
                  </div>
                  <p style={{fontSize:12,lineHeight:1.7,color:"var(--color-text-secondary)",margin:0}}>{result.recommendation}</p>
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <Box title="Score Breakdown">
                  <ResponsiveContainer width="100%" height={190}>
                    <RadarChart data={rd}>
                      <PolarGrid stroke="var(--color-border-tertiary)"/>
                      <PolarAngleAxis dataKey="subject" tick={{fontSize:10,fill:"var(--color-text-secondary)"}}/>
                      <PolarRadiusAxis angle={30} domain={[0,10]} tick={{fontSize:9}} tickCount={3}/>
                      <Radar name="Score" dataKey="A" stroke={PU} fill={PU} fillOpacity={0.3} strokeWidth={2}/>
                    </RadarChart>
                  </ResponsiveContainer>
                </Box>
                <Box title="Key Metrics">
                  {result.key_metrics && Object.entries(result.key_metrics).map(([k,v]) => (
                    <div key={k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                      <span style={{fontSize:11,color:"var(--color-text-secondary)",textTransform:"capitalize"}}>{k.replace(/_/g," ")}</span>
                      <span style={{fontSize:12,fontWeight:500}}>{v}</span>
                    </div>
                  ))}
                </Box>
              </div>
              <Box title="Dimension Analysis">
                <div style={{display:"flex",flexDirection:"column",gap:10}}>
                  {result.scores && Object.values(result.scores).map((s,i) => (
                    <div key={i}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                        <span style={{fontSize:12,fontWeight:500}}>{s.label}</span>
                        <span style={{fontSize:12,fontWeight:700,color:s.score>=7?G:s.score>=5?AM:RD}}>{s.score}/10</span>
                      </div>
                      <div style={{height:5,background:"var(--color-background-secondary)",borderRadius:10,overflow:"hidden",marginBottom:3}}>
                        <div style={{height:"100%",width:(s.score*10)+"%",background:s.score>=7?G:s.score>=5?AM:RD,borderRadius:10}}/>
                      </div>
                      <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:0,lineHeight:1.6}}>{s.comment}</p>
                    </div>
                  ))}
                </div>
              </Box>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div style={{background:"#fff5f5",border:"0.5px solid #fca5a5",borderRadius:"var(--border-radius-lg)",padding:"1rem"}}>
                  <p style={{fontWeight:500,fontSize:12,color:"#9b1c1c",margin:"0 0 8px"}}>Strengths</p>
                  {(result.pros||[]).map((p,i) => <p key={i} style={{fontSize:11,color:"#7f1d1d",margin:"3px 0",lineHeight:1.6}}>- {p}</p>)}
                </div>
                <div style={{background:"#fef2f2",border:"0.5px solid #fca5a5",borderRadius:"var(--border-radius-lg)",padding:"1rem"}}>
                  <p style={{fontWeight:500,fontSize:12,color:"#991b1b",margin:"0 0 8px"}}>Risks</p>
                  {(result.cons||[]).map((c,i) => <p key={i} style={{fontSize:11,color:"#991b1b",margin:"3px 0",lineHeight:1.6}}>- {c}</p>)}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
var DATA_TABS = ["ov","tx","res","hpi","com","ind","hot","dem"]

export default function Dashboard() {
  var [tab, setTab] = useState("ov")
  var [dist, setDist] = useState("all")
  var [yearRange, setYearRange] = useState({from:"2015", to:"2024"})
  var [propType, setPropType] = useState("all")
  var showFilter = DATA_TABS.includes(tab)
  function renderContent() {
    if (tab === "ov")  return <Overview dist={dist} yearRange={yearRange}/>
    if (tab === "tx")  return <Transactions dist={dist} yearRange={yearRange} propType={propType}/>
    if (tab === "res") return <Residential dist={dist}/>
    if (tab === "hpi") return <HPI/>
    if (tab === "com") return <Commercial/>
    if (tab === "ind") return <Industrial dist={dist}/>
    if (tab === "hot") return <HotelTab dist={dist}/>
    if (tab === "dem") return <Demographics/>
    if (tab === "fc")  return <ForecastingAgent/>
    if (tab === "scr") return <Screener/>
    if (tab === "rfr") return <DataRefresh/>
    if (tab === "ai")  return <AIAnalysis/>
    return <Overview dist={dist} yearRange={yearRange}/>
  }
  var cur = TABS.find(t => t.id === tab)
  var isAI = ["fc","scr","rfr","ai"].includes(tab)
  return (
    <div style={{display:"flex",minHeight:"100vh",background:"var(--color-background-tertiary)",fontFamily:"var(--font-sans)"}}>
      <div style={{width:218,flexShrink:0,background:"#1A0508",display:"flex",flexDirection:"column",minHeight:"100vh",position:"sticky",top:0}}>
        <div style={{padding:"20px 16px 14px",borderBottom:"1px solid rgba(255,255,255,0.08)"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:32,height:32,background:"#C0272D",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center"}}><BarChart3 size={18} color="#fff"/></div>
            <div>
              <p style={{color:"#fff",fontWeight:500,fontSize:13,margin:0,lineHeight:1.2}}>Gamuda Intelligence</p>
              <p style={{color:"rgba(255,255,255,0.5)",fontSize:10,margin:0}}>Property Intelligence</p>
            </div>
          </div>
        </div>
        <nav style={{flex:1,padding:"10px 8px",display:"flex",flexDirection:"column",gap:2}}>
          {TABS.map((t,i) => {
            if (t.id === "div") return <div key={i} style={{height:1,background:"rgba(255,255,255,0.08)",margin:"6px 4px"}}/>
            var Icon = t.icon, active = tab === t.id
            return (
              <button key={t.id} onClick={() => setTab(t.id)} style={{display:"flex",alignItems:"center",gap:9,padding:"8px 10px",borderRadius:8,cursor:"pointer",background:active?"rgba(192,39,45,0.25)":"transparent",border:`1px solid ${active?"#E05555":"transparent"}`,color:active?"#FF8080":"rgba(255,255,255,0.6)",width:"100%",textAlign:"left"}}>
                <Icon size={14}/>
                <span style={{fontSize:12,fontWeight:active?500:400,flex:1}}>{t.label}</span>
                {t.badge && <span style={{padding:"1px 5px",borderRadius:20,fontSize:9,fontWeight:600,background:"rgba(139,92,246,0.3)",color:"#c4b5fd"}}>{t.badge}</span>}
                {active && <ChevronRight size={11}/>}
              </button>
            )
          })}
        </nav>
        <div style={{padding:"10px 12px",borderTop:"1px solid rgba(255,255,255,0.08)"}}>
          <div style={{padding:"7px 10px",background:"rgba(255,255,255,0.04)",borderRadius:6}}>
            <p style={{fontSize:9,color:"rgba(255,255,255,0.35)",lineHeight:1.5,margin:0}}>Gamuda Property Intelligence{"\n"}Per-section PDF reports with AI{"\n"}Forecast · Screener · Data Refresh</p>
          </div>
        </div>
      </div>
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{background:"var(--color-background-primary)",borderBottom:"0.5px solid var(--color-border-tertiary)",padding:"13px 24px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <h1 style={{fontSize:15,fontWeight:500,margin:0,color:"var(--color-text-primary)"}}>{cur && cur.label}</h1>
              {isAI && <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:600,background:"rgba(139,92,246,0.1)",color:"#7c3aed",border:"1px solid rgba(139,92,246,0.2)"}}>AGENTIC AI</span>}
            </div>
            <p style={{fontSize:11,color:"var(--color-text-secondary)",margin:0,marginTop:1}}>Gamuda Property Market Intelligence · Q1 2025</p>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <SlidesButton/>
            <div style={{padding:"5px 12px",borderRadius:20,background:"#FEF2F2",fontSize:11,fontWeight:500,color:"#C0272D",display:"flex",alignItems:"center",gap:6}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:"#C0272D",display:"inline-block"}}/>Live Dashboard
            </div>
          </div>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"18px 24px 40px"}}>
          {showFilter && <DistrictFilter dist={dist} setDist={setDist} yearRange={yearRange} setYearRange={setYearRange} propType={propType} setPropType={setPropType}/>}
          {renderContent()}
        </div>
      </div>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--color-border-secondary);border-radius:10px}*{box-sizing:border-box}select,input{outline:none}`}</style>
    </div>
  )
}
