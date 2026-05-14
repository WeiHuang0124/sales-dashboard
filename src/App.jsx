import { useState, useEffect, useRef, useCallback } from "react"
import * as XLSX from "xlsx"
import {
  AreaChart, Area, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, ComposedChart, Line,
} from "recharts"
import storage from "./storage.js"

/* ─── Storage keys ─── */
const IDX_KEY  = "sales2:index"
const DAY_KEY  = (d) => `sales2:day:${d}`
const PROD_KEY = "sales2:products"

/* ─── Design tokens ─── */
const C = {
  bg:"#0b0d1a", panel:"#111326", card:"#161829",
  border:"rgba(255,255,255,0.06)",
  cyan:"#00e5ff", pink:"#ff3cac", purple:"#7c3aed",
  amber:"#f59e0b", text:"#e2e8f0", muted:"#64748b",
  ok:"#10b981", err:"#ef4444",
}
const fmt  = (n) => Math.round(n).toLocaleString("zh-TW")
const fmtK = (v) => v>=10000?`${(v/10000).toFixed(1)}萬`:v>=1000?`${(v/1000).toFixed(0)}k`:`${Math.round(v)}`
const today = new Date().toISOString().slice(0,10)

/* ──────────── Parse helpers ──────────── */
function parseShopeeText(text) {
  const lines = text.trim().split("\n").filter(l=>l.trim())
  if (lines.length<2) return null
  const sep = lines[0].includes("\t")?"\t":","
  const hdr = lines[0].split(sep).map(h=>h.trim().replace(/^["']|["']$/g,""))
  const di=hdr.findIndex(h=>h==="日期")
  if (di<0) return null

  // 支援兩種蝦皮格式：
  // 格式A：店鋪統計  → 總銷售額 (TWD)、訂單總數、訪客數
  // 格式B：商品概覽  → 銷售額(全部訂單) (TWD)、買家(全部訂單)、商品訪客數
  const ri = hdr.findIndex(h=>h.includes("總銷售額")||h==="銷售額(全部訂單) (TWD)")
  if (ri<0) return null

  const oi = hdr.findIndex(h=>h==="訂單總數"||h==="買家(全部訂單)")
  const vi = hdr.findIndex(h=>h==="訪客數"||h==="商品訪客數")
  const ci = hdr.findIndex(h=>h.includes("訂單轉換率")||h==="轉換率(全部訂單)")
  const ai = hdr.findIndex(h=>h.includes("廣告費用")||h.includes("廣告花費")||h.includes("廣告支出"))

  const rows=[]
  for (let i=1;i<lines.length;i++) {
    const p=lines[i].split(sep).map(x=>x.trim().replace(/^["']|["']$/g,""))
    const date=p[di]
    if (!date||!date.match(/^\d{4}-\d{2}-\d{2}$/)) continue
    rows.push({
      date,
      revenue: parseFloat(p[ri]?.replace(/,/g,""))||0,
      orders:  oi>=0?(parseInt(p[oi]?.replace(/,/g,""))||0):0,
      visitors:vi>=0?(parseInt(p[vi]?.replace(/,/g,""))||0):0,
      conv:    ci>=0?(p[ci]||"0%"):"0%",
      adSpend: ai>=0?(parseFloat(p[ai]?.replace(/,/g,""))||0):0,
    })
  }
  return rows.length>0?{type:"shopee",rows}:null
}

function parseProductText(text) {
  const lines = text.trim().split("\n").filter(l=>l.trim())
  if (lines.length<2) return null
  const sep=lines[0].includes("\t")?"\t":","
  const hdr=lines[0].split(sep).map(h=>h.trim().replace(/^["']|["']$/g,""))

  // 拒絕第一欄是「日期」的檔案（那是蝦皮每日報表，不是商品明細）
  if (hdr[0]==="日期") return null

  // 必須要有能辨識為「商品名稱」的欄位
  const nameIdx = hdr.findIndex(h=>h.includes("商品名稱")||h.includes("商品")||h.includes("品名")||h.includes("名稱"))
  const revenueIdx = hdr.findIndex(h=>h.includes("銷售額")||h.includes("金額")||h.includes("收入")||h.includes("營業額"))
  const qtyIdx = hdr.findIndex(h=>h.includes("數量")||h.includes("件數")||h.includes("銷量"))

  // 如果有明確的商品名稱欄，用欄位對應；否則用位置對應（舊格式：名稱,數量,金額）
  const useMapping = nameIdx>=0

  const rows=[]
  for (let i=1;i<lines.length;i++) {
    const p=lines[i].split(sep).map(x=>x.trim().replace(/^["']|["']$/g,""))
    if (!p[0]) continue

    let product, qty=0, revenue=0
    if (useMapping) {
      product = p[nameIdx]||""
      qty     = qtyIdx>=0?(parseInt(p[qtyIdx]?.replace(/,/g,""))||0):0
      revenue = revenueIdx>=0?(parseFloat(p[revenueIdx]?.replace(/,/g,""))||0):0
    } else {
      // 舊格式：商品名稱, [數量,] 金額
      const clean = p.map(x=>x.replace(/,/g,""))
      product = clean[0]
      if (clean.length===2) revenue=parseFloat(clean[1])||0
      else if (clean.length>=3){qty=parseInt(clean[1])||0;revenue=parseFloat(clean[2])||0}
    }
    if (product&&revenue>0) rows.push({product,qty,revenue})
  }
  return rows.length>0?{type:"product",rows}:null
}

async function parseFile(file) {
  return new Promise((resolve)=>{
    const reader=new FileReader()
    const isXlsx=/\.(xlsx|xls)$/i.test(file.name)
    reader.onload=(e)=>{
      try {
        if (isXlsx) {
          const wb=XLSX.read(e.target.result,{type:"array"})
          for (const sn of wb.SheetNames) {
            const tsv=XLSX.utils.sheet_to_csv(wb.Sheets[sn],{FS:"\t",blankrows:false})
            const s=parseShopeeText(tsv)
            const p=s?null:parseProductText(tsv)
            if (s||p){resolve(s||p);return}
          }
          resolve(null)
        } else {
          const t=e.target.result
          const s=parseShopeeText(t)
          resolve(s||(parseProductText(t))||null)
        }
      } catch{resolve(null)}
    }
    reader.onerror=()=>resolve(null)
    if (isXlsx) reader.readAsArrayBuffer(file)
    else reader.readAsText(file)
  })
}

/* ──────────── AI Insights ──────────── */
async function callClaudeInsights(dataPayload) {
  const prompt = `你是一位專業的電商經營顧問，專精於蝦皮平台的數據分析。
請根據以下店鋪數據，提供具體可執行的經營建議。

${JSON.stringify(dataPayload, null, 2)}

請用繁體中文，以JSON格式回覆，結構如下：
{
  "summary": "整體店鋪狀況的一句話摘要（20字以內）",
  "health_score": 整體健康分數1-100的數字,
  "insights": [
    {
      "category": "類別名稱",
      "icon": "emoji圖示",
      "priority": "high" | "medium" | "low",
      "title": "建議標題（15字以內）",
      "analysis": "數據分析說明（60字以內）",
      "actions": ["具體行動1", "具體行動2", "具體行動3"]
    }
  ]
}

建議類別必須涵蓋（如有數據）：
- 廣告投放效益（ROAS分析、預算調整）
- 商品結構優化（主力商品、長尾商品）
- 流量與轉換提升（訪客、轉換率）
- 營業額成長策略（趨勢、季節性）
- 定價與促銷建議

只回傳JSON，不要有其他文字。`

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1500,
      messages: [{ role:"user", content: prompt }]
    })
  })
  const data = await res.json()
  const text = data.content?.find(b=>b.type==="text")?.text || ""
  const clean = text.replace(/```json|```/g,"").trim()
  return JSON.parse(clean)
}

/* ──────────── UI helpers ──────────── */
function GradientCard({label,value,sub,from,to}) {
  return (
    <div style={{background:`linear-gradient(135deg,${from},${to})`,borderRadius:16,padding:"20px 22px",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-20,right:-20,width:80,height:80,borderRadius:"50%",background:"rgba(255,255,255,0.08)"}}/>
      <div style={{fontSize:11,color:"rgba(255,255,255,0.7)",letterSpacing:"0.07em",textTransform:"uppercase",marginBottom:10}}>{label}</div>
      <div style={{fontSize:30,fontWeight:700,color:"#fff",letterSpacing:"-0.02em"}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:"rgba(255,255,255,0.6)",marginTop:5}}>{sub}</div>}
    </div>
  )
}

function StatCard({label,value,icon,color}) {
  return (
    <div style={{background:C.card,borderRadius:14,padding:"16px 18px",border:`1px solid ${C.border}`,display:"flex",alignItems:"center",gap:14}}>
      <div style={{width:42,height:42,borderRadius:12,background:`${color}20`,display:"flex",alignItems:"center",justifyContent:"center"}}>
        <i className={`ti ${icon}`} style={{fontSize:20,color}}/>
      </div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:11,color:C.muted,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:4}}>{label}</div>
        <div style={{fontSize:20,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{value}</div>
      </div>
    </div>
  )
}

const CustomTooltip = ({active,payload,label}) => {
  if (!active||!payload?.length) return null
  return (
    <div style={{background:"#1e2130",border:"1px solid rgba(255,255,255,0.08)",borderRadius:10,fontSize:12,color:C.text,padding:"10px 14px"}}>
      <div style={{color:C.muted,fontSize:11,marginBottom:6}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}>
          <span style={{width:8,height:8,borderRadius:2,background:p.color||p.stroke,display:"inline-block"}}/>
          <span style={{color:C.muted,fontSize:11}}>{p.name==="revenue"?"營業額":p.name==="orders"?"訂單":p.name==="visitors"?"訪客":p.name}</span>
          <span style={{fontWeight:600,color:C.text,marginLeft:"auto",paddingLeft:16}}>{p.name==="revenue"?`$${fmt(p.value)}`:p.value}</span>
        </div>
      ))}
    </div>
  )
}

function PanelBox({title,children}) {
  return (
    <div style={{background:C.card,borderRadius:16,border:`1px solid ${C.border}`,padding:"20px 22px"}}>
      {title&&<div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:18}}>{title}</div>}
      {children}
    </div>
  )
}

function DropZone({onFile,processing}) {
  const [hover,setHover]=useState(false)
  const ref=useRef()
  const onDrop=useCallback((e)=>{e.preventDefault();setHover(false);const f=e.dataTransfer.files?.[0];if(f) onFile(f)},[onFile])
  return (
    <div onDragOver={e=>{e.preventDefault();setHover(true)}} onDragLeave={()=>setHover(false)} onDrop={onDrop} onClick={()=>ref.current?.click()}
      style={{border:`2px dashed ${hover?C.cyan:`${C.purple}60`}`,borderRadius:16,padding:"40px 24px",textAlign:"center",cursor:"pointer",transition:"all 0.2s",background:hover?`${C.cyan}08`:`${C.purple}08`}}>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv,.tsv,.txt" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f)onFile(f);e.target.value=""}}/>
      {processing?(
        <><div style={{fontSize:36,marginBottom:12,color:C.cyan}}><i className="ti ti-loader" style={{animation:"spin 1s linear infinite",display:"inline-block"}}/></div><div style={{fontSize:14,fontWeight:600,color:C.cyan}}>解析中...</div></>
      ):(
        <>
          <div style={{width:60,height:60,borderRadius:18,background:hover?`${C.cyan}20`:`${C.purple}20`,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 18px"}}>
            <i className="ti ti-cloud-upload" style={{fontSize:30,color:hover?C.cyan:C.purple}}/>
          </div>
          <div style={{fontSize:16,fontWeight:600,color:C.text,marginBottom:8}}>{hover?"放開以上傳":"拖曳檔案到這裡"}</div>
          <div style={{fontSize:13,color:C.muted,marginBottom:14}}>或點擊選擇檔案</div>
          <div style={{display:"flex",gap:8,justifyContent:"center"}}>
            {[".xlsx",".csv",".tsv"].map(e=><span key={e} style={{fontSize:11,padding:"3px 12px",borderRadius:20,background:`${C.purple}20`,color:C.purple,border:`1px solid ${C.purple}40`}}>{e}</span>)}
          </div>
        </>
      )}
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  )
}

function GlobalDragOverlay({visible}) {
  if (!visible) return null
  return (
    <div style={{position:"fixed",inset:0,zIndex:9999,background:"rgba(11,13,26,0.9)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",border:`2px solid ${C.cyan}`,pointerEvents:"none"}}>
      <i className="ti ti-cloud-upload" style={{fontSize:64,color:C.cyan,marginBottom:20}}/>
      <div style={{fontSize:22,fontWeight:700,color:C.text}}>放開以匯入報表</div>
      <div style={{fontSize:14,color:C.muted,marginTop:10}}>支援 .xlsx · .csv · .tsv</div>
    </div>
  )
}

/* ──────────── Insights Panel ──────────── */
const PRIORITY_CONFIG = {
  high:   { label:"優先處理", bg:"rgba(239,68,68,0.12)",   border:"rgba(239,68,68,0.3)",   color:"#ef4444" },
  medium: { label:"建議執行", bg:"rgba(245,158,11,0.12)",  border:"rgba(245,158,11,0.3)",  color:"#f59e0b" },
  low:    { label:"長期優化", bg:"rgba(16,185,129,0.12)",  border:"rgba(16,185,129,0.3)",  color:"#10b981" },
}

function InsightCard({insight, index}) {
  const [open, setOpen] = useState(index===0)
  const cfg = PRIORITY_CONFIG[insight.priority] || PRIORITY_CONFIG.medium
  return (
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,overflow:"hidden",marginBottom:10}}>
      <button onClick={()=>setOpen(o=>!o)} style={{
        width:"100%",display:"flex",alignItems:"center",gap:12,padding:"14px 18px",
        background:"none",border:"none",cursor:"pointer",textAlign:"left",
      }}>
        <span style={{fontSize:22,flexShrink:0}}>{insight.icon}</span>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text}}>{insight.title}</div>
          <div style={{fontSize:11,color:C.muted,marginTop:2}}>{insight.category}</div>
        </div>
        <span style={{fontSize:11,fontWeight:600,padding:"3px 10px",borderRadius:20,background:cfg.bg,border:`1px solid ${cfg.border}`,color:cfg.color,flexShrink:0}}>
          {cfg.label}
        </span>
        <i className={`ti ${open?"ti-chevron-up":"ti-chevron-down"}`} style={{fontSize:14,color:C.muted,marginLeft:4}}/>
      </button>
      {open&&(
        <div style={{padding:"0 18px 16px",borderTop:`1px solid ${C.border}`}}>
          <p style={{fontSize:13,color:"#94a3b8",lineHeight:1.7,margin:"12px 0 14px"}}>{insight.analysis}</p>
          <div style={{fontSize:11,color:C.muted,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:8}}>具體行動</div>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {insight.actions.map((a,i)=>(
              <div key={i} style={{display:"flex",alignItems:"flex-start",gap:10,background:`${cfg.color}08`,border:`1px solid ${cfg.color}20`,borderRadius:9,padding:"9px 12px"}}>
                <span style={{fontSize:13,fontWeight:700,color:cfg.color,flexShrink:0,marginTop:1}}>{i+1}</span>
                <span style={{fontSize:13,color:C.text,lineHeight:1.6}}>{a}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}

function HealthMeter({score}) {
  const color = score>=75?C.ok:score>=50?C.amber:C.err
  const sectors = [
    {label:"危險",color:C.err,range:[0,40]},
    {label:"待優化",color:C.amber,range:[40,70]},
    {label:"健康",color:C.ok,range:[70,100]},
  ]
  return (
    <div style={{textAlign:"center",padding:"8px 0"}}>
      <div style={{position:"relative",width:120,height:60,margin:"0 auto 12px"}}>
        <svg viewBox="0 0 120 65" width="120" height="65">
          <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke="rgba(255,255,255,0.06)" strokeWidth="10" strokeLinecap="round"/>
          <path d="M10,60 A50,50 0 0,1 110,60" fill="none" stroke={color} strokeWidth="10" strokeLinecap="round"
            strokeDasharray={`${(score/100)*157} 157`} style={{transition:"stroke-dasharray 1s ease"}}/>
          <text x="60" y="56" textAnchor="middle" fill={color} fontSize="22" fontWeight="700">{score}</text>
        </svg>
      </div>
      <div style={{fontSize:11,color:C.muted}}>店鋪健康分數</div>
    </div>
  )
}

function InsightsView({dayData,prodData}) {
  const [result,setResult]=useState(null)
  const [loading,setLoading]=useState(false)
  const [err,setErr]=useState(null)

  const dates=Object.keys(dayData).sort()
  const hasData = dates.length>0

  function buildPayload() {
    const totalRev=dates.reduce((s,d)=>s+dayData[d].revenue,0)
    const totalOrds=dates.reduce((s,d)=>s+dayData[d].orders,0)
    const totalVis=dates.reduce((s,d)=>s+dayData[d].visitors,0)
    const totalAds=dates.reduce((s,d)=>s+(dayData[d].adSpend||0),0)
    const activeDays=dates.filter(d=>dayData[d].revenue>0).length
    const convRate=totalVis>0?((totalOrds/totalVis)*100).toFixed(2):0
    const roas=totalAds>0?(totalRev/totalAds).toFixed(2):"無廣告數據"
    const avgOrderVal=totalOrds>0?(totalRev/totalOrds).toFixed(0):0
    const monthlyMap={}
    dates.forEach(d=>{const m=d.slice(0,7);if(!monthlyMap[m])monthlyMap[m]={rev:0,ord:0};monthlyMap[m].rev+=dayData[d].revenue;monthlyMap[m].ord+=dayData[d].orders})
    const monthly=Object.entries(monthlyMap).map(([m,v])=>({month:m,revenue:v.rev,orders:v.ord}))
    const recentDays=dates.slice(-7)
    const recentRev=recentDays.reduce((s,d)=>s+dayData[d].revenue,0)
    const prevDays=dates.slice(-14,-7)
    const prevRev=prevDays.reduce((s,d)=>s+dayData[d].revenue,0)
    const trend=prevRev>0?((recentRev-prevRev)/prevRev*100).toFixed(1):null
    const dowMap={0:0,1:0,2:0,3:0,4:0,5:0,6:0}
    const dowCount={0:0,1:0,2:0,3:0,4:0,5:0,6:0}
    dates.forEach(d=>{const dow=new Date(d).getDay();dowMap[dow]+=dayData[d].revenue;dowCount[dow]++})
    const dowNames=["日","一","二","三","四","五","六"]
    const dowAvg=Object.entries(dowMap).map(([k,v])=>({day:`週${dowNames[k]}`,avg:dowCount[k]>0?Math.round(v/dowCount[k]):0})).sort((a,b)=>b.avg-a.avg)
    return {
      分析期間:`${dates[0]} 至 ${dates[dates.length-1]}（共${dates.length}天）`,
      有銷售天數:`${activeDays}天（佔比${Math.round(activeDays/dates.length*100)}%）`,
      總營業額:`NT$${fmt(totalRev)}`,
      總訂單數:`${totalOrds}筆`,
      總訪客數:`${totalVis}人`,
      整體轉換率:`${convRate}%`,
      平均客單價:`NT$${fmt(avgOrderVal)}`,
      廣告總花費:totalAds>0?`NT$${fmt(totalAds)}`:"無數據",
      廣告投報比:roas,
      近7天趨勢:trend?`${trend>0?"+":""}${trend}%`:"資料不足",
      每月數據:monthly,
      星期表現:dowAvg.slice(0,3).map(d=>`${d.day}平均NT$${fmt(d.avg)}`),
      商品排名前5:prodData.slice(0,5).map(p=>({商品:p.product,總營業額:`NT$${fmt(p.revenue)}`,銷售天數:p.days})),
    }
  }

  async function run() {
    setLoading(true); setErr(null); setResult(null)
    try {
      const payload=buildPayload()
      const res=await callClaudeInsights(payload)
      setResult(res)
    } catch(e) {
      setErr("分析失敗，請稍後再試")
      console.error(e)
    }
    setLoading(false)
  }

  if (!hasData) return (
    <PanelBox title="AI 經營建議">
      <div style={{textAlign:"center",padding:"3rem",color:C.muted}}>
        <i className="ti ti-brain" style={{fontSize:40,display:"block",marginBottom:12,color:`${C.purple}80`}}/>
        <div style={{fontSize:14,marginBottom:8}}>尚未匯入資料</div>
        <div style={{fontSize:12}}>請先匯入蝦皮報表，AI 才能分析你的數據</div>
      </div>
    </PanelBox>
  )

  return (
    <div>
      {/* Header */}
      <div style={{background:`linear-gradient(135deg,${C.purple}30,${C.cyan}15)`,border:`1px solid ${C.purple}40`,borderRadius:16,padding:"20px 22px",marginBottom:16,display:"flex",alignItems:"center",gap:16}}>
        <div style={{width:48,height:48,borderRadius:14,background:`linear-gradient(135deg,${C.purple},${C.cyan})`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
          <i className="ti ti-brain" style={{fontSize:24,color:"#fff"}}/>
        </div>
        <div style={{flex:1}}>
          <div style={{fontSize:15,fontWeight:700,color:C.text}}>AI 經營建議</div>
          <div style={{fontSize:12,color:C.muted,marginTop:3}}>基於 {Object.keys(dayData).length} 天銷售數據，生成個人化建議</div>
        </div>
        <button onClick={run} disabled={loading} style={{
          background:loading?`${C.purple}50`:`linear-gradient(135deg,${C.purple},${C.cyan})`,
          border:"none",borderRadius:11,padding:"10px 20px",color:"#fff",cursor:loading?"wait":"pointer",
          fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:7,flexShrink:0,
          transition:"opacity 0.2s",
        }}>
          {loading
            ?<><i className="ti ti-loader" style={{fontSize:15,animation:"spin 1s linear infinite"}}/>分析中...</>
            :<><i className="ti ti-sparkles" style={{fontSize:15}}/>{result?"重新分析":"開始分析"}</>
          }
        </button>
      </div>

      {err && <div style={{background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",borderRadius:10,padding:"12px 16px",color:C.err,fontSize:13,marginBottom:16}}>{err}</div>}

      {loading && (
        <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:"40px",textAlign:"center"}}>
          <div style={{marginBottom:20}}>
            {["分析銷售趨勢...","計算廣告效益...","評估商品表現...","生成個人化建議..."].map((t,i)=>(
              <div key={i} style={{fontSize:13,color:i===0?C.cyan:C.muted,marginBottom:8,transition:"color 0.5s",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                <i className={`ti ${i===0?"ti-circle-check":"ti-circle"}`} style={{fontSize:14}}/>
                {t}
              </div>
            ))}
          </div>
          <div style={{fontSize:12,color:C.muted}}>AI 正在深度分析你的數據，約需 10 秒...</div>
        </div>
      )}

      {result && !loading && (
        <div>
          {/* Score + summary */}
          <div style={{display:"grid",gridTemplateColumns:"140px 1fr",gap:12,marginBottom:16}}>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"16px 12px"}}>
              <HealthMeter score={result.health_score||70}/>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"18px 20px",display:"flex",alignItems:"center"}}>
              <div>
                <div style={{fontSize:11,color:C.muted,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:8}}>整體診斷</div>
                <div style={{fontSize:16,fontWeight:600,color:C.text,lineHeight:1.6}}>{result.summary}</div>
                <div style={{fontSize:12,color:C.muted,marginTop:8}}>共 {result.insights?.length||0} 項建議 · 點各項展開查看</div>
              </div>
            </div>
          </div>

          {/* Insight cards */}
          {(result.insights||[]).sort((a,b)=>{const o={high:0,medium:1,low:2};return o[a.priority]-o[b.priority]}).map((ins,i)=>(
            <InsightCard key={i} insight={ins} index={i}/>
          ))}

          <div style={{textAlign:"center",marginTop:16,fontSize:11,color:C.muted}}>
            <i className="ti ti-info-circle" style={{marginRight:5}}/>
            建議基於現有數據生成，執行前請結合實際情況判斷
          </div>
        </div>
      )}

      {!result && !loading && !err && (
        <div style={{background:C.card,border:`1px dashed ${C.border}`,borderRadius:16,padding:"40px",textAlign:"center",color:C.muted}}>
          <div style={{fontSize:36,marginBottom:12}}>🤖</div>
          <div style={{fontSize:14,marginBottom:8,color:C.text}}>點「開始分析」讓 AI 診斷你的店鋪</div>
          <div style={{fontSize:12}}>廣告效益 · 商品策略 · 流量優化 · 成長建議</div>
        </div>
      )}
    </div>
  )
}

/* ──────────── Main App ──────────── */
export default function App() {
  const [dayData,   setDayData]   = useState({})
  const [prodData,  setProdData]  = useState([])
  const [view,      setView]      = useState("dashboard")
  const [importDate,setImportDate]= useState(today)
  const [importText,setImportText]= useState("")
  const [msg,       setMsg]       = useState(null)
  const [loading,   setLoading]   = useState(true)
  const [deleting,  setDeleting]  = useState(null)
  const [processing,setProcessing]= useState(false)
  const [dragging,  setDragging]  = useState(false)
  const dragCounter = useRef(0)

  useEffect(()=>{loadAll()},[])

  useEffect(()=>{
    const onEnter=(e)=>{e.preventDefault();dragCounter.current++;if(e.dataTransfer.items?.[0]?.kind==="file")setDragging(true)}
    const onLeave=()=>{dragCounter.current--;if(dragCounter.current===0)setDragging(false)}
    const onOver=(e)=>e.preventDefault()
    const onDrop=async(e)=>{e.preventDefault();dragCounter.current=0;setDragging(false);const f=e.dataTransfer.files?.[0];if(f)await processFile(f)}
    window.addEventListener("dragenter",onEnter)
    window.addEventListener("dragleave",onLeave)
    window.addEventListener("dragover",onOver)
    window.addEventListener("drop",onDrop)
    return()=>{window.removeEventListener("dragenter",onEnter);window.removeEventListener("dragleave",onLeave);window.removeEventListener("dragover",onOver);window.removeEventListener("drop",onDrop)}
  },[])

  async function loadAll() {
    try {
      const idx=await storage.get(IDX_KEY)
      if (idx) {
        const dates=JSON.parse(idx.value)
        const data={}
        for (const d of dates){try{const r=await storage.get(DAY_KEY(d));if(r)data[d]=JSON.parse(r.value)}catch{}}
        setDayData(data)
      }
      try{const pr=await storage.get(PROD_KEY);if(pr)setProdData(JSON.parse(pr.value))}catch{}
    }catch{}
    setLoading(false)
  }

  async function saveDays(rows) {
    let dates=[]
    try{const idx=await storage.get(IDX_KEY);if(idx)dates=JSON.parse(idx.value)}catch{}
    const nd={...dayData}
    for (const r of rows){
      const val={revenue:r.revenue,orders:r.orders,visitors:r.visitors,conv:r.conv,adSpend:r.adSpend||0}
      await storage.set(DAY_KEY(r.date),JSON.stringify(val))
      nd[r.date]=val
      if(!dates.includes(r.date))dates.push(r.date)
    }
    dates.sort()
    await storage.set(IDX_KEY,JSON.stringify(dates))
    setDayData(nd)
    return rows.length
  }

  async function saveProducts(date,rows) {
    const pm={}
    prodData.forEach(p=>{pm[p.product]={...p}})
    rows.forEach(({product,qty,revenue})=>{
      if(!pm[product])pm[product]={product,qty:0,revenue:0,days:0}
      pm[product].qty+=qty;pm[product].revenue+=revenue;pm[product].days+=1
    })
    const sorted=Object.values(pm).sort((a,b)=>b.revenue-a.revenue)
    await storage.set(PROD_KEY,JSON.stringify(sorted))
    setProdData(sorted)
    return rows.length
  }

  async function deleteDay(date) {
    setDeleting(date)
    try{
      await storage.delete(DAY_KEY(date))
      let dates=[]
      try{const idx=await storage.get(IDX_KEY);if(idx)dates=JSON.parse(idx.value)}catch{}
      dates=dates.filter(d=>d!==date)
      await storage.set(IDX_KEY,JSON.stringify(dates))
      setDayData(prev=>{const n={...prev};delete n[date];return n})
    }catch{}
    setDeleting(null)
  }

  function showMsg(text,ok){setMsg({text,ok});setTimeout(()=>setMsg(null),6000)}

  async function processFile(file) {
    setProcessing(true)
    const result=await parseFile(file)
    if(!result){showMsg(`無法解析「${file.name}」`,false)}
    else if(result.type==="shopee"){const n=await saveDays(result.rows);showMsg(`✓ 匯入 ${n} 天資料`,true);setView("dashboard")}
    else{const n=await saveProducts(importDate,result.rows);showMsg(`✓ 匯入 ${n} 項商品`,true);setView("ranking")}
    setProcessing(false)
  }

  async function handleTextImport() {
    if(!importText.trim()){showMsg("請貼上報表資料",false);return}
    setProcessing(true)
    const s=parseShopeeText(importText)
    const p=!s&&parseProductText(importText)
    if(s){const n=await saveDays(s.rows);showMsg(`✓ 已匯入 ${n} 天`,true)}
    else if(p){const n=await saveProducts(importDate,p.rows);showMsg(`✓ 已儲存 ${n} 項商品`,true)}
    else showMsg("無法辨識格式",false)
    setImportText("");setProcessing(false)
  }

  const dates=Object.keys(dayData).sort()
  const dailyChart=dates.map(d=>({date:d.slice(5),fullDate:d,revenue:dayData[d].revenue,orders:dayData[d].orders,visitors:dayData[d].visitors}))
  const monthlyMap={}
  dates.forEach(d=>{const m=d.slice(0,7);if(!monthlyMap[m])monthlyMap[m]={revenue:0,orders:0,visitors:0};monthlyMap[m].revenue+=dayData[d].revenue;monthlyMap[m].orders+=dayData[d].orders;monthlyMap[m].visitors+=dayData[d].visitors})
  const monthlyChart=Object.entries(monthlyMap).sort().map(([m,v])=>({month:`${parseInt(m.slice(5))}月`,...v}))
  const todayRev=dayData[today]?.revenue??0
  const curMonth=today.slice(0,7)
  const monthRev=dates.filter(d=>d.startsWith(curMonth)).reduce((s,d)=>s+dayData[d].revenue,0)
  const totalOrds=dates.reduce((s,d)=>s+dayData[d].orders,0)
  const totalVis=dates.reduce((s,d)=>s+dayData[d].visitors,0)
  const convRate=totalVis>0?((totalOrds/totalVis)*100).toFixed(1)+"%":"—"

  const NAV=[
    ["dashboard","ti-layout-dashboard","儀表板"],
    ["ranking","ti-trophy","商品排名"],
    ["insights","ti-brain","經營建議"],
    ["import","ti-file-upload","匯入報表"],
  ]

  if(loading) return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"50vh",color:C.muted,fontSize:14}}><i className="ti ti-loader" style={{fontSize:20,marginRight:8}}/>載入中...</div>

  return (
    <>
      <GlobalDragOverlay visible={dragging}/>
      <div style={{background:C.bg,borderRadius:20,display:"flex",overflow:"hidden",minHeight:650,border:`1px solid ${C.border}`}}>

        {/* Sidebar */}
        <div style={{width:190,background:C.panel,flexShrink:0,display:"flex",flexDirection:"column",borderRight:`1px solid ${C.border}`}}>
          <div style={{padding:"22px 18px 18px",borderBottom:`1px solid ${C.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:34,height:34,borderRadius:10,background:`linear-gradient(135deg,${C.purple},${C.cyan})`,display:"flex",alignItems:"center",justifyContent:"center"}}>
                <i className="ti ti-chart-bar" style={{fontSize:18,color:"#fff"}}/>
              </div>
              <div><div style={{fontSize:13,fontWeight:700,color:C.text}}>銷售總覽</div><div style={{fontSize:10,color:C.muted}}>蝦皮後台</div></div>
            </div>
          </div>
          <div style={{padding:"14px 10px",flex:1}}>
            <div style={{fontSize:10,color:C.muted,letterSpacing:"0.08em",textTransform:"uppercase",padding:"0 8px",marginBottom:8}}>主選單</div>
            {NAV.map(([v,icon,label])=>{
              const active=view===v
              return (
                <button key={v} onClick={()=>setView(v)} style={{
                  width:"100%",display:"flex",alignItems:"center",gap:10,padding:"9px 12px",
                  borderRadius:10,border:"none",cursor:"pointer",textAlign:"left",marginBottom:2,
                  background:active?`linear-gradient(135deg,${C.purple}40,${C.cyan}20)`:"transparent",
                  color:active?C.cyan:C.muted,
                  boxShadow:active?`inset 0 0 0 1px ${C.cyan}30`:"none",transition:"all 0.15s",
                }}>
                  <i className={`ti ${icon}`} style={{fontSize:16}}/>
                  <span style={{fontSize:12,fontWeight:active?600:400}}>{label}</span>
                  {active&&<div style={{marginLeft:"auto",width:5,height:5,borderRadius:"50%",background:C.cyan}}/>}
                  {v==="insights"&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:4,background:`${C.purple}40`,color:C.cyan,marginLeft:v==="insights"&&!active?4:0}}>AI</span>}
                </button>
              )
            })}
          </div>
          <div style={{padding:"14px 16px",borderTop:`1px solid ${C.border}`}}>
            <div style={{background:`${C.purple}15`,border:`1px dashed ${C.purple}40`,borderRadius:10,padding:"10px 12px",textAlign:"center",marginBottom:10}}>
              <i className="ti ti-drag-drop" style={{fontSize:18,color:C.purple,display:"block",marginBottom:4}}/>
              <div style={{fontSize:10,color:C.muted,lineHeight:1.6}}>拖曳 .xlsx / .csv<br/>到任意位置匯入</div>
            </div>
            <div style={{fontSize:11,color:C.muted}}>已儲存 <span style={{color:C.text,fontWeight:600}}>{dates.length}</span> 天資料</div>
          </div>
        </div>

        {/* Content */}
        <div style={{flex:1,padding:"24px",overflowY:"auto",maxHeight:650}}>
          {msg&&(
            <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px",borderRadius:12,marginBottom:16,fontSize:13,fontWeight:500,background:msg.ok?"rgba(16,185,129,0.12)":"rgba(239,68,68,0.12)",border:`1px solid ${msg.ok?"rgba(16,185,129,0.3)":"rgba(239,68,68,0.3)"}`,color:msg.ok?C.ok:C.err}}>
              <i className={`ti ${msg.ok?"ti-circle-check":"ti-alert-circle"}`} style={{fontSize:16}}/>
              {msg.text}
              <button onClick={()=>setMsg(null)} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",color:"inherit",padding:0,fontSize:14}}>✕</button>
            </div>
          )}

          {/* 儀表板 */}
          {view==="dashboard"&&(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
                <GradientCard label="本月營業額" value={`$${fmt(monthRev)}`} sub={`${curMonth.slice(5)}月累計`} from="#7c3aed" to="#3b82f6"/>
                <GradientCard label="今日營業額" value={`$${fmt(todayRev)}`} sub={today} from="#0ea5e9" to="#06b6d4"/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:10,marginBottom:18}}>
                <StatCard label="訂單總數" value={fmt(totalOrds)} icon="ti-shopping-bag" color={C.cyan}/>
                <StatCard label="訪客總數" value={fmt(totalVis)} icon="ti-users" color={C.pink}/>
                <StatCard label="整體轉換率" value={convRate} icon="ti-arrows-exchange" color={C.purple}/>
              </div>
              {dates.length===0?(
                <div style={{textAlign:"center",padding:"3rem",color:C.muted,background:C.card,borderRadius:16,border:`1px solid ${C.border}`}}>
                  <i className="ti ti-database-off" style={{fontSize:40,display:"block",marginBottom:12,color:`${C.purple}90`}}/>
                  <div style={{fontSize:14,marginBottom:8}}>尚未匯入任何報表</div>
                  <div style={{fontSize:12,marginBottom:20}}>把蝦皮的 .xlsx 直接拖進視窗</div>
                  <button onClick={()=>setView("import")} style={{background:`linear-gradient(135deg,${C.purple},${C.cyan})`,border:"none",borderRadius:10,padding:"9px 22px",color:"#fff",cursor:"pointer",fontWeight:600,fontSize:13}}>前往匯入 →</button>
                </div>
              ):(
                <div style={{display:"flex",flexDirection:"column",gap:12}}>
                  <PanelBox title="每日營業額趨勢">
                    <div style={{height:200}}>
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={dailyChart} margin={{top:4,right:4,left:-10,bottom:0}}>
                          <defs><linearGradient id="gRev" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.cyan} stopOpacity={0.3}/><stop offset="95%" stopColor={C.cyan} stopOpacity={0}/></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false}/>
                          <XAxis dataKey="date" tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false} interval={Math.max(0,Math.floor(dailyChart.length/12)-1)}/>
                          <YAxis tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false} tickFormatter={fmtK} width={36}/>
                          <Tooltip content={<CustomTooltip/>}/>
                          <Area type="monotone" dataKey="revenue" stroke={C.cyan} strokeWidth={2} fill="url(#gRev)" dot={false} activeDot={{r:4,fill:C.cyan,strokeWidth:0}}/>
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </PanelBox>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                    {monthlyChart.length>0&&(
                      <PanelBox title="每月彙總">
                        <div style={{height:160}}>
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={monthlyChart} margin={{top:4,right:4,left:-14,bottom:0}}>
                              <defs>{monthlyChart.map((_,i)=><linearGradient key={i} id={`bg${i}`} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.purple} stopOpacity={1}/><stop offset="100%" stopColor={C.cyan} stopOpacity={0.7}/></linearGradient>)}</defs>
                              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false}/>
                              <XAxis dataKey="month" tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false}/>
                              <YAxis tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false} tickFormatter={fmtK} width={34}/>
                              <Tooltip content={<CustomTooltip/>}/>
                              <Bar dataKey="revenue" radius={[6,6,0,0]}>{monthlyChart.map((_,i)=><Cell key={i} fill={`url(#bg${i})`}/>)}</Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </PanelBox>
                    )}
                    <PanelBox title="每日流量 & 訂單">
                      <div style={{height:160}}>
                        <ResponsiveContainer width="100%" height="100%">
                          <ComposedChart data={dailyChart} margin={{top:4,right:4,left:-14,bottom:0}}>
                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false}/>
                            <XAxis dataKey="date" tick={{fontSize:9,fill:C.muted}} axisLine={false} tickLine={false} interval={Math.max(0,Math.floor(dailyChart.length/8)-1)}/>
                            <YAxis tick={{fontSize:9,fill:C.muted}} axisLine={false} tickLine={false} width={28}/>
                            <Tooltip content={<CustomTooltip/>}/>
                            <Bar dataKey="visitors" fill={`${C.purple}55`} radius={[4,4,0,0]}/>
                            <Line type="monotone" dataKey="orders" stroke={C.pink} strokeWidth={2} dot={false} activeDot={{r:3}}/>
                          </ComposedChart>
                        </ResponsiveContainer>
                      </div>
                      <div style={{display:"flex",gap:12,marginTop:8,fontSize:11,color:C.muted}}>
                        <span style={{display:"flex",alignItems:"center",gap:5}}><span style={{width:8,height:8,borderRadius:2,background:`${C.purple}55`,display:"inline-block"}}/>訪客</span>
                        <span style={{display:"flex",alignItems:"center",gap:5}}><span style={{width:12,height:2,background:C.pink,display:"inline-block",borderRadius:1}}/>訂單</span>
                      </div>
                    </PanelBox>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* 商品排名 */}
          {view==="ranking"&&(
            <div>
              {prodData.length===0?(
                <div>
                  {/* 說明目前上傳的是什麼 */}
                  <div style={{background:"rgba(245,158,11,0.08)",border:"1px solid rgba(245,158,11,0.25)",borderRadius:14,padding:"16px 18px",marginBottom:14,display:"flex",gap:14,alignItems:"flex-start"}}>
                    <span style={{fontSize:22,flexShrink:0}}>⚠️</span>
                    <div>
                      <div style={{fontSize:13,fontWeight:600,color:C.amber,marginBottom:6}}>「商品概覽」不含個別商品名稱</div>
                      <div style={{fontSize:13,color:"#94a3b8",lineHeight:1.8}}>
                        你上傳的 <code style={{background:"rgba(255,255,255,0.07)",padding:"1px 7px",borderRadius:4,color:C.text}}>productoverview</code> 是每日流量彙總，已正確存入每日數據。<br/>
                        商品排名需要另一份報表，請依下方步驟匯出。
                      </div>
                    </div>
                  </div>

                  {/* 步驟說明 */}
                  <PanelBox title="如何取得商品銷售排名">
                    <div style={{display:"flex",flexDirection:"column",gap:10}}>
                      {[
                        {step:"1",text:"蝦皮賣家中心 → 上方「數據」選單",icon:"ti-layout-dashboard"},
                        {step:"2",text:"左側選「商品分析」→ 切到「商品銷售」頁籤",icon:"ti-chart-bar"},
                        {step:"3",text:"設定日期範圍 → 點「匯出」→ 下載 .xlsx 或 .csv",icon:"ti-file-download"},
                        {step:"4",text:"把下載的檔案拖進這個頁面，商品排名就會出現",icon:"ti-drag-drop"},
                      ].map(({step,text,icon})=>(
                        <div key={step} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 14px",background:`${C.purple}10`,borderRadius:11,border:`1px solid ${C.purple}25`}}>
                          <div style={{width:28,height:28,borderRadius:8,background:`linear-gradient(135deg,${C.purple},${C.cyan})`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                            <span style={{fontSize:12,fontWeight:700,color:"#fff"}}>{step}</span>
                          </div>
                          <div style={{flex:1}}>
                            <div style={{fontSize:13,color:C.text}}>{text}</div>
                          </div>
                          <i className={`ti ${icon}`} style={{fontSize:16,color:C.muted}}/>
                        </div>
                      ))}
                    </div>

                    <div style={{marginTop:16,padding:"12px 14px",background:"rgba(0,229,255,0.05)",borderRadius:11,border:"1px solid rgba(0,229,255,0.15)"}}>
                      <div style={{fontSize:12,fontWeight:600,color:C.cyan,marginBottom:6}}>📋 匯出的檔案應包含這些欄位</div>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                        {["商品名稱","銷售額","銷售數量","訂單數"].map(f=>(
                          <span key={f} style={{fontSize:11,padding:"3px 10px",borderRadius:20,background:"rgba(0,229,255,0.1)",color:C.cyan,border:"1px solid rgba(0,229,255,0.2)"}}>{f}</span>
                        ))}
                      </div>
                      <div style={{fontSize:11,color:C.muted,marginTop:8}}>如果欄位名稱不同，也可以手動貼入格式：<code style={{color:C.text}}>商品名稱, 數量, 金額</code></div>
                    </div>
                  </PanelBox>
                </div>
              ):(
                <div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:10,marginBottom:16}}>
                    <StatCard label="商品種類" value={`${prodData.length} 項`} icon="ti-box" color={C.cyan}/>
                    <StatCard label="最暢銷" value={prodData[0]?.product||"—"} icon="ti-crown" color={C.amber}/>
                    <StatCard label="累計營業額" value={`$${fmt(prodData.reduce((s,p)=>s+p.revenue,0))}`} icon="ti-coin" color={C.purple}/>
                  </div>
                  <PanelBox title="銷售排行榜">
                    {prodData.map((p,i)=>{
                      const pct=Math.round(p.revenue/prodData[0].revenue*100)
                      const barBg=i===0?`linear-gradient(90deg,${C.amber},${C.pink})`:i===1?`linear-gradient(90deg,${C.cyan},${C.purple})`:i===2?`linear-gradient(90deg,${C.purple},${C.pink})`:`${C.purple}45`
                      return (
                        <div key={p.product} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 0",borderBottom:i<prodData.length-1?`1px solid ${C.border}`:"none"}}>
                          <div style={{fontSize:18,width:28,textAlign:"center",flexShrink:0}}>{i<3?["🥇","🥈","🥉"][i]:<span style={{fontSize:13,color:C.muted,fontWeight:600}}>{i+1}</span>}</div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,color:C.text,fontSize:13,marginBottom:6,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.product}</div>
                            <div style={{height:4,background:"rgba(255,255,255,0.06)",borderRadius:4}}><div style={{width:`${pct}%`,height:"100%",background:barBg,borderRadius:4}}/></div>
                          </div>
                          <div style={{textAlign:"right",flexShrink:0}}>
                            <div style={{fontSize:13,fontWeight:700,color:i<3?C.amber:C.text,fontVariantNumeric:"tabular-nums"}}>${fmt(p.revenue)}</div>
                            <div style={{fontSize:10,color:C.muted,marginTop:2}}>{p.days}天 · {p.qty||"—"}件</div>
                          </div>
                        </div>
                      )
                    })}
                    <div style={{textAlign:"right",marginTop:12}}>
                      <button onClick={async()=>{await storage.delete(PROD_KEY);setProdData([])}} style={{fontSize:11,color:C.err,background:"rgba(239,68,68,0.1)",border:`1px solid rgba(239,68,68,0.2)`,borderRadius:8,padding:"4px 12px",cursor:"pointer"}}>
                        <i className="ti ti-trash" style={{fontSize:12,marginRight:4}}/>清除商品資料
                      </button>
                    </div>
                  </PanelBox>
                </div>
              )}
            </div>
          )}

          {/* AI 經營建議 */}
          {view==="insights"&&<InsightsView dayData={dayData} prodData={prodData}/>}

          {/* 匯入 */}
          {view==="import"&&(
            <div>
              <DropZone onFile={processFile} processing={processing}/>
              <div style={{display:"flex",alignItems:"center",gap:12,margin:"20px 0"}}>
                <div style={{flex:1,height:1,background:C.border}}/>
                <span style={{fontSize:12,color:C.muted}}>或手動貼上文字</span>
                <div style={{flex:1,height:1,background:C.border}}/>
              </div>
              <PanelBox>
                <div style={{marginBottom:12}}>
                  <label style={{fontSize:12,color:C.muted,display:"block",marginBottom:6}}>商品資料日期</label>
                  <input type="date" value={importDate} onChange={e=>setImportDate(e.target.value)} style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:"8px 12px",color:C.text,fontSize:13,outline:"none"}}/>
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,color:C.muted,display:"block",marginBottom:6}}>貼上報表內容</label>
                  <textarea value={importText} onChange={e=>setImportText(e.target.value)}
                    placeholder={"蝦皮店鋪統計（含標題行）：\n日期\t總銷售額 (TWD)\t訂單總數\t訪客數\n\n或商品 CSV：\n商品名稱,數量,金額\n貓咪飼料,10,2500"}
                    style={{width:"100%",height:140,fontSize:12,fontFamily:"monospace",resize:"vertical",padding:"12px",background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,color:C.text,lineHeight:1.6,boxSizing:"border-box",outline:"none"}}/>
                </div>
                <button onClick={handleTextImport} disabled={processing} style={{background:`linear-gradient(135deg,${C.purple},${C.cyan})`,border:"none",borderRadius:10,padding:"10px 22px",color:"#fff",cursor:"pointer",fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:7}}>
                  <i className="ti ti-upload" style={{fontSize:15}}/>貼上並匯入
                </button>
              </PanelBox>
              {dates.length>0&&(
                <div style={{marginTop:16}}>
                  <PanelBox title={`已儲存日報（${dates.length} 天）`}>
                    {dates.slice().reverse().slice(0,15).map((date,i,arr)=>{
                      const d=dayData[date]
                      return (
                        <div key={date} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:i<arr.length-1?`1px solid ${C.border}`:"none"}}>
                          <div style={{fontSize:12,fontWeight:500,color:C.text,minWidth:92,fontVariantNumeric:"tabular-nums"}}>{date}</div>
                          <div style={{fontSize:11,color:C.muted,flex:1}}>{d.orders} 筆 · {d.visitors} 訪客</div>
                          <div style={{fontSize:13,fontWeight:700,color:d.revenue>0?C.cyan:C.muted,fontVariantNumeric:"tabular-nums"}}>${fmt(d.revenue)}</div>
                          <button onClick={()=>deleteDay(date)} disabled={deleting===date} style={{width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(239,68,68,0.1)",border:`1px solid rgba(239,68,68,0.2)`,borderRadius:8,cursor:"pointer",color:C.err,fontSize:12}}>
                            {deleting===date?"…":"✕"}
                          </button>
                        </div>
                      )
                    })}
                    {dates.length>15&&<div style={{textAlign:"center",padding:"10px 0",fontSize:11,color:C.muted}}>共 {dates.length} 天，顯示最近 15 筆</div>}
                  </PanelBox>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  )
}
