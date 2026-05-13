import { useState, useEffect, useRef, useCallback } from "react"
import * as XLSX from "xlsx"
import {
  AreaChart, Area, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell,
  ComposedChart, Line,
} from "recharts"
import storage from "./storage.js"

/* ─── Storage keys ─── */
const IDX_KEY  = "sales2:index"
const DAY_KEY  = (d) => `sales2:day:${d}`
const PROD_KEY = "sales2:products"

/* ─── Design tokens ─── */
const C = {
  bg:     "#0b0d1a",
  panel:  "#111326",
  card:   "#161829",
  border: "rgba(255,255,255,0.06)",
  cyan:   "#00e5ff",
  pink:   "#ff3cac",
  purple: "#7c3aed",
  amber:  "#f59e0b",
  text:   "#e2e8f0",
  muted:  "#64748b",
  ok:     "#10b981",
  err:    "#ef4444",
}

const fmt   = (n) => Math.round(n).toLocaleString("zh-TW")
const fmtK  = (v) => v >= 10000 ? `${(v/10000).toFixed(1)}萬` : v >= 1000 ? `${(v/1000).toFixed(0)}k` : `${Math.round(v)}`
const today = new Date().toISOString().slice(0, 10)

/* ──────────────────────────────────────────
   Parse helpers
────────────────────────────────────────── */
function parseShopeeText(text) {
  const lines = text.trim().split("\n").filter(l => l.trim())
  if (lines.length < 2) return null
  const sep = lines[0].includes("\t") ? "\t" : ","
  const hdr = lines[0].split(sep).map(h => h.trim().replace(/^["']|["']$/g, ""))
  const di = hdr.findIndex(h => h === "日期")
  const ri = hdr.findIndex(h => h.includes("總銷售額"))
  const oi = hdr.findIndex(h => h === "訂單總數")
  const vi = hdr.findIndex(h => h === "訪客數")
  const ci = hdr.findIndex(h => h === "訂單轉換率")
  if (di < 0 || ri < 0) return null
  const rows = []
  for (let i = 1; i < lines.length; i++) {
    const p = lines[i].split(sep).map(x => x.trim().replace(/^["']|["']$/g, ""))
    const date = p[di]
    if (!date || !date.match(/^\d{4}-\d{2}-\d{2}$/)) continue
    rows.push({
      date,
      revenue:  parseFloat(p[ri]?.replace(/,/g, "")) || 0,
      orders:   parseInt(p[oi]?.replace(/,/g, ""))   || 0,
      visitors: parseInt(p[vi]?.replace(/,/g, ""))   || 0,
      conv:     p[ci] || "0%",
    })
  }
  return rows.length > 0 ? { type: "shopee", rows } : null
}

function parseProductText(text) {
  const lines = text.trim().split("\n").filter(l => l.trim())
  if (lines.length < 2) return null
  const sep = lines[0].includes("\t") ? "\t" : ","
  const rows = []
  for (let i = 1; i < lines.length; i++) {
    const p = lines[i].split(sep).map(x => x.trim().replace(/^["']|["']$/g, "").replace(/,/g, ""))
    if (!p[0]) continue
    let qty = 0, revenue = 0
    if (p.length === 2)      revenue = parseFloat(p[1]) || 0
    else if (p.length >= 3) { qty = parseInt(p[1]) || 0; revenue = parseFloat(p[2]) || 0 }
    if (p[0] && revenue > 0) rows.push({ product: p[0], qty, revenue })
  }
  return rows.length > 0 ? { type: "product", rows } : null
}

async function parseFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader()
    const isXlsx = /\.(xlsx|xls)$/i.test(file.name)

    reader.onload = (e) => {
      try {
        if (isXlsx) {
          const wb = XLSX.read(e.target.result, { type: "array" })
          for (const sheetName of wb.SheetNames) {
            const ws  = wb.Sheets[sheetName]
            const tsv = XLSX.utils.sheet_to_csv(ws, { FS: "\t", blankrows: false })
            const shopee  = parseShopeeText(tsv)
            const product = shopee ? null : parseProductText(tsv)
            if (shopee || product) { resolve(shopee || product); return }
          }
          resolve(null)
        } else {
          const text    = e.target.result
          const shopee  = parseShopeeText(text)
          const product = shopee ? null : parseProductText(text)
          resolve(shopee || product || null)
        }
      } catch { resolve(null) }
    }
    reader.onerror = () => resolve(null)
    if (isXlsx) reader.readAsArrayBuffer(file)
    else         reader.readAsText(file)
  })
}

/* ──────────────────────────────────────────
   UI helpers
────────────────────────────────────────── */
function GradientCard({ label, value, sub, from, to }) {
  return (
    <div style={{ background:`linear-gradient(135deg,${from},${to})`, borderRadius:16, padding:"20px 22px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:-20, right:-20, width:80, height:80, borderRadius:"50%", background:"rgba(255,255,255,0.08)" }}/>
      <div style={{ fontSize:11, color:"rgba(255,255,255,0.7)", letterSpacing:"0.07em", textTransform:"uppercase", marginBottom:10 }}>{label}</div>
      <div style={{ fontSize:30, fontWeight:700, color:"#fff", letterSpacing:"-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize:12, color:"rgba(255,255,255,0.6)", marginTop:5 }}>{sub}</div>}
    </div>
  )
}

function StatCard({ label, value, icon, color }) {
  return (
    <div style={{ background:C.card, borderRadius:14, padding:"16px 18px", border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:14 }}>
      <div style={{ width:42, height:42, borderRadius:12, background:`${color}20`, display:"flex", alignItems:"center", justifyContent:"center" }}>
        <i className={`ti ${icon}`} style={{ fontSize:20, color }} />
      </div>
      <div style={{ flex:1, minWidth:0 }}>
        <div style={{ fontSize:11, color:C.muted, letterSpacing:"0.05em", textTransform:"uppercase", marginBottom:4 }}>{label}</div>
        <div style={{ fontSize:20, fontWeight:600, color:C.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{value}</div>
      </div>
    </div>
  )
}

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null
  return (
    <div style={{ background:"#1e2130", border:"1px solid rgba(255,255,255,0.08)", borderRadius:10, fontSize:12, color:C.text, padding:"10px 14px" }}>
      <div style={{ color:C.muted, fontSize:11, marginBottom:6 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:2 }}>
          <span style={{ width:8, height:8, borderRadius:2, background:p.color||p.stroke, display:"inline-block" }}/>
          <span style={{ color:C.muted, fontSize:11 }}>
            {p.name==="revenue"?"營業額":p.name==="orders"?"訂單":p.name==="visitors"?"訪客":p.name}
          </span>
          <span style={{ fontWeight:600, color:C.text, marginLeft:"auto", paddingLeft:16 }}>
            {p.name==="revenue" ? `$${fmt(p.value)}` : p.value}
          </span>
        </div>
      ))}
    </div>
  )
}

function PanelBox({ title, children }) {
  return (
    <div style={{ background:C.card, borderRadius:16, border:`1px solid ${C.border}`, padding:"20px 22px" }}>
      {title && <div style={{ fontSize:13, fontWeight:600, color:C.text, marginBottom:18 }}>{title}</div>}
      {children}
    </div>
  )
}

function DropZone({ onFile, processing }) {
  const [hover, setHover] = useState(false)
  const inputRef = useRef()

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setHover(false)
    const file = e.dataTransfer.files?.[0]
    if (file) onFile(file)
  }, [onFile])

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setHover(true) }}
      onDragLeave={() => setHover(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      style={{
        border:`2px dashed ${hover ? C.cyan : `${C.purple}60`}`,
        borderRadius:16, padding:"40px 24px", textAlign:"center",
        cursor:"pointer", transition:"all 0.2s",
        background: hover ? `${C.cyan}08` : `${C.purple}08`,
      }}
    >
      <input
        ref={inputRef} type="file" accept=".xlsx,.xls,.csv,.tsv,.txt"
        style={{ display:"none" }}
        onChange={(e) => { const f=e.target.files?.[0]; if(f) onFile(f); e.target.value="" }}
      />
      {processing ? (
        <>
          <div style={{ fontSize:36, marginBottom:12, color:C.cyan }}>
            <i className="ti ti-loader" style={{ animation:"spin 1s linear infinite", display:"inline-block" }}/>
          </div>
          <div style={{ fontSize:14, fontWeight:600, color:C.cyan }}>解析中...</div>
        </>
      ) : (
        <>
          <div style={{ width:60, height:60, borderRadius:18, background:hover?`${C.cyan}20`:`${C.purple}20`, display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 18px" }}>
            <i className="ti ti-cloud-upload" style={{ fontSize:30, color:hover?C.cyan:C.purple }}/>
          </div>
          <div style={{ fontSize:16, fontWeight:600, color:C.text, marginBottom:8 }}>
            {hover ? "放開以上傳" : "拖曳檔案到這裡"}
          </div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:14 }}>或點擊選擇檔案</div>
          <div style={{ display:"flex", gap:8, justifyContent:"center" }}>
            {[".xlsx",".csv",".tsv"].map(ext => (
              <span key={ext} style={{ fontSize:11, padding:"3px 12px", borderRadius:20, background:`${C.purple}20`, color:C.purple, border:`1px solid ${C.purple}40` }}>{ext}</span>
            ))}
          </div>
        </>
      )}
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  )
}

function GlobalDragOverlay({ visible }) {
  if (!visible) return null
  return (
    <div style={{
      position:"fixed", inset:0, zIndex:9999,
      background:"rgba(11,13,26,0.9)",
      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
      border:`2px solid ${C.cyan}`,
      pointerEvents:"none",
    }}>
      <i className="ti ti-cloud-upload" style={{ fontSize:64, color:C.cyan, marginBottom:20 }}/>
      <div style={{ fontSize:22, fontWeight:700, color:C.text }}>放開以匯入報表</div>
      <div style={{ fontSize:14, color:C.muted, marginTop:10 }}>支援 .xlsx · .csv · .tsv</div>
    </div>
  )
}

/* ──────────────────────────────────────────
   Main Dashboard
────────────────────────────────────────── */
export default function App() {
  const [dayData,    setDayData]    = useState({})
  const [prodData,   setProdData]   = useState([])
  const [view,       setView]       = useState("dashboard")
  const [importDate, setImportDate] = useState(today)
  const [importText, setImportText] = useState("")
  const [msg,        setMsg]        = useState(null)
  const [loading,    setLoading]    = useState(true)
  const [deleting,   setDeleting]   = useState(null)
  const [processing, setProcessing] = useState(false)
  const [dragging,   setDragging]   = useState(false)
  const dragCounter = useRef(0)

  useEffect(() => { loadAll() }, [])

  /* Global drag */
  useEffect(() => {
    const onEnter = (e) => { e.preventDefault(); dragCounter.current++; if (e.dataTransfer.items?.[0]?.kind==="file") setDragging(true) }
    const onLeave = ()  => { dragCounter.current--; if (dragCounter.current === 0) setDragging(false) }
    const onOver  = (e) => e.preventDefault()
    const onDrop  = async (e) => {
      e.preventDefault(); dragCounter.current = 0; setDragging(false)
      const file = e.dataTransfer.files?.[0]
      if (file) await processFile(file)
    }
    window.addEventListener("dragenter", onEnter)
    window.addEventListener("dragleave", onLeave)
    window.addEventListener("dragover",  onOver)
    window.addEventListener("drop",      onDrop)
    return () => {
      window.removeEventListener("dragenter", onEnter)
      window.removeEventListener("dragleave", onLeave)
      window.removeEventListener("dragover",  onOver)
      window.removeEventListener("drop",      onDrop)
    }
  }, [])

  /* Storage */
  async function loadAll() {
    try {
      const idx = await storage.get(IDX_KEY)
      if (idx) {
        const dates = JSON.parse(idx.value)
        const data  = {}
        for (const d of dates) {
          try { const r = await storage.get(DAY_KEY(d)); if (r) data[d] = JSON.parse(r.value) } catch {}
        }
        setDayData(data)
      }
      try { const pr = await storage.get(PROD_KEY); if (pr) setProdData(JSON.parse(pr.value)) } catch {}
    } catch {}
    setLoading(false)
  }

  async function saveDays(rows) {
    let dates = []
    try { const idx = await storage.get(IDX_KEY); if (idx) dates = JSON.parse(idx.value) } catch {}
    const nd = { ...dayData }
    for (const r of rows) {
      const val = { revenue:r.revenue, orders:r.orders, visitors:r.visitors, conv:r.conv }
      await storage.set(DAY_KEY(r.date), JSON.stringify(val))
      nd[r.date] = val
      if (!dates.includes(r.date)) dates.push(r.date)
    }
    dates.sort()
    await storage.set(IDX_KEY, JSON.stringify(dates))
    setDayData(nd)
    return rows.length
  }

  async function saveProducts(date, rows) {
    const pm = {}
    prodData.forEach(p => { pm[p.product] = { ...p } })
    rows.forEach(({ product, qty, revenue }) => {
      if (!pm[product]) pm[product] = { product, qty:0, revenue:0, days:0 }
      pm[product].qty += qty; pm[product].revenue += revenue; pm[product].days += 1
    })
    const sorted = Object.values(pm).sort((a,b) => b.revenue - a.revenue)
    await storage.set(PROD_KEY, JSON.stringify(sorted))
    setProdData(sorted)
    return rows.length
  }

  async function deleteDay(date) {
    setDeleting(date)
    try {
      await storage.delete(DAY_KEY(date))
      let dates = []
      try { const idx = await storage.get(IDX_KEY); if (idx) dates = JSON.parse(idx.value) } catch {}
      dates = dates.filter(d => d !== date)
      await storage.set(IDX_KEY, JSON.stringify(dates))
      setDayData(prev => { const n = { ...prev }; delete n[date]; return n })
    } catch {}
    setDeleting(null)
  }

  function showMsg(text, ok) {
    setMsg({ text, ok })
    setTimeout(() => setMsg(null), 6000)
  }

  async function processFile(file) {
    setProcessing(true)
    const result = await parseFile(file)
    if (!result) {
      showMsg(`無法解析「${file.name}」，請確認格式`, false)
    } else if (result.type === "shopee") {
      const n = await saveDays(result.rows)
      showMsg(`✓ 從「${file.name}」匯入 ${n} 天資料`, true)
      setView("dashboard")
    } else {
      const n = await saveProducts(importDate, result.rows)
      showMsg(`✓ 從「${file.name}」匯入 ${n} 項商品`, true)
      setView("ranking")
    }
    setProcessing(false)
  }

  async function handleTextImport() {
    if (!importText.trim()) { showMsg("請貼上報表資料", false); return }
    setProcessing(true)
    const shopee  = parseShopeeText(importText)
    const product = !shopee && parseProductText(importText)
    if (shopee)        { const n = await saveDays(shopee.rows);           showMsg(`✓ 已匯入 ${n} 天蝦皮報表`, true) }
    else if (product)  { const n = await saveProducts(importDate, product.rows); showMsg(`✓ 已儲存 ${importDate}（${n} 項商品）`, true) }
    else                 showMsg("無法辨識格式", false)
    setImportText("")
    setProcessing(false)
  }

  /* Derived data */
  const dates        = Object.keys(dayData).sort()
  const dailyChart   = dates.map(d => ({ date:d.slice(5), fullDate:d, revenue:dayData[d].revenue, orders:dayData[d].orders, visitors:dayData[d].visitors }))
  const monthlyMap   = {}
  dates.forEach(d => {
    const m = d.slice(0,7)
    if (!monthlyMap[m]) monthlyMap[m] = { revenue:0, orders:0, visitors:0 }
    monthlyMap[m].revenue  += dayData[d].revenue
    monthlyMap[m].orders   += dayData[d].orders
    monthlyMap[m].visitors += dayData[d].visitors
  })
  const monthlyChart = Object.entries(monthlyMap).sort().map(([m,v]) => ({ month:`${parseInt(m.slice(5))}月`, ...v }))
  const todayRev     = dayData[today]?.revenue ?? 0
  const curMonth     = today.slice(0,7)
  const monthRev     = dates.filter(d=>d.startsWith(curMonth)).reduce((s,d)=>s+dayData[d].revenue,0)
  const totalOrds    = dates.reduce((s,d)=>s+dayData[d].orders,0)
  const totalVis     = dates.reduce((s,d)=>s+dayData[d].visitors,0)
  const convRate     = totalVis>0 ? ((totalOrds/totalVis)*100).toFixed(1)+"%" : "—"

  const NAV = [
    ["dashboard","ti-layout-dashboard","儀表板"],
    ["ranking",  "ti-trophy",          "商品排名"],
    ["import",   "ti-file-upload",     "匯入報表"],
  ]

  if (loading) return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"center", height:"50vh", color:C.muted, fontSize:14 }}>
      <i className="ti ti-loader" style={{fontSize:20,marginRight:8}}/>載入中...
    </div>
  )

  return (
    <>
      <GlobalDragOverlay visible={dragging}/>
      <div style={{ background:C.bg, borderRadius:20, display:"flex", overflow:"hidden", minHeight:650, border:`1px solid ${C.border}` }}>

        {/* Sidebar */}
        <div style={{ width:190, background:C.panel, flexShrink:0, display:"flex", flexDirection:"column", borderRight:`1px solid ${C.border}` }}>
          <div style={{ padding:"22px 18px 18px", borderBottom:`1px solid ${C.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <div style={{ width:34, height:34, borderRadius:10, background:`linear-gradient(135deg,${C.purple},${C.cyan})`, display:"flex", alignItems:"center", justifyContent:"center" }}>
                <i className="ti ti-chart-bar" style={{fontSize:18,color:"#fff"}}/>
              </div>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:C.text}}>銷售總覽</div>
                <div style={{fontSize:10,color:C.muted}}>蝦皮後台</div>
              </div>
            </div>
          </div>

          <div style={{ padding:"14px 10px", flex:1 }}>
            <div style={{fontSize:10,color:C.muted,letterSpacing:"0.08em",textTransform:"uppercase",padding:"0 8px",marginBottom:8}}>主選單</div>
            {NAV.map(([v,icon,label]) => {
              const active = view===v
              return (
                <button key={v} onClick={()=>setView(v)} style={{
                  width:"100%", display:"flex", alignItems:"center", gap:10, padding:"9px 12px",
                  borderRadius:10, border:"none", cursor:"pointer", textAlign:"left", marginBottom:2,
                  background: active ? `linear-gradient(135deg,${C.purple}40,${C.cyan}20)` : "transparent",
                  color: active ? C.cyan : C.muted,
                  boxShadow: active ? `inset 0 0 0 1px ${C.cyan}30` : "none",
                  transition:"all 0.15s",
                }}>
                  <i className={`ti ${icon}`} style={{fontSize:16}}/>
                  <span style={{fontSize:12,fontWeight:active?600:400}}>{label}</span>
                  {active && <div style={{marginLeft:"auto",width:5,height:5,borderRadius:"50%",background:C.cyan}}/>}
                </button>
              )
            })}
          </div>

          <div style={{ padding:"14px 16px", borderTop:`1px solid ${C.border}` }}>
            <div style={{ background:`${C.purple}15`, border:`1px dashed ${C.purple}40`, borderRadius:10, padding:"10px 12px", textAlign:"center", marginBottom:10 }}>
              <i className="ti ti-drag-drop" style={{fontSize:18,color:C.purple,display:"block",marginBottom:4}}/>
              <div style={{fontSize:10,color:C.muted,lineHeight:1.6}}>拖曳 .xlsx / .csv<br/>到任意位置匯入</div>
            </div>
            <div style={{fontSize:11,color:C.muted}}>已儲存 <span style={{color:C.text,fontWeight:600}}>{dates.length}</span> 天資料</div>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex:1, padding:"24px", overflowY:"auto", maxHeight:650 }}>

          {/* Toast */}
          {msg && (
            <div style={{
              display:"flex", alignItems:"center", gap:10, padding:"10px 16px",
              borderRadius:12, marginBottom:16, fontSize:13, fontWeight:500,
              background: msg.ok?"rgba(16,185,129,0.12)":"rgba(239,68,68,0.12)",
              border:`1px solid ${msg.ok?"rgba(16,185,129,0.3)":"rgba(239,68,68,0.3)"}`,
              color: msg.ok?C.ok:C.err,
            }}>
              <i className={`ti ${msg.ok?"ti-circle-check":"ti-alert-circle"}`} style={{fontSize:16}}/>
              {msg.text}
              <button onClick={()=>setMsg(null)} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",color:"inherit",padding:0,fontSize:14}}>✕</button>
            </div>
          )}

          {/* ── 儀表板 ── */}
          {view==="dashboard" && (
            <div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:12 }}>
                <GradientCard label="本月營業額" value={`$${fmt(monthRev)}`}  sub={`${curMonth.slice(5)}月累計`} from="#7c3aed" to="#3b82f6"/>
                <GradientCard label="今日營業額" value={`$${fmt(todayRev)}`}  sub={today}                        from="#0ea5e9" to="#06b6d4"/>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,minmax(0,1fr))", gap:10, marginBottom:18 }}>
                <StatCard label="訂單總數"   value={fmt(totalOrds)} icon="ti-shopping-bag"    color={C.cyan}/>
                <StatCard label="訪客總數"   value={fmt(totalVis)}  icon="ti-users"           color={C.pink}/>
                <StatCard label="整體轉換率" value={convRate}        icon="ti-arrows-exchange" color={C.purple}/>
              </div>

              {dates.length===0 ? (
                <div style={{ textAlign:"center", padding:"3rem", color:C.muted, background:C.card, borderRadius:16, border:`1px solid ${C.border}` }}>
                  <i className="ti ti-database-off" style={{fontSize:40,display:"block",marginBottom:12,color:`${C.purple}90`}}/>
                  <div style={{fontSize:14,marginBottom:8}}>尚未匯入任何報表</div>
                  <div style={{fontSize:12,color:C.muted,marginBottom:20}}>把蝦皮的 .xlsx 直接拖進視窗，或前往匯入頁面</div>
                  <button onClick={()=>setView("import")} style={{ background:`linear-gradient(135deg,${C.purple},${C.cyan})`, border:"none", borderRadius:10, padding:"9px 22px", color:"#fff", cursor:"pointer", fontWeight:600, fontSize:13 }}>
                    前往匯入 →
                  </button>
                </div>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:12}}>
                  <PanelBox title="每日營業額趨勢">
                    <div style={{height:200}}>
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={dailyChart} margin={{top:4,right:4,left:-10,bottom:0}}>
                          <defs>
                            <linearGradient id="gRev" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%"  stopColor={C.cyan} stopOpacity={0.3}/>
                              <stop offset="95%" stopColor={C.cyan} stopOpacity={0}/>
                            </linearGradient>
                          </defs>
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
                    {monthlyChart.length>0 && (
                      <PanelBox title="每月彙總">
                        <div style={{height:160}}>
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={monthlyChart} margin={{top:4,right:4,left:-14,bottom:0}}>
                              <defs>
                                {monthlyChart.map((_,i)=>(
                                  <linearGradient key={i} id={`bg${i}`} x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%"   stopColor={C.purple} stopOpacity={1}/>
                                    <stop offset="100%" stopColor={C.cyan}   stopOpacity={0.7}/>
                                  </linearGradient>
                                ))}
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" vertical={false}/>
                              <XAxis dataKey="month" tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false}/>
                              <YAxis tick={{fontSize:10,fill:C.muted}} axisLine={false} tickLine={false} tickFormatter={fmtK} width={34}/>
                              <Tooltip content={<CustomTooltip/>}/>
                              <Bar dataKey="revenue" radius={[6,6,0,0]}>
                                {monthlyChart.map((_,i)=>(<Cell key={i} fill={`url(#bg${i})`}/>))}
                              </Bar>
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

          {/* ── 商品排名 ── */}
          {view==="ranking" && (
            <div>
              {prodData.length===0 ? (
                <PanelBox title="商品銷售排名">
                  <div style={{padding:"14px 16px",background:`${C.purple}10`,borderRadius:12,border:`1px dashed ${C.purple}40`,marginBottom:16,lineHeight:1.9,fontSize:13,color:C.muted}}>
                    <strong style={{color:C.text}}>如何取得商品排名？</strong><br/>
                    蝦皮 → 數據中心 → <strong style={{color:C.cyan}}>商品分析</strong> → 匯出，再拖入儀表板<br/>
                    或手動格式：<code style={{fontSize:12,background:"rgba(0,229,255,0.1)",color:C.cyan,padding:"1px 8px",borderRadius:4}}>商品名稱, 數量, 金額</code>
                  </div>
                  <div style={{textAlign:"center",padding:"2rem",color:C.muted}}>
                    <i className="ti ti-box-off" style={{fontSize:40,display:"block",marginBottom:12,color:`${C.purple}80`}}/>尚無商品資料
                  </div>
                </PanelBox>
              ) : (
                <div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:10,marginBottom:16}}>
                    <StatCard label="商品種類"   value={`${prodData.length} 項`}                                  icon="ti-box"   color={C.cyan}/>
                    <StatCard label="最暢銷"     value={prodData[0]?.product||"—"}                                icon="ti-crown" color={C.amber}/>
                    <StatCard label="累計營業額" value={`$${fmt(prodData.reduce((s,p)=>s+p.revenue,0))}`}         icon="ti-coin"  color={C.purple}/>
                  </div>
                  <PanelBox title="銷售排行榜">
                    {prodData.map((p,i)=>{
                      const pct   = Math.round(p.revenue/prodData[0].revenue*100)
                      const barBg = i===0?`linear-gradient(90deg,${C.amber},${C.pink})`:i===1?`linear-gradient(90deg,${C.cyan},${C.purple})`:i===2?`linear-gradient(90deg,${C.purple},${C.pink})`:`${C.purple}45`
                      return (
                        <div key={p.product} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 0",borderBottom:i<prodData.length-1?`1px solid ${C.border}`:"none"}}>
                          <div style={{fontSize:18,width:28,textAlign:"center",flexShrink:0}}>
                            {i<3?["🥇","🥈","🥉"][i]:<span style={{fontSize:13,color:C.muted,fontWeight:600}}>{i+1}</span>}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,color:C.text,fontSize:13,marginBottom:6,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.product}</div>
                            <div style={{height:4,background:"rgba(255,255,255,0.06)",borderRadius:4}}>
                              <div style={{width:`${pct}%`,height:"100%",background:barBg,borderRadius:4}}/>
                            </div>
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

          {/* ── 匯入 ── */}
          {view==="import" && (
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
                  <input type="date" value={importDate} onChange={e=>setImportDate(e.target.value)}
                    style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:"8px 12px",color:C.text,fontSize:13,outline:"none"}}/>
                </div>
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:12,color:C.muted,display:"block",marginBottom:6}}>貼上報表內容</label>
                  <textarea value={importText} onChange={e=>setImportText(e.target.value)}
                    placeholder={"蝦皮店鋪統計（含標題行）：\n日期\t總銷售額 (TWD)\t訂單總數\t訪客數\n\n或商品 CSV：\n商品名稱,數量,金額\n貓咪飼料,10,2500"}
                    style={{width:"100%",height:140,fontSize:12,fontFamily:"monospace",resize:"vertical",padding:"12px",background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,color:C.text,lineHeight:1.6,boxSizing:"border-box",outline:"none"}}/>
                </div>
                <button onClick={handleTextImport} disabled={processing} style={{background:`linear-gradient(135deg,${C.purple},${C.cyan})`,border:"none",borderRadius:10,padding:"10px 22px",color:"#fff",cursor:"pointer",fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:7,opacity:processing?0.6:1}}>
                  <i className="ti ti-upload" style={{fontSize:15}}/>貼上並匯入
                </button>
              </PanelBox>

              {dates.length>0 && (
                <div style={{marginTop:16}}>
                  <PanelBox title={`已儲存日報（${dates.length} 天）`}>
                    {dates.slice().reverse().slice(0,15).map((date,i,arr)=>{
                      const d = dayData[date]
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
                    {dates.length>15 && <div style={{textAlign:"center",padding:"10px 0",fontSize:11,color:C.muted}}>共 {dates.length} 天，顯示最近 15 筆</div>}
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
