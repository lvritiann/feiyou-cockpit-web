# -*- coding: utf-8 -*-
"""
build_dashboard.py — 把「非油销售汇总驾驶舱V1.xlsx」图表驾驶舱里的 10 张图表
重做成纯静态网页 (web/index.html)，数据内嵌、零 Excel 依赖、加载飞快。

用法:
    python build_dashboard.py
依赖: openpyxl (项目隔离环境已带)
说明:
    - 生成的 index.html 可单独丢到任意静态托管 (GitHub Pages / Nginx / OSS)。
    - 每次 Excel 数据刷新后，重跑本脚本即可让网页同步最新快照。
"""
import os, json, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.join(HERE, "..", "非油销售汇总驾驶舱V1.xlsx")
OUT  = os.path.join(HERE, "index.html")

wb = openpyxl.load_workbook(DASH, data_only=False)
cd = wb["图表数据源"]
cp = wb["图表驾驶舱"]

def g(r, c):  # 图表数据源 单元格
    return cd.cell(r, c).value
def p(r, c):  # 图表驾驶舱 单元格
    return cp.cell(r, c).value

# ---------- 图1-4: 各站年任务与累计完成 (块6, 行93-107) ----------
stations   = [g(r, 1) for r in range(93, 108)]      # 15 站
year_task  = [g(r, 2) for r in range(93, 108)]      # 年任务(万)
comp_tob   = [g(r, 3) for r in range(93, 108)]      # 含烟草累计完成(万)
comp_non   = [g(r, 5) for r in range(93, 108)]      # 不含烟草累计完成(万)
rate_tob   = [g(r, 4) for r in range(93, 108)]      # 含烟草完成率
rate_non   = [g(r, 6) for r in range(93, 108)]      # 不含烟草完成率

# ---------- 图5/6: 各站分会计月销售趋势 (块1 含烟草 行6-11 / 块2 不含 行25-30) ----------
MONTHS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
trend_stations = [g(r, 1) for r in range(6, 12)]     # 趋势图含的 6 站
trend_tob   = [[g(r, 1 + m) for m in range(1, 9)] for r in range(6, 12)]   # 含烟草 元
trend_non   = [[g(r, 1 + m) for m in range(1, 9)] for r in range(25, 31)]  # 不含烟草 元

# ---------- 图7/8: 全公司月度变化 (行86 含烟草 / 87 不含 / 88 含烟草环比) ----------
month_tob = [g(86, 1 + m) for m in range(1, 13)]     # 含烟草月度金额 元
month_non = [g(87, 1 + m) for m in range(1, 13)]     # 不含烟草月度金额 元
month_tob_mom = [g(88, 1 + m) for m in range(1, 13)] # 含烟草环比
# 不含烟草环比现场推算
month_non_mom = [None] + [
    round((month_non[i] - month_non[i-1]) / month_non[i-1], 4) if (month_non[i] is not None and month_non[i-1]) else None
    for i in range(1, 12)
]
# 只取有数据的月份 (1-8)
def trim(row):
    return [row[i] for i in range(8)]
month_labels = MONTHS[:8]
month_tob = trim(month_tob); month_non = trim(month_non)
month_tob_mom = trim(month_tob_mom); month_non_mom = trim(month_non_mom)
trend_tob = [trim(r) for r in trend_tob]; trend_non = [trim(r) for r in trend_non]

# ---------- 图9: 含烟草金额构成 (饼) 源头 图表数据源 N82:N85 /1e4 ----------
pie_raw = [g(82, 14), g(83, 14), g(84, 14), g(85, 14)]   # 元
pie = [["汽车用品", round(pie_raw[0]/1e4, 2)],
       ["便利百货", round(pie_raw[1]/1e4, 2)],
       ["香烟",     round(pie_raw[2]/1e4, 2)],
       ["咖啡",     round(pie_raw[3]/1e4, 2)]]

# ---------- 图10: 双口径对比 (金额/毛利, 万元) ----------
dual = {
    "cats":   ["金额", "毛利"],
    "nontob": [round(g(40, 14)/1e4, 2), round(g(78, 14)/1e4, 2)],  # N40/N78
    "tob":    [round(g(21, 14)/1e4, 2), round(g(59, 14)/1e4, 2)],  # N21/N59
}

DATA = {
    "stations": stations, "year_task": year_task,
    "comp_tob": comp_tob, "comp_non": comp_non,
    "rate_tob": rate_tob, "rate_non": rate_non,
    "trend_stations": trend_stations, "trend_months": month_labels,
    "trend_tob": trend_tob, "trend_non": trend_non,
    "month_labels": month_labels,
    "month_tob": month_tob, "month_non": month_non,
    "month_tob_mom": month_tob_mom, "month_non_mom": month_non_mom,
    "pie": pie, "dual": dual,
}

# ============================ 生成 HTML ============================
HTML = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>非油销售汇总驾驶舱 · 图表视图</title>
<script src="https://lib.baomitu.com/echarts/5.5.0/echarts.min.js"></script>
<script>
if (typeof echarts === 'undefined') {
  document.write('<scr'+'ipt src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"><\\/scr'+'ipt>');
}
</script>
<style>
  :root{ --bg:#0f1420; --card:#1a2030; --line:#2a3346; --txt:#e6edf3; --sub:#9aa7bd; --acc:#4ea1ff; --acc2:#ff9f43; --good:#3ddc97; }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--txt);font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif;padding:18px}
  header{margin-bottom:14px;border-bottom:1px solid var(--line);padding-bottom:12px}
  header h1{font-size:20px;font-weight:600}
  header p{color:var(--sub);font-size:12px;margin-top:4px}
  .grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
  @media(max-width:900px){.grid{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px 4px}
  .card h3{font-size:13px;font-weight:500;color:var(--sub);margin-bottom:6px;padding-left:6px;border-left:3px solid var(--acc)}
  .chart{width:100%;height:300px}
  footer{margin-top:16px;color:var(--sub);font-size:11px;text-align:center}
</style>
</head>
<body>
<header>
  <h1>非油销售汇总驾驶舱 · 图表视图</h1>
  <p>静态快照 · 数据取自「图表驾驶舱」10 张图表 · 由 build_dashboard.py 从驾驶舱 Excel 生成</p>
</header>
<main class="grid">
  <div class="card"><h3>任务完成情况（含烟草）· 各站年任务 vs 累计完成（万元）</h3><div id="c1" class="chart"></div></div>
  <div class="card"><h3>各站年任务 vs 累计完成（不含烟草，万元）</h3><div id="c2" class="chart"></div></div>
  <div class="card"><h3>各站任务完成比（含烟草）· 累计完成 ÷ 年任务</h3><div id="c3" class="chart"></div></div>
  <div class="card"><h3>各站任务完成比（不含烟草）· 累计完成 ÷ 年任务</h3><div id="c4" class="chart"></div></div>
  <div class="card"><h3>各任务销售趋势（含烟草）· 各站分会计月非油金额（元）</h3><div id="c5" class="chart"></div></div>
  <div class="card"><h3>各任务销售趋势（不含烟草）· 各站分会计月非油金额（元）</h3><div id="c6" class="chart"></div></div>
  <div class="card"><h3>非油月度变化（含烟草）· 月度金额与环比</h3><div id="c7" class="chart"></div></div>
  <div class="card"><h3>非油月度变化（不含烟草）· 月度金额与环比</h3><div id="c8" class="chart"></div></div>
  <div class="card"><h3>含烟草金额构成（汽车用品 / 便利百货 / 香烟 / 咖啡，万元）</h3><div id="c9" class="chart"></div></div>
  <div class="card"><h3>含烟草 vs 不含烟草（金额 / 毛利，万元）</h3><div id="c10" class="chart"></div></div>
</main>
<footer>本页为静态快照，不含动态计算能力；数据更新后重跑 build_dashboard.py 重新生成即可。</footer>
<script>
const DATA = __DATA__;
const grid = {left:48,right:18,top:36,bottom:30};
const baseGrid = {left:50,right:20,top:40,bottom:34};
const cat = (data)=>({type:'category',data:data,axisLine:{lineStyle:{color:'#3a4660'}},axisLabel:{color:'#9aa7bd',fontSize:10}});
const val = (name,fmt)=>({type:'value',name:name,axisLine:{lineStyle:{color:'#3a4660'}},splitLine:{lineStyle:{color:'#222b3d'}},axisLabel:{color:'#9aa7bd',fontSize:10,formatter:fmt||undefined}});
const tip = {trigger:'axis',backgroundColor:'#0d1220',borderColor:'#2a3346',textStyle:{color:'#e6edf3'}};
const tipItem = {trigger:'item',backgroundColor:'#0d1220',borderColor:'#2a3346',textStyle:{color:'#e6edf3'}};
const legend = {textStyle:{color:'#9aa7bd',fontSize:11},top:6};
const PALETTE = ['#4ea1ff','#ff9f43','#3ddc97','#b07bff','#ff6b6b','#f7d154','#5ad1e0'];
const charts = [];
function mk(id,opt){ const el=document.getElementById(id); const c=echarts.init(el,null,{renderer:'canvas'}); c.setOption(opt); charts.push(c); }

// 图1
mk('c1',{tooltip:tip,legend:legend,grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.stations),yAxis:val('万元'),
  series:[
    {name:'年任务',type:'bar',data:DATA.year_task,barWidth:14},
    {name:'含烟草累计完成',type:'bar',data:DATA.comp_tob,barWidth:14}
  ]});
// 图2
mk('c2',{tooltip:tip,legend:legend,grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.stations),yAxis:val('万元'),
  series:[
    {name:'年任务',type:'bar',data:DATA.year_task,barWidth:14},
    {name:'不含烟草累计完成',type:'bar',data:DATA.comp_non,barWidth:14}
  ]});
// 图3
mk('c3',{tooltip:tip,grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.stations),yAxis:val('完成率',v=>(v*100).toFixed(0)+'%'),
  series:[{name:'含烟草完成率',type:'bar',data:DATA.rate_tob.map(v=>v==null?null:+(v*100).toFixed(1)),
    label:{show:true,position:'top',color:'#9aa7bd',fontSize:9,formatter:p=>(p.value==null?'':p.value+'%')}}]});
// 图4
mk('c4',{tooltip:tip,grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.stations),yAxis:val('完成率',v=>(v*100).toFixed(0)+'%'),
  series:[{name:'不含烟草完成率',type:'bar',data:DATA.rate_non.map(v=>v==null?null:+(v*100).toFixed(1)),
    label:{show:true,position:'top',color:'#9aa7bd',fontSize:9,formatter:p=>(p.value==null?'':p.value+'%')}}]});
// 图5
mk('c5',{tooltip:tip,legend:Object.assign({data:DATA.trend_stations},legend),grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.trend_months),yAxis:val('元'),
  series:DATA.trend_stations.map((s,i)=>({name:s,type:'line',smooth:true,data:DATA.trend_tob[i]}))});
// 图6
mk('c6',{tooltip:tip,legend:Object.assign({data:DATA.trend_stations},legend),grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.trend_months),yAxis:val('元'),
  series:DATA.trend_stations.map((s,i)=>({name:s,type:'line',smooth:true,data:DATA.trend_non[i]}))});
// 图7 含烟草 月度金额(柱)+环比(线,次轴)
mk('c7',{tooltip:tip,legend:Object.assign({data:['含烟草金额','含烟草环比']},legend),grid:Object.assign({},baseGrid,{right:55}),color:PALETTE,
  xAxis:cat(DATA.month_labels),
  yAxis:[val('元'),{type:'value',name:'环比',axisLine:{lineStyle:{color:'#3a4660'}},splitLine:{show:false},axisLabel:{color:'#9aa7bd',fontSize:10,formatter:v=>(v*100).toFixed(0)+'%'}}],
  series:[
    {name:'含烟草金额',type:'bar',data:DATA.month_tob,barWidth:20},
    {name:'含烟草环比',type:'line',yAxisIndex:1,smooth:true,data:DATA.month_tob_mom,lineStyle:{width:2},itemStyle:{color:'#ff9f43'}}
  ]});
// 图8 不含烟草 月度金额(柱)+环比(线,次轴)
mk('c8',{tooltip:tip,legend:Object.assign({data:['不含烟草金额','不含烟草环比']},legend),grid:Object.assign({},baseGrid,{right:55}),color:PALETTE,
  xAxis:cat(DATA.month_labels),
  yAxis:[val('元'),{type:'value',name:'环比',axisLine:{lineStyle:{color:'#3a4660'}},splitLine:{show:false},axisLabel:{color:'#9aa7bd',fontSize:10,formatter:v=>(v*100).toFixed(0)+'%'}}],
  series:[
    {name:'不含烟草金额',type:'bar',data:DATA.month_non,barWidth:20},
    {name:'不含烟草环比',type:'line',yAxisIndex:1,smooth:true,data:DATA.month_non_mom,lineStyle:{width:2},itemStyle:{color:'#ff9f43'}}
  ]});
// 图9 饼
mk('c9',{tooltip:tipItem,legend:Object.assign({orient:'vertical',left:'left'},legend),color:PALETTE,
  series:[{type:'pie',radius:['38%','66%'],center:['58%','54%'],
    label:{color:'#e6edf3',fontSize:11,formatter:'{b}\\n{c}万'},
    data:DATA.pie.map(d=>({name:d[0],value:d[1]}))}]});
// 图10 双口径
mk('c10',{tooltip:tip,legend:Object.assign({data:['不含烟草','含烟草']},legend),grid:baseGrid,color:PALETTE,
  xAxis:cat(DATA.dual.cats),yAxis:val('万元'),
  series:[
    {name:'不含烟草',type:'bar',data:DATA.dual.nontob,barWidth:34},
    {name:'含烟草',type:'bar',data:DATA.dual.tob,barWidth:34}
  ]});

window.addEventListener('resize',()=>charts.forEach(c=>c.resize()));
</script>
</body>
</html>
"""

HTML = HTML.replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML)
print("✅ 已生成:", OUT)
print("   数据规模: 站点", len(stations), "站 | 趋势", len(trend_stations), "站×",
      len(month_labels), "月 | 饼", len(pie), "类 | 双口径", len(dual["cats"]), "项")
