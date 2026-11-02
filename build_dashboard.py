# -*- coding: utf-8 -*-
"""
build_dashboard.py — 把「非油销售汇总驾驶舱V1.xlsx」图表数据源里的数据
重做成纯静态网页 (web/index.html) + 数据层 (web/data.json)。

设计要点:
- 全部用【标签扫描】定位「图表数据源」的 块1~块6 与「按月汇总各站」的分类明细块，
  绝不硬编码行号(此前硬编码曾导致尾部站点错位，见项目记忆)。
- 提取覆盖: 15 站 × 12 月 金额/毛利(含+不含) / 各站年任务与累计完成 / 公司月度
  金额+环比 / 品类构成(5类) / 双口径 / 【品类明细】(5品类+非油合计 × 15站 × 12月 金额+毛利)。
- 生成的 index.html: 数据内嵌 + 联动(站点选择器 + 点击图表联动筛选) + 站点明细弹窗
  (品类×12月金额/毛利堆叠柱 + 构成饼 + 合计表) + porcelain 青瓷蓝主题，加载飞快。
- 公司级图表(构成/双口径/任务/月度)在选中某站后自动切换为该站数据。
- 每日自动更新由 GitHub Action(10点) 重跑本脚本(读 data.json) + AirScript 云端
  推送 data.json 共同完成。

用法:
    python build_dashboard.py                     # 读 Excel(默认2027版) -> data.json + index.html
    python build_dashboard.py --year 2026         # 读 2026 版驾驶舱 -> data_2026.json + index.html
    python build_dashboard.py --from-json         # 只读 data.json -> 重渲染 index.html (Action 用)
    python build_dashboard.py --snapshot-2026     # 2026收官(12-25)归档: data.json -> data_2026.json
依赖: openpyxl

双年度说明(2026-10-08 规划):
- 2027会计年 = 2026-12-26 ~ 2027-12-25, 12-26 启动;
- 2027 数据固定写 data.json → GitHub Action(--from-json) 与 AirScript 推送链路零改动;
- 2026 会计年收官后 data.json 快照为 data_2026.json, 页面顶部年度切换器可回看;
- 2027 版驾驶舱若含「2026同比」sheet(收官夜注入), 页面自动叠加 2026 同期对比线。
"""
import os, json, sys, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.join(HERE, "..", "非油销售汇总驾驶舱V1.xlsx")
DATA_JSON = os.path.join(HERE, "data.json")
INDEX = os.path.join(HERE, "index.html")

STATION_ORDER = ["秦岭", "宁陕", "洋县", "汉中", "富平", "韩城", "富县", "南沙",
                 "华山", "白河", "旬阳", "略阳", "玉华宫", "照金", "天汉水城"]
CAT_DETAIL = ["汽车用品", "便利百货", "香烟零售", "烟草批发", "咖啡"]
CAT_DETAIL_ALL = CAT_DETAIL + ["非油合计"]
TASK_TOTAL_2026 = 1130  # 2026会计年全年任务(万元)

# ── 双年度配置(2026-10-08 规划v2: 2027会计年=2026-12-26~2027-12-25, 12-26启动) ──
# 【原地清空方案】不新建文件: 12-25 收官夜把现行 16 个文件云盘复制归档(改名+2026归档),
#   12-26 原地清空动态数据区续用 → FeiyouSubmit/pipeline/AirScript/各站链接全部零改动。
# 2027 的 json 固定叫 data.json: Action(--from-json) 与 AirScript 推送链路零改动;
# 2026 收官后由 --snapshot-2026 把最终 data.json 快照为 data_2026.json, 此后永不更新。
_ARCH_2026_DASH = os.path.join(HERE, "..", "非油销售汇总驾驶舱V1-2026归档.xlsx")
YEAR_CFG = {
    "2026": {
        "dash": _ARCH_2026_DASH,  # 12-25 归档副本(内容=2026全年原样), 仅重建快照时读
        "json": os.path.join(HERE, "data_2026.json"),
        "y0": datetime.datetime(2025, 12, 26),
        "task": TASK_TOTAL_2026,
    },
    "2027": {
        "dash": os.path.join(HERE, "..", "非油销售汇总驾驶舱V1.xlsx"),  # 原地续用, 文件名不变
        "json": os.path.join(HERE, "data.json"),
        "y0": datetime.datetime(2026, 12, 26),
        "task": None,  # 2027任务未下达 → 页面显示「任务待下达」
    },
}
DASH = YEAR_CFG["2027"]["dash"]      # 兼容旧引用
DATA_JSON = YEAR_CFG["2027"]["json"]  # 兼容旧引用
INDEX = os.path.join(HERE, "index.html")


def _num(v):
    """把单元格值安全转成 float。

    背景(2026-09-10 实测)：「按日汇总」第 4 行(=2025-12-26 手修行)的整行数值
    被以 opType:"formula" 写成了纯数字公式串，如 '=22084.02'。openpyxl 以
    data_only=False 读取时返回的是字符串 '=22084.02'，`isinstance(v,(int,float))`
    判定失败 → 该行被静默跳过，导致公司年度累计少了 22,084.02 元
    (网页显示 977.76万 vs 驾驶舱权威 979.97万)。
    故数值读取一律走本函数，兼容 数字 / '123' / '=123' 三种形态。
    """
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("="):
            s = s[1:].strip()
        try:
            return float(s)
        except Exception:
            return None
    return None


def load_excel(dash=None):
    import openpyxl
    return openpyxl.load_workbook(dash or DASH, data_only=False)


def parse_yoy_sheet(wb):
    """解析 2027 版驾驶舱新增的「2026同比」sheet(收官夜由 _snapshot_2026_yoy.py 注入静态值)。

    布局与「图表数据源」块1/块3 同构:
      A列含「块1 ·」标题 → 其后 15 站行 × 12 会计月列 = 2026同期各站含烟草金额(元)
      A列含「块2 ·」标题 → 其后 15 站行 × 12 会计月列 = 2026同期各站含烟草毛利(元)
    sheet 不存在或解析失败 → 返回 None(页面自动隐藏同比线, 不报错)。
    """
    if "2026同比" not in wb.sheetnames:
        return None
    ws = wb["2026同比"]
    try:
        b1 = find_row(ws, "块1 ·")
        b2 = find_row(ws, "块2 ·")
        if not b1:
            return None
        r1 = station_rows(ws, b1)
        r2 = station_rows(ws, b2) if b2 else {}
        if not r1:
            return None

        def _row_sum(rmap):
            per = {}
            for st, r in rmap.items():
                per[st] = [_num(ws.cell(r, 1 + m).value) or 0 for m in range(1, 13)]
            return per

        amt = _row_sum(r1)
        prof = _row_sum(r2)
        comp_amt = [round(sum(amt.get(st, [0] * 12)[m] for st in STATION_ORDER), 2) for m in range(12)]
        comp_prof = [round(sum(prof.get(st, [0] * 12)[m] for st in STATION_ORDER), 2) for m in range(12)]
        return {
            "monthLabels": [f"{m}月" for m in range(1, 13)],
            "monthlyTob": amt,
            "monthlyTobProfit": prof,
            "companyTob": comp_amt,
            "companyTobProfit": comp_prof,
        }
    except Exception:
        return None


def find_row(ws, kw, start=1, end=None):
    end = end or ws.max_row
    for r in range(start, end + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and kw in v:
            return r
    return None


def station_rows(ws, after_title):
    """标题行之后, 站名所在的行号映射(扫描到非站名/合计/下一块为止)。"""
    rows = {}
    for r in range(after_title + 2, after_title + 40):
        v = ws.cell(r, 1).value
        if v in STATION_ORDER:
            rows[v] = r
        elif isinstance(v, str) and ("合计" in v or "块" in v):
            break
    return rows


def total_row(ws, after_title):
    for r in range(after_title + 2, after_title + 40):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "合计" in v:
            return r
    return None


def build_data(year="2027"):
    cfg = YEAR_CFG.get(year) or YEAR_CFG["2027"]
    wb = load_excel(cfg["dash"])
    cd = wb["图表数据源"]
    ms = wb["按月汇总各站"]
    dd = wb["按日汇总"]  # 近30天日变化数据源

    # 探测最新有数据月(块5 含烟草合计行, 1月=col2...12月=col13): 供所有月度数据截断
    b5_probe = find_row(cd, "块5 ·")
    r_tob_probe = find_row(cd, "含烟草合计", start=b5_probe)
    tob12 = [cd.cell(r_tob_probe, 1 + m).value for m in range(1, 13)]
    LAST_MONTH = 12
    for m in range(11, -1, -1):
        if tob12[m] not in (None, ""):
            LAST_MONTH = m + 1
            break

    b1 = find_row(cd, "块1 ·"); r1 = station_rows(cd, b1)
    b2 = find_row(cd, "块2 ·"); r2 = station_rows(cd, b2)
    b3 = find_row(cd, "块3 ·"); r3 = station_rows(cd, b3)
    b4 = find_row(cd, "块4 ·"); r4 = station_rows(cd, b4)

    def monthly(rmap):
        return {st: [cd.cell(r, 1 + m).value for m in range(1, LAST_MONTH + 1)] for st, r in rmap.items()}

    monthly_data = {
        "tob": monthly(r1), "non": monthly(r2),
        "tobProfit": monthly(r3), "nonProfit": monthly(r4),
    }

    # ── 近30天日变化: 读「按日汇总」每站每日 金额(col2-16)+毛利(col18-32) + 合计(col17/33) ──
    daily30 = build_daily30(dd)

    # 块6: 各站年任务与累计完成(万元)
    b6 = find_row(cd, "块6 ·")
    cr = station_rows(cd, b6)
    stations = []
    for st in STATION_ORDER:
        r = cr[st]
        stations.append({
            "name": st,
            "yearTask": cd.cell(r, 2).value,
            "compTob": cd.cell(r, 3).value,
            "rateTob": cd.cell(r, 4).value,
            "compNon": cd.cell(r, 5).value,
            "rateNon": cd.cell(r, 6).value,
        })

    # 块5: 公司月度 + 环比（保留，供公司视角月度图）
    b5 = find_row(cd, "块5 ·")
    r_tob = find_row(cd, "含烟草合计", start=b5)
    r_non = find_row(cd, "不含烟草合计", start=b5)
    r_tobm = find_row(cd, "含烟草环比", start=b5)
    r_nonm = find_row(cd, "不含烟草环比", start=b5)
    # 月份轴截断到最新有数据月(LAST_MONTH)
    n_m = LAST_MONTH
    company = {
        "monthLabels": [f"{m}月" for m in range(1, n_m + 1)],
        "tob": [cd.cell(r_tob, 1 + m).value for m in range(1, n_m + 1)],
        "non": [cd.cell(r_non, 1 + m).value for m in range(1, n_m + 1)],
        "tobMom": [cd.cell(r_tobm, 1 + m).value for m in range(1, n_m + 1)],
        "nonMom": [cd.cell(r_nonm, 1 + m).value for m in range(1, n_m + 1)],
    }

    # ── 分类明细块：按月汇总各站「各站逐月分类明细」(5品类 + 非油合计) ×15站×12月 金额+毛利 ──
    def cat_titles():
        titles = {}
        for r in range(40, ms.max_row + 1):
            v = ms.cell(r, 1).value
            if isinstance(v, str) and "·" in v:
                name = v.split("·")[0].strip()
                if name in CAT_DETAIL_ALL:
                    titles[name] = r
        return titles

    def cat_matrix(title_row):
        """标题行后按序收集站名行：前 15 行为金额区、后 15 行为毛利区。"""
        st_rows = []
        for r in range(title_row + 1, title_row + 45):
            v = ms.cell(r, 1).value
            if v in STATION_ORDER:
                st_rows.append((v, r))
        amt = {name: [ms.cell(r, 1 + m).value for m in range(1, LAST_MONTH + 1)] for name, r in st_rows[:15]}
        prof = {name: [ms.cell(r, 1 + m).value for m in range(1, LAST_MONTH + 1)] for name, r in st_rows[15:30]}
        return amt, prof

    ct = cat_titles()
    cat_amount, cat_profit = {}, {}
    for c in CAT_DETAIL_ALL:
        amt, prof = cat_matrix(ct[c])
        cat_amount[c] = amt
        cat_profit[c] = prof
    category_detail = {"cats": CAT_DETAIL_ALL, "amount": cat_amount, "profit": cat_profit}

    # 公司品类年度合计（5 品类，万元）—— 由分类明细块汇总，保证与明细一致
    comp5 = []
    for c in CAT_DETAIL:
        total = sum(v or 0 for st in STATION_ORDER for v in cat_amount[c].get(st, [0] * 12))
        comp5.append(round(total / 1e4, 2))
    composition = {"cats": CAT_DETAIL, "tob": comp5}

    # 双口径 + 公司总额: 直接读「按日汇总」(已对齐权威源) 会计年累计，
    # 保证网页年度总额与驾驶舱周/日简报一致(均源自按日汇总)，不再受图表数据源(原始填报汇总)漂移影响。
    _Y0 = cfg["y0"]
    def _ytd(col):
        s = 0.0
        for _r in range(4, dd.max_row + 1):
            _a = dd.cell(_r, 1).value
            if not isinstance(_a, datetime.datetime) or _a < _Y0:
                continue
            _v = _num(dd.cell(_r, col).value)
            if _v is not None:
                s += _v
        return s
    _tob_amt = _ytd(17)                       # Q 公司合计金额(含烟草)
    _tob_pro = _ytd(33)                       # AG 公司合计毛利(含烟草)
    _non_amt = _ytd(34) + _ytd(35) + _ytd(37) # AH汽车+AI便利+AK咖啡 金额(不含烟草)
    _non_pro = _ytd(38) + _ytd(39) + _ytd(44) # AL汽车+AM便利+AR咖啡 毛利(不含烟草)
    totals = {
        "tobAmount": round(_tob_amt / 1e4, 2),
        "nonAmount": round(_non_amt / 1e4, 2),
        "tobProfit": round(_tob_pro / 1e4, 2),
        "nonProfit": round(_non_pro / 1e4, 2),
    }
    dual = {
        "cats": ["金额", "毛利"],
        "tob": [totals["tobAmount"], totals["tobProfit"]],
        "non": [totals["nonAmount"], totals["nonProfit"]],
    }

    # ── 预测模型(目标达成短信) → forecast 节点 ──
    forecast = build_forecast(wb, task_fallback=cfg["task"])

    return {
        "generatedAt": datetime.date.today().isoformat(),
        "year": year,
        "taskTotal": cfg["task"],
        "stations": stations,
        "monthly": monthly_data,
        "company": company,
        "composition": composition,
        "categoryDetail": category_detail,
        "dual": dual,
        "totals": totals,
        "forecast": forecast,
        "daily30": daily30,
        "yoy2026": parse_yoy_sheet(wb) if year == "2027" else None,
    }


def build_daily30(ws):
    """读「按日汇总」近30天(截至最新有数据日) 每站每日 金额/毛利 + 公司合计。
    ws: 「按日汇总」sheet。列结构(1-based): 1=日期, 2-16=15站金额, 17=金额合计,
       18-32=15站毛利, 33=毛利合计, 45=辅助日期。
    """
    import datetime as _dt
    # 找最新有数据的日期行(金额合计>0)
    rows = []
    for r in range(4, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if not isinstance(d, (_dt.datetime, _dt.date)):
            continue
        amt_total = _num(ws.cell(r, 17).value)
        if amt_total is not None and amt_total > 0:
            rows.append(r)
    if not rows:
        return {"dates": [], "amount": {}, "profit": {}}
    end_r = rows[-1]
    start_r = rows[-30] if len(rows) >= 30 else rows[0]

    dates, amount, profit = [], {}, {}
    for i, st in enumerate(STATION_ORDER):
        amount[st] = []
        profit[st] = []
    amount["合计"] = []
    profit["合计"] = []
    for r in range(start_r, end_r + 1):
        d = ws.cell(r, 1).value
        if isinstance(d, _dt.datetime):
            dates.append(d.strftime("%m-%d"))
        elif isinstance(d, _dt.date):
            dates.append(d.strftime("%m-%d"))
        else:
            dates.append("")
        for i, st in enumerate(STATION_ORDER):
            amount[st].append(ws.cell(r, 2 + i).value)
            profit[st].append(ws.cell(r, 18 + i).value)
        amount["合计"].append(ws.cell(r, 17).value)
        profit["合计"].append(ws.cell(r, 33).value)
    return {"dates": dates, "amount": amount, "profit": profit}


def build_forecast(wb, task_fallback=TASK_TOTAL_2026):
    """读取「预测模型」Sheet 参数 + 「目标达成」Sheet 预测短信，供『预测』标签页使用。
    task_fallback: 「预测模型」无「全年任务万」参数时的兜底任务值(2026=1130, 2027=None=待下达)。
    """
    pm = wb["预测模型"]
    L = {}
    for r in range(1, pm.max_row + 1):
        a = pm.cell(r, 1).value
        if isinstance(a, str) and a.strip():
            L[a.strip()] = pm.cell(r, 2).value

    def numv(k):
        v = L.get(k)
        try:
            return None if v is None else round(float(v), 4)
        except (TypeError, ValueError):
            return None

    # 短信正文:「目标达成」找到标题行后, 正文在下一行 A 列(或本行 B 列)
    sms = ""
    da = wb["目标达成"]
    for r in range(1, da.max_row + 1):
        a = da.cell(r, 1).value
        if isinstance(a, str) and "【每日销量预测短信】" in a:
            for cand in (da.cell(r + 1, 1).value, da.cell(r, 2).value):
                if isinstance(cand, str) and len(cand) > 50:
                    sms = cand.strip()
                    break
            break

    as_of = L.get("截至日期")
    if hasattr(as_of, "strftime"):
        as_of = as_of.strftime("%Y-%m-%d")
    pred_end = L.get("预测截止")
    if hasattr(pred_end, "strftime"):
        pred_end = pred_end.strftime("%Y-%m-%d")
    huashan = L.get("华山停业起")
    if hasattr(huashan, "strftime"):
        huashan = huashan.strftime("%Y-%m-%d")

    # 超额% 存的是小数(0.1956), 转成百分比数(19.56)
    def pct(v):
        return None if v is None else round(float(v) * 100, 2)

    # 模型里 毛利率/红线 已是小数形式(0.2408=24.08%), 直接乘100显示
    def pct100(v):
        return None if v is None else round(float(v) * 100, 2)

    return {
        "asOf": as_of,
        "predEnd": pred_end,
        "huashanFrom": huashan,
        "remainingMonths": numv("剩余月数"),
        "task": numv("全年任务万") if numv("全年任务万") is not None else task_fallback,
        "kpis": {
            "累计含去化销售额万": numv("累计含去化销售额万"),
            "自然口径累计万": numv("自然口径累计万"),
            "累计毛利万": numv("累计毛利万"),
            "累计去化额万": numv("累计去化额(烟批)万"),
            "整体毛利率": pct100(numv("整体毛利率")),
            "华山扣除万": numv("华山扣除万"),
            "剩余预测万": numv("剩余预测万"),
            "自然口径全年万": numv("S0自然预测万"),
        },
        "scenarios": {
            "labels": ["全年烟草去化", "剩余可去化", "全口径全年", "公司整体毛利", "期末香烟库存"],
            "A": [numv("X_A最大去化万(23.5%)"), numv("剩余可去化A万"), numv("全口径A万"),
                  numv("公司毛利A万"), numv("期末库存A万(23.5%)")],
            "B": [numv("X_B最大去化万(24%)"), numv("剩余可去化B万"), numv("全口径B万"),
                  numv("公司毛利B万"), numv("期末库存B万(24%)")],
        },
        "overrun": {"A": pct(numv("超额A%")), "B": pct(numv("超额B%"))},
        "extra": {
            "当前香烟库存万": numv("当前库存万"),
            "库存约束去化上限万": numv("库存约束去化上限X_cap万"),
            "库存是否封顶": numv("库存是否封顶(1=是,0=否)"),
            "剩余购进万": numv("剩余购进万"),
            "自然毛利率": pct100(numv("自然毛利率")),
            "红线A": pct100(numv("公司整体毛利红线A")),
            "红线B": pct100(numv("红线B")),
        },
        "sms": sms,
    }


HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>陕西高速延长石油有限责任公司非油销售汇总驾驶舱 · 图表视图</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<script>
if (typeof echarts === 'undefined') {
  document.write('<scr'+'ipt src="https://lib.baomitu.com/echarts/5.5.0/echarts.min.js"><\/scr'+'ipt>');
}
</script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#0d1320; --card:#162032; --ink:#eef2ff; --txt:#d7e0f0; --mut:#8fa1c5;
    --faint:#64748b; --grid:#2b3752; --line:#232f4a;
    --data:#3b82f6; --data2:#60a5fa; --hero:#3b82f6; --accent:#fbbf24;
    --good:#34d399;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--txt);font-family:'Inter','Microsoft YaHei','PingFang SC',sans-serif;
    padding:14px 10px 24px;-webkit-font-smoothing:antialiased}
  .wrap{max-width:none;margin:0 auto;width:100%}
  header{margin-bottom:18px}
  header h1{font-size:22px;font-weight:700;color:var(--ink);letter-spacing:-.02em}
  header p{color:var(--mut);font-size:12.5px;margin-top:5px}
  header p.byline{color:var(--faint);font-size:11px;margin-top:3px}

  /* 页头：左上角双 Logo（陕西交控 + 延长石油），右侧标题组 */
  .headtop{display:flex;align-items:center;gap:22px;flex-wrap:wrap}
  .brand{display:flex;align-items:center;gap:18px;flex:none;padding-right:20px;
    border-right:1px solid var(--line)}
  .brand img{height:84px;width:auto;display:block}
  .headtext{min-width:0}
  @media(max-width:1100px){
    .brand img{height:64px}
    .brand{gap:14px;padding-right:16px}
    .headtop{gap:16px}
  }
  @media(max-width:700px){
    .brand{border-right:none;padding-right:0;gap:10px}
    .brand img{height:48px}
  }
  .upd{color:var(--faint);font-size:11px;margin-top:2px}

  /* 年度切换器 */
  .yearbar{display:flex;gap:8px;margin-top:10px}
  .ytab{padding:7px 18px;border-radius:999px;border:1px solid var(--line);background:var(--card);
    color:var(--mut);font-size:13px;font-weight:600;cursor:pointer;transition:.18s;user-select:none}
  .ytab:hover{color:var(--data);border-color:var(--data2)}
  .ytab.active{background:var(--hero);color:#fff;border-color:var(--hero)}

  /* 联动站点选择器 */
  .stations{display:flex;flex-wrap:wrap;gap:7px;margin:14px 0 18px}
  .chip{padding:6px 13px;border-radius:999px;border:1px solid var(--line);background:var(--card);
    color:var(--mut);font-size:12.5px;cursor:pointer;transition:.18s;user-select:none}
  .chip:hover{border-color:var(--data2);color:var(--data)}
  .chip.active{background:var(--hero);color:#fff;border-color:var(--hero);font-weight:600}
  .chip.all.active{background:var(--accent);border-color:var(--accent);color:#1e293b}

  /* KPI 行 */
  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:16px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:16px 18px}
  .kpi .k{font-size:12px;color:var(--mut);font-weight:500}
  .kpi .v{font-size:27px;font-weight:800;color:var(--ink);margin-top:6px;letter-spacing:-.02em}
  .kpi .v small{font-size:13px;font-weight:600;color:var(--mut);margin-left:3px}
  .kpi .s{font-size:11.5px;color:var(--good);margin-top:3px}

  .grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
  @media(min-width:1500px){.grid{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:980px){.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}}
  .stack{display:flex;flex-direction:column;gap:14px;margin-top:14px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:14px 16px 8px}
  .card.wide{grid-column:1/-1}
  .card h3{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:2px;padding-left:8px;border-left:3px solid var(--data)}
  .card .sub{font-size:11px;color:var(--faint);margin-bottom:8px;padding-left:8px}
  .chart{width:100%;height:390px}
  .chart.tall{height:450px}
  footer{margin-top:18px;color:var(--faint);font-size:11px;text-align:center;line-height:1.6}

  .m-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .m-card{background:var(--bg);border:1px solid var(--line);border-radius:14px;padding:10px 12px 6px}
  .m-card h4{font-size:12.5px;color:var(--ink);font-weight:600;margin-bottom:4px;padding-left:8px;border-left:3px solid var(--data2)}
  .m-chart{width:100%;height:300px}
  .m-table{width:100%;border-collapse:collapse;font-size:12px}
  .m-table th,.m-table td{border:1px solid var(--line);padding:7px 10px;text-align:left}
  .m-table th{background:#22304a;color:#eef2ff;font-weight:600}
  .m-table .r{text-align:right}
  .m-table tr.tot td{background:#3a3118;font-weight:700;color:#fde68a}
  @media(max-width:760px){.m-grid{grid-template-columns:1fr}}

  /* 标签页 */
  .tabs{display:flex;gap:6px;margin:14px 0 4px;border-bottom:1px solid var(--line);padding-bottom:10px}
  .tab{padding:8px 22px;border-radius:10px;cursor:pointer;font-size:13.5px;font-weight:600;
    color:var(--mut);background:var(--card);border:1px solid var(--line);transition:.18s;user-select:none}
  .tab:hover{color:var(--data)}
  .tab.active{background:var(--hero);color:#fff;border-color:var(--hero)}
  .tab.active.fc{background:var(--accent);border-color:var(--accent);color:#1e293b}
  .view{display:none}
  .view.active{display:block}

  /* 预测页 */
  .fc-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:16px 0}
  @media(max-width:980px){.fc-kpis{grid-template-columns:repeat(2,1fr)}}
  .fc-kpi{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:16px 18px}
  .fc-kpi .k{font-size:12px;color:var(--mut);font-weight:500}
  .fc-kpi .v{font-size:25px;font-weight:800;color:var(--ink);margin-top:6px;letter-spacing:-.02em}
  .fc-kpi .v small{font-size:13px;font-weight:600;color:var(--mut);margin-left:3px}
  .fc-kpi .s{font-size:11.5px;color:var(--mut);margin-top:3px}
  .sms-card{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:18px 20px;margin-top:14px}
  .sms-card h3{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:10px;padding-left:8px;border-left:3px solid var(--accent)}
  .sms-card .sms{font-size:12.8px;color:var(--txt);line-height:1.9;white-space:pre-wrap}
  .sms-card .sms b{color:var(--accent);font-weight:700}
  .note{font-size:11px;color:var(--faint);margin-top:8px;line-height:1.7}

  /* ══════════ 销售概览（前一日 · 仅在总览页） ══════════ */
  .so{margin:2px 0 20px}
  .so-strip{display:flex;flex-wrap:wrap;align-items:center;gap:8px 14px;background:var(--card);
    border:1px solid var(--line);border-radius:14px;padding:11px 16px;font-size:12.5px;color:var(--mut)}
  .so-strip b{color:var(--ink);font-weight:700;font-variant-numeric:tabular-nums}
  .so-strip .sep{width:1px;height:13px;background:var(--line);flex:none}
  .so-strip .tagname{color:var(--accent);font-weight:700}
  /* 品类明细小卡：名称 + 金额 + 占比 */
  .so-cats{display:inline-flex;flex-wrap:wrap;align-items:center;gap:6px}
  .so-pill{display:inline-flex;align-items:center;gap:5px;background:rgba(255,255,255,.035);
    border:1px solid var(--line);border-radius:7px;padding:2px 9px;font-size:11.5px;color:var(--mut)}
  .so-pill i{width:7px;height:7px;border-radius:2px;flex:none}
  .so-pill b{color:var(--ink);font-weight:700;font-variant-numeric:tabular-nums}
  .so-pill em{font-style:normal;color:var(--faint);font-size:11px;margin-left:1px}
  .so-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:12px}
  .so-kpi{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:13px 16px}
  .so-kpi .k{font-size:11.5px;color:var(--mut)}
  .so-kpi .v{font-size:23px;font-weight:800;color:var(--ink);margin-top:4px;letter-spacing:-.02em;
    font-variant-numeric:tabular-nums}
  .so-kpi .v small{font-size:12px;color:var(--mut);font-weight:600;margin-left:3px}
  .so-kpi .s{font-size:11px;color:var(--faint);margin-top:3px}
  .so-body{display:grid;grid-template-columns:290px 1fr;gap:12px;margin-top:12px}
  .so-donutbox{background:var(--card);border:1px solid var(--line);border-radius:16px;
    padding:10px 12px 6px;position:relative}
  .so-chart{width:100%;height:186px}
  .so-donut-center{position:absolute;left:0;right:0;top:92px;text-align:center;pointer-events:none}
  .so-donut-center b{display:block;font-size:17px;font-weight:800;color:var(--ink);letter-spacing:-.02em}
  .so-donut-center em{display:block;font-style:normal;font-size:10.5px;color:var(--faint);margin-top:1px}
  .so-catlist{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:12px 18px;
    display:flex;flex-direction:column;justify-content:center}
  .so-catlist .cap{font-size:11px;color:var(--faint);padding-bottom:7px;border-bottom:1px solid var(--line)}
  .so-catlist #soCatList{display:grid;grid-template-rows:repeat(3,auto);grid-auto-flow:column;
    column-gap:34px;align-content:center;flex:1;padding-top:5px}
  .so-crow{display:flex;align-items:center;gap:9px;padding:6px 0;font-size:12.5px;
    border-bottom:1px solid rgba(35,47,74,.5)}
  .so-crow:nth-child(3),.so-crow:last-child{border-bottom:none}
  .so-crow i{width:9px;height:9px;border-radius:3px;flex:none}
  .so-crow span{color:var(--txt);flex:1;white-space:nowrap}
  .so-crow em{font-style:normal;color:var(--faint);font-size:11px;width:46px;text-align:right}
  .so-crow b{color:var(--ink);font-weight:700;font-variant-numeric:tabular-nums;width:88px;text-align:right}
  .so-crow.tot b{color:var(--accent)}
  /* 站卡：一排 5 个 × 3 排（15 站），等行高 + 底部信息对齐 */
  .so-stations{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));grid-auto-rows:1fr;
    gap:12px;margin-top:14px}
  .so-card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:11px 13px;
    cursor:pointer;transition:.16s;display:flex;flex-direction:column;min-width:0}
  .so-card:hover{border-color:var(--data2);transform:translateY(-1px)}
  .so-card.on{border-color:var(--accent);box-shadow:inset 0 0 0 1px var(--accent)}
  .so-c-top{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px}
  .so-c-name{font-size:14px;font-weight:700;color:var(--ink)}
  .so-c-share{font-size:11px;color:var(--faint);font-variant-numeric:tabular-nums}
  .so-c-body{display:flex;gap:10px;align-items:center}
  .so-c-donut{width:84px;height:84px;flex:none}
  .so-c-list{flex:1;min-width:0}
  .so-li{display:flex;align-items:center;gap:6px;font-size:11.5px;line-height:1.8}
  .so-li i{width:7px;height:7px;border-radius:2px;flex:none}
  .so-li span{color:var(--mut);flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .so-li b{color:var(--txt);font-weight:600;font-variant-numeric:tabular-nums}
  .so-li.tot{padding-top:3px;border-top:1px solid var(--line);margin-top:2px}
  .so-li.tot b{color:var(--accent);font-weight:700}
  .so-c-foot{display:flex;flex-wrap:wrap;justify-content:space-between;gap:2px 6px;
    margin-top:auto;padding-top:7px;border-top:1px solid var(--line);
    font-size:11.5px;color:var(--faint)}
  .so-c-foot span{white-space:nowrap}
  .so-c-foot b{color:var(--ink);font-weight:700;font-variant-numeric:tabular-nums}
  .so-bad{color:#fca5a5;font-weight:700}
  .so-na{color:#64748b;font-weight:600}
  .so-off{display:inline-block;color:#fbbf24;background:rgba(251,191,36,.12);
    border:1px solid rgba(251,191,36,.38);border-radius:5px;padding:1px 7px;font-size:10.5px;font-weight:700}
  .so-card.off{border-color:rgba(251,191,36,.25)}
  .so-card.off .so-c-name{color:#e8d9a8}
  .so-empty{background:var(--card);border:1px dashed var(--line);border-radius:14px;padding:14px 16px;
    font-size:12.5px;color:var(--faint);text-align:center}
  @media(max-width:1180px){.so-body{grid-template-columns:1fr}.so-donut-center{top:92px}}
  @media(max-width:980px){.so-kpis{grid-template-columns:repeat(2,1fr)}}
  /* 站卡列数：默认 5（一排 5 个 × 3 排）；窄屏逐级降列，始终等宽等行高 */
  @media(max-width:1300px){.so-stations{grid-template-columns:repeat(4,minmax(0,1fr))}}
  @media(max-width:1060px){.so-stations{grid-template-columns:repeat(3,minmax(0,1fr))}}
  @media(max-width:800px){.so-stations{grid-template-columns:repeat(2,minmax(0,1fr))}}
  @media(max-width:520px){.so-stations{grid-template-columns:1fr}}
  @media(max-width:760px){
    .so-strip{font-size:12px}
    /* 窄屏：品类改单列，否则双列最小内容宽度会把整页撑出横向滚动 */
    .so-catlist #soCatList{grid-template-columns:1fr;grid-template-rows:none;grid-auto-flow:row;
      column-gap:0}
    .so-crow{padding:5px 0}
    .so-crow:nth-child(3){border-bottom:1px solid rgba(35,47,74,.5)}
    .so-crow em{width:42px}
    .so-crow b{width:82px}
    .so-c-body{gap:8px}
    .so-c-donut{width:72px;height:72px}
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="headtop">
      <div class="brand" id="brand"><img src="data:image/png;base64,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" alt="陕西交控"><img src="data:image/png;base64,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" alt="延长石油"></div>
      <div class="headtext">
        <h1>陕西高速延长石油有限责任公司非油销售汇总驾驶舱 · 图表视图</h1>
        <p>联动仪表盘 · 点击站点可查看该站品类明细（数据取自驾驶舱按月汇总/图表数据源）</p>
        <p class="byline">Design by 刘龙伟 · Email: liulongwei@126.com</p>
      </div>
    </div>
    <div class="yearbar" id="yearbar">
      <div class="ytab active" data-year="2027" title="2027会计年: 2026-12-26 ~ 2027-12-25（预留，12-26 启用）">2027.预留</div>
      <div class="ytab" data-year="2026" title="2026会计年: 2025-12-26 ~ 2026-12-25">2026年</div>
    </div>
    <div class="upd" id="upd"></div>
  </header>

  <div class="tabs" id="tabs">
    <div class="tab active" data-view="overview">总览</div>
    <div class="tab fc" data-view="forecast">预测</div>
  </div>

  <!-- ═══ 总览视图 ═══ -->
  <div class="view active" id="view-overview">

  <div class="stations" id="stations"></div>

  <div class="kpis" id="kpis"></div>

  <div style="display:flex;align-items:center;gap:10px;margin:2px 0 12px">
    <span style="font-size:12.5px;font-weight:600;color:var(--mut)">车流与转化 · 单数统计</span>
    <span style="flex:1;height:1px;background:var(--line)"></span>
  </div>
  <div class="kpis" id="tkpis"></div>

  <!-- ── 销售概览（前一日 · 仅总览页；位置：车流与转化 KPI 之下） ── -->
  <section class="so" id="so">
    <div class="so-strip" id="soStrip"></div>
    <div class="so-kpis" id="soKpis"></div>
    <div class="so-body">
      <div class="so-donutbox">
        <div id="soDonut" class="so-chart"></div>
        <div class="so-donut-center"><b id="soDonutV">—</b><em id="soDonutCap">万元 · 总销售</em></div>
      </div>
      <div class="so-catlist">
        <div class="cap">品类构成 · 金额（元） / 占比</div>
        <div id="soCatList"></div>
      </div>
    </div>
    <div class="so-stations" id="soStations"></div>
  </section>

  <div class="grid">
    <div class="card"><h3 id="t1">各站任务完成排行 · 年任务 vs 累计完成（含烟草）</h3>
      <div class="sub">万元 · 深色=当前选中站 · 点击柱体可切换选中并查看明细</div><div id="c1" class="chart"></div></div>
    <div class="card"><h3 id="t2">各站任务完成率（含烟草）</h3>
      <div class="sub">累计完成 ÷ 年任务 · 深色=当前选中站</div><div id="c2" class="chart"></div></div>
    <div class="card"><h3 id="t9">毛利贡献额排名（含烟草 · 年度累计）</h3>
      <div class="sub">万元 · 各站累计毛利降序 · 金色描边=当前选中站</div><div id="c9" class="chart"></div></div>
    <div class="card"><h3 id="t10">毛利率排名（含烟草 · 年度累计）</h3>
      <div class="sub">% · 各站累计毛利率降序 · 红线=23.5% · 金色描边=当前选中站</div><div id="c10" class="chart"></div></div>
    <div class="card"><h3 id="t5">金额构成（含烟草 · 年度累计，万元）</h3>
      <div class="sub">汽车用品 / 便利百货 / 香烟零售 / 烟草批发 / 咖啡</div><div id="c5" class="chart"></div></div>
    <div class="card"><h3 id="t6">双口径对比（金额 / 毛利，万元）</h3>
      <div class="sub">含烟草 vs 不含烟草</div><div id="c6" class="chart"></div></div>
    <div class="card"><h3 id="t11">各站车流量（车次）</h3>
      <div class="sub">汽油 + 柴油堆叠 · 金色描边=当前选中站</div><div id="c11" class="chart"></div></div>
    <div class="card"><h3 id="t12">各站单加升数</h3>
      <div class="sub">升/单 · 汽油 vs 柴油 · 金色描边=当前选中站</div><div id="c12" class="chart"></div></div>
    <div class="card"><h3 id="t13">各站非油转化率</h3>
      <div class="sub">% · 非油笔数 ÷ 总车次 · 虚线=公司均值 · 金色描边=当前选中站</div><div id="c13" class="chart"></div></div>
  </div>
  <div class="stack">
    <div class="card wide"><h3 id="t3">各站分会计月销售趋势（含烟草金额）</h3>
      <div class="sub">元 · 选中单站时该站加粗高亮、其余淡显；选「全公司」则均衡展示</div><div id="c3" class="chart tall"></div></div>
    <div class="card wide"><h3 id="t7">近30天日变化 · 非油销售额（含烟草）</h3>
      <div class="sub">元/日 · 最新30天 · 选「全公司」看公司合计，选中站看该站</div><div id="c7" class="chart"></div></div>
    <div class="card wide"><h3 id="t8">近30天日变化 · 非油毛利（含烟草）</h3>
      <div class="sub">元/日 · 最新30天 · 选「全公司」看公司合计，选中站看该站</div><div id="c8" class="chart"></div></div>
    <div class="card wide"><h3 id="t4">全公司月度非油金额与环比（含烟草）</h3>
      <div class="sub">柱=月度金额(元) · 线=环比(%)</div><div id="c4" class="chart"></div></div>
  </div>

  <div class="card wide" style="margin-top:14px">
    <h3 id="detailTitle">品类构成与明细</h3>
    <div class="sub" id="detailSub" style="margin-bottom:10px">5 品类 × 12 会计月 金额/毛利堆叠、合计表（品类构成见上方「金额构成」卡）</div>
    <div class="m-grid">
      <div class="m-card"><h4 id="dAmtH">品类 × 12 月 金额（元）</h4><div id="dAmt" class="m-chart" style="height:320px"></div></div>
      <div class="m-card"><h4 id="dProfH">品类 × 12 月 毛利（元）</h4><div id="dProf" class="m-chart" style="height:320px"></div></div>
      <div class="m-card" style="grid-column:1/-1"><h4>金额 / 毛利 年度合计（万元）</h4>
        <div style="overflow:auto"><table class="m-table" id="dTable"></table></div></div>
    </div>
  </div>
  </div>

  <!-- ═══ 预测视图 ═══ -->
  <div class="view" id="view-forecast">
    <div class="fc-kpis" id="fcKpis"></div>
    <div class="card wide">
      <h3>全年预测 · 双情景对比（毛利红线 23.5% vs 24%）</h3>
      <div class="sub">单位：万元 · 23.5%=守住公司整体毛利红线情景 · 24%=提高红线情景（烟草去化更保守）</div>
      <div id="fc1" class="chart tall" style="height:400px"></div>
    </div>
    <div class="card wide" style="margin-top:14px">
      <h3>销量任务超额（双情景）</h3>
      <div class="sub">全年任务 1130 万元 · 超额率 %</div>
      <div id="fc2" class="chart" style="height:220px"></div>
    </div>
    <div class="sms-card">
      <h3>预测短信原文（取自驾驶舱「目标达成」）</h3>
      <div class="sms" id="fcSms"></div>
    </div>
  </div>

</div>

<script>
let DATA = null;
let CAT_DETAIL = [];
const C = { ink:'#eef2ff', txt:'#d7e0f0', mut:'#8fa1c5', faint:'#64748b', grid:'#2b3752',
            data:'#3b82f6', data2:'#60a5fa', hero:'#ffffff', sheet:'#162032' };
const PALETTE = ['#ef4444','#3b82f6','#facc15','#06b6d4','#ec4899','#22c55e','#a855f7','#f97316',
                 '#14b8a6','#eab308','#8b5cf6','#10b981','#f43f5e','#0ea5e9','#d946ef'];
const MODAL_PALETTE = ['#ef4444','#3b82f6','#facc15','#06b6d4','#ec4899','#a855f7'];
const fmtW = v => (v==null?'—':(v/10000).toFixed(2)+'万');
const sum12 = arr => (arr||[]).reduce((a,b)=>a+(b||0),0);
const charts = [];
function mk(id, opt){ const el=document.getElementById(id);
  const c=echarts.init(el,null,{renderer:'canvas'}); c.setOption(opt); charts.push(c); return c; }
const num2 = v => (typeof v === 'number') ? Number(v.toFixed(2)) : v;   // 图表数字统一保留两位小数
const fmt2 = v => (typeof v === 'number') ? v.toLocaleString('zh-CN',{minimumFractionDigits:2, maximumFractionDigits:2}) : v; // 千分位+两位小数
const baseTip = {trigger:'axis', backgroundColor:'#ffffff', borderWidth:0, padding:[9,13],
  textStyle:{color:'#111827', fontFamily:'Inter', fontSize:12},
  valueFormatter: v => Array.isArray(v) ? v.map(num2) : num2(v)};
const tipItem = {trigger:'item', backgroundColor:'#ffffff', borderWidth:0, padding:[9,13],
  textStyle:{color:'#111827', fontFamily:'Inter', fontSize:12},
  valueFormatter: v => num2(v)};
const baseGrid = {left:54,right:22,top:30,bottom:46};
const cat = d => ({type:'category',data:d,axisLine:{lineStyle:{color:C.grid}},
  axisTick:{show:false}, axisLabel:{color:C.mut,fontSize:11}});
// 站名类目轴：强制显示全部标签 + 旋转 30° 避免 15 站名重叠截断（排行/完成率图专用）
const catRot = d => ({type:'category',data:d,axisLine:{lineStyle:{color:C.grid}},
  axisTick:{show:false}, axisLabel:{color:C.mut,fontSize:11,interval:0,rotate:30,hideOverlap:false}});
const rankGrid = {left:54,right:24,top:34,bottom:54,containLabel:true};
const val = (name,fmt) => ({type:'value',name:name,nameTextStyle:{color:C.mut,fontSize:10},
  axisLine:{show:false},splitLine:{lineStyle:{color:C.grid}},axisLabel:{color:C.mut,fontSize:11,formatter:fmt}});

let SEL = '全公司';

// ---------- KPI ----------
function renderKPI(){
  const t = DATA.totals, task = DATA.taskTotal;
  const stObj = SEL === '全公司' ? null : DATA.stations.find(s => s.name === SEL);
  const comp   = stObj ? stObj.compTob   : t.tobAmount;
  const rate   = stObj ? (stObj.rateTob!=null ? stObj.rateTob*100 : null)
                       : (task!=null && t.tobAmount!=null ? t.tobAmount/task*100 : null);
  const profit = stObj ? Math.round(sum12(DATA.monthly.tobProfit[SEL])/1e2)/100 : t.tobProfit;
  const profRate = stObj ? (comp ? profit/comp*100 : 0) : (t.tobProfit/t.tobAmount*100);
  const yearTask  = stObj ? stObj.yearTask : task;
  const noTask = (yearTask==null);
  const subt1 = stObj ? (noTask ? '任务待下达 · 累计完成 '+ (comp==null?'—':comp.toFixed(2)) +'万'
                                : `年任务 ${yearTask}万 · 进度 ${rate.toFixed(2)}%`)
                      : (noTask ? '年度任务待下达' : `全年任务 ${task}万 · 进度 ${(t.tobAmount/task*100).toFixed(2)}%`);
  const yoyNote = (YEAR==='2027' && DATA.yoy2026) ? ' · 对比2026同期见月度图虚线' : '';
  const cards = [
    {k: stObj ? '本站累计完成(含烟草)' : '累计完成(含烟草)',
      v: comp==null?'—':comp.toFixed(2), u:'万', s: subt1 + yoyNote},
    {k: '完成率(含烟草)', v: (rate==null||!isFinite(rate))?'—':rate.toFixed(2), u:'%',
      s: noTask ? '任务待下达' : (rate>=100 ? '已达标 ✔' : ((100-rate).toFixed(2)+'% 待完成'))},
    {k: stObj ? '本站累计毛利(含烟草)' : '累计毛利(含烟草)',
      v: profit==null?'—':profit.toFixed(2), u:'万',
      s: '毛利率 ' + profRate.toFixed(2) + '%'},
    {k: stObj ? '本站年任务' : '全年任务',
      v: noTask ? '待下达' : yearTask, u: noTask?'':'万',
      s: noTask ? '任务下达后自动更新' : (stObj ? '本站任务盘（不含烟草去化）' : '烟草去化已计入')},
  ];
  document.getElementById('kpis').innerHTML = cards.map(c=>
    `<div class="kpi"><div class="k">${c.k}</div><div class="v">${c.v}<small>${c.u}</small></div><div class="s">${c.s}</div></div>`).join('');
}

// ---------- 车流与转化（单数统计） ----------
function _cnt(){ return (DATA.counts) || { totals:null, stations:[], monthly:{}, daily30:{dates:[],tot:[],nb:[]} }; }
function _stripInt(v){ return (v==null||v==='') ? '—' : Number(v).toLocaleString('zh-CN'); }
function renderTrafficKPI(){
  const cd = _cnt();
  const st = (SEL === '全公司') ? null : (cd.stations||[]).find(x => x.name === SEL);
  const t = st || cd.totals || {};
  const pre = st ? (SEL + ' ') : '';
  const cards = [
    {k: pre + '总服务车次', v: _stripInt(t.tot), u:'车次',
     s: '汽油 ' + _stripInt(t.gas) + ' · 柴油 ' + _stripInt(t.die)},
    {k: pre + '总加油升数', v: (t.vol ? (t.vol/1e4).toFixed(1) : '—'), u:'万升',
     s: '汽油 ' + (t.gasL ? (t.gasL/1e4).toFixed(1) : '—') + '万 · 柴油 ' + (t.dieL ? (t.dieL/1e4).toFixed(1) : '—') + '万' + (((t.gasL||0)+(t.dieL||0)) ? '<br>约 ' + Math.round(((t.gasL||0)*0.738 + (t.dieL||0)*0.828)/1000).toLocaleString('zh-CN') + ' 吨（密度按汽油0.738、柴油0.828折算）' : '')},
    {k: pre + '汽油单加升数', v: (t.gasPer==null?'—':t.gasPer), u:'升/单',
     s: '汽油升数 ÷ 汽油车次'},
    {k: pre + '柴油单加升数', v: (t.diePer==null?'—':t.diePer), u:'升/单',
     s: '柴油升数 ÷ 柴油车次'},
  ];
  document.getElementById('tkpis').innerHTML = cards.map(c =>
    `<div class="kpi"><div class="k">${c.k}</div><div class="v">${c.v}<small>${c.u}</small></div><div class="s">${c.s}</div></div>`).join('');
}
function renderTraffic(){
  const cd = _cnt(); if(!(cd.stations||[]).length) return;
  const arr = cd.stations.slice().sort((a,b)=>a.tot-b.tot);
  const names = arr.map(x=>x.name);
  const gas = arr.map(x=>x.gas), die = arr.map(x=>x.die);
  mk('c11', {
    tooltip: Object.assign({}, baseTip, {formatter: ps => {
      let h = '<b>' + ps[0].name + '</b><br>', t = 0;
      ps.forEach(x=>{ if(x.value==null) return; t += x.value; h += x.marker + x.seriesName + '：' + fmt2(x.value) + ' 车次<br>'; });
      return h + '<span style="color:#8fa1c5">合计 ' + fmt2(t) + ' 车次</span>';
    }}),
    legend:{data:['汽油车次','柴油车次'], textStyle:{color:C.txt,fontSize:11.5,fontWeight:600}, top:2, itemWidth:12, itemHeight:8},
    grid:{left:66, right:56, top:34, bottom:18},
    xAxis: Object.assign(val('车次'), {axisLabel:{color:C.mut,fontSize:10,formatter:v=>(v/1000)+'k'}}),
    yAxis: Object.assign(cat(names), {axisLabel:{color:C.txt, fontSize:11}}),
    series:[
      {name:'汽油车次', type:'bar', stack:'v', barWidth:13, color:'#ef4444',
       itemStyle:{color:'#ef4444'},
       data: gas.map((x,i)=>({value:x, itemStyle:{color:'#ef4444',
         borderColor:(SEL!=='全公司'&&names[i]===SEL)?'#fbbf24':'transparent',
         borderWidth:(SEL!=='全公司'&&names[i]===SEL)?2:0}}))},
      {name:'柴油车次', type:'bar', stack:'v', barWidth:13, color:'#60a5fa',
       itemStyle:{color:'#60a5fa'},
       data: die.map((x,i)=>({value:x, itemStyle:{color:'#60a5fa',
         borderColor:(SEL!=='全公司'&&names[i]===SEL)?'#fbbf24':'transparent',
         borderWidth:(SEL!=='全公司'&&names[i]===SEL)?2:0,
         borderRadius:[0,3,3,0]}})),
       label:{show:true, position:'right', color:C.mut, fontSize:9.5,
              formatter:p=>{const t=(gas[p.dataIndex]+die[p.dataIndex]); return t>0?(t/1000).toFixed(0)+'k':'';}}}
    ]
  });
}
function renderFuelPer(){
  const cd = _cnt(); if(!(cd.stations||[]).length) return;
  const arr = cd.stations.slice().sort((a,b)=>a.diePer-b.diePer);
  const names = arr.map(x=>x.name);
  const gap = arr.map(x=>x.gasPer), dap = arr.map(x=>x.diePer);
  mk('c12', {
    tooltip: Object.assign({}, baseTip, {valueFormatter:v=>num2(v)+' 升/单'}),
    legend:{data:['汽油单加','柴油单加'], textStyle:{color:C.txt,fontSize:11.5,fontWeight:600}, top:2, itemWidth:12, itemHeight:8},
    grid:{left:66, right:48, top:34, bottom:18},
    xAxis: Object.assign(val('升/单'), {axisLabel:{color:C.mut,fontSize:10}}),
    yAxis: Object.assign(cat(names), {axisLabel:{color:C.txt, fontSize:11}}),
    series:[
      {name:'汽油单加', type:'bar', barWidth:8, barGap:'15%', color:'#ef4444',
       itemStyle:{color:'#ef4444'},
       data: gap.map((x,i)=>({value:x, itemStyle:{color:'#ef4444',
         borderColor:(SEL!=='全公司'&&names[i]===SEL)?'#fbbf24':'transparent',
         borderWidth:(SEL!=='全公司'&&names[i]===SEL)?2:0,
         borderRadius:[0,3,3,0]}}))},
      {name:'柴油单加', type:'bar', barWidth:8, color:'#34d399',
       itemStyle:{color:'#34d399'},
       data: dap.map((x,i)=>({value:x, itemStyle:{color:'#34d399',
         borderColor:(SEL!=='全公司'&&names[i]===SEL)?'#fbbf24':'transparent',
         borderWidth:(SEL!=='全公司'&&names[i]===SEL)?2:0,
         borderRadius:[0,3,3,0]}})),
       label:{show:true, position:'right', color:C.mut, fontSize:9.5, formatter:p=>num2(p.value)}}
    ]
  });
}
function renderConv(){
  const cd = _cnt(); if(!(cd.stations||[]).length) return;
  const avg = (cd.totals && cd.totals.conv) || 0;
  const arr = cd.stations.slice().sort((a,b)=>a.conv-b.conv);
  const names = arr.map(x=>x.name), convs = arr.map(x=>x.conv);
  const hi = Math.max(20, Math.max.apply(null, convs));
  mk('c13', {
    tooltip: Object.assign({}, baseTip, {valueFormatter:v=>num2(v)+'%'}),
    grid:{left:66, right:52, top:14, bottom:18},
    xAxis: Object.assign(val('转化率', v=>v+'%'), {max: Math.ceil(hi*1.18)}),
    yAxis: Object.assign(cat(names), {axisLabel:{color:C.txt, fontSize:11}}),
    series:[{
      type:'bar', barWidth:13,
      data: convs.map((x,i)=>({value:x, itemStyle:{
        color: (x>=30 ? '#34d399' : (x<15 ? '#ef4444' : '#3b82f6')),
        borderColor: (SEL!=='全公司' && names[i]===SEL) ? '#fbbf24' : 'transparent',
        borderWidth: (SEL!=='全公司' && names[i]===SEL) ? 2 : 0,
        borderRadius:[0,3,3,0]}})),
      label:{show:true, position:'right', color:C.txt, fontSize:10, formatter:p=>num2(p.value)+'%'},
      markLine:{symbol:'none', silent:true, data:[{xAxis:avg, lineStyle:{color:'#fbbf24', type:'dashed', width:1.5},
        label:{formatter:'均值 '+num2(avg)+'%', color:'#fbbf24', fontSize:9.5, position:'end'}}]}
    }]
  });
}// ---------- 图1 任务（公司=排行；选中站=该站任务进度） ----------
function renderRank(){
  const t = document.getElementById('t1');
  if(SEL === '全公司'){
    const st = DATA.stations.slice().sort((a,b)=>(b.compTob||0)-(a.compTob||0));
    const hasTask = st.some(s=>s.yearTask!=null);
    const names = st.map(s=>s.name), task = st.map(s=>s.yearTask), done = st.map(s=>s.compTob);
    const colors = st.map(s=> s.name===SEL ? C.hero : C.data2);
    if(hasTask){
      t.textContent = '各站任务完成排行 · 年任务 vs 累计完成（含烟草）';
      mk('c1',{tooltip:baseTip,legend:{data:['年任务','累计完成'],textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
        grid:rankGrid, xAxis:catRot(names), yAxis:val('万元'),
        series:[
          {name:'年任务',type:'bar',data:task,barWidth:13,itemStyle:{color:'#475569',borderRadius:[4,4,0,0]}},
          {name:'累计完成',type:'bar',data:done,barWidth:13,
            itemStyle:{color:p=>colors[p.dataIndex]!==undefined?colors[p.dataIndex]:C.data,borderRadius:[4,4,0,0]},
            label:{show:true,position:'top',color:C.mut,fontSize:9,formatter:p=>Number(p.value).toFixed(2)}}
        ]});
    }else{
      t.textContent = '各站累计完成排行（含烟草 · 任务待下达）';
      mk('c1',{tooltip:baseTip, grid:rankGrid, xAxis:catRot(names), yAxis:val('万元'),
        series:[{name:'累计完成',type:'bar',data:done,barWidth:18,
          itemStyle:{color:C.data,borderRadius:[4,4,0,0]},
          label:{show:true,position:'top',color:C.mut,fontSize:9,formatter:p=>Number(p.value).toFixed(2)}}]});
    }
    return;
  }
  const s = DATA.stations.find(x=>x.name===SEL);
  if(s.yearTask==null){
    t.textContent = SEL + ' · 累计完成（含烟草 · 任务待下达）';
    mk('c1',{tooltip:baseTip, grid:baseGrid, xAxis:cat(['累计完成']), yAxis:val('万元'),
      series:[{type:'bar',data:[s.compTob||0],barWidth:44,
        itemStyle:{color:C.hero,borderRadius:[4,4,0,0]},
        label:{show:true,position:'top',color:C.mut,fontSize:10,formatter:p=>Number(p.value).toFixed(2)}}]});
    return;
  }
  t.textContent = SEL + ' · 任务进度（年任务 vs 累计完成，万元）';
  mk('c1',{tooltip:baseTip, grid:baseGrid, xAxis:cat(['年任务','累计完成']), yAxis:val('万元'),
    series:[{type:'bar',data:[s.yearTask, s.compTob],barWidth:44,
      itemStyle:{color:p=>p.dataIndex===0?'#475569':C.hero,borderRadius:[4,4,0,0]},
      label:{show:true,position:'top',color:C.mut,fontSize:10,formatter:p=>Number(p.value).toFixed(2)}}]});
}
// ---------- 图2 完成率（公司=排行；选中站=含/不含烟草对比） ----------
function renderRate(){
  const t = document.getElementById('t2');
  if(SEL === '全公司'){
    t.textContent = '各站任务完成率（含烟草）';
    const st = DATA.stations.slice().sort((a,b)=>(b.rateTob||0)-(a.rateTob||0));
    const hasTask = st.some(s=>s.rateTob!=null);
    const names = st.map(s=>s.name);
    if(!hasTask){
      t.textContent = '各站任务完成率（任务待下达）';
      mk('c2',{tooltip:baseTip, grid:rankGrid, xAxis:catRot(names), yAxis:val('完成率',v=>v+'%'),
        graphic:{type:'text',left:'center',top:'middle',
          style:{text:'任务下达后展示各站完成率',fill:'#8fa1c5',fontSize:14}},
        series:[]});
      return;
    }
    const rate = st.map(s=>+(s.rateTob*100).toFixed(2));
    const colors = st.map(s=> s.name===SEL ? C.hero : C.data2);
    mk('c2',{tooltip:{...baseTip,formatter:p=>p[0].name+'：'+Number(p[0].value).toFixed(2)+'%'},grid:rankGrid,
      xAxis:catRot(names), yAxis:val('完成率',v=>v+'%'),
      series:[{type:'bar',data:rate.map((v,i)=>({value:v,itemStyle:{color:colors[i],borderRadius:[4,4,0,0]}})),
        barWidth:18, label:{show:true,position:'top',color:C.mut,fontSize:9,formatter:p=>Number(p.value).toFixed(2)+'%'}}]});
    return;
  }
  const s = DATA.stations.find(x=>x.name===SEL);
  if(s.rateTob==null){
    t.textContent = SEL + ' · 完成率（任务待下达）';
    mk('c2',{tooltip:baseTip, grid:baseGrid, xAxis:cat(['含烟草','不含烟草']), yAxis:val('完成率',v=>v+'%'),
      graphic:{type:'text',left:'center',top:'middle',
        style:{text:'任务下达后展示完成率',fill:'#8fa1c5',fontSize:14}},
      series:[]});
    return;
  }
  t.textContent = SEL + ' · 完成率（含烟草 vs 不含烟草）';
  mk('c2',{tooltip:{...baseTip,formatter:p=>p[0].name+'：'+Number(p[0].value).toFixed(2)+'%'},grid:baseGrid,
    xAxis:cat(['含烟草','不含烟草']), yAxis:val('完成率',v=>v+'%'),
    series:[{type:'bar',data:[+(s.rateTob*100).toFixed(2), +(s.rateNon*100).toFixed(2)],barWidth:44,
      itemStyle:{color:p=>p.dataIndex===0?C.hero:C.data2,borderRadius:[4,4,0,0]},
      label:{show:true,position:'top',color:C.mut,fontSize:10,formatter:p=>Number(p.value).toFixed(2)+'%'}}]});
}
// ---------- 图3 趋势(联动) ----------
function renderTrend(){
  const t = document.getElementById('t3');
  t.textContent = SEL==='全公司' ? '各站分会计月销售趋势（含烟草金额）' : SEL + ' · 月度销售趋势（含烟草金额）';
  const months = DATA.company.monthLabels;
  const m = DATA.monthly.tob;
  const series = DATA.stations.map((s,i)=>{
    const isSel = (SEL===s.name);
    const dim = (SEL!=='全公司' && !isSel);
    return {name:s.name, type:'line', smooth:true, symbol:'none',
      data:m[s.name],
      lineStyle:{width:isSel?3.4:1.4, color:isSel?C.hero:PALETTE[i%PALETTE.length], opacity:dim?0.18:1},
      itemStyle:{color:isSel?C.hero:PALETTE[i%PALETTE.length], opacity:dim?0.18:1},
      z:isSel?10:1, emphasis:{focus:'series'}};
  });
  // 2026 同期叠加(虚线) —— 仅 2027 视图且存在同比块时
  const yoy = DATA.yoy2026;
  if(yoy && YEAR==='2027'){
    if(SEL==='全公司' && yoy.companyTob){
      t.textContent = '各站分会计月销售趋势（含烟草 · 金色虚线=2026同期公司合计）';
      series.push({name:'2026同期·公司合计', type:'line', smooth:true, symbol:'none',
        data:yoy.companyTob,
        lineStyle:{width:2.6, color:'#fbbf24', type:'dashed'},
        itemStyle:{color:'#fbbf24'}, z:11, emphasis:{focus:'series'}});
    }else if(SEL!=='全公司' && yoy.monthlyTob && yoy.monthlyTob[SEL]){
      t.textContent = SEL + ' · 月度销售趋势（含烟草 · 金色虚线=2026同期）';
      series.push({name:'2026同期·'+SEL, type:'line', smooth:true, symbol:'none',
        data:yoy.monthlyTob[SEL],
        lineStyle:{width:2.6, color:'#fbbf24', type:'dashed'},
        itemStyle:{color:'#fbbf24'}, z:11, emphasis:{focus:'series'}});
    }
  }
  mk('c3',{tooltip:baseTip, legend:{show:false}, grid:{left:60,right:24,top:18,bottom:40},
    xAxis:cat(months), yAxis:val('元'), series});
}
// ---------- 图4 月度（公司=公司月度；选中站=该站月度） ----------
function renderCompany(){
  const t = document.getElementById('t4');
  const months = DATA.company.monthLabels;
  if(SEL === '全公司'){
    const yoy = DATA.yoy2026;
    const hasYoy = (YEAR==='2027' && yoy && yoy.companyTob);
    t.textContent = hasYoy ? '全公司月度非油金额与环比（含烟草 · 虚线=2026同期）'
                           : '全公司月度非油金额与环比（含烟草）';
    const c = DATA.company;
    const legendData = ['含烟草金额','含烟草环比'];
    const series = [
      {name:'含烟草金额',type:'bar',data:c.tob,barWidth:20,itemStyle:{color:C.data,borderRadius:[4,4,0,0]}},
      {name:'含烟草环比',type:'line',yAxisIndex:1,smooth:true,data:c.tobMom,
        lineStyle:{width:2,color:C.accent},itemStyle:{color:C.accent}}
    ];
    if(hasYoy){
      legendData.push('2026同期');
      series.push({name:'2026同期',type:'line',smooth:true,symbol:'circle',symbolSize:5,
        data:yoy.companyTob,
        lineStyle:{width:2.2,color:'#fbbf24',type:'dashed'},itemStyle:{color:'#fbbf24'}});
    }
    mk('c4',{tooltip:baseTip, legend:{data:legendData,textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
      grid:Object.assign({},baseGrid,{right:56}), xAxis:cat(months),
      yAxis:[val('元'),{type:'value',name:'环比',nameTextStyle:{color:C.mut,fontSize:10},
        axisLine:{show:false},splitLine:{show:false},axisLabel:{color:C.mut,fontSize:11,formatter:v=>(v*100).toFixed(0)+'%'}}],
      series});
    return;
  }
  t.textContent = SEL + ' · 月度非油金额与环比（含烟草）';
  const arr = DATA.monthly.tob[SEL] || [];
  const mom = arr.map((v,i)=>{ if(i===0) return null; const p=arr[i-1]; if(!p) return null; return (v-p)/p; });
  mk('c4',{tooltip:baseTip, legend:{data:['月度金额','环比'],textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
    grid:Object.assign({},baseGrid,{right:56}), xAxis:cat(months),
    yAxis:[val('元'),{type:'value',name:'环比',nameTextStyle:{color:C.mut,fontSize:10},
      axisLine:{show:false},splitLine:{show:false},axisLabel:{color:C.mut,fontSize:11,formatter:v=>v==null?'':(v*100).toFixed(0)+'%'}}],
    series:[
      {name:'月度金额',type:'bar',data:arr,barWidth:20,itemStyle:{color:C.data,borderRadius:[4,4,0,0]}},
      {name:'环比',type:'line',yAxisIndex:1,smooth:true,data:mom,
        lineStyle:{width:2,color:C.accent},itemStyle:{color:C.accent}}
    ]});
}
// ---------- 图5 构成（公司=全公司5类；选中站=该站5类） ----------
function renderComp(){
  const t = document.getElementById('t5');
  let names, vals, profits;
  if(SEL === '全公司'){
    t.textContent = '金额构成（含烟草 · 年度累计，万元）';
    names = DATA.composition.cats; vals = DATA.composition.tob;
    // 公司级各品类毛利 = 各站该品类年度毛利合计
    profits = names.map(c=> Math.round(
      DATA.stations.reduce((s,st)=> s+sum12(DATA.categoryDetail.profit[c][st.name]||[]), 0)/1e2)/100);
  }else{
    t.textContent = SEL + ' · 品类构成（年度累计，万元）';
    names = CAT_DETAIL;
    vals = CAT_DETAIL.map(c=> Math.round(sum12(DATA.categoryDetail.amount[c][SEL])/1e2)/100);
    profits = CAT_DETAIL.map(c=> Math.round(sum12(DATA.categoryDetail.profit[c][SEL])/1e2)/100);
  }
  const total = vals.reduce((s,v)=>s+(v||0),0) || 1;
  const data = names.map((n,i)=>{
    const amt = vals[i], prof = profits[i];
    const rate = amt ? (prof/amt*100) : 0;
    const pct  = amt ? (amt/total*100) : 0;
    return {name:n, value:amt, profit:prof, rate:rate, pct:pct};
  });
  mk('c5',{tooltip:{...tipItem, formatter:p=>{
      const d=p.data;
      return d.name+'<br>金额：'+fmt2(d.value)+' 万<br>占全年：'+Number(d.pct).toFixed(1)+'%<br>毛利：'+fmt2(d.profit)+' 万<br>毛利率：'+Number(d.rate).toFixed(2)+'%';
    }},
    legend:{orient:'vertical',left:'left',textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600}},
    color:MODAL_PALETTE,
    series:[{type:'pie',radius:['42%','68%'],center:['62%','54%'],
      label:{color:C.txt,fontSize:10.5,formatter:p=>{
        const d=p.data;
        return d.name+'\n'+num2(d.value)+'万 · 占'+Number(d.pct).toFixed(1)+'% · 毛利 '+num2(d.profit)+'万 ('+Number(d.rate).toFixed(1)+'%)';
      }},
      labelLine:{lineStyle:{color:C.faint}},
      data:data}]});
}
// ---------- 图6 双口径（公司=公司；选中站=该站） ----------
function renderDual(){
  const t = document.getElementById('t6');
  let d;
  if(SEL === '全公司'){
    t.textContent = '双口径对比（金额 / 毛利，万元）';
    d = DATA.dual;
  }else{
    t.textContent = SEL + ' · 双口径对比（金额 / 毛利，万元）';
    const m = DATA.monthly;
    d = {cats:['金额','毛利'],
      non:[ Math.round(sum12(m.non[SEL])/1e2)/100, Math.round(sum12(m.nonProfit[SEL])/1e2)/100 ],
      tob:[ Math.round(sum12(m.tob[SEL])/1e2)/100, Math.round(sum12(m.tobProfit[SEL])/1e2)/100 ]};
  }
  mk('c6',{tooltip:baseTip, legend:{data:['不含烟草','含烟草'],textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
    grid:baseGrid, xAxis:cat(d.cats), yAxis:val('万元'),
    series:[
      {name:'不含烟草',type:'bar',data:d.non,barWidth:34,itemStyle:{color:C.data2,borderRadius:[4,4,0,0]}},
      {name:'含烟草',type:'bar',data:d.tob,barWidth:34,itemStyle:{color:C.data,borderRadius:[4,4,0,0]}}
    ]});
}

// ---------- 近30天日变化(联动) ----------
function renderDaily30(){
  const d = DATA.daily30 || {dates:[], amount:{}, profit:{}};
  if(!d.dates.length){ return; }
  const key = SEL === '全公司' ? '合计' : SEL;
  const amtArr = d.amount[key] || [];
  const profArr = d.profit[key] || [];
  const t7 = document.getElementById('t7');
  const t8 = document.getElementById('t8');
  t7.textContent = '近30天日变化 · 非油销售额（含烟草）' + (SEL==='全公司' ? '' : ' · ' + SEL);
  t8.textContent = '近30天日变化 · 非油毛利（含烟草）' + (SEL==='全公司' ? '' : ' · ' + SEL);
  // 7日均线
  const ma7 = arr => arr.map((_,i)=>{
    if(i<6) return null;
    let s=0; for(let j=i-6;j<=i;j++) s+=(arr[j]||0);
    return Math.round(s/7*100)/100;
  });
  const mkDaily = (id, arr, color) => {
    mk(id,{tooltip:{...baseTip,formatter:ps=>{
        const p=ps[0]; let html=p.name+'<br>';
        ps.forEach(x=>{ html+=x.marker+x.seriesName+'：'+fmt2(x.value==null?0:x.value)+' 元<br>'; });
        return html;
      }},
      legend:{data:['日值','7日均线'],textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
      grid:{left:70,right:26,top:34,bottom:46}, xAxis:cat(d.dates),
      yAxis:val('元'),
      series:[
        {name:'日值',type:'line',smooth:true,symbol:'circle',symbolSize:4,data:arr,
          lineStyle:{width:2,color:color},itemStyle:{color:color},areaStyle:{color:color,opacity:0.08}},
        {name:'7日均线',type:'line',smooth:true,symbol:'none',data:ma7(arr),
          lineStyle:{width:1.6,color:'#fbbf24',type:'dashed'},itemStyle:{color:'#fbbf24'}}
      ]});
  };
  mkDaily('c7', amtArr, C.data);
  mkDaily('c8', profArr, C.data2);
}

// ---------- 毛利贡献额 / 毛利率 排名榜(公司视角,横向条形) ----------
function renderRank2(){
  const prof = DATA.stations.map(s=>({
    name:s.name,
    amt: Math.round(sum12(DATA.monthly.tob[s.name])/1e2)/100,          // 万元
    profit: Math.round(sum12(DATA.monthly.tobProfit[s.name])/1e2)/100, // 万元
  }));
  prof.forEach(p=>{ p.rate = p.amt ? +(p.profit/p.amt*100).toFixed(2) : 0; });
  const byProfit = prof.slice().sort((a,b)=>b.profit-a.profit);
  const byRate   = prof.slice().sort((a,b)=>b.rate-a.rate);

  // 毛利贡献额榜(横向)
  mk('c9',{tooltip:{...tipItem,valueFormatter:v=>num2(v)+' 万'},
    grid:{left:70,right:44,top:16,bottom:26},
    xAxis:{type:'value',axisLine:{show:false},splitLine:{lineStyle:{color:C.grid}},
      axisLabel:{color:C.mut,fontSize:11,formatter:v=>v+'万'}},
    yAxis:{type:'category',inverse:true,data:byProfit.map(x=>x.name),
      axisLine:{show:false},axisTick:{show:false},axisLabel:{color:C.txt,fontSize:11.5}},
    series:[{type:'bar',data:byProfit.map((x,i)=>({
        value:x.profit, itemStyle:{color:i===0?'#fbbf24':(i===1?'#60a5fa':'#3b82f6'),
          borderColor:(SEL!=='全公司'&&x.name===SEL)?'#fbbf24':'transparent',
          borderWidth:(SEL!=='全公司'&&x.name===SEL)?2:0,
          borderRadius:[0,4,4,0]}})),
      barWidth:16,
      label:{show:true,position:'right',color:C.txt,fontSize:10.5,formatter:p=>num2(p.value)+'万'}}]});

  // 毛利率榜(横向 + 23.5%红线)
  mk('c10',{tooltip:{...tipItem,valueFormatter:v=>num2(v)+'%'},
    grid:{left:70,right:44,top:16,bottom:26},
    xAxis:{type:'value',axisLine:{show:false},splitLine:{lineStyle:{color:C.grid}},
      axisLabel:{color:C.mut,fontSize:11,formatter:v=>v+'%'},
      max: Math.max(...byRate.map(x=>x.rate), 23.5)*1.12},
    yAxis:{type:'category',inverse:true,data:byRate.map(x=>x.name),
      axisLine:{show:false},axisTick:{show:false},axisLabel:{color:C.txt,fontSize:11.5}},
    series:[{type:'bar',data:byRate.map((x,i)=>({
        value:x.rate,
        itemStyle:{color: x.rate>=23.5 ? '#34d399' : '#ef4444',
          borderColor:(SEL!=='全公司'&&x.name===SEL)?'#fbbf24':'transparent',
          borderWidth:(SEL!=='全公司'&&x.name===SEL)?2:0,
          borderRadius:[0,4,4,0]}})),
      barWidth:16,
      label:{show:true,position:'right',color:C.txt,fontSize:10.5,formatter:p=>num2(p.value)+'%'},
      markLine:{symbol:'none',data:[{xAxis:23.5,lineStyle:{color:'#fbbf24',type:'dashed',width:1.5},
        label:{formatter:'红线 23.5%',color:'#fbbf24',fontSize:10,position:'end'}}]}}]});
}

// ---------- 品类明细（内嵌到主页，4 块：金额堆叠/毛利堆叠/构成饼/合计表） ----------
function renderDetail(){
  const isCo = SEL === '全公司';
  const t = document.getElementById('detailTitle');
  const sub = document.getElementById('detailSub');
  const aH = document.getElementById('dAmtH');
  const pH = document.getElementById('dProfH');
  const months = DATA.company.monthLabels;
  const amt = {}, prof = {};
  if (isCo) {
    t.textContent = '全公司 · 品类构成与明细';
    sub.textContent = '5 品类 × 12 会计月 金额/毛利堆叠（公司合计）、年度构成、合计表';
    aH.textContent = '品类 × 12 月 金额（元，公司合计）';
    pH.textContent = '品类 × 12 月 毛利（元，公司合计）';
    for (const c of CAT_DETAIL) {
      const aArr = new Array(12).fill(0), pArr = new Array(12).fill(0);
      for (const st of DATA.stations) {
        const a = DATA.categoryDetail.amount[c][st.name] || [];
        const p = DATA.categoryDetail.profit[c][st.name] || [];
        for (let m = 0; m < 12; m++) { aArr[m] += (a[m] || 0); pArr[m] += (p[m] || 0); }
      }
      amt[c] = aArr; prof[c] = pArr;
    }
  } else {
    t.textContent = SEL + ' · 品类构成与明细';
    sub.textContent = '5 品类 × 12 会计月 金额/毛利堆叠、年度构成、合计表';
    aH.textContent = '品类 × 12 月 金额（元）';
    pH.textContent = '品类 × 12 月 毛利（元）';
    for (const c of CAT_DETAIL) {
      amt[c]  = DATA.categoryDetail.amount[c][SEL] || [];
      prof[c] = DATA.categoryDetail.profit[c][SEL]  || [];
    }
  }
  // 各品类全年合计 & 占总额比例（用于图例标注 + 合计表）
  const annA = {}, annP = {};
  CAT_DETAIL.forEach(c=>{ annA[c]=sum12(amt[c]); annP[c]=sum12(prof[c]); });
  const totA = CAT_DETAIL.reduce((s,c)=>s+annA[c],0) || 1;
  const totP = CAT_DETAIL.reduce((s,c)=>s+annP[c],0) || 1;
  const pctA = c => (annA[c]/totA*100);
  const pctP = c => (annP[c]/totP*100);

  // 金额堆叠：图例标全年占比%，tooltip 补当月占比% + 当月合计
  const totName = isCo ? '公司合计' : '本站合计';
  mk('dAmt',{tooltip:{...baseTip, formatter:ps=>{
      const idx = ps[0].dataIndex;
      const mSum = CAT_DETAIL.reduce((s,c)=>s+(amt[c][idx]||0),0);
      const mTot = mSum || 1;
      let h = ps[0].name+'<br>';
      ps.forEach(x=>{ const v=x.value||0; h+=x.marker+x.seriesName+'：'+fmt2(v)+' 元（当月 '+(v/mTot*100).toFixed(1)+'%）'+'<br>'; });
      h += '<div style="margin-top:5px;padding-top:5px;border-top:1px solid #e5e7eb;font-weight:700;color:#111827">'
         + totName+'：'+fmt2(mSum)+' 元</div>';
      return h;
    }},
    legend:{textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600}, top:6,
      formatter:name=>name+' '+pctA(name).toFixed(1)+'%'},
    grid:{left:54,right:22,top:38,bottom:46},
    xAxis:cat(months), yAxis:val('元'),
    series:CAT_DETAIL.map((c,i)=>({name:c,type:'bar',stack:'a',data:amt[c],barMaxWidth:24,
      itemStyle:{color:MODAL_PALETTE[i]}}))});
  // 毛利堆叠：图例标全年占比%，tooltip 补当月占比% + 当月合计
  mk('dProf',{tooltip:{...baseTip, formatter:ps=>{
      const idx = ps[0].dataIndex;
      const mSum = CAT_DETAIL.reduce((s,c)=>s+(prof[c][idx]||0),0);
      const mTot = mSum || 1;
      let h = ps[0].name+'<br>';
      ps.forEach(x=>{ const v=x.value||0; h+=x.marker+x.seriesName+'：'+fmt2(v)+' 元（当月 '+(v/mTot*100).toFixed(1)+'%）'+'<br>'; });
      h += '<div style="margin-top:5px;padding-top:5px;border-top:1px solid #e5e7eb;font-weight:700;color:#111827">'
         + totName+'：'+fmt2(mSum)+' 元</div>';
      return h;
    }},
    legend:{textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600}, top:6,
      formatter:name=>name+' '+pctP(name).toFixed(1)+'%'},
    grid:{left:54,right:22,top:38,bottom:46},
    xAxis:cat(months), yAxis:val('元'),
    series:CAT_DETAIL.map((c,i)=>({name:c,type:'bar',stack:'p',data:prof[c],barMaxWidth:24,
      itemStyle:{color:MODAL_PALETTE[i]}}))});

  const ta = Math.round(totA/1e2)/100;
  const tp = Math.round(totP/1e2)/100;
  const rows = CAT_DETAIL.map(c=>{
    const a = Math.round(annA[c]/1e2)/100, p = Math.round(annP[c]/1e2)/100;
    return `<tr><td>${c}</td><td class="r">${fmt2(a)}</td><td class="r">${pctA(c).toFixed(1)}%</td><td class="r">${fmt2(p)}</td><td class="r">${pctP(c).toFixed(1)}%</td></tr>`;
  }).join('');
  document.getElementById('dTable').innerHTML =
    `<tr><th>品类</th><th>金额（万元）</th><th>金额占比</th><th>毛利（万元）</th><th>毛利占比</th></tr>` + rows +
    `<tr class="tot"><td>合计</td><td class="r">${fmt2(ta)}</td><td class="r">100%</td><td class="r">${fmt2(tp)}</td><td class="r">100%</td></tr>`;
}

// ---------- 销售概览（前一日 · 公司汇总 + 15 站分列） ----------
// 数据源：data.json.salesOverview（由 AirScript v3.7 buildSalesOverview 生成）
// 口径：总额=5品类之和；车流=汽油单+柴油单；笔数=汽用+便+咖；单笔=非烟金额÷笔数；
//       转化率=笔数÷车流。含烟草口径单笔另列参考。
const SO_COLORS = ['#3b82f6','#34d399','#fbbf24','#f97316','#a78bfa'];
// 非正常营业站点（关停 / 停业 / 未复营）：站卡右上角标注「不正常营业」，并从「在营站数」中排除。
// 站点复营后从此数组移除即可，无需改其它代码。
const SO_INACTIVE = ['华山'];
const soFmtY = v => (v==null||isNaN(v)) ? '—' : Number(v).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
const soFmtN = v => (v==null||isNaN(v)) ? '—' : Number(v).toLocaleString('zh-CN');

function renderSalesOverview(){
  const sec = document.getElementById('so');
  if(!sec) return;
  const so = DATA && DATA.salesOverview;
  if(!so || !so.company){
    sec.style.display = 'none';
    return;
  }
  sec.style.display = '';

  const cats = (so.cats && so.cats.length) ? so.cats : (CAT_DETAIL.length ? CAT_DETAIL : ['汽车用品','便利百货','烟草零售','烟草批发','咖啡']);
  const list = so.stations || [];
  const isOff = n => SO_INACTIVE.indexOf(n) >= 0;
  const offNames = list.filter(s => isOff(s.name)).map(s => s.name);
  const liveN = list.filter(s => !isOff(s.name) && (s.amt>0 || s.carTot>0)).length;

  // ── 统计口径随选中站点切换 ──
  //   SEL = 全公司 → 用 salesOverview.company（15 站汇总）
  //   SEL = 某站   → 用 salesOverview.stations 里该站的前一日数据
  const stSel = (SEL === '全公司') ? null : list.filter(s => s.name === SEL)[0];
  const isCompany = !stSel;
  const co = stSel || so.company;
  const total = Number(co.amt) || 0;
  const catAmt = co.catAmt || [];
  const scopeName = stSel ? stSel.name : '';
  const stOff = !!stSel && isOff(stSel.name);

  // ── 顶部条 ──
  // 品类部分直接列出「名称 + 金额 + 占比」，不单独放提示文案
  const catPills = cats.map((n,i)=>{
    const v = Number(catAmt[i])||0;
    const p = total ? (v/total*100).toFixed(1) : '0.0';
    return `<span class="so-pill"><i style="background:${SO_COLORS[i]||'#64748b'}"></i>${n}`
         + `<b>${soFmtY(v)}</b>元<em>${p}%</em></span>`;
  }).join('');
  document.getElementById('soStrip').innerHTML =
    `<span class="tagname">${so.isYesterday ? '前一日' : '最近有数日'} ${String(so.asOf||'').slice(5)}${isCompany ? '' : ' · ' + scopeName}</span>
     <span class="sep"></span>
     <span>${isCompany
        ? `${liveN} 站在营合计 <b>${soFmtY(total)}</b> 元${offNames.length ? `（${offNames.join('、')}不正常营业）` : ''}`
        : `本站合计 <b>${soFmtY(total)}</b> 元${stOff ? '（不正常营业）' : ''}`}</span>
     <span class="sep"></span>
     <span>车流 <b>${soFmtN(co.carTot)}</b> 车次 · 非油 <b>${soFmtN(co.nb)}</b> 笔</span>
     <span class="sep"></span>
     <span class="so-cats">${catPills}</span>`;

  // ── 四个 KPI（单站视图用「本站」前缀，避免与全公司口径混淆） ──
  const P = isCompany ? '' : '本站';
  const kp = [
    {k:P+'总销售金额', v:soFmtY(co.amt), u:'元', s:'含烟草 · 五品类合计'},
    {k:P+'总车流量', v:soFmtN(co.carTot), u:'车次', s:`汽油 ${soFmtN(co.gasCnt)} · 柴油 ${soFmtN(co.dieCnt)}`},
    {k:P+'平均单笔消费', v:soFmtY(co.price), u:'元', s:`非烟口径 · 含烟草 ${soFmtY(co.priceTob)} 元`},
    {k:P+'车次→非油转化率', v:(co.conv==null?'—':Number(co.conv).toFixed(2)), u:'%',
     s:`非油笔数 ${soFmtN(co.nb)} ÷ 车次 ${soFmtN(co.carTot)}`},
  ];
  document.getElementById('soKpis').innerHTML = kp.map(c=>
    `<div class="so-kpi"><div class="k">${c.k}</div><div class="v">${c.v}<small>${c.u}</small></div><div class="s">${c.s}</div></div>`).join('');

  // ── 品类明细（与环形图同色） ──
  document.getElementById('soCatList').innerHTML =
    cats.map((n,i)=>{
      const v = Number(catAmt[i])||0;
      return `<div class="so-crow"><i style="background:${SO_COLORS[i]||'#64748b'}"></i><span>${n}</span>`
           + `<em>${total ? (v/total*100).toFixed(1) : '0.0'}%</em><b>${soFmtY(v)}</b></div>`;
    }).join('')
    + `<div class="so-crow tot"><i style="background:var(--accent)"></i><span>合计</span><em>100.0%</em><b>${soFmtY(total)}</b></div>`;

  document.getElementById('soDonutV').textContent = (total/1e4).toFixed(2);
  const capEl = document.getElementById('soDonutCap');
  if(capEl) capEl.textContent = '万元 · ' + (isCompany ? '总销售' : scopeName + '销售');

  // ── 公司品类环形图 ──
  mk('soDonut', {
    tooltip: {...tipItem, formatter: p => `${p.name}<br/><b>${soFmtY(p.value)}</b> 元 · ${p.percent}%`},
    series: [{
      type:'pie', radius:['60%','90%'], center:['50%','50%'], avoidLabelOverlap:true,
      itemStyle:{borderColor:'#162032', borderWidth:2},
      label:{show:false}, labelLine:{show:false}, emphasis:{scale:true, scaleSize:4},
      data: cats.map((n,i)=>({name:n, value:Number(catAmt[i])||0,
        itemStyle:{color:SO_COLORS[i]||'#64748b'}}))
    }]
  });

  // ── 15 站卡片：仅「全公司」总览视图展示 ──
  //    单站视图下不渲染各站分类明细（含品类金额数字与环形图），避免与单站数据混淆
  const box = document.getElementById('soStations');
  if (SEL !== '全公司') {
    box.innerHTML = '';
    box.style.display = 'none';
    return;
  }
  box.style.display = '';
  box.innerHTML = list.map((s,idx)=>{
    const amt = Number(s.amt)||0;
    const share = total ? (amt/total*100) : 0;
    const rows = cats.map((n,i)=>{
      const v = Number((s.catAmt||[])[i])||0;
      return `<div class="so-li"><i style="background:${SO_COLORS[i]||'#64748b'}"></i><span>${n}</span><b>${soFmtY(v)}</b></div>`;
    }).join('');
    const warn = s.warn ? '<span class="so-bad" title="单加升数越界，疑似「单数/升数」列错位"> ⚠</span>' : '';
    const alive = (amt > 0 || Number(s.carTot) > 0);
    const off = isOff(s.name);
    // 右上角标记优先级：不正常营业 > 占比 / 无数据
    const tag = off ? '<span class="so-off">不正常营业</span>'
              : (alive ? ('占比 ' + share.toFixed(1) + '%') : '<span class="so-na">无数据</span>');
    return `<div class="so-card${SEL===s.name?' on':''}${off?' off':''}" data-so-station="${s.name}">
      <div class="so-c-top"><span class="so-c-name">${s.name}</span>
        <span class="so-c-share">${tag}</span></div>
      <div class="so-c-body">
        <div class="so-c-donut" id="soD${idx}"></div>
        <div class="so-c-list">${rows}
          <div class="so-li tot"><i style="background:var(--accent)"></i><span>合计</span><b>${soFmtY(amt)}</b></div>
        </div>
      </div>
      <div class="so-c-foot">
        <span>车流 <b>${soFmtN(s.carTot)}</b> 车次${warn}</span>
        <span>转化 <b>${(s.conv==null?'—':Number(s.conv).toFixed(2))}%</b></span>
        <span>单笔 <b>${soFmtY(s.price)}</b> 元</span>
      </div>
    </div>`;
  }).join('');

  // 站卡小环形图
  list.forEach((s,idx)=>{
    const el = document.getElementById('soD'+idx);
    if(!el) return;
    const ca = s.catAmt || [];
    const sTot = Number(s.amt)||0;
    mk('soD'+idx, {
      tooltip: {...tipItem, formatter: p => `${p.name}<br/><b>${soFmtY(p.value)}</b> 元 · ${p.percent}%`},
      series: [{
        type:'pie', radius:['62%','94%'], center:['50%','50%'],
        itemStyle:{borderColor:'#162032', borderWidth:1.5},
        label:{show:false}, labelLine:{show:false}, emphasis:{scale:true, scaleSize:3},
        data: sTot > 0
          ? cats.map((n,i)=>({name:n, value:Number(ca[i])||0, itemStyle:{color:SO_COLORS[i]||'#64748b'}}))
          : [{name:'无销售', value:1, itemStyle:{color:'#243049'}}]
      }]
    });
  });

  // 点站卡 → 联动选中（改数据后走统一刷新入口，禁止渲染函数互调）
  box.querySelectorAll('[data-so-station]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const n = el.getAttribute('data-so-station');
      SEL = (SEL === n) ? '全公司' : n;
      syncChips(); renderAll();
    });
  });
}

// ---------- 渲染总入口 ----------
function renderAll(){
  charts.forEach(c=>c.dispose()); charts.length=0;
  renderSalesOverview();
  renderKPI(); renderTrafficKPI(); renderRank(); renderRate(); renderRank2(); renderTrend(); renderDaily30(); renderCompany(); renderComp(); renderDual();
  renderTraffic(); renderFuelPer(); renderConv();
  renderDetail();
  document.getElementById('upd').textContent = '数据生成于 ' + DATA.generatedAt
    + ' · ' + YEAR + ' 会计年（' + (YEAR==='2027' ? '2026-12-26 ~ 2027-12-25' : '2025-12-26 ~ 2026-12-25') + '）'
    + ' · 当前视角：' + SEL;
}

// ---------- 站点选择器 ----------
function buildChips(){
  const box = document.getElementById('stations');
  box.innerHTML = '';   // 年度切换时重建, 防止累积
  const all = document.createElement('div');
  all.className='chip all'+(SEL==='全公司'?' active':'');
  all.textContent='全公司'; all.onclick=()=>{SEL='全公司';syncChips();renderAll();};
  box.appendChild(all);
  DATA.stations.forEach(s=>{
    const c=document.createElement('div');
    c.className='chip'+(SEL===s.name?' active':''); c.textContent=s.name;
    c.onclick=()=>{SEL=s.name;syncChips();renderAll();};
    box.appendChild(c);
  });
}
function syncChips(){ document.querySelectorAll('#stations .chip').forEach(c=>{
  c.classList.toggle('active', (c.classList.contains('all')&&SEL==='全公司') || c.textContent===SEL); }); }

// ---------- 标签页切换 ----------
function switchView(name){
  document.querySelectorAll('#tabs .tab').forEach(t=>t.classList.toggle('active', t.dataset.view===name));
  document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active', v.id==='view-'+name));
  if(name==='forecast'){ renderForecast(); }
  setTimeout(()=>charts.forEach(c=>c.resize()), 60);
}

// ---------- 预测视图 ----------
function renderForecast(){
  const fc = DATA.forecast || {};
  const k = fc.kpis || {};
  const ex = fc.extra || {};
  // KPI —— 主数字 = 全年预计（全口径，含烟草批发去化）；副标题 = 自然经营口径 / 实际完成数（截至 asOf）
  const sc2 = fc.scenarios || {labels:[], A:[], B:[]};
  const task = Number(fc.task) || 1130;
  const natCum = k['自然口径累计万'];
  const cumInc = k['累计含去化销售额万'];
  const hm = k['华山扣除万'];
  const S0 = (k['自然口径全年万']!=null) ? Number(k['自然口径全年万'])
           : ((sc2.A[2]!=null && sc2.A[0]!=null) ? sc2.A[2]-sc2.A[0] : null);
  // 全口径全年 = 自然口径全年 + 全年烟草去化（= 驾驶舱预测模型 B31「全口径A万」）
  const fullYear = (sc2.A[2]!=null) ? Number(sc2.A[2])
                 : ((S0!=null && sc2.A[0]!=null) ? S0+Number(sc2.A[0]) : null);
  const cumGp = (k['累计毛利万']!=null) ? Number(k['累计毛利万'])
              : ((k['整体毛利率']!=null && cumInc!=null) ? Number(k['整体毛利率'])/100*Number(cumInc) : null);
  const cumQu = (k['累计去化额万']!=null) ? Number(k['累计去化额万'])
              : ((cumInc!=null && natCum!=null) ? Number(cumInc)-Number(natCum) : null);
  const remPred = (k['剩余预测万']!=null) ? Number(k['剩余预测万'])
              : ((S0!=null && natCum!=null && hm!=null) ? S0-Number(natCum)+Number(hm) : null);
  const rateFull = (fullYear!=null) ? fullYear/task*100 : null;
  const ratePlan = (S0!=null) ? S0/task*100 : null;
  const rateReal = (natCum!=null) ? Number(natCum)/task*100 : null;
  const N = v => (v==null || !isFinite(v)) ? '—' : Number(v).toFixed(2);
  const cards = [
    {k:'预计全年 · 全口径', v:N(fullYear), u:'万',
     s:`自然口径 ${N(S0)} 万 ＋ 烟草批发去化 ${N(sc2.A[0])} 万`},
    {k:'预计任务达成率 · 全口径', v:(rateFull==null?'—':rateFull.toFixed(2)), u:'%',
     warn:(rateFull!=null && rateFull<100),
     s:`超额 ${rateFull==null?'—':(rateFull-100).toFixed(2)}% · 自然口径 ${ratePlan==null?'—':ratePlan.toFixed(2)}% · 实际进度 ${rateReal==null?'—':rateReal.toFixed(2)}%`},
    {k:'预计全年公司毛利', v:N(sc2.A[3]), u:'万',
     s:`实际累计 ${N(cumGp)} 万 · 预计毛利率 23.50%（守住红线情景）`},
    {k:'预计期末香烟库存', v:N(sc2.A[4]), u:'万',
     s:`当前 ${N(ex['当前香烟库存万'])} 万 · 全年去化 ${N(sc2.A[0])} 万（已去化 ${N(cumQu)} 万）`},
  ];
  document.getElementById('fcKpis').innerHTML = cards.map(c=>
    `<div class="fc-kpi"><div class="k">${c.k}</div><div class="v"${c.warn?' style="color:#ef4444"':''}>${c.v}<small>${c.u}</small></div><div class="s">${c.s}</div></div>`).join('')
    + `<div class="note" style="grid-column:1/-1;margin-top:0">卡片主数字 = <b>全口径全年</b>（含烟草批发去化 ${N(sc2.A[0])} 万，即驾驶舱「全口径A」）；副标题给<b>自然经营口径</b>与<b>实际完成数</b>。自然口径全年 ${N(S0)} 万 = 实际累计 ${N(natCum)} ＋ 剩余预测 ${N(remPred)} − 华山扣除 ${N(hm)} 万；再加全年去化 ${N(sc2.A[0])} 万 = 全口径 ${N(fullYear)} 万。数据截至 ${fc.asOf||'—'}。</div>`;

  // 双情景对比图
  const sc = fc.scenarios || {labels:[], A:[], B:[]};
  mk('fc1',{tooltip:{...baseTip,valueFormatter: v=>num2(v)+' 万'},
    legend:{data:['23.5%红线','24%红线'],textStyle:{color:'#dbe3f0',fontSize:12,fontWeight:600},top:2},
    grid:{left:56,right:24,top:36,bottom:40}, xAxis:cat(sc.labels), yAxis:val('万元'),
    series:[
      {name:'23.5%红线',type:'bar',data:sc.A,barWidth:26,itemStyle:{color:'#34d399',borderRadius:[4,4,0,0]},
        label:{show:true,position:'top',color:C.mut,fontSize:9.5,formatter:p=>Number(p.value).toFixed(1)}},
      {name:'24%红线',type:'bar',data:sc.B,barWidth:26,itemStyle:{color:'#fbbf24',borderRadius:[4,4,0,0]},
        label:{show:true,position:'top',color:C.mut,fontSize:9.5,formatter:p=>Number(p.value).toFixed(1)}}
    ]});

  // 销量超额
  const ov = fc.overrun || {};
  mk('fc2',{tooltip:{...baseTip,formatter:p=>p[0].name+'：'+Number(p[0].value).toFixed(2)+'%'},
    grid:{left:56,right:24,top:24,bottom:34}, xAxis:cat(['23.5%红线','24%红线']),
    yAxis:val('超额率',v=>v+'%'),
    series:[{type:'bar',data:[ov.A, ov.B],barWidth:52,
      itemStyle:{color:p=>p.dataIndex===0?'#34d399':'#fbbf24',borderRadius:[4,4,0,0]},
      label:{show:true,position:'top',color:C.mut,fontSize:11,formatter:p=>Number(p.value).toFixed(2)+'%'}}]});

  // 短信原文
  const sms = fc.sms || '';
  const lines = sms.split('\n').filter(Boolean);
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  document.getElementById('fcSms').innerHTML = lines.map((ln,i)=>`<div>${esc(ln)}</div>`).join('')
    + `<div class="note">数据截至 ${fc.asOf||'—'} · 预测截止 ${fc.predEnd||'—'} · 剩余 ${fc.remainingMonths??'—'} 个月 · 当前香烟库存 ${ex['当前香烟库存万']??'—'}万（库存约束去化上限 ${ex['库存约束去化上限万']??'—'}万${ex['库存是否封顶']===1?'，已封顶':''}）</div>`;
}

// 图1/图2 柱体点击 -> 选中该站 (联动)
function bindChartClick(){
  ['c1','c2'].forEach(id=>{
    const c = echarts.getInstanceByDom(document.getElementById(id));
    if(c) c.on('click', p=>{ if(p.name && DATA.stations.some(s=>s.name===p.name)){ SEL=p.name; syncChips(); renderAll(); }});
  });
}

document.querySelectorAll('#tabs .tab').forEach(t=>{
  t.addEventListener('click', ()=>switchView(t.dataset.view));
});
window.addEventListener('resize',()=>charts.forEach(c=>c.resize()));

// ---------- 双年度数据层 ----------
let YEAR = '2027';
const JSON_BY_YEAR = {'2027':'data.json', '2026':'data_2026.json'};
const YEAR_CACHE = {};

function syncYearTabs(){
  document.querySelectorAll('#yearbar .ytab').forEach(t=>
    t.classList.toggle('active', t.dataset.year===YEAR));
}

const detectYear = d => (d && d.year==='2027') ? '2027' : '2026';

async function loadYear(y){
  if(YEAR_CACHE[y]){ DATA = YEAR_CACHE[y]; }
  else{
    const r = await fetch(JSON_BY_YEAR[y], {cache:'no-store'});
    if(!r.ok) throw new Error('HTTP '+r.status);
    DATA = await r.json();
    // 按数据自带 year 字段归位缓存: 过渡期 data.json 仍是 2026 数据时,
    // 点「2027」Tab 会自动落回 2026 视图, 避免数据年份与标签错位。
    y = detectYear(DATA);
    YEAR_CACHE[y] = DATA;
  }
  CAT_DETAIL = (DATA.composition && DATA.composition.cats) || [];
  SEL = '全公司';
  YEAR = y;
  buildChips(); renderAll(); bindChartClick(); syncYearTabs();
}

async function switchYear(y){
  if(y===YEAR && DATA) return;
  const u = document.getElementById('upd');
  try{
    u.textContent = '⏳ 正在加载 ' + y + ' 会计年数据…';
    await loadYear(y);
  }catch(e){
    u.textContent = '⚠️ ' + y + ' 会计年数据暂不可用 (' + e + ')';
    syncYearTabs();
  }
}

function boot(){
  // 首屏固定读 data.json(="当前会计年"数据, AirScript 推送目标), 再按数据自带年份落视图:
  // 12-26 切换前 data.json=2026数据 → 自动高亮「2026」; 切换后 AirScript-2027 推送
  // 带 year:"2027" 的 data.json → 自动高亮「2027」。归档快照 data_2026.json 仅由 Tab 读取。
  const u = document.getElementById('upd');
  fetch('data.json', {cache:'no-store'})
    .then(r=>{ if(!r.ok) throw new Error('HTTP '+r.status); return r.json(); })
    .then(d=>{
      const y = (d && d.year==='2027') ? '2027' : '2026';
      YEAR_CACHE[y] = d;
      DATA = d; CAT_DETAIL = (d.composition && d.composition.cats) || [];
      SEL = '全公司'; YEAR = y;
      buildChips(); renderAll(); bindChartClick(); syncYearTabs();
    })
    .catch(e=>{ if(u) u.textContent='⚠️ 数据加载失败: '+e; });
}
document.querySelectorAll('#yearbar .ytab').forEach(t=>{
  t.addEventListener('click', ()=>switchYear(t.dataset.year));
});
boot();
</script>
<script src="hub-home.js" defer></script><!-- HUB-HOME v1 --></body>
</html>
"""


def merge_yoy_from_archived_json(data):
    """渲染前把 2026 归档 json 的月度数据合并为 yoy2026 块（页面同比虚线数据源）。
    驾驶舱「2026同比」sheet 若已提供则不覆盖；归档件不存在则静默跳过。
    ——收官夜只需 --snapshot-2026 归档 data.json，同比能力即自动生成，无需写驾驶舱。"""
    if data.get("yoy2026"):
        return data
    p = YEAR_CFG["2026"]["json"]
    if not os.path.exists(p):
        return data
    try:
        with open(p, encoding="utf-8") as f:
            d26 = json.load(f)
        comp = (d26 or {}).get("company") or {}
        mon = (d26 or {}).get("monthly") or {}
        if mon.get("tob"):
            data["yoy2026"] = {
                "monthLabels": comp.get("monthLabels") or [f"{m}月" for m in range(1, 13)],
                "monthlyTob": mon.get("tob", {}),
                "monthlyTobProfit": mon.get("tobProfit", {}),
                "companyTob": comp.get("tob") or [],
                "companyTobProfit": comp.get("tobProfit") or [],
            }
    except Exception:
        pass
    return data


def render(data, out_path=INDEX):
    html = HTML.replace("__DATA__", json.dumps(data, ensure_ascii=False))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def main():
    # 参数: --year 2026|2027 (默认2027) · --from-json · --snapshot-2026 (收官归档)
    year = "2027"
    if "--year" in sys.argv:
        i = sys.argv.index("--year")
        if i + 1 < len(sys.argv):
            year = sys.argv[i + 1] if sys.argv[i + 1] in YEAR_CFG else "2027"
    cfg = YEAR_CFG[year]

    if "--snapshot-2026" in sys.argv:
        # 2026会计年收官(12-25)后执行一次: 当前 data.json → data_2026.json 快照, 此后不再更新
        import shutil
        if not os.path.exists(DATA_JSON):
            print("❌ 未找到 data.json, 无法归档")
            return
        shutil.copyfile(DATA_JSON, YEAR_CFG["2026"]["json"])
        print("✅ 已归档: data.json -> data_2026.json (2026会计年最终快照)")
        return

    out_json = cfg["json"]
    if "--from-json" in sys.argv and os.path.exists(out_json):
        with open(out_json, encoding="utf-8") as f:
            data = json.load(f)
        print(f"▶ 仅从 {os.path.basename(out_json)} 重渲染 index.html ({year})")
    else:
        data = build_data(year)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        print(f"▶ 已从 Excel({year}) 提取数据 -> {os.path.basename(out_json)}")
    # 页面运行时 fetch json（HTML 不内嵌数据），yoy 合并必须回写 json 才能生效。
    # 仅 2027 视图合并；Action 每天10点 --from-json 时与 AirScript 推送自然衔接
    # （AirScript 覆盖后下次 Action 重渲染会把 yoy 补回）。
    if year == "2027":
        data = merge_yoy_from_archived_json(data)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    render(data)
    print("✅ 已生成:", INDEX)
    print("   站点", len(data["stations"]), "站 | 趋势", len(data["monthly"]["tob"]),
          "站 | 构成", len(data["composition"]["cats"]), "类 | 品类明细",
          len(data["categoryDetail"]["cats"]), "类×", len(data["categoryDetail"]["amount"]["汽车用品"]), "站×12月")
    print("   公司累计(含烟草):", data["totals"]["tobAmount"], "万 | 毛利:",
          data["totals"]["tobProfit"], "万 | 任务:",
          data["taskTotal"] if data["taskTotal"] is not None else "待下达")


if __name__ == "__main__":
    main()
