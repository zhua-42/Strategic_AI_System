# -*- coding: utf-8 -*-
"""
龙头公司对比（Leader Comparison）
=================================
按行业选取 3~4 家龙头公司，抓取真实财务指标做横向对比，
口径参考《企业财务报表分析》（资产质量 / 资本结构 / 利润质量 / 现金流 /
杜邦分解）与券商行研「公司分析模版」（市场位势 / 盈利能力 / 供应链话语权 /
成长后劲 / 政策抗性 / 经营效率）。

数据来源（按优先级）：
1. akshare 东方财富「财务摘要」接口（真实、实时，含 ROE/毛利率/净利率/营收/净利润/每股收益）
2. 本地 SQLite company_financial 表
3. 内置兜底基准（龙头名单 + 历史公开区间值，标注「兜底口径」）

对外接口：
    get_leaders(industry) -> [公司名, ...]  （3~4 家）
    compare_leaders(industry) -> {companies, metrics, values, notes, sources}
    build_leader_payload(industry) -> 图表/表格用的结构化 payload
"""
import os
import sqlite3
import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial_research.db")

# 行业 -> 龙头公司（优先真实抓取；抓不到时用兜底基准）
LEADER_MAP = {
    "新能源汽车": ["宁德时代", "比亚迪", "亿纬锂能", "国轩高科"],
    "汽车零部件": ["宁德时代", "比亚迪", "福耀玻璃", "三花智控"],
    "白酒行业": ["贵州茅台", "五粮液", "泸州老窖", "山西汾酒"],
    "家电制造": ["美的集团", "格力电器", "海尔智家", "海信视像"],
    "房地产": ["保利发展", "万科A", "招商蛇口", "华润置地"],
    "银行业": ["招商银行", "工商银行", "建设银行", "宁波银行"],
    "医药生物": ["恒瑞医药", "迈瑞医疗", "药明康德", "爱尔眼科"],
    "化学制药": ["恒瑞医药", "复星医药", "科伦药业", "人福医药"],
    "医疗器械": ["迈瑞医疗", "联影医疗", "乐普医疗", "鱼跃医疗"],
    "半导体": ["中芯国际", "韦尔股份", "北方华创", "中微公司"],
    "光伏": ["隆基绿能", "通威股份", "晶科能源", "天合光能"],
    "机器人": ["汇川技术", "埃斯顿", "绿的谐波", "双环传动"],
    "人工智能": ["海康威视", "科大讯飞", "浪潮信息", "寒武纪"],
    "电力设备": ["宁德时代", "阳光电源", "国电南瑞", "特变电工"],
    "通信设备": ["中兴通讯", "烽火通信", "中际旭创", "新易盛"],
    "消费电子": ["立讯精密", "歌尔股份", "传音控股", "蓝思科技"],
}

# 兜底基准：公司 -> 关键财务指标（公开历史口径，仅当实时抓取失败时使用）
_FALLBACK = {
    "贵州茅台": {"roe": 30.0, "gross_margin": 91.0, "net_margin": 50.0, "revenue_yoy": 15.0, "profit_yoy": 16.0, "eps": 65.0},
    "五粮液": {"roe": 24.0, "gross_margin": 76.0, "net_margin": 36.0, "revenue_yoy": 12.0, "profit_yoy": 13.0, "eps": 8.5},
    "泸州老窖": {"roe": 28.0, "gross_margin": 86.0, "net_margin": 45.0, "revenue_yoy": 20.0, "profit_yoy": 22.0, "eps": 9.5},
    "山西汾酒": {"roe": 33.0, "gross_margin": 75.0, "net_margin": 30.0, "revenue_yoy": 18.0, "profit_yoy": 25.0, "eps": 8.0},
    "宁德时代": {"roe": 22.0, "gross_margin": 24.0, "net_margin": 12.0, "revenue_yoy": 18.0, "profit_yoy": 20.0, "eps": 11.0},
    "比亚迪": {"roe": 18.5, "gross_margin": 20.0, "net_margin": 5.2, "revenue_yoy": 25.0, "profit_yoy": 30.0, "eps": 11.5},
    "亿纬锂能": {"roe": 15.0, "gross_margin": 17.0, "net_margin": 9.0, "revenue_yoy": 20.0, "profit_yoy": 18.0, "eps": 2.2},
    "国轩高科": {"roe": 8.0, "gross_margin": 16.0, "net_margin": 3.0, "revenue_yoy": 15.0, "profit_yoy": 12.0, "eps": 0.5},
    "美的集团": {"roe": 22.0, "gross_margin": 26.0, "net_margin": 9.0, "revenue_yoy": 8.0, "profit_yoy": 12.0, "eps": 4.9},
    "格力电器": {"roe": 25.0, "gross_margin": 30.0, "net_margin": 13.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 6.0},
    "海尔智家": {"roe": 18.0, "gross_margin": 31.0, "net_margin": 7.0, "revenue_yoy": 6.0, "profit_yoy": 10.0, "eps": 1.8},
    "海信视像": {"roe": 14.0, "gross_margin": 18.0, "net_margin": 5.0, "revenue_yoy": 12.0, "profit_yoy": 15.0, "eps": 1.9},
    "保利发展": {"roe": 8.0, "gross_margin": 18.0, "net_margin": 4.0, "revenue_yoy": -10.0, "profit_yoy": -20.0, "eps": 1.4},
    "万科A": {"roe": 5.0, "gross_margin": 15.0, "net_margin": 2.0, "revenue_yoy": -15.0, "profit_yoy": -30.0, "eps": 0.6},
    "招商蛇口": {"roe": 7.0, "gross_margin": 17.0, "net_margin": 3.5, "revenue_yoy": -5.0, "profit_yoy": -10.0, "eps": 0.9},
    "华润置地": {"roe": 10.0, "gross_margin": 22.0, "net_margin": 8.0, "revenue_yoy": 5.0, "profit_yoy": 3.0, "eps": 4.5},
    "招商银行": {"roe": 15.5, "gross_margin": 0.0, "net_margin": 36.0, "revenue_yoy": 3.0, "profit_yoy": 5.0, "eps": 5.6},
    "工商银行": {"roe": 10.5, "gross_margin": 0.0, "net_margin": 40.0, "revenue_yoy": 2.0, "profit_yoy": 3.0, "eps": 1.0},
    "建设银行": {"roe": 11.0, "gross_margin": 0.0, "net_margin": 38.0, "revenue_yoy": 2.0, "profit_yoy": 3.5, "eps": 1.3},
    "宁波银行": {"roe": 14.5, "gross_margin": 0.0, "net_margin": 35.0, "revenue_yoy": 8.0, "profit_yoy": 10.0, "eps": 3.9},
    "恒瑞医药": {"roe": 12.0, "gross_margin": 85.0, "net_margin": 20.0, "revenue_yoy": 15.0, "profit_yoy": 25.0, "eps": 0.9},
    "迈瑞医疗": {"roe": 30.0, "gross_margin": 65.0, "net_margin": 32.0, "revenue_yoy": 12.0, "profit_yoy": 15.0, "eps": 10.0},
    "药明康德": {"roe": 15.0, "gross_margin": 40.0, "net_margin": 20.0, "revenue_yoy": -5.0, "profit_yoy": -10.0, "eps": 2.4},
    "爱尔眼科": {"roe": 20.0, "gross_margin": 50.0, "net_margin": 18.0, "revenue_yoy": 15.0, "profit_yoy": 20.0, "eps": 0.5},
    "复星医药": {"roe": 8.0, "gross_margin": 45.0, "net_margin": 8.0, "revenue_yoy": 3.0, "profit_yoy": 5.0, "eps": 1.2},
    "科伦药业": {"roe": 12.0, "gross_margin": 55.0, "net_margin": 12.0, "revenue_yoy": 10.0, "profit_yoy": 15.0, "eps": 1.5},
    "人福医药": {"roe": 10.0, "gross_margin": 40.0, "net_margin": 8.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 0.8},
    "联影医疗": {"roe": 12.0, "gross_margin": 48.0, "net_margin": 18.0, "revenue_yoy": 20.0, "profit_yoy": 25.0, "eps": 2.0},
    "乐普医疗": {"roe": 10.0, "gross_margin": 60.0, "net_margin": 15.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 0.7},
    "鱼跃医疗": {"roe": 15.0, "gross_margin": 50.0, "net_margin": 22.0, "revenue_yoy": 10.0, "profit_yoy": 12.0, "eps": 1.6},
    "中芯国际": {"roe": 5.0, "gross_margin": 22.0, "net_margin": 8.0, "revenue_yoy": 15.0, "profit_yoy": 10.0, "eps": 0.6},
    "韦尔股份": {"roe": 12.0, "gross_margin": 30.0, "net_margin": 12.0, "revenue_yoy": 25.0, "profit_yoy": 35.0, "eps": 1.8},
    "北方华创": {"roe": 14.0, "gross_margin": 42.0, "net_margin": 18.0, "revenue_yoy": 30.0, "profit_yoy": 35.0, "eps": 3.0},
    "中微公司": {"roe": 10.0, "gross_margin": 45.0, "net_margin": 15.0, "revenue_yoy": 25.0, "profit_yoy": 30.0, "eps": 1.5},
    "隆基绿能": {"roe": 6.0, "gross_margin": 15.0, "net_margin": 3.0, "revenue_yoy": -10.0, "profit_yoy": -30.0, "eps": 0.8},
    "通威股份": {"roe": 8.0, "gross_margin": 20.0, "net_margin": 5.0, "revenue_yoy": -5.0, "profit_yoy": -15.0, "eps": 1.9},
    "晶科能源": {"roe": 10.0, "gross_margin": 14.0, "net_margin": 4.0, "revenue_yoy": 8.0, "profit_yoy": 10.0, "eps": 0.9},
    "天合光能": {"roe": 9.0, "gross_margin": 15.0, "net_margin": 4.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 1.3},
    "汇川技术": {"roe": 20.0, "gross_margin": 34.0, "net_margin": 17.0, "revenue_yoy": 20.0, "profit_yoy": 18.0, "eps": 1.8},
    "埃斯顿": {"roe": 8.0, "gross_margin": 30.0, "net_margin": 5.0, "revenue_yoy": 15.0, "profit_yoy": 20.0, "eps": 0.3},
    "绿的谐波": {"roe": 10.0, "gross_margin": 45.0, "net_margin": 25.0, "revenue_yoy": 10.0, "profit_yoy": 15.0, "eps": 0.9},
    "双环传动": {"roe": 12.0, "gross_margin": 22.0, "net_margin": 9.0, "revenue_yoy": 18.0, "profit_yoy": 22.0, "eps": 1.1},
    "海康威视": {"roe": 18.0, "gross_margin": 44.0, "net_margin": 16.0, "revenue_yoy": 8.0, "profit_yoy": 10.0, "eps": 1.6},
    "科大讯飞": {"roe": 5.0, "gross_margin": 42.0, "net_margin": 4.0, "revenue_yoy": 15.0, "profit_yoy": 20.0, "eps": 0.5},
    "浪潮信息": {"roe": 8.0, "gross_margin": 11.0, "net_margin": 2.5, "revenue_yoy": 20.0, "profit_yoy": 25.0, "eps": 1.2},
    "寒武纪": {"roe": -5.0, "gross_margin": 55.0, "net_margin": -10.0, "revenue_yoy": 100.0, "profit_yoy": 80.0, "eps": 0.2},
    "阳光电源": {"roe": 20.0, "gross_margin": 30.0, "net_margin": 12.0, "revenue_yoy": 25.0, "profit_yoy": 30.0, "eps": 5.0},
    "国电南瑞": {"roe": 15.0, "gross_margin": 30.0, "net_margin": 15.0, "revenue_yoy": 10.0, "profit_yoy": 12.0, "eps": 1.1},
    "特变电工": {"roe": 12.0, "gross_margin": 25.0, "net_margin": 10.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 1.9},
    "中兴通讯": {"roe": 12.0, "gross_margin": 40.0, "net_margin": 7.0, "revenue_yoy": 5.0, "profit_yoy": 10.0, "eps": 1.9},
    "烽火通信": {"roe": 6.0, "gross_margin": 22.0, "net_margin": 3.0, "revenue_yoy": 5.0, "profit_yoy": 8.0, "eps": 0.6},
    "中际旭创": {"roe": 18.0, "gross_margin": 30.0, "net_margin": 15.0, "revenue_yoy": 50.0, "profit_yoy": 60.0, "eps": 3.0},
    "新易盛": {"roe": 15.0, "gross_margin": 28.0, "net_margin": 15.0, "revenue_yoy": 40.0, "profit_yoy": 50.0, "eps": 1.5},
    "立讯精密": {"roe": 18.0, "gross_margin": 12.0, "net_margin": 4.5, "revenue_yoy": 12.0, "profit_yoy": 15.0, "eps": 1.8},
    "歌尔股份": {"roe": 8.0, "gross_margin": 12.0, "net_margin": 3.0, "revenue_yoy": 10.0, "profit_yoy": 15.0, "eps": 0.8},
    "传音控股": {"roe": 25.0, "gross_margin": 24.0, "net_margin": 9.0, "revenue_yoy": 15.0, "profit_yoy": 18.0, "eps": 8.0},
    "蓝思科技": {"roe": 10.0, "gross_margin": 18.0, "net_margin": 5.0, "revenue_yoy": 10.0, "profit_yoy": 15.0, "eps": 1.1},
    "福耀玻璃": {"roe": 18.0, "gross_margin": 35.0, "net_margin": 18.0, "revenue_yoy": 15.0, "profit_yoy": 20.0, "eps": 2.8},
    "三花智控": {"roe": 15.0, "gross_margin": 26.0, "net_margin": 12.0, "revenue_yoy": 15.0, "profit_yoy": 18.0, "eps": 0.9},
}

# 用于行业名模糊匹配（与 industry_chain_data 保持一致）
_KEYWORDS = {
    "新能源": "新能源汽车", "锂": "新能源汽车", "汽车": "新能源汽车", "电池": "新能源汽车",
    "白酒": "白酒行业", "酿酒": "白酒行业", "酒": "白酒行业",
    "家电": "家电制造", "电器": "家电制造", "空调": "家电制造",
    "房": "房地产", "地产": "房地产",
    "银": "银行业", "金融": "银行业",
    "医药": "医药生物", "药": "医药生物", "医疗": "医药生物",
    "半导体": "半导体", "芯片": "半导体", "集成电路": "半导体",
    "光伏": "光伏", "太阳能": "光伏",
    "机器": "机器人", "自动化": "机器人",
    "人工智能": "人工智能", "AI": "人工智能", "计算机": "人工智能",
    "电力": "电力设备", "电网": "电力设备",
    "通信": "通信设备", "光模块": "通信设备",
    "消费电子": "消费电子", "电子": "消费电子",
}


def _norm_industry(industry):
    if not industry:
        return ""
    ind = str(industry).strip()
    if ind in LEADER_MAP:
        return ind
    for key, chain in LEADER_MAP.items():
        if key in ind or ind in key:
            return key
    for kw, target in _KEYWORDS.items():
        if kw in ind:
            return target
    return ind


def get_leaders(industry):
    """返回 3~4 家龙头公司名列表。"""
    ind = _norm_industry(industry)
    leaders = LEADER_MAP.get(ind, [])
    if not leaders:
        # 尝试从产业链知识库的龙头字段提取
        try:
            import industry_chain_data as icd
            chain, matched = icd.get_chain(industry)
            names = []
            for seg in chain:
                for part in str(seg.get("leaders", "")).split("、"):
                    part = part.strip()
                    if part and "（" not in part and len(part) <= 6:
                        names.append(part)
            seen = []
            for n in names:
                if n not in seen:
                    seen.append(n)
            leaders = seen[:4]
        except Exception:
            pass
    return leaders[:4]


def _fetch_live_metrics(company):
    """通过 akshare 财务摘要抓取真实指标。失败返回 None。"""
    try:
        from news_fetcher import lookup_stock_code
        import akshare as ak
        code = lookup_stock_code(company)
        if not code:
            return None
        df = ak.stock_financial_abstract(symbol=code)
        if df is None or len(df) == 0:
            return None
        # 找最新报告期列
        cols = [c for c in df.columns if str(c).isdigit() and len(str(c)) == 8]
        if not cols:
            return None
        latest = sorted(cols)[-1]
        def pick(name):
            row = df[df["指标"].astype(str).str.contains(name, na=False)]
            if row.empty:
                return None
            val = row.iloc[0].get(latest)
            try:
                return float(val)
            except Exception:
                return None
        return {
            "roe": pick("净资产收益率"),
            "gross_margin": pick("销售毛利率"),
            "net_margin": pick("销售净利率"),
            "revenue": pick("营业总收入"),
            "net_profit": pick("归母净利润"),
            "eps": pick("基本每股收益"),
            "report_date": latest,
        }
    except Exception:
        return None


def _db_metrics(company):
    """从本地 SQLite company_financial 取指标。"""
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("SELECT roe, gross_margin, margin, eps, data_as_of FROM company_financial "
                    "WHERE company_name LIKE ?", (f"%{company}%",))
        row = cur.fetchone()
        conn.close()
        if row:
            return {"roe": row[0] or 0, "gross_margin": row[1] or 0,
                    "net_margin": row[2] or 0, "eps": row[3] or 0,
                    "report_date": row[4] or "本地库"}
    except Exception:
        pass
    return None


def compare_leaders(industry):
    """
    龙头对比主函数。返回：
    {
      "industry": 归一行业,
      "companies": [名字...],
      "metrics": ["ROE(%)", "毛利率(%)", "净利率(%)", "营收同比(%)", "净利同比(%)", "EPS(元)"],
      "values": {公司: {指标: 值}},
      "notes": {公司: 说明},
      "sources": [来源说明],
      "ok": bool
    }
    """
    ind = _norm_industry(industry)
    leaders = get_leaders(ind)
    if not leaders:
        return {"industry": ind, "companies": [], "metrics": [],
                "values": {}, "notes": {}, "sources": [], "ok": False}

    metrics = ["ROE(%)", "毛利率(%)", "净利率(%)", "营收同比(%)", "净利同比(%)", "EPS(元)"]
    values = {}
    notes = {}
    live_used = 0

    for comp in leaders:
        live = _fetch_live_metrics(comp)
        if live and (live.get("roe") is not None or live.get("gross_margin") is not None):
            live_used += 1
            note = f"实时抓取（财务摘要 {live.get('report_date', '')}）"
        else:
            live = _db_metrics(comp) or {}
            note = "本地数据库/兜底基准口径"
        fb = _FALLBACK.get(comp, {})
        roe = live.get("roe") if live.get("roe") is not None else fb.get("roe", 0)
        gm = live.get("gross_margin") if live.get("gross_margin") is not None else fb.get("gross_margin", 0)
        nm = live.get("net_margin") if live.get("net_margin") is not None else fb.get("net_margin", 0)
        rev = live.get("revenue") if live.get("revenue") is not None else fb.get("revenue", 0)
        profit = live.get("net_profit") if live.get("net_profit") is not None else fb.get("profit_yoy", 0)
        eps = live.get("eps") if live.get("eps") is not None else fb.get("eps", 0)
        # 营收/净利同比：实时接口给出的是绝对额，取兜底同比（或估算）
        rev_yoy = fb.get("revenue_yoy", 0)
        profit_yoy = fb.get("profit_yoy", 0)
        values[comp] = {
            "ROE(%)": round(roe, 2),
            "毛利率(%)": round(gm, 2),
            "净利率(%)": round(nm, 2),
            "营收同比(%)": round(rev_yoy, 1),
            "净利同比(%)": round(profit_yoy, 1),
            "EPS(元)": round(eps, 2),
        }
        notes[comp] = note

    sources = [
        "龙头名单：行业产业链知识库/公开龙头口径",
        "实时财务指标：东方财富财务摘要接口（若可用）",
        "同比增速与缺失项：上市公司公开报告兜底口径（区间值，仅供对比参考）",
        "对比框架参考：《企业财务报表分析》杜邦/盈利质量/现金流 + 券商行研公司分析模版",
    ]
    if live_used > 0:
        sources.insert(0, f"{live_used}/{len(leaders)} 家公司已抓取到实时财务摘要数据")

    return {"industry": ind, "companies": leaders, "metrics": metrics,
            "values": values, "notes": notes, "sources": sources, "ok": True}


def build_leader_payload(industry):
    """构建用于 Plotly/导出图表的 payload。"""
    r = compare_leaders(industry)
    if not r["ok"]:
        return r
    # 转成 series 结构：指标 -> [各公司值]
    series = {}
    for m in r["metrics"]:
        series[m] = [r["values"][c].get(m, 0) for c in r["companies"]]
    r["series"] = series
    return r


def format_leader_markdown(industry):
    """把龙头对比格式化为报告正文可引用的 markdown 表格文本。"""
    r = compare_leaders(industry)
    if not r["ok"]:
        return "（该行业暂无龙头对比数据）"
    lines = [f"#### 龙头公司横向对比（{r['industry']}）", ""]
    lines.append("| 指标 | " + " | ".join(r["companies"]) + " |")
    lines.append("| --- | " + " | ".join(["---"] * len(r["companies"])) + " |")
    for m in r["metrics"]:
        vals = [str(r["values"][c].get(m, "—")) for c in r["companies"]]
        lines.append(f"| {m} | " + " | ".join(vals) + " |")
    lines.append("")
    for c in r["companies"]:
        lines.append(f"- **{c}**：{r['notes'].get(c, '')}")
    lines.append("")
    lines.append("数据来源：" + "；".join(r["sources"]))
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    ind = sys.argv[1] if len(sys.argv) > 1 else "新能源汽车"
    r = build_leader_payload(ind)
    print("ok:", r["ok"], "| industry:", r.get("industry"), "| companies:", r.get("companies"))
    for c, v in r.get("values", {}).items():
        print(" ", c, v, "|", r.get("notes", {}).get(c))
