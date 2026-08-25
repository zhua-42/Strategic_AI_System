# -*- coding: utf-8 -*-
"""
数据更新器（Deloitte/Forage 数据底座增强版）
=============================================
功能：
1. 通过 akshare「东方财富-业绩报表」拉取全市场最新报告期真实财务数据
2. 行业聚合：CR4、平均 ROE、净利润率、毛利率、经营现金流、样本数（全市场真实计算）
3. 个股精细指标：通过「财务摘要」接口为头部/核心公司补全杜邦与资产结构（总资产、净资产、负债）
4. 把结果写入 financial_research.db（industry_benchmark / company_financial / data_meta）
5. 记录数据截至日期与来源，供前端展示「数据新鲜度」

用法：
    python data_updater.py            # 增量/按需刷新（6 小时内有缓存则跳过）
    python data_updater.py --force    # 强制刷新
"""
import os
import sys
import time
import json
import sqlite3
import argparse
import datetime

import pandas as pd

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial_research.db")
META_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "market_meta.json")
FRESH_SECONDS = 6 * 3600  # 6 小时内视为新鲜

# 报告期候选（优先最新）
REPORT_PERIODS = [
    "20260630", "20260331", "20251231", "20250930",
    "20250630", "20250331", "20241231", "20240930",
    "20240630", "20240331", "20231231",
]

# 东方财富行业名 → 系统常用行业名 归一映射（未列出的行业保留原名）
IND_NAME_MAP = {
    "汽车整车": "新能源汽车", "汽车零部件": "新能源汽车", "电池": "新能源汽车",
    "乘用车": "新能源汽车", "商用车": "新能源汽车", "汽车服务": "新能源汽车",
    "白酒": "白酒行业", "啤酒": "白酒行业", "饮料制造": "白酒行业",
    "白色家电": "家电制造", "黑色家电": "家电制造", "小家电": "家电制造", "厨卫电器": "家电制造",
    "房地产开发": "房地产", "房地产服务": "房地产", "房产开发": "房地产",
    "股份制银行": "银行业", "国有大型银行": "银行业", "城商行": "银行业", "农商行": "银行业",
    "银行": "银行业",
    "半导体": "半导体", "电子化学品": "半导体", "元件": "半导体",
    "光伏设备": "光伏", "风电设备": "光伏", "电池设备": "光伏",
    "化学制药": "医药生物", "中药": "医药生物", "生物制品": "医药生物",
    "医疗器械": "医药生物", "医疗服务": "医药生物", "医药商业": "医药生物",
    "机器人": "机器人", "自动化设备": "机器人",
    "软件开发": "人工智能", "计算机设备": "人工智能", "IT服务": "人工智能",
    "电网设备": "电力设备", "电力": "电力设备",
    "证券": "证券", "保险": "保险", "多元金融": "金融",
    "煤炭开采": "煤炭", "焦炭加工": "煤炭",
    "炼化及贸易": "石油石化", "油气开采": "石油石化",
    "普钢": "钢铁", "特钢": "钢铁",
    "水泥": "建筑材料", "玻璃玻纤": "建筑材料", "装修建材": "建筑材料",
    "物流": "物流", "铁路公路": "交通运输", "航空机场": "交通运输", "港口": "交通运输",
    "食品加工": "食品饮料", "休闲食品": "食品饮料", "调味发酵品": "食品饮料",
    "种植业": "农业", "养殖业": "农业", "农产品加工": "农业", "饲料": "农业",
    "军工电子": "国防军工", "航天装备": "国防军工", "航空装备": "国防军工", "地面兵装": "国防军工",
    "游戏": "传媒", "广告营销": "传媒", "影视院线": "传媒", "数字媒体": "传媒",
    "纺织制造": "纺织服装", "服装家纺": "纺织服装",
    "造纸": "轻工制造", "包装印刷": "轻工制造", "家居用品": "轻工制造",
    "电力设备": "电力设备",
}


def _norm_industry(name):
    """行业名归一：映射到系统常用名；未映射则保留原东财名。"""
    name = str(name or "").strip()
    return IND_NAME_MAP.get(name, name)


def _log(msg):
    print(f"[data_updater] {msg}", flush=True)


def _conn():
    return sqlite3.connect(DB_PATH)


def init_schema():
    """确保行业基准/公司表具备新增真实数据字段。"""
    conn = _conn()
    cur = conn.cursor()
    for col, ddl in [
        ("gross_margin", "REAL DEFAULT 0"),
        ("sample_size", "INTEGER DEFAULT 0"),
        ("data_as_of", "TEXT DEFAULT ''"),
    ]:
        try:
            cur.execute(f"ALTER TABLE industry_benchmark ADD COLUMN {col} {ddl}")
        except Exception:
            pass
    for col, ddl in [
        ("gross_margin", "REAL DEFAULT 0"),
        ("total_assets", "REAL DEFAULT 0"),
        ("total_liability", "REAL DEFAULT 0"),
        ("eps", "REAL DEFAULT 0"),
        ("data_as_of", "TEXT DEFAULT ''"),
        ("data_source", "TEXT DEFAULT ''"),
    ]:
        try:
            cur.execute(f"ALTER TABLE company_financial ADD COLUMN {col} {ddl}")
        except Exception:
            pass
    cur.execute(
        """CREATE TABLE IF NOT EXISTS data_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )"""
    )
    conn.commit()
    conn.close()


def read_meta():
    try:
        if os.path.exists(META_CACHE_FILE):
            with open(META_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def write_meta(meta):
    os.makedirs(os.path.dirname(META_CACHE_FILE), exist_ok=True)
    with open(META_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def _ak_import():
    import akshare as ak
    return ak


def fetch_performance_report():
    """拉取全市场业绩报表：优先最新报告期。返回 (df, period)。"""
    ak = _ak_import()
    for period in REPORT_PERIODS:
        try:
            df = ak.stock_yjbb_em(date=period)
            if df is not None and len(df) > 0:
                _log(f"业绩报表拉取成功: 报告期={period} 行数={len(df)}")
                return df, period
        except Exception as e:
            _log(f"报告期 {period} 失败: {e}")
    return None, None


def to_num(v):
    try:
        if v is None or pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def fetch_kpi_abstract(symbol):
    """个股财务摘要（含总资产/净资产/负债等），失败返回 None。"""
    ak = _ak_import()
    try:
        df = ak.stock_financial_abstract(symbol=symbol)
        if df is None or len(df) == 0:
            return None
        data = df.set_index("指标") if "指标" in df.columns else df
        latest_col = None
        for col in reversed(list(data.columns)):
            if col not in ("选项", "指标") and str(col).isdigit() and len(str(col)) == 8:
                latest_col = col
                break
        if latest_col is None:
            return None
        series = data[latest_col] if latest_col in data.columns else None
        if series is None:
            return None
        return {
            "report_date": latest_col,
            "total_assets": to_num(_pick(series, "总资产")),
            "total_liability": to_num(_pick(series, "总负债")),
            "net_assets": to_num(_pick(series, "净资产")),
            "roe": to_num(_pick(series, "净资产收益率")),
            "gross_margin": to_num(_pick(series, "销售毛利率")),
            "net_margin": to_num(_pick(series, "销售净利率")),
            "revenue": to_num(_pick(series, "营业总收入")),
            "net_profit": to_num(_pick(series, "归母净利润")),
            "eps": to_num(_pick(series, "基本每股收益")),
        }
    except Exception as e:
        _log(f"个股 {symbol} 财务摘要失败: {e}")
        return None


def _pick(series, name):
    for idx, val in series.items():
        if str(idx).strip().replace(" ", "") == name.replace(" ", ""):
            return val
    for idx, val in series.items():
        if name in str(idx):
            return val
    return 0.0


def refresh_market_data(force=False, verbose=True):
    """主入口：刷新全市场行业/个股基准数据。返回 meta。"""
    meta = read_meta()
    if not force and meta.get("last_refresh"):
        last = meta["last_refresh"]
        try:
            last_ts = datetime.datetime.fromisoformat(last).timestamp()
            if time.time() - last_ts < FRESH_SECONDS:
                _log(f"数据仍新鲜（{last}），跳过。--force 可强制刷新。")
                return meta
        except Exception:
            pass

    init_schema()

    df, period = fetch_performance_report()
    if df is None:
        _log("警告：全市场业绩报表拉取失败，保留旧数据。")
        meta["last_error"] = "stock_yjbb_em 拉取失败"
        meta["last_try"] = datetime.datetime.now().isoformat()
        write_meta(meta)
        return meta

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    colmap = {}
    for c in df.columns:
        if "股票代码" in c:
            colmap[c] = "code"
        elif "股票简称" in c:
            colmap[c] = "name"
        elif "营业总收入-营业总收入" in c:
            colmap[c] = "revenue"
        elif "营业总收入-同比增长" in c:
            colmap[c] = "revenue_yoy"
        elif "净利润-净利润" in c:
            colmap[c] = "net_profit"
        elif "净利润-同比增长" in c:
            colmap[c] = "profit_yoy"
        elif "净资产收益率" in c:
            colmap[c] = "roe"
        elif "销售毛利率" in c:
            colmap[c] = "gross_margin"
        elif "每股经营现金流量" in c:
            colmap[c] = "cashflow_ps"
        elif "每股净资产" in c:
            colmap[c] = "bvps"
        elif "所处行业" in c:
            colmap[c] = "industry"
        elif "最新公告日期" in c:
            colmap[c] = "ann_date"
        elif "每股收益" in c:
            colmap[c] = "eps"
    df = df.rename(columns=colmap)
    df = df.dropna(subset=[c for c in ["name", "revenue", "industry"] if c in df.columns])

    conn = _conn()
    cur = conn.cursor()

    # 2) 行业聚合（归一行业名后合并）
    industry_agg = {}
    if "industry" in df.columns and "revenue" in df.columns:
        df["_ind_norm"] = df["industry"].map(_norm_industry)
        grouped = df.dropna(subset=["_ind_norm"]).groupby("_ind_norm")
        for ind, g in grouped:
            g = g.copy()
            g["revenue"] = pd.to_numeric(g["revenue"], errors="coerce").fillna(0)
            g["net_profit"] = pd.to_numeric(g.get("net_profit", 0), errors="coerce").fillna(0)
            g["roe"] = pd.to_numeric(g.get("roe", 0), errors="coerce").fillna(0)
            g["gross_margin"] = pd.to_numeric(g.get("gross_margin", 0), errors="coerce").fillna(0)
            g["cashflow_ps"] = pd.to_numeric(g.get("cashflow_ps", 0), errors="coerce").fillna(0)
            tot_rev = float(g["revenue"].sum())
            tot_profit = float(g["net_profit"].sum())
            n = int(len(g))
            if tot_rev <= 0 or n < 3:
                continue
            top4 = g.nlargest(4, "revenue")
            cr4 = float(top4["revenue"].sum()) / tot_rev * 100.0 if tot_rev > 0 else 0.0
            roe_valid = g[g["roe"] > 0]
            avg_roe = float(roe_valid["roe"].mean()) if len(roe_valid) > 0 else 0.0
            avg_gm = float(g[g["gross_margin"] > 0]["gross_margin"].mean()) if (g["gross_margin"] > 0).any() else 0.0
            avg_npm = (tot_profit / tot_rev * 100.0) if tot_rev > 0 else 0.0
            top10 = g.nlargest(10, "revenue")
            avg_cf = float(top10["cashflow_ps"].mean()) if len(top10) > 0 and (top10["cashflow_ps"] != 0).any() else 0.0
            industry_agg[ind] = {
                "cr4": round(cr4, 2),
                "avg_roe": round(avg_roe, 2),
                "net_profit_margin": round(avg_npm, 2),
                "gross_margin": round(avg_gm, 2),
                "asset_turnover": 0.0,
                "equity_multiplier": 1.0,
                "operating_cash_flow": round(avg_cf, 2),
                "sample_size": n,
            }

    # 3) 行业样本公司（营收前 5）精细指标：补周转率/权益乘数
    top_industries = sorted(industry_agg.items(), key=lambda kv: -kv[1]["sample_size"])[:40]
    detailed_industries = {}
    for ind, _ in top_industries:
        sub = df[df["industry"] == ind].nlargest(5, "revenue")
        assets = 0.0
        liab = 0.0
        revs = 0.0
        n = 0
        for _, row in sub.iterrows():
            code = str(row.get("code", "")).zfill(6)
            kpi = fetch_kpi_abstract(code)
            if kpi is None:
                continue
            assets += kpi["total_assets"]
            liab += kpi["total_liability"]
            revs += kpi["revenue"]
            n += 1
            time.sleep(0.05)
        if n >= 1 and assets > 0:
            ind_turnover = (revs / assets) if assets > 0 else 0.0
            eq_mult = assets / max(1.0, assets - liab) if (assets - liab) > 0 else 1.0
            detailed_industries[ind] = {
                "asset_turnover": round(min(ind_turnover, 5.0), 2),
                "equity_multiplier": round(min(eq_mult, 15.0), 2),
            }
    for ind, det in detailed_industries.items():
        if ind in industry_agg:
            industry_agg[ind].update(det)

    # 4) 写行业基准
    written = 0
    for ind, agg in industry_agg.items():
        src = f"东方财富业绩报表(报告期{period})聚合, {agg['sample_size']}家样本"
        cur.execute(
            """INSERT INTO industry_benchmark
               (industry_name, cr4, avg_roe, net_profit_margin, asset_turnover,
                equity_multiplier, operating_cash_flow, data_source,
                gross_margin, sample_size, data_as_of)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(industry_name) DO UPDATE SET
                 cr4=excluded.cr4, avg_roe=excluded.avg_roe,
                 net_profit_margin=excluded.net_profit_margin,
                 asset_turnover=excluded.asset_turnover,
                 equity_multiplier=excluded.equity_multiplier,
                 operating_cash_flow=excluded.operating_cash_flow,
                 data_source=excluded.data_source, gross_margin=excluded.gross_margin,
                 sample_size=excluded.sample_size, data_as_of=excluded.data_as_of""",
            (ind, agg["cr4"], agg["avg_roe"], agg["net_profit_margin"],
             agg.get("asset_turnover", 0), agg.get("equity_multiplier", 1.0),
             agg["operating_cash_flow"], src, agg["gross_margin"], agg["sample_size"], period),
        )
        written += 1
    conn.commit()
    _log(f"行业基准更新: {written} 个行业")

    # 5) 个股表：标记来源（已有公司）
    cur.execute("SELECT company_name FROM company_financial")
    known = [r[0] for r in cur.fetchall()]
    updated = 0
    for cname in known[:60]:
        sub = df[df["name"].astype(str).str.contains(cname[:2], na=False)]
        if sub.empty:
            continue
        row = sub.iloc[0]
        ind = str(row.get("industry", ""))
        cur.execute(
            """UPDATE company_financial SET
                 industry=COALESCE(?, industry), gross_margin=?, eps=?,
                 data_as_of=?, data_source=?
               WHERE company_name=?""",
            (ind, to_num(row.get("gross_margin")), to_num(row.get("eps")), period,
             f"东方财富业绩报表(报告期{period})", cname),
        )
        updated += 1
    conn.commit()
    _log(f"个股财务标注更新: {updated} 家（含毛利率/每股现金流/报告期）")

    conn.close()

    meta = {
        "last_refresh": datetime.datetime.now().isoformat(),
        "report_period": period,
        "industries": written,
        "companies_total": int(len(df)) if df is not None else 0,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    write_meta(meta)
    _log(f"完成: 报告期={period} 行业={written} 公司={meta['companies_total']}")
    return meta


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true", help="强制刷新")
    args = parser.parse_args()
    refresh_market_data(force=args.force)
