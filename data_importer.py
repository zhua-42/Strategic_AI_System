# -*- coding: utf-8 -*-
"""
456 目录新数据清洗入库（Data Importer 2026）
============================================
把 C:\\Users\\DELL\\Desktop\\456 下的行业财务/行情 Excel 清洗后写入
financial_research.db，包括：

1. financial_statements  财务报表明细表（资产负债表/利润表/现金流量表）
   列：stkcd, short_name, accper, typrep, stmt_type, item, value
2. company_financial     聚合核心指标（ROE/毛利率/净利率/EPS/总资产/负债）
   —— 按最新报告期从三大表计算，供 get_company_data 使用
3. stock_daily_quotes    个股日行情（品牌类文件：中国传媒/保险行业主要品牌）
   列：code, name, trade_date, open, high, low, close, volume, amount
4. sector_quotes         板块行情（同花顺行情数据）
   列：sector, pct_chg, turnover, net_inflow, up_count, down_count, leader

用法：
    python data_importer.py                 # 导入全部
    python data_importer.py --dry-run       # 只统计不写入
"""
import os
import sys
import sqlite3
import datetime

import pandas as pd

SRC_DIR = r"C:\Users\DELL\Desktop\456"
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "financial_research.db")

# 三类文件识别
STMT_FILES = [  # 财务报表（三张表）
    "体育业.xlsx", "农、林、牧、渔业.xlsx", "半导体行业.xlsx",
    "广播、电视、电影和影视录音制作业.xlsx", "批发业.xlsx", "文化艺术业.xlsx",
    "新能源汽车整车生产.xlsx", "新闻与传媒行业现金流量表.xlsx",
    "新闻与传媒行业资产负债表.xlsx", "新闻和出版业.xlsx", "白酒行业.xlsx", "综合.xlsx",
]
QUOTE_FILES = [  # 个股日行情（每个 sheet 一只股票）
    "中国传媒行业品牌.xlsx", "保险行业主要品牌.xlsx",
]
SECTOR_FILE = "同花顺行情数据.xlsx"

# 利润表 -> company_financial 的关键科目（中文列名，自动兼容编码型文件）
KEY_ITEMS = {
    "营业总收入": "revenue",
    "营业收入": "revenue",
    "净利润": "net_profit",
    "归母净利润": "net_profit",
    "营业成本": "oper_cost",
    "销售费用": "sell_exp",
    "管理费用": "admin_exp",
    "财务费用": "fin_exp",
    "研发费用": "rd_exp",
    "基本每股收益": "eps",
    "净资产收益率": "roe",
    "总资产报酬率": "roa",
    "销售毛利率": "gross_margin",
    "销售净利率": "net_margin",
}
BALANCE_ITEMS = {
    "资产总计": "total_assets",
    "负债合计": "total_liability",
    "所有者权益合计": "total_equity",
    "股东权益合计": "total_equity",
    "货币资金": "cash_balance",
    "存货": "inventory",
    "应收账款净额": "receivable",
    "应收账款": "receivable",
}
CASH_ITEMS = {
    "经营活动产生的现金流量净额": "oper_cashflow",
    "销售商品、提供劳务收到的现金": "sales_cash",
}


def _load_sheet_raw(xl, sheet, max_cols=200):
    """读取 sheet 原始行：动态检测表头行（含'没有单位'/'元'的行=单位行），
    自动确定数据起始行；科目列名取表头里的中文名。"""
    raw = xl.parse(sheet, header=None, nrows=6)
    # 找单位行：包含 '没有单位' 或大量 '元' 的行
    unit_row = None
    for i in range(min(4, len(raw))):
        row_vals = [str(v) for v in raw.iloc[i].values]
        if any("没有单位" in v for v in row_vals):
            unit_row = i
            break
        yuan_cnt = sum(1 for v in row_vals if v == "元")
        if yuan_cnt >= 3:
            unit_row = i
            break
    if unit_row is None:
        # 找不到单位行：假设 row0 是中文列名、row1 起为数据
        hdr = 1
        names_row = 0
    else:
        hdr = unit_row + 1            # 数据从单位行下一行开始
        names_row = unit_row - 1 if unit_row >= 1 else 0  # 中文列名在单位行上一行
    # 中文列名行
    names = [str(v) for v in raw.iloc[names_row].values] if names_row < len(raw) else []
    df = xl.parse(sheet, header=None, skiprows=hdr)
    df.columns = [str(c).strip() for c in df.columns]
    # 标准化前 4 列（无论哪种表头，前 4 列语义一致）
    base_names = ["证券代码", "证券简称", "统计截止日期", "报表类型"]
    for i, nm in enumerate(base_names):
        if i < len(df.columns):
            df.columns.values[i] = nm
    # 科目列（第 4 列之后）：优先用中文名（names 行），列名若是纯数字/编码则替换
    for i in range(len(base_names), min(len(df.columns), len(names))):
        nm = str(names[i]).strip()
        cur = str(df.columns[i]).strip()
        if nm and nm != "nan" and "没有单位" not in nm and ("证券代码" not in nm):
            # 当前列名是编码/数字时才替换；已是中文名则保留
            if not any("\u4e00" <= ch <= "\u9fff" for ch in cur):
                df.columns.values[i] = nm
    return df


def _norm_col_map(df):
    """构建 中文科目名 -> 实际列名 映射。"""
    colmap = {}
    for col in df.columns:
        s = str(col).strip()
        colmap[s] = col
        # 去空格匹配（如 '营业总收入' 与 '营业总收入 '）
        colmap[s.replace(" ", "")] = col
    return colmap


def _to_num(v):
    try:
        if v is None or pd.isna(v):
            return None
        return float(v)
    except Exception:
        return None


def import_statement_file(fn, conn, dry=False):
    """导入一个财务报表文件（三张表 -> financial_statements + 聚合公司指标）。"""
    fp = os.path.join(SRC_DIR, fn)
    xl = pd.ExcelFile(fp)
    industry = os.path.splitext(fn)[0]
    cur = conn.cursor()
    stats = {"rows": 0, "companies": set(), "latest": {}}

    for sheet in xl.sheet_names:
        sheet_l = sheet.lower()
        if "资产负债" in sheet_l:
            stmt = "balance"
        elif "利润" in sheet_l:
            stmt = "income"
        elif "现金流" in sheet_l:
            stmt = "cashflow"
        elif sheet_l.startswith("sheet"):
            # 单 sheet 文件：用文件名判断报表类型
            fn_l = fn.lower()
            if "资产负债" in fn_l:
                stmt = "balance"
            elif "利润" in fn_l:
                stmt = "income"
            elif "现金流" in fn_l:
                stmt = "cashflow"
            else:
                continue
        else:
            continue
        try:
            df = _load_sheet_raw(xl, sheet)
        except Exception as e:
            print(f"   [skip] {fn}/{sheet}: {e}")
            continue
        if df is None or len(df) == 0:
            continue
        colmap = _norm_col_map(df)
        # 跳过单位行
        df = df[~df["证券简称"].astype(str).str.contains("没有单位|单位", na=False)]
        df = df.dropna(subset=["证券代码", "证券简称"])

        for _, r in df.iterrows():
            code = str(r.get("证券代码", "")).strip()
            name = str(r.get("证券简称", "")).strip()
            accper = str(r.get("统计截止日期", "")).strip()
            typrep = str(r.get("报表类型", "")).strip()
            if not code or not name or not accper:
                continue
            # 容错：nan / 日期错位 / 空值一律视为合并报表 A
            if typrep in ("nan", "None", "", "NaT"):
                typrep = "A"
            elif "-" in typrep and len(typrep) >= 8:
                typrep = "A"
            # 报表类型取 A（合并报表年报优先），避免重复叠加
            if typrep and typrep not in ("A", "1", "A-合并"):
                continue
            accper_d = accper[:10]
            if not dry:
                for item_cn, _ in (list(KEY_ITEMS.items()) if stmt == "income"
                                   else list(BALANCE_ITEMS.items()) if stmt == "balance"
                                   else list(CASH_ITEMS.items())):
                    col = colmap.get(item_cn)
                    if col is None:
                        continue
                    val = _to_num(r.get(col))
                    if val is None:
                        continue
                    cur.execute(
                        """INSERT OR REPLACE INTO financial_statements
                           (stkcd, short_name, accper, typrep, stmt_type, item, value, source)
                           VALUES (?,?,?,?,?,?,?,?)""",
                        (code, name, accper_d, typrep or "A", stmt, item_cn, val,
                         f"456/{industry}/{sheet}"),
                    )
                    stats["rows"] += 1
            stats["companies"].add((code, name))
            # 记录每个公司最新报告期
            cur_key = code
            if cur_key not in stats["latest"] or accper_d > stats["latest"][cur_key][0]:
                stats["latest"][cur_key] = (accper_d, name, stmt)
    if not dry:
        conn.commit()
    print(f"  {fn}: 明细 {stats['rows']} 行 / 公司 {len(stats['companies'])} 家")
    return stats


def aggregate_company_metrics(conn):
    """从 financial_statements 按最新报告期聚合公司核心指标 -> company_financial。"""
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT stkcd, short_name FROM financial_statements")
    companies = cur.fetchall()
    n = 0
    for code, name in companies:
        # 最新报告期
        cur.execute("SELECT MAX(accper) FROM financial_statements WHERE stkcd=?", (code,))
        latest = cur.fetchone()[0]
        if not latest:
            continue
        vals = {}
        cur.execute("SELECT stmt_type, item, value FROM financial_statements "
                    "WHERE stkcd=? AND accper=?", (code, latest))
        for stmt, item, value in cur.fetchall():
            key = KEY_ITEMS.get(item) or BALANCE_ITEMS.get(item) or CASH_ITEMS.get(item)
            if key:
                vals.setdefault(key, value)
        if not vals.get("revenue") and not vals.get("total_assets"):
            continue
        roe = vals.get("roe")
        if roe is None and vals.get("net_profit") and vals.get("total_equity"):
            try:
                roe = vals["net_profit"] / vals["total_equity"] * 100.0
            except Exception:
                roe = None
        margin = None
        if vals.get("net_profit") is not None and vals.get("revenue"):
            try:
                margin = vals["net_profit"] / vals["revenue"] * 100.0
            except Exception:
                margin = None
        gross = None
        if vals.get("revenue") is not None and vals.get("oper_cost") is not None:
            try:
                gross = (vals["revenue"] - vals["oper_cost"]) / vals["revenue"] * 100.0
            except Exception:
                gross = None
        try:
            cur.execute(
                """INSERT OR REPLACE INTO company_financial
                   (company_name, industry, year, roe, margin, turnover, multiplier,
                    cashflow, pain_point, gross_margin, total_assets, total_liability,
                    eps, data_as_of, data_source)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (name, "456导入", latest[:4], round(roe, 2) if roe is not None else 0,
                 round(margin, 2) if margin is not None else 0, 0, 0,
                 vals.get("oper_cashflow") or 0, "",
                 round(gross, 2) if gross is not None else 0,
                 vals.get("total_assets") or 0, vals.get("total_liability") or 0,
                 vals.get("eps") or 0, latest,
                 "456行业财务数据导入（CSMAR口径，最新报告期）"),
            )
            n += 1
        except Exception as e:
            print(f"  [聚合失败] {name}: {e}")
    conn.commit()
    print(f"  公司核心指标聚合更新: {n} 家")
    return n


def import_quote_file(fn, conn, dry=False):
    """导入个股日行情（每个 sheet 一只股票）。"""
    fp = os.path.join(SRC_DIR, fn)
    xl = pd.ExcelFile(fp)
    cur = conn.cursor()
    rows = 0
    for sheet in xl.sheet_names:
        if sheet.lower().startswith("sheet"):
            continue
        try:
            df = xl.parse(sheet)
        except Exception as e:
            print(f"   [skip] {fn}/{sheet}: {e}")
            continue
        df.columns = [str(c).strip() for c in df.columns]
        code = str(df.iloc[0].get("证券代码", "")).strip()
        name = sheet.split(" ", 1)[-1] if " " in sheet else sheet
        for _, r in df.iterrows():
            d = str(r.get("交易日期", ""))[:10]
            if not d:
                continue
            if not dry:
                cur.execute(
                    """INSERT OR REPLACE INTO stock_daily_quotes
                       (code, name, trade_date, open, high, low, close, volume, amount, source)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (code, name, d,
                     _to_num(r.get("开盘价")), _to_num(r.get("最高价")),
                     _to_num(r.get("最低价")), _to_num(r.get("收盘价")),
                     _to_num(r.get("成交数量(股)")), _to_num(r.get("成交金额(元)")),
                     f"456/{fn}"),
                )
            rows += 1
    if not dry:
        conn.commit()
    print(f"  {fn}: 行情 {rows} 行")
    return rows


def import_sector_file(fn, conn, dry=False):
    fp = os.path.join(SRC_DIR, fn)
    xl = pd.ExcelFile(fp)
    df = xl.parse(xl.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]
    cur = conn.cursor()
    rows = 0
    for _, r in df.iterrows():
        if not dry:
            cur.execute(
                """INSERT OR REPLACE INTO sector_quotes
                   (sector, pct_chg, turnover_wan, amount_yi, net_inflow_yi,
                    up_count, down_count, leader_stock, latest_price, leader_pct,
                    quote_date, source)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (str(r.get("板块", "")), _to_num(r.get("涨跌幅(%)")),
                 _to_num(r.get("总成交量(万手)")), _to_num(r.get("总成交额(亿元)")),
                 _to_num(r.get("净流入(亿元)")), _to_num(r.get("上涨家数")),
                 _to_num(r.get("下跌家数")), str(r.get("领涨股", "")),
                 _to_num(r.get("最新价")), _to_num(r.get("涨跌幅(%).1")),
                 datetime.date.today().strftime("%Y-%m-%d"), f"456/{fn}"),
            )
        rows += 1
    if not dry:
        conn.commit()
    print(f"  {fn}: 板块行情 {rows} 行")
    return rows


def init_tables(conn):
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS financial_statements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            stkcd TEXT, short_name TEXT, accper TEXT, typrep TEXT,
            stmt_type TEXT, item TEXT, value REAL, source TEXT
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock_daily_quotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT, name TEXT, trade_date TEXT,
            open REAL, high REAL, low REAL, close REAL,
            volume REAL, amount REAL, source TEXT
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sector_quotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sector TEXT, pct_chg REAL, turnover_wan REAL, amount_yi REAL,
            net_inflow_yi REAL, up_count INTEGER, down_count INTEGER,
            leader_stock TEXT, latest_price REAL, leader_pct REAL,
            quote_date TEXT, source TEXT
        )""")
    # company_financial 已有扩展列检查
    for col, ddl in [("gross_margin", "REAL DEFAULT 0"), ("total_assets", "REAL DEFAULT 0"),
                     ("total_liability", "REAL DEFAULT 0"), ("eps", "REAL DEFAULT 0"),
                     ("data_as_of", "TEXT DEFAULT ''"), ("data_source", "TEXT DEFAULT ''")]:
        try:
            cur.execute(f"ALTER TABLE company_financial ADD COLUMN {col} {ddl}")
        except Exception:
            pass
    conn.commit()


def main():
    dry = "--dry-run" in sys.argv
    conn = sqlite3.connect(DB_PATH)
    init_tables(conn)
    print(f"源目录: {SRC_DIR} | dry-run={dry}")
    total = {"stmt": 0, "quote": 0, "sector": 0}
    # 1) 财务报表
    for fn in STMT_FILES:
        fp = os.path.join(SRC_DIR, fn)
        if os.path.exists(fp):
            st = import_statement_file(fn, conn, dry)
            total["stmt"] += st["rows"]
    # 2) 个股行情
    for fn in QUOTE_FILES:
        fp = os.path.join(SRC_DIR, fn)
        if os.path.exists(fp):
            total["quote"] += import_quote_file(fn, conn, dry)
    # 3) 板块行情
    if os.path.exists(os.path.join(SRC_DIR, SECTOR_FILE)):
        total["sector"] += import_sector_file(SECTOR_FILE, conn, dry)
    # 4) 聚合公司指标
    if not dry:
        aggregate_company_metrics(conn)
    conn.close()
    print("完成:", total)


if __name__ == "__main__":
    main()
