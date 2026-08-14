# -*- coding: utf-8 -*-
"""
产品化功能模块（模块 B）：
1. 首屏示例问题一键填充
2. 上传年报 / 财务数据表解析（PDF / XLSX / CSV）
3. 报告收藏与分享链接（?report=<id> 直接加载）
4. 使用数据看板（研究次数 / 平均耗时 / Token 成本 / Agent 耗时分布）
"""
import io
import json
import os
import time

import pandas as pd
import pdfplumber
import streamlit as st

# 数据持久化目录（运行时自动创建，不提交到仓库）
DATA_DIR = "data"
FAVORITES_FILE = os.path.join(DATA_DIR, "favorites.json")
USAGE_FILE = os.path.join(DATA_DIR, "usage.json")
MAX_USAGE_RECORDS = 200

# 在线演示地址（用于生成分享链接）
APP_URL = "https://jbmelfwhrlgagcsrpvhkqy.streamlit.app"

# DeepSeek 定价（元 / 百万 token，按官方公开价格估算，可随时调整）
DEEPSEEK_INPUT_PRICE_PER_M = 2.0
DEEPSEEK_OUTPUT_PRICE_PER_M = 3.0

# 首屏示例问题：一键填充侧边栏参数并触发研究
EXAMPLE_QUESTIONS = [
    {
        "label": "🚗 比亚迪 · 投资价值分析",
        "target": "公司",
        "company": "比亚迪",
        "query": "",
        "purpose": "投资价值分析",
        "report_type": "年度策略",
        "period_type": "年度",
        "year": "2024",
    },
    {
        "label": "🔋 新能源汽车 · 行业趋势",
        "target": "行业",
        "company": "",
        "query": "新能源汽车",
        "purpose": "行业趋势分析",
        "report_type": "年度策略",
        "period_type": "年度",
        "year": "2024",
    },
    {
        "label": "🍶 贵州茅台 · 财务质量",
        "target": "公司",
        "company": "贵州茅台",
        "query": "",
        "purpose": "财务质量分析",
        "report_type": "年度策略",
        "period_type": "年度",
        "year": "2024",
    },
]

# 财务数据表字段别名（兼容中英文表头）
_COLUMN_ALIASES = {
    "name": ["公司", "公司名称", "公司简称", "名称", "company", "company_name"],
    "industry": ["行业", "所属行业", "industry"],
    "roe": ["ROE", "净资产收益率", "roe"],
    "margin": ["净利润率", "净利率", "margin"],
    "turnover": ["资产周转率", "总资产周转率", "turnover"],
    "multiplier": ["权益乘数", "multiplier"],
    "cash": ["经营现金流", "现金流", "cash", "cashflow"],
    "pain_point": ["核心痛点", "痛点", "pain_point"],
}


# ------------------------- 数据持久化 -------------------------
def _ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default


def save_json(path, obj):
    try:
        _ensure_data_dir()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ------------------------- 初始化与链接加载 -------------------------
def init_product_state():
    """初始化产品化功能所需的所有 session 状态。"""
    if "favorites" not in st.session_state:
        st.session_state["favorites"] = load_json(FAVORITES_FILE, [])
    if "usage_history" not in st.session_state:
        st.session_state["usage_history"] = load_json(USAGE_FILE, [])
    if "agent_times" not in st.session_state:
        st.session_state["agent_times"] = []
    if "usage_stats" not in st.session_state:
        st.session_state["usage_stats"] = {"input_tokens": 0, "output_tokens": 0}
    if "uploaded_companies" not in st.session_state:
        st.session_state["uploaded_companies"] = []
    if "uploaded_report_text" not in st.session_state:
        st.session_state["uploaded_report_text"] = ""
    if "share_text" not in st.session_state:
        st.session_state["share_text"] = ""
    if "share_link" not in st.session_state:
        st.session_state["share_link"] = ""

    # 通过分享链接 ?report=<id> 直接加载收藏的报告
    try:
        rid = st.query_params.get("report")
        if rid and not st.session_state.get("_share_loaded"):
            st.session_state["_share_loaded"] = True
            for fav in st.session_state.get("favorites", []):
                if str(fav.get("id")) == str(rid):
                    st.session_state["current_query"] = fav.get("query", "")
                    st.session_state["current_report"] = fav.get("report", "")
                    st.session_state["current_data"] = fav.get("data", {})
                    break
    except Exception:
        pass


# ------------------------- Agent 耗时与使用记录 -------------------------
def record_agent_time(agent, seconds):
    st.session_state.setdefault("agent_times", []).append(
        {"agent": agent, "seconds": round(float(seconds), 2)}
    )


def record_usage(query, duration_sec, tokens_in, tokens_out, agent_times):
    record = {
        "query": query,
        "duration_sec": round(float(duration_sec), 2),
        "tokens_in": int(tokens_in or 0),
        "tokens_out": int(tokens_out or 0),
        "agent_times": agent_times or [],
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    history = st.session_state.setdefault("usage_history", [])
    history.insert(0, record)
    st.session_state["usage_history"] = history[:MAX_USAGE_RECORDS]
    save_json(USAGE_FILE, st.session_state["usage_history"])


def estimate_cost(tokens_in, tokens_out):
    return (
        int(tokens_in or 0) / 1e6 * DEEPSEEK_INPUT_PRICE_PER_M
        + int(tokens_out or 0) / 1e6 * DEEPSEEK_OUTPUT_PRICE_PER_M
    )


# ------------------------- 收藏与分享 -------------------------
def add_favorite(query, report, data):
    fav = {
        "id": str(int(time.time() * 1000)),
        "query": query,
        "report": report,
        "data": data,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    st.session_state.setdefault("favorites", []).insert(0, fav)
    save_json(FAVORITES_FILE, st.session_state["favorites"])
    return fav


def delete_favorite(fid):
    st.session_state["favorites"] = [
        f for f in st.session_state.get("favorites", [])
        if str(f.get("id")) != str(fid)
    ]
    save_json(FAVORITES_FILE, st.session_state["favorites"])


def build_share_text(query, app_url=APP_URL):
    return (
        f"【AI 智能投研】《{query}》深度研究报告已生成\n"
        "由 7-Agent 多智能体流水线产出：行业研究 × 财务分析 × DCF 估值 × 多空辩论 × 专家委员会\n"
        "每条结论均带数据证据链（来源 / 字段 / 页码 / 可信等级），可溯源、可复核\n"
        f"在线体验：{app_url}"
    )


# ------------------------- 上传文件解析 -------------------------
def _to_float(value, default=0.0):
    try:
        if value is None:
            return default
        s = str(value).replace(",", "").replace("%", "").strip()
        if not s:
            return default
        return float(s)
    except Exception:
        return default


def _df_to_companies(df):
    """把 XLSX / CSV 表格转换为公司财务数据结构。"""
    if df is None or df.empty:
        return []
    df.columns = [str(c).strip() for c in df.columns]
    col = {}
    for key, aliases in _COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in df.columns:
                col[key] = alias
                break
    if "name" not in col:
        return []
    companies = []
    for _, row in df.iterrows():
        name = str(row.get(col["name"], "")).strip()
        if not name or name.lower() in ("nan", "none"):
            continue
        companies.append({
            "name": name,
            "industry": str(row.get(col.get("industry"), "")).strip(),
            "roe": _to_float(row.get(col.get("roe"))),
            "margin": _to_float(row.get(col.get("margin"))),
            "turnover": _to_float(row.get(col.get("turnover"))),
            "multiplier": _to_float(row.get(col.get("multiplier")), default=1.5),
            "cash": _to_float(row.get(col.get("cash"))),
            "pain_point": str(row.get(col.get("pain_point"), "")).strip(),
        })
    return companies


def _llm_extract_company(text, client):
    """用大模型从年报 PDF 文本中抽取结构化财务指标。"""
    prompt = (
        "请从下面这份上市公司年报摘要文本中，抽取关键财务指标，并严格只输出 JSON：\n"
        '{"company_name": "公司简称", "industry": "所属行业", "year": "报告年度", '
        '"roe": 净资产收益率(%), "margin": 净利润率(%), "turnover": 总资产周转率, '
        '"multiplier": 权益乘数, "cash": 经营现金流(万元), "pain_point": "一句话核心风险/痛点"}\n'
        "找不到的字段填 0 或空字符串。\n\n"
        f"年报文本（节选）：\n{text[:12000]}"
    )
    try:
        resp = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
        )
        raw = (resp.choices[0].message.content or "").strip()
        if raw.startswith("```"):
            raw = raw.strip("`")
            if raw.startswith("json"):
                raw = raw[4:]
        data = json.loads(raw)
        if not data.get("company_name"):
            return []
        return [{
            "name": str(data.get("company_name", "")).strip(),
            "industry": str(data.get("industry", "")).strip(),
            "year": str(data.get("year", "")).strip(),
            "roe": _to_float(data.get("roe")),
            "margin": _to_float(data.get("margin")),
            "turnover": _to_float(data.get("turnover")),
            "multiplier": _to_float(data.get("multiplier"), default=1.5),
            "cash": _to_float(data.get("cash")),
            "pain_point": str(data.get("pain_point", "")).strip(),
        }]
    except Exception:
        return []


def parse_uploaded_file(uploaded_file, client):
    """解析上传文件，返回 {"companies": [...], "report_text": "...", "error": "..."}。"""
    result = {"companies": [], "report_text": "", "error": ""}
    try:
        name = (uploaded_file.name or "").lower()
        raw = uploaded_file.getvalue()
        if name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw))
            result["companies"] = _df_to_companies(df)
        elif name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(raw))
            result["companies"] = _df_to_companies(df)
        elif name.endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
            text = text.strip()
            if len(text) < 100:
                result["error"] = "PDF 可提取文本过少，可能是扫描版，请改用 XLSX/CSV 数据表"
            else:
                result["report_text"] = text[:30000]
                result["companies"] = _llm_extract_company(text, client)
        else:
            result["error"] = "仅支持 PDF / XLSX / CSV 文件"
    except Exception as e:
        result["error"] = f"解析失败：{e}"
    return result
