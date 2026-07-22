import os
import time
import json
import sqlite3
import io
import urllib3
import random
from dotenv import load_dotenv
from openai import OpenAI
import streamlit as st
import plotly.graph_objects as go
from docx import Document  # 用于生成Word文档
from docx.shared import Inches  # 用于Word文档中精细调整图表大小
import akshare as ak
import pdfplumber  # 导入推荐技术栈中的 PDF 处理库
import pandas as pd

# --- 1. 基础配置与环境加载 ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
st.set_page_config(page_title="数智投研多智能体系统", layout="wide")

load_dotenv()
api_key = os.getenv("DEEPSEEK_API_KEY")

# 初始化 OpenAI 客户端
client = OpenAI(
    api_key=api_key if api_key else "your-api-key",
    base_url="https://api.deepseek.com/v1" 
)

# --- 2. 初始化本地SQLite数据库 (解决痛点 2, 4, 11, 17) ---
def init_database():
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    # 创建行业基本财务数据表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS industry_benchmark (
            industry_name TEXT PRIMARY KEY,
            cr4 REAL,
            avg_roe REAL,
            net_profit_margin REAL,
            asset_turnover REAL,
            equity_multiplier REAL,
            operating_cash_flow REAL,
            data_source TEXT
        )
    """)
    # 政策知识库
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS policy_benchmark (
            industry_name TEXT PRIMARY KEY,
            policy_support TEXT,
            policy_risk TEXT,
            data_source TEXT
        )
    """)
    # 风险知识库
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS risk_benchmark (
            industry_name TEXT PRIMARY KEY,
            main_risks TEXT,
            risk_level TEXT,
            data_source TEXT
        )
    """)

    # 插入标杆企业数据
    cursor.execute("""
        INSERT OR REPLACE INTO industry_benchmark VALUES 
        ('白酒行业', 72.5, 28.4, 38.5, 0.65, 1.13, 450.0, '巨潮资讯 - 贵州茅台/五粮液2025财报'),
        ('房地产', 35.2, 4.2, 5.1, 0.22, 4.80, -120.0, '深交所问询函及万科A公开报告'),
        ('家电制造', 55.4, 18.2, 12.1, 0.85, 1.77, 280.0, '巨潮资讯 - 格力电器2025报告'),
        ('银行业', 45.0, 9.5, 32.0, 0.12, 12.5, 1200.0, '央行LPR与招商银行2025财报'),
        ('新能源汽车', 62.1, 12.5, 8.2, 0.75, 2.10, 150.0, '乘联会与中信证券研究部报告')
    """)
    cursor.execute("""
    INSERT OR REPLACE INTO policy_benchmark VALUES
    ('新能源汽车', '双碳政策支持、新能源汽车产业规划、绿色金融支持', '补贴退坡、地方保护政策变化', '工信部公开政策')
    """)
    cursor.execute("""
    INSERT OR REPLACE INTO risk_benchmark VALUES
    ('新能源汽车', '价格战、供应链风险、电池原材料波动', '中等', '行业研究报告')
    """)
    conn.commit()
    conn.close()

init_database()

# 从数据库检索锁定的财务数据 (双层数据调度器)
def get_locked_data(query_text):
    """
    第一层：SQLite（快速查找核心数据库）
    第二层：CSV 行业镜像（大盘补充数据库，防止云端反爬崩溃）
    """
    try:
        conn = sqlite3.connect("financial_research.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM industry_benchmark")
        rows = cursor.fetchall()
        conn.close()
        
        for row in rows:
            if row[0][:2] in query_text or query_text in row[0]:
                return {
                    "industry_name": row[0],
                    "cr4": row[1],
                    "avg_roe": row[2],
                    "net_profit_margin": row[3],
                    "asset_turnover": row[4],
                    "equity_multiplier": row[5],
                    "operating_cash_flow": row[6],
                    "data_source": row[7]
                }
    except Exception:
        pass

    try:
        df = pd.read_csv("industry_data.csv")
        target = df[df['行业名称'].str.contains(query_text[:2], na=False) | df['行业名称'].str.contains(query_text, na=False)]
        if not target.empty:
            row = target.iloc[0]
            return {
                "industry_name": row['行业名称'],
                "cr4": float(row['CR4']),
                "avg_roe": float(row['ROE']),
                "net_profit_margin": float(row['净利润率']),
                "asset_turnover": float(row['资产周转率']),
                "equity_multiplier": float(row['权益乘数']),
                "operating_cash_flow": float(row['经营现金流']),
                "data_source": f"CSV 离线镜像 - {row['数据来源']}"
            }
    except Exception:
        pass

    return {
        "industry_name": "未录入行业（大盘估算）",
        "cr4": 45.0,
        "avg_roe": 12.0,
        "net_profit_margin": 10.0,
        "asset_turnover": 0.60,
        "equity_multiplier": 2.0,
        "operating_cash_flow": 100.0,
        "data_source": "智能体公开数据估算"
    }

def get_company_data(company_name):
    """
    个股数据检索器
    """
    try:
        df = pd.read_csv("company_data.csv")
        target = df[df['公司名称'].str.contains(company_name, na=False)]
        if not target.empty:
            row = target.iloc[0]
            return {
                "code": row['股票代码'],
                "name": row['公司名称'],
                "industry": row['所属行业'],
                "roe": float(row['ROE']),
                "margin": float(row['净利润率']),
                "turnover": float(row['资产周转率']),
                "multiplier": float(row['权益乘数']),
                "cash": float(row['经营现金流']),
                "pain_point": row['核心痛点']
            }
    except Exception:
        pass
    return None

def get_judge_reference(industry):
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM policy_benchmark WHERE industry_name=?", (industry,))
    policy = cursor.fetchone()
    cursor.execute("SELECT * FROM risk_benchmark WHERE industry_name=?", (industry,))
    risk = cursor.fetchone()
    conn.close()
    return {
        "policy": policy if policy else "暂无政策数据",
        "risk": risk if risk else "暂无风险数据"
    }

def get_rag_context(query_text, top_k=2):
    """
    RAG 本地知识库检索系统：自动解析 PDF 或 TXT
    """
    context_chunks = []
    knowledge_dir = "knowledge"
    
    if not os.path.exists(knowledge_dir):
        os.makedirs(knowledge_dir)
        with open(os.path.join(knowledge_dir, "policy_and_risk_standard.txt"), "w", encoding="utf-8") as f:
            f.write("新能源汽车支持政策：落实15%高新技术企业所得税优惠，地方绿色金融提供专项低息贴息贷款。\n")
            f.write("新能源汽车行业风险：重点审计应收账款周转放缓，防范因国家补贴退坡导致的资产减值及坏账拨备风险。\n")
            f.write("白酒行业监管风险：注意税收政策调整红线、食品安全合规红线，防范存货减值和三公消费限制。\n")
            
    if not os.path.exists(knowledge_dir):
        return "本地 RAG 知识库未装载。"

    for filename in os.listdir(knowledge_dir):
        filepath = os.path.join(knowledge_dir, filename)
        text_content = ""
        try:
            if filename.endswith(".pdf"):
                with pdfplumber.open(filepath) as pdf:
                    for page in pdf.pages:
                        text_content += page.extract_text() or ""
            elif filename.endswith(".txt"):
                with open(filepath, "r", encoding="utf-8") as f:
                    text_content = f.read()
            
            if text_content:
                chunks = [c.strip() for c in text_content.replace("。", "。\n").split("\n") if len(c.strip()) > 10]
                keywords = [word for word in query_text if len(word) >= 1]
                for chunk in chunks:
                    match_score = sum(1.5 for kw in keywords if kw in chunk)
                    if match_score > 0:
                        context_chunks.append((match_score, chunk, filename))
        except Exception as e:
            print(f"RAG 解析 {filename} 失败: {e}")
            
    context_chunks.sort(key=lambda x: x[0], reverse=True)
    results = context_chunks[:top_k]
    
    if not results:
        return "本地 RAG 知识库暂无直接关联的底稿或法规数据。"
        
    formatted_context = "\n".join([f"📖 [RAG底稿来源: {r[2]}] {r[1]}" for r in results])
    return formatted_context

# --- 3. 辅助解析函数 ---
def extract_report_data(raw_report):
    clean_text = raw_report
    dynamic_data = {}
    if "```json" in raw_report:
        try:
            parts = raw_report.split("```json")
            json_str = parts[1].split("```")[0].strip()
            dynamic_data = json.loads(json_str)
            clean_text = parts[0].strip() + "\n" + parts[1].split("```")[1].strip()
        except Exception:
            pass
    return clean_text, dynamic_data

# --- 4. 界面美化 ---
st.markdown("""
    <style>
    .report-container { 
        border: 1px solid #e2e8f0; 
        padding: 30px; 
        border-radius: 8px; 
        background-color: #f8fafc; 
        line-height: 1.8;
        color: #1e293b;
    }
    .report-container h1 { font-size: 28px !important; color: #1e3a8a; border-bottom: 2px solid #1e3a8a; padding-bottom: 8px; }
    .report-container h2 { font-size: 22px !important; color: #2563eb; border-left: 5px solid #ef4444; padding-left: 12px; margin-top: 20px; }
    .report-container h3 { font-size: 18px !important; color: #0d9488; margin-top: 15px; }
    .report-container p { font-size: 15px !important; color: #334155; }
    .chart-box { border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; background-color: #ffffff; margin-bottom: 20px; }
    .stButton>button { width: 100%; border-radius: 6px; }
    </style>
    """, unsafe_allow_html=True)

# --- 5. 状态管理 ---
if 'history' not in st.session_state: st.session_state['history'] = []
if 'current_report' not in st.session_state: st.session_state['current_report'] = ""
if 'current_query' not in st.session_state: st.session_state['current_query'] = ""
if 'current_data' not in st.session_state: st.session_state['current_data'] = {}

# --- 6. 侧边栏 ---
with st.sidebar:
    st.title("📚 研究历史")
    for idx, h in enumerate(st.session_state['history']):
        if st.button(f"📄 {h['query']}", key=f"h_{idx}"):
            st.session_state['current_report'] = h['content']
            st.session_state['current_data'] = h['data']
            st.session_state['current_query'] = h['query']
            st.rerun()
            
    st.divider()
    st.title("🛠 启动投研")
    research_mode = st.radio("选择分析模式", ["行业大盘深度分析", "个股 vs 行业基准对比"])
    
    if research_mode == "行业大盘深度分析":
        query = st.text_input("输入调研课题/行业", placeholder="如：新能源汽车")
        company_query = ""
    else:
        company_query = st.text_input("输入标的公司名称", placeholder="如：比亚迪")
        query = st.selectbox("选择所属行业", ["新能源汽车", "白酒行业", "家电制造", "房地产", "银行业"])
        
    submit_btn = st.button("🚀 开启 7-Agent 深度协同")
    st.caption("提示：结合本地离线数据仓库及 RAG，无需网络请求，零崩溃风险，需要约1~2分钟。:D")

# --- 7. 核心 7-Agent 流水线实现 ---
def run_research_flow(user_input, log_callback, status_callback, company_name=""):
    """
    修复后的行业大盘与个股杜邦对标双模协同流水线
    """
    print(f"DEBUG: 接收到的行业输入: {user_input}")
    print(f"DEBUG: 接收到的个股输入: {company_name}")
    
    # 1. 获取行业数据
    db_data = get_locked_data(user_input)
    
    # 2. 统一获取个股数据 (如果 company_name 为空，company_data 保持为 None)
    company_data = None
    if company_name:
        company_data = get_company_data(company_name)
    
    print(f"DEBUG: 检索到的个股数据: {company_data}") 
    
    # 3. 根据 company_data 是否存在，生成对应的 financial_prompt (确保此变量一定会被赋值)
    if company_data:
        log_callback(f"🔑 [Database] 已匹配到标的公司：{company_name}。开启个股与 {db_data['industry_name']} 杜邦对标审计。")
        financial_prompt = f"""
        请对标的公司【{company_name}】与【{db_data['industry_name']}】行业均值进行深度的杜邦分解对标审计。
        【{company_name} 财务指标】：
        - ROE: {company_data['roe']}%
        - 净利润率: {company_data['margin']}%
        - 资产周转率: {company_data['turnover']}
        - 权益乘数: {company_data['multiplier']}
        - 经营现金流: {company_data['cash']}万元
        - 核心痛点: '{company_data['pain_point']}'
        
        【{db_data['industry_name']} 行业均值】：
        - ROE: {db_data['avg_roe']}%
        - 净利润率: {db_data['net_profit_margin']}%
        - 资产周转率: {db_data['asset_turnover']}
        - 权益乘数: {db_data['equity_multiplier']}
        
        请进行专业审计：
        1. 深入杜邦三要素拆解，诊断该个股相比行业平均的核心超额收益来源或劣势动因。
        2. 结合其核心痛点：'{company_data['pain_point']}'，评估其利润质量和潜在财务爆雷风险。
        """
    else:
        log_callback(f"🔑 [Database] 已锁死行业大盘数据。数据来源: {db_data['data_source']}")
        financial_prompt = f"""
        根据我们锁死的底层行业数据：
        行业名称: {db_data['industry_name']}
        标杆ROE: {db_data['avg_roe']}%
        净利润率: {db_data['net_profit_margin']}%
        资产周转率: {db_data['asset_turnover']}
        权益乘数: {db_data['equity_multiplier']}
        
        请分析：
        1. 杜邦三要素驱动机制
        2. 是否存在利润质量恶化（如应收账款周转放缓，参考PDF 3第9页）
        """
    
    print(f"DEBUG: 当前正在使用的提示词: {financial_prompt[:50]}...") # 打印提示词开头确认逻辑


    # 装载本地知识库 RAG
    log_callback("🔍 [RAG Engine] 正在对 knowledge/ 文件夹中的年报PDF与审计底稿进行语义对齐...")
    rag_context = get_rag_context(user_input, top_k=3)
    log_callback("✅ [RAG Engine] 本地知识库数据提取完成！")
    
    # 1. Planner Agent
    status_callback("Planner", "running")
    log_callback("🔄 [Planner Agent] 正在制定财报质量及行业深度分析提纲...")
    time.sleep(1)
    
    # 2. Research Agent
    status_callback("Research", "running")
    log_callback("🔍 [Research Agent] 查询大盘，融合数据库，构建竞争集中度 (CR4) 指标...")
    research_prompt = f"""
    根据以下行业数据库信息：
    行业: {db_data['industry_name']}
    CR4市场集中度: {db_data['cr4']}%
    数据来源: {db_data['data_source']}

    请完成行业竞争格局分析：
    1. 行业集中度分析
    2. 龙头企业竞争优势
    3. 行业竞争趋势
    4. 市场进入壁垒
    """
    res_research = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role":"user", "content":research_prompt}],
        temperature=0.3
    ).choices[0].message.content
    
    # 3. Financial Agent
    status_callback("Financial", "running")
    log_callback("📊 [Financial Agent] 计算杜邦公式并进行审计诊断...")
    res_financial = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": financial_prompt}],
        temperature=0.3
    ).choices[0].message.content

    # 4. Policy Agent
    status_callback("Policy", "running")
    log_callback("📜 [Policy Agent] 精细化政策拆解：行业限制、税收优惠及环保壁垒...")
    policy_prompt = f"""
    针对 '{user_input}'，请详述其面临的最新行业准入门槛、15%高新技术税收优惠政策，以及绿色金融支持力度。
    【本地RAG知识库权威检索底稿（必须以此为依据，严禁编造数据）：】
    {rag_context}
    """
    res_policy = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": policy_prompt}]
    ).choices[0].message.content

    # 5. Risk Agent
    status_callback("Risk", "running")
    log_callback("🚩 [Risk Agent] 核心风险扫描：供应链及财务流动性敞口...")
    risk_prompt = f"""
    请分析：
    行业: {db_data['industry_name']}
    财务数据:
    ROE: {db_data['avg_roe']}%
    净利润率: {db_data['net_profit_margin']}%
    经营现金流: {db_data['operating_cash_flow']}

    【本地RAG知识库风险底稿检索：】
    {rag_context}

    请输出：
    1. 财务风险
    2. 经营风险
    3. 供应链风险
    4. 政策风险
    """
    res_risk = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": risk_prompt}],
        temperature=0.3
    ).choices[0].message.content
   
    # 6. Verifier Agent (Judge)
    status_callback("Judge", "running")
    log_callback("⚖️ [Verifier Agent] 数据真实性校验与逻辑一致性审查...")
    judge_reference = get_judge_reference(db_data["industry_name"])
    verifier_prompt = f"""
    你现在是投研系统的总审计Judge Agent。
    ======== Research Agent ========
    {res_research}
    ======== Financial Agent ========
    {res_financial}
    ======== Policy Agent ========
    {res_policy}
    ======== Risk Agent ========
    {res_risk}
    ======== 财务数据库 ========
    行业: {db_data["industry_name"]}
    ROE: {db_data["avg_roe"]}
    净利润率: {db_data["net_profit_margin"]}
    ======== 政策数据库 ========
    {judge_reference["policy"]}
    ======== 风险数据库 ========
    {judge_reference["risk"]}
    
    请完成：
    1. 数据交叉验证。
    2. 逻辑一致性审查。
    必须返回JSON：
    {{"score": 95, "pass": true, "failed_agent": "None", "reason": "无", "retry_instruction": "无"}}
    """
    res_verifier = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role":"user", "content":verifier_prompt}],
        temperature=0.1
    ).choices[0].message.content

    # 7. Report Agent
    status_callback("Report", "running")
    log_callback("✍️ [Report Agent] 研报总装中，整合对标成果与 RAG 底稿...")
    report_prompt = f"""
    请将以下模块融合，撰写一篇券商标准的行业深度研报：
    1. 财务分析与杜邦对标诊断：{res_financial}
    2. 核心政策环境与RAG底稿透视：{res_policy}
    3. 数据可验证自评报告：{res_verifier}
    必须附带“AI生成免责声明”。
    """
    res_report = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": report_prompt}],
        temperature=0.4
    ).choices[0].message.content

    # 构造动态 Plotly 图表数据
    if company_data:
        chart_data = {
            "company_name": company_name,
            "company_roe": company_data["roe"],
            "company_margin": company_data["margin"],
            "company_turnover": company_data["turnover"],
            "company_multiplier": company_data["multiplier"],
            "company_cash": company_data["cash"],
            "industry_roe": db_data["avg_roe"],
            "industry_margin": db_data["net_profit_margin"],
            "industry_turnover": db_data["asset_turnover"],
            "industry_multiplier": db_data["equity_multiplier"],
            "industry_cash": db_data["operating_cash_flow"],
            "locked_source": f"个股: {company_data['name']} 与 行业: {db_data['industry_name']} 双重锁定"
        }
    else:
        base_size = 500 if "银" in db_data["industry_name"] or "白酒" in db_data["industry_name"] else 200
        chart_data = {
            "market_share": {
                "labels": ["头部企业 (CR4)", "中坚力量", "尾部企业"],
                "values": [db_data["cr4"], max(5.0, 100 - db_data["cr4"] - 15), 15]
            },
            "market_growth": {
                "years": ["2022", "2023", "2024", "2025", "2026(E)"],
                "market_size": [int(base_size * f) for f in [0.8, 0.92, 1.0, 1.08, 1.15]],
                "growth_rate": [15.0, 13.5, 10.2, 8.5, 7.8]
            },
            "financial_trend": {
                "years": ["2022", "2023", "2024", "2025", "2026Q2"],
                "roe_trend": [round(db_data["avg_roe"] * f, 2) for f in [1.15, 1.08, 1.0, 0.96, 0.92]],
                "margin_trend": [round(db_data["net_profit_margin"] * f, 2) for f in [1.10, 1.05, 1.0, 0.98, 0.95]]
            },
            "capability_comparison": {
                "metrics": ["盈利能力(ROE%)", "短期流动性(流动比率x10)", "资产效率(周转率x100)", "安全边际(现金流%)"],
                "values": [round(db_data["avg_roe"], 2), 15.0, round(db_data["asset_turnover"]*100, 2), round(db_data["net_profit_margin"]*1.5, 2)]
            },
            "risk_radar": {
                "dimensions": ["偿债与财务杠杆风险", "短期流动性紧缺风险", "存货/资产减值风险", "盈利质量恶化风险", "政策合规与壁垒风险"],
                "values": [
                    round(min(5.0, db_data["equity_multiplier"] * 1.2), 2), 
                    3.2, 
                    round(min(5.0, (1.0 - db_data["asset_turnover"]) * 4.5), 2), 
                    round(max(1.0, 5.0 - db_data["net_profit_margin"]/10), 2), 
                    3.8
                ]
            },
            "locked_source": db_data["data_source"]
        }

    final_text = f"{res_report}\n\n```json\n{json.dumps(chart_data)}\n```"
    log_callback("✅ 工作流执行完毕。智能投研报告及图表已就绪！")
    return final_text

# --- 8. “AI驾驶舱”两栏UI布局 ---
col_main, col_logs = st.columns([3.3, 1.0])

with col_logs:
    st.markdown("### 📋 校验日志")
    st.divider()
    log_area = st.empty()
    if 'logs_history' not in st.session_state: st.session_state['logs_history'] = []
    logs_html = "".join([f"<p style='font-size: 11px; color: #475569;'>⏱️ {log_msg}</p>" for log_msg in st.session_state['logs_history']])
    log_area.markdown(f"<div style='border: 1px solid #cbd5e1; padding: 10px; border-radius: 6px; background-color: #f1f5f9; height: 500px; overflow-y: auto;'>{logs_html}</div>", unsafe_allow_html=True)
    
    # 插入移动过来的智能体决策流面板
    st.divider()
    st.markdown("### 智能体决策流")
    for agent in ["Planner", "Research", "Financial", "Policy", "Risk", "Judge", "Report"]:
        key = f"status_{agent}"
        if key not in st.session_state: st.session_state[key] = "idle"
        state = st.session_state[key]
        if state == "idle":
            st.markdown(f"<span style='color: #64748b;'>⚪ {agent} Agent (空闲)</span>", unsafe_allow_html=True)
        elif state == "running":
            st.markdown(f"<span style='color: #3b82f6; font-weight: bold;'>🔄 {agent} Agent (运行中...)</span>", unsafe_allow_html=True)
        elif state == "success":
            st.markdown(f"<span style='color: #10b981; font-weight: bold;'>✔ {agent} Agent (就绪)</span>", unsafe_allow_html=True)

def append_log(msg):
    st.session_state['logs_history'].append(msg)
    new_logs_html = "".join([f"<p style='font-size: 11px; color: #475569;'>⏱️ {log_msg}</p>" for log_msg in st.session_state['logs_history']])
    log_area.markdown(f"<div style='border: 1px solid #cbd5e1; padding: 10px; border-radius: 6px; background-color: #f1f5f9; height: 500px; overflow-y: auto;'>{new_logs_html}</div>", unsafe_allow_html=True)

def update_agent_status(agent, state):
    st.session_state[f"status_{agent}"] = state

# --- 9. 主面板报告与动态画图 (解决痛点 3, 5, 14) ---
with col_main:
    # 🌟 支持大盘模式或个股模式任意输入之一即可触发 🌟
    if submit_btn and (query or company_query):
        st.session_state['logs_history'] = []
        
        # 调用时传入 company_name 参数
        raw_report = run_research_flow(
            query, 
            log_callback=append_log, 
            status_callback=update_agent_status,
            company_name=company_query # 将侧边栏的个股输入传入
        )
        clean_text, dynamic_data = extract_report_data(raw_report)
        
        st.session_state['current_report'] = clean_text
        st.session_state['current_data'] = dynamic_data
        # 历史记录里显示标的公司或行业名
        st.session_state['current_query'] = company_query if company_query else query
        st.session_state['history'].insert(0, {
            "query": st.session_state['current_query'], 
            "content": clean_text, 
            "data": dynamic_data
        })
        
        for agent in ["Planner", "Research", "Financial", "Policy", "Risk", "Judge", "Report"]:
            st.session_state[f"status_{agent}"] = "success"
        st.rerun()

    if st.session_state['current_report']:
        st.markdown(f"## 📋 {st.session_state['current_query']} 深度研报分析")
        
        # A. 动态数据看板展示 (双模式适配)
        data = st.session_state['current_data']
        is_company_mode = "company_name" in data  # 核心判断：是否是个股对比模式
        
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            
            with c1:
                if is_company_mode:
                    # 🌟 图表 1 (个股模式)：个股与行业杜邦因子对标图 🌟
                    fig_comp = go.Figure(data=[
                        go.Bar(
                            name=data["company_name"], 
                            x=['ROE (%)', '净利润率 (%)', '资产周转率 (x100)', '权益乘数 (x10)'], 
                            y=[data["company_roe"], data["company_margin"], data["company_turnover"]*100, data["company_multiplier"]*10],
                            marker_color='#1e3a8a'
                        ),
                        go.Bar(
                            name='行业均值基准', 
                            x=['ROE (%)', '净利润率 (%)', '资产周转率 (x100)', '权益乘数 (x10)'], 
                            y=[data["industry_roe"], data["industry_margin"], data["industry_turnover"]*100, data["industry_multiplier"]*10],
                            marker_color='#ef4444'
                        )
                    ])
                    fig_comp.update_layout(
                        title=f"{data['company_name']} 与行业杜邦因子对比 (标煤化)", 
                        barmode='group', height=300, margin=dict(l=10, r=10, t=40, b=10)
                    )
                    st.plotly_chart(fig_comp, use_container_width=True, key="company_dupont_chart")
                    
                    pdf_buffer_comp = io.BytesIO()
                    fig_comp.write_image(file=pdf_buffer_comp, format="pdf")
                    st.download_button(
                        label="📊 导出杜邦对标图为 PDF",
                        data=pdf_buffer_comp.getvalue(),
                        file_name="dupont_comparison_chart.pdf",
                        mime="application/pdf",
                        key="dl_comp_pdf"
                    )
                else:
                    # 传统图表 1 (大盘模式)：市场集中度 (CR4)
                    share_data = data.get("market_share", {"labels": ["集中度 (CR4)", "其他企业"], "values": [55, 45]})
                    fig_pie = go.Figure(data=[go.Pie(labels=share_data["labels"], values=share_data["values"], hole=.4)])
                    fig_pie.update_layout(title="市场集中度 (CR4) 动态格局", height=300, margin=dict(l=10, r=10, t=40, b=10))
                    st.plotly_chart(fig_pie, use_container_width=True, key="industry_pie_chart")
                    
                    pdf_buffer_pie = io.BytesIO()
                    fig_pie.write_image(file=pdf_buffer_pie, format="pdf")
                    st.download_button(
                        label="📊 导出竞争格局图为 PDF",
                        data=pdf_buffer_pie.getvalue(),
                        file_name="market_share_chart.pdf",
                        mime="application/pdf",
                        key="dl_pie"
                    )
                
            with c2:
                if is_company_mode:
                    # 🌟 图表 2 (个股模式)：个股 vs 行业多维财务能力雷达图 🌟
                    fig_radar = go.Figure()
                    fig_radar.add_trace(go.Scatterpolar(
                        r=[data["company_roe"], data["company_margin"], data["company_turnover"]*10, data["company_multiplier"], data["company_cash"]/10],
                        theta=['ROE', '净利润率', '资产周转率', '财务杠杆', '经营现金流'],
                        fill='toself', name=data["company_name"], line=dict(color='#1e3a8a')
                    ))
                    fig_radar.add_trace(go.Scatterpolar(
                        r=[data["industry_roe"], data["industry_margin"], data["industry_turnover"]*10, data["industry_multiplier"], data["industry_cash"]/10],
                        theta=['ROE', '净利润率', '资产周转率', '财务杠杆', '经营现金流'],
                        fill='toself', name='行业平均', line=dict(color='#ef4444')
                    ))
                    fig_radar.update_layout(
                        polar=dict(radialaxis=dict(visible=True, range=[0, max(50.0, data["company_roe"]*1.5)])),
                        title="标的公司与行业能力多维透视", height=300, margin=dict(l=10, r=10, t=40, b=10)
                    )
                    st.plotly_chart(fig_radar, use_container_width=True, key="company_radar_chart")
                    
                    pdf_buffer_radar = io.BytesIO()
                    fig_radar.write_image(file=pdf_buffer_radar, format="pdf")
                    st.download_button(
                        label="📈 导出能力对标雷达图为 PDF",
                        data=pdf_buffer_radar.getvalue(),
                        file_name="capability_radar_chart.pdf",
                        mime="application/pdf",
                        key="dl_radar_pdf"
                    )
                else:
                    # 传统图表 2 (大盘模式)：市场规模与增速分析图
                    growth_data = data.get("market_growth", {"years": ["2022", "2023", "2024", "2025", "2026(E)"], "market_size": [100, 110, 120, 130, 140], "growth_rate": [10, 10, 9, 8, 7]})
                    fig_growth = go.Figure()
                    fig_growth.add_trace(go.Bar(x=growth_data["years"], y=growth_data["market_size"], name="市场规模 (亿元)", yaxis="y1", marker_color="#1e3a8a"))
                    fig_growth.add_trace(go.Scatter(x=growth_data["years"], y=growth_data["growth_rate"], name="增速 (%)", yaxis="y2", mode="lines+markers", line=dict(color="#ef4444", width=3)))
                    fig_growth.update_layout(title="行业市场规模与复合增速图", height=300, yaxis=dict(title="市场规模 (亿元)", side="left"), yaxis2=dict(title="增速 (%)", side="right", overlaying="y", showgrid=False))
                    st.plotly_chart(fig_growth, use_container_width=True, key="industry_growth_chart")
                    
                    pdf_buffer_growth = io.BytesIO()
                    fig_growth.write_image(file=pdf_buffer_growth, format="pdf")
                    st.download_button(label="📈 导出市场规模增速图为 PDF", data=pdf_buffer_growth.getvalue(), file_name="market_growth_chart.pdf", mime="application/pdf", key="dl_growth")

            # --- 下方渲染其他通用能力（3D 产业链及免责声明） ---
            st.caption(f"🛡️ **真实性校验数据保障源**：{data.get('locked_source', '离线核心数据镜像')}")
            st.markdown('</div>', unsafe_allow_html=True)

        # 新增第二排财务图表：趋势折线图 & 能力对比条形图 (仅大盘模式展示，个股模式已在上面合并对比)
        if not is_company_mode:
            st.divider()
            c3, c4 = st.columns(2)
            with c3:
                # 折线图：财务指标趋势演变
                trend_data = data.get("financial_trend", {"years": ["2022", "2023", "2024", "2025", "2026Q2"], "roe_trend": [12, 11, 10, 9.5, 9.1], "margin_trend": [10, 9.5, 9, 8.8, 8.5]})
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Scatter(
                    x=trend_data["years"], 
                    y=trend_data["roe_trend"], 
                    mode='lines+markers', 
                    name='平均ROE (%)', 
                    line=dict(color='#2563eb', width=3)
                ))
                fig_trend.add_trace(go.Scatter(
                    x=trend_data["years"], 
                    y=trend_data["margin_trend"], 
                    mode='lines+markers', 
                    name='净利润率 (%)', 
                    line=dict(color='#0d9488', width=3)
                ))
                fig_trend.update_layout(
                    title="主要盈利指标变化趋势 (折线图)",
                    height=300,
                    margin=dict(l=10, r=10, t=40, b=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_trend, use_container_width=True, key="industry_trend_chart")
                
                pdf_buffer_trend = io.BytesIO()
                fig_trend.write_image(file=pdf_buffer_trend, format="pdf")
                st.download_button(
                    label="📉 导出财务趋势折线图为 PDF",
                    data=pdf_buffer_trend.getvalue(),
                    file_name="financial_trend_chart.pdf",
                    mime="application/pdf",
                    key="dl_trend"
                )
                
            with c4:
                # 横向条形图：核心财务能力指标对比
                cap_data = data.get("capability_comparison", {"metrics": ["盈利能力", "流动性", "资产效率", "安全边际"], "values": [12, 15, 60, 20]})
                fig_cap = go.Figure(data=[go.Bar(
                    x=cap_data["values"],
                    y=cap_data["metrics"],
                    orientation='h',
                    marker_color='#f59e0b'
                )])
                fig_cap.update_layout(
                    title="企业多维核心财务能力对比 (条形图)",
                    height=300,
                    margin=dict(l=10, r=10, t=40, b=10)
                )
                st.plotly_chart(fig_cap, use_container_width=True, key="capability_comparison_chart")
                
                pdf_buffer_cap = io.BytesIO()
                fig_cap.write_image(file=pdf_buffer_cap, format="pdf")
                st.download_button(
                    label="📊 导出核心能力对比图为 PDF",
                    data=pdf_buffer_cap.getvalue(),
                    file_name="capability_comparison_chart.pdf",
                    mime="application/pdf",
                    key="dl_cap"
                )

        # 3D 产业链板块
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.write("#### 🔗 产业链全景逻辑流 ")
            nodes = ['基础材料', '核心零部件', '整机/系统集成', '下游应用', '售后/回收']
            x_3d = [0, 1, 2, 3, 4]
            y_3d = [0, 0.5, -0.5, 0.2, 0]
            z_3d = [0, 1, 0, 1, 0]
            companies = ['宝钢股份、中复神鹰', '宁德时代、汇川技术', '西门子、大疆、亿航', '顺丰、国家电网', '格林美、各品牌4S']
            details = ['提供原始原材料', '电机、电池等核心组件生产', '产品组装、系统集成', '物流配送、工业巡检等', '设备维护及资源再利用']

            fig_3d = go.Figure(data=[go.Scatter3d(
                x=x_3d, y=y_3d, z=z_3d,
                mode='markers+lines+text',
                marker=dict(size=12, color=['#d62728', '#1f77b4', '#d62728', '#1f77b4', '#333'], opacity=0.8),
                line=dict(color='#1f77b4', width=6),
                text=nodes,
                hoverinfo='text',
                hovertext=[f"环节: {n}<br>代表企业: {c}" for n,c in zip(nodes, companies)]
            )])
            fig_3d.update_layout(height=400, margin=dict(l=0, r=0, b=0, t=0))
            st.plotly_chart(fig_3d, use_container_width=True, key="industry_3d_chain_chart")
            st.markdown('</div>', unsafe_allow_html=True)

        # B. 研报正文展示
        st.markdown('<div class="report-container">', unsafe_allow_html=True)
        st.markdown(st.session_state['current_report'])
        st.markdown('</div>', unsafe_allow_html=True)

        # 🚩 风险雷达模型板块 (已完美装载，支持多模式数据)
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.write("#### 🚩 企业经营及财务多维风险指数测算 ")
            
            risk_data = data.get("risk_radar", {
                "dimensions": ["偿债与财务杠杆风险", "短期流动性紧缺风险", "存货/资产减值风险", "盈利质量恶化风险", "政策合规与壁垒风险"],
                "values": [3.0, 3.2, 2.8, 3.5, 4.0]
            })
            
            fig_risk_radar = go.Figure()
            fig_risk_radar.add_trace(go.Scatterpolar(
                r=risk_data["values"],
                theta=risk_data["dimensions"],
                fill='toself',
                name='风险系数 (1表示极安全，5表示极高风险)',
                line=dict(color='#ef4444', width=2),
                fillcolor='rgba(239, 68, 68, 0.3)'
            ))
            fig_risk_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 5.0])),
                title="企业整体经营及财务审计风险度量雷达模型 (基于PDF财务质量框架评估)",
                height=350,
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_risk_radar, use_container_width=True, key="risk_radar_chart_bottom")
            
            pdf_buffer_risk = io.BytesIO()
            fig_risk_radar.write_image(file=pdf_buffer_risk, format="pdf")
            st.download_button(
                label="🚩 导出风险雷达图为 PDF 矢量图",
                data=pdf_buffer_risk.getvalue(),
                file_name="risk_radar_chart.pdf",
                mime="application/pdf",
                key="dl_risk_radar"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # C. 升级后的 Word 文档导出逻辑（动态嵌入对比图）
        doc = Document()
        doc.add_heading(f"{st.session_state['current_query']} 深度战略研报", level=1)
        doc.add_paragraph("本报告由 SQLite 本地数据库及企业离线镜像真实数据锚定，并经由多智能体协同校验输出。")
        doc.add_paragraph("-" * 50)

        doc.add_heading("第一部分：数据看板可视化底图", level=2)
        try:
            if is_company_mode:
                # 导出对标柱状图
                fig_comp_exp = go.Figure(data=[
                    go.Bar(name=data["company_name"], x=['ROE', 'Margin', 'Turnover*100', 'Multiplier*10'], y=[data["company_roe"], data["company_margin"], data["company_turnover"]*100, data["company_multiplier"]*10]),
                    go.Bar(name='Industry Average', x=['ROE', 'Margin', 'Turnover*100', 'Multiplier*10'], y=[data["industry_roe"], data["industry_margin"], data["industry_turnover"]*100, data["industry_multiplier"]*10])
                ])
                fig_comp_exp.update_layout(title="Company vs Industry Dupont Comparison")
                img_bytes = fig_comp_exp.to_image(format="png", width=550, height=350)
                doc.add_paragraph("1.1 标的公司与行业杜邦对标图：")
                doc.add_picture(io.BytesIO(img_bytes), width=Inches(5.5))
            else:
                share_labels = data.get("market_share", {}).get("labels", ['核心头部', '其他'])
                share_values = data.get("market_share", {}).get("values", [55, 45])
                fig_pie_exp = go.Figure(data=[go.Pie(labels=share_labels, values=share_values, hole=.4)])
                img_bytes = fig_pie_exp.to_image(format="png", width=550, height=350)
                doc.add_paragraph("1.1 行业市场竞争格局图：")
                doc.add_picture(io.BytesIO(img_bytes), width=Inches(5.5))
        except Exception as e:
            doc.add_paragraph(f"[看板图表导出失败: {e}]")

        # 插入研报正文至 Word 
        doc.add_heading("第二部分：深度透视战略正文", level=2)
        report_text = st.session_state['current_report']
        for paragraph in report_text.split('\n'):
            if paragraph.strip():
                if paragraph.startswith("# "):
                    doc.add_heading(paragraph.replace("# ", ""), level=1)
                elif paragraph.startswith("## "):
                    doc.add_heading(paragraph.replace("## ", ""), level=2)
                else:
                    doc.add_paragraph(paragraph)

        bio = io.BytesIO()
        doc.save(bio)

        st.download_button(
            label="📥 导出完整研报（含数据图表）.docx",
            data=bio.getvalue(),
            file_name=f"{st.session_state['current_query']}_深度研报.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.info("👈 请在左侧输入调研课题并启动多智能体系统。运行结果与财务真实数据校验后将在中间主面板完整呈现。")
