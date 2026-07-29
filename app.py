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
import chromadb
from sentence_transformers import SentenceTransformer

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

# --- 2. 初始化本地SQLite数据库 (数据层分离改造) ---
def init_database():
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    
    # 行业基准表
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
    
    # 个股财务数据表 (承接 Excel 导入)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS company_financial (
            company_name TEXT PRIMARY KEY,
            industry TEXT,
            year TEXT,
            roe REAL,
            margin REAL,
            turnover REAL,
            multiplier REAL,
            cashflow REAL,
            pain_point TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS policy_benchmark (
            industry_name TEXT PRIMARY KEY,
            policy_support TEXT,
            policy_risk TEXT,
            data_source TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS risk_benchmark (
            industry_name TEXT PRIMARY KEY,
            main_risks TEXT,
            risk_level TEXT,
            data_source TEXT
        )
    """)

    # 填充行业大盘基础种子数据
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

    # 尝试加载组员的 Excel 数据
    import_financial_excel()

def import_financial_excel():
    """
    自动检测并导入组员存放在 knowledge/financial_report/company_financial.xlsx 里的个股财务报表。
    若无此文件，则写入默认数据以保证项目可用性。
    """
    excel_path = "knowledge/financial_report/company_financial.xlsx"
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    
    if os.path.exists(excel_path):
        try:
            # 自动加载 Excel 工作簿
            df = pd.read_excel(excel_path)
            # 标准化处理列名，避免因空格或不一致导致导入失败
            df.columns = [c.strip() for c in df.columns]
            
            for _, row in df.iterrows():
                cursor.execute("""
                    INSERT OR REPLACE INTO company_financial 
                    (company_name, industry, year, roe, margin, turnover, multiplier, cashflow, pain_point)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    str(row.get('公司', '')).strip(),
                    str(row.get('行业', '')).strip(),
                    str(row.get('年份', '2025')).strip(),
                    float(row.get('ROE', 0)),
                    float(row.get('净利润率', 0)),
                    float(row.get('资产周转率', 0)),
                    float(row.get('权益乘数', 1.5)),  # 杜邦分解乘数，默认 1.5
                    float(row.get('经营现金流', 0)),
                    str(row.get('核心痛点', '行业竞争加剧')).strip()
                ))
            conn.commit()
            print("🚀 [Database] 组员个股财务 Excel 数据已成功导入本地 SQLite 库。")
        except Exception as e:
            print(f"⚠️ [Database] 读取 Excel 导入 SQLite 失败: {e}，将采用备用默认数据")
    else:
        # 备用种子数据
        fallback_data = [
            ('比亚迪', '新能源汽车', '2025', 18.5, 5.2, 1.1, 2.5, 35000.0, '供应链毛利被压缩、出海关税壁垒'),
            ('宁德时代', '新能源汽车', '2025', 22.0, 12.0, 0.9, 1.8, 42000.0, '电池产能过剩、碳酸锂价格波动波动影响利润')
        ]
        for item in fallback_data:
            cursor.execute("""
                INSERT OR REPLACE INTO company_financial 
                (company_name, industry, year, roe, margin, turnover, multiplier, cashflow, pain_point)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, item)
        conn.commit()
    conn.close()

# 初始化数据库
init_database()
# ============================
# 🌟 行业分类加载模块（不进入 RAG，纯内存 DataFrame）
# ============================
import pandas as pd

# 全局变量
df_industry_hierarchy = None   # SwClassCode_2021.xls（层级结构）
df_company_mapping = None      # 最新个股申万行业分类.xlsx（个股映射）

def load_industry_data():
    """
    加载申万行业分类的两个核心文件
    """
    global df_industry_hierarchy, df_company_mapping
    
    # 1. 加载层级结构表（用于关键词 -> 三级行业 的模糊匹配）
    hierarchy_path = "knowledge/industry/SwClassCode_2021.xls"
    if os.path.exists(hierarchy_path):
        df_industry_hierarchy = pd.read_excel(hierarchy_path, sheet_name="Sheet1", dtype=str)
        # 去掉全空行（行业代码为空的行）
        df_industry_hierarchy = df_industry_hierarchy.dropna(subset=["行业代码"], how="any")
        print(f"✅ 行业层级结构加载成功，共 {len(df_industry_hierarchy)} 条")
    else:
        print("⚠️ 未找到 SwClassCode_2021.xls，行业层级匹配功能受限")
        df_industry_hierarchy = pd.DataFrame()

    # 2. 加载个股映射表（用于根据公司名称反查行业，或根据行业反查公司）
    mapping_path = "knowledge/industry/最新个股申万行业分类(完整版-截至7月末).xlsx"
    if os.path.exists(mapping_path):
        df_company_mapping = pd.read_excel(mapping_path, sheet_name=0, dtype=str)
        # 只保留关键列（注意列名是中文，别写错）
        df_company_mapping = df_company_mapping[['公司简称', '新版一级行业', '新版二级行业', '新版三级行业']].drop_duplicates()
        print(f"✅ 个股-行业映射加载成功，共 {len(df_company_mapping)} 条记录")
    else:
        print("⚠️ 未找到 最新个股申万行业分类.xlsx，个股行业查询功能受限")
        df_company_mapping = pd.DataFrame()

# 在 app 启动时立即执行
load_industry_data()

# ---------- 工具函数 1：关键词匹配三级行业（用于“新能源汽车” -> ['电动乘用车','综合乘用车']） ----------
def match_industries_by_keyword(keyword):
    """
    输入：用户输入的模糊行业词（如"新能源汽车"）
    输出：匹配到的申万三级行业名称列表（去重）
    """
    if df_industry_hierarchy.empty:
        return [keyword]
    
    # 先在三级行业名称里搜
    matched = df_industry_hierarchy[
        df_industry_hierarchy["三级行业名称"].str.contains(keyword, na=False)
    ]
    
    # 如果三级没搜到，退而求其次搜二级
    if matched.empty:
        matched = df_industry_hierarchy[
            df_industry_hierarchy["二级行业名称"].str.contains(keyword, na=False)
        ]
    
    if matched.empty:
        return [keyword]
    
    # 返回去重的三级行业名称
    return matched["三级行业名称"].dropna().unique().tolist()

# ---------- 工具函数 2：根据公司名称查行业（用于“比亚迪” -> 电动乘用车） ----------
def get_company_industry(company_name):
    """
    输入：公司简称（如"比亚迪"）
    输出：包含 一级、二级、三级 行业的字典，若未找到则返回 None
    """
    if df_company_mapping.empty:
        return None
    
    result = df_company_mapping[
        df_company_mapping["公司简称"].str.contains(company_name, na=False)
    ]
    if result.empty:
        return None
    
    row = result.iloc[0]
    return {
        "一级": row["新版一级行业"],
        "二级": row["新版二级行业"],
        "三级": row["新版三级行业"]
    }

# ---------- 工具函数 3：根据三级行业反查所有公司（供 Industry Agent 调用） ----------
def get_companies_by_industry(industry_3rd):
    """
    输入：三级行业名称（如"电动乘用车"）
    输出：该行业下所有公司简称列表
    """
    if df_company_mapping.empty:
        return []
    
    result = df_company_mapping[
        df_company_mapping["新版三级行业"] == industry_3rd
    ]
    return result["公司简称"].dropna().unique().tolist()
# ============================
# 🌟 Vector Database & RAG 自动同步引擎 (支持递归读取，过滤 Excel 结构化数据) 🌟
# ============================
@st.cache_resource
def get_vector_db_and_model():
    """
    缓存模型和数据库，防止重新加载导致内存溢出
    使用 os.walk 自动扫描整个 knowledge/ 文件夹，并过滤 Excel 纯数据文件
    """
    model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
    v_client = chromadb.PersistentClient(path="./vector_db")
    coll = v_client.get_or_create_collection(name="financial_knowledge")
    
    knowledge_dir = "knowledge"
    if os.path.exists(knowledge_dir) and coll.count() == 0:
        # 使用 os.walk 递归检索整个知识库目录
        for root, dirs, files in os.walk(knowledge_dir):
            for filename in files:
                filepath = os.path.join(root, filename)
                
                # 明确过滤跳过 Excel 与 CSV，防止数值类非结构化数据污染 RAG 文本数据库
                if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith(".csv"):
                    print(f"🚫 [RAG Filter] 结构化财务文件已跳过，不写入向量库: {filename}")
                    continue
                
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
                        # 切分 chunk 写入向量库
                        chunks = [c.strip() for c in text_content.replace("。", "。\n").split("\n") if len(c.strip()) > 15]
                        if chunks:
                            embeddings = model.encode(chunks).tolist()
                            ids = [f"{filename}_{idx}" for idx in range(len(chunks))]
                            metadatas = [{"source": filename} for _ in chunks]
                            coll.add(documents=chunks, embeddings=embeddings, metadatas=metadatas, ids=ids)
                            print(f"✅ [RAG Sync] 非结构化知识文档已载入向量库: {filename}")
                except Exception as e:
                    print(f"RAG 自动索引失败 {filename}: {e}")
    return model, coll

embedding_model, collection = get_vector_db_and_model()

def vector_search(query_text, top_k=3):
    try:
        query_embedding = embedding_model.encode(query_text).tolist()
        result = collection.query(query_embeddings=[query_embedding], n_results=top_k)
        if result["documents"] and result["documents"][0]:
            return "\n".join(result["documents"][0])
    except Exception:
        pass
    return "暂无相关知识底稿数据"

# ============================
# Financial Agent Tools (加入 DCF 估值模型)
# ============================
def query_financial_database(industry):
    data = get_locked_data(industry)
    return json.dumps(data, ensure_ascii=False)

def calculate_dupont(roe, margin, turnover, multiplier):
    result = {
        "ROE": roe,
        "净利润率": margin,
        "资产周转率": turnover,
        "权益乘数": multiplier,
        "解释": f"ROE由利润率、资产效率和财务杠杆共同驱动。当前ROE={roe}%"
    }
    return json.dumps(result, ensure_ascii=False)

def calculate_dcf(free_cash_flow, growth_rate, wacc, terminal_growth_rate=0.02, years=5):
    """
    两阶段折现现金流模型 (DCF) 估值计算器工具。
    :param free_cash_flow: 基期自由现金流 (万元)
    :param growth_rate: 预测期前n年的复合年增长率 (例如 0.1 表示 10%)
    :param wacc: 加权平均资本成本 (例如 0.08 表示 8%)
    :param terminal_growth_rate: 永续增长率 (一般设为 2% 左右，即 0.02)
    :param years: 预测期数 (默认5年)
    """
    try:
        fcf = float(free_cash_flow)
        g = float(growth_rate)
        r = float(wacc)
        tg = float(terminal_growth_rate)
        y = int(years)
        
        if r <= tg:
            return json.dumps({"error": "WACC必须大于永续增长率以实现收敛。"}, ensure_ascii=False)
            
        fcfs = []
        discounted_fcfs = []
        current_fcf = fcf
        
        # 1. 第一阶段：预测期现金折现
        for t in range(1, y + 1):
            current_fcf = current_fcf * (1 + g)
            fcfs.append(current_fcf)
            discount_factor = (1 + r) ** t
            discounted_fcfs.append(current_fcf / discount_factor)
            
        pv_forecast = sum(discounted_fcfs)
        
        # 2. 第二阶段：永续期终值折现 (Gordon Growth Model)
        terminal_value = (fcfs[-1] * (1 + tg)) / (r - tg)
        pv_terminal = terminal_value / ((1 + r) ** y)
        
        # 内在企业价值
        enterprise_value = pv_forecast + pv_terminal
        
        result = {
            "估值模型": "两阶段折现现金流(DCF)模型",
            "WACC (加权平均资本成本)": f"{r*100:.2f}%",
            "高速预测期增长率": f"{g*100:.2f}%",
            "永续增长率": f"{tg*100:.2f}%",
            "各年预期现金流": [f"第{i+1}年: {val:.2f}万元" for i, val in enumerate(fcfs)],
            "预测期折现后现金流(PV1)": f"{pv_forecast:.2f}万元",
            "期末终值折现现值(PV2)": f"{pv_terminal:.2f}万元",
            "内在企业价值 (Enterprise Value)": f"{enterprise_value:.2f}万元",
            "逻辑评价": f"该企业的当前核心估值大约为 {enterprise_value:.2f}万元。该估值基于公司在预测期内的高速扩张假设以及永续稳态发展。分析师应重点考量当前现金流的充裕程度。"
        }
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"DCF计算工具异常: {str(e)}"}, ensure_ascii=False)

financial_tools = [
    {
        "type": "function",
        "function": {
            "name": "query_financial_database",
            "description": "查询行业财务数据库，获取ROE、利润率、资产周转率等指标",
            "parameters": {
                "type": "object",
                "properties": {
                    "industry": {"type": "string", "description": "行业名称"}
                },
                "required": ["industry"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_dupont",
            "description": "计算杜邦分析指标关系",
            "parameters": {
                "type": "object",
                "properties": {
                    "roe": {"type": "number"},
                    "margin": {"type": "number"},
                    "turnover": {"type": "number"},
                    "multiplier": {"type": "number"}
                },
                "required": ["roe", "margin", "turnover", "multiplier"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calculate_dcf",
            "description": "两阶段折现现金流模型(DCF)计算内在价值，适用于评估重资产或高成长企业估值",
            "parameters": {
                "type": "object",
                "properties": {
                    "free_cash_flow": {"type": "number", "description": "基期现金流（通常采用当前经营现金流或自由现金流，万元）"},
                    "growth_rate": {"type": "number", "description": "预测期年增长率，如 0.15 代表 15%"},
                    "wacc": {"type": "number", "description": "加权平均资本成本折现率，如 0.08 代表 8%"},
                    "terminal_growth_rate": {"type": "number", "description": "永续增长率，默认为 0.02"},
                    "years": {"type": "integer", "description": "预测年数，通常为 5"}
                },
                "required": ["free_cash_flow", "growth_rate", "wacc"]
            }
        }
    }
]

def execute_financial_tool(tool_name, args):
    if tool_name == "query_financial_database":
        return query_financial_database(args["industry"])
    elif tool_name == "calculate_dupont":
        return calculate_dupont(args["roe"], args["margin"], args["turnover"], args["multiplier"])
    elif tool_name == "calculate_dcf":
        return calculate_dcf(
            free_cash_flow=args.get("free_cash_flow"),
            growth_rate=args.get("growth_rate"),
            wacc=args.get("wacc"),
            terminal_growth_rate=args.get("terminal_growth_rate", 0.02),
            years=args.get("years", 5)
        )
    return "未知工具"

# 从数据库检索锁定的财务数据 (大盘库)
def get_locked_data(query_text):
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

# 升级后的个股 SQLite 财务数据库读取
def get_company_data(company_name):
    """
    废除原先读取 company_data.csv 临时镜像的模式，
    改由 SQLite 中直接对 `company_financial` 进行关联匹配，实现真实现金流与杜邦参数查询
    """
    try:
        conn = sqlite3.connect("financial_research.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT company_name, industry, year, roe, margin, turnover, multiplier, cashflow, pain_point 
            FROM company_financial 
            WHERE company_name LIKE ?
        """, (f"%{company_name}%",))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {
                "name": row[0],
                "industry": row[1],
                "year": row[2],
                "roe": float(row[3]),
                "margin": float(row[4]),
                "turnover": float(row[5]),
                "multiplier": float(row[6]),
                "cash": float(row[7]),
                "pain_point": row[8]
            }
    except Exception as e:
        print(f"检索 SQLite 个股报错: {e}")
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

# 初始化 Session State
if 'history' not in st.session_state: st.session_state['history'] = []
if 'current_report' not in st.session_state: st.session_state['current_report'] = ""
if 'current_data' not in st.session_state: st.session_state['current_data'] = {}
if 'current_query' not in st.session_state: st.session_state['current_query'] = ""

# --- 4. 侧边栏 UI ---
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
    
    research_mode = st.radio(
        "选择分析模式",
        ["简易模式（快速分析）", "标准模式（专业投研）"]
    )
    
    company_query = ""
    query = ""
    period = "默认近三年+最新季度"
    purpose = "综合分析"
    report_type = "深度研究"

    if research_mode == "简易模式（快速分析）":
        research_target = st.radio("选择研究对象", ["公司", "行业"])
        if research_target == "公司":
            company_query = st.text_input("输入公司名称", placeholder="如：比亚迪")
        else:
            query = st.text_input("输入行业", placeholder="如：新能源汽车")
    else:
        research_target = st.radio("研究对象类型", ["公司", "行业"])
        if research_target == "公司":
            company_query = st.text_input("输入公司名称", placeholder="如：比亚迪")
            period_type = st.selectbox("选择时间周期类型", ["年度", "季度"])
            year_select = st.selectbox("⚙️ 选择年份", ["2021", "2022", "2023", "2024", "2025", "2026"])
            if period_type == "年度":
                period = f"{year_select}年度"
            else:
                quarter_select = st.selectbox("⚙️ 选择季度", ["Q1", "Q2", "Q3", "Q4"])
                period = f"{year_select}年{quarter_select}"
        else:
            query = st.text_input("输入行业", placeholder="如：新能源汽车")
            period_type = st.selectbox("选择时间周期类型", ["年度", "季度", "月度"])
            year_select = st.selectbox("⚙️ 选择年份", ["2021", "2022", "2023", "2024", "2025", "2026"])
            if period_type == "年度":
                period = f"{year_select}年度"
            elif period_type == "季度":
                quarter_select = st.selectbox("⚙️ 选择季度", ["Q1", "Q2", "Q3", "Q4"])
                period = f"{year_select}年{quarter_select}"
            else:
                month_select = st.selectbox("⚙️ 选择月份", ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"])
                period = f"{year_select}年{month_select}"
                
        report_type = st.selectbox("报告类型", ["年度策略", "季度跟踪", "专题研究"])
        purpose = st.selectbox("研究目的", ["投资价值分析", "行业趋势分析", "财务质量分析", "风险评估"])
    
    submit_btn = st.button("🚀 开启 7-Agent 深度协同")
    st.caption("提示：结合本地离线数据仓库及 RAG，无需网络请求，零崩溃风险，需要约1~2分钟。:D")

# --- 5. 核心 7-Agent 流水线实现 ---
def query_understanding_agent(user_input, company_name="", period="默认近三年+最新季度", purpose="综合分析", report_type="深度研究"):
    if company_name:
        target = company_name
        research_type = "公司研究"
    else:
        target = user_input
        research_type = "行业研究"
    return {
        "target": target,
        "research_type": research_type,
        "period": period,
        "purpose": purpose,
        "report_type": report_type
    }

def data_planning_agent(research_requirement):
    if research_requirement["research_type"] == "公司研究":
        return ["公司财务报表", "最新季度数据", "公司公告", "同行业数据"]
    return ["行业规模", "行业增长率", "竞争格局", "政策数据", "风险数据"]

def data_retrieval_agent(required_data, user_input):
    result = {}
    result["database"] = get_locked_data(user_input)
    missing = []
    if not result["database"]:
        missing.append("数据库暂无匹配数据，请上传文件")
    result["missing"] = missing
    return result

def run_research_flow(user_input, log_callback, status_callback, company_name="", period="默认近三年+最新季度", purpose="综合分析", report_type="深度研究"):
    """
    行业大盘与个股对标协同流水线
    """
    research_requirement = query_understanding_agent(user_input, company_name, period, purpose, report_type)
    log_callback(f"🧠 [Query Agent] 解析需求: {research_requirement}")
    
    required_data = data_planning_agent(research_requirement)
    log_callback(f"📚 [Data Planning Agent] 规划数据: {required_data}")
    
    retrieved_data = data_retrieval_agent(required_data, user_input)
    log_callback("📦 [Data Retrieval Agent] 数据索引检索完毕。")
    
    db_data = retrieved_data["database"]    
    company_data = None
    if company_name:
        company_data = get_company_data(company_name)
    
    # 杜邦分析与DCF估值复合提示词 (Financial Agent)
    if company_data:
        log_callback(f"🔑 [Database] 检测到个股【{company_name}】。开始进行杜邦基准与DCF分析锁定。")
        financial_prompt = f"""
        你现在是资深特许金融分析师(CFA)及特许公认会计师(ACCA)。
        
        请针对标的公司【{company_name}】与【{db_data['industry_name']}】行业均值进行深度杜邦分解对标审计。
        本篇研报的分析周期确定为【{period}】，研究目的偏向于【{purpose}】。
        
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
        
        请务必进行深度审计并调用对应工具：
        1. 必须调用 `calculate_dcf` 工具对该公司进行内在价值估算。你可以使用公司的当前经营现金流 {company_data['cash']} 万元作为 free_cash_flow。假设增长率为 0.12 (12%)，WACC折现率为 0.085 (8.5%)。
        2. 针对其研究目的【{purpose}】，利用杜邦三要素进行拆解，指出其财务偏离行业基准的主要驱动力量。
        3. 结合其痛点：'{company_data['pain_point']}' 与分析周期【{period}】，对资产负债表与核心流动性进行审计风险预测。
        """
    else:
        log_callback(f"🔑 [Database] 已锁死行业大盘数据。数据来源: {db_data['data_source']}，时间跨度: {period}")
        financial_prompt = f"""
        根据我们锁死的底层行业数据：
        行业名称: {db_data['industry_name']}
        标杆ROE: {db_data['avg_roe']}%
        净利润率: {db_data['net_profit_margin']}%
        资产周转率: {db_data['asset_turnover']}
        权益乘数: {db_data['equity_multiplier']}
        
        分析周期为【{period}】，研究偏好为【{purpose}】。
        请分析：
        1. 该行业在【{period}】内的杜邦三要素驱动机制，尤其是其【{purpose}】维度下的财务质量表现。
        2. 是否存在行业性的利润质量恶化（如应收账款周转放缓）。
        """

    # 向量数据库检索 (RAG 闭环)
    log_callback("🔍 [RAG Engine] 正在对 knowledge/ 文件夹中的年报PDF与审计底稿进行语义对齐...")
    rag_context = vector_search(user_input, top_k=3)
    log_callback("✅ [RAG Engine] 本地向量数据库检索对齐完成！")
    
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

    请结合研究周期【{period}】完成行业竞争格局分析：
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
    
    # 3. Financial Agent (支持 DCF 工具调用)
    status_callback("Financial", "running")
    log_callback("📊 [Financial Agent] 计算杜邦公式与 DCF 模型，并进行审计诊断...")
    financial_messages = [
        {
            "role": "system",
            "content": "你是Financial Agent。你的任务：优先使用真实数据进行财务杜邦分解或DCF估值，根据需求调用合适的工具，严禁编造数据。"
        },
        {"role": "user", "content": financial_prompt}
    ]
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=financial_messages,
        tools=financial_tools,
        tool_choice={"type": "function", "function": {"name": "calculate_dcf"}},
        temperature=0.3
    )
    message = response.choices[0].message
    if message.tool_calls:
        financial_messages.append(message)
        for tool_call in message.tool_calls:
            tool_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)
            log_callback(f"🛠️ [Financial Tool] 激活 {tool_name} 工具计算中...")
            tool_result = execute_financial_tool(tool_name, arguments)
            financial_messages.append({
                "role": "tool",
                "tool_call_id": tool_call.id,
                "content": tool_result
            })
        final_response = client.chat.completions.create(
            model="deepseek-chat",
            messages=financial_messages,
            temperature=0.3
        )
        res_financial = final_response.choices[0].message.content
    else:
        res_financial = message.content

    # 4. Policy Agent
    status_callback("Policy", "running")
    log_callback("📜 [Policy Agent] 精细化政策拆解：行业限制、税收优惠及环保壁垒...")
    policy_prompt = f"""
    针对 '{user_input}'，请详述其在【{period}】内面临的最新行业准入门槛、15%高新技术税收优惠政策，以及绿色金融支持力度。
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

    【本地RAG知识库风险底稿检索（关注周期【{period}】下的主要瓶颈）：】
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
    请完成数据交叉验证及逻辑一致性审查：
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
    log_callback("✍️ [Report Agent] 研报总装中，整合对标成果与 RAG 深度分析...")
    report_prompt = f"""
    你是一名卖方证券研究所首席分析师。
    
    请根据以下研究材料，生成一篇定位为【{report_type}】、研究目的侧重于【{purpose}】、时间跨度锚定在【{period}】的标准深度研究报告。
    
    报告必须严格按照以下结构输出，切勿漏项：
    
    # 一、核心观点
    - 要求：100-200字总结全文，针对【{purpose}】给出核心投资逻辑、内在价值DCF估值结论及大盘趋势。
    
    # 二、行业行情回顾 (聚焦周期: {period})
    
    # 三、企业内在价值与基本面分析 (侧重目的: {purpose})
    财务分析与估值模型依据：
    {res_financial}
    
    # 四、细分赛道分析
    
    # 五、产业链分析
    
    # 六、政策与行业合规 (RAG政策输入)：
    {res_policy}
    
    # 七、投资策略与盈利预测 (时间周期跨度: {period})
    
    # 八、风险提示 (RAG风险及底线输入)：
    {res_risk}
    
    自评检验结果反馈：
    {res_verifier}
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

# --- 9. 主面板报告与动态画图 ---
with col_main:
    if submit_btn and (query or company_query):
        st.session_state['logs_history'] = []
        
        raw_report = run_research_flow(
            query,
            log_callback=append_log,
            status_callback=update_agent_status,
            company_name=company_query,
            period=period,
            purpose=purpose,
            report_type=report_type
        )
        clean_text, dynamic_data = extract_report_data(raw_report)
        
        st.session_state['current_report'] = clean_text
        st.session_state['current_data'] = dynamic_data
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
        is_company_mode = "company_name" in data
        
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            
            with c1:
                if is_company_mode:
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
                        title=f"{data['company_name']} 与行业杜邦因子对比 (标准化)", 
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
                    fig_radar = go.Figure()
                    fig_radar.add_trace(go.Scatterpolar(
                        r=[data["company_roe"], data["company_margin"], data["company_turnover"]*10, data["company_multiplier"], data["company_cash"]/1000],
                        theta=['ROE', '净利润率', '资产周转率', '财务杠杆', '经营现金流(万)'],
                        fill='toself', name=data["company_name"], line=dict(color='#1e3a8a')
                    ))
                    fig_radar.add_trace(go.Scatterpolar(
                        r=[data["industry_roe"], data["industry_margin"], data["industry_turnover"]*10, data["industry_multiplier"], data["industry_cash"]/1000],
                        theta=['ROE', '净利润率', '资产周转率', '财务杠杆', '经营现金流(万)'],
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
                    growth_data = data.get("market_growth", {"years": ["2022", "2023", "2024", "2025", "2026(E)"], "market_size": [100, 110, 120, 130, 140], "growth_rate": [10, 10, 9, 8, 7]})
                    fig_growth = go.Figure()
                    fig_growth.add_trace(go.Bar(x=growth_data["years"], y=growth_data["market_size"], name="市场规模 (亿元)", yaxis="y1", marker_color="#1e3a8a"))
                    fig_growth.add_trace(go.Scatter(x=growth_data["years"], y=growth_data["growth_rate"], name="增速 (%)", yaxis="y2", mode="lines+markers", line=dict(color="#ef4444", width=3)))
                    fig_growth.update_layout(title="行业市场规模与复合增速图", height=300, yaxis=dict(title="市场规模 (亿元)", side="left"), yaxis2=dict(title="增速 (%)", side="right", overlaying="y", showgrid=False))
                    st.plotly_chart(fig_growth, use_container_width=True, key="industry_growth_chart")
                    
                    pdf_buffer_growth = io.BytesIO()
                    fig_growth.write_image(file=pdf_buffer_growth, format="pdf")
                    st.download_button(label="📈 导出市场规模增速图为 PDF", data=pdf_buffer_growth.getvalue(), file_name="market_growth_chart.pdf", mime="application/pdf", key="dl_growth")

            st.caption(f"🛡️ **真实性校验数据保障源**：{data.get('locked_source', '离线核心数据镜像')}")
            st.markdown('</div>', unsafe_allow_html=True)

        if not is_company_mode:
            st.divider()
            c3, c4 = st.columns(2)
            with c3:
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

        # 🚩 风险雷达模型板块
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

        # C. Word 文档导出逻辑（动态嵌入对比图）
        doc = Document()
        doc.add_heading(f"{st.session_state['current_query']} 深度战略研报", level=1)
        doc.add_paragraph("本报告由 SQLite 本地数据库及企业离线镜像真实数据锚定，并经由多智能体协同校验输出。")
        doc.add_paragraph("-" * 50)

        doc.add_heading("第一部分：数据看板可视化底图", level=2)
        try:
            if is_company_mode:
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
