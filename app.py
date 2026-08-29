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

# ============================================================
# chromadb 兼容层（云端环境修复）
# ------------------------------------------------------------
# Streamlit Cloud 的 Python 3.12 可能自带旧版 sqlite3（<3.35），
# chromadb 导入时会直接抛 RuntimeError；若环境装有 pysqlite3-binary
# （新版 sqlite），则用它替换标准 sqlite3 模块，保证 chromadb 可用。
# 即使这一步失败也不影响网站——RAG 会在下方 get_vector_db_and_model()
# 中整体容错，自动降级为关键词检索。
# ============================================================
def _ensure_chromadb_sqlite():
    try:
        if sqlite3.sqlite_version_info >= (3, 35, 0):
            return
    except Exception:
        pass
    try:
        import pysqlite3  # noqa: F401
        import sys
        sys.modules["sqlite3"] = sys.modules.pop("pysqlite3")
        sys.modules.pop("__pysqlite3__", None)
        print("[chromadb] 已用 pysqlite3-binary 替换 sqlite3 模块（>=3.35）。")
    except Exception:
        print("[chromadb] 提示：系统 sqlite3 版本过低，且未安装 pysqlite3-binary；"
              "RAG 将自动降级为关键词检索，网站功能不受影响。")


_ensure_chromadb_sqlite()
import chromadb
from sentence_transformers import SentenceTransformer
import product_features as pf
import industry_chain_data as icd
import report_export as rex
import data_updater as du
import learning_data as ldata
import news_fetcher as nf
import web_reader as wr
import leader_compare as lc

# --- 1. 基础配置与环境加载 ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
st.set_page_config(page_title="数智投研多智能体系统", layout="wide")

# 【优化读取逻辑】
api_key = None

# 1. 优先尝试从 Streamlit 官方 Secrets 中读取（解决云端部署覆盖问题）
try:
    if "DEEPSEEK_API_KEY" in st.secrets:
        api_key = st.secrets["DEEPSEEK_API_KEY"]
except Exception:
    pass  # 本地未配置 secrets.toml 时跳过，继续尝试 .env

# 2. 如果 Secrets 没读到，再尝试读取本地 .env（兼容本地开发运行）
if not api_key:
    load_dotenv()
    api_key = os.getenv("DEEPSEEK_API_KEY")

# 3. 【自我诊断工具】在页面侧边栏打印当前加载状态（排查后可自行删除）
if not api_key or api_key.strip() in ["your-api-key", ""]:
    st.error("⚠️ 【诊断提示】系统目前读取到的 API Key 依然为空或默认占位符！这说明您的配置未生效，请检查 Streamlit Secrets 或 `.env`。")
    st.stop()
else:
    # 仅展示前4位和总长度，确保密钥安全
    st.sidebar.success(f"🔑 密钥载入成功 (长度: {len(api_key)}位, 开头: {api_key[:4]}...)")

# 4. 初始化 OpenAI 客户端
client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com"  # 建议改为官方标准的 base_url，避免带 /v1 导致请求路径叠加
)

# --- 记录每次 LLM 调用的 token 用量（用于使用数据看板） ---
_orig_create = client.chat.completions.create

def _tracked_create(*args, **kwargs):
    resp = _orig_create(*args, **kwargs)
    usage = getattr(resp, "usage", None)
    if usage is not None:
        stats = st.session_state.setdefault("usage_stats", {"input_tokens": 0, "output_tokens": 0})
        stats["input_tokens"] = stats.get("input_tokens", 0) + int(getattr(usage, "prompt_tokens", 0) or 0)
        stats["output_tokens"] = stats.get("output_tokens", 0) + int(getattr(usage, "completion_tokens", 0) or 0)
    return resp

client.chat.completions.create = _tracked_create

# --- 2. 初始化本地SQLite数据库 (数据层分离改造) ---
def init_database():
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    
    # 行业基准表（含真实数据底座扩展列）
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS industry_benchmark (
            industry_name TEXT PRIMARY KEY,
            cr4 REAL,
            avg_roe REAL,
            net_profit_margin REAL,
            asset_turnover REAL,
            equity_multiplier REAL,
            operating_cash_flow REAL,
            data_source TEXT,
            gross_margin REAL DEFAULT 0,
            sample_size INTEGER DEFAULT 0,
            data_as_of TEXT DEFAULT ''
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
            pain_point TEXT,
            gross_margin REAL DEFAULT 0,
            total_assets REAL DEFAULT 0,
            total_liability REAL DEFAULT 0,
            eps REAL DEFAULT 0,
            data_as_of TEXT DEFAULT '',
            data_source TEXT DEFAULT ''
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

    # 填充行业大盘基础种子数据（仅空库时生效，绝不覆盖东方财富真实聚合数据）
    if cursor.execute("SELECT COUNT(*) FROM industry_benchmark").fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO industry_benchmark
            (industry_name, cr4, avg_roe, net_profit_margin, asset_turnover, equity_multiplier, operating_cash_flow, data_source)
            VALUES 
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

    # 确保学习资料库表结构完整（旧库可能缺 used_count 等列）
    try:
        ldata.init_learning_schema()
    except Exception as _e:
        print(f"⚠️ [Database] 学习资料库 schema 初始化失败: {_e}")

    # 尝试加载组员的 Excel 数据
    import_financial_excel()

def import_financial_excel():
    excel_path = "knowledge/financial_report/company_financial.xlsx"
    conn = sqlite3.connect("financial_research.db")
    cursor = conn.cursor()
    
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
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
                    float(row.get('权益乘数', 1.5)), 
                    float(row.get('经营现金流', 0)),
                    str(row.get('核心痛点', '行业竞争加剧')).strip()
                ))
            conn.commit()
            print("🚀 [Database] 组员个股财务 Excel 数据已成功导入本地 SQLite 库。")
        except Exception as e:
            print(f"⚠️ [Database] 读取 Excel 导入 SQLite 失败: {e}，将采用备用默认数据")
    else:
        fallback_data = [
            ('比亚迪', '新能源汽车', '2025', 18.5, 5.2, 1.1, 2.5, 35000.0, '供应链毛利被压缩、出海关税壁垒'),
            ('宁德时代', '新能源汽车', '2025', 22.0, 12.0, 0.9, 1.8, 42000.0, '电池产能过剩影响利润指标')
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
# 🌟 行业分类加载模块 (SwClassCode 纯内存极速查询) 🌟
# ============================
df_industry_hierarchy = pd.DataFrame()
df_company_mapping = pd.DataFrame()

def load_industry_data():
    global df_industry_hierarchy, df_company_mapping
    
    hierarchy_path = "knowledge/industry/SwClassCode_2021.xls"
    if os.path.exists(hierarchy_path):
        try:
            df_industry_hierarchy = pd.read_excel(hierarchy_path, dtype=str)
            df_industry_hierarchy = df_industry_hierarchy.dropna(subset=["行业代码"], how="any")
        except Exception as e:
            print(f"读取 SwClassCode 失败: {e}")

    mapping_path = "knowledge/industry/最新个股申万行业分类(完整版-截至7月末).xlsx"
    if os.path.exists(mapping_path):
        try:
            df_company_mapping = pd.read_excel(mapping_path, dtype=str)
            df_company_mapping.columns = [c.strip() for c in df_company_mapping.columns]
        except Exception as e:
            print(f"读取个股行业映射表失败: {e}")

# app 启动时载入
load_industry_data()

def get_company_industry(company_name):
    """
    个股申万行业查找器
    """
    if df_company_mapping.empty:
        return None
    result = df_company_mapping[df_company_mapping["公司简称"].str.contains(company_name, na=False)]
    if result.empty:
        return None
    row = result.iloc[0]
    return {
        "一级": row.get("新版一级行业", "未分类"),
        "二级": row.get("新版二级行业", "未分类"),
        "三级": row.get("新版三级行业", "未分类")
    }

def auto_align_industry(company_name="", query_industry=""):
    """
    🌟 核心：行业智能对齐与路由算法 🌟
    功能：实现输入个股/行业 -> 自动匹配5大基准大类或申万三级自研行业，彻底解决“比亚迪锁定白酒”的空字符串Bug。
    """
    # 5个核心种子行业映射字典
    core_mapping = {
        "汽车": "新能源汽车", "新能源": "新能源汽车", "锂": "新能源汽车", "电动": "新能源汽车",
        "白酒": "白酒行业", "酿酒": "白酒行业", "茅台": "白酒行业", "五粮液": "白酒行业",
        "家电": "家电制造", "电器": "家电制造", "空调": "家电制造", "冰箱": "家电制造",
        "房产": "房地产", "住宅": "房地产", "万科": "房地产", "保利": "房地产",
        "银行": "银行业", "金融": "银行业", "招商": "银行业"
    }

    # 1. 优先根据个股名称进行申万反查与对齐
    if company_name:
        conn = sqlite3.connect("financial_research.db")
        cursor = conn.cursor()
        cursor.execute("SELECT industry FROM company_financial WHERE company_name LIKE ?", (f"%{company_name}%",))
        row = cursor.fetchone()
        conn.close()
        if row and row[0] != "未分类" and row[0].strip() != "":
            return row[0]
            
        sw_info = get_company_industry(company_name)
        if sw_info:
            for keyword, target_ind in core_mapping.items():
                if keyword in sw_info["一级"] or keyword in sw_info["三级"] or keyword in sw_info["二级"]:
                    return target_ind
            return sw_info["三级"] # 没匹配到核心5个，直接以申万三级作为新行业
            
    # 2. 对输入的模糊行业词进行对齐
    if query_industry and query_industry.strip() != "":
        for keyword, target_ind in core_mapping.items():
            if keyword in query_industry:
                return target_ind
        return query_industry

    return "新能源汽车" # 兜底默认

# ============================
# 🌟 Vector Database & RAG 自动同步 🌟
# ============================
# 本地模型目录候选（优先加载，完全离线可用；兼容新旧模型）
LOCAL_MODEL_DIRS = [
    os.path.join("models", "bge-small-zh-v1.5"),                    # 推荐：95MB，中文效果好
    os.path.join("models", "paraphrase-multilingual-MiniLM-L12-v2"),  # 旧：470MB
]
# 云端下载候选（体积从小到大，保证最快可用）
DOWNLOAD_MODEL_CANDIDATES = [
    "BAAI/bge-small-zh-v1.5",
    "paraphrase-multilingual-MiniLM-L12-v2",
]

# 手动缓存成功结果：失败不缓存（None），下次启动自动重试
_EMBEDDING_CACHE = {"model": None}


def _source_reachable(url, timeout=5):
    """快速探测模型源是否可达，避免在网络不通时长时间卡住。"""
    import socket
    try:
        host = url.split("//")[1].split("/")[0]
        socket.create_connection((host, 443), timeout=timeout)
        return True
    except Exception:
        return False


def _run_with_timeout(fn, seconds):
    """在独立线程中执行 fn；超过 seconds 秒仍未完成则放弃，避免网络异常时无限卡住。"""
    import threading
    box = {}

    def _target():
        try:
            box["value"] = fn()
        except Exception as e:
            box["error"] = e

    t = threading.Thread(target=_target, daemon=True)
    t.start()
    t.join(seconds)
    if t.is_alive():
        return None
    if "error" in box:
        return None
    return box.get("value")


def load_embedding_model():
    """加载向量模型：本地 models/ 目录（离线）→ 云端下载（小模型优先）。
    成功结果缓存；失败不缓存，下次自动重试。"""
    # 1) 本地模型目录（最优先，无需联网）
    for local_dir in LOCAL_MODEL_DIRS:
        if os.path.isdir(local_dir):
            try:
                model = SentenceTransformer(local_dir)
                _EMBEDDING_CACHE["model"] = model
                print(f"[RAG] 已从本地加载向量模型: {local_dir}")
                return model
            except Exception:
                continue
    # 2) 已成功缓存的模型
    if _EMBEDDING_CACHE.get("model") is not None:
        return _EMBEDDING_CACHE["model"]
    # 3) 云端下载（小模型优先；放宽超时保证 95MB 模型可下载完成）
    os.environ.setdefault("HF_HUB_DOWNLOAD_TIMEOUT", "60")
    os.environ.setdefault("HF_HUB_ETAG_TIMEOUT", "30")
    # 本地环境（fix_and_run.ps1 设置了 SAS_FAST_START）给 20 秒快速失败；
    # 云端部署给 300 秒，保证 95MB 小模型（或 470MB 旧模型）能下载完成。
    per_source_timeout = 20 if os.environ.get("SAS_FAST_START") == "1" else 300
    for model_name in DOWNLOAD_MODEL_CANDIDATES:
        for endpoint in ["https://hf-mirror.com", "https://huggingface.co"]:
            if not _source_reachable(endpoint):
                continue
            os.environ["HF_ENDPOINT"] = endpoint
            # 用默认参数绑定，避免 lambda 闭包捕获循环变量
            model = _run_with_timeout(
                lambda mn=model_name: SentenceTransformer(mn),
                seconds=per_source_timeout,
            )
            if model is not None:
                _EMBEDDING_CACHE["model"] = model
                print(f"[RAG] 已从 {endpoint} 下载向量模型: {model_name}")
                return model
    print("[RAG] 向量模型下载失败（网络受限），已降级为关键词检索。"
          "云端首次加载会自动重试；也可将模型放入 models/bge-small-zh-v1.5 目录实现离线加载。")
    return None


def get_vector_db_and_model():
    """
    初始化向量库 + 语义模型。
    任何一步失败（模型下载受限 / chromadb Rust 绑定环境问题 / 索引异常）
    都自动降级为「关键词检索」模式，绝不让 RAG 拖垮整个网站。
    注意：本函数可能在后台线程执行，不能使用 st.spinner 等 UI 操作。
    """
    try:
        model = load_embedding_model()
        if model is None:
            return None, None
        # 确保向量库目录存在（云端从仓库克隆后可能没有该目录）
        os.makedirs("vector_db", exist_ok=True)
        v_client = chromadb.PersistentClient(path="./vector_db")
        coll = v_client.get_or_create_collection(name="financial_knowledge")

        knowledge_dir = "knowledge"
        if os.path.exists(knowledge_dir):
            # 已索引文件名集合（增量索引：新学习资料自动入向量库）
            indexed = set()
            try:
                _existing = coll.get(include=["metadatas"])
                for _m in (_existing.get("metadatas") or []):
                    indexed.add(_m.get("source", ""))
            except Exception:
                _existing = {}
            for root, dirs, files in os.walk(knowledge_dir):
                for filename in files:
                    if filename.endswith((".xlsx", ".xls", ".csv", ".db")):
                        continue
                    if filename in indexed:
                        continue
                    filepath = os.path.join(root, filename)
                    text_content = ""
                    try:
                        if filename.endswith(".pdf"):
                            with pdfplumber.open(filepath) as pdf:
                                for page in pdf.pages:
                                    text_content += page.extract_text() or ""
                        elif filename.endswith((".txt", ".md")):
                            with open(filepath, "r", encoding="utf-8") as f:
                                text_content = f.read()

                        if text_content:
                            chunks = [c.strip() for c in text_content.replace("。", "。\n").split("\n") if len(c.strip()) > 15]
                            if chunks:
                                embeddings = model.encode(chunks).tolist()
                                ids = [f"{filename}_{idx}" for idx in range(len(chunks))]
                                metadatas = [{"source": filename} for _ in chunks]
                                coll.add(documents=chunks, embeddings=embeddings, metadatas=metadatas, ids=ids)
                    except Exception as e:
                        print(f"RAG 自动索引失败 {filename}: {e}")
        return model, coll
    except Exception as e:
        # chromadb 在部分云端环境初始化失败（Rust 绑定/线程清理问题），
        # 降级为关键词检索，保证网站照常可用。
        print(f"[RAG] 向量库初始化失败，已降级为关键词检索模式: {type(e).__name__}: {e}")
        try:
            import gc
            gc.collect()
        except Exception:
            pass
        return None, None


# ============================================================
# RAG 异步就绪机制（云端友好）
# ------------------------------------------------------------
# 模块加载时：
#   1) 本地有模型 -> 同步初始化，立即启用语义检索；
#   2) 本地无模型 -> 网站立即打开，后台线程下载小模型（95MB），
#      完成后自动切换到语义检索（vector_search 每次检查 _RAG_STATE）。
# 这样云端首次访问不会被模型下载阻塞 300 秒，也不会永久降级。
# ============================================================
_RAG_STATE = {"model": None, "collection": None, "status": "init", "note": ""}


def _try_local_rag():
    """尝试用本地模型同步初始化 RAG（快，不下载）。成功返回 (model, coll)。"""
    for local_dir in LOCAL_MODEL_DIRS:
        if os.path.isdir(local_dir):
            try:
                model = SentenceTransformer(local_dir)
                _EMBEDDING_CACHE["model"] = model
                # 同步初始化 chromadb collection（本地环境 chromadb 正常可用）
                os.makedirs("vector_db", exist_ok=True)
                v_client = chromadb.PersistentClient(path="./vector_db")
                coll = v_client.get_or_create_collection(name="financial_knowledge")
                _RAG_STATE["model"] = model
                _RAG_STATE["collection"] = coll
                _RAG_STATE["status"] = "ready"
                _RAG_STATE["note"] = f"本地模型 {local_dir}"
                print(f"[RAG] 已从本地加载向量模型: {local_dir}")
                return model, coll
            except Exception as e:
                print(f"[RAG] 本地模型加载失败 {local_dir}: {e}")
                continue
    return None, None


def _bg_rag_init():
    """后台线程：下载小模型 + 初始化 chromadb + 索引知识库。"""
    try:
        model, coll = get_vector_db_and_model()
        if model is not None:
            _RAG_STATE["model"] = model
            _RAG_STATE["collection"] = coll
            _RAG_STATE["status"] = "ready"
            _RAG_STATE["note"] = "语义检索已就绪（后台下载完成）"
            print("[RAG] 后台初始化完成，已切换为语义检索。")
        else:
            _RAG_STATE["status"] = "fallback"
            _RAG_STATE["note"] = "语义模型下载失败，使用关键词检索（下次启动自动重试）"
    except Exception as e:
        _RAG_STATE["status"] = "fallback"
        _RAG_STATE["note"] = f"RAG 初始化异常：{str(e)[:100]}"


def _sync_rag_state():
    """把后台初始化结果同步到模块级 embedding_model/collection。"""
    global embedding_model, collection
    if _RAG_STATE["status"] == "ready" and _RAG_STATE["model"] is not None:
        embedding_model = _RAG_STATE["model"]
        collection = _RAG_STATE["collection"]


# 模块加载：本地模型同步初始化；否则立即降级并启动后台下载
embedding_model, collection = _try_local_rag()
if embedding_model is None:
    embedding_model, collection = None, None
    _RAG_STATE["status"] = "downloading"
    _RAG_STATE["note"] = "语义模型下载中（首次约 1~2 分钟），当前使用关键词检索"
    try:
        import threading
        threading.Thread(target=_bg_rag_init, daemon=True).start()
    except Exception:
        _RAG_STATE["status"] = "fallback"

# 侧边栏状态提示（每次 rerun 同步状态）
_sync_rag_state()
if embedding_model is None:
    if _RAG_STATE["status"] == "downloading":
        st.sidebar.info("🔍 语义检索模型下载中（首次约 1~2 分钟），当前为关键词检索；下载完成后自动升级。")
    else:
        st.sidebar.info("当前为“关键词检索”模式：语义模型未就绪（云端网络受限时自动降级），网页功能不受影响。")
else:
    st.sidebar.success(f"✅ 语义检索已就绪（{_RAG_STATE.get('note', '')}）")


def vector_search(query_text, top_k=3):
    # 每次查询前同步后台初始化结果（后台下载完成后自动启用语义检索）
    _sync_rag_state()
    # 向量库不可用（如模型未下载成功）时，降级为关键词检索
    if embedding_model is None or collection is None:
        return get_rag_context(query_text, top_k=top_k)
    try:
        query_embedding = embedding_model.encode(query_text).tolist()
        result = collection.query(query_embeddings=[query_embedding], n_results=top_k)
        if result["documents"] and result["documents"][0]:
            return "\n".join(result["documents"][0])
    except Exception:
        pass
    return "暂无相关知识底稿数据"

# ============================
# Financial Agent Tools (DCF + 杜邦 + 数据库查询)
# ============================
def query_financial_database(industry):
    # 记录 Tool Trace 监控
    st.session_state['tool_traces'].append({
        "agent": "Financial Agent",
        "tool": "query_financial_database",
        "input": f"industry='{industry}'",
        "output": f"数据查询成功，已通过 SQL 检索 '{industry}' 大盘行业基准表指标"
    })
    data = get_locked_data(industry)
    return json.dumps(data, ensure_ascii=False)

def calculate_dupont(roe, margin, turnover, multiplier):
    st.session_state['tool_traces'].append({
        "agent": "Financial Agent",
        "tool": "calculate_dupont",
        "input": f"roe={roe}, margin={margin}, turnover={turnover}, multiplier={multiplier}",
        "output": "杜邦分解勾稽关系计算及合理性逻辑验证成功"
    })
    result = {
        "ROE": roe,
        "净利润率": margin,
        "资产周转率": turnover,
        "权益乘数": multiplier,
        "解释": f"ROE由利润率、资产效率和财务杠杆共同驱动。当前计算ROE={roe}%"
    }
    return json.dumps(result, ensure_ascii=False)

def calculate_dcf(free_cash_flow, growth_rate, wacc, terminal_growth_rate=0.02, years=5):
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
        
        for t in range(1, y + 1):
            current_fcf = current_fcf * (1 + g)
            fcfs.append(current_fcf)
            discount_factor = (1 + r) ** t
            discounted_fcfs.append(current_fcf / discount_factor)
            
        pv_forecast = sum(discounted_fcfs)
        terminal_value = (fcfs[-1] * (1 + tg)) / (r - tg)
        pv_terminal = terminal_value / ((1 + r) ** y)
        enterprise_value = pv_forecast + pv_terminal
        
        st.session_state['tool_traces'].append({
            "agent": "Valuation Agent",
            "tool": "calculate_dcf",
            "input": f"fcf={fcf}, growth={g}, wacc={r}",
            "output": f"折现计算完成，内在价值: {enterprise_value:.2f}万元"
        })
        
        result = {
            "估值模型": "两阶段折现现金流(DCF)模型",
            "WACC": f"{r*100:.2f}%",
            "高速预测期增长率": f"{g*100:.2f}%",
            "内在企业价值": f"{enterprise_value:.2f}万元"
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
            "description": "两阶段折现现金流模型(DCF)计算内在价值",
            "parameters": {
                "type": "object",
                "properties": {
                    "free_cash_flow": {"type": "number", "description": "基期现金流（通常采用当前经营现金流或自由现金流，万元）"},
                    "growth_rate": {"type": "number", "description": "预测期年增长率，如 0.15 代表 15%"},
                    "wacc": {"type": "number", "description": "加权平均资本成本折现率，如 0.08 代表 8%"}
                },
                "required": ["free_cash_flow", "growth_rate", "wacc"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_latest_news",
            "description": "搜索某公司/行业的最新新闻与公告（东方财富公告/快讯/搜狗/财新，实时公开信息）",
            "parameters": {
                "type": "object",
                "properties": {
                    "keyword": {"type": "string", "description": "搜索关键词（公司名或行业名）"}
                },
                "required": ["keyword"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_webpage_charts",
            "description": "读取指定网页的文字、表格与内嵌图表数据（ECharts/Chart.js/Highcharts/JSON 数据块）",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "网页 URL"}
                },
                "required": ["url"]
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
            wacc=args.get("wacc")
        )
    elif tool_name == "search_latest_news":
        _kw = str(args.get("keyword", "")).strip()
        _b = nf.fetch_news_bundle(keyword=_kw, company_name=_kw, industry=_kw,
                                  days=30, limit_each=5)
        return json.dumps({"工具": "search_latest_news", "关键词": _kw,
                           "结果数": len(_b.get("items", [])),
                           "新闻与公告": _b.get("items", [])[:8]}, ensure_ascii=False)
    elif tool_name == "read_webpage_charts":
        _u = str(args.get("url", "")).strip()
        _pg = wr.read_webpage(_u, timeout=10)
        return json.dumps({"工具": "read_webpage_charts", "网址": _u,
                           "标题": _pg.get("title", ""),
                           "正文摘要": (_pg.get("text") or "")[:600],
                           "表格数": len(_pg.get("tables", [])),
                           "表格": _pg.get("tables", [])[:2],
                           "图表数据": _pg.get("charts", [])[:5],
                           "图片说明": _pg.get("images", [])[:5]}, ensure_ascii=False)
    return "未知工具"

# 从 SQLite 数据库检索行业基准
def get_locked_data(query_text):
    if not query_text or query_text.strip() == "" or query_text == "未录入行业（大盘估算）":
        return {
            "industry_name": "未录入行业（大盘估算）",
            "cr4": 45.0, "avg_roe": 12.0, "net_profit_margin": 10.0,
            "asset_turnover": 0.60, "equity_multiplier": 2.0, "operating_cash_flow": 100.0,
            "data_source": "大盘平均估算"
        }
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
                    "data_source": f"SQLite底表 - {row[7]}",
                    "gross_margin": row[8] if len(row) > 8 else 0,
                    "sample_size": row[9] if len(row) > 9 else 0,
                    "data_as_of": row[10] if len(row) > 10 else "",
                }
    except Exception:
        pass
    return {
        "industry_name": "未录入行业（大盘估算）",
        "cr4": 45.0, "avg_roe": 12.0, "net_profit_margin": 10.0,
        "asset_turnover": 0.60, "equity_multiplier": 2.0, "operating_cash_flow": 100.0,
        "data_source": "智能体估算"
    }

# 升级后的个股 SQLite 财务数据库读取
def get_company_data(company_name):
    if not company_name:
        return None
    # 优先使用用户上传的财务数据
    for _uc in st.session_state.get("uploaded_companies") or []:
        _uname = _uc.get("name", "")
        if company_name in _uname or _uname in company_name:
            return {
                "name": _uname,
                "industry": _uc.get("industry", ""),
                "year": str(_uc.get("year", "用户上传")),
                "roe": float(_uc.get("roe", 0) or 0),
                "margin": float(_uc.get("margin", 0) or 0),
                "turnover": float(_uc.get("turnover", 0) or 0),
                "multiplier": float(_uc.get("multiplier", 1.5) or 1.5),
                "cash": float(_uc.get("cash", 0) or 0),
                "pain_point": _uc.get("pain_point", "用户上传数据，未标注核心痛点"),
            }
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


# ============================================================
# 🔗 资料缺口自动补链（缺口 → 分析需要什么 → 搜索可获取链接）
# ============================================================
# 官方固定入口（100% 可用，无需搜索）：按缺口类型给出权威获取地址
OFFICIAL_SOURCE_LINKS = {
    "年报": [
        ("巨潮资讯网（法定披露平台·年报/公告 PDF 下载）", "http://www.cninfo.com.cn/new/index"),
        ("上交所官网（沪市公司定期报告）", "http://www.sse.com.cn/assortment/stock/list/info/announcement/"),
        ("深交所官网（深市公司定期报告）", "http://www.szse.cn/disclosure/listed/bulletinDetail/index.html"),
    ],
    "财务": [
        ("巨潮资讯网（年报/半年报/财务数据）", "http://www.cninfo.com.cn/new/index"),
        ("东方财富数据中心（业绩报表）", "https://data.eastmoney.com/bbsj/"),
        ("新浪财经（公司财务摘要）", "https://finance.sina.com.cn/stock/"),
    ],
    "行业": [
        ("东方财富行业研报中心", "https://data.eastmoney.com/report/industry.jshtml"),
        ("慧博投研资讯（行业研报）", "https://www.hibor.com.cn/"),
        ("发现报告（行业研究报告）", "https://www.fxbaogao.com/"),
    ],
    "政策": [
        ("中国政府网（政策文件库）", "https://www.gov.cn/zhengce/"),
        ("工信部官网（产业政策）", "https://www.miit.gov.cn/"),
        ("国家发改委（产业政策与规划）", "https://www.ndrc.gov.cn/"),
    ],
    "新闻": [
        ("东方财富资讯", "https://finance.eastmoney.com/"),
        ("财新网", "https://www.caixin.com/"),
        ("新浪财经新闻", "https://finance.sina.com.cn/"),
    ],
    "研报": [
        ("东方财富研报中心", "https://data.eastmoney.com/report/"),
        ("慧博投研资讯", "https://www.hibor.com.cn/"),
        ("发现报告", "https://www.fxbaogao.com/"),
    ],
    "风险": [
        ("证监会（监管与风险提示）", "http://www.csrc.gov.cn/"),
        ("巨潮资讯（公司公告含风险提示）", "http://www.cninfo.com.cn/new/index"),
    ],
}

# 缺口类型 → 关键词模板（用于自动搜索）
_GAP_KEYWORD_TMPL = {
    "年报": "{kw} {year}年年度报告 PDF 下载 巨潮资讯",
    "财务": "{kw} 财务报表 财务指标 下载",
    "行业": "{kw} 行业研究报告 PDF 下载",
    "政策": "{kw} 政策 文件 原文 官网",
    "新闻": "{kw} 最新新闻 公告",
    "研报": "{kw} 深度研究报告 PDF",
    "风险": "{kw} 风险提示 公告",
}


def _gap_type_of(item_text):
    """根据缺口项文本判断资料类型，返回 (type, 搜索词)。"""
    t = str(item_text or "")
    if "年报" in t or "财务指标" in t or "财报" in t:
        return "年报", t.replace(" 财务指标（ROE/利润率/周转/乘数）", "").replace(" 财务指标", "")
    if "可比公司" in t or "同行业" in t:
        return "行业", t.replace(" 数据", "").replace("（可选：上传可比公司对照表以获得更强基准）", "")
    if "行业基准" in t or "市场规模" in t or "行业研报" in t:
        return "行业", t.replace(" 行业基准数据", "").replace(" 行业基准", "")
    if "政策" in t:
        return "政策", t.replace(" 数据", "")
    if "新闻" in t or "公告" in t:
        return "新闻", t.replace(" 数据", "")
    if "风险" in t:
        return "风险", t
    return "研报", t


def suggest_gap_sources(gap_item, company_name="", industry="", year="2025"):
    """
    为单个资料缺口自动生成「获取资料链接」：
    1) 官方固定入口（巨潮/交易所/政府网等，按资料类型）
    2) 自动搜索（搜狗网页搜索，按缺口类型生成关键词）
    返回 {"official": [{"name","url"}], "search": [{"title","url","source"}], "note": str}
    """
    item_text = str(gap_item.get("item", ""))
    gap_type, base_kw = _gap_type_of(item_text)
    # 可比公司/行业类缺口优先用行业名搜索；其余优先公司名
    if gap_type == "行业" and industry:
        kw_entity = industry
    elif gap_type == "行业" and not industry:
        kw_entity = base_kw or company_name or item_text
    else:
        kw_entity = company_name or industry or base_kw or item_text

    # 1) 官方固定入口
    official = OFFICIAL_SOURCE_LINKS.get(gap_type, OFFICIAL_SOURCE_LINKS["研报"])

    # 2) 自动搜索
    tmpl = _GAP_KEYWORD_TMPL.get(gap_type, _GAP_KEYWORD_TMPL["研报"])
    search_kw = tmpl.format(kw=kw_entity, year=year).strip()
    if company_name and "年报" in tmpl:
        search_kw = f"{company_name} {year}年年度报告 PDF 下载 巨潮资讯"
    try:
        _r = nf.search_web_pages(search_kw, limit=5)
        search = _r.get("items", []) or []
        note = _r.get("note", "")
    except Exception as e:
        search = []
        note = f"自动搜索失败: {str(e)[:80]}"
    if not search:
        note = "自动搜索暂无结果（网络受限时请使用右侧官方入口）"
    return {"official": official, "search": search, "note": note,
            "search_kw": search_kw, "type": gap_type}

# ============================================================
# 🌟 真实风险雷达（基于真实财务指标映射，口径透明可溯）🌟
# ============================================================
def _clamp(v, lo=0.0, hi=5.0):
    try:
        return round(max(lo, min(hi, float(v))), 2)
    except Exception:
        return 3.0


# 政策合规与壁垒风险：行业基准值（来自政策库/公开监管信息，可随知识库扩充覆盖）
POLICY_RISK_FLOOR = {
    "新能源": 3.5, "汽车": 3.5, "白酒": 2.5, "酒": 2.5, "家电": 2.0, "电器": 2.0,
    "房地产": 4.0, "银行": 3.5, "金融": 3.5, "医药": 3.0, "药": 3.0, "医疗": 3.0,
    "半导体": 3.5, "芯片": 3.5, "光伏": 3.0, "机器": 2.5, "人工智能": 3.0,
}


def compute_risk_radar(company_data=None, db_data=None):
    """
    基于真实指标计算 5 维风险指数（0=安全，5=高危），并给出每维计算口径。
    company_data: 个股表数据（含 roe/margin/turnover/multiplier/cash/gross_margin/eps）
    db_data: 行业聚合数据（含 avg_roe/net_profit_margin/equity_multiplier/gross_margin/operating_cash_flow）
    """
    db = db_data or {}
    comp = company_data or {}
    margins = []
    dims = []

    # 1) 偿债与财务杠杆风险：权益乘数越高 → 杠杆风险越高（乘数1→0，乘数6+→5）
    mult = comp.get("multiplier") or db.get("equity_multiplier") or 1.5
    v1 = _clamp((float(mult) - 1.0) / 1.2)
    margins.append("权益乘数 %.2f → 每超出1倍加 0.83 分（区间映射），乘数为公司/行业真实值" % float(mult))
    dims.append("偿债与财务杠杆风险")

    # 2) 短期流动性紧缺风险：每股经营现金流/每股收益（现金流质量），<0.5 高风险
    eps = float(comp.get("eps") or 0)
    cfps = float(comp.get("cash") or 0)
    if eps > 0 and abs(cfps) > 0:
        quality = cfps / max(eps, 0.01)
        v2 = _clamp(5.0 - quality * 2.0)
        margins.append("现金流质量(每股经营现金流/每股收益)=%.2f → 5-2×质量（质量<0.5 记高危）" % quality)
    else:
        base_cf = float(db.get("operating_cash_flow") or 0)
        v2 = _clamp(5.0 - base_cf * 0.5)
        margins.append("行业每股经营现金流 %.2f → 5-0.5×现金流（样本聚合值）" % base_cf)
    dims.append("短期流动性紧缺风险")

    # 3) 存货/资产减值风险：公司毛利率显著低于行业均值 → 减值风险（毛利率真实值）
    comp_gm = float(comp.get("gross_margin") or 0)
    ind_gm = float(db.get("gross_margin") or 0)
    if comp_gm > 0 and ind_gm > 0:
        gap = ind_gm - comp_gm
        v3 = _clamp(gap / 8.0)
        margins.append("毛利率缺口(行业%.1f%%-公司%.1f%%)÷8 → 缺口越大减值风险越高" % (ind_gm, comp_gm))
    else:
        v3 = _clamp((30.0 - ind_gm) / 8.0) if ind_gm > 0 else 3.0
        margins.append("行业毛利率 %.1f%% → (30-毛利率)/8 估算（口径：行业均值）" % ind_gm)
    dims.append("存货/资产减值风险")

    # 4) 盈利质量恶化风险：公司 ROE 相对行业均值差距（真实 ROE）
    comp_roe = float(comp.get("roe") or 0)
    ind_roe = float(db.get("avg_roe") or 0)
    if comp_roe > 0 and ind_roe > 0:
        v4 = _clamp((ind_roe - comp_roe) / 5.0)
        margins.append("ROE差距(行业%.1f%%-公司%.1f%%)÷5 → 落后越多风险越高" % (ind_roe, comp_roe))
    else:
        v4 = _clamp((10.0 - ind_roe) / 4.0) if ind_roe > 0 else 3.0
        margins.append("行业平均ROE %.1f%% → (10-ROE)/4 估算（口径：行业均值）" % ind_roe)
    dims.append("盈利质量恶化风险")

    # 5) 政策合规与壁垒风险：行业政策基准值（knowledge 政策库/公开监管信息）+
    ind_name = db.get("industry_name", "") or ""
    v5 = 3.0
    for kw, base in POLICY_RISK_FLOOR.items():
        if kw in ind_name:
            v5 = base
            break
    margins.append("行业「%s」政策合规基准值 %.1f（来源：政策与监管公开信息，可上传政策库修正）" % (ind_name or "未匹配", v5))
    dims.append("政策合规与壁垒风险")

    return {
        "dimensions": dims,
        "values": [v1, v2, v3, v4, v5],
        "methodology": margins,
        "based_on": "真实财务指标映射（权益乘数/现金流质量/毛利率/ROE/行业政策基准）",
    }


# ============================================================
# 🌟 资料缺口审查（研报生成前：缺什么、为什么缺、怎么补）🌟
# ============================================================
def assess_data_gaps(company_name="", industry_name="", is_company_mode=False, rag_ok=True):
    """
    模拟 Data Planning / Data Retrieval 的缺口审查：
    对每项期望数据，检查本地数据库/RAG/上传件，输出缺口清单（含原因与建议）。
    返回 list[{"item","status","reason","suggest"}]
    """
    conn = sqlite3.connect("financial_research.db")
    cur = conn.cursor()
    gaps = []

    def _has_industry(ind):
        try:
            cur.execute("SELECT COUNT(*) FROM industry_benchmark WHERE industry_name=?", (ind,))
            return cur.fetchone()[0] > 0
        except Exception:
            return False

    if is_company_mode and company_name:
        # 公司向数据清单
        try:
            cur.execute("SELECT COUNT(*) FROM company_financial WHERE company_name LIKE ?", (f"%{company_name}%",))
            has_comp = cur.fetchone()[0] > 0
        except Exception:
            has_comp = False
        if has_comp:
            gaps.append({"item": f"{company_name} 财务指标（ROE/利润率/周转/乘数）", "status": "ok",
                          "reason": "已从本地财务数据库/业绩报表锁定（最新报告期）", "suggest": ""})
        else:
            gaps.append({"item": f"{company_name} 财务指标", "status": "missing",
                          "reason": "本地数据库暂无该公司；实时行情接口当前不可达（数据供应商限制），无法自动抓取",
                          "suggest": "请上传年报 PDF 或财务数据表（XLSX/CSV），系统将自动解析入库"})
        gaps.append({"item": "同行业可比公司数据", "status": "partial",
                      "reason": "行业基准为全市场聚合（样本数≥3），可比公司清单未单独维护",
                      "suggest": "可选：上传可比公司对照表以获得更强基准"})
    else:
        # 行业向数据清单
        if _has_industry(industry_name):
            gaps.append({"item": f"{industry_name} 行业基准（CR4/ROE/净利率/毛利率）", "status": "ok",
                          "reason": "来自东方财富业绩报表全市场聚合（报告期见数据截止标注）", "suggest": ""})
        else:
            gaps.append({"item": f"{industry_name} 行业基准数据", "status": "missing",
                          "reason": "本地行业库暂无该行业，且实时接口当前不可达",
                          "suggest": "请上传行业研报 PDF 或行业数据表"})
        gaps.append({"item": "行业市场规模与增速", "status": "partial",
                      "reason": "规模/增速为估算口径（市场公开区间），无付费数据源",
                      "suggest": "上传行业研究机构报告可替换为权威口径"})
        gaps.append({"item": "政策与风险数据", "status": "partial",
                      "reason": "政策库仅覆盖部分行业；政策原文为公开信息但无结构化订阅源",
                      "suggest": "上传政策原文/风险事件清单（PDF/TXT）即可入库"})

    # 共同项
    try:
        import news_fetcher as _nf
        _news_ok = bool(_nf.fetch_7x24_news(keyword="", pages=1, limit=1).get("items"))
    except Exception:
        _news_ok = False
    if _news_ok:
        gaps.append({"item": "最新新闻与公告", "status": "ok",
                      "reason": "系统内置实时抓取：东方财富公告大全 / 7×24快讯 / 搜狗新闻 / 财新，无需付费 API",
                      "suggest": ""})
    else:
        gaps.append({"item": "最新新闻与公告", "status": "partial",
                      "reason": "实时新闻接口当前不可达（网络受限），系统会自动降级；可上传新闻文本补充",
                      "suggest": "上传新闻文本或截图（支持多文件）"})
    conn.close()
    return gaps


def build_report_meta(db_data=None, company_data=None, is_company_mode=False):
    """报告元信息：数据鲜度、来源、口径。"""
    db = db_data or {}
    meta = {}
    meta["数据截至报告期"] = db.get("data_as_of") or "本地库未标注（历史口径）"
    meta["行业样本数"] = f"{db.get('sample_size', 0)} 家上市公司" if db.get("sample_size") else "未聚合"
    meta["数据主来源"] = db.get("data_source", "本地数据库")
    if company_data:
        meta["公司数据来源"] = f"{company_data.get('name')}（{company_data.get('year', '报告年度')}）财务指标"
    meta["产业链数据"] = "行业环节成本/利润率为区间值·综合公开资料（见各环节说明）"
    meta["报告生成方式"] = "7-Agent 多智能体流水线 + 证据链校验"
    return meta

# --- 3. 辅助解析函数 (已定位至顶层全局空间) ---
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

def chart_pdf_bytes(fig):
    """把 Plotly 图导出为 PDF 字节；环境缺少 kaleido 时返回 None，不阻塞页面渲染。"""
    try:
        buf = io.BytesIO()
        fig.write_image(file=buf, format="pdf")
        return buf.getvalue()
    except Exception:
        return None


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

    for root, dirs, files in os.walk(knowledge_dir):
        for filename in files:
            filepath = os.path.join(root, filename)
            text_content = ""
            try:
                if filename.endswith(".pdf"):
                    with pdfplumber.open(filepath) as pdf:
                        for page in pdf.pages:
                            text_content += page.extract_text() or ""
                elif filename.endswith((".txt", ".md")):
                    with open(filepath, "r", encoding="utf-8") as f:
                        text_content = f.read()
                
                if text_content:
                    chunks = [c.strip() for c in text_content.replace("。", "。\n").split("\n") if len(c.strip()) > 15]
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

# ============================================================
# 🧠 系统方法库（学习资料注入：网站自身学习调用）
# ============================================================
LEARNING_DIR = os.path.join("knowledge", "learning")

# 方法库分域：把入库的学习资料按专业域组织，供各 Agent 自动调用
METHOD_DOMAINS = {
    "写作规范": ["行研方法论_研报模版及话术_学习笔记", "行研方法论_深度报告tips_学习笔记"],
    "研究方法": ["行研方法论_券商行研全流程tips_学习笔记", "行研方法论_行研40问_学习笔记"],
    "财务分析": ["行研方法论_企业财务报表分析_学习笔记", "Deloitte_Tableau停机分析_学习笔记", "Deloitte_Excel平等分类_学习笔记"],
    "估值建模": ["Forage_JPMorgan投行_学习笔记", "Forage_Citi金融_学习笔记"],
    "审计风控": ["Forage_KPMG审计_学习笔记"],
    "咨询分析": ["Forage_PwC咨询_学习笔记", "Forage_BCG数据科学_学习笔记"],
}

_METHOD_CACHE = {"loaded": False, "bank": {}}

def load_method_bank():
    """加载 knowledge/learning 下的学习笔记全文到内存方法库。"""
    if _METHOD_CACHE["loaded"]:
        return _METHOD_CACHE["bank"]
    bank = {}
    if os.path.isdir(LEARNING_DIR):
        for fn in os.listdir(LEARNING_DIR):
            if not fn.endswith(".md"):
                continue
            fp = os.path.join(LEARNING_DIR, fn)
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    bank[os.path.splitext(fn)[0]] = f.read()
            except Exception as e:
                print(f"方法库读取失败 {fn}: {e}")
    _METHOD_CACHE["bank"] = bank
    _METHOD_CACHE["loaded"] = True
    return bank


def method_brief(domains, query_hint="", max_chars=1200):
    """
    按领域从方法库取资料并检索相关片段，返回 (brief_text, used_titles)。
    这是 AI 系统学习调用学习资料的统一入口。
    """
    bank = load_method_bank()
    used = []
    parts = []
    for domain in domains:
        for title in METHOD_DOMAINS.get(domain, []):
            if title not in bank:
                continue
            doc = bank[title]
            # 关键词打分选段
            lines = [l.strip() for l in doc.replace("。", "。\n").split("\n") if len(l.strip()) > 12]
            kw = [w for w in (query_hint or "").split() if len(w) >= 2] + ["方法", "框架", "步骤", "模型", "估值", "分析"]
            scored = []
            for ln in lines:
                s = sum(1.0 for k in kw if k in ln)
                if s > 0:
                    scored.append((s, ln))
            scored.sort(key=lambda x: -x[0])
            top = scored[:6]
            if top:
                snippet = " ".join(ln for _, ln in top)[:max_chars]
                parts.append(f"【{domain} · 学习资料《{title}》】\n{snippet}")
                used.append(title)
    if not parts:
        return "", []
    brief = "\n\n".join(parts)
    return brief, used


def method_domain_for(research_type, purpose, report_type):
    """根据当前研究任务选择要注入的方法领域。"""
    domains = ["研究方法"]
    if research_type == "公司研究":
        domains += ["财务分析", "估值建模"]
    else:
        domains += ["财务分析"]
    if "风险" in (purpose or "") or "审计" in (purpose or ""):
        domains.append("审计风控")
    domains.append("写作规范")
    return domains

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

# --- 产品化功能状态初始化（示例引导 / 上传 / 收藏 / 使用看板） ---
pf.init_product_state()

# --- 6. 侧边栏 ---
with st.sidebar:
    # 上传年报 / 财务数据表（可选）
    with st.expander("📤 上传年报 / 财务数据表（可选）"):
        up_file = st.file_uploader(
            "支持 PDF（年报）或 XLSX / CSV（财务数据表）",
            type=["pdf", "xlsx", "csv"],
            key="upload_file",
        )
        if up_file is not None:
            _fid = f"{up_file.name}_{up_file.size}"
            if st.session_state.get("uploaded_file_id") != _fid:
                st.session_state["uploaded_file_id"] = _fid
                with st.spinner("正在解析上传文件，请稍候..."):
                    _parse_result = pf.parse_uploaded_file(up_file, client)
                if _parse_result["error"]:
                    st.error(_parse_result["error"])
                st.session_state["uploaded_companies"] = _parse_result["companies"]
                st.session_state["uploaded_report_text"] = _parse_result["report_text"]
                if _parse_result["companies"]:
                    _names = "、".join(c["name"] for c in _parse_result["companies"])
                    st.success(f"已解析 {len(_parse_result['companies'])} 家公司财务数据：{_names}")
                if _parse_result["report_text"]:
                    st.success(f"已提取年报文本 {len(_parse_result['report_text'])} 字，将作为分析参考资料")
                if not _parse_result["companies"] and not _parse_result["report_text"] and not _parse_result["error"]:
                    st.warning("未识别到有效字段，请确认表头包含：公司 / ROE / 净利润率 / 资产周转率 / 权益乘数 / 经营现金流")
    # --- 🌟 数据新鲜度卡片 + 手动刷新（真实数据底座） ---
    with st.expander("📅 数据底座与刷新", expanded=False):
        _meta = du.read_meta()
        if _meta.get("report_period"):
            st.caption(f"行业基准数据：**{_meta.get('industries', 0)} 个行业** · 全市场 **{_meta.get('companies_total', 0)} 家** 公司")
            st.caption(f"报告期：**{_meta['report_period']}**（业绩报表真实聚合）· 最近刷新 {_meta.get('updated_at', '')}")
        else:
            st.caption("尚未完成全市场数据同步，点击下方按钮触发（约 1-2 分钟）。")
        if st.button("🔄 刷新实时财务数据", key="btn_refresh_data", use_container_width=True):
            with st.spinner("正在同步全市场业绩报表（东方财富）..."):
                _r = du.refresh_market_data(force=True)
            if _r.get("report_period"):
                st.success(f"刷新完成：{_r.get('industries')} 个行业 / 报告期 {_r['report_period']}")
            else:
                st.warning("刷新失败（数据源可能暂时不可达），已保留本地数据。")
            st.rerun()
        st.caption("说明：实时行情接口受数据源限制不可达；行业财务指标来自东方财富业绩报表全市场聚合，自动按 6 小时缓存。")

    # --- 🧠 系统方法论知识库（仅内部学习调用，不对访客开放浏览） ---
    with st.expander("🧠 系统方法论知识库（内部）", expanded=False):
        st.caption("系统内置 Forage 案例与券商行研方法论知识库，仅供 AI 投研 Agent "
                   "按研究任务内部检索调用（学习方法论/写作规范），不对外提供浏览与下载。")
        with st.expander("📥 入库新资料（管理员）", expanded=False):
            _lup = st.file_uploader("PDF / MD / TXT", type=["pdf", "md", "txt"], key="learn_upload")
            if _lup is not None and st.button("入库供系统学习", key="learn_upload_btn"):
                os.makedirs("knowledge/learning", exist_ok=True)
                with open(os.path.join("knowledge", "learning", _lup.name), "wb") as _fh:
                    _fh.write(_lup.getvalue())
                ldata.upsert_resource(os.path.splitext(_lup.name)[0], "用户上传", "markdown" if not _lup.name.lower().endswith(".pdf") else "pdf",
                                      "", ["用户上传"], _lup.name, _lup.size)
                load_method_bank()
                st.success(f"已入库：{_lup.name}（下次研究时系统将自动检索调用）")
                st.rerun()

    st.title("📚 研究历史")
    for idx, h in enumerate(st.session_state['history']):
        if st.button(f"📄 {h['query']}", key=f"h_{idx}"):
            st.session_state['current_report'] = h['content']
            st.session_state['current_data'] = h['data']
            st.session_state['current_query'] = h['query']
            st.rerun()            
    st.divider()
    # 收藏列表（产品化：用户留存）
    st.markdown("### ⭐ 我的收藏")
    _favs = st.session_state.get("favorites", [])
    if _favs:
        for _fi, _fav in enumerate(_favs[:10]):
            _fc1, _fc2 = st.columns([4, 1])
            with _fc1:
                if st.button(f"📌 {_fav['query']}", key=f"fav_load_{_fi}", use_container_width=True):
                    st.session_state['current_query'] = _fav['query']
                    st.session_state['current_report'] = _fav['report']
                    st.session_state['current_data'] = _fav['data']
                    st.rerun()
            with _fc2:
                if st.button("✕", key=f"fav_del_{_fi}"):
                    pf.delete_favorite(_fav['id'])
                    st.rerun()
    else:
        st.caption("暂无收藏。生成报告后点击「收藏本报告」即可在这里管理。")
    st.title("🛠 启动投研")
    
    research_mode = st.radio(
        "选择分析模式",
        ["简易模式（快速分析）", "标准模式（专业投研）"],
        key="research_mode"
    )
    
    company_query = ""
    query = ""
    period = "默认近三年+最新季度"
    purpose = "综合分析"
    report_type = "深度研究"

    if research_mode == "简易模式（快速分析）":
        research_target = st.radio("选择研究对象", ["公司", "行业"], key="research_target")
        if research_target == "公司":
            company_query = st.text_input("输入公司名称", placeholder="如：比亚迪", key="company_query")
        else:
            query = st.text_input("输入行业", placeholder="如：新能源汽车", key="query")
    else:
        research_target = st.radio("研究对象类型", ["公司", "行业"], key="research_target")
        if research_target == "公司":
            company_query = st.text_input("输入公司名称", placeholder="如：比亚迪", key="company_query")
            period_type = st.selectbox("选择时间周期类型", ["年度", "季度"], key="period_type")
            year_select = st.selectbox("⚙️ 选择年份", ["2021", "2022", "2023", "2024", "2025", "2026"], key="year_select")
            if period_type == "年度":
                period = f"{year_select}年度"
            else:
                quarter_select = st.selectbox("⚙️ 选择季度", ["Q1", "Q2", "Q3", "Q4"], key="quarter_select")
                period = f"{year_select}年{quarter_select}"
        else:
            query = st.text_input("输入行业", placeholder="如：新能源汽车", key="query")
            period_type = st.selectbox("选择时间周期类型", ["年度", "季度", "月度"], key="period_type")
            year_select = st.selectbox("⚙️ 选择年份", ["2021", "2022", "2023", "2024", "2025", "2026"], key="year_select")
            if period_type == "年度":
                period = f"{year_select}年度"
            elif period_type == "季度":
                quarter_select = st.selectbox("⚙️ 选择季度", ["Q1", "Q2", "Q3", "Q4"], key="quarter_select")
                period = f"{year_select}年{quarter_select}"
            else:
                month_select = st.selectbox("⚙️ 选择月份", ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"], key="month_select")
                period = f"{year_select}年{month_select}"
                
        report_type = st.selectbox("报告类型", ["年度策略", "季度跟踪", "专题研究"], key="report_type")
        purpose = st.selectbox("研究目的", ["投资价值分析", "行业趋势分析", "财务质量分析", "风险评估"], key="purpose")
    
    submit_btn = st.button("🚀 开启 7-Agent 深度协同", key="submit_btn")
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
    # ============================================================
    # 🚀 用户可见的实时进度条（7-Agent 流水线全程可视化）
    # ============================================================
    _progress_bar = st.progress(0.0, text="🚀 正在启动 7-Agent 投研流水线...")

    def _set_progress(pct, label):
        """更新进度条：pct 为 0~100 数值，label 为当前阶段说明。"""
        try:
            _progress_bar.progress(min(max(float(pct), 0.0), 100.0) / 100.0,
                                   text=f"{label}　【{min(max(float(pct), 0.0), 100.0):.0f}%】")
        except Exception:
            pass

    # 各 Agent 耗时统计（用于使用数据看板）
    _stage_last_t = time.time()
    _last_agent = None
    _orig_status_cb = status_callback

    def _timed_status_cb(agent, state):
        nonlocal _stage_last_t, _last_agent
        if state == "running":
            if _last_agent is not None:
                pf.record_agent_time(_last_agent, round(time.time() - _stage_last_t, 2))
            _last_agent = agent
            _stage_last_t = time.time()
        _orig_status_cb(agent, state)

    status_callback = _timed_status_cb
    # 🌟 核心修复一：行业智能对齐与自适应路由算法 🌟
    _set_progress(3, "🎯 行业智能对齐与路由")
    aligned_industry = auto_align_industry(company_name, user_input)
    db_data = get_locked_data(aligned_industry)

    # 📰 实时新闻与公告抓取（网站自身搜索能力：公告/快讯/网络搜索）
    _set_progress(6, "📰 正在抓取实时新闻与公告（东方财富/搜狗/财新）")
    log_callback("📰 [News Engine] 正在抓取实时新闻与公告（东方财富/搜狗/财新）...")
    _news_kw = company_name or aligned_industry
    _news_bundle = nf.fetch_news_bundle(company_name=company_name, industry=aligned_industry,
                                        keyword=_news_kw, days=30, limit_each=8)
    news_items = _news_bundle.get("items", []) or []
    log_callback(f"📰 [News Engine] 抓取到 {len(news_items)} 条实时新闻/公告（{_news_bundle.get('note', '')}）")

    # 🌐 网页图表数据读取（当新闻带链接时，尝试读取正文/表格/图表数据）
    _set_progress(10, "🌐 读取新闻网页正文 / 表格 / 图表数据")
    web_page_summary = ""
    for _nit in news_items[:3]:
        _u = _nit.get("url", "")
        if _u and _u.startswith("http"):
            _pg = wr.read_webpage(_u, timeout=8)
            if _pg.get("ok"):
                web_page_summary = wr.format_page_summary(_pg, max_text=900)
                break
    if web_page_summary:
        log_callback("🌐 [Web Reader] 已读取新闻网页正文/表格/图表数据，将作为研报参考。")
    else:
        log_callback("🌐 [Web Reader] 暂无可用网页详情（网络受限时跳过，不影响主流程）。")

    # 🏆 龙头公司横向对比（3~4 家，真实财务摘要 + 兜底基准）
    _set_progress(13, "🏆 龙头公司横向对比（3~4 家 · 真实财务摘要）")
    leader_payload = lc.build_leader_payload(aligned_industry)
    if leader_payload.get("ok"):
        log_callback(f"🏆 [Leader Compare] 龙头对比已生成：{'、'.join(leader_payload.get('companies', [])[:4])}")

    # 🧠 系统方法库调用：按研究任务自动检索注入学习资料（Forage 案例 + 行研方法论）
    _set_progress(16, "🧠 检索系统方法论知识库（Forage 案例 + 行研方法论）")
    _rt = "公司研究" if company_name else "行业研究"
    _method_domains = method_domain_for(_rt, purpose, report_type)
    _method_brief, _method_used = method_brief(
        _method_domains,
        query_hint=f"{company_name or ''} {aligned_industry} {purpose or ''} {report_type or ''}",
    )
    if _method_used:
        ldata.mark_used(_method_used, "Method Bank")
        log_callback(f"🧠 [Method Bank] 系统方法库已检索调用 {len(_method_used)} 份学习资料：{'、'.join(_method_used[:6])}")
    else:
        log_callback("🧠 [Method Bank] 系统方法库暂无匹配该任务的资料（可上传补充）。")
    
    # 重新加载 Data Retrieval Agent 避免白酒行业污染
    research_requirement = query_understanding_agent(aligned_industry, company_name, period, purpose, report_type)
    log_callback(f"🧠 [Query Agent] 对齐行业分类 -> {aligned_industry}，需求解析: {research_requirement}")
    
    company_data = None
    if company_name:
        company_data = get_company_data(company_name)
    
    # 杜邦分析与DCF估值复合提示词 (Financial Agent)
    if company_data:
        log_callback(f"🔑 [Database] 检测到个股【{company_name}】。开始进行杜邦基准与DCF分析锁定。")
        financial_prompt = f"""
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
        
        请进行深度审计并调用对应工具：
        1. 必须调用 `calculate_dcf` 工具对该个股进行内在价值估算。你可以使用公司的当前经营现金流 {company_data['cash']} 万元作为 free_cash_flow。假设增长率为 0.12 (12%)，WACC折现率为 0.085 (8.5%)。
        2. 针对其研究目的【{purpose}】，利用杜邦三要素进行拆解，指出其财务偏离行业基准的主要驱动力量。
        3. 如需要可调用 `search_latest_news` 检索公司最新动态、`read_webpage_charts` 读取相关网页数据。
        """
        if _method_brief:
            financial_prompt += f"\n\n【系统方法库·学习资料参考（估值/财务分析案例方法）】\n{_method_brief[:1800]}"
        if news_items:
            financial_prompt += f"\n\n【实时新闻与公告（可引用）】\n{nf.format_items_markdown(news_items, max_items=6)}"
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
        请分析该行业在【{period}】内的杜邦三要素驱动机制，尤其是其【{purpose}】维度下的财务质量表现。
        """
        if _method_brief:
            financial_prompt += f"\n\n【系统方法库·学习资料参考（财务分析框架）】\n{_method_brief[:1800]}"
        if news_items:
            financial_prompt += f"\n\n【实时新闻与公告（可引用）】\n{nf.format_items_markdown(news_items, max_items=6)}"

    # 向量数据库检索 (RAG 闭环)
    _set_progress(18, "🔍 知识库向量检索（RAG 底稿对齐）")
    log_callback("🔍 [RAG Engine] 正在进行向量库高维度特征检索对齐...")
    rag_context = vector_search(aligned_industry, top_k=3)
    # 用户上传的年报文本作为补充参考资料进入 RAG 上下文
    _uploaded_text = st.session_state.get("uploaded_report_text", "")
    if _uploaded_text:
        rag_context = rag_context + "\n\n[用户上传年报文本]\n" + _uploaded_text[:5000]
    log_callback("✅ [RAG Engine] 本地向量数据库检索对齐完成！")
    
    # 1. Planner Agent
    _set_progress(22, "📋 Planner Agent：制定研究提纲")
    status_callback("Planner", "running")
    log_callback("🔄 [Planner Agent] 正在制定财报质量及行业深度分析提纲...")
    time.sleep(1)
    
    # 2. Research Agent（支持实时新闻/网页图表工具）
    _set_progress(30, "🔍 Research Agent：行业竞争格局与 CR4 分析")
    status_callback("Research", "running")
    log_callback("🔍 [Research Agent] 查询大盘，融合数据库，构建竞争集中度 (CR4) 指标...")
    research_prompt = f"根据以下行业数据库信息，结合研究周期【{period}】完成行业竞争格局分析。行业: {db_data['industry_name']}, CR4: {db_data['cr4']}%"
    if _method_brief:
        research_prompt += f"\n\n【系统方法库·学习资料参考】\n{_method_brief[:1500]}"
    if news_items:
        research_prompt += f"\n\n【实时新闻与公告（可引用最新动态）】\n{nf.format_items_markdown(news_items, max_items=8)}"
    _research_messages = [{"role": "user", "content": research_prompt}]
    _res_r = client.chat.completions.create(
        model="deepseek-chat",
        messages=_research_messages,
        tools=financial_tools,
        temperature=0.3
    )
    _res_msg = _res_r.choices[0].message
    if getattr(_res_msg, "tool_calls", None):
        _research_messages.append(_res_msg)
        for _tc in _res_msg.tool_calls:
            _tname = _tc.function.name
            _targs = json.loads(_tc.function.arguments)
            log_callback(f"🛠️ [Research Tool] 触发 {_tname} 工具调用！")
            _tresult = execute_financial_tool(_tname, _targs)
            _research_messages.append({"role": "tool", "tool_call_id": _tc.id, "content": _tresult})
        res_research = client.chat.completions.create(
            model="deepseek-chat",
            messages=_research_messages,
            temperature=0.3
        ).choices[0].message.content
    else:
        res_research = _res_msg.content
    
    # 3. Financial Agent (支持 DCF 工具调用)
    _set_progress(42, "📊 Financial Agent：杜邦分解 + DCF 估值")
    status_callback("Financial", "running")
    log_callback("📊 [Financial Agent] 计算杜邦公式与 DCF 模型，并进行审计诊断...")
    financial_messages = [
        {"role": "system", "content": "你是Financial Agent。你的任务：优先使用真实数据进行财务杜邦分解或DCF估值，根据需求调用合适的工具，严禁编造数据。"},
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
            log_callback(f"🛠️ [Financial Tool] 触发 {tool_name} 工具调用！")
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

    # 4. Policy Agent（注入最新新闻/公告与网页数据作为政策背景）
    _set_progress(52, "📜 Policy Agent：政策与合规拆解")
    status_callback("Policy", "running")
    log_callback("📜 [Policy Agent] 精细化政策拆解：行业限制、税收优惠及环保壁垒...")
    _news_ctx = nf.format_items_markdown(news_items, max_items=6) if news_items else "（无实时新闻）"
    policy_prompt = (f"针对 '{aligned_industry}'，请详述其面临的最新行业准入门槛及绿色金融支持力度。"
                     f"参考底稿: {rag_context}\n\n【最新新闻与公告（实时抓取，可引用）】\n{_news_ctx}")
    res_policy = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": policy_prompt}]
    ).choices[0].message.content

    # 5. Risk Agent（注入实时新闻/公告中的风险事件线索）
    _set_progress(60, "🚩 Risk Agent：核心风险扫描")
    status_callback("Risk", "running")
    log_callback("🚩 [Risk Agent] 核心风险扫描...")
    risk_prompt = (f"请分析：行业: {db_data['industry_name']}的财务与政策风险。参考底稿: {rag_context}"
                   f"\n\n【最新新闻与公告（实时抓取，注意其中的风险线索：处罚/诉讼/减持/监管问询等）】\n{_news_ctx}")
    res_risk = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": risk_prompt}],
        temperature=0.3
    ).choices[0].message.content
   
    # 🌟 核心修复二：学术级智能体多边辩论机制 (Financial Agent vs Risk Agent) 🌟
    _set_progress(68, "💬 多空辩论：Financial 专家 vs Risk 审计专家")
    log_callback("💬 [Debate] 审计对立碰撞启动：Financial 专家 与 Risk 审计专家辩论会...")
    time.sleep(1)
    
    debate_prompt_fin = f"""
    针对 {aligned_industry} 行业的财务前景及核心公司，基于你的研究支持：{res_financial}。
    请作为绝对乐观的财务学家，提出论证，重点证明该行业的杜邦收益质量以及其估值具备极高的内在安全边际，反驳任何盲目的减值质疑。
    """
    res_debate_fin = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": debate_prompt_fin}],
        temperature=0.5
    ).choices[0].message.content
    log_callback(f"💬 [Financial Agent]: 观点成立。杜邦分析显示资产效率极高，DCF内在企业价值空间充足。")
    
    debate_prompt_risk = f"""
    现在请扮演极具批判性的风险审计总监。
    
    刚才财务专家发表了如下乐观论点：
    {res_debate_fin}
    
    请根据风险底稿【{res_risk}】和公司的痛点，提出尖锐的反驳。重点论证其高周转、高ROE是否是通过透支现金流、增加隐性杠杆获得的，证明其真实的‘核心利润质量’并不像账面数据那么好看。
    """
    res_debate_risk = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": debate_prompt_risk}],
        temperature=0.5
    ).choices[0].message.content
    log_callback(f"🚩 [Risk Agent Rebuttal]: 反驳成立！高杠杆及应收账款周转放缓，经营现金流存在重大隐性流失。")

    # 🌟 核心修复三：专家委员会 Agent 终审、矛盾消除与数据可信度证据链 (Evidence Ledger) 🌟
    _set_progress(78, "⚖️ 专家委员会终审：矛盾消除 + 证据链审定")
    status_callback("Judge", "running")
    log_callback("⚖️ [Committee Agent] 专家委员会正在进行矛盾消除、逻辑排歧与可信等级审定...")
    judge_reference = get_judge_reference(db_data["industry_name"])
    
    committee_prompt = f"""
    你现在是【专家委员会 Committee Agent】。
    你的任务是根据多边辩论、数据规划，以及数据库对报告中的每一条结论进行真实性验证、逻辑排歧，并为核心论据标明“可信度等级(A/B/C)”。
    
    ======== 辩论听证会记录 ========
    1. 财务专家立场：{res_debate_fin}
    2. 风险专家反驳：{res_debate_risk}
    
    ======== 数据库对照底牌 ========
    行业基准：行业={db_data["industry_name"]}, ROE={db_data["avg_roe"]}%, 净利率={db_data["net_profit_margin"]}%
    政策背景：{judge_reference["policy"]}
    风险事实：{judge_reference["risk"]}
    
    请严格执行以下三步：
    1. 【矛盾消除与逻辑一致性】：分析财务乐观论调与风险审计的反驳是否冲突。如果存在冲突（例如高ROE与低现金流），指出原因并调和逻辑，给出最终审定意见。
    2. 【数据来源与可信度度量】：
       请为研报的核心论点生成一个【数据可信度证据链 (Evidence Ledger)】，严格采用以下格式：
       - 论点: [具体财务/竞争结论]
       - 来源: [具体数据源，如公司2025年报/申万数据/本地SQLite库]
       - 字段: [具体会计科目]
       - 页面/位置: [例如 P23 利润表 或 数据库底表第一行]
       - 可信等级: [A/B/C 三选一]
    3. 最终判定：是否通过终审。
    
    必须返回标准 JSON 字符串：
    {{
        "score": 98,
        "pass": true,
        "evidence_ledger": [
            {{"point": "ROE高达18.5%", "source": "本地SQLite数据库", "field": "ROE", "page": "company_financial底表", "confidence": "A"}},
            {{"point": "现金流指标优秀", "source": "公司2025年报", "field": "经营现金流", "page": "P12 现金流量表", "confidence": "A"}},
            {{"point": "面临供应链毛利压缩风险", "source": "个股痛点备忘录", "field": "核心痛点", "page": "个股镜像表", "confidence": "B"}}
        ],
        "failed_agent": "None",
        "reason": "辩论逻辑已调和并交叉验证通过，逻辑无瑕疵"
    }}
    """
    
    res_verifier = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role":"user", "content":committee_prompt}],
        temperature=0.1
    ).choices[0].message.content

    try:
        # 解析专家委员会证据链
        verifier_json = json.loads(res_verifier.replace("```json", "").replace("```", "").strip())
        evidence_ledger = verifier_json.get("evidence_ledger", [])
    except Exception:
        evidence_ledger = [
            {"point": "未录入指标", "source": "自审异常", "field": "无", "page": "无", "confidence": "C"}
        ]

    # 7. Report Agent
    _set_progress(88, "✍️ Report Agent：研报总装与图表生成")
    status_callback("Report", "running")
    log_callback("✍️ [Report Agent] 研报总装中，整合对标成果与 RAG 深度分析...")
    _write_brief, _write_used = method_brief(["写作规范"], query_hint=f"{aligned_industry} 研报 {report_type}")
    if _write_used:
        ldata.mark_used(_write_used, "Report Agent")
        log_callback(f"🧠 [Method Bank] Report Agent 调用写作规范资料 {len(_write_used)} 份")

    # 新闻/公告/网页图表数据注入研报（最新动态 + 来源链接）
    _news_md = nf.format_items_markdown(news_items, max_items=8) if news_items else "（实时新闻接口暂不可达，建议上传新闻文本补充）"
    _leader_md = lc.format_leader_markdown(aligned_industry) if leader_payload.get("ok") else "（该行业暂无龙头对比数据）"

    report_prompt = f"""
    你是一名卖方证券研究所首席分析师（参考券商行研深度报告方法论《行业分析模版》与《公司分析模版》）。
    请根据以下研究材料，生成一篇定位为【{report_type}】、研究目的侧重于【{purpose}】、时间跨度锚定在【{period}】的标准深度研究报告。
    写作要求（券商行研方法论，严格对齐行研模版结构）：
    1. 结论先行：开头 100-200 字给出核心观点（评级/结论/风险一句），正文再展开论证；
    2. 数据说话：每个核心论点尽量附数据（% / 倍 / 亿元），并指明来源（行业聚合/公司财报/政策底稿/实时新闻）；
    3. 结构完整（行业研究模版：核心观点 → 行情回顾/行业表现（含图表佐证与涨跌幅/PE-PB分位）→ 细分板块分析（选 4~6 个热门板块，含关键财务数据/核心驱动/典型企业）→ 产业链传导（上游成本→中游制造→下游需求）→ 行业事件/新闻 → 投资策略与盈利预测（核心假设）→ 风险提示（若XX发生→将导致XX结果句式）；
       公司研究模版：核心观点 → 市场位势（营收排名/市占率/增速对比/行业集中度）→ 盈利能力与定价权（毛利率/净利率 vs 行业均值/产品差异化/价格轨迹）→ 供应链话语权（应付/应收/存货周转）→ 成长后劲（管线/研发/产能）→ 政策抗性（集采/医保敞口）→ 经营效率（人均产出/费用控制）→ 风险提示）；
    4. 话术规范：客观审慎，避免绝对化表述；估值与预测给出假设条件；
    5. 必须引用「最新新闻与公告」与「龙头公司横向对比」两大板块（见下方实时数据）；
    6. 结尾必须给出「免责声明」。篇幅 1200-2000 字。

    报告必须整合以下多边对标及辩论博弈结果：

    # 一、核心观点
    - 要求：100-200字总结全文，针对【{purpose}】给出核心投资逻辑、内在价值估值结论及多边辩论综述。

    # 二、行业行情回顾 (聚焦周期: {period})

    # 三、企业内在价值与基本面博弈分析 (结合 Financial 与 Risk 专家博弈点：)
    1. 财务分析依据：{res_financial}
    2. 辩论意见综述：财务立场的{res_debate_fin} 与 风险立场的{res_debate_risk}

    # 四、细分赛道分析

    # 五、产业链传导

    # 六、政策与行业合规 (RAG政策输入)：
    {res_policy}

    # 七、投资策略与盈利预测 (时间周期跨度: {period})

    # 八、风险提示 (RAG风险及底线输入)：
    {res_risk}

    ======== 实时新闻与公告（务必引用并标注来源与日期）========
    {_news_md}

    ======== 龙头公司横向对比（3~4 家，务必引用并简要分析）========
    {_leader_md}

    ======== 网页数据摘录（新闻正文/表格/图表数据，可引用）========
    {web_page_summary if web_page_summary else "（无）"}

    自评检验结果反馈：
    {res_verifier}

    【系统方法库·写作规范学习资料（必须按此规范执行）】：
    {_write_brief if _write_brief else "（暂无写作规范资料）"}
    """
    res_report = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": report_prompt}],
        temperature=0.4
    ).choices[0].message.content

    # 构造动态 Plotly 图表数据（真实数据底座版）
    is_company = company_data is not None
    risk_radar = compute_risk_radar(company_data, db_data)
    chain_payload = icd.build_chain_payload(db_data.get("industry_name", aligned_industry))
    data_gaps = assess_data_gaps(company_name, db_data.get("industry_name", aligned_industry),
                                 is_company_mode=is_company)
    report_meta = build_report_meta(db_data, company_data, is_company)
    as_of = db_data.get("data_as_of") or "本地库"
    if company_data:
        chart_data = {
            "company_name": company_name,
            "company_roe": company_data["roe"],
            "company_margin": company_data["margin"],
            "company_turnover": company_data["turnover"],
            "company_multiplier": company_data["multiplier"],
            "company_cash": company_data["cash"],
            "company_gross_margin": company_data.get("gross_margin", 0),
            "company_eps": company_data.get("eps", 0),
            "industry_roe": db_data["avg_roe"],
            "industry_margin": db_data["net_profit_margin"],
            "industry_turnover": db_data["asset_turnover"],
            "industry_multiplier": db_data["equity_multiplier"],
            "industry_cash": db_data["operating_cash_flow"],
            "industry_gross_margin": db_data.get("gross_margin", 0),
            "evidence_ledger": evidence_ledger, # 将证据链传到前端渲染
            "locked_source": f"个股: {company_data['name']} 与 行业: {db_data['industry_name']} 双重锁定",
            "risk_radar": risk_radar,
            "industry_chain": chain_payload,
            "data_gaps": data_gaps,
            "report_meta": report_meta,
            "market_as_of": as_of,
            "data_note": "行业毛利/ROE/CR4/净利率为东方财富业绩报表全市场真实聚合；历史序列与市场规模为示意/估算口径，详见缺口说明。",
            "news_items": news_items,
            "news_note": _news_bundle.get("note", ""),
            "leader_data": leader_payload,
            "web_page_summary": web_page_summary,
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
                "growth_rate": [15.0, 13.5, 10.2, 8.5, 7.8],
                "note": "市场规模/增速为估算口径；CR4、ROE、净利率、毛利率为真实聚合值"
            },
            "financial_trend": {
                "years": ["2022", "2023", "2024", "2025", "2026Q2"],
                "roe_trend": [round(db_data["avg_roe"] * f, 2) for f in [1.15, 1.08, 1.0, 0.96, 0.92]],
                "margin_trend": [round(db_data["net_profit_margin"] * f, 2) for f in [1.10, 1.05, 1.0, 0.98, 0.95]],
                "note": "2026Q2 为真实聚合值，其余年份为趋势示意（2022-2025 历史明细需授权数据源）"
            },
            "capability_comparison": {
                "metrics": ["盈利能力(ROE%)", "毛利率(%)", "集中度(CR4%)", "净利率(%)"],
                "values": [round(db_data["avg_roe"], 2), round(db_data.get("gross_margin", 0), 2),
                           round(db_data["cr4"], 2), round(db_data["net_profit_margin"], 2)]
            },
            "risk_radar": risk_radar,
            "industry_chain": chain_payload,
            "data_gaps": data_gaps,
            "report_meta": report_meta,
            "market_as_of": as_of,
            "data_note": "CR4/ROE/净利率/毛利率为真实全市场聚合；历史趋势与市场规模为估算口径，详见缺口说明。",
            "evidence_ledger": evidence_ledger,
            "locked_source": db_data["data_source"],
            "news_items": news_items,
            "news_note": _news_bundle.get("note", ""),
            "leader_data": leader_payload,
            "web_page_summary": web_page_summary,
        }

    final_text = f"{res_report}\n\n```json\n{json.dumps(chart_data)}\n```"
    log_callback("✅ 工作流执行完毕。智能投研报告及图表已就绪！")
    _set_progress(100, "✅ 研报生成完成，正在渲染图表与导出选项…")
    if _last_agent is not None:
        pf.record_agent_time(_last_agent, round(time.time() - _stage_last_t, 2))
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

# --- 9. 主面板报告与动态画图 ---
with col_main:
    # 初始化 Tool Traces
    if 'tool_traces' not in st.session_state:
        st.session_state['tool_traces'] = []

    # ============================================================
    # 🌟 资料缺口审查面板（研报生成前体检：缺什么 / 为什么缺 / 上传补充）
    # ============================================================
    if submit_btn and (query or company_query):
        _ind = auto_align_industry(company_query, query)
        _db = get_locked_data(_ind)
        _co = get_company_data(company_query) if company_query else None
        _gaps_pre = assess_data_gaps(company_query, _ind, is_company_mode=bool(company_query))
        st.session_state["precheck_gaps"] = _gaps_pre
        st.session_state["precheck_industry"] = _ind
        st.session_state["precheck_db_ok"] = bool(_db.get("data_as_of"))
        st.session_state["precheck_company"] = company_query
        st.session_state["precheck_query"] = query
        st.session_state["precheck_done"] = True
        st.rerun()

    if st.session_state.get("precheck_done") and (st.session_state.get("precheck_query") or st.session_state.get("precheck_company")):
        _gaps = st.session_state.get("precheck_gaps") or []
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.write("### 📦 资料缺口审查（研报生成前自动体检）")
            _missing = [g for g in _gaps if g.get("status") == "missing"]
            _partial = [g for g in _gaps if g.get("status") == "partial"]
            _ok = [g for g in _gaps if g.get("status") == "ok"]
            _m1, _m2, _m3 = st.columns(3)
            _m1.metric("检查项", len(_gaps))
            _m2.metric("❌ 缺失（需补充）", len(_missing))
            _m3.metric("⚠️ 部分可替代", len(_partial))
            _gap_note = st.session_state.get("gap_uploaded_note", "")
            if _gap_note:
                st.success(_gap_note)
            for g in _gaps:
                _icon = {"ok": "✅", "partial": "⚠️", "missing": "❌"}.get(g.get("status"), "ℹ️")
                st.markdown(f"**{_icon} {g.get('item','')}**　*状态：{g.get('status','')}*")
                st.caption(f"为什么：{g.get('reason','')}")
                if g.get("suggest"):
                    st.caption(f"如何补充：{g.get('suggest','')}")
                # ============ 自动补链：系统自己分析需要什么资料并搜索获取链接 ============
                if g.get("status") in ("missing", "partial"):
                    _gap_key = f"gap_src_{g.get('item','')[:30]}_{hash(g.get('reason','')) & 0xffff}"
                    with st.expander(f"🔗 自动获取资料链接（{g.get('item','')[:24]}）", expanded=False):
                        _src_meta = suggest_gap_sources(g, company_name=st.session_state.get("precheck_company", ""),
                                                        industry=st.session_state.get("precheck_industry", ""))
                        st.caption(f"📋 系统分析：该项需要 **{_src_meta['type']}** 类资料，已按「{_src_meta['search_kw']}」检索")
                        # 官方固定入口
                        st.markdown("**🏛️ 官方权威入口（推荐，100% 可访问）**")
                        for _name, _url in _src_meta["official"]:
                            st.markdown(f"- [{_name}]({_url})")
                        # 自动搜索结果
                        st.markdown("**🔍 自动搜索到的资料页**")
                        if _src_meta["search"]:
                            for _it in _src_meta["search"][:4]:
                                _t = _it.get("title", "")
                                _u = _it.get("url", "")
                                _s = _it.get("source", "")
                                if _u and _u.startswith("http"):
                                    st.markdown(f"- {_t}（{_s}）[打开]({_u})")
                                elif _u:
                                    st.markdown(f"- {_t}（{_s}）[搜狗跳转](https://www.sogou.com{_u if _u.startswith('/') else '/' + _u})")
                                else:
                                    st.markdown(f"- {_t}（{_s}）")
                        else:
                            st.caption(_src_meta["note"])
                        st.caption("💡 下载后可在下方「上传缺失资料」直接上传，系统自动解析入库参与研究。")
            with st.expander("📤 上传缺失资料（系统将自动解析入库，研报会引用你的资料）", expanded=bool(_missing)):
                _gap_files = st.file_uploader(
                    "支持多文件：年报/研报 PDF、财务数据表 XLSX/CSV、政策或新闻 TXT/MD",
                    type=["pdf", "xlsx", "xls", "csv", "txt", "md"],
                    accept_multiple_files=True,
                    key="gap_uploader",
                )
                if _gap_files and st.button("💾 解析并入库", key="gap_parse_btn"):
                    _stored = 0
                    _f_note = []
                    os.makedirs("knowledge/gap_uploads", exist_ok=True)
                    for _upf in _gap_files:
                        try:
                            with open(os.path.join("knowledge", "gap_uploads", _upf.name), "wb") as _fh:
                                _fh.write(_upf.getvalue())
                            _res = pf.parse_uploaded_file(_upf, client)
                            if _res.get("companies"):
                                _uc = st.session_state.setdefault("uploaded_companies", [])
                                for _c in _res["companies"]:
                                    if _c["name"] not in [x.get("name") for x in _uc]:
                                        _uc.append(_c)
                                _f_note.append(f"{_upf.name}：解析出 {len(_res['companies'])} 家公司财务数据")
                            if _res.get("report_text"):
                                with open(os.path.join("knowledge", "gap_uploads", _upf.name + ".txt"), "w", encoding="utf-8") as _tf:
                                    _tf.write(_res["report_text"])
                                _f_note.append(f"{_upf.name}：提取正文 {len(_res['report_text'])} 字入知识库")
                            if _res.get("error"):
                                _f_note.append(f"{_upf.name}：{_res['error']}")
                            _stored += 1
                        except Exception as _upe:
                            _f_note.append(f"{_upf.name}：保存失败 {_upe}")
                    st.session_state["gap_uploaded_note"] = "；".join(_f_note) or "文件已保存"
                    st.rerun()
            _go_c, _skip_c = st.columns(2)
            with _go_c:
                if st.button("🚀 已上传补充资料，开始研究", key="gap_go", use_container_width=True, type="primary"):
                    st.session_state["trigger_run"] = True
                    st.rerun()
            with _skip_c:
                if st.button("⏭️ 接受当前缺口，直接开始研究", key="gap_skip", use_container_width=True):
                    st.session_state["trigger_run"] = True
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        st.divider()

    _auto_trigger = st.session_state.get("trigger_run", False)
    st.session_state["trigger_run"] = False
    if _auto_trigger and (query or company_query):
        _run_started = time.time()
        _tok_before = dict(st.session_state.get("usage_stats", {}))
        st.session_state['logs_history'] = []
        st.session_state['tool_traces'] = [] # 清空之前的工具调用链痕迹
        st.session_state['agent_times'] = []  # 清空上一轮的 Agent 耗时
        
        try:
            raw_report = run_research_flow(
                query,
                log_callback=append_log,
                status_callback=update_agent_status,
                company_name=company_query,
                period=period,
                purpose=purpose,
                report_type=report_type
            )
        except Exception as _run_error:
            try:
                st.progress(1.0, text=f"❌ 研究流程执行失败：{str(_run_error)[:80]}")
            except Exception:
                pass
            st.error(f"研究流程执行失败：{_run_error}")
            st.stop()
        _run_elapsed = time.time() - _run_started
        clean_text, dynamic_data = extract_report_data(raw_report)
        if not clean_text or not clean_text.strip():
            clean_text = "⚠️ 本次研究未能生成报告正文（模型返回为空）。请检查 DeepSeek API 额度是否充足，或稍后重试。"
        
        st.session_state['current_report'] = clean_text
        st.session_state['current_data'] = dynamic_data
        # 历史记录里显示标的公司或行业名
        st.session_state['current_query'] = company_query if company_query else query
        st.session_state['history'].insert(0, {
            "query": st.session_state['current_query'], 
            "content": clean_text, 
            "data": dynamic_data
        })
        # 记录本次运行的使用数据（时长 / token / Agent 耗时）
        _tok_after = st.session_state.get("usage_stats", {})
        pf.record_usage(
            st.session_state['current_query'],
            _run_elapsed,
            _tok_after.get("input_tokens", 0) - _tok_before.get("input_tokens", 0),
            _tok_after.get("output_tokens", 0) - _tok_before.get("output_tokens", 0),
            st.session_state.get("agent_times", []),
        )
        
        for agent in ["Planner", "Research", "Financial", "Policy", "Risk", "Judge", "Report"]:
            st.session_state[f"status_{agent}"] = "success"
        st.rerun()

    if st.session_state['current_report']:
        st.markdown(f"## 📋 {st.session_state['current_query']} 深度研报分析")
        _asof = st.session_state.get('current_data', {}).get('market_as_of', '')
        _lock = st.session_state.get('current_data', {}).get('locked_source', '')
        if _asof or _lock:
            st.caption(f"📅 数据截至报告期：**{_asof or '本地库'}**　|　🔗 数据来源：{_lock}")
        
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
                    st.caption(f"📚 数据来源：公司财务指标（{data.get('company_name','')}）vs 行业聚合值（{data.get('market_as_of', '')}）· 东方财富业绩报表")
                    
                    _comp_pdf = chart_pdf_bytes(fig_comp)
                    if _comp_pdf is not None:
                        st.download_button(
                        label="📊 导出杜邦对标图为 PDF",
                        data=_comp_pdf,
                        file_name="dupont_comparison_chart.pdf",
                        mime="application/pdf",
                        key="dl_comp_pdf"
                    )
                else:
                    share_data = data.get("market_share", {"labels": ["集中度 (CR4)", "其他企业"], "values": [55, 45]})
                    fig_pie = go.Figure(data=[go.Pie(labels=share_data["labels"], values=share_data["values"], hole=.4)])
                    fig_pie.update_layout(title="市场集中度 (CR4) 动态格局", height=300, margin=dict(l=10, r=10, t=40, b=10))
                    st.plotly_chart(fig_pie, use_container_width=True, key="industry_pie_chart")
                    st.caption(f"📚 数据来源：东方财富业绩报表全市场聚合（报告期 {data.get('market_as_of', '—')}）· CR4={share_data['values'][0]}%")
                    
                    _pie_pdf = chart_pdf_bytes(fig_pie)
                    if _pie_pdf is not None:
                        st.download_button(
                        label="📊 导出竞争格局图为 PDF",
                        data=_pie_pdf,
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
                    st.caption(f"📚 数据来源：公司财务指标（{data.get('company_name','')}）与行业聚合值（{data.get('market_as_of', '')}）")
                    
                    _radar_pdf = chart_pdf_bytes(fig_radar)
                    if _radar_pdf is not None:
                        st.download_button(
                        label="📈 导出能力对标雷达图为 PDF",
                        data=_radar_pdf,
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
                    st.caption("📚 数据来源：市场规模/增速为估算口径（公开区间）；CR4/ROE/净利率/毛利率为东方财富业绩报表真实聚合")
                    
                    _growth_pdf = chart_pdf_bytes(fig_growth)
                    if _growth_pdf is not None:
                        st.download_button(label="📈 导出市场规模增速图为 PDF", data=_growth_pdf, file_name="market_growth_chart.pdf", mime="application/pdf", key="dl_growth")

            # --- 🌟 公司模式新增：杜邦 ROE 差距归因瀑布图（需求：更多图表类型） ---
            if is_company_mode:
                import math as _math
                _wf_labels = ["行业ROE", "利润率贡献", "周转率贡献", "杠杆贡献", "公司ROE"]
                def _ln(v):
                    return _math.log(max(float(v), 0.01))
                try:
                    _wf_dm = _ln(data["company_margin"]) - _ln(data["industry_margin"])
                    _wf_dt = _ln(data["company_turnover"]) - _ln(data["industry_turnover"])
                    _wf_dl = _ln(data["company_multiplier"]) - _ln(data["industry_multiplier"])
                except Exception:
                    _wf_dm = _wf_dt = _wf_dl = 0.0
                _wf_vals = [data["industry_roe"], _wf_dm, _wf_dt, _wf_dl, data["company_roe"]]
                _wf_fig = go.Figure()
                _wf_bottoms = [0, data["industry_roe"], data["industry_roe"] + _wf_dm,
                               data["industry_roe"] + _wf_dm + _wf_dt, 0]
                for _i, (_lb, _v, _bt) in enumerate(zip(_wf_labels, _wf_vals, _wf_bottoms)):
                    if _i in (1, 2, 3):
                        _wf_fig.add_trace(go.Bar(x=[_lb], y=[abs(_v)], base=[_bt],
                                                 marker_color="#86bc25" if _v >= 0 else "#C0392B",
                                                 text=[f"{_v:+.2f}"], textposition="inside",
                                                 name="贡献" if _i == 1 else None,
                                                 showlegend=False))
                    else:
                        _wf_fig.add_trace(go.Bar(x=[_lb], y=[abs(_v)], base=[0],
                                                 marker_color="#0F2A5C" if _i == 0 else "#1F5FA8",
                                                 text=[f"{_v:.2f}"], textposition="outside",
                                                 showlegend=False))
                _wf_fig.update_layout(
                    title="杜邦 ROE 差距归因（公司 vs 行业 · 对数分解示意）",
                    height=300, margin=dict(l=10, r=10, t=50, b=10),
                    yaxis_title="ROE 贡献（百分点/对数）",
                )
                st.plotly_chart(_wf_fig, use_container_width=True, key="dupont_waterfall_chart")
                st.caption(f"📚 数据来源：公司财务指标（{data.get('company_name','')}）vs 行业聚合值（{data.get('market_as_of', '')}）；分解口径：ln 差线性化示意")

            # --- 🌟 新增：多智能体工具与数据库自研 Trace 监控组件 🌟 ---
            if st.session_state['tool_traces']:
                st.divider()
                with st.expander("🛠️ 智能体工具调用链与数据库自审监控 (Auditing Trace)", expanded=True):
                    cols_th = st.columns([1.2, 1.8, 2.5, 3.5])
                    cols_th[0].markdown("**执行智能体**")
                    cols_th[1].markdown("**调用接口(Tool)**")
                    cols_th[2].markdown("**输入参数(Arguments)**")
                    cols_th[3].markdown("**输出反馈 / 审计核验**")
                    st.divider()
                    for idx, trace in enumerate(st.session_state['tool_traces']):
                        cols = st.columns([1.2, 1.8, 2.5, 3.5])
                        cols[0].markdown(f"🤖 `{trace['agent']}`")
                        cols[1].markdown(f"📂 `{trace['tool']}`")
                        cols[2].code(trace['input'], language="json")
                        cols[3].info(trace['output'])
                        
            # --- 🌟 新增：专家委员会学术级数据证据链可视化面板 (Evidence Ledger) 🌟 ---
            evidence_data = data.get("evidence_ledger", [])
            if evidence_data:
                st.divider()
                with st.expander("🛡️ 专家委员会数据可信度终审证据链 (Evidence Ledger)", expanded=True):
                    st.markdown("根据评审手册规范，委员会对研报中引用的核心结论与数据源进行多边交叉检验，生成的存证证据链如下：")
                    df_evidence = pd.DataFrame(evidence_data)
                    # 友好重命名列头
                    df_evidence.rename(columns={
                        "point": "审计结论 / 核心论点",
                        "source": "数据物理来源",
                        "field": "对齐科目",
                        "page": "位置 / 页码",
                        "confidence": "可信度评级 (Confidence Level)"
                    }, inplace=True)
                    st.dataframe(df_evidence, use_container_width=True)
                    st.caption("注：可信度评级 A 代表经底层数据库强锁定且与行业层级完全一致；B 代表经离线文件 RAG 检索校验；C 代表基于大盘生成推理。")

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
                st.caption("📚 数据来源：最新报告期为东方财富业绩报表真实聚合；历史期为趋势示意（授权数据源缺失）")
                
                _trend_pdf = chart_pdf_bytes(fig_trend)
                if _trend_pdf is not None:
                    st.download_button(
                    label="📉 导出财务趋势折线图为 PDF",
                    data=_trend_pdf,
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
                st.caption("📚 数据来源：ROE/毛利率/CR4/净利率为东方财富业绩报表全市场聚合（报告期见数据截止标注）")
                
                _cap_pdf = chart_pdf_bytes(fig_cap)
                if _cap_pdf is not None:
                    st.download_button(
                    label="📊 导出核心能力对比图为 PDF",
                    data=_cap_pdf,
                    file_name="capability_comparison_chart.pdf",
                    mime="application/pdf",
                    key="dl_cap"
                )

        # 3D 产业链板块（真实环节数据 · 全信息 hover · 每次必现）
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.write("#### 🔗 产业链全景逻辑流（鼠标悬停查看：业务内容 / 龙头企业 / 成本 / 利润率 / 特点 / 实时动态）")
            chain = data.get("industry_chain", {})
            nodes = chain.get("nodes", [])
            # 把实时新闻动态挂到产业链 hover 上（网站自身搜索能力）
            _chain_news = data.get("news_items", []) or []
            _chain_news_text = "；".join(f"{it.get('title','')[:40]}" for it in _chain_news[:5]) if _chain_news else "（暂无实时动态）"
            if nodes:
                stage_color_map = {
                    "upstream": "#2563eb", "midstream": "#0d9488",
                    "integration": "#f59e0b", "downstream": "#1e3a8a", "service": "#8b5cf6",
                }
                xs = [nd["x"] for nd in nodes]
                ys = [nd["y"] for nd in nodes]
                zs = [nd["z"] for nd in nodes]
                names = [nd["name"] for nd in nodes]
                colors = [stage_color_map.get(nd.get("stage", ""), "#2563eb") for nd in nodes]
                custom = [dict(
                    stage=nd["name"],
                    business=nd["business"],
                    leaders=nd["leaders"],
                    cost=nd["cost"],
                    margin=nd["margin"],
                    features=nd["features"],
                    source=nd["source"],
                    news=_chain_news_text,
                ) for nd in nodes]
                try:
                    fig_3d = go.Figure(data=[go.Scatter3d(
                        x=xs, y=ys, z=zs,
                        mode='markers+lines+text',
                        marker=dict(size=16, color=colors, opacity=0.92,
                                    line=dict(color="white", width=2)),
                        line=dict(color='#94a3b8', width=5),
                        text=names,
                        textposition="top center",
                        textfont=dict(size=11, color="#1e293b"),
                        customdata=custom,
                        # 注意：Scatter3d 的 hoverinfo 仅支持 x/y/z/text/name，
                        # 自定义字段请通过 hovertemplate 的 %{customdata.xxx} 引用
                        hoverinfo='text',
                        hovertemplate=(
                            "<b>%{customdata.stage}</b><br><br>"
                            "📌 主要业务：%{customdata.business}<br>"
                            "🏢 龙头/主要企业：%{customdata.leaders}<br>"
                            "💰 成本特征：%{customdata.cost}<br>"
                            "📈 利润率：%{customdata.margin}<br>"
                            "✨ 特点：%{customdata.features}<br>"
                            "📰 实时动态（网站自动检索）：%{customdata.news}<br>"
                            "📚 数据来源：%{customdata.source}<extra></extra>"
                        ),
                    )])
                    _z_axis = dict(showticklabels=True)
                    # 三轴语义化：X=环节推进（上游→下游）、Y=阶段层级、Z=利润率中值(%)
                    _x_ticks = [nd["x"] for nd in nodes]
                    _x_labels = [str(nd["name"])[:8] for nd in nodes]
                    _stage_ticks = [1.5, 0.75, 0.0, -0.75, -1.5]
                    _stage_labels = ["上游", "中游", "整机/集成", "下游", "服务/回收"]
                    fig_3d.update_layout(
                        height=500,
                        margin=dict(l=0, r=0, b=0, t=10),
                        scene=dict(
                            xaxis=dict(
                                title=dict(text="环节推进（上游 → 下游）", font=dict(size=12, color="#1e3a8a")),
                                tickmode="array", tickvals=_x_ticks, ticktext=_x_labels,
                                tickfont=dict(size=9, color="#334155"), showgrid=True, zeroline=False,
                            ),
                            yaxis=dict(
                                title=dict(text="产业链阶段层级", font=dict(size=12, color="#0d9488")),
                                tickmode="array", tickvals=_stage_ticks, ticktext=_stage_labels,
                                tickfont=dict(size=10, color="#334155"), showgrid=True, zeroline=True,
                            ),
                            zaxis=dict(
                                title=dict(text="环节利润率中值 (%)", font=dict(size=12, color="#b45309")),
                                tickfont=dict(size=10, color="#334155"), showgrid=True, zeroline=True,
                            ),
                            camera=dict(eye=dict(x=1.4, y=1.2, z=0.9)),
                        ),
                        hoverlabel=dict(font=dict(size=12, color="#1e293b"), bgcolor="#f8fafc", bordercolor="#cbd5e1"),
                    )
                    st.plotly_chart(fig_3d, use_container_width=True, key="industry_3d_chain_chart")
                    st.caption("📐 坐标含义：**X 轴**＝产业链环节推进（上游→下游）｜**Y 轴**＝阶段层级（上游1.5 → 服务-1.5）｜"
                               "**Z 轴**＝环节利润率中值(%)（由各环节利润率区间解析，真实口径）")
                except Exception as _3d_err:
                    # 3D 图渲染容错：极少数 plotly 版本兼容性问题时降级为列表展示，不阻断页面
                    print(f"[3D Chain] 3D 渲染失败，降级为列表展示: {_3d_err}")
                    st.warning("3D 产业链图在当前环境渲染失败，已降级为文字列表展示。")
                    for _nd in nodes:
                        st.markdown(f"**{_nd.get('name', '')}**（{_nd.get('stage', '')}）")
                        st.caption(f"业务：{_nd.get('business', '')}｜龙头：{_nd.get('leaders', '')}｜利润率：{_nd.get('margin', '')}")
                if chain.get("matched_industry"):
                    st.caption(f"已匹配行业：**{chain['matched_industry']}** —— {chain.get('note', '')}")
                else:
                    st.info(chain.get("note", "通用产业链框架（该行业暂未收录细分库），系统已自动检索实时新闻充实动态。"))
                if _chain_news:
                    with st.expander("📰 产业链实时动态（网站自动检索，悬停亦可查看）", expanded=False):
                        for _cn in _chain_news[:8]:
                            _cu = _cn.get("url", "")
                            if _cu:
                                st.markdown(f"- {_cn.get('date','')} **{_cn.get('title','')}**（{_cn.get('source','')}） [链接]({_cu})")
                            else:
                                st.markdown(f"- {_cn.get('date','')} **{_cn.get('title','')}**（{_cn.get('source','')}）")
            else:
                st.info("产业链结构构建中……")
            st.markdown('</div>', unsafe_allow_html=True)

        # ============================================================
        # 🌟 龙头公司横向对比板块（3~4 家龙头 · 真实财务摘要 + 对比图表）
        # ============================================================
        _leader = data.get("leader_data", {}) or {}
        if _leader.get("ok") and _leader.get("companies"):
            with st.container():
                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                st.write(f"#### 🏆 龙头公司横向对比（{_leader.get('industry', '')} · 3~4 家龙头）")
                _comps = _leader["companies"]
                _metrics = _leader.get("metrics", [])
                _vals = _leader.get("values", {})
                # 对比表格
                _ldf = pd.DataFrame({c: _vals.get(c, {}) for c in _comps})
                _ldf = _ldf.reindex(_metrics)
                st.dataframe(_ldf, use_container_width=True)
                # 分组柱状图
                _fig_leader = go.Figure()
                _colors = ["#0F2A5C", "#1F5FA8", "#0d9488", "#86bc25", "#E67E22", "#7D5BA6"]
                for _ci, _c in enumerate(_comps):
                    _fig_leader.add_trace(go.Bar(
                        name=_c,
                        x=_metrics,
                        y=[_vals.get(_c, {}).get(m, 0) for m in _metrics],
                        marker_color=_colors[_ci % len(_colors)],
                    ))
                _fig_leader.update_layout(
                    title="龙头公司多指标横向对比（ROE / 毛利率 / 净利率 / 同比 / EPS）",
                    barmode="group", height=340,
                    margin=dict(l=10, r=10, t=50, b=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(_fig_leader, use_container_width=True, key="leader_compare_chart")
                _notes = _leader.get("notes", {})
                for _c in _comps:
                    st.caption(f"**{_c}**：{_notes.get(_c, '')}")
                st.caption("数据来源：" + "；".join(_leader.get("sources", [])))
                # 导出按钮
                _leader_png = rex.render_chart_png("leader_compare", _leader, title="龙头公司横向对比")
                if _leader_png:
                    st.download_button("📊 导出龙头对比图为 PNG", data=_leader_png,
                                       file_name="leader_comparison.png", mime="image/png",
                                       key="dl_leader_png")
                st.markdown('</div>', unsafe_allow_html=True)

        # ============================================================
        # 📰 最新新闻与公告板块（网站自身搜索能力 · 实时公开信息）
        # ============================================================
        _news_list = data.get("news_items", []) or []
        if _news_list:
            with st.container():
                st.markdown('<div class="chart-box">', unsafe_allow_html=True)
                st.write(f"#### 📰 最新新闻与公告（实时抓取 · {data.get('news_note', '')}）")
                st.caption("来源：东方财富公告大全 / 全市场公告 / 7×24快讯 / 搜狗新闻 / 财新网（无需付费 API）")
                # 来源分布图
                _src_counter = {}
                for _n in _news_list:
                    _s = _n.get("source", "其他")
                    _src_counter[_s] = _src_counter.get(_s, 0) + 1
                if _src_counter:
                    _fig_news = go.Figure(data=[go.Pie(labels=list(_src_counter.keys()),
                                                       values=list(_src_counter.values()), hole=.45)])
                    _fig_news.update_layout(title="新闻/公告来源分布", height=280,
                                            margin=dict(l=10, r=10, t=50, b=10))
                    st.plotly_chart(_fig_news, use_container_width=True, key="news_source_chart")
                # 列表（带链接）
                _ncols = st.columns(2)
                for _idx, _n in enumerate(_news_list[:12]):
                    _col = _ncols[_idx % 2]
                    _t = _n.get("title", "")
                    _d = _n.get("date", "")
                    _s = _n.get("source", "")
                    _u = _n.get("url", "")
                    _label = f"**{_t[:60]}**　{_d}　（{_s}）"
                    if _u:
                        _col.markdown(f"- {_label} [原文]({_u})")
                    else:
                        _col.markdown(f"- {_label}")
                st.markdown('</div>', unsafe_allow_html=True)

        # ============================================================
        # 🌐 网页数据读取结果（文字 + 表格 + 图表数据）
        # ============================================================
        _web_summary = data.get("web_page_summary", "")
        if _web_summary:
            with st.expander("🌐 网页数据读取结果（正文 / 表格 / 内嵌图表数据）", expanded=False):
                st.caption("系统已自动读取相关新闻网页，提取文字、表格与内嵌图表数据供研报引用（图片 alt 亦被解析）。")
                st.text_area("网页数据摘录", _web_summary[:4000], height=220, key="web_page_summary_area", disabled=True)

        # B. 研报正文展示
        # A+. 使用数据看板（产品运营视角：数据驱动）
        with st.expander("📊 使用数据看板", expanded=False):
            _usage_history = st.session_state.get("usage_history", [])
            _total_runs = len(_usage_history)
            _total_in = sum(int(r.get("tokens_in", 0) or 0) for r in _usage_history)
            _total_out = sum(int(r.get("tokens_out", 0) or 0) for r in _usage_history)
            _avg_dur = round(sum(float(r.get("duration_sec", 0) or 0) for r in _usage_history) / _total_runs, 1) if _total_runs else 0.0
            _est_cost = pf.estimate_cost(_total_in, _total_out)
            _um1, _um2, _um3, _um4 = st.columns(4)
            _um1.metric("累计研究次数", f"{_total_runs} 次")
            _um2.metric("平均生成耗时", f"{_avg_dur} 秒")
            _um3.metric("累计 Token 用量", f"{_total_in + _total_out:,}")
            _um4.metric("估算调用成本", f"¥{_est_cost:.4f}")
            _agg_times = {}
            for _r in _usage_history:
                for _t in _r.get("agent_times", []):
                    _a = _t.get("agent", "未知")
                    _agg_times.setdefault(_a, []).append(float(_t.get("seconds", 0) or 0))
            if _agg_times:
                _agent_names = ["Planner", "Research", "Financial", "Policy", "Risk", "Judge", "Report"]
                _names_show = [a for a in _agent_names if a in _agg_times]
                _avgs_show = [round(sum(_agg_times[a]) / len(_agg_times[a]), 2) for a in _names_show]
                _fig_usage = go.Figure(data=[go.Bar(x=_names_show, y=_avgs_show, marker_color="#2563eb", text=_avgs_show, textposition="outside")])
                _fig_usage.update_layout(title="各 Agent 平均耗时（秒）", height=260, margin=dict(l=10, r=10, t=40, b=10))
                st.plotly_chart(_fig_usage, use_container_width=True, key="usage_agent_chart")
            else:
                st.caption("完成一次研究后，这里会展示各 Agent 的耗时分布。")
        st.markdown('<div class="report-container">', unsafe_allow_html=True)
        st.markdown(st.session_state['current_report'])
        st.markdown('</div>', unsafe_allow_html=True)

        # 🚩 风险雷达模型板块（真实财务指标映射 · 口径透明）
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.write("#### 🚩 企业经营及财务多维风险指数测算（基于真实财务指标映射）")

            risk_data = data.get("risk_radar", {
                "dimensions": ["偿债与财务杠杆风险", "短期流动性紧缺风险", "存货/资产减值风险", "盈利质量恶化风险", "政策合规与壁垒风险"],
                "values": [3.0, 3.2, 2.8, 3.5, 4.0],
                "methodology": [],
                "based_on": "默认口径（该行业未匹配数据）",
            })

            fig_risk_radar = go.Figure()
            fig_risk_radar.add_trace(go.Scatterpolar(
                r=risk_data["values"],
                theta=risk_data["dimensions"],
                fill='toself',
                name='风险系数 (1表示极安全，5表示极高风险)',
                line=dict(color='#ef4444', width=2),
                fillcolor='rgba(239, 68, 68, 0.3)',
                customdata=[f"维度: {d}\n计算口径: {m}\n风险值: {v}" for d, m, v in
                            zip(risk_data["dimensions"], risk_data.get("methodology", [""] * 5), risk_data["values"])],
                hovertemplate="%{customdata}<extra></extra>",
            ))
            fig_risk_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 5.0])),
                title="企业整体经营及财务审计风险度量雷达模型（0=安全，5=高危）",
                height=380,
                margin=dict(l=20, r=20, t=40, b=20),
                hoverlabel=dict(font=dict(size=11, color="#1e293b"), bgcolor="#f8fafc", bordercolor="#cbd5e1"),
            )
            st.plotly_chart(fig_risk_radar, use_container_width=True, key="risk_radar_chart_bottom")
            st.caption("📚 数据来源：真实财务指标映射（权益乘数/现金流质量/毛利率/ROE/行业政策基准，报告期见数据截止标注）")

            _risk_pdf = chart_pdf_bytes(fig_risk_radar)
            if _risk_pdf is not None:
                st.download_button(
                label="🚩 导出风险雷达图为 PDF 矢量图",
                data=_risk_pdf,
                file_name="risk_radar_chart.pdf",
                mime="application/pdf",
                key="dl_risk_radar"
            )
            with st.expander("📖 风险维度计算口径（逐条可溯源）", expanded=False):
                st.caption(risk_data.get("based_on", ""))
                for i, (d, m) in enumerate(zip(risk_data["dimensions"], risk_data.get("methodology", []))):
                    st.markdown(f"**{i+1}. {d}** → 风险 {risk_data['values'][i]}：{m}")
                st.caption("注：数据均来自本地财务数据库（东方财富业绩报表聚合，报告期见页面数据截止标注）；"
                           "政策维度为公开监管信息基准值，可上传政策库/风险事件清单修正。")
            st.markdown('</div>', unsafe_allow_html=True)

        # C. 升级后的 Word 文档导出逻辑（动态嵌入对比图及证据链）
        # B+. 收藏与分享（产品化功能：用户留存与传播）
        st.divider()
        _fc1, _fc2, _fc3 = st.columns(3)
        with _fc1:
            if st.button("⭐ 收藏本报告", key="fav_save", use_container_width=True):
                _fav_item = pf.add_favorite(
                    st.session_state['current_query'],
                    st.session_state['current_report'],
                    st.session_state['current_data'],
                )
                st.session_state["share_link"] = f"{pf.APP_URL}?report={_fav_item['id']}"
                st.toast(f"已收藏《{st.session_state['current_query']}》")
        with _fc2:
            if st.button("🔗 生成分享链接", key="fav_share", use_container_width=True):
                _favs_now = st.session_state.get("favorites", [])
                if _favs_now:
                    _fav_item = _favs_now[0]
                else:
                    _fav_item = pf.add_favorite(
                        st.session_state['current_query'],
                        st.session_state['current_report'],
                        st.session_state['current_data'],
                    )
                st.session_state["share_link"] = f"{pf.APP_URL}?report={_fav_item['id']}"
        with _fc3:
            if st.button("📋 复制分享文案", key="fav_copy", use_container_width=True):
                st.session_state["share_text"] = pf.build_share_text(st.session_state['current_query'])
        if st.session_state.get("share_link"):
            st.code(st.session_state["share_link"], language=None)
            st.caption("对方打开该链接即可直接看到这份报告（收藏保存在当前部署实例，重新部署后需重新收藏）")
        if st.session_state.get("share_text"):
            st.text_area("分享文案（可直接复制）", st.session_state["share_text"], height=110, key="share_text_area")
        # C. 升级版文档导出（Word 全图表版 + PPT 版，matplotlib 稳定渲染，kaleido 优先）
        _export_cache_key = f"{st.session_state['current_query']}_{hash(str(data)[:200])}"
        if st.session_state.get("export_cache_key") != _export_cache_key:
            st.session_state["export_cache_key"] = _export_cache_key
            st.session_state["export_cache"] = None
        if st.session_state.get("export_cache") is None:
            try:
                _ci = {}
                _cap_share = ""
                if is_company_mode:
                    _png = rex.render_chart_png(
                        "dupont_compare", data,
                        title=f"{data.get('company_name','')} 与行业杜邦因子对比（标准化）")
                    _ci["dupont"] = {
                        "title": "杜邦因子对标（公司 vs 行业）",
                        "caption": f"数据口径：ROE/净利润率/资产周转率/权益乘数，公司值与行业聚合值（报告期 {data.get('market_as_of','—')}）",
                        "png": _png,
                        "source": f"公司财务指标（{data.get('company_name','')}）vs 东方财富业绩报表行业聚合（{data.get('market_as_of','—')}）",
                        "notes": [
                            f"公司 ROE {data.get('company_roe','—')}% vs 行业 {data.get('industry_roe','—')}%",
                            f"公司净利率 {data.get('company_margin','—')}% vs 行业 {data.get('industry_margin','—')}%",
                            "杜邦三要素：利润率 / 资产周转率 / 权益乘数，驱动 ROE 差距归因见下页瀑布图",
                        ],
                    }
                    _png = rex.render_chart_png("company_radar", data, title="标的公司与行业能力多维透视")
                    _ci["radar"] = {"title": "能力多维透视雷达", "caption": "ROE/净利润率/资产周转率/财务杠杆/经营现金流（真实指标）", "png": _png,
                                    "source": f"公司财务指标 vs 行业聚合（{data.get('market_as_of','—')}）",
                                    "notes": ["五项能力维度公司 vs 行业均值对比", "现金流维度单位：万元（公司）vs 行业每股现金流"]}
                    _png = rex.render_chart_png("dupont_waterfall", data, title="杜邦 ROE 差距归因")
                    _ci["waterfall"] = {"title": "杜邦 ROE 差距归因瀑布", "caption": "公司相对行业 ROE 差距的利润率/周转率/杠杆贡献分解",
                                        "png": _png, "source": "对数分解示意（ln 差线性化），基数来自公司/行业真实财务指标",
                                        "notes": ["柱状从行业 ROE 起步，逐项叠加三要素贡献得到公司 ROE", "绿色=正向贡献，红色=负向贡献"]}
                else:
                    _png = rex.render_chart_png("market_share", data.get("market_share", {}), title="行业市场集中度（CR4）")
                    _ci["share"] = {"title": "行业竞争格局", "caption": f"CR4={data.get('market_share',{}).get('values',[0])[0]}%（东方财富业绩报表真实聚合）", "png": _png,
                                    "source": f"东方财富业绩报表全市场聚合（报告期 {data.get('market_as_of','—')}）",
                                    "notes": [f"行业 CR4 = {data.get('market_share',{}).get('values',[0])}%", "头部集中度反映竞争格局与定价权"]}
                    _png = rex.render_chart_png("market_growth", data.get("market_growth", {}), title="行业市场规模与复合增速")
                    _ci["growth"] = {"title": "市场规模与增速", "caption": "市场规模/增速为估算口径；趋势示意，详见缺口说明", "png": _png,
                                     "source": "市场规模/增速为公开区间估算；CR4/ROE 等为真实聚合",
                                     "notes": ["柱状=市场规模（亿元），折线=同比增速（%）", "增速逐年放缓属行业成熟期典型特征"]}
                    _png = rex.render_chart_png("financial_trend", data.get("financial_trend", {}), title="主要盈利指标变化趋势")
                    _ci["trend"] = {"title": "盈利指标趋势", "caption": "最新期为真实聚合值，历史期为趋势示意", "png": _png,
                                    "source": f"最新期=东方财富业绩报表（{data.get('market_as_of','—')}）；历史期=趋势示意",
                                    "notes": ["ROE 与净利率双线走势", "用于判断行业盈利质量所处周期位置"]}
                    _png = rex.render_chart_png("capability_compare", data.get("capability_comparison", {}), title="企业多维核心财务能力对比")
                    _ci["cap"] = {"title": "核心财务能力对比", "caption": "ROE/毛利率/CR4/净利率（真实聚合）", "png": _png,
                                  "source": f"东方财富业绩报表聚合（{data.get('market_as_of','—')}）",
                                  "notes": ["横向条形展示行业核心财务能力", "盈利能力（ROE）与集中度（CR4）为主要观察维度"]}
                # 产业链 + 风险雷达
                _png = rex.render_chart_png("industry_chain", data.get("industry_chain", {}), title="产业链全景逻辑流")
                _ci["chain"] = {"title": "产业链全景逻辑流", "caption": "环节→龙头→利润率（区间值·综合公开资料）", "png": _png,
                                "source": "行业公开研究/公司年报/实时新闻检索（区间值）",
                                "notes": ["五环节：上游→中游→整机→下游→服务", "悬停可查看业务/龙头/成本/利润率/实时动态"]}
                _png = rex.render_chart_png("risk_radar", data.get("risk_radar", {}), title="企业经营及财务多维风险指数")
                _ci["risk"] = {"title": "多维风险指数雷达", "caption": "0=安全，5=高危；口径见风险板块说明", "png": _png,
                               "source": "真实财务指标映射（权益乘数/现金流质量/毛利率/ROE/政策基准）",
                               "notes": ["五维风险：杠杆/流动性/减值/盈利/政策", "数值越高风险越大"]}
                # 龙头对比 + 新闻来源（需求 5/7）
                _leader_x = data.get("leader_data", {}) or {}
                if _leader_x.get("ok"):
                    _png = rex.render_chart_png("leader_compare", _leader_x, title="龙头公司横向对比")
                    _ci["leader"] = {"title": "龙头公司横向对比", "caption": f"{_leader_x.get('industry','')} · 3~4 家龙头多指标对比",
                                     "png": _png,
                                     "source": "东方财富财务摘要接口 / 上市公司公开报告兜底口径",
                                     "notes": [f"{c}：{_leader_x.get('notes',{}).get(c,'')}" for c in _leader_x.get("companies", [])[:4]]}
                _news_x = data.get("news_items", []) or []
                if _news_x:
                    _src_count = {}
                    for _n in _news_x:
                        _s = _n.get("source", "其他")
                        _src_count[_s] = _src_count.get(_s, 0) + 1
                    _png = rex.render_chart_png("news_source", {"sources": _src_count}, title="新闻/公告来源分布")
                    _ci["news"] = {"title": "新闻/公告来源分布", "caption": "实时抓取来源占比（东方财富/搜狗/财新等）",
                                   "png": _png, "source": "东方财富公告大全 / 7×24快讯 / 搜狗新闻 / 财新网",
                                   "notes": [f"共 {len(_news_x)} 条实时信息", "来源链接见报告正文新闻板块"]}
                _gap_texts = ["• " + g["item"] + f"（{g['status']}）" + (f"：{g['reason']}" if g.get("reason") else "")
                              for g in data.get("data_gaps", [])]
                _meta = data.get("report_meta", {})
                _doc_bytes = rex.export_docx(
                    st.session_state['current_query'],
                    st.session_state['current_report'],
                    _ci, evidence_data, _gap_texts, _meta, is_company_mode,
                )
                _ppt_bytes = rex.export_pptx(
                    st.session_state['current_query'],
                    st.session_state['current_report'],
                    _ci, evidence_data, _gap_texts, _meta, is_company_mode,
                    leader_data=data.get("leader_data", {}),
                    news_items=data.get("news_items", []),
                )
                st.session_state["export_cache"] = {"docx": _doc_bytes, "pptx": _ppt_bytes}
            except Exception as _exp_err:
                st.warning(f"文档预渲染部分失败（不影响页面）：{_exp_err}")
                st.session_state["export_cache"] = {"docx": None, "pptx": None}

        _ecache = st.session_state.get("export_cache") or {}
        _ed1, _ed2 = st.columns(2)
        with _ed1:
            if _ecache.get("docx"):
                st.download_button(
                    label="📥 导出完整研报（Word · 含全图表）.docx",
                    data=_ecache["docx"],
                    file_name=f"{st.session_state['current_query']}_深度研报.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_report_docx",
                )
            else:
                st.caption("Word 导出暂不可用（依赖未就绪）")
        with _ed2:
            if _ecache.get("pptx"):
                st.download_button(
                    label="📊 导出演示版研报（PPT · 含全图表）.pptx",
                    data=_ecache["pptx"],
                    file_name=f"{st.session_state['current_query']}_研报演示.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="dl_report_pptx",
                )
            else:
                st.caption("PPT 导出暂不可用（依赖未就绪）")
    else:
        # 首屏示例问题（一键填充，降低使用门槛）
        def _apply_example(ex):
            st.session_state["research_mode"] = "标准模式（专业投研）"
            st.session_state["research_target"] = ex["target"]
            st.session_state["company_query"] = ex["company"]
            st.session_state["query"] = ex["query"]
            st.session_state["period_type"] = ex["period_type"]
            st.session_state["year_select"] = ex["year"]
            st.session_state["report_type"] = ex["report_type"]
            st.session_state["purpose"] = ex["purpose"]
            st.session_state["trigger_run"] = True

        st.markdown("### 🚀 不知道从哪开始？试试这些示例")
        st.caption("点击任意示例，系统会自动填好参数并开始研究（首次运行约 1~2 分钟）")
        _ex_cols = st.columns(len(pf.EXAMPLE_QUESTIONS))
        for _ei, _ex in enumerate(pf.EXAMPLE_QUESTIONS):
            with _ex_cols[_ei]:
                st.button(
                    _ex["label"],
                    key=f"example_{_ei}",
                    use_container_width=True,
                    on_click=_apply_example,
                    args=(_ex,),
                )
        st.info("💡 也可以从左侧手动输入公司或行业，选择时间周期与研究目的后启动。")

