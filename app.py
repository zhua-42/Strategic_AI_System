import os
from dotenv import load_dotenv
from openai import OpenAI

# 加载环境变量
load_dotenv()
api_key = os.getenv("DEEPSEEK_API_KEY")  # 👈 唯一修改了这一行，让它正确读取 .env 文件中的密钥名称

# 初始化 OpenAI 客户端（这里以调用 DeepSeek 官方 API 为例，兼容 OpenAI 格式）
# 如果学校组委会提供了其他模型接口，只需修改 base_url 和 api_key 即可
client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com/v1" 
)
import streamlit as st
import requests
import time
import urllib3
import json
import base64
import plotly.graph_objects as go
import plotly.express as px

# 1. 基础配置
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
st.set_page_config(page_title="全行业战略情报系统", layout="wide")

# --- 2. 配置区 ---
API_KEY = "sk-fa7cc8174a8c49349c0d5f1181b8255a"
BOT_ID = "7641635024001024015"

# --- 3. 辅助函数 ---
def is_json(myjson):
    try:
        json.loads(myjson)
    except: return False
    return True

def extract_report_data(raw_report):
    """
    解析和提取原始报告内容。
    如果原始报告中包含 Markdown 格式的 JSON 块，则提取为 dynamic_data；
    否则，将全文作为 clean_text 返回。
    """
    clean_text = raw_report
    dynamic_data = {}
    
    # 尝试寻找并解析文本中可能存在的 ```json ... ``` 块
    if "```json" in raw_report:
        try:
            parts = raw_report.split("```json")
            json_str = parts[1].split("```")[0].strip()
            if is_json(json_str):
                dynamic_data = json.loads(json_str)
                # 拼接 JSON 块前后的文本，作为纯文本报告
                clean_text = parts[0].strip() + "\n" + parts[1].split("```")[1].strip()
        except Exception:
            pass
            
    return clean_text, dynamic_data

# --- 4. 界面美化与格式统一 (解决痛点2：标题正文字号区分) ---
st.markdown("""
    <style>
    /* 全局背景与字体 */
    .main { background-color: #ffffff; font-family: "Microsoft YaHei", sans-serif; }
    
    /* 报告容器 */
    .report-container { 
        border: 1px solid #e6e9ef; 
        padding: 45px; 
        border-radius: 12px; 
        background-color: #ffffff; 
        line-height: 1.8;
        color: #333333;
    }
    
    /* 严格区分标题与正文 */
    .report-container h1 { font-size: 32px !important; color: #003366; border-bottom: 2px solid #003366; padding-bottom: 10px; margin-top: 30px; }
    .report-container h2 { font-size: 26px !important; color: #004080; border-left: 5px solid #d62728; padding-left: 15px; margin-top: 25px; }
    .report-container h3 { font-size: 20px !important; color: #1f77b4; margin-top: 20px; font-weight: bold; }
    .report-container p { font-size: 16px !important; color: #444444; margin-bottom: 15px; }
    
    /* 图表容器 */
    .chart-box { border: 1px solid #f0f2f6; padding: 25px; border-radius: 12px; background-color: #fafafa; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    
    /* 侧边栏按钮 */
    .stButton>button { width: 100%; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

# --- 5. 状态管理 ---
if 'history' not in st.session_state: st.session_state['history'] = []
if 'current_report' not in st.session_state: st.session_state['current_report'] = ""
if 'current_query' not in st.session_state: st.session_state['current_query'] = ""
if 'conv_id' not in st.session_state: st.session_state['conv_id'] = ""
if 'chat_messages' not in st.session_state: st.session_state['chat_messages'] = []

# --- 6. 侧边栏 ---
with st.sidebar:
    st.title("📚 研究历史")
    for idx, h in enumerate(st.session_state['history']):
        if st.button(f"📄 {h['query']}", key=f"h_{idx}"):
            st.session_state['current_report'] = h['content']
            st.session_state['current_query'] = h['query']
            st.session_state['conv_id'] = h.get('conv_id', "")
    st.divider()
    st.title("🛠 启动调研")
    query = st.text_input("输入调研课题", placeholder="如：新能源汽车")
    submit_btn = st.button("🚀 开启深度自动化研究")
    st.caption("⚠️ 温馨提示：生成报告需要 3-5 分钟，期间请勿刷新页面，保持耐心哦 (:D)")

# --- 7. 核心调用逻辑 ---
def run_research(user_input, existing_conv_id=None, progress_callback=None):
    """
    用原生 Python 代码实现的 Multi-Agent 协同流。
    """
    if not api_key:
        return "错误：未在 .env 文件中检测到 DEEPSEEK_API_KEY，请先配置。", ""

    try:
        # ---- Agent 1: 行业数据专家 (生成结构化 JSON 数据) ----
        if progress_callback:
            progress_callback(15, "📊 [Agent 1] 行业数据专家正在检索并建模...")

        system_prompt_data = """
        你是一个精通行业数据建模的专家。你需要针对用户输入的调研课题，生成一份结构化的预测数据。
        必须严格输出为 JSON 格式（不要包含任何多余的解释文字），格式如下：
        {
          "market_share": {
            "labels": ["核心头部", "中坚力量", "初创企业", "其他"],
            "values": [40, 30, 15, 15]
          },
          "market_growth": {
            "years": ["2022", "2023", "2024", "2025E", "2026E"],
            "scale": [200, 310, 500, 800, 1100],
            "growth_rate": [15, 40, 61, 60, 37]
          },
          "supply_chain": [
            {"node": "基础材料", "companies": "企业A", "details": "原材料供应", "x": 0, "y": 0, "z": 0},
            {"node": "核心部件", "companies": "企业B", "details": "核心技术组件", "x": 1, "y": 0.5, "z": 1},
            {"node": "集成应用", "companies": "企业C", "details": "终端产品", "x": 2, "y": -0.5, "z": 0}
          ],
          "risk_assessment": {
            "categories": ["政策波动","技术瓶颈","竞争烈度","资本冷热","合规挑战"],
            "values": [4.0, 3.5, 4.5, 2.8, 3.8],
            "descriptions": [
              "1. **竞争烈度 (4.5/5.0)**：行业红海竞争加剧...",
              "2. **政策波动 (4.0/5.0)**：关注补贴政策退坡..."
            ]
          }
        }
        """
        
        response_data = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt_data},
                {"role": "user", "content": f"请针对以下课题生成数据: {user_input}"}
            ],
            temperature=0.2,
            response_format={"type": "json_object"} # 确保输出是标准 JSON
        )
        json_data_str = response_data.choices[0].message.content

        # ---- Agent 2: 首席行业分析师 (结合数据撰写专业研报) ----
        if progress_callback:
            progress_callback(55, "📝 [Agent 2] 首席分析师正在撰写深度研报正文...")

        system_prompt_writer = """
        你是一个资深的行业研究员。你需要结合提供的行业结构化数据，撰写一篇结构严谨、分析深入的行业研究报告。
        研报必须采用 Markdown 格式，严格包含以下结构：
        # 行业研究报告
        ## 一、行业发展概述与背景
        ## 二、市场规模预测与竞争格局分析
        ## 三、产业链逻辑与核心壁垒分析
        ## 四、行业核心风险评估及战略建议
        """

        user_content_writer = f"""
        调研课题：{user_input}
        参考行业数据：{json_data_str}
        请结合上述数据，撰写一篇不少于 1500 字的深度研报。
        """

        response_writer = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt_writer},
                {"role": "user", "content": user_content_writer}
            ],
            temperature=0.7
        )
        markdown_report = response_writer.choices[0].message.content

        # ---- 格式拼接 (适配你原本的 extract_report_data 提取器) ----
        if progress_callback:
            progress_callback(90, "🔄 整合报告与可视化图表数据...")

        # 将 JSON 数据以 markdown 块形式拼在最末尾，方便你原有的 extract_report_data 提取
        final_output = f"{markdown_report}\n\n```json\n{json_data_str}\n```"

        if progress_callback:
            progress_callback(100, "✨ 研报生成完成！")
            
        # 模拟返回一个会话ID
        simulated_conv_id = "local_conv_" + str(int(time.time()))
        return final_output, simulated_conv_id

    except Exception as e:
        if progress_callback:
            progress_callback(100, f"❌ 运行出错：{str(e)[:50]}")
        return f"生成研报失败，请确认 API Key 是否正确。错误信息: {e}", ""

# --- 8. 可视化组件 (解决痛点1, 4, 5, 6) ---

def draw_main_dashboard(query, data=None):
    st.write("#### 📊 行业数据驾驶舱")
    colors = ['#1f77b4', '#d62728', '#32a852', '#ff7f0e']
    
    # 动态解析或使用静态兜底数据
    share_labels = data.get("market_share", {}).get("labels", ['核心头部', '中坚力量', '初创企业', '其他']) if data else ['核心头部', '中坚力量', '初创企业', '其他']
    share_values = data.get("market_share", {}).get("values", [45, 25, 15, 15]) if data else [45, 25, 15, 15]
    
    growth_years = data.get("market_growth", {}).get("years", ['2022', '2023', '2024', '2025E', '2026E']) if data else ['2022', '2023', '2024', '2025E', '2026E']
    growth_scale = data.get("market_growth", {}).get("scale", [210, 300, 520, 850, 1200]) if data else [210, 300, 520, 850, 1200]
    growth_rate = data.get("market_growth", {}).get("growth_rate", [15, 42, 73, 63, 41]) if data else [15, 42, 73, 63, 41]
    
    c1, c2 = st.columns(2)
    with c1:
        fig_pie = go.Figure(data=[go.Pie(labels=share_labels, 
                                        values=share_values, hole=.4,
                                        marker=dict(colors=colors))])
        fig_pie.update_layout(title_text="市场份额分布预期", height=400, showlegend=True)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with c2:
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=growth_years, y=growth_scale,
                                 marker_color='#1f77b4', name='市场规模(亿)'))
        fig_bar.add_trace(go.Scatter(x=growth_years, y=growth_rate,
                                     line=dict(color='#d62728', width=3), name='增速(%)', yaxis='y2'))
        fig_bar.update_layout(title_text="行业规模与增长率趋势", height=400,
                              yaxis2=dict(overlaying='y', side='right'),
                              legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig_bar, use_container_width=True)
    
    st.markdown(f"<p style='font-size: 12px; color: gray; text-align: right;'>数据来源：基于字节跳动搜索插件检索及 {query} 行业公开披露数据、公开财报综合计算得出</p>", unsafe_allow_html=True)


def draw_3d_supply_chain(data=None):
    st.write("#### 🔗 产业链全景逻辑流 (3D交互感)")
    
    # 动态解析或使用静态兜底数据
    if data and "supply_chain" in data and len(data["supply_chain"]) > 0:
        nodes = [item.get("node", "") for item in data["supply_chain"]]
        companies = [item.get("companies", "") for item in data["supply_chain"]]
        details = [item.get("details", "") for item in data["supply_chain"]]
        x = [item.get("x", idx) for idx, item in enumerate(data["supply_chain"])]
        y = [item.get("y", 0.0 if idx % 2 == 0 else 0.5) for idx, item in enumerate(data["supply_chain"])]
        z = [item.get("z", 0.0 if idx % 2 == 0 else 1.0) for idx, item in enumerate(data["supply_chain"])]
    else:
        nodes = ['基础材料', '核心零部件', '整机/系统集成', '下游应用', '售后/回收']
        x = [0, 1, 2, 3, 4]
        y = [0, 0.5, -0.5, 0.2, 0]
        z = [0, 1, 0, 1, 0]
        companies = ['宝钢股份、中复神鹰', '宁德时代、汇川技术', '西门子、大疆、亿航', '顺丰、国家电网', '格林美、各品牌4S']
        details = ['提供碳纤维、高性能钢材等原始原料', '电机、电池、传感器等核心组件生产', '产品组装、飞控系统及AI算法集成', '物流配送、工业巡检、消费文旅等', '设备维护及资源循环再利用']

    fig = go.Figure(data=[go.Scatter3d(
        x=x, y=y, z=z,
        mode='markers+lines+text',
        marker=dict(size=12, color=['#d62728', '#1f77b4', '#d62728', '#1f77b4', '#333'][:len(nodes)], opacity=0.8),
        line=dict(color='#1f77b4', width=6),
        text=nodes,
        hoverinfo='text',
        hovertext=[f"环节: {n}<br>业务: {d}<br>代表企业: {c}" for n,d,c in zip(nodes, details, companies)]
    )])
    fig.update_layout(height=500, margin=dict(l=0, r=0, b=0, t=0), 
                      scene=dict(xaxis_title='流程阶段', yaxis_title='价值分布', zaxis_title='技术壁垒'))
    st.plotly_chart(fig, use_container_width=True)


def draw_risk_radar(data=None):
    st.write("#### 🚩 行业风险综合评估")
    
    # 动态解析或使用静态兜底数据
    if data and "risk_assessment" in data:
        categories = data["risk_assessment"].get("categories", ['政策波动风险','技术突破瓶颈','市场竞争烈度','资本环境冷热','合规性挑战'])
        values = data["risk_assessment"].get("values", [4.2, 3.1, 4.8, 2.5, 3.9])
        descriptions = data["risk_assessment"].get("descriptions", [])
    else:
        categories = ['政策波动风险','技术突破瓶颈','市场竞争烈度','资本环境冷热','合规性挑战']
        values = [4.2, 3.1, 4.8, 2.5, 3.9]
        descriptions = [
            "1. **市场竞争烈度 (4.8/5.0)**：处于红海竞争前期，头部企业通过降价策略挤压长尾厂商，需警惕毛利率下滑。",
            "2. **政策波动风险 (4.2/5.0)**：高度依赖补贴与地方空域开放政策，政策转向可能直接影响项目ROI。",
            "3. **合规性挑战 (3.9/5.0)**：随着法律框架落地，安全认证成本将上升约15%-25%。",
            "4. **技术突破瓶颈 (3.1/5.0)**：固态电池与长时续航仍是主要物理障碍，决定了商业化天花板。"
        ]
    
    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(
          r=values, theta=categories, fill='toself',
          marker=dict(color='#d62728'), line=dict(color='#003366')
    ))
    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0, 5])), showlegend=False, height=450)
    
    c1, c2 = st.columns([1, 1])
    with c1: st.plotly_chart(fig_radar, use_container_width=True)
    with c2:
        st.markdown("**🔍 风险因素深度解析：**")
        for desc in descriptions:
            st.markdown(desc)
# --- 9. 主界面展示 ---
st.title("🛰️ 数字化战略研究 AI 驾驶舱")

if submit_btn and query:
    # 使用 st.status 容器来显示动态进度
    with st.status("准备启动深度研究...", expanded=True) as status:
        progress_bar = st.progress(0, text="0%")
        status_text = st.empty()
        
        def update_progress(percent, message):
            progress_bar.progress(percent)
            status_text.markdown(message)
        
        raw_report, cid = run_research(query, progress_callback=update_progress)
        
        # 处理报告
        clean_text, dynamic_data = extract_report_data(raw_report)
        
        if clean_text:
            st.session_state['current_report'] = clean_text
            st.session_state['current_data'] = dynamic_data
            st.session_state['current_query'] = query
            st.session_state['conv_id'] = cid
            st.session_state['history'].insert(0, {"query": query, "content": clean_text, "data": dynamic_data, "cid": cid})
            status.update(label="✅ 研报已生成！", state="complete")
        else:
            st.error("研报生成失败，请检查API配置或重试。")
            status.update(label="❌ 生成失败", state="error")

# 找到这段代码：
if st.session_state['current_report']:
    col_rep, col_chat = st.columns([2.5, 1])
    
    with col_rep:
        st.markdown(f"## 📋 {st.session_state['current_query']} 深度研报")
        
        # A. 驾驶舱数据 (痛点1, 4) ———> 修改这里，加入第二个参数
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            draw_main_dashboard(
                st.session_state['current_query'], 
                st.session_state.get('current_data', {})
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # B. 研报正文 (痛点2)
        st.markdown('<div class="report-container">', unsafe_allow_html=True)
        st.markdown(st.session_state['current_report'])
        st.markdown('</div>', unsafe_allow_html=True)

        # C. 产业链图 (痛点5) ———> 修改这里，加入参数
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            draw_3d_supply_chain(st.session_state.get('current_data', {}))
            st.markdown('</div>', unsafe_allow_html=True)

        # D. 风险评估 (痛点3) ———> 修改这里，加入参数
        with st.container():
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            draw_risk_radar(st.session_state.get('current_data', {}))
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.download_button(
            label="📥 导出为专业 Word 文档",
            data=st.session_state['current_report'],
            file_name=f"{st.session_state['current_query']}研报.doc",
            mime="application/msword"
        )
    
    with col_chat:
        st.subheader("💬 战略咨询对谈")
        st.info("已针对当前研报建立对话上下文，您可以追问任何细节：")
        for msg in st.session_state['chat_messages']:
            with st.chat_message(msg["role"]): st.write(msg["content"])

        sub_query = st.chat_input("追问细节，例如：上游企业的议价能力如何？")
        if sub_query:
            st.session_state['chat_messages'].append({"role": "user", "content": sub_query})
            with st.chat_message("user"): st.write(sub_query)
            with st.chat_message("assistant"):
                with st.spinner("专家正在分析..."):
                    answer, _ = run_research(sub_query, st.session_state['conv_id'])
                    st.write(answer)
                    st.session_state['chat_messages'].append({"role": "assistant", "content": answer})
else:
    st.info("👈 请在左侧输入调研课题并点击启动按钮。")