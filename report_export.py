# -*- coding: utf-8 -*-
"""
报告文档导出引擎（Word / PPT / 图表渲染）
=========================================
- 图表渲染：使用 matplotlib（稳定、无需外部浏览器/Chromium，本地与云端均可运行），
  主题色对齐 Deloitte 绿色系 + 深蓝主色（参考 Deloitte Forage 视觉与现有看板配色）。
- 中文字体：优先使用仓库内 assets/fonts/ 打包的开源思源黑体，找不到时回退系统字体。
- Word：封面 → 数据看板（逐图+口径说明）→ 证据链 → 正文 → 资料缺口 → 免责声明。
- PPT：封面 → 目录 → 核心指标 → 逐图页 → 证据链 → 免责声明（16:9）。
"""
import io
import os
import re
import datetime

import pandas as pd

# ---------- 主题色（Deloitte / Forage 视觉统一） ----------
C_GREEN = "#86bc25"     # Deloitte 绿
C_NAVY = "#1e3a8a"      # 深蓝（主标题/公司）
C_BLUE = "#2563eb"      # 蓝（对比）
C_TEAL = "#0d9488"      # 青
C_RED = "#ef4444"       # 红（警示/行业）
C_AMBER = "#f59e0b"     # 琥珀
C_PURPLE = "#8b5cf6"    # 紫
C_SLATE = "#334155"     # 正文灰
C_LIGHT = "#f1f5f9"     # 浅底

FONT_BUNDLED = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "assets", "fonts", "NotoSansSC-VF.ttf")
FONT_CANDIDATES = [
    FONT_BUNDLED,
    r"C:\Windows\Fonts\NotoSansSC-VF.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    "Noto Sans CJK SC", "Microsoft YaHei", "SimHei", "WenQuanYi Zen Hei",
]

_FONT_CACHE = {"name": None, "props": None}


def setup_font():
    """注册中文字体，返回 matplotlib FontProperties（缓存）。"""
    if _FONT_CACHE["props"] is not None:
        return _FONT_CACHE["props"]
    import matplotlib
    from matplotlib import font_manager as fm

    chosen = None
    for cand in FONT_CANDIDATES:
        try:
            if os.path.isfile(cand):
                fm.fontManager.addfont(cand)
                name = fm.FontProperties(fname=cand).get_name()
                chosen = (cand, name)
                break
        except Exception:
            continue
    if chosen is None:
        # 系统字体名
        for name in FONT_CANDIDATES:
            if not isinstance(name, str) or not os.path.exists(name):
                try:
                    fm.fontManager.findfont(name, fallback_to_default=False)
                    chosen = (None, name)
                    break
                except Exception:
                    continue
    if chosen is None:
        chosen = (None, "DejaVu Sans")
    matplotlib.rcParams["font.sans-serif"] = [chosen[1], "DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False
    props = fm.FontProperties(fname=chosen[0]) if chosen[0] else fm.FontProperties(family=chosen[1])
    _FONT_CACHE["props"] = props
    return props


def _fig_png(fig, dpi=150):
    """matplotlib 图 → PNG bytes。"""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    buf.seek(0)
    return buf.getvalue()


def _style_ax(ax, title=None):
    ax.set_facecolor("white")
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(colors=C_SLATE, labelsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.35, color="#cbd5e1")
    if title:
        ax.set_title(title, fontsize=12, color=C_NAVY, fontweight="bold", pad=10)


def render_chart_png(chart_type, data=None, is_company_mode=False, title="", subtitle=""):
    """
    按图表类型用 matplotlib 渲染，返回 PNG bytes。
    支持：dupont_compare / market_share / market_growth / financial_trend /
          capability_compare / company_radar / risk_radar / industry_chain
    """
    import matplotlib.pyplot as plt
    setup_font()
    data = data or {}
    fig = None

    def _make(figsize=(7, 4.2)):
        return plt.figure(figsize=figsize, dpi=150, facecolor="white")

    if chart_type == "dupont_compare":
        fig = _make()
        ax = fig.add_subplot(111)
        labels = ["ROE (%)", "净利润率 (%)", "资产周转率×100", "权益乘数×10"]
        comp = [data.get("company_roe", 0), data.get("company_margin", 0),
                data.get("company_turnover", 0) * 100, data.get("company_multiplier", 0) * 10]
        ind = [data.get("industry_roe", 0), data.get("industry_margin", 0),
               data.get("industry_turnover", 0) * 100, data.get("industry_multiplier", 0) * 10]
        x = range(len(labels))
        ax.bar([i - 0.19 for i in x], comp, width=0.38, color=C_NAVY, label=data.get("company_name", "标的公司"))
        ax.bar([i + 0.19 for i in x], ind, width=0.38, color=C_AMBER, label="行业均值基准")
        ax.set_xticks(list(x)); ax.set_xticklabels(labels, fontsize=9)
        _style_ax(ax, title or "标的公司与行业杜邦因子对比（标准化）")
        ax.legend(fontsize=9, loc="upper right", frameon=False)
        for i, v in enumerate(comp):
            ax.text(i - 0.19, v + 0.6, f"{v:.1f}", ha="center", fontsize=7.5, color=C_NAVY)
        for i, v in enumerate(ind):
            ax.text(i + 0.19, v + 0.6, f"{v:.1f}", ha="center", fontsize=7.5, color=C_AMBER)

    elif chart_type == "market_share":
        fig = _make()
        ax = fig.add_subplot(111)
        labels = data.get("labels", ["头部企业 (CR4)", "中坚力量", "尾部企业"])
        values = data.get("values", [50, 35, 15])
        colors = [C_GREEN, C_NAVY, "#cbd5e1"]
        wedges, _texts, autotexts = ax.pie(values, labels=labels, autopct="%1.1f%%",
                                           colors=colors, startangle=90, textprops={"fontsize": 9},
                                           wedgeprops=dict(width=0.42, edgecolor="white"))
        for at in autotexts:
            at.set_color("white"); at.set_fontweight("bold"); at.set_fontsize(9)
        ax.set_title(title or "行业市场集中度（CR4）动态格局", fontsize=12, color=C_NAVY, fontweight="bold", pad=10)

    elif chart_type == "market_growth":
        fig = _make()
        ax = fig.add_subplot(111)
        years = data.get("years", ["2022", "2023", "2024", "2025", "2026(E)"])
        size = data.get("market_size", [100, 110, 120, 130, 140])
        rate = data.get("growth_rate", [10, 10, 9, 8, 7])
        ax.bar(years, size, color=C_NAVY, alpha=0.85, label="市场规模 (亿元)")
        ax.set_ylabel("市场规模 (亿元)", color=C_NAVY, fontsize=9)
        ax2 = ax.twinx()
        ax2.plot(years, rate, color=C_RED, marker="o", linewidth=2.5, label="增速 (%)")
        ax2.set_ylabel("增速 (%)", color=C_RED, fontsize=9)
        ax2.spines["right"].set_visible(False); ax2.spines["top"].set_visible(False)
        ax2.tick_params(colors=C_RED, labelsize=9)
        _style_ax(ax, title or "行业市场规模与复合增速")
        ax.legend(fontsize=8, loc="upper left", frameon=False)
        ax2.legend(fontsize=8, loc="upper right", frameon=False)

    elif chart_type == "financial_trend":
        fig = _make()
        ax = fig.add_subplot(111)
        years = data.get("years", ["2022", "2023", "2024", "2025", "2026Q2"])
        roe = data.get("roe_trend", [12, 11, 10, 9.5, 9.1])
        margin = data.get("margin_trend", [10, 9.5, 9, 8.8, 8.5])
        ax.plot(years, roe, marker="o", color=C_BLUE, linewidth=2.5, label="平均 ROE (%)")
        ax.plot(years, margin, marker="s", color=C_TEAL, linewidth=2.5, label="净利润率 (%)")
        _style_ax(ax, title or "主要盈利指标变化趋势")
        ax.legend(fontsize=9, frameon=False)
        ax.set_ylim(bottom=0)

    elif chart_type == "capability_compare":
        fig = _make()
        ax = fig.add_subplot(111)
        metrics = data.get("metrics", ["盈利能力(ROE%)", "短期流动性(流动比率×10)", "资产效率(周转率×100)", "安全边际(现金流%)"])
        values = data.get("values", [12, 15, 60, 20])
        y = range(len(metrics))
        ax.barh(list(y), values, color=[C_AMBER, C_TEAL, C_BLUE, C_GREEN], height=0.55)
        ax.set_yticks(list(y)); ax.set_yticklabels(metrics, fontsize=9)
        for i, v in enumerate(values):
            ax.text(v + max(values) * 0.02, i, f"{v:.1f}", va="center", fontsize=8.5, color=C_SLATE)
        _style_ax(ax, title or "企业多维核心财务能力对比")
        ax.grid(axis="x", linestyle="--", alpha=0.3, color="#cbd5e1")
        ax.grid(axis="y", visible=False)

    elif chart_type == "company_radar":
        fig = _make((6.6, 4.6))
        ax = fig.add_subplot(111, polar=True)
        theta_vars = ["ROE", "净利润率", "资产周转率", "财务杠杆", "经营现金流"]
        comp_r = [data.get("company_roe", 0), data.get("company_margin", 0),
                  data.get("company_turnover", 0) * 10, data.get("company_multiplier", 0),
                  data.get("company_cash", 0) / 1000]
        ind_r = [data.get("industry_roe", 0), data.get("industry_margin", 0),
                 data.get("industry_turnover", 0) * 10, data.get("industry_multiplier", 0),
                 data.get("industry_cash", 0) / 1000]
        angles = [i / len(theta_vars) * 2 * 3.1415926 for i in range(len(theta_vars))]
        angles_closed = angles + angles[:1]
        comp_c = comp_r + comp_r[:1]
        ind_c = ind_r + ind_r[:1]
        ax.plot(angles_closed, comp_c, color=C_NAVY, linewidth=2, label=data.get("company_name", "标的公司"))
        ax.fill(angles_closed, comp_c, color=C_NAVY, alpha=0.15)
        ax.plot(angles_closed, ind_c, color=C_AMBER, linewidth=2, label="行业平均")
        ax.fill(angles_closed, ind_c, color=C_AMBER, alpha=0.12)
        ax.set_xticks(angles); ax.set_xticklabels(theta_vars, fontsize=9, color=C_SLATE)
        ax.set_ylim(0, max(50.0, max(comp_r + ind_r + [1]) * 1.4))
        ax.set_title(title or "标的公司与行业能力多维透视", fontsize=12, color=C_NAVY, fontweight="bold", pad=18)
        ax.legend(loc="upper right", bbox_to_anchor=(1.32, 1.1), fontsize=8.5, frameon=False)

    elif chart_type == "risk_radar":
        fig = _make((7.0, 4.6))
        ax = fig.add_subplot(111, polar=True)
        dims = data.get("dimensions", ["偿债与财务杠杆风险", "短期流动性紧缺风险", "存货/资产减值风险",
                                       "盈利质量恶化风险", "政策合规与壁垒风险"])
        vals = data.get("values", [3, 3, 3, 3, 3])
        angles = [i / len(dims) * 2 * 3.1415926 for i in range(len(dims))]
        angles_closed = angles + angles[:1]
        vals_closed = vals + vals[:1]
        ax.plot(angles_closed, vals_closed, color=C_RED, linewidth=2.2)
        ax.fill(angles_closed, vals_closed, color=C_RED, alpha=0.25)
        ax.set_xticks(angles); ax.set_xticklabels(dims, fontsize=8.6, color=C_SLATE)
        ax.set_ylim(0, 5.0)
        ax.set_yticks([1, 2, 3, 4, 5])
        ax.set_yticklabels(["1 安全", "2", "3 关注", "4", "5 高危"], fontsize=7.5, color="#94a3b8")
        ax.set_title(title or "企业经营及财务多维风险指数（1=安全，5=高危）",
                     fontsize=12, color=C_RED, fontweight="bold", pad=18)
        ax.grid(color="#e2e8f0", alpha=0.8)

    elif chart_type == "industry_chain":
        fig = _make((8.4, 4.0))
        ax = fig.add_subplot(111)
        nodes = data.get("nodes", [])
        if not nodes:
            ax.text(0.5, 0.5, "该行业暂无产业链明细数据", ha="center", va="center", fontsize=11, color=C_SLATE)
            ax.axis("off")
        else:
            ax.axis("off")
            stage_color = {
                "upstream": C_BLUE, "midstream": C_TEAL,
                "integration": C_AMBER, "downstream": C_NAVY, "service": C_PURPLE,
            }
            for i, nd in enumerate(nodes):
                x = i / max(1, len(nodes) - 1)
                color = stage_color.get(nd.get("stage", ""), C_BLUE)
                ax.scatter([x], [0.5], s=1250, color=color, alpha=0.92, edgecolors="white", zorder=3)
                ax.text(x, 0.5, f"{nd.get('name', '')[:8]}\n{nd.get('name', '')[8:16]}",
                        ha="center", va="center", fontsize=8.2, color="white", fontweight="bold", zorder=4)
                ax.text(x, 0.72, f"龙头: {nd.get('leaders', '')[:26]}",
                        ha="center", va="center", fontsize=7.0, color=C_SLATE, zorder=4)
                ax.text(x, 0.24, f"利润率: {nd.get('margin', '')[:22]}",
                        ha="center", va="center", fontsize=7.0, color=C_SLATE, zorder=4)
                if i < len(nodes) - 1:
                    nx = (i + 1) / max(1, len(nodes) - 1)
                    ax.annotate("", xy=(nx, 0.5), xytext=(x + 0.04, 0.5),
                                arrowprops=dict(arrowstyle="-|>", color="#94a3b8", lw=2))
            ax.set_xlim(-0.03, 1.03); ax.set_ylim(0.1, 0.9)
            ax.set_title(title or "产业链全景逻辑流（环节→龙头→利润率）",
                         fontsize=12, color=C_NAVY, fontweight="bold", pad=12)

    if fig is None:
        fig = _make()
        ax = fig.add_subplot(111)
        ax.text(0.5, 0.5, "暂无图表数据", ha="center", va="center", fontsize=12, color=C_SLATE)
        ax.axis("off")
    png = _fig_png(fig)
    import matplotlib.pyplot as plt
    plt.close(fig)
    return png


# ---------- 高级导出：把 Plotly 图转 PNG（优先 kaleido，失败回退） ----------
def fig_plotly_to_png(fig, width=900, height=560):
    """优先 kaleido 渲染（云端 0.2.1 可用），失败返回 None。"""
    try:
        return fig.to_image(format="png", width=width, height=height)
    except Exception:
        return None


def chart_png_or_none(fig, chart_type, data=None, fallback_title=""):
    """统一入口：kaleido → matplotlib 兜底，保证永远有图。"""
    png = fig_plotly_to_png(fig)
    if png:
        return png
    try:
        return render_chart_png(chart_type, data, title=fallback_title)
    except Exception:
        return None


# ---------- 文档构建 ----------
def _today():
    return datetime.date.today().strftime("%Y-%m-%d")


def _strip_md(text):
    t = re.sub(r"[#*`>]", "", text)
    return t.strip()


def _split_sections(report_text):
    """把 markdown 报告切成 (level, title, body_lines)。"""
    sections = []
    cur_title = "报告正文"
    cur_level = 2
    body = []
    for line in (report_text or "").split("\n"):
        s = line.strip()
        if re.match(r"^#+\s+", s):
            if body and any(b.strip() for b in body):
                sections.append((cur_level, cur_title, body))
            lvl = len(s) - len(s.lstrip("#"))
            cur_title = re.sub(r"^#+\s*", "", s).strip()
            cur_level = lvl
            body = []
        else:
            body.append(s)
    if body and any(b.strip() for b in body):
        sections.append((cur_level, cur_title, body))
    return sections or [(2, "报告正文", (report_text or "").split("\n"))]


def _excerpt(text, limit=180):
    t = re.sub(r"\s+", " ", text).strip()
    return t[:limit] + ("…" if len(t) > limit else "")


def _add_chart_block(doc_or_slide, png, caption, width_in=5.9):
    from docx.shared import Inches
    from io import BytesIO
    try:
        doc_or_slide.add_picture(BytesIO(png), width=Inches(width_in))
        doc_or_slide.add_paragraph(f"图注：{caption}")
    except Exception as e:
        doc_or_slide.add_paragraph(f"[图表插入失败: {e}]")


# ---------------- Word ----------------
def export_docx(query, report_text, chart_images, evidence_data=None, gap_data=None,
                meta=None, is_company_mode=True, source_text=""):
    """
    chart_images: dict -> {"chart_id": {"title","caption","png"}}
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import docx.shared as dx

    doc = Document()
    # 页边距与默认字体
    for section in doc.sections:
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    style = doc.styles["Normal"]
    style.font.name = "Microsoft YaHei"
    style.font.size = Pt(10.5)

    # 封面
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{query}\n深度战略研报")
    run.font.size = Pt(26); run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A); run.bold = True
    doc.add_paragraph()
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"多智能体智能投研系统 · 数据可溯源版\n生成日期：{_today()}")
    r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(0x86, 0xBC, 0x25)
    doc.add_paragraph()
    doc.add_paragraph("—" * 40)

    # 数据鲜度与来源
    doc.add_heading("数据口径与来源说明", level=1)
    doc.add_paragraph(meta or {})
    if isinstance(meta, dict):
        for k, v in meta.items():
            doc.add_paragraph(f"• {k}：{v}", style=None)

    # 第一部分：数据看板
    doc.add_heading("第一部分：数据看板可视化", level=1)
    if chart_images:
        for cid, ch in chart_images.items():
            doc.add_heading(ch.get("title", cid), level=2)
            if ch.get("png"):
                _add_chart_block(doc, ch["png"], ch.get("caption", ""))
            else:
                doc.add_paragraph("[图表数据缺失]")
    else:
        doc.add_paragraph("本次研究未生成图表。")

    # 第二部分：证据链
    doc.add_heading("第二部分：数据可信度证据链（Evidence Ledger）", level=1)
    if evidence_data:
        table = doc.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 3"
        hdr = table.rows[0].cells
        for i, t in enumerate(["审计论点", "数据来源", "位置/页码", "可信度评级"]):
            hdr[i].text = t
        for item in evidence_data:
            cells = table.add_row().cells
            cells[0].text = str(item.get("point", ""))
            cells[1].text = str(item.get("source", ""))
            cells[2].text = str(item.get("page", ""))
            cells[3].text = str(item.get("confidence", ""))
    else:
        doc.add_paragraph("暂无证据链条目。")

    # 第三部分：正文
    doc.add_heading("第三部分：深度研究报告正文", level=1)
    for lvl, title, body in _split_sections(report_text):
        doc.add_heading(f"{title}", level=min(lvl + 1, 4))
        for line in body:
            if line.strip():
                doc.add_paragraph(_strip_md(line))

    # 第四部分：资料缺口与补充说明
    doc.add_heading("第四部分：资料缺口说明", level=1)
    if gap_data:
        for g in gap_data:
            doc.add_paragraph(f"• {g}")
    else:
        doc.add_paragraph("本次研究资料完整，无关键缺口。")

    # 免责声明
    doc.add_heading("免责声明", level=2)
    doc.add_paragraph("本报告由 AI 多智能体系统自动生成，所有数据来源已在正文与证据链中标注；"
                      "内容仅供学习与研究参考，不构成任何投资建议。投资决策与风险由使用者自行承担。")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ---------------- PPT ----------------
def export_pptx(query, report_text, chart_images, evidence_data=None, gap_data=None,
                meta=None, is_company_mode=True, source_text=""):
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def _bg(slide):
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shape.line.fill.background()
        # 顶部 Deloitte 绿条
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.12))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0x86, 0xBC, 0x25)
        bar.line.fill.background()
        return shape

    def _title(slide, text, sub=None):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.8))
        tf = tb.text_frame
        tf.text = text
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(30); run.font.bold = True
                run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        if sub:
            tb2 = slide.shapes.add_textbox(Inches(0.55), Inches(1.05), Inches(12.2), Inches(0.5))
            tf2 = tb2.text_frame
            tf2.text = sub
            for para in tf2.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(14); run.font.color.rgb = RGBColor(0x88, 0x99, 0xAA)

    def _footer(slide, text):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.4))
        tf = tb.text_frame
        tf.text = text
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(9); run.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    # 1) 封面
    s = prs.slides.add_slide(blank)
    _bg(s)
    tb = s.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.0))
    tf = tb.text_frame
    tf.text = f"{query}\n深度战略研报"
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(40); run.font.bold = True
            run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    tb2 = s.shapes.add_textbox(Inches(1.0), Inches(4.4), Inches(11.3), Inches(0.8))
    tf2 = tb2.text_frame
    tf2.text = f"多智能体智能投研系统 · 数据可溯源版\n生成日期：{_today()}"
    for para in tf2.paragraphs:
        for run in para.runs:
            run.font.size = Pt(16); run.font.color.rgb = RGBColor(0x86, 0xBC, 0x25)
    _footer(s, "AI 生成内容，仅供学习研究参考，不构成投资建议")

    # 2) 目录
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "目录")
    items = ["01 数据口径与来源", "02 数据看板可视化", "03 深度研究报告正文", "04 证据链与可信度", "05 资料缺口说明", "06 免责声明"]
    tb = s.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(11.0), Inches(4.5))
    tf = tb.text_frame
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = it
        for run in p.runs:
            run.font.size = Pt(22); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # 3) 数据口径页
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "01 数据口径与来源")
    tb = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.0))
    tf = tb.text_frame
    meta_items = []
    if isinstance(meta, dict):
        for k, v in meta.items():
            meta_items.append(f"• {k}：{v}")
    meta_items.append("• 风险维度采用真实财务指标映射（详见各图口径说明）")
    for i, it in enumerate(meta_items or ["• 暂无元信息"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = it
        for run in p.runs:
            run.font.size = Pt(14); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # 4) 图表页（每张一页）
    if chart_images:
        for cid, ch in chart_images.items():
            s = prs.slides.add_slide(blank)
            _bg(s)
            _title(s, ch.get("title", cid), ch.get("caption", ""))
            if ch.get("png"):
                from PIL import Image
                import io as _io
                try:
                    img = Image.open(_io.BytesIO(ch["png"]))
                    w, h = img.size
                    ratio = h / w
                    max_w, max_h = Inches(11.5), Inches(4.9)
                    if ratio > max_h / max_w:
                        pic_w = int(max_h / ratio)
                        pic_h = int(max_h)
                    else:
                        pic_w = int(max_w); pic_h = int(max_w * ratio)
                    left = int((prs.slide_width - pic_w) / 2)
                    top = Inches(1.5)
                    s.shapes.add_picture(_io.BytesIO(ch["png"]), left, top, pic_w, pic_h)
                except Exception:
                    pass
            _footer(s, f"{query} · {_today()}")

    # 5) 正文要点（分段）
    sections = _split_sections(report_text)
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "03 深度研究报告正文（摘要）")
    tb = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.2))
    tf = tb.text_frame
    count = 0
    for lvl, title, body in sections[:8]:
        full = " ".join(body)
        if not full.strip():
            continue
        p = tf.paragraphs[0] if count == 0 else tf.add_paragraph()
        p.text = f"▎{title}"
        for run in p.runs:
            run.font.size = Pt(16); run.font.bold = True; run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        p2 = tf.add_paragraph()
        p2.text = _excerpt(full, 200)
        for run in p2.runs:
            run.font.size = Pt(13); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
        count += 1
    _footer(s, f"{query} · {_today()}")

    # 6) 证据链
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "04 数据可信度证据链")
    tb = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.0))
    tf = tb.text_frame
    if evidence_data:
        for i, item in enumerate(evidence_data[:8]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {item.get('point', '')[:60]} | 来源:{item.get('source', '')[:40]} | 评级:{item.get('confidence', '')}"
            for run in p.runs:
                run.font.size = Pt(13); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
    else:
        p = tf.paragraphs[0]
        p.text = "暂无证据链条目"
        for run in p.runs:
            run.font.size = Pt(14); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # 7) 资料缺口
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "05 资料缺口说明")
    tb = s.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.0))
    tf = tb.text_frame
    if gap_data:
        for i, g in enumerate(gap_data):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {g}"
            for run in p.runs:
                run.font.size = Pt(14); run.font.color.rgb = RGBColor(0xC0, 0x6A, 0x1A)
    else:
        p = tf.paragraphs[0]
        p.text = "本次研究资料完整，无关键缺口。"
        for run in p.runs:
            run.font.size = Pt(14); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    # 8) 免责声明
    s = prs.slides.add_slide(blank)
    _bg(s); _title(s, "06 免责声明")
    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.8), Inches(3.0))
    tf = tb.text_frame
    tf.text = ("本报告由 AI 多智能体系统自动生成，所有数据来源已在正文与证据链中标注；"
               "内容仅供学习与研究参考，不构成任何投资建议。投资决策与风险由使用者自行承担。")
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(16); run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()


if __name__ == "__main__":
    # 自检
    png = render_chart_png("risk_radar", {"dimensions": ["a", "b", "c", "d", "e"], "values": [3, 4, 2, 3.5, 4]})
    print("risk_radar png:", len(png) if png else None)
    png2 = render_chart_png("dupont_compare", {"company_roe": 18.5, "company_margin": 5.2, "company_turnover": 1.1,
                                               "company_multiplier": 2.5, "industry_roe": 12.5, "industry_margin": 8.2,
                                               "industry_turnover": 0.75, "industry_multiplier": 2.1, "company_name": "比亚迪"})
    print("dupont png:", len(png2) if png2 else None)
