# -*- coding: utf-8 -*-
"""Generic transaction, comps and scenario-model utilities.

The original app is strongest at A-share industry research.  This module adds
an instrument-agnostic workbench: users can paste or upload the same columns
for a private company, a US-listed issuer, or a Chinese company, while every
calculated result keeps an explicit source/as-of field.
"""
from __future__ import annotations

import io
import math
import re
from datetime import date
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

try:
    import pandas as pd
except Exception:  # pragma: no cover - the Streamlit app already depends on pandas
    pd = None

try:
    from chart_utils import apply_plotly_theme, PLOTLY_CONFIG
except Exception:  # pragma: no cover - keep pure model functions importable
    apply_plotly_theme = None
    PLOTLY_CONFIG = {}


SCENARIOS: Dict[str, Dict[str, float]] = {
    "Bear": {"revenue_growth": 0.15, "fcf_margin": 0.16, "wacc": 0.11, "terminal_growth": 0.02},
    "Base": {"revenue_growth": 0.25, "fcf_margin": 0.22, "wacc": 0.09, "terminal_growth": 0.025},
    "Bull": {"revenue_growth": 0.35, "fcf_margin": 0.28, "wacc": 0.08, "terminal_growth": 0.03},
}

SOURCE_MODES = {
    "中国大陆可用优先": "优先使用本地库、用户上传、巨潮/交易所/东方财富/AkShare；外网失败时不阻断流程。",
    "全球公开披露": "优先使用公司 IR、SEC/交易所披露和公开行情；必须记录报告期、抓取时间与来源 URL。",
    "用户上传优先": "适合 Snowflake、Meta、私有公司和受网络限制环境；上传 10-K/10-Q/年报、共识和估值表后计算。",
}


def source_quality(source: Any, as_of: Any = "") -> str:
    """Classify provenance without implying that a URL or upload is audited.

    A source label is evidence of where a number came from; it is not proof that
    the number is correct.  The UI and exported workbook use this conservative
    vocabulary so placeholder values cannot be mistaken for market facts.
    """
    text = str(source or "").strip()
    lower = text.lower()
    if not text or any(token in text for token in ("待核实", "请填入", "占位", "示例")):
        return "未核实/占位"
    if lower.startswith(("http://", "https://")):
        return "URL 已提供，需核验"
    if as_of:
        return "已提供来源与日期，需核验"
    return "已提供来源，日期缺失"


def source_caption(st: Any, source: Any, as_of: Any = "", quality: Any = "") -> None:
    """Render the small, consistent provenance line required below every chart."""
    text = str(source or "未提供来源，待核实").strip()
    parts = [f"来源：{text}"]
    if as_of:
        parts.append(f"截至：{as_of}")
    if quality:
        parts.append(f"数据质量：{quality}")
    st.caption("；".join(parts))


def render_workbench_chart(st: Any, fig: Any, **kwargs: Any) -> bool:
    """Render a chart defensively so one malformed optional series does not stop a page."""
    try:
        if apply_plotly_theme is not None:
            apply_plotly_theme(fig)
        kwargs.setdefault("config", PLOTLY_CONFIG)
        st.plotly_chart(fig, **kwargs)
        return True
    except Exception as exc:
        st.warning(f"图表渲染失败，已保留表格数据：{str(exc)[:120]}")
        return False


def _num(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or (isinstance(value, str) and not value.strip()):
            return default
        number = float(str(value).replace(",", "").replace("%", "").strip())
        return number if math.isfinite(number) else default
    except (TypeError, ValueError, OverflowError):
        return default


def _pct(value: Any, default: float = 0.0) -> float:
    number = _num(value, default)
    return number / 100.0 if abs(number) > 1.0 else number


def _safe_div(numerator: float, denominator: float) -> Optional[float]:
    return numerator / denominator if denominator not in (0, None) else None


def normalize_comps(rows: Any) -> List[Dict[str, Any]]:
    """Normalize common comps column aliases into a stable schema."""
    if pd is not None and isinstance(rows, pd.DataFrame):
        records = rows.to_dict("records")
    elif isinstance(rows, Mapping):
        records = [dict(rows)]
    else:
        records = list(rows or [])
    aliases = {
        "company": ["company", "公司", "name", "名称", "issuer", "标的"],
        "ticker": ["ticker", "代码", "symbol", "证券代码"],
        "revenue": ["revenue", "营收", "营业收入", "ARR", "arr"],
        "gross_profit": ["gross_profit", "gross profit", "毛利", "毛利润"],
        "ev": ["ev", "enterprise_value", "企业价值", "EV"],
        "equity_value": ["equity_value", "market_cap", "市值", "股权价值"],
        "net_debt": ["net_debt", "净债务", "debt_minus_cash"],
        "growth": ["growth", "revenue_growth", "营收增速", "收入增速"],
        "profitability": ["profitability", "ebitda_margin", "fcf_margin", "利润率", "EBITDA利润率"],
        "source": ["source", "来源", "source_url", "数据来源"],
        "as_of": ["as_of", "截至日期", "report_date", "报告期"],
    }

    def pick(row: Mapping[str, Any], key: str) -> Any:
        lowered = {str(k).strip().lower(): v for k, v in row.items()}
        for alias in aliases[key]:
            if alias.lower() in lowered:
                return lowered[alias.lower()]
        return ""

    output: List[Dict[str, Any]] = []
    for row in records:
        company = str(pick(row, "company") or "未命名公司").strip()
        revenue = _num(pick(row, "revenue"))
        gross_profit = _num(pick(row, "gross_profit"))
        equity_value = _num(pick(row, "equity_value"))
        net_debt = _num(pick(row, "net_debt"))
        ev = _num(pick(row, "ev"), equity_value + net_debt)
        growth = _pct(pick(row, "growth"))
        profitability = _pct(pick(row, "profitability"))
        output.append({
            "company": company,
            "ticker": str(pick(row, "ticker") or "").strip(),
            "revenue": revenue,
            "gross_profit": gross_profit,
            "ev": ev,
            "equity_value": equity_value,
            "net_debt": net_debt,
            "growth": growth,
            "profitability": profitability,
            "ev_revenue": _safe_div(ev, revenue),
            "ev_gross_profit": _safe_div(ev, gross_profit),
            "rule_of_40": growth * 100 + profitability * 100,
            "source": str(pick(row, "source") or "用户输入/待核实").strip(),
            "as_of": str(pick(row, "as_of") or "").strip(),
        })
        output[-1]["source_quality"] = source_quality(output[-1]["source"], output[-1]["as_of"])
    return output


def comps_summary(rows: Any) -> Dict[str, Any]:
    peers = normalize_comps(rows)
    def median(key: str) -> Optional[float]:
        values = sorted(float(x[key]) for x in peers if x.get(key) is not None)
        if not values:
            return None
        mid = len(values) // 2
        return values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2
    return {
        "peers": peers,
        "median_ev_revenue": median("ev_revenue"),
        "median_ev_gross_profit": median("ev_gross_profit"),
        "median_rule_of_40": median("rule_of_40"),
        "source_count": len({x.get("source") for x in peers if x.get("source")}),
        "unverified_count": sum(1 for x in peers if x.get("source_quality") == "未核实/占位"),
    }


def _scenario_assumptions(overrides: Optional[Mapping[str, Mapping[str, Any]]] = None) -> Dict[str, Dict[str, float]]:
    result = {name: dict(values) for name, values in SCENARIOS.items()}
    for name, values in (overrides or {}).items():
        if name not in result:
            result[name] = {}
        for key, value in values.items():
            if key in ("revenue_growth", "fcf_margin", "wacc", "terminal_growth"):
                result[name][key] = _pct(value) if key != "wacc" and key != "terminal_growth" else _pct(value)
    return result


def build_dcf_scenarios(
    base_revenue: Any,
    years: int = 5,
    net_debt: Any = 0,
    overrides: Optional[Mapping[str, Mapping[str, Any]]] = None,
) -> Dict[str, Any]:
    """Build Bear/Base/Bull DCF outputs and WACC/terminal-growth sensitivities."""
    years = max(1, min(int(_num(years, 5)), 15))
    revenue0 = max(0.0, _num(base_revenue))
    net_debt_value = _num(net_debt)
    cases = _scenario_assumptions(overrides)
    results: Dict[str, Any] = {}
    validation_notes: List[str] = []
    for name, assumptions in cases.items():
        g = _pct(assumptions.get("revenue_growth"))
        fcf_margin = _pct(assumptions.get("fcf_margin"))
        wacc = _pct(assumptions.get("wacc"))
        terminal_growth = _pct(assumptions.get("terminal_growth"))
        original_wacc = wacc
        if wacc <= terminal_growth:
            wacc = terminal_growth + 0.005
            validation_notes.append(f"{name}: WACC 不高于终值增长率，已自动调整为至少高 0.5 个百分点以避免终值分母为零。")
        if original_wacc < 0:
            validation_notes.append(f"{name}: WACC 输入为负值，已按终值约束修正。")
        revenues, fcfs, pv_fcfs = [], [], []
        for t in range(1, years + 1):
            revenue = revenue0 * ((1 + g) ** t)
            fcf = revenue * fcf_margin
            revenues.append(revenue)
            fcfs.append(fcf)
            pv_fcfs.append(fcf / ((1 + wacc) ** t))
        terminal_value = fcfs[-1] * (1 + terminal_growth) / (wacc - terminal_growth) if fcfs else 0.0
        pv_terminal = terminal_value / ((1 + wacc) ** years) if fcfs else 0.0
        enterprise_value = sum(pv_fcfs) + pv_terminal
        results[name] = {
            "assumptions": {"revenue_growth": g, "fcf_margin": fcf_margin, "wacc": wacc, "terminal_growth": terminal_growth},
            "validation_notes": [n for n in validation_notes if n.startswith(f"{name}:")],
            "years": list(range(1, years + 1)),
            "revenue": revenues,
            "fcf": fcfs,
            "pv_fcf": pv_fcfs,
            "terminal_value": terminal_value,
            "pv_terminal": pv_terminal,
            "enterprise_value": enterprise_value,
            "equity_value": enterprise_value - net_debt_value,
            "terminal_value_share": pv_terminal / enterprise_value if enterprise_value else None,
        }

        wacc_grid = [max(0.01, wacc + delta) for delta in (-0.02, -0.01, 0, 0.01, 0.02)]
        tg_grid = [max(-0.02, terminal_growth + delta) for delta in (-0.01, -0.005, 0, 0.005, 0.01)]
        sensitivity = []
        for sensitivity_wacc in wacc_grid:
            row = []
            for sensitivity_tg in tg_grid:
                if sensitivity_wacc <= sensitivity_tg:
                    row.append(None)
                    continue
                pv = sum(fcfs[t - 1] / ((1 + sensitivity_wacc) ** t) for t in range(1, years + 1))
                tv = fcfs[-1] * (1 + sensitivity_tg) / (sensitivity_wacc - sensitivity_tg)
                row.append(pv + tv / ((1 + sensitivity_wacc) ** years))
            sensitivity.append(row)
        results[name]["wacc_grid"] = wacc_grid
        results[name]["terminal_growth_grid"] = tg_grid
        results[name]["sensitivity"] = sensitivity
    return {"base_revenue": revenue0, "years": years, "net_debt": net_debt_value, "scenarios": results,
            "validation_notes": validation_notes}


def build_lbo_model(
    entry_revenue: Any,
    entry_multiple: Any = 8.0,
    exit_multiple: Any = 9.0,
    debt_percent: Any = 0.45,
    revenue_growth: Any = 0.20,
    debt_paydown_percent: Any = 0.60,
    hold_years: int = 5,
) -> Dict[str, float]:
    """Transparent first-pass LBO return bridge for screening, not a full debt schedule."""
    revenue = max(0.0, _num(entry_revenue))
    entry_multiple = max(0.0, _num(entry_multiple, 8.0))
    exit_multiple = max(0.0, _num(exit_multiple, 9.0))
    debt_percent = max(0.0, min(0.95, _pct(debt_percent)))
    growth = _pct(revenue_growth)
    paydown = max(0.0, min(1.0, _pct(debt_paydown_percent)))
    years = max(1, min(15, int(_num(hold_years, 5))))
    entry_ev = revenue * entry_multiple
    initial_debt = entry_ev * debt_percent
    initial_equity = entry_ev - initial_debt
    exit_revenue = revenue * ((1 + growth) ** years)
    exit_ev = exit_revenue * exit_multiple
    debt_remaining = initial_debt * (1 - paydown)
    exit_equity = exit_ev - debt_remaining
    moic = _safe_div(exit_equity, initial_equity) or 0.0
    irr = moic ** (1 / years) - 1 if moic > 0 else -1.0
    return {
        "entry_ev": entry_ev, "initial_debt": initial_debt, "initial_equity": initial_equity,
        "exit_revenue": exit_revenue, "exit_ev": exit_ev, "debt_remaining": debt_remaining,
        "exit_equity": exit_equity, "moic": moic, "irr": irr, "hold_years": years,
    }


def build_transaction_materials(company: str, arr: Any, description: str = "企业数据平台", sector: str = "企业软件") -> Dict[str, Any]:
    """Return editable, anonymized teaser/CIM/buyer-list scaffolding."""
    name = company.strip() or "Project Atlas"
    arr_value = _num(arr)
    teaser = (
        "# Project Atlas | Anonymous Teaser\n\n"
        f"An enterprise data platform in the {sector} sector with approximately ${arr_value:,.0f}M ARR.\n\n"
        "*Draft input only: ARR and all operating claims must be validated against company records before circulation.*\n\n"
        "## Investment highlights\n- Recurring revenue model with enterprise customers\n"
        "- Scalable data infrastructure and workflow product\n- Additional detail available under NDA\n\n"
        "## Process\nQualified parties may request the full information memorandum and management presentation."
    )
    cim_sections = [
        "Executive Summary", "Company Overview", "Market and Competitive Landscape", "Product and Technology",
        "Customers and Go-to-Market", "Historical Financials", "Operating Metrics", "Forecast and Key Assumptions",
        "Valuation", "Transaction Considerations", "Risk Factors", "Appendices and Sources",
    ]
    buyer_types = ["Strategic cloud/data platform", "Enterprise software consolidator", "Growth equity", "Technology-focused buyout fund", "Infrastructure software investor"]
    buyers = [{"buyer_type": b, "status": "候选池，需人工筛选", "rationale": f"与{name}的{description}存在潜在协同"} for b in buyer_types]
    return {
        "teaser": teaser,
        "cim_sections": cim_sections,
        "buyers": buyers,
        "arr": arr_value,
        "data_quality": "用户输入草稿，尚未由审计/公司披露核验",
        "source": "用户输入/待核实",
        "as_of": "",
    }


def _write_df(writer: Any, sheet: str, frame: Any, title: str = "") -> None:
    if frame is None or pd is None:
        return
    frame.to_excel(writer, sheet_name=sheet, index=False, startrow=2 if title else 0)
    workbook = writer.book
    worksheet = writer.sheets[sheet]
    header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#0F2A5C", "border": 0})
    title_fmt = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#0F2A5C"})
    if title:
        worksheet.write(0, 0, title, title_fmt)
    for col, value in enumerate(frame.columns):
        worksheet.write(2 if title else 0, col, value, header)
    worksheet.freeze_panes(3 if title else 1, 0)
    worksheet.hide_gridlines(2)
    for idx, column in enumerate(frame.columns):
        worksheet.set_column(idx, idx, min(32, max(12, len(str(column)) + 2)))


def workbook_bytes(bundle: Mapping[str, Any]) -> Optional[bytes]:
    """Create a compact multi-sheet xlsx for the workbench download button."""
    if pd is None:
        return None
    try:
        import xlsxwriter  # noqa: F401
    except Exception:
        return None
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        try:
            writer.book.set_calc_mode("auto")
        except Exception:
            pass
        summary = pd.DataFrame(bundle.get("summary", []))
        _write_df(writer, "Summary", summary, "Investment Workbench Summary")
        if bundle.get("comps"):
            _write_df(writer, "Comps", pd.DataFrame(bundle["comps"]), "Comparable Companies")
        dcf = bundle.get("dcf") or {}
        if dcf.get("scenarios"):
            assumption_rows = []
            for scenario, result in dcf["scenarios"].items():
                a = result["assumptions"]
                assumption_rows.append({
                    "Scenario": scenario,
                    "Revenue CAGR": a["revenue_growth"],
                    "FCF Margin": a["fcf_margin"],
                    "WACC": a["wacc"],
                    "Terminal Growth": a["terminal_growth"],
                    "Base Revenue": dcf.get("base_revenue", 0),
                    "Net Debt": dcf.get("net_debt", 0),
                    "Years": dcf.get("years", 5),
                })
            _write_df(writer, "Assumptions", pd.DataFrame(assumption_rows), "DCF Assumptions (editable)")
            build_sheet = writer.book.add_worksheet("DCF Build")
            writer.sheets["DCF Build"] = build_sheet
            header_fmt = writer.book.add_format({"bold": True, "font_color": "white", "bg_color": "#0F2A5C"})
            number_fmt = writer.book.add_format({"num_format": "#,##0.0"})
            factor_fmt = writer.book.add_format({"num_format": "0.000x"})
            headers = ["Scenario", "Year", "Revenue", "FCF", "Discount Factor", "PV FCF", "Terminal Value", "PV Terminal", "Enterprise Value"]
            for col, header in enumerate(headers):
                build_sheet.write(0, col, header, header_fmt)
            row = 1
            scenario_names = list(dcf["scenarios"].keys())
            for s_idx, scenario in enumerate(scenario_names):
                assumption_row = 4 + s_idx  # title row + blank row + header row + zero-based data row
                result = dcf["scenarios"][scenario]
                for year in range(1, dcf.get("years", 5) + 1):
                    excel_row = row + 1  # XlsxWriter indexes rows from zero; formulas from one.
                    build_sheet.write(row, 0, scenario)
                    build_sheet.write(row, 1, year)
                    if year == 1:
                        revenue_formula = f"=Assumptions!$F${assumption_row}*(1+Assumptions!$B${assumption_row})"
                    else:
                        revenue_formula = f"=C{row}*(1+Assumptions!$B${assumption_row})"
                    build_sheet.write_formula(row, 2, revenue_formula, number_fmt, result["revenue"][year - 1])
                    build_sheet.write_formula(row, 3, f"=C{excel_row}*Assumptions!$C${assumption_row}", number_fmt, result["fcf"][year - 1])
                    build_sheet.write_formula(row, 4, f"=1/(1+Assumptions!$D${assumption_row})^B{excel_row}", factor_fmt, 1 / ((1 + result["assumptions"]["wacc"]) ** year))
                    build_sheet.write_formula(row, 5, f"=D{excel_row}*E{excel_row}", number_fmt, result["pv_fcf"][year - 1])
                    if year == dcf.get("years", 5):
                        build_sheet.write_formula(row, 6, f"=D{excel_row}*(1+Assumptions!$E${assumption_row})/(Assumptions!$D${assumption_row}-Assumptions!$E${assumption_row})", number_fmt, result["terminal_value"])
                        build_sheet.write_formula(row, 7, f"=G{excel_row}*E{excel_row}", number_fmt, result["pv_terminal"])
                        first_row = row - dcf.get("years", 5) + 2
                        build_sheet.write_formula(row, 8, f"=SUM(F{first_row}:F{excel_row})+H{excel_row}", number_fmt, result["enterprise_value"])
                    row += 1
            build_sheet.freeze_panes(1, 2)
            build_sheet.hide_gridlines(2)
            build_sheet.set_column("A:A", 12)
            build_sheet.set_column("B:B", 8)
            build_sheet.set_column("C:I", 16)
        rows = []
        for scenario, result in (dcf.get("scenarios") or {}).items():
            rows.append({"Scenario": scenario, "Revenue CAGR": result["assumptions"]["revenue_growth"], "FCF Margin": result["assumptions"]["fcf_margin"], "WACC": result["assumptions"]["wacc"], "Terminal Growth": result["assumptions"]["terminal_growth"], "Enterprise Value": result["enterprise_value"], "Equity Value": result["equity_value"], "TV Share": result["terminal_value_share"]})
        if rows:
            _write_df(writer, "DCF Scenarios", pd.DataFrame(rows), "DCF Scenario Summary")
        if dcf.get("scenarios"):
            base = dcf["scenarios"].get("Base") or next(iter(dcf["scenarios"].values()))
            sens = pd.DataFrame(base.get("sensitivity", []), index=[f"WACC {x:.1%}" for x in base.get("wacc_grid", [])], columns=[f"TGR {x:.1%}" for x in base.get("terminal_growth_grid", [])])
            sens.index.name = "WACC / TGR"
            _write_df(writer, "Sensitivity", sens.reset_index(), "Base Case EV Sensitivity")
        if bundle.get("lbo"):
            _write_df(writer, "LBO", pd.DataFrame([bundle["lbo"]]), "LBO Screening Bridge")
        # Do not stamp today's date on placeholders: it falsely suggests that
        # a live source was retrieved today.  A real as-of date must come from
        # the uploaded filing/quote row or an explicitly entered source.
        sources = pd.DataFrame(bundle.get("sources", [{
            "source": "用户输入/待核实",
            "as_of": "",
            "data_quality": "未核实/占位",
            "note": "无外部数据；请替换为可审计来源并填写报告期/抓取日期",
        }]))
        _write_df(writer, "Sources", sources, "Sources and Data Quality")
        for _format in writer.book.formats:
            _format.set_font_name("LXGW WenKai")
    output.seek(0)
    return output.getvalue()


def _default_peer_frame() -> Any:
    """Illustrative rows only; the UI labels them as data to replace."""
    if pd is None:
        return []
    return pd.DataFrame([
        {"company": "Snowflake", "ticker": "SNOW", "revenue": 0, "gross_profit": 0, "ev": 0, "growth": 0, "profitability": 0, "source": "示例占位：请填入最新 10-K/行情数据", "as_of": ""},
        {"company": "Datadog", "ticker": "DDOG", "revenue": 0, "gross_profit": 0, "ev": 0, "growth": 0, "profitability": 0, "source": "示例占位：请填入最新 10-K/行情数据", "as_of": ""},
        {"company": "MongoDB", "ticker": "MDB", "revenue": 0, "gross_profit": 0, "ev": 0, "growth": 0, "profitability": 0, "source": "示例占位：请填入最新 10-K/行情数据", "as_of": ""},
        {"company": "Confluent", "ticker": "CFLT", "revenue": 0, "gross_profit": 0, "ev": 0, "growth": 0, "profitability": 0, "source": "示例占位：请填入最新 10-K/行情数据", "as_of": ""},
    ])


def _read_upload(uploaded_file: Any) -> Any:
    if uploaded_file is None or pd is None:
        return None
    try:
        name = (uploaded_file.name or "").lower()
        raw = uploaded_file.getvalue()
        if name.endswith(".csv"):
            return pd.read_csv(io.BytesIO(raw))
        if name.endswith((".xlsx", ".xls")):
            return pd.read_excel(io.BytesIO(raw))
    except Exception:
        return None
    return None


def _fmt_pct(value: Any) -> str:
    return "—" if value is None else f"{_num(value) * 100:.1f}%"


def source_guidance(mode: str) -> str:
    return SOURCE_MODES.get(mode, SOURCE_MODES["用户上传优先"])


def render_workbench(st: Any, page: str, chain_builder: Any = None, news_fetcher: Any = None,
                     webpage_reader: Any = None, pdf_exporter: Any = None) -> None:
    """Render the additive workbench pages without changing the legacy flow."""
    st.title("分析工作台")
    st.caption("通用输入、来源可追溯、模型可切换。未填入官方或用户数据的字段会明确标记为待核实。")
    source_mode = st.selectbox("数据来源模式", list(SOURCE_MODES), index=0, key=f"wb_source_mode_{page}")
    st.info(source_guidance(source_mode))
    with st.expander("数据真实性与可审计性说明", expanded=False):
        st.markdown(
            "- **公开披露**：应来自公司 IR、SEC/交易所、巨潮或正式公告，并填写报告期/抓取日期。\n"
            "- **用户上传**：系统只负责清洗和计算，不验证文件是否完整或经审计。\n"
            "- **模型假设**：DCF/LBO 的增长率、利润率、WACC 和终值是输入假设，不是外部事实。\n"
            "- **示例占位**：没有上传或可追溯来源时，页面只显示占位值并标注“未核实/占位”，不会当作真实行情。"
        )

    if page == "交易材料与估值":
        st.subheader("交易材料与估值包")
        c1, c2, c3 = st.columns(3)
        with c1:
            company = st.text_input("项目名称", "DataFlow Analytics", key="wb_deal_company")
            arr = st.number_input("ARR / 当前收入（$M）", min_value=0.0, value=200.0, step=5.0, key="wb_deal_arr")
        with c2:
            sector = st.text_input("行业", "企业数据平台", key="wb_deal_sector")
            entry_multiple = st.number_input("LBO 入场 EV/Revenue", min_value=0.0, value=8.0, step=0.5, key="wb_entry_multiple")
        with c3:
            net_debt = st.number_input("净债务（$M）", value=0.0, step=5.0, key="wb_deal_debt")
            exit_multiple = st.number_input("LBO 出场 EV/Revenue", min_value=0.0, value=9.0, step=0.5, key="wb_exit_multiple")
        uploaded = st.file_uploader("上传可比公司 CSV/XLSX（可选）", type=["csv", "xlsx", "xls"], key="wb_deal_upload")
        uploaded_frame = _read_upload(uploaded)
        frame = uploaded_frame if uploaded_frame is not None else _default_peer_frame()
        if pd is not None and frame is not None:
            if uploaded_frame is None:
                st.warning("可比公司表当前为示例占位数据；估值区间只能作为模型演示，不能直接用于交易定价。")
            edited = st.data_editor(frame, num_rows="dynamic", use_container_width=True, key="wb_deal_comps_editor")
            comps = comps_summary(edited)
            peers = comps["peers"]
            if peers:
                st.dataframe(pd.DataFrame(peers), use_container_width=True)
                st.caption(f"来源状态：{comps['source_count']} 个来源；未核实/占位 {comps['unverified_count']} 行。")
        else:
            comps = comps_summary([])
            peers = []
        dcf = build_dcf_scenarios(arr, net_debt=net_debt)
        lbo = build_lbo_model(arr, entry_multiple, exit_multiple)
        materials = build_transaction_materials(company, arr, sector=sector)
        st.subheader("估值区间与回报筛选")
        summary_rows = [{"场景": name, "企业价值（$M）": round(result["enterprise_value"], 1), "股权价值（$M）": round(result["equity_value"], 1), "营收 CAGR": _fmt_pct(result["assumptions"]["revenue_growth"])} for name, result in dcf["scenarios"].items()]
        summary_rows.append({"场景": "LBO 筛选", "企业价值（$M）": round(lbo["entry_ev"], 1), "股权价值（$M）": round(lbo["exit_equity"], 1), "营收 CAGR": _fmt_pct(0.2)})
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)
        st.caption("估值输入来源：项目名称、ARR、净债务和 LBO 参数为用户输入或默认演示值；正式交易材料需用经审计财务、可比公司行情和报告期来源替换。")
        if st.button("生成交易材料草稿", type="primary", key="wb_build_deal"):
            st.markdown(materials["teaser"])
            st.markdown("### CIM 章节目录")
            st.write("；".join(materials["cim_sections"]))
            st.markdown("### 潜在买方候选池（需人工筛选与合规核验）")
            st.dataframe(pd.DataFrame(materials["buyers"]), use_container_width=True)
            cim_text = "# Confidential Information Memorandum\n\n" + "\n".join(f"## {idx}. {section}\n\n待补充经核实数据、来源与管理层访谈记录。" for idx, section in enumerate(materials["cim_sections"], 1))
            st.download_button("下载匿名 One-pager", materials["teaser"].encode("utf-8"), "anonymous_teaser.md", "text/markdown", key="wb_teaser_md")
            st.download_button("下载 CIM 草稿目录", cim_text.encode("utf-8"), "cim_draft.md", "text/markdown", key="wb_cim_md")
            if pdf_exporter:
                try:
                    import report_export as _rex
                    _valuation_png = _rex.render_chart_png(
                        "capability_compare",
                        {"metrics": ["Bear EV", "Base EV", "Bull EV", "LBO 入场 EV"],
                         "values": [float(dcf["scenarios"][s]["enterprise_value"]) for s in ["Bear", "Base", "Bull"]] + [float(lbo["entry_ev"])]},
                        title="交易材料估值区间（$M）",
                    )
                    _meta = {"研究主体": company, "报告类型": "交易材料 / 匿名 One-pager / CIM 草稿",
                             "分析周期": "用户输入与模型情景", "数据口径": "用户输入、模型假设和待核验可比公司数据"}
                    _pdf = pdf_exporter(company, materials["teaser"] + "\n\n" + cim_text,
                                        {"valuation": {"title": "交易材料估值区间", "caption": "Bear/Base/Bull DCF 与 LBO 入场企业价值；正式材料需替换为可审计来源。", "png": _valuation_png, "source": "用户输入与模型假设；可比公司数据需逐行核验"}},
                                        gap_data=["可比公司行情和经审计财务仍需补充", "管理层访谈和买方名单需人工核验"], meta=_meta)
                    st.download_button("下载匿名 One-pager / CIM PDF", _pdf, "transaction_materials.pdf", "application/pdf", key="wb_deal_pdf")
                except Exception as _pdf_err:
                    st.caption(f"交易材料 PDF 暂不可用：{str(_pdf_err)[:100]}")
        bundle = {"summary": summary_rows, "comps": peers, "dcf": dcf, "lbo": lbo, "sources": [{"source": source_mode, "as_of": str(date.today()), "note": "正式交易材料需替换为经审计和可引用来源"}]}
        xlsx = workbook_bytes(bundle)
        if xlsx:
            st.download_button("下载估值与交易材料工作簿", xlsx, "transaction_workbench.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="wb_deal_xlsx")
        else:
            st.info("当前环境未安装 XlsxWriter，部署时请按 requirements.txt 安装后启用 Excel 下载。")
        return

    if page == "可比公司分析":
        st.subheader("可比公司分析")
        st.caption("适用于 SNOW、Meta、私有 SaaS 或任意行业。EV/Revenue、EV/GP 和 Rule of 40 使用同一套输入口径。")
        source_url = st.text_input("公开来源 URL（可选）", placeholder="公司 IR、交易所、SEC 或公告页面", key="wb_comps_url")
        if webpage_reader and source_url and st.button("读取网页表格与图表数据", key="wb_comps_read_url"):
            with st.spinner("正在读取公开网页…"):
                page_data = webpage_reader(source_url)
            st.session_state["wb_comps_web_data"] = page_data
        page_data = st.session_state.get("wb_comps_web_data")
        if page_data:
            st.caption(f"网页状态：{page_data.get('note', '')}；来源：{page_data.get('source_url') or source_url or '未提供'}；抓取时间：{page_data.get('fetched_at', '未记录')}")
            if page_data.get("tables"):
                first = page_data["tables"][0]
                headers = first.get("headers") or []
                rows = first.get("rows") or []
                if headers and rows and pd is not None:
                    st.session_state["wb_comps_web_frame"] = pd.DataFrame(rows, columns=headers[:len(rows[0])])
        uploaded = st.file_uploader("上传公司财务与估值数据", type=["csv", "xlsx", "xls"], key="wb_comps_upload")
        uploaded_frame = _read_upload(uploaded)
        frame = uploaded_frame if uploaded_frame is not None else st.session_state.get("wb_comps_web_frame", _default_peer_frame())
        if pd is not None and frame is not None:
            if uploaded_frame is None and st.session_state.get("wb_comps_web_frame") is None:
                st.warning("当前使用的是示例占位行；EV/Revenue、EV/GP 和 Rule of 40 在填入最新财务与行情数据前不代表真实公司数据。")
            edited = st.data_editor(frame, num_rows="dynamic", use_container_width=True, key="wb_comps_editor")
            summary = comps_summary(edited)
            peer_frame = pd.DataFrame(summary["peers"])
            st.dataframe(peer_frame, use_container_width=True)
            st.caption(f"来源状态：{summary['source_count']} 个来源；未核实/占位 {summary['unverified_count']} 行。请在正式分析前逐行核对来源与报告期。")
            m1, m2, m3 = st.columns(3)
            m1.metric("中位数 EV/Revenue", "—" if summary["median_ev_revenue"] is None else f"{summary['median_ev_revenue']:.1f}x")
            m2.metric("中位数 EV/GP", "—" if summary["median_ev_gross_profit"] is None else f"{summary['median_ev_gross_profit']:.1f}x")
            m3.metric("中位数 Rule of 40", "—" if summary["median_rule_of_40"] is None else f"{summary['median_rule_of_40']:.1f}")
            try:
                import plotly.graph_objects as go
                valid = [x for x in summary["peers"] if x.get("ev_revenue") is not None]
                if valid:
                    fig = go.Figure(go.Bar(x=[x["company"] for x in valid], y=[x["ev_revenue"] for x in valid], text=[f"{x['ev_revenue']:.1f}x" for x in valid], textposition="outside"))
                    fig.update_layout(title="EV / Revenue", height=330, template="plotly_white")
                    render_workbench_chart(st, fig, use_container_width=True, key="wb_comps_chart")
                    st.caption("来源：当前上传文件、公开网页表格或手工输入；未提供可核验来源的行标记为示例/待核实。")
            except Exception as exc:
                st.info(f"图表暂不可用：{str(exc)[:100]}")
        return

    if page == "DCF 情景模型":
        st.subheader("Bear / Base / Bull DCF")
        c1, c2, c3 = st.columns(3)
        with c1:
            revenue = st.number_input("基期收入（$M）", min_value=0.0, value=1000.0, step=50.0, key="wb_dcf_revenue")
            net_debt = st.number_input("净债务（$M）", value=0.0, step=50.0, key="wb_dcf_net_debt")
        with c2:
            base_growth = st.number_input("Base 未来 5 年 CAGR", min_value=-0.9, max_value=3.0, value=0.25, step=0.01, format="%.2f", key="wb_dcf_growth")
            years = st.number_input("预测年限", min_value=1, max_value=15, value=5, step=1, key="wb_dcf_years")
        with c3:
            base_wacc = st.number_input("Base WACC", min_value=0.01, max_value=0.5, value=0.09, step=0.005, format="%.3f", key="wb_dcf_wacc")
            base_tg = st.number_input("Base 终值增长率", min_value=-0.05, max_value=0.08, value=0.025, step=0.005, format="%.3f", key="wb_dcf_tg")
        with st.expander("编辑 Bear / Base / Bull 假设", expanded=False):
            a1, a2, a3 = st.columns(3)
            with a1:
                bear_growth = st.number_input("Bear 收入 CAGR", min_value=-0.9, max_value=3.0, value=0.15, step=0.01, format="%.2f", key="wb_dcf_bear_growth")
                bear_margin = st.number_input("Bear FCF Margin", min_value=-1.0, max_value=1.0, value=0.16, step=0.01, format="%.2f", key="wb_dcf_bear_margin")
            with a2:
                base_margin = st.number_input("Base FCF Margin", min_value=-1.0, max_value=1.0, value=0.22, step=0.01, format="%.2f", key="wb_dcf_base_margin")
                bull_growth = st.number_input("Bull 收入 CAGR", min_value=-0.9, max_value=3.0, value=0.35, step=0.01, format="%.2f", key="wb_dcf_bull_growth")
            with a3:
                bull_margin = st.number_input("Bull FCF Margin", min_value=-1.0, max_value=1.0, value=0.28, step=0.01, format="%.2f", key="wb_dcf_bull_margin")
                st.caption("比例输入 0.25 = 25%。")
        overrides = {
            "Bear": {"revenue_growth": bear_growth, "fcf_margin": bear_margin},
            "Base": {"revenue_growth": base_growth, "fcf_margin": base_margin, "wacc": base_wacc, "terminal_growth": base_tg},
            "Bull": {"revenue_growth": bull_growth, "fcf_margin": bull_margin},
        }
        dcf = build_dcf_scenarios(revenue, years, net_debt, overrides)
        rows = []
        for name, result in dcf["scenarios"].items():
            rows.append({"Scenario": name, "Revenue CAGR": _fmt_pct(result["assumptions"]["revenue_growth"]), "FCF Margin": _fmt_pct(result["assumptions"]["fcf_margin"]), "WACC": _fmt_pct(result["assumptions"]["wacc"]), "Terminal Growth": _fmt_pct(result["assumptions"]["terminal_growth"]), "Enterprise Value": round(result["enterprise_value"], 1), "Equity Value": round(result["equity_value"], 1), "TV Share": _fmt_pct(result["terminal_value_share"])} )
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
        for note in dcf.get("validation_notes", []):
            st.warning(note)
        try:
            import plotly.graph_objects as go
            fig = go.Figure(go.Bar(x=list(dcf["scenarios"].keys()), y=[v["enterprise_value"] for v in dcf["scenarios"].values()], text=[f"{v['enterprise_value']:,.0f}" for v in dcf["scenarios"].values()], textposition="outside"))
            fig.update_layout(title="DCF 企业价值（$M）", height=330, template="plotly_white")
            render_workbench_chart(st, fig, use_container_width=True, key="wb_dcf_ev_chart")
            st.caption("来源：用户输入与模型默认假设（Base CAGR 默认 25%）；不代表外部已核验财务数据。")
            base = dcf["scenarios"]["Base"]
            heat = go.Figure(go.Heatmap(z=base["sensitivity"], x=[f"{x:.1%}" for x in base["terminal_growth_grid"]], y=[f"{x:.1%}" for x in base["wacc_grid"]], colorscale="Blues", hovertemplate="WACC %{y}, TGR %{x}<br>EV %{z:,.1f}<extra></extra>"))
            heat.update_layout(title="Base Case WACC / 终值增长率敏感性", xaxis_title="终值增长率", yaxis_title="WACC", height=380, template="plotly_white")
            render_workbench_chart(st, heat, use_container_width=True, key="wb_dcf_heatmap")
            st.caption("来源：DCF 模型计算；WACC 与终值增长率网格由当前输入生成，属于情景分析而非市场报价。")
        except Exception as exc:
            st.info(f"图表暂不可用：{str(exc)[:100]}")
        xlsx = workbook_bytes({"summary": rows, "dcf": dcf, "sources": [{"source": source_mode, "as_of": str(date.today()), "note": "Base CAGR 默认 25%，请替换为公司/行业底稿"}]})
        if xlsx:
            st.download_button("下载 DCF Excel", xlsx, "scenario_dcf.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="wb_dcf_xlsx")
        return

    if page == "财报更新":
        st.subheader("通用季度财报更新")
        st.caption("支持 Meta、Snowflake 或任意公司；数值由用户上传/输入，系统不会把缺失项伪装成最新事实。")
        company = st.text_input("公司", "Meta Platforms", key="wb_earnings_company")
        quarter = st.text_input("报告期", "最新季度", key="wb_earnings_period")
        earnings_url = st.text_input("公司 IR / SEC / 交易所来源 URL（可选）", key="wb_earnings_url")
        if webpage_reader and earnings_url and st.button("读取来源正文与图表", key="wb_earnings_read_url"):
            with st.spinner("正在读取财报来源…"):
                st.session_state["wb_earnings_web_data"] = webpage_reader(earnings_url)
        earnings_web = st.session_state.get("wb_earnings_web_data")
        if earnings_web:
            st.caption(f"来源读取状态：{earnings_web.get('note', '')}；来源：{earnings_web.get('source_url') or earnings_url or '未提供'}；抓取时间：{earnings_web.get('fetched_at', '未记录')}")
            if earnings_web.get("text"):
                st.text_area("来源摘录", earnings_web.get("text", "")[:5000], height=180, key="wb_earnings_web_excerpt", disabled=True)
        labels = ["广告收入", "街头预期", "Reality Labs 亏损", "Capex 指引", "DAU", "MAU", "用户 engagement"]
        values = {}
        cols = st.columns(2)
        for idx, label in enumerate(labels):
            with cols[idx % 2]:
                values[label] = st.text_input(label, key=f"wb_earn_{idx}")
        if st.button("生成 Earnings Update 大纲", type="primary", key="wb_earnings_generate"):
            display_values = {k: (v or "待补充官方披露/一致预期") for k, v in values.items()}
            report_text = f"""# {company} {quarter} Earnings Update

## 1. Executive Summary
本报告根据当前输入和上传资料生成。所有缺失数据均保留为待补充项。

## 2. Revenue and advertising performance
- 广告收入：{display_values['广告收入']}
- 街头预期：{display_values['街头预期']}
- 差异与驱动：需要补充公司新闻稿、10-Q/10-K 和一致预期来源。

## 3. Reality Labs loss trend
- Reality Labs 亏损：{display_values['Reality Labs 亏损']}
- 趋势判断：需要至少三个季度的同口径数据后再判断改善或恶化。

## 4. AI investment, capex and returns
- Capex 指引：{display_values['Capex 指引']}
- AI 投资回报：需要将资本开支、折旧、收入增量和利润率变化放在同一期间比较。

## 5. Users and engagement
- DAU：{display_values['DAU']}
- MAU：{display_values['MAU']}
- Engagement：{display_values['用户 engagement']}

## 6. Guidance and estimates
需要补充管理层指引、市场一致预期、汇率和报告期定义。

## 7. Risks and catalysts
风险包括资本开支回报不确定、广告需求变化、监管和产品投入执行偏差；催化剂包括广告定价、用户时长和 AI 产品商业化改善。

## 8. Valuation and conclusion
本节需接入市值、净现金、收入和利润预测后完成估值。当前版本是结构化草稿，不构成最新事实判断。
"""
            st.markdown(report_text)
            st.dataframe(pd.DataFrame([{"Metric": k, "Value": v, "Source": source_mode} for k, v in display_values.items()]), use_container_width=True)
            st.download_button("下载 Earnings Update 草稿", report_text.encode("utf-8"), f"{company}_{quarter}_earnings_update.md", "text/markdown", key="wb_earnings_md")
            st.info("要生成可发布的 8–10 页报告，还需要上传公司 10-Q/10-K、业绩新闻稿、电话会纪要和一致预期数据。")
        return

    if page == "产业链地图":
        st.subheader("产业链地图")
        industry = st.text_input("行业或公司", "半导体", key="wb_chain_industry")
        chain = chain_builder(industry) if chain_builder else {"nodes": [], "note": "未加载产业链数据"}
        nodes = chain.get("nodes", []) or []
        st.caption(chain.get("note", ""))
        if not nodes:
            st.info("暂无节点。系统保留通用五环节框架，请补充行业名称或上传行业资料。")
            return
        view3d, view2d = st.tabs(["3D 全景", "2D 可读视图"])
        with view3d:
            try:
                import plotly.graph_objects as go
                colors = {"upstream": "#2563eb", "midstream": "#0d9488", "integration": "#f59e0b", "downstream": "#1e3a8a", "service": "#8b5cf6"}
                fig = go.Figure(go.Scatter3d(x=[_num(n.get("x")) for n in nodes], y=[_num(n.get("y")) for n in nodes], z=[_num(n.get("z")) for n in nodes], mode="lines+markers+text", text=[str(n.get("name", ""))[:16] for n in nodes], textposition="top center", marker={"size": 13, "color": [colors.get(n.get("stage"), "#94a3b8") for n in nodes]}, line={"color": "#cbd5e1", "width": 5}, customdata=[[n.get("stage", ""), n.get("business", ""), n.get("leaders", ""), n.get("margin", ""), n.get("source", "")] for n in nodes], hovertemplate="<b>%{text}</b><br>阶段：%{customdata[0]}<br>业务：%{customdata[1]}<br>主要企业：%{customdata[2]}<br>利润率：%{customdata[3]}<br>来源：%{customdata[4]}<extra></extra>"))
                fig.update_layout(title="产业链节点与利润率", height=560, template="plotly_white", scene={"xaxis_title": "环节顺序", "yaxis_title": "阶段层级", "zaxis_title": "利润率中值 (%)", "camera": {"eye": {"x": 1.5, "y": 1.3, "z": 0.9}}})
                render_workbench_chart(st, fig, use_container_width=True, key="wb_chain_3d")
                st.caption("来源：产业链节点 source 字段（公开资料/用户上传/数据库）；利润率缺失时按 0 显示并标注数据质量。")
            except Exception as exc:
                st.warning(f"3D 图渲染失败，已保留 2D 视图：{str(exc)[:100]}")
        with view2d:
            table = pd.DataFrame([{k: n.get(k, "") for k in ["name", "stage", "business", "leaders", "margin", "source"]} for n in nodes])
            st.dataframe(table, use_container_width=True)
            try:
                import plotly.graph_objects as go
                fig2 = go.Figure(go.Bar(x=[str(n.get("name", ""))[:14] for n in nodes], y=[_num(n.get("z")) for n in nodes], marker_color="#1F5FA8"))
                fig2.update_layout(title="环节利润率中值（缺失数据显示为 0，请查看来源）", height=320, template="plotly_white")
                render_workbench_chart(st, fig2, use_container_width=True, key="wb_chain_2d")
                st.caption("来源：同上产业链节点 source 字段；数值为节点利润率中值，缺失项不视为真实 0。")
            except Exception:
                pass
        if news_fetcher and st.button("刷新产业链公开动态", key="wb_chain_news"):
            with st.spinner("正在按多来源顺序获取公开信息…"):
                news = news_fetcher.fetch_news_bundle(industry=industry, keyword=industry, limit_each=8)
            st.session_state["wb_chain_news_result"] = news
        news = st.session_state.get("wb_chain_news_result")
        if news:
            st.write(f"数据源状态：{'成功' if news.get('ok') else '未成功'}；{news.get('note', '')}")
            if news.get("source_status"):
                st.dataframe(pd.DataFrame(news["source_status"]), use_container_width=True)
            if news.get("items"):
                st.dataframe(pd.DataFrame(news["items"]), use_container_width=True)
        return
