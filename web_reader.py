# -*- coding: utf-8 -*-
"""
网页读取器（Web Page Reader）
=============================
让投研系统能够读取网页上的「文字 + 表格 + 内嵌图表数据」：

1. 文字内容：正文段落提取（BeautifulSoup）
2. 表格数据：HTML <table> -> 结构化 rows
3. 内嵌图表数据：从 <script> 中解析常见图表库的数据结构
   - ECharts: option.series[].data / xAxis.data
   - Chart.js: chart.data.datasets / labels
   - Highcharts: series.data / categories
   - 通用 JSON 数据块：window.__INITIAL_STATE__ / chartData 等
4. 图片说明：<img> 的 alt 与 src（图表截图的可读替代信息）

对外统一接口：
    read_webpage(url) -> {"ok", "title", "text", "tables": [...], "charts": [...], "images": [...], "note"}
"""
import json
import re

import requests

UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}


def _safe_json(text):
    """尽力解析一段 JS 对象/JSON（容忍前后空格、注释等）。"""
    if not text:
        return None
    t = text.strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    # 去掉尾随分号、前后括号包裹
    t = t.rstrip(";").strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    return None


def _extract_echarts(script_text):
    """从 ECharts 初始化代码里提取 series 数据。"""
    charts = []
    # 形如 series: [{name:.., data:[..]}] / xAxis: {data:[..]}
    for m in re.finditer(r"series\s*:\s*(\[[^\]]*\])", script_text, re.S):
        chunk = m.group(1)
        names = re.findall(r"name\s*:\s*['\"]([^'\"]+)['\"]", chunk)
        datas = re.findall(r"data\s*:\s*(\[[^\]]*\])", chunk, re.S)
        if datas:
            for i, d in enumerate(datas):
                parsed = _safe_json(d)
                if isinstance(parsed, list) and parsed:
                    charts.append({
                        "type": "ECharts series",
                        "name": names[i] if i < len(names) else f"系列{i + 1}",
                        "data": parsed[:60],
                    })
    # xAxis 类别
    for m in re.finditer(r"xAxis\s*:\s*\{[^}]*?data\s*:\s*(\[[^\]]*\])", script_text, re.S):
        parsed = _safe_json(m.group(1))
        if isinstance(parsed, list) and parsed:
            charts.append({"type": "ECharts xAxis", "name": "横轴类别",
                           "data": [str(x) for x in parsed[:60]]})
    return charts


def _extract_chartjs(script_text):
    """从 Chart.js 代码里提取 datasets。"""
    charts = []
    for m in re.finditer(r"data\s*:\s*\{[^{}]*labels\s*:\s*(\[[^\]]*\])\s*,[^{}]*datasets\s*:\s*(\[[^\]]*\])", script_text, re.S):
        labels = _safe_json(m.group(1))
        datasets_raw = _safe_json(m.group(2))
        if isinstance(labels, list) and isinstance(datasets_raw, list):
            for i, ds in enumerate(datasets_raw[:6]):
                if isinstance(ds, dict) and "data" in ds:
                    charts.append({
                        "type": "Chart.js dataset",
                        "name": ds.get("label", f"系列{i + 1}"),
                        "labels": [str(x) for x in labels[:60]],
                        "data": ds.get("data", [])[:60],
                    })
    return charts


def _extract_highcharts(script_text):
    charts = []
    for m in re.finditer(r"categories\s*:\s*(\[[^\]]*\])", script_text, re.S):
        parsed = _safe_json(m.group(1))
        if isinstance(parsed, list) and parsed:
            charts.append({"type": "Highcharts categories", "name": "横轴类别",
                           "data": [str(x) for x in parsed[:60]]})
    for m in re.finditer(r"name\s*:\s*['\"]([^'\"]+)['\"]\s*,\s*data\s*:\s*(\[[^\]]*\])", script_text, re.S):
        parsed = _safe_json(m.group(2))
        if isinstance(parsed, list) and parsed:
            charts.append({"type": "Highcharts series", "name": m.group(1),
                           "data": parsed[:60]})
    return charts


def _extract_json_state(script_text):
    """window.__INITIAL_STATE__ / chartData / seriesData 等全局数据块。"""
    charts = []
    for key in ["__INITIAL_STATE__", "chartData", "seriesData", "optionData",
                "echartsOption", "initData", "statData"]:
        m = re.search(r"(?:window\.)?%s\s*=\s*(.+?);" % key, script_text, re.S)
        if m:
            parsed = _safe_json(m.group(1))
            if isinstance(parsed, (dict, list)):
                charts.append({"type": f"JSON {key}", "name": key, "data": parsed})
    return charts


def extract_charts_from_html(html):
    """从 HTML 的所有 <script> 中提取图表数据。"""
    charts = []
    for m in re.finditer(r"<script[^>]*>(.*?)</script>", html, re.S):
        script = m.group(1)
        if not script or len(script) < 20:
            continue
        charts += _extract_echarts(script)
        charts += _extract_chartjs(script)
        charts += _extract_highcharts(script)
        charts += _extract_json_state(script)
        if len(charts) > 20:
            break
    # 去重（按 name+data 前 10 项）
    out, seen = [], set()
    for c in charts:
        sig = str(c.get("name", "")) + str(c.get("data", []))[:10]
        if sig in seen:
            continue
        seen.add(sig)
        out.append(c)
    return out[:20]


def extract_tables_from_html(soup):
    """提取 HTML 表格为 [{"caption", "headers", "rows"}]。"""
    tables = []
    for i, tbl in enumerate(soup.select("table")[:10]):
        headers, rows = [], []
        # thead / th
        ths = tbl.select("th")
        if ths:
            headers = [th.get_text(strip=True)[:30] for th in ths[:12]]
        for tr in tbl.select("tr")[:30]:
            cells = [td.get_text(strip=True)[:40] for td in tr.select("td, th")]
            if cells and any(cells):
                rows.append(cells[:12])
        if rows:
            caption = ""
            cap = tbl.find_previous("caption") or tbl.select_one("caption")
            if cap:
                caption = cap.get_text(strip=True)[:60]
            tables.append({"caption": caption or f"表格{i + 1}", "headers": headers, "rows": rows[:20]})
    return tables


def read_webpage(url, timeout=15):
    """读取网页：文字 + 表格 + 图表数据 + 图片说明。"""
    try:
        from bs4 import BeautifulSoup
        r = requests.get(url, headers=UA, timeout=timeout)
        if r.status_code != 200:
            return {"ok": False, "title": "", "text": "",
                    "tables": [], "charts": [], "images": [],
                    "note": f"HTTP {r.status_code}"}
        r.encoding = r.apparent_encoding or r.encoding
        html = r.text
        soup = BeautifulSoup(html, "lxml")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        title = (soup.title.get_text(strip=True) if soup.title else "")[:120]
        # 正文：p / li / h1-h6 / td
        blocks = []
        for sel in ["p", "h1", "h2", "h3", "h4", "li", "blockquote"]:
            for node in soup.select(sel):
                t = node.get_text(" ", strip=True)
                if len(t) >= 12:
                    blocks.append(t)
        text = "\n".join(blocks)[:8000]
        tables = extract_tables_from_html(soup)
        charts = extract_charts_from_html(html)
        images = [{"alt": (img.get("alt") or "")[:80], "src": img.get("src", "")[:150]}
                  for img in soup.select("img")[:12] if img.get("src") or img.get("alt")]
        return {"ok": True, "title": title, "text": text,
                "tables": tables, "charts": charts, "images": images,
                "note": f"已读取 {len(blocks)} 段正文 / {len(tables)} 张表格 / {len(charts)} 组图表数据"}
    except Exception as e:
        return {"ok": False, "title": "", "text": "",
                "tables": [], "charts": [], "images": [],
                "note": f"网页读取失败: {str(e)[:150]}"}


def format_page_summary(page, max_text=1200):
    """把网页读取结果格式化为报告可引用的文本。"""
    parts = [f"【网页】{page.get('title', '')}（{page.get('note', '')}）"]
    txt = (page.get("text") or "").strip()
    if txt:
        parts.append("正文摘录：\n" + txt[:max_text])
    for tbl in (page.get("tables") or [])[:3]:
        head = " | ".join(tbl.get("headers", [])) or "（无表头）"
        parts.append(f"表格《{tbl.get('caption', '')}》：表头 {head}")
        for row in tbl.get("rows", [])[:5]:
            parts.append("  " + " | ".join(row[:8]))
    for ch in (page.get("charts") or [])[:5]:
        name = ch.get("name", "")
        data = ch.get("data", [])
        parts.append(f"图表数据[{ch.get('type', '')} · {name}]：{str(data)[:300]}")
    return "\n".join(parts)[:5000]


if __name__ == "__main__":
    import sys
    url = sys.argv[1] if len(sys.argv) > 1 else "https://www.eastmoney.com/"
    page = read_webpage(url)
    print("ok:", page["ok"], "|", page["note"])
    print("title:", page["title"])
    if page["charts"]:
        print("charts:", len(page["charts"]))
        for c in page["charts"][:3]:
            print("  -", c["type"], c["name"], str(c["data"])[:120])
