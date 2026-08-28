# -*- coding: utf-8 -*-
"""
新闻与公告抓取器（News & Announcement Fetcher）
==============================================
为投研系统提供「实时新闻 / 公告 / 网络搜索」能力（无需额外付费 API Key）：

1. 个股公告：东方财富公告大全（akshare stock_individual_notice_report / stock_notice_report）
2. 全市场 7x24 快讯：东方财富 np-listapi getFastNewsList（实时财经快讯）
3. 宏观/财经新闻：财新网（akshare stock_news_main_cx）、新闻联播文字稿（news_cctv）
4. 网络搜索：搜狗新闻搜索（免费、无需 Key，带来源与链接）
5. 个股代码查询：akshare stock_info_a_code_name（名称 -> 代码，带缓存）

所有接口均带超时与容错，单个数据源失败不影响整体，返回统一结构：
    {"ok": bool, "items": [{"title","date","source","url","summary"}], "note": str}

说明：本模块是「网站自己的搜索能力」的一部分，与 RAG 知识库互补——
RAG 覆盖已上传文档，本模块覆盖实时公开信息。
"""
import json
import os
import time
import datetime

import requests

UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
    "Accept": "text/html,application/json,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# 名称 -> 代码缓存（避免每次研究都重新拉取 5500+ 家）
_CODE_CACHE = {"ts": 0, "df": None}
_CODE_CACHE_TTL = 6 * 3600

# 磁盘缓存目录（全市场公告 6000+ 行，首次拉取约 25 秒，缓存 12 小时）
_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
_MARKET_ANN_CACHE = os.path.join(_CACHE_DIR, "news_market_announcements.json")
_MARKET_ANN_TTL = 12 * 3600


def _cache_get(path, ttl):
    try:
        if os.path.exists(path) and time.time() - os.path.getmtime(path) < ttl:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None


def _cache_set(path, obj):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False)
    except Exception:
        pass


def _today(days=0):
    return (datetime.date.today() + datetime.timedelta(days=days)).strftime("%Y%m%d")


def _norm(items):
    """统一为 [{title,date,source,url,summary}]，去重、去空。"""
    out, seen = [], set()
    for it in items:
        title = str(it.get("title") or "").strip()
        if not title or title in seen:
            continue
        seen.add(title)
        out.append({
            "title": title[:200],
            "date": str(it.get("date") or "")[:20],
            "source": str(it.get("source") or "")[:40],
            "url": str(it.get("url") or ""),
            "summary": str(it.get("summary") or "")[:400],
        })
    return out[:30]


# ---------------------------------------------------------------- 公告
def fetch_company_announcements(company_name="", stock_code="", days=30):
    """个股公告（东方财富公告大全）。company_name 或 stock_code 任一即可。"""
    try:
        import akshare as ak
        if not stock_code:
            stock_code = lookup_stock_code(company_name)
        if not stock_code:
            return {"ok": False, "items": [], "note": "未找到该公司的股票代码，无法抓取公告"}
        begin = _today(-days)
        end = _today()
        df = ak.stock_individual_notice_report(security=stock_code, symbol="全部",
                                               begin_date=begin, end_date=end)
        items = []
        for _, r in df.iterrows():
            items.append({
                "title": str(r.get("公告标题", "")),
                "date": str(r.get("公告日期", "")),
                "source": "东方财富·公告大全",
                "url": str(r.get("网址", "")),
                "summary": str(r.get("公告类型", "")),
            })
        return {"ok": True, "items": _norm(items),
                "note": f"东方财富公告大全 · {company_name or stock_code} 近{days}天"}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"公告抓取失败: {str(e)[:120]}"}


def fetch_market_announcements(days=1, keyword="", limit=20):
    """全市场公告（按日期，带 12 小时磁盘缓存），可按关键词过滤（行业/公司名）。"""
    try:
        cached = _cache_get(_MARKET_ANN_CACHE, _MARKET_ANN_TTL)
        if cached is None:
            import akshare as ak
            df = ak.stock_notice_report(symbol="全部", date=_today())
            if df is None or len(df) == 0:
                df = ak.stock_notice_report(symbol="全部", date=_today(-1))
            cached = df.to_dict("records")
            _cache_set(_MARKET_ANN_CACHE, cached)
        items = []
        for r in cached:
            title = str(r.get("公告标题", ""))
            name = str(r.get("名称", ""))
            if keyword and keyword not in title and keyword not in name:
                continue
            items.append({
                "title": title,
                "date": str(r.get("公告日期", "")),
                "source": "东方财富·全市场公告",
                "url": str(r.get("网址", "")),
                "summary": f"{name} · {r.get('公告类型', '')}",
            })
        return {"ok": True, "items": _norm(items)[:limit],
                "note": "东方财富全市场公告" + (f"· 关键词「{keyword}」" if keyword else "")}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"全市场公告抓取失败: {str(e)[:120]}"}


# ---------------------------------------------------------------- 快讯
def fetch_7x24_news(keyword="", pages=3, limit=15):
    """东方财富 7x24 实时快讯（多页）；可传关键词过滤，未命中时返回最新头条兜底。"""
    try:
        items = []
        url = "https://np-listapi.eastmoney.com/comm/web/getFastNewsList"
        for page in range(1, pages + 1):
            params = {"client": "web", "biz": "web_7x24", "fastColumn": "102",
                      "sortEnd": "", "pageSize": "50", "pageIndex": str(page),
                      "req_trace": "1"}
            r = requests.get(url, params=params, headers=UA, timeout=12)
            data = r.json()
            lst = ((data.get("data") or {}).get("fastNewsList")) or []
            if not lst:
                break
            for it in lst:
                title = str(it.get("title") or "").strip()
                if not title:
                    continue
                items.append({
                    "title": title,
                    "date": str(it.get("showTime") or "")[:16],
                    "source": "东方财富·7×24快讯",
                    "url": str(it.get("url") or ""),
                    "summary": str(it.get("summary") or "")[:300],
                })
        if keyword:
            hits = [it for it in items if keyword in it["title"]]
            if not hits:
                hits = [it for it in items if keyword in it.get("summary", "")]
            items = hits or items  # 未命中则返回最新头条
        return {"ok": True, "items": _norm(items)[:limit],
                "note": "东方财富 7×24 实时快讯" + (f"· 关键词「{keyword}」" if keyword else "")}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"快讯抓取失败: {str(e)[:120]}"}


# ---------------------------------------------------------------- 财经新闻
def fetch_caixin_news(limit=15):
    """财新网财经新闻（akshare）。"""
    try:
        import akshare as ak
        df = ak.stock_news_main_cx()
        items = []
        for _, r in df.iterrows():
            items.append({
                "title": str(r.get("summary", "")),
                "date": "",
                "source": "财新网",
                "url": str(r.get("url", "")),
                "summary": str(r.get("tag", "")),
            })
        return {"ok": True, "items": _norm(items)[:limit], "note": "财新网财经新闻"}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"财新新闻抓取失败: {str(e)[:120]}"}


def fetch_cctv_news(limit=15):
    """新闻联播文字稿（akshare）。"""
    try:
        import akshare as ak
        df = ak.news_cctv(date=_today())
        if df is None or len(df) == 0:
            df = ak.news_cctv(date=_today(-1))
        items = []
        for _, r in df.iterrows():
            items.append({
                "title": str(r.get("title", "")),
                "date": str(r.get("date", "")),
                "source": "新闻联播文字稿",
                "url": "",
                "summary": str(r.get("content", ""))[:200],
            })
        return {"ok": True, "items": _norm(items)[:limit], "note": "新闻联播文字稿"}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"新闻联播抓取失败: {str(e)[:120]}"}


# ---------------------------------------------------------------- 网络搜索
def search_web_news(keyword, limit=10):
    """搜狗新闻搜索（免费、无需 Key），返回标题/来源/链接/摘要。"""
    try:
        from bs4 import BeautifulSoup
        params = {"query": keyword, "mode": "1", "sourceid": "inttime"}
        r = requests.get("https://news.sogou.com/news", params=params,
                         headers=UA, timeout=15)
        soup = BeautifulSoup(r.text, "lxml")
        items = []
        # 兼容两种页面结构：.news-item 与 .vrwrap
        nodes = soup.select(".news-item") or soup.select(".vrwrap")
        for node in nodes[:limit]:
            h3 = node.select_one("h3 a") or node.select_one("h3")
            if h3 is None:
                continue
            title = h3.get_text(strip=True)
            link = h3.get("href", "") if h3.has_attr("href") else ""
            src = node.select_one(".news-from") or node.select_one(".news-source")
            source = src.get_text(strip=True) if src else "搜狗新闻"
            summary_node = (node.select_one(".news-summary") or node.select_one(".txt-info")
                            or node.select_one(".news-detail"))
            summary = summary_node.get_text(strip=True) if summary_node else ""
            # 搜狗跳转链接 -> 尽量还原真实地址
            if link.startswith("/link?") and "url=" in link:
                import urllib.parse
                q = urllib.parse.parse_qs(urllib.parse.urlparse(link).query)
                if q.get("url"):
                    link = q["url"][0]
            items.append({"title": title, "date": "", "source": source,
                          "url": link, "summary": summary[:200]})
        return {"ok": True, "items": _norm(items), "note": f"搜狗新闻搜索 · {keyword}"}
    except Exception as e:
        return {"ok": False, "items": [], "note": f"网络搜索失败: {str(e)[:120]}"}


# ---------------------------------------------------------------- 代码查询
def lookup_stock_code(company_name):
    """公司名称 -> 股票代码（akshare 全市场名单，带缓存）。"""
    if not company_name:
        return ""
    try:
        import akshare as ak
        now = time.time()
        if (_CODE_CACHE["df"] is None or now - _CODE_CACHE["ts"] > _CODE_CACHE_TTL):
            _CODE_CACHE["df"] = ak.stock_info_a_code_name()
            _CODE_CACHE["ts"] = now
        df = _CODE_CACHE["df"]
        hit = df[df["name"].astype(str).str.contains(company_name.strip(), na=False)]
        if not hit.empty:
            return str(hit.iloc[0]["code"]).zfill(6)
        # 去掉空格再匹配（如 万 科Ａ）
        hit2 = df[df["name"].astype(str).str.replace(" ", "").str.contains(
            company_name.strip().replace(" ", ""), na=False)]
        if not hit2.empty:
            return str(hit2.iloc[0]["code"]).zfill(6)
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------- 统一入口
def fetch_news_bundle(company_name="", industry="", keyword="", days=30, limit_each=8):
    """
    一键获取：个股公告 + 全市场公告(按行业/公司关键词过滤) + 7x24 快讯 + 网络搜索。
    返回 {"items": [...], "sources": [说明], "ok": True}
    """
    kw = keyword or company_name or industry
    bundle = []

    # 1) 个股公告（公司模式）
    if company_name:
        ann = fetch_company_announcements(company_name, days=days)
        if ann["items"]:
            bundle += ann["items"]

    # 2) 全市场公告按关键词过滤（行业模式也能命中公司公告）
    if industry or company_name:
        mk = fetch_market_announcements(days=1, keyword=kw, limit=limit_each)
        if mk["items"]:
            bundle += mk["items"]

    # 3) 7x24 快讯按关键词过滤
    flash = fetch_7x24_news(keyword=kw, pages=2, limit=limit_each)
    if flash["items"]:
        bundle += flash["items"]

    # 4) 搜狗网络搜索（关键词）
    web = search_web_news(kw, limit=limit_each)
    if web["items"]:
        bundle += web["items"]

    # 5) 兜底：行业关键词的财新/央视新闻（当上面都空时）
    if not bundle:
        for fn in (fetch_caixin_news, fetch_cctv_news):
            r = fn(limit=limit_each)
            if r["items"]:
                bundle += r["items"]

    sources = ["东方财富公告大全", "东方财富全市场公告", "东方财富7×24快讯",
               "搜狗新闻搜索", "财新网", "新闻联播"]
    return {"ok": bool(bundle), "items": _norm(bundle)[:30], "sources": sources,
            "note": "实时公开信息（公告/快讯/网络搜索，无需付费 API）"}


def format_items_markdown(items, max_items=10):
    """把新闻/公告列表格式化为报告可引用的 markdown 文本。"""
    if not items:
        return "（暂无实时新闻/公告数据）"
    lines = []
    for it in items[:max_items]:
        d = it.get("date", "")
        s = it.get("source", "")
        t = it.get("title", "")
        u = it.get("url", "")
        line = f"- {t}（{s}" + (f" · {d}" if d else "") + "）"
        if u:
            line += f" 来源链接: {u}"
        lines.append(line)
    return "\n".join(lines)


if __name__ == "__main__":
    # 自检
    r = fetch_news_bundle(company_name="宁德时代", industry="新能源汽车", keyword="宁德时代")
    print("bundle ok:", r["ok"], "items:", len(r["items"]))
    for it in r["items"][:8]:
        print("-", it["date"], it["title"][:50], "|", it["source"])
