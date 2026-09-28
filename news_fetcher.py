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

try:
    import requests
except Exception as _requests_error:
    requests = None
    print(f"[News] requests 不可用，将只使用本地缓存/可用数据源: {_requests_error}")

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


def _retrieved_at():
    """Return an ISO-8601 UTC timestamp for provenance in UI/export rows."""
    return datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()


def _result(ok, items=None, note="", source="", *, attempted=True,
            fetched_at=None, from_cache=False):
    """Build a consistent source result and never call an empty response live."""
    clean_items = _norm(items or [])
    return {
        "ok": bool(ok and clean_items),
        "items": clean_items,
        "note": note,
        "source": source,
        "attempted": bool(attempted),
        "fetched_at": fetched_at or _retrieved_at(),
        "from_cache": bool(from_cache),
        "has_citable_url": any(
            str(x.get("url") or "").lower().startswith(("http://", "https://"))
            for x in clean_items
        ),
    }


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
        source = str(it.get("source") or "未知来源（待核实）")[:80]
        raw_url = str(it.get("url") or "").strip()
        # pandas/akshare frequently serializes missing links as literal
        # strings such as "nan" or "None"; these are not citable URLs.
        url = "" if raw_url.lower() in {"nan", "none", "null", "n/a", "-"} else raw_url
        out.append({
            "title": title[:200],
            "date": str(it.get("date") or "")[:20],
            "source": source,
            "url": url,
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
            return _result(False, note="未找到该公司的股票代码，无法抓取公告",
                           source="东方财富·个股公告")
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
        return _result(bool(items), items,
                       f"东方财富公告大全 · {company_name or stock_code} 近{days}天" if items
                       else f"东方财富公告大全 · {company_name or stock_code} 近{days}天无结果",
                       source="东方财富·个股公告")
    except Exception as e:
        return _result(False, note=f"公告抓取失败: {str(e)[:120]}",
                       source="东方财富·个股公告")


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
            from_cache = False
        else:
            from_cache = True
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
        return _result(bool(items), items[:limit],
                       "东方财富全市场公告" + (f"· 关键词「{keyword}」" if keyword else "")
                       + ("（无匹配结果）" if not items else ""),
                       source="东方财富·全市场公告", from_cache=from_cache)
    except Exception as e:
        return _result(False, note=f"全市场公告抓取失败: {str(e)[:120]}",
                       source="东方财富·全市场公告")


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
        return _result(bool(items), items[:limit],
                       "东方财富 7×24 实时快讯" + (f"· 关键词「{keyword}」" if keyword else "")
                       + ("（无匹配结果）" if not items else ""),
                       source="东方财富·7×24快讯")
    except Exception as e:
        return _result(False, note=f"快讯抓取失败: {str(e)[:120]}",
                       source="东方财富·7×24快讯")


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
        return _result(bool(items), items[:limit],
                       "财新网财经新闻" + ("（无结果）" if not items else ""),
                       source="财新网")
    except Exception as e:
        return _result(False, note=f"财新新闻抓取失败: {str(e)[:120]}", source="财新网")


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
        return _result(bool(items), items[:limit],
                       "新闻联播文字稿" + ("（无结果）" if not items else ""),
                       source="新闻联播文字稿")
    except Exception as e:
        return _result(False, note=f"新闻联播抓取失败: {str(e)[:120]}",
                       source="新闻联播文字稿")


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
        return _result(bool(items), items, f"搜狗新闻搜索 · {keyword}"
                       + ("（无结果）" if not items else ""), source="搜狗新闻搜索")
    except Exception as e:
        return _result(False, note=f"网络搜索失败: {str(e)[:120]}", source="搜狗新闻搜索")


def _resolve_sogou_link(link):
    """把搜狗跳转链接（/link?url=加密串）还原为可访问的真实地址。
    搜狗对 /link?url=... 的加密串需要请求其跳转接口解码，失败则返回原链接。"""
    if not link:
        return ""
    if not link.startswith("/link?"):
        return link
    import urllib.parse
    q = urllib.parse.parse_qs(urllib.parse.urlparse(link).query)
    if q.get("url") and str(q["url"][0]).startswith("http"):
        return q["url"][0]
    # 加密跳转：请求搜狗跳转接口拿 Location
    try:
        ua = {"User-Agent": UA["User-Agent"], "Referer": "https://www.sogou.com/"}
        r = requests.get("https://www.sogou.com" + link, headers=ua,
                         timeout=10, allow_redirects=False)
        loc = r.headers.get("Location", "")
        if loc and loc.startswith("http"):
            return loc
        if loc.startswith("/link?"):
            return _resolve_sogou_link(loc)
    except Exception:
        pass
    return link


def search_web_pages(keyword, limit=8):
    """
    搜狗通用网页搜索（免费、无需 Key）。
    用于「资料缺口审查」：自动搜索年报 PDF / 行业研报 / 政策原文等资料的获取链接。
    返回 {"ok", "items": [{"title","url","source","summary"}], "note"}
    """
    try:
        from bs4 import BeautifulSoup
        params = {"query": keyword}
        r = requests.get("https://www.sogou.com/web", params=params,
                         headers=UA, timeout=15)
        soup = BeautifulSoup(r.text, "lxml")
        items = []
        nodes = soup.select(".vrwrap") or soup.select(".rb") or []
        for node in nodes[:limit]:
            h3 = node.select_one("h3 a") or node.select_one("h3")
            if h3 is None:
                continue
            title = h3.get_text(strip=True)
            link = h3.get("href", "") if h3.has_attr("href") else ""
            src = node.select_one(".citeurl") or node.select_one(".fz-mid")
            source = src.get_text(strip=True) if src else "搜狗网页"
            summary_node = node.select_one(".text-layout") or node.select_one(".str_info")
            summary = summary_node.get_text(strip=True) if summary_node else ""
            items.append({"title": title, "date": "", "source": source,
                          "url": _resolve_sogou_link(link), "summary": summary[:200]})
        return _result(bool(items), items, f"搜狗网页搜索 · {keyword}"
                       + ("（无结果）" if not items else ""), source="搜狗网页搜索")
    except Exception as e:
        return _result(False, note=f"网页搜索失败: {str(e)[:120]}", source="搜狗网页搜索")


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
    source_status = []

    def _record(label, result):
        source_status.append({
            "source": label,
            "ok": bool(result.get("ok")),
            "count": len(result.get("items") or []),
            "note": result.get("note", ""),
            "fetched_at": result.get("fetched_at", ""),
            "from_cache": bool(result.get("from_cache", False)),
            "has_citable_url": bool(result.get("has_citable_url", False)),
        })
        return result

    # 1) 个股公告（公司模式）
    if company_name:
        ann = _record("东方财富·个股公告", fetch_company_announcements(company_name, days=days))
        if ann["items"]:
            bundle += ann["items"]

    # 2) 全市场公告按关键词过滤（行业模式也能命中公司公告）
    if industry or company_name:
        mk = _record("东方财富·全市场公告", fetch_market_announcements(days=1, keyword=kw, limit=limit_each))
        if mk["items"]:
            bundle += mk["items"]

    # 3) 7x24 快讯按关键词过滤
    flash = _record("东方财富·7×24快讯", fetch_7x24_news(keyword=kw, pages=2, limit=limit_each))
    if flash["items"]:
        bundle += flash["items"]

    # 4) 搜狗网络搜索（关键词）
    web = _record("搜狗新闻搜索", search_web_news(kw, limit=limit_each))
    if web["items"]:
        bundle += web["items"]

    # 5) 兜底：行业关键词的财新/央视新闻（当上面都空时）
    if not bundle:
        for fn in (fetch_caixin_news, fetch_cctv_news):
            r = _record(fn.__name__, fn(limit=limit_each))
            if r["items"]:
                bundle += r["items"]

    # Only advertise sources that were actually attempted.  This prevents a
    # chart/report from implying that an uncalled provider supplied data.
    sources = [row["source"] for row in source_status]
    ok_count = sum(1 for item in source_status if item["ok"] and item["count"] > 0)
    normalized = _norm(bundle)[:30]
    return {"ok": bool(normalized), "items": normalized, "sources": sources,
            "source_status": source_status, "successful_sources": ok_count,
            "fetched_at": _retrieved_at(),
            "note": f"实时公开信息；成功返回 {ok_count}/{len(source_status)} 个来源（失败来源不阻断流程）"
                    if normalized else
                    f"未获取到可引用公开信息；成功返回 {ok_count}/{len(source_status)} 个来源，不能视为最新事实"}


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
