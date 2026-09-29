"""Bounded, auditable research tools. Search snippets are never treated as filings."""
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from io import BytesIO
import ipaddress
import re
import socket
import time
import os
from urllib.parse import urljoin, urlparse, quote
from xml.etree import ElementTree

import requests
from bs4 import BeautifulSoup


def public_url(url):
    p = urlparse(url)
    if p.scheme not in ("http", "https") or not p.hostname or p.username or p.password:
        raise ValueError("只接受公开 HTTP(S) 资料")
    if p.port not in (None, 80, 443):
        raise ValueError("资料端口不受支持")
    for result in socket.getaddrinfo(p.hostname, p.port or 443, type=socket.SOCK_STREAM):
        if not ipaddress.ip_address(result[4][0]).is_global:
            raise ValueError("资料地址不是公网地址")
    return url


def download(url, max_bytes=18_000_000):
    """Validate every redirect, enforce byte/time budgets, preserve TLS verification."""
    deadline = time.monotonic() + 35
    for _ in range(5):
        public_url(url)
        agent = os.getenv("SEC_USER_AGENT", "StrategicResearch/1.0 public-filings-reader") if urlparse(url).hostname.endswith("sec.gov") else "Mozilla/5.0 (compatible; ResearchReader/1.0)"
        with requests.get(url, headers={"User-Agent": agent},
                          timeout=(6, 12), stream=True, allow_redirects=False) as response:
            if 300 <= response.status_code < 400:
                url = urljoin(url, response.headers.get("Location", ""))
                continue
            response.raise_for_status()
            chunks, size = [], 0
            for part in response.iter_content(65536):
                size += len(part)
                if size > max_bytes or time.monotonic() > deadline:
                    raise ValueError("资料超过读取大小或时限")
                chunks.append(part)
            return b"".join(chunks), response.headers.get("Content-Type", ""), url
    raise ValueError("资料跳转次数过多")


def cninfo_filings(target):
    """Resolve the issuer before fetching filings from CNINFO's official disclosure service."""
    endpoint="https://www.cninfo.com.cn/new/information/topSearch/query"
    public_url(endpoint)
    response=requests.post(endpoint,data={"keyWord":target,"maxNum":10},timeout=(6,15))
    response.raise_for_status()
    matches=response.json()
    if not isinstance(matches,list):
        return []
    exact=[m for m in matches if m.get("zwjc","").casefold()==target.casefold() or m.get("code")==target]
    if not exact:
        # Ambiguous company names must not silently select a similarly named company.
        return []
    issuer=next((m for m in exact if m.get("category")=="A股"),exact[0])
    column="hke" if issuer.get("category")=="港股" else "szse"
    data={"stock":issuer["code"]+","+issuer["orgId"],"tabName":"fulltext","pageSize":30,"pageNum":1,
          "column":column,"searchkey":"","seDate":f"{datetime.now().year-3}-01-01~{datetime.now().date()}",
          "sortName":"time","sortType":"desc","isHLtitle":"false"}
    if column!="hke": data["category"]="category_ndbg_szsh;category_bndbg_szsh;category_sjdbg_szsh;category_yjdbg_szsh"
    response=requests.post("https://www.cninfo.com.cn/new/hisAnnouncement/query",data=data,timeout=(6,15))
    response.raise_for_status()
    items=[]
    for item in response.json().get("announcements") or []:
        title=re.sub("<[^>]+>","",item.get("announcementTitle",""))
        if not any(word in title for word in ("年度报告","季度报告","半年度报告","年報","中期報告","业绩","業績")):
            continue
        items.append({"title":issuer["zwjc"]+" · "+title,"url":"https://static.cninfo.com.cn/"+item["adjunctUrl"],
                      "provider":"巨潮资讯公告","issuer":issuer["zwjc"],"ticker":issuer["code"]})
    # Include compact summaries for recent metrics and full statements for model inputs.
    return items[:8]


def sec_filings(target):
    import json
    raw,_,_=download("https://www.sec.gov/files/company_tickers.json",5_000_000)
    companies=json.loads(raw).values()
    tokens=set(re.findall(r"[A-Za-z]{1,8}",target.upper()))
    exact=[c for c in companies if c["ticker"] in tokens or c["title"].casefold()==target.casefold()]
    if len(exact)!=1:
        exact=[c for c in json.loads(raw).values() if target.casefold() in c["title"].casefold()]
    if len(exact)!=1:
        return []
    company=exact[0]
    cik=str(company["cik_str"]).zfill(10)
    raw,_,_=download(f"https://data.sec.gov/submissions/CIK{cik}.json",5_000_000)
    recent=json.loads(raw).get("filings",{}).get("recent",{})
    result=[]
    for i,form in enumerate(recent.get("form",[])):
        if form not in ("10-K","10-Q","20-F","6-K"):
            continue
        accession=recent["accessionNumber"][i].replace("-","")
        result.append({"title":company["title"]+" "+form+" "+recent["filingDate"][i],
                       "url":f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession}/{recent['primaryDocument'][i]}",
                       "provider":"SEC EDGAR","issuer":company["title"],"ticker":company["ticker"]})
        if len(result)>=6: break
    return result


def search(query, limit=5):
    """Domestic Sogou + Bing RSS, no subscription required; failures remain visible."""
    items, errors = [], []
    for engine, url in [
        ("搜狗", "https://www.sogou.com/web?query=" + quote(query)),
        ("Bing", "https://www.bing.com/search?format=rss&q=" + quote(query)),
    ]:
        try:
            raw, _, _ = download(url, 2_000_000)
            if engine == "Bing":
                nodes = ElementTree.fromstring(raw).findall(".//item")
                found = [{"title": n.findtext("title", ""), "url": n.findtext("link", "")} for n in nodes]
            else:
                soup = BeautifulSoup(raw, "html.parser")
                found = [{"title": a.get_text(" ", strip=True), "url": urljoin(url, a.get("href", ""))}
                         for a in soup.select(".vrwrap h3 a, .rb h3 a")]
            items.extend(dict(x, search_query=query, search_engine=engine) for x in found[:limit])
            if not found:
                errors.append(engine + "：无结果或触发验证")
        except Exception as exc:
            errors.append(engine + "：" + type(exc).__name__)
    return items, errors


def read_document(item):
    raw, mime, url = download(item["url"])
    pages, links = [], []
    if "pdf" in mime.lower() or raw[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(BytesIO(raw)) as pdf:
            total = len(pdf.pages)
            # Long annual reports: sample beginning and financial statements near end.
            indices = set(range(min(total, 40))) | set(range(max(40, total - 45), total))
            # Follow the actual table of contents into the financial section; statements
            # are often in the middle of a long annual report, before hundreds of notes.
            for idx in range(min(total,10)):
                first=pdf.pages[idx]
                text=first.extract_text() or ""
                for match in re.finditer(r"(?:财务报告|财务报表|合并利润表|合并资产负债表)[^\n\d]{0,180}(\d{1,3})\s*(?:\n|$)",text):
                    start=int(match.group(1))-1
                    if 0<=start<total:
                        indices.update(range(max(0,start-2),min(total,start+25)))
                first.close()
            indices=sorted(indices)[:125]
            for idx in indices:
                page=pdf.pages[idx]
                text = page.extract_text() or ""
                pages.append({"page": idx + 1, "text": text[:12000]})
                page.close()
        text = "\n".join(f"[第{p['page']}页] {p['text']}" for p in pages)
    else:
        soup = BeautifulSoup(raw, "html.parser")
        for a in soup.select("a[href]"):
            href = urljoin(url, a["href"])
            label = a.get_text(" ", strip=True)
            if re.search(r"\.pdf(?:\?|$)", href, re.I):
                links.append({"url": href, "title": label or item["title"]})
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        text = soup.get_text(" ", strip=True)
    if len(text) < 200:
        raise ValueError("正文为空、扫描件或需要登录")
    return {**item, "url": url, "text": text[:400000], "pages": pages,
            "pdf_links": links[:4], "retrieved_at": datetime.now(timezone.utc).isoformat(),
            "domain": urlparse(url).hostname, "method": "PDF正文" if pages else "网页正文"}


def collect(target, kind, period, progress=lambda *_: None, uploaded_text=""):
    year = datetime.now().year
    suffixes = ([f"{year} 最新 财报 年报 投资者关系 revenue", "annual report financial statements investor relations",
                 "产品 业务分部 毛利率 竞争对手", "研发 投资 产能 capex guidance", "股权 风险 招股说明书"]
                if kind == "公司" else
                [f"{year} 最新 统计 数据 协会", "市场规模 竞争格局 出货量", "产业链 技术路线 商业化",
                 "政策 监管 风险", "龙头公司 年报 收入 利润"])
    if "默认" not in period:
        suffixes.append(period + " 公告 数据")
    candidates, audit = [], []
    progress(0.03, "检索财报、行业统计、产品与风险资料")
    with ThreadPoolExecutor(max_workers=4) as pool:
        filing_job = pool.submit(cninfo_filings if re.search(r"[\u4e00-\u9fff]|^\d{5,6}$",target) else sec_filings,target) if kind=="公司" else None
        jobs = {pool.submit(search, target + " " + q, 3): q for q in suffixes}
        for job in as_completed(jobs):
            items, errors = job.result()
            candidates.extend(items)
            audit.append({"query": target + " " + jobs[job], "results": len(items), "errors": errors})
        if filing_job:
            try:
                filings=filing_job.result()
                candidates=filings+candidates
                audit.append({"provider":"官方公告","results":len(filings)})
            except Exception as exc:
                audit.append({"provider":"官方公告","error":type(exc).__name__+": "+str(exc)[:100]})
    # Round-robin per query, favor primary-source domains without declaring other pages official.
    unique = {x["url"]: x for x in candidates if x.get("url", "").startswith("http")}
    primary = ("cninfo.com.cn", "sse.com.cn", "szse.cn", "hkexnews.hk", "sec.gov", "stats.gov.cn")
    selected = sorted(unique.values(), key=lambda x: not any(d in urlparse(x["url"]).netloc for d in primary))[:24]
    docs = []
    with ThreadPoolExecutor(max_workers=5) as pool:
        jobs = {pool.submit(read_document, item): item for item in selected}
        for job in as_completed(jobs):
            try:
                doc = job.result()
                # A search redirect that ends on a search page is not a source document.
                if doc["domain"] in ("www.sogou.com", "www.bing.com"):
                    raise ValueError("未解析到资料原文")
                words=[target.casefold()] + re.findall(r"[a-zA-Z]{2,}",target.casefold())
                if not doc.get("provider") and not any(w in (doc["title"]+" "+doc["text"]).casefold() for w in words):
                    raise ValueError("正文未匹配研究主体，已剔除无关结果")
                docs.append(doc)
            except Exception as exc:
                audit.append({"url": jobs[job]["url"], "error": str(exc)[:120]})
            progress(0.10 + .12 * (len(docs) / max(1, len(selected))), f"已读取 {len(docs)} 份原文")
    pdfs = {x["url"]: x for d in docs for x in d["pdf_links"] if x["url"] not in unique}
    with ThreadPoolExecutor(max_workers=4) as pool:
        jobs = {pool.submit(read_document, item): item for item in list(pdfs.values())[:8]}
        for job in as_completed(jobs):
            try:
                docs.append(job.result())
            except Exception as exc:
                audit.append({"url": jobs[job]["url"], "error": str(exc)[:120]})
    if uploaded_text:
        docs.append({"title": "用户上传资料（独立核验前仅作用户提供信息）", "url": "用户上传",
                     "text": uploaded_text[:100000], "domain": "user-upload", "pages": [],
                     "method": "用户上传正文", "retrieved_at": datetime.now(timezone.utc).isoformat()})
    # Stable ordering makes retry/resume references deterministic.
    docs = list({d["url"]: d for d in docs}.values())
    for i, doc in enumerate(sorted(docs, key=lambda d: d["url"])):
        doc["id"] = "S" + str(i + 1)
    return sorted(docs, key=lambda d: int(d["id"][1:])), audit
