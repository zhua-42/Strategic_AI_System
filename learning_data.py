# -*- coding: utf-8 -*-
"""
学习资料库（Learning Library）
===============================
将 Forage 模拟成果、Deloitte 参考资料、券商行研方法论（qweqwe）整理为
可学习的资料并入库 SQLite（learning_resources 表），供网站浏览/检索/下载。

资料文件存放于 knowledge/learning/（md/pdf），库表存元数据+摘要+标签。
"""
import os
import json
import sqlite3
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LEARNING_DIR = os.path.join(BASE_DIR, "knowledge", "learning")
DB_PATH = os.path.join(BASE_DIR, "financial_research.db")

TAG_CATALOG = {
    "forage": "Forage 模拟练习",
    "deloitte": "Deloitte 案例",
    "报告方法": "券商行研方法论",
    "财务分析": "财务报表分析",
}


def init_learning_schema():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """CREATE TABLE IF NOT EXISTS learning_resources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            source TEXT DEFAULT '',
            resource_type TEXT DEFAULT 'markdown',
            summary TEXT DEFAULT '',
            tags TEXT DEFAULT '',
            file_path TEXT DEFAULT '',
            file_size INTEGER DEFAULT 0,
            created_at TEXT DEFAULT ''
        )"""
    )
    conn.commit()
    conn.close()


def _norm_tags(tags):
    if isinstance(tags, (list, tuple)):
        return ",".join(str(t) for t in tags)
    return str(tags or "")


def upsert_resource(title, source, resource_type, summary, tags, file_path, file_size=0):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id FROM learning_resources WHERE title=?", (title,))
    exists = cur.fetchone()
    if exists:
        cur.execute(
            """UPDATE learning_resources SET
                 source=?, resource_type=?, summary=?, tags=?, file_path=?, file_size=?
               WHERE id=?""",
            (source or "", resource_type or "markdown", summary or "",
             _norm_tags(tags) or "", file_path or "", int(file_size or 0), exists[0]),
        )
    else:
        cur.execute(
            """INSERT INTO learning_resources
               (title, source, resource_type, summary, tags, file_path, file_size, created_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (title, source or "", resource_type or "markdown", summary or "",
             _norm_tags(tags) or "", file_path or "", int(file_size or 0),
             datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
    conn.commit()
    conn.close()


def search_resources(keyword="", tag="", limit=100):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    sql = "SELECT id, title, source, resource_type, summary, tags, file_path, file_size, created_at FROM learning_resources WHERE 1=1"
    params = []
    if keyword:
        sql += " AND (title LIKE ? OR summary LIKE ? OR tags LIKE ?)"
        like = f"%{keyword}%"
        params += [like, like, like]
    if tag:
        sql += " AND tags LIKE ?"
        params.append(f"%{tag}%")
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(int(limit))
    rows = cur.execute(sql, params).fetchall()
    conn.close()
    return [
        {
            "id": r[0], "title": r[1], "source": r[2], "resource_type": r[3],
            "summary": r[4], "tags": r[5], "file_path": r[6], "file_size": r[7],
            "created_at": r[8],
        }
        for r in rows
    ]


def count_resources():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        n = cur.execute("SELECT COUNT(*) FROM learning_resources").fetchone()[0]
    except Exception:
        n = 0
    conn.close()
    return n


def scan_and_sync_learning_dir():
    """扫描 knowledge/learning/ 下的文件，自动入库（文件级）。"""
    if not os.path.isdir(LEARNING_DIR):
        return 0
    n = 0
    for fn in sorted(os.listdir(LEARNING_DIR)):
        fp = os.path.join(LEARNING_DIR, fn)
        if not os.path.isfile(fp):
            continue
        size = os.path.getsize(fp)
        title = os.path.splitext(fn)[0]
        rtype = "pdf" if fn.lower().endswith(".pdf") else "markdown"
        upsert_resource(title, "网站学习资料库", rtype, "", ["网站资料"], fn, size)
        n += 1
    return n
