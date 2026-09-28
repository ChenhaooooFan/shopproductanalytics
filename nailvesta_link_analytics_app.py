#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NailVesta 链接数据深度分析 (Streamlit)
=====================================
输入：TikTok Shop "product analysis list" 每日导出 xlsx（每天一个文件，175 列）
覆盖：3/20 - 9/20 六个月，逐日 × 逐链接 × 全指标

功能
----
1. 全指标归档：175 列自动分类为【前端(流量/漏斗)】与【后端(成交/履约/售后)】
2. 两段对比·归因：任意两段时间（默认最近 7 天 vs 前 7 天），全指标记分卡 +
   『谁导致的』——按链接拆出占跌量、按渠道拆出掉在哪、漏斗四因子、比率型指标的
   比率效应 / 结构效应分解
3. 多粒度对比：日 / 周 / 月 / 季，规模指标一律按『日均』比（各期天数不同）
4. 异常检测与归因：前 N 期中位数 + MAD 检测突增突降，自动给出主责链接与原因
5. 日环比链条 & 回暖：大跌后逐日追踪（基线 = 跌前 7 日中位数）
6. 改动效果溯源：自动识别改名/重组、件单价突变（改价/改促销）+ 你登记的改动

运行：  streamlit run nailvesta_link_analytics_app.py
依赖：  streamlit pandas numpy plotly openpyxl
"""
from __future__ import annotations

import hashlib
import html as _html
import io
import json
import os
import re
import zipfile

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="NailVesta 链接深度分析", page_icon="📊", layout="wide")

# 全局外观：配色主题在 .streamlit/config.toml 里（primaryColor 等），这里补 config.toml
# 管不到的细节——数字卡片不截断、按钮/卡片/侧边栏的圆角阴影、字距，让整体更像一个统一设计的产品
# 而不是默认 Streamlit 模板的感觉。
st.markdown("""
<style>
/* 颜色一律用半透明灰 + 继承文字色：深色 / 浅色主题下都看得清（Cloud 上没传 config.toml 时是浅色）。 */
div[data-testid="stMetricValue"] {
    font-size: clamp(0.85rem, 1.6vw, 1.6rem) !important; font-weight: 650 !important;
    white-space: normal !important; overflow-wrap: break-word !important; line-height: 1.25 !important;
}
div[data-testid="stMetricLabel"] { font-size: 0.78rem !important; white-space: normal !important; opacity: .72; }
div[data-testid="stMetricDelta"] { font-size: 0.78rem !important; white-space: normal !important; }
div[data-testid="stMetric"] {
    background: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.18);
    border-radius: 14px; padding: 14px 16px 10px;
}
h1 { letter-spacing: -0.01em; font-weight: 750 !important; }
h2, h3, h4 { letter-spacing: -0.005em; font-weight: 650 !important; }
.block-container { padding-top: 2.2rem !important; max-width: 1320px; }

/* 按钮 / 输入 / 表格：圆角、轻阴影 */
.stButton > button, .stDownloadButton > button, .stLinkButton > a {
    border-radius: 10px !important; font-weight: 550 !important;
    transition: transform .12s ease, box-shadow .12s ease, filter .12s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 14px rgba(0,0,0,0.18); }
.stButton > button[kind="primary"] { box-shadow: 0 2px 10px rgba(194,65,107,0.30); }
div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, .stTextInput input,
div[data-testid="stExpander"] { border-radius: 10px !important; }
div[data-testid="stDataFrame"], div[data-testid="stTable"] { border-radius: 12px; overflow: hidden; }
div[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px !important; transition: border-color .15s ease, background .15s ease; }
div[data-testid="stVerticalBlockBorderWrapper"]:hover { border-color: rgba(194,65,107,0.45) !important; }

/* 分区标签：做成 APP 里的分段切换 */
div[data-baseweb="tab-list"] { gap: 6px; border-bottom: 1px solid rgba(128,128,128,0.18); }
button[data-baseweb="tab"] { border-radius: 10px 10px 0 0 !important; padding: 8px 16px !important; font-weight: 600 !important; }
button[data-baseweb="tab"][aria-selected="true"] { background: rgba(194,65,107,0.10) !important; }

/* 侧边栏 */
section[data-testid="stSidebar"] { border-right: 1px solid rgba(128,128,128,0.18); }
section[data-testid="stSidebar"] [data-testid="stPageLink"] p { font-weight: 500; }
header[data-testid="stHeader"] { background: transparent; }

/* 页头 */
.nv-head { display: flex; justify-content: space-between; align-items: flex-end; gap: 16px; flex-wrap: wrap;
    padding: 2px 0 14px; margin-bottom: 14px; border-bottom: 1px solid rgba(128,128,128,0.22); }
.nv-title { font-size: 1.7rem; font-weight: 760; letter-spacing: -0.01em; line-height: 1.2; }
.nv-sub { opacity: .68; font-size: .92rem; margin-top: 4px; }
.nv-chips { display: flex; flex-wrap: wrap; gap: 6px; justify-content: flex-end; }
.nv-chip { font-size: .78rem; padding: 3px 10px; border-radius: 999px; background: rgba(128,128,128,0.10);
    border: 1px solid rgba(128,128,128,0.20); white-space: nowrap; }

/* 状态胶囊 */
.nv-sum { display: flex; gap: 8px; flex-wrap: wrap; margin: 2px 0 12px; }
.nv-pill { padding: 6px 12px; border-radius: 10px; font-weight: 650; font-size: .9rem; }
.nv-pill.down { background: rgba(217,83,79,0.14); color: #D9534F; }
.nv-pill.up { background: rgba(46,158,122,0.14); color: #1F9A70; }
.nv-pill.warn { background: rgba(224,160,48,0.16); color: #C98A12; }
.nv-pill.flat, .nv-pill.info { background: rgba(128,128,128,0.12); }

/* 结论卡 */
.nv-cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 12px; margin: 4px 0 14px; }
.nv-card { border-radius: 14px; padding: 13px 16px; background: rgba(128,128,128,0.06);
    border: 1px solid rgba(128,128,128,0.18); border-left: 4px solid rgba(128,128,128,0.5); }
.nv-card.down { border-left-color: #D9534F; } .nv-card.up { border-left-color: #2E9E7A; }
.nv-card.warn { border-left-color: #E0A030; } .nv-card.info { border-left-color: #4C6EF5; }
.nv-card-t { font-weight: 680; font-size: .96rem; margin-bottom: 6px; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.nv-card-b { font-size: .9rem; line-height: 1.65; opacity: .92; }
.nv-tag { font-size: .72rem; padding: 1px 8px; border-radius: 999px; background: rgba(128,128,128,0.16); font-weight: 500; }
.nv-sec { font-size: 1.08rem; font-weight: 680; margin: 18px 0 8px; }
.nv-sec-note { font-size: .8rem; font-weight: 400; opacity: .62; margin-left: 10px; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────
# 1. 指标归档：前端 / 后端 分类规则
# ──────────────────────────────────────────────────────────────────────
ID_COLS = ["Product Name", "Product ID", "GMV range"]

# 前端 = 流量获取与漏斗效率（曝光→点击→加购→下单转化率），以及内容供给
FRONT_PATTERNS = [
    "impression", "click", "ctr", "ctor", "add-to-cart", "atc",
    "new live counts", "new video count", "posted content",
]
# 后端 = 成交金额/订单/件数/客户/履约/售后
BACK_PATTERNS = [
    "gmv", "order", "items", "customer", "aov", "refund",
    "shipping", "tax", "merchandise",
]

# 指标类型（决定聚合方式与展示格式）
MONEY_HINTS = ["gmv", "aov", "refund", "shipping fee", "tax", "merchandise"]
RATE_HINTS = ["ctr", "ctor", "rate", "%"]


def classify(metric: str) -> str:
    """把单个指标名归类为 前端 / 后端 / 其他。前端优先（CTR 等比率属漏斗）。"""
    m = metric.lower()
    for p in FRONT_PATTERNS:
        if p in m:
            return "前端"
    for p in BACK_PATTERNS:
        if p in m:
            return "后端"
    return "其他"


def value_kind(metric: str) -> str:
    m = metric.lower()
    if any(h in m for h in RATE_HINTS):
        return "比率"
    if any(h in m for h in MONEY_HINTS):
        return "金额"
    return "计数"


def parse_num(v) -> float:
    """'$1,234.56' / '4.33%' / 1234 -> float。空/异常 -> 0.0"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return 0.0 if pd.isna(v) else float(v)
    s = str(v).strip().replace("$", "").replace(",", "").replace("%", "")
    if s in ("", "-", "nan", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def short_name(name) -> str:
    """链接简称：去掉 'NailVesta' 和 ' – 100% Handcrafted 3D Press-on…' 这类描述尾巴。"""
    n = re.sub(r"\s+", " ", str(name).replace("NailVesta", "").strip())
    parts = [x.strip() for x in re.split(r"\s[–—-]\s|–|\|", n) if x.strip()]
    s = parts[0] if parts else n
    s = re.sub(r"^100% ", "", s)
    s = re.sub(r"\s*3D Design.*$", "", s)
    s = re.sub(r"\s*Handcrafted.*$", "", s).strip()
    return (s or n)[:34]


def round_num(d: pd.DataFrame, n: int = 2) -> pd.DataFrame:
    """只对数值列四舍五入（带日期列的表直接 .round 在 pandas 3 会报警告）。"""
    return d.round({c: n for c in d.select_dtypes("number").columns})


# ──────────────────────────────────────────────────────────────────────
# 2. 解析单个日报文件
# ──────────────────────────────────────────────────────────────────────
DATE_RE = re.compile(r"(\d{2})/(\d{2})/(\d{4})")


def _read_one(buf, fname: str):
    """返回 (date, DataFrame[识别列 + '区段::指标' 列], catalog_records)"""
    head = pd.read_excel(buf, header=None, nrows=4, engine="openpyxl")
    a1 = str(head.iloc[0, 0]) if head.shape[0] > 0 else ""
    m = DATE_RE.search(a1)
    if m:
        d = pd.Timestamp(f"{m.group(3)}-{m.group(2)}-{m.group(1)}")
    else:  # 回退：从文件名 26MMDD / 2026MMDD 猜
        m2 = re.search(r"(20\d{2})(\d{2})(\d{2})", fname) or re.search(r"\b26(\d{2})(\d{2})\b", fname)
        if not m2:
            return None, None, None
        g = m2.groups()
        d = pd.Timestamp(f"{g[0]}-{g[1]}-{g[2]}") if len(g[0]) == 4 else pd.Timestamp(f"2026-{g[0]}-{g[1]}")

    sections = head.iloc[2].tolist() if head.shape[0] > 2 else []
    headers = head.iloc[3].tolist() if head.shape[0] > 3 else []

    if isinstance(buf, io.BytesIO):
        buf.seek(0)
    body = pd.read_excel(buf, header=None, skiprows=4, engine="openpyxl")
    if body.empty:
        return None, None, None

    ncol = min(len(headers), body.shape[1])
    keys, catalog, seen = [], [], {}
    for i in range(ncol):
        hdr = str(headers[i]).strip() if headers[i] is not None else f"col{i}"
        sec_raw = sections[i] if i < len(sections) else None
        sec = str(sec_raw).strip() if sec_raw not in (None, "") and str(sec_raw) != "nan" else "识别"
        if hdr in ID_COLS:
            key = hdr
        else:
            key = f"{sec}::{hdr}"
            if key in seen:                       # 同区段重名 -> 加序号
                seen[key] += 1
                key = f"{key}#{seen[key]}"
            else:
                seen[key] = 1
        keys.append(key)
        if hdr not in ID_COLS:
            catalog.append(dict(列号=i + 1, 区段=sec, 指标=hdr, key=key,
                                分类=classify(hdr), 类型=value_kind(hdr)))

    body = body.iloc[:, :ncol]
    body.columns = keys
    body = body[body.get("Product ID").notna()] if "Product ID" in body.columns else body
    if body.empty:
        return None, None, None

    # 向量化数值清洗（逐格 map 在 185 个文件 × 6M 单元格下太慢且会造成 DataFrame 碎片化）
    num = {}
    for k in keys:
        if k in ID_COLS:
            continue
        s = body[k]
        # ⚠️ pandas 3.0 起字符串列的 dtype 是 'str'（StringDtype）而不是 'object'，
        #    用 `dtype == object` 判断会漏掉，导致 '$439.70' 直接被 to_numeric 变成 NaN→0。
        #    这里一律用 is_numeric_dtype 反向判断。
        if not pd.api.types.is_numeric_dtype(s):
            s = (s.astype("string")
                  .str.replace(",", "", regex=False)
                  .str.replace("$", "", regex=False)
                  .str.replace("%", "", regex=False)
                  .str.strip()
                  .replace({"": pd.NA, "-": pd.NA, "nan": pd.NA, "None": pd.NA}))
        num[k] = pd.to_numeric(s, errors="coerce").fillna(0.0).astype("float64")

    ident = pd.DataFrame({
        "日期": d,
        "product_id": body["Product ID"].astype(str).str.strip(),
        "product_name": body["Product Name"].astype(str).str.strip(),
    }, index=body.index)
    out = pd.concat([ident, pd.DataFrame(num, index=body.index)], axis=1).reset_index(drop=True)
    return d, out, catalog


PARSE_VER = "ops-v1"          # 改了 _read_one 的识别逻辑就把这个版本号 +1，旧的识别结果会自动作废、重新识别


def _parse_dir() -> str:
    return os.path.join(os.path.expanduser("~/nailvesta_lark_cache"), f"_parsed_{PARSE_VER}")


def _read_one_cached(raw: bytes, name: str):
    """
    识别一个日报文件，结果按「文件内容指纹」存到磁盘：同一个文件第二次起直接读结果（0.00x 秒），
    只有新增的、或者被重新上传改过内容的文件才真正去识别（每个约 0.2 秒）。
    返回 (日期, 数据, 指标目录, 这次是不是新识别的)。
    """
    key = hashlib.sha1(raw).hexdigest()[:24]
    pdir = _parse_dir()
    meta_p, data_p = os.path.join(pdir, key + ".json"), os.path.join(pdir, key + ".parquet")
    try:
        if os.path.exists(meta_p):
            with open(meta_p, encoding="utf-8") as fh:
                meta = json.load(fh)
            if meta.get("skip"):
                return None, None, None, False
            if os.path.exists(data_p):
                return pd.Timestamp(meta["date"]), pd.read_parquet(data_p), meta.get("catalog"), False
    except Exception:                               # noqa: BLE001  缓存坏了就当没有，重新识别
        pass
    d, df, cat = _read_one(io.BytesIO(raw), name)
    try:
        os.makedirs(pdir, exist_ok=True)
        if df is not None:
            df.to_parquet(data_p, index=False)
        with open(meta_p, "w", encoding="utf-8") as fh:
            json.dump({"name": name, "skip": df is None, "date": str(d.date()) if d is not None else None,
                       "catalog": cat}, fh, ensure_ascii=False, default=str)
    except Exception:                               # noqa: BLE001  写不了缓存不影响这次结果
        pass
    return d, df, cat, True


@st.cache_data(show_spinner=False)
def load_files(payloads: list[tuple[str, bytes]]):
    """payloads: [(filename, bytes)] -> (long_df, catalog_df, log, stats)"""
    frames, catalog, log = [], None, []
    t0 = __import__("time").time()
    n_new = n_skip = 0
    prog = st.progress(0.0, text="读取中…")
    for i, (name, raw) in enumerate(payloads):
        try:
            d, df, cat, fresh = _read_one_cached(raw, name)
            n_new += fresh
            if df is None:
                n_skip += 1
                log.append(f"跳过（无法识别日期/内容）：{name}")
            else:
                frames.append(df)
                if catalog is None and cat:
                    catalog = pd.DataFrame(cat)
                log.append(f"{'✓ 新识别' if fresh else '· 已识别'} {d.date()}  {name}  ({len(df)} 链接)")
        except Exception as e:                      # noqa: BLE001
            log.append(f"✗ {name}: {e}")
        prog.progress((i + 1) / max(len(payloads), 1),
                      text=f"读取中… {i+1}/{len(payloads)}（新识别 {n_new} 个）")
    prog.empty()
    stats = {"files": len(payloads), "new": n_new, "cached": len(payloads) - n_new, "skipped": n_skip,
             "secs": __import__("time").time() - t0}
    if not frames:
        return pd.DataFrame(), pd.DataFrame(), log, stats

    df = pd.concat(frames, ignore_index=True)
    df = df.drop_duplicates(subset=["日期", "product_id"], keep="last")
    df = df.sort_values(["日期", "product_id"]).reset_index(drop=True)
    return df, (catalog if catalog is not None else pd.DataFrame()), log, stats


def gather_from_folder(folder: str):
    out = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                p = os.path.join(root, f)
                try:
                    with open(p, "rb") as fh:
                        out.append((f, fh.read()))
                except OSError:
                    pass
    return out


def gather_from_zip(raw: bytes):
    out = []
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        for n in z.namelist():
            if n.lower().endswith((".xlsx", ".xls")) and not os.path.basename(n).startswith("~$") \
               and "__MACOSX" not in n:
                out.append((os.path.basename(n), z.read(n)))
    return out


# ──────────────────────────────────────────────────────────────────────
# 2b. 飞书 / Lark Base 直连：自动拉取「链接文件」附件
# ──────────────────────────────────────────────────────────────────────
LARK_HOSTS = {
    "Lark 国际版 (larksuite.com)": "https://open.larksuite.com",
    "飞书 国内版 (feishu.cn)": "https://open.feishu.cn",
}
DEFAULT_CACHE = os.path.expanduser("~/nailvesta_lark_cache")


def parse_base_url(url: str):
    """从 Base 链接里抠出 app_token；支持 /base/xxx 与 /wiki/xxx"""
    m = re.search(r"/(?:base|wiki)/([A-Za-z0-9]+)", url or "")
    return m.group(1) if m else (url or "").strip()


class NetworkError(RuntimeError):
    """连不上 Lark 服务器（SSL 被重置 / 连接中断 / 超时）——与权限问题区分开"""


def _lark_api(host, path, token=None, method="GET", retries: int = 4, **kw):
    """
    带自动重试：SSL EOF、连接重置、超时、5xx 属于网络抖动，指数退避重试（1s,2s,4s,8s）。
    实测：偶发一次 SSLEOFError 后立即重试即成功，185 个文件连续下载必须有这层保护。
    """
    import time
    import requests
    h = {"Authorization": f"Bearer {token}"} if token else {}
    h.update(kw.pop("headers", {}))
    transient = (requests.exceptions.SSLError, requests.exceptions.ConnectionError,
                 requests.exceptions.Timeout, requests.exceptions.ChunkedEncodingError)
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.request(method, f"{host}/open-apis{path}", headers=h, timeout=60, **kw)
            if (r.status_code >= 500 or r.status_code == 429) and attempt < retries:     # 5xx / 限流：等一下再试
                time.sleep(2 ** attempt)
                continue
            return r
        except transient as e:
            last = e
            if attempt < retries:
                time.sleep(2 ** attempt)
    raise NetworkError(f"连不上 Lark 服务器（已自动重试 {retries} 次）：{type(last).__name__}") from last


def _lark_fetch_media(host, token, ftok, local, min_size: int = 500):
    """下载一个附件到 local（先写 .part 再改名，下载一半断掉不会留下坏文件）。成功返回 None，失败返回原因。
    在线程里跑，不能碰 st.*。"""
    try:
        r = _lark_api(host, f"/drive/v1/medias/{ftok}/download", token=token)
        if r.status_code == 200 and len(r.content) > min_size:
            tmp = local + ".part"
            with open(tmp, "wb") as fh:
                fh.write(r.content)
            os.replace(tmp, local)
            return None
        try:
            j = r.json()
            return f"{j.get('code')} {j.get('msg')}"
        except Exception:                                   # noqa: BLE001
            return f"HTTP {r.status_code}"
    except NetworkError as e:
        return str(e)


def _lark_fetch_many(host, token, jobs, label: str, min_size: int = 500, workers: int = 6) -> dict:
    """jobs = [(显示名, file_token, 本地路径)]，几个一起下：Lark 单个附件下载要 3–6 秒，一个一个下几百个文件太慢。
    返回 {显示名: 失败原因}。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    fails: dict = {}
    if not jobs:
        return fails
    prog = st.progress(0.0, text=f"{label}… 0/{len(jobs)}")
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_lark_fetch_media, host, token, ft, lp, min_size): nm for nm, ft, lp in jobs}
        for i, f in enumerate(as_completed(futs)):
            err = f.result()
            if err:
                fails[futs[f]] = err
            prog.progress((i + 1) / len(jobs), text=f"{label}… {i + 1}/{len(jobs)}")
    prog.empty()
    return fails


@st.cache_data(ttl=5400, show_spinner=False)
def lark_token(host: str, app_id: str, app_secret: str) -> str:
    r = _lark_api(host, "/auth/v3/tenant_access_token/internal", method="POST",
                  json={"app_id": app_id, "app_secret": app_secret})
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"取 token 失败: {d.get('code')} {d.get('msg')}")
    return d["tenant_access_token"]


def lark_tables(host, token, app_token):
    r = _lark_api(host, f"/bitable/v1/apps/{app_token}/tables?page_size=100", token=token)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"列表失败: {d.get('msg')}")
    return [(t["table_id"], t.get("name", "")) for t in d["data"].get("items", [])]


def lark_records(host, token, app_token, table_id):
    """翻页拉全部记录"""
    items, page = [], None
    while True:
        q = "?page_size=500" + (f"&page_token={page}" if page else "")
        r = _lark_api(host, f"/bitable/v1/apps/{app_token}/tables/{table_id}/records/search{q}",
                      token=token, method="POST", json={})
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"读记录失败: {d.get('msg')}")
        data = d.get("data", {})
        items += data.get("items", [])
        if not data.get("has_more"):
            break
        page = data.get("page_token")
    return items


def lark_download_attachments(host, app_id, app_secret, base_url, cache_dir, attach_field=None):
    """把 Base 里所有附件下载到 cache_dir（已存在则跳过），返回 [(filename, bytes)]"""
    os.makedirs(cache_dir, exist_ok=True)
    app_token = parse_base_url(base_url)
    token = lark_token(host, app_id, app_secret)
    tables = lark_tables(host, token, app_token)
    if not tables:
        raise RuntimeError("这个 Base 里没有找到数据表")
    table_id = tables[0][0]
    recs = lark_records(host, token, app_token, table_id)

    files, todo = [], []
    for rec in recs:
        for fname, val in rec.get("fields", {}).items():
            if attach_field and fname != attach_field:
                continue
            if isinstance(val, list):
                for a in val:
                    if isinstance(a, dict) and a.get("file_token") and a.get("name", "").lower().endswith((".xlsx", ".xls")):
                        todo.append((a["name"], a["file_token"]))
    todo = list({t[1]: t for t in todo}.values())          # 按 token 去重

    local_of = {ftok: os.path.join(cache_dir, f"{ftok[:8]}_{name}") for name, ftok in todo}
    need = [(name, ftok, local_of[ftok]) for name, ftok in todo
            if not (os.path.exists(local_of[ftok]) and os.path.getsize(local_of[ftok]) > 1000)]
    fails = _lark_fetch_many(host, token, need, f"从 Base 下载新增附件（共 {len(todo)} 个，缺 {len(need)} 个）",
                             min_size=1000)
    failed = [f"{nm} → {err}" for nm, err in fails.items()]
    n_new = len(need) - len(fails)
    for name, ftok in todo:
        lp = local_of[ftok]
        if os.path.exists(lp) and os.path.getsize(lp) > 1000:
            with open(lp, "rb") as fh:
                files.append((name, fh.read()))
    st.session_state["_lark_failed"] = failed               # 失败清单交给界面展示，绝不静默丢
    st.session_state["_lark_new"] = n_new
    return files


# ──────────────────────────────────────────────────────────────────────
# 2c. 飞书 wiki / 电子表格直连：达人组数据（跟上面的 Base 附件是两回事——
#     这里读的是表格单元格本身，不是下载附件）
# ──────────────────────────────────────────────────────────────────────
def lark_resolve_doc(host, token, url):
    """
    把一个飞书/Lark 链接解析成 (obj_token, obj_type)。
    /wiki/xxx 链接要先转一次（wiki 节点包着真实文档）；/sheets/xxx、/base/xxx 等链接本身就是真实 token。
    """
    tok = parse_base_url(url)
    if "/wiki/" in (url or ""):
        r = _lark_api(host, f"/wiki/v2/spaces/get_node?token={tok}", token=token)
        d = r.json()
        if d.get("code") != 0:
            raise RuntimeError(f"解析 wiki 节点失败：{d.get('code')} {d.get('msg')}"
                               "（多半是这个应用还没被加为该文档/知识库的协作者）")
        node = d["data"]["node"]
        return node["obj_token"], node["obj_type"]
    kind = ("sheet" if "/sheets/" in url else "bitable" if "/base/" in url else
            "docx" if "/docx/" in url else "doc")
    return tok, kind


def lark_sheet_tabs(host, token, spreadsheet_token):
    """列出电子表格里的全部子表：[{sheet_id, title, row_count, col_count}]"""
    r = _lark_api(host, f"/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query", token=token)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"读取子表列表失败：{d.get('code')} {d.get('msg')}")
    out = []
    for s in d["data"]["sheets"]:
        gp = s.get("grid_properties", {}) or {}
        out.append(dict(sheet_id=s["sheet_id"], title=s.get("title", ""),
                        row_count=gp.get("row_count", 0), col_count=gp.get("column_count", 0)))
    return out


def _col_letters(n: int) -> str:
    """1 -> A, 27 -> AA ……（Excel 风格列号转字母）"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def lark_sheet_raw_values(host, token, spreadsheet_token, sheet_id, row_count, col_count):
    """
    读一个子表的全部单元格原始值（不传 valueRenderOption——传 ToString 会把附件单元格简化成
    {'id':1,'type':'attachment'}，拿不到 fileToken；不传才会给完整的 {fileToken, name, size, ...}）。
    """
    row_count, col_count = max(int(row_count or 1), 1), max(int(col_count or 1), 1)
    rng = f"{sheet_id}!A1:{_col_letters(col_count)}{row_count}"
    r = _lark_api(host, f"/sheets/v2/spreadsheets/{spreadsheet_token}/values/{rng}", token=token)
    d = r.json()
    if d.get("code") != 0:
        raise RuntimeError(f"读取「{sheet_id}」单元格失败：{d.get('code')} {d.get('msg')}")
    return d["data"]["valueRange"].get("values") or []


# 达人组数据保存表的固定结构（2026-09-21 实测确认）：每个月一个子表（"6月"…），
# 行1=表头（日期/Creators数据/Products数据/Vidoes数据/LIVE Streams数据），行2="月数据"（月度汇总附件），
# 行3="周数据"（目前空），行5 起每行一天，B~E 列各放一个 xlsx 附件（当天没数据就是空）。
CREATOR_COL_SECTION = {1: "Creators", 2: "Products", 3: "Videos", 4: "LIVE"}


def lark_creator_scan_cells(host, token, spreadsheet_token):
    """扫描『达人组数据保存』全部月份子表，列出每个附件单元格：月数据/周数据/某天 × 4 个部分。"""
    cells = []
    for tab in lark_sheet_tabs(host, token, spreadsheet_token):
        rows = lark_sheet_raw_values(host, token, spreadsheet_token, tab["sheet_id"],
                                     max(tab["row_count"], 60), max(tab["col_count"], 5))
        for i, row in enumerate(rows, start=1):
            label = row[0] if row else None
            if label is None or i == 1:                          # 跳过表头行、空行
                continue
            for j, sec in CREATOR_COL_SECTION.items():
                cell = row[j] if j < len(row) else None
                if isinstance(cell, list) and cell and isinstance(cell[0], dict) and cell[0].get("fileToken"):
                    a = cell[0]
                    cells.append(dict(month=tab["title"], row=i, label=str(label), section=sec,
                                      file_token=a["fileToken"], name=a.get("text", ""), size=a.get("size", 0)))
    return cells


_MONTH_NUM = {"6月": 6, "7月": 7, "8月": 8, "9月": 9, "10月": 10, "11月": 11, "12月": 12}
_YEAR_HINT = 2026                                                  # 这批数据都在这一年


def _creator_cell_date(month_title: str, label: str):
    """'月数据'/'周数据' -> (None, 'monthly'/'weekly')；'8月10号' -> ('2026-08-10','daily')"""
    if label in ("月数据", "周数据"):
        return None, {"月数据": "monthly", "周数据": "weekly"}[label]
    m = re.match(r"(\d+)月(\d+)号", label)
    if m:
        try:
            return f"{_YEAR_HINT}-{int(m.group(1)):02d}-{int(m.group(2)):02d}", "daily"
        except ValueError:
            return None, "invalid"
    return None, "unknown"


def _to_num(s: pd.Series) -> pd.Series:
    """'$205.33'/'4.11%'/'1,234' -> float；已经是数字就原样返回。"""
    if pd.api.types.is_numeric_dtype(s):
        return s
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False)
                         .str.replace("$", "", regex=False).str.replace("%", "", regex=False)
                         .str.strip().replace({"": None, "-": None}), errors="coerce")


def _parse_duration(s: pd.Series) -> pd.Series:
    """'1m28s'/'5s'/'0s' -> 秒数（float）。"""
    ss = s.astype("string")
    m = ss.str.extract(r"(?:(\d+)m)?(?:(\d+)s)?")
    mins = pd.to_numeric(m[0], errors="coerce").fillna(0)
    secs = pd.to_numeric(m[1], errors="coerce").fillna(0)
    out = mins * 60 + secs
    out[ss.isna()] = np.nan
    return out


def _clean_pid(s: pd.Series) -> pd.Series:
    """Product ID 统一成字符串、去掉尾巴 .0，跟运营数据的 pid 对上号，才能两边联查。"""
    return (s.astype("string").str.strip().str.replace(r"\.0$", "", regex=True))


def _creator_clean_df(raw: pd.DataFrame) -> pd.DataFrame:
    """
    金额/百分比字符串("$205.33"/"4.11%"/"2.72%")转数字；"1m28s"这种时长转成秒；
    Product ID/Video ID/LIVE ID 统一成字符串（Product ID 要跟运营数据的 pid 对得上，才能两边联查）。
    """
    out = raw.copy()
    for col in out.columns:
        cl = str(col)
        if cl in ("Avg. viewing duration",):
            out[col] = _parse_duration(out[col])
        elif cl == "Product ID":
            out[col] = _clean_pid(out[col])
        elif cl in ("Video ID", "LIVE ID"):
            out[col] = out[col].astype("string")
        elif any(k in cl for k in ("GMV", "Refunds", "AOV", "commission", "fee", "GPM", "CTR", "CTOR",
                                    "rate", "Rate", "Engagement")):
            out[col] = _to_num(out[col])
    return out


def lark_creator_fetch_all(host, app_id, app_secret, url, cache_dir):
    """
    达人组数据总入口：解析 wiki 链接 → 扫描 4 个月份子表的附件网格 → 下载（本地有缓存就跳过）→
    每个文件行0=表头/行1=英文说明(跳过)/行2+=数据 → 按 4 个部分(Creators/Products/Videos/LIVE)
    分别拼成一张长表，多加 _date（None=非逐日）/_gran(daily/weekly/monthly)/_month/_src_file 四列。
    返回 ({section: DataFrame}, 扫描到的 cells 明细, 下载失败清单)。
    """
    os.makedirs(cache_dir, exist_ok=True)
    token = lark_token(host, app_id, app_secret)
    obj_token, obj_type = lark_resolve_doc(host, token, url)
    if obj_type != "sheet":
        raise RuntimeError(f"这个链接是「{obj_type}」类型，达人组数据这块目前只认电子表格(sheet)。")
    cells = lark_creator_scan_cells(host, token, obj_token)
    if not cells:
        raise RuntimeError("扫描完了但一个附件都没找到——检查一下表格结构是不是变了。")

    frames = {sec: [] for sec in set(CREATOR_COL_SECTION.values())}
    need = []
    for c in cells:
        local = os.path.join(cache_dir, f"{c['file_token'][:10]}.xlsx")
        if not os.path.exists(local + ".parquet") and not (os.path.exists(local) and os.path.getsize(local) > 500):
            need.append((f"{c['month']} {c['label']} {c['section']}（{c['name']}）", c["file_token"], local))
    dl_fail = _lark_fetch_many(host, token, need, f"下载达人组新增附件（共 {len(cells)} 个，缺 {len(need)} 个）")
    failed, n_new = [f"{nm} → {err}" for nm, err in dl_fail.items()], 0
    prog = st.progress(0.0, text=f"检查达人组数据… 0/{len(cells)}")
    for i, c in enumerate(cells):
        date, gran = _creator_cell_date(c["month"], c["label"])
        local = os.path.join(cache_dir, f"{c['file_token'][:10]}.xlsx")
        parsed_cache = local + ".parquet"                        # 解析结果也存一份，下次不用重新读 xlsx
        try:
            if os.path.exists(parsed_cache):
                raw = pd.read_parquet(parsed_cache)
                frames[c["section"]].append(raw)
                continue
            if not (os.path.exists(local) and os.path.getsize(local) > 500):
                continue                                          # 下载失败的已经记在 failed 里
            raw = pd.read_excel(local, header=0, skiprows=[1]).dropna(how="all")
            n_new += 1
            if raw.empty:
                continue
            raw = _creator_clean_df(raw)
            raw.insert(0, "_date", date)
            raw.insert(1, "_gran", gran)
            raw.insert(2, "_month", _MONTH_NUM.get(c["month"]))
            raw.insert(3, "_src_file", c["name"])
            try:
                raw.to_parquet(parsed_cache, index=False)          # 缓存解析结果，下次同一个附件不用再读 xlsx
            except Exception:                                       # noqa: BLE001
                pass                                                # 写缓存失败不影响本次结果，就是下次会重读一遍
            frames[c["section"]].append(raw)
        except Exception as e:                                    # noqa: BLE001
            failed.append(f"{c['month']} {c['label']} {c['section']} → {type(e).__name__}: {e}")
        prog.progress((i + 1) / len(cells), text=f"检查达人组数据… {i+1}/{len(cells)}（新识别 {n_new} 个）")
    prog.empty()

    out = {}
    for sec, fl in frames.items():
        if fl:
            out[sec] = _creator_slim(pd.concat(fl, ignore_index=True, sort=False))
    return out, cells, failed, {"new": n_new, "checked": pd.Timestamp.now()}


_CR_DROP_COLS = ("_src_file", "Video link")                  # 程序里没用到的长文本，451 个附件时约占 95MB
_CR_CAT_COLS = ("Video title", "Post date", "LIVE title", "LIVE start time", "LIVE end time",
                "Product name", "Product category")         # 同一条视频每天一行、标题重复几十次 → 分类存储只存一份


def _creator_slim(t: pd.DataFrame) -> pd.DataFrame:
    """
    达人组原始表瘦身（Streamlit Cloud 内存有上限，达人附件到 451 个时原始表约 520MB，超了就「Oh no」）：
    去掉没用到的长文本列、日期从文字转成真正的日期、重复很多的文字列用分类存储。数值一个都不动。
    分类列都不拿来 groupby（pandas 2.x 对分类列 groupby 会把所有类别两两组合），只用来显示。
    """
    t = t.drop(columns=[c for c in _CR_DROP_COLS if c in t.columns])
    if "_date" in t.columns:
        t["_date"] = pd.to_datetime(t["_date"], errors="coerce")
    for c in _CR_CAT_COLS:
        if c in t.columns:
            t[c] = t[c].astype("category")
    return t


# ──────────────────────────────────────────────────────────────────────
# 3. 关键列定位 + 派生指标（比率必须用分子/分母重算，不能对比率求平均）
# ──────────────────────────────────────────────────────────────────────
ALL = "All"
K = {
    "gmv":      f"{ALL}::GMV",
    "orders":   f"{ALL}::Orders",
    "sku":      f"{ALL}::SKU orders",
    "items":    f"{ALL}::Items sold",
    "cust":     f"{ALL}::Est. customers",
    "impr":     f"{ALL}::Product impressions",
    "clicks":   f"{ALL}::Product clicks",
    "atc":      f"{ALL}::Add-to-cart count",
    "uimpr":    f"{ALL}::Unique product impressions",
    "uclicks":  f"{ALL}::Unique clicks",
    "refund":   f"{ALL}::Refunds",
    "iref":     f"{ALL}::Items refunded",
    "ship":     f"{ALL}::Shipping fees",
    # 渠道 GMV（来自 All 区段的归因拆分列）
    "ch_slive": f"{ALL}::Seller LIVE-attributed GMV",
    "ch_svid":  f"{ALL}::Seller video-attributed GMV",
    "ch_cre":   f"{ALL}::Creator-attributed GMV",
    "ch_card":  f"{ALL}::Seller product card GMV",
    "atc_users":   f"{ALL}::Add-to-cart users",
    "shop_impr":   f"{ALL}::Shop tab product impressions",
    "shop_clicks": f"{ALL}::Shop tab product clicks",
    "shop_gmv":    f"{ALL}::Shop tab GMV",
    # 渠道曝光 / 点击（四个渠道的曝光加起来 = All 曝光，可以精确拆分）
    "i_cre":  "Affiliate::Product impressions",
    "i_live": "Seller LIVE::Product impressions",
    "i_card": "Seller Product card::Product impressions",
    "i_vid":  "Seller video::Product impressions",
    "c_cre":  "Affiliate::Product clicks",
    "c_live": "Seller LIVE::Product clicks",
    "c_card": "Seller Product card::Product clicks",
    "c_vid":  "Seller video::Product clicks",
    "i_cre_live": "Affiliate::Product impressions (LIVE)",
    "i_cre_vid":  "Affiliate::Product impressions (video)",
    "new_live":   "Affiliate::New LIVE counts",
    "new_vid":    "Affiliate::New video count",
}
CHANNELS = [("达人 Affiliate", "ch_cre"), ("商家自播 LIVE", "ch_slive"),
            ("商品卡 Product card", "ch_card"), ("商家视频 Video", "ch_svid")]
# 同一渠道的曝光 / 点击 key（与 CHANNELS 一一对应）
CH_IMPR = {"达人 Affiliate": "i_cre", "商家自播 LIVE": "i_live",
           "商品卡 Product card": "i_card", "商家视频 Video": "i_vid"}
CH_CLICK = {"达人 Affiliate": "c_cre", "商家自播 LIVE": "c_live",
            "商品卡 Product card": "c_card", "商家视频 Video": "c_vid"}
# 渠道色（已做色盲模拟校验：红绿色盲下 4 色仍可区分）
CH_COLOR = {"达人 Affiliate": "#C2416B", "商家自播 LIVE": "#F08C00",
            "商品卡 Product card": "#4C6EF5", "商家视频 Video": "#1B3A8C"}

# 派生比率：名称 -> (分子key, 分母key, 格式, 分类)
DERIVED = {
    "CTR 点击率":      ("clicks", "impr", "pct", "前端"),
    "唯一CTR":         ("uclicks", "uimpr", "pct", "前端"),
    "加购率 ATC":      ("atc", "clicks", "pct", "前端"),
    "CTOR 点击成单率": ("sku", "clicks", "pct", "前端"),
    "商城页CTR":       ("shop_clicks", "shop_impr", "pct", "前端"),
    "达人渠道CTR":     ("c_cre", "i_cre", "pct", "前端"),
    "商品卡CTR":       ("c_card", "i_card", "pct", "前端"),
    "自播CTR":         ("c_live", "i_live", "pct", "前端"),
    "AOV 客单价":      ("gmv", "sku", "money", "后端"),
    "件/单 连带率":    ("items", "sku", "num", "后端"),
    "退款率":          ("refund", "gmv", "pct", "后端"),
    "退货件率":        ("iref", "items", "pct", "后端"),
    "件均价 ASP":      ("gmv", "items", "money", "后端"),
}


def have(df: pd.DataFrame, key: str) -> bool:
    return K.get(key, key) in df.columns


def col(df: pd.DataFrame, key: str) -> pd.Series:
    c = K.get(key, key)
    return df[c] if c in df.columns else pd.Series(0.0, index=df.index)


PERIODS = {"日": "D", "周": "W-MON", "月": "MS", "季": "QS"}


def to_period(s: pd.Series, gran: str) -> pd.Series:
    """日期 → 所属期的起始日（周 = 周一，月 = 1 号，季 = 季初）。"""
    if gran == "日":
        return s.dt.normalize()
    u = pd.Series(pd.to_datetime(s.dropna().unique()))
    freq = {"周": "W-SUN", "月": "M", "季": "Q"}[gran]
    starts = u.dt.to_period(freq).dt.start_time
    return s.map(dict(zip(u, starts)))


def aggregate(df: pd.DataFrame, gran: str) -> pd.DataFrame:
    """按粒度聚合：计数/金额求和，比率用分子分母重算。"""
    base = [k for k in K.values() if k in df.columns]
    d = df[["日期"] + base].copy()
    d["_p"] = to_period(d["日期"], gran)
    g = d.groupby("_p")[base].sum().reset_index().rename(columns={"_p": "期间"})
    # 天数（用于日均）
    days = d.groupby("_p")["日期"].nunique().reset_index(name="天数").rename(columns={"_p": "期间"})
    g = g.merge(days, on="期间")
    for name, (num, den, fmt, _c) in DERIVED.items():
        n, dn = K.get(num), K.get(den)
        if n in g.columns and dn in g.columns:
            g[name] = np.where(g[dn] > 0, g[n] / g[dn] * (100 if fmt == "pct" else 1), np.nan)
    if K["gmv"] in g.columns:
        g["日均GMV"] = g[K["gmv"]] / g["天数"]
    if K["impr"] in g.columns:
        g["日均曝光"] = g[K["impr"]] / g["天数"]
    return g.sort_values("期间").reset_index(drop=True)


def pct(a, b):
    return (a / b - 1) * 100 if (b not in (0, None) and pd.notna(b) and b != 0) else np.nan


def md(s: str) -> str:
    """Streamlit 的 markdown 会把成对的 $ 当成数学公式（金额会被吃掉、变斜体），要转义。"""
    return s.replace("$", "\\$")


def usd(v, sign: bool = True) -> str:
    """金额：+$1,112 / -$990（正负号放在 $ 前面）。"""
    if pd.isna(v):
        return "—"
    return ("+" if v >= 0 else "-") + f"${abs(v):,.0f}" if sign else f"${v:,.0f}"


def fmt_val(v, kind):
    if pd.isna(v):
        return "—"
    if kind == "money":
        return f"${v:,.2f}"
    if kind == "pct":
        return f"{v:.2f}%"
    if kind == "num":
        return f"{v:,.2f}"
    return f"{v:,.0f}"


# ──────────────────────────────────────────────────────────────────────
# 4. 归因引擎
# ──────────────────────────────────────────────────────────────────────
def funnel_attribution(cur: dict, base: dict) -> pd.DataFrame:
    """GMV = 曝光 × CTR × CTOR × AOV 的对数乘法归因。返回各因子对 ΔGMV 的贡献。"""
    def safe(x):
        return x if (x and x > 0 and np.isfinite(x)) else np.nan

    f_cur = {
        "曝光": safe(cur["impr"]),
        "CTR": safe(cur["clicks"] / cur["impr"] if cur["impr"] else np.nan),
        "CTOR": safe(cur["sku"] / cur["clicks"] if cur["clicks"] else np.nan),
        "AOV": safe(cur["gmv"] / cur["sku"] if cur["sku"] else np.nan),
    }
    f_base = {
        "曝光": safe(base["impr"]),
        "CTR": safe(base["clicks"] / base["impr"] if base["impr"] else np.nan),
        "CTOR": safe(base["sku"] / base["clicks"] if base["clicks"] else np.nan),
        "AOV": safe(base["gmv"] / base["sku"] if base["sku"] else np.nan),
    }
    d_gmv = cur["gmv"] - base["gmv"]
    rows, logs = [], {}
    for k in f_cur:
        if pd.isna(f_cur[k]) or pd.isna(f_base[k]) or f_base[k] == 0:
            logs[k] = np.nan
        else:
            logs[k] = np.log(f_cur[k] / f_base[k])
    tot = np.nansum(list(logs.values()))
    for k, lg in logs.items():
        share = (lg / tot) if (tot and np.isfinite(tot) and tot != 0 and np.isfinite(lg)) else np.nan
        rows.append(dict(因子=k,
                         本期=f_cur[k], 基期=f_base[k],
                         变化率=(f_cur[k] / f_base[k] - 1) * 100 if (f_base[k] and np.isfinite(f_cur[k])) else np.nan,
                         贡献占比=share * 100 if pd.notna(share) else np.nan,
                         贡献金额=share * d_gmv if pd.notna(share) else np.nan))
    return pd.DataFrame(rows)


def channel_attribution(df: pd.DataFrame, cur_mask, base_mask) -> pd.DataFrame:
    rows = []
    for label, key in CHANNELS:
        c = col(df[cur_mask], key).sum()
        b = col(df[base_mask], key).sum()
        rows.append(dict(渠道=label, 本期=c, 基期=b, 变化=c - b, 变化率=pct(c, b)))
    return pd.DataFrame(rows).sort_values("变化")


def link_attribution(df: pd.DataFrame, cur_mask, base_mask, topn=10) -> pd.DataFrame:
    gc = df[cur_mask].groupby(["product_id"]).agg(
        本期=(K["gmv"], "sum"), 名称=("product_name", "last"),
        曝光本期=(K["impr"], "sum")).reset_index()
    gb = df[base_mask].groupby(["product_id"]).agg(
        基期=(K["gmv"], "sum"), 曝光基期=(K["impr"], "sum")).reset_index()
    m = gc.merge(gb, on="product_id", how="outer").fillna({"本期": 0, "基期": 0, "曝光本期": 0, "曝光基期": 0})
    m["名称"] = m["名称"].fillna("(已下架)")
    m["变化"] = m["本期"] - m["基期"]
    m["曝光变化率"] = m.apply(lambda r: pct(r["曝光本期"], r["曝光基期"]), axis=1)
    m = m.sort_values("变化")
    return pd.concat([m.head(topn), m.tail(topn)]).drop_duplicates("product_id")


def _funnel_from(s) -> dict:
    """s = 一段数据各原始列的合计（Series/dict，键是原始列名）→ 漏斗四要素 + 渠道 GMV。分母为 0 给 nan。"""
    g = {k2: float(s.get(K[k2], 0.0)) for k2 in ["gmv", "impr", "clicks", "sku", "orders"]}
    g["ctr"] = g["clicks"] / g["impr"] * 100 if g["impr"] else np.nan
    g["ctor"] = g["sku"] / g["clicks"] * 100 if g["clicks"] else np.nan
    g["aov"] = g["gmv"] / g["sku"] if g["sku"] else np.nan
    for lbl, kk2 in CHANNELS:
        g[lbl] = float(s.get(K[kk2], 0.0))
    return g


def _funnel_of(d: pd.DataFrame) -> dict:
    """把一段数据压成漏斗四要素。"""
    return _funnel_from(d[[c for c in K.values() if c in d.columns]].sum())


def diagnose(cur: dict, base: dict) -> tuple[str, str]:
    """
    对比本期 vs 基期的漏斗，判断『为什么掉/涨』。
    GMV 在跌 → 找跌得最多的因子；GMV 在涨 → 找涨得最多的因子（涨和跌分开，否则突增会被说成『某某下滑』）。
    返回 (主因标签, 人话解释)
    """
    def rel(a, b):
        return (a / b - 1) * 100 if (b and np.isfinite(b) and np.isfinite(a) and b != 0) else np.nan

    up = rel(cur["gmv"], base["gmv"])
    rising = pd.notna(up) and up >= 0
    txt = {
        "曝光": ("曝光上升——平台分发变多/达人内容带量", "曝光下滑——平台分发减少/内容供给断档/链接被降权"),
        "CTR": ("点击率提升——主图/标题/售价更吸引人", "点击率下滑——主图/标题/售价对人群吸引力变弱"),
        "CTOR": ("点击成单率提升——详情页/价格/促销起作用", "点击成单率下滑——详情页/价格/库存/评价环节掉链子"),
        "AOV": ("客单价提升——折扣变浅或多件成交变多", "客单价下滑——折扣加深或成交结构下移"),
    }
    cands = [("曝光", rel(cur["impr"], base["impr"])), ("CTR", rel(cur["ctr"], base["ctr"])),
             ("CTOR", rel(cur["ctor"], base["ctor"])), ("AOV", rel(cur["aov"], base["aov"]))]
    valid = [c for c in cands if pd.notna(c[1])]
    if not valid:
        if base["gmv"] == 0 and cur["gmv"] > 0:
            return "新出单", "基期没有成交（新上架/新起量的链接）。"
        if cur["gmv"] == 0 and base["gmv"] > 0:
            return "停止出单", "本期没有成交（下架/断流）。"
        return "数据不足", "前后期数据不足，无法判断。"
    k, v = (max if rising else min)(valid, key=lambda x: x[1])
    ch = sorted([(lbl, cur.get(lbl, 0) - base.get(lbl, 0)) for lbl, _ in CHANNELS], key=lambda x: x[1])
    lead = ch[-1] if rising else ch[0]
    ch_txt = ""
    if lead[1] != 0:
        ch_txt = f"；渠道上主要是{lead[0].split()[0]}{'多了' if lead[1] > 0 else '少了'} ${abs(lead[1]):,.0f}"
    return k, f"{txt[k][0 if rising else 1]}（{k} {v:+.0f}%）{ch_txt}"


def link_sums(df_: pd.DataFrame, cur_mask, base_mask, base_scale: float = 1.0):
    """
    每条链接在本期 / 基期的全部求和列（外连接：新上架、已下架的链接也在里面，缺的记 0）。
    base_scale：基期是多期合计或天数不同时，基期除以它折算成与本期同口径。
    返回 (本期表, 基期表, 链接名)，行 = product_id，列 = 原始列名。
    """
    cols = [c for c in K.values() if c in df_.columns]
    gc = df_[cur_mask].groupby("product_id")[cols].sum()
    gb = df_[base_mask].groupby("product_id")[cols].sum() / base_scale
    ids = gc.index.union(gb.index)
    names = (df_[cur_mask | base_mask].groupby("product_id")["product_name"].last()
             .reindex(ids).fillna("").map(lambda s: short_name(s) or "(无名称)"))
    dup = names.duplicated(keep=False)                   # 同名链接加 ID 尾号区分，否则图上会叠在一起
    names[dup] = [f"{n} ·{str(p)[-4:]}" for p, n in names[dup].items()]
    return gc.reindex(ids, fill_value=0.0), gb.reindex(ids, fill_value=0.0), names


def sum_contrib(gc, gb, names, key: str) -> pd.DataFrame:
    """
    求和型指标（曝光/点击/GMV/订单…）的链接贡献：全店 Δ = Σ 各链接 Δ（严格相加）。
    占跌量% = 该链接的下跌量 ÷ 所有下跌链接的下跌量之和（有涨有跌时比『占净变化』更能看出谁在拖）。
    """
    c = K[key]
    t = pd.DataFrame({"链接": names, "基期": gb[c], "本期": gc[c]})
    t["Δ"] = t["本期"] - t["基期"]
    t["自身变化%"] = np.where(t["基期"] > 0, (t["本期"] / t["基期"] - 1) * 100, np.nan)
    down, up = t.loc[t["Δ"] < 0, "Δ"].sum(), t.loc[t["Δ"] > 0, "Δ"].sum()
    t["占跌量%"] = np.where(t["Δ"] < 0, t["Δ"] / down * 100 if down else np.nan, np.nan)
    t["占涨量%"] = np.where(t["Δ"] > 0, t["Δ"] / up * 100 if up else np.nan, np.nan)
    chmap = CH_IMPR if key == "impr" else CH_CLICK if key == "clicks" else (
        {lbl: kk for lbl, kk in CHANNELS} if key == "gmv" else None)
    if chmap:
        for lbl, kk in chmap.items():
            if K[kk] in gc.columns:
                t[f"Δ{lbl.split()[0]}"] = gc[K[kk]] - gb[K[kk]]
        chcols = [f"Δ{lbl.split()[0]}" for lbl in chmap if f"Δ{lbl.split()[0]}" in t.columns]
        if chcols:
            sub = t[chcols]
            t["主要渠道"] = np.where(t["Δ"] < 0, sub.idxmin(axis=1), sub.idxmax(axis=1))
            t["主要渠道"] = t["主要渠道"].str.lstrip("Δ")
    t = t[(t["基期"] != 0) | (t["本期"] != 0)]
    return t.sort_values("Δ")


def ratio_contrib(gc, gb, names, ratio: str) -> pd.DataFrame:
    """
    比率型指标（CTR / CTOR / 加购率 / AOV…）的链接贡献，单位 = 百分点（金额型为 $）。
    设全店基期比率 r_b。链接 i 的贡献 = s_ic·(r_ic − r_b) − s_ib·(r_ib − r_b)，
    其中 s = 该链接占全店分母的比重，r = 该链接自身比率。所有链接贡献之和 = 全店比率变化（严格相加）。
    再拆成：比率效应 = s_ic·(r_ic − r_ib)（链接自己变好/变差）
          结构效应 = (s_ic − s_ib)·(r_ib − r_b)（高/低比率的链接流量占比变了）
    """
    num, den, fmt, _ = DERIVED[ratio]
    mult = 100 if fmt == "pct" else 1
    nc, dc, nb_, db = gc[K[num]], gc[K[den]], gb[K[num]], gb[K[den]]
    Dc, Db = dc.sum(), db.sum()
    if not Dc or not Db:
        return pd.DataFrame()
    rb = nb_.sum() / Db
    sc, sb = dc / Dc, db / Db
    ric = np.where(dc > 0, nc / dc.replace(0, np.nan), np.nan)
    rib = np.where(db > 0, nb_ / db.replace(0, np.nan), np.nan)
    contrib = sc * (np.nan_to_num(ric) - rb) - sb * (np.nan_to_num(rib) - rb)
    rate_eff = np.where((dc > 0) & (db > 0), sc * (np.nan_to_num(ric) - np.nan_to_num(rib)), 0.0)
    t = pd.DataFrame({"链接": names, "贡献": contrib * mult, "比率效应": rate_eff * mult,
                      "自身比率·基期": rib * mult, "自身比率·本期": ric * mult,
                      "分母占比·基期%": sb * 100, "分母占比·本期%": sc * 100})
    t["结构效应"] = t["贡献"] - t["比率效应"]
    t = t[(db > 0) | (dc > 0)]
    return t.sort_values("贡献")


DEN_NAME = {"impr": "曝光", "uimpr": "独立曝光", "clicks": "点击", "sku": "SKU 订单", "items": "件数",
            "gmv": "GMV", "shop_impr": "商城页曝光", "i_cre": "达人曝光", "i_card": "商品卡曝光", "i_live": "自播曝光"}


def reason_for(key: str, r: pd.Series, gc_row, gb_row) -> tuple[str, str]:
    """
    一条链接『为什么』变（按所看的指标给原因，避免看曝光却讲成交）：
      流量指标（曝光/点击）→ 掉/涨在哪个渠道，并附同期 GMV 变化；
      比率指标 → 比率效应（链接自己变了）/ 结构效应（它的流量占比变了）；
      成交指标（GMV/订单/加购…）→ 漏斗四因子诊断。
    """
    if key in DERIVED:
        num, den, f_, _ = DERIVED[key]
        u = " pp" if f_ == "pct" else ""
        rate, mix = r.get("比率效应", np.nan), r.get("结构效应", np.nan)
        tag = "比率效应" if abs(rate) >= abs(mix) else "结构效应"
        return tag, (f"自身 {key.split()[0]} {fmt_val(r.get('自身比率·基期'), f_)} → {fmt_val(r.get('自身比率·本期'), f_)}"
                     f"（比率效应 {rate:+.3f}{u}）；{DEN_NAME.get(den, den)}占全店 {r.get('分母占比·基期%', np.nan):.1f}% → "
                     f"{r.get('分母占比·本期%', np.nan):.1f}%（结构效应 {mix:+.3f}{u}）")
    if key in ("impr", "clicks") or (key in K and key not in ("gmv", "orders", "sku", "items", "cust",
                                                               "refund", "comm", "atc")):
        nm = {"impr": "曝光", "clicks": "点击"}.get(key) or dict((k, l) for k, l, _ in PROF["granular"]).get(key, key)
        own = r.get("自身变化%", np.nan)
        txt = f"{nm} {fv(key, r['基期'])} → {fv(key, r['本期'])}（{'新出现' if not r['基期'] else f'{own:+.0f}%'}）"
        ch = r.get("主要渠道", "")
        if ch and f"Δ{ch}" in r.index:
            txt += f"，主要{'掉' if r['Δ'] < 0 else '涨'}在{ch}（{fdelta(key, r[f'Δ{ch}'])}）"
        txt += f"；同期 GMV {usd(float(gc_row.get(K['gmv'], 0) - gb_row.get(K['gmv'], 0)))}"
        verb = ("下滑", "上升") if key in ("impr", "clicks") else ("减少", "增加")
        return f"{nm}{verb[0] if r['Δ'] < 0 else verb[1]}", txt
    return diagnose(_funnel_from(gc_row), _funnel_from(gb_row))


def culprits(df_: pd.DataFrame, cur_mask, base_mask, base_scale: float = 1.0, topn: int = 3,
             by: str = "gmv", direction: int = -1) -> pd.DataFrame:
    """
    找出对 Δ指标 贡献最大的链接 + 原因。
    by：K 里的求和型 key（gmv/impr/clicks/orders…）或 DERIVED 里的比率名（CTR 点击率…）
    direction：<0 找拖累最多的（突降），>0 找贡献最多的（突增）。
    """
    gc, gb, names = link_sums(df_, cur_mask, base_mask, base_scale)
    t = ratio_contrib(gc, gb, names, by) if by in DERIVED else sum_contrib(gc, gb, names, by)
    if t.empty:
        return pd.DataFrame()
    dcol = "贡献" if by in DERIVED else "Δ"
    t = t[t[dcol] < 0] if direction < 0 else t[t[dcol] > 0].iloc[::-1]
    out = []
    for pid, r in t.head(topn).iterrows():
        tag, why = reason_for(by, r, gc.loc[pid], gb.loc[pid])
        row = dict(链接=r["链接"], Δ=r[dcol], 主因=tag, 原因=why, product_id=pid)
        if by not in DERIVED:
            row.update(本期=r["本期"], 基期=r["基期"], 自身变化=r["自身变化%"],
                       占比=r["占跌量%"] if direction < 0 else r["占涨量%"],
                       主要渠道=r.get("主要渠道", ""))
        else:
            row.update(比率效应=r["比率效应"], 结构效应=r["结构效应"])
        out.append(row)
    return pd.DataFrame(out)


def detect_anomalies(s: pd.DataFrame, valcol: str, win: int = 7, k: float = 3.0,
                     pct_thr: float | None = None) -> pd.DataFrame:
    """
    稳健异常检测（只用过去，不偷看未来）：
      基线 = 前 win 期（不含当期）的中位数
      偏离度 z = (当期 − 基线) ÷ (近 2·win 期『正常波动幅度』的 MAD × 1.4826)
      |z| ≥ k，或 |偏离%| ≥ pct_thr，判为异常。
    """
    x = s[valcol].astype(float)
    med = x.shift(1).rolling(win, min_periods=3).median()
    dev = x - med
    mad = dev.abs().shift(1).rolling(win * 2, min_periods=3).median()
    z = dev / (mad.replace(0, np.nan) * 1.4826)
    out = s.copy()
    out["基线"] = med
    out["偏离度"] = z
    out["偏离%"] = (x / med - 1) * 100
    flag = z.abs() >= k
    if pct_thr:
        flag = flag | (out["偏离%"].abs() >= pct_thr)
    out["异常"] = flag.fillna(False) & med.notna()
    return out


# ──────────────────────────────────────────────────────────────────────
# 4b. 逐日环比链条 / 回暖追踪 / 改动事件
# ──────────────────────────────────────────────────────────────────────
def series_of(df: pd.DataFrame, metric_key: str, pid: str | None = None) -> pd.DataFrame:
    """取某指标的逐日序列（可限定单链接）。比率型自动用分子/分母重算。"""
    d = df if pid is None else df[df["product_id"] == pid]
    if metric_key in DERIVED:
        num, den, fmt, _ = DERIVED[metric_key]
        g = d.groupby("日期")[[K[num], K[den]]].sum().reset_index()
        g["值"] = np.where(g[K[den]] > 0, g[K[num]] / g[K[den]] * (100 if fmt == "pct" else 1), np.nan)
        return g[["日期", "值"]]
    c = K.get(metric_key, metric_key)
    if c not in d.columns:
        return pd.DataFrame(columns=["日期", "值"])
    return d.groupby("日期")[c].sum().reset_index().rename(columns={c: "值"})


def daily_chain(s: pd.DataFrame) -> pd.DataFrame:
    """逐日环比链条：每天 vs 前一天，并给出相对区间起点的累计指数。"""
    out = s.sort_values("日期").reset_index(drop=True).copy()
    out["前一日"] = out["值"].shift(1)
    out["日环比%"] = (out["值"] / out["前一日"] - 1) * 100
    base0 = out["值"].iloc[0] if len(out) and out["值"].iloc[0] else np.nan
    out["相对起点%"] = (out["值"] / base0 - 1) * 100 if pd.notna(base0) else np.nan
    out["方向"] = np.where(out["日环比%"] >= 0, "▲", "▼")
    return out


def recovery_after(s: pd.DataFrame, event_date, pre: int = 7, post: int = 14) -> dict:
    """
    大跌后的逐日回暖追踪。
    基线 = 跌前 `pre` 天的**中位数**——均值会被跌前某一天的冲高拉高，
    把『冲高后回落』误判成『持续阴跌』。
    返回：基线、事件后每日 值 / vs 前一天% / vs 基线%、判定、跌前一天是不是冲高。
    """
    s = s.sort_values("日期").reset_index(drop=True).copy()
    s["日环比%"] = s["值"].pct_change() * 100          # 在整条序列上算，事件当天也是 vs 真正的前一天
    ed = pd.Timestamp(event_date)
    before = s[(s["日期"] < ed) & (s["日期"] >= ed - pd.Timedelta(days=pre))]
    after = s[(s["日期"] >= ed) & (s["日期"] <= ed + pd.Timedelta(days=post))].copy()
    if len(before) < 3 or len(after) < 2:
        return {}
    baseline = before["值"].median()
    after["恢复度%"] = (after["值"] / baseline - 1) * 100 if baseline else np.nan
    h = len(after) // 2
    first, second = after["值"].iloc[:h].mean(), after["值"].iloc[h:].mean()
    end = after["恢复度%"].iloc[-1]
    if pd.notna(end) and end >= -5:
        verdict = "✅ 已恢复到跌前水平"
    elif len(after) >= 4 and second > first * 1.03:
        verdict = "🟡 逐步回暖（后半段高于前半段），但还没回到跌前"
    elif len(after) >= 4 and second < first * 0.97:
        verdict = "❌ 继续阴跌"
    else:
        verdict = "➖ 低位横盘，未见回暖"
    ref = before["值"].iloc[:-1].median() if len(before) > 3 else np.nan
    after_spike = bool(pd.notna(ref) and ref > 0 and before["值"].iloc[-1] >= ref * 1.3)

    def at(n):
        r = after[after["日期"] == ed + pd.Timedelta(days=n)]
        return r["恢复度%"].iloc[0] if len(r) else np.nan

    return dict(baseline=baseline, pre_n=len(before), after=after, verdict=verdict,
                after_spike=after_spike, prev_day=before["值"].iloc[-1],
                trough=after["值"].min(), trough_day=after.loc[after["值"].idxmin(), "日期"],
                end_recovery=end, d7=at(7), d14=at(14))


def asp_breaks(g: pd.DataFrame, win: int = 7, thr: float = 15.0, min_items: int = 3) -> list[dict]:
    """
    件单价（GMV ÷ 件数）结构突变 = 改价 / 改促销的痕迹。
    只用出单 ≥ min_items 件的日子（件数太少时单价会乱跳）；
    某天起的 win 个有效日中位数 vs 之前 win 个有效日中位数，相差 ≥ thr% 记一次，两次至少隔 14 天。
    """
    d = g.groupby("日期")[[K["gmv"], K["items"]]].sum()
    d = d[d[K["items"]] >= min_items]
    if len(d) < 3 * win:
        return []
    asp = d[K["gmv"]] / d[K["items"]]
    pre = asp.rolling(win).median().shift(1)          # 之前 win 个有效日（不含当天）
    post = asp[::-1].rolling(win).median()[::-1]      # 当天起 win 个有效日
    ch = (post / pre - 1) * 100
    out, last = [], None
    for dt, v in ch.items():
        if pd.notna(v) and abs(v) >= thr and (last is None or (dt - last).days >= 14):
            out.append(dict(日期=dt, 事件="改价/改促销（件单价突变）",
                            详情=f"件单价 ${pre[dt]:.2f} → ${post[dt]:.2f}（{v:+.0f}%，前后各{win}个出单日中位数）"))
            last = dt
    return out


@st.cache_data(show_spinner=False)
def detect_events(df: pd.DataFrame) -> pd.DataFrame:
    """
    从数据本身自动识别链接『改动事件』：
      · 改名/重组   —— 同一 Product ID 的 product_name 变了（改标题/重组总链会体现在这里）
      · 改价/改促销 —— 件单价出现 ≥15% 的结构性跳变（改价、换折扣、换 promotion 都会留下这个痕迹）
      · 上架        —— 该链接首次出现
      · 断更恢复    —— 连续 ≥3 天零曝光后重新出现
    """
    ev = []
    for pid, g in df.sort_values("日期").groupby("product_id"):
        names = g["product_name"].tolist()
        dates = g["日期"].tolist()
        ev.append(dict(日期=dates[0], product_id=pid, 事件="上架/首次出现",
                       详情=str(names[0])[:60]))
        for i in range(1, len(names)):
            if str(names[i]).strip() != str(names[i - 1]).strip():
                ev.append(dict(日期=dates[i], product_id=pid, 事件="改名/重组",
                               详情=f"{str(names[i-1])[:34]} → {str(names[i])[:34]}"))
        if K["gmv"] in g.columns and K["items"] in g.columns:
            ev += [dict(e, product_id=pid) for e in asp_breaks(g)]
        if K["impr"] in g.columns:
            has = (g[K["impr"]] > 0).tolist()
            run = 0
            for dt, h_ in zip(dates, has):
                if not h_:
                    run += 1
                else:
                    if run >= 3:
                        ev.append(dict(日期=dt, product_id=pid, 事件="断更后恢复曝光",
                                       详情=f"此前连续 {run} 天零曝光"))
                    run = 0
    if not ev:
        return pd.DataFrame(columns=["日期", "product_id", "事件", "详情"])
    return pd.DataFrame(ev).sort_values("日期").reset_index(drop=True)


def change_effect(df_: pd.DataFrame, pid: str, date, win: int = 7) -> dict:
    """
    改动效果：改前 win 天日均 vs 改后 win 天日均（另给改后 2·win 天），缺数据的天按 0 计。
    判定以曝光为准（曝光是索引列），GMV 只做补充：
      曝光 ≥ +10%：成交跟上 = ✅ 有效；成交没跟上 = 🟡 看转化
      曝光 ≤ −10%：成交靠转化撑住 = 🟡 流量问题还在；否则 = ❌ 变差
      曝光 ±10% 内：看成交有没有明显变化
    """
    g = df_[df_["product_id"] == pid]
    cols = [c for c in K.values() if c in g.columns]
    days = pd.date_range(df_["日期"].min(), df_["日期"].max(), freq="D")
    s = g.groupby("日期")[cols].sum().reindex(days, fill_value=0.0)
    d0 = pd.Timestamp(date)
    pre = s[(s.index < d0) & (s.index >= d0 - pd.Timedelta(days=win))]
    post = s[(s.index >= d0) & (s.index < d0 + pd.Timedelta(days=win))]
    post2 = s[(s.index >= d0) & (s.index < d0 + pd.Timedelta(days=2 * win))]
    if len(pre) < 3 or len(post) < 2:
        return {}
    fp, fq = _funnel_from(pre.sum()), _funnel_from(post.sum())
    r = dict(改前天数=len(pre), 改后天数=len(post),
             曝光改前=fp["impr"] / len(pre), 曝光改后=fq["impr"] / len(post),
             GMV改前=fp["gmv"] / len(pre), GMV改后=fq["gmv"] / len(post),
             GMV改后14=post2[K["gmv"]].sum() / len(post2), 曝光改后14=post2[K["impr"]].sum() / len(post2),
             点击改前=fp["clicks"] / len(pre), 点击改后=fq["clicks"] / len(post),
             订单改前=fp["sku"] / len(pre), 订单改后=fq["sku"] / len(post),
             CTR改前=fp["ctr"], CTR改后=fq["ctr"], CTOR改前=fp["ctor"], CTOR改后=fq["ctor"])
    r["曝光变化%"] = pct(r["曝光改后"], r["曝光改前"])
    r["曝光变化%·14天"] = pct(r["曝光改后14"], r["曝光改前"])
    r["GMV变化%"] = pct(r["GMV改后"], r["GMV改前"])
    r["GMV变化%·14天"] = pct(r["GMV改后14"], r["GMV改前"])
    r["基数太小"] = bool(r["GMV改前"] < 20 or r["曝光改前"] < 500)
    gch, ich = r["GMV变化%"], r["曝光变化%"]
    g_up, g_dn = pd.notna(gch) and gch >= 10, pd.notna(gch) and gch <= -10
    if pd.isna(ich):
        r["判定"] = "数据不足（改前没有曝光）"
    elif ich >= 10:
        r["判定"] = "✅ 有效：曝光涨了，成交也跟上" if not g_dn else "🟡 曝光涨了但成交没跟上（看 CTR / CTOR / 价格）"
    elif ich <= -10:
        r["判定"] = "🟡 曝光掉了，靠转化撑住成交（流量问题还在）" if g_up else "❌ 改后变差：曝光掉了"
    elif g_up:
        r["判定"] = "🟡 曝光没动、成交涨了（转化 / 价格起作用）"
    elif g_dn:
        r["判定"] = "🟡 曝光没动、成交掉了（转化变差）"
    else:
        r["判定"] = "➖ 变化不大（曝光、成交都在 ±10% 以内）"
    return r


# 已知的人工改动（来自历史复盘，可在页面上编辑/补充；保存后下次打开还在）
KNOWN_CHANGES = [
    # 2026-09-21 用户逐条确认（8/24 那批：Last Chance / TOP TREND / ANIMALS，其余各自单独的日期）
    {"日期": "2026-08-20", "链接关键词": "TOP TREND", "改动": "改为达人链：Top30 + 新款"},
    {"日期": "2026-08-24", "链接关键词": "LAST CHANCE", "改动": "8/24 修链批次之一"},
    {"日期": "2026-08-24", "链接关键词": "TOP TREND", "改动": "8/24 修链批次之一"},
    {"日期": "2026-08-24", "链接关键词": "ANIMALS", "改动": "8/24 修链批次之一（复盘记录：改后下滑）"},
    {"日期": "2026-09-10", "链接关键词": "FLORAL", "改动": "Floral 系列链调整"},
    {"日期": "2026-09-11", "链接关键词": "TOP TREND", "改动": "在 8/20 基础上加入五折款"},
    {"日期": "2026-09-19", "链接关键词": "DreamWear", "改动": "母链重组：保留高销款、按销量排名"},
]
CHANGES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit", "nailvesta_changes.csv")


# ──────────────────────────────────────────────────────────────────────
# 5. 侧边栏：数据载入
# ──────────────────────────────────────────────────────────────────────
NAV_BOX = st.sidebar.container()          # 数据源开关 + 页面菜单（脚本最后才填，但排在侧边栏最上面）
st.sidebar.title("📊 数据源")
st.sidebar.caption("TikTok Shop · product analysis list 每日导出")

SECRETS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit", "secrets.toml")


def _sec(key: str, default: str = "") -> str:
    """凭证优先级：.streamlit/secrets.toml → 环境变量 → 空"""
    try:
        v = st.secrets.get(key, None)
        if v:
            return str(v)
    except Exception:                                       # noqa: BLE001
        pass
    return os.environ.get(key, default)


def _secret_field(container, label: str, secret_key: str, **kw) -> str:
    """
    密码类输入框，公开部署安全版：Streamlit 的 `type="password"` 只是界面上显示圆点，
    真实值仍会整个发到浏览器（随便打开一次控制台/查看网络请求就能看到），本地自己用没事，
    **公开部署给别人访问就是把 App Secret 泄露给所有访问者**。所以这里输入框永远从空白开始；
    用户没输入时，悄悄在服务器这边用 st.secrets 里存的值，那个值本身不会被塞进前端组件。
    """
    server_val = _sec(secret_key)
    typed = container.text_input(label, value="", type="password",
                                 placeholder="●●●●●●（已用 Secrets 里保存的值）" if server_val else "", **kw)
    return typed or server_val


def _save_secrets(updates: dict) -> None:
    """
    合并写入 secrets.toml（不覆盖已有的其它 key）——这个程序接了两个不同的飞书/Lark 数据源，
    各自的凭证都存在这一个文件里，写的时候要先读旧内容再合并，不能整份覆盖。
    """
    cur: dict = {}
    if os.path.exists(SECRETS_PATH):
        try:
            import tomllib
            with open(SECRETS_PATH, "rb") as fh:
                cur = tomllib.load(fh)
        except Exception:                                   # noqa: BLE001
            cur = {}
    cur.update(updates)
    os.makedirs(os.path.dirname(SECRETS_PATH), exist_ok=True)
    with open(SECRETS_PATH, "w", encoding="utf-8") as fh:
        for k, v in cur.items():
            fh.write(f'{k} = "{v}"\n')


mode = st.sidebar.radio("载入方式",
                        ["飞书 Base 直连", "本地文件夹", "上传 ZIP", "上传多个 xlsx"], index=0)
payloads: list[tuple[str, bytes]] = []

if mode == "飞书 Base 直连":
    st.sidebar.caption("凭证填一次即可（存本地 secrets，不会上传）")

    _def_host = _sec("LARK_HOST", "Lark 国际版 (larksuite.com)")
    hostname = st.sidebar.selectbox(
        "服务器", list(LARK_HOSTS.keys()),
        index=list(LARK_HOSTS.keys()).index(_def_host) if _def_host in LARK_HOSTS else 0)
    base_url = st.sidebar.text_input(
        "Base 链接 / app_token",
        value=_sec("LARK_BASE_URL",
                   "https://qusjdzq12vah.sg.larksuite.com/base/XwZwbY0NTairFWsejC3lr0zjg0t"))
    app_id = _secret_field(st.sidebar, "App ID", "LARK_APP_ID",
                           help="开放平台 → 企业自建应用 → 凭证与基础信息。"
                                "在 Streamlit Cloud 上部署时留空即可，去 Settings → Secrets 配一次，"
                                "以后每次打开都会自动生效，不用手填。")
    app_secret = _secret_field(st.sidebar, "App Secret", "LARK_APP_SECRET")
    def _mask(v: str) -> str:
        """只给看首尾各4位+长度，够她自己核对是不是复制错/粘了引号/漏删 xxx，不会把 Secret 整个亮出来。"""
        return f"{v[:4]}…{v[-4:]}（{len(v)}位）" if len(v) > 10 else f"（{len(v)}位，太短了不太对）"

    if app_id and app_secret:
        st.sidebar.caption(f"🔑 App ID: {_mask(app_id)}　App Secret: {_mask(app_secret)}")
        if "xxx" in app_id.lower() or "xxx" in app_secret.lower():
            st.sidebar.error("⚠️ 里面还留着占位符 `xxx`——Secrets 里没真的换成你的值。")
        if app_id.startswith(('"', "'")) or app_secret.startswith(('"', "'")):
            st.sidebar.error("⚠️ 值开头是引号——粘贴时把 TOML 的 `\"` 也粘进去了，去掉。")
    elif _sec("LARK_APP_ID") or _sec("LARK_APP_SECRET"):
        st.sidebar.caption("⚠️ Secrets 里好像只配了一半（App ID / App Secret 缺一个），检查一下。")
    else:
        st.sidebar.caption("还没配 Secrets（两个框都是空的，这不是 bug，去 Settings → Secrets 配一次）。")
    cache_dir = st.sidebar.text_input("本地缓存目录", value=_sec("LARK_CACHE_DIR", DEFAULT_CACHE),
                                      help="已下载的文件会跳过，不会重复下载")

    with st.sidebar.expander("📖 怎么配（第一次看这里）", expanded=not bool(app_id)):
        st.markdown(
            "**1. 建应用**　开放平台 → 创建**企业自建应用** → 复制 App ID / App Secret\n"
            "- Lark 国际版：https://open.larksuite.com/app\n"
            "- 飞书国内版：https://open.feishu.cn/app\n\n"
            "**2. 开权限**　权限管理 → 添加以下 API 权限：\n"
            "- `bitable:app:readonly` 读多维表格\n"
            "- `drive:drive:readonly` 下载云文件（**附件必须靠它**）\n\n"
            "**3. 发布**　版本管理与发布 → 创建版本 → 申请发布（**不发布权限不生效**）\n\n"
            "**4. 授权 Base**　打开你的 Base → 右上「···」→ **添加文档应用** → 搜到这个应用加进来\n"
            "（漏了这步会报 `no permission` 或拿到 0 条记录）\n\n"
            "**5. 免得每次输**　在下面点「保存凭证到本地」。")

    cta, ctb = st.sidebar.columns(2)
    if cta.button("① 测试连接"):
        for _attempt in (0, 1):                              # token 本地缓存可能比 Lark 实际有效期活得久，遇到失效重试一次
            try:
                _h = LARK_HOSTS[hostname]
                _tk = lark_token(_h, app_id, app_secret)
                _tbls = lark_tables(_h, _tk, parse_base_url(base_url))
                _recs = lark_records(_h, _tk, parse_base_url(base_url), _tbls[0][0]) if _tbls else []
                _atts = [a for r in _recs for v in r.get("fields", {}).values() if isinstance(v, list)
                         for a in v if isinstance(a, dict)
                         and str(a.get("name", "")).lower().endswith((".xlsx", ".xls"))]
                # 真下载 1 个文件，验证 drive 下载权限（只数附件不够，下载权限是单独的）
                _dl = "（没有附件可测）"
                if _atts:
                    _rr = _lark_api(_h, f"/drive/v1/medias/{_atts[0]['file_token']}/download", token=_tk)
                    if _rr.status_code == 200 and len(_rr.content) > 1000:
                        _dl = f"✅ 可下载（试下 {_atts[0]['name']}，{len(_rr.content)//1024} KB）"
                    else:
                        try:
                            _j = _rr.json()
                            _dl = f"❌ {_j.get('code')} {_j.get('msg')}"
                        except Exception:                      # noqa: BLE001
                            _dl = f"❌ HTTP {_rr.status_code}"
                        _dl += "\n\n→ 去开 `drive:drive:readonly` 权限，并重新创建版本发布"
                _box = st.sidebar.success if _dl.startswith("✅") else st.sidebar.warning
                _box(f"✅ 连接成功\n\n表：{len(_tbls)} 个（{_tbls[0][1] if _tbls else '—'}）\n\n"
                     f"记录：{len(_recs)} 条\n\n附件：{len(_atts)} 个\n\n下载测试：{_dl}")
                break
            except NetworkError as e:
                st.sidebar.error(f"🌐 {e}\n\n**这是网络问题，不是权限问题**——通常是 VPN 在切换或网络抖动。"
                                 "直接再点一次；一直失败再检查 VPN 是否开着。")
                break
            except Exception as e:                              # noqa: BLE001
                if _attempt == 0 and "access token" in str(e).lower():
                    lark_token.clear()                            # 本地缓存的 token 可能已经过期，清掉重试
                    continue
                st.sidebar.error(f"🔑 {e}\n\n这是**凭证/权限问题**：App ID/Secret 复制错 / 权限没发布 / "
                                 "应用没加为 Base 协作者 / 服务器选错（你的是 Lark 国际版）")
    if ctb.button("💾 保存凭证"):
        try:
            _save_secrets({"LARK_HOST": hostname, "LARK_BASE_URL": base_url,
                           "LARK_APP_ID": app_id, "LARK_APP_SECRET": app_secret,
                           "LARK_CACHE_DIR": cache_dir})
            st.sidebar.success("已存到 .streamlit/secrets.toml（记得别提交到 git）")
        except Exception as e:                              # noqa: BLE001
            st.sidebar.error(f"保存失败：{e}")

    def _pull_ops(auto: bool = False) -> bool:
        """从 Base 拉日报：本地已有的附件不重下，只下新增的；识别也只识别新增的（见 load_files）。"""
        for _attempt in (0, 1):
            try:
                got = lark_download_attachments(LARK_HOSTS[hostname], app_id, app_secret,
                                                base_url, cache_dir)
                fails = st.session_state.get("_lark_failed", [])
                if got:
                    st.session_state["_payloads"] = got
                    st.session_state["_ops_checked"] = pd.Timestamp.now()
                    if not auto:
                        st.sidebar.success(f"已检查 {len(got)} 个文件，新下载 {st.session_state.get('_lark_new', 0)} 个")
                else:
                    st.sidebar.error("一个附件都没拿到——多半是 `drive:drive:readonly` 权限没开/没发布")
                if fails:
                    st.sidebar.warning(f"⚠️ 有 {len(fails)} 个文件下载失败（已跳过，其余照常分析）。"
                                       "网络抖动造成的话再点一次『拉取全部』即可，已下好的会走缓存不重下。")
                    with st.sidebar.expander(f"失败清单（{len(fails)}）"):
                        st.text("\n".join(fails))
                return bool(got)
            except NetworkError as e:
                st.sidebar.error(f"🌐 {e}\n\n网络问题，再点一次即可（已下好的文件有缓存）。")
                return False
            except Exception as e:                              # noqa: BLE001
                if _attempt == 0 and "access token" in str(e).lower():
                    lark_token.clear()
                    continue
                st.sidebar.error(f"🔑 {e}")
                return False
        return False

    if st.sidebar.button("② 🔄 从 Base 拉取（只补新增）", type="primary",
                         help="检查 Base 里的附件：已经下载、识别过的直接用，只下载和识别新增的。"):
        _pull_ops()
    elif (not st.session_state.get("_payloads") and not st.session_state.get("_ops_auto_tried")
          and app_id and app_secret and base_url):
        # 打开页面自动跑一次：读已识别的结果 + 只补 Base 里新增的文件
        st.session_state["_ops_auto_tried"] = True
        if not _pull_ops(auto=True):
            _local = gather_from_folder(cache_dir) if os.path.isdir(cache_dir) else []
            if _local:
                st.session_state["_payloads"] = _local
                st.sidebar.info(f"连不上飞书，先用本地已下载的 {len(_local)} 个文件（可能缺最新几天）。")
    if os.path.isdir(st.session_state.get("_cache_dir", DEFAULT_CACHE)) and \
            not st.session_state.get("_payloads"):
        if st.sidebar.button("📂 直接用本地缓存（离线）"):
            st.session_state["_payloads"] = gather_from_folder(cache_dir)

elif mode == "本地文件夹":
    folder = st.sidebar.text_input("文件夹路径（含所有日报 xlsx，可含子目录）",
                                   value=os.path.expanduser("~/Desktop"))
    if st.sidebar.button("扫描并载入", type="primary"):
        st.session_state["_payloads"] = gather_from_folder(folder)
elif mode == "上传 ZIP":
    up = st.sidebar.file_uploader("上传 zip", type=["zip"])
    if up is not None:
        st.session_state["_payloads"] = gather_from_zip(up.read())
else:
    ups = st.sidebar.file_uploader("上传 xlsx（可多选）", type=["xlsx"], accept_multiple_files=True)
    if ups:
        st.session_state["_payloads"] = [(u.name, u.read()) for u in ups]

# ── 5b. 达人组数据（独立第二数据源：飞书国内版，跟上面的运营数据 Base 不是同一个租户/平台）──
st.sidebar.markdown("---")
st.sidebar.subheader("🎨 达人组数据")
st.sidebar.caption("达人组自己维护的独立文档，跟上面运营数据同一个 Lark 租户，默认沿用同一份凭证")
DEFAULT_CREATOR_URL = ("https://qusjdzq12vah.sg.larksuite.com/wiki/BL3UwdvEXi6fe3kMWahln8uPgNh"
                       "?from=from_copylink&sheet=WlSXr8")
DEFAULT_CREATOR_CACHE = os.path.expanduser("~/nailvesta_creator_cache")
with st.sidebar.expander("达人组数据连接", expanded=not bool(st.session_state.get("_creator_data"))):
    st.caption("跟上面运营数据是**同一个 Lark 国际版租户**，已确认可以直接用同一个 App ID/Secret——"
              "留空就会自动沿用上面「飞书 Base 直连」里填的凭证。")
    _cr_host_name = st.selectbox(
        "服务器 ", list(LARK_HOSTS.keys()), key="cr_host",
        index=list(LARK_HOSTS.keys()).index(_sec("CREATOR_HOST", "Lark 国际版 (larksuite.com)"))
        if _sec("CREATOR_HOST", "Lark 国际版 (larksuite.com)") in LARK_HOSTS else 0)
    cr_url = st.text_input("文档链接", value=_sec("CREATOR_URL", DEFAULT_CREATOR_URL), key="cr_url")
    cr_app_id = _secret_field(st, "App ID（留空=沿用上面）", "CREATOR_APP_ID", key="cr_app_id")
    cr_app_secret = _secret_field(st, "App Secret（留空=沿用上面）", "CREATOR_APP_SECRET", key="cr_app_secret")
    cr_cache = st.text_input("本地缓存目录", value=_sec("CREATOR_CACHE_DIR", DEFAULT_CREATOR_CACHE),
                             key="cr_cache")
    _cr_id = cr_app_id or app_id if mode == "飞书 Base 直连" else cr_app_id
    _cr_secret = cr_app_secret or app_secret if mode == "飞书 Base 直连" else cr_app_secret
    crc1, crc2 = st.columns(2)
    if crc1.button("💾 保存凭证", key="cr_save"):
        try:
            _save_secrets({"CREATOR_HOST": _cr_host_name, "CREATOR_URL": cr_url,
                           "CREATOR_APP_ID": cr_app_id, "CREATOR_APP_SECRET": cr_app_secret,
                           "CREATOR_CACHE_DIR": cr_cache})
            st.success("已保存")
        except Exception as e:                                # noqa: BLE001
            st.error(f"保存失败：{e}")
    def _pull_creator(force: bool) -> bool:
        """达人组数据：每个附件识别一次就存结果，之后只下载、识别新增的；force = 不用 30 分钟内的检查结果。"""
        for _attempt in (0, 1):
            try:
                secs, cells, failed, cst = lark_creator_fetch_all(LARK_HOSTS[_cr_host_name], _cr_id, _cr_secret,
                                                                   cr_url, cr_cache)
                st.session_state["_creator_data"] = secs
                st.session_state["_creator_ver"] = pd.Timestamp.now().value     # 新数据 → 页面结构要重算
                st.session_state["_creator_rows"] = {k: len(v) for k, v in secs.items()}
                st.session_state["_cr_stats"] = {"cells": len(cells), "ok": len(cells) - len(failed), **cst}
                if force:
                    st.success(f"✅ 扫到 {len(cells)} 个附件，新识别 {cst['new']} 个，其余直接用已识别结果")
                if failed:
                    st.warning(f"⚠️ {len(failed)} 个下载/解析失败（其余照常用）")
                    with st.expander("失败清单"):
                        st.text("\n".join(failed))
                return True
            except NetworkError as e:
                st.error(f"🌐 {e}")
                return False
            except Exception as e:                              # noqa: BLE001
                if _attempt == 0 and "access token" in str(e).lower():
                    lark_token.clear()
                    continue
                st.error(f"🔑 {e}")
                return False
        return False

    if crc2.button("🔄 拉取（只补新增）", key="cr_pull", type="primary",
                   help="重新检查达人组文档：识别过的附件直接用，只下载和识别新增的。"):
        if not _cr_id or not _cr_secret:
            st.error("没有可用的 App ID / Secret——去上面「飞书 Base 直连」填一份，或在这里单独填。")
        else:
            _pull_creator(force=True)
    elif (not st.session_state.get("_creator_data") and not st.session_state.get("_cr_auto_tried")
          and _cr_id and _cr_secret and cr_url):
        st.session_state["_cr_auto_tried"] = True          # 打开页面自动检查一次（增量）
        _pull_creator(force=False)
if st.session_state.get("_creator_data"):
    _cs = st.session_state.get("_cr_stats")
    _rows = st.session_state.get("_creator_rows") or {n: len(d) for n, d in st.session_state["_creator_data"].items()}
    st.sidebar.caption("达人组数据：已载入 " + "、".join(f"{n}{v:,}行" for n, v in _rows.items())
                       + (f"。共 {_cs['cells']} 个附件，{_cs['checked']:%m-%d %H:%M} 检查时新识别 {_cs['new']} 个，"
                          f"其余直接用已识别结果" if _cs else ""))

# 达人组数据的真实列名（2026-09-21 用真实文件核对过，不是猜的）。
# 结构跟运营数据完全不同：这里的表头在第 0 行、第 1 行是英文说明（已在读取时跳过）。
CREATOR_COLS = {
    "Creators": dict(gmv="Creator-attributed GMV", videos="Videos", live="LIVE streams",
                     impr="Product impressions", posted=None, orders="Attributed orders",
                     name="Creator name", creator="Creator name"),
    "Products": dict(gmv="Creator-attributed GMV", videos="Videos", live="LIVE streams",
                     impr="Product impressions", posted="Creators posted content",
                     orders="Attributed orders", name="Product name", creator=None,
                     with_sales="Creators with sales"),
    "Videos": dict(gmv="Creator video-attributed GMV", videos=None, live=None,
                   impr="Video product impressions", posted=None, orders="Video-attributed orders",
                   name="Video title", creator="Creator name", id_col="Video ID"),
    "LIVE": dict(gmv="Creator LIVE-attributed GMV", videos=None, live=None,
                impr="LIVE product impressions", posted=None, orders="LIVE-attributed orders",
                name="LIVE title", creator="Creator name", id_col="LIVE ID"),
}
CREATOR_ORDER = ["Creators", "Products", "Videos", "LIVE"]


def _creator_daily(d: pd.DataFrame) -> pd.DataFrame:
    """只取逐日粒度的行（月度汇总跟逐日混在一张表里，算总量必须先剔掉月度，否则会翻倍）。"""
    t = d[d["_gran"] == "daily"].copy()
    t["_d"] = pd.to_datetime(t["_date"])
    return t


def _creator_period_key(t: pd.DataFrame, gran: str) -> pd.Series:
    if gran == "日":
        return t["_d"].dt.date
    if gran == "周":
        return t["_d"].dt.to_period("W-MON").dt.start_time.dt.date
    return t["_d"].dt.to_period("M").dt.start_time.dt.date


# 达人组·全字段目录：每张表除了 ID/名称/日期以外的字段，自动归到 前端(曝光/互动/内容产出) 或 后端(成交/成本/售后)。
# 跟运营数据的『指标归档』是同一个思路，只是达人组这边维度更细分。
_CREATOR_NONMETRIC = {"_date", "_gran", "_month", "_src_file", "Creator name", "Product name", "Product ID",
                     "Product category", "Video title", "Video ID", "Post date", "Video link", "LIVE title",
                     "LIVE ID", "LIVE start time", "LIVE end time"}


def creator_classify(col: str) -> tuple[str, str]:
    """返回 (前端/后端, 细分类)。英文原始字段 + 程序自己派生的中文字段都要能归类。"""
    c = col.lower()
    if "flat fee" in c:
        return "后端", "成本（该字段目前恒为0）"
    if "commission" in c or "佣金" in col:
        return "后端", "成本"
    if "refund" in c or "退款" in col:
        return "后端", "售后"
    if "with sales" in c or "出单" in col:           # 出单达人 / 出单视频 / 出单链接：内容 → 成交的转化，跟 CTOR 同类
        return "前端", "内容转化"
    if any(k in c for k in ("gmv", "order", "items sold", "aov", "customer", "products sold", "gpm",
                            "avg. gmv")):
        return "后端", "成交"
    if any(k in c for k in ("like", "comment", "share", "engagement", "completion", "duration")):
        return "前端", "互动质量"
    # 曝光/播放/点击类要先判——"Video views"/"Video product impressions" 这类字段名里也带 "video"，
    # 但量级是千万级曝光而不是内容条数，混进"内容产出"会在图上把内容条数(个位数~千)压成一条 0 线。
    if any(k in c for k in ("impression", "view", "click", "ctr", "ctor", "tap-through", "viewer")) or \
            any(k in col for k in ("曝光", "点击", "观众")):
        return "前端", "曝光与转化"
    if any(k in c for k in ("video", "live stream", "sample", "showcase", "posted content")) or \
            any(k in col for k in ("视频数", "直播场数", "直播数", "达人数", "链接数", "寄样", "样品")):
        return "前端", "内容产出"
    return "后端", "其他"


# ──────────────────────────────────────────────────────────────────────
# 3b. 达人组数据集：把 Creators / Products / Videos / LIVE 四张表整理成跟运营数据同构的
#     「日期 × 链接 × 指标」长表（链接汇总 = Products 表，逐链逐日跟运营 Affiliate 区段 GMV 99.6% 一致），
#     另外给出 运营没有的维度：达人 × 天（Creators 表，精确）、视频/直播 × 天（精确）、
#     链接 × 达人 × 天（挂车关系，近似：视频 GMV 是这条视频带来的全部成交，买家可能买了别的链接）。
# ──────────────────────────────────────────────────────────────────────
CR_LINK, CR_CH, CR_VA, CR_LA = "链接汇总", "渠道拆分", "视频挂车", "直播挂车"

K_CR = {
    "gmv": f"{CR_LINK}::Creator-attributed GMV",
    "orders": f"{CR_LINK}::Attributed orders",
    "sku": f"{CR_LINK}::Attributed orders（漏斗分子）",
    "items": f"{CR_LINK}::Creator-attributed items sold",
    "cust": f"{CR_LINK}::Customers",
    "impr": f"{CR_LINK}::Product impressions",
    "clicks": f"{CR_LINK}::Product clicks",
    "refund": f"{CR_LINK}::Refunds",
    "iref": f"{CR_LINK}::Items refunded",
    "comm": f"{CR_LINK}::Est. commission",
    "new_vid": f"{CR_LINK}::Videos",
    "new_live": f"{CR_LINK}::LIVE streams",
    "posted": f"{CR_LINK}::Creators posted content",
    "cws": f"{CR_LINK}::Creators with sales",
    "vws": f"{CR_LINK}::Videos with sales",
    "lws": f"{CR_LINK}::LIVE streams with sales",
    "samp_ct": f"{CR_LINK}::Total sample content",
    "samples": f"{CR_LINK}::Samples shipped",
    "ch_vid": f"{CR_CH}::达人视频 GMV",
    "ch_live": f"{CR_CH}::达人直播 GMV",
    "ch_card": f"{CR_CH}::商品卡及其他 GMV",
    "i_vid": f"{CR_VA}::Video product impressions",
    "i_live": f"{CR_LA}::LIVE product impressions",
    "i_card": f"{CR_CH}::商品卡及其他曝光",
    "c_vid": f"{CR_VA}::Video product clicks",
    "c_live": f"{CR_LA}::Product clicks",
    "c_card": f"{CR_CH}::商品卡及其他点击",
    "v_active": f"{CR_VA}::活跃视频数",
    "v_selling": f"{CR_VA}::出单视频数",
    "v_creators": f"{CR_VA}::活跃视频达人数",
    "v_views": f"{CR_VA}::Video views",
    "v_views_k": f"{CR_VA}::Video views（千次）",
    "v_likes": f"{CR_VA}::Likes",
    "v_comments": f"{CR_VA}::Comments",
    "v_shares": f"{CR_VA}::Shares",
    "v_eng": f"{CR_VA}::赞评转合计",
    "v_comp_w": f"{CR_VA}::完播加权分子",
    "l_active": f"{CR_LA}::活跃直播场数",
    "l_viewers": f"{CR_LA}::LIVE product viewers",
    "l_room": f"{CR_LA}::Impressions",
    "l_likes": f"{CR_LA}::Likes",
    "l_comments": f"{CR_LA}::Comments",
    "l_shares": f"{CR_LA}::Shares",
    "l_dur_w": f"{CR_LA}::观看时长加权分子",
}
# 链接汇总（Products 表）字段 → K_CR 短名
_CR_P_MAP = {"Creator-attributed GMV": "gmv", "Attributed orders": "orders",
             "Creator-attributed items sold": "items", "Customers": "cust",
             "Product impressions": "impr", "Product clicks": "clicks", "Refunds": "refund",
             "Items refunded": "iref", "Est. commission": "comm", "Videos": "new_vid",
             "LIVE streams": "new_live", "Creators posted content": "posted", "Creators with sales": "cws",
             "Videos with sales": "vws", "LIVE streams with sales": "lws",
             "Total sample content": "samp_ct", "Samples shipped": "samples"}
# 不可加总的原始字段（比率/均值），全量表里按加权或均值给出
_CR_AVG_FIELDS = {"CTR", "CTOR", "AOV", "Engagement", "Completion rate", "Tap-through rate", "Video GPM",
                  "Show GPM", "Avg. GMV per customer", "Avg. viewing duration"}

DERIVED_CR = {
    "CTR 点击率": ("clicks", "impr", "pct", "前端"),
    "CTOR 点击成单率": ("sku", "clicks", "pct", "前端"),
    "视频挂车CTR": ("c_vid", "i_vid", "pct", "前端"),
    "直播挂车CTR": ("c_live", "i_live", "pct", "前端"),
    "视频互动率": ("v_eng", "v_views", "pct", "前端"),
    "视频完播率": ("v_comp_w", "v_views", "pct", "前端"),
    "直播平均观看时长(秒)": ("l_dur_w", "l_viewers", "num", "前端"),
    "发布达人出单率": ("cws", "posted", "pct", "前端"),
    "视频出单率": ("v_selling", "v_active", "pct", "前端"),
    "样品转化率": ("samp_ct", "samples", "pct", "前端"),
    "每千次播放GMV": ("ch_vid", "v_views_k", "money", "后端"),
    "AOV 客单价": ("gmv", "sku", "money", "后端"),
    "件均价 ASP": ("gmv", "items", "money", "后端"),
    "件/单 连带率": ("items", "sku", "num", "后端"),
    "退款率": ("refund", "gmv", "pct", "后端"),
    "退货件率": ("iref", "items", "pct", "后端"),
    "佣金率": ("comm", "gmv", "pct", "后端"),
}
CHANNELS_CR = [("达人视频", "ch_vid"), ("达人直播", "ch_live"), ("商品卡及其他", "ch_card")]
CH_IMPR_CR = {"达人视频": "i_vid", "达人直播": "i_live", "商品卡及其他": "i_card"}
CH_CLICK_CR = {"达人视频": "c_vid", "达人直播": "c_live", "商品卡及其他": "c_card"}
CH_COLOR_CR = {"达人视频": "#1B3A8C", "达人直播": "#F08C00", "商品卡及其他": "#4C6EF5"}
DEN_NAME_CR = {"impr": "曝光", "clicks": "点击", "sku": "订单", "items": "件数", "gmv": "GMV",
               "i_vid": "视频挂车曝光", "i_live": "直播挂车曝光", "v_views": "视频播放", "v_active": "活跃视频",
               "posted": "发布达人", "samples": "寄样", "l_viewers": "直播观众", "v_views_k": "千次播放"}
SCORE_CR = [
    ("前端", "曝光", "impr", "count"), ("前端", "点击", "clicks", "count"),
    ("前端", "CTR 点击率", "CTR 点击率", "pct"), ("前端", "CTOR 点击成单率", "CTOR 点击成单率", "pct"),
    ("前端", "视频挂车曝光", "i_vid", "count"), ("前端", "直播挂车曝光", "i_live", "count"),
    ("前端", "商品卡及其他曝光", "i_card", "count"),
    ("前端", "视频挂车CTR", "视频挂车CTR", "pct"), ("前端", "直播挂车CTR", "直播挂车CTR", "pct"),
    ("前端", "新发视频（按链接累计）", "new_vid", "count"), ("前端", "新开直播（按链接累计）", "new_live", "count"),
    ("前端", "发布达人（按链接累计）", "posted", "count"), ("前端", "出单达人（按链接累计）", "cws", "count"),
    ("前端", "发布达人出单率", "发布达人出单率", "pct"),
    ("前端", "活跃视频（按链接累计）", "v_active", "count"), ("前端", "出单视频（按链接累计）", "v_selling", "count"),
    ("前端", "视频出单率", "视频出单率", "pct"),
    ("前端", "视频播放", "v_views", "count"), ("前端", "视频互动率", "视频互动率", "pct"),
    ("前端", "视频完播率", "视频完播率", "pct"),
    ("前端", "直播观众", "l_viewers", "count"), ("前端", "直播平均观看时长(秒)", "直播平均观看时长(秒)", "num"),
    ("前端", "寄样", "samples", "count"), ("前端", "样品产出内容", "samp_ct", "count"),
    ("前端", "样品转化率", "样品转化率", "pct"),
    ("后端", "GMV", "gmv", "money"), ("后端", "订单", "orders", "count"), ("后端", "件数", "items", "count"),
    ("后端", "买家数", "cust", "count"), ("后端", "AOV 客单价", "AOV 客单价", "money"),
    ("后端", "件均价 ASP", "件均价 ASP", "money"), ("后端", "件/单 连带率", "件/单 连带率", "num"),
    ("后端", "达人视频 GMV", "ch_vid", "money"), ("后端", "达人直播 GMV", "ch_live", "money"),
    ("后端", "商品卡及其他 GMV", "ch_card", "money"), ("后端", "每千次播放GMV", "每千次播放GMV", "money"),
    ("后端", "预估佣金", "comm", "money"), ("后端", "佣金率", "佣金率", "pct"),
    ("后端", "退款额（有滞后，近期偏低）", "refund", "money"), ("后端", "退款率（有滞后）", "退款率", "pct"),
]
METRIC_PICK_CR = {"曝光": "impr", "GMV": "gmv", "点击": "clicks", "订单": "orders", "件数": "items",
                  "新发视频": "new_vid", "发布达人": "posted", "出单达人": "cws", "视频播放": "v_views",
                  "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率", "AOV 客单价": "AOV 客单价",
                  "视频互动率": "视频互动率", "视频完播率": "视频完播率", "佣金率": "佣金率"}


def _cr_days(t: pd.DataFrame) -> pd.DataFrame:
    """只取逐日行（月度汇总混在同一张表里，不剔会翻倍），日期转 Timestamp。"""
    t = t[t["_gran"] == "daily"].copy()
    t["日期"] = pd.to_datetime(t["_date"])
    return t


def _f(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").astype("float64").fillna(0.0)


def _anchor_explode(t: pd.DataFrame) -> pd.DataFrame:
    """Videos/LIVE 的 Product ID 可能是逗号分隔的多个链接（一条内容挂多个商品）→ 一行拆成每个链接一行，
    可加总的量按 1/n 平摊（_w），计数类（这条链被几条视频挂了）每条链接都记 1。"""
    ids = t["Product ID"].astype("string").fillna("").str.split(",")
    t = t.assign(_pids=ids, _n=ids.str.len().clip(lower=1)).explode("_pids")
    t["product_id"] = t["_pids"].astype("string").str.strip()
    t = t[t["product_id"].notna() & (t["product_id"] != "")].copy()
    t["_w"] = 1.0 / t["_n"].astype(float)
    return t


def _cr_kind(field: str) -> str:
    f = field.lower()
    if field in _CR_AVG_FIELDS or "率" in field or "rate" in f:
        return "比率"
    if "item" in f:                                          # Items refunded 是件数不是金额
        return "计数"
    if any(k in f for k in ("gmv", "refund", "commission", "flat fee", "aov", "gpm")):
        return "金额"
    return "计数"


# 达人组四个 Excel：中文名、去重 ID 列、判定「出单」的列、名词、GMV/订单/点击列、比率字段的加权列
CR_TABLES = {
    "Creators": ("达人表", "Creator name", "Attributed orders", "达人", "Creator-attributed GMV",
                 "Attributed orders", None, {"CTR": "Product impressions"}, "Product impressions"),
    "Products": ("链接表", "Product ID", "Creator-attributed GMV", "链接", "Creator-attributed GMV",
                 "Attributed orders", "Product clicks", {"CTR": "Product impressions"}, "Product impressions"),
    "Videos": ("视频表", "Video ID", "Creator video-attributed GMV", "视频", "Creator video-attributed GMV",
               "Video-attributed orders", "Video product clicks",
               {"CTR": "Video product impressions", "Avg. GMV per customer": "Video-attributed orders"},
               "Video views"),
    "LIVE": ("直播表", "LIVE ID", "Creator LIVE-attributed GMV", "直播", "Creator LIVE-attributed GMV",
             "LIVE-attributed orders", "Product clicks",
             {"CTR": "LIVE product impressions", "Tap-through rate": "Impressions", "Show GPM": "Impressions",
              "Avg. GMV per customer": "LIVE-attributed orders"}, "LIVE product viewers"),
}
CR_PCT_FIELDS = {"CTR", "CTOR", "Engagement", "Completion rate", "Tap-through rate"}
CR_FIELD_CN = {
    "Creator-attributed GMV": "达人带货 GMV", "Creator LIVE-attributed GMV": "达人直播 GMV",
    "Creator video-attributed GMV": "达人视频 GMV", "Affiliate product card-attributed GMV": "商品卡 GMV",
    "Refunds": "退款金额", "Items refunded": "退款件数", "Attributed orders": "订单数",
    "Creator-attributed items sold": "售出件数", "AOV": "客单价", "CTOR": "点击成单率", "CTR": "点击率",
    "LIVE streams": "直播场数（新开）", "Videos": "视频数（新发）", "Total sample content": "样品产出内容数",
    "Samples shipped": "寄样数", "Products added to showcase": "加入橱窗的商品数",
    "Product impressions": "商品曝光", "Product clicks": "商品点击", "Video views": "视频播放",
    "Customers": "买家数", "Products sold": "售出商品数", "Est. commission": "预估佣金",
    "Est. flat fee": "预估固定费用（恒为 0）", "Videos with sales": "出单视频数", "LIVE streams with sales": "出单直播数",
    "Creators posted content": "发布内容的达人数", "Creators with sales": "出单达人数",
    "Video-attributed orders": "视频订单数", "Video-attributed items sold": "视频售出件数", "Likes": "点赞",
    "Comments": "评论", "Shares": "分享", "Video product impressions": "视频商品曝光",
    "Video product clicks": "视频商品点击", "Completion rate": "完播率", "Video GPM": "千次播放 GMV",
    "Engagement": "互动率", "Avg. GMV per customer": "人均 GMV", "LIVE-attributed items sold": "直播售出件数",
    "LIVE-attributed orders": "直播订单数", "Avg. viewing duration": "平均观看时长（秒）",
    "LIVE product impressions": "直播商品曝光", "Impressions": "直播间曝光", "Show GPM": "千次曝光 GMV",
    "LIVE product viewers": "直播商品观众", "Tap-through rate": "点进率",
}


def cr_agg(T: pd.DataFrame, tname: str, key: pd.Series) -> pd.DataFrame:
    """一张达人 Excel 的全部数值字段按 key（日期 / 周 / 月 / 本期·上一段…）汇总，列顺序 = Excel 原始顺序。
    可加的字段求和；比率/均值字段按曝光、播放、观众等加权重算（客单价、CTOR 用合计重算）；
    最后两列是这一组里的去重个数：表里出现的 X 数、出单 X 数。"""
    _, idc, sellc, noun, gmvc, ordc, clkc, wmap, wdef = CR_TABLES[tname]
    key = pd.Series(np.asarray(key), index=T.index)
    num_cols = [c for c in T.columns if c not in _CREATOR_NONMETRIC and c != "日期" and not c.startswith("_")
                and pd.api.types.is_numeric_dtype(T[c])]
    out = {}
    for c in num_cols:
        if _cr_kind(c) != "比率":
            out[c] = _f(T[c]).groupby(key).sum()
        elif c == "AOV":
            out[c] = _f(T[gmvc]).groupby(key).sum() / _f(T[ordc]).groupby(key).sum().replace(0, np.nan)
        elif c == "CTOR":
            clicks = _f(T[clkc]) if clkc else _f(T["CTR"]) / 100 * _f(T["Product impressions"])
            out[c] = _f(T[ordc]).groupby(key).sum() / clicks.groupby(key).sum().replace(0, np.nan) * 100
        else:
            wt = _f(T[wmap.get(c, wdef)])
            out[c] = (_f(T[c]) * wt).groupby(key).sum() / wt.groupby(key).sum().replace(0, np.nan)
    g = pd.DataFrame(out)
    sell = _f(T[sellc]) > 0
    g[f"表里出现的{noun}数"] = T.groupby(key)[idc].nunique()
    g[f"出单{noun}数"] = T[sell].groupby(key[sell])[idc].nunique().reindex(g.index, fill_value=0)
    return g


def build_creator_base(cd: dict) -> dict:
    """达人组四张表 → 分析用的全部结构（只依赖达人数据本身；按链接的渠道 GMV 另外用运营数据校准）。"""
    out = {}
    P = _cr_days(cd["Products"]) if "Products" in cd else pd.DataFrame()
    V = _cr_days(cd["Videos"]) if "Videos" in cd else pd.DataFrame()
    L = _cr_days(cd["LIVE"]) if "LIVE" in cd else pd.DataFrame()
    C = _cr_days(cd["Creators"]) if "Creators" in cd else pd.DataFrame()

    # ① 链接 × 天（链接汇总，精确）
    link = pd.DataFrame({"日期": P["日期"].values, "product_id": P["Product ID"].astype(str).values,
                         "product_name": P["Product name"].astype(str).values})
    for raw, key in _CR_P_MAP.items():
        link[K_CR[key]] = _f(P[raw]).values if raw in P.columns else 0.0
    link[K_CR["sku"]] = link[K_CR["orders"]]
    for raw in ("CTR", "CTOR", "Est. flat fee"):                    # 原始比率/恒 0 字段，只进目录不参与加总
        if raw in P.columns:
            link[f"{CR_LINK}::{raw}"] = _f(P[raw]).values
    link = link.groupby(["日期", "product_id"], as_index=False).agg(
        {**{c: "sum" for c in link.columns if c not in ("日期", "product_id", "product_name")},
         "product_name": "last"})

    # ② 视频挂车 × 链接 × 天
    ve = _anchor_explode(V) if len(V) else pd.DataFrame()
    if len(ve):
        w = ve["_w"]
        ve = ve.assign(
            _gmv=_f(ve["Creator video-attributed GMV"]) * w, _ord=_f(ve["Video-attributed orders"]) * w,
            _impr=_f(ve["Video product impressions"]) * w, _clk=_f(ve["Video product clicks"]) * w,
            _views=_f(ve["Video views"]) * w, _likes=_f(ve["Likes"]) * w, _com=_f(ve["Comments"]) * w,
            _shr=_f(ve["Shares"]) * w, _compw=_f(ve["Completion rate"]) / 100 * _f(ve["Video views"]) * w,
            _sell=np.where(_f(ve["Creator video-attributed GMV"]) > 0, ve["Video ID"].astype("string"), pd.NA))
        va = ve.groupby(["日期", "product_id"]).agg(
            i=("_impr", "sum"), c=("_clk", "sum"), views=("_views", "sum"), likes=("_likes", "sum"),
            com=("_com", "sum"), shr=("_shr", "sum"), compw=("_compw", "sum"),
            active=("Video ID", "nunique"), selling=("_sell", "nunique"), creators=("Creator name", "nunique"),
            gmv_anchor=("_gmv", "sum")).reset_index()
        va = va.rename(columns={"i": K_CR["i_vid"], "c": K_CR["c_vid"], "views": K_CR["v_views"],
                                "likes": K_CR["v_likes"], "com": K_CR["v_comments"], "shr": K_CR["v_shares"],
                                "compw": K_CR["v_comp_w"], "active": K_CR["v_active"],
                                "selling": K_CR["v_selling"], "creators": K_CR["v_creators"],
                                "gmv_anchor": "_v_gmv_anchor"})
        link = link.merge(va, on=["日期", "product_id"], how="left")
    # ③ 直播挂车 × 链接 × 天
    le = _anchor_explode(L) if len(L) else pd.DataFrame()
    if len(le):
        w = le["_w"]
        le = le.assign(
            _gmv=_f(le["Creator LIVE-attributed GMV"]) * w, _impr=_f(le["LIVE product impressions"]) * w,
            _clk=_f(le["Product clicks"]) * w, _view=_f(le["LIVE product viewers"]) * w,
            _room=_f(le["Impressions"]) * w, _likes=_f(le["Likes"]) * w, _com=_f(le["Comments"]) * w,
            _shr=_f(le["Shares"]) * w,
            _durw=_f(le["Avg. viewing duration"]) * _f(le["LIVE product viewers"]) * w)
        la = le.groupby(["日期", "product_id"]).agg(
            i=("_impr", "sum"), c=("_clk", "sum"), view=("_view", "sum"), room=("_room", "sum"),
            likes=("_likes", "sum"), com=("_com", "sum"), shr=("_shr", "sum"), durw=("_durw", "sum"),
            active=("LIVE ID", "nunique"), gmv_anchor=("_gmv", "sum")).reset_index()
        la = la.rename(columns={"i": K_CR["i_live"], "c": K_CR["c_live"], "view": K_CR["l_viewers"],
                                "room": K_CR["l_room"], "likes": K_CR["l_likes"], "com": K_CR["l_comments"],
                                "shr": K_CR["l_shares"], "durw": K_CR["l_dur_w"], "active": K_CR["l_active"],
                                "gmv_anchor": "_l_gmv_anchor"})
        link = link.merge(la, on=["日期", "product_id"], how="left")
    for k in ("i_vid", "c_vid", "v_views", "v_likes", "v_comments", "v_shares", "v_comp_w", "v_active",
              "v_selling", "v_creators", "i_live", "c_live", "l_viewers", "l_room", "l_likes", "l_comments",
              "l_shares", "l_dur_w", "l_active"):
        c = K_CR[k]
        link[c] = _f(link[c]) if c in link.columns else 0.0
    for c in ("_v_gmv_anchor", "_l_gmv_anchor"):
        link[c] = _f(link[c]) if c in link.columns else 0.0
    link[K_CR["v_views_k"]] = link[K_CR["v_views"]] / 1000
    link[K_CR["v_eng"]] = link[K_CR["v_likes"]] + link[K_CR["v_comments"]] + link[K_CR["v_shares"]]
    # 商品卡及其他曝光/点击 = 链接总量 − 视频挂车 − 直播挂车（挂车量来自内容表，口径略有出入时截到 0）
    link[K_CR["i_card"]] = (link[K_CR["impr"]] - link[K_CR["i_vid"]] - link[K_CR["i_live"]]).clip(lower=0)
    link[K_CR["c_card"]] = (link[K_CR["clicks"]] - link[K_CR["c_vid"]] - link[K_CR["c_live"]]).clip(lower=0)
    out["link"] = link.sort_values(["日期", "product_id"]).reset_index(drop=True)

    # ④ 链接 × 达人 × 天（挂车关系，近似）
    parts = []
    if len(ve):
        ve["_new"] = (pd.to_datetime(ve["Post date"], format="%m/%d/%Y %H:%M", errors="coerce").dt.normalize()
                      == ve["日期"])
        parts.append(ve.groupby(["日期", "product_id", "Creator name"]).agg(
            视频GMV=("_gmv", "sum"), 视频订单=("_ord", "sum"), 挂车曝光=("_impr", "sum"), 播放=("_views", "sum"),
            活跃视频=("Video ID", "nunique"), 新发视频=("_new", "sum")).reset_index())
    if len(le):
        parts.append(le.groupby(["日期", "product_id", "Creator name"]).agg(
            直播GMV=("_gmv", "sum"), 挂车曝光=("_impr", "sum"), 直播场数=("LIVE ID", "nunique")).reset_index())
    if parts:
        lc = pd.concat(parts, ignore_index=True).groupby(["日期", "product_id", "Creator name"], as_index=False).sum()
        for c in ("视频GMV", "直播GMV", "视频订单", "挂车曝光", "播放", "活跃视频", "新发视频", "直播场数"):
            lc[c] = _f(lc[c]) if c in lc.columns else 0.0
        lc["挂车GMV"] = lc["视频GMV"] + lc["直播GMV"]
        out["lc"] = lc.rename(columns={"Creator name": "达人"})
    else:
        out["lc"] = pd.DataFrame(columns=["日期", "product_id", "达人", "挂车GMV"])

    # ⑤ 达人 × 天（Creators 表，精确，全部链接合计）
    if len(C):
        cr = pd.DataFrame({"日期": C["日期"].values, "达人": C["Creator name"].astype(str).values})
        for raw, nm in [("Creator-attributed GMV", "GMV"), ("Creator video-attributed GMV", "视频GMV"),
                        ("Creator LIVE-attributed GMV", "直播GMV"),
                        ("Affiliate product card-attributed GMV", "商品卡GMV"), ("Attributed orders", "订单"),
                        ("Creator-attributed items sold", "件数"), ("Refunds", "退款"),
                        ("Product impressions", "曝光"), ("Video views", "播放"), ("Videos", "新发视频"),
                        ("LIVE streams", "新开直播"), ("Samples shipped", "寄样"),
                        ("Total sample content", "样品内容"), ("Products added to showcase", "加橱窗"),
                        ("Customers", "买家"), ("Est. commission", "佣金")]:
            cr[nm] = _f(C[raw]).values if raw in C.columns else 0.0
        cr["点击"] = (_f(C["CTR"]) / 100 * _f(C["Product impressions"])).values if "CTR" in C.columns else 0.0
        out["creator_day"] = cr.groupby(["日期", "达人"], as_index=False).sum()
    else:
        out["creator_day"] = pd.DataFrame(columns=["日期", "达人", "GMV"])

    # ⑥ 内容（视频 / 直播）× 天，精确
    cont = []
    if len(V):
        cont.append(pd.DataFrame({
            "日期": V["日期"].values, "类型": "视频", "内容ID": V["Video ID"].astype(str).values,
            "标题": V["Video title"].astype(str).str.slice(0, 80).values, "达人": V["Creator name"].astype(str).values,
            "挂车链接": V["Product ID"].astype(str).values, "GMV": _f(V["Creator video-attributed GMV"]).values,
            "订单": _f(V["Video-attributed orders"]).values, "曝光": _f(V["Video product impressions"]).values,
            "点击": _f(V["Video product clicks"]).values, "播放/观众": _f(V["Video views"]).values,
            "赞": _f(V["Likes"]).values, "评": _f(V["Comments"]).values, "转": _f(V["Shares"]).values,
            "完播率%": _f(V["Completion rate"]).values, "发布时间": V["Post date"].astype(str).values}))
    if len(L):
        cont.append(pd.DataFrame({
            "日期": L["日期"].values, "类型": "直播", "内容ID": L["LIVE ID"].astype(str).values,
            "标题": L["LIVE title"].astype(str).str.slice(0, 80).values, "达人": L["Creator name"].astype(str).values,
            "挂车链接": L["Product ID"].astype(str).values, "GMV": _f(L["Creator LIVE-attributed GMV"]).values,
            "订单": _f(L["LIVE-attributed orders"]).values, "曝光": _f(L["LIVE product impressions"]).values,
            "点击": _f(L["Product clicks"]).values, "播放/观众": _f(L["LIVE product viewers"]).values,
            "赞": _f(L["Likes"]).values, "评": _f(L["Comments"]).values, "转": _f(L["Shares"]).values,
            "完播率%": np.nan, "发布时间": L["LIVE start time"].astype(str).values}))
    out["content"] = pd.concat(cont, ignore_index=True) if cont else pd.DataFrame()
    for c in ("标题", "发布时间", "挂车链接"):          # 同一条内容每天一行，这些文字重复几十次 → 分类存储（只显示、不当 groupby 键）
        if c in out["content"]:
            out["content"][c] = out["content"][c].astype("category")

    # ⑦ 全量逐日表：三张原始表（达人/视频/直播）的每一个数值字段按天汇总（链接表的字段在 link 里已有，只补去重数）
    dt = []
    for tname, T in (("Creators", C), ("Videos", V), ("LIVE", L)):
        if len(T):
            g = cr_agg(T, tname, T["日期"])
            g.columns = [f"{CR_TABLES[tname][0]}::{c}" for c in g.columns]
            dt.append(g)
    if len(P):
        g = cr_agg(P, "Products", P["日期"])[["表里出现的链接数", "出单链接数"]]
        g.columns = [f"链接表::{c}" for c in g.columns]
        dt.append(g)
    out["daytot"] = pd.concat(dt, axis=1).fillna(0.0).rename_axis("日期").reset_index() if dt \
        else pd.DataFrame(columns=["日期"])
    # ⑧ 原始逐日表 + 他们另外上传的「月数据」Excel（逐日 / 逐周 / 逐月 / 自选区间对比页用）
    # 「多粒度对比」用的原始字段表：只按日期汇总数值 + 数去重个数，标题 / 发布时间这类长文本用不到，不留在内存里
    out["raw"] = {k: T.drop(columns=[c for c in _CR_CAT_COLS if c in T.columns])
                  for k, T in (("Creators", C), ("Products", P), ("Videos", V), ("LIVE", L)) if len(T)}
    out["raw_monthly"] = {k: cd[k][cd[k]["_gran"] == "monthly"].copy() for k in cd
                          if (cd[k]["_gran"] == "monthly").any()}
    return out


def creator_catalog_full(link: pd.DataFrame, daytot: pd.DataFrame) -> pd.DataFrame:
    """达人侧指标归档：跟运营的 catalog 同结构（列号/区段/指标/key/分类/类型），外加『层级』——
    『链接』= 可以按链接归因的字段；『全量』= 只有每天合计（达人表/视频表/直播表的原始字段，不带链接维度）。"""
    rows = []
    hidden = {K_CR["sku"], K_CR["v_views_k"], K_CR["v_comp_w"], K_CR["l_dur_w"], K_CR["v_eng"]}
    for c in link.columns:
        if c in ("日期", "product_id", "product_name") or c.startswith("_") or c in hidden:
            continue
        sec, name = c.split("::", 1)
        rows.append(dict(区段=sec, 指标=name, key=c, 分类=creator_classify(name)[0],
                         细分类=creator_classify(name)[1], 类型=_cr_kind(name), 层级="链接"))
    for c in daytot.columns:
        if c == "日期":
            continue
        sec, name = c.split("::", 1)
        k = "均值" if _cr_kind(name) == "比率" else _cr_kind(name)
        rows.append(dict(区段=sec, 指标=name, key=c, 分类=creator_classify(name)[0],
                         细分类=creator_classify(name)[1], 类型=k, 层级="全量"))
    cat = pd.DataFrame(rows)
    cat.insert(0, "列号", range(1, len(cat) + 1))
    return cat


def creator_link_with_channels(base: dict, ops_df: pd.DataFrame | None) -> tuple[pd.DataFrame, bool]:
    """链接 × 天长表 + 渠道 GMV。
    运营数据在时：用运营日报 All 区段逐链逐日的「达人视频/直播 GMV」（精确，跟达人链接 GMV 逐格 99.6% 一致）；
    不在时：用视频/直播挂车平摊估算（视频 GMV 是这条视频带来的全部成交，逐链误差较大，页面上会标注）。"""
    link = base["link"].copy()
    need = ["All::Creator video-attributed GMV", "All::Creator LIVE-attributed GMV"]
    exact = ops_df is not None and not ops_df.empty and all(c in ops_df.columns for c in need)
    if exact:
        o = (ops_df[ops_df["日期"].isin(link["日期"].unique())]
             .groupby(["日期", "product_id"])[need].sum().reset_index())
        link = link.merge(o, on=["日期", "product_id"], how="left")
        link[K_CR["ch_vid"]] = _f(link.pop(need[0]))
        link[K_CR["ch_live"]] = _f(link.pop(need[1]))
    else:
        link[K_CR["ch_vid"]] = link["_v_gmv_anchor"]
        link[K_CR["ch_live"]] = link["_l_gmv_anchor"]
    link[K_CR["ch_card"]] = link[K_CR["gmv"]] - link[K_CR["ch_vid"]] - link[K_CR["ch_live"]]
    return link, exact


# 记分卡：(分类, 显示名, K 短名 或 DERIVED 名, 格式)
SCORE = [
    ("前端", "曝光", "impr", "count"), ("前端", "独立曝光", "uimpr", "count"),
    ("前端", "点击", "clicks", "count"), ("前端", "CTR 点击率", "CTR 点击率", "pct"),
    ("前端", "加购次数", "atc", "count"), ("前端", "加购率（加购/点击）", "加购率 ATC", "pct"),
    ("前端", "CTOR 点击成单率", "CTOR 点击成单率", "pct"),
    ("前端", "达人渠道曝光", "i_cre", "count"), ("前端", "　其中·达人直播曝光", "i_cre_live", "count"),
    ("前端", "　其中·达人视频曝光", "i_cre_vid", "count"), ("前端", "商品卡曝光", "i_card", "count"),
    ("前端", "自播曝光", "i_live", "count"), ("前端", "商家视频曝光", "i_vid", "count"),
    ("前端", "达人渠道CTR", "达人渠道CTR", "pct"), ("前端", "商品卡CTR", "商品卡CTR", "pct"),
    ("前端", "达人新开直播数", "new_live", "count"), ("前端", "达人新发视频数", "new_vid", "count"),
    ("后端", "GMV", "gmv", "money"), ("后端", "订单", "orders", "count"), ("后端", "SKU 订单", "sku", "count"),
    ("后端", "件数", "items", "count"), ("后端", "买家数", "cust", "count"),
    ("后端", "AOV 客单价", "AOV 客单价", "money"), ("后端", "件均价 ASP", "件均价 ASP", "money"),
    ("后端", "件/单 连带率", "件/单 连带率", "num"),
    ("后端", "达人 GMV", "ch_cre", "money"), ("后端", "自播 GMV", "ch_slive", "money"),
    ("后端", "商品卡 GMV", "ch_card", "money"), ("后端", "商家视频 GMV", "ch_svid", "money"),
    ("后端", "退款额（有滞后，近期偏低）", "refund", "money"), ("后端", "退款率（有滞后）", "退款率", "pct"),
]

METRIC_PICK = {"曝光": "impr", "点击": "clicks", "GMV": "gmv", "订单": "orders",
               "加购": "atc", "件数": "items",
               "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率",
               "AOV 客单价": "AOV 客单价"}

# 两套数据集共用同一套页面：页面代码里的 df / K / DERIVED / CHANNELS / SCORE … 都是模块级变量，
# 切到「达人」时把它们整体换成达人版（_activate_creator），页面函数一行不用分叉。
PROF_OPS = dict(
    key="ops", name="运营", title="NailVesta 链接数据深度分析", scope="全店",
    periods=["日", "周", "月", "季"],
    granular=[("gmv", "GMV", "money"), ("impr", "曝光", "count"), ("clicks", "点击", "count"),
              ("atc", "加购", "count"), ("orders", "订单", "count"), ("sku", "SKU订单", "count"),
              ("items", "件数", "count"), ("cust", "买家数", "count"), ("refund", "退款额", "money"),
              ("i_cre", "达人渠道曝光", "count"), ("i_card", "商品卡曝光", "count"),
              ("i_live", "自播曝光", "count"), ("i_vid", "商家视频曝光", "count"),
              ("ch_cre", "达人GMV", "money"), ("ch_slive", "自播GMV", "money"),
              ("ch_card", "商品卡GMV", "money"), ("ch_svid", "商家视频GMV", "money"),
              ("new_live", "达人新开直播", "count"), ("new_vid", "达人新发视频", "count")],
    drill=[("gmv", "GMV"), ("impr", "曝光"), ("clicks", "点击"), ("sku", "SKU订单"), ("orders", "订单"),
           ("items", "件数"), ("atc", "加购"), ("i_cre", "达人曝光"), ("i_card", "商品卡曝光"),
           ("i_live", "自播曝光"), ("i_vid", "视频曝光")],
    compare_opts={"曝光": "impr", "GMV": "gmv", "点击": "clicks", "订单": "orders", "加购": "atc",
                  "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率", "AOV 客单价": "AOV 客单价"},
    anomaly_targets={"曝光": "impr", "点击": "clicks", "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率",
                     "订单": "orders", "GMV": "gmv"},
    front_desc="**前端 = 流量获取与漏斗效率**：曝光 / 点击 / CTR / 加购 / CTOR / 内容供给",
    back_desc="**后端 = 成交与履约售后**：GMV / 订单 / 件数 / 客户 / AOV / 退款 / 运费 / 税",
)
PROF_CR = dict(
    key="cr", name="达人", title="NailVesta 达人数据深度分析", scope="达人合计",
    periods=["日", "周", "月"],
    granular=[("gmv", "GMV", "money"), ("impr", "曝光", "count"), ("clicks", "点击", "count"),
              ("orders", "订单", "count"), ("items", "件数", "count"), ("cust", "买家数", "count"),
              ("ch_vid", "达人视频GMV", "money"), ("ch_live", "达人直播GMV", "money"),
              ("ch_card", "商品卡及其他GMV", "money"), ("i_vid", "视频挂车曝光", "count"),
              ("i_live", "直播挂车曝光", "count"), ("i_card", "商品卡及其他曝光", "count"),
              ("new_vid", "新发视频", "count"), ("new_live", "新开直播", "count"),
              ("posted", "发布达人", "count"), ("cws", "出单达人", "count"),
              ("v_active", "活跃视频", "count"), ("v_selling", "出单视频", "count"),
              ("v_views", "视频播放", "count"), ("l_viewers", "直播观众", "count"),
              ("samples", "寄样", "count"), ("samp_ct", "样品产出内容", "count"),
              ("comm", "预估佣金", "money"), ("refund", "退款额", "money")],
    drill=[("gmv", "GMV"), ("impr", "曝光"), ("clicks", "点击"), ("sku", "订单"), ("items", "件数"),
           ("i_vid", "视频挂车曝光"), ("i_live", "直播挂车曝光"), ("i_card", "商品卡及其他曝光"),
           ("new_vid", "新发视频"), ("new_live", "新开直播"), ("posted", "发布达人"), ("cws", "出单达人"),
           ("v_views", "视频播放"), ("comm", "预估佣金")],
    compare_opts={"曝光": "impr", "GMV": "gmv", "点击": "clicks", "订单": "orders", "新发视频": "new_vid",
                  "发布达人": "posted", "出单达人": "cws", "视频播放": "v_views",
                  "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率", "AOV 客单价": "AOV 客单价",
                  "视频互动率": "视频互动率", "佣金率": "佣金率"},
    anomaly_targets={"曝光": "impr", "点击": "clicks", "CTR 点击率": "CTR 点击率", "CTOR 点击成单率": "CTOR 点击成单率",
                     "订单": "orders", "GMV": "gmv", "视频播放": "v_views", "新发视频": "new_vid",
                     "发布达人": "posted", "出单达人": "cws", "视频互动率": "视频互动率", "视频完播率": "视频完播率"},
    front_desc="**前端 = 达人带来的流量与内容**：曝光 / 点击 / CTR / CTOR / 视频·直播挂车曝光 / 新发视频 / "
               "发布与出单达人 / 播放、互动、完播 / 寄样转化",
    back_desc="**后端 = 达人带来的成交与成本**：GMV / 订单 / 件数 / 买家 / 客单价 / 视频·直播·商品卡 GMV / "
              "每千次播放 GMV / 佣金 / 退款",
)
PROF, SCOPE = PROF_OPS, PROF_OPS["scope"]
DAYTOT: pd.DataFrame | None = None       # 达人：不带链接维度的全量逐日表（达人表/视频表/直播表原始字段）
CR_CH_EXACT = False                      # 达人：按链接的视频/直播 GMV 是否已用运营数据精确校准


def _creator_base() -> dict | None:
    """达人组四张表 → 分析结构。每个会话只整理一次（拉了新数据才重算），切页面不重复算。"""
    cd = st.session_state.get("_creator_data")
    if not cd:
        return None
    ver = st.session_state.setdefault("_creator_ver", id(cd))
    memo = st.session_state.get("_cr_base_memo")
    if memo and memo[0] == ver:
        return memo[1]
    with st.spinner("整理达人组数据（链接 / 达人 / 视频 / 直播）…"):
        base = build_creator_base(cd)
    st.session_state["_cr_base_memo"] = (ver, base)
    # 原始四张表整理完就用不到了（页面只读整理后的结构），换成空表，省下一整份内存
    st.session_state.setdefault("_creator_rows", {k: len(v) for k, v in cd.items()})
    st.session_state["_creator_data"] = {k: v.iloc[:0] for k, v in cd.items()}
    return base


def _activate_creator(base: dict):
    """把页面用到的口径整体切成达人版。"""
    global K, DERIVED, CHANNELS, CH_IMPR, CH_CLICK, CH_COLOR, SCORE, METRIC_PICK, DEN_NAME, PROF, SCOPE, DAYTOT
    K, DERIVED, CHANNELS = K_CR, DERIVED_CR, CHANNELS_CR
    CH_IMPR, CH_CLICK, CH_COLOR = CH_IMPR_CR, CH_CLICK_CR, CH_COLOR_CR
    SCORE, METRIC_PICK, DEN_NAME = SCORE_CR, METRIC_PICK_CR, DEN_NAME_CR
    PROF, SCOPE, DAYTOT = PROF_CR, PROF_CR["scope"], base["daytot"]


payloads = st.session_state.get("_payloads", [])
has_creator = bool(st.session_state.get("_creator_data"))
df = catalog = pd.DataFrame()
log: list[str] = []

ops_stats: dict = {}
if payloads:
    # 这次要新识别几个：只在「这批文件」第一次出现时数一次（load_files 有跨会话缓存，它自己的计数可能是别人那次的）
    _memo = st.session_state.get("_ops_new_memo")
    if not _memo or _memo[0] != id(payloads):
        _pd = _parse_dir()
        _n_new = sum(not os.path.exists(os.path.join(_pd, hashlib.sha1(raw).hexdigest()[:24] + ".json"))
                     for _n, raw in payloads)
        st.session_state["_ops_new_memo"] = (id(payloads), _n_new, pd.Timestamp.now())
    df, catalog, log, ops_stats = load_files(payloads)
    if df.empty:
        st.title("NailVesta 链接数据深度分析")
        st.error("没有解析到有效数据。请检查文件格式。")
        with st.expander("解析日志"):
            st.write("\n".join(log))
        st.stop()

has_ops = not df.empty

if not has_ops and not has_creator:
    st.title("NailVesta 链接数据深度分析")
    st.info("👈 左侧选择数据源：载入每日 product analysis 导出文件（3/20–9/20 每天一个 xlsx），"
           "和/或拉「🎨 达人组数据」。两边都有时，侧边栏最上面可以一键切换运营 / 达人，页面是同一套。")
    st.markdown(
        "**说明**：飞书 Base 里的日报是*附件文件*。把它们下载到一个文件夹（或打包成 zip），"
        "在左侧指定路径/上传即可。程序会自动识别每个文件内部的 `Analysis date`。")
    st.stop()

if has_ops and has_creator:
    SRC = NAV_BOX.radio("看哪边的数据", ["运营", "达人"], horizontal=True, key="src",
                        help="下面是同一套页面：切到「达人」后，每一页都换成达人组数据来分析。")
else:
    SRC = "运营" if has_ops else "达人"
    NAV_BOX.caption(f"当前只载入了{SRC}数据" + ("——拉了达人组数据就能切换到达人" if has_ops
                                              else "——载入运营日报后可切换到运营"))

OPS_ALL = df                                  # 运营全量：运营页里的「达人侧原因」、达人渠道 GMV 校准都要用
CRB = _creator_base() if has_creator else None
if SRC == "达人":
    df, CR_CH_EXACT = creator_link_with_channels(CRB, OPS_ALL if has_ops else None)
    catalog = creator_catalog_full(df, CRB["daytot"])
    _activate_creator(CRB)

dmin, dmax = df["日期"].min(), df["日期"].max()
st.sidebar.markdown("---")
st.sidebar.success(f"{PROF['name']}数据：{df['日期'].nunique()} 天 · {df['product_id'].nunique()} 个链接")
st.sidebar.caption(f"{dmin.date()} → {dmax.date()}")
if SRC == "运营" and st.session_state.get("_ops_new_memo"):
    _m = st.session_state["_ops_new_memo"]
    _dl = st.session_state.get("_lark_new")
    st.sidebar.caption(f"运营日报 {len(payloads)} 个文件：{_m[2]:%m-%d %H:%M} 载入时"
                       + (f"新下载 {_dl} 个、" if _dl is not None else "")
                       + f"新识别 {_m[1]} 个，其余 {len(payloads) - _m[1]} 个直接用已识别结果。")

df_all = df          # 「两段对比」用全量数据，不受下面的分析区间限制
rng = st.sidebar.date_input(f"分析区间（{PROF['name']}）", value=(dmin.date(), dmax.date()),
                            min_value=dmin.date(), max_value=dmax.date(), key=f"rng_{PROF['key']}")
if isinstance(rng, tuple) and len(rng) == 2:
    df = df[(df["日期"] >= pd.Timestamp(rng[0])) & (df["日期"] <= pd.Timestamp(rng[1]))]

if SRC == "运营":
    with st.sidebar.expander("解析日志"):
        st.text("\n".join(log[-200:]))

# 缺口检查
alldays = pd.date_range(df["日期"].min(), df["日期"].max(), freq="D")
missing = sorted(set(alldays) - set(df["日期"].unique()))
if missing:
    st.sidebar.warning(f"缺 {len(missing)} 天，如：{', '.join(str(pd.Timestamp(m).date()) for m in list(missing)[:5])}…")

daily = aggregate(df, "日")
EV_ALL = detect_events(df)          # 改动事件：链接下钻 和 改动溯源 共用，只算一次


def sums_of(d: pd.DataFrame) -> dict:
    """一段数据的全部求和列合计，键 = K 的短名。"""
    return {k: float(d[c].sum()) for k, c in K.items() if c in d.columns}


def value_of(s: dict, key: str) -> float:
    """从合计里取指标值：比率型用分子/分母重算。"""
    if key in DERIVED:
        num, den, f_, _ = DERIVED[key]
        n_, d_ = s.get(num, 0.0), s.get(den, 0.0)
        return n_ / d_ * (100 if f_ == "pct" else 1) if d_ else np.nan
    return s.get(key, np.nan)


def chg_txt(a, b, f_) -> str:
    """本期 a vs 基期 b：比率型给百分点 pp，其余给 %。"""
    if pd.isna(a) or pd.isna(b):
        return "—"
    if f_ == "pct":
        return f"{a - b:+.2f} pp"
    return f"{pct(a, b):+.1f}%" if b else "—"


def contrib_bars(t: pd.DataFrame, dcol: str, n: int = 12, xtitle: str = "Δ", label_col: str = "链接") -> go.Figure:
    """拖累最多的 n 条（红）+ 拉动最多的 n 条（绿），横向条形图。"""
    s = pd.concat([t[t[dcol] < 0].head(n), t[t[dcol] > 0].tail(n)]).sort_values(dcol)
    f = go.Figure(go.Bar(x=s[dcol], y=s[label_col], orientation="h",
                         marker_color=np.where(s[dcol] >= 0, "#1F8A6B", "#C24338"),
                         hovertemplate="%{y}<br>%{x:,.2f}<extra></extra>"))
    f.add_vline(x=0, line_color="#8A727C")
    f.update_layout(height=max(320, 24 * len(s) + 40), margin=dict(t=10, b=10),
                    xaxis_title=xtitle, yaxis=dict(automargin=True))
    return f



# ──────────────────────────────────────────────────────────────────────
# 6a. APP 版式部件（每页统一：页头 → 结论卡 → 分区标签）+ 单位 + 大白话句式
# ──────────────────────────────────────────────────────────────────────
def unit_of(key: str) -> str:
    """K 短名 → 单位：金额 $，曝光 IM，点击「次」，订单「单」，件数「件」…"""
    k = key or ""
    if k in ("gmv", "refund", "ship", "comm", "shop_gmv") or k.startswith("ch_"):
        return "$"
    if k in ("impr", "uimpr", "shop_impr", "l_room") or k.startswith("i_"):
        return "IM"
    if k in ("clicks", "uclicks", "shop_clicks", "atc", "atc_users") or k.startswith("c_"):
        return "次"
    return {"orders": "单", "sku": "单", "items": "件", "iref": "件", "cust": "人", "new_vid": "条",
            "v_active": "条", "v_selling": "条", "vws": "条", "samp_ct": "条", "new_live": "场", "l_active": "场",
            "lws": "场", "posted": "人", "cws": "人", "v_creators": "人", "l_viewers": "人", "samples": "个",
            "v_views": "次播放"}.get(k, "")


def fv(key: str, v, digits: int = 0) -> str:
    """带单位的数值：$1,038 / 285,242 IM / 3.68% / 1,258 单。"""
    if v is None or pd.isna(v):
        return "—"
    if key in DERIVED:
        f_ = DERIVED[key][2]
        return f"{v:.2f}%" if f_ == "pct" else (f"${v:,.2f}" if f_ == "money" else f"{v:,.2f}")
    u = unit_of(key)
    if u == "$":
        return f"${v:,.{digits}f}"
    return f"{v:,.{digits}f}" + (f" {u}" if u else "")


def fdelta(key: str, d, digits: int = 0) -> str:
    """带单位的变化量：+$1,112 / -59,718 IM / -0.09 个百分点。"""
    if d is None or pd.isna(d):
        return "—"
    if key in DERIVED:
        f_ = DERIVED[key][2]
        if f_ == "pct":
            return f"{d:+.2f} 个百分点"
        if f_ == "money":
            return ("+" if d >= 0 else "-") + f"${abs(d):,.2f}"
        return f"{d:+,.2f}"
    u = unit_of(key)
    if u == "$":
        return ("+" if d >= 0 else "-") + f"${abs(d):,.{digits}f}"
    return f"{d:+,.{digits}f}" + (f" {u}" if u else "")


def _sp(lbl: str) -> str:
    """日期 / 括号结尾的标签后面补个空格：「09-21 是」而不是「09-21是」。"""
    return lbl + (" " if lbl[-1:].isdigit() else "")


def say_change(name: str, key: str, cur, base, cur_lbl: str, base_lbl: str) -> str:
    """大白话：「曝光：09-14 是 285,242 IM，前 7 天日均是 344,960 IM，少了 59,718 IM，降了 17.3%。」"""
    if cur is None or base is None or pd.isna(cur) or pd.isna(base):
        return f"{name}：数据不足。"
    d = cur - base
    if key in DERIVED and DERIVED[key][2] == "pct":
        if abs(d) < 0.005:
            return f"{name}：{_sp(cur_lbl)}是 {fv(key, cur)}，跟{base_lbl}（{fv(key, base)}）基本一样。"
        return (f"{name}：{_sp(cur_lbl)}是 {fv(key, cur)}，{_sp(base_lbl)}是 {fv(key, base)}，"
                f"{'高' if d > 0 else '低'}了 {abs(d):.2f} 个百分点。")
    if d == 0:
        return f"{name}：{_sp(cur_lbl)}是 {fv(key, cur)}，跟{base_lbl}一样。"
    p = pct(cur, base)
    tail = f"，{'涨' if d > 0 else '降'}了 {abs(p):.1f}%" if pd.notna(p) else "（基期是 0，没法算百分比）"
    return (f"{name}：{_sp(cur_lbl)}是 {fv(key, cur)}，{_sp(base_lbl)}是 {fv(key, base)}，"
            f"{'多' if d > 0 else '少'}了 {fv(key, abs(d))}{tail}。")


def level_of(d, key: str = "impr", thr: float = 5.0) -> str:
    """变化方向 → 卡片颜色。比率用百分点，其余用 %。"""
    if d is None or pd.isna(d):
        return "flat"
    if abs(d) < thr:
        return "flat"
    return "up" if d > 0 else "down"


def _esc(s) -> str:
    """HTML 转义；**x** → 粗体；$ 换成实体（不然会被 Markdown 当公式）。"""
    s = _html.escape(str(s))
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    return s.replace("$", "&#36;")


def page_header(title: str, sub: str = "", chips: list[str] | None = None):
    chips_html = "".join(f'<span class="nv-chip">{_esc(c)}</span>' for c in (chips or []) if c)
    st.html(f'<div class="nv-head"><div><div class="nv-title">{_esc(title)}</div>'
            f'<div class="nv-sub">{_esc(sub)}</div></div><div class="nv-chips">{chips_html}</div></div>')


def summary_pills(items: list[tuple[str, str]]):
    """一行状态胶囊：[(down/up/flat/warn, 文字)]"""
    st.html('<div class="nv-sum">' + "".join(f'<span class="nv-pill {lv}">{_esc(t)}</span>'
                                             for lv, t in items if t) + "</div>")


_CARD_ICON = {"down": "🔴", "up": "🟢", "warn": "🟡", "info": "🔵", "flat": "⚪"}


def insight_cards(cards: list[dict], cols: int = 0):
    """结论卡：{"level": down/up/warn/info/flat, "title": 标题, "body": 一句或多句大白话, "tag": 小标签}"""
    parts = []
    for c in cards:
        body = c.get("body", "")
        body = "<br>".join(_esc(x) for x in body if x) if isinstance(body, (list, tuple)) else _esc(body)
        tag = f'<span class="nv-tag">{_esc(c["tag"])}</span>' if c.get("tag") else ""
        lv = c.get("level", "info")
        parts.append(f'<div class="nv-card {lv}"><div class="nv-card-t">{_CARD_ICON.get(lv, "🔵")} '
                     f'{_esc(c.get("title", ""))}{tag}</div><div class="nv-card-b">{body}</div></div>')
    style = f' style="grid-template-columns: repeat({cols}, minmax(0, 1fr))"' if cols else ""
    st.html(f'<div class="nv-cards"{style}>' + "".join(parts) + "</div>")


def section_title(text: str, note: str = ""):
    st.html(f'<div class="nv-sec">{_esc(text)}' + (f'<span class="nv-sec-note">{_esc(note)}</span>' if note else "")
            + "</div>")


# ── ① 总览 ────────────────────────────────────────────────────────────
def _nav_cards(items: list[tuple[str, str, str]]):
    """首页快捷入口卡片：items = [(page_key, 一句话说明, emoji)]"""
    cols = st.columns(len(items))
    for cc, (key, desc, emoji) in zip(cols, items):
        p = PAGE_REFS.get(key)
        if not p:
            continue
        with cc.container(border=True):
            st.page_link(p, label=f"**{p.title}**", icon=emoji)
            st.caption(desc)


def page_overview():
    D = df_all
    d_last = D["日期"].max()
    a0, a1 = d_last - pd.Timedelta(days=6), d_last
    b0, b1 = d_last - pd.Timedelta(days=13), d_last - pd.Timedelta(days=7)
    if PROF["key"] == "ops":
        note = "达人组数据已接入（侧边栏最上面可切到「达人」）" if has_creator else ""
    else:
        note = "按链接的视频/直播 GMV 已用运营日报精确校准" if CR_CH_EXACT else "⚠️ 未载入运营日报：按链接的视频/直播 GMV 为挂车估算"
    page_header(PROF["title"], "打开先看这一页：最近 7 天整体怎么样、有没有预警、接下来该看哪一页。曝光是索引列，先看曝光。",
                [f"{PROF['name']}数据 {D['日期'].min():%m/%d} → {d_last:%m/%d}", f"{D['日期'].nunique()} 天",
                 f"{D['product_id'].nunique()} 个链接", f"{len(catalog)} 个指标", note])
    cm, bm = D["日期"].between(a0, a1), D["日期"].between(b0, b1)
    na, nb = D.loc[cm, "日期"].nunique(), D.loc[bm, "日期"].nunique()
    R = radar_all(7, 3.0, 30.0)
    today = R[R["日期"] == d_last]
    if not na or not nb:
        st.info("数据不足 14 天，没法做最近 7 天 vs 前 7 天。")
        sc = sb = None
    else:
        sc = sums_of(D[cm])
        sb = {k: v * na / nb for k, v in sums_of(D[bm]).items()}
        p_im, p_g = pct(sc["impr"], sb["impr"]), pct(sc["gmv"], sb["gmv"])
        d_ctr = value_of(sc, "CTR 点击率") - value_of(sb, "CTR 点击率")
        d_ctor = value_of(sc, "CTOR 点击成单率") - value_of(sb, "CTOR 点击成单率")
        summary_pills([("info", f"最近 7 天（{a0:%m/%d}–{a1:%m/%d}）vs 前 7 天"),
                       (level_of(p_im), f"曝光 {p_im:+.1f}%"), (level_of(d_ctr, thr=0.1), f"CTR {d_ctr:+.2f} 个百分点"),
                       (level_of(d_ctor, thr=0.1), f"CTOR {d_ctor:+.2f} 个百分点"), (level_of(p_g), f"GMV {p_g:+.1f}%"),
                       ("down" if (today["方向"] == "下跌").any() else "flat",
                        f"{d_last:%m-%d} 预警：🔴 {int((today['方向'] == '下跌').sum())} · 🟢 {int((today['方向'] == '上升').sum())}")])
    tabs = st.tabs(["结论", "核心指标", "趋势图"])
    with tabs[0]:
        if sc is not None:
            cur_lbl, base_lbl = f"最近 7 天（{a0:%m-%d}～{a1:%m-%d}）", f"前 7 天（{b0:%m-%d}～{b1:%m-%d}）"
            gc, gb, names = link_sums(D, cm, bm, nb / na)
            ti, tg = sum_contrib(gc, gb, names, "impr"), sum_contrib(gc, gb, names, "gmv")
            ch_i = sorted(((lbl.split()[0], sc.get(k, 0) - sb.get(k, 0)) for lbl, k in CH_IMPR.items()), key=lambda x: x[1])
            g_lines = [say_change("GMV", "gmv", sc["gmv"], sb["gmv"], cur_lbl, base_lbl),
                       _top_links_line(tg, "gmv", sc["gmv"] < sb["gmv"], 2)]
            if PROF["key"] == "cr":
                cd_ = CRB["creator_day"]
                dd = (cd_[cd_["日期"].between(a0, a1)].groupby("达人")["GMV"].sum()
                      .sub(cd_[cd_["日期"].between(b0, b1)].groupby("达人")["GMV"].sum(), fill_value=0).sort_values())
                tc_ = dd.head(2) if sc["gmv"] < sb["gmv"] else dd.tail(2).iloc[::-1]
                if len(tc_):
                    g_lines.append("达人上变化最大：" + "、".join(f"@{n}（{fdelta('gmv', v)}）" for n, v in tc_.items()) + "。")
            cards = [
                {"level": level_of(p_im), "title": f"曝光{'下降' if p_im < 0 else '上升'} {abs(p_im):.1f}%", "tag": "索引指标",
                 "body": [say_change("曝光", "impr", sc["impr"], sb["impr"], cur_lbl, base_lbl),
                          "按渠道：" + "、".join(f"{n} {fdelta('impr', v)}" for n, v in ch_i) + "。",
                          _top_links_line(ti, "impr", sc["impr"] < sb["impr"], 2)]},
                {"level": level_of(d_ctr, thr=0.1) if abs(d_ctr) >= abs(d_ctor) else level_of(d_ctor, thr=0.1),
                 "title": "转化：CTR / CTOR / 客单价", "tag": "漏斗中段",
                 "body": [say_change("CTR", "CTR 点击率", value_of(sc, "CTR 点击率"), value_of(sb, "CTR 点击率"), cur_lbl, base_lbl),
                          say_change("CTOR", "CTOR 点击成单率", value_of(sc, "CTOR 点击成单率"),
                                     value_of(sb, "CTOR 点击成单率"), cur_lbl, base_lbl),
                          say_change("客单价", "AOV 客单价", value_of(sc, "AOV 客单价"), value_of(sb, "AOV 客单价"),
                                     cur_lbl, base_lbl)]},
                {"level": level_of(p_g), "title": f"GMV {'下降' if p_g < 0 else '上升'} {abs(p_g):.1f}%", "tag": "结果",
                 "body": g_lines},
            ]
            if len(today):
                cards.append({"level": "down" if (today["方向"] == "下跌").any() else "up",
                              "title": f"{d_last:%m-%d} 有 {len(today)} 条预警", "tag": "去「异常预警」看全部",
                              "body": [f"{a['对象']} · {a['指标']}：{alert_headline(a, 7)}" for _, a in today.head(4).iterrows()]})
            else:
                cards.append({"level": "flat", "title": f"{d_last:%m-%d} 没有预警", "tag": "异常预警",
                              "body": "所有链接、渠道" + ("、达人" if PROF["key"] == "cr" else "") + "的曝光、点击、CTR、CTOR、订单、GMV "
                                      "都在正常波动范围里（跟前 7 天中位数比）。"})
            insight_cards(cards, cols=2)
        section_title("接下来看哪一页")
        _nav_cards([("compare", "谁涨谁跌、掉在哪个渠道：本期 vs 基期全指标归因", "🔍"),
                    ("anomaly", "每条链接 / 渠道" + ("/ 达人" if PROF["key"] == "cr" else "") + "的异常，曝光优先", "🚨"),
                    ("chain", "大跌之后有没有回暖，逐日追踪", "📉"),
                    ("effect", "你改过的链接，改完曝光有没有涨", "🧪")])
        if PROF["key"] == "cr":
            _nav_cards([("granular", "四个 Excel 的每个字段：逐日 / 逐周 / 逐月 / 自选区间环比", "🗓️"),
                        ("rank", "谁在卖、哪条视频/直播在卖：带环比和贡献度的排行", "🏆"),
                        ("drilldown", "任意一条链接：全指标趋势、占达人合计的比重、是哪些达人在推", "🔗")])
    with tabs[1]:
        tot = {k: col(df, k).sum() for k in K}
        st.caption(f"当前分析区间累计（左侧可以改区间）：{df['日期'].min():%m/%d} → {df['日期'].max():%m/%d}")
        section_title("前端 · 流量")
        a, b, c, d2 = st.columns(4)
        a.metric("曝光", fv("impr", tot["impr"]))
        b.metric("点击", fv("clicks", tot["clicks"]))
        c.metric("CTR", f"{tot['clicks'] / tot['impr'] * 100:.2f}%" if tot["impr"] else "—")
        d2.metric("CTOR", f"{tot['sku'] / tot['clicks'] * 100:.2f}%" if tot["clicks"] else "—")
        section_title("后端 · 成交")
        a, b, c, d2 = st.columns(4)
        a.metric("GMV", fv("gmv", tot["gmv"]))
        b.metric("订单", fv("sku", tot["sku"]))
        c.metric("件数", fv("items", tot["items"]))
        d2.metric("客单价", f"${tot['gmv'] / tot['sku']:,.2f}" if tot["sku"] else "—")
        if PROF["key"] == "cr":
            section_title("达人 · 内容供给与成本", "去重口径：同一条视频 / 同一个达人只算一次")
            days_in = df["日期"].unique()
            cday = CRB["creator_day"][CRB["creator_day"]["日期"].isin(days_in)]
            cont = CRB["content"][CRB["content"]["日期"].isin(days_in)]
            a, b, c, d2, e = st.columns(5)
            a.metric("新发视频", f"{cday['新发视频'].sum():,.0f} 条")
            b.metric("新开直播", f"{cday['新开直播'].sum():,.0f} 场")
            c.metric("出单达人", f"{cday.loc[cday['订单'] > 0, '达人'].nunique():,} 人",
                     help="区间内至少出过 1 单的达人（Creators 表去重）")
            d2.metric("出单视频", f"{cont.loc[(cont['类型'] == '视频') & (cont['GMV'] > 0), '内容ID'].nunique():,} 条")
            e.metric("预估佣金", f"${tot['comm']:,.0f}",
                     f"佣金率 {tot['comm'] / tot['gmv'] * 100:.1f}%" if tot["gmv"] else None, delta_color="off",
                     delta_arrow="off")
    with tabs[2]:
        section_title("曝光 / CTR / CTOR / GMV 走势", "都换算成指数：起点 7 天平均 = 100，方便放在一张图里比")
        idx = daily.copy()
        for nm, kk in [("曝光", K["impr"]), ("GMV", K["gmv"])]:
            if kk in idx:
                base0 = idx[kk].iloc[:7].mean()
                idx[nm + "_指数"] = idx[kk] / base0 * 100 if base0 else np.nan
        for nm in ["CTR 点击率", "CTOR 点击成单率"]:
            if nm in idx:
                base0 = idx[nm].iloc[:7].mean()
                idx[nm.split()[0] + "_指数"] = idx[nm] / base0 * 100 if base0 else np.nan
        cols_idx = [c for c in idx.columns if c.endswith("_指数")]
        if cols_idx:
            fig = px.line(idx, x="期间", y=cols_idx, labels={"value": "指数（起点 7 天平均 = 100）", "期间": ""})
            fig.update_layout(height=380, legend_title="", hovermode="x unified", margin=dict(t=20, b=10))
            st.plotly_chart(fig, key="ov_idx")
        section_title("曝光按渠道（每日堆叠，IM）")
        chi = df.groupby("日期").agg(**{lbl.split()[0]: (K[k], "sum") for lbl, k in CH_IMPR.items()}).reset_index()
        fi = px.area(chi, x="日期", y=[l.split()[0] for l in CH_IMPR], labels={"value": "曝光 IM", "日期": ""},
                     color_discrete_map={l.split()[0]: CH_COLOR[l] for l in CH_IMPR})
        fi.update_layout(height=340, legend_title="", hovermode="x unified", margin=dict(t=20, b=10))
        st.plotly_chart(fi, key="ov_chi")
        section_title("GMV 按渠道（每日堆叠，$）")
        if PROF["key"] == "cr" and not CR_CH_EXACT:
            st.caption("⚠️ 没载入运营日报，按链接的视频 / 直播 GMV 用挂车关系估算，逐链误差较大。")
        ch_d = df.groupby("日期").agg(**{lbl: (K[k], "sum") for lbl, k in CHANNELS}).reset_index()
        figc = px.area(ch_d, x="日期", y=[l for l, _ in CHANNELS], labels={"value": "GMV $", "日期": ""},
                       color_discrete_map=CH_COLOR)
        figc.update_layout(height=340, legend_title="", hovermode="x unified", margin=dict(t=20, b=10))
        st.plotly_chart(figc, key="ov_chg")


# ── 两段对比·归因 ──────────────────────────────────────────────────────
def _delta_txt(key: str, a, b) -> str:
    """KPI 卡下面那行：比率给「个百分点」，其余给 %，都写明跟谁比。"""
    if pd.isna(a) or pd.isna(b):
        return "—"
    if key in DERIVED and DERIVED[key][2] == "pct":
        return f"{a - b:+.2f} 个百分点"
    p = pct(a, b)
    return "—" if pd.isna(p) else f"{p:+.1f}%"


def kpi_row(cur: dict, base: dict, keys: list[tuple[str, str]], vs: str = "vs 基期"):
    """一排指标卡：带单位的数值 + 相对基期的变化。"""
    cols = st.columns(len(keys))
    for cc, (lbl, key) in zip(cols, keys):
        a, b = value_of(cur, key), value_of(base, key)
        d = _delta_txt(key, a, b)
        cc.metric(lbl, fv(key, a, 2 if key in DERIVED else 0), None if d == "—" else f"{d} {vs}")


def _top_links_line(t: pd.DataFrame, key: str, falling: bool, n: int = 3) -> str:
    x = t[t["Δ"] < 0].head(n) if falling else t[t["Δ"] > 0].tail(n).iloc[::-1]
    if x.empty:
        return ""
    share = "占跌量%" if falling else "占涨量%"
    return (f"{'拖累' if falling else '拉动'}最多的链接：" + "、".join(
        f"{r['链接']}（{fdelta(key, r['Δ'])}，占{'跌' if falling else '涨'}量 {r[share]:.0f}%"
        + (f"，主要在{r['主要渠道']}" if r.get("主要渠道") else "") + "）" for _, r in x.iterrows()) + "。")


def render_compare():
    D = df_all
    d_first, d_last = D["日期"].min(), D["日期"].max()
    head = st.container()
    picked = _pick_two_periods(d_first, d_last, "cmp")
    if picked is None:
        with head:
            page_header("两段对比 · 归因", "本期跟基期比：变了多少、掉在哪个渠道、是哪几条链接")
        return
    a0, a1, b0, b1 = picked
    cm, bm = D["日期"].between(a0, a1), D["日期"].between(b0, b1)
    na, nb = D.loc[cm, "日期"].nunique(), D.loc[bm, "日期"].nunique()
    with head:
        page_header("两段对比 · 归因",
                    "本期跟基期比：变了多少、掉 / 涨在哪个渠道、是哪几条链接" + ("、哪几个达人" if PROF["key"] == "cr" else "")
                    + "。曝光是索引列，先看曝光。",
                    [f"{PROF['name']}数据", f"本期 {a0:%m/%d}–{a1:%m/%d}（{na} 天）", f"基期 {b0:%m/%d}–{b1:%m/%d}（{nb} 天）",
                     "不受左侧分析区间限制"])
    if not na or not nb:
        st.warning("所选区间里没有数据。")
        return
    if (cm & bm).any():
        st.warning("本期和基期有重叠的日子，差异会被稀释。")
    scale = nb / na
    if na != nb:
        st.caption(f"两段有数据的天数不同（本期 {na} 天、基期 {nb} 天），基期已按天数折算成跟本期一样长再比。")
    cur_lbl = f"本期（{a0:%m-%d}～{a1:%m-%d}）"
    base_lbl = f"基期（{b0:%m-%d}～{b1:%m-%d}{'，已折算' if na != nb else ''}）"
    lab_c, lab_b = f"本期 {a0:%m/%d}–{a1:%m/%d}", f"基期 {b0:%m/%d}–{b1:%m/%d}" + ("（已折算）" if na != nb else "")
    sc = sums_of(D[cm])
    sb = {k: v / scale for k, v in sums_of(D[bm]).items()}
    gc, gb, names = link_sums(D, cm, bm, scale)
    ti, tg = sum_contrib(gc, gb, names, "impr"), sum_contrib(gc, gb, names, "gmv")
    p_im, p_g = pct(sc["impr"], sb["impr"]), pct(sc["gmv"], sb["gmv"])
    d_ctr = value_of(sc, "CTR 点击率") - value_of(sb, "CTR 点击率")
    d_ctor = value_of(sc, "CTOR 点击成单率") - value_of(sb, "CTOR 点击成单率")
    summary_pills([(level_of(p_im), f"曝光 {p_im:+.1f}%" if pd.notna(p_im) else "曝光 —"),
                   (level_of(d_ctr, thr=0.1), f"CTR {d_ctr:+.2f} 个百分点" if pd.notna(d_ctr) else ""),
                   (level_of(d_ctor, thr=0.1), f"CTOR {d_ctor:+.2f} 个百分点" if pd.notna(d_ctor) else ""),
                   (level_of(p_g), f"GMV {p_g:+.1f}%" if pd.notna(p_g) else "GMV —")])
    tabs = st.tabs(["结论", "全指标归因", "图表", "明细表"])

    # ── 结论 ──
    with tabs[0]:
        kpi_row(sc, sb, [("曝光", "impr"), ("点击", "clicks"), ("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率")])
        kpi_row(sc, sb, [("订单", "sku"), ("客单价", "AOV 客单价"), ("GMV", "gmv"), ("件数", "items")])
        ch_i = sorted(((lbl, sc.get(CH_IMPR[lbl], 0) - sb.get(CH_IMPR[lbl], 0)) for lbl, _ in CHANNELS),
                      key=lambda x: x[1])
        fa = funnel_attribution(sc, sb).dropna(subset=["贡献金额"])
        fa = fa.reindex(fa["贡献金额"].abs().sort_values(ascending=False).index)
        cards = [
            {"level": level_of(p_im), "title": f"曝光{'下降' if (p_im or 0) < 0 else '上升'} {abs(p_im or 0):.1f}%",
             "tag": "索引指标",
             "body": [say_change("曝光", "impr", sc["impr"], sb["impr"], cur_lbl, base_lbl),
                      "按渠道：" + "、".join(f"{l.split()[0]} {fdelta('impr', v)}" for l, v in ch_i) + "。",
                      _top_links_line(ti, "impr", sc["impr"] < sb["impr"])]},
            {"level": level_of(d_ctr, thr=0.1) if abs(d_ctr or 0) >= abs(d_ctor or 0) else level_of(d_ctor, thr=0.1),
             "title": "转化：CTR / CTOR / 客单价", "tag": "漏斗中段",
             "body": [say_change("CTR（点击 ÷ 曝光）", "CTR 点击率", value_of(sc, "CTR 点击率"), value_of(sb, "CTR 点击率"),
                                 cur_lbl, base_lbl),
                      say_change("CTOR（订单 ÷ 点击）", "CTOR 点击成单率", value_of(sc, "CTOR 点击成单率"),
                                 value_of(sb, "CTOR 点击成单率"), cur_lbl, base_lbl),
                      say_change("客单价", "AOV 客单价", value_of(sc, "AOV 客单价"), value_of(sb, "AOV 客单价"),
                                 cur_lbl, base_lbl)]},
            {"level": level_of(p_g), "title": f"GMV {'下降' if (p_g or 0) < 0 else '上升'} {abs(p_g or 0):.1f}%",
             "tag": "结果",
             "body": [say_change("GMV", "gmv", sc["gmv"], sb["gmv"], cur_lbl, base_lbl),
                      "拆到漏斗四环：" + "、".join(f"{r['因子']}{'带来' if r['贡献金额'] > 0 else '拖掉'} "
                                               f"{fv('gmv', abs(r['贡献金额']))}" for _, r in fa.iterrows()) + "。",
                      _top_links_line(tg, "gmv", sc["gmv"] < sb["gmv"])]},
        ]
        if PROF["key"] == "ops":
            cre = [("达人直播曝光", "i_cre_live"), ("达人视频曝光", "i_cre_vid"),
                   ("达人新开直播（按链接累计）", "new_live"), ("达人新发视频（按链接累计）", "new_vid")]
            lines = [f"{lbl}：{fv(k, sb.get(k, 0))} → {fv(k, sc.get(k, 0))}（{_delta_txt(k, sc.get(k), sb.get(k))}）"
                     for lbl, k in cre if k in sc]
            if lines:
                cards.append({"level": level_of(pct(sc.get("i_cre", 0), sb.get("i_cre", 0))), "title": "达人渠道再拆",
                              "tag": "基期 → 本期", "body": lines})
        else:
            cre = [("新发视频", "new_vid"), ("新开直播", "new_live"), ("发布达人", "posted"), ("出单达人", "cws"),
                   ("活跃视频", "v_active"), ("出单视频", "v_selling")]
            cards.append({"level": level_of(pct(sc.get("new_vid", 0), sb.get("new_vid", 0))), "title": "内容供给",
                          "tag": "按链接累计，基期 → 本期",
                          "body": [f"{lbl}：{fv(k, sb.get(k, 0))} → {fv(k, sc.get(k, 0))}（{_delta_txt(k, sc.get(k), sb.get(k))}）"
                                   for lbl, k in cre if k in sc]})
        insight_cards(cards, cols=2)
        if PROF["key"] == "cr" and not CR_CH_EXACT:
            st.caption("⚠️ 没载入运营日报：按链接的视频/直播 GMV 用挂车关系估算，逐链误差较大。")

    # ── 全指标归因 ──
    with tabs[1]:
        opts_m = {k: v for k, v in PROF["compare_opts"].items() if v in DERIVED or v in K}
        pick_m = st.radio("按哪个指标排、看谁导致的", list(opts_m), horizontal=True, key="cmp_metric")
        key_m = opts_m[pick_m]
        section_title("按链接拆：每条链接的全部指标", f"按「{pick_m}」的变化排序；{SCOPE}变化 = 各链接变化之和")
        st.dataframe(link_multi_attr(D, cm, bm, scale,
                                     key_m if key_m in ("impr", "clicks", "sku", "orders", "gmv") else "impr")
                     .drop(columns="_sort"), hide_index=True)
        if key_m in DERIVED:
            f_ = DERIVED[key_m][2]
            unit = "个百分点" if f_ == "pct" else ("$" if f_ == "money" else "")
            t = ratio_contrib(gc, gb, names, key_m)
            if len(t):
                va, vb = value_of(sc, key_m), value_of(sb, key_m)
                insight_cards([{"level": level_of(va - vb, thr=0.05), "title": f"{pick_m} 是被哪些链接带动的",
                                "body": [say_change(pick_m, key_m, va, vb, cur_lbl, base_lbl),
                                         f"下面每条链接的『贡献』加起来正好等于这个变化（{t['贡献'].sum():+.3f} {unit}）。",
                                         "比率效应 = 链接自己的比率变了（主图 / 价格 / 详情页）；结构效应 = 流量在链接之间重新分配"
                                         f"（比如低 {pick_m} 的链接拿到更多曝光，也会拉低{SCOPE}的 {pick_m}）。"]}], cols=1)
                st.plotly_chart(contrib_bars(t, "贡献", xtitle=f"对{SCOPE} {pick_m} 变化的贡献（{unit}）"), key="cmp_rc")
                tbl = pd.concat([t[t["贡献"] < 0].head(15), t[t["贡献"] > 0].tail(10).iloc[::-1]])
                rs_ = [reason_for(key_m, r, gc.loc[p], gb.loc[p]) for p, r in tbl.iterrows()]
                disp = tbl.assign(主因=[x[0] for x in rs_], 原因=[x[1] for x in rs_])
                for c in ["贡献", "比率效应", "结构效应"]:
                    disp[c] = disp[c].map(lambda v: f"{v:+.3f} {unit}")
                for c in ["自身比率·基期", "自身比率·本期"]:
                    disp[c] = disp[c].map(lambda v: fmt_val(v, f_))
                for c in ["分母占比·基期%", "分母占比·本期%"]:
                    disp[c] = disp[c].map(lambda v: f"{v:.1f}%")
                st.dataframe(disp, hide_index=True)
        else:
            t = sum_contrib(gc, gb, names, key_m)
            net, down, up = t["Δ"].sum(), t.loc[t["Δ"] < 0, "Δ"].sum(), t.loc[t["Δ"] > 0, "Δ"].sum()
            top5 = t[t["Δ"] < 0].head(5)["占跌量%"].sum()
            insight_cards([{"level": level_of(pct(value_of(sc, key_m), value_of(sb, key_m))),
                            "title": f"{pick_m} 是被哪些链接带动的",
                            "body": [say_change(pick_m, key_m, value_of(sc, key_m), value_of(sb, key_m), cur_lbl, base_lbl),
                                     f"净变化 {fdelta(key_m, net)} = 下降的链接合计 {fdelta(key_m, down)} ＋ 上涨的链接合计 "
                                     f"{fdelta(key_m, up)}；拖累最多的 5 条占总跌量 {top5:.0f}%。"]}], cols=1)
            st.plotly_chart(contrib_bars(t, "Δ", xtitle=f"{pick_m}变化（{unit_of(key_m) or '数量'}，本期 − 基期）"),
                            key="cmp_sc")
            top = pd.concat([t[t["Δ"] < 0].head(15), t[t["Δ"] > 0].tail(10).iloc[::-1]])
            dg = [reason_for(key_m, r, gc.loc[p], gb.loc[p]) for p, r in top.iterrows()]
            disp = top.assign(主因=[x[0] for x in dg], 原因=[x[1] for x in dg])
            for c in ["基期", "本期"]:
                disp[c] = disp[c].map(lambda v: fv(key_m, v))
            for c in [c for c in disp.columns if c.startswith("Δ")]:
                disp[c] = disp[c].map(lambda v: fdelta(key_m, v))
            for c in ["自身变化%", "占跌量%", "占涨量%"]:
                disp[c] = disp[c].map(lambda v: "" if pd.isna(v) else (f"{v:+.1f}%" if c == "自身变化%" else f"{v:.1f}%"))
            st.dataframe(disp, hide_index=True)
            if PROF["key"] == "ops" and CRB is not None and "主要渠道" in top.columns:
                aff = top[top["主要渠道"] == "达人"]
                cov = set(pd.to_datetime(CRB["link"]["日期"].unique()))
                if len(aff) and not any(pd.Timestamp(d) in cov for d in D.loc[cm, "日期"].unique()):
                    st.caption(f"有 {len(aff)} 条链接的变化主要在「达人」渠道，但达人组数据只到 {max(cov):%m-%d}，"
                               "这一段对不了达人侧原因——把本期挪到那之前就能看到。")
                elif len(aff):
                    section_title(f"其中变化主要在「达人」渠道的 {len(aff)} 条链接：达人组数据里为什么",
                                  "日均；挂车达人按视频 / 直播 GMV")
                    st.dataframe(pd.DataFrame({"链接": aff["链接"].values, f"{pick_m}变化": [fdelta(key_m, v) for v in aff["Δ"]],
                                               "达人侧原因": [cr_link_reason(p, D.loc[cm, "日期"].unique(),
                                                                          D.loc[bm, "日期"].unique()) for p in aff.index]}),
                                 hide_index=True)
            chcols = [c for c in t.columns if c.startswith("Δ") and c != "Δ"]
            hm = t[t["Δ"] < 0].head(12).copy()
            if chcols and len(hm):
                if key_m == "impr" and K.get("i_cre_live") in gc.columns:
                    hm["Δ其中·达人直播"] = (gc[K["i_cre_live"]] - gb[K["i_cre_live"]]).reindex(hm.index)
                    hm["Δ其中·达人视频"] = (gc[K["i_cre_vid"]] - gb[K["i_cre_vid"]]).reindex(hm.index)
                    chcols = chcols + ["Δ其中·达人直播", "Δ其中·达人视频"]
                section_title("拖累最多的链接，各自掉在哪个渠道", f"格子 = 该渠道 本期 − 基期（{unit_of(key_m) or '数量'}）；红 = 掉，蓝 = 涨")
                fh = px.imshow(hm[chcols].values, x=[c[1:] for c in chcols], y=hm["链接"].tolist(),
                               color_continuous_scale="RdBu", color_continuous_midpoint=0, text_auto=",.0f", aspect="auto")
                fh.update_layout(height=max(300, 30 * len(hm) + 80), margin=dict(t=10, b=10), coloraxis_showscale=False)
                st.plotly_chart(fh, key="cmp_hm")
        if PROF["key"] == "ops":
            if CRB is not None:
                render_creator_cross_check(D.loc[cm, "日期"].unique(), D.loc[bm, "日期"].unique(), sc, sb)
        else:
            render_creator_dimensions(D.loc[cm, "日期"].unique(), D.loc[bm, "日期"].unique(), pick_m)

    # ── 图表 ──
    with tabs[2]:
        section_title("GMV 拆到漏斗四环", "GMV = 曝光 × CTR × CTOR × 客单价；对数拆分，四项加起来 = GMV 总变化")
        fa_all = funnel_attribution(sc, sb)
        d_gmv = sc["gmv"] - sb["gmv"]
        fw = go.Figure(go.Waterfall(
            x=fa_all["因子"].tolist() + ["GMV 总变化"], y=fa_all["贡献金额"].tolist() + [d_gmv],
            measure=["relative"] * len(fa_all) + ["total"],
            text=[fdelta("gmv", v) for v in fa_all["贡献金额"].tolist() + [d_gmv]], textposition="outside",
            decreasing=dict(marker=dict(color="#D9534F")), increasing=dict(marker=dict(color="#2E9E7A")),
            totals=dict(marker=dict(color="#4C6EF5"))))
        fw.update_layout(height=320, margin=dict(t=20, b=10), yaxis_title="$")
        st.plotly_chart(fw, key="cmp_fw")
        section_title("按渠道：曝光和 GMV 各掉在哪 / 涨在哪", "本期 − 基期")
        chd = pd.DataFrame([{"渠道": lbl, "Δ曝光": sc.get(CH_IMPR[lbl], 0) - sb.get(CH_IMPR[lbl], 0),
                             "ΔGMV": sc.get(gk, 0) - sb.get(gk, 0)} for lbl, gk in CHANNELS])
        cc1, cc2 = st.columns(2)
        for cc, col_, ukey in [(cc1, "Δ曝光", "impr"), (cc2, "ΔGMV", "gmv")]:
            f = go.Figure(go.Bar(x=[l.split()[0] for l in chd["渠道"]], y=chd[col_],
                                 marker_color=[CH_COLOR[l] for l in chd["渠道"]],
                                 text=[fdelta(ukey, v) for v in chd[col_]], textposition="outside"))
            f.add_hline(y=0, line_color="#8A727C")
            f.update_layout(height=300, margin=dict(t=40, b=10), title=f"各渠道{col_[1:]}变化（{unit_of(ukey)}）")
            cc.plotly_chart(f, key=f"cmp_ch_{ukey}")
        section_title("逐日：两段里每天怎么走的", "柱 = 各渠道曝光（IM）；线 = GMV（$，右轴）")
        lo, hi = min(a0, b0), max(a1, b1)
        need = [K[k] for k in ["gmv", "impr", "clicks", "sku"]] + [K[CH_IMPR[l]] for l, _ in CHANNELS]
        need = [c for c in need if c in D.columns]
        dall = D.groupby("日期")[need].sum()
        w = dall[(dall.index >= lo) & (dall.index <= hi)]
        fd = go.Figure()
        for lbl, _ in CHANNELS:
            if K[CH_IMPR[lbl]] in w.columns:
                fd.add_bar(x=w.index, y=w[K[CH_IMPR[lbl]]], name=f"{lbl.split()[0]}曝光", marker_color=CH_COLOR[lbl])
        fd.add_trace(go.Scatter(x=w.index, y=w[K["gmv"]], name="GMV（右轴）", mode="lines+markers",
                                line=dict(color="#C2416B", width=2), yaxis="y2"))
        for x0, x1, nm_ in [(b0, b1, "基期"), (a0, a1, "本期")]:
            fd.add_vrect(x0=x0 - pd.Timedelta(hours=12), x1=x1 + pd.Timedelta(hours=12), fillcolor="#8A727C",
                         opacity=.06 if nm_ == "基期" else .12, line_width=0, annotation_text=nm_,
                         annotation_position="top left")
        fd.update_layout(barmode="stack", height=380, hovermode="x unified", margin=dict(t=30, b=10),
                         yaxis=dict(title="曝光 IM"), yaxis2=dict(title="GMV $", overlaying="y", side="right"),
                         legend=dict(orientation="h", y=1.14))
        st.plotly_chart(fd, key="cmp_daily")

    # ── 明细表 ──
    with tabs[3]:
        section_title(f"{SCOPE}记分卡：前端 / 后端全部指标")
        rows = []
        for cat_, lbl, key, f_ in SCORE:
            a_, b_ = value_of(sc, key), value_of(sb, key)
            if pd.isna(a_) and pd.isna(b_):
                continue
            rows.append({"分类": cat_, "指标": lbl, lab_b: fv(key, b_, 2 if key in DERIVED else 0),
                         lab_c: fv(key, a_, 2 if key in DERIVED else 0), "变化（vs 基期）": _delta_txt(key, a_, b_)})
        st.dataframe(pd.DataFrame(rows), hide_index=True, height=min(1100, 38 + 35 * len(rows)))
        section_title("逐日明细", "每个 % 都写明了跟谁比")
        avg7 = dall.rolling(7, min_periods=3).mean().shift(1)
        part = np.where(w.index.to_series().between(a0, a1), "本期", np.where(w.index.to_series().between(b0, b1), "基期", "—"))
        st.dataframe(pd.DataFrame({
            "日期": w.index.strftime("%m-%d (%a)"), "所属": part,
            "曝光": w[K["impr"]].map(lambda v: fv("impr", v)),
            "曝光 vs前一天": (w[K["impr"]].pct_change() * 100).map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%"),
            "曝光 vs前7日均": ((dall[K["impr"]] / avg7[K["impr"]] - 1) * 100).reindex(w.index)
            .map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%"),
            "CTR": (w[K["clicks"]] / w[K["impr"]].replace(0, np.nan) * 100).map(lambda v: fmt_val(v, "pct")),
            "CTOR": (w[K["sku"]] / w[K["clicks"]].replace(0, np.nan) * 100).map(lambda v: fmt_val(v, "pct")),
            "GMV": w[K["gmv"]].map(lambda v: fv("gmv", v)),
            "GMV vs前一天": (w[K["gmv"]].pct_change() * 100).map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%"),
            "GMV vs前7日均": ((dall[K["gmv"]] / avg7[K["gmv"]] - 1) * 100).reindex(w.index)
            .map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%")}), hide_index=True)
        thr = st.slider("『特别好 / 特别差』的门槛：当天 vs 前 7 日均 偏离 ≥", 10, 60, 20, step=5, key="cmp_thr")
        days_all = sorted(D["日期"].unique())
        out = []
        for d in [x for x in days_all if a0 <= x <= a1]:
            prev = [x for x in days_all if x < d][-7:]
            if len(prev) < 3:
                continue
            for mk, lbl in [("impr", "曝光"), ("gmv", "GMV")]:
                v, b = dall.at[d, K[mk]], dall.loc[prev, K[mk]].mean()
                p = pct(v, b)
                if pd.isna(p) or abs(p) < thr:
                    continue
                dm, pm = D["日期"] == d, D["日期"].isin(prev)
                cp = culprits(D, dm, pm, base_scale=len(prev), topn=3, by=mk, direction=1 if p > 0 else -1)
                out.append({"日期": pd.Timestamp(d).strftime("%m-%d"), "指标": lbl,
                            "说明": f"{lbl} {fv(mk, v)}，前 {len(prev)} 天日均 {fv(mk, b)}，"
                                    f"{'多' if v > b else '少'}了 {fv(mk, abs(v - b))}（{p:+.0f}%）",
                            "主责链接": "、".join(f"{r['链接']}（{fdelta(mk, r['Δ'])}"
                                              + (f"，{r['主要渠道']}" if r.get("主要渠道") else "") + "）"
                                              for _, r in cp.iterrows()) if len(cp) else "—"})
        if out:
            section_title(f"本期里偏离 ≥ {thr}% 的日子，以及是哪几条链导致的")
            st.dataframe(pd.DataFrame(out), hide_index=True)
        else:
            st.caption(f"本期没有哪一天偏离前 7 日均 ≥ {thr}%。")
        with st.expander("📋 全部链接明细：本期 / 基期 / 变化（可下载 CSV）"):
            full = pd.DataFrame({"链接": names})
            for lbl, k_ in [("曝光", "impr"), ("点击", "clicks"), ("加购", "atc"), ("GMV", "gmv"),
                            ("订单", "orders"), ("件数", "items")]:
                if k_ not in K:
                    continue
                full[f"{lbl}·基期"], full[f"{lbl}·本期"] = gb[K[k_]], gc[K[k_]]
                full[f"{lbl}·变化%"] = np.where(gb[K[k_]] > 0, (gc[K[k_]] / gb[K[k_]].replace(0, np.nan) - 1) * 100, np.nan)
            for lbl, (n_, d_), m_ in [("CTR%", ("clicks", "impr"), 100), ("加购率%", ("atc", "clicks"), 100),
                                      ("CTOR%", ("sku", "clicks"), 100), ("件单价$", ("gmv", "items"), 1)]:
                if n_ not in K or d_ not in K:
                    continue
                rb_ = gb[K[n_]] / gb[K[d_]].replace(0, np.nan) * m_
                rc_ = gc[K[n_]] / gc[K[d_]].replace(0, np.nan) * m_
                full[f"{lbl}·基期"], full[f"{lbl}·本期"] = rb_, rc_
                full[f"{lbl}·变化{'个百分点' if m_ == 100 else '%'}"] = (rc_ - rb_) if m_ == 100 else (rc_ / rb_ - 1) * 100
            for lbl, k_ in CH_IMPR.items():
                full[f"{lbl.split()[0]}曝光变化 IM"] = gc[K[k_]] - gb[K[k_]]
            full["Product ID"] = full.index
            full = full.sort_values("曝光·本期", ascending=False)
            st.dataframe(round_num(full), hide_index=True, height=480)
            st.download_button("⬇️ 下载 CSV", full.to_csv(index=False).encode("utf-8-sig"),
                               f"两段对比_{a0:%m%d}-{a1:%m%d}_vs_{b0:%m%d}-{b1:%m%d}.csv", "text/csv", key="cmp_dl")


def page_compare():
    render_compare()

# ── ③ 多粒度对比 ──────────────────────────────────────────────────────
def page_granular():
    if PROF["key"] == "cr":
        page_granular_creator()
        return
    head = st.container()
    gran = st.radio("粒度", PROF["periods"], horizontal=True, index=1, key="gran_ops")
    agg = aggregate(df, gran)
    unit_g = {"日": "天", "周": "周", "月": "月", "季": "季"}[gran]
    use_avg = gran != "日"
    with head:
        page_header("多粒度对比", f"每一{unit_g}跟上一{unit_g}比：规模指标" + ("按日均比（各期天数不同）" if use_avg else "按当天比")
                    + "，比率的变化用「个百分点」。曝光先看。",
                    [f"{PROF['name']}数据", f"粒度：{gran}", f"分析区间 {df['日期'].min():%m/%d} → {df['日期'].max():%m/%d}"])
    metric_opts = {}
    for k, lbl, f_ in PROF["granular"]:
        if k in K and K[k] in agg.columns:
            if use_avg:
                agg[f"{lbl}（日均）"] = agg[K[k]] / agg["天数"]
                metric_opts[f"{lbl}（日均）"] = (f"{lbl}（日均）", f_, False, k)
            else:
                metric_opts[lbl] = (K[k], f_, False, k)
    for nm, (_n, _d, f_, _c) in DERIVED.items():
        if nm in agg.columns:
            metric_opts[nm] = (nm, f_, True, nm)

    def chg_series(s, f_, is_ratio):
        return s.diff() if (is_ratio and f_ == "pct") else s.pct_change() * 100

    tabs = st.tabs(["结论", "图表", "全部指标表"])
    with tabs[0]:
        if len(agg) < 2:
            st.info("只有一期数据，没法比。")
        else:
            p_now, p_prev = agg["期间"].iloc[-1], agg["期间"].iloc[-2]
            n_now, n_prev = int(agg["天数"].iloc[-1]), int(agg["天数"].iloc[-2])
            per_all_g = to_period(df["日期"], gran)
            cmk, bmk = per_all_g == p_now, per_all_g == p_prev
            gcg, gbg, nmg = link_sums(df, cmk, bmk, base_scale=n_prev / n_now)
            cur = {k: v / n_now for k, v in sums_of(df[cmk]).items()}
            base = {k: v / n_prev for k, v in sums_of(df[bmk]).items()}
            lbl_c = f"这一{unit_g}（{_plabel_any(p_now, gran if gran != '季' else '月')} 起，{n_now} 天）" + ("日均" if use_avg else "")
            lbl_b = f"上一{unit_g}（{n_prev} 天）" + ("日均" if use_avg else "")

            def lead(key_):
                t = sum_contrib(gcg, gbg, nmg, key_)
                t["Δ"] = t["Δ"] / n_now
                falling = cur.get(key_, 0) < base.get(key_, 0)
                x = t[t["Δ"] < 0].head(3) if falling else t[t["Δ"] > 0].tail(3).iloc[::-1]
                return (f"{'拖累' if falling else '拉动'}最多：" + "、".join(f"{r['链接']}（{fdelta(key_, r['Δ'])}/天）"
                                                                       for _, r in x.iterrows()) + "。") if len(x) else ""
            p_im, p_g = pct(cur["impr"], base["impr"]), pct(cur["gmv"], base["gmv"])
            summary_pills([("info", f"{lbl_c} vs {lbl_b}"), (level_of(p_im), f"曝光 {p_im:+.1f}%"),
                           (level_of(p_g), f"GMV {p_g:+.1f}%")])
            insight_cards([
                {"level": level_of(p_im), "title": f"曝光{'下降' if p_im < 0 else '上升'} {abs(p_im):.1f}%", "tag": "索引指标",
                 "body": [say_change("曝光", "impr", cur["impr"], base["impr"], lbl_c, lbl_b), lead("impr")]},
                {"level": level_of(value_of(cur, "CTR 点击率") - value_of(base, "CTR 点击率"), thr=0.1), "title": "转化",
                 "body": [say_change(nm, kk, value_of(cur, kk), value_of(base, kk), lbl_c, lbl_b)
                          for nm, kk in (("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"), ("客单价", "AOV 客单价"))]},
                {"level": level_of(p_g), "title": f"GMV {'下降' if p_g < 0 else '上升'} {abs(p_g):.1f}%", "tag": "结果",
                 "body": [say_change("GMV", "gmv", cur["gmv"], base["gmv"], lbl_c, lbl_b), lead("gmv")]},
                {"level": "info", "title": "其它规模指标", "tag": "日均" if use_avg else "",
                 "body": [say_change(lbl, k, cur.get(k), base.get(k), "这一期", "上一期") for k, lbl, _ in PROF["granular"]
                          if k in cur and k not in ("impr", "gmv")][:8]},
            ], cols=2)
    with tabs[1]:
        pick = st.multiselect("选择指标（可多选，每个都会出图）", list(metric_opts.keys()),
                              default=[x for x in metric_opts if x.split("（")[0] in ("曝光", "CTR 点击率", "CTOR 点击成单率", "GMV")],
                              key="gran_pick")
        for lbl in pick:
            cname, kind, is_ratio, key = metric_opts[lbl]
            sub = agg[["期间", "天数", cname]].dropna().copy()
            sub["环比"] = chg_series(sub[cname], kind, is_ratio)
            unit = "个百分点" if (is_ratio and kind == "pct") else "%"
            section_title(lbl, f"单位：{unit_of(key) or ('%' if is_ratio and kind == 'pct' else '')}；右图 = 环比（vs 上一{unit_g}，{unit}）")
            cc1, cc2 = st.columns([3, 2])
            f = go.Figure()
            f.add_bar(x=sub["期间"], y=sub[cname], name=lbl, marker_color="#C2416B", opacity=.85)
            f.add_trace(go.Scatter(x=sub["期间"], y=sub[cname].rolling(3, min_periods=1).mean(),
                                   name="3 期移动平均", mode="lines", line=dict(color="#4C6EF5", width=2)))
            f.update_layout(height=280, margin=dict(t=10, b=10), hovermode="x unified", legend=dict(orientation="h", y=1.15))
            cc1.plotly_chart(f, key=f"gr_v_{lbl}")
            f2 = go.Figure(go.Bar(x=sub["期间"], y=sub["环比"],
                                  marker_color=["#2E9E7A" if (pd.notna(v) and v >= 0) else "#D9534F" for v in sub["环比"]]))
            f2.update_layout(height=280, margin=dict(t=10, b=10), showlegend=False, hovermode="x unified")
            cc2.plotly_chart(f2, key=f"gr_r_{lbl}")
    with tabs[2]:
        section_title(f"每{unit_g}全指标汇总", ("规模指标是日均，" if use_avg else "") + "比率已用分子 / 分母重算；环比 = vs 上一期")
        summ = agg[["期间", "天数"]].copy()
        for lbl, (cname, kind, is_ratio, key) in metric_opts.items():
            summ[lbl] = agg[cname]
            summ[f"{lbl}｜环比{'（个百分点）' if (is_ratio and kind == 'pct') else '%'}"] = chg_series(agg[cname], kind, is_ratio)
        st.dataframe(round_num(summ.iloc[::-1], 3), hide_index=True, height=460)
        st.download_button(f"⬇️ 下载每{unit_g}汇总 CSV", summ.to_csv(index=False).encode("utf-8-sig"),
                           f"nailvesta_{PROF['name']}_每{unit_g}汇总.csv", "text/csv", key="dl_summ")


def _raw_unit_key(name: str) -> str:
    """原始字段名 → 借一个 K 短名来带单位（曝光 IM、点击 次、金额 $）。"""
    n = name.lower()
    if "impression" in n or "曝光" in name:
        return "impr"
    if "click" in n or "点击" in name:
        return "clicks"
    if any(x in n for x in ("gmv", "refund", "commission", "shipping", "tax", "merchandise", "flat fee")):
        return "gmv"
    return ""


def section_view(kind: str, gran_key: str):
    cat = catalog[catalog["分类"] == kind] if not catalog.empty else pd.DataFrame()
    head = st.container()
    secs = sorted(cat["区段"].unique()) if not cat.empty else []
    c1, c2 = st.columns([2, 1])
    sec = c1.selectbox("区段（数据表里的一组字段）", secs, key=f"sec_{kind}") if secs else None
    gran2 = c2.radio("粒度", PROF["periods"], horizontal=True, index=2, key=f"g_{kind}_{PROF['key']}")
    unit2 = {"日": "天", "周": "周", "月": "月", "季": "季"}[gran2]
    with head:
        page_header(f"{kind}指标", PROF["front_desc" if kind == "前端" else "back_desc"].replace("**", ""),
                    [f"{PROF['name']}数据", f"{kind}共 {len(cat)} 个字段", f"粒度：{gran2}" + ("（日均）" if gran2 != "日" else "")])
    tabs = st.tabs(["结论", "图表", "明细表"])
    gg = pd.DataFrame()
    if sec:
        src = df
        full_level = "层级" in cat.columns and (cat.loc[cat["区段"] == sec, "层级"] == "全量").all()
        if full_level:
            src = DAYTOT[DAYTOT["日期"].isin(df["日期"].unique())]
        opts = cat[(cat["区段"] == sec) & (cat["类型"] != "比率")]["key"].tolist()
        opts = [k for k in opts if k in src.columns]
        labels = {k: k.split("::", 1)[-1] for k in opts}
        imp_first = sorted(opts, key=lambda k: 0 if _raw_unit_key(labels[k]) == "impr" else 1)
        with tabs[1]:
            chosen = st.multiselect("指标", opts, default=imp_first[:4], format_func=lambda k: labels.get(k, k),
                                    key=f"m_{kind}")
            if full_level:
                st.caption("这个区段是原始表的每日合计（不带链接维度）：可加的字段求和；比率 / 均值类按播放、观众或曝光加权。")
        if chosen:
            d = src[["日期"] + chosen].copy()
            d["_p"] = to_period(d["日期"], gran2)
            gg = d.groupby("_p")[chosen].sum()
            nd = d.groupby("_p")["日期"].nunique()
            if gran2 != "日":
                gg = gg.div(nd, axis=0)
            gg = gg.reset_index().rename(columns={"_p": "期间"})
            with tabs[0]:
                if len(gg) >= 2:
                    cl, bl = f"最近一{unit2}（{gg['期间'].iloc[-1]:%m-%d} 起）", f"上一{unit2}"
                    cards = []
                    for k in chosen:
                        uk = _raw_unit_key(labels[k])
                        c_, b_ = gg[k].iloc[-1], gg[k].iloc[-2]
                        cards.append({"level": level_of(pct(c_, b_)), "title": labels[k],
                                      "tag": "日均" if gran2 != "日" else "",
                                      "body": say_change(labels[k], uk, c_, b_, cl, bl)})
                    insight_cards(cards, cols=2)
                else:
                    st.info("只有一期数据，没法比。")
            with tabs[1]:
                fig = px.line(gg, x="期间", y=chosen, markers=True, labels={"value": "", "期间": ""})
                fig.for_each_trace(lambda t: t.update(name=labels.get(t.name, t.name)))
                fig.update_layout(height=400, legend_title="", hovermode="x unified", title="日均" if gran2 != "日" else "")
                st.plotly_chart(fig, key=f"sv_{kind}")
            with tabs[2]:
                tbl = gg.copy()
                tbl.insert(1, "天数", nd.values)
                for c in chosen:
                    tbl[labels[c] + f" 环比%（vs 上一{unit2}）"] = tbl[c].pct_change() * 100
                tbl = tbl.rename(columns={c: labels[c] + ("（日均）" if gran2 != "日" else "") for c in chosen})
                st.dataframe(round_num(tbl.iloc[::-1]), hide_index=True)
    rat = [nm for nm, v in DERIVED.items() if v[3] == kind]
    agg2 = aggregate(df, gran2)
    rat = [r for r in rat if r in agg2.columns]
    if rat:
        with tabs[0]:
            if len(agg2) >= 2:
                section_title(f"{kind}比率", "由分子 / 分母合计后重算")
                insight_cards([{"level": level_of(agg2[r].iloc[-1] - agg2[r].iloc[-2],
                                                  thr=0.1 if DERIVED[r][2] == "pct" else 5),
                                "title": r, "body": say_change(r, r, agg2[r].iloc[-1], agg2[r].iloc[-2],
                                                               f"最近一{unit2}", f"上一{unit2}")} for r in rat], cols=3)
        with tabs[1]:
            section_title(f"{kind}比率指标", "由分子 / 分母合计后重算")
            fr = px.line(agg2, x="期间", y=rat, markers=True, labels={"value": "", "期间": ""})
            fr.update_layout(height=340, legend_title="", hovermode="x unified")
            st.plotly_chart(fr, key=f"sv_r_{kind}")


def page_frontend():
    section_view("前端", "front")


def page_backend():
    section_view("后端", "back")

# ──────────────────────────────────────────────────────────────────────
# 6c. 异常预警雷达：所有细分层级（大盘 / 渠道 / 每条链接 / 每个达人）× 各项指标，曝光是索引列
# ──────────────────────────────────────────────────────────────────────
_METRIC_RANK = {"曝光": 0, "点击": 1, "CTR": 2, "CTOR": 3, "订单": 4, "GMV": 5, "播放": 6}
_LEVEL_RANK = {"大盘": 0, "渠道": 1, "链接": 2, "达人": 3}
_METRIC_KEY = {"曝光": "impr", "点击": "clicks", "CTR": "CTR 点击率", "CTOR": "CTOR 点击成单率", "订单": "sku",
               "GMV": "gmv", "播放": "v_views"}
_RATIO_METRICS = {"CTR", "CTOR"}
# 基线（前 N 天中位数）低于门槛的不报：量太小时一点波动就是几十个百分点，报了也没意义
_FLOOR = {"链接": {"曝光": 1000, "点击": 30, "订单": 3, "GMV": 50, "CTR": 1000, "CTOR": 50},
          "渠道": {"曝光": 5000, "点击": 100, "订单": 3, "GMV": 50, "CTR": 5000, "CTOR": 100},
          "达人": {"曝光": 1000, "点击": 30, "订单": 3, "GMV": 50, "CTR": 1000, "播放": 1000},
          "大盘": {}}


def _scan_wide(w: pd.DataFrame, win: int, k: float, pthr: float, floor: float = 0.0,
               den: pd.DataFrame | None = None, den_floor: float = 0.0) -> pd.DataFrame:
    """w：行 = 日期，列 = 对象。基线 = 前 win 天中位数（只看过去）；|偏离| ≥ pthr% 且 |稳健 z| ≥ k 才算异常。"""
    mp = max(3, win // 2)
    med = w.shift(1).rolling(win, min_periods=mp).median()
    dev = w - med
    mad = dev.abs().shift(1).rolling(win * 2, min_periods=mp).median()
    z = dev / (mad * 1.4826).where(mad > 0)
    rel = (w / med.where(med != 0) - 1) * 100
    ok = med.notna() & w.notna()
    ok &= (den.shift(1).rolling(win, min_periods=mp).median() >= den_floor) if den is not None else (med.abs() >= floor)
    flag = (ok & (rel.abs() >= pthr) & ((z.abs() >= k) | mad.eq(0))).to_numpy()
    ii, jj = np.nonzero(flag)
    if not len(ii):
        return pd.DataFrame(columns=["日期", "obj", "当天", "基线", "偏离%", "z"])
    return pd.DataFrame({"日期": w.index[ii], "obj": w.columns[jj], "当天": w.to_numpy()[ii, jj],
                         "基线": med.to_numpy()[ii, jj], "偏离%": rel.to_numpy()[ii, jj], "z": z.to_numpy()[ii, jj]})


def _wide(frame: pd.DataFrame, obj: str, col: str, dates) -> pd.DataFrame:
    """长表 → 行日期 × 列对象；数据里有的日子缺的对象记 0，整天没数据的日子保持空（不参与基线）。"""
    present = pd.Index(sorted(frame["日期"].unique()))
    g = frame.groupby(["日期", obj])[col].sum().unstack(obj).reindex(present).fillna(0.0)
    return g.reindex(dates)


def _link_names_all() -> dict:
    nm = df_all.groupby("product_id")["product_name"].last().map(lambda s: short_name(s) or "(无名称)")
    dup = nm.duplicated(keep=False)
    nm[dup] = [f"{n} ·{str(p)[-4:]}" for p, n in nm[dup].items()]
    return nm.to_dict()


def radar_all(win: int = 7, k: float = 3.0, pthr: float = 30.0) -> pd.DataFrame:
    """全部层级 × 全部指标的异常（全历史，每天一套）。会话里缓存，切页面不重算。"""
    D = df_all
    memo_key = (PROF["key"], len(D), str(D["日期"].max()), win, k, pthr, id(CRB))
    memo = st.session_state.setdefault("_radar_memo", {})
    if memo_key in memo:
        return memo[memo_key]
    dates = pd.date_range(D["日期"].min(), D["日期"].max(), freq="D")
    parts = []

    def add(res, level, metric, names=None):
        if len(res):
            res = res.assign(层级=level, 指标=metric)
            res["对象"] = res["obj"].map(names) if names else res["obj"]
            parts.append(res)

    def scan_level(level, frame, obj, cols: dict, names=None):
        """cols：指标 → 列名（可加的）；CTR / CTOR 用 点击/曝光、订单/点击 现算。"""
        wides = {m: _wide(frame, obj, c, dates) for m, c in cols.items() if c in frame.columns}
        fl = _FLOOR.get(level, {})
        for m, w in wides.items():
            add(_scan_wide(w, win, k, pthr, floor=fl.get(m, 0)), level, m, names)
        if "点击" in wides and "曝光" in wides:
            add(_scan_wide(wides["点击"] / wides["曝光"].where(wides["曝光"] > 0) * 100, win, k, pthr,
                           den=wides["曝光"], den_floor=fl.get("CTR", 0)), level, "CTR", names)
        if "订单" in wides and "点击" in wides:
            add(_scan_wide(wides["订单"] / wides["点击"].where(wides["点击"] > 0) * 100, win, k, pthr,
                           den=wides["点击"], den_floor=fl.get("CTOR", 0)), level, "CTOR", names)

    base_cols = {"曝光": K["impr"], "点击": K["clicks"], "订单": K["sku"], "GMV": K["gmv"]}
    # 大盘
    tot = D.assign(_all=SCOPE)
    scan_level("大盘", tot, "_all", base_cols)
    # 渠道：每个渠道的曝光 / 点击 / CTR / GMV
    rows = []
    for lbl, ik in CH_IMPR.items():
        sub = pd.DataFrame({"日期": D["日期"], "渠道": lbl,
                            "曝光": D[K[ik]] if K.get(ik) in D.columns else 0.0,
                            "点击": D[K[CH_CLICK[lbl]]] if K.get(CH_CLICK.get(lbl)) in D.columns else np.nan,
                            "GMV": D[K[dict(CHANNELS)[lbl]]] if K.get(dict(CHANNELS).get(lbl)) in D.columns else 0.0})
        rows.append(sub)
    chf = pd.concat(rows, ignore_index=True)
    scan_level("渠道", chf, "渠道", {"曝光": "曝光", "点击": "点击", "GMV": "GMV"})
    # 每条链接
    scan_level("链接", D, "product_id", base_cols, names=_link_names_all())
    # 每个达人（达人数据）
    if PROF["key"] == "cr" and CRB is not None:
        cd = CRB["creator_day"]
        cd = cd[cd["日期"].between(dates.min(), dates.max())]
        big = cd.groupby("达人")["曝光"].max()
        cd = cd[cd["达人"].isin(big[big >= _FLOOR["达人"]["曝光"]].index) | cd["达人"].isin(
            cd.groupby("达人")["GMV"].max().loc[lambda s: s >= _FLOOR["达人"]["GMV"]].index)]
        scan_level("达人", cd, "达人", {"曝光": "曝光", "点击": "点击", "订单": "订单", "GMV": "GMV", "播放": "播放"},
                   names={n: f"@{n}" for n in cd["达人"].unique()})
    out = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(
        columns=["日期", "obj", "当天", "基线", "偏离%", "z", "层级", "指标", "对象"])
    out["方向"] = np.where(out["偏离%"] >= 0, "上升", "下跌")
    out["_m"] = out["指标"].map(_METRIC_RANK).fillna(9)
    out["_l"] = out["层级"].map(_LEVEL_RANK).fillna(9)
    out["_abs"] = (out["当天"] - out["基线"]).abs()
    out = out.sort_values(["日期", "_m", "_l", "_abs"], ascending=[False, True, True, False]).reset_index(drop=True)
    memo[memo_key] = out
    return out


def _prev_days(day, win: int) -> list:
    return [d for d in sorted(df_all["日期"].unique()) if d < day][-win:]


def alert_headline(a: pd.Series, win: int) -> str:
    """一条预警的大白话：「09-21 是 42,110 IM，前 7 天中位数 71,300 IM，少了 29,190 IM（-41%）。」"""
    key, day = _METRIC_KEY.get(a["指标"], ""), pd.Timestamp(a["日期"])
    cur, base = a["当天"], a["基线"]
    if a["指标"] in _RATIO_METRICS:
        return (f"{day:%m-%d} 是 {cur:.2f}%，前 {win} 天中位数 {base:.2f}%，"
                f"{'高' if cur > base else '低'}了 {abs(cur - base):.2f} 个百分点（相对 {a['偏离%']:+.0f}%）。")
    return (f"{day:%m-%d} 是 {fv(key, cur)}，前 {win} 天中位数 {fv(key, base)}，"
            f"{'多' if cur > base else '少'}了 {fv(key, abs(cur - base))}（{a['偏离%']:+.0f}%）。")


def _day_masks(day, prev):
    return df_all["日期"] == day, df_all["日期"].isin(prev)


def alert_reason(a: pd.Series, win: int) -> list[str]:
    """一条预警为什么：掉 / 涨在哪个渠道、同一天漏斗其它环节、哪几条链 / 哪几个达人在动。"""
    day = pd.Timestamp(a["日期"])
    prev = _prev_days(day, win)
    n = len(prev)
    if not n:
        return []
    m, lv, obj = a["指标"], a["层级"], a["obj"]
    falling = a["当天"] < a["基线"]
    lines = []
    if lv == "达人":
        cd = CRB["creator_day"]
        x = cd[cd["达人"] == obj]
        cur, base = x[x["日期"] == day], x[x["日期"].isin(prev)]
        bits = []
        for col, key in (("新发视频", "new_vid"), ("播放", "v_views"), ("曝光", "impr"), ("订单", "sku"), ("GMV", "gmv")):
            if col == m:
                continue
            bits.append(f"{col} {fv(key, cur[col].sum())}（前{n}天日均 {fv(key, base[col].sum() / n)}）")
        lines.append("同一天：" + "、".join(bits))
        ml = _main_link_of([obj], [day] + prev)
        if len(ml):
            lines.append(f"主推链接：{ml.iloc[0]}")
        return lines
    cm, bm = _day_masks(day, prev)
    if lv == "链接":
        cm, bm = cm & (df_all["product_id"] == obj), bm & (df_all["product_id"] == obj)
    cur = sums_of(df_all[cm])
    base = {kk: v / n for kk, v in sums_of(df_all[bm]).items()}
    if lv in ("大盘", "链接"):
        chmap = CH_CLICK if m == "点击" else ({l: kk for l, kk in CHANNELS} if m in ("GMV", "订单") else CH_IMPR)
        ukey = "clicks" if m == "点击" else ("gmv" if m in ("GMV", "订单") else "impr")
        chd = sorted(((l, cur.get(kk, 0) - base.get(kk, 0)) for l, kk in chmap.items() if kk in cur),
                     key=lambda x: x[1] if falling else -x[1])
        if chd and chd[0][1] != 0:
            lines.append(f"主要{'掉' if chd[0][1] < 0 else '涨'}在「{chd[0][0]}」渠道（{fdelta(ukey, chd[0][1])}，"
                         f"跟前 {n} 天日均比）")
        funnel = [("曝光", "impr"), ("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"), ("GMV", "gmv")]
        lines.append("同一天：" + "、".join(f"{nm} {fv(kk, value_of(cur, kk))}（前{n}天日均 {fv(kk, value_of(base, kk))}）"
                                         for nm, kk in funnel if nm != m))
        if lv == "链接" and PROF["key"] == "ops" and CRB is not None and chd and chd[0][0].split()[0] == "达人":
            lines.append("达人侧：" + cr_link_reason(obj, [day], prev))
        if lv == "链接" and PROF["key"] == "cr" and CRB is not None:
            col = "挂车曝光" if m in ("曝光", "点击", "CTR") else "挂车GMV"
            t = _dim_delta(CRB["lc"][CRB["lc"]["product_id"] == obj], "达人", col, [day], prev)
            t = (t[t["Δ"] < 0] if falling else t[t["Δ"] > 0].iloc[::-1]).head(2)
            ck = "impr" if col == "挂车曝光" else "gmv"
            if len(t):
                lines.append("挂车达人：" + "、".join(f"@{nn} {fv(ck, r['基期'])}→{fv(ck, r['本期'])}"
                                                   for nn, r in t.iterrows()))
    if lv in ("大盘", "渠道"):
        key = _METRIC_KEY[m]
        if lv == "渠道":
            key = {"曝光": CH_IMPR, "点击": CH_CLICK, "GMV": dict(CHANNELS)}.get(m, CH_IMPR).get(obj, key)
        gc, gb, names = link_sums(df_all, cm, bm, n)
        if key in DERIVED:
            t = ratio_contrib(gc, gb, names, key)
            dcol = "贡献"
        else:
            t = pd.DataFrame({"链接": names, "Δ": gc[K[key]] - gb[K[key]]}) if K.get(key) in gc.columns else pd.DataFrame()
            dcol = "Δ"
        if len(t):
            t = (t[t[dcol] < 0].sort_values(dcol) if falling else t[t[dcol] > 0].sort_values(dcol, ascending=False)).head(3)
            if len(t):
                fmt_ = (lambda v: f"{v:+.2f} 个百分点") if key in DERIVED else (lambda v: fdelta(key, v))
                lines.append(f"{'拖累' if falling else '拉动'}最多的链接：" +
                             "、".join(f"{r['链接']}（{fmt_(r[dcol])}）" for _, r in t.iterrows()))
    return lines


def link_multi_attr(D: pd.DataFrame, cm, bm, scale: float, sort_key: str = "impr", topn: int = 15) -> pd.DataFrame:
    """按链接的全指标归因：每条链接 曝光 / 点击 / CTR / CTOR / 订单 / 客单价 / GMV 各变了多少（基期按天数折算）。"""
    gc, gb, names = link_sums(D, cm, bm, scale)
    col = lambda g, k: g[K[k]] if K.get(k) in g.columns else pd.Series(0.0, index=g.index)  # noqa: E731
    t = pd.DataFrame({"链接": names})
    for nm, k in (("曝光", "impr"), ("点击", "clicks"), ("订单", "sku"), ("GMV", "gmv")):
        t[f"_{k}_b"], t[f"_{k}_c"] = col(gb, k), col(gc, k)
        t[f"Δ{nm}"] = t[f"_{k}_c"] - t[f"_{k}_b"]
    t["_ctr_b"] = np.where(t["_impr_b"] > 0, t["_clicks_b"] / t["_impr_b"].replace(0, np.nan) * 100, np.nan)
    t["_ctr_c"] = np.where(t["_impr_c"] > 0, t["_clicks_c"] / t["_impr_c"].replace(0, np.nan) * 100, np.nan)
    t["_ctor_b"] = np.where(t["_clicks_b"] > 0, t["_sku_b"] / t["_clicks_b"].replace(0, np.nan) * 100, np.nan)
    t["_ctor_c"] = np.where(t["_clicks_c"] > 0, t["_sku_c"] / t["_clicks_c"].replace(0, np.nan) * 100, np.nan)
    t["_aov_b"] = np.where(t["_sku_b"] > 0, t["_gmv_b"] / t["_sku_b"].replace(0, np.nan), np.nan)
    t["_aov_c"] = np.where(t["_sku_c"] > 0, t["_gmv_c"] / t["_sku_c"].replace(0, np.nan), np.nan)
    chd = pd.DataFrame({l: col(gc, k) - col(gb, k) for l, k in CH_IMPR.items()})
    t["_main_ch"] = np.where(t["Δ曝光"] < 0, chd.idxmin(axis=1), chd.idxmax(axis=1)) if len(chd.columns) else ""
    sk = {"impr": "Δ曝光", "clicks": "Δ点击", "sku": "Δ订单", "orders": "Δ订单", "gmv": "ΔGMV"}.get(sort_key, "Δ曝光")
    t = t[(t[["_impr_b", "_impr_c", "_gmv_b", "_gmv_c"]].abs().sum(axis=1)) > 0]
    t = pd.concat([t.sort_values(sk).head(topn), t.sort_values(sk, ascending=False).head(topn // 2)])
    t = t[~t.index.duplicated()]
    out = pd.DataFrame({
        "链接": t["链接"],
        "曝光 前→今": [f"{fv('impr', b)} → {fv('impr', c)}" for b, c in zip(t["_impr_b"], t["_impr_c"])],
        "Δ曝光": [fdelta("impr", v) for v in t["Δ曝光"]],
        "主要掉/涨在": [str(x).split()[0] if isinstance(x, str) else "" for x in t["_main_ch"]],
        "Δ点击": [fdelta("clicks", v) for v in t["Δ点击"]],
        "CTR 前→今": [f"{fv('CTR 点击率', b)} → {fv('CTR 点击率', c)}" for b, c in zip(t["_ctr_b"], t["_ctr_c"])],
        "CTOR 前→今": [f"{fv('CTOR 点击成单率', b)} → {fv('CTOR 点击成单率', c)}" for b, c in zip(t["_ctor_b"], t["_ctor_c"])],
        "Δ订单": [fdelta("sku", v) for v in t["Δ订单"]],
        "客单价 前→今": [f"{fv('AOV 客单价', b)} → {fv('AOV 客单价', c)}" for b, c in zip(t["_aov_b"], t["_aov_c"])],
        "ΔGMV": [fdelta("gmv", v) for v in t["ΔGMV"]],
    }, index=t.index)
    out["_sort"] = t[sk].values
    return out


def creator_multi_attr(cur_days, base_days, sort_col: str = "曝光", topn: int = 15) -> pd.DataFrame:
    """按达人的全指标归因（Creators 表）：曝光 / 点击 / CTR / 播放 / 订单 / GMV / 新发视频 各变了多少。"""
    cd = CRB["creator_day"]
    parts = {c: _dim_delta(cd, "达人", c, cur_days, base_days) for c in ("曝光", "点击", "播放", "订单", "GMV", "新发视频")}
    idx = parts[sort_col].index if sort_col in parts else parts["曝光"].index
    for p in parts.values():
        idx = idx.union(p.index)
    g = lambda c, w: parts[c][w].reindex(idx).fillna(0.0)  # noqa: E731
    t = pd.DataFrame(index=idx)
    for c in parts:
        t[f"_{c}_b"], t[f"_{c}_c"] = g(c, "基期"), g(c, "本期")
        t[f"Δ{c}"] = t[f"_{c}_c"] - t[f"_{c}_b"]
    sk = f"Δ{sort_col}" if f"Δ{sort_col}" in t else "Δ曝光"
    t = pd.concat([t.sort_values(sk).head(topn), t.sort_values(sk, ascending=False).head(topn // 2)])
    t = t[~t.index.duplicated()]
    ctr = lambda cl, im: cl / im * 100 if im else np.nan  # noqa: E731
    main = _main_link_of(t.index, list(cur_days) + list(base_days)).reindex(t.index).fillna("—")
    return pd.DataFrame({
        "达人": "@" + t.index.astype(str),
        "曝光 前→今": [f"{fv('impr', b)} → {fv('impr', c)}" for b, c in zip(t["_曝光_b"], t["_曝光_c"])],
        "Δ曝光": [fdelta("impr", v) for v in t["Δ曝光"]],
        "Δ点击": [fdelta("clicks", v) for v in t["Δ点击"]],
        "CTR 前→今": [f"{fv('CTR 点击率', ctr(a, b))} → {fv('CTR 点击率', ctr(c, d))}"
                      for a, b, c, d in zip(t["_点击_b"], t["_曝光_b"], t["_点击_c"], t["_曝光_c"])],
        "Δ播放": [fdelta("v_views", v) for v in t["Δ播放"]],
        "Δ订单": [fdelta("sku", v) for v in t["Δ订单"]],
        "ΔGMV": [fdelta("gmv", v) for v in t["ΔGMV"]],
        "新发视频 前→今": [f"{b:.1f} → {c:.0f}" for b, c in zip(t["_新发视频_b"], t["_新发视频_c"])],
        "主推链接": main.values,
    })


def object_metric_table(a: pd.Series, win: int) -> pd.DataFrame:
    """一条预警对应对象的全部指标：当天 / 前 N 天日均 / 变化 / 大白话。"""
    day = pd.Timestamp(a["日期"])
    prev = _prev_days(day, win)
    n = len(prev)
    lv, obj, bad = a["层级"], a["obj"], a["指标"]
    day_lbl, base_lbl = f"{day:%m-%d}", f"前 {n} 天日均"
    rows = []
    if lv == "达人":
        cd = CRB["creator_day"]
        x = cd[cd["达人"] == obj]
        cur, base = x[x["日期"] == day].sum(numeric_only=True), x[x["日期"].isin(prev)].sum(numeric_only=True) / max(n, 1)
        spec = [("曝光", "impr", cur["曝光"], base["曝光"]), ("点击", "clicks", cur["点击"], base["点击"]),
                ("CTR", "CTR 点击率", cur["点击"] / cur["曝光"] * 100 if cur["曝光"] else np.nan,
                 base["点击"] / base["曝光"] * 100 if base["曝光"] else np.nan),
                ("播放", "v_views", cur["播放"], base["播放"]), ("订单", "sku", cur["订单"], base["订单"]),
                ("GMV", "gmv", cur["GMV"], base["GMV"]), ("新发视频", "new_vid", cur["新发视频"], base["新发视频"]),
                ("预估佣金", "comm", cur["佣金"], base["佣金"])]
    else:
        cm, bm = _day_masks(day, prev)
        if lv == "链接":
            cm, bm = cm & (df_all["product_id"] == obj), bm & (df_all["product_id"] == obj)
        cur = sums_of(df_all[cm])
        base = {kk: v / max(n, 1) for kk, v in sums_of(df_all[bm]).items()}
        keys = [("曝光", "impr"), ("点击", "clicks"), ("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"),
                ("订单", "sku"), ("客单价", "AOV 客单价"), ("GMV", "gmv")]
        if lv == "渠道":
            keys = [("曝光", CH_IMPR.get(obj)), ("点击", CH_CLICK.get(obj)), ("GMV", dict(CHANNELS).get(obj))]
            keys = [(nm, kk) for nm, kk in keys if kk and kk in cur]
        else:
            keys += [(f"{l} 曝光", kk) for l, kk in CH_IMPR.items() if kk in cur]
        spec = [(nm, kk, value_of(cur, kk), value_of(base, kk)) for nm, kk in keys]
    for nm, kk, c, b in spec:
        d = c - b if pd.notna(c) and pd.notna(b) else np.nan
        rows.append({"": "⚠️" if nm == bad else "", "指标": nm, day_lbl: fv(kk, c, 2 if kk in DERIVED else 0),
                     base_lbl: fv(kk, b, 2 if kk in DERIVED else 0), "变化": fdelta(kk, d),
                     "变化%": "—" if (kk in DERIVED and DERIVED[kk][2] == "pct") or pd.isna(pct(c, b)) else f"{pct(c, b):+.1f}%",
                     "大白话": say_change(nm, kk, c, b, day_lbl, base_lbl), "_rel": pct(c, b) if pd.notna(d) else np.nan})
    return pd.DataFrame(rows)


def _alert_label(a: pd.Series) -> str:
    arrow = "🔴" if a["方向"] == "下跌" else "🟢"
    return f"{arrow} {a['层级']} · {a['对象']} · {a['指标']} {a['偏离%']:+.0f}%"


def page_anomaly():
    D = df_all
    last = D["日期"].max()
    head = st.container()
    c0, c1, c2, c3 = st.columns([1.4, 1, 1, 1])
    day = pd.Timestamp(c0.date_input("看哪一天", value=last.date(), min_value=D["日期"].min().date(),
                                     max_value=last.date(), key=f"rad_day_{PROF['key']}"))
    win = c1.slider("基线 = 前 N 天中位数", 3, 21, 7, key="rad_win")
    kk = c2.slider("灵敏度 k（越小越敏感）", 1.0, 5.0, 3.0, step=0.5, key="rad_k")
    pthr = c3.slider("偏离至少 %", 10, 80, 30, step=5, key="rad_p")
    with head:
        page_header("异常预警", f"每一条链接、每个渠道{'、每个达人' if PROF['key'] == 'cr' else ''}逐项检测；曝光是索引列，"
                             "先看曝光，再看点击、CTR、CTOR、订单、GMV",
                    [f"{PROF['name']}数据", f"看 {day:%m-%d}", f"基线 = 前 {win} 天中位数", f"偏离 ≥ {pthr}% 且 z ≥ {kk}"])
    R = radar_all(win, kk, pthr)
    today = R[R["日期"] == day]
    # 大盘曝光不管异不异常都要说一句
    cm, bm = _day_masks(day, _prev_days(day, win))
    tot_now = D.loc[cm, K["impr"]].sum()
    tot_med = D[bm].groupby("日期")[K["impr"]].sum().median() if bm.any() else np.nan
    rel0 = pct(tot_now, tot_med)
    n_dn, n_up = int((today["方向"] == "下跌").sum()), int((today["方向"] == "上升").sum())
    n_im = int((today["指标"] == "曝光").sum())
    summary_pills([("down" if n_dn else "flat", f"🔴 {n_dn} 条下跌预警"), ("up" if n_up else "flat", f"🟢 {n_up} 条上升预警"),
                   ("warn" if n_im else "flat", f"其中曝光类 {n_im} 条"),
                   (level_of(rel0, thr=pthr), f"{SCOPE}曝光 {fv('impr', tot_now)}（vs 前{win}天中位数 {rel0:+.1f}%）"
                    if pd.notna(rel0) else "")])
    tabs = st.tabs(["结论", "全指标归因", "图表", "明细表", "按日/周/月看大盘"])

    with tabs[0]:
        if today.empty:
            insight_cards([{"level": "flat", "title": f"{day:%m-%d} 没有触发预警",
                            "body": [f"{SCOPE}曝光 {fv('impr', tot_now)}，前 {win} 天中位数 {fv('impr', tot_med)}"
                                     + (f"（{rel0:+.1f}%）" if pd.notna(rel0) else "") + "。",
                                     f"所有链接{'、渠道、达人' if PROF['key'] == 'cr' else '、渠道'}的各项指标都在正常波动范围里"
                                     f"（偏离 < {pthr}% 或者没超过平时的波动幅度）。"]}])
        else:
            show = today.head(12)
            cards = []
            for _, a in show.iterrows():
                cards.append({"level": "down" if a["方向"] == "下跌" else "up",
                              "title": f"{a['对象']} · {a['指标']}", "tag": a["层级"],
                              "body": [alert_headline(a, win)] + alert_reason(a, win)})
            insight_cards(cards)
            if len(today) > len(show):
                st.caption(f"还有 {len(today) - len(show)} 条，见「明细表」。排序：先曝光、再点击 / CTR / CTOR / 订单 / GMV；"
                           "同一指标里先大盘、渠道，再链接、达人，量大的在前。")
        last14 = R[R["日期"] > day - pd.Timedelta(days=14)]
        last14 = last14[last14["日期"] <= day]
        if len(last14):
            cnt = last14.groupby(["日期", "方向"]).size().unstack(fill_value=0).reindex(
                pd.date_range(day - pd.Timedelta(days=13), day), fill_value=0)
            fc = go.Figure()
            for dname, color in (("下跌", "#D9534F"), ("上升", "#2E9E7A")):
                if dname in cnt:
                    fc.add_bar(x=cnt.index, y=cnt[dname], name=f"{dname}预警", marker_color=color)
            fc.update_layout(barmode="stack", height=240, margin=dict(t=30, b=10), title="近 14 天每天的预警条数",
                             legend=dict(orientation="h", y=1.2))
            st.plotly_chart(fc, key="rad_cnt")

    with tabs[1]:
        if today.empty:
            st.info("这一天没有预警。可以在「按日/周/月看大盘」里看大盘层面的深度归因。")
        else:
            opts = list(today.index)
            pick = st.selectbox("选一条预警", opts, format_func=lambda i: _alert_label(today.loc[i]), key="rad_pick")
            a = today.loc[pick]
            insight_cards([{"level": "down" if a["方向"] == "下跌" else "up", "title": f"{a['对象']} · {a['指标']}",
                            "tag": a["层级"], "body": [alert_headline(a, win)] + alert_reason(a, win)}], cols=1)
            section_title(f"{a['对象']}：这一天的全部指标", f"跟前 {len(_prev_days(day, win))} 天日均比；⚠️ = 触发预警的指标")
            om = object_metric_table(a, win)
            st.dataframe(_style_rows(om.drop(columns="_rel"), om["_rel"], ["变化", "变化%"]), hide_index=True)
            prev = _prev_days(day, win)
            if a["层级"] in ("大盘", "渠道"):
                sort_key = {"曝光": "impr", "点击": "clicks", "订单": "sku", "GMV": "gmv"}.get(a["指标"], "impr")
                section_title("按链接拆：每条链接的全部指标怎么变的", f"按「{a['指标']}」变化排序；基期 = 前 {len(prev)} 天日均")
                cmx, bmx = _day_masks(day, prev)
                la = link_multi_attr(df_all, cmx, bmx, len(prev), sort_key)
                st.dataframe(la.drop(columns="_sort"), hide_index=True)
                if PROF["key"] == "cr":
                    section_title("按达人拆：每个达人的全部指标怎么变的", "Creators 表；按曝光变化排序")
                    st.dataframe(creator_multi_attr([day], prev, {"GMV": "GMV", "订单": "订单"}.get(a["指标"], "曝光")),
                                 hide_index=True)
            elif a["层级"] == "链接" and PROF["key"] == "cr":
                section_title("挂这条链的达人：挂车曝光 / 播放 / 订单 / GMV 怎么变的", "挂车关系，近似")
                x = CRB["lc"][CRB["lc"]["product_id"] == a["obj"]]
                parts = {c: _dim_delta(x, "达人", c, [day], prev) for c in ("挂车曝光", "播放", "视频订单", "挂车GMV", "新发视频")}
                idx = parts["挂车曝光"].index
                for p_ in parts.values():
                    idx = idx.union(p_.index)
                t = pd.DataFrame({f"Δ{c}": parts[c]["Δ"].reindex(idx).fillna(0) for c in parts})
                t = t.sort_values("Δ挂车曝光").head(15)
                st.dataframe(pd.DataFrame({"达人": "@" + t.index.astype(str),
                                           "Δ挂车曝光": [fdelta("impr", v) for v in t["Δ挂车曝光"]],
                                           "Δ播放": [fdelta("v_views", v) for v in t["Δ播放"]],
                                           "Δ视频订单": [fdelta("sku", v, 1) for v in t["Δ视频订单"]],
                                           "Δ挂车GMV": [fdelta("gmv", v) for v in t["Δ挂车GMV"]],
                                           "Δ新发视频": [f"{v:+.1f}" for v in t["Δ新发视频"]]}), hide_index=True)
            elif a["层级"] == "链接" and PROF["key"] == "ops" and CRB is not None:
                st.caption("达人侧：" + cr_link_reason(a["obj"], [day], prev))
            elif a["层级"] == "达人":
                section_title(f"{a['对象']} 推的链接（挂车）", "挂车曝光 / 挂车 GMV 变化")
                x = CRB["lc"][CRB["lc"]["达人"] == a["obj"]]
                t = _dim_delta(x, "product_id", "挂车曝光", [day], prev)
                g2 = _dim_delta(x, "product_id", "挂车GMV", [day], prev)
                names = _link_names_all()
                st.dataframe(pd.DataFrame({"链接": [names.get(p, p) for p in t.index],
                                           "挂车曝光 前→今": [f"{fv('impr', b)} → {fv('impr', c)}" for b, c in zip(t["基期"], t["本期"])],
                                           "Δ挂车曝光": [fdelta("impr", v) for v in t["Δ"]],
                                           "Δ挂车GMV": [fdelta("gmv", g2["Δ"].get(p, 0)) for p in t.index]}), hide_index=True)

    with tabs[2]:
        objs = today if len(today) else R[R["日期"] <= day].head(30)
        if objs.empty:
            st.info("没有可画的预警。")
        else:
            pick2 = st.selectbox("选一条预警看走势", list(objs.index), format_func=lambda i: _alert_label(objs.loc[i]),
                                 key="rad_pick2")
            a = objs.loc[pick2]
            ser = _alert_series(a)
            ser = ser[ser.index >= day - pd.Timedelta(days=60)]
            ser = ser[ser.index <= day]
            med = ser.shift(1).rolling(win, min_periods=max(3, win // 2)).median()
            hits = R[(R["层级"] == a["层级"]) & (R["obj"] == a["obj"]) & (R["指标"] == a["指标"])]
            hits = hits[hits["日期"].isin(ser.index)]
            key = _METRIC_KEY.get(a["指标"], "")
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=ser.index, y=ser.values, mode="lines+markers", name=a["指标"],
                                     line=dict(color="#C2416B", width=2)))
            fig.add_trace(go.Scatter(x=med.index, y=med.values, mode="lines", name=f"基线（前{win}天中位数）",
                                     line=dict(color="#8A727C", dash="dash")))
            fig.add_trace(go.Scatter(x=hits["日期"], y=hits["当天"], mode="markers", name="预警",
                                     marker=dict(size=12, color=np.where(hits["方向"] == "下跌", "#D9534F", "#2E9E7A"),
                                                 symbol="circle-open", line=dict(width=3))))
            unit = unit_of(key) if key not in DERIVED else "%"
            fig.update_layout(height=380, hovermode="x unified", margin=dict(t=40, b=10),
                              title=f"{a['对象']} · {a['指标']}（{unit}）近 60 天", legend=dict(orientation="h", y=1.15))
            st.plotly_chart(fig, key="rad_series")

    with tabs[3]:
        c1_, c2_, c3_, c4_ = st.columns(4)
        rng = c1_.selectbox("时间", ["选中这一天", "最近 7 天", "最近 30 天"], key="rad_rng")
        lv = c2_.multiselect("层级", list(_LEVEL_RANK), default=[x for x in _LEVEL_RANK if x in set(R["层级"])],
                             key="rad_lv")
        ms = c3_.multiselect("指标", list(_METRIC_RANK), default=[x for x in _METRIC_RANK if x in set(R["指标"])],
                             key="rad_ms")
        dr = c4_.multiselect("方向", ["下跌", "上升"], default=["下跌", "上升"], key="rad_dr")
        n_days = {"选中这一天": 1, "最近 7 天": 7, "最近 30 天": 30}[rng]
        T = R[(R["日期"] <= day) & (R["日期"] > day - pd.Timedelta(days=n_days)) & R["层级"].isin(lv)
              & R["指标"].isin(ms) & R["方向"].isin(dr)]
        st.caption(f"共 {len(T)} 条。排序：日期新的在前，同一天先曝光。")
        if len(T):
            disp = pd.DataFrame({"日期": T["日期"].dt.strftime("%m-%d"), "方向": np.where(T["方向"] == "下跌", "🔴 下跌", "🟢 上升"),
                                 "层级": T["层级"], "对象": T["对象"], "指标": T["指标"],
                                 "说明": [alert_headline(r, win) for _, r in T.iterrows()]})
            st.dataframe(disp, hide_index=True, height=min(700, 40 + 35 * len(disp)))
            st.download_button("⬇️ 下载预警清单 CSV", disp.to_csv(index=False).encode("utf-8-sig"),
                               f"异常预警_{PROF['name']}_{day:%m%d}.csv", "text/csv", key="rad_dl")

    with tabs[4]:
        _anomaly_periodic()


def _alert_series(a: pd.Series) -> pd.Series:
    """一条预警对应对象 × 指标的逐日序列。"""
    m, lv, obj = a["指标"], a["层级"], a["obj"]
    if lv == "达人":
        x = CRB["creator_day"][CRB["creator_day"]["达人"] == obj].groupby("日期").sum(numeric_only=True)
        if m == "CTR":
            return (x["点击"] / x["曝光"].where(x["曝光"] > 0) * 100)
        return x[m] if m in x else pd.Series(dtype=float)
    d = df_all if lv != "链接" else df_all[df_all["product_id"] == obj]
    if lv == "渠道":
        cols = {"曝光": CH_IMPR.get(obj), "点击": CH_CLICK.get(obj), "GMV": dict(CHANNELS).get(obj)}
        g = d.groupby("日期")[[K[v] for v in cols.values() if v and K.get(v) in d.columns]].sum()
        imp, clk = K.get(cols["曝光"]), K.get(cols["点击"]) if cols["点击"] else None
        if m == "CTR" and clk in g and imp in g:
            return g[clk] / g[imp].where(g[imp] > 0) * 100
        c = K.get(cols.get(m))
        return g[c] if c in g else pd.Series(dtype=float)
    key = _METRIC_KEY[m]
    s = series_of(d, key)
    return s.set_index("日期")["值"]


def _period_text(days) -> str:
    days = _ts_list(days)
    return f"{days[0]:%m-%d}" if len(days) == 1 else f"{days[0]:%m-%d}～{days[-1]:%m-%d}"


def plain_summary(cur: dict, base: dict, cur_lbl: str, base_lbl: str, lead: str = "impr") -> list[str]:
    """大白话总结：先说索引指标（默认曝光），再把漏斗每一环各说一句。"""
    order = [("曝光", "impr"), ("点击", "clicks"), ("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"),
             ("订单", "sku"), ("客单价", "AOV 客单价"), ("GMV", "gmv")]
    order.sort(key=lambda x: 0 if x[1] == lead else 1)
    return [say_change(nm, kk, value_of(cur, kk), value_of(base, kk), cur_lbl, base_lbl) for nm, kk in order
            if pd.notna(value_of(cur, kk)) or pd.notna(value_of(base, kk))]


def _anomaly_periodic():
    st.caption("这里看大盘整体：按 日 / 周 / 月 检测（周 / 月的规模指标按日均比）。前面几个标签是逐条链接 / 渠道"
               + ("/ 达人" if PROF["key"] == "cr" else "") + "的预警。")
    c0, c1, c2, c3, c4 = st.columns([1.1, 1.3, 1, 1, 1])
    gran5 = c0.radio("维度", ["日", "周", "月"], horizontal=True, index=0, key="g5")
    target = c1.selectbox("检测指标", list(PROF["anomaly_targets"]), key=f"a_target_{PROF['key']}")
    win = c2.slider("基线 = 前 N 期", 3, 21, 7 if gran5 == "日" else 4, step=1, key=f"a_win_{gran5}")
    kk = c3.slider("灵敏度 k", 1.0, 5.0, 3.0 if gran5 == "日" else 2.0, step=0.5, key=f"a_k_{gran5}")
    pthr = c4.slider("或 |偏离| ≥ %", 10, 80, 30 if gran5 == "日" else 20, step=5, key=f"a_p_{gran5}")
    by5 = PROF["anomaly_targets"][target]
    tc = by5 if by5 in DERIVED else K[by5]
    D = df
    agg5 = aggregate(D, gran5)
    if by5 not in DERIVED and gran5 != "日":
        agg5[tc] = agg5[tc] / agg5["天数"]
    ser = agg5[["期间", tc]].dropna().rename(columns={"期间": "日期"})
    det = detect_anomalies(ser, tc, win=win, k=kk, pct_thr=pthr)
    per_all = to_period(D["日期"], gran5)
    periods_sorted = sorted(agg5["期间"].tolist())
    unit5 = {"日": "天", "周": "周", "月": "月"}[gran5]
    ulab = unit_of(by5) if by5 not in DERIVED else "%"
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=det["日期"], y=det[tc], mode="lines", line=dict(color="#C2416B", width=2),
                             name=target + ("（日均）" if gran5 != "日" and by5 not in DERIVED else "")))
    fig.add_trace(go.Scatter(x=det["日期"], y=det["基线"], mode="lines", name=f"基线（前{win}{unit5}中位数）",
                             line=dict(color="#8A727C", width=1.5, dash="dash")))
    an = det[det["异常"]]
    fig.add_trace(go.Scatter(x=an["日期"], y=an[tc], mode="markers", name="异常",
                             marker=dict(color="#D9534F", size=11, symbol="circle-open", line=dict(width=3))))
    fig.update_layout(height=380, hovermode="x unified", legend=dict(orientation="h", y=1.12),
                      title=f"{SCOPE} · {target}（{ulab}）", margin=dict(t=50, b=10))
    st.plotly_chart(fig, key="a_series")
    if an.empty:
        st.info(f"没有检出异常{unit5}（|z| ≥ {kk} 或 |偏离| ≥ {pthr}%）。")
        return
    st.markdown(f"**检出 {len(an)} 个异常{unit5}**")
    lst = []
    for _, r in an.sort_values("日期", ascending=False).iterrows():
        cur_v, base_v = r[tc], r["基线"]
        lbl = _plabel_any(pd.Timestamp(r["日期"]), gran5)
        if by5 in DERIVED and DERIVED[by5][2] == "pct":
            txt = f"{target} {cur_v:.2f}%，前 {win} {unit5}中位数 {base_v:.2f}%，{'高' if cur_v > base_v else '低'}了 {abs(cur_v - base_v):.2f} 个百分点"
        else:
            txt = (f"{target}{'（日均）' if gran5 != '日' and by5 not in DERIVED else ''} {fv(by5, cur_v)}，前 {win} {unit5}中位数 "
                   f"{fv(by5, base_v)}，{'多' if cur_v > base_v else '少'}了 {fv(by5, abs(cur_v - base_v))}（{r['偏离%']:+.0f}%）")
        lst.append({"期间": lbl, "方向": "🟢 突增" if r["偏离%"] >= 0 else "🔴 突降", "说明": txt})
    st.dataframe(pd.DataFrame(lst), hide_index=True)

    pickd = st.selectbox(f"选一个异常{unit5}看深度归因", an["日期"].tolist()[::-1],
                         format_func=lambda d: _plabel_any(pd.Timestamp(d), gran5), key=f"a_pick_{gran5}")
    pd_ = pd.Timestamp(pickd)
    prevs = [x for x in periods_sorted if x < pd_][-win:]
    cur_mask, base_mask = per_all == pd_, per_all.isin(prevs)
    n_c, n_b = D.loc[cur_mask, "日期"].nunique(), D.loc[base_mask, "日期"].nunique()
    if not prevs or not n_b:
        st.warning(f"这一{unit5}之前没有可用的基线。")
        return
    scale5 = n_b / n_c
    cur = sums_of(D[cur_mask])
    base = {k2: v / scale5 for k2, v in sums_of(D[base_mask]).items()}
    cur_d, base_d = D.loc[cur_mask, "日期"].unique(), D.loc[base_mask, "日期"].unique()
    cur_lbl = f"{_period_text(cur_d)}（{n_c} 天）" if gran5 != "日" else _period_text(cur_d)
    base_lbl = (f"前 {len(prevs)} {unit5}（{_period_text(base_d)}）的平均"
                + ("，已按天数折算" if gran5 != "日" else ""))
    lines = plain_summary(cur, base, cur_lbl, base_lbl, lead=by5 if by5 in ("impr", "clicks", "sku", "gmv") else "impr")
    chd = sorted(((l.split()[0], cur.get(k2, 0) - base.get(k2, 0)) for l, k2 in CH_IMPR.items() if k2 in cur),
                 key=lambda x: x[1])
    lines.append("曝光按渠道：" + "、".join(f"{nm} {fdelta('impr', v)}" for nm, v in chd) + "。")
    gcx, gbx, nmx = link_sums(D, cur_mask, base_mask, scale5)
    if by5 in DERIVED:
        la = ratio_contrib(gcx, gbx, nmx, by5)
        dcol = "贡献"
    else:
        la = sum_contrib(gcx, gbx, nmx, by5)
        dcol = "Δ"
    if len(la):
        top_l = la.reindex(la[dcol].abs().sort_values(ascending=False).index).iloc[0]
        v_txt = f"{top_l[dcol]:+.2f} 个百分点" if by5 in DERIVED else fdelta(by5, top_l[dcol])
        lines.append(f"链接上：{target}变化最大的是 {top_l['链接']}（{v_txt}）。")
    d_main = value_of(cur, by5) - value_of(base, by5)
    insight_cards([{"level": level_of(pct(value_of(cur, by5), value_of(base, by5)), thr=pthr / 3),
                    "title": f"{cur_lbl} 对比 {base_lbl}", "tag": f"{target}{'上升' if d_main > 0 else '下降'}",
                    "body": lines}], cols=1)

    section_title("① GMV 拆到漏斗四环", "GMV = 曝光 × CTR × CTOR × 客单价；四项加起来 = GMV 的总变化")
    fa = funnel_attribution(cur, base)
    d_gmv = cur["gmv"] - base["gmv"]
    fw = go.Figure(go.Waterfall(x=fa["因子"].tolist() + ["GMV 总变化"], y=fa["贡献金额"].tolist() + [d_gmv],
                                measure=["relative"] * len(fa) + ["total"],
                                text=[fdelta("gmv", v) for v in fa["贡献金额"].tolist() + [d_gmv]], textposition="outside",
                                decreasing=dict(marker=dict(color="#D9534F")), increasing=dict(marker=dict(color="#2E9E7A")),
                                totals=dict(marker=dict(color="#4C6EF5"))))
    fw.update_layout(height=320, margin=dict(t=20, b=10), yaxis_title="$")
    st.plotly_chart(fw, key="a_fw")
    section_title("② 按渠道拆", f"本{unit5} − {base_lbl}")
    ca = pd.DataFrame([{"渠道": l, "曝光变化": cur.get(CH_IMPR[l], 0) - base.get(CH_IMPR[l], 0),
                        "GMV变化": cur.get(k2, 0) - base.get(k2, 0)} for l, k2 in CHANNELS if CH_IMPR.get(l) in cur])
    cc1, cc2 = st.columns(2)
    for cc, colv, ukey in ((cc1, "曝光变化", "impr"), (cc2, "GMV变化", "gmv")):
        f = go.Figure(go.Bar(x=ca[colv], y=[l.split()[0] for l in ca["渠道"]], orientation="h",
                             marker_color=[CH_COLOR[l] for l in ca["渠道"]],
                             text=[fdelta(ukey, v) for v in ca[colv]], textposition="auto"))
        f.add_vline(x=0, line_color="#8A727C")
        f.update_layout(height=240, margin=dict(t=30, b=10), title=f"各渠道{colv}（{unit_of(ukey)}）")
        cc.plotly_chart(f, key=f"a_ch_{ukey}")
    section_title("③ 按链接拆：全部指标", f"按「{target}」的变化排序；基期 = {base_lbl}")
    if len(la):
        st.plotly_chart(contrib_bars(la, dcol, xtitle=(f"对{SCOPE} {target} 的贡献（个百分点）" if by5 in DERIVED
                                                       else f"{target}变化（{unit_of(by5)}）")), key="a_la")
    st.dataframe(link_multi_attr(D, cur_mask, base_mask, scale5,
                                 by5 if by5 in ("impr", "clicks", "sku", "gmv") else "impr").drop(columns="_sort"),
                 hide_index=True)
    if PROF["key"] == "cr":
        section_title("④ 按达人拆：全部指标", "Creators 表；曝光 / 点击 / CTR / 播放 / 订单 / GMV / 新发视频")
        st.dataframe(creator_multi_attr(cur_d, base_d, _CR_DIM_COL.get(by5, "曝光")), hide_index=True)
    elif CRB is not None and len(la) and "主要渠道" in la.columns:
        aff = la[la["主要渠道"] == "达人"]
        aff = aff.reindex(aff[dcol].abs().sort_values(ascending=False).index).head(5)
        cov4 = set(pd.to_datetime(CRB["link"]["日期"].unique()))
        if len(aff) and not any(pd.Timestamp(d) in cov4 for d in cur_d):
            st.caption(f"④ 达人侧原因：这一{unit5}不在达人组数据范围内（达人数据只到 {max(cov4):%m-%d}）。")
        elif len(aff):
            section_title("④ 变化主要在达人渠道的链接：达人侧原因", "达人组数据，日均")
            st.dataframe(pd.DataFrame({"链接": aff["链接"].values,
                                       f"{target}变化": [fdelta(by5, v) for v in aff[dcol]],
                                       "达人侧原因": [cr_link_reason(p, cur_d, base_d) for p in aff.index]}),
                         hide_index=True)


def _plabel_any(p: pd.Timestamp, gran: str) -> str:
    if gran == "日":
        return f"{p:%m-%d}（周{'一二三四五六日'[p.weekday()]}）"
    if gran == "周":
        return f"{p:%m-%d}～{p + pd.Timedelta(days=6):%m-%d}"
    return f"{p:%Y-%m}"


# ── ⑦ 链接下钻 ────────────────────────────────────────────────────────
def page_drilldown():
    head = st.container()
    tot_by = df.groupby(["product_id"]).agg(i=(K["impr"], "sum"), g=(K["gmv"], "sum"),
                                            n=("product_name", "last")).reset_index()
    tot_by = tot_by.sort_values("i", ascending=False)
    namemap = dict(zip(tot_by["product_id"], tot_by["n"]))
    imap, gmap = dict(zip(tot_by["product_id"], tot_by["i"])), dict(zip(tot_by["product_id"], tot_by["g"]))
    c1, c2 = st.columns([3, 1])
    pid = c1.selectbox("选择链接（按曝光从高到低）", tot_by["product_id"].tolist(), key=f"dd_pid_{PROF['key']}",
                       format_func=lambda p: f"{short_name(namemap.get(p, ''))}　·　{fv('impr', imap.get(p, 0))}　·　"
                                             f"{fv('gmv', gmap.get(p, 0))}")
    gran3 = c2.radio("粒度", PROF["periods"], horizontal=True, index=0, key=f"g_link_{PROF['key']}")
    lname = short_name(namemap.get(pid, ""))
    sub = df[df["product_id"] == pid].copy()
    unit3 = {"日": "天", "周": "周", "月": "月", "季": "季"}[gran3]
    sub["_p"] = to_period(sub["日期"], gran3)
    keys3 = list(dict.fromkeys(k for k, _ in PROF["drill"] if k in K and K[k] in sub.columns))
    gs = sub.groupby("_p")[[K[k] for k in keys3]].sum()
    gs.columns = keys3
    nd3 = df.assign(_p=to_period(df["日期"], gran3)).groupby("_p")["日期"].nunique().reindex(gs.index)
    gs["CTR%"] = np.where(gs["impr"] > 0, gs["clicks"] / gs["impr"] * 100, np.nan)
    if "atc" in gs:
        gs["加购率%"] = np.where(gs["clicks"] > 0, gs["atc"] / gs["clicks"] * 100, np.nan)
    gs["CTOR%"] = np.where(gs["clicks"] > 0, gs["sku"] / gs["clicks"] * 100, np.nan)
    gs["件单价"] = np.where(gs["items"] > 0, gs["gmv"] / gs["items"], np.nan)
    if gran3 != "日":
        gs[keys3] = gs[keys3].div(nd3, axis=0)
    gs = gs.reset_index().rename(columns={"_p": "期间"})
    gs.insert(1, "天数", nd3.values)
    avg_t = "（日均）" if gran3 != "日" else ""
    ev_l = EV_ALL[(EV_ALL["product_id"] == pid) & (EV_ALL["事件"].isin(["改名/重组", "改价/改促销（件单价突变）"]))]

    # 最近 7 天 vs 前 7 天（全量数据，不受分析区间影响）
    D = df_all
    d_last = D["日期"].max()
    a0, b0, b1 = d_last - pd.Timedelta(days=6), d_last - pd.Timedelta(days=13), d_last - pd.Timedelta(days=7)
    lk = D[D["product_id"] == pid]
    cm, bm = lk["日期"].between(a0, d_last), lk["日期"].between(b0, b1)
    na, nb = D.loc[D["日期"].between(a0, d_last), "日期"].nunique(), D.loc[D["日期"].between(b0, b1), "日期"].nunique()
    sc = sums_of(lk[cm])
    sb = {k: v * na / nb for k, v in sums_of(lk[bm]).items()} if nb else {}
    R = radar_all(7, 3.0, 30.0)
    alerts = R[(R["层级"] == "链接") & (R["obj"] == pid) & (R["日期"] > d_last - pd.Timedelta(days=14))]
    p_im, p_g = pct(sc.get("impr"), sb.get("impr")), pct(sc.get("gmv"), sb.get("gmv"))
    with head:
        page_header(lname or "链接下钻", "单条链接：最近 7 天怎么样、掉在哪个渠道、全部指标的走势" +
                    ("、是哪些达人在推" if PROF["key"] == "cr" else ""),
                    [f"{PROF['name']}数据", f"Product ID …{str(pid)[-6:]}", f"粒度：{gran3}"])
        summary_pills([("info", f"最近 7 天（{a0:%m/%d}–{d_last:%m/%d}）vs 前 7 天"),
                       (level_of(p_im), f"曝光 {p_im:+.1f}%" if pd.notna(p_im) else "曝光 —"),
                       (level_of(p_g), f"GMV {p_g:+.1f}%" if pd.notna(p_g) else "GMV —"),
                       ("down" if (alerts["方向"] == "下跌").any() else "flat", f"近 14 天预警 {len(alerts)} 条")])
    tab_names = ["结论", "图表", "明细表"] + (["达人与内容"] if PROF["key"] == "cr" else [])
    tabs = st.tabs(tab_names)
    with tabs[0]:
        if sb:
            cl, bl = f"最近 7 天（{a0:%m-%d}～{d_last:%m-%d}）", f"前 7 天（{b0:%m-%d}～{b1:%m-%d}）"
            chd = sorted(((l.split()[0], sc.get(k, 0) - sb.get(k, 0)) for l, k in CH_IMPR.items() if k in sc),
                         key=lambda x: x[1])
            cards = [{"level": level_of(p_im), "title": "曝光", "tag": "索引指标",
                      "body": [say_change("曝光", "impr", sc.get("impr"), sb.get("impr"), cl, bl),
                               "按渠道：" + "、".join(f"{n} {fdelta('impr', v)}" for n, v in chd) + "。"]},
                     {"level": level_of(value_of(sc, "CTR 点击率") - value_of(sb, "CTR 点击率"), thr=0.1), "title": "转化",
                      "body": [say_change(nm, kk, value_of(sc, kk), value_of(sb, kk), cl, bl)
                               for nm, kk in (("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"), ("客单价", "AOV 客单价"))]},
                     {"level": level_of(p_g), "title": "GMV", "tag": "结果",
                      "body": say_change("GMV", "gmv", sc.get("gmv"), sb.get("gmv"), cl, bl)}]
            if PROF["key"] == "ops" and CRB is not None:
                cards.append({"level": "info", "title": "达人侧", "tag": "达人组数据",
                              "body": cr_link_reason(pid, D.loc[D["日期"].between(a0, d_last), "日期"].unique(),
                                                     D.loc[D["日期"].between(b0, b1), "日期"].unique())})
            if len(alerts):
                cards.append({"level": "down" if (alerts["方向"] == "下跌").any() else "up",
                              "title": f"近 14 天 {len(alerts)} 条预警", "tag": "异常预警",
                              "body": [f"{a['指标']}：{alert_headline(a, 7)}" for _, a in alerts.head(5).iterrows()]})
            if len(ev_l):
                cards.append({"level": "warn", "title": f"识别到 {len(ev_l)} 次改动", "tag": "改动效果溯源",
                              "body": [f"{e['日期']:%m-%d} {e['事件']}：{e['详情']}" for _, e in ev_l.tail(4).iterrows()]})
            insight_cards(cards, cols=2)
        else:
            st.info("前 7 天没有数据，没法比。")

    def mark(fig):
        for _, e in ev_l.iterrows():
            fig.add_vline(x=e["日期"], line_color="#9A6512", line_dash="dot", opacity=.7)
        return fig

    with tabs[1]:
        c1, c2 = st.columns(2)
        f = go.Figure()
        for lbl, k_ in CH_IMPR.items():
            if k_ in gs:
                f.add_bar(x=gs["期间"], y=gs[k_], name=lbl.split()[0], marker_color=CH_COLOR[lbl])
        f.update_layout(barmode="stack", height=300, title=f"曝光{avg_t}（按渠道，IM）", margin=dict(t=40, b=10),
                        legend=dict(orientation="h", y=-0.2))
        c1.plotly_chart(mark(f), key="dd_imp")
        f = go.Figure(go.Bar(x=gs["期间"], y=gs["gmv"], marker_color="#C2416B"))
        f.update_layout(height=300, title=f"GMV{avg_t}（$）", margin=dict(t=40, b=10))
        c2.plotly_chart(mark(f), key="dd_gmv")
        c3, c4 = st.columns(2)
        rate_cols = ["CTR%", "CTOR%"] + (["加购率%"] if "atc" in keys3 else [])
        f3 = px.line(gs, x="期间", y=rate_cols, markers=True, labels={"value": "%", "期间": ""})
        f3.update_layout(height=300, title=" / ".join(c.rstrip("%") for c in rate_cols), legend_title="",
                         margin=dict(t=40, b=10))
        c3.plotly_chart(mark(f3), key="dd_rate")
        f4 = px.line(gs, x="期间", y="件单价", markers=True, labels={"期间": ""})
        f4.update_traces(line_color="#1B3A8C")
        f4.update_layout(height=300, title="件单价（GMV ÷ 件数，$）", margin=dict(t=40, b=10))
        c4.plotly_chart(mark(f4), key="dd_asp")
        if len(ev_l):
            st.caption("虚线 = 自动识别到的改动：" + "；".join(f"{e['日期']:%m-%d} {e['事件']}（{e['详情']}）"
                                                        for _, e in ev_l.iterrows()))
    with tabs[2]:
        tshow = gs.copy()
        for k, lbl in PROF["drill"]:
            if k in tshow:
                tshow[k] = tshow[k].map(lambda v, kk=k: fv(kk, v, 1 if gran3 != "日" else 0))
        tshow = tshow.rename(columns={k: f"{lbl}{avg_t}" for k, lbl in PROF["drill"]})
        tshow[f"曝光 环比%（vs 上一{unit3}）"] = gs["impr"].pct_change() * 100
        tshow[f"GMV 环比%（vs 上一{unit3}）"] = gs["gmv"].pct_change() * 100
        st.dataframe(round_num(tshow.iloc[::-1]), hide_index=True)
    if PROF["key"] == "cr":
        with tabs[3]:
            render_creator_link_extras(pid, sub, gran3)

# ── ⑧ 指标归档 ────────────────────────────────────────────────────────
def page_catalog():
    page_header("指标归档", "这个程序读到的全部指标：属于前端（流量）还是后端（成交）、来自哪个区段 / 哪张表。",
                [f"{PROF['name']}数据", f"{len(catalog)} 个指标"])
    if catalog.empty:
        st.warning("未生成指标目录。")
        return
    n_f, n_b = int((catalog["分类"] == "前端").sum()), int((catalog["分类"] == "后端").sum())
    summary_pills([("info", f"指标总数 {len(catalog)}"), ("info", f"前端 · 流量 {n_f}"), ("info", f"后端 · 成交 {n_b}"),
                   ("flat", f"其他 {len(catalog) - n_f - n_b}")])
    tabs = st.tabs(["全部指标", "分布"])
    with tabs[0]:
        fsel = st.multiselect("筛选分类", ["前端", "后端", "其他"], default=["前端", "后端", "其他"], key="cat_sel")
        show_cols = ["列号", "区段", "指标", "分类", "类型"] + [c for c in ("细分类", "层级") if c in catalog.columns]
        if "层级" in catalog.columns:
            st.caption("层级：「链接」= 每条链接每天都有，可以按链接归因；「全量」= 原始表的每日合计"
                       "（达人表 / 视频表 / 直播表的字段本身不带链接维度），在「前端 / 后端」页按天看趋势。")
        st.dataframe(catalog[catalog["分类"].isin(fsel)][show_cols], hide_index=True, height=520)
        st.download_button("⬇️ 导出指标归档 CSV", catalog.to_csv(index=False).encode("utf-8-sig"),
                           f"nailvesta_{PROF['name']}_指标归档.csv", "text/csv")
    with tabs[1]:
        c1, c2 = st.columns([1, 2])
        pie = catalog.groupby(["分类"]).size().reset_index(name="数量")
        fp = px.pie(pie, names="分类", values="数量", hole=.55,
                    color="分类", color_discrete_map={"前端": "#4C6EF5", "后端": "#C2416B", "其他": "#8A727C"})
        fp.update_layout(height=300, margin=dict(t=10))
        c1.plotly_chart(fp, key="cat_pie")
        heat = catalog.groupby(["区段", "分类"]).size().reset_index(name="数量")
        fh = px.bar(heat, x="区段", y="数量", color="分类", barmode="stack",
                    color_discrete_map={"前端": "#4C6EF5", "后端": "#C2416B", "其他": "#8A727C"})
        fh.update_layout(height=300, margin=dict(t=10), legend_title="")
        c2.plotly_chart(fh, key="cat_bar")


def link_selector(key: str, allow_all=True, label: str = "链接（按曝光从高到低）"):
    tot_by = df.groupby("product_id").agg(i=(K["impr"], "sum"), g=(K["gmv"], "sum"),
                                          n=("product_name", "last")).reset_index()
    tot_by = tot_by.sort_values("i", ascending=False)
    ids = (["__ALL__"] if allow_all else []) + tot_by["product_id"].tolist()
    nm = dict(zip(tot_by["product_id"], tot_by["n"]))
    im, gm = dict(zip(tot_by["product_id"], tot_by["i"])), dict(zip(tot_by["product_id"], tot_by["g"]))

    def lab(p):
        if p == "__ALL__":
            return f"【{SCOPE}合计】"
        return f"{short_name(nm.get(p, ''))}　·　{fv('impr', im.get(p, 0))}　·　{fv('gmv', gm.get(p, 0))}"
    pick = st.selectbox(label, ids, format_func=lab, key=key)
    return (None if pick == "__ALL__" else pick), (None if pick == "__ALL__" else nm.get(pick, ""))


def _fmt_culprits(cp: pd.DataFrame, mk: str) -> pd.DataFrame:
    """culprits() 的结果 → 带单位的展示表。"""
    cpd = cp.copy()
    if mk in DERIVED:
        for c_ in ("Δ", "比率效应", "结构效应"):
            if c_ in cpd:
                cpd[c_] = cpd[c_].map(lambda v: "—" if pd.isna(v) else f"{v:+.2f} 个百分点")
        return cpd.drop(columns=["product_id"], errors="ignore").rename(columns={"Δ": "贡献"})
    cpd["Δ"] = cpd["Δ"].map(lambda v: fdelta(mk, v))
    for c_ in ("本期", "基期"):
        if c_ in cpd:
            cpd[c_] = cpd[c_].map(lambda v: fv(mk, v))
    if "自身变化" in cpd:
        cpd["自身变化"] = cpd["自身变化"].map(lambda v: "—" if pd.isna(v) else f"{v:+.0f}%")
    if "占比" in cpd:
        cpd["占比"] = cpd["占比"].map(lambda v: "—" if pd.isna(v) else f"{v:.0f}%")
    return cpd.drop(columns=["product_id"], errors="ignore")


# ── ⑨ 日环比链条 & 回暖 ───────────────────────────────────────────────
def page_chain():
    head = st.container()
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        pid8, pname8 = link_selector(f"lnk8_{PROF['key']}")
    mname = c2.selectbox("指标", list(METRIC_PICK.keys()), index=0, key=f"chain_m_{PROF['key']}")
    thr = c3.slider("大跌线：比前一天跌 ≥ %", 5, 50, 15, step=5, key="chain_thr")
    mk8 = METRIC_PICK[mname]
    is_rate8 = mk8 in DERIVED and DERIVED[mk8][2] == "pct"
    title_obj = short_name(pname8) if pname8 else (SCOPE if SCOPE.endswith("合计") else f"{SCOPE}合计")
    s8 = series_of(df, mk8, pid8)
    if s8.empty or s8["值"].dropna().empty:
        with head:
            page_header("日环比 & 回暖", "每一天比前一天怎么动、大跌之后有没有回来", [title_obj, mname])
        st.warning("这条链接在当前区间里没有这个指标的数据。")
        return
    ch = daily_chain(s8)
    d_first, d_last = ch["日期"].iloc[0], ch["日期"].iloc[-1]
    drops = ch[ch["日环比%"] <= -thr]
    recs = {}
    for _, dr in drops.iterrows():
        r = recovery_after(s8, dr["日期"], pre=7, post=14)
        if r:
            recs[dr["日期"]] = r
    last_dc = ch["日环比%"].iloc[-1]
    cum = ch["相对起点%"].iloc[-1]

    def dtxt(cur, prev):
        """比率型给百分点，其余给 %。"""
        if pd.isna(cur) or pd.isna(prev):
            return "—"
        return f"{cur - prev:+.2f} 个百分点" if is_rate8 else (f"{pct(cur, prev):+.1f}%" if prev else "—")

    with head:
        page_header("日环比 & 回暖", "每一天比前一天怎么动；大跌之后有没有回到跌前水平。每个 % 都写明是跟谁比的。",
                    [f"{PROF['name']}数据", title_obj, f"指标：{mname}", f"{d_first:%m/%d} → {d_last:%m/%d}"])
        last_rec = recs[max(recs)] if recs else None
        summary_pills([("info", f"{title_obj} · {mname}"),
                       (level_of(last_dc), f"{d_last:%m-%d} vs 前一天 {dtxt(ch['值'].iloc[-1], ch['前一日'].iloc[-1])}"),
                       (level_of(cum), f"vs 区间首日 {d_first:%m-%d}：{dtxt(ch['值'].iloc[-1], ch['值'].iloc[0])}"),
                       ("down" if len(drops) else "flat", f"单日跌 ≥ {thr}%：{len(drops)} 次"),
                       (("up" if last_rec["verdict"].startswith("✅") else "down" if last_rec["verdict"].startswith("❌")
                         else "warn"), f"最近一次大跌 {max(recs):%m-%d}：{last_rec['verdict'][2:]}") if last_rec else
                       ("flat", "")])
    tabs = st.tabs(["结论", "逐日链条", "大跌与回暖", "明细表"])

    # ── 结论
    with tabs[0]:
        sub = df_all if pid8 is None else df_all[df_all["product_id"] == pid8]
        D = df_all
        dl = D["日期"].max()
        a0, b0, b1 = dl - pd.Timedelta(days=6), dl - pd.Timedelta(days=13), dl - pd.Timedelta(days=7)
        na = D.loc[D["日期"].between(a0, dl), "日期"].nunique()
        nb = D.loc[D["日期"].between(b0, b1), "日期"].nunique()
        cards = []
        vl, vp = ch["值"].iloc[-1], ch["前一日"].iloc[-1]
        cards.append({"level": level_of(vl - vp if is_rate8 else pct(vl, vp), thr=0.1 if is_rate8 else 5),
                      "title": f"最近一天（{d_last:%m-%d}）", "tag": "vs 前一天",
                      "body": say_change(mname, mk8, vl, vp, f"{d_last:%m-%d}", f"前一天（{ch['日期'].iloc[-2]:%m-%d}）")
                      if len(ch) > 1 else "只有一天数据。"})
        if na and nb:
            sc = {k: v / na for k, v in sums_of(sub[sub["日期"].between(a0, dl)]).items()}
            sb = {k: v / nb for k, v in sums_of(sub[sub["日期"].between(b0, b1)]).items()}
            c7, b7 = value_of(sc, mk8), value_of(sb, mk8)
            cards.append({"level": level_of(c7 - b7 if is_rate8 else pct(c7, b7), thr=0.1 if is_rate8 else 5),
                          "title": "最近 7 天", "tag": "日均 vs 前 7 天日均",
                          "body": say_change(mname + ("" if mk8 in DERIVED else "（日均）"), mk8, c7, b7,
                                             f"最近 7 天（{a0:%m-%d}～{dl:%m-%d}）", f"前 7 天（{b0:%m-%d}～{b1:%m-%d}）")})
        if recs:
            dd = max(recs)
            r = recs[dd]
            drow = ch[ch["日期"] == dd].iloc[0]
            lv = "up" if r["verdict"].startswith("✅") else "down" if r["verdict"].startswith("❌") else "warn"
            body = [f"{dd:%m-%d} 当天 {fv(mk8, drow['值'])}，比前一天（{fv(mk8, drow['前一日'])}）跌了 {abs(drow['日环比%']):.0f}%。",
                    f"跌前 7 天中位数是 {fv(mk8, r['baseline'])}；之后最低到 {fv(mk8, r['trough'])}"
                    f"（{r['trough_day']:%m-%d}，比跌前低 {abs((r['trough'] / r['baseline'] - 1) * 100) if r['baseline'] else 0:.0f}%）。",
                    f"到 {r['after']['日期'].iloc[-1]:%m-%d} 是 {fv(mk8, r['after']['值'].iloc[-1])}，"
                    f"跟跌前比 {r['end_recovery']:+.0f}%。" if pd.notna(r["end_recovery"]) else ""]
            if r["after_spike"]:
                body.append("注意：大跌前一天本身是冲高，这次「跌」多半只是冲高后的回落。")
            cards.append({"level": lv, "title": f"最近一次大跌：{r['verdict'][2:]}", "tag": f"{dd:%m-%d}", "body": body})
            ok = sum(1 for x in recs.values() if x["verdict"].startswith("✅"))
            bad = sum(1 for x in recs.values() if x["verdict"].startswith("❌"))
            cards.append({"level": "warn" if bad else "info", "title": f"区间里一共 {len(recs)} 次大跌", "tag": f"单日跌 ≥ {thr}%",
                          "body": [f"已恢复 {ok} 次、继续阴跌 {bad} 次、其余在回暖或横盘。",
                                   "逐次看：「大跌与回暖」标签。"]})
        else:
            cards.append({"level": "flat", "title": f"区间里没有单日跌 ≥ {thr}% 的情况", "tag": "大跌",
                          "body": "可以把上面的「大跌线」调低一点再看。"})
        worst = ch.dropna(subset=["日环比%"]).nsmallest(1, "日环比%")
        best = ch.dropna(subset=["日环比%"]).nlargest(1, "日环比%")
        if len(worst) and len(best):
            w, b_ = worst.iloc[0], best.iloc[0]
            cards.append({"level": "info", "title": "区间里单日变化最大的两天", "tag": "vs 前一天",
                          "body": [f"跌得最多：{w['日期']:%m-%d}，{fv(mk8, w['前一日'])} → {fv(mk8, w['值'])}（{w['日环比%']:+.0f}%）。",
                                   f"涨得最多：{b_['日期']:%m-%d}，{fv(mk8, b_['前一日'])} → {fv(mk8, b_['值'])}（{b_['日环比%']:+.0f}%）。"]})
        insight_cards(cards, cols=2)

    # ── 逐日链条
    with tabs[1]:
        unit = unit_of(mk8) if mk8 not in DERIVED else ("%" if is_rate8 else "$")
        f = go.Figure()
        f.add_trace(go.Scatter(x=ch["日期"], y=ch["值"], mode="lines+markers", name=mname,
                               line=dict(color="#C2416B", width=2), marker=dict(size=5)))
        f.update_layout(height=300, hovermode="x unified", margin=dict(t=40, b=10),
                        title=f"{title_obj} · {mname}（{unit}）" if unit else f"{title_obj} · {mname}")
        st.plotly_chart(f, key="ch_abs")
        colors = ["#1F8A6B" if (pd.notna(v) and v >= 0) else "#C24338" for v in ch["日环比%"]]
        f2 = go.Figure(go.Bar(x=ch["日期"], y=ch["日环比%"], marker_color=colors,
                              hovertemplate="%{x|%m-%d}<br>比前一天 %{y:+.1f}%<extra></extra>"))
        f2.add_hline(y=0, line_color="#8A727C")
        f2.add_hline(y=-thr, line_color="#C24338", line_dash="dot", opacity=.6)
        f2.update_layout(height=280, margin=dict(t=40, b=10), yaxis_title="比前一天 %",
                         title="逐日环比（红色虚线 = 大跌线）")
        st.plotly_chart(f2, key="ch_bar")
        f3 = go.Figure(go.Scatter(x=ch["日期"], y=ch["相对起点%"], mode="lines", fill="tozeroy",
                                  line=dict(color="#4C6EF5", width=2)))
        f3.add_hline(y=0, line_color="#8A727C", line_dash="dash")
        f3.update_layout(height=260, margin=dict(t=40, b=10), yaxis_title=f"比 {d_first:%m-%d} 高/低 %",
                         title=f"累计：每天比区间首日（{d_first:%m-%d}）高还是低")
        st.plotly_chart(f3, key="ch_cum")

    # ── 大跌与回暖
    with tabs[2]:
        st.caption("基线 = 跌前 7 天的**中位数**（均值会被跌前某一天的冲高拉高）。"
                   "「跌前一天是冲高」= 大跌前一天比再之前几天的中位数高 30% 以上——这种「跌」多半只是冲高后的回落。")
        if drops.empty:
            st.info(f"区间内没有单日跌幅 ≥ {thr}% 的情况。")
        else:
            post_n = st.slider("看之后多少天", 5, 30, 14, key="chain_post")
            ov = []
            for _, dr in drops.iterrows():
                r = recovery_after(s8, dr["日期"], pre=7, post=post_n)
                if not r:
                    continue
                pf = (lambda v: "—" if pd.isna(v) else f"{v:+.0f}%")
                ov.append({"大跌日": dr["日期"].strftime("%m-%d (%a)"),
                           "当天": fv(mk8, dr["值"]), "前一天": fv(mk8, dr["前一日"]),
                           "比前一天": f"{dr['日环比%']:+.0f}%",
                           "跌前一天是冲高": "是（多半是回落）" if r["after_spike"] else "",
                           "跌前 7 天中位数": fv(mk8, r["baseline"]),
                           "最低点 vs 跌前": pf((r["trough"] / r["baseline"] - 1) * 100 if r["baseline"] else np.nan),
                           "D+7 vs 跌前": pf(r["d7"]), "D+14 vs 跌前": pf(r["d14"]),
                           f"第{post_n}天 vs 跌前": pf(r["end_recovery"]), "判定": r["verdict"]})
            if ov:
                section_title(f"所有大跌日（{len(ov)} 次）")
                st.dataframe(pd.DataFrame(ov).iloc[::-1], hide_index=True)
            dsel = st.selectbox("选一个大跌日，看逐日回暖", drops["日期"].dt.date.tolist()[::-1], key="chain_dsel")
            rec = recovery_after(s8, dsel, pre=7, post=post_n)
            if not rec:
                st.warning("该日前后数据不足。")
            else:
                msg = f"判定：{rec['verdict']}" + ("（注意：跌前一天是冲高，这次「大跌」多半是冲高后的回落）"
                                                  if rec["after_spike"] else "")
                lv = "up" if rec["verdict"].startswith("✅") else "down" if rec["verdict"].startswith("❌") else "warn"
                aft = rec["after"]
                insight_cards([{"level": lv, "title": f"{pd.Timestamp(dsel):%m-%d} 大跌之后", "tag": title_obj,
                                "body": [msg,
                                         f"跌前 {rec['pre_n']} 天中位数 {fv(mk8, rec['baseline'])}；最低 {fv(mk8, rec['trough'])}"
                                         f"（{rec['trough_day']:%m-%d}）；最后一天（D+{len(aft) - 1}，{aft['日期'].iloc[-1]:%m-%d}）"
                                         f" {fv(mk8, aft['值'].iloc[-1])}，跟跌前比 {rec['end_recovery']:+.0f}%。"
                                         if pd.notna(rec["end_recovery"]) else ""]}])
                fr = go.Figure()
                fr.add_trace(go.Bar(x=aft["日期"], y=aft["日环比%"], name="比前一天 %",
                                    marker_color=["#1F8A6B" if (pd.notna(v) and v >= 0) else "#C24338"
                                                  for v in aft["日环比%"]]))
                fr.add_trace(go.Scatter(x=aft["日期"], y=aft["恢复度%"], name="比跌前 7 天中位数 %",
                                        mode="lines+markers", line=dict(color="#4C6EF5", width=2), yaxis="y2"))
                fr.add_hline(y=0, line_color="#8A727C")
                fr.update_layout(height=340, hovermode="x unified", yaxis=dict(title="比前一天 %"),
                                 yaxis2=dict(title="比跌前 %", overlaying="y", side="right"),
                                 legend=dict(orientation="h", y=1.15), margin=dict(t=40))
                st.plotly_chart(fr, key="ch_rec")

                d0 = pd.Timestamp(dsel)
                prev_days = sorted(x for x in df["日期"].unique() if x < d0)[-7:]
                if prev_days:
                    section_title(f"{d0:%m-%d} 是谁掉的", f"跟前 {len(prev_days)} 天日均比")
                    dm, pm = df["日期"] == d0, df["日期"].isin(prev_days)
                    if pid8:
                        dm, pm = dm & (df["product_id"] == pid8), pm & (df["product_id"] == pid8)
                    cs, bs = sums_of(df[dm]), {k: v / len(prev_days) for k, v in sums_of(df[pm]).items()}
                    chx = sorted([(l.split()[0], cs.get(k, 0) - bs.get(k, 0)) for l, k in CH_IMPR.items()],
                                 key=lambda x: x[1])
                    lines = ["曝光按渠道：" + "、".join(f"{n} {fdelta('impr', v)}" for n, v in chx) + "。"]
                    for nm_, kk in (("曝光", "impr"), ("CTR", "CTR 点击率"), ("CTOR", "CTOR 点击成单率"), ("GMV", "gmv")):
                        lines.append(say_change(nm_, kk, value_of(cs, kk), value_of(bs, kk), f"{d0:%m-%d}",
                                                f"前 {len(prev_days)} 天日均"))
                    insight_cards([{"level": "info", "title": "当天全漏斗", "body": lines}])
                    if not pid8:
                        cp = culprits(df, df["日期"] == d0, df["日期"].isin(prev_days), base_scale=len(prev_days),
                                      topn=8, by=mk8, direction=-1)
                        if len(cp):
                            cpd = _fmt_culprits(cp, mk8)
                            if PROF["key"] == "ops" and CRB is not None and "主要渠道" in cp:
                                cpd["达人侧原因"] = [cr_link_reason(r["product_id"], [d0], prev_days)
                                                    if r.get("主要渠道") == "达人" else ""
                                                    for _, r in cp.iterrows()]
                            st.markdown(f"**拖累最多的链接（{mname}）**")
                            st.dataframe(cpd.rename(columns={"本期": "当天", "基期": "前7天日均", "占比": "占跌量"}),
                                         hide_index=True)
                        if PROF["key"] == "cr":
                            col_c = _CR_DIM_COL.get(mk8, "曝光")
                            tcd = _dim_delta(CRB["creator_day"], "达人", col_c, [d0], prev_days)
                            tcd = tcd[tcd["Δ"] < 0].head(8)
                            if len(tcd):
                                st.markdown(f"**拖累最多的达人（{col_c}）**")
                                st.dataframe(_fmt_dim(tcd, col_c).reset_index(), hide_index=True)
                show = aft.copy()
                show["D+N"] = [f"D+{i}" if i else "D0 大跌日" for i in range(len(show))]
                show["日期"] = show["日期"].dt.strftime("%m-%d (%a)")
                show["比前一天"] = show["日环比%"].map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%")
                show["比跌前 7 天中位数"] = show["恢复度%"].map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%")
                show["值"] = show["值"].map(lambda v: fv(mk8, v))
                section_title("大跌后逐日")
                st.dataframe(show[["D+N", "日期", "值", "比前一天", "比跌前 7 天中位数"]], hide_index=True)

    # ── 明细表
    with tabs[3]:
        t = ch.copy()
        big = t["日环比%"].abs() >= thr
        t["日期"] = t["日期"].dt.strftime("%m-%d (%a)")
        t[mname] = t["值"].map(lambda v: fv(mk8, v))
        t["前一天"] = t["前一日"].map(lambda v: fv(mk8, v))
        t["变化量"] = [fdelta(mk8, a - b) if pd.notna(a) and pd.notna(b) else "—" for a, b in zip(ch["值"], ch["前一日"])]
        t["比前一天"] = t["日环比%"].map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%")
        t[f"比区间首日（{d_first:%m-%d}）"] = t["相对起点%"].map(lambda v: "—" if pd.isna(v) else f"{v:+.1f}%")
        st.caption(f"其中有 {int(big.sum())} 天比前一天变化 ≥ {thr}%。")
        st.dataframe(t[["日期", mname, "前一天", "变化量", "比前一天", f"比区间首日（{d_first:%m-%d}）"]].iloc[::-1],
                     hide_index=True, height=460)


# ── ⑩ 改动效果溯源 ────────────────────────────────────────────────────
def _effect_row(e: dict, win: int) -> dict:
    """change_effect 结果 → 一行带单位的展示。"""
    def pf(v):
        return "—" if pd.isna(v) else f"{v:+.0f}%"
    return {f"日均曝光 改前→改后{win}天": f"{fv('impr', e['曝光改前'])} → {fv('impr', e['曝光改后'])}",
            "曝光变化": pf(e["曝光变化%"]), "曝光·改后14天": pf(e["曝光变化%·14天"]),
            "CTR 改前→改后": f"{fv('CTR 点击率', e['CTR改前'])} → {fv('CTR 点击率', e['CTR改后'])}",
            "CTOR 改前→改后": f"{fv('CTOR 点击成单率', e['CTOR改前'])} → {fv('CTOR 点击成单率', e['CTOR改后'])}",
            f"日均GMV 改前→改后{win}天": f"{fv('gmv', e['GMV改前'])} → {fv('gmv', e['GMV改后'])}",
            "GMV变化": pf(e["GMV变化%"]), "GMV·改后14天": pf(e["GMV变化%·14天"]),
            "判定": e["判定"] + ("（改前基数太小，% 仅供参考）" if e["基数太小"] else "")}


def _effect_lines(e: dict, win: int) -> list[str]:
    cl, bl = f"改后 {e['改后天数']} 天日均", f"改前 {e['改前天数']} 天日均"
    out = [say_change("曝光", "impr", e["曝光改后"], e["曝光改前"], cl, bl),
           say_change("CTR", "CTR 点击率", e["CTR改后"], e["CTR改前"], "改后", "改前"),
           say_change("CTOR", "CTOR 点击成单率", e["CTOR改后"], e["CTOR改前"], "改后", "改前"),
           say_change("GMV", "gmv", e["GMV改后"], e["GMV改前"], cl, bl)]
    if pd.notna(e["曝光变化%·14天"]):
        out.append(f"拉长到改后 14 天：曝光 {e['曝光变化%·14天']:+.0f}%、GMV "
                   + ("—" if pd.isna(e["GMV变化%·14天"]) else f"{e['GMV变化%·14天']:+.0f}%") + "（都跟改前日均比）。")
    if e["基数太小"]:
        out.append("改前基数太小（日均 GMV < $20 或曝光 < 500 IM），百分比只作参考。")
    return out


def _effect_level(verdict: str) -> str:
    return "up" if verdict.startswith("✅") else "down" if verdict.startswith("❌") else \
        "warn" if verdict.startswith("🟡") else "flat"


def page_effect():
    head = st.container()
    ev = EV_ALL.copy()
    nm_all = df_all.groupby("product_id")["product_name"].last().to_dict()
    ev["链接"] = ev["product_id"].map(lambda p: short_name(nm_all.get(p, "")))
    ev_main = ev[ev["事件"].isin(["改名/重组", "改价/改促销（件单价突变）"])]

    # 全部自动识别的改动：改后曝光怎么样（会话里缓存）
    memo_key = ("effect", PROF["key"], len(df_all), str(df_all["日期"].max()), len(ev_main))
    memo = st.session_state.setdefault("_effect_memo", {})
    if memo_key not in memo:
        res = []
        for _, r in ev_main.iterrows():
            e = change_effect(df_all, r["product_id"], r["日期"])
            if not e or pd.isna(e["曝光变化%"]):
                continue
            res.append(dict(日期=r["日期"].date(), 类型=r["事件"], 链接=r["链接"][:40], 详情=r["详情"], **e))
        memo[memo_key] = pd.DataFrame(res)
    rd = memo[memo_key]

    tabs = st.tabs(["结论", "单条评估", "全部改动", "我的改动清单", "自动识别的事件"])

    # ── 我的改动清单（先渲染，结论要用）
    with tabs[3]:
        st.caption("把你记得的改动写进来（日期 YYYY-MM-DD；链接关键词 = 链接名里的一段，比如 TOP TREND；"
                   "留空日期 = 待补）。点「保存」后下次打开还在，「结论」里会自动评估每一条。")
        try:
            base_ch = pd.read_csv(CHANGES_FILE, dtype=str).fillna("") if os.path.exists(CHANGES_FILE) \
                else pd.DataFrame(KNOWN_CHANGES)
        except Exception:                                   # noqa: BLE001
            base_ch = pd.DataFrame(KNOWN_CHANGES)
        man = st.data_editor(base_ch, num_rows="dynamic", key="manual_changes")
        if st.button("💾 保存改动清单"):
            try:
                os.makedirs(os.path.dirname(CHANGES_FILE), exist_ok=True)
                man.to_csv(CHANGES_FILE, index=False, encoding="utf-8-sig")
                st.success(f"已保存到 {CHANGES_FILE}")
            except Exception as e:                          # noqa: BLE001
                st.error(f"保存失败：{e}")
        isum = df_all.groupby("product_id")[K["impr"]].sum()
        res_m, man_cards = [], []
        for _, r in man.iterrows():
            d_, kw = str(r.get("日期", "")).strip(), str(r.get("链接关键词", "")).strip()
            if not d_ or not kw:
                continue
            try:
                d0 = pd.Timestamp(d_)
            except Exception:                               # noqa: BLE001
                res_m.append({"日期": d_, "链接": kw, "判定": "日期格式不对"})
                continue
            cands = [p for p, n in nm_all.items() if kw.upper() in str(n).upper()]
            if not cands:
                res_m.append({"日期": d_, "链接": kw, "判定": "没找到名称含这个关键词的链接"})
                continue
            p = max(cands, key=lambda x: isum.get(x, 0))       # 同名多条时取曝光最高的
            e = change_effect(df_all, p, d0)
            if not e:
                res_m.append({"日期": d_, "链接": short_name(nm_all[p]), "改动": r.get("改动", ""),
                              "判定": f"改动前后数据不足（前 ≥3 天、后 ≥2 天；数据到 {df_all['日期'].max():%m-%d}）"})
                continue
            res_m.append({"日期": d_, "链接": short_name(nm_all[p]), "改动": r.get("改动", ""), **_effect_row(e, 7)})
            man_cards.append({"level": _effect_level(e["判定"]), "title": f"{d0:%m-%d} {short_name(nm_all[p])}",
                              "tag": e["判定"][2:], "body": [str(r.get("改动", "")) or "（没写改动内容）"] + _effect_lines(e, 7)})
        if res_m:
            section_title("你登记的改动 → 改后效果", "改后 7 天日均 vs 改前 7 天日均")
            st.dataframe(pd.DataFrame(res_m), hide_index=True)

    with head:
        page_header("改动效果溯源", "你改过的链接，改完曝光有没有涨。曝光是索引列，先看曝光；成交（GMV）做补充。",
                    [f"{PROF['name']}数据", "口径：改后 7 天日均 vs 改前 7 天日均", f"数据到 {df_all['日期'].max():%m/%d}"])
        n_up = int((rd["曝光变化%"] >= 10).sum()) if len(rd) else 0
        n_dn = int((rd["曝光变化%"] <= -10).sum()) if len(rd) else 0
        summary_pills([("info", f"自动识别改动 {len(ev_main)} 次（能评估 {len(rd)} 次）"),
                       ("up" if n_up else "flat", f"改后曝光涨 ≥10%：{n_up} 次"),
                       ("down" if n_dn else "flat", f"改后曝光跌 ≥10%：{n_dn} 次"),
                       ("info", f"你登记的改动：{len(man_cards)} 条可评估")])

    # ── 结论
    with tabs[0]:
        if man_cards:
            section_title("你登记的改动", "改后 7 天日均 vs 改前 7 天日均；在「我的改动清单」里可以增删")
            insight_cards(man_cards, cols=2)
        if len(rd):
            section_title("自动识别的改动（改名 / 重组 + 改价 / 改促销）")
            ok_ = rd[~rd["基数太小"]]                    # 改前基数太小的百分比会很夸张，不进「最多」榜
            best = ok_.sort_values("曝光变化%").iloc[::-1].head(3)
            worst = ok_.sort_values("曝光变化%").head(3)
            cards = [{"level": "info", "title": f"{len(rd)} 次改动，改后曝光：涨 {n_up} 次 / 跌 {n_dn} 次",
                      "tag": "±10% 以内算没动",
                      "body": [f"改后曝光 **和** 成交都涨（✅）：{int(rd['判定'].str.startswith('✅').sum())} 次；"
                               f"改后曝光掉了（❌）：{int(rd['判定'].str.startswith('❌').sum())} 次；"
                               f"一涨一跌（🟡）：{int(rd['判定'].str.startswith('🟡').sum())} 次。"]},
                     {"level": "up", "title": "改完曝光涨得最多", "tag": "改后 7 天 · 不含基数太小的",
                      "body": [f"{r['日期']:%m-%d} {r['链接']}（{r['类型'].split('（')[0]}）：日均曝光 {fv('impr', r['曝光改前'])} → "
                               f"{fv('impr', r['曝光改后'])}（{r['曝光变化%']:+.0f}%），GMV "
                               + ("—" if pd.isna(r["GMV变化%"]) else f"{r['GMV变化%']:+.0f}%") + "。"
                               for _, r in best.iterrows() if r["曝光变化%"] > 0] or ["没有改后曝光上涨的改动。"]},
                     {"level": "down", "title": "改完曝光掉得最多", "tag": "改后 7 天 · 不含基数太小的",
                      "body": [f"{r['日期']:%m-%d} {r['链接']}（{r['类型'].split('（')[0]}）：日均曝光 {fv('impr', r['曝光改前'])} → "
                               f"{fv('impr', r['曝光改后'])}（{r['曝光变化%']:+.0f}%），GMV "
                               + ("—" if pd.isna(r["GMV变化%"]) else f"{r['GMV变化%']:+.0f}%") + "。"
                               for _, r in worst.iterrows() if r["曝光变化%"] < 0] or ["没有改后曝光下跌的改动。"]}]
            insight_cards(cards, cols=3)
        if not man_cards and not len(rd):
            st.info("还没有能评估的改动：自动识别的事件前后数据不足，你登记的改动也没对上。")

    # ── 单条评估
    with tabs[1]:
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            pid9, pname9 = link_selector(f"lnk9_{PROF['key']}", allow_all=False)
        evd = ev[(ev["product_id"] == pid9) & (ev["事件"] != "上架/首次出现")]["日期"].dt.date.tolist() if pid9 else []
        default_dates = sorted(set(evd))
        with c2:
            if default_dates:
                edate = st.selectbox("改动日（自动识别）", default_dates[::-1], key="eff_date")
            else:
                edate = st.date_input("改动日（手填）", value=(df_all["日期"].max() - pd.Timedelta(days=14)).date(),
                                      min_value=df_all["日期"].min().date(), max_value=df_all["日期"].max().date(),
                                      key=f"eff_date_in_{PROF['key']}")
        win9 = c3.slider("前后各看几天", 3, 21, 7, key="eff_win")
        if pid9:
            e9 = change_effect(df_all, pid9, edate, win=win9)
            if not e9:
                st.warning(f"{edate} 离数据的开头或结尾太近（至少要改前 3 天、改后 2 天数据，当前数据到 "
                           f"{df_all['日期'].max():%m-%d}）。换一个改动日，或等新数据进来再看。")
            else:
                insight_cards([{"level": _effect_level(e9["判定"]), "title": f"{short_name(pname9)} · {pd.Timestamp(edate):%m-%d} 改动",
                                "tag": e9["判定"][2:], "body": _effect_lines(e9, win9)}])
                cols9 = st.columns(4)
                cols9[0].metric(f"日均曝光（改后{win9}天）", fv("impr", e9["曝光改后"]),
                                "—" if pd.isna(e9["曝光变化%"]) else f"{e9['曝光变化%']:+.1f}% vs 改前")
                cols9[1].metric("CTR（改后）", fmt_val(e9["CTR改后"], "pct"),
                                chg_txt(e9["CTR改后"], e9["CTR改前"], "pct") + " vs 改前")
                cols9[2].metric("CTOR（改后）", fmt_val(e9["CTOR改后"], "pct"),
                                chg_txt(e9["CTOR改后"], e9["CTOR改前"], "pct") + " vs 改前")
                cols9[3].metric(f"日均GMV（改后{win9}天）", fv("gmv", e9["GMV改后"]),
                                "—" if pd.isna(e9["GMV变化%"]) else f"{e9['GMV变化%']:+.1f}% vs 改前")
            s_main = series_of(df_all, "impr", pid9)
            fm = go.Figure()
            fm.add_trace(go.Scatter(x=s_main["日期"], y=s_main["值"], mode="lines", name="曝光（IM）",
                                    line=dict(color="#4C6EF5", width=2)))
            sg = series_of(df_all, "gmv", pid9)
            fm.add_trace(go.Scatter(x=sg["日期"], y=sg["值"], mode="lines", name="GMV（$）",
                                    line=dict(color="#C2416B", width=2), yaxis="y2"))
            fm.add_vline(x=pd.Timestamp(edate), line_color="#9A6512", line_dash="dash")
            fm.add_annotation(x=pd.Timestamp(edate), y=1, yref="paper", text="改动",
                              showarrow=False, bgcolor="#FBEFD9", font=dict(color="#9A6512"))
            for d0 in default_dates:
                if d0 != edate:
                    fm.add_vline(x=pd.Timestamp(d0), line_color="#B8A4AC", line_dash="dot", opacity=.6)
            fm.update_layout(height=380, hovermode="x unified",
                             yaxis=dict(title="曝光 IM"), yaxis2=dict(title="GMV $", overlaying="y", side="right"),
                             legend=dict(orientation="h", y=1.14), margin=dict(t=40))
            st.plotly_chart(fm, key="eff_line")
            rec9 = recovery_after(series_of(df_all, "impr", pid9), edate, pre=win9, post=win9 * 2)
            if rec9 and rec9["baseline"]:
                st.caption(f"曝光逐日追踪（基线 = 改前 {rec9['pre_n']} 天中位数 {fv('impr', rec9['baseline'])}）："
                           f"{rec9['verdict']}；最低 {fv('impr', rec9['trough'])}"
                           f"（{(rec9['trough'] / rec9['baseline'] - 1) * 100:+.0f}% vs 基线）。")

    # ── 全部改动
    with tabs[2]:
        if rd.empty:
            st.info("没有前后数据都够的改动事件。")
        else:
            rs = rd.sort_values("曝光变化%")
            fb = px.bar(rs, x="曝光变化%", y=rs["链接"] + " · " + rs["日期"].astype(str), orientation="h",
                        color=np.where(rs["曝光变化%"] >= 0, "改后曝光涨", "改后曝光跌"),
                        color_discrete_map={"改后曝光涨": "#1F8A6B", "改后曝光跌": "#C24338"},
                        hover_data={"GMV变化%": ":+.0f"},
                        labels={"曝光变化%": "改后 7 天日均 vs 改前 7 天日均 · 曝光变化 %", "y": ""})
            fb.update_layout(height=max(320, 24 * len(rs)), margin=dict(t=10), legend_title="")
            st.plotly_chart(fb, key="eff_all")
            nice = pd.concat([rs[["日期", "类型", "链接", "详情"]].reset_index(drop=True),
                              pd.DataFrame([_effect_row(r, 7) for _, r in rs.iterrows()])], axis=1)
            st.dataframe(nice, hide_index=True)
            st.caption("口径：改后 7 天日均 vs 改前 7 天日均；判定先看曝光（±10%），再看成交跟没跟上。")

    # ── 自动识别的事件
    with tabs[4]:
        kinds = st.multiselect("事件类型", sorted(ev["事件"].unique()),
                               default=[k for k in ev["事件"].unique() if k != "上架/首次出现"], key="eff_kinds")
        evf = ev[ev["事件"].isin(kinds)] if kinds else ev
        st.dataframe(evf[["日期", "事件", "链接", "详情"]].iloc[::-1], hide_index=True, height=360)
        st.caption(f"共识别 {len(evf)} 个事件。改名/重组 = 你对总链的调整；件单价突变 = 改价 / 换折扣 / 换 promotion。")

# ──────────────────────────────────────────────────────────────────────
# 6b. 达人侧：按达人 / 按内容 / 按挂车关系 的归因，以及运营页里附带的「达人侧原因」
# ──────────────────────────────────────────────────────────────────────
# 页面上选的指标 → Creators 表（达人 × 天）/ 内容表里对应的列；没有对应字段的按 GMV 拆
_CR_DIM_COL = {"gmv": "GMV", "orders": "订单", "sku": "订单", "impr": "曝光", "new_vid": "新发视频",
               "new_live": "新开直播", "v_views": "播放", "items": "件数", "comm": "佣金", "cust": "买家",
               "samples": "寄样"}
_CR_CONTENT_COL = {"gmv": "GMV", "orders": "订单", "sku": "订单", "impr": "曝光", "clicks": "点击",
                   "v_views": "播放/观众"}
_CR_MONEY = {"GMV", "挂车GMV", "视频GMV", "直播GMV", "佣金", "退款"}


def _ts_list(days) -> list:
    return sorted({pd.Timestamp(d) for d in days})


def _dim_delta(frame: pd.DataFrame, key, val: str, cur_days, base_days) -> pd.DataFrame:
    """某个维度（达人 / 内容 / 挂车达人 / 链接）上的两段对比：本期合计 vs 基期（按天数折算成本期长度）。
    所有行的 Δ 相加 = 合计变化；返回 本期 / 基期 / Δ / 自身变化% / 占跌量% / 占涨量%，按 Δ 升序。"""
    cd, bd = _ts_list(cur_days), _ts_list(base_days)
    scale = len(cd) / len(bd) if bd else np.nan
    cur = frame[frame["日期"].isin(cd)].groupby(key)[val].sum()
    base = frame[frame["日期"].isin(bd)].groupby(key)[val].sum() * scale
    t = pd.concat([base.rename("基期"), cur.rename("本期")], axis=1).fillna(0.0)
    t = t[(t["基期"] != 0) | (t["本期"] != 0)].copy()
    t["Δ"] = t["本期"] - t["基期"]
    t["自身变化%"] = np.where(t["基期"] > 0, (t["本期"] / t["基期"].replace(0, np.nan) - 1) * 100, np.nan)
    down, up = t.loc[t["Δ"] < 0, "Δ"].sum(), t.loc[t["Δ"] > 0, "Δ"].sum()
    t["占跌量%"] = np.where(t["Δ"] < 0, t["Δ"] / down * 100 if down else np.nan, np.nan)
    t["占涨量%"] = np.where(t["Δ"] > 0, t["Δ"] / up * 100 if up else np.nan, np.nan)
    return t.sort_values("Δ")


def _fmt_dim(t: pd.DataFrame, val: str) -> pd.DataFrame:
    money = val in _CR_MONEY
    u = " IM" if "曝光" in val else ""
    d = t.copy()
    for c in ("基期", "本期"):
        d[c] = d[c].map(lambda v: f"${v:,.0f}" if money else f"{v:,.0f}{u}")
    d["Δ"] = d["Δ"].map(lambda v: usd(v) if money else f"{v:+,.0f}{u}")
    d["自身变化%"] = d["自身变化%"].map(lambda v: "新出现" if pd.isna(v) else f"{v:+.0f}%")
    for c in ("占跌量%", "占涨量%"):
        d[c] = d[c].map(lambda v: "" if pd.isna(v) else f"{v:.1f}%")
    return d.rename(columns={"基期": "基期（已折算）"})


def _link_short_names() -> dict:
    src = df_all if PROF["key"] == "cr" else CRB["link"]
    return {p: short_name(n) for p, n in src.groupby("product_id")["product_name"].last().items()}


def _anchor_names(ids: str, names: dict, n: int = 2) -> str:
    parts = [names.get(p.strip(), "…" + p.strip()[-4:]) for p in str(ids).split(",") if p.strip()]
    return "、".join(parts[:n]) + (f" 等{len(parts)}条" if len(parts) > n else "")


def _main_link_of(creators, days) -> pd.Series:
    """每个达人挂车 GMV 最高的那条链接（没有挂车成交的按挂车曝光）。"""
    lc = CRB["lc"]
    x = lc[lc["日期"].isin(_ts_list(days)) & lc["达人"].isin(list(creators))]
    if x.empty:
        return pd.Series(dtype=str)
    g = x.groupby(["达人", "product_id"])[["挂车GMV", "挂车曝光"]].sum().reset_index()
    g = g.sort_values(["挂车GMV", "挂车曝光"]).drop_duplicates("达人", keep="last").set_index("达人")["product_id"]
    names = _link_short_names()
    return g.map(lambda p: names.get(p, p))


def cr_link_reason(pid: str, cur_days, base_days) -> str:
    """运营页用：某条链接的变化主要在「达人」渠道 → 达人组数据里这条链的内容供给、挂车达人怎么变的（全部日均）。"""
    L = CRB["link"]
    cov = set(pd.to_datetime(L["日期"].unique()))
    cur_all = _ts_list(cur_days)
    cd, bd = [d for d in cur_all if d in cov], [d for d in _ts_list(base_days) if d in cov]
    if not cd or not bd:
        return f"达人数据没覆盖这段（只到 {max(cov):%m-%d}）"
    if len(cd) * 2 < len(cur_all):
        return f"达人数据只到 {max(cov):%m-%d}，本期只覆盖 {len(cd)}/{len(cur_all)} 天，先不下结论"
    lk = L[L["product_id"] == pid]

    def per_day(days, k):
        return lk.loc[lk["日期"].isin(days), K_CR[k]].sum() / len(days)

    bits = []
    ia, ib = per_day(cd, "impr"), per_day(bd, "impr")
    if ia or ib:
        bits.append(f"曝光 {fv('impr', ib)}→{fv('impr', ia)}/天" + (f"（{pct(ia, ib):+.0f}%）" if ib else ""))
        vi, vb = per_day(cd, "i_vid"), per_day(bd, "i_vid")
        li, lb = per_day(cd, "i_live"), per_day(bd, "i_live")
        if vi or vb or li or lb:
            bits.append(f"其中视频挂车 {fdelta('impr', vi - vb)}/天、直播挂车 {fdelta('impr', li - lb)}/天")
    for k, nm in (("new_vid", "新发视频"), ("posted", "发布达人"), ("cws", "出单达人")):
        a, b = per_day(cd, k), per_day(bd, k)
        if a or b:
            bits.append(f"{nm} {b:.1f}→{a:.1f}/天")
    ga, gb = per_day(cd, "gmv"), per_day(bd, "gmv")
    if ga or gb:
        bits.append(f"达人 GMV ${gb:,.0f}→${ga:,.0f}/天")
    falling = ia < ib
    t = _dim_delta(CRB["lc"][CRB["lc"]["product_id"] == pid], "达人", "挂车曝光", cd, bd)
    t = t[t["Δ"] < 0] if falling else t[t["Δ"] > 0].iloc[::-1]
    t = t[t["Δ"].abs() >= 1].head(2)
    who = "、".join(f"@{n} {fv('impr', r['基期'] / len(cd))}→{fv('impr', r['本期'] / len(cd))}/天" for n, r in t.iterrows())
    note = f"（达人数据只覆盖其中 {len(cd)}/{len(cur_all)} 天）" if len(cd) < len(cur_all) else ""
    out = "；".join(bits) + (f"；曝光{'掉' if falling else '涨'}得最多的挂车达人：{who}" if who else "")
    return note + (out or "达人侧内容供给没有明显变化")


def _anomaly_extra(cp: pd.DataFrame, cur_days, base_days, direction: int, by: str) -> dict:
    """异常归因清单的附加列：运营 → 主责链接里变化在达人渠道的，附达人侧原因；达人 → 主责达人。"""
    if PROF["key"] == "ops":
        if CRB is None or cp is None or not len(cp) or "主要渠道" not in cp.columns:
            return {}
        rows = [f"{r['链接']}：{cr_link_reason(r['product_id'], cur_days, base_days)}"
                for _, r in cp.iterrows() if r.get("主要渠道") == "达人"][:2]
        return {"达人侧原因（主责链接里变化在达人渠道的）": "；".join(rows) if rows else "—"}
    col_ = _CR_DIM_COL.get(by, "GMV")
    t = _dim_delta(CRB["creator_day"], "达人", col_, cur_days, base_days)
    t = t[t["Δ"] < 0] if direction < 0 else t[t["Δ"] > 0].iloc[::-1]
    f_ = usd if col_ in _CR_MONEY else (lambda v: f"{v:+,.0f}")
    return {f"主责达人（{col_}）": "、".join(f"@{n}（{f_(r['Δ'])}）" for n, r in t.head(3).iterrows()) or "—"}


def render_creator_cross_check(cur_days, base_days, sc: dict, sb: dict):
    """运营「两段对比」：运营只看得到「达人 Affiliate 渠道」的结果数字；同样的两段时间换到达人组数据里看
    供给端——曝光、内容量、出单达人、内容质量、具体哪些达人在动——跟渠道涨跌方向对不对得上。曝光先看。"""
    L = CRB["link"]
    cov = set(pd.to_datetime(L["日期"].unique()))
    cur_all, base_all = _ts_list(cur_days), _ts_list(base_days)
    cd, bd = [d for d in cur_all if d in cov], [d for d in base_all if d in cov]
    section_title("达人侧同期对照：运营看到「达人渠道」变了，达人组数据里为什么",
                  "同样的本期 / 基期换到达人组自己的数据，全部按日均比（两段天数、覆盖天数可能不同）")
    if not cd or not bd:
        st.caption(f"达人组数据（{min(cov):%m-%d} → {max(cov):%m-%d}）没覆盖这两段，对照不了。")
        return
    if len(cd) < len(cur_all) or len(bd) < len(base_all):
        st.warning(md(f"⚠️ 达人组数据只到 {max(cov):%m-%d}：本期覆盖 {len(cd)}/{len(cur_all)} 天、"
                      f"基期覆盖 {len(bd)}/{len(base_all)} 天，下面只用覆盖到的天算日均。"))

    def ops_pd(days, k):
        return OPS_ALL.loc[OPS_ALL["日期"].isin(days), K[k]].sum() / len(days)

    cday = CRB["creator_day"]

    def per_day(days, c):
        return cday.loc[cday["日期"].isin(days), c].sum() / len(days)

    def sellers(days):
        x = cday[cday["日期"].isin(days) & (cday["订单"] > 0)].groupby("日期")["达人"].nunique()
        return x.reindex(days, fill_value=0).mean()

    vid = CRB["content"][CRB["content"]["类型"] == "视频"]

    def vq(days):
        v = vid[vid["日期"].isin(days)]
        w = v["播放/观众"].sum()
        if not w:
            return np.nan, np.nan
        return (v["完播率%"] * v["播放/观众"]).sum() / w, (v["赞"] + v["评"] + v["转"]).sum() / w * 100

    oi_c, oi_b = ops_pd(cd, "i_cre"), ops_pd(bd, "i_cre")
    og_c, og_b = ops_pd(cd, "ch_cre"), ops_pd(bd, "ch_cre")
    ci_c, ci_b = per_day(cd, "曝光"), per_day(bd, "曝光")
    ck_c, ck_b = per_day(cd, "点击"), per_day(bd, "点击")
    g_c = L.loc[L["日期"].isin(cd), K_CR["gmv"]].sum() / len(cd)
    ctr_c = ck_c / ci_c * 100 if ci_c else np.nan
    ctr_b = ck_b / ci_b * 100 if ci_b else np.nan
    s_c, s_b = sellers(cd), sellers(bd)
    nv_c, nv_b = per_day(cd, "新发视频"), per_day(bd, "新发视频")
    nl_c, nl_b = per_day(cd, "新开直播"), per_day(bd, "新开直播")
    (cp_c, en_c), (cp_b, en_b) = vq(cd), vq(bd)
    r1 = st.columns(4)
    r1[0].metric("运营·达人渠道曝光/天", fv("impr", oi_c), f"{chg_txt(oi_c, oi_b, 'count')} vs 基期")
    r1[1].metric("达人组·曝光/天", fv("impr", ci_c), f"{chg_txt(ci_c, ci_b, 'count')} vs 基期",
                 help="达人组的曝光口径比运营日报的「达人渠道曝光」高约 16%，看方向和幅度，不要直接比绝对值")
    r1[2].metric("达人组·CTR", fmt_val(ctr_c, "pct"), f"{chg_txt(ctr_c, ctr_b, 'pct')} vs 基期")
    r1[3].metric("运营·达人渠道 GMV/天", usd(og_c, sign=False), f"{chg_txt(og_c, og_b, 'money')} vs 基期")
    r2 = st.columns(4)
    r2[0].metric("出单达人/天（去重）", f"{s_c:.1f} 人", f"{s_c - s_b:+.1f} vs 基期")
    r2[1].metric("新发视频/天", f"{nv_c:.1f} 条", f"{chg_txt(nv_c, nv_b, 'count')} vs 基期")
    r2[2].metric("新开直播/天", f"{nl_c:.1f} 场", f"{chg_txt(nl_c, nl_b, 'count')} vs 基期")
    r2[3].metric("视频完播率（播放加权）", fmt_val(cp_c, "pct"), f"{chg_txt(cp_c, cp_b, 'pct')} vs 基期")
    diff = pct(g_c, og_c)
    if pd.notna(diff):
        st.caption(f"同一批天里两个数据源的达人 GMV 差 {diff:+.1f}%（正常应接近 0：逐链逐日 99.6% 一致；差得多说明窗口没对齐）。")

    if len(cd) * 2 < len(cur_all):
        st.info(md(f"本期只有 {len(cd)}/{len(cur_all)} 天有达人数据，上面的数字只代表这 {len(cd)} 天，"
                   f"**先不下归因结论**。把「本期」挪到 {max(cov):%m-%d} 之前（或等达人组数据更新）就能看到完整对照。"))
        return
    ich = pct(oi_c, oi_b)
    cl, bl = f"本期（{cd[0]:%m-%d}～{cd[-1]:%m-%d}）日均", f"基期（{bd[0]:%m-%d}～{bd[-1]:%m-%d}）日均"
    lines = [say_change("达人渠道曝光（运营）", "impr", oi_c, oi_b, cl, bl),
             say_change("达人组曝光", "impr", ci_c, ci_b, cl, bl),
             say_change("达人组 CTR", "CTR 点击率", ctr_c, ctr_b, cl.replace("日均", ""), bl.replace("日均", "")),
             say_change("达人渠道 GMV（运营）", "gmv", og_c, og_b, cl, bl)]
    if pd.notna(ich) and abs(ich) >= 5:
        drivers = []
        for nm, a, b, u in (("新发视频", nv_c, nv_b, "条"), ("新开直播", nl_c, nl_b, "场"), ("出单达人", s_c, s_b, "人")):
            ch = pct(a, b)
            if pd.notna(ch) and abs(ch) >= 5 and (ch > 0) == (ich > 0):
                drivers.append(f"{nm} {b:.1f}→{a:.1f} {u}/天（{ch:+.0f}%）")
        if drivers:
            lines.append(f"跟曝光同方向的供给变化：{'、'.join(drivers)}——**内容供给**大概率是主因。")
        elif pd.notna(cp_c) and pd.notna(cp_b) and abs(cp_c - cp_b) >= 1:
            lines.append(f"内容量、出单达人都没同向变化，但视频完播率 {cp_b:.1f}%→{cp_c:.1f}%——更像**内容质量**在变。")
        else:
            lines.append("内容量、出单达人、完播率都没有明显同向变化——更像**头部达人个体**或达人组以外的因素（看下表）。")
        if pd.notna(en_c) and pd.notna(en_b) and abs(en_c - en_b) >= 0.3:
            lines.append(f"视频互动率（赞评转 ÷ 播放）{en_b:.2f}% → {en_c:.2f}%。")
        t = _dim_delta(cday, "达人", "曝光", cd, bd)
        mover = (t[t["Δ"] < 0] if ich < 0 else t[t["Δ"] > 0].iloc[::-1]).head(3)
        if len(mover):
            lines.append(f"曝光{'掉' if ich < 0 else '涨'}得最多的达人：" +
                         "、".join(f"@{n}（{fdelta('impr', r['Δ'] / len(cd))}/天）" for n, r in mover.iterrows()) + "。")
    insight_cards([{"level": level_of(ich), "title": f"达人渠道曝光 {ich:+.1f}%" if pd.notna(ich) else "达人渠道曝光",
                    "tag": "日均 vs 基期", "body": lines}])
    st.markdown("**变化最大的达人（全部指标，Creators 表；按 Δ曝光 排，基期按天数折算）**")
    st.dataframe(creator_multi_attr(cd, bd, "曝光", topn=10), hide_index=True)


def render_creator_dimensions(cur_days, base_days, pick_m: str):
    """达人版「两段对比」：链接之外，再按达人、按视频/直播拆，以及单条链接是哪些达人在推。"""
    key_m = PROF["compare_opts"].get(pick_m, "gmv")
    col_c = _CR_DIM_COL.get(key_m, "GMV")
    cd, bd = _ts_list(cur_days), _ts_list(base_days)
    names = _link_short_names()

    # ① 按达人
    st.markdown("---")
    st.markdown(f"#### 谁导致的：按达人拆（{col_c}；Creators 表精确，所有达人 Δ 相加 = 达人合计变化）")
    if col_c == "GMV" and key_m != "gmv":
        st.caption(f"「{pick_m}」在达人层级没有对应字段，这里按 GMV 拆。")
    t = _dim_delta(CRB["creator_day"], "达人", col_c, cd, bd)
    if t.empty:
        st.info("两段里都没有数据。")
    else:
        money = col_c in _CR_MONEY
        sgn = usd if money else (lambda v: f"{v:+,.0f}")
        net, down, up = t["Δ"].sum(), t.loc[t["Δ"] < 0, "Δ"].sum(), t.loc[t["Δ"] > 0, "Δ"].sum()
        top5 = t[t["Δ"] < 0].head(5)["占跌量%"].sum()
        st.markdown(md(f"达人合计 **{col_c}** 净变化 **{sgn(net)}** = {int((t['Δ'] < 0).sum())} 个达人下降合计 {sgn(down)}"
                       f" ＋ {int((t['Δ'] > 0).sum())} 个达人上升合计 {sgn(up)}；拖累最多的 5 个达人占总跌量 **{top5:.0f}%**。"))
        st.plotly_chart(contrib_bars(t.reset_index(), "Δ", xtitle=f"Δ{col_c}（本期 − 基期）", label_col="达人"))
        tbl = pd.concat([t[t["Δ"] < 0].head(15), t[t["Δ"] > 0].tail(10).iloc[::-1]])
        nv = _dim_delta(CRB["creator_day"][CRB["creator_day"]["达人"].isin(tbl.index)], "达人", "新发视频", cd, bd)
        disp = _fmt_dim(tbl, col_c)
        disp["新发视频 基期→本期"] = [f"{nv.at[n, '基期']:.0f}→{nv.at[n, '本期']:.0f}" if n in nv.index else "0→0"
                                 for n in tbl.index]
        disp["主推链接（挂车GMV最高）"] = _main_link_of(tbl.index, cd + bd).reindex(tbl.index).fillna("—").values
        st.dataframe(disp.reset_index(), hide_index=True)

    # ② 按视频 / 直播
    col_v = _CR_CONTENT_COL.get(key_m, "GMV")
    st.markdown(f"#### 谁导致的：按视频 / 直播拆（{col_v}；每条内容精确）")
    cont = CRB["content"]
    t2 = _dim_delta(cont, ["类型", "内容ID"], col_v, cd, bd)
    if t2.empty:
        st.caption("两段里都没有数据。")
    else:
        meta = (cont[cont["日期"].isin(cd + bd)].drop_duplicates(["类型", "内容ID"], keep="last")
                .set_index(["类型", "内容ID"])[["达人", "标题", "挂车链接", "发布时间"]])
        t2 = t2.join(meta)
        t2["内容"] = ("@" + t2["达人"].astype(str) + " · " + t2["标题"].astype(str).str.slice(0, 20)
                      + " ·" + t2.index.get_level_values("内容ID").astype(str).str[-4:])
        st.plotly_chart(contrib_bars(t2.reset_index(), "Δ", n=10, xtitle=f"Δ{col_v}（本期 − 基期）", label_col="内容"))
        tb2 = pd.concat([t2[t2["Δ"] < 0].head(12), t2[t2["Δ"] > 0].tail(8).iloc[::-1]])
        disp2 = _fmt_dim(tb2, col_v)
        disp2["挂车链接"] = tb2["挂车链接"].map(lambda x: _anchor_names(x, names))
        st.dataframe(disp2.reset_index()[["类型", "达人", "标题", "挂车链接", "发布时间", "基期（已折算）", "本期",
                                          "Δ", "自身变化%", "占跌量%", "占涨量%"]], hide_index=True)

    # ③ 链接 → 达人（挂车口径）
    st.markdown("#### 选一条链接：是哪些达人在推它（挂车口径）")
    st.caption("『挂车 GMV』= 挂了这条链接的视频/直播带来的成交（一条内容挂多条链接时平摊）。视频带来的成交里买家"
               "可能最后买了别的链接，所以加起来不一定等于这条链接的 GMV——用来看『谁在推这条链、谁停了』。")
    lg = _dim_delta(df_all, "product_id", K["impr"], cd, bd)
    opts = lg.reindex(lg["Δ"].abs().sort_values(ascending=False).index).head(40).index.tolist()
    if not opts:
        return
    pid = st.selectbox("链接（按曝光变化幅度排序）", opts, key="cr_dim_link",
                       format_func=lambda p: f"{names.get(p, p)}　（曝光 {fdelta('impr', lg.at[p, 'Δ'])} vs 基期）")
    x = CRB["lc"][CRB["lc"]["product_id"] == pid]
    t3 = _dim_delta(x, "达人", "挂车GMV", cd, bd)
    if t3.empty:
        st.caption("这两段里没有挂这条链接的内容。")
        return
    tb3 = pd.concat([t3[t3["Δ"] < 0].head(12), t3[t3["Δ"] > 0].tail(8).iloc[::-1]])
    disp3 = _fmt_dim(tb3, "挂车GMV")
    for c, nm in (("新发视频", "新发视频"), ("活跃视频", "活跃视频·天")):
        z = _dim_delta(x[x["达人"].isin(tb3.index)], "达人", c, cd, bd)
        disp3[f"{nm} 基期→本期"] = [f"{z.at[n, '基期']:.0f}→{z.at[n, '本期']:.0f}" if n in z.index else "0→0"
                                  for n in tb3.index]
    st.dataframe(disp3.reset_index(), hide_index=True)


def render_creator_link_extras(pid: str, sub: pd.DataFrame, gran: str):
    """达人版「链接下钻」追加：这条链占达人合计的比重、内容供给、是哪些达人在推、卖得最好的内容。"""
    days = df["日期"].unique()
    unit = {"日": "天", "周": "周", "月": "月"}.get(gran, gran)
    section_title("这条链占达人合计的比重", "曝光占比先看")
    ks = [("impr", "曝光"), ("gmv", "GMV"), ("orders", "订单"), ("new_vid", "新发视频")]
    tot = df.assign(_p=to_period(df["日期"], gran)).groupby("_p")[[K[k] for k, _ in ks]].sum()
    me = (sub.assign(_p=to_period(sub["日期"], gran)).groupby("_p")[[K[k] for k, _ in ks]].sum()
          .reindex(tot.index, fill_value=0.0))
    share = me / tot.replace(0, np.nan) * 100
    share.columns = [f"{n}占比%" for _, n in ks]
    share = share.reset_index().rename(columns={"_p": "期间"})
    fs = px.line(share, x="期间", y=list(share.columns[1:]), markers=True, labels={"value": "%", "期间": ""})
    fs.update_layout(height=320, legend_title="", hovermode="x unified", margin=dict(t=10, b=10))
    st.plotly_chart(fs, key="cr_dd_share")
    tb = share.copy()
    tb.insert(1, "本链曝光", me[K["impr"]].map(lambda v: fv("impr", v)).values)
    tb.insert(2, "达人合计曝光", tot[K["impr"]].map(lambda v: fv("impr", v)).values)
    tb.insert(3, "本链 GMV", me[K["gmv"]].map(lambda v: fv("gmv", v)).values)
    tb.insert(4, "达人合计 GMV", tot[K["gmv"]].map(lambda v: fv("gmv", v)).values)
    for c in share.columns[1:]:
        tb[c] = tb[c].map(lambda v: "—" if pd.isna(v) else f"{v:.1f}%")
    st.dataframe(tb.iloc[::-1], hide_index=True)

    section_title(f"内容供给：每{unit}新发视频 / 新开直播 / 发布达人 / 出单达人（这条链）")
    sup = sub.assign(_p=to_period(sub["日期"], gran)).groupby("_p")[
        [K["new_vid"], K["new_live"], K["posted"], K["cws"]]].sum()
    sup.columns = ["新发视频", "新开直播", "发布达人", "出单达人"]
    sup = sup.reset_index().rename(columns={"_p": "期间"})
    fsu = px.bar(sup, x="期间", y=["新发视频", "新开直播", "发布达人", "出单达人"], barmode="group",
                 labels={"value": "", "期间": ""})
    fsu.update_layout(height=300, legend_title="", hovermode="x unified", margin=dict(t=10, b=10))
    st.plotly_chart(fsu, key="cr_dd_sup")

    names = _link_short_names()
    section_title("是哪些达人在推这条链", "挂车口径，当前分析区间，按挂车曝光排")
    x = CRB["lc"][(CRB["lc"]["product_id"] == pid) & CRB["lc"]["日期"].isin(days)]
    if x.empty:
        st.caption("区间内没有挂这条链的视频/直播。")
    else:
        g = x.groupby("达人").agg(挂车曝光=("挂车曝光", "sum"), 播放=("播放", "sum"), 挂车GMV=("挂车GMV", "sum"),
                                 视频订单=("视频订单", "sum"), 新发视频=("新发视频", "sum"), 活跃视频天=("活跃视频", "sum"))
        g = g.sort_values(["挂车曝光", "挂车GMV"], ascending=False)
        tot_i, tot_g = g["挂车曝光"].sum(), g["挂车GMV"].sum()
        g["占这条链挂车曝光%"] = g["挂车曝光"] / tot_i * 100 if tot_i else np.nan
        g["占这条链挂车GMV%"] = g["挂车GMV"] / tot_g * 100 if tot_g else np.nan
        top = g.head(20).copy()
        top["挂车曝光"] = top["挂车曝光"].map(lambda v: fv("impr", v))
        top["挂车GMV"] = top["挂车GMV"].map(lambda v: fv("gmv", v))
        for c in ("占这条链挂车曝光%", "占这条链挂车GMV%"):
            top[c] = top[c].map(lambda v: "—" if pd.isna(v) else f"{v:.1f}%")
        top = top[["挂车曝光", "占这条链挂车曝光%", "播放", "挂车GMV", "占这条链挂车GMV%", "视频订单", "新发视频", "活跃视频天"]]
        st.caption(f"共 {len(g)} 个达人挂过这条链；挂车曝光前 5 的达人占 "
                   f"{g['挂车曝光'].head(5).sum() / tot_i * 100:.0f}%。" if tot_i else f"共 {len(g)} 个达人挂过这条链。")
        st.dataframe(round_num(top.reset_index(), 1), hide_index=True)

    section_title("挂这条链的视频 / 直播", "当前分析区间，按曝光排")
    cont = CRB["content"]
    c = cont[cont["日期"].isin(days) & cont["挂车链接"].str.contains(str(pid), regex=False)]
    if c.empty:
        st.caption("区间内没有挂这条链的内容。")
        return
    tc = c.groupby(["类型", "内容ID"]).agg(内容GMV=("GMV", "sum"), 订单=("订单", "sum"), 播放观众=("播放/观众", "sum"),
                                          曝光=("曝光", "sum"), 达人=("达人", "last"), 标题=("标题", "last"),
                                          发布时间=("发布时间", "last"), 挂车链接=("挂车链接", "last"))
    tc = tc.sort_values(["曝光", "内容GMV"], ascending=False).head(15)
    tc["挂车链接"] = tc["挂车链接"].map(lambda v: _anchor_names(v, names))
    tc["曝光"] = tc["曝光"].map(lambda v: fv("impr", v))
    tc["内容GMV"] = tc["内容GMV"].map(lambda v: fv("gmv", v))
    tc = tc[["曝光", "播放观众", "内容GMV", "订单", "达人", "标题", "发布时间", "挂车链接"]]
    st.caption("『内容 GMV』是这条视频/直播带来的全部成交（挂多条链接时包含别的链接）。")
    st.dataframe(round_num(tc.reset_index(), 0), hide_index=True)


# ── 通用：本期自己选，基期默认自动 = 紧挨着本期前面、同样天数的一段 ──────────
def _pick_two_periods(d_first: pd.Timestamp, d_last: pd.Timestamp, key: str):
    """返回 (a0, a1, b0, b1)；没选完整返回 None。控件 key 按数据源区分（两边日期范围不同）。"""
    c1, c2 = st.columns(2)
    r_cur = c1.date_input("本期", value=((d_last - pd.Timedelta(days=6)).date(), d_last.date()),
                          min_value=d_first.date(), max_value=d_last.date(), key=f"{key}_cur_{PROF['key']}")
    auto = c2.toggle("基期自动 = 本期前面同样天数", value=True, key=f"{key}_auto_{PROF['key']}")
    if not (isinstance(r_cur, tuple) and len(r_cur) == 2):
        st.info("请把本期的起止日期选完整。")
        return None
    a0, a1 = pd.Timestamp(r_cur[0]), pd.Timestamp(r_cur[1])
    if auto:
        n = (a1 - a0).days + 1
        b1 = a0 - pd.Timedelta(days=1)
        b0 = b1 - pd.Timedelta(days=n - 1)
        c2.markdown(f"基期（跟谁比）：**{b0:%Y/%m/%d} – {b1:%Y/%m/%d}**，紧挨着本期、同样 {n} 天")
        if b0 < d_first:
            c2.caption(f"⚠️ 基期有一部分早于数据开始日 {d_first:%m/%d}，只用有数据的那几天（按天数折算）")
        return a0, a1, b0, b1
    r_base = c2.date_input("基期（跟谁比）",
                           value=((d_last - pd.Timedelta(days=13)).date(), (d_last - pd.Timedelta(days=7)).date()),
                           min_value=d_first.date(), max_value=d_last.date(), key=f"{key}_base_{PROF['key']}")
    if not (isinstance(r_base, tuple) and len(r_base) == 2):
        st.info("请把基期的起止日期选完整。")
        return None
    return a0, a1, pd.Timestamp(r_base[0]), pd.Timestamp(r_base[1])


# ── 达人版「多粒度对比」：四个 Excel 的全部字段，逐日 / 逐周 / 逐月 / 自选区间 ──
_WEEKDAY = "一二三四五六日"
_CR_LINK_CUM = {"Videos", "LIVE streams", "Creators posted content", "Creators with sales", "Videos with sales",
                "LIVE streams with sales"}
_CR_MONEY_AVG = {"AOV", "Avg. GMV per customer", "Video GPM", "Show GPM"}


def _cr_plabel(p: pd.Timestamp, gran: str, days=None) -> str:
    if gran == "日":
        return f"{p:%m-%d}（周{_WEEKDAY[p.weekday()]}）"
    if gran == "周":
        return f"{p:%m-%d}~{p + pd.Timedelta(days=6):%m-%d}" + (f"（{int(days)}天）" if days and days < 7 else "")
    full = (p + pd.offsets.MonthEnd(0)).day
    return f"{p:%Y-%m}" + (f"（{int(days)}天）" if days and days < full else "")


def _cr_is_pct(field: str, t: str | None = None) -> bool:
    """LIVE 表的 Engagement 是 TikTok 自己的互动指数（原始值 200+），不是百分比。"""
    return field in CR_PCT_FIELDS and not (t == "LIVE" and field == "Engagement")


def _cr_is_level(field: str) -> bool:
    """不能按天数摊的字段：比率/均值、去重个数。"""
    return _cr_kind(field) == "比率" or field.startswith(("表里出现的", "出单"))


def _cr_is_money(field: str) -> bool:
    return field in _CR_MONEY_AVG or (_cr_kind(field) == "金额" and field not in CR_PCT_FIELDS)


def _cr_fmt(field: str, v, avg: bool = False, t: str | None = None) -> str:
    if pd.isna(v):
        return "—"
    if field in CR_PCT_FIELDS and not _cr_is_pct(field, t):
        return f"{v:,.2f}"
    if field in CR_PCT_FIELDS:
        return f"{v:.2f}%"
    if field == "Avg. viewing duration":
        return f"{v:,.1f} 秒"
    if _cr_is_money(field):
        return f"${v:,.2f}"
    u = " IM" if "impression" in field.lower() else ""
    return (f"{v:,.1f}" if avg and not _cr_is_level(field) else f"{v:,.0f}") + u


def _cr_delta(field: str, cur, prev, t: str | None = None) -> tuple[str, float]:
    """(变化的文字, 环比%)。百分比字段的变化用 pp。"""
    if pd.isna(cur) or pd.isna(prev):
        return "—", np.nan
    d = cur - prev
    rel = (cur / prev - 1) * 100 if prev else np.nan
    if field in CR_PCT_FIELDS and not _cr_is_pct(field, t):
        return f"{d:+,.2f}", rel
    if field in CR_PCT_FIELDS:
        return f"{d:+.2f} pp", rel
    if field == "Avg. viewing duration":
        return f"{d:+.1f} 秒", rel
    if _cr_is_money(field):
        return ("+" if d >= 0 else "-") + f"${abs(d):,.2f}", rel
    u = " IM" if "impression" in field.lower() else ""
    return (f"{d:+,.1f}" if d != round(d) else f"{d:+,.0f}") + u, rel


def _cr_cn(tname: str, field: str, gran: str) -> str:
    if field.startswith(("表里出现的", "出单")):
        return "当天去重" if gran == "日" else "这一期里去重"
    if tname == "LIVE" and field == "Engagement":
        return "互动指数（TikTok 原始口径，不是百分比）"
    cn = CR_FIELD_CN.get(field, field)
    return cn + ("（按链接累计）" if tname == "Products" and field in _CR_LINK_CUM else "")


def _cr_header(field: str) -> str:
    return field if field.startswith(("表里出现的", "出单")) else f"{CR_FIELD_CN.get(field, field)}｜{field}"


def _color_css(v) -> str:
    if pd.isna(v) or abs(v) < 0.05:
        return ""
    return "color: #2E9E7A; font-weight: 600" if v > 0 else "color: #D9534F; font-weight: 600"


def _style_rows(disp: pd.DataFrame, rel: pd.Series, cols) -> "pd.io.formats.style.Styler":
    css = pd.DataFrame("", index=disp.index, columns=disp.columns)
    for c in cols:
        css[c] = [_color_css(v) for v in rel.values]
    return disp.style.apply(lambda _: css, axis=None)


def cr_period_table(tname: str, gran: str) -> pd.DataFrame:
    """行 = 期间，列 = 天数、来源 + 这张 Excel 的全部字段（合计口径）。
    逐月时，已经过完的月份如果有他们上传的「月数据」Excel 就直接用它（逐日数据 7/31 才开始，靠它才有完整的 6、7 月）。"""
    memo = st.session_state.setdefault("_crg_memo", {})
    mk = (id(CRB), tname, gran)
    if mk in memo:
        return memo[mk]
    T = CRB["raw"][tname]
    key = to_period(T["日期"], gran)
    g = cr_agg(T, tname, key)
    g.insert(0, "天数", T.groupby(key)["日期"].nunique().reindex(g.index).astype(int))
    g.insert(1, "来源", "逐日汇总")
    if gran == "月" and tname in CRB["raw_monthly"]:
        last = T["日期"].max()
        for m, Tm in CRB["raw_monthly"][tname].groupby("_month"):
            start = pd.Timestamp(year=last.year, month=int(m), day=1)
            end = start + pd.offsets.MonthEnd(0)
            if end >= last:                                  # 还没过完的月份用逐日汇总
                continue
            gm = cr_agg(Tm, tname, pd.Series(start, index=Tm.index))
            gm.insert(0, "天数", end.day)
            gm.insert(1, "来源", "月数据 Excel")
            g = pd.concat([g.drop(index=start, errors="ignore"), gm])
    g = g.sort_index()
    memo[mk] = g
    return g


def cr_window_table(tname: str, a0, a1, b0, b1) -> pd.DataFrame:
    """自选区间：本期 / 上一段 两行，列 = 天数 + 全部字段（合计口径）。"""
    T = CRB["raw"][tname]
    lab = np.where(T["日期"].between(a0, a1), "本期", np.where(T["日期"].between(b0, b1), "上一段", ""))
    m = lab != ""
    Tm = T[m]
    if Tm.empty:
        return pd.DataFrame(index=["上一段", "本期"])
    key = pd.Series(lab[m], index=Tm.index)
    g = cr_agg(Tm, tname, key)
    g.insert(0, "天数", Tm.groupby(key)["日期"].nunique())
    return g.reindex(["上一段", "本期"])


def cr_compare_frame(tname: str, cur: pd.Series, prev: pd.Series, avg: bool, gran: str) -> pd.DataFrame:
    """一张 Excel 的全部字段：上一期 / 本期 / 变化 / 环比%（字段 = 行，顺序同 Excel）。"""
    dc = cur.get("天数") if pd.notna(cur.get("天数")) and cur.get("天数") else np.nan
    dp = prev.get("天数") if pd.notna(prev.get("天数")) and prev.get("天数") else np.nan
    rows = []
    for f in [c for c in cur.index if c not in ("天数", "来源")]:
        c, p = cur[f], prev[f]
        if avg and not _cr_is_level(f):
            c, p = c / dc, p / dp
        txt, rel = _cr_delta(f, c, p, tname)
        rows.append({"字段": f, "说明": _cr_cn(tname, f, gran), "上一期": _cr_fmt(f, p, avg, tname),
                     "本期": _cr_fmt(f, c, avg, tname), "变化": txt, "环比%": "—" if pd.isna(rel) else f"{rel:+.1f}%",
                     "_rel": rel, "_prev": p, "_cur": c})
    return pd.DataFrame(rows)


def _cr_movers(frames: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """四张表合看：金额 / 计数类字段里环比变化最大的（基数太小的不算：计数 < 5、金额 < $20；
    比率类字段样本小时波动很大，放在下面各表里看，不进这张榜）。"""
    parts = []
    for t, cf in frames.items():
        x = cf.copy()
        x.insert(0, "表", t)
        parts.append(x)
    if not parts:
        return pd.DataFrame(), pd.DataFrame()
    a = pd.concat(parts, ignore_index=True)
    a = a[~a["字段"].map(_cr_is_level)]
    floor = a["字段"].map(lambda f: 20 if _cr_is_money(f) else 5)
    a = a[a["_rel"].notna() & (a["_prev"].abs() >= floor)]
    a = a.drop_duplicates(["字段", "上一期", "本期"])
    cols = ["表", "字段", "说明", "上一期", "本期", "环比%"]
    up = a[a["_rel"] >= 5].sort_values("_rel", ascending=False).head(8)
    down = a[a["_rel"] <= -5].sort_values("_rel").head(8)
    return up[cols + ["_rel"]], down[cols + ["_rel"]]


def _cr_matrix(t: str, g: pd.DataFrame, gran: str, avg: bool):
    """每一期 × 全部字段（最新在上），可切换 数值 / 环比。"""
    fields = [c for c in g.columns if c not in ("天数", "来源")]
    val = g[fields].astype(float).copy()
    if avg:
        for f in fields:
            if not _cr_is_level(f):
                val[f] = val[f] / g["天数"]
    rel = pd.DataFrame({f: (val[f].diff() if _cr_is_pct(f, t) else val[f].pct_change(fill_method=None) * 100)
                        for f in fields}).replace([np.inf, -np.inf], np.nan)
    view = st.radio("显示", ["数值", "环比（vs 上一期）"], horizontal=True, key=f"crg_view_{t}")
    if view == "数值":
        body = pd.DataFrame({_cr_header(f): [_cr_fmt(f, v, avg, t) for v in val[f]] for f in fields}, index=g.index)
        css = None
    else:
        body = pd.DataFrame({_cr_header(f): ["—" if pd.isna(v) else (f"{v:+.2f} pp" if _cr_is_pct(f, t) else f"{v:+.1f}%")
                                             for v in rel[f]] for f in fields}, index=g.index)
        css = pd.DataFrame({_cr_header(f): [_color_css(v) for v in rel[f]] for f in fields}, index=g.index)
    lead = pd.DataFrame({"期间": [_cr_plabel(p, gran, d) for p, d in zip(g.index, g["天数"])],
                         "天数": g["天数"].astype(int).values}, index=g.index)
    if gran == "月":
        lead["来源"] = g["来源"].values
    disp = pd.concat([lead, body], axis=1).iloc[::-1]
    if css is not None:
        css = pd.concat([pd.DataFrame("", index=lead.index, columns=lead.columns), css], axis=1).iloc[::-1]
        st.dataframe(disp.style.apply(lambda _: css, axis=None), hide_index=True, height=420)
    else:
        st.dataframe(disp, hide_index=True, height=420)
    out = lead.copy()
    for f in fields:
        out[f] = val[f].round(4)
        out[f"{f}｜环比{'pp' if _cr_is_pct(f, t) else '%'}"] = rel[f].round(2)
    st.download_button(f"⬇️ 下载 {t} 每期全部字段 + 环比 CSV", out.to_csv(index=False).encode("utf-8-sig"),
                       f"达人组_{t}_{gran}{'_日均' if avg else '_合计'}.csv", "text/csv", key=f"crg_dl_{t}")


def _cr_field_trend(t: str, fields: list, gran: str | None, avg: bool, g: pd.DataFrame | None, win=None):
    """单个字段的走势：逐日/周/月 = 每期数值 + 环比；自选区间 = 两段的逐日走势。"""
    f = st.selectbox("看某一个字段的走势", fields, key=f"crg_field_{t}",
                     format_func=lambda x: x if x.startswith(("表里出现的", "出单")) else f"{CR_FIELD_CN.get(x, x)}（{x}）")
    if gran:
        v = g[f].astype(float)
        if avg and not _cr_is_level(f):
            v = v / g["天数"]
        r = v.diff() if _cr_is_pct(f, t) else v.pct_change(fill_method=None).replace([np.inf, -np.inf], np.nan) * 100
        # 逐日用真实日期轴（缺的天会留空，不会被挤掉）；周 / 月用期间标签
        x = list(g.index) if gran == "日" else [_cr_plabel(p, gran, d) for p, d in zip(g.index, g["天数"])]
        a_, b_ = st.columns([3, 2])
        fv = go.Figure(go.Bar(x=x, y=v, marker_color="#C2416B"))
        fv.update_layout(height=300, margin=dict(t=30, b=10),
                         title=f"{CR_FIELD_CN.get(f, f)}{'（日均）' if avg and not _cr_is_level(f) else ''}")
        a_.plotly_chart(fv, key=f"crg_v_{t}")
        fr = go.Figure(go.Bar(x=x, y=r, marker_color=["#2E9E7A" if (pd.notna(z) and z >= 0) else "#D9534F" for z in r]))
        fr.add_hline(y=0, line_color="#8A727C")
        fr.update_layout(height=300, margin=dict(t=30, b=10),
                         title=f"环比（vs 上一期，{'pp' if _cr_is_pct(f, t) else '%'}）")
        b_.plotly_chart(fr, key=f"crg_r_{t}")
        return
    a0, a1, b0, b1 = win
    T = CRB["raw"][t]
    T = T[T["日期"].between(b0, a1)]
    if T.empty:
        st.caption("这两段里没有数据。")
        return
    s = cr_agg(T, t, T["日期"])[f]
    fig = go.Figure(go.Scatter(x=s.index, y=s.values, mode="lines+markers", line=dict(color="#C2416B", width=2)))
    for x0, x1, nm in ((b0, b1, "上一段"), (a0, a1, "本期")):
        fig.add_vrect(x0=x0 - pd.Timedelta(hours=12), x1=x1 + pd.Timedelta(hours=12), fillcolor="#8A727C",
                      opacity=.08 if nm == "上一段" else .16, line_width=0, annotation_text=nm,
                      annotation_position="top left")
    fig.update_layout(height=320, margin=dict(t=30, b=10), hovermode="x unified",
                      title=f"{CR_FIELD_CN.get(f, f)}：两段的逐日走势")
    st.plotly_chart(fig, key=f"crg_w_{t}")


def page_granular_creator():
    head = st.container()
    avail = [t for t in ("Creators", "Products", "Videos", "LIVE") if t in CRB["raw"]]
    last_by = {t: CRB["raw"][t]["日期"].max() for t in avail}
    common_last = min(last_by.values())
    c1, c2, c3 = st.columns([1.7, 1.1, 2.2])
    mode = c1.radio("对比方式", ["逐日", "逐周", "逐月", "自选区间"], horizontal=True, key="crg_mode")
    gran = {"逐日": "日", "逐周": "周", "逐月": "月"}.get(mode)
    if mode == "逐日":
        avg = False
        c2.caption("口径：当天数值")
    else:
        avg = c2.radio("口径", ["日均", "合计"], horizontal=True, index=0 if gran else 1,
                       key=f"crg_avg_{mode}",
                       help="周 / 月天数不同（首尾常常不满），默认按日均比；自选区间两段天数一样，默认按合计比。"
                            "比率和去重个数两种口径下都一样。") == "日均"
    unit_note = "（日均）" if avg else ""

    frames, heads, tables, win = {}, {}, {}, None
    if gran:
        tables = {t: cr_period_table(t, gran) for t in avail}
        periods = sorted(set().union(*[set(g.index) for g in tables.values()]), reverse=True)
        common = sorted(set.intersection(*[set(g.index) for g in tables.values()]), reverse=True)
        full = [p for p in common if gran != "周" or all(tables[t].at[p, "天数"] >= 7 for t in avail)]
        default = full[0] if full else (common[0] if common else periods[0])
        days_of = {}
        for g in tables.values():
            for p, d in g["天数"].items():
                days_of[p] = max(days_of.get(p, 0), d)
        pick = c3.selectbox("本期（跟它的上一期比）", periods, index=periods.index(default), key=f"crg_pick_{gran}",
                            format_func=lambda p: _cr_plabel(p, gran, days_of.get(p)))
        lbls = {}
        for t in avail:
            g = tables[t]
            if pick not in g.index:
                heads[t] = f"这张表没有这一期的数据（最新到 {last_by[t]:%m-%d}）。"
                continue
            i = list(g.index).index(pick)
            if i == 0:
                heads[t] = "这是这张表的第一期，没有上一期可比。"
                continue
            cur, prev = g.iloc[i], g.iloc[i - 1]
            lbls[t] = (_cr_plabel(g.index[i], gran, cur["天数"]),
                       f"上一期 {_cr_plabel(g.index[i - 1], gran, prev['天数'])}")
            heads[t] = (f"本期 **{_cr_plabel(g.index[i], gran, cur['天数'])}**（{cur['来源']}，{int(cur['天数'])} 天）"
                        f" vs 上一期 **{_cr_plabel(g.index[i - 1], gran, prev['天数'])}**（{prev['来源']}，{int(prev['天数'])} 天）")
            frames[t] = cr_compare_frame(t, cur, prev, avg, gran)
    else:
        d_first = min(CRB["raw"][t]["日期"].min() for t in avail)
        rng = c3.date_input("本期（选一段日期，自动跟前面同样天数的一段比）",
                            value=((common_last - pd.Timedelta(days=6)).date(), common_last.date()),
                            min_value=d_first.date(), max_value=max(last_by.values()).date(), key="crg_rng")
        if not (isinstance(rng, tuple) and len(rng) == 2):
            st.info("请把本期的起止日期选完整。")
            return
        a0, a1 = pd.Timestamp(rng[0]), pd.Timestamp(rng[1])
        n = (a1 - a0).days + 1
        b1 = a0 - pd.Timedelta(days=1)
        b0 = b1 - pd.Timedelta(days=n - 1)
        win = (a0, a1, b0, b1)
        lbls = {t: (f"本期（{a0:%m-%d}～{a1:%m-%d}）", f"上一段（{b0:%m-%d}～{b1:%m-%d}）") for t in avail}
        st.markdown(f"本期 **{a0:%Y/%m/%d} – {a1:%m/%d}**（{n} 天） vs 上一段 **{b0:%Y/%m/%d} – {b1:%m/%d}**"
                    f"（紧挨着本期、同样 {n} 天）")
        for t in avail:
            w = cr_window_table(t, a0, a1, b0, b1)
            dc = int(w.at["本期", "天数"]) if "天数" in w and pd.notna(w.at["本期", "天数"]) else 0
            db = int(w.at["上一段", "天数"]) if "天数" in w and pd.notna(w.at["上一段", "天数"]) else 0
            if not dc or not db:
                heads[t] = (f"这张表在{'本期' if not dc else '上一段'}没有数据（这张表的数据范围 "
                            f"{CRB['raw'][t]['日期'].min():%m-%d} → {last_by[t]:%m-%d}）。")
                continue
            heads[t] = f"本期有数据 {dc}/{n} 天，上一段 {db}/{n} 天" + \
                ("——⚠️ 两段有数据的天数不同，建议口径选「日均」" if dc != db and not avg else "")
            frames[t] = cr_compare_frame(t, w.loc["本期"], w.loc["上一段"], avg, "区间")

    if max(last_by.values()) > common_last:
        st.caption("各表更新到的日期：" + "　".join(f"{t} {d:%m-%d}" for t, d in last_by.items()))

    # 每张表的关键字段：曝光（索引）→ 点击 → CTR → 订单 → GMV
    cards, pills = [], [("info", f"{mode} · {'日均' if avg else '合计'}口径")]
    for t in avail:
        if t not in frames:
            continue
        cf = frames[t].set_index("字段")
        cl, bl = lbls[t]
        f_im = CR_TABLES[t][7].get("CTR")
        spec = [("曝光", "impr", f_im), ("点击", "clicks", CR_TABLES[t][6]), ("CTR", "CTR 点击率", "CTR"),
                ("订单", "sku", CR_TABLES[t][5]), ("GMV", "gmv", CR_TABLES[t][4])]
        lines = [say_change(nm + ("" if kk in DERIVED else unit_note), kk, cf.at[f, "_cur"], cf.at[f, "_prev"], cl, bl)
                 for nm, kk, f in spec if f and f in cf.index]
        p_im = cf.at[f_im, "_rel"] if f_im in cf.index else np.nan
        p_g = cf.at[CR_TABLES[t][4], "_rel"] if CR_TABLES[t][4] in cf.index else np.nan
        cards.append({"level": level_of(p_im), "title": f"{t} · {CR_TABLES[t][0]}",
                      "tag": f"曝光 {p_im:+.1f}%" if pd.notna(p_im) else "曝光 —", "body": lines})
        if t in ("Creators", "Products") and pd.notna(p_im):
            pills.append((level_of(p_im), f"{t} 曝光 {p_im:+.1f}%"))
            if pd.notna(p_g):
                pills.append((level_of(p_g), f"{t} GMV {p_g:+.1f}%"))
    with head:
        page_header("多粒度对比", "达人组上传的 Creators / Products / Videos / LIVE 四个 Excel，每个字段逐日 / 逐周 / 逐月 / 自选区间环比。"
                    "比率类（CTR、完播率…）按曝光、播放加权重算；「表里出现的 / 出单 X 数」是这一期里的去重个数。",
                    ["达人数据"] + [f"{t} 到 {d:%m/%d}" for t, d in last_by.items()])
        summary_pills(pills)

    tabs = st.tabs(["结论"] + [f"{t} · {CR_TABLES[t][0]}" + (f"（{len(frames[t])} 个字段）" if t in frames else "")
                               for t in avail])
    with tabs[0]:
        if cards:
            section_title("每张表的关键字段", "先看曝光（索引列），再看点击、CTR、订单、GMV")
            insight_cards(cards, cols=2)
        else:
            st.info("所选期间没有可比的数据。")
        section_title(f"变化最大的字段{unit_note}", "四张表合看金额 / 计数类；环比 ≥ 5%，基数太小的不算；比率类见各表")
        up, down = _cr_movers(frames)
        ca, cb = st.columns(2)
        for cc, tb, ttl in ((ca, up, "上升最多"), (cb, down, "下降最多")):
            cc.markdown(f"**{ttl}**")
            if tb.empty:
                cc.caption("没有。")
            else:
                cc.dataframe(_style_rows(tb.drop(columns="_rel"), tb["_rel"], ["环比%"]), hide_index=True)

    for tab, t in zip(tabs[1:], avail):
        with tab:
            st.markdown(heads.get(t, ""))
            if t in frames:
                cf = frames[t]
                n_up, n_dn = int((cf["_rel"] >= 0.05).sum()), int((cf["_rel"] <= -0.05).sum())
                st.caption(f"{n_up} 个字段上升、{n_dn} 个下降（绿 = 上升，红 = 下降，只表示方向——退款、佣金上升不是好事）。"
                           + ("金额和计数是日均值。" if avg else ""))
                disp = cf.drop(columns=["_rel", "_prev", "_cur"])
                st.dataframe(_style_rows(disp, cf["_rel"], ["变化", "环比%"]), hide_index=True,
                             height=min(1000, 40 + 35 * len(disp)))
            if gran:
                st.markdown(f"**每一期 × 全部字段**{unit_note}（最新在上）")
                _cr_matrix(t, tables[t], gran, avg)
            fields = (frames[t]["字段"].tolist() if t in frames
                      else [c for c in tables[t].columns if c not in ("天数", "来源")] if gran else [])
            if fields:
                _cr_field_trend(t, fields, gran, avg, tables.get(t), win)


# ── 达人独有：达人 / 内容排行 ───────────────────────────────────────────
RANK_DIMS = {
    "达人": ("达人", ["曝光", "点击", "GMV", "订单", "件数", "播放", "新发视频", "新开直播", "寄样", "样品内容", "加橱窗",
                    "买家", "佣金", "退款"]),
    "视频": ("内容ID", ["曝光", "点击", "GMV", "订单", "播放/观众", "赞", "评", "转"]),
    "直播": ("内容ID", ["曝光", "点击", "GMV", "订单", "播放/观众", "赞", "评", "转"]),
}
_RANK_KEY = {"曝光": "impr", "点击": "clicks", "GMV": "gmv", "订单": "sku", "件数": "items", "播放": "v_views",
             "新发视频": "new_vid", "新开直播": "new_live", "佣金": "comm", "退款": "refund", "买家": "cust"}


def _rank_multi(frame: pd.DataFrame, key: str, cd, bd, names: pd.Series, sort_col: str, topn: int = 30) -> pd.DataFrame:
    """达人 / 内容的全指标两段对比：曝光 → 点击 → CTR → 订单 → GMV（按 sort_col 的变化排）。"""
    cols = [c for c in ("曝光", "点击", "播放", "播放/观众", "订单", "GMV") if c in frame.columns]
    parts = {c: _dim_delta(frame, key, c, cd, bd) for c in cols}
    idx = pd.Index([])
    for p in parts.values():
        idx = idx.union(p.index)
    t = pd.DataFrame(index=idx)
    for c in cols:
        t[f"_{c}_b"] = parts[c]["基期"].reindex(idx).fillna(0.0)
        t[f"_{c}_c"] = parts[c]["本期"].reindex(idx).fillna(0.0)
        t[f"Δ{c}"] = t[f"_{c}_c"] - t[f"_{c}_b"]
    sk = f"Δ{sort_col}" if f"Δ{sort_col}" in t else f"Δ{cols[0]}"
    t = pd.concat([t.sort_values(sk).head(topn), t.sort_values(sk, ascending=False).head(topn // 2)])
    t = t[~t.index.duplicated()]
    ctr = lambda cl, im: cl / im * 100 if im else np.nan  # noqa: E731
    out = pd.DataFrame({"名称": names.reindex(t.index).fillna(pd.Series(t.index.astype(str), index=t.index)).values})
    if "曝光" in cols:
        out["曝光 前→今"] = [f"{fv('impr', b)} → {fv('impr', c)}" for b, c in zip(t["_曝光_b"], t["_曝光_c"])]
        out["Δ曝光"] = [fdelta("impr", v) for v in t["Δ曝光"]]
    if "点击" in cols:
        out["Δ点击"] = [fdelta("clicks", v) for v in t["Δ点击"]]
        if "曝光" in cols:
            out["CTR 前→今"] = [f"{fv('CTR 点击率', ctr(a, b))} → {fv('CTR 点击率', ctr(c, d))}"
                                for a, b, c, d in zip(t["_点击_b"], t["_曝光_b"], t["_点击_c"], t["_曝光_c"])]
    for c in ("播放", "播放/观众"):
        if c in cols:
            out[f"Δ{c}"] = [f"{v:+,.0f}" for v in t[f"Δ{c}"]]
    if "订单" in cols:
        out["Δ订单"] = [fdelta("sku", v) for v in t["Δ订单"]]
    if "GMV" in cols:
        out["GMV 前→今"] = [f"{fv('gmv', b)} → {fv('gmv', c)}" for b, c in zip(t["_GMV_b"], t["_GMV_c"])]
        out["ΔGMV"] = [fdelta("gmv", v) for v in t["ΔGMV"]]
    return out


def page_rank():
    head = st.container()
    if PROF["key"] != "cr":
        with head:
            page_header("达人 / 内容排行", "这一页只针对达人数据", ["运营数据"])
        st.info("这一页只针对达人数据：侧边栏最上面切到「达人」。")
        return
    D = df_all
    d_first, d_last = D["日期"].min(), D["日期"].max()
    c0, c3 = st.columns([2, 1])
    dim = c0.radio("维度", list(RANK_DIMS), key="rk_dim", horizontal=True)
    key, mets = RANK_DIMS[dim]
    val = c3.selectbox("排序指标（默认曝光）", mets, key=f"rk_val_{dim}")
    picked = _pick_two_periods(d_first, d_last, "rk")
    if picked is None:
        return
    a0, a1, b0, b1 = picked
    avail = _ts_list(D["日期"].unique())
    cd = [d for d in avail if a0 <= d <= a1]
    bd = [d for d in avail if b0 <= d <= b1]
    frame = CRB["creator_day"] if dim == "达人" else CRB["content"][CRB["content"]["类型"] == dim]
    t = _dim_delta(frame, key, val, cd, bd) if cd and bd else pd.DataFrame()
    if t.empty:
        with head:
            page_header("达人 / 内容排行", "谁在卖、哪条视频 / 直播在卖", ["达人数据"])
        st.warning("所选区间里没有数据。")
        return
    vk = _RANK_KEY.get(val, "")
    fmt = (lambda v: fv(vk, v)) if vk else (lambda v: f"{v:,.0f}")
    sgn = (lambda v: fdelta(vk, v)) if vk else (lambda v: f"{v:+,.0f}")
    names = _link_short_names()
    if dim == "达人":
        t["说明"] = _main_link_of(t.index, cd + bd).reindex(t.index).fillna("—").map(lambda v: f"主推：{v}").values
        t["名称"] = "@" + t.index.astype(str)
        name_all = pd.Series("@" + frame["达人"].astype(str).unique(), index=frame["达人"].unique())
    else:
        meta = (frame[frame["日期"].isin(cd + bd)].drop_duplicates("内容ID", keep="last")
                .set_index("内容ID")[["达人", "标题", "挂车链接", "发布时间"]])
        t = t.join(meta)
        t["名称"] = ("@" + t["达人"].astype(str) + " · " + t["标题"].astype(str).str.slice(0, 20)
                    + " ·" + t.index.astype(str).str[-4:])
        t["说明"] = t["挂车链接"].map(lambda x: "挂车：" + _anchor_names(x, names))
        name_all = ("@" + meta["达人"].astype(str) + " · " + meta["标题"].astype(str).str.slice(0, 20)
                    + " ·" + meta.index.astype(str).str[-4:])

    cur_tot, base_tot = t["本期"].sum(), t["基期"].sum()
    n_cur, n_base = int((t["本期"] > 0).sum()), int((t["基期"] > 0).sum())
    top10 = lambda s: s.nlargest(10).sum() / s.sum() * 100 if s.sum() else np.nan  # noqa: E731
    new, lost = t[(t["基期"] == 0) & (t["本期"] > 0)], t[(t["基期"] > 0) & (t["本期"] == 0)]
    cl, bl = f"本期（{a0:%m-%d}～{a1:%m-%d}）", f"基期（{b0:%m-%d}～{b1:%m-%d}，按天数折算）"
    # 合计的曝光 / GMV（不管排序指标是什么，曝光都先报）
    tot = {}
    for c in ("曝光", "点击", "GMV"):
        if c in frame.columns:
            tt = _dim_delta(frame, key, c, cd, bd)
            tot[c] = (tt["本期"].sum(), tt["基期"].sum())
    p_val = pct(cur_tot, base_tot)
    with head:
        page_header("达人 / 内容排行", f"谁在卖、哪条视频 / 直播在卖：按「{val}」排，带本期 vs 基期和贡献度。"
                    f"每个{dim}的 Δ 加起来 = 合计的变化。",
                    ["达人数据", f"维度：{dim}", f"本期 {a0:%m/%d}–{a1:%m/%d}", f"基期 {b0:%m/%d}–{b1:%m/%d}"])
        pills = [("info", f"{dim} · 按{val}")]
        if "曝光" in tot:
            pi = pct(*tot["曝光"])
            pills.append((level_of(pi), f"合计曝光 {pi:+.1f}%" if pd.notna(pi) else "合计曝光 —"))
        if val != "曝光":
            pills.append((level_of(p_val), f"合计{val} {p_val:+.1f}%" if pd.notna(p_val) else f"合计{val} —"))
        pills += [("info", f"有{val}的{dim} {n_cur:,} 个（基期 {n_base:,}）"),
                  ("up" if len(new) else "flat", f"新起量 {len(new)}"), ("down" if len(lost) else "flat", f"掉线 {len(lost)}")]
        summary_pills(pills)

    tabs = st.tabs(["结论", "排行", "谁拉动 / 拖累", "全指标对比", "集中度 · 新起 · 掉线"])
    with tabs[0]:
        tot_lines = [say_change(f"{dim}合计{c}", {"曝光": "impr", "点击": "clicks", "GMV": "gmv"}[c], a, b, cl, bl)
                     for c, (a, b) in tot.items()]
        if val not in tot:
            tot_lines.append(say_change(f"{dim}合计{val}", vk, cur_tot, base_tot, cl, bl) if vk else
                             f"{dim}合计{val}：本期 {cur_tot:,.0f}，基期 {base_tot:,.0f}（{p_val:+.1f}%）。")
        dn, up_ = t[t["Δ"] < 0].head(3), t[t["Δ"] > 0].tail(3).iloc[::-1]

        def who(x):
            return [f"{r['名称']}：{fmt(r['基期'])} → {fmt(r['本期'])}（{sgn(r['Δ'])}"
                    + (f"，{r['自身变化%']:+.0f}%" if pd.notna(r["自身变化%"]) else "，新出现") + f"）。{r['说明']}"
                    for _, r in x.iterrows()]
        t10c, t10b = top10(t["本期"]), top10(t["基期"])
        cards = [{"level": level_of(pct(*tot["曝光"])) if "曝光" in tot else level_of(p_val),
                  "title": f"{dim}合计", "tag": "曝光优先", "body": tot_lines},
                 {"level": "down", "title": f"{val}拖累最多的{dim}", "tag": f"合计 {sgn(t.loc[t['Δ'] < 0, 'Δ'].sum())}",
                  "body": who(dn) or ["没有下降的。"]},
                 {"level": "up", "title": f"{val}拉动最多的{dim}", "tag": f"合计 {sgn(t.loc[t['Δ'] > 0, 'Δ'].sum())}",
                  "body": who(up_) or ["没有上升的。"]},
                 {"level": "warn" if pd.notna(t10c) and pd.notna(t10b) and t10c - t10b >= 5 else "info",
                  "title": "集中度与新老更替", "tag": f"{val}",
                  "body": [f"前 10 个{dim}占本期{val}的 {t10c:.1f}%（基期 {t10b:.1f}%）。" if pd.notna(t10c) and pd.notna(t10b) else "",
                           f"新起量 {len(new)} 个（合计 {sgn(new['Δ'].sum())}），掉线 {len(lost)} 个（合计 {sgn(lost['Δ'].sum())}）。"]}]
        insight_cards(cards, cols=2)

    with tabs[1]:
        section_title(f"本期排行 Top 20（{val}）")
        top = t.sort_values("本期", ascending=False).head(20)
        fb = go.Figure(go.Bar(x=top["本期"], y=top["名称"], orientation="h", marker_color="#C2416B",
                              text=[fmt(v) for v in top["本期"]], textposition="outside"))
        fb.update_layout(height=max(360, 26 * len(top) + 60), margin=dict(t=10, b=10),
                         yaxis=dict(autorange="reversed", automargin=True))
        st.plotly_chart(fb, key="rk_top")
        show = _fmt_dim(top, val)
        show["本期占比%"] = top["本期"].map(lambda v: f"{v / cur_tot * 100:.1f}%" if cur_tot else "—").values
        st.dataframe(show[["名称", "说明", "本期", "本期占比%", "基期（已折算）", "Δ", "自身变化%"]], hide_index=True)

    with tabs[2]:
        section_title(f"谁拉动、谁拖累（Δ{val}）", f"所有{dim}的 Δ 相加 = {sgn(t['Δ'].sum())}")
        st.plotly_chart(contrib_bars(t, "Δ", xtitle=f"Δ{val}（本期 − 基期）", label_col="名称"), key="rk_contrib")
        mv = pd.concat([t[t["Δ"] < 0].head(15), t[t["Δ"] > 0].tail(10).iloc[::-1]])
        st.dataframe(_fmt_dim(mv, val)[["名称", "说明", "基期（已折算）", "本期", "Δ", "自身变化%", "占跌量%", "占涨量%"]],
                     hide_index=True)

    with tabs[3]:
        sc_ = val if val in ("曝光", "点击", "播放", "播放/观众", "订单", "GMV") else "曝光"
        section_title(f"每个{dim}的全部指标", f"按 Δ{sc_} 排：跌得最多的 30 个 + 涨得最多的 15 个；基期按天数折算")
        st.dataframe(_rank_multi(frame, key, cd, bd, name_all, sc_), hide_index=True, height=560)

    with tabs[4]:
        section_title(f"前 N 个{dim}占了多少（本期 vs 基期）")
        fc = go.Figure()
        for lbl_, col_, color in (("本期", "本期", "#C2416B"), ("基期", "基期", "#8A727C")):
            s_ = t[col_].sort_values(ascending=False)
            s_ = s_[s_ > 0].head(50)
            if s_.sum():
                fc.add_trace(go.Scatter(x=list(range(1, len(s_) + 1)), y=s_.cumsum() / t[col_].sum() * 100,
                                        mode="lines+markers", name=lbl_, line=dict(color=color, width=2)))
        fc.update_layout(height=300, margin=dict(t=10, b=10), hovermode="x unified",
                         xaxis_title=f"前 N 个{dim}", yaxis_title=f"累计占 {val} %")
        st.plotly_chart(fc, key="rk_conc")
        a_, b_ = st.columns(2)
        with a_:
            st.markdown(f"**新起量**（基期没有、本期有）：{len(new)} 个，合计 {sgn(new['Δ'].sum())}")
            st.dataframe(_fmt_dim(new.sort_values("本期", ascending=False).head(10), val)[["名称", "说明", "本期"]],
                         hide_index=True)
        with b_:
            st.markdown(f"**掉线**（基期有、本期没有）：{len(lost)} 个，合计 {sgn(lost['Δ'].sum())}")
            st.dataframe(_fmt_dim(lost.sort_values("基期").head(10), val)[["名称", "说明", "基期（已折算）"]],
                         hide_index=True)
    st.download_button(f"⬇️ 导出{dim}排行全表 CSV", t.reset_index().to_csv(index=False).encode("utf-8-sig"),
                       f"达人组_{dim}_{val}_排行.csv", "text/csv", key="rk_dl")


# ──────────────────────────────────────────────────────────────────────
# 6c. 问问小猫：右下角小猫头 → 对话框。Claude（Sonnet 5）听懂问题 → 调下面这些「查数工具」
#     （都是程序里现成的分析函数）在程序内部真跑一遍 → 用大白话回答。
#     凭证：ANTHROPIC_API_KEY + NV_CHAT_PASSWORD（Streamlit Secrets 或本地 .streamlit/secrets.toml）。
# ──────────────────────────────────────────────────────────────────────
CHAT_MODEL = "claude-sonnet-5"                 # 她选的：Sonnet 5（便宜、快）
CHAT_EFFORT = "medium"
CHAT_MAX_ROUNDS = 8                            # 一个问题最多查几轮数
CHAT_MAX_QUESTIONS = 60                        # 每个浏览器会话最多问多少题（防止误刷）
CHAT_KEEP_TURNS = 10                           # 多轮对话最多带最近几问的上下文
CHAT_PRICE = {"in": 2.0, "out": 10.0, "cache_w": 2.5, "cache_r": 0.2}   # Sonnet 5 $/百万 token

_CHAT_METRICS = {"曝光": "impr", "点击": "clicks", "CTR": "CTR 点击率", "CTOR": "CTOR 点击成单率",
                 "订单": "sku", "GMV": "gmv", "客单价": "AOV 客单价", "件数": "items"}


def _chat_metric(name) -> str | None:
    """「曝光」「CTR」「GMV」… → K / DERIVED 的键；认 METRIC_PICK 里的所有名字。"""
    s = str(name or "").strip()
    if not s:
        return None
    for m in (_CHAT_METRICS, METRIC_PICK):
        if s in m:
            return m[s]
    if s in K or s in DERIVED:
        return s
    low = s.lower()
    for m in (_CHAT_METRICS, METRIC_PICK):
        for k, v in m.items():
            if low in k.lower() or k.lower() in low:
                return v
    return None


def _chat_date(s, default=None) -> pd.Timestamp | None:
    """'2026-09-08' / '09-08' / '9/8' → Timestamp（没写年份用数据最新日期的年份）。"""
    if s is None or str(s).strip() == "":
        return default
    t = str(s).strip().replace("/", "-").replace(".", "-")
    if re.fullmatch(r"\d{1,2}-\d{1,2}", t):
        t = f"{df_all['日期'].max().year}-{t}"
    return pd.Timestamp(t).normalize()


def _chat_names() -> pd.Series:
    return df_all.groupby("product_id")["product_name"].last().map(lambda s: short_name(s) or "(无名称)")


def _chat_link(q) -> tuple[str | None, str]:
    """链接关键词 / Product ID → (product_id, 名称)；找不到或有歧义时返回 (None, 说明)。"""
    s = str(q or "").strip()
    if not s:
        return None, "没给链接"
    names = _chat_names()
    if s in names.index:
        return s, names[s]
    hit = [p for p, n in names.items() if s.upper() in str(n).upper() or str(p).endswith(s)]
    if not hit:
        return None, f"没找到名称含「{s}」的链接。可以先用 find_links 搜一下。"
    imp = df_all[df_all["product_id"].isin(hit)].groupby("product_id")[K["impr"]].sum()
    pid = imp.idxmax() if len(imp) else hit[0]
    return pid, names[pid]


def _chat_window(start, end, days: int = 7):
    """本期：没给就用数据最后 days 天；返回 (a0, a1)。"""
    last = df_all["日期"].max()
    a1 = min(_chat_date(end, last), last)
    a0 = _chat_date(start, a1 - pd.Timedelta(days=days - 1))
    if a0 > a1:
        a0, a1 = a1, a0
    return a0, a1


def _chat_base(a0, a1, base_start, base_end):
    if base_start and base_end:
        return _chat_date(base_start), _chat_date(base_end)
    n = (a1 - a0).days + 1
    b1 = a0 - pd.Timedelta(days=1)
    return b1 - pd.Timedelta(days=n - 1), b1


def _chat_table(t: pd.DataFrame, max_rows: int = 15) -> str:
    if t is None or len(t) == 0:
        return "（无）"
    return t.head(max_rows).to_string(index=False, max_colwidth=40)


def _chat_clip(s: str, n: int = 7000) -> str:
    return s if len(s) <= n else s[:n] + "\n…（结果太长，后面截掉了）"


# ── 工具 ① 数据概况
def tool_data_info(**_) -> str:
    D = df_all
    d0, d1 = D["日期"].min(), D["日期"].max()
    lines = [f"当前数据源：{PROF['name']}数据；日期 {d0:%Y-%m-%d} → {d1:%Y-%m-%d}（{D['日期'].nunique()} 天）；"
             f"链接 {D['product_id'].nunique()} 个；渠道：" + "、".join(l for l, _ in CHANNELS) + "。",
             "可问的指标：" + "、".join(METRIC_PICK.keys()) + "。"]
    if CRB is not None:
        cov = CRB["link"]["日期"]
        lines.append(f"达人组数据：{cov.min():%Y-%m-%d} → {cov.max():%Y-%m-%d}"
                     + ("（比运营日报更新慢，最近几天可能没覆盖）" if PROF["key"] == "ops" else "") + "。")
    lines.append(tool_compare_periods())
    return "\n".join(lines)


# ── 工具 ② 两段对比（全指标 + 渠道 + 链接 + 达人）
def tool_compare_periods(start=None, end=None, base_start=None, base_end=None, link=None, **_) -> str:
    a0, a1 = _chat_window(start, end)
    b0, b1 = _chat_base(a0, a1, base_start, base_end)
    D = df_all
    title = f"{SCOPE}"
    if link:
        pid, nm = _chat_link(link)
        if pid is None:
            return nm
        D, title = D[D["product_id"] == pid], f"链接「{nm}」"
    cm, bm = D["日期"].between(a0, a1), D["日期"].between(b0, b1)
    na = df_all.loc[df_all["日期"].between(a0, a1), "日期"].nunique()
    nb = df_all.loc[df_all["日期"].between(b0, b1), "日期"].nunique()
    if not na or not nb:
        return (f"本期（{a0:%m-%d}～{a1:%m-%d}）或基期（{b0:%m-%d}～{b1:%m-%d}）没有数据；"
                f"数据范围 {df_all['日期'].min():%m-%d} → {df_all['日期'].max():%m-%d}。")
    sc, sb = sums_of(D[cm]), {k: v * na / nb for k, v in sums_of(D[bm]).items()}
    out = [f"【{title}·两段对比】本期 = {a0:%Y-%m-%d}～{a1:%m-%d}（{na} 天），基期 = {b0:%Y-%m-%d}～{b1:%m-%d}"
           f"（{nb} 天{'，已按天数折算成本期长度' if na != nb else ''}）"]
    for nm_, kk in _CHAT_METRICS.items():
        if kk in DERIVED or kk in sc:
            out.append("· " + say_change(nm_, kk, value_of(sc, kk), value_of(sb, kk), "本期", "基期"))
    for key, lab in (("impr", "曝光"), ("gmv", "GMV")):
        chm = CH_IMPR if key == "impr" else {l: k for l, k in CHANNELS}
        parts = [f"{l.split()[0]} {fdelta(key, sc.get(k, 0) - sb.get(k, 0))}" for l, k in chm.items() if k in sc]
        if parts:
            out.append(f"· {lab}按渠道变化：" + "、".join(parts))
    if not link:
        t = link_multi_attr(df_all, df_all["日期"].between(a0, a1), df_all["日期"].between(b0, b1), nb / na,
                            sort_key="impr", topn=8)
        out.append("\n按链接的全指标归因（按 Δ曝光 排，跌得最多的在前，后面是涨得最多的）：\n" +
                   _chat_table(t.drop(columns=[c for c in t.columns if c.startswith("_")], errors="ignore"), 14))
        gc, gb, names = link_sums(df_all, df_all["日期"].between(a0, a1), df_all["日期"].between(b0, b1), nb / na)
        tg = sum_contrib(gc, gb, names, "gmv")
        if len(tg):
            out.append("\nGMV 拖累最多：" + "；".join(f"{r['链接']} {fdelta('gmv', r['Δ'])}" for _, r in tg.head(3).iterrows())
                       + "\nGMV 拉动最多：" + "；".join(f"{r['链接']} {fdelta('gmv', r['Δ'])}"
                                                  for _, r in tg.iloc[::-1].head(3).iterrows()))
    if CRB is not None:
        cov = set(pd.to_datetime(CRB["creator_day"]["日期"].unique()))
        cd = [d for d in pd.date_range(a0, a1) if d in cov]
        bd = [d for d in pd.date_range(b0, b1) if d in cov]
        if len(cd) * 2 >= na and bd:
            if link:
                pid, _nm = _chat_link(link)
                out.append("\n达人侧（这条链）：" + cr_link_reason(pid, cd, bd))
            else:
                out.append(f"\n按达人的全指标对比（达人组 Creators 表，覆盖本期 {len(cd)}/{na} 天，按 Δ曝光 排）：\n" +
                           _chat_table(creator_multi_attr(cd, bd, "曝光", topn=6), 10))
        else:
            out.append(f"\n达人组数据只到 {max(cov):%m-%d}，本期覆盖不到一半，没法做达人侧归因。")
    return _chat_clip("\n".join(out))


# ── 工具 ③ 异常预警
def tool_anomalies(date=None, days=1, level=None, metric=None, direction=None, limit=15, **_) -> str:
    R = radar_all(7, 3.0, 30.0)
    if R.empty:
        return "全历史都没有预警。"
    d1 = _chat_date(date, df_all["日期"].max())
    days = max(1, min(int(days or 1), 60))
    d0 = d1 - pd.Timedelta(days=days - 1)
    x = R[R["日期"].between(d0, d1)]
    if level:
        x = x[x["层级"] == str(level)]
    if metric:
        x = x[x["指标"].str.upper() == str(metric).upper()]
    if direction in ("下跌", "上升"):
        x = x[x["方向"] == direction]
    head = (f"{d0:%m-%d}～{d1:%m-%d} 共 {len(x)} 条预警（判定：跟前 7 天中位数比偏离 ≥30% 且超出正常波动；"
            f"下跌 {int((x['方向'] == '下跌').sum())}、上升 {int((x['方向'] == '上升').sum())}；按曝光优先排）。")
    if x.empty:
        return head
    lines = [head]
    for i, (_, a) in enumerate(x.head(max(1, min(int(limit or 15), 30))).iterrows()):
        lines.append(f"· [{a['层级']}] {a['对象']} · {a['指标']}：{alert_headline(a, 7)}")
        if i < 4:
            lines += [f"    ↳ {r}" for r in alert_reason(a, 7)[:4]]
    return _chat_clip("\n".join(lines))


# ── 工具 ④ 单条链接报告
def tool_link_report(link, days=7, end=None, **_) -> str:
    pid, nm = _chat_link(link)
    if pid is None:
        return nm
    days = max(1, min(int(days or 7), 60))
    a0, a1 = _chat_window(None, end, days)
    b0, b1 = _chat_base(a0, a1, None, None)
    out = [tool_compare_periods(a0.strftime("%Y-%m-%d"), a1.strftime("%Y-%m-%d"), link=pid)]
    R = radar_all(7, 3.0, 30.0)
    al = R[(R["层级"] == "链接") & (R["obj"] == pid) & (R["日期"] > a1 - pd.Timedelta(days=14))]
    out.append(f"\n近 14 天这条链的预警 {len(al)} 条：" +
               ("".join(f"\n· {a['指标']}：{alert_headline(a, 7)}" for _, a in al.head(6).iterrows()) or "无"))
    ev = EV_ALL[(EV_ALL["product_id"] == pid) & (EV_ALL["事件"] != "上架/首次出现")]
    if len(ev):
        out.append("\n自动识别的改动：" + "；".join(f"{e['日期']:%m-%d} {e['事件']}（{e['详情']}）" for _, e in ev.tail(5).iterrows()))
    if CRB is not None and PROF["key"] == "cr":
        lc = CRB["lc"][CRB["lc"]["product_id"] == pid]
        t = _dim_delta(lc, "达人", "挂车曝光", list(pd.date_range(a0, a1)), list(pd.date_range(b0, b1)))
        if len(t):
            out.append("\n挂车曝光变化最大的达人：" + "；".join(
                f"@{n} {fv('impr', r['基期'])}→{fv('impr', r['本期'])}" for n, r in
                pd.concat([t.head(3), t.iloc[::-1].head(2)]).iterrows()))
    return _chat_clip("\n".join(out))


# ── 工具 ⑤ 走势（逐日 / 周 / 月）
def tool_trend(metric="曝光", link=None, start=None, end=None, granularity="日", **_) -> str:
    mk = _chat_metric(metric)
    if mk is None:
        return f"不认识指标「{metric}」。可用：" + "、".join(METRIC_PICK.keys())
    D, title = df_all, SCOPE
    if link:
        pid, nm = _chat_link(link)
        if pid is None:
            return nm
        D, title = D[D["product_id"] == pid], f"链接「{nm}」"
    a0 = _chat_date(start, D["日期"].min())
    a1 = _chat_date(end, D["日期"].max())
    D = D[D["日期"].between(a0, a1)]
    gran = granularity if granularity in ("日", "周", "月") else "日"
    g = aggregate(D, gran)
    if g.empty:
        return "这段没有数据。"
    col = mk if mk in g.columns else K.get(mk)
    if col not in g.columns:
        return f"{title}没有「{metric}」这个指标的数据。"
    v = g[col] if (mk in DERIVED or gran == "日") else g[col] / g["天数"]
    lab = f"{metric}" + ("" if (mk in DERIVED or gran == "日") else "（日均）")
    full_n = {"日": 1, "周": 7}
    rows = [f"{_plabel_any(p, gran)}：{fv(mk, x, 0)}"
            + (f"（vs 上一{'天' if gran == '日' else gran} {pct(x, y):+.1f}%）" if pd.notna(y) and y else "")
            + (f"〔这一{gran}只有 {int(n)} 天数据〕" if n < full_n.get(gran, pd.Timestamp(p).days_in_month) else "")
            for p, x, y, n in zip(g["期间"], v, v.shift(1), g["天数"])]
    rows = rows[-90:]
    return _chat_clip(f"【{title} · {lab} · 按{gran}】{a0:%m-%d} → {a1:%m-%d}\n" + "\n".join(rows))


# ── 工具 ⑥ 大跌与回暖
def tool_drops_recovery(metric="曝光", link=None, threshold=15, **_) -> str:
    mk = _chat_metric(metric) or "impr"
    pid, nm = (None, SCOPE)
    if link:
        pid, nm = _chat_link(link)
        if pid is None:
            return nm
    s = series_of(df_all, mk, pid)
    if s.empty:
        return "没有数据。"
    ch = daily_chain(s)
    thr = float(threshold or 15)
    drops = ch[ch["日环比%"] <= -thr]
    out = [f"【{nm} · {metric}】单日比前一天跌 ≥ {thr:.0f}% 的有 {len(drops)} 次（基线 = 跌前 7 天中位数）："]
    for _, dr in drops.tail(8).iterrows():
        r = recovery_after(s, dr["日期"], pre=7, post=14)
        if not r:
            continue
        out.append(f"· {dr['日期']:%m-%d}：{fv(mk, dr['前一日'])} → {fv(mk, dr['值'])}（{dr['日环比%']:+.0f}%）；"
                   f"跌前中位数 {fv(mk, r['baseline'])}；最低 {fv(mk, r['trough'])}（{r['trough_day']:%m-%d}）；"
                   f"到 {r['after']['日期'].iloc[-1]:%m-%d} 比跌前 {r['end_recovery']:+.0f}%；判定：{r['verdict']}"
                   + ("（跌前一天是冲高，多半是回落）" if r["after_spike"] else ""))
    last = ch.iloc[-1]
    out.append(f"最新一天 {last['日期']:%m-%d}：{fv(mk, last['值'])}，比前一天 "
               + ("—" if pd.isna(last["日环比%"]) else f"{last['日环比%']:+.1f}%") + "。")
    return _chat_clip("\n".join(out))


# ── 工具 ⑦ 改动效果
def tool_change_effects(link=None, **_) -> str:
    names = _chat_names()
    rows = []
    try:
        man = pd.read_csv(CHANGES_FILE, dtype=str).fillna("") if os.path.exists(CHANGES_FILE) \
            else pd.DataFrame(KNOWN_CHANGES)
    except Exception:                                       # noqa: BLE001
        man = pd.DataFrame(KNOWN_CHANGES)
    imp = df_all.groupby("product_id")[K["impr"]].sum()
    for _, r in man.iterrows():
        d_, kw = str(r.get("日期", "")).strip(), str(r.get("链接关键词", "")).strip()
        cands = [p for p, n in names.items() if kw and kw.upper() in str(n).upper()]
        if d_ and cands:
            p = max(cands, key=lambda x: imp.get(x, 0))
            rows.append(("你登记的", pd.Timestamp(d_), p, str(r.get("改动", ""))))
    for _, e in EV_ALL[EV_ALL["事件"].isin(["改名/重组", "改价/改促销（件单价突变）"])].iterrows():
        rows.append(("自动识别", e["日期"], e["product_id"], f"{e['事件']}：{e['详情']}"))
    if link:
        pid, nm = _chat_link(link)
        if pid is None:
            return nm
        rows = [r for r in rows if r[2] == pid]
    out = ["口径：改后 7 天日均 vs 改前 7 天日均；判定先看曝光（±10%），再看成交跟没跟上。"]
    mine = sorted([r for r in rows if r[0] == "你登记的"], key=lambda r: r[1], reverse=True)
    auto = sorted([r for r in rows if r[0] != "你登记的"], key=lambda r: r[1], reverse=True)
    n_auto = 0
    for src, d0, p, what in mine + auto:
        if src != "你登记的":
            if n_auto >= 10:
                break
        e = change_effect(df_all, p, d0)
        if not e:
            if src == "你登记的":
                out.append(f"· [{src}] {d0:%m-%d} {names.get(p, p)}（{what}）：前后数据不足")
            continue
        n_auto += src != "你登记的"
        out.append(f"· [{src}] {d0:%m-%d} {names.get(p, p)}（{what}）：曝光 {fv('impr', e['曝光改前'])}→{fv('impr', e['曝光改后'])}/天"
                   f"（{e['曝光变化%']:+.0f}%）；CTR {e['CTR改前']:.2f}%→{e['CTR改后']:.2f}%；CTOR {e['CTOR改前']:.2f}%→{e['CTOR改后']:.2f}%；"
                   f"GMV {fv('gmv', e['GMV改前'])}→{fv('gmv', e['GMV改后'])}/天"
                   + ("" if pd.isna(e["GMV变化%"]) else f"（{e['GMV变化%']:+.0f}%）") + f"；判定：{e['判定']}"
                   + ("（改前基数太小）" if e["基数太小"] else ""))
    return _chat_clip("\n".join(out)) if len(out) > 1 else "没有找到改动记录。"


# ── 工具 ⑧ 链接排行
def tool_rank_links(metric="曝光", start=None, end=None, top=10, sort="value", **_) -> str:
    mk = _chat_metric(metric)
    if mk is None:
        return f"不认识指标「{metric}」。"
    a0, a1 = _chat_window(start, end)
    b0, b1 = _chat_base(a0, a1, None, None)
    na = df_all.loc[df_all["日期"].between(a0, a1), "日期"].nunique()
    nb = df_all.loc[df_all["日期"].between(b0, b1), "日期"].nunique()
    if not na:
        return "本期没有数据。"
    gc, gb, names = link_sums(df_all, df_all["日期"].between(a0, a1), df_all["日期"].between(b0, b1),
                              (nb / na) if nb else 1.0)
    if mk in DERIVED:
        num, den, _f, _ = DERIVED[mk]
        mul = 100 if DERIVED[mk][2] == "pct" else 1
        cur = gc[K[num]] / gc[K[den]].replace(0, np.nan) * mul
        base = gb[K[num]] / gb[K[den]].replace(0, np.nan) * mul
        t = pd.DataFrame({"链接": names, "本期": cur, "基期": base, "分母": gc[K[den]]})
        t = t[t["分母"] >= (1000 if den == "impr" else 30)]            # 量太小的比率不排
    else:
        t = pd.DataFrame({"链接": names, "本期": gc[K[mk]], "基期": gb[K[mk]]})
    t["Δ"] = t["本期"] - t["基期"]
    t = t[(t["本期"].fillna(0) != 0) | (t["基期"].fillna(0) != 0)]
    key = {"value": ("本期", False), "drop": ("Δ", True), "rise": ("Δ", False)}.get(sort, ("本期", False))
    t = t.sort_values(key[0], ascending=key[1]).head(max(1, min(int(top or 10), 30)))
    show = pd.DataFrame({"链接": t["链接"], "本期": t["本期"].map(lambda v: fv(mk, v)),
                         "基期（折算）": t["基期"].map(lambda v: fv(mk, v)), "变化": t["Δ"].map(lambda v: fdelta(mk, v)),
                         "变化%": [("—" if pd.isna(b) or not b else f"{pct(c, b):+.0f}%") for c, b in zip(t["本期"], t["基期"])]})
    return (f"【链接排行 · {metric} · 本期 {a0:%m-%d}～{a1:%m-%d} vs 基期 {b0:%m-%d}～{b1:%m-%d}（按天数折算）· 排序={sort}】\n"
            + _chat_table(show, 30))


# ── 工具 ⑨ 达人 / 视频 / 直播排行（达人组数据）
def tool_rank_creators(dimension="达人", metric="曝光", start=None, end=None, top=10, sort="value", **_) -> str:
    if CRB is None:
        return "没有载入达人组数据。"
    dim = dimension if dimension in RANK_DIMS else "达人"
    key, mets = RANK_DIMS[dim]
    val = metric if metric in mets else "曝光"
    cov = CRB["creator_day"]["日期"]
    last = cov.max()
    a1 = min(_chat_date(end, last), last)
    a0 = _chat_date(start, a1 - pd.Timedelta(days=6))
    b0, b1 = _chat_base(a0, a1, None, None)
    avail = _ts_list(cov.unique())
    cd, bd = [d for d in avail if a0 <= d <= a1], [d for d in avail if b0 <= d <= b1]
    if not cd or not bd:
        return f"达人组数据范围 {cov.min():%m-%d} → {last:%m-%d}，所选时段没数据。"
    frame = CRB["creator_day"] if dim == "达人" else CRB["content"][CRB["content"]["类型"] == dim]
    t = _dim_delta(frame, key, val, cd, bd)
    if t.empty:
        return "两段都没有数据。"
    k2 = _RANK_KEY.get(val, "")
    if dim == "达人":
        t["名称"] = "@" + t.index.astype(str)
        t["说明"] = _main_link_of(t.index, cd + bd).reindex(t.index).fillna("—").map(lambda v: f"主推 {v}").values
    else:
        meta = (frame[frame["日期"].isin(cd + bd)].drop_duplicates("内容ID", keep="last").set_index("内容ID")
                [["达人", "标题", "挂车链接"]])
        t = t.join(meta)
        t["名称"] = "@" + t["达人"].astype(str) + " · " + t["标题"].astype(str).str.slice(0, 18)
        t["说明"] = t["挂车链接"].map(lambda x: "挂车 " + _anchor_names(x, _link_short_names()))
    key_ = {"value": ("本期", False), "drop": ("Δ", True), "rise": ("Δ", False)}.get(sort, ("本期", False))
    t = t.sort_values(key_[0], ascending=key_[1]).head(max(1, min(int(top or 10), 30)))
    f = (lambda v: fv(k2, v)) if k2 else (lambda v: f"{v:,.0f}")
    d = (lambda v: fdelta(k2, v)) if k2 else (lambda v: f"{v:+,.0f}")
    show = pd.DataFrame({"名称": t["名称"], "本期": t["本期"].map(f), "基期（折算）": t["基期"].map(f),
                         "变化": t["Δ"].map(d), "变化%": t["自身变化%"].map(lambda v: "新出现" if pd.isna(v) else f"{v:+.0f}%"),
                         "说明": t["说明"]})
    return (f"【{dim}排行 · {val} · 本期 {a0:%m-%d}～{a1:%m-%d} vs 基期 {b0:%m-%d}～{b1:%m-%d}（按天数折算）· 排序={sort}】\n"
            + _chat_table(show, 30))


# ── 工具 ⑩ 达人组四个 Excel 的字段环比
def tool_creator_fields(table="Creators", granularity="周", keyword=None, **_) -> str:
    if CRB is None:
        return "没有载入达人组数据。"
    t = table if table in CRB["raw"] else "Creators"
    gran = granularity if granularity in ("日", "周", "月") else "周"
    g = cr_period_table(t, gran)
    if len(g) < 2:
        return "期数不够，没法环比。"
    full = [p for p in g.index if gran != "周" or g.at[p, "天数"] >= 7]
    i = list(g.index).index(full[-1]) if full else len(g) - 1
    if i == 0:
        return "这是第一期，没有上一期。"
    cur, prev = g.iloc[i], g.iloc[i - 1]
    cf = cr_compare_frame(t, cur, prev, gran != "日", gran)
    if keyword:
        m = cf["字段"].str.contains(str(keyword), case=False) | cf["说明"].str.contains(str(keyword), case=False)
        cf = cf[m] if m.any() else cf
    show = cf[["字段", "说明", "上一期", "本期", "变化", "环比%"]]
    return _chat_clip(f"【{t} 表 · 本期 {_cr_plabel(g.index[i], gran, cur['天数'])} vs 上一期 "
                      f"{_cr_plabel(g.index[i - 1], gran, prev['天数'])}"
                      f"{'（金额、计数按日均）' if gran != '日' else ''}】\n" + _chat_table(show, 45))


# ── 工具 ⑪ 搜链接
def tool_find_links(keyword="", **_) -> str:
    names = _chat_names()
    kw = str(keyword or "").upper()
    hit = [p for p, n in names.items() if kw in str(n).upper() or kw in str(p)]
    if not hit:
        return f"没有名称含「{keyword}」的链接。"
    g = df_all[df_all["product_id"].isin(hit)].groupby("product_id").agg(i=(K["impr"], "sum"), g=(K["gmv"], "sum"),
                                                                           last=("日期", "max"))
    g = g.sort_values("i", ascending=False).head(15)
    return "\n".join(f"· {names[p]}（ID …{str(p)[-6:]}）：累计曝光 {fv('impr', r['i'])}，GMV {fv('gmv', r['g'])}，"
                     f"最后有数据 {r['last']:%m-%d}" for p, r in g.iterrows())


_S = {"type": "string"}
_I = {"type": "integer"}
_DATE = {"type": "string", "description": "日期，YYYY-MM-DD 或 MM-DD"}
CHAT_TOOLS = [
    ("data_info", tool_data_info, "当前数据源的概况：日期范围、链接数、渠道、可用指标、达人数据覆盖到哪天，以及最近 7 天 vs 前 7 天的全指标变化。"
     "不确定数据范围时先调它。", {}, []),
    ("compare_periods", tool_compare_periods,
     "两段对比·归因：本期 vs 基期（默认最近 7 天 vs 紧挨着的前 7 天，基期按天数折算）。返回曝光/点击/CTR/CTOR/订单/客单价/GMV/件数的变化、"
     "曝光和 GMV 按渠道拆、每条链接的全指标归因（谁拖累谁拉动）、有达人数据时按达人拆。可只看某一条链接。",
     {"start": _DATE, "end": _DATE, "base_start": _DATE, "base_end": _DATE,
      "link": {"type": "string", "description": "可选：链接名称关键词或 Product ID"}}, []),
    ("anomalies", tool_anomalies,
     "异常预警：某天（或截至某天的最近几天）所有层级（大盘/渠道/链接/达人）× 所有指标（曝光/点击/CTR/CTOR/订单/GMV）的异常，"
     "跟前 7 天中位数比，按曝光优先排；前几条附原因（掉在哪个渠道、哪几条链、哪几个达人）。",
     {"date": _DATE, "days": {"type": "integer", "description": "看截至 date 的最近几天，默认 1"},
      "level": {"type": "string", "enum": ["大盘", "渠道", "链接", "达人"]},
      "metric": {"type": "string", "enum": ["曝光", "点击", "CTR", "CTOR", "订单", "GMV", "播放"]},
      "direction": {"type": "string", "enum": ["下跌", "上升"]}, "limit": _I}, []),
    ("link_report", tool_link_report,
     "单条链接报告：最近 N 天 vs 前 N 天的全指标、渠道拆分、达人侧原因，近 14 天预警，自动识别的改动。",
     {"link": {"type": "string", "description": "链接名称关键词或 Product ID"}, "days": _I, "end": _DATE}, ["link"]),
    ("trend", tool_trend, "某个指标的走势（按日/周/月，每期带环比），可限定某条链接和日期范围。",
     {"metric": {"type": "string", "description": "曝光/点击/CTR/CTOR/订单/GMV/客单价/件数 等"},
      "link": _S, "start": _DATE, "end": _DATE, "granularity": {"type": "string", "enum": ["日", "周", "月"]}}, ["metric"]),
    ("drops_recovery", tool_drops_recovery, "大跌与回暖：单日比前一天跌超过阈值的日子，之后有没有回到跌前水平。",
     {"metric": _S, "link": _S, "threshold": {"type": "number", "description": "跌幅阈值 %，默认 15"}}, []),
    ("change_effects", tool_change_effects,
     "改动效果：用户登记的改动 + 自动识别的改名/改价，改后 7 天日均 vs 改前 7 天日均（曝光、CTR、CTOR、GMV 和判定）。",
     {"link": _S}, []),
    ("rank_links", tool_rank_links, "链接排行：某指标在本期的排名，或跌得最多 / 涨得最多（vs 前面同样天数）。",
     {"metric": _S, "start": _DATE, "end": _DATE, "top": _I,
      "sort": {"type": "string", "enum": ["value", "drop", "rise"]}}, []),
    ("rank_creators", tool_rank_creators, "达人组数据：达人 / 视频 / 直播排行（本期值，或跌得最多 / 涨得最多），带主推链接或挂车链接。",
     {"dimension": {"type": "string", "enum": ["达人", "视频", "直播"]}, "metric": _S, "start": _DATE, "end": _DATE,
      "top": _I, "sort": {"type": "string", "enum": ["value", "drop", "rise"]}}, []),
    ("creator_fields", tool_creator_fields,
     "达人组上传的 Excel（Creators/Products/Videos/LIVE）每个字段的最新一期 vs 上一期（日/周/月），可按关键词筛字段（如 完播、CTR、样品）。",
     {"table": {"type": "string", "enum": ["Creators", "Products", "Videos", "LIVE"]},
      "granularity": {"type": "string", "enum": ["日", "周", "月"]}, "keyword": _S}, []),
    ("find_links", tool_find_links, "按名称关键词搜链接（名称、ID、累计曝光和 GMV）。", {"keyword": _S}, ["keyword"]),
]
_CHAT_FUNCS = {n: f for n, f, *_ in CHAT_TOOLS}
_CHAT_SCHEMAS = {n: (props, req) for n, _f, _d, props, req in CHAT_TOOLS}


def _chat_tool_defs() -> list[dict]:
    defs = []
    for n, _f, desc, props, req in CHAT_TOOLS:
        if n in ("rank_creators", "creator_fields") and CRB is None:
            continue
        defs.append({"name": n, "description": desc, "eager_input_streaming": True,
                     "input_schema": {"type": "object", "properties": props, "required": req}})
    return defs


def _chat_valid(name: str, args) -> str | None:
    """eager 流式下服务器不校验工具参数，这里自己校验一遍：返回错误说明或 None。"""
    if name not in _CHAT_FUNCS:
        return f"没有叫 {name} 的工具"
    if not isinstance(args, dict):
        return "参数不是对象"
    props, req = _CHAT_SCHEMAS[name]
    for r in req:
        if r not in args:
            return f"缺少参数 {r}"
    for k, v in args.items():
        spec = props.get(k)
        if spec is None:
            return f"多了不认识的参数 {k}"
        want = spec.get("type")
        ok = (isinstance(v, str) if want == "string" else
              (isinstance(v, int) and not isinstance(v, bool)) if want == "integer" else
              (isinstance(v, (int, float)) and not isinstance(v, bool)) if want == "number" else True)
        if not ok:
            return f"参数 {k} 类型不对（要 {want}）"
        if "enum" in spec and v not in spec["enum"]:
            return f"参数 {k} 只能是 {spec['enum']}"
    return None


_CHAT_LABEL = {"data_info": "数据概况", "compare_periods": "两段对比·归因", "anomalies": "异常预警",
               "link_report": "链接下钻", "trend": "走势", "drops_recovery": "大跌与回暖", "change_effects": "改动效果",
               "rank_links": "链接排行", "rank_creators": "达人 / 内容排行", "creator_fields": "达人 Excel 字段环比",
               "find_links": "搜链接"}


def _chat_step_label(name: str, args: dict) -> str:
    bits = [str(v) for k, v in (args or {}).items() if v not in (None, "")]
    return f"{_CHAT_LABEL.get(name, name)}" + (f"（{'，'.join(bits)}）" if bits else "")


CHAT_SYSTEM = """你是「小猫」，NailVesta（穿戴甲品牌，在 TikTok Shop 卖货）链接数据深度分析程序里的问答助手。提问的人是店铺运营负责人，用中文回答。

你能用的数据都在工具里：工具会在程序内部真跑一遍分析并返回结果。每个数字都必须来自工具结果；工具没给的数不要估、不要编。问题需要数据时先调工具，可以连着调几个（比如先看两段对比，再看异常预警或某条链接）。

回答规则：
1. 判断好坏、有没有异常，先看曝光（曝光是索引列），再看点击、CTR、CTOR、订单、客单价、GMV。GMV 在涨但曝光在掉，要明确指出流量在掉。
2. 说变化要写清楚：哪个指标、跟哪段日期比、现在多少、差多少、百分之几。比率（CTR、CTOR）的变化写「个百分点」。
3. 单位：曝光写 IM，金额写 $，点击写「次」，订单写「单」。
4. 下跌要点名是哪个渠道、哪几条链接、哪几个达人（工具结果里有的话）。
5. 大白话。先一句话结论，再分几点说依据；一般 300 字以内，用户要详细再展开。可以用 markdown 小表格。
6. 问题说得含糊时，按最常见的理解去查（默认最近 7 天 vs 前 7 天），并在回答里说一句你是按什么口径查的。
7. 数据没覆盖到（比如达人组数据比运营日报更新慢）就直说，不下结论。
8. 需要的话，最后一句建议去程序里哪一页继续看：总览、两段对比·归因、多粒度对比、前端指标、后端指标、异常预警、链接下钻、日环比 & 回暖、改动效果溯源、指标归档、达人 / 内容排行（只有达人数据有）。
9. 只回答跟这个程序、这些数据、店铺运营有关的问题；别的问题礼貌说明你只管这个程序的数据。"""


def _chat_context() -> str:
    D = df_all
    s = (f"当前数据源：{PROF['name']}数据（侧边栏最上面可以切运营 / 达人）。数据日期 {D['日期'].min():%Y-%m-%d} → "
         f"{D['日期'].max():%Y-%m-%d}，{D['product_id'].nunique()} 个链接；渠道：" + "、".join(l for l, _ in CHANNELS) + "。")
    if CRB is not None:
        cov = CRB["link"]["日期"]
        s += f"达人组数据 {cov.min():%Y-%m-%d} → {cov.max():%Y-%m-%d}。"
    return s + f"今天是 {pd.Timestamp.today():%Y-%m-%d}。"


_ON_CLOUD = os.path.abspath(__file__).startswith("/mount/")     # Streamlit Community Cloud 的代码目录


def _chat_secret(key: str) -> str:
    return st.session_state.get(f"_nv_{key}") or _sec(key)


def _chat_cost(u: dict) -> float:
    return (u["in"] * CHAT_PRICE["in"] + u["out"] * CHAT_PRICE["out"] + u["cw"] * CHAT_PRICE["cache_w"]
            + u["cr"] * CHAT_PRICE["cache_r"]) / 1e6


def _chat_answer(question: str, api_msgs: list) -> tuple[str, list[str], dict, str]:
    """跑一个问题：查数进度（st.status）在上、回答流式输出在下，需要时调工具。返回 (回答, 查数步骤, 用量, 错误)。"""
    import anthropic
    client = anthropic.Anthropic(api_key=_chat_secret("ANTHROPIC_API_KEY"), max_retries=2)
    tools = _chat_tool_defs()
    system = [{"type": "text", "text": CHAT_SYSTEM}, {"type": "text", "text": _chat_context()}]
    msgs = api_msgs + [{"role": "user", "content": question}]
    steps: list[str] = []
    use = {"in": 0, "out": 0, "cw": 0, "cr": 0}
    status = st.status("小猫在想…", expanded=False)
    live = st.empty()
    json_retries = 0
    for _round in range(CHAT_MAX_ROUNDS):
        text = ""
        try:
            with client.messages.stream(model=CHAT_MODEL, max_tokens=64000, system=system, tools=tools,
                                        messages=msgs, thinking={"type": "adaptive"},
                                        output_config={"effort": CHAT_EFFORT},
                                        cache_control={"type": "ephemeral"}) as stream:
                for ev in stream:
                    if ev.type == "text":
                        text += ev.text
                        live.markdown(md(text) + " ▌")
                resp = stream.get_final_message()
            json_retries = 0
        except ValueError:                                  # 工具参数 JSON 完全解析不了：这一轮重来
            json_retries += 1
            if json_retries > 2:
                return "", steps, use, "查数参数出错了，换个问法再试一次。"
            continue
        except anthropic.AuthenticationError:
            return "", steps, use, "API key 不对或已失效：去 Streamlit Secrets（或本地 secrets.toml）检查 ANTHROPIC_API_KEY。"
        except anthropic.PermissionDeniedError:
            return "", steps, use, "这个 API key 没有权限用 Sonnet 5，去 console.anthropic.com 看一下账户。"
        except anthropic.RateLimitError:
            return "", steps, use, "请求太频繁（或额度用完了），等一分钟再问；也可能要去 console.anthropic.com 充值。"
        except anthropic.APIStatusError as e:
            return "", steps, use, f"Claude 服务返回错误（{e.status_code}），稍后再试。"
        except anthropic.APIConnectionError:
            return "", steps, use, "连不上 Claude 服务，检查一下网络。"
        u = resp.usage
        use["in"] += u.input_tokens or 0
        use["out"] += u.output_tokens or 0
        use["cw"] += getattr(u, "cache_creation_input_tokens", 0) or 0
        use["cr"] += getattr(u, "cache_read_input_tokens", 0) or 0
        tool_uses = [b for b in resp.content if b.type == "tool_use"]
        if resp.stop_reason == "refusal":
            return text, steps, use, "这个问题小猫没法回答，换个问法试试。"
        if not tool_uses:
            status.update(label=f"查了 {len(steps)} 项数据" if steps else "没有查数，直接回答", state="complete")
            live.markdown(md(text))
            return text, steps, use, ""
        if resp.stop_reason == "max_tokens":
            return text, steps, use, "回答太长被截断了，换个更具体的问法。"
        if text.strip():
            status.write(md(text.strip()))
        live.empty()
        results = []
        for b in tool_uses:
            err = _chat_valid(b.name, b.input)
            if err:
                results.append({"type": "tool_result", "tool_use_id": b.id, "is_error": True, "content": err})
                continue
            lab = _chat_step_label(b.name, b.input)
            steps.append(lab)
            status.update(label=f"小猫在查：{lab}")
            status.write(f"🔎 {lab}")
            try:
                out = _CHAT_FUNCS[b.name](**b.input)
                results.append({"type": "tool_result", "tool_use_id": b.id, "content": out or "（无结果）"})
            except Exception as e:                          # noqa: BLE001
                results.append({"type": "tool_result", "tool_use_id": b.id, "is_error": True,
                                "content": f"查数出错：{type(e).__name__}: {e}"})
        msgs = msgs + [{"role": "assistant", "content": resp.content}, {"role": "user", "content": results}]
    status.update(label="查数轮数到上限了", state="error")
    return "", steps, use, f"这个问题要查的东西太多（超过 {CHAT_MAX_ROUNDS} 轮），拆成几个小问题再问。"


def _chat_trim(api_msgs: list) -> list:
    """只带最近 CHAT_KEEP_TURNS 问的上下文（按「用户提问」切，工具调用和结果成对保留）。"""
    starts = [i for i, m in enumerate(api_msgs) if m["role"] == "user" and isinstance(m["content"], str)]
    return api_msgs[starts[-CHAT_KEEP_TURNS]:] if len(starts) > CHAT_KEEP_TURNS else api_msgs


def _chat_rerun():
    """只重跑对话框（不关掉它）；对话框是在整页重跑里打开的那一次不能这样重跑，就跳过（内容已经画出来了）。"""
    from streamlit.errors import StreamlitAPIException
    try:
        st.rerun(scope="fragment")
    except StreamlitAPIException:
        pass


_CHAT_EXAMPLES = {
    "ops": ["最近 7 天整体怎么样？曝光掉了吗？", "昨天有哪些异常？", "DreamWear 最近一周表现怎么样？",
            "8/24 修的那几条链接改完有用吗？", "上周曝光掉得最多的 5 条链接是哪些？", "达人渠道最近为什么在掉？"],
    "cr": ["最近 7 天达人合计怎么样？", "曝光掉得最多的达人是谁？", "哪条视频最近卖得最好？",
           "Creators 表的完播率周环比怎么样？", "LAST CHANCE 这条链是哪些达人在推？", "9/12 达人数据有什么异常？"],
}


@st.dialog("问问小猫", width="large", icon="🐱")
def chat_dialog():
    key = _chat_secret("ANTHROPIC_API_KEY")
    pwd = _chat_secret("NV_CHAT_PASSWORD")
    if not key or not pwd:
        st.markdown("**第一次用要先配两样东西**：Claude 的 API key（按量付费）和一个你自己定的问答口令"
                    "（公开链接上防止别人用你的额度）。")
        st.markdown("- **Streamlit Cloud 上**：App 右下角 ⋮ → Settings → Secrets，加两行：")
        st.code('ANTHROPIC_API_KEY = "sk-ant-…"\nNV_CHAT_PASSWORD = "你自己定的口令"', language="toml")
        if _ON_CLOUD:
            st.info("配好 Secrets 后刷新页面即可。（公开部署上不提供在这里填写，免得别人替你设口令。）")
            return
        st.markdown("- **本地电脑上**：直接填在下面，存进本地 `.streamlit/secrets.toml`（不会上传）。")
        with st.form("nv_chat_setup"):
            k_in = st.text_input("Anthropic API key", value="", type="password",
                                 placeholder="已配置" if key else "sk-ant-…（在 console.anthropic.com → API Keys 创建）")
            p_in = st.text_input("问答口令（自己定）", value="", type="password",
                                 placeholder="已配置" if pwd else "")
            if st.form_submit_button("💾 保存"):
                upd = {}
                if k_in.strip():
                    upd["ANTHROPIC_API_KEY"] = k_in.strip()
                if p_in.strip():
                    upd["NV_CHAT_PASSWORD"] = p_in.strip()
                if upd:
                    try:
                        _save_secrets(upd)
                    except Exception as e:                  # noqa: BLE001
                        st.error(f"保存失败：{e}")
                    for k_, v_ in upd.items():
                        st.session_state[f"_nv_{k_}"] = v_
                    st.session_state["_nv_chat_ok"] = "NV_CHAT_PASSWORD" in upd
                    _chat_rerun()
        return
    if not st.session_state.get("_nv_chat_ok"):
        import hmac
        with st.form("nv_chat_login"):
            p = st.text_input("问答口令", type="password", placeholder="输入你在 Secrets 里设的 NV_CHAT_PASSWORD")
            if st.form_submit_button("进入"):
                if hmac.compare_digest(p.encode(), pwd.encode()):
                    st.session_state["_nv_chat_ok"] = True
                    _chat_rerun()
                else:
                    st.error("口令不对。")
        return

    D = df_all
    store = st.session_state.setdefault("_nv_chat", {})
    conv = store.setdefault(PROF["key"], {"api": [], "ui": []})
    top = st.columns([5, 1])
    top[0].caption(f"现在问的是 **{PROF['name']}数据**（{D['日期'].min():%m/%d} → {D['日期'].max():%m/%d}）。"
                   "想问另一套数据，关掉对话框，用侧边栏最上面的开关切换。")
    if top[1].button("清空对话", key="nv_chat_clear", disabled=not conv["ui"]):
        conv["api"], conv["ui"] = [], []
        _chat_rerun()
    box = st.container(height=520, border=False)
    q = None
    with box:
        if not conv["ui"]:
            st.markdown("🐱 喵～我能帮你查这个程序里的数据：两段对比、异常预警、链接下钻、回暖、改动效果、排行……"
                        "用大白话问就行，我会在程序里跑一遍再回答。试试：")
            ex = _CHAT_EXAMPLES[PROF["key"]]
            cols = st.columns(2)
            for i, e in enumerate(ex):
                if cols[i % 2].button(e, key=f"nv_chat_ex_{i}", width="stretch"):
                    q = e
        for m in conv["ui"]:
            with st.chat_message(m["role"], avatar="🐱" if m["role"] == "assistant" else None):
                if m.get("steps"):
                    st.caption("🔎 查了：" + "；".join(m["steps"]))
                st.markdown(md(m["text"]) if m["role"] == "assistant" else m["text"])
                if m.get("err"):
                    st.warning(m["err"])
                if m.get("cost"):
                    st.caption(f"本题花费约 ${m['cost']:.3f}")
    typed = st.chat_input("用大白话问，比如：上周曝光为什么掉了？", key="nv_chat_in")
    q = typed or q
    if not q:
        return
    total = sum(sum(1 for m in c["ui"] if m["role"] == "user") for c in store.values())
    if total >= CHAT_MAX_QUESTIONS:
        with box:
            st.warning(f"这个会话已经问了 {total} 题（上限 {CHAT_MAX_QUESTIONS}），刷新页面再继续。")
        return
    with box:
        with st.chat_message("user"):
            st.markdown(q)
        with st.chat_message("assistant", avatar="🐱"):
            ans, steps, use, err = _chat_answer(q, _chat_trim(conv["api"]))
            if err:
                st.warning(err)
    conv["ui"].append({"role": "user", "text": q})
    conv["ui"].append({"role": "assistant", "text": ans or "", "steps": steps, "err": err, "cost": _chat_cost(use)})
    if ans and not err:
        conv["api"] += [{"role": "user", "content": q}, {"role": "assistant", "content": ans}]
    _chat_rerun()


def _cat_svg_uri() -> str:
    from urllib.parse import quote
    svg = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
           "<path d='M9 30 L8 5 L27 17 Z' fill='#C2416B'/><path d='M55 30 L56 5 L37 17 Z' fill='#C2416B'/>"
           "<path d='M13 24 L12.5 11 L22 18 Z' fill='#F6B5C8'/><path d='M51 24 L51.5 11 L42 18 Z' fill='#F6B5C8'/>"
           "<ellipse cx='32' cy='37' rx='25' ry='21' fill='#C2416B'/>"
           "<ellipse cx='23' cy='34' rx='4.2' ry='5' fill='#fff'/><ellipse cx='41' cy='34' rx='4.2' ry='5' fill='#fff'/>"
           "<ellipse cx='24' cy='35' rx='2.2' ry='3' fill='#2b1a22'/><ellipse cx='42' cy='35' rx='2.2' ry='3' fill='#2b1a22'/>"
           "<path d='M29.5 42 L34.5 42 L32 45 Z' fill='#F6B5C8'/>"
           "<path d='M32 45 Q29 48.5 26 46.5 M32 45 Q35 48.5 38 46.5' stroke='#fff' stroke-width='1.6' fill='none' "
           "stroke-linecap='round'/>"
           "<path d='M5 39 L17 41 M5 45 L17 44 M59 39 L47 41 M59 45 L47 44' stroke='#fff' stroke-width='1.3' "
           "stroke-linecap='round'/></svg>")
    return "data:image/svg+xml;utf8," + quote(svg)


def render_chat_cat():
    """每一页右下角的小猫头：点它打开「问问小猫」。"""
    st.html(f"""<style>
.st-key-nv_cat_btn {{ position: fixed; right: 28px; bottom: 26px; z-index: 999990; width: auto !important; }}
.st-key-nv_cat_btn button {{
  width: 68px; height: 68px; min-height: 68px; border-radius: 50%; padding: 0;
  background: #fff url("{_cat_svg_uri()}") center 58% / 54px 54px no-repeat;
  border: 2px solid #C2416B; color: transparent; box-shadow: 0 6px 18px rgba(194,65,107,.35);
  transition: transform .15s ease, box-shadow .15s ease; }}
.st-key-nv_cat_btn button:hover {{ transform: translateY(-2px) scale(1.06); box-shadow: 0 10px 24px rgba(194,65,107,.45);
  border-color: #C2416B; color: transparent; }}
.st-key-nv_cat_btn button:focus:not(:active) {{ border-color: #C2416B; color: transparent; }}
.st-key-nv_cat_btn button p {{ font-size: 0; }}
</style>""")
    if st.button("🐱", key="nv_cat_btn", help="问问小猫：用大白话问这个程序里的数据"):
        chat_dialog()


# ──────────────────────────────────────────────────────────────────────
# 7. 页面导航：侧边栏最上面是「运营 / 达人」开关，下面是同一套页面（开关切的是数据源，不是页面）。
#    st.navigation 设成 hidden，菜单自己画在 NAV_BOX 里，这样开关能放在菜单上面。
# ──────────────────────────────────────────────────────────────────────
PAGE_REFS: dict[str, "st.navigation.StreamlitPage"] = {}


def _reg(key: str, func, title: str, icon: str):
    p = st.Page(func, title=title, icon=icon, url_path=key)
    PAGE_REFS[key] = p
    return p


NAV_GROUPS = {
    "总览": [_reg("overview", page_overview, "总览", "🏠")],
    "深度分析": [_reg("compare", page_compare, "两段对比 · 归因", "🔍"),
               _reg("granular", page_granular, "多粒度对比", "🗓️"),
               _reg("frontend", page_frontend, "前端指标", "📈"),
               _reg("backend", page_backend, "后端指标", "💰")],
    "监控与溯源": [_reg("anomaly", page_anomaly, "异常预警", "🚨"),
                _reg("drilldown", page_drilldown, "链接下钻", "🔗"),
                _reg("chain", page_chain, "日环比 & 回暖", "📉"),
                _reg("effect", page_effect, "改动效果溯源", "🧪")],
    "参考": [_reg("catalog", page_catalog, "指标归档", "📚")],
    "达人独有": [_reg("rank", page_rank, "达人 / 内容排行", "🏆")],
}
pg = st.navigation(NAV_GROUPS, position="hidden")
with NAV_BOX:
    for grp, pages in NAV_GROUPS.items():
        if grp == "达人独有" and PROF["key"] != "cr":
            continue
        st.caption(grp)
        for p in pages:
            st.page_link(p, label=p.title, icon=p.icon)
    st.markdown("---")
pg.run()
render_chat_cat()

st.sidebar.markdown("---")
st.sidebar.download_button(f"⬇️ 导出{PROF['name']}明细(长表) CSV",
                           df.to_csv(index=False).encode("utf-8-sig"),
                           f"nailvesta_{PROF['name']}_daily_long.csv", "text/csv")
