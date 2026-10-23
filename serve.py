#!/usr/bin/env python3
"""智能报销录入系统 v2.0 — 桌面产品版
双击 start.command 或运行 python3 serve.py 启动
"""

import json
import os
import sqlite3
import sys
import time
import webbrowser
import urllib.request
import urllib.error
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

def _resolve_dirs():
    """
    PyInstaller 打包后路径策略：
    - 资源文件（template.xlsx / index.html）：在 _MEIPASS（onefile）或 exe 同级（onedir）
    - 数据文件（data.db / exports / invoices）：写入用户家目录 ~/.reimbursement
      （.app 内部不可写）
    非打包状态下，两者都指向 APP_DIR。
    """
    src_dir = Path(__file__).parent
    if getattr(sys, 'frozen', False):
        # PyInstaller onefile：解包资源在 _MEIPASS；onedir：在 exe 同级
        resource_dir = Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
        data_dir = Path.home() / ".reimbursement"
        data_dir.mkdir(exist_ok=True)
        return resource_dir, data_dir
    return src_dir, src_dir

RESOURCE_DIR, DATA_DIR = _resolve_dirs()
APP_DIR = RESOURCE_DIR   # 兼容历史代码中的 APP_DIR 引用（只读资源）
DB_PATH = DATA_DIR / "data.db"
TPL_PATH = RESOURCE_DIR / "template.xlsx"
EXPORTS_DIR = DATA_DIR / "exports"   # 统一存放导出的 Excel
INVOICES_DIR = DATA_DIR / "invoices"  # 发票文件夹
DESKTOP = Path.home() / "Desktop"
PORT = 8765
HOST = "127.0.0.1"
_lock_fp = None  # 单实例锁文件句柄

# ═══════════════════════════════════
# 数据库
# ═══════════════════════════════════

def get_db():
    db = sqlite3.connect(str(DB_PATH), timeout=10)
    db.row_factory = sqlite3.Row
    # WAL 模式：多人并发读写不锁库
    db.execute("PRAGMA journal_mode=WAL")
    db.execute("PRAGMA busy_timeout=5000")
    db.execute("PRAGMA journal_mode=WAL")
    return db

def init_db():
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_name TEXT DEFAULT 'default',
            seq INTEGER DEFAULT 0,
            date TEXT DEFAULT '',
            category TEXT DEFAULT '',
            category_detail TEXT DEFAULT '',
            customer TEXT DEFAULT '',
            approve TEXT DEFAULT '',
            project TEXT DEFAULT '',
            people_n TEXT DEFAULT '',
            count INTEGER,
            amount REAL DEFAULT 0,
            invoice_type TEXT DEFAULT '',
            pay_img TEXT DEFAULT '',
            recv_img TEXT DEFAULT '',
            merchant TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS reimbursements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT DEFAULT '',
            month TEXT DEFAULT '',
            status TEXT DEFAULT 'draft',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            exported_at TIMESTAMP,
            excel_filename TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT DEFAULT '',
            stored_name TEXT DEFAULT '',
            mime_type TEXT DEFAULT '',
            size INTEGER DEFAULT 0,
            remark TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)

    # ── records 表字段扩展（幂等迁移：用 ALTER TABLE ADD COLUMN）──
    cur = db.execute("PRAGMA table_info(records)")
    existing_record_cols = {row["name"] for row in cur.fetchall()}
    record_new_cols = {
        "merchant": "TEXT DEFAULT ''",       # 商家/销售方（星巴克、海底捞）
    }
    for col, ddl in record_new_cols.items():
        if col not in existing_record_cols:
            db.execute(f"ALTER TABLE records ADD COLUMN {col} {ddl}")

    # ── invoices 表字段扩展（幂等迁移：用 ALTER TABLE ADD COLUMN）──
    cur = db.execute("PRAGMA table_info(invoices)")
    existing_cols = {row["name"] for row in cur.fetchall()}
    new_cols = {
        "img_type": "TEXT DEFAULT ''",        # payment/receipt/approval/invoice/project
        "rec_date": "TEXT DEFAULT ''",
        "rec_amount": "REAL",
        "rec_customer": "TEXT DEFAULT ''",
        "rec_approve": "TEXT DEFAULT ''",     # SP 开头申请编号
        "rec_project": "TEXT DEFAULT ''",     # WB 项目编号 / KH 客户编号
        "rec_invoice_no": "TEXT DEFAULT ''",  # 发票号
        "rec_raw": "TEXT DEFAULT ''",         # 完整识别 JSON
        "group_id": "TEXT DEFAULT ''",
        "display_name": "TEXT DEFAULT ''",    # 规范化展示名（无扩展名）
        "group_amount": "REAL",
        "group_date": "TEXT DEFAULT ''",
        "flags": "TEXT DEFAULT ''",           # 组差异标记 JSON
        "rec_guests_n": "INTEGER",            # 客方人数（招待对象）
        "rec_hosts_n": "INTEGER",             # 我方陪同人数
        "rec_opportunity": "TEXT DEFAULT ''", # 商机标题（子品牌，如旺旺GEO）
        "rec_merchant": "TEXT DEFAULT ''",    # 商家/销售方（星巴克、海底捞）
    }
    for col, ddl in new_cols.items():
        if col not in existing_cols:
            db.execute(f"ALTER TABLE invoices ADD COLUMN {col} {ddl}")
    db.commit()

    # 数据迁移：把旧的 default 批次记录收编为一张草稿报销单（只迁移一次）
    orphan = db.execute(
        "SELECT COUNT(*) as c FROM records WHERE batch_name='default'"
    ).fetchone()["c"]
    exists_reimb = db.execute("SELECT COUNT(*) as c FROM reimbursements").fetchone()["c"]
    if orphan > 0 and exists_reimb == 0:
        user = db.execute(
            "SELECT value FROM settings WHERE key='user_name'"
        ).fetchone()
        uname = user["value"] if user else "报销"
        db.execute(
            "INSERT INTO reimbursements(title,month,status) VALUES(?,?, 'draft')",
            (f"{uname}报销单", str(time.localtime().tm_mon)),
        )
        new_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        db.execute(
            "UPDATE records SET batch_name=? WHERE batch_name='default'",
            (f"reimb_{new_id}",),
        )

    defaults = {
        "api_key": "",
        "api_endpoint": "https://api.xiaomimimo.com/v1/chat/completions",
        "api_model": "mimo-v2.5",
        "user_name": "报销用户",
        "export_month": str(time.localtime().tm_mon),
        # ── 偏好设置 ──
        "theme_mode": "light",      # light / dark / auto
        "accent_color": "blue",     # blue / purple / green / orange
        "font_size": "medium",      # small / medium / large
        "desktop_path": "",         # 留空 = 桌面默认路径
    }
    for k, v in defaults.items():
        db.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))

    # ── api_profiles 表：多 Provider Profile 管理（CC Switch 风格）──
    db.execute("""
        CREATE TABLE IF NOT EXISTS api_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            provider TEXT NOT NULL,
            endpoint TEXT NOT NULL,
            api_key TEXT NOT NULL,
            model TEXT NOT NULL,
            is_active INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # 一次性迁移：把现有 settings 里的 api_* 配置转为「默认」Profile 并激活
    cnt = db.execute("SELECT COUNT(*) FROM api_profiles").fetchone()[0]
    if cnt == 0:
        old_key = get_setting("api_key")
        old_endpoint = get_setting("api_endpoint", "https://api.xiaomimimo.com/v1/chat/completions")
        old_model = get_setting("api_model", "mimo-v2.5")
        provider_guess = "xiaomi"
        if "glm" in old_endpoint.lower() or "zhipu" in old_endpoint.lower():
            provider_guess = "glm"
        elif "kimi" in old_endpoint.lower() or "moonshot" in old_endpoint.lower():
            provider_guess = "kimi"
        elif "openai" in old_endpoint.lower():
            provider_guess = "openai"
        elif "anthropic" in old_endpoint.lower() or "claude" in old_endpoint.lower():
            provider_guess = "claude"
        if old_key:
            db.execute("""
                INSERT INTO api_profiles(name, provider, endpoint, api_key, model, is_active)
                VALUES(?,?,?,?,?,1)
            """, ("默认", provider_guess, old_endpoint, old_key, old_model))
    db.commit()
    db.close()

    # 确保导出目录存在
    EXPORTS_DIR.mkdir(exist_ok=True)
    INVOICES_DIR.mkdir(exist_ok=True)

def get_setting(key: str, default: str = "") -> str:
    db = get_db()
    row = db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    db.close()
    return row["value"] if row else default

def set_setting(key: str, value: str):
    db = get_db()
    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, value))
    db.commit()
    db.close()

# ═══════════════════════════════════
# 识别引擎
# ═══════════════════════════════════

EXTRACT_PROMPT = """You are a receipt recognition assistant for a Chinese expense reimbursement system.
First identify the image type, then extract ONLY the fields that this type is responsible for. Return null for all other fields.

## Image Types and Their Responsibilities

### payment (支付截图 - WeChat/Alipay/bank payment completion or bill detail page)
PRIMARY FOCUS: amount(s), date. These will be matched against the closest receipt (水单).
IMPORTANT - detect single vs multiple payments:
(A) SINGLE payment page (支付完成页): one large amount under "支付金额"/"实付".
(B) MULTIPLE payments (账单详情/交易记录, same merchant, within a short time like a few minutes): e.g. tea-break ordering paid in 3 separate transactions.
Extract ONLY:
- amount: for SINGLE payment -> the paid amount (number, e.g. 188.00). For MULTIPLE payments -> the SUM of all amounts (number, e.g. 3 transactions of 35+40+35 = 110.00).
- amounts: ONLY for MULTIPLE payments -> a JSON array of each individual amount, e.g. [35.00, 40.00, 35.00]. For SINGLE payment -> null.
- date: payment date YYYY-MM-DD. For MULTIPLE payments -> use the date of the FIRST/earliest transaction.
This is the AUTHORITATIVE source for the actual paid amount.
- amount: use the PAID amount line ("支付金额"/"实付"/"转账金额"), NOT "商品总额"/"原价".
- date: from "支付时间"/"转账时间"/"交易时间". If only "今天/昨天" shown, use null.
DO NOT extract transaction id or payment method.
- merchant: the payee/merchant name shown on the payment page (e.g. field "收款方"/"商家"/"商品" -> "星巴克"). Used for invoice matching.
All other fields = null.

### receipt (水单/餐厅小票/消费明细单 - e.g. Starbucks order, restaurant bill)
PRIMARY FOCUS: amount, date, merchant. These are the authoritative source for this image.
- amount: the TOTAL paid amount (number, e.g. 110.00). Use "总计/实付/实收" line, NOT "产品小计" if they differ.
- date: consumption/order date YYYY-MM-DD (e.g. from "下单时间 2026/06/18 15:07:30" -> 2026-06-18)
- merchant: the SELLER/restaurant/shop name at the top of the bill (e.g. "星巴克", "海底捞", "瑞幸咖啡"). This is the merchant, NOT the client.
OPTIONAL (only if clearly visible, otherwise null):
- count: number of people/diners if printed on the bill (人数), else null
- people_n: attendee names if printed, else null
IMPORTANT:
- customer = null for receipt. The merchant is the SELLER; the client company (e.g. 旺旺) is NOT on the bill.
- Do NOT confuse "产品小计/subtotal" with the paid total.
All other fields = null.

### approval (事前申请单/审批系统截图 - e.g. "招待费事前申请" detail page)
NOTE: This image type often spans TWO screenshots (page top + page bottom). Extract ALL of the fields below from whatever portion is visible.
Extract ONLY these fields:
- approve: value of the field labeled "审批编号". Format is SP + YYYYMMDD + "-" + 4-digit seq, e.g. "SP20260617-0020".
- date: value of the field labeled "申请时间". Format YYYY-MM-DD HH:MM:SS, keep full precision, e.g. "2026-06-17 15:29:46".
- customer: value of the field labeled "CRM系统-客户名称" (or just "CRM客户名称"). Copy it EXACTLY as shown, do NOT expand/complete it. e.g. "旺旺" stays "旺旺", do NOT write "旺旺集团".
- opportunity: value of the field labeled "项目名称" (项目/商机). Copy EXACTLY. Used later to link to the project sheet.
- amount: value of the field labeled "申请金额" (only this label). Strip the currency unit (元/RMB/人民币), keep the number, e.g. "380.00". Used to match against receipt/payment to group them into one record.
- count: SUM of "客方人数" (来宾/招待对象人数) + "我方陪同人数" (陪同人员人数). Output the TOTAL integer. e.g. 客方2 + 我方1 = 3. This is the total 招待人数.
- guests_n: the "客方人数" value alone (integer), if visible, else null.
- hosts_n: the "我方陪同人数" value alone (integer), if visible, else null.
DO NOT extract applicant name / department / approval status / reason.
If any field is not visible on THIS screenshot, set it to null (the other half may be on a sibling screenshot).
All other fields = null.

### project (项目编号单/客户编号单 - 珍客CRM客户详情页截图)
PRIMARY FOCUS: extract the project number (项目编号) as the key identifier.
Extract ONLY these fields:
- project: the project number (项目编号). Look for a code starting with "WB" (e.g. WB001) FIRST; if no WB code, look for "KH"-prefixed customer number (e.g. KH001). Copy it EXACTLY as shown on the page.
- project_code: the WB number itself (e.g. WB001). null if only KH number is available.
- customer_code: the KH number (e.g. KH001). null if only WB number is available.
- customer: the value of the field labeled "客户名称" (e.g. "旺旺集团"). Copy EXACTLY.
- opportunity: the value of the field labeled "商机标题" (e.g. "旺旺GEO", "旺旺Agent"). Copy EXACTLY. This is the SUB-BRAND/PROJECT name used for fuzzy matching against approval's customer.
IMPORTANT:
- If you cannot find any WB/KH-prefixed code, set project = null and project_code/customer_code = null. Do NOT invent one.
- "客户名称" is the parent company (e.g. 旺旺集团); "商机标题" is the sub-brand/opportunity (e.g. 旺旺GEO). Keep them separate.
All other fields = null.

### invoice (发票 - e.g. 增值税电子普通发票/专用发票)
PRIMARY FOCUS: merchant, amount. (Invoice is handled in a SEPARATE module, not the main matching pipeline.)
Extract ONLY:
- merchant: the SELLER/开票方 company name (销售方名称), e.g. "星巴克咖啡(上海)有限公司". This is who ISSUED the invoice.
- amount: the invoice total (价税合计/合计金额), number with decimals, e.g. 110.00.
- invoice_no: the invoice number (发票号码), if visible.
- date: invoice date (开票日期) YYYY-MM-DD, if visible.
DO NOT extract buyer/purchaser company. DO NOT extract customer.
All other fields = null.

## Rules
- date format: YYYY-MM-DD (e.g. 2026-04-02)
- customer must be a COMPANY name (公司/集团/品牌), NEVER a person name
- amount: number with decimals, e.g. 184.00
- Return ONLY JSON, no explanation. Unrecognized fields = null.

Return format:
{"amount":null,"amounts":null,"date":null,"merchant":null,"customer":null,"approve":null,"project":null,"project_code":null,"customer_code":null,"opportunity":null,"count":null,"people_n":null,"guests_n":null,"hosts_n":null,"invoice_no":null,"invoice_type":null,"img_type":null}
Note: "amounts" is a JSON array of individual amounts for multi-payment screenshots only; otherwise null. "amount" is always the effective total (sum for multi, single for single).
Return ONLY JSON, no explanation."""

def parse_model_json(content: str) -> dict:
    content = content.strip()
    if content.startswith("```"):
        content = content.split("```")[1] if "```" in content else content
        if content.startswith("json"): content = content[4:]
        content = content.strip()
    start = content.find("{")
    end = content.rfind("}")
    if start >= 0 and end > start:
        content = content[start:end + 1]
    try:
        return json.loads(content)
    except Exception:
        return {"error": "模型返回非 JSON", "raw": content[:300]}

def get_active_profile():
    """返回当前激活的 Profile（dict），无则 None"""
    db = get_db()
    row = db.execute("SELECT * FROM api_profiles WHERE is_active=1 LIMIT 1").fetchone()
    db.close()
    return dict(row) if row else None

def call_vision(data_uri: str) -> dict:
    # 优先读激活的 Profile；兼容旧 settings
    profile = get_active_profile()
    if profile:
        api_key = profile["api_key"]
        endpoint = profile["endpoint"]
        model = profile["model"]
    else:
        api_key = get_setting("api_key")
        endpoint = get_setting("api_endpoint", "https://api.xiaomimimo.com/v1/chat/completions")
        model = get_setting("api_model", "mimo-v2.5")
    if not api_key:
        return {"error": "请先在「模型API接口」页面配置并激活一个 API Profile"}

    body = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": EXTRACT_PROMPT},
            {"type": "image_url", "image_url": {"url": data_uri}},
        ]}],
        "temperature": 0.1,
    }

    body_json = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=body_json,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    last_error = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
            resp_obj = json.loads(raw)
            if "error" in resp_obj:
                err = resp_obj["error"]
                last_error = f"API 错误: {err.get('message', str(err)[:200])}"
                code = err.get("code", "")
                if code in ("429", "502", "503") or "rate" in str(err).lower():
                    time.sleep(2 ** attempt)
                    continue
                break
            content = resp_obj["choices"][0]["message"]["content"]
            return parse_model_json(content)
        except urllib.error.HTTPError as e:
            # 4xx/5xx：读 body 看是否限流
            try:
                err_body = e.read().decode("utf-8", errors="replace")[:300]
            except Exception:
                err_body = ""
            last_error = f"HTTP {e.code}: {err_body}"
            if e.code in (429, 502, 503):
                time.sleep(2 ** attempt)
                continue
            break
        except urllib.error.URLError as e:
            last_error = f"网络错误: {e.reason}"
            time.sleep(1.5)
            continue
        except json.JSONDecodeError:
            last_error = f"模型返回非 JSON: {raw[:300]}"
            break
        except Exception as e:
            last_error = f"识别失败: {e}"
            break

    return {"error": last_error or "识别失败（已重试3次）"}


def call_llm(messages: list, max_tokens: int = 1024, temperature: float = 0.3) -> dict:
    """
    通用文本对话接口（复用激活的 Profile）。
    输入：messages = [{role, content}, ...]
    输出：{"ok": True, "content": "..."} 或 {"ok": False, "error": "..."}
    """
    profile = get_active_profile()
    if profile:
        api_key = profile["api_key"]
        endpoint = profile["endpoint"]
        model = profile["model"]
    else:
        api_key = get_setting("api_key")
        endpoint = get_setting("api_endpoint", "https://api.xiaomimimo.com/v1/chat/completions")
        model = get_setting("api_model", "mimo-v2.5")
    if not api_key:
        return {"ok": False, "error": "请先在「模型API接口」页面配置并激活一个 Profile"}

    body = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    body_json = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        endpoint, data=body_json,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            obj = json.loads(resp.read().decode("utf-8", errors="replace"))
        if "error" in obj:
            return {"ok": False, "error": f"API 错误: {obj['error']}"}
        content = obj["choices"][0]["message"]["content"]
        return {"ok": True, "content": content}
    except urllib.error.HTTPError as e:
        try: err_body = e.read().decode("utf-8", errors="replace")[:200]
        except Exception: err_body = ""
        return {"ok": False, "error": f"HTTP {e.code}: {err_body}"}
    except Exception as e:
        return {"ok": False, "error": str(e)}

TYPE_SUFFIX = {
    "payment": "支付", "receipt": "水单", "approval": "申请",
    "invoice": "发票", "project": "项目",
}

def _parse_date(s):
    """解析日期字符串为 (y,m,d) 整数，失败返回 None"""
    if not s:
        return None
    import re
    m = re.search(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", str(s))
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

def _date_eq(d1, d2, tol=1):
    """日期近似相等（tol 天内）。任一为 None 返回 False"""
    a, b = _parse_date(d1), _parse_date(d2)
    if not a or not b:
        return False
    import datetime
    try:
        da = datetime.date(*a); db = datetime.date(*b)
        return abs((da - db).days) <= tol
    except Exception:
        return False

def _amount_eq(a1, a2, tol=1.0):
    try:
        return abs(float(a1) - float(a2)) <= tol
    except Exception:
        return False

def _pick_customer(*items):
    """按 receipt > payment > invoice > approval > project 优先级取客户名；
    客户名都缺时退到商家名（receipt/payment 都会提取 merchant）；都没有才返回"未命名"。"""
    priority = {"receipt": 0, "payment": 1, "invoice": 2, "approval": 3, "project": 4}
    for it in sorted(items, key=lambda x: priority.get(x.get("img_type", ""), 9)):
        c = (it.get("rec_customer") or "").strip()
        if c:
            return c
    # 兜底 1：取商家名（receipt/payment 都会提取 merchant 字段）
    for it in sorted(items, key=lambda x: priority.get(x.get("img_type", ""), 9)):
        m = (it.get("rec_merchant") or "").strip()
        if m:
            return m
    # 兜底 2：从 raw 取 location（保留旧逻辑，虽然 EXTRACT_PROMPT 通常不返回此字段）
    for it in items:
        raw = it.get("rec_raw_dict") or {}
        loc = (raw.get("location") or "").strip()
        if loc:
            return loc
    return "未命名"

def _format_dot_date(s):
    """2026-05-20 → '2026.5.20'（去前导零）"""
    p = _parse_date(s)
    if not p:
        return ""
    return f"{p[0]}.{p[1]}.{p[2]}"

import re
_CUST_SUFFIX_RE = re.compile(r'(集团|公司|有限|科技|股份|控股|网络|技术|Agent|GEO|项目)$')
def _customer_match(name1, name2):
    """模糊客户名匹配：去后缀取核心词，看包含关系。例：旺旺GEO ↔ 旺旺 → True"""
    if not name1 or not name2:
        return False
    n1, n2 = name1.strip(), name2.strip()
    if n1 == n2:
        return True
    shorter, longer = (n1, n2) if len(n1) <= len(n2) else (n2, n1)
    core = _CUST_SUFFIX_RE.sub('', shorter).strip()
    if len(core) >= 2 and core in longer:
        return True
    core2 = _CUST_SUFFIX_RE.sub('', longer).strip()
    if len(core2) >= 2 and core2 in shorter:
        return True
    return False

def _normalize_customer(name):
    """清理商家名用于文件名"""
    if not name:
        return "未命名"
    # 去除文件名非法字符
    for ch in '/\\:*?"<>|':
        name = name.replace(ch, "")
    return name.strip()[:30] or "未命名"

def group_invoices(invoice_rows):
    """
    输入：list of dict（每张发票的 db row，含 rec_* 字段）
    输出：list of {
        group_id: str,
        members: [row, ...],
        date: str,           # 组最终日期（消费核心日期，无核心取最早）
        amount: float,       # 组最终金额（取更低）
        customer: str,
        flags: {approved, paid, diff} 或 None,
    }

    分组规则（按你定的交叉校验）：
    1. receipt+payment 同日期(±1天)同客户 → 消费核心
    2. approval 按 ±1 天归入最近核心；无匹配单独成组
    3. invoice 按 同客户+金额≈水单(±1) 归入核心；无匹配单独成组
    4. project 按 WB/KH 编号归入已有同编号组；无匹配单独成组
    """
    rows = [dict(r) for r in invoice_rows if r.get("img_type")]
    # 解析 rec_raw 给每条，便于取 location
    for r in rows:
        try:
            r["rec_raw_dict"] = json.loads(r.get("rec_raw") or "{}")
        except Exception:
            r["rec_raw_dict"] = {}

    receipts  = [r for r in rows if r["img_type"] == "receipt"]
    payments  = [r for r in rows if r["img_type"] == "payment"]
    approvals = [r for r in rows if r["img_type"] == "approval"]
    invoices  = [r for r in rows if r["img_type"] == "invoice"]
    projects  = [r for r in rows if r["img_type"] == "project"]

    groups = []  # 每组：{members:[...], customer, dates:[...], core_amount, approval_amount}
    used_ids = set()

    def _new_group(seed):
        members = list(seed)
        g = {
            "members": members,
            "_ids": {m["id"] for m in members},
            "_customer_candidates": members,
        }
        groups.append(g)
        return g

    # ── Step 1: 强绑定消费核心 payment↔receipt（金额+日期匹配）──
    # 注意：payment 和 receipt 都只提取 amount+date，不提取客户名。
    # 匹配规则：金额相近(±1元) + 日期相近(±1天) → 绑定
    # 退而求其次：仅金额完全相等（无日期信息时）
    matched_receipts = set()  # 存放已被匹配的 receipt 的主键 id（不要用 id()，那是 Python 对象地址）
    for p in payments:
        best = None
        best_score = 999
        p_amt = p.get("rec_amount")
        p_date = _parse_date(p.get("rec_date"))
        for rc in receipts:
            if rc["id"] in matched_receipts:
                continue
            rc_amt = rc.get("rec_amount")
            rc_date = _parse_date(rc.get("rec_date"))
            same_amt = _amount_eq(p_amt, rc_amt)
            # 日期差
            days = None
            if p_date and rc_date:
                import datetime
                try:
                    days = abs((datetime.date(*p_date) - datetime.date(*rc_date)).days)
                except Exception:
                    pass
            # 打分：金额相等+日期近=0；金额相等(无日期)=1；金额近+日期近=2；仅日期近=3
            if same_amt and days is not None and days <= 1:
                score = 0
            elif same_amt and days is None:
                score = 1
            elif _amount_eq(p_amt, rc_amt, tol=2.0) and days is not None and days <= 1:
                score = 2
            elif days is not None and days <= 1:
                # 同日但金额差较大：仍可作为候选（可能多笔招待）
                score = 3
            else:
                continue
            if score < best_score:
                best_score = score; best = rc
        if best:
            matched_receipts.add(best["id"])
            _new_group([p, best])
            used_ids.add(p["id"]); used_ids.add(best["id"])
        else:
            # payment 单独成组（无匹配水单）
            _new_group([p]); used_ids.add(p["id"])

    # 剩余未配 receipt 单独成组
    for rc in receipts:
        if rc["id"] not in matched_receipts:
            _new_group([rc]); used_ids.add(rc["id"])

    # ── Step 2: approval 按日期+金额归入最近核心 ──
    # 申请单金额是预估，但与水单/支付金额应接近；同时日期接近(±1天)
    # 优先：同日期 + 金额相近；其次：同日期；最后：仅金额相近
    for ap in approvals:
        best_g = None; best_score = 999
        ap_amt = ap.get("rec_amount")
        ap_d = _parse_date(ap.get("rec_date"))
        for g in groups:
            for m in g["members"]:
                d2 = _parse_date(m.get("rec_date"))
                # 日期差
                days = None
                if ap_d and d2:
                    import datetime
                    try:
                        days = abs((datetime.date(*ap_d) - datetime.date(*d2)).days)
                    except Exception:
                        pass
                # 金额相近（与 receipt/payment 的实付金额）
                amt_close = _amount_eq(ap_amt, m.get("rec_amount"), tol=2.0) if ap_amt else False
                is_core = m["img_type"] in ("receipt", "payment")
                # 打分：同日+金额近=0；同日=1；金额近(跨日)=2；其他不考虑
                if days is not None and days <= 1 and amt_close:
                    score = 0
                elif days is not None and days <= 1:
                    score = 1
                elif amt_close and is_core:
                    score = 2
                else:
                    continue
                if score < best_score:
                    best_score = score; best_g = g
        if best_g:
            best_g["members"].append(ap)
            best_g["_ids"].add(ap["id"])
            best_g["_customer_candidates"] = list(best_g["members"])
        else:
            _new_group([ap])
        used_ids.add(ap["id"])

    # ── Step 3: invoice 按 同客户+金额≈水单(±1) 归入核心 ──
    for iv in invoices:
        best_g = None; best_score = 999
        for g in groups:
            for m in g["members"]:
                same_cust = (iv.get("rec_customer") or "") and (m.get("rec_customer") or "") and \
                            iv["rec_customer"].strip() == m["rec_customer"].strip()
                amt_eq = _amount_eq(iv.get("rec_amount"), m.get("rec_amount"))
                # 优先匹配 receipt/payment（核心），其次 approval
                is_core = m["img_type"] in ("receipt", "payment")
                if same_cust and amt_eq:
                    score = 0 if is_core else 1
                    if score < best_score:
                        best_score = score; best_g = g
                        break
            if best_g and best_score == 0:
                break
        if best_g:
            best_g["members"].append(iv)
            best_g["_ids"].add(iv["id"])
            best_g["_customer_candidates"] = list(best_g["members"])
        else:
            # 退而求其次：仅按客户匹配
            for g in groups:
                for m in g["members"]:
                    if (iv.get("rec_customer") or "") and (m.get("rec_customer") or "") and \
                       iv["rec_customer"].strip() == m["rec_customer"].strip():
                        best_g = g; break
                if best_g:
                    break
            if best_g:
                best_g["members"].append(iv)
                best_g["_ids"].add(iv["id"])
            else:
                _new_group([iv])
        used_ids.add(iv["id"])

    # ── Step 4: project 匹配 ──
    # 优先级：①同 WB/KH 编号 > ②商机标题与 approval 客户模糊匹配 > ③客户名模糊匹配
    for pj in projects:
        code = (pj.get("rec_project") or "").strip()
        pj_opp = (pj.get("rec_opportunity") or "").strip()
        pj_cust = (pj.get("rec_customer") or "").strip()
        best_g = None
        best_score = 999
        # ① 同 WB/KH 编号精确匹配
        if code:
            for g in groups:
                for m in g["members"]:
                    if (m.get("rec_project") or "").strip() == code:
                        best_g = g; best_score = 0; break
                if best_g:
                    break
        # ② 商机标题(opportunity) ↔ approval/组的客户名 模糊匹配
        if not best_g and pj_opp:
            for g in groups:
                for m in g["members"]:
                    m_cust = (m.get("rec_customer") or "").strip()
                    if m["img_type"] == "approval" and _customer_match(pj_opp, m_cust):
                        best_g = g; best_score = 1; break
                if best_g:
                    break
        # ③ 客户名模糊匹配（兜底）
        if not best_g and pj_cust:
            for g in groups:
                for m in g["members"]:
                    m_cust = (m.get("rec_customer") or "").strip()
                    if m_cust and _customer_match(pj_cust, m_cust):
                        best_g = g; best_score = 2; break
                if best_g:
                    break
        if best_g:
            best_g["members"].append(pj)
            best_g["_ids"].add(pj["id"])
        else:
            _new_group([pj])
        used_ids.add(pj["id"])

    # ── 计算每组最终值 ──
    import datetime
    results = []
    for i, g in enumerate(groups):
        members = g["members"]
        # 客户名（优先级 receipt>payment>invoice>approval>project）
        customer = _pick_customer(*members)
        # 日期：优先消费核心日期，无核心取最早非空日期
        core_dates = [m.get("rec_date") for m in members
                      if m["img_type"] in ("receipt", "payment") and m.get("rec_date")]
        all_dates = [m.get("rec_date") for m in members if m.get("rec_date")]
        date_pool = core_dates or all_dates
        if date_pool:
            parsed = [(_parse_date(d), d) for d in date_pool]
            parsed = [p for p in parsed if p[0]]
            if core_dates:
                # 多数票：取消费核心最常见的日期
                from collections import Counter
                cnt = Counter(p[0] for p in parsed)
                final_date_tuple = cnt.most_common(1)[0][0]
            else:
                final_date_tuple = min(p[0] for p in parsed)
            final_date = f"{final_date_tuple[0]:04d}-{final_date_tuple[1]:02d}-{final_date_tuple[2]:02d}"
        else:
            final_date = ""

        # 金额：取组内所有非空 amount 的最小值（你的「取更低」规则）
        amounts = [float(m.get("rec_amount") or 0) for m in members
                   if m.get("rec_amount") not in (None, "", 0)]
        final_amount = min(amounts) if amounts else 0.0

        # 差异标记：申请金额 vs 实付金额
        approval_amt = None
        for m in members:
            if m["img_type"] == "approval" and m.get("rec_amount"):
                try:
                    approval_amt = float(m["rec_amount"])
                except Exception:
                    pass
        paid_amt = None
        for m in members:
            if m["img_type"] in ("receipt", "payment", "invoice") and m.get("rec_amount"):
                try:
                    paid_amt = float(m["rec_amount"])
                    break
                except Exception:
                    pass
        flags = None
        if approval_amt is not None and paid_amt is not None and approval_amt != paid_amt:
            flags = {
                "approved": round(approval_amt, 2),
                "paid": round(paid_amt, 2),
                "diff": round(paid_amt - approval_amt, 2),
            }

        # 多笔支付信息：从 payment 成员的 raw JSON 取 amounts 数组
        multi_pay = None
        for m in members:
            if m["img_type"] == "payment":
                try:
                    raw = json.loads(m.get("rec_raw") or "{}")
                    arr = raw.get("amounts")
                    if isinstance(arr, list) and len(arr) >= 2:
                        multi_pay = {
                            "items": arr,
                            "count": len(arr),
                            "sum": round(sum(float(x) for x in arr), 2),
                        }
                        break
                except Exception:
                    pass

        # 招待人数：优先用 approval 的 count（已客方+我方求和）
        guest_count = None
        guests_n = None
        hosts_n = None
        for m in members:
            if m["img_type"] == "approval":
                if m.get("count") is not None and guest_count is None:
                    guest_count = m["count"]
                if m.get("rec_guests_n") is not None and guests_n is None:
                    guests_n = m["rec_guests_n"]
                if m.get("rec_hosts_n") is not None and hosts_n is None:
                    hosts_n = m["rec_hosts_n"]
        # 兜底：水单若打印了人数，approval 没有则取水单
        if guest_count is None:
            for m in members:
                if m["img_type"] == "receipt" and m.get("count"):
                    guest_count = m["count"]
                    break

        results.append({
            "group_id": f"g{i+1}",
            "members": members,
            "date": final_date,
            "amount": round(final_amount, 2),
            "customer": customer,
            "flags": flags,
            "multi_pay": multi_pay,
            "count": guest_count,
            "guests_n": guests_n,
            "hosts_n": hosts_n,
        })

    return results


def build_display_name(group, member):
    """
    根据组最终值 + 单张类型，生成展示名（无扩展名）。
    例：'2026.5.20 184.00 星巴克_水单'
    """
    date_str = _format_dot_date(group["date"])
    amount_str = f"{group['amount']:.2f}" if group["amount"] else "0.00"
    customer = _normalize_customer(group["customer"])
    suffix = TYPE_SUFFIX.get(member.get("img_type", ""), "其他")
    base = f"{date_str} {amount_str} {customer}_{suffix}" if date_str else \
           f"{amount_str} {customer}_{suffix}"
    return base


def rename_invoice_file(stored_name, display_name):
    """
    实际重命名磁盘文件，返回新的 stored_name。
    保持原扩展名。冲突时追加 (2)、(3)...
    """
    if not stored_name:
        return stored_name
    ext = Path(stored_name).suffix
    # 清理 display_name 用于文件系统
    safe_name = display_name
    for ch in '/\\:*?"<>|':
        safe_name = safe_name.replace(ch, "")
    new_name = f"{safe_name}{ext}"
    src = INVOICES_DIR / stored_name
    if not src.is_file():
        return stored_name
    dst = INVOICES_DIR / new_name
    n = 2
    while dst.exists() and dst.resolve() != src.resolve():
        dst = INVOICES_DIR / f"{safe_name} ({n}){ext}"
        n += 1
    try:
        src.rename(dst)
        return dst.name
    except OSError:
        return stored_name


# ═══════════════════════════════════
# HTTP 服务
# ═══════════════════════════════════

class Handler(SimpleHTTPRequestHandler):
    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        return self.rfile.read(length) if length else b""

    def _get_db_schema(self):
        """返回数据库 schema 描述（供 Text-to-SQL）"""
        return """Table records (报销记录):
  id, batch_name, seq, date (YYYY-MM-DD), category, category_detail, customer,
  approve, project, people_n, count, amount (use CAST(amount AS REAL) for math),
  invoice_type, pay_img, recv_img, merchant, remark, created_at.
  default batch_name is 'default'. Reimbursement batches use values like 'reimb_3'.

Table reimbursements (报销单):
  id, title, month, status (draft/exported), created_at, exported_at, excel_filename.

Table invoices (发票库):
  id, filename, stored_name, mime_type, size, remark, created_at,
  img_type (payment/receipt/approval/invoice/project), rec_date, rec_amount,
  rec_customer, rec_approve, rec_project, rec_invoice_no, rec_raw, group_id,
  display_name, group_amount, group_date, flags, rec_guests_n, rec_hosts_n,
  rec_opportunity, rec_merchant."""

    def _run_text_to_sql(self, question: str) -> str:
        """
        Text-to-SQL 流程：
        1. 让 AI 生成 SQL（仅 SELECT）
        2. 安全校验：禁止非 SELECT、限制 100 行
        3. 执行 SQL
        4. 把结果交回 AI 总结成自然语言
        """
        schema = self._get_db_schema()
        gen_messages = [
            {
                "role": "system",
                "content": (
                    "You are a SQLite SQL generator. Given a natural language question, output ONE SQLite SELECT query.\n"
                    "Schema:\n" + schema +
                    "\n\nRules:\n"
                    "1. Output ONLY SELECT statements. Never INSERT/UPDATE/DELETE/DROP/ALTER.\n"
                    "2. amount column is TEXT - use CAST(amount AS REAL) for math.\n"
                    "3. date format is 'YYYY-MM-DD'. Use strftime('%m', date) for month.\n"
                    "4. Add LIMIT 100 at the end.\n"
                    "5. Return ONLY the raw SQL, no explanation, no markdown fences.\n"
                    "6. If the question cannot be answered with SQL, return exactly: CANNOT_ANSWER"
                ),
            },
            {"role": "user", "content": question},
        ]
        gen = call_llm(gen_messages, max_tokens=300, temperature=0.0)
        # 重试一次（MiMo 偶发返回空）
        if not gen.get("ok") or not gen.get("content", "").strip():
            time.sleep(0.5)
            gen = call_llm(gen_messages, max_tokens=300, temperature=0.2)
        if not gen.get("ok"):
            return f"查询生成失败：{gen.get('error')}"

        sql = gen["content"].strip()
        # 用正则提取 SQL：去除 markdown 代码块包裹
        import re as _re
        m = _re.search(r"```(?:sql)?\s*([\s\S]+?)```", sql)
        if m:
            sql = m.group(1).strip()
        else:
            # 没有代码块，去掉首尾可能的 ` 和 sql 标记
            sql = sql.strip("`").strip()
            if sql.lower().startswith("sql\n"):
                sql = sql[4:].strip()
        if sql == "CANNOT_ANSWER" or not sql:
            return ("抱歉，这个问题暂时无法精确查询（AI 未能生成有效 SQL）。\n\n"
                    "建议换种问法试试：\n"
                    "- 「所有报销记录有多少条」\n"
                    "- 「按费用类别分组统计金额」\n"
                    "- 「总金额合计是多少」\n"
                    "- 「按月份分组统计」")

        # 安全校验：多层防御，避免通过 NL→SQL 读取敏感表（尤其 api_profiles.api_key 明文）
        sql_lower = sql.lower().strip()
        # 1) 禁止多语句 / 注释（防止堆叠和注释掉校验）
        if ";" in sql_lower or "--" in sql_lower or "/*" in sql_lower:
            return "⚠ 检测到不安全字符（;/--/*），已拒绝执行"
        # 2) 黑名单：危险动词 + 敏感表 + 元数据 + 联合注入
        blacklist = [
            "insert", "update", "delete", "drop", "alter", "create", "truncate",
            "attach", "detach", "pragma", "replace", "vacuum", "reindex",
            "union", "intersect", "except",             # 联合注入窃取他表
            "sqlite_master", "sqlite_",                  # 元数据（含建表语句/表名）
            "api_profiles",                              # 含明文 API key
            "settings",                                  # 可能含其他敏感配置
        ]
        for kw in blacklist:
            if kw in sql_lower:
                return f"⚠ 检测到不安全操作或受限表（{kw}），已拒绝执行"
        # 3) 必须是 SELECT 或 WITH 开头
        if not (sql_lower.startswith("select") or sql_lower.startswith("with")):
            return f"⚠ 仅允许 SELECT 查询。生成的 SQL：\n```\n{sql[:200]}\n```"

        # 执行（用只读事务包起来，进一步防止任何写操作）
        try:
            db = get_db()
            db.execute("BEGIN")  # 开启事务，下面 ROLLBACK 保证不产生副作用
            try:
                rows = db.execute(sql).fetchall()
            finally:
                db.execute("ROLLBACK")
            db.close()
        except Exception as e:
            return f"SQL 执行出错：{e}\n\n生成的 SQL：\n```\n{sql}\n```"

        # 转为可读结构
        if not rows:
            summary_data = "查询结果为空（没有匹配的数据）"
        else:
            cols = rows[0].keys()
            table_lines = [" | ".join(cols)]
            table_lines.append("-" * 40)
            for r in rows[:50]:
                table_lines.append(" | ".join(str(r[c]) for c in cols))
            summary_data = f"共 {len(rows)} 行：\n" + "\n".join(table_lines)

        # 交回 AI 总结
        sum_messages = [
            {
                "role": "system",
                "content": (
                    "你是数据解读助手。我会给你一个 SQL 查询结果，请用简洁的中文自然语言回答用户的问题。"
                    "要点：\n"
                    "- 直接回答问题，不要复述 SQL\n"
                    "- 金额保留 2 位小数并加 ¥ 符号\n"
                    "- 列举时用编号或分点\n"
                    "- 数据为空时明确说『没有相关记录』"
                ),
            },
            {
                "role": "user",
                "content": f"用户问题：{question}\n\nSQL：{sql}\n\n查询结果：\n{summary_data}",
            },
        ]
        s = call_llm(sum_messages, max_tokens=500, temperature=0.2)
        if not s.get("ok"):
            return f"数据已查询到，但总结失败：{s.get('error')}\n\n原始结果：\n{summary_data}"
        return s["content"]

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path
        qs = parse_qs(urlparse(self.path).query)

        if path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            with open(APP_DIR / "index.html", "rb") as f:
                self.wfile.write(f.read())
            return

        # GET /api/stats —— 数据报表统计
        if path == "/api/stats":
            db = get_db()
            # 1. 总览
            overview = db.execute("""
                SELECT
                    COUNT(*) as total_records,
                    COALESCE(SUM(CAST(amount AS REAL)), 0) as total_amount,
                    COALESCE(SUM(CASE WHEN CAST(amount AS REAL) > 0 THEN 1 ELSE 0 END), 0) as expense_count,
                    COALESCE(AVG(CAST(amount AS REAL)), 0) as avg_amount,
                    COALESCE(MIN(CAST(amount AS REAL)), 0) as min_amount,
                    COALESCE(MAX(CAST(amount AS REAL)), 0) as max_amount
                FROM records
            """).fetchone()

            # 2. 按费用类别分组
            by_category = db.execute("""
                SELECT
                    COALESCE(NULLIF(category, ''), '未分类') as category,
                    COUNT(*) as count,
                    SUM(CAST(amount AS REAL)) as amount
                FROM records
                GROUP BY COALESCE(NULLIF(category, ''), '未分类')
                ORDER BY amount DESC
            """).fetchall()

            # 3. 按月份分组
            by_month = db.execute("""
                SELECT
                    COALESCE(strftime('%Y-%m', date), '未确定') as month,
                    COUNT(*) as count,
                    SUM(CAST(amount AS REAL)) as amount
                FROM records
                WHERE date != ''
                GROUP BY month
                ORDER BY month
            """).fetchall()

            # 4. 按客户分组（Top 10）
            by_customer = db.execute("""
                SELECT
                    COALESCE(NULLIF(customer, ''), '未指定') as customer,
                    COUNT(*) as count,
                    SUM(CAST(amount AS REAL)) as amount
                FROM records
                GROUP BY COALESCE(NULLIF(customer, ''), '未指定')
                ORDER BY amount DESC
                LIMIT 10
            """).fetchall()

            # 5. 按招待人分组
            by_person = db.execute("""
                SELECT
                    COALESCE(NULLIF(approve, ''), '未指定') as person,
                    COUNT(*) as count,
                    SUM(CAST(amount AS REAL)) as amount
                FROM records
                WHERE approve != ''
                GROUP BY approve
                ORDER BY amount DESC
                LIMIT 10
            """).fetchall()

            # 6. 最近 30 天日趋势
            daily = db.execute("""
                SELECT
                    date,
                    SUM(CAST(amount AS REAL)) as amount,
                    COUNT(*) as count
                FROM records
                WHERE date != ''
                  AND date >= date('now', '-90 days')
                GROUP BY date
                ORDER BY date
            """).fetchall()

            # 7. 报销单状态分布
            reimb_status = db.execute("""
                SELECT status, COUNT(*) as count FROM reimbursements GROUP BY status
            """).fetchall()

            db.close()
            self._json({
                "overview": dict(overview),
                "by_category": [dict(r) for r in by_category],
                "by_month": [dict(r) for r in by_month],
                "by_customer": [dict(r) for r in by_customer],
                "by_person": [dict(r) for r in by_person],
                "daily": [dict(r) for r in daily],
                "reimb_status": [dict(r) for r in reimb_status],
            })
            return

        if path == "/api/config":
            profile = get_active_profile()
            self._json({
                "api_ready": bool(profile and profile["api_key"]) or bool(get_setting("api_key")),
                "model": profile["model"] if profile else get_setting("api_model"),
                "endpoint": profile["endpoint"] if profile else get_setting("api_endpoint"),
                "active_profile": profile["name"] if profile else None,
                "provider": profile["provider"] if profile else None,
                "user_name": get_setting("user_name"),
            })
            return

        if path == "/api/settings":
            db = get_db()
            rows = db.execute("SELECT key, value FROM settings").fetchall()
            db.close()
            settings = {r["key"]: r["value"] for r in rows}
            if "api_key" in settings:
                settings["api_key"] = ""
            self._json(settings)
            return

        # GET /api/profiles —— 列出所有 Profile
        if path == "/api/profiles":
            db = get_db()
            rows = db.execute(
                "SELECT id, name, provider, endpoint, api_key, model, is_active, created_at "
                "FROM api_profiles ORDER BY is_active DESC, id ASC"
            ).fetchall()
            db.close()
            # 对 api_key 脱敏（前端只显示尾 4 位）
            def mask(k):
                if not k: return ""
                return (k[:4] + "****" + k[-4:]) if len(k) > 8 else "****"
            self._json([{
                "id": r["id"], "name": r["name"], "provider": r["provider"],
                "endpoint": r["endpoint"], "api_key": "", "api_key_mask": mask(r["api_key"]),
                "model": r["model"], "is_active": bool(r["is_active"]),
            } for r in rows])
            return

        if path == "/api/records":
            batch = qs.get("batch", ["default"])[0]
            db = get_db()
            rows = db.execute(
                "SELECT * FROM records WHERE batch_name=? ORDER BY seq", (batch,)
            ).fetchall()
            db.close()
            self._json([dict(r) for r in rows])
            return

        if path == "/api/batches":
            db = get_db()
            rows = db.execute(
                "SELECT batch_name, COUNT(*) as count, MAX(created_at) as latest FROM records GROUP BY batch_name ORDER BY latest DESC"
            ).fetchall()
            db.close()
            self._json([dict(r) for r in rows])
            return

        if path == "/api/reimbursements":
            db = get_db()
            rows = db.execute("""
                SELECT r.id, r.title, r.month, r.status, r.created_at,
                       r.exported_at, r.excel_filename,
                       (SELECT COUNT(*) FROM records rec WHERE rec.batch_name='reimb_'||r.id) AS rec_count,
                       (SELECT COALESCE(SUM(amount),0) FROM records rec WHERE rec.batch_name='reimb_'||r.id) AS total_amount
                FROM reimbursements r
                ORDER BY r.created_at DESC
            """).fetchall()
            db.close()
            self._json([dict(r) for r in rows])
            return

        # GET /api/reimbursements/<id>
        if path.startswith("/api/reimbursements/"):
            try:
                rid = int(path.rsplit("/", 1)[-1])
            except ValueError:
                self._json({"error": "无效的报销单 ID"}, 400)
                return
            db = get_db()
            r = db.execute(
                "SELECT * FROM reimbursements WHERE id=?", (rid,)
            ).fetchone()
            if not r:
                db.close()
                self._json({"error": "报销单不存在"}, 404)
                return
            recs = db.execute(
                "SELECT * FROM records WHERE batch_name=? ORDER BY seq",
                (f"reimb_{rid}",),
            ).fetchall()
            db.close()
            self._json({
                "id": r["id"], "title": r["title"], "month": r["month"],
                "status": r["status"], "created_at": r["created_at"],
                "exported_at": r["exported_at"], "excel_filename": r["excel_filename"],
                "records": [dict(x) for x in recs],
            })
            return

        # GET /api/exports/<filename> —— 下载已导出的 Excel
        if path.startswith("/api/exports/"):
            import urllib.parse
            raw = path[len("/api/exports/"):]
            fname = urllib.parse.unquote(raw)
            # 防目录穿越
            if "/" in fname or ".." in fname:
                self.send_response(400); self.end_headers(); return
            fpath = EXPORTS_DIR / fname
            if not fpath.is_file():
                self.send_response(404); self.end_headers(); return
            file_data = fpath.read_bytes()
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Length", str(len(file_data)))
            # 中文文件名用 RFC 5987 编码，避免 latin-1 header 编码错误
            ascii_fallback = fname.encode("ascii", "ignore").decode("ascii") or "export.xlsx"
            encoded = urllib.parse.quote(fname)
            self.send_header(
                "Content-Disposition",
                f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}",
            )
            self.end_headers()
            self.wfile.write(file_data)
            return

        # GET /api/reveal/<filename> —— 在 Finder 中显示已导出的 Excel
        # （pywebview 的 WKWebView 不能可靠触发下载，改用系统命令显示文件）
        if path.startswith("/api/reveal/"):
            import urllib.parse, subprocess as _sp
            raw = path[len("/api/reveal/"):]
            fname = urllib.parse.unquote(raw)
            if "/" in fname or ".." in fname:
                self._json({"error": "非法文件名"}, 400); return
            fpath = EXPORTS_DIR / fname
            if not fpath.is_file():
                self._json({"error": "文件不存在"}, 404); return
            try:
                # macOS：在 Finder 中显示文件（-R 高亮选中）
                _sp.run(["open", "-R", str(fpath)], check=False)
                self._json({"ok": True, "path": str(fpath)})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # GET /api/invoices —— 发票列表（含识别与分组信息）
        if path == "/api/invoices":
            db = get_db()
            rows = db.execute(
                "SELECT id, filename, stored_name, mime_type, size, remark, created_at, "
                "img_type, rec_date, rec_amount, rec_customer, rec_approve, rec_project, "
                "rec_invoice_no, rec_raw, group_id, display_name, group_amount, group_date, flags "
                "FROM invoices ORDER BY COALESCE(group_id,'~'), created_at DESC"
            ).fetchall()
            db.close()
            self._json([dict(r) for r in rows])
            return

        # GET /api/invoices/file/<stored_name> —— 读取发票文件（图片内联/PDF内联）
        if path.startswith("/api/invoices/file/"):
            import urllib.parse as _up
            raw = path[len("/api/invoices/file/"):]
            stored = _up.unquote(raw)
            if "/" in stored or ".." in stored:
                self.send_response(400); self.end_headers(); return
            fpath = INVOICES_DIR / stored
            if not fpath.is_file():
                self.send_response(404); self.end_headers(); return
            # 从 DB 查原始文件名和 mime
            db = get_db()
            row = db.execute(
                "SELECT filename, mime_type FROM invoices WHERE stored_name=?", (stored,)
            ).fetchone()
            db.close()
            mime = (row["mime_type"] if row else "") or "application/octet-stream"
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            with open(fpath, "rb") as f:
                self.wfile.write(f.read())
            return

        # Never expose local databases, credentials or source via static fallback.
        self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        body = self._read_body()

        # POST /api/invoices —— 上传发票文件（multipart/form-data）
        if path == "/api/invoices":
            try:
                ctype = self.headers.get("Content-Type", "")
                if "multipart/form-data" not in ctype:
                    self._json({"error": "需要 multipart/form-data 请求"}, 400)
                    return
                # 解析边界
                import re as _re
                m = _re.search(r"boundary=(.+)", ctype)
                if not m:
                    self._json({"error": "缺少 boundary"}, 400)
                    return
                boundary = m.group(1).strip().encode()
                # 切分 parts
                parts = body.split(b"--" + boundary)
                file_content = None
                filename = ""
                file_mime = ""
                remark = ""
                for part in parts:
                    if b"Content-Disposition" not in part:
                        continue
                    # 分离头部和正文
                    head_body = part.split(b"\r\n\r\n", 1)
                    if len(head_body) != 2:
                        continue
                    head, content = head_body
                    content = content.rstrip(b"\r\n")
                    head_str = head.decode("utf-8", "ignore")
                    nm = _re.search(r'name="([^"]+)"', head_str)
                    fm = _re.search(r'filename="([^"]*)"', head_str)
                    mt = _re.search(r"Content-Type:\s*(\S+)", head_str)
                    field = nm.group(1) if nm else ""
                    if field == "file" and fm:
                        file_content = content
                        filename = fm.group(1)
                        file_mime = mt.group(1) if mt else "application/octet-stream"
                    elif field == "remark":
                        remark = content.decode("utf-8", "ignore")
                if not file_content or not filename:
                    self._json({"error": "未收到文件"}, 400)
                    return

                # 生成唯一存储名：时间戳 + 原始扩展名
                ts = time.strftime("%Y%m%d_%H%M%S")
                ext = Path(filename).suffix or ""
                # 防止扩展名含路径
                ext = ext if ext and len(ext) <= 12 else ""
                stored_name = f"{ts}_{int(time.time()*1000)%10000}{ext}"
                stored_path = INVOICES_DIR / stored_name
                with open(stored_path, "wb") as f:
                    f.write(file_content)

                db = get_db()
                db.execute(
                    "INSERT INTO invoices(filename, stored_name, mime_type, size, remark) "
                    "VALUES(?,?,?,?,?)",
                    (filename, stored_name, file_mime, len(file_content), remark),
                )
                db.commit()
                new_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
                db.close()
                self._json({
                    "ok": True, "id": new_id, "filename": filename,
                    "stored_name": stored_name, "size": len(file_content),
                })
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/invoices/recognize —— 对指定发票调 AI 识别，写回 rec_* 字段
        if path == "/api/invoices/recognize":
            try:
                data = json.loads(body or "{}")
                ids = data.get("ids", [])  # list[int]
                if not isinstance(ids, list) or not ids:
                    self._json({"error": "缺少 ids 参数"}, 400)
                    return
                db = get_db()
                results = []
                for iid in ids:
                    row = db.execute(
                        "SELECT id, stored_name, mime_type FROM invoices WHERE id=?",
                        (int(iid),),
                    ).fetchone()
                    if not row:
                        results.append({"id": int(iid), "error": "发票不存在"})
                        continue
                    fpath = INVOICES_DIR / row["stored_name"]
                    if not fpath.is_file():
                        results.append({"id": int(iid), "error": "文件缺失"})
                        continue
                    # 读文件 → data URI
                    import base64
                    mime = row["mime_type"] or "image/jpeg"
                    with open(fpath, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode()
                    ext = Path(row["stored_name"]).suffix.lower()
                    if ext == ".pdf" or "pdf" in mime:
                        data_uri = f"data:application/pdf;base64,{b64}"
                    else:
                        data_uri = f"data:{mime};base64,{b64}"
                    rec = call_vision(data_uri)
                    if "error" in rec:
                        results.append({"id": int(iid), "error": rec["error"]})
                        continue
                    # 写回 rec_* 字段
                    def _to_int(v):
                        try: return int(float(v))
                        except Exception: return None
                    db.execute("""
                        UPDATE invoices SET
                            img_type=?, rec_date=?, rec_amount=?, rec_customer=?,
                            rec_approve=?, rec_project=?, rec_invoice_no=?, rec_raw=?,
                            count=?, rec_guests_n=?, rec_hosts_n=?, rec_opportunity=?, rec_merchant=?
                        WHERE id=?
                    """, (
                        rec.get("img_type") or "",
                        rec.get("date") or "",
                        float(rec["amount"]) if rec.get("amount") not in (None, "") else None,
                        rec.get("customer") or "",
                        rec.get("approve") or "",
                        rec.get("project_code") or rec.get("project") or rec.get("customer_code") or "",
                        rec.get("invoice_no") or "",
                        json.dumps(rec, ensure_ascii=False),
                        _to_int(rec.get("count")),
                        _to_int(rec.get("guests_n")),
                        _to_int(rec.get("hosts_n")),
                        rec.get("opportunity") or "",
                        rec.get("merchant") or "",
                        int(iid),
                    ))
                    results.append({"id": int(iid), "ok": True, "img_type": rec.get("img_type")})
                db.commit()
                db.close()
                self._json({"ok": True, "results": results})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/invoices/group —— 对所有发票跑分组算法 + 实际重命名
        if path == "/api/invoices/group":
            try:
                db = get_db()
                rows = db.execute(
                    "SELECT * FROM invoices WHERE img_type != '' ORDER BY created_at"
                ).fetchall()
                rows = [dict(r) for r in rows]
                if not rows:
                    self._json({"ok": True, "groups": [], "message": "没有可分组的发票（需先识别）"})
                    db.close()
                    return

                groups = group_invoices(rows)

                # 收集已用文件名，处理冲突
                used_names = set()
                # 加载目录内现有文件名（不含被处理的）
                existing_files = {f.name for f in INVOICES_DIR.glob("*") if f.is_file()}

                renamed = []
                for g in groups:
                    gid = g["group_id"]
                    for member in g["members"]:
                        display = build_display_name(g, member)
                        # 跨组重名追加序号
                        base_display = display
                        n = 2
                        while display in used_names:
                            display = f"{base_display} ({n})"
                            n += 1
                        used_names.add(display)
                        old_stored = member["stored_name"]
                        new_stored = rename_invoice_file(old_stored, display)
                        db.execute("""
                            UPDATE invoices SET
                                group_id=?, display_name=?, group_amount=?, group_date=?, flags=?,
                                stored_name=?
                            WHERE id=?
                        """, (
                            gid, display,
                            g["amount"], g["date"],
                            json.dumps(g["flags"], ensure_ascii=False) if g["flags"] else "",
                            new_stored,
                            member["id"],
                        ))
                        renamed.append({
                            "id": member["id"],
                            "old": old_stored,
                            "new": new_stored,
                            "display": display,
                        })

                db.commit()
                db.close()
                # 返回分组摘要
                summary = [{
                    "group_id": g["group_id"],
                    "date": g["date"],
                    "amount": g["amount"],
                    "customer": g["customer"],
                    "flags": g["flags"],
                    "count": len(g["members"]),
                    "members": [{"id": m["id"], "img_type": m["img_type"]} for m in g["members"]],
                } for g in groups]
                self._json({"ok": True, "groups": summary, "renamed": renamed})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/invoices/rename_invoice —— 发票专属命名+匹配
        # 对 img_type='invoice' 的发票：按「商家 金额」命名，
        # 可选匹配报销单(merchant+amount) → 命中则把支付日期写到文件名最前面
        if path == "/api/invoices/rename_invoice":
            try:
                db = get_db()
                rows = db.execute(
                    "SELECT * FROM invoices WHERE img_type='invoice' ORDER BY created_at"
                ).fetchall()
                rows = [dict(r) for r in rows]
                if not rows:
                    self._json({"ok": True, "renamed": [], "message": "没有发票文件"})
                    db.close()
                    return

                # 加载所有报销记录，用于按 商家+金额 匹配
                rec_rows = db.execute(
                    "SELECT date, amount, customer, merchant, pay_img, recv_img FROM records ORDER BY date"
                ).fetchall()
                rec_rows = [dict(r) for r in rec_rows]

                used_names = set()
                renamed = []
                for inv in rows:
                    raw = {}
                    try: raw = json.loads(inv.get("rec_raw") or "{}")
                    except Exception: raw = {}
                    merchant = (inv.get("rec_merchant") or raw.get("merchant") or "").strip()
                    amount = inv.get("rec_amount")
                    inv_no = (inv.get("rec_invoice_no") or raw.get("invoice_no") or "").strip()

                    # 匹配报销单：商家模糊匹配 + 金额接近
                    pay_date = ""
                    matched_rec = None
                    if amount:
                        best_rec = None
                        best_score = 999
                        for rec in rec_rows:
                            try:
                                rec_amt = float(rec.get("amount") or 0)
                            except Exception:
                                continue
                            if rec_amt <= 0:
                                continue
                            amt_diff = abs(rec_amt - float(amount))
                            if amt_diff > 2:
                                continue
                            # 商家匹配：发票 merchant vs 报销单 merchant
                            rec_merchant = (rec.get("merchant") or "").strip()
                            merchant_match = False
                            if merchant and rec_merchant:
                                merchant_match = _customer_match(merchant, rec_merchant)
                            # 打分：金额差为主，商家匹配为加分
                            score = amt_diff + (0 if merchant_match else (5 if merchant and rec_merchant else 2.5))
                            if score < best_score:
                                best_score = score
                                best_rec = rec
                        if best_rec:
                            matched_rec = best_rec
                            pay_date = best_rec.get("date") or ""

                    # 构造文件名：[日期] 商家 金额[_发票号]
                    date_prefix = _format_dot_date(pay_date) if pay_date else ""
                    merchant_safe = _normalize_customer(merchant) if merchant else "未知商家"
                    amt_str = f"{float(amount):.2f}" if amount else "0.00"
                    parts = []
                    if date_prefix:
                        parts.append(date_prefix)
                    parts.append(amt_str)
                    parts.append(merchant_safe)
                    base = " ".join(parts) + "_发票"
                    if inv_no:
                        base += f"_{inv_no}"

                    # 处理重名
                    display = base
                    n = 2
                    while display in used_names:
                        display = f"{base} ({n})"
                        n += 1
                    used_names.add(display)

                    old_stored = inv["stored_name"]
                    new_stored = rename_invoice_file(old_stored, display)
                    db.execute("""
                        UPDATE invoices SET
                            display_name=?, group_date=?, group_amount=?, stored_name=?
                        WHERE id=?
                    """, (
                        display, pay_date,
                        float(amount) if amount else None,
                        new_stored,
                        inv["id"],
                    ))
                    renamed.append({
                        "id": inv["id"],
                        "old": old_stored,
                        "new": new_stored,
                        "display": display,
                        "matched": bool(matched_rec),
                        "pay_date": pay_date,
                        "merchant": merchant,
                        "amount": float(amount) if amount else 0,
                    })

                db.commit()
                db.close()
                self._json({"ok": True, "renamed": renamed})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # ── Profile 管理 ──
        # POST /api/profiles          新建
        # PUT  /api/profiles/<id>     更新
        # POST /api/profiles/<id>/activate   激活
        # POST /api/profiles/test     测试连接
        if path == "/api/profiles":
            try:
                d = json.loads(body or "{}")
                name = (d.get("name") or "").strip()
                provider = (d.get("provider") or "custom").strip()
                endpoint = (d.get("endpoint") or "").strip()
                api_key = (d.get("api_key") or "").strip()
                model = (d.get("model") or "").strip()
                if not name or not endpoint or not api_key or not model:
                    self._json({"error": "请填写名称、Endpoint、API Key、Model"}, 400); return
                db = get_db()
                try:
                    cur = db.execute(
                        "INSERT INTO api_profiles(name,provider,endpoint,api_key,model,is_active) VALUES(?,?,?,?,?,0)",
                        (name, provider, endpoint, api_key, model),
                    )
                    db.commit()
                    new_id = cur.lastrowid
                    # 若是第一个 Profile，自动激活
                    cnt = db.execute("SELECT COUNT(*) FROM api_profiles").fetchone()[0]
                    if cnt == 1:
                        db.execute("UPDATE api_profiles SET is_active=1 WHERE id=?", (new_id,))
                        db.commit()
                    db.close()
                    self._json({"ok": True, "id": new_id})
                except sqlite3.IntegrityError:
                    db.close()
                    self._json({"error": "名称已存在"}, 400)
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path.startswith("/api/profiles/") and path.endswith("/activate"):
            try:
                pid = int(path.split("/")[3])
                db = get_db()
                # 清空其它激活，激活指定的
                db.execute("UPDATE api_profiles SET is_active=0")
                cur = db.execute("UPDATE api_profiles SET is_active=1 WHERE id=?", (pid,))
                db.commit()
                if cur.rowcount == 0:
                    db.close()
                    self._json({"error": "Profile 不存在"}, 404); return
                # 同步到 settings（兼容旧逻辑）
                row = db.execute("SELECT * FROM api_profiles WHERE id=?", (pid,)).fetchone()
                if row:
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_key',?)", (row["api_key"],))
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_endpoint',?)", (row["endpoint"],))
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_model',?)", (row["model"],))
                    db.commit()
                db.close()
                self._json({"ok": True})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path == "/api/profiles/test":
            try:
                d = json.loads(body or "{}")
                endpoint = (d.get("endpoint") or "").strip()
                api_key = (d.get("api_key") or "").strip()
                model = (d.get("model") or "").strip()
                profile_id = d.get("profile_id")
                if profile_id and not api_key:
                    db = get_db()
                    row = db.execute("SELECT endpoint, api_key, model FROM api_profiles WHERE id=?", (int(profile_id),)).fetchone()
                    db.close()
                    if not row:
                        self._json({"error": "Profile 不存在"}, 404); return
                    api_key = row["api_key"]
                    endpoint = endpoint or row["endpoint"]
                    model = model or row["model"]
                if not endpoint or not api_key or not model:
                    self._json({"error": "缺少 endpoint/api_key/model"}, 400); return
                # 发一个极简请求验证可用性（用 urllib，不依赖 curl）
                test_body = json.dumps({
                    "model": model,
                    "messages": [{"role": "user", "content": "ping"}],
                    "max_tokens": 5,
                }).encode("utf-8")
                req = urllib.request.Request(
                    endpoint, data=test_body,
                    headers={
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": "application/json",
                    },
                    method="POST",
                )
                try:
                    with urllib.request.urlopen(req, timeout=20) as resp:
                        code = resp.getcode()
                    self._json({"ok": True, "http_code": str(code)})
                except urllib.error.HTTPError as e:
                    self._json({"ok": False, "http_code": str(e.code), "error": f"HTTP {e.code}"})
                except Exception as e:
                    self._json({"ok": False, "http_code": "0", "error": str(e)})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/transcribe —— 语音转文字（浏览器不支持 Web Speech 时的 AI 回退）
        if path == "/api/transcribe":
            try:
                ct = self.headers.get("Content-Type", "")
                if not ct.startswith("audio/") and not ct.startswith("multipart/"):
                    self._json({"error": "需要音频数据"}, 400); return

                # multipart：从 form 中取 audio 字段
                audio_bytes = None
                audio_fmt = "webm"
                if ct.startswith("multipart/"):
                    form = self._read_body()
                    # 简单解析 multipart
                    boundary = ct.split("boundary=")[-1].encode()
                    parts = form.split(b"--" + boundary)
                    for p in parts:
                        if b'name="audio"' in p:
                            # 跳过头部分
                            idx = p.find(b"\r\n\r\n")
                            if idx > 0:
                                audio_bytes = p[idx+4:].rstrip(b"\r\n-")
                                if b"audio/" in p[:idx]:
                                    m = re.search(rb"audio/(\w+)", p[:idx])
                                    if m: audio_fmt = m.group(1).decode()
                            break
                else:
                    audio_bytes = self._read_body()

                if not audio_bytes:
                    self._json({"error": "未收到音频数据"}, 400); return

                # 调 AI 转写（OpenAI 兼容的 /audio/transcriptions）
                profile = get_active_profile()
                if not profile:
                    self._json({"error": "未配置 API"}, 400); return
                endpoint = profile["endpoint"].replace("/chat/completions", "/audio/transcriptions")
                api_key = profile["api_key"]
                model = "whisper-1"  # 多数服务商兼容

                # 构建 multipart/form-data
                boundary = "----ReimBoundary" + str(int(time.time()*1000))
                body_parts = []
                body_parts.append(f"--{boundary}".encode())
                body_parts.append(f'Content-Disposition: form-data; name="model"'.encode())
                body_parts.append(b"")
                body_parts.append(model.encode())
                body_parts.append(f"--{boundary}".encode())
                body_parts.append(
                    f'Content-Disposition: form-data; name="file"; filename="audio.{audio_fmt}"'.encode()
                )
                body_parts.append(f"Content-Type: audio/{audio_fmt}".encode())
                body_parts.append(b"")
                body_parts.append(audio_bytes)
                body_parts.append(f"--{boundary}--".encode())
                body_data = b"\r\n".join(body_parts)

                req = urllib.request.Request(
                    endpoint, data=body_data,
                    headers={
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": f"multipart/form-data; boundary={boundary}",
                    },
                    method="POST",
                )
                try:
                    with urllib.request.urlopen(req, timeout=30) as resp:
                        obj = json.loads(resp.read().decode("utf-8", errors="replace"))
                    text = obj.get("text", "")
                    self._json({"ok": True, "text": text})
                except urllib.error.HTTPError as e:
                    try: err = e.read().decode("utf-8", errors="replace")[:200]
                    except Exception: err = ""
                    # 大多数视觉模型不支持音频接口，回退提示
                    self._json({"error": f"语音转写失败（HTTP {e.code}）：{err}。建议使用 Chrome 浏览器的本地语音识别。"}, 502)
                except Exception as e:
                    self._json({"error": f"语音转写失败：{e}"}, 500)
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/ask_data —— 智能问数（Text-to-SQL）
        if path == "/api/ask_data":
            try:
                d = json.loads(body or "{}")
                question = (d.get("question") or "").strip()
                if not question:
                    self._json({"error": "请输入问题"}, 400); return
                answer = self._run_text_to_sql(question)
                self._json({"ok": True, "answer": answer, "route": "data"})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # POST /api/chat —— 通用智能问答（带报销数据上下文）
        if path == "/api/chat":
            try:
                d = json.loads(body or "{}")
                question = (d.get("question") or "").strip()
                history = d.get("history") or []  # [{role,content}, ...]
                # 隐私保护：仅 DEBUG=1 时打印问题（可能含客户名等敏感信息）
                if os.environ.get("DEBUG"):
                    print(f"[chat] 收到问题: {question[:50]} (历史{len(history)}轮)", flush=True)
                if not question:
                    self._json({"error": "请输入问题"}, 400); return

                # 收集数据上下文（让 AI 感知当前数据状态）
                db = get_db()
                ctx_parts = []
                # 1. 当前批次记录数与合计
                try:
                    row = db.execute(
                        "SELECT COUNT(*) c, COALESCE(SUM(CAST(amount AS REAL)),0) s FROM records WHERE batch_name='default'"
                    ).fetchone()
                    ctx_parts.append(f"当前报销批次共 {row['c']} 条记录，金额合计 ¥{row['s']:.2f}")
                except Exception: pass
                # 2. 报销单数与状态分布
                try:
                    rows = db.execute(
                        "SELECT status, COUNT(*) c FROM reimbursements GROUP BY status"
                    ).fetchall()
                    if rows:
                        ctx_parts.append("报销单状态：" + "、".join(f"{r['status']}{r['c']}个" for r in rows))
                except Exception: pass
                # 3. 发票数量
                try:
                    cnt = db.execute("SELECT COUNT(*) c FROM invoices").fetchone()
                    ctx_parts.append(f"发票库共 {cnt['c']} 张发票")
                except Exception: pass
                # 4. 表结构（让 AI 知道有哪些字段可查）
                ctx_parts.append("数据库表：records(报销记录: 发生日期date/费用类别category/费用内容detail/招待客户customer/金额amount等), reimbursements(报销单), invoices(发票), api_profiles(API配置)")
                db.close()
                context = "\n".join(ctx_parts)

                messages = [
                    {
                        "role": "system",
                        "content": (
                            "你是一个报销助手。用户会问关于报销、发票、操作方法或通用问题。"
                            "你有以下当前数据上下文，回答时可参考但不要泄露原始 SQL：\n" + context +
                            "\n\n回答要求：简洁实用，中文。涉及数据时用自然语言描述，不要输出 SQL。"
                            "如果用户的问题明显是查数据（如『本月花了多少』『最近有哪些报销』），"
                            "请回复一句：[ASK_DATA]，系统会改走结构化查询接口，不要硬猜数据。"
                        ),
                    },
                ]
                # 加最近 6 轮历史
                for h in history[-6:]:
                    if h.get("role") in ("user", "assistant") and h.get("content"):
                        messages.append({"role": h["role"], "content": h["content"]})
                messages.append({"role": "user", "content": question})

                r = call_llm(messages, max_tokens=800, temperature=0.4)
                if not r.get("ok"):
                    self._json({"error": r.get("error", "AI 调用失败")}, 500); return
                answer = r["content"]
                # 若 AI 提示要走结构化查询
                if "[ASK_DATA]" in answer:
                    # 改走 Text-to-SQL
                    sql_res = self._run_text_to_sql(question)
                    self._json({"ok": True, "answer": sql_res, "route": "data"})
                else:
                    self._json({"ok": True, "answer": answer, "route": "chat"})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return
        if path == "/api/reimbursements":
            try:
                data = json.loads(body or "{}")
                db = get_db()
                title = data.get("title", "").strip()
                month = str(data.get("month", "")).strip()
                if not title:
                    uname = get_setting("user_name", "报销")
                    title = f"{uname}报销单"
                db.execute(
                    "INSERT INTO reimbursements(title, month, status) VALUES(?,?, 'draft')",
                    (title, month),
                )
                db.commit()
                new_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
                db.close()
                self._json({"ok": True, "id": new_id, "batch_name": f"reimb_{new_id}"})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path == "/api/settings":
            try:
                data = json.loads(body)
                db = get_db()
                for k, v in data.items():
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (k, str(v)))
                db.commit()
                db.close()
                self._json({"ok": True})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path == "/api/recognize":
            try:
                payload = json.loads(body or "{}")
                image = payload.get("image", "")
                if not image:
                    self._json({"error": "缺少 image 参数"}, 400)
                    return
                result = call_vision(image)
                self._json(result)
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path == "/api/records":
            try:
                data = json.loads(body)
                batch = data.get("batch_name", "default")
                records = data.get("records", [])
                # 后端二次防御：过滤掉完全空的记录（前端 rowsToState 已过滤，这里再兜底）
                def _is_empty_rec(r):
                    try:
                        amt = float(r.get("金额") or 0)
                    except (TypeError, ValueError):
                        amt = 0.0
                    return not (
                        amt > 0
                        or str(r.get("招待客户名称", "")).strip()
                        or str(r.get("发生日期", "")).strip()
                        or str(r.get("事前申请编号", "")).strip()
                        or str(r.get("项目编号", "")).strip()
                        or str(r.get("备注", "")).strip()
                    )
                clean = [r for r in records if not _is_empty_rec(r)]
                db = get_db()
                # 用 BEGIN IMMEDIATE 把 DELETE + INSERT 包成原子事务，
                # 避免并发保存导致行重复或丢失
                db.execute("BEGIN IMMEDIATE")
                try:
                    db.execute("DELETE FROM records WHERE batch_name=?", (batch,))
                    for i, r in enumerate(clean):
                        db.execute("""
                            INSERT INTO records(batch_name,seq,date,category,category_detail,customer,approve,project,people_n,count,amount,invoice_type,pay_img,recv_img,merchant,remark)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """, (
                            batch, i+1,
                            r.get("发生日期", ""), r.get("费用类别", ""), r.get("费用类别简称", ""),
                            r.get("招待客户名称", ""), r.get("事前申请编号", ""), r.get("项目编号", ""),
                            r.get("招待人同乘人", ""), r.get("人数"), r.get("金额", 0),
                            r.get("发票类型", ""), r.get("付款截图路径", ""), r.get("水单截图路径", ""),
                            r.get("商家", ""),
                            r.get("备注", "")
                        ))
                    db.commit()
                except Exception:
                    db.execute("ROLLBACK")
                    raise
                finally:
                    try: db.close()
                    except Exception: pass
                self._json({"ok": True, "count": len(clean), "filtered": len(records) - len(clean)})
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        if path == "/api/export":
            try:
                data = json.loads(body)
                month = data.get("month", str(time.localtime().tm_mon))
                batch = data.get("batch_name", "default")

                db = get_db()
                # 导出是只读操作：优先用前端传来的 records，
                # 若为空则从 DB 读取该批次已有记录（避免覆盖丢失数据）
                records = data.get("records", [])
                if not records:
                    rows_db = db.execute(
                        "SELECT * FROM records WHERE batch_name=? ORDER BY seq",
                        (batch,),
                    ).fetchall()
                    records = [{
                        "序号": r["seq"],
                        "发生日期": r["date"],
                        "费用类别": r["category_detail"],
                        "费用类别简称": r["category"],
                        "招待客户名称": r["customer"],
                        "事前申请编号": r["approve"],
                        "项目编号": r["project"],
                        "招待人同乘人": r["people_n"],
                        "人数": r["count"],
                        "金额": r["amount"],
                        "发票类型": r["invoice_type"],
                        "付款截图路径": r["pay_img"],
                        "水单截图路径": r["recv_img"],
                        "备注": r["remark"],
                    } for r in rows_db]
                db.close()

                if not records:
                    self._json({"ok": False, "error": "没有可导出的记录"}, 400)
                    return

                # 同步到 state.json（向后兼容 / 离线脚本）
                # 注意：必须写到 DATA_DIR（~/.reimbursement/），不能写到 APP_DIR，
                # 因为打包后 APP_DIR 指向 _MEIPASS 只读临时目录，重启即丢。
                try:
                    with open(DATA_DIR / "state.json", "w", encoding="utf-8") as f:
                        json.dump(records, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass  # state.json 仅用于离线脚本兼容，写失败不影响导出

                # 直接 import 调用 export（不再 subprocess，PyInstaller 友好）
                import export_expense
                export_res = export_expense.export(
                    month=str(month).replace("月", "") if month else None,
                    records_override=records,
                    return_result=True,
                    exports_dir=EXPORTS_DIR,
                    db_path=DB_PATH,
                )

                ok = export_res.get("ok", False)
                excel_filename = export_res.get("filename", "") if ok else ""
                err_msg = export_res.get("error", "") if not ok else ""

                # 更新报销单状态为已导出（batch 形如 reimb_3）
                if ok and excel_filename and batch.startswith("reimb_"):
                    try:
                        rid = int(batch.split("_", 1)[1])
                        db = get_db()
                        db.execute(
                            """UPDATE reimbursements
                               SET status='exported', exported_at=CURRENT_TIMESTAMP,
                                   excel_filename=?, month=COALESCE(NULLIF(?, ''), month)
                               WHERE id=?""",
                            (excel_filename, str(month).replace("月", ""), rid),
                        )
                        db.commit()
                        db.close()
                    except Exception:
                        pass

                self._json({
                    "ok": ok,
                    "error": err_msg if not ok else None,
                    "excel_filename": excel_filename,
                    # 透传截断信息，前端据此提示用户
                    "truncated": export_res.get("truncated", 0) if ok else 0,
                    "total_input": export_res.get("total_input", 0) if ok else 0,
                })
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        self.send_response(404)
        self.end_headers()

    def do_DELETE(self):
        path = urlparse(self.path).path
        if path == "/api/records":
            qs = parse_qs(urlparse(self.path).query)
            batch = qs.get("batch", ["default"])[0]
            db = get_db()
            db.execute("DELETE FROM records WHERE batch_name=?", (batch,))
            db.commit()
            db.close()
            self._json({"ok": True})
            return

        # DELETE /api/reimbursements/<id> —— 删除报销单及其记录
        if path.startswith("/api/reimbursements/"):
            try:
                rid = int(path.rsplit("/", 1)[-1])
            except ValueError:
                self._json({"error": "无效的报销单 ID"}, 400)
                return
            db = get_db()
            db.execute("DELETE FROM records WHERE batch_name=?", (f"reimb_{rid}",))
            db.execute("DELETE FROM reimbursements WHERE id=?", (rid,))
            db.commit()
            db.close()
            self._json({"ok": True})
            return

        # DELETE /api/invoices/<id> —— 删除发票（DB记录 + 磁盘文件）
        if path.startswith("/api/invoices/"):
            try:
                iid = int(path.rsplit("/", 1)[-1])
            except ValueError:
                self._json({"error": "无效的发票 ID"}, 400)
                return
            db = get_db()
            row = db.execute(
                "SELECT stored_name FROM invoices WHERE id=?", (iid,)
            ).fetchone()
            if row:
                fpath = INVOICES_DIR / row["stored_name"]
                if fpath.is_file():
                    try:
                        fpath.unlink()
                    except OSError:
                        pass
                db.execute("DELETE FROM invoices WHERE id=?", (iid,))
                db.commit()
            db.close()
            self._json({"ok": True})
            return

        # DELETE /api/profiles/<id> —— 删除 Profile
        if path.startswith("/api/profiles/"):
            try:
                pid = int(path.rsplit("/", 1)[-1])
            except ValueError:
                self._json({"error": "无效的 Profile ID"}, 400); return
            db = get_db()
            row = db.execute("SELECT is_active FROM api_profiles WHERE id=?", (pid,)).fetchone()
            if row and row["is_active"]:
                db.close()
                self._json({"error": "不能删除当前激活的 Profile"}, 400); return
            db.execute("DELETE FROM api_profiles WHERE id=?", (pid,))
            db.commit()
            db.close()
            self._json({"ok": True})
            return

        self.send_response(404)
        self.end_headers()

    def do_PATCH(self):
        # PATCH /api/invoices/<id> —— 修改发票备注
        path = urlparse(self.path).path
        body = self._read_body()
        if path.startswith("/api/invoices/"):
            try:
                iid = int(path.rsplit("/", 1)[-1])
                data = json.loads(body or "{}")
                remark = str(data.get("remark", ""))
                db = get_db()
                db.execute("UPDATE invoices SET remark=? WHERE id=?", (remark, iid))
                db.commit()
                db.close()
                self._json({"ok": True})
            except ValueError:
                self._json({"error": "无效的发票 ID"}, 400)
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        # PATCH /api/profiles/<id> —— 更新 Profile 字段
        if path.startswith("/api/profiles/"):
            try:
                pid = int(path.rsplit("/", 1)[-1])
                d = json.loads(body or "{}")
                fields = []
                vals = []
                for k in ("name", "provider", "endpoint", "api_key", "model"):
                    if k in d and d[k] is not None:
                        fields.append(f"{k}=?")
                        vals.append(str(d[k]).strip() if k != "name" else d[k].strip())
                if not fields:
                    self._json({"error": "无可更新字段"}, 400); return
                vals.append(pid)
                db = get_db()
                try:
                    db.execute(f"UPDATE api_profiles SET {','.join(fields)} WHERE id=?", vals)
                    db.commit()
                except sqlite3.IntegrityError:
                    db.close()
                    self._json({"error": "名称已存在"}, 400); return
                # 若是激活的，同步 settings
                row = db.execute("SELECT * FROM api_profiles WHERE id=? AND is_active=1", (pid,)).fetchone()
                if row:
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_key',?)", (row["api_key"],))
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_endpoint',?)", (row["endpoint"],))
                    db.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('api_model',?)", (row["model"],))
                    db.commit()
                db.close()
                self._json({"ok": True})
            except ValueError:
                self._json({"error": "无效的 Profile ID"}, 400)
            except Exception as e:
                self._json({"error": str(e)}, 500)
            return

        self.send_response(404)
        self.end_headers()


# ═══════════════════════════════════
# 主入口
# ═══════════════════════════════════

def _setup_app_logging():
    """在 .app（frozen）模式下，把 print/stderr 重定向到日志文件，便于排查问题"""
    if not getattr(sys, 'frozen', False):
        return  # 命令行模式不需要
    log_dir = Path.home() / ".reimbursement"
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / "app.log"
    # 重定向 stdout/stderr 到日志文件（追加模式）
    sys.stdout = open(log_file, "a", buffering=1, encoding="utf-8")
    sys.stderr = sys.stdout
    print(f"\n{'='*50}\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] 应用启动")


def _acquire_single_instance_lock():
    """单实例锁：已运行则返回 False。frozen 状态用文件锁，命令行模式跳过"""
    if not getattr(sys, 'frozen', False):
        return True
    import fcntl
    lock_file = DATA_DIR / ".app.lock"
    lock_file.parent.mkdir(exist_ok=True)
    try:
        global _lock_fp
        _lock_fp = open(lock_file, "w")
        fcntl.flock(_lock_fp, fcntl.LOCK_EX | fcntl.LOCK_NB)
        _lock_fp.write(str(os.getpid()))
        _lock_fp.flush()
        return True
    except (IOError, OSError):
        # 锁被占用。检查锁持有进程是否还活着，若已死则清理后重试
        try:
            old_pid = lock_file.read_text().strip()
            if old_pid.isdigit():
                os.kill(int(old_pid), 0)  # 检查进程是否存在（不发信号）
                # 进程还活着 → 真的有实例在跑
                return False
        except (ProcessLookupError, FileNotFoundError, ValueError, OSError):
            # 进程已死或读不到 → 锁是残留的，清理后重试
            pass
        # 清理残留锁，重试一次
        try:
            if '_lock_fp' in dir() or '_lock_fp' in globals():
                pass
        except Exception:
            pass
        try:
            _lock_fp = open(lock_file, "w")
            fcntl.flock(_lock_fp, fcntl.LOCK_EX | fcntl.LOCK_NB)
            _lock_fp.write(str(os.getpid()))
            _lock_fp.flush()
            print("清理了残留锁，重新获取成功")
            return True
        except (IOError, OSError):
            return False


def _start_server_thread():
    """启动 HTTP 服务到守护线程"""
    url = f"http://{HOST}:{PORT}"
    try:
        server = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"端口 {PORT} 已被占用，请关闭已有实例或选择其他端口。")
            return None, url
        else:
            raise

    import threading
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    return server, url


def _wait_for_server(url, timeout=10):
    """等服务就绪"""
    for _ in range(timeout * 2):
        try:
            urllib.request.urlopen(url + "/api/config", timeout=1)
            return True
        except Exception:
            time.sleep(0.5)
    return False


def _launch_native_window(url):
    """用 pywebview 打开原生窗口（替代浏览器）"""
    try:
        import webview
    except ImportError:
        # 没装 pywebview，回退浏览器
        print("pywebview 未安装，回退浏览器模式")
        webbrowser.open(url)
        # 命令行模式：保持运行
        try:
            while True: time.sleep(1)
        except KeyboardInterrupt:
            pass
        return

    # 等服务就绪
    _wait_for_server(url)

    api_ready = bool(get_setting("api_key"))
    title = f"智能报销系统" + ("" if api_ready else " · 未配置 API Key")

    def _on_close():
        # 关窗即彻底退出：清理锁文件 + 退出进程
        print("窗口已关闭，应用退出")
        try:
            (DATA_DIR / ".app.lock").unlink(missing_ok=True)
        except Exception:
            pass
        os._exit(0)

    # 创建原生窗口
    webview.create_window(
        title=title,
        url=url,
        width=1440,
        height=900,
        min_size=(1100, 700),
        text_select=True,
        on_top=False,
    )
    # 启动 GUI 主循环（阻塞，关闭窗口后 _on_close 触发退出）
    webview.start(debug=False)


def main():
    _setup_app_logging()
    os.chdir(APP_DIR)
    init_db()

    # 启动前清理可能的残留锁文件（避免历史锁导致无法启动）
    try:
        (DATA_DIR / ".app.lock").unlink(missing_ok=True)
    except Exception:
        pass

    # 启动 HTTP 服务（线程）。端口被占 = 已有实例 → 不再开新窗口
    server, url = _start_server_thread()
    if server is None:
        print("已有实例运行（端口被占），本次启动退出")
        return

    api_ready = bool(get_setting("api_key"))
    model = get_setting("api_model")
    print("=" * 50)
    print("  智能报销录入系统 v2.0")
    print(f"  地址: {url}")
    print(f"  模型: {model}  |  Key: {'已配置 ✓' if api_ready else '未配置 ⚠'}")
    print(f"  数据目录: {DATA_DIR}")
    print("=" * 50)

    # 启动原生窗口（主线程，阻塞）
    _launch_native_window(url)


if __name__ == "__main__":
    main()
