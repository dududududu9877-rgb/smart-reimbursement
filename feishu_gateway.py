#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能报销系统 · 飞书网关
======================
独立轻量服务，零侵入接入报销系统。员工在飞书发报销截图 → 自动AI识别 → 回复卡片 → 入库。

工作流：
  飞书事件 → /webhook/event 验签 → 下载图片 → 调报销系统 /api/recognize → 回复卡片 → 入库

依赖：仅 Python 标准库（无需 pip install）
启动：python3 feishu_gateway.py
配置：feishu_config.json（从 feishu_config.example.json 复制后填写）
"""

import base64
import hashlib
import hmac
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

# ============================================================
# 配置加载
# ============================================================
HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "feishu_config.json")

def load_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"❌ 找不到配置文件 {CONFIG_PATH}")
        print("   请复制 feishu_config.example.json 为 feishu_config.json 并填写飞书应用凭证")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

CFG = load_config()
APP_ID = CFG.get("app_id", "")
APP_SECRET = CFG.get("app_secret", "")
ENCRYPT_KEY = CFG.get("encrypt_key", "")  # 事件订阅加密策略 Key，可为空
VERIFY_TOKEN = CFG.get("verification_token", "")
GATEWAY_PORT = int(CFG.get("gateway_port", 9900))
# 报销系统地址（本地）
BAOXIAO_BASE = CFG.get("baoxiao_base", "http://127.0.0.1:8765")
# 是否识别后自动入库（写入 data.db 的 feishu_<日期> 批次）
AUTO_SAVE = bool(CFG.get("auto_save", True))
# 只响应这些用户的开頭（open_id 前缀）；空=所有人可用，建议测试时限定自己
ALLOW_OPENID_PREFIXES = CFG.get("allow_openid_prefixes", [])

# ============================================================
# 飞书 AES-256-CBC 解密（事件加密用，Encrypt Key 非空时需要）
# ============================================================
def aes_decrypt(encrypt: str, key: str) -> str:
    """飞书事件加密：AES-256-CBC，key=SHA256(key)[:32]，IV=密文前16字节"""
    try:
        from Crypto.Cipher import AES  # pycryptodome
        from Crypto.Util.Padding import unpad
    except ImportError:
        # 退而用标准库？Python 标准库无 AES。提示安装。
        print("⚠️  事件启用了加密，但缺少 pycryptodome。请运行: pip3 install pycryptodome")
        print("   或者到飞书后台「事件订阅」把加密策略改为「不加密」(明文模式，VERIFY_TOKEN 验签即可)")
        raise
    key_bytes = hashlib.sha256(key.encode("utf-8")).digest()  # AES-256 → 32字节key
    cipher_data = base64.b64decode(encrypt)
    iv, encrypted = cipher_data[:16], cipher_data[16:]
    cipher = AES.new(key_bytes, AES.MODE_CBC, iv)
    decrypted = unpad(cipher.decrypt(encrypted), AES.block_size)
    return decrypted.decode("utf-8")

# ============================================================
# 飞书 API 调用（token 缓存 + 图片下载 + 发消息）
# ============================================================
_token_cache = {"token": "", "expire_at": 0}

def get_tenant_access_token() -> str:
    """获取并缓存 tenant_access_token"""
    if _token_cache["token"] and time.time() < _token_cache["expire_at"] - 60:
        return _token_cache["token"]
    body = json.dumps({"app_id": APP_ID, "app_secret": APP_SECRET}).encode("utf-8")
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if data.get("code") != 0:
        raise RuntimeError(f"获取 token 失败: {data}")
    _token_cache["token"] = data["tenant_access_token"]
    _token_cache["expire_at"] = time.time() + data.get("expire", 7200)
    return _token_cache["token"]

def download_message_resource(message_id: str, message_type: str, file_key: str) -> bytes:
    """
    下载飞书消息中的图片/文件资源。
    飞书接口：GET /open-apis/im/v1/messages/{message_id}/resources/{file_key}?type=image
    """
    token = get_tenant_access_token()
    url = (f"https://open.feishu.cn/open-apis/im/v1/messages/{urllib.parse.quote(message_id)}"
           f"/resources/{urllib.parse.quote(file_key)}?type=image")
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"}, method="GET")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()

def send_message(receive_id: str, receive_id_type: str, msg_type: str, content: str):
    """发送消息。content 是 JSON 字符串。"""
    token = get_tenant_access_token()
    body = json.dumps({
        "receive_id": receive_id,
        "msg_type": msg_type,
        "content": content,
    }, ensure_ascii=False).encode("utf-8")
    url = ("https://open.feishu.cn/open-apis/im/v1/messages"
           f"?receive_id_type={receive_id_type}")
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }, method="POST")
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))

def reply_message(message_id: str, msg_type: str, content: str):
    """回复指定消息（用 reply 接口，更自然）。"""
    token = get_tenant_access_token()
    body = json.dumps({"msg_type": msg_type, "content": content}, ensure_ascii=False).encode("utf-8")
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{urllib.parse.quote(message_id)}/reply"
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }, method="POST")
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))

# ============================================================
# 调用报销系统识别接口（复用现有 /api/recognize）
# ============================================================
def recognize_image(data_uri: str) -> dict:
    """调报销系统的识别接口，返回字段 dict"""
    body = json.dumps({"image": data_uri}).encode("utf-8")
    req = urllib.request.Request(
        BAOXIAO_BASE.rstrip("/") + "/api/recognize",
        data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode("utf-8"))

def save_record(fields: dict, source: str = "feishu") -> bool:
    """把识别结果写入报销系统 data.db（调 /api/records）"""
    batch = f"feishu_{time.strftime('%Y%m%d')}"
    record = {
        "发生日期": fields.get("发生日期", ""),
        "费用类别": fields.get("费用类别", ""),
        "费用类别简称": fields.get("费用类别简称", ""),
        "招待客户名称": fields.get("招待客户名称", ""),
        "事前申请编号": fields.get("事前申请编号", ""),
        "项目编号": fields.get("项目编号", ""),
        "招待人同乘人": fields.get("招待人同乘人", ""),
        "人数": fields.get("人数", ""),
        "金额": fields.get("金额", 0),
        "发票类型": fields.get("发票类型", ""),
        "商家": fields.get("商家", ""),
        "备注": f"[飞书识别] {fields.get('备注', '')}".strip(),
    }
    body = json.dumps({"batch_name": batch, "records": [record]}).encode("utf-8")
    req = urllib.request.Request(
        BAOXIAO_BASE.rstrip("/") + "/api/records",
        data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("ok", False)
    except Exception as e:
        print(f"  ⚠️ 入库失败: {e}")
        return False

# ============================================================
# 识别结果 → 飞书交互卡片
# ============================================================
def build_card(fields: dict, saved: bool, error: str = "") -> str:
    """构造飞书消息卡片 JSON 字符串"""
    def val(k):
        v = fields.get(k, "")
        return v if v not in (None, "") else "—"
    amount = val("金额")
    category = val("费用类别")
    customer = val("招待客户名称")
    date = val("发生日期")
    merchant = val("商家")
    people = val("招待人同乘人")

    if error:
        content_lines = [
            {"tag": "div", "text": {"tag": "lark_md",
                "content": f"❌ **识别失败**\n{error}"}}
        ]
    else:
        status = "✅ 已录入报销系统" if saved else "⚠️ 已识别，但入库失败（请查桌面端）"
        content_lines = [
            {"tag": "div", "text": {"tag": "lark_md",
                "content": f"📊 **报销信息识别结果**\n{s}"}},
            {"tag": "hr"},
            {"tag": "div", "fields": [
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**金额**\n{amount}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**费用类别**\n{category}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**发生日期**\n{date}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**招待客户**\n{customer}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**商家**\n{merchant}"}},
                {"is_short": True, "text": {"tag": "lark_md", "content": f"**招待/同乘人**\n{people}"}},
            ]},
        ]
    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "title": {"tag": "plain_text", "content": "智能报销识别"},
            "template": "red" if error else ("green" if saved else "orange")
        },
        "elements": content_lines
    }
    return json.dumps(card, ensure_ascii=False)

# ============================================================
# 消息去重（同一 message_id 只处理一次）
# ============================================================
_processed = set()
MAX_CACHE = 500

def is_processed(msg_id: str) -> bool:
    if msg_id in _processed:
        return True
    _processed.add(msg_id)
    if len(_processed) > MAX_CACHE:
        # 简单清理：保留后半部分
        _processed.clear()
        _processed.update(list(_processed)[-MAX_CACHE//2:])
    return False

# ============================================================
# 处理一条图片消息
# ============================================================
def handle_image_message(message_id: str, image_key: str, receive_id: str, receive_id_type: str, sender_open_id: str):
    """下载图片 → 识别 → 入库 → 回复卡片"""
    print(f"[{time.strftime('%H:%M:%S')}] 处理图片 message_id={message_id[:20]}... key={image_key}")
    # 权限过滤
    if ALLOW_OPENID_PREFIXES:
        if not any(sender_open_id.startswith(p) for p in ALLOW_OPENID_PREFIXES):
            send_message(receive_id, receive_id_type, "text",
                         json.dumps({"text": "抱歉，您没有使用权限。"}, ensure_ascii=False))
            return
    # 1. 下载图片
    try:
        img_bytes = download_message_resource(message_id, "image", image_key)
        print(f"  ✓ 下载图片 {len(img_bytes)} 字节")
    except Exception as e:
        reply_text(message_id, f"下载图片失败: {e}")
        return
    # 2. 转 data URI
    data_uri = "data:image/jpeg;base64," + base64.b64encode(img_bytes).decode("ascii")
    # 3. 识别
    try:
        result = recognize_image(data_uri)
        print(f"  ✓ 识别返回: {str(result)[:120]}")
    except Exception as e:
        reply_card(message_id, {}, False, error=f"识别服务异常: {e}")
        return
    if isinstance(result, dict) and result.get("error"):
        reply_card(message_id, {}, False, error=result["error"])
        return
    # 4. 入库
    saved = False
    if AUTO_SAVE:
        try:
            saved = save_record(result)
            print(f"  {'✓ 入库成功' if saved else '✗ 入库失败'}")
        except Exception as e:
            print(f"  ⚠️ 入库异常: {e}")
    # 5. 回复卡片
    reply_card(message_id, result, saved)

def reply_card(message_id, fields, saved, error=""):
    try:
        card_json = build_card(fields, saved, error)
        reply_message(message_id, "interactive", card_json)
    except Exception as e:
        print(f"  ⚠️ 发送卡片失败({e})，改发纯文本")
        try:
            txt = error or json.dumps(fields, ensure_ascii=False, indent=2)[:1500]
            reply_message(message_id, "text", json.dumps({"text": txt}, ensure_ascii=False))
        except Exception as e2:
            print(f"  ✗ 纯文本也失败: {e2}")

def reply_text(message_id, text):
    try:
        reply_message(message_id, "text", json.dumps({"text": text}, ensure_ascii=False))
    except Exception as e:
        print(f"  ⚠️ 回复失败: {e}")

# ============================================================
# HTTP 服务 + 飞书事件处理
# ============================================================
class FeishuHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # 静默默认日志，自己 print

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        # 健康检查
        if self.path in ("/", "/health"):
            self._send_json({"ok": True, "service": "feishu-gateway", "time": int(time.time())})
            return
        self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b""

        if path == "/webhook/event":
            self._handle_event(raw)
        else:
            self._send_json({"error": "not found"}, 404)

    def _handle_event(self, raw: bytes):
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            self._send_json({"error": "invalid json"}, 400)
            return

        # 情况1：URL 验证（飞书配置回调地址时发的 challenge）
        if "challenge" in payload and ("type" in payload and payload["type"] == "url_verification"):
            print("[验证] 收到 URL 验证请求，返回 challenge")
            self._send_json({"challenge": payload["challenge"]})
            return

        # 情况2：加密事件（encrypt 字段）
        if "encrypt" in payload:
            if not ENCRYPT_KEY:
                print("⚠️ 收到加密事件但未配置 encrypt_key")
                self._send_json({"error": "encrypt_key missing"}, 400)
                return
            try:
                decrypted = aes_decrypt(payload["encrypt"], ENCRYPT_KEY)
                payload = json.loads(decrypted)
            except Exception as e:
                print(f"⚠️ 解密失败: {e}")
                self._send_json({"error": "decrypt failed"}, 400)
                return

        # 飞书 v2 事件结构：{ "schema": "2.0", "header": {...}, "event": {...} }
        header = payload.get("header", {})
        event = payload.get("event", payload)  # 兼容旧版

        # 情况3：URL 验证（v2 schema 也有可能是这种）
        if header.get("event_type") == "url_verification" or "challenge" in payload:
            self._send_json({"challenge": payload.get("challenge", "")})
            return

        # 先返回 200，避免飞书重试；后续异步处理
        self._send_json({"code": 0})

        # 验 token
        token = header.get("token", "")
        if VERIFY_TOKEN and token != VERIFY_TOKEN:
            print(f"⚠️ token 校验失败: {token[:10]}...")
            return

        event_type = header.get("event_type", "")
        if event_type != "im.message.receive_v1":
            return  # 非消息事件，忽略

        # 解析消息
        message = event.get("message", {})
        msg_type = message.get("message_type", "")
        message_id = message.get("message_id", "")
        if not message_id or is_processed(message_id):
            return

        sender = event.get("sender", {})
        sender_open_id = sender.get("sender_id", {}).get("open_id", "")

        # chat 信息（用于主动发消息，receive_id）
        chat_id = message.get("chat_id", "")
        receive_id = chat_id or sender_open_id
        receive_id_type = "chat_id" if chat_id else "open_id"

        # 只处理图片
        if msg_type != "image":
            reply_text(message_id, "请发送报销截图/发票图片，我会自动识别 📸")
            return

        # 取 image_key
        try:
            content = json.loads(message.get("content", "{}"))
            image_key = content.get("image_key", "")
        except Exception:
            image_key = ""
        if not image_key:
            reply_text(message_id, "未识别到图片内容")
            return

        # 异步处理（这里直接同步，单条延迟可接受；飞书已收到200不会重试）
        try:
            handle_image_message(message_id, image_key, receive_id, receive_id_type, sender_open_id)
        except Exception as e:
            print(f"✗ 处理异常: {e}")
            try:
                reply_text(message_id, f"处理出错: {e}")
            except Exception:
                pass


# ============================================================
# 启动
# ============================================================
def main():
    # 启动前自检
    if not APP_ID or not APP_SECRET:
        print("❌ 配置缺少 app_id 或 app_secret，请检查 feishu_config.json")
        sys.exit(1)
    print("═" * 50)
    print("  智能报销系统 · 飞书网关")
    print("═" * 50)
    print(f"  监听端口      : {GATEWAY_PORT}")
    print(f"  报销系统地址  : {BAOXIAO_BASE}")
    print(f"  自动入库      : {'开' if AUTO_SAVE else '关'}")
    print(f"  事件加密      : {'开 (需 pycryptodome)' if ENCRYPT_KEY else '关 (明文+token验签)'}")
    print(f"  App ID        : {APP_ID[:8]}...")
    print(f"  回调地址(填飞书后台): http://<你的公网穿透URL>:{GATEWAY_PORT}/webhook/event")
    print("─" * 50)
    # 检查报销系统是否在跑
    try:
        urllib.request.urlopen(BAOXIAO_BASE.rstrip("/") + "/api/config", timeout=3).read()
        print("  ✓ 报销系统已连接")
    except Exception:
        print(f"  ⚠️ 连不上报销系统 {BAOXIAO_BASE}，请先启动 serve.py (端口 8765)")
    print("═" * 50)
    server = ThreadingHTTPServer(("0.0.0.0", GATEWAY_PORT), FeishuHandler)
    print(f"\n🚀 网关已启动，等待飞书消息...\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
        server.shutdown()


if __name__ == "__main__":
    main()
