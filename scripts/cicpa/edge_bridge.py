# 关联方识别与核查 — Edge/Chrome OpenCLI 本地桥接
# 本文件改编自 jackwener/OpenCLI(extension)(Apache-2.0)
# 上游项目: OpenCLI | 仓库: https://gitee.com/github_dep/opencli
# 上游协议: Apache-2.0
#
# 本地修改: CPA-Q(quanfanpro-code)
# 本文件的修改部分同样以 Apache-2.0 协议发布,与上游保持一致。
#"""通过 OpenCLI 浏览器扩展读取当前 Edge 或 Chrome 登录态。"""

from __future__ import annotations

import base64
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import struct
import threading
from typing import Callable, Dict, Optional
import uuid


OPENCLI_HOST = "127.0.0.1"
OPENCLI_PORT = 19825
ZSK_URL = "https://zsk-cmis.cicpa.org.cn/"
_WEBSOCKET_GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"


class OpenCliBridgeError(RuntimeError):
    """OpenCLI 扩展连接或返回内容无效。"""


def _read_exact(stream, length: int) -> bytes:
    chunks = bytearray()
    while len(chunks) < length:
        chunk = stream.read(length - len(chunks))
        if not chunk:
            raise ConnectionError("OpenCLI 扩展连接已断开")
        chunks.extend(chunk)
    return bytes(chunks)


def _read_frame(stream) -> tuple[int, bytes]:
    first, second = _read_exact(stream, 2)
    opcode = first & 0x0F
    length = second & 0x7F
    if length == 126:
        length = struct.unpack("!H", _read_exact(stream, 2))[0]
    elif length == 127:
        length = struct.unpack("!Q", _read_exact(stream, 8))[0]
    mask = _read_exact(stream, 4) if second & 0x80 else b""
    payload = bytearray(_read_exact(stream, length))
    if mask:
        for index in range(length):
            payload[index] ^= mask[index % 4]
    return opcode, bytes(payload)


def _frame(opcode: int, payload: bytes) -> bytes:
    first = 0x80 | opcode
    length = len(payload)
    if length < 126:
        header = bytes((first, length))
    elif length <= 0xFFFF:
        header = bytes((first, 126)) + struct.pack("!H", length)
    else:
        header = bytes((first, 127)) + struct.pack("!Q", length)
    return header + payload


class _BridgeState:
    def __init__(self):
        self.command_id = "rpi-cookies-" + uuid.uuid4().hex
        self.cookies: Dict[str, str] = {}
        self.error = ""
        self.done = threading.Event()

    def command(self) -> Dict[str, str]:
        return {
            "id": self.command_id,
            "session": "related-party-identification",
            "action": "cookies",
            "url": ZSK_URL,
        }

    def accept(self, payload: object) -> None:
        if not isinstance(payload, dict) or payload.get("id") != self.command_id:
            return
        if not payload.get("ok"):
            self.error = str(payload.get("error") or "OpenCLI 未能读取登录状态")
            self.done.set()
            return
        data = payload.get("data")
        if not isinstance(data, list):
            self.error = "OpenCLI 返回的登录状态格式无效"
            self.done.set()
            return
        for item in data:
            if not isinstance(item, dict):
                continue
            name = item.get("name")
            value = item.get("value")
            domain = str(item.get("domain", "")).lstrip(".")
            if (
                isinstance(name, str)
                and isinstance(value, str)
                and domain.endswith("cicpa.org.cn")
            ):
                self.cookies[name] = value
        self.done.set()


class _BridgeServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, address, state: _BridgeState):
        self.state = state
        super().__init__(address, _BridgeHandler)


class _BridgeHandler(BaseHTTPRequestHandler):
    server: _BridgeServer

    def log_message(self, _format, *_args) -> None:
        return

    def _plain_response(self, body: bytes, content_type: str = "text/plain") -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if self.path == "/ping":
            self._plain_response(b"pong")
            return
        if self.path == "/status":
            body = json.dumps(
                {"daemonVersion": "related-party-identification"},
                separators=(",", ":"),
            ).encode("utf-8")
            self._plain_response(body, "application/json")
            return
        if self.path != "/ext" or self.headers.get("Upgrade", "").lower() != "websocket":
            self.send_error(404)
            return

        key = self.headers.get("Sec-WebSocket-Key")
        if not key:
            self.send_error(400)
            return
        accept = base64.b64encode(
            hashlib.sha1((key + _WEBSOCKET_GUID).encode("ascii")).digest()
        ).decode("ascii")
        self.send_response(101, "Switching Protocols")
        self.send_header("Upgrade", "websocket")
        self.send_header("Connection", "Upgrade")
        self.send_header("Sec-WebSocket-Accept", accept)
        self.end_headers()
        self.wfile.flush()

        try:
            while not self.server.state.done.is_set():
                opcode, raw = _read_frame(self.rfile)
                if opcode == 0x8:
                    break
                if opcode == 0x9:
                    self.wfile.write(_frame(0xA, raw))
                    self.wfile.flush()
                    continue
                if opcode != 0x1:
                    continue
                payload = json.loads(raw.decode("utf-8"))
                if isinstance(payload, dict) and payload.get("type") == "hello":
                    command = json.dumps(
                        self.server.state.command(),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ).encode("utf-8")
                    self.wfile.write(_frame(0x1, command))
                    self.wfile.flush()
                    continue
                self.server.state.accept(payload)
        except (ConnectionError, json.JSONDecodeError, OSError, UnicodeError) as exc:
            if not self.server.state.done.is_set():
                self.server.state.error = "OpenCLI 连接中断：{}".format(
                    type(exc).__name__
                )
                self.server.state.done.set()


def collect_edge_cookies(
    *,
    timeout: float = 45.0,
    on_ready: Optional[Callable[[], None]] = None,
    host: str = OPENCLI_HOST,
    port: int = OPENCLI_PORT,
) -> Dict[str, str]:
    """等待已安装的 OpenCLI 扩展并取得当前注协 Cookie。"""
    state = _BridgeState()
    try:
        server = _BridgeServer((host, port), state)
    except OSError as exc:
        raise OpenCliBridgeError(
            "OpenCLI 连接端口被其他程序占用，请关闭已运行的 OpenCLI 后重试"
        ) from exc
    thread = threading.Thread(
        target=server.serve_forever,
        name="related-party-opencli-bridge",
        daemon=True,
    )
    thread.start()
    try:
        if on_ready is not None:
            on_ready()
        if not state.done.wait(timeout):
            raise OpenCliBridgeError(
                "未等到 OpenCLI 扩展连接，请确认扩展已启用后重试"
            )
        if state.error:
            raise OpenCliBridgeError(state.error)
        if not state.cookies.get("cicpa_token") or not state.cookies.get("XSRF-TOKEN"):
            raise OpenCliBridgeError(
                "Edge 或 Chrome 尚未完成注协登录，请登录并进入行业执业知识库"
            )
        return state.cookies
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=2.0)
