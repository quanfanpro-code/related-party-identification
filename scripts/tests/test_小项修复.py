# -*- coding: utf-8 -*-
# 批次五小项修复的行为测试：秘密扫描正则、登录轮询降速、页面漂移上限、
# 客户端缓存只在业务成功时写入、Retry-After 小写兼容、登录状态原子写。
# 注意：假 token 必须拼接构造，原样写进文件会被项目自身的秘密扫描扫出。
import inspect
import json
import tempfile
import unittest
from pathlib import Path

from scripts import sync_upstream
from scripts.cicpa import auth
from scripts.cicpa.browser_transport import OpenCliBridgeError, OpenCliTransport
from scripts.cicpa.client import CicpaClient

FAKE_TOKEN = "abcd1234" + "efgh5678" + "ijkl90"  # 22 位假 token，拼接以避免触发秘密扫描


class Test秘密扫描(unittest.TestCase):
    def test_假token能被正则命中(self):
        text = json.dumps({"cicpa_token": FAKE_TOKEN}, ensure_ascii=False)
        self.assertTrue(any(p.search(text) for p in sync_upstream.SECRET_PATTERNS),
                        "双重转义导致正则失效")

    def test_普通文本不误报(self):
        self.assertFalse(any(p.search('{"note": "今天天气不错"}') for p in sync_upstream.SECRET_PATTERNS))

    def test_扫描后缀包含py(self):
        self.assertIn(".py", sync_upstream.SECRET_SCAN_SUFFIXES)

    def test_临时目录py文件里的假token被扫出(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            (root / "scripts").mkdir()
            (root / "scripts" / "leak.py").write_text(
                'cicpa_token = "' + FAKE_TOKEN + '"', encoding="utf-8")
            checks = sync_upstream._default_checks(root)
            self.assertFalse(checks["secret_leaks"], ".py 里的假 token 没被扫出来")

    def test_敏感变更分类的正则能命中(self):
        with tempfile.TemporaryDirectory() as d:
            staging = Path(d)
            (staging / "upstream").mkdir()
            (staging / "upstream" / "a.py").write_text("requests.post(url)", encoding="utf-8")
            (staging / "upstream" / "b.md").write_text("导出维度结果", encoding="utf-8")
            categories = sync_upstream._classify_sensitive_changes(staging, ["a.py", "b.md"])
            self.assertIn("api", categories, "requests. 调用没被识别为 api 类敏感变更")
            self.assertIn("export", categories)


class Test浏览器传输(unittest.TestCase):
    def test_登录轮询默认5秒(self):
        signature = inspect.signature(OpenCliTransport.wait_for_login)
        self.assertEqual(signature.parameters["poll"].default, 5.0)

    def test_页面漂移重试上限2次后报错(self):
        transport = OpenCliTransport.__new__(OpenCliTransport)
        transport._eval = lambda *a, **k: json.dumps({"nav": True})
        opened = []
        transport.open_page = lambda url: opened.append(url)
        with self.assertRaises(OpenCliBridgeError):
            transport.request("GET", "https://zsk-cmis.cicpa.org.cn/open/x")
        self.assertEqual(len(opened), 2, "漂移重试应正好打开页面 2 次后放弃")


class FakeResponse:
    def __init__(self, payload, status=200, headers=None):
        self._payload = payload
        self.status_code = status
        self.headers = headers or {}

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, response):
        self.response = response
        self.calls = 0

    def request(self, *args, **kwargs):
        self.calls += 1
        return self.response


class NoWaitLimiter:
    def wait(self, kind):
        return 0.0


class Test客户端缓存(unittest.TestCase):
    def test_业务错误载荷不进缓存(self):
        session = FakeSession(FakeResponse({"status_code": 50001, "status_msg": "业务错误"}))
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=lambda s: None)
        client.request_json("GET", "https://zsk-cmis.cicpa.org.cn/x", cache_key="k")
        client.request_json("GET", "https://zsk-cmis.cicpa.org.cn/x", cache_key="k")
        self.assertEqual(session.calls, 2, "业务错误被缓存，第二次没有真正请求")

    def test_业务成功进缓存(self):
        session = FakeSession(FakeResponse({"status_code": 0, "data": {}}))
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=lambda s: None)
        client.request_json("GET", "https://zsk-cmis.cicpa.org.cn/x", cache_key="k")
        client.request_json("GET", "https://zsk-cmis.cicpa.org.cn/x", cache_key="k")
        self.assertEqual(session.calls, 1)

    def test_Retry_After小写键兼容(self):
        response = FakeResponse({}, headers={"retry-after": "7"})
        self.assertEqual(CicpaClient._retry_after_seconds(response), 7.0)


class Test登录状态写入(unittest.TestCase):
    def test_写入后可回读且目录自动创建(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "sub" / "login.json"
            auth.write_login_status(path, "authenticated", "ok")
            payload = auth.read_login_status(path)
            self.assertEqual(payload["status"], "authenticated")


if __name__ == "__main__":
    unittest.main()
