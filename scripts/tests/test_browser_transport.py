# 关联方识别与核查 — OpenCLI 页面取数通道测试
#"""browser_transport 模块的单元测试,全部经注入假 runner,不联网。"""

import base64
import json
from types import SimpleNamespace
import unittest

from scripts.cicpa.browser_transport import (
    OPENCLI_SESSION,
    OpenCliBridgeError,
    OpenCliTransport,
    locate_cli,
    parse_profile_list_output,
)


class FakeRunner:
    """记录调用并按序返回预设 stdout 的假 CLI 运行器。"""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def __call__(self, args, timeout=None):
        self.calls.append(list(args))
        item = self.responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return SimpleNamespace(stdout=item, returncode=0, stderr="")


def _cli_payload(status, body=b"", headers=None):
    return json.dumps(
        {"s": status, "h": headers or {}, "b": base64.b64encode(body).decode("ascii")},
        ensure_ascii=False,
    )


class LocateCliTests(unittest.TestCase):
    def test_找不到opencli时给出安装指引(self):
        with self.assertRaises(OpenCliBridgeError) as ctx:
            locate_cli(which=lambda name: None)
        message = str(ctx.exception)
        self.assertIn("opencli", message)
        self.assertIn("安装", message)

    def test_找到opencli时返回路径(self):
        self.assertEqual(locate_cli(which=lambda name: "C:/bin/opencli.cmd"), "C:/bin/opencli.cmd")


class ProfileListTests(unittest.TestCase):
    SAMPLE = (
        "Connected Browser Bridge profiles\n"
        "\n"
        "  vzymxgmh — connected v1.0.24\n"
    )

    def test_样例输出解析出连接中的配置名(self):
        self.assertEqual(parse_profile_list_output(self.SAMPLE), ["vzymxgmh"])

    def test_空输出返回空列表(self):
        self.assertEqual(parse_profile_list_output("No connected profiles\n"), [])

    def test_多个连接配置都能列出(self):
        text = self.SAMPLE.replace("vzymxgmh", "aaa") + "  bbb — connected v1.0.24\n"
        self.assertEqual(parse_profile_list_output(text), ["aaa", "bbb"])


class TransportProfileTests(unittest.TestCase):
    def test_显式配置优先且不查询profile列表(self):
        runner = FakeRunner([_cli_payload(200, b"{}")])
        transport = OpenCliTransport(profile="fixed", runner=runner)
        transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        joined = " ".join(" ".join(call) for call in runner.calls)
        self.assertNotIn("profile", joined.split("--profile")[1][:2])
        self.assertIn("--profile", joined)

    def test_单个连接配置自动采用(self):
        runner = FakeRunner(
            [
                "Connected Browser Bridge profiles\n  vzymxgmh — connected v1.0.24\n",
                _cli_payload(200, b"{}"),
            ]
        )
        transport = OpenCliTransport(runner=runner, cli_path="opencli")
        transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        self.assertIn("vzymxgmh", " ".join(runner.calls[1]))

    def test_没有连接配置时报错并给指引(self):
        runner = FakeRunner(["No connected profiles\n"])
        transport = OpenCliTransport(runner=runner, cli_path="opencli")
        with self.assertRaises(OpenCliBridgeError) as ctx:
            transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        self.assertIn("RPI_OPENCLI_PROFILE", str(ctx.exception))

    def test_多个连接配置时报错提示环境变量(self):
        runner = FakeRunner(["  a — connected v1\n  b — connected v1\n"])
        transport = OpenCliTransport(runner=runner, cli_path="opencli")
        with self.assertRaises(OpenCliBridgeError):
            transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")


class RequestContractTests(unittest.TestCase):
    def _transport(self, responses):
        return OpenCliTransport(profile="p", runner=FakeRunner(responses))

    def test_get请求把params并入查询串(self):
        transport = self._transport([_cli_payload(200, b"{}")])
        transport.request(
            "GET",
            "https://zsk-cmis.cicpa.org.cn/open/api/v1/thing",
            params={"internal": "false"},
        )
        js = runner_eval_arg(transport, 0)
        self.assertIn("/open/api/v1/thing?internal=false", js)

    def test_post_json请求嵌入序列化体并设置内容类型(self):
        transport = self._transport([_cli_payload(200, b"{}")])
        transport.request(
            "POST",
            "https://zsk-cmis.cicpa.org.cn/open/api/v1/search",
            json={"query": "公司"},
        )
        js = runner_eval_arg(transport, 0)
        self.assertIn('"query"', js.replace("\\", ""))
        self.assertIn("application/json", js)
        self.assertIn("POST", js)

    def test_multipart文件上传编译为FormData(self):
        transport = self._transport([_cli_payload(200, b"{}")])
        transport.request(
            "POST",
            "https://zsk-cmis.cicpa.org.cn/open/api/v1/upload_local",
            files={"file": ("companies.xlsx", b"PK", "application/vnd.xlsx")},
        )
        js = runner_eval_arg(transport, 0)
        self.assertIn("FormData", js)
        self.assertIn("companies.xlsx", js)
        self.assertIn("application/vnd.xlsx", js)
        self.assertIn(base64.b64encode(b"PK").decode("ascii"), js)

    def test_自动注入csrf令牌头(self):
        transport = self._transport([_cli_payload(200, b"{}")])
        transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        js = runner_eval_arg(transport, 0)
        self.assertIn("XSRF-TOKEN", js)
        self.assertIn("X-XSRF-TOKEN", js)

    def test_非知识库域名被拒绝(self):
        transport = self._transport([])
        with self.assertRaises(OpenCliBridgeError):
            transport.request("GET", "https://evil.example.com/x")

    def test_响应对象完整暴露状态头和json(self):
        body = json.dumps({"status_code": 0}).encode("utf-8")
        transport = self._transport(
            [_cli_payload(200, body, headers={"Retry-After": "30"})]
        )
        response = transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers.get("Retry-After"), "30")
        self.assertEqual(response.json(), {"status_code": 0})
        self.assertEqual(response.content, body)

    def test_页面不在知识库域时导航回来重发(self):
        transport = self._transport(
            [
                '{"s":0,"nav":true,"h":{},"b":""}',
                "",
                _cli_payload(200, b"{}"),
            ]
        )
        response = transport.request("GET", "https://zsk-cmis.cicpa.org.cn/x")
        self.assertEqual(response.status_code, 200)
        self.assertTrue(any("open" in " ".join(call) for call in transport._runner.calls))


def runner_eval_arg(transport, call_index):
    args = transport._runner.calls[call_index]
    return args[args.index("eval") + 1]


class LoginFlowTests(unittest.TestCase):
    def test_open_page使用open子命令(self):
        runner = FakeRunner([""])
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        transport.open_page("https://zsk-cmis.cicpa.org.cn/")
        args = runner.calls[0]
        self.assertIn("open", args)
        self.assertIn(OPENCLI_SESSION, args)

    def test_get_url读取当前地址(self):
        runner = FakeRunner(["https://zsk-cmis.cicpa.org.cn/#/home_local\n"])
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        self.assertTrue(transport.get_url().startswith("https://zsk-cmis.cicpa.org.cn"))

    def test_会话验证以官方接口结果为准(self):
        ok = json.dumps({"status_code": 0}).encode("utf-8")
        runner = FakeRunner([ok.decode("utf-8")])
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        self.assertTrue(transport.verify_session())
        self.assertIn("get_user_index", runner_eval_arg(transport, 0))

    def test_等待登录成功返回且轮询验证(self):
        ok = json.dumps({"status_code": 0}).encode("utf-8")
        runner = FakeRunner(
            [
                "https://zsk-cmis.cicpa.org.cn/#/home_local\n",
                ok.decode("utf-8"),
            ]
        )
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        transport.wait_for_login(timeout=5.0, poll=0.0)
        self.assertEqual(len(runner.calls), 2)

    def test_等待登录超时抛出明确错误(self):
        runner = FakeRunner(
            [
                "https://zsk-cmis.cicpa.org.cn/#/home_local\n",
                '{"status_code":401}',
            ]
        )
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        clock = iter([0.0, 1.0, 2.0, 3.0])
        with self.assertRaises(OpenCliBridgeError) as ctx:
            transport.wait_for_login(timeout=2.0, poll=0.0, clock=lambda: next(clock))
        self.assertIn("登录", str(ctx.exception))

    def test_等待登录先回到知识库域(self):
        ok = json.dumps({"status_code": 0}).encode("utf-8")
        runner = FakeRunner(
            [
                "about:blank\n",
                "",
                '{"status_code": 0}',
            ]
        )
        transport = OpenCliTransport(profile="p", runner=runner, cli_path="opencli")
        transport.wait_for_login(timeout=5.0, poll=0.0)
        self.assertIn("open", " ".join(runner.calls[1]))


if __name__ == "__main__":
    unittest.main()
