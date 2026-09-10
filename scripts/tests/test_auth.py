# 关联方识别与核查 — 原创编排层
# Copyright (C) 2026 CPA-Q (quanfanpro-code)
#
# 本文件是 related-party-identification 的原创编排层,采用 GNU Affero General
# Public License v3.0 (AGPL-3.0) 发布。
# 完整协议见项目根目录 LICENSE 文件(AGPL-3.0)。
#
# 本项目的 scripts/cicpa/ 目录包含改编自 nigo/nigo-skills(MIT) 和
# jackwener/OpenCLI(Apache-2.0) 的代码,分别保留原始许可证。
# 详见 NOTICE 和 references/SOURCES.json。
#"""OpenCLI + Edge 唯一登录路线的认证编排测试,全部注入假桥,不联网。"""

import json
from pathlib import Path
import tempfile
import unittest

from scripts.cicpa.auth import (
    AuthError,
    auth_status,
    collect_browser_session,
    default_browser_profile_path,
    default_status_path,
    detect_browser,
    public_login_status,
    read_login_status,
    run_guided_login,
    start_guided_login,
    write_login_status,
)
from scripts.cicpa.browser_transport import OpenCliBridgeError


class FakeBridge:
    """按预设行为模拟 opencli 页面桥。"""

    def __init__(self, *, verify=None, wait_error=None):
        self.verify = verify
        self.wait_error = wait_error
        self.wait_calls = []

    def wait_for_login(self, timeout=900.0, poll=2.5, clock=None, sleep=None):
        self.wait_calls.append(timeout)
        if self.wait_error is not None:
            raise self.wait_error

    def verify_session(self):
        if isinstance(self.verify, Exception):
            raise self.verify
        return bool(self.verify)


class AuthPathTests(unittest.TestCase):
    def test_状态文件只位于当前用户本地应用数据目录(self):
        local_app_data = r"C:\Users\测试\AppData\Local"

        self.assertEqual(
            default_status_path(local_app_data),
            Path(local_app_data) / "related-party-identification" / "login-status.json",
        )

    def test_公开状态只能包含允许字段且_json_不带_bom(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_auth_status_"))
        status_path = artifact_dir / "login-status.json"

        written = write_login_status(status_path, "waiting_user", "请在 Edge 中完成登录")
        payload = read_login_status(status_path)
        raw = status_path.read_bytes()

        self.assertEqual(set(payload), {"status", "message_zh", "updated_at"})
        self.assertEqual(payload["status"], "waiting_user")
        self.assertEqual(payload["updated_at"], written["updated_at"])
        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        serialized = json.dumps(payload, ensure_ascii=False).lower()
        self.assertNotIn("cookie", serialized)
        self.assertNotIn("password", serialized)
        self.assertNotIn("header", serialized)

    def test_非法公开状态会被拒绝(self):
        with self.assertRaises(ValueError):
            public_login_status("has_cookie", "不应接受")


class BrowserDetectTests(unittest.TestCase):
    def test_识别六十四位_edge_安装位置(self):
        program_files = r"C:\Program Files"
        expected = Path(program_files) / "Microsoft" / "Edge" / "Application" / "msedge.exe"

        browser = detect_browser(
            env={"ProgramFiles": program_files, "ProgramFiles(x86)": r"C:\Program Files (x86)"},
            exists=lambda path: Path(path) == expected,
        )

        self.assertEqual(browser["name"], "Microsoft Edge")
        self.assertEqual(Path(browser["executable_path"]), expected)

    def test_firefox不再作为候选返回(self):
        browser = detect_browser(
            env={"ProgramFiles": r"C:\Program Files"},
            exists=lambda path: "firefox" in str(path).lower(),
        )

        self.assertIsNone(browser)

    def test_浏览器配置就是当前用户原有_edge_配置(self):
        local_app_data = r"C:\Users\测试\AppData\Local"

        self.assertEqual(
            default_browser_profile_path(local_app_data),
            Path(local_app_data)
            / "Microsoft"
            / "Edge"
            / "User Data",
        )


class SessionLoginTests(unittest.TestCase):
    def test_登录编排等待浏览器会话就绪(self):
        waits = []

        def factory(**kwargs):
            bridge = FakeBridge()
            bridge.wait_for_login = lambda **kw: waits.append(kw.get("timeout"))
            return bridge

        collect_browser_session(bridge_factory=factory, timeout=60.0)

        self.assertEqual(waits, [60.0])

    def test_登录编排把桥接错误翻译成认证错误(self):
        def factory(**kwargs):
            return FakeBridge(wait_error=OpenCliBridgeError("登录等待超时:请在 Edge 中完成登录"))

        with self.assertRaisesRegex(AuthError, "登录等待超时"):
            collect_browser_session(bridge_factory=factory)

    def test_引导登录成功写入认证状态(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_login_ok_"))
        status_path = artifact_dir / "login-status.json"

        ok = run_guided_login(status_path, bridge_factory=lambda **kw: FakeBridge(verify=True))

        self.assertTrue(ok)
        self.assertEqual(read_login_status(status_path)["status"], "authenticated")

    def test_引导登录失败写入具体原因(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_login_fail_"))
        status_path = artifact_dir / "login-status.json"

        ok = run_guided_login(
            status_path,
            bridge_factory=lambda **kw: FakeBridge(
                wait_error=OpenCliBridgeError("未找到 opencli 命令")
            ),
        )

        payload = read_login_status(status_path)
        self.assertFalse(ok)
        self.assertEqual(payload["status"], "failed")
        self.assertIn("opencli", payload["message_zh"])

    def test_状态验证三分支(self):
        self.assertEqual(
            auth_status(bridge_factory=lambda **kw: FakeBridge(verify=True))["status"],
            "authenticated",
        )
        self.assertEqual(
            auth_status(bridge_factory=lambda **kw: FakeBridge(verify=False))["status"],
            "expired",
        )
        failed = auth_status(
            bridge_factory=lambda **kw: FakeBridge(
                verify=OpenCliBridgeError("OpenCLI 扩展未连接")
            )
        )
        self.assertEqual(failed["status"], "failed")
        self.assertIn("OpenCLI", failed["message_zh"])

    def test_认证实现不再依赖_playwright(self):
        cicpa_dir = Path(__file__).resolve().parents[1] / "cicpa"
        source = "\n".join(
            path.read_text(encoding="utf-8-sig")
            for path in (
                cicpa_dir / "auth.py",
                cicpa_dir / "browser_transport.py",
                cicpa_dir / "opencli_setup.py",
            )
        )

        self.assertNotIn("playwright", source.lower())
        self.assertNotIn("remote-debugging", source)
        self.assertNotIn("extension://", source)

    def test_后台帮助进程的命令行不包含认证秘密(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_auth_process_"))
        captured = {}

        class Process:
            pid = 4321

        def fake_popen(command, **kwargs):
            captured["command"] = command
            captured["kwargs"] = kwargs
            return Process()

        result = start_guided_login(
            status_path=artifact_dir / "login-status.json",
            popen=fake_popen,
        )

        command_text = " ".join(str(part) for part in captured["command"]).lower()
        self.assertEqual(result["pid"], 4321)
        self.assertNotIn("cookie=", command_text)
        self.assertNotIn("password=", command_text)
        self.assertNotIn("token-value", command_text)


if __name__ == "__main__":
    unittest.main()
