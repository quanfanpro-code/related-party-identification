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
#import base64
import json
import os
from pathlib import Path
import sqlite3
import tempfile
from types import SimpleNamespace
import unittest

from scripts.cicpa.auth import (
    CredentialStore,
    DpapiProtector,
    collect_cookies_with_browser,
    default_auth_path,
    default_browser_profile_path,
    default_cookie_verifier,
    default_status_path,
    detect_browser,
    open_default_browser,
    public_login_status,
    read_firefox_cookies,
    read_login_status,
    start_guided_login,
    write_login_status,
)


class FakeProtector:
    def protect(self, value):
        return b"protected:" + base64.b64encode(value)

    def unprotect(self, value):
        return base64.b64decode(value.removeprefix(b"protected:"))


class FakeResponseClient:
    def __init__(self, payload):
        self.payload = payload
        self.calls = []

    def request_json(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        return self.payload


class AuthPathTests(unittest.TestCase):
    def test_认证文件只位于当前用户本地应用数据目录(self):
        local_app_data = r"C:\Users\测试\AppData\Local"

        self.assertEqual(
            default_auth_path(local_app_data),
            Path(local_app_data) / "related-party-identification" / "auth.bin",
        )
        self.assertEqual(
            default_status_path(local_app_data),
            Path(local_app_data) / "related-party-identification" / "login-status.json",
        )

    def test_加密文件不包含明文_cookie(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_auth_fake_"))
        auth_path = artifact_dir / "auth.bin"
        store = CredentialStore(path=auth_path, protector=FakeProtector())
        cookies = {"cicpa_token": "plain-secret-token", "XSRF-TOKEN": "plain-xsrf"}

        store.save_cookies(cookies)

        self.assertNotIn(b"plain-secret-token", auth_path.read_bytes())
        self.assertEqual(store.load_cookies(), cookies)

    def test_公开状态只能包含允许字段且_json_不带_bom(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_auth_status_"))
        status_path = artifact_dir / "login-status.json"

        written = write_login_status(status_path, "waiting_user", "请在官方浏览器中完成登录")
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


class BrowserLoginTests(unittest.TestCase):
    def test_优先识别_windows_默认_firefox(self):
        expected = Path(r"C:\Program Files\Mozilla Firefox\firefox.exe")

        browser = detect_browser(
            default_command_reader=lambda: (
                r'"C:\Program Files\Mozilla Firefox\firefox.exe" -osint -url "%1"'
            ),
            exists=lambda path: Path(path) == expected,
        )

        self.assertEqual(browser["name"], "Mozilla Firefox")
        self.assertEqual(browser["engine"], "firefox-profile")
        self.assertEqual(Path(browser["executable_path"]), expected)

    def test_优先识别六十四位_edge_安装位置(self):
        program_files = r"C:\Program Files"
        expected = Path(program_files) / "Microsoft" / "Edge" / "Application" / "msedge.exe"

        browser = detect_browser(
            env={"ProgramFiles": program_files, "ProgramFiles(x86)": r"C:\Program Files (x86)"},
            exists=lambda path: Path(path) == expected,
        )

        self.assertEqual(browser["name"], "Microsoft Edge")
        self.assertEqual(browser["engine"], "opencli")
        self.assertEqual(Path(browser["executable_path"]), expected)

    def test_默认_firefox_按正常方式打开并读取当前配置(self):
        profile_path = Path(tempfile.mkdtemp(prefix="rpi_firefox_profile_"))
        opened = []

        cookies = collect_cookies_with_browser(
            browser={
                "name": "Mozilla Firefox",
                "engine": "firefox-profile",
                "executable_path": r"C:\Firefox\firefox.exe",
            },
            profile_path=profile_path,
            browser_opener=lambda url: opened.append(url),
            cookie_reader=lambda profile: {
                "XSRF-TOKEN": "xsrf-value",
                "cicpa_token": "token-value",
            } if Path(profile) == profile_path else {},
            clock=lambda: 0.0,
            sleep=lambda _seconds: None,
            timeout=10.0,
        )

        self.assertEqual(opened, ["https://cmis.cicpa.org.cn/#/login"])
        self.assertEqual(cookies["cicpa_token"], "token-value")

    def test_edge_通过_opencli_读取平时浏览器的登录态(self):
        profile_root = Path(tempfile.mkdtemp(prefix="rpi_edge_profile_"))
        opened = []
        captured = {}

        def edge_cookie_reader(**kwargs):
            captured.update(kwargs)
            kwargs["on_ready"]()
            return {
                "XSRF-TOKEN": "xsrf-value",
                "cicpa_token": "token-value",
            }

        cookies = collect_cookies_with_browser(
            browser={
                "name": "Microsoft Edge",
                "engine": "opencli",
                "executable_path": r"C:\Edge\msedge.exe",
            },
            profile_path=profile_root,
            browser_opener=lambda url: opened.append(url),
            edge_cookie_reader=edge_cookie_reader,
            extension_checker=lambda path: Path(path) == profile_root,
            timeout=30.0,
        )

        self.assertEqual(opened, ["https://cmis.cicpa.org.cn/#/login"])
        self.assertEqual(captured["timeout"], 30.0)
        self.assertEqual(cookies["cicpa_token"], "token-value")

    def test_edge_缺少_opencli_时给出可执行的安装选择(self):
        profile_root = Path(tempfile.mkdtemp(prefix="rpi_edge_no_opencli_"))
        opened = []

        with self.assertRaisesRegex(Exception, "推荐安装 Firefox.*OpenCLI"):
            collect_cookies_with_browser(
                browser={
                    "name": "Microsoft Edge",
                    "engine": "opencli",
                    "executable_path": r"C:\Edge\msedge.exe",
                },
                profile_path=profile_root,
                browser_opener=lambda url: opened.append(url),
                extension_checker=lambda _path: False,
            )

        self.assertEqual(opened, [])

    def test_浏览器配置就是当前用户原有_edge_配置(self):
        local_app_data = r"C:\Users\测试\AppData\Local"

        self.assertEqual(
            default_browser_profile_path(local_app_data),
            Path(local_app_data)
            / "Microsoft"
            / "Edge"
            / "User Data",
        )

    def test_firefox_使用当前用户原有配置(self):
        roaming = Path(tempfile.mkdtemp(prefix="rpi_firefox_profile_"))
        firefox_dir = roaming / "Mozilla" / "Firefox"
        firefox_dir.mkdir(parents=True)
        (firefox_dir / "profiles.ini").write_text(
            "[InstallABC]\nDefault=Profiles/main.default\nLocked=1\n",
            encoding="utf-8",
        )

        profile = default_browser_profile_path(
            browser_name="Mozilla Firefox",
            roaming_app_data=str(roaming),
        )

        self.assertEqual(
            profile,
            firefox_dir / "Profiles" / "main.default",
        )

    def test_系统默认浏览器只按普通链接方式打开一次(self):
        opened = []

        open_default_browser(
            "https://cmis.cicpa.org.cn/#/login",
            opener=lambda url: opened.append(url),
        )

        self.assertEqual(opened, ["https://cmis.cicpa.org.cn/#/login"])

    def test_直接读取_firefox_当前注协登录状态(self):
        profile_path = Path(tempfile.mkdtemp(prefix="rpi_firefox_cookie_"))
        connection = sqlite3.connect(profile_path / "cookies.sqlite")
        try:
            connection.execute(
                "CREATE TABLE moz_cookies "
                "(name TEXT, value TEXT, host TEXT, expiry INTEGER)"
            )
            connection.executemany(
                "INSERT INTO moz_cookies VALUES (?, ?, ?, ?)",
                [
                    (
                        "cicpa_token",
                        "token-value",
                        "zsk-cmis.cicpa.org.cn",
                        4102444800,
                    ),
                    ("other", "ignore", "example.com", 4102444800),
                ],
            )
            connection.commit()
        finally:
            connection.close()
        session_dir = profile_path / "sessionstore-backups"
        session_dir.mkdir()
        session_payload = json.dumps(
            {
                "cookies": [
                    {
                        "name": "XSRF-TOKEN",
                        "value": "xsrf-value",
                        "host": "zsk-cmis.cicpa.org.cn",
                    }
                ]
            },
            separators=(",", ":"),
        ).encode("utf-8")
        literal_length = len(session_payload)
        compressed = bytearray([min(literal_length, 15) << 4])
        remaining = literal_length - 15
        if remaining >= 0:
            while remaining >= 255:
                compressed.append(255)
                remaining -= 255
            compressed.append(remaining)
        compressed.extend(session_payload)
        (session_dir / "recovery.jsonlz4").write_bytes(
            b"mozLz40\0"
            + literal_length.to_bytes(4, "little")
            + bytes(compressed)
        )

        cookies = read_firefox_cookies(profile_path, now=lambda: 0)

        self.assertEqual(
            cookies,
            {
                "cicpa_token": "token-value",
                "XSRF-TOKEN": "xsrf-value",
            },
        )

    def test_认证实现不再依赖_playwright(self):
        cicpa_dir = Path(__file__).resolve().parents[1] / "cicpa"
        source = "\n".join(
            path.read_text(encoding="utf-8-sig")
            for path in (
                cicpa_dir / "auth.py",
                cicpa_dir / "edge_bridge.py",
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
            auth_path=artifact_dir / "auth.bin",
            status_path=artifact_dir / "login-status.json",
            popen=fake_popen,
        )

        command_text = " ".join(str(part) for part in captured["command"]).lower()
        self.assertEqual(result["pid"], 4321)
        self.assertNotIn("cookie=", command_text)
        self.assertNotIn("password=", command_text)
        self.assertNotIn("token-value", command_text)


class CookieVerificationTests(unittest.TestCase):
    def test_登录成功以轻量官方接口结果为准(self):
        fake_client = FakeResponseClient({"status_code": 0, "data": {"user": "ok"}})
        captured = {}

        def client_factory(cookies):
            captured["cookies"] = cookies
            return fake_client

        verified = default_cookie_verifier(
            {"cicpa_token": "secret"},
            client_factory=client_factory,
        )

        self.assertTrue(verified)
        self.assertEqual(captured["cookies"], {"cicpa_token": "secret"})
        self.assertEqual(len(fake_client.calls), 1)


@unittest.skipUnless(os.name == "nt", "Windows DPAPI 仅在 Windows 运行")
class DpapiRoundTripTests(unittest.TestCase):
    def test_当前_windows_用户可往返解密且磁盘无明文(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_auth_dpapi_"))
        auth_path = artifact_dir / "auth.bin"
        store = CredentialStore(path=auth_path, protector=DpapiProtector())
        cookies = {"cicpa_token": "dpapi-plain-secret"}

        store.save_cookies(cookies)

        self.assertEqual(store.load_cookies(), cookies)
        self.assertNotIn(b"dpapi-plain-secret", auth_path.read_bytes())


if __name__ == "__main__":
    unittest.main()
