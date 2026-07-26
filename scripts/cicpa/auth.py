# 关联方识别与核查 — 注协登录凭据管理
# 本文件改编自 nigo/nigo-skills/cicpa-company-query(MIT)
# 上游作者: nigo(涂佳兵) | 原始仓库: https://github.com/nigo81/nigo-skills
# 上游协议: MIT
#
# 本地修改: CPA-Q(quanfanpro-code)
# 本文件的修改部分同样以 MIT 协议发布,与上游保持一致。
#"""Windows 当前用户认证状态与可见浏览器登录。"""

import argparse
import configparser
import ctypes
from ctypes import wintypes
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import time
from typing import Callable, Dict, Optional

from .client import CicpaClient, CicpaError, RateLimiter, RatePolicy, requests
from .edge_bridge import OpenCliBridgeError, collect_edge_cookies
from .opencli_setup import opencli_extension_installed


LOGIN_URL = "https://cmis.cicpa.org.cn/#/login"
ZSK_HOST = "zsk-cmis.cicpa.org.cn"
ZSK_BASE = "https://zsk-cmis.cicpa.org.cn"
VALID_STATUSES = {"idle", "waiting_user", "authenticated", "expired", "failed"}


class AuthError(RuntimeError):
    """认证流程基础错误。"""


class DependencyMissing(AuthError):
    """登录所需依赖不存在。"""


class BrowserUnavailable(AuthError):
    """没有找到可使用的系统浏览器。"""


class LoginTimeout(AuthError):
    """用户未在限定时间内完成登录。"""


class _DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


class DpapiProtector:
    """使用 Windows DPAPI 的当前用户密钥保护字节。"""

    _UI_FORBIDDEN = 0x01

    @staticmethod
    def _blob(value: bytes):
        buffer = ctypes.create_string_buffer(value, len(value))
        blob = _DataBlob(
            len(value),
            ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
        )
        return blob, buffer

    def protect(self, value: bytes) -> bytes:
        if os.name != "nt":
            raise AuthError("DPAPI 仅支持 Windows")
        input_blob, input_buffer = self._blob(value)
        output_blob = _DataBlob()
        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        success = crypt32.CryptProtectData(
            ctypes.byref(input_blob),
            "related-party-identification",
            None,
            None,
            None,
            self._UI_FORBIDDEN,
            ctypes.byref(output_blob),
        )
        _ = input_buffer
        if not success:
            raise ctypes.WinError()
        try:
            return ctypes.string_at(output_blob.pbData, output_blob.cbData)
        finally:
            kernel32.LocalFree(output_blob.pbData)

    def unprotect(self, value: bytes) -> bytes:
        if os.name != "nt":
            raise AuthError("DPAPI 仅支持 Windows")
        input_blob, input_buffer = self._blob(value)
        output_blob = _DataBlob()
        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        success = crypt32.CryptUnprotectData(
            ctypes.byref(input_blob),
            None,
            None,
            None,
            None,
            self._UI_FORBIDDEN,
            ctypes.byref(output_blob),
        )
        _ = input_buffer
        if not success:
            raise ctypes.WinError()
        try:
            return ctypes.string_at(output_blob.pbData, output_blob.cbData)
        finally:
            kernel32.LocalFree(output_blob.pbData)


def _local_app_data(local_app_data: Optional[str] = None) -> Path:
    value = local_app_data or os.environ.get("LOCALAPPDATA")
    if not value:
        raise AuthError("无法确定当前用户的 LOCALAPPDATA 目录")
    return Path(value) / "related-party-identification"


def default_auth_path(local_app_data: Optional[str] = None) -> Path:
    return _local_app_data(local_app_data) / "auth.bin"


def default_status_path(local_app_data: Optional[str] = None) -> Path:
    return _local_app_data(local_app_data) / "login-status.json"


def default_browser_profile_path(
    local_app_data: Optional[str] = None,
    browser_name: str = "Microsoft Edge",
    roaming_app_data: Optional[str] = None,
) -> Path:
    """返回当前 Windows 用户原有浏览器配置目录。"""
    if browser_name == "Mozilla Firefox":
        roaming = roaming_app_data or os.environ.get("APPDATA")
        if not roaming:
            raise AuthError("无法确定当前用户的 APPDATA 目录")
        firefox_root = Path(roaming) / "Mozilla" / "Firefox"
        parser = configparser.RawConfigParser()
        if not parser.read(firefox_root / "profiles.ini", encoding="utf-8"):
            raise BrowserUnavailable("未找到当前用户的 Firefox 配置")
        profile_value = next(
            (
                parser.get(section, "Default")
                for section in parser.sections()
                if section.startswith("Install") and parser.has_option(section, "Default")
            ),
            None,
        )
        if profile_value is None:
            profile_value = next(
                (
                    parser.get(section, "Path")
                    for section in parser.sections()
                    if section.startswith("Profile")
                    and parser.getboolean(section, "Default", fallback=False)
                ),
                None,
            )
        if profile_value is None:
            raise BrowserUnavailable("未找到当前用户正在使用的 Firefox 配置")
        profile = Path(profile_value)
        return profile if profile.is_absolute() else firefox_root / profile
    value = local_app_data or os.environ.get("LOCALAPPDATA")
    if not value:
        raise AuthError("无法确定当前用户的 LOCALAPPDATA 目录")
    if browser_name == "Google Chrome":
        return Path(value) / "Google" / "Chrome" / "User Data"
    return Path(value) / "Microsoft" / "Edge" / "User Data"


class CredentialStore:
    """只保存 DPAPI 加密后的 Cookie 字典。"""

    def __init__(self, path: Optional[Path] = None, protector=None):
        self.path = Path(path) if path is not None else default_auth_path()
        self.protector = protector or DpapiProtector()

    def save_cookies(self, cookies: Dict[str, str]) -> None:
        if not isinstance(cookies, dict) or not cookies:
            raise ValueError("Cookie 字典不能为空")
        normalized = {}
        for name, value in cookies.items():
            if not isinstance(name, str) or not isinstance(value, str):
                raise ValueError("Cookie 名称和值必须是字符串")
            normalized[name] = value
        plain = json.dumps(
            normalized,
            ensure_ascii=False,
            separators=(",", ":"),
            sort_keys=True,
        ).encode("utf-8")
        encrypted = self.protector.protect(plain)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        with open(temporary, "wb") as stream:
            stream.write(encrypted)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, self.path)

    def load_cookies(self) -> Dict[str, str]:
        if not self.path.exists():
            return {}
        encrypted = self.path.read_bytes()
        plain = self.protector.unprotect(encrypted)
        payload = json.loads(plain.decode("utf-8"))
        if not isinstance(payload, dict):
            raise AuthError("认证文件内容无效")
        return {str(name): str(value) for name, value in payload.items()}


def public_login_status(status: str, message_zh: str) -> Dict[str, str]:
    if status not in VALID_STATUSES:
        raise ValueError("不允许的登录状态")
    return {
        "status": status,
        "message_zh": str(message_zh),
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }


def write_login_status(path: Path, status: str, message_zh: str) -> Dict[str, str]:
    destination = Path(path)
    payload = public_login_status(status, message_zh)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    with open(temporary, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(text)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, destination)
    return payload


def read_login_status(path: Optional[Path] = None) -> Dict[str, str]:
    source = Path(path) if path is not None else default_status_path()
    if not source.exists():
        return public_login_status("idle", "尚未启动登录")
    with open(source, "r", encoding="utf-8") as stream:
        payload = json.load(stream)
    status = payload.get("status")
    if status not in VALID_STATUSES:
        raise AuthError("登录状态文件内容无效")
    return {
        "status": status,
        "message_zh": str(payload.get("message_zh", "")),
        "updated_at": str(payload.get("updated_at", "")),
    }


def _read_default_browser_command() -> str:
    if os.name != "nt":
        return ""
    import winreg

    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\Shell\Associations"
            r"\UrlAssociations\https\UserChoice",
        ) as key:
            prog_id = winreg.QueryValueEx(key, "ProgId")[0]
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            str(prog_id) + r"\shell\open\command",
        ) as key:
            return str(winreg.QueryValueEx(key, "")[0])
    except OSError:
        return ""


def _command_executable(command: str) -> Optional[Path]:
    value = str(command).strip()
    if not value:
        return None
    if value.startswith('"'):
        end = value.find('"', 1)
        return Path(value[1:end]) if end > 1 else None
    return Path(value.split(maxsplit=1)[0])


def detect_browser(
    env=None,
    exists: Callable[[Path], bool] = os.path.exists,
    default_command_reader: Callable[[], str] = _read_default_browser_command,
    preferred_name: Optional[str] = None,
):
    """优先识别 Windows 默认浏览器，再回退到已安装的支持浏览器。"""
    if env is None and preferred_name is None:
        default_executable = _command_executable(default_command_reader())
        if default_executable is not None and exists(default_executable):
            executable_name = default_executable.name.lower()
            if executable_name == "firefox.exe":
                return {
                    "name": "Mozilla Firefox",
                    "engine": "firefox-profile",
                    "executable_path": str(default_executable),
                }
            if executable_name == "msedge.exe":
                return {
                    "name": "Microsoft Edge",
                    "engine": "opencli",
                    "executable_path": str(default_executable),
                }
            if executable_name == "chrome.exe":
                return {
                    "name": "Google Chrome",
                    "engine": "opencli",
                    "executable_path": str(default_executable),
                }
    environment = os.environ if env is None else env
    candidates = []
    for variable, relative, name in (
        ("ProgramFiles", r"Mozilla Firefox\firefox.exe", "Mozilla Firefox"),
        ("ProgramFiles(x86)", r"Mozilla Firefox\firefox.exe", "Mozilla Firefox"),
        ("ProgramFiles", r"Microsoft\Edge\Application\msedge.exe", "Microsoft Edge"),
        ("ProgramFiles(x86)", r"Microsoft\Edge\Application\msedge.exe", "Microsoft Edge"),
        ("LOCALAPPDATA", r"Microsoft\Edge\Application\msedge.exe", "Microsoft Edge"),
        ("ProgramFiles", r"Google\Chrome\Application\chrome.exe", "Google Chrome"),
        ("ProgramFiles(x86)", r"Google\Chrome\Application\chrome.exe", "Google Chrome"),
        ("LOCALAPPDATA", r"Google\Chrome\Application\chrome.exe", "Google Chrome"),
    ):
        root = environment.get(variable)
        if root:
            candidates.append((Path(root) / Path(relative), name))
    for executable, name in candidates:
        if (preferred_name is None or name == preferred_name) and exists(executable):
            return {
                "name": name,
                "engine": (
                    "firefox-profile"
                    if name == "Mozilla Firefox"
                    else "opencli"
                ),
                "executable_path": str(executable),
            }
    return None


def open_default_browser(url: str, *, opener=None) -> None:
    """按 Windows 日常方式把链接交给当前默认浏览器。"""
    if opener is None:
        if os.name != "nt":
            raise BrowserUnavailable("此登录流程仅支持 Windows")
        opener = os.startfile
    opener(url)


def _read_mozlz4(path: Path) -> bytes:
    raw = Path(path).read_bytes()
    if not raw.startswith(b"mozLz40\0") or len(raw) < 13:
        raise ValueError("Firefox 当前会话文件格式无效")
    expected_length = int.from_bytes(raw[8:12], "little")
    source = memoryview(raw)[12:]
    output = bytearray()
    index = 0
    while index < len(source):
        token = source[index]
        index += 1
        literal_length = token >> 4
        if literal_length == 15:
            while True:
                extra = source[index]
                index += 1
                literal_length += extra
                if extra != 255:
                    break
        output.extend(source[index:index + literal_length])
        index += literal_length
        if index >= len(source):
            break
        offset = source[index] | (source[index + 1] << 8)
        index += 2
        if offset <= 0 or offset > len(output):
            raise ValueError("Firefox 当前会话文件内容无效")
        match_length = token & 15
        if match_length == 15:
            while True:
                extra = source[index]
                index += 1
                match_length += extra
                if extra != 255:
                    break
        match_length += 4
        start = len(output) - offset
        for position in range(match_length):
            output.append(output[start + position])
    if len(output) != expected_length:
        raise ValueError("Firefox 当前会话文件尚未写完")
    return bytes(output)


def read_firefox_cookies(
    profile_path: Path,
    *,
    now: Callable[[], float] = time.time,
) -> Dict[str, str]:
    """从当前 Firefox 配置的只读数据库读取仍有效的注协 Cookie。"""
    cookies: Dict[str, str] = {}
    database = Path(profile_path) / "cookies.sqlite"
    if database.exists():
        connection = sqlite3.connect(
            database.as_uri() + "?mode=ro",
            uri=True,
            timeout=1,
        )
        try:
            rows = connection.execute(
                "SELECT name, value FROM moz_cookies "
                "WHERE host LIKE ? AND (expiry = 0 OR expiry > ?)",
                ("%cicpa.org.cn", int(now())),
            ).fetchall()
        finally:
            connection.close()
        cookies.update({str(name): str(value) for name, value in rows if name})

    session_path = (
        Path(profile_path)
        / "sessionstore-backups"
        / "recovery.jsonlz4"
    )
    if session_path.exists():
        session = json.loads(_read_mozlz4(session_path).decode("utf-8"))
        for item in session.get("cookies", []):
            if (
                isinstance(item, dict)
                and item.get("name")
                and str(item.get("host", "")).lstrip(".").endswith("cicpa.org.cn")
            ):
                cookies[str(item["name"])] = str(item.get("value", ""))
    return cookies


def collect_cookies_with_browser(
    *,
    browser=None,
    profile_path: Optional[Path] = None,
    browser_opener: Callable[[str], None] = open_default_browser,
    cookie_reader: Callable[[Path], Dict[str, str]] = read_firefox_cookies,
    edge_cookie_reader: Callable[..., Dict[str, str]] = collect_edge_cookies,
    extension_checker: Callable[[Path], bool] = opencli_extension_installed,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
    timeout: float = 900.0,
) -> Dict[str, str]:
    """用当前用户的 Firefox 或带 OpenCLI 的 Edge/Chrome 完成登录。"""
    selected = browser or detect_browser()
    if not selected:
        raise BrowserUnavailable(
            "未检测到 Firefox、Microsoft Edge 或 Google Chrome"
        )
    browser_name = str(selected.get("name", ""))
    if browser_name in {"Microsoft Edge", "Google Chrome"}:
        user_data_root = (
            Path(profile_path)
            if profile_path is not None
            else default_browser_profile_path(browser_name=browser_name)
        )
        if not extension_checker(user_data_root):
            raise BrowserUnavailable(
                "推荐安装 Firefox；若继续使用 Edge 或 Chrome，"
                "需要先安装 OpenCLI 浏览器扩展"
            )
        try:
            return edge_cookie_reader(
                timeout=timeout,
                on_ready=lambda: browser_opener(LOGIN_URL),
            )
        except OpenCliBridgeError as exc:
            raise AuthError(str(exc)) from exc
    if browser_name != "Mozilla Firefox":
        raise BrowserUnavailable(
            "当前浏览器不支持；推荐 Firefox，其次使用带 OpenCLI 扩展的 Edge 或 Chrome"
        )
    current_profile = (
        Path(profile_path)
        if profile_path is not None
        else default_browser_profile_path(browser_name=browser_name)
    )
    started = clock()
    browser_opener(LOGIN_URL)
    consecutive_read_errors = 0
    while clock() - started < timeout:
        try:
            cookies = cookie_reader(current_profile)
        except (OSError, sqlite3.Error, ValueError):
            consecutive_read_errors += 1
            if consecutive_read_errors >= 30:
                raise AuthError("无法读取当前默认浏览器登录状态")
        else:
            consecutive_read_errors = 0
            if cookies.get("cicpa_token") and cookies.get("XSRF-TOKEN"):
                return cookies
        sleep(2.0)
    raise LoginTimeout("登录等待超时")


def _client_from_cookies(cookies: Dict[str, str]) -> CicpaClient:
    if requests is None:
        raise DependencyMissing("缺少 requests，需先完成依赖预检")
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": ZSK_BASE + "/companylibrarynew/",
            "Origin": ZSK_BASE,
        }
    )
    zsk_names = {
        "XSRF-TOKEN",
        "yuqing_whole_jsessionid",
        "cicpa_token",
        "cicpa_ticket",
        "companyVerifyCode",
        "userid",
        "u_name",
    }
    for name, value in cookies.items():
        domain = ZSK_HOST if name in zsk_names else ".cicpa.org.cn"
        session.cookies.set(name, value, domain=domain)
    xsrf = cookies.get("XSRF-TOKEN")
    if xsrf:
        session.headers["X-Xsrf-Token"] = xsrf
    return CicpaClient(session=session, limiter=RateLimiter(RatePolicy()))


def default_cookie_verifier(
    cookies: Dict[str, str],
    client_factory: Callable[[Dict[str, str]], CicpaClient] = _client_from_cookies,
) -> bool:
    """以轻量官方接口验证 Cookie，不按固定保存时长猜测。"""
    try:
        client = client_factory(cookies)
        payload = client.request_json(
            "GET",
            ZSK_BASE + "/open/industry_chain_api/v1/search/get_user_index",
            cache_key=("auth", "get_user_index"),
        )
    except (AuthError, CicpaError, OSError):
        return False
    return payload.get("status_code") == 0


def run_guided_login(
    auth_path: Path,
    status_path: Path,
    *,
    browser_flow: Callable[..., Dict[str, str]] = collect_cookies_with_browser,
    verifier: Callable[[Dict[str, str]], bool] = default_cookie_verifier,
    browser_name: Optional[str] = None,
) -> bool:
    """后台帮助进程执行登录、验证和加密保存。"""
    write_login_status(status_path, "waiting_user", "请在已经打开的系统浏览器中登录注协并进入行业执业知识库")
    try:
        selected = (
            detect_browser(preferred_name=browser_name)
            if browser_name is not None
            else None
        )
        if browser_name is not None and selected is None:
            raise BrowserUnavailable("没有找到所选浏览器")
        cookies = (
            browser_flow(browser=selected)
            if selected is not None
            else browser_flow()
        )
        write_login_status(status_path, "waiting_user", "已检测到知识库页面，正在验证登录状态")
        if not verifier(cookies):
            write_login_status(status_path, "failed", "登录状态未通过官方接口验证，请重新登录或确认账号权限")
            return False
        CredentialStore(path=auth_path).save_cookies(cookies)
        write_login_status(status_path, "authenticated", "登录成功，认证状态已由当前 Windows 用户加密保存")
        return True
    except DependencyMissing:
        write_login_status(status_path, "failed", "缺少登录依赖，需要先取得用户同意后安装")
    except BrowserUnavailable as exc:
        write_login_status(status_path, "failed", str(exc))
    except LoginTimeout:
        write_login_status(status_path, "failed", "登录等待超时，可以重新启动登录流程")
    except AuthError as exc:
        write_login_status(status_path, "failed", str(exc))
    except Exception:
        write_login_status(status_path, "failed", "登录流程出现错误，未保存认证状态")
    return False


def start_guided_login(
    auth_path: Optional[Path] = None,
    status_path: Optional[Path] = None,
    *,
    popen: Callable = subprocess.Popen,
    browser_name: Optional[str] = None,
):
    """启动无终端输入的后台帮助进程，浏览器窗口仍保持可见。"""
    auth_destination = Path(auth_path) if auth_path is not None else default_auth_path()
    status_destination = Path(status_path) if status_path is not None else default_status_path()
    write_login_status(status_destination, "waiting_user", "正在打开注协官方登录页面")
    skill_root = Path(__file__).resolve().parents[2]
    command = [
        sys.executable,
        "-m",
        "scripts.cicpa.auth",
        "--worker",
        "--auth-path",
        str(auth_destination),
        "--status-path",
        str(status_destination),
    ]
    if browser_name is not None:
        command.extend(["--browser-name", browser_name])
    process = popen(
        command,
        cwd=str(skill_root),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return {
        "status": "waiting_user",
        "message_zh": "已打开注协官方登录流程，请在平时使用的浏览器中完成登录",
        "pid": process.pid,
    }


def _main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="关联方识别登录帮助进程")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--auth-path", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--status-path", type=Path, help=argparse.SUPPRESS)
    parser.add_argument(
        "--browser-name",
        choices=["Mozilla Firefox", "Microsoft Edge", "Google Chrome"],
        help=argparse.SUPPRESS,
    )
    args = parser.parse_args(argv)
    if not args.worker or not args.auth_path or not args.status_path:
        parser.error("此模块由关联方识别工作流调用")
    return (
        0
        if run_guided_login(
            args.auth_path,
            args.status_path,
            browser_name=args.browser_name,
        )
        else 1
    )


if __name__ == "__main__":
    raise SystemExit(_main())
