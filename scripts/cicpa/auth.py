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
#"""Windows 当前用户认证状态与 OpenCLI + Edge 登录编排。"""

import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Callable, Optional

from .browser_transport import OpenCliBridgeError, OpenCliTransport
from .opencli_setup import opencli_extension_installed  # noqa: F401 供工作流复用


VALID_STATUSES = {"idle", "waiting_user", "authenticated", "expired", "failed"}


class AuthError(RuntimeError):
    """认证流程基础错误。"""


def _local_app_data(local_app_data: Optional[str] = None) -> Path:
    value = local_app_data or os.environ.get("LOCALAPPDATA")
    if not value:
        raise AuthError("无法确定当前用户的 LOCALAPPDATA 目录")
    return Path(value) / "related-party-identification"


def default_status_path(local_app_data: Optional[str] = None) -> Path:
    return _local_app_data(local_app_data) / "login-status.json"


def default_browser_profile_path(
    local_app_data: Optional[str] = None,
    browser_name: str = "Microsoft Edge",
    roaming_app_data: Optional[str] = None,
) -> Path:
    """返回当前 Windows 用户原有浏览器配置目录。"""
    _ = roaming_app_data
    value = local_app_data or os.environ.get("LOCALAPPDATA")
    if not value:
        raise AuthError("无法确定当前用户的 LOCALAPPDATA 目录")
    if browser_name == "Google Chrome":
        return Path(value) / "Google" / "Chrome" / "User Data"
    return Path(value) / "Microsoft" / "Edge" / "User Data"


def detect_browser(
    env=None,
    exists: Callable[[Path], bool] = os.path.exists,
    preferred_name: Optional[str] = None,
):
    """识别已安装的 Edge 或 Chrome;Firefox 不在本技能路线内。"""
    environment = os.environ if env is None else env
    candidates = []
    for variable, relative, name in (
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
                "engine": "opencli",
                "executable_path": str(executable),
            }
    return None


def public_login_status(status: str, message_zh: str) -> dict:
    if status not in VALID_STATUSES:
        raise ValueError("不允许的登录状态")
    return {
        "status": status,
        "message_zh": str(message_zh),
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }


def write_login_status(path: Path, status: str, message_zh: str) -> dict:
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


def read_login_status(path: Optional[Path] = None) -> dict:
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


def collect_browser_session(
    *,
    bridge_factory: Callable[..., OpenCliTransport] = OpenCliTransport,
    timeout: float = 900.0,
    sleep: Callable[[float], None] = time.sleep,
) -> None:
    """等待 Edge 中的注协会话就绪;会话即登录态,不复制任何 Cookie。"""
    bridge = bridge_factory()
    try:
        bridge.wait_for_login(timeout=timeout, sleep=sleep)
    except OpenCliBridgeError as exc:
        raise AuthError(str(exc)) from exc


def auth_status(
    *,
    bridge_factory: Callable[..., OpenCliTransport] = OpenCliTransport,
) -> dict:
    """实时验证浏览器内注协会话,不读取任何保存的凭据。"""
    try:
        bridge = bridge_factory()
        verified = bridge.verify_session()
    except OpenCliBridgeError as exc:
        return {"status": "failed", "message_zh": str(exc)}
    except OSError as exc:
        return {"status": "failed", "message_zh": "opencli 调用失败:{}".format(exc)}
    if verified:
        return {"status": "authenticated", "message_zh": "Edge 内注协会话有效"}
    return {
        "status": "expired",
        "message_zh": "Edge 内尚未进入注协知识库或会话已失效,需要重新执行登录引导",
    }


def run_guided_login(
    status_path: Path,
    *,
    bridge_factory: Callable[..., OpenCliTransport] = OpenCliTransport,
) -> bool:
    """后台帮助进程执行登录编排、验证并写入公开状态。"""
    write_login_status(status_path, "waiting_user", "正在打开 Edge 中的注协官方登录页面")
    try:
        collect_browser_session(bridge_factory=bridge_factory)
        write_login_status(
            status_path,
            "authenticated",
            "登录成功,Edge 内注协会话有效;本技能不保存任何凭据",
        )
        return True
    except AuthError as exc:
        write_login_status(status_path, "failed", str(exc))
    except Exception:
        write_login_status(status_path, "failed", "登录流程出现错误,请重新启动登录")
    return False


def start_guided_login(
    status_path: Optional[Path] = None,
    *,
    popen: Callable = subprocess.Popen,
):
    """启动无终端输入的后台帮助进程;登录页固定由 opencli 驱动的 Edge 打开。"""
    status_destination = (
        Path(status_path) if status_path is not None else default_status_path()
    )
    write_login_status(status_destination, "waiting_user", "正在打开 Edge 中的注协官方登录页面")
    skill_root = Path(__file__).resolve().parents[2]
    command = [
        sys.executable,
        "-m",
        "scripts.cicpa.auth",
        "--worker",
        "--status-path",
        str(status_destination),
    ]
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
        "message_zh": "正在通过 OpenCLI 在 Edge 中打开注协登录页面,请在该页面完成登录",
        "pid": process.pid,
    }


def _main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="关联方识别登录帮助进程")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--status-path", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if not args.worker or not args.status_path:
        parser.error("此模块由关联方识别工作流调用")
    return 0 if run_guided_login(args.status_path) else 1


if __name__ == "__main__":
    raise SystemExit(_main())
