# 关联方识别与核查 — OpenCLI 页面取数通道
#"""经 opencli CLI 在 Edge 页面上下文调用注协接口,浏览器自动携带全量登录态。"""

from __future__ import annotations

import base64
import json as jsonlib
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
from urllib.parse import urlencode


OPENCLI_SESSION = "related-party"
ZSK_HOST = "zsk-cmis.cicpa.org.cn"
ZSK_BASE = "https://" + ZSK_HOST


class OpenCliBridgeError(RuntimeError):
    """opencli 通道不可用或返回内容无效。"""


def locate_cli(which: Callable[[str], Optional[str]] = shutil.which) -> str:
    found = which("opencli")
    if not found:
        raise OpenCliBridgeError(
            "未找到 opencli 命令。请先安装 OpenCLI 命令行工具"
            "(npm install -g @jackwener/opencli)"
            "并在 Edge 中启用 OpenCLI 扩展,这是一次性动作。"
        )
    return found


def build_cli_command(cli_path: str) -> List[str]:
    """构造调用 opencli 的命令前缀。

    Windows 下 npm 生成的是 .cmd 批处理包装,参数需先经命令解释器解析:
    请求地址里的 & 会被当成命令分隔符,超长参数(例如把整份企业名单文件
    编码后塞进同一个参数)会被截断。两种情况下页面收到的 JS 都不完整,
    表现为 SyntaxError: missing ) after argument list。
    因此遇到 .cmd/.bat 包装时,改为用 node 直接执行其入口脚本,绕开命令
    解释器;    条件不满足时保持原命令不变。可用环境变量 RPI_OPENCLI_NODE 指定
    node 可执行文件位置,未指定时按 PATH 查找。
    """
    if os.name == "nt" and cli_path.lower().endswith((".cmd", ".bat")):
        entry = (
            Path(cli_path).resolve().parent
            / "node_modules"
            / "@jackwener"
            / "opencli"
            / "dist"
            / "src"
            / "main.js"
        )
        node = os.environ.get("RPI_OPENCLI_NODE") or shutil.which("node")
        if node and entry.is_file():
            return [node, str(entry)]
    return [cli_path]


def parse_profile_list_output(text: str) -> List[str]:
    profiles = []
    for line in str(text).splitlines():
        stripped = line.strip()
        if "— connected" in stripped or "- connected" in stripped:
            name = stripped.split("—")[0].split("- connected")[0].strip()
            if name and " " not in name:
                profiles.append(name)
    return profiles


class _Response:
    """鸭子类型兼容 requests 响应的最小子集。"""

    def __init__(self, status_code: int, headers: Dict[str, str], content: bytes):
        self.status_code = status_code
        self.headers = headers
        self.content = content

    def json(self) -> Any:
        try:
            return jsonlib.loads(self.content.decode("utf-8"))
        except (UnicodeDecodeError, jsonlib.JSONDecodeError) as exc:
            raise OpenCliBridgeError("接口返回的不是有效 JSON") from exc


class OpenCliTransport:
    """以 requests.Session 的 request 契约,把请求编译进浏览器页面上下文。"""

    def __init__(
        self,
        profile: Optional[str] = None,
        session: str = OPENCLI_SESSION,
        runner: Optional[Callable[..., Any]] = None,
        cli_path: Optional[str] = None,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.session_name = session
        self._runner = runner or self._default_runner
        self._cli_path = cli_path
        self._explicit_profile = profile
        self._resolved_profile: Optional[str] = None
        self._clock = clock

    @staticmethod
    def _default_runner(args: List[str], timeout: Optional[float] = None):
        try:
            return subprocess.run(
                args,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as exc:
            raise OSError("opencli 调用超时") from exc

    @property
    def cli_path(self) -> str:
        if self._cli_path is None:
            self._cli_path = locate_cli()
        return self._cli_path

    @property
    def profile(self) -> str:
        if self._resolved_profile:
            return self._resolved_profile
        if self._explicit_profile:
            self._resolved_profile = self._explicit_profile
            return self._resolved_profile
        output = self._run(["profile", "list"], timeout=30).stdout
        profiles = parse_profile_list_output(output or "")
        if not profiles:
            raise OpenCliBridgeError(
                "OpenCLI 扩展未连接。请确认 Edge 已打开且 OpenCLI 扩展已启用;"
                "有多个浏览器桥配置时,设置环境变量 RPI_OPENCLI_PROFILE 指定其一。"
            )
        if len(profiles) > 1:
            raise OpenCliBridgeError(
                "检测到多个 OpenCLI 连接配置({}),请设置环境变量 RPI_OPENCLI_PROFILE 指定其一。".format(
                    "、".join(profiles)
                )
            )
        self._resolved_profile = profiles[0]
        return self._resolved_profile

    def _run(self, args: List[str], timeout: Optional[float] = None):
        return self._runner([*build_cli_command(self.cli_path), *args], timeout=timeout)

    def _eval(self, js: str, timeout: float = 60.0) -> str:
        result = self._run(
            ["--profile", self.profile, "browser", self.session_name, "eval", js],
            timeout=timeout,
        )
        if getattr(result, "returncode", 0) != 0:
            raise OSError("opencli 调用失败:{}".format(getattr(result, "stderr", "") or ""))
        return _unwrap_eval_output(result.stdout or "")

    def open_page(self, url: str) -> None:
        self._run(
            ["--profile", self.profile, "browser", self.session_name, "open", url],
            timeout=60.0,
        )

    def get_url(self) -> str:
        result = self._run(
            ["--profile", self.profile, "browser", self.session_name, "get", "url"],
            timeout=30.0,
        )
        return str(result.stdout or "").strip()

    def _verify_js(self) -> str:
        return (
            "(async()=>{try{const r=await fetch("
            "'/open/industry_chain_api/v1/search/get_user_index',{credentials:'include'});"
            "const t=await r.text();return t;}catch(e){return 'ERR '+e.message;}})()"
        )

    def verify_session(self) -> bool:
        try:
            payload = jsonlib.loads(self._eval(self._verify_js()))
        except (jsonlib.JSONDecodeError, OpenCliBridgeError):
            return False
        return payload.get("status_code") == 0

    def wait_for_login(
        self,
        timeout: float = 900.0,
        poll: float = 5.0,
        clock: Optional[Callable[[], float]] = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        clock = clock or self._clock
        if not self.get_url().startswith(ZSK_BASE):
            self.open_page(ZSK_BASE + "/#/home_local")
        started = clock()
        while clock() - started < timeout:
            if self.verify_session():
                return
            sleep(poll)
        raise OpenCliBridgeError(
            "登录等待超时:请在 Edge 中的注协页面完成登录后重新执行登录验证。"
        )

    def request(
        self,
        method: str,
        url: str,
        *,
        params: Optional[Dict[str, Any]] = None,
        json: Any = None,
        files: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
        timeout: float = 30.0,
        _nav_retries: int = 0,
        **_: Any,
    ) -> _Response:
        if not url.startswith(ZSK_BASE + "/") and url != ZSK_BASE:
            raise OpenCliBridgeError("取数请求只允许发往注协知识库域名:" + ZSK_HOST)
        path = url[len(ZSK_BASE):]
        if params:
            path = path + ("&" if "?" in path else "?") + urlencode(params)
        extra_headers = dict(headers or {})
        body_js = "undefined"
        if files is not None:
            name, (filename, content, mime) = next(iter(files.items()))
            encoded = base64.b64encode(content).decode("ascii")
            # 注意:占位符外面不能再加引号。下面的替换值由 json.dumps 生成,
            # 本身就带引号;模板里若再包一层,会得到 ""file"" 这样的三重引号,
            # 页面端报 SyntaxError: missing ) after argument list。
            # atob 里的 __DATA__ 是原始 base64 文本,不经 json.dumps,故保留引号。
            body_js = (
                '(()=>{const fd=new FormData();'
                'fd.append(__NAME__,new File('
                '[Uint8Array.from(atob("__DATA__"))],__FILENAME__,'
                '{type:__MIME__}));return fd;})()'
            )
            body_js = (
                body_js.replace("__NAME__", jsonlib.dumps(name))
                .replace("__DATA__", encoded)
                .replace("__FILENAME__", jsonlib.dumps(filename))
                .replace("__MIME__", jsonlib.dumps(mime))
            )
            extra_headers.pop("Content-Type", None)
        elif json is not None:
            body_js = "JSON.stringify(" + jsonlib.dumps(json) + ")"
            extra_headers.setdefault("Content-Type", "application/json")
        header_items = ",".join(
            jsonlib.dumps(key) + ":" + jsonlib.dumps(value)
            for key, value in extra_headers.items()
        )
        if header_items:
            header_items = "," + header_items
        js = (
            '(async()=>{'
            'if(location.hostname!=="__HOST__")return JSON.stringify({nav:true});'
            'const xm=document.cookie.match(/XSRF-TOKEN=([^;]+)/);'
            'const h={"X-XSRF-TOKEN":(xm?xm[1]:"")__EXTRA__};'
            'try{'
            'const r=await fetch("__PATH__",{method:"__METHOD__",credentials:"include",'
            'headers:h,body:__BODY__});'
            'const buf=await r.arrayBuffer();'
            'const bytes=new Uint8Array(buf);'
            "let bin='';const chunk=0x8000;"
            'for(let i=0;i<bytes.length;i+=chunk)'
            'bin+=String.fromCharCode.apply(null,bytes.subarray(i,i+chunk));'
            'const hh={};r.headers.forEach((v,k)=>{hh[k]=v;});'
            'return JSON.stringify({s:r.status,h:hh,b:btoa(bin)});'
            '}catch(e){return JSON.stringify({s:0,h:{},b:"",e:String(e)});}})()'
        )
        js = (
            js.replace("__HOST__", ZSK_HOST)
            .replace("__EXTRA__", header_items)
            .replace("__PATH__", path)
            .replace("__METHOD__", method.upper())
            .replace("__BODY__", body_js)
        )
        output = self._eval(js, timeout=max(timeout, 60.0))
        try:
            payload = jsonlib.loads(output)
        except jsonlib.JSONDecodeError as exc:
            raise OpenCliBridgeError("页面取数返回内容无法解析") from exc
        if payload.get("nav"):
            if _nav_retries >= 2:
                raise OpenCliBridgeError("页面连续偏离注协域名，重新打开 2 次后仍未恢复")
            self.open_page(ZSK_BASE + "/#/home_local")
            return self.request(
                method,
                url,
                params=params,
                json=json,
                files=files,
                headers=headers,
                timeout=timeout,
                _nav_retries=_nav_retries + 1,
            )
        content = base64.b64decode(payload.get("b") or "")
        return _Response(int(payload.get("s", 0)), dict(payload.get("h") or {}), content)


def _unwrap_eval_output(stdout: str) -> str:
    text = stdout.strip()
    if not text:
        return ""
    if text.startswith('"') and text.endswith('"'):
        try:
            return jsonlib.loads(text)
        except jsonlib.JSONDecodeError:
            return text
    return text
