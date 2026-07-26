# 关联方识别与核查 — OpenCLI 安装引导
# 本文件改编自 jackwener/OpenCLI(extension)(Apache-2.0)
# 上游项目: OpenCLI | 仓库: https://gitee.com/github_dep/opencli
# 上游协议: Apache-2.0
#
# 本地修改: CPA-Q(quanfanpro-code)
# 本文件的修改部分同样以 Apache-2.0 协议发布,与上游保持一致。
#"""从国内 Gitee 固定版本准备 OpenCLI 浏览器扩展。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
from typing import Callable, Dict, Optional
import urllib.request


OPENCLI_EXTENSION_ID = "ildkmabpimmkaediidaifkhjpohdnifk"
OPENCLI_GITEE_REPOSITORY = "https://gitee.com/github_dep/opencli"
OPENCLI_GITEE_COMMIT = "10baf02060bda736234c160927ccd4c9d0af82b6"
OPENCLI_RAW_BASE = (
    OPENCLI_GITEE_REPOSITORY
    + "/raw/"
    + OPENCLI_GITEE_COMMIT
    + "/extension/"
)
OPENCLI_EXTENSION_FILES: Dict[str, str] = {
    "manifest.json": "0f897bac24f1dfbe835d0fc9fbc6f32ae3750316acf7a6fcbe978668f4e6e28c",
    "popup.html": "9e0572175c7619056d77a618425cc7e6888eccfad7a1d5ffa74e8c784eece918",
    "popup.js": "c9e89f350ce1c97af3ff538debe63cda3dc98cbac119b99f41d4d411a26d5776",
    "dist/background.js": "871ca5a89cd869c5b91d255680923dd939d560d40d7c5763c121674af44b6140",
    "icons/icon-16.png": "41f055dfbcfa48a06d81c2733d686a39ab6c4aae3efa1e26247dddf4fa00ec95",
    "icons/icon-32.png": "415856e54bd712b13bc369b404cc0676080ad9fe5836bd82f5e066097ebc33e7",
    "icons/icon-48.png": "60a028b636ae832d995158ecec742c47cad7784187ab06108410154194217184",
    "icons/icon-128.png": "4b146c873a1f6e21f129e9d9ad7a6e56d48f64f536ef160468e967f265c424a4",
}


class OpenCliSetupError(RuntimeError):
    """OpenCLI 扩展下载或校验失败。"""


def default_extension_dir(local_app_data: Optional[str] = None) -> Path:
    value = local_app_data or os.environ.get("LOCALAPPDATA")
    if not value:
        raise OpenCliSetupError("无法确定当前用户的本地应用数据目录")
    return (
        Path(value)
        / "related-party-identification"
        / ("opencli-extension-" + OPENCLI_GITEE_COMMIT[:12])
    )


def _profiles(user_data_root: Path):
    root = Path(user_data_root)
    for path in root.iterdir() if root.is_dir() else ():
        if path.is_dir() and (path.name == "Default" or path.name.startswith("Profile ")):
            yield path


def opencli_extension_installed(user_data_root: Path) -> bool:
    """检查 Edge 或 Chrome 当前用户配置是否已登记 OpenCLI 扩展。"""
    for profile in _profiles(Path(user_data_root)):
        if (profile / "Extensions" / OPENCLI_EXTENSION_ID).is_dir():
            return True
        secure_preferences = profile / "Secure Preferences"
        if not secure_preferences.is_file():
            continue
        try:
            payload = json.loads(secure_preferences.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        settings = payload.get("extensions", {}).get("settings", {})
        extension = settings.get(OPENCLI_EXTENSION_ID)
        if isinstance(extension, dict) and not extension.get("disable_reasons"):
            return True
    return False


def _download_bytes(
    url: str,
    *,
    opener: Callable = urllib.request.urlopen,
) -> bytes:
    with opener(url, timeout=30) as response:
        return response.read()


def download_opencli_extension(
    destination: Optional[Path] = None,
    *,
    opener: Callable = urllib.request.urlopen,
) -> Path:
    """下载固定版本扩展并逐文件校验，不执行浏览器安装。"""
    root = Path(destination) if destination is not None else default_extension_dir()
    for relative, expected_hash in OPENCLI_EXTENSION_FILES.items():
        target = root / relative
        if target.is_file() and hashlib.sha256(target.read_bytes()).hexdigest() == expected_hash:
            continue
        data = _download_bytes(OPENCLI_RAW_BASE + relative, opener=opener)
        actual_hash = hashlib.sha256(data).hexdigest()
        if actual_hash != expected_hash:
            raise OpenCliSetupError("OpenCLI 扩展文件校验失败：{}".format(relative))
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(
            target.name + ".download." + str(os.getpid())
        )
        temporary.write_bytes(data)
        os.replace(temporary, target)
    return root


def open_install_guide(
    browser: Dict[str, str],
    extension_dir: Path,
    *,
    popen: Callable = subprocess.Popen,
    folder_opener: Optional[Callable[[str], None]] = None,
) -> None:
    """打开浏览器扩展管理页和已经准备好的扩展文件夹。"""
    executable = browser.get("executable_path")
    if not executable:
        raise OpenCliSetupError("未找到浏览器程序")
    page = (
        "edge://extensions"
        if browser.get("name") == "Microsoft Edge"
        else "chrome://extensions"
    )
    popen([str(executable), page])
    opener = folder_opener
    if opener is None:
        if os.name != "nt":
            raise OpenCliSetupError("此安装引导仅支持 Windows")
        opener = os.startfile
    opener(str(extension_dir))


def _main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="准备 OpenCLI 浏览器扩展")
    parser.add_argument("--download", action="store_true")
    args = parser.parse_args(argv)
    if not args.download:
        parser.error("请由关联方识别工作流调用")
    path = download_opencli_extension()
    print(
        json.dumps(
            {
                "status": "downloaded",
                "path": str(path),
                "source": OPENCLI_GITEE_REPOSITORY,
                "commit": OPENCLI_GITEE_COMMIT,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
