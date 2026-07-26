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
#import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from scripts.cicpa import opencli_setup


class FakeResponse:
    def __init__(self, data):
        self.data = data

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        return self.data


class OpenCliSetupTests(unittest.TestCase):
    def test_扩展只从_gitee_固定版本下载并校验(self):
        destination = Path(tempfile.mkdtemp(prefix="rpi_opencli_download_"))
        data = b"built-extension-file"
        digest = hashlib.sha256(data).hexdigest()
        requested = []

        def opener(url, timeout):
            requested.append((url, timeout))
            return FakeResponse(data)

        with mock.patch.object(
            opencli_setup,
            "OPENCLI_EXTENSION_FILES",
            {"dist/background.js": digest},
        ):
            path = opencli_setup.download_opencli_extension(
                destination,
                opener=opener,
            )

        self.assertEqual(path, destination)
        self.assertEqual(
            (destination / "dist" / "background.js").read_bytes(),
            data,
        )
        self.assertEqual(len(requested), 1)
        self.assertTrue(requested[0][0].startswith("https://gitee.com/"))
        self.assertIn(opencli_setup.OPENCLI_GITEE_COMMIT, requested[0][0])

    def test_可以识别当前浏览器已经登记的_opencli_扩展(self):
        user_data = Path(tempfile.mkdtemp(prefix="rpi_opencli_profile_"))
        profile = user_data / "Default"
        profile.mkdir()
        (profile / "Secure Preferences").write_text(
            json.dumps(
                {
                    "extensions": {
                        "settings": {
                            opencli_setup.OPENCLI_EXTENSION_ID: {
                                "disable_reasons": []
                            }
                        }
                    }
                }
            ),
            encoding="utf-8",
        )

        self.assertTrue(opencli_setup.opencli_extension_installed(user_data))

    def test_安装引导打开浏览器管理页和扩展文件夹(self):
        extension_dir = Path(tempfile.mkdtemp(prefix="rpi_opencli_guide_"))
        processes = []
        folders = []

        opencli_setup.open_install_guide(
            {
                "name": "Microsoft Edge",
                "executable_path": r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            },
            extension_dir,
            popen=lambda command: processes.append(command),
            folder_opener=lambda path: folders.append(path),
        )

        self.assertEqual(processes[0][-1], "edge://extensions")
        self.assertEqual(folders, [str(extension_dir)])


if __name__ == "__main__":
    unittest.main()
