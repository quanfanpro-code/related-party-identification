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
#import unittest

from scripts.cicpa.edge_bridge import ZSK_URL, _BridgeState


class OpenCliBridgeTests(unittest.TestCase):
    def test_发送给扩展的命令只读取注协站点_cookie(self):
        state = _BridgeState()

        command = state.command()

        self.assertEqual(command["action"], "cookies")
        self.assertEqual(command["url"], ZSK_URL)
        self.assertEqual(command["session"], "related-party-identification")

    def test_只接受注协域名且不把_cookie_写入公开状态(self):
        state = _BridgeState()

        state.accept(
            {
                "id": state.command_id,
                "ok": True,
                "data": [
                    {
                        "name": "cicpa_token",
                        "value": "token-value",
                        "domain": ".cicpa.org.cn",
                    },
                    {
                        "name": "XSRF-TOKEN",
                        "value": "xsrf-value",
                        "domain": "zsk-cmis.cicpa.org.cn",
                    },
                    {
                        "name": "other",
                        "value": "ignore",
                        "domain": "example.com",
                    },
                ],
            }
        )

        self.assertTrue(state.done.is_set())
        self.assertEqual(
            state.cookies,
            {
                "cicpa_token": "token-value",
                "XSRF-TOKEN": "xsrf-value",
            },
        )
        self.assertEqual(state.error, "")


if __name__ == "__main__":
    unittest.main()
