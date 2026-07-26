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
from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace
import unittest

from scripts.cicpa.client import CompanyMatch, EquityEdge, KeyPerson
from scripts.discovery import (
    DiscoveryPolicy,
    SeedExportCandidate,
    discover,
)


class FakeDiscoveryClient:
    def __init__(self, *, matches, equity=None, personnel=None):
        self.matches = matches
        self.equity = equity or {}
        self.personnel = personnel or {}
        self.equity_calls = []
        self.personnel_calls = []

    def search_companies(self, _name):
        return list(self.matches)

    def get_equity_relations(self, company_id):
        self.equity_calls.append(company_id)
        return list(self.equity.get(company_id, []))

    def get_key_personnel(self, company_id):
        self.personnel_calls.append(company_id)
        return list(self.personnel.get(company_id, []))


def company(org_id, name):
    return CompanyMatch(org_id=org_id, name=name)


def investment(org_id, name, ratio):
    return EquityEdge(
        company_id=org_id,
        company_name=name,
        ratio=ratio,
        direction="investment",
    )


class DiscoveryDepthTests(unittest.TestCase):
    def test_默认先扩展两层并保留关系路径(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={
                "a": [investment("b", "乙公司", 60.0)],
                "b": [investment("c", "丙公司", 55.0)],
                "c": [investment("d", "丁公司", 51.0)],
            },
        )

        result = discover("甲公司", client=client, policy=DiscoveryPolicy())

        self.assertEqual(result.status, "completed")
        self.assertEqual(set(result.candidates), {"乙公司", "丙公司"})
        self.assertEqual(result.candidates["丙公司"].depth, 2)
        self.assertIn("甲公司", result.candidates["乙公司"].paths[0])
        self.assertNotIn("d", client.equity_calls)

    def test_用户确认后可继续到第五层(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={
                "a": [investment("b", "乙公司", 60.0)],
                "b": [investment("c", "丙公司", 60.0)],
                "c": [investment("d", "丁公司", 60.0)],
            },
        )

        result = discover(
            "甲公司",
            client=client,
            policy=DiscoveryPolicy(),
            approved_depth=5,
        )

        self.assertIn("丁公司", result.candidates)
        self.assertEqual(result.candidates["丁公司"].depth, 3)

    def test_循环持股只保留一次候选(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={
                "a": [investment("b", "乙公司", 60.0)],
                "b": [
                    investment("a", "甲公司", 60.0),
                    investment("c", "丙公司", 60.0),
                ],
            },
        )

        result = discover("甲公司", client=client)

        self.assertEqual(set(result.candidates), {"乙公司", "丙公司"})


class DiscoveryRedFlagTests(unittest.TestCase):
    def test_低持股但共同关键人员仍保留且不继续穿透(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={
                "a": [investment("b", "乙公司", 10.0)],
                "b": [investment("c", "丙公司", 80.0)],
            },
            personnel={
                "a": [KeyPerson(name="张三", role="董事")],
                "b": [KeyPerson(name="张三", role="监事")],
            },
        )

        result = discover("甲公司", client=client)

        candidate = result.candidates["乙公司"]
        self.assertIn("共同关键人员：张三", candidate.reasons)
        self.assertIn("低持股比例但存在其他红旗", candidate.notes)
        self.assertNotIn("c", client.equity_calls)

    def test_被审计单位导出中的客户供应商只作为一层候选(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={"a": []},
        )
        export_candidates = [
            SimpleNamespace(
                name="客户公司",
                relation_type="公开客户关系",
                marked_related=True,
                red_flags=(),
            ),
            SeedExportCandidate(
                name="供应商公司",
                relation_type="公开供应商关系",
            ),
        ]

        result = discover(
            "甲公司",
            client=client,
            seed_export_candidates=export_candidates,
        )

        self.assertEqual(set(result.candidates), {"客户公司", "供应商公司"})
        self.assertIn(
            "被审计单位完整维度导出：公开客户关系",
            result.candidates["客户公司"].reasons,
        )
        self.assertNotIn(
            "导出表内关联标注",
            result.candidates["客户公司"].reasons,
        )
        self.assertEqual(client.equity_calls, ["a"])


class DiscoveryBoundaryTests(unittest.TestCase):
    def test_第百零一个候选出现时暂停并保留前一百个(self):
        edges = [investment("org-{}".format(i), "候选{:03d}公司".format(i), 30.0) for i in range(101)]
        client = FakeDiscoveryClient(
            matches=[company("a", "甲公司")],
            equity={"a": edges},
        )

        result = discover("甲公司", client=client)

        self.assertEqual(result.status, "needs_user_confirmation")
        self.assertEqual(len(result.candidates), 100)
        self.assertTrue(result.stopped_at_cap)

    def test_重名且没有唯一精确结果时请求用户确认(self):
        client = FakeDiscoveryClient(
            matches=[company("a", "甲集团"), company("b", "甲集团有限公司")]
        )

        result = discover("甲", client=client)

        self.assertEqual(result.status, "needs_company_confirmation")
        self.assertEqual(len(result.company_matches), 2)
        self.assertEqual(result.candidates, {})


class CompatibilityEntryTests(unittest.TestCase):
    def test_旧入口_help_不再依赖外部_skill(self):
        skill_root = Path(__file__).resolve().parents[2]
        script = skill_root / "scripts" / "discover_scope.py"

        completed = subprocess.run(
            [sys.executable, "-X", "utf8", str(script), "--help"],
            cwd=skill_root,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertIn("初始穿透深度", completed.stdout)


if __name__ == "__main__":
    unittest.main()
