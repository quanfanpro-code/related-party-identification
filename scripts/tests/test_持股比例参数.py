# -*- coding: utf-8 -*-
"""持股穿透比例参数：默认 20%，任务开始前在对话中向用户确认是否调整。"""

from pathlib import Path
from types import SimpleNamespace
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.cicpa.client import CompanyMatch, EquityEdge  # noqa: E402
from scripts.cicpa.exporter import ExportState, save_export_state  # noqa: E402
from scripts.discovery import (  # noqa: E402
    DiscoveryPolicy,
    DiscoveryResult,
    discover,
)
from scripts.related_party_workflow import (  # noqa: E402
    TaskState,
    resume_or_init_state,
    run_workflow,
    save_task_state,
)


class _FakeClient:
    def __init__(self, equity):
        self.equity = equity

    def search_companies(self, _name):
        return [CompanyMatch(org_id="a", name="甲公司")]

    def get_equity_relations(self, company_id):
        return list(self.equity.get(company_id, []))

    def get_key_personnel(self, _company_id):
        return []


def _edge(org_id, name, ratio):
    return EquityEdge(company_id=org_id, company_name=name, ratio=ratio, direction="investment")


class DiscoverThresholdTests(unittest.TestCase):
    def _client(self):
        return _FakeClient({"a": [_edge("b", "乙公司", 15.0)]})

    def test_默认阈值百分之二十时低持股不纳入候选(self):
        result = discover("甲公司", client=self._client(), policy=DiscoveryPolicy())
        self.assertNotIn("乙公司", result.candidates)

    def test_对话确认低比例后低持股纳入候选(self):
        result = discover(
            "甲公司", client=self._client(), policy=DiscoveryPolicy(equity_threshold=10.0)
        )
        self.assertIn("乙公司", result.candidates)
        self.assertEqual(result.candidates["乙公司"].ratio, 15.0)


class StateThresholdTests(unittest.TestCase):
    def _args(self, path, **overrides):
        base = dict(
            state_path=path,
            mode="discovery",
            audited_entity="甲公司",
            input=[],
            output_dir="",
            approved_depth=5,
            candidate_cap=100,
            as_of_date="",
        )
        base.update(overrides)
        return SimpleNamespace(**base)

    def test_新建任务未指定时默认百分之二十(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            state = resume_or_init_state(self._args(path))
            self.assertEqual(state.equity_threshold, 20.0)

    def test_新建任务按对话确认的比例执行(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            state = resume_or_init_state(self._args(path, equity_threshold=10.0))
            self.assertEqual(state.equity_threshold, 10.0)

    def test_续跑未指定时保留任务原比例(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            save_task_state(
                path,
                TaskState(
                    task_id="任务状态",
                    mode="discovery",
                    audited_entity="甲公司",
                    equity_threshold=30.0,
                ),
            )
            state = resume_or_init_state(self._args(path))
            self.assertEqual(state.equity_threshold, 30.0)

    def test_续跑按本次对话确认的比例更新(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            save_task_state(
                path,
                TaskState(
                    task_id="任务状态",
                    mode="discovery",
                    audited_entity="甲公司",
                    equity_threshold=30.0,
                ),
            )
            state = resume_or_init_state(self._args(path, equity_threshold=15.0))
            self.assertEqual(state.equity_threshold, 15.0)


class _FakeExporter:
    def __init__(self, extract_dir):
        self.extract_dir = Path(extract_dir)

    def start_export(self, company_names, state_path):
        state = ExportState(
            batch_no="batch-1",
            company_names=list(company_names),
            created_at="2026-10-27 10:00",
            triggered_at="2026-10-27 10:00",
            staging_dir=str(Path(state_path).parent),
            status="waiting",
            required_dimensions=["S0000002"],
            pending_groups=[["S0000002"]],
        )
        save_export_state(state_path, state)
        return state

    def run_to_completion(self, state, state_path, max_polls=18):
        state.dimension_results = {"S0000002": "completed"}
        state.status = "completed"
        state.extract_dir = str(self.extract_dir)
        save_export_state(state_path, state)
        return self.extract_dir


class WorkflowThresholdTests(unittest.TestCase):
    def test_发现引擎与报告范围使用任务持股比例(self):
        captured = {}

        def fake_discoverer(_entity, client=None, policy=None, **_kwargs):
            captured["policy"] = policy
            return DiscoveryResult(status="completed", seed_name="甲公司", candidates={})

        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_ratio_"))
        data_dir = artifact_dir / "甲公司_注协原始导出"
        data_dir.mkdir()
        checker_calls = []

        def fake_checker(**kwargs):
            checker_calls.append(kwargs)
            return SimpleNamespace(
                output_path=kwargs["output_path"], summary=[], errors=[], limitations=[]
            )

        state = TaskState(
            task_id="ratio-1",
            mode="discovery",
            audited_entity="甲公司",
            output_dir=str(artifact_dir),
            equity_threshold=10.0,
        )

        run_workflow(
            state,
            artifact_dir / "task.json",
            client_factory=lambda: object(),
            exporter_factory=lambda _client: _FakeExporter(data_dir),
            discoverer=fake_discoverer,
            checker=fake_checker,
        )

        self.assertEqual(captured["policy"].equity_threshold, 10.0)
        self.assertEqual(checker_calls[0]["scope_metadata"]["持股穿透比例"], 10.0)


if __name__ == "__main__":
    unittest.main()
