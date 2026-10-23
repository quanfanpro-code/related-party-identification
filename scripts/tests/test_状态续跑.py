# -*- coding: utf-8 -*-
"""续跑与状态保护：确认加大候选上限后不得重跑覆盖已有进度。"""

from pathlib import Path
from types import SimpleNamespace
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.cicpa.client import CompanyMatch, EquityEdge  # noqa: E402
from scripts.discovery import DiscoveryPolicy, discover  # noqa: E402
from scripts.related_party_workflow import (  # noqa: E402
    StateConflict,
    TaskState,
    resume_or_init_state,
    save_task_state,
)


class _FakeClient:
    def __init__(self, equity):
        self.equity = equity
        self.equity_calls = []
        self.search_calls = 0

    def search_companies(self, _name):
        self.search_calls += 1
        return [CompanyMatch(org_id="a", name="甲公司")]

    def get_equity_relations(self, company_id):
        self.equity_calls.append(company_id)
        return list(self.equity.get(company_id, []))

    def get_key_personnel(self, _company_id):
        return []


def _edge(org_id, name, ratio):
    return EquityEdge(company_id=org_id, company_name=name, ratio=ratio, direction="investment")


class DiscoveryResumeTests(unittest.TestCase):
    def _equity(self):
        return {
            "a": [_edge("b", "乙公司", 60.0), _edge("c", "丙公司", 70.0)],
            "b": [_edge("d", "丁公司", 55.0)],
            "c": [],
            "d": [],
        }

    def test_续跑从待展开队列继续且不重查已展开主体(self):
        snapshots = []
        paused = discover(
            "甲公司",
            client=_FakeClient(self._equity()),
            policy=DiscoveryPolicy(candidate_cap=1),
            on_progress=snapshots.append,
        )
        self.assertEqual(paused.status, "needs_user_confirmation")
        self.assertTrue(snapshots, "暂停前必须留下可续跑的快照")
        self.assertIn("乙公司", [node["name"] for node in snapshots[-1]["queue"]])

        client = _FakeClient(self._equity())
        resumed = discover(
            "甲公司",
            client=client,
            policy=DiscoveryPolicy(candidate_cap=10),
            resume=snapshots[-1],
        )
        self.assertEqual(resumed.status, "completed")
        self.assertEqual(set(resumed.candidates), {"乙公司", "丙公司", "丁公司"})
        self.assertEqual(client.search_calls, 0, "续跑不应重新检索被审计单位")
        self.assertEqual(sorted(client.equity_calls), ["a", "b", "c", "d"], "每个主体只展开一次")


class StateResumeTests(unittest.TestCase):
    def _args(self, path, **overrides):
        base = dict(
            state_path=path,
            mode="discovery",
            audited_entity="甲公司",
            input=[],
            output_dir="",
            approved_depth=5,
            candidate_cap=300,
            as_of_date="",
        )
        base.update(overrides)
        return SimpleNamespace(**base)

    def test_已有状态按续跑处理并保留候选与队列(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            save_task_state(
                path,
                TaskState(
                    task_id="任务状态",
                    mode="discovery",
                    audited_entity="甲公司",
                    candidate_cap=100,
                    candidate_count=1,
                    candidate_records=[{"name": "乙公司"}],
                    discovery_progress={"queue": [{"name": "乙公司", "org_id": "b", "depth": 1, "path": []}]},
                ),
            )

            state = resume_or_init_state(self._args(path))

            self.assertEqual(state.candidate_cap, 300)
            self.assertEqual(state.candidate_count, 1)
            self.assertEqual([item["name"] for item in state.candidate_records], ["乙公司"])
            self.assertEqual(state.discovery_progress["queue"][0]["name"], "乙公司")

    def test_状态文件属于其他单位时拒绝覆盖(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "任务状态.json"
            save_task_state(path, TaskState(task_id="任务状态", mode="discovery", audited_entity="乙公司"))

            with self.assertRaises(StateConflict):
                resume_or_init_state(self._args(path))


if __name__ == "__main__":
    unittest.main()
