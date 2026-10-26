# -*- coding: utf-8 -*-
"""128K 串行上下文调度外壳测试：初始化、唯一占用、短状态与候选完整分页。"""

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.串行上下文调度 import (  # noqa: E402
    DispatchError,
    candidate_page,
    complete_dispatch,
    handoff,
    initialize,
    recover_dispatch,
    sha256_file,
    start_dispatch,
    status,
)


class SerialContextTests(unittest.TestCase):
    def test_1m_bypass_and_second_active_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            state_path = Path(folder) / "任务状态.json"
            state_path.write_text('{"task_id":"demo","candidate_records":[]}', encoding="utf-8")
            backup = Path(folder) / "backup"
            self.assertEqual(initialize(state_path, "1M", "current-model", backup)["status"], "bypassed")
            self.assertFalse((state_path.parent / "上下文调度状态.json").exists())
            initialize(state_path, "128K", "current-model", backup)
            packet = Path(folder) / "候选发现记录.json"
            packet.write_text('{"candidates":[]}', encoding="utf-8")
            before = sha256_file(packet)
            start_dispatch(state_path, "confirm-scope", packet, "window-A", backup)
            with self.assertRaisesRegex(DispatchError, "活动任务"):
                start_dispatch(state_path, "next", packet, "window-B", backup)
            self.assertEqual(sha256_file(packet), before)

    def test_status_is_compact_and_candidate_pages_cover_original_set(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            records = [{"name": f"公司{i}", "paths": [f"路径{i}"]} for i in range(120)]
            state_path = root / "任务状态.json"
            state_path.write_text(json.dumps({
                "stage": "needs_user_confirmation", "status": "needs_user_confirmation",
                "message_zh": "确认范围", "candidate_count": 120,
                "report_paths": [], "candidate_records": records,
            }, ensure_ascii=False), encoding="utf-8")
            initialize(state_path, "128K", "current-model", root / "backup")
            summary = status(state_path)
            self.assertNotIn("candidate_records", summary)
            self.assertEqual(summary["candidate_count"], 120)
            source = root / "候选发现记录.json"
            source.write_text(json.dumps({"candidates": records}, ensure_ascii=False), encoding="utf-8")
            before = sha256_file(source)
            items = []
            pages = []
            for page in range(1, 4):
                payload = candidate_page(source, page=page, page_size=50)
                pages.append((payload["page"], payload["page_size"], payload["total"]))
                items.extend(payload["items"])
            self.assertEqual(pages, [(1, 50, 120), (2, 50, 120), (3, 50, 120)])
            self.assertEqual(items, records)
            self.assertEqual(sha256_file(source), before)
            with self.assertRaises(DispatchError):
                candidate_page(source, page=4, page_size=50)
            with self.assertRaises(DispatchError):
                candidate_page(source, page=0, page_size=50)

    def test_router_contract_and_controller_rotation(self):
        root = Path(__file__).resolve().parents[2]
        skill = (root / "SKILL.md").read_text(encoding="utf-8-sig")
        guide = (root / "references" / "128K串行上下文.md").read_text(encoding="utf-8-sig")
        self.assertIn("本次使用 1M 还是 128K 上下文模型", skill)
        for phrase in ("前台", "串行", "同一模型", "候选页", "result_sha256"):
            self.assertIn(phrase, guide)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            state_path = root / "任务状态.json"
            state_path.write_text(json.dumps({
                "task_id": "demo", "stage": "running", "status": "running",
                "message_zh": "匿名调度", "candidate_count": 0,
                "report_paths": [], "candidate_records": [],
            }, ensure_ascii=False), encoding="utf-8")
            initialize(state_path, "128K", "current-model", root / "backup")
            packet = root / "输入.json"
            packet.write_text('{"items":[]}', encoding="utf-8")
            for index in range(20):
                task_id = f"task-{index}"
                start_dispatch(state_path, task_id, packet, f"window-{index}", root / "backup")
                result = root / f"结果-{index}.json"
                result.write_text('{"saved":true}', encoding="utf-8")
                complete_dispatch(state_path, task_id, result, root / "backup")
            self.assertTrue(status(state_path)["rotation_required"])

    def test_cli_subcommands_dispatch(self):
        with tempfile.TemporaryDirectory() as folder:
            from scripts.串行上下文调度 import main

            root = Path(folder)
            state_path = root / "任务状态.json"
            state_path.write_text('{"stage":"running","status":"running","message_zh":"匿名","candidate_count":0,"report_paths":[]}', encoding="utf-8")
            backup = root / "backup"
            base = ["--state-path", str(state_path)]
            self.assertEqual(self.captured_main(main, ["初始化", *base, "--mode", "128K", "--model-binding", "current-model", "--backup-root", str(backup)]), 0)
            packet = root / "输入.json"
            packet.write_text('{"items":[]}', encoding="utf-8")
            self.assertEqual(self.captured_main(main, ["开始", *base, "--task-id", "T1", "--packet", str(packet), "--window-id", "W1", "--backup-root", str(backup)]), 0)
            result = root / "结果.json"
            result.write_text('{"saved":true}', encoding="utf-8")
            self.assertEqual(self.captured_main(main, ["完成", *base, "--task-id", "T1", "--result", str(result), "--backup-root", str(backup)]), 0)
            self.assertEqual(self.captured_main(main, ["状态", *base]), 0)
            self.assertNotIn("candidate_records", self.last_stdout)
            source = root / "候选发现记录.json"
            source.write_text('{"candidates":[{"name":"甲公司"},{"name":"乙公司"}]}', encoding="utf-8")
            self.assertEqual(self.captured_main(main, ["候选页", "--source", str(source), "--page", "1", "--page-size", "1"]), 0)
            page = json.loads(self.last_stdout)
            self.assertEqual(page["total"], 2)
            self.assertEqual(len(page["items"]), 1)
            self.assertEqual(self.captured_main(main, ["交接", *base]), 0)
            self.assertEqual(json.loads(self.last_stdout)["context_mode"], "128K")

    def captured_main(self, func, argv):
        import contextlib
        import io

        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            code = func(argv)
        self.last_stdout = buffer.getvalue()
        return code


if __name__ == "__main__":
    unittest.main()
