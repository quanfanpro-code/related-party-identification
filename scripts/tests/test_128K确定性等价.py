# -*- coding: utf-8 -*-
"""同一离线导出目录走原核查与 128K 调度的确定性等价验证。"""

import json
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.related_party_check import run_check  # noqa: E402
from scripts.串行上下文调度 import (  # noqa: E402
    complete_dispatch,
    initialize,
    start_dispatch,
    status,
)
from scripts.tests.test_报告追溯 import write_basic, write_dimension  # noqa: E402


def report_business_signature(path: Path) -> list:
    workbook = openpyxl.load_workbook(path, data_only=False)
    signature = []
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows():
            if row[0].value == "报告生成时间":
                continue
            rows.append([
                (
                    cell.value,
                    cell.data_type,
                    cell.number_format,
                    str(cell.font),
                    str(cell.fill),
                    str(cell.border),
                    str(cell.alignment),
                    str(cell.protection),
                    cell.hyperlink.target if cell.hyperlink else None,
                )
                for cell in row
            ])
        validations = [
            (item.type, item.formula1, str(item.sqref))
            for item in sheet.data_validations.dataValidation
        ]
        signature.append({
            "title": sheet.title,
            "state": sheet.sheet_state,
            "freeze": str(sheet.freeze_panes),
            "merged": sorted(str(item) for item in sheet.merged_cells.ranges),
            "rows": rows,
            "validations": validations,
        })
    workbook.close()
    if not signature:
        raise AssertionError("报告签名为空，未实际读取工作表")
    return signature


class OfflineEquivalenceTests(unittest.TestCase):
    def test_same_offline_export_produces_same_report_with_serial_context(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source"
            write_basic(source, ["甲公司", "乙公司"], shared_phone=True)
            write_dimension(source, "主要人员（高管）.xlsx", [[1, "甲公司", "王六", "董事"], [2, "乙公司", "王六", "经理"]])
            baseline = root / "baseline.xlsx"
            optimized = root / "optimized.xlsx"
            common = dict(data_dir=source, target_names=["甲公司"], as_of_date=date(2026, 10, 7))
            run_check(output_path=baseline, **common)
            state_path = root / "任务状态.json"
            state_path.write_text(json.dumps({
                "task_id": "offline", "stage": "running", "status": "running",
                "message_zh": "离线等价", "candidate_count": 0,
                "report_paths": [], "candidate_records": [],
            }, ensure_ascii=False), encoding="utf-8")
            initialize(state_path, "128K", "test-current-model", root / "backup")
            start_dispatch(state_path, "offline-check", source / "基础工商信息.xlsx", "window-A", root / "backup")
            run_check(output_path=optimized, **common)
            complete_dispatch(state_path, "offline-check", optimized, root / "backup")
            baseline_signature = report_business_signature(baseline)
            optimized_signature = report_business_signature(optimized)
            self.assertEqual(8, len(baseline_signature), f"报告应固定八张工作表，实际 {len(baseline_signature)} 张")
            self.assertEqual(baseline_signature, optimized_signature)
            self.assertIsNone(status(state_path)["active_dispatch"])


if __name__ == "__main__":
    unittest.main()
