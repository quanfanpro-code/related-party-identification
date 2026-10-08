# -*- coding: utf-8 -*-
# 批次二：导出器容错修复的行为测试。
# 缺文件两次后记 missing 放行、完成消息如实计数、空工作簿不判 no_data、
# 同名冲突改名留双份、跨盘符 ZIP 路径不裸抛。
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
import zipfile

import openpyxl

import scripts.cicpa.exporter as exporter_module

from scripts.cicpa.exporter import (
    CicpaExporter,
    ExportValidationError,
    ExportWorkbookIncomplete,
)
from scripts.tests.test_exporter import (
    FakeClient,
    make_dimension_zip,
    make_statistics_only_zip,
    waiting_state,
)


def make_workbook_bytes(rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in rows:
        sheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def make_partial_zip(company_names, dimension_rows, statistics_counts=None):
    """组装同时含维度工作簿与统计表的批次 ZIP。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for code, has_rows in dimension_rows:
            rows = [["企业名称", "核查字段"]]
            if has_rows:
                rows += [[name, "有效数据"] for name in company_names]
            archive.writestr(
                exporter_module.DIMENSION_NAMES[code] + ".xlsx",
                make_workbook_bytes(rows),
            )
        if statistics_counts:
            rows = [
                ["导出维度记录数统计"],
                ["公司名称", "现用名"]
                + [exporter_module.DIMENSION_NAMES[code] for code in statistics_counts],
            ]
            for name in company_names:
                rows.append(
                    [name, name] + [statistics_counts[code] for code in statistics_counts]
                )
            archive.writestr("统计表.xlsx", make_workbook_bytes(rows))
    return output.getvalue()


def make_empty_workbook_zip(code):
    """维度工作簿连表头都没有（空白工作簿）。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            exporter_module.DIMENSION_NAMES[code] + ".xlsx",
            make_workbook_bytes([]),
        )
    return output.getvalue()


def make_garbage_workbook_zip(code):
    """维度文件是损坏的字节流，openpyxl 无法作为工作簿打开。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            exporter_module.DIMENSION_NAMES[code] + ".xlsx",
            b"not a real workbook",
        )
    return output.getvalue()


def make_cross_drive_zip(anchor_dir):
    """成员的绝对路径落在 anchor_dir 所在盘符之外，触发 commonpath 的 ValueError。"""
    anchor_drive = Path(anchor_dir).resolve().drive.upper()
    other_drive = "D:" if anchor_drive != "D:" else "E:"
    output = BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr(other_drive + "\\outside-evidence.txt", "unsafe")
    return output.getvalue()


def make_task(task_id):
    return {
        "task_id": task_id,
        "batch_id": "batch-new",
        "type": "批量查询",
        "status": 1,
        "url": "/{}.zip".format(task_id),
    }


def single_dimension_state(artifact_dir, codes):
    state = waiting_state(artifact_dir)
    state.required_dimensions = list(codes)
    state.active_group = list(codes)
    state.task_id = "task-1"
    return state


class MissingRetryLimitTests(unittest.TestCase):
    """缺文件维度：重新触发仍拿不到后记 missing 并继续，不再无限重排。"""

    def test_缺文件两次后记未取得且不再重排(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_missing_limit_"))
        state_path = artifact_dir / "state.json"
        code = "S0000107"
        state = single_dimension_state(artifact_dir, [code])
        # 批次 ZIP 只有统计表且该维度计数为正：是"没拿到文件"而非"无数据"
        exporter = CicpaExporter(
            FakeClient(binary_result=make_statistics_only_zip("甲公司", code, 3)),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-1"), state_path)

        self.assertEqual(completed, [])
        self.assertNotIn(code, state.dimension_results)
        self.assertEqual(state.missing_counts[code], 1)
        self.assertEqual(state.pending_groups, [[code]])

        state.active_group = state.pending_groups.pop(0)
        state.task_id = "task-2"
        completed = exporter.download_current_group(state, make_task("task-2"), state_path)

        self.assertEqual(completed, [])
        self.assertEqual(state.dimension_results[code], "missing")
        self.assertEqual(state.pending_groups, [])
        self.assertEqual(state.status, "completed")
        self.assertIn("未取得", state.message_zh)
        self.assertNotIn("已全部完成", state.message_zh)

    def test_完成消息如实统计取得与未取得(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_honest_message_"))
        state_path = artifact_dir / "state.json"
        codes = ["S0000002", "S0000107", "S0000018"]
        state = single_dimension_state(artifact_dir, codes)
        # 第一批：基础工商有数据、发票信息经统计表确认为零、经营异常缺文件
        exporter = CicpaExporter(
            FakeClient(
                binary_result=make_partial_zip(
                    ["甲公司"],
                    [("S0000002", True)],
                    statistics_counts={"S0000107": 0, "S0000018": 2},
                )
            ),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-1"), state_path)

        self.assertEqual(set(completed), {"S0000002", "S0000107"})
        self.assertEqual(state.dimension_results["S0000107"], "no_data")
        self.assertNotIn("S0000018", state.dimension_results)

        state.active_group = state.pending_groups.pop(0)
        state.task_id = "task-2"
        exporter.download_current_group(state, make_task("task-2"), state_path)

        self.assertEqual(state.dimension_results["S0000018"], "missing")
        self.assertEqual(state.status, "completed")
        self.assertEqual(state.message_zh, "完成 2 个维度，1 个维度未取得")
        self.assertTrue((artifact_dir / "基础工商信息.xlsx").exists())

    def test_单维度两个轮询窗口失败后记未取得并继续(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_poll_missing_"))
        state_path = artifact_dir / "state.json"
        state = waiting_state(artifact_dir)
        state.required_dimensions = ["S0000002", "S0000006"]
        state.active_group = ["S0000002"]
        state.task_id = "task-stuck"
        state.pending_groups = [["S0000006"]]
        client = FakeClient(
            json_results=[
                {"status_code": 0, "data": {"list": []}},
                {"status_code": 0, "data": {"task_id": "task-next"}},
            ]
        )
        exporter = CicpaExporter(client, staging_root=artifact_dir)

        exporter.handle_poll_window_failure(state, state_path)
        exporter.handle_poll_window_failure(state, state_path)

        self.assertNotEqual(state.status, "paused")
        self.assertEqual(state.dimension_results["S0000002"], "missing")
        self.assertEqual(state.active_group, [])
        self.assertIn("基础工商信息", state.message_zh)
        self.assertIn("未取得", state.message_zh)

        triggered = exporter.trigger_next_group(state, state_path)

        self.assertTrue(triggered)
        self.assertEqual(state.active_group, ["S0000006"])


class EmptyWorkbookTests(unittest.TestCase):
    """空工作簿区分：连表头都没有按缺文件走 missing；表头完整无数据行才是 no_data。"""

    def test_连表头都没有的工作簿校验报残缺(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_headerless_unit_"))
        workbook_path = artifact_dir / "发票信息.xlsx"
        workbook_path.write_bytes(make_workbook_bytes([]))

        with self.assertRaises(ExportWorkbookIncomplete):
            CicpaExporter._validate_dimension_workbook(workbook_path, ["甲公司"])

    def test_表头完整但没有数据行仍记无数据(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_header_only_"))
        workbook_path = artifact_dir / "发票信息.xlsx"
        workbook_path.write_bytes(make_workbook_bytes([["企业名称", "核查字段"]]))

        result = CicpaExporter._validate_dimension_workbook(workbook_path, ["甲公司"])

        self.assertEqual(result, "no_data")

    def test_连表头都没有的工作簿不判无数据(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_headerless_"))
        state_path = artifact_dir / "state.json"
        code = "S0000107"
        state = single_dimension_state(artifact_dir, [code])
        exporter = CicpaExporter(
            FakeClient(binary_result=make_empty_workbook_zip(code)),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-1"), state_path)

        self.assertEqual(completed, [])
        self.assertNotIn(code, state.dimension_results)
        self.assertEqual(state.missing_counts[code], 1)
        self.assertEqual(state.pending_groups, [[code]])

        state.active_group = state.pending_groups.pop(0)
        state.task_id = "task-2"
        exporter.download_current_group(state, make_task("task-2"), state_path)

        self.assertEqual(state.dimension_results[code], "missing")
        self.assertNotEqual(state.dimension_results[code], "no_data")
        # 残缺文件不进证据主目录，只留在批次解压目录里
        self.assertFalse((artifact_dir / "发票信息.xlsx").exists())

    def test_损坏无法打开的工作簿按缺文件处理(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_garbage_"))
        state_path = artifact_dir / "state.json"
        code = "S0000107"
        state = single_dimension_state(artifact_dir, [code])
        exporter = CicpaExporter(
            FakeClient(binary_result=make_garbage_workbook_zip(code)),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-1"), state_path)

        self.assertEqual(completed, [])
        self.assertNotIn(code, state.dimension_results)
        self.assertEqual(state.missing_counts[code], 1)
        self.assertEqual(state.pending_groups, [[code]])


class ConflictKeepBothTests(unittest.TestCase):
    """同名不同内容：新文件改名保留、沿用旧文件、留痕后继续，不暂停不覆盖。"""

    def test_冲突改名留双份且任务继续(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_conflict_keep_"))
        state_path = artifact_dir / "state.json"
        code = "S0000002"
        existing = artifact_dir / "基础工商信息.xlsx"
        existing.write_bytes(b"old evidence")
        state = single_dimension_state(artifact_dir, [code])
        exporter = CicpaExporter(
            FakeClient(binary_result=make_dimension_zip(["甲公司"], [(code, True)])),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-new"), state_path)

        self.assertEqual(completed, [code])
        self.assertEqual(state.dimension_results[code], "completed")
        self.assertNotEqual(state.status, "paused")
        self.assertEqual(existing.read_bytes(), b"old evidence")
        renamed = artifact_dir / "基础工商信息.本批.xlsx"
        self.assertTrue(renamed.exists())
        self.assertNotEqual(renamed.read_bytes(), b"old evidence")
        conflicts = [
            item for item in state.batch_history if item.get("result") == "file_conflict"
        ]
        self.assertEqual(len(conflicts), 1)
        self.assertIn("不一致", conflicts[0]["detail"])
        self.assertIn("基础工商信息.本批.xlsx", conflicts[0]["detail"])


class CrossDriveZipTests(unittest.TestCase):
    """ZIP 成员落在其他盘符时 commonpath 会抛 ValueError：按校验失败处理，不裸抛。"""

    def test_跨盘符路径不裸抛(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_cross_drive_"))
        archive_path = artifact_dir / "cross.zip"
        archive_path.write_bytes(make_cross_drive_zip(artifact_dir))
        extract_dir = artifact_dir / "extract"
        extract_dir.mkdir()

        try:
            CicpaExporter._safe_extract(archive_path, extract_dir)
        except ExportValidationError:
            pass
        except ValueError as exc:
            self.fail("跨盘符路径裸抛 ValueError：{}".format(exc))
        else:
            self.fail("跨盘符路径未被拒绝")

    def test_整批无法安全解包时按缺文件继续(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_archive_invalid_"))
        state_path = artifact_dir / "state.json"
        code = "S0000002"
        state = single_dimension_state(artifact_dir, [code])
        exporter = CicpaExporter(
            FakeClient(binary_result=make_cross_drive_zip(artifact_dir)),
            artifact_dir=artifact_dir,
        )

        completed = exporter.download_current_group(state, make_task("task-1"), state_path)

        self.assertEqual(completed, [])
        self.assertNotIn(code, state.dimension_results)
        self.assertEqual(state.missing_counts[code], 1)
        self.assertEqual(state.pending_groups, [[code]])
        self.assertNotEqual(state.status, "paused")
        self.assertIn(
            "archive_invalid",
            [item.get("result") for item in state.batch_history],
        )

        state.active_group = state.pending_groups.pop(0)
        state.task_id = "task-2"
        exporter.download_current_group(state, make_task("task-2"), state_path)

        self.assertEqual(state.dimension_results[code], "missing")
        self.assertEqual(state.status, "completed")


if __name__ == "__main__":
    unittest.main()
