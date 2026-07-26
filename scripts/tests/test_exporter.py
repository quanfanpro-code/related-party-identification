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
from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

import openpyxl

from scripts.cicpa.exporter import (
    CicpaExporter,
    ExportState,
    ExportTaskNotCorrelated,
    ExportValidationError,
    load_export_state,
    save_export_state,
)


class FakeClient:
    def __init__(self, json_results=None, binary_result=b""):
        self.json_results = list(json_results or [])
        self.binary_result = binary_result
        self.json_calls = []
        self.binary_calls = []

    def request_json(self, method, url, **kwargs):
        self.json_calls.append((method, url, kwargs))
        return self.json_results.pop(0)

    def request_bytes(self, method, url, **kwargs):
        self.binary_calls.append((method, url, kwargs))
        return self.binary_result


def make_export_zip(company_names, artifact_dir):
    workbook_path = Path(artifact_dir) / "基础工商信息.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["企业名称", "统一社会信用代码"])
    for index, name in enumerate(company_names, 1):
        sheet.append([name, "91510000TEST{:06d}".format(index)])
    workbook.save(workbook_path)
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(workbook_path, arcname="基础工商信息.xlsx")
    return output.getvalue()


def waiting_state(artifact_dir):
    return ExportState(
        batch_no="batch-new",
        company_names=["甲公司"],
        created_at="2026-07-26T15:00:00+08:00",
        triggered_at="2026-07-26 15:00",
        preexisting_task_ids=["task-old"],
        staging_dir=str(artifact_dir),
        status="waiting",
    )


class ExportFlowTests(unittest.TestCase):
    def test_真实流程保存批次号并只下载对应任务(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_flow_"))
        state_path = artifact_dir / "export-state.json"
        zip_bytes = make_export_zip(["甲公司"], artifact_dir)
        delivery_dir = artifact_dir / "甲公司_注协原始导出"
        client = FakeClient(
            json_results=[
                {
                    "status_code": 0,
                    "status_msg": "success",
                    "data": {
                        "batch_id": "batch-new",
                        "all_num": 1,
                        "hit_num": 1,
                        "not_hit_num": 0,
                    },
                },
                {
                    "status_code": 0,
                    "status_msg": "success",
                    "data": {"total": 1, "list": [{"name": "甲公司"}]},
                },
                {
                    "status_code": 0,
                    "status_msg": "success",
                    "data": [
                        {
                            "name": "工商",
                            "children": [
                                {"second_dimension_code": "basic"},
                                {"second_dimension_code": "customer"},
                            ],
                        }
                    ],
                },
                {
                    "status_code": 0,
                    "data": {
                        "list": [
                            {
                                "task_id": "task-old",
                                "type": "批量查询",
                                "status": 1,
                                "date": "2026-07-26 14:00",
                                "url": "/old.zip",
                            }
                        ]
                    },
                },
                {"status_code": 0, "status_msg": "success", "data": {}},
                {
                    "status_code": 0,
                    "data": {
                        "list": [
                            {
                                "task_id": "task-new",
                                "batch_id": "batch-new",
                                "type": "批量查询",
                                "name": "批量查询-甲公司",
                                "status": 1,
                                "date": "2026-07-26 15:01",
                                "url": "/download/task-new.zip",
                            },
                            {
                                "task_id": "task-old",
                                "type": "批量查询",
                                "status": 1,
                                "date": "2026-07-26 14:00",
                                "url": "/old.zip",
                            },
                        ]
                    },
                },
            ],
            binary_result=zip_bytes,
        )
        try:
            exporter = CicpaExporter(
                client,
                artifact_dir=delivery_dir,
                now=lambda: "2026-07-26 15:00",
            )
        except TypeError as exc:
            self.fail("导出器尚未支持精确成果目录：{}".format(exc))

        state = exporter.start_export(["甲公司"], state_path)
        task = exporter.wait_for_task(state, state_path, max_polls=1)
        extract_dir = exporter.download_and_validate(state, task, state_path)

        reloaded = load_export_state(state_path)
        self.assertEqual(Path(state.staging_dir), delivery_dir)
        self.assertTrue(state.direct_delivery)
        self.assertEqual(reloaded.batch_no, "batch-new")
        self.assertEqual(reloaded.task_id, "task-new")
        self.assertEqual(reloaded.status, "completed")
        self.assertEqual(extract_dir, delivery_dir)
        self.assertTrue((extract_dir / "基础工商信息.xlsx").exists())
        self.assertTrue((delivery_dir / "complete-dimensions.zip").exists())
        self.assertTrue((delivery_dir / "upload-companies.xlsx").exists())
        self.assertFalse((delivery_dir / "complete-dimensions_files").exists())
        self.assertEqual(
            client.binary_calls[0][1],
            "https://zsk-cmis.cicpa.org.cn/download/task-new.zip",
        )
        poll_calls = [
            kwargs
            for _method, url, kwargs in client.json_calls
            if url.endswith("/download/task_list")
        ]
        self.assertEqual(poll_calls[-1]["kind"], "poll")

    def test_超时后绝不下载旧的最新任务(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_timeout_"))
        state = waiting_state(artifact_dir)
        client = FakeClient(
            json_results=[
                {
                    "status_code": 0,
                    "data": {
                        "list": [
                            {
                                "task_id": "task-old",
                                "type": "批量查询",
                                "status": 1,
                                "date": "2026-07-26 14:00",
                                "url": "/old.zip",
                            }
                        ]
                    },
                }
            ]
        )
        exporter = CicpaExporter(client, staging_root=artifact_dir)

        with self.assertRaises(ExportTaskNotCorrelated):
            exporter.wait_for_task(state, artifact_dir / "state.json", max_polls=1)

        self.assertEqual(client.binary_calls, [])

    def test_多个无法唯一对应的新任务会立即阻断(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_ambiguous_"))
        state = waiting_state(artifact_dir)
        tasks = [
            {
                "task_id": "task-a",
                "type": "批量查询",
                "status": 1,
                "date": "2026-07-26 15:01",
                "url": "/a.zip",
            },
            {
                "task_id": "task-b",
                "type": "批量查询",
                "status": 1,
                "date": "2026-07-26 15:02",
                "url": "/b.zip",
            },
        ]
        exporter = CicpaExporter(FakeClient(), staging_root=artifact_dir)

        with self.assertRaises(ExportTaskNotCorrelated):
            exporter.correlate_task(state, tasks)


class ExportValidationTests(unittest.TestCase):
    def test_下载文件中的公司不一致时拒绝完成(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_mismatch_"))
        state_path = artifact_dir / "state.json"
        state = waiting_state(artifact_dir)
        task = {
            "task_id": "task-new",
            "batch_id": "batch-new",
            "type": "批量查询",
            "status": 1,
            "date": "2026-07-26 15:01",
            "url": "/wrong.zip",
        }
        client = FakeClient(binary_result=make_export_zip(["乙公司"], artifact_dir))
        exporter = CicpaExporter(client, staging_root=artifact_dir)

        with self.assertRaises(ExportValidationError):
            exporter.download_and_validate(state, task, state_path)

        self.assertNotEqual(state.status, "completed")

    def test_zip_路径穿越在写文件前被拒绝(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_zip_slip_"))
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as archive:
            archive.writestr("../outside.txt", "unsafe")
        state = waiting_state(artifact_dir)
        task = {
            "task_id": "task-new",
            "batch_id": "batch-new",
            "type": "批量查询",
            "status": 1,
            "date": "2026-07-26 15:01",
            "url": "/unsafe.zip",
        }
        exporter = CicpaExporter(
            FakeClient(binary_result=output.getvalue()),
            staging_root=artifact_dir,
        )

        with self.assertRaises(ExportValidationError):
            exporter.download_and_validate(state, task, artifact_dir / "state.json")

        self.assertFalse((artifact_dir.parent / "outside.txt").exists())

    def test_状态文件不带_bom_且不接受认证秘密字段(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_export_state_"))
        state_path = artifact_dir / "state.json"
        state = waiting_state(artifact_dir)

        save_export_state(state_path, state)
        raw = state_path.read_bytes()
        payload = json.loads(raw.decode("utf-8"))

        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        self.assertNotIn("cookie", json.dumps(payload, ensure_ascii=False).lower())
        self.assertNotIn("password", json.dumps(payload, ensure_ascii=False).lower())
        self.assertEqual(load_export_state(state_path), state)


if __name__ == "__main__":
    unittest.main()
