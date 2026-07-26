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
from datetime import date
from pathlib import Path
import tempfile
import unittest

import openpyxl

from scripts.related_party_check import (
    Company,
    build_company,
    rule3_counterparty_profile,
    run_check,
    write_report,
)


def blank_basic_row(name="甲公司"):
    row = [None] * 23
    row[0] = name
    row[2] = "存续"
    row[3] = "张三"
    row[4] = "1000 万元"
    row[5] = "2020-01-01"
    row[15] = 10
    row[21] = "成都市测试路 1 号"
    row[22] = "软件开发与技术服务"
    return tuple(row)


def make_basic_export(directory):
    path = Path(directory) / "基础工商信息.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["字段{}".format(index) for index in range(1, 24)])
    target = list(blank_basic_row("甲公司"))
    counterparty = list(blank_basic_row("乙公司"))
    counterparty[3] = "李四"
    counterparty[5] = "2025-01-01"
    counterparty[15] = 0
    sheet.append(target)
    sheet.append(counterparty)
    workbook.save(path)
    return path


def make_check_basic_export(directory, company_names):
    path = Path(directory) / "基础工商信息.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["字段{}".format(index) for index in range(1, 24)])
    for company_name in company_names:
        sheet.append(blank_basic_row(company_name))
    workbook.save(path)
    return path


def make_check_counterparty_export(
    path,
    owner_name,
    relation_label,
    counterparty_name,
):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append([
        "序号",
        "公司名称",
        "公告时间",
        "金额",
        "占比（%）",
        "与本公司关系",
        "货币代码",
        "关联方名称",
        "关联方ID",
        "关联理由",
    ])
    sheet.append([
        1,
        owner_name,
        "2026-07-01",
        100,
        10,
        relation_label,
        "CNY",
        counterparty_name,
        "org-1",
        "公开公告",
    ])
    workbook.save(path)
    return path


class RowBoundaryTests(unittest.TestCase):
    def test_最终受益人最短合法行不会越界(self):
        dimensions = {
            "ultimate_beneficiary": {
                "甲公司": [
                    ("序号", "甲公司", "字段3", "字段4", "字段5", "张三")
                ]
            }
        }

        company = build_company("甲公司", dimensions, blank_basic_row())

        self.assertEqual(company.beneficiaries, [("张三", "")])

    def test_变更记录缺少变更后字段时保留已取得内容(self):
        dimensions = {
            "change": {
                "甲公司": [
                    ("序号", "甲公司", "2026-01-01", "地址", "旧地址")
                ]
            }
        }

        company = build_company("甲公司", dimensions, blank_basic_row())

        self.assertEqual(
            company.changes,
            [("2026-01-01", "地址", "旧地址", "")],
        )


class CounterpartyProfileTests(unittest.TestCase):
    def test_成立时间按核查基准日动态计算二十四个月(self):
        target = Company(name="甲公司", business_scope="软件开发与技术服务")
        party = Company(
            name="乙公司",
            found_date="2025-01",
            insured=10,
            capital=10_000_000,
            business_scope="软件开发与技术服务",
        )

        hits = rule3_counterparty_profile(
            target,
            party,
            {"甲公司"},
            as_of_date=date(2026, 7, 1),
        )

        self.assertEqual(len(hits), 1)
        self.assertIn("成立不足24个月", hits[0][2])
        self.assertNotIn("2018", hits[0][2])

    def test_零参保使用专门红旗(self):
        target = Company(name="甲公司", business_scope="软件开发与技术服务")
        party = Company(
            name="乙公司",
            found_date="2010-01",
            insured=0,
            capital=10_000_000,
            business_scope="软件开发与技术服务",
        )

        hits = rule3_counterparty_profile(
            target,
            party,
            {"甲公司"},
            as_of_date=date(2026, 7, 1),
        )

        self.assertIn("参保人数0(空壳特征)", hits[0][2])
        self.assertNotIn("参保人数0(疑似空壳)", hits[0][2])


class DisclosedPartySemanticsTests(unittest.TestCase):
    def test_公开客户供应商不会生成关联风险或披露差异(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_public_counterparty_"))
        make_check_basic_export(artifact_dir, ["甲公司"])
        make_check_counterparty_export(
            artifact_dir / "客户.xlsx",
            owner_name="甲公司",
            relation_label="客户",
            counterparty_name="乙公司",
        )

        for label, disclosed_parties in (
            ("未提供自报名单", None),
            ("提供自报名单", {"丙公司"}),
        ):
            with self.subTest(label=label):
                output_path = artifact_dir / "{}_关联方核查报告.xlsx".format(label)
                result = run_check(
                    data_dir=artifact_dir,
                    target_names=["甲公司"],
                    output_path=output_path,
                    as_of_date=date(2026, 7, 1),
                    disclosed_parties=disclosed_parties,
                )

                self.assertEqual(result.hits, [])
                self.assertEqual(result.summary, [])
                workbook = openpyxl.load_workbook(output_path, read_only=True)
                try:
                    self.assertNotIn("04_注协标记关联", workbook.sheetnames)
                    self.assertNotIn("04b_用户披露差异", workbook.sheetnames)
                finally:
                    workbook.close()


class RunCheckIntegrationTests(unittest.TestCase):
    def test_汇总第三行是表头且第一条数据使用风险格式(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_report_style_"))
        output_path = artifact_dir / "甲公司_关联方核查报告.xlsx"
        summary = [{
            "company_a": "乙公司",
            "company_b": "甲公司",
            "relation_type": "审计对象-交易对手",
            "is_related": "是（建议确认）",
            "max_level": "🔴硬关联",
            "dimensions": "股权控制穿透",
            "hit_count": 1,
            "evidence": "甲公司投资乙公司",
            "suggestion": "核对股权资料",
        }]

        write_report(
            output_path,
            summary,
            [],
            {},
            {},
            "甲公司",
            {},
        )

        workbook = openpyxl.load_workbook(output_path)
        try:
            sheet = workbook["汇总判断"]
            self.assertEqual(sheet["A3"].value, "公司A")
            self.assertEqual(sheet["A3"].fill.fgColor.rgb, "00305496")
            self.assertEqual(sheet["A3"].font.color.rgb, "00FFFFFF")
            self.assertEqual(sheet["A4"].value, "乙公司")
            self.assertNotEqual(sheet["A4"].fill.fgColor.rgb, "00305496")
            self.assertEqual(sheet.freeze_panes, "A4")
        finally:
            workbook.close()

    def test_读取警告进入结果和数据质量工作表(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_check_report_"))
        data_dir = artifact_dir / "complete-dimensions_files"
        data_dir.mkdir()
        make_basic_export(data_dir)
        (data_dir / "主要人员（高管）.xlsx").write_bytes(b"not-an-xlsx")
        output_path = artifact_dir / "关联方核查报告.xlsx"

        result = run_check(
            data_dir=data_dir,
            target_names=["甲公司"],
            output_path=output_path,
            as_of_date=date(2026, 7, 1),
            task_mode="existing_export",
            scope_metadata={"初始穿透层数": 2, "候选上限": 100},
        )

        self.assertEqual(result.output_path, output_path)
        self.assertTrue(output_path.exists())
        self.assertTrue(
            any("主要人员（高管）.xlsx" in item["source"] for item in result.errors)
        )
        workbook = openpyxl.load_workbook(output_path, read_only=True)
        try:
            self.assertIn("数据质量与限制", workbook.sheetnames)
            self.assertIn("任务范围", workbook.sheetnames)
        finally:
            workbook.close()


if __name__ == "__main__":
    unittest.main()
