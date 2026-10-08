# 用业务反例复核识别逻辑，既检查漏报，也检查误报与证据等级。
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
import tempfile
import unittest

from scripts.related_party_check import (
    Company, HARD, MEDIUM, address_relationship, normalize_phones, parse_capital,
    parse_insured, parse_found_date, months_since_found, build_company, compare_pair,
    rule1_fingerprint, rule2_personnel, rule5_equity, run_check,
)
from scripts.discovery import discover
from scripts.cicpa.client import CompanyMatch
from scripts.tests.test_discovery import FakeDiscoveryClient, investment
from scripts.tests.test_related_party_check import blank_basic_row
from scripts.tests.test_报告追溯 import write_basic, write_dimension


class IdentificationLogicTests(unittest.TestCase):
    def test_候选发现不把未知姓名当共同人员(self):
        from scripts.cicpa.client import KeyPerson
        client = FakeDiscoveryClient(matches=[CompanyMatch("a", "甲公司")],
            equity={"a": [investment("b", "乙公司", 1)]},
            personnel={"a": [KeyPerson("未知")], "b": [KeyPerson("未知")]})
        self.assertEqual(discover("甲公司", client=client).candidates, {})

    def test_自然人长姓名保留且机构标志正确(self):
        natural = "阿卜杜拉阿卜杜拉"
        self.assertTrue(rule2_personnel(Company("甲", shareholders=[(natural, "30%", False)]), Company("乙", shareholders=[(natural, "30%", False)])))
        dim = {"shareholder_new": {"甲公司": [(1, "甲公司", "", "中金公司", "是", "30%")], "乙公司": [(2, "乙公司", "", "中金公司", "是", "30%")]} }
        a = build_company("甲公司", dim, blank_basic_row("甲公司"))
        b = build_company("乙公司", dim, blank_basic_row("乙公司"))
        self.assertFalse(any(h[0] == "person" and "中金公司" in h[2] for h in rule2_personnel(a, b)))

    def test_仅同名最终受益人不等同共同控制(self):
        hits = rule2_personnel(Company("甲", beneficiaries=[("张三", "5%")]), Company("乙", beneficiaries=[("张三", "5%")]))
        self.assertTrue(hits)
        self.assertTrue(all(h[1] == MEDIUM for h in hits))

    def test_同一母公司的两条投资记录能连接兄弟公司(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_siblings_"))
        write_basic(root, ["甲公司", "乙公司", "丙公司"], shared_phone=False)
        write_dimension(root, "对外投资（新）.xlsx", [[1, "乙公司", "甲公司", "", "80%"], [2, "乙公司", "丙公司", "", "90%"]])
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx")
        self.assertTrue(any(h["field"] == "equity_path" and {h["company_a"], h["company_b"]} == {"甲公司", "丙公司"} for h in result.hits))

    def test_多条质押线索不在第一条后停止(self):
        from scripts.related_party_check import rule7_guarantee
        a = Company("甲公司", pledges_out=[("张三", "乙公司"), ("李四", "乙公司")])
        self.assertEqual(len(rule7_guarantee(a, Company("乙公司"))), 2)

    def test_结构残缺的高管表不能显示资料完整(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_schema_"))
        write_basic(root, ["甲公司", "乙公司"], shared_phone=False)
        write_dimension(root, "主要人员（高管）.xlsx", [[1, "甲公司"], [2, "乙公司"]])
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx")
        self.assertTrue(any("字段不足" in str(item) for item in result.errors))

    def test_多层多数持股链与共同母公司保留逐段证据(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_chain_"))
        write_basic(root, ["甲公司", "乙公司", "丙公司"], shared_phone=False)
        write_dimension(root, "对外投资（新）.xlsx", [[1, "甲公司", "乙公司", "", "60%"], [2, "乙公司", "丙公司", "", "70%"]])
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx")
        chain = [hit for hit in result.hits if hit["field"] == "equity_path" and {hit["company_a"], hit["company_b"]} == {"甲公司", "丙公司"}]
        self.assertEqual(len(chain), 1)
        self.assertIn("乙公司", chain[0]["evidence"])
        self.assertEqual({s["cell"] for s in chain[0]["sources"] if s["file"].endswith("对外投资（新）.xlsx")}, {"C2", "E2", "C3", "E3"})

    def test_少数投资不能沿链推定控制(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_minority_"))
        write_basic(root, ["甲公司", "乙公司", "丙公司"], shared_phone=False)
        write_dimension(root, "对外投资（新）.xlsx", [[1, "甲公司", "乙公司", "", "20%"], [2, "乙公司", "丙公司", "", "60%"]])
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx")
        self.assertFalse(any(h["field"] == "equity_path" and {h["company_a"], h["company_b"]} == {"甲公司", "丙公司"} for h in result.hits))

    def test_未来成立日期列为数据异常而非经营风险(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_date_"))
        write_basic(root, ["甲公司", "乙公司"], shared_phone=False)
        import openpyxl
        workbook = openpyxl.load_workbook(root / "基础工商信息.xlsx")
        workbook.active["F3"] = "2030-01-01"
        workbook.save(root / "基础工商信息.xlsx"); workbook.close()
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx", as_of_date=date(2026, 10, 7))
        self.assertTrue(any("成立日期" in str(item) for item in result.errors))
        self.assertFalse(any("成立" in hit["evidence"] for hit in result.hits))

    def test_接口未知比例不转零且优先实际字段(self):
        from scripts.cicpa.client import CicpaClient
        from scripts.tests.test_client import FakeSession, FakeResponse, NoWaitLimiter
        for raw in ("", "待核实", "NaN", "-1", "150%"):
            self.assertIsNone(CicpaClient._ratio({"czbl": raw}, "czbl"))
        client = CicpaClient(session=FakeSession([FakeResponse(payload={"status_code": 0, "data": {"invests": {"children": [{"entName": "乙公司", "orgId": "b", "czbl": "60%", "investRatio": "1%"}]}}})]), limiter=NoWaitLimiter())
        self.assertEqual(client.get_equity_relations("a")[0].ratio, 60)

    def test_跨城市相同尾号座机不判相同(self):
        a = Company("甲", phones=normalize_phones("010-88881234"))
        b = Company("乙", phones=normalize_phones("028-88881234"))
        self.assertEqual(rule1_fingerprint(a, b), [])

    def test_带格式与国家码的同一完整号码仍命中(self):
        a = Company("甲", phones=normalize_phones("+86(028)88881234"))
        b = Company("乙", phones=normalize_phones("028-88881234"))
        self.assertTrue(any(hit[0] == "phone" for hit in rule1_fingerprint(a, b)))

    def test_缺区号号码及相邻号码只作为待核实线索(self):
        for first, second in (("88881234", "88881234"), ("02888881234", "02888881235")):
            with self.subTest(first=first):
                hits = rule1_fingerprint(Company("甲", phones={first}), Company("乙", phones={second}))
                self.assertTrue(hits)
                self.assertTrue(all(hit[1] == MEDIUM for hit in hits))

    def test_一方缺区号尾段一致只作为待核实线索(self):
        # 工商登记常见写法差异：一方带区号、一方只写本地号，不能整条漏掉
        hits = rule1_fingerprint(Company("甲", phones={"02888881234"}), Company("乙", phones={"88881234"}))
        self.assertTrue(hits, "一方缺区号时尾段一致应列待核实线索")
        self.assertTrue(all(hit[1] == MEDIUM for hit in hits))
        self.assertTrue(any("缺区号" in hit[2] for hit in hits))
        reverse = rule1_fingerprint(Company("甲", phones={"88881234"}), Company("乙", phones={"02888881234"}))
        self.assertTrue(reverse, "方向反过来同样应命中")

    def test_手机号尾八位不与本地号互判缺区号(self):
        # 手机号尾 8 位与本地座机同形，参与尾段比对会大量误报
        self.assertEqual(rule1_fingerprint(Company("甲", phones={"13888881234"}), Company("乙", phones={"88881234"})), [])

    def test_低比例持股单独出现降为轻微_伴随其他红旗保留中级(self):
        from scripts.related_party_check import LOW
        alone = compare_pair(Company("甲公司", investments=[("乙公司", "1%")]), Company("乙公司"), {}, set())
        invest_alone = [hit for hit in alone if hit["field"] == "invest"]
        self.assertTrue(invest_alone)
        self.assertTrue(all(hit["level"] == LOW for hit in invest_alone), "1% 持股单独出现应降为轻微")
        flagged = compare_pair(
            Company("甲公司", investments=[("乙公司", "1%")], main_persons=[("陈甲", "董事")]),
            Company("乙公司", main_persons=[("陈甲", "监事")]), {}, set())
        invest_flagged = [hit for hit in flagged if hit["field"] == "invest"]
        self.assertTrue(invest_flagged)
        self.assertTrue(all(hit["level"] == MEDIUM for hit in invest_flagged), "伴随人员重合红旗时应保留中级")

    def test_比例未知的持股不按低比例降级(self):
        hits = compare_pair(Company("甲公司", investments=[("乙公司", "待核实")]), Company("乙公司"), {}, set())
        invest = [hit for hit in hits if hit["field"] == "invest"]
        self.assertTrue(invest)
        self.assertTrue(all(hit["level"] == MEDIUM for hit in invest), "未知比例不当作低比例处理")

    def test_地址不能丢失城市或六位门牌(self):
        for first, second in (("成都市中山路100号A座101室", "北京市中山路100号A座102室"),
                              ("成都市大道123456号", "成都市大道654321号")):
            self.assertIsNone(address_relationship(first, second))
        self.assertIsNotNone(address_relationship("成都市中山路100号A座101室", "成都市中山路100号A座102室"))

    def test_缺失占位值不会成为共同人员或资源(self):
        dim = {"trademark": {name: [(1, name, "暂无数据")] for name in ("甲公司", "乙公司")}}
        rows = []
        for name in ("甲公司", "乙公司"):
            row = list(blank_basic_row(name)); row[3] = "未知"; row[21] = None
            rows.append(build_company(name, dim, row))
        hits = compare_pair(*rows, dim, set())
        self.assertFalse(any(hit["dimension"] in {"关键人员重合", "无形资产共用"} for hit in hits))

    def test_人员跨角色命中等级不受公司先后顺序影响(self):
        a = Company("甲", shareholders=[("陈甲", "10%", False)])
        b = Company("乙", main_persons=[("陈甲", "董事")])
        forward = [(h[0], h[1], h[2]) for h in rule2_personnel(a, b)]
        reverse = [(h[0], h[1], h[2]) for h in rule2_personnel(b, a)]
        self.assertEqual(forward, reverse)
        self.assertTrue(forward)

    def test_知名姓名不作为排除关联线索的白名单(self):
        a = Company("甲", main_persons=[("马云", "董事")])
        b = Company("乙", main_persons=[("马云", "董事")])
        self.assertTrue(rule2_personnel(a, b))

    def test_政府共同控制例外不扩大到国有母公司(self):
        public = "成都市人民政府国有资产监督管理委员会"
        self.assertEqual(rule2_personnel(Company("甲", actual_controllers=[public]), Company("乙", actual_controllers=[public])), [])
        parent = "模拟国有控股集团有限公司"
        self.assertTrue(rule2_personnel(Company("甲", actual_controllers=[parent]), Company("乙", actual_controllers=[parent])))

    def test_共同机构股东不能因名称长被丢掉(self):
        parent = "模拟共同母公司有限公司"
        a = Company("甲", shareholders=[(parent, "80%", True)])
        b = Company("乙", shareholders=[(parent, "90%", True)])
        self.assertTrue(rule5_equity(a, b))

    def test_直接股东记录也识别投资方向(self):
        a = Company("甲公司")
        b = Company("乙公司", shareholders=[("甲公司", "60%", True)])
        self.assertTrue(rule5_equity(a, b))

    def test_低比例投资不能直接按高风险控制线索处理(self):
        hits = rule5_equity(Company("甲公司", investments=[("乙公司", "1%")]), Company("乙公司"))
        self.assertTrue(hits)
        self.assertTrue(all(hit[1] == MEDIUM for hit in hits))

    def test_千位分隔和币种不误作人民币小资本(self):
        self.assertEqual(parse_capital("1,234.56万元人民币"), 12345600)
        self.assertIsNone(parse_capital("50万美元"))
        self.assertIsNone(parse_capital("待确认"))
        self.assertEqual(parse_capital(0), 0)
        self.assertEqual(parse_insured("1,234人"), 1234)
        self.assertIsNone(parse_insured("少于5人"))

    def test_成立日期保留日且周年边界不提前(self):
        self.assertEqual(parse_found_date("2025-10-20"), "2025-10-20")
        self.assertEqual(months_since_found("2025-10-20", date(2026, 10, 7)), 11)
        self.assertIsNone(parse_found_date("2025-02-30"))

    def test_唯一模糊搜索结果也不能替换被审计主体(self):
        result = discover("甲公司", client=FakeDiscoveryClient(matches=[CompanyMatch("x", "甲子公司")]))
        self.assertEqual(result.status, "needs_company_confirmation")
        self.assertEqual(result.candidates, {})

    def test_未知持股比例保留候选且明确标记待核实(self):
        client = FakeDiscoveryClient(matches=[CompanyMatch("a", "甲公司")],
            equity={"a": [investment("b", "乙公司", None)]})
        result = discover("甲公司", client=client)
        self.assertIn("乙公司", result.candidates)
        self.assertTrue(any("比例" in note for note in result.candidates["乙公司"].notes))

    def test_发票电话参与指纹且能定位原表(self):
        root = Path(tempfile.mkdtemp(prefix="rpi_logic_invoice_"))
        write_basic(root, ["甲公司", "乙公司"], shared_phone=False)
        write_dimension(root, "发票信息.xlsx", [["甲公司", "", "", "", "", "", "020-87651234", ""],
                                                 ["乙公司", "", "", "", "", "", "020-87651234", ""]])
        with redirect_stdout(StringIO()):
            result = run_check(data_dir=root, target_names=["甲公司"], output_path=root / "报告.xlsx")
        phones = [hit for hit in result.hits if hit["field"] == "phone"]
        self.assertTrue(phones)
        self.assertEqual({s["cell"] for s in phones[0]["sources"]}, {"G2", "G3"})

    def test_历史地址不会被当成当前地址高风险(self):
        a = Company("甲", addresses=["北京市现在大街200号"], changes=[("2020", "地址", "成都市老地址路100号", "北京市现在大街200号")])
        b = Company("乙", addresses=["成都市老地址路100号"])
        hits = compare_pair(a, b, {}, set())
        self.assertTrue(any(hit["dimension"] == "历史关联痕迹" for hit in hits))
        self.assertFalse(any(hit["field"] == "address" and hit["level"] == HARD for hit in hits))


if __name__ == "__main__":
    unittest.main()
