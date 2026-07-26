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

from scripts.cicpa.client import (
    AuthenticationRequired,
    CicpaClient,
    CicpaError,
    CompanyDetail,
    CompanyMatch,
    EquityEdge,
    KeyPerson,
    PermissionDenied,
    RateLimitExceeded,
    RateLimiter,
    RatePolicy,
    Shareholder,
)


class FakeTime:
    def __init__(self):
        self.now = 0.0
        self.sleeps = []

    def clock(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


class FakeResponse:
    def __init__(self, status_code=200, payload=None, headers=None, text="", content=b""):
        self.status_code = status_code
        self._payload = payload if payload is not None else {"status_code": 0}
        self.headers = headers or {}
        self.text = text
        self.content = content

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, BaseException):
            raise outcome
        return outcome


class NoWaitLimiter:
    def wait(self, kind="normal"):
        return 0.0


class RateLimiterTests(unittest.TestCase):
    def test_普通请求与轮询使用各自随机等待区间(self):
        fake_time = FakeTime()
        values = iter([3.5, 12.0])
        limiter = RateLimiter(
            policy=RatePolicy(),
            clock=fake_time.clock,
            sleep=fake_time.sleep,
            random_uniform=lambda _start, _end: next(values),
        )

        self.assertEqual(limiter.wait("normal"), 3.5)
        self.assertEqual(limiter.wait("poll"), 12.0)
        self.assertEqual(fake_time.sleeps, [3.5, 12.0])

    def test_一分钟第十六次请求必须等待窗口释放(self):
        fake_time = FakeTime()
        limiter = RateLimiter(
            policy=RatePolicy(normal_delay=(0.0, 0.0), poll_delay=(0.0, 0.0)),
            clock=fake_time.clock,
            sleep=fake_time.sleep,
            random_uniform=lambda _start, _end: 0.0,
        )

        for _ in range(15):
            limiter.wait()
        waited = limiter.wait()

        self.assertGreaterEqual(waited, 60.0)
        self.assertGreaterEqual(fake_time.now, 60.0)

    def test_一小时第三百零一次请求必须等待窗口释放(self):
        fake_time = FakeTime()
        limiter = RateLimiter(
            policy=RatePolicy(
                max_per_minute=1000,
                normal_delay=(0.0, 0.0),
                poll_delay=(0.0, 0.0),
            ),
            clock=fake_time.clock,
            sleep=fake_time.sleep,
            random_uniform=lambda _start, _end: 0.0,
        )

        for _ in range(300):
            limiter.wait()
        waited = limiter.wait()

        self.assertGreaterEqual(waited, 3600.0)
        self.assertGreaterEqual(fake_time.now, 3600.0)


class CicpaClientTests(unittest.TestCase):
    def test_相同缓存键只访问一次外部接口(self):
        session = FakeSession([FakeResponse(payload={"status_code": 0, "data": {"name": "甲公司"}})])
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=lambda _seconds: None)

        first = client.request_json("GET", "/company/detail", cache_key=("detail", "甲公司"))
        second = client.request_json("GET", "/company/detail", cache_key=("detail", "甲公司"))

        self.assertEqual(first, second)
        self.assertEqual(len(session.calls), 1)

    def test_认证与权限错误不重复请求(self):
        for status_code, error_type in (
            (401, AuthenticationRequired),
            (403, PermissionDenied),
        ):
            with self.subTest(status_code=status_code):
                session = FakeSession([FakeResponse(status_code=status_code)])
                client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=lambda _seconds: None)

                with self.assertRaises(error_type):
                    client.request_json("GET", "/protected")

                self.assertEqual(len(session.calls), 1)

    def test_优先遵守_retry_after(self):
        fake_time = FakeTime()
        session = FakeSession(
            [
                FakeResponse(status_code=429, headers={"Retry-After": "7"}),
                FakeResponse(payload={"status_code": 0, "data": {"ok": True}}),
            ]
        )
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=fake_time.sleep)

        result = client.request_json("GET", "/limited")

        self.assertEqual(result["data"], {"ok": True})
        self.assertEqual(fake_time.sleeps, [7.0])
        self.assertEqual(len(session.calls), 2)

    def test_三次退避后停止继续撞限流(self):
        fake_time = FakeTime()
        session = FakeSession([FakeResponse(status_code=429) for _ in range(4)])
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=fake_time.sleep)

        with self.assertRaises(RateLimitExceeded):
            client.request_json("GET", "/limited")

        self.assertEqual(fake_time.sleeps, [30.0, 60.0, 120.0])
        self.assertEqual(len(session.calls), 4)

    def test_网络与服务端错误只做有限重试(self):
        fake_time = FakeTime()
        session = FakeSession(
            [
                OSError("network down"),
                FakeResponse(status_code=503),
                FakeResponse(payload={"status_code": 0, "data": {"ok": True}}),
            ]
        )
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=fake_time.sleep)

        result = client.request_json("GET", "/temporary")

        self.assertEqual(result["data"], {"ok": True})
        self.assertEqual(len(session.calls), 3)
        self.assertEqual(fake_time.sleeps, [1.0, 2.0])

    def test_二进制下载复用同一认证与重试入口(self):
        session = FakeSession(
            [
                FakeResponse(status_code=503),
                FakeResponse(content=b"zip-bytes"),
            ]
        )
        fake_time = FakeTime()
        client = CicpaClient(session=session, limiter=NoWaitLimiter(), sleep=fake_time.sleep)

        content = client.request_bytes("GET", "/download/task.zip")

        self.assertEqual(content, b"zip-bytes")
        self.assertEqual(fake_time.sleeps, [1.0])
        self.assertEqual(len(session.calls), 2)


class CompanyApiTests(unittest.TestCase):
    def test_搜索保留全部候选并把精确名称排在前面(self):
        response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "total": 2,
                    "list": [
                        {
                            "org_id": "other-id",
                            "name": "甲公司分公司",
                            "credit_code": "91510000OTHER00001",
                            "legal_person": "李四",
                            "status": "存续",
                        },
                        {
                            "org_id": "exact-id",
                            "name": "甲公司",
                            "credit_code": "91510000TEST000001",
                            "legal_person": "张三",
                            "status": "存续",
                        },
                    ],
                },
            }
        )
        client = CicpaClient(
            session=FakeSession([response]),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        matches = client.search_companies("甲公司")

        self.assertEqual(
            matches,
            [
                CompanyMatch(
                    org_id="exact-id",
                    name="甲公司",
                    credit_code="91510000TEST000001",
                    legal_person="张三",
                    status="存续",
                ),
                CompanyMatch(
                    org_id="other-id",
                    name="甲公司分公司",
                    credit_code="91510000OTHER00001",
                    legal_person="李四",
                    status="存续",
                ),
            ],
        )

    def test_企业详情映射稳定字段(self):
        response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "org_id": "org-1",
                    "entName": "甲公司",
                    "uniscid": "91510000TEST000001",
                    "legalPerson": "张三",
                    "entStatus": "存续",
                    "address": "成都市测试路 1 号",
                },
            }
        )
        client = CicpaClient(
            session=FakeSession([response]),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        detail = client.get_company_detail("org-1")

        self.assertEqual(
            detail,
            CompanyDetail(
                org_id="org-1",
                name="甲公司",
                credit_code="91510000TEST000001",
                legal_person="张三",
                status="存续",
                address="成都市测试路 1 号",
            ),
        )

    def test_股东和主要人员解析真实列表容器(self):
        shareholders_response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "total": 1,
                    "list": [
                        {
                            "shareholderName": "乙公司",
                            "subConam": "6000 万元",
                            "fundedRatio": "60%",
                            "orgId": "org-2",
                        }
                    ],
                },
            }
        )
        personnel_response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "total": 1,
                    "list": [{"name": "王五", "position": "董事长"}],
                },
            }
        )
        client = CicpaClient(
            session=FakeSession([shareholders_response, personnel_response]),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        self.assertEqual(
            client.get_shareholders("org-1"),
            [Shareholder(name="乙公司", ratio=60.0, org_id="org-2")],
        )
        self.assertEqual(
            client.get_key_personnel("org-1"),
            [KeyPerson(name="王五", role="董事长")],
        )

    def test_股权结构分别标记投资和股东方向(self):
        response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "invests": {
                        "children": [
                            {
                                "entName": "丙公司",
                                "investRatio": "55%",
                                "orgId": "org-3",
                            }
                        ]
                    },
                    "holders": {
                        "children": [
                            {
                                "entName": "乙公司",
                                "investRatio": "60%",
                                "orgId": "org-2",
                            }
                        ]
                    },
                },
            }
        )
        client = CicpaClient(
            session=FakeSession([response]),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        relations = client.get_equity_relations("org-1")

        self.assertEqual(
            relations,
            [
                EquityEdge(
                    company_id="org-3",
                    company_name="丙公司",
                    ratio=55.0,
                    direction="investment",
                ),
                EquityEdge(
                    company_id="org-2",
                    company_name="乙公司",
                    ratio=60.0,
                    direction="holder",
                ),
            ],
        )

    def test_子公司发现只返回对外投资方向(self):
        response = FakeResponse(
            payload={
                "status_code": 0,
                "status_msg": "success",
                "data": {
                    "invests": {
                        "children": [
                            {"name": "丙公司", "ratio": "55", "orgid": "org-3"}
                        ]
                    },
                    "holders": {
                        "children": [
                            {"name": "乙公司", "ratio": "60", "orgid": "org-2"}
                        ]
                    },
                },
            }
        )
        client = CicpaClient(
            session=FakeSession([response]),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        subsidiaries = client.discover_subsidiaries("org-1")

        self.assertEqual(
            subsidiaries,
            [CompanyMatch(org_id="org-3", name="丙公司")],
        )

    def test_业务失败不会伪装成空结果(self):
        client = CicpaClient(
            session=FakeSession(
                [
                    FakeResponse(
                        payload={
                            "status_code": 1001,
                            "status_msg": "查询条件无效",
                            "data": {},
                        }
                    )
                ]
            ),
            limiter=NoWaitLimiter(),
            sleep=lambda _seconds: None,
        )

        with self.assertRaises(CicpaError):
            client.search_companies("甲公司")


if __name__ == "__main__":
    unittest.main()
