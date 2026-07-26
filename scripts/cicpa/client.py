# 关联方识别与核查 — 注协查询客户端
# 本文件改编自 nigo/nigo-skills/cicpa-company-query(MIT)
# 上游作者: nigo(涂佳兵) | 原始仓库: https://github.com/nigo81/nigo-skills
# 上游协议: MIT
#
# 本地修改: CPA-Q(quanfanpro-code)
# 本文件的修改部分同样以 MIT 协议发布,与上游保持一致。
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
#"""中注协接口的统一请求、限速和缓存入口。"""

from collections import deque
from dataclasses import dataclass
import random
import time
from typing import Any, Callable, Deque, Dict, Hashable, List, Optional, Tuple

try:
    import requests
except ImportError:  # 预检会在真正联网前给出安装引导
    requests = None


class CicpaError(RuntimeError):
    """中注协访问的基础错误。"""


class AuthenticationRequired(CicpaError):
    """登录态失效，需要重新登录。"""


class PermissionDenied(CicpaError):
    """当前账号没有所需权限。"""


class RateLimitExceeded(CicpaError):
    """多次触发限流后停止访问。"""


class TemporaryApiError(CicpaError):
    """网络或服务端临时错误达到重试上限。"""


@dataclass(frozen=True)
class CompanyMatch:
    org_id: str
    name: str
    credit_code: str = ""
    legal_person: str = ""
    status: str = ""


@dataclass(frozen=True)
class CompanyDetail:
    org_id: str
    name: str = ""
    credit_code: str = ""
    legal_person: str = ""
    status: str = ""
    address: str = ""


@dataclass(frozen=True)
class Shareholder:
    name: str
    ratio: float = 0.0
    org_id: str = ""


@dataclass(frozen=True)
class KeyPerson:
    name: str
    role: str = ""


@dataclass(frozen=True)
class EquityEdge:
    company_id: str
    company_name: str
    ratio: float
    direction: str


@dataclass(frozen=True)
class RatePolicy:
    """本 skill 的保守访问频率自我约束。"""

    max_per_minute: int = 15
    max_per_hour: int = 300
    normal_delay: Tuple[float, float] = (2.0, 5.0)
    poll_delay: Tuple[float, float] = (10.0, 20.0)

    def __post_init__(self):
        if self.max_per_minute <= 0 or self.max_per_hour <= 0:
            raise ValueError("请求上限必须大于 0")
        for start, end in (self.normal_delay, self.poll_delay):
            if start < 0 or end < start:
                raise ValueError("随机等待区间无效")


class RateLimiter:
    """串行滚动窗口限速器。"""

    def __init__(
        self,
        policy: RatePolicy,
        clock: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
        random_uniform: Callable[[float, float], float] = random.uniform,
    ):
        self.policy = policy
        self._clock = clock
        self._sleep = sleep
        self._random_uniform = random_uniform
        self._minute: Deque[float] = deque()
        self._hour: Deque[float] = deque()

    @staticmethod
    def _purge(queue: Deque[float], now: float, window: float) -> None:
        while queue and now - queue[0] >= window:
            queue.popleft()

    def wait(self, kind: str = "normal") -> float:
        """等待配额和随机间隔，返回本次总等待秒数。"""
        if kind not in {"normal", "poll"}:
            raise ValueError("请求类型只能是 normal 或 poll")

        total_wait = 0.0
        while True:
            now = self._clock()
            self._purge(self._minute, now, 60.0)
            self._purge(self._hour, now, 3600.0)
            waits = []
            if len(self._minute) >= self.policy.max_per_minute:
                waits.append(60.0 - (now - self._minute[0]))
            if len(self._hour) >= self.policy.max_per_hour:
                waits.append(3600.0 - (now - self._hour[0]))
            quota_wait = max(waits, default=0.0)
            if quota_wait <= 0:
                break
            self._sleep(quota_wait)
            total_wait += quota_wait

        delay_range = self.policy.normal_delay if kind == "normal" else self.policy.poll_delay
        jitter = float(self._random_uniform(*delay_range))
        if jitter > 0:
            self._sleep(jitter)
            total_wait += jitter

        timestamp = self._clock()
        self._minute.append(timestamp)
        self._hour.append(timestamp)
        return total_wait


class CicpaClient:
    """所有中注协 HTTP 请求必须经过此类。"""

    ZSK_BASE = "https://zsk-cmis.cicpa.org.cn"
    RATE_BACKOFFS = (30.0, 60.0, 120.0)
    TEMP_BACKOFFS = (1.0, 2.0)

    def __init__(
        self,
        session=None,
        limiter: Optional[RateLimiter] = None,
        sleep: Callable[[float], None] = time.sleep,
    ):
        if session is None:
            if requests is None:
                raise CicpaError("缺少 requests，需先完成依赖预检")
            session = requests.Session()
        self.session = session
        self.limiter = limiter or RateLimiter(RatePolicy())
        self._sleep = sleep
        self._cache: Dict[Hashable, Any] = {}

    @staticmethod
    def _retry_after_seconds(response) -> Optional[float]:
        value = response.headers.get("Retry-After")
        if value is None:
            return None
        try:
            seconds = float(value)
        except (TypeError, ValueError):
            return None
        return max(0.0, seconds)

    def _request_response(
        self,
        method: str,
        url: str,
        *,
        kind: str = "normal",
        **kwargs,
    ):
        """发送一次可重试请求，JSON 与二进制下载共用。"""
        kwargs.setdefault("timeout", 30)
        rate_retries = 0
        temporary_retries = 0

        while True:
            self.limiter.wait(kind)
            try:
                response = self.session.request(method, url, **kwargs)
            except Exception as exc:
                request_error = requests is not None and isinstance(
                    exc, requests.exceptions.RequestException
                )
                if not isinstance(exc, OSError) and not request_error:
                    raise
                if temporary_retries >= len(self.TEMP_BACKOFFS):
                    raise TemporaryApiError("网络错误达到重试上限") from exc
                delay = self.TEMP_BACKOFFS[temporary_retries]
                temporary_retries += 1
                self._sleep(delay)
                continue

            if response.status_code == 401:
                raise AuthenticationRequired("登录态已失效")
            if response.status_code == 403:
                raise PermissionDenied("当前账号无权访问该功能")
            if response.status_code == 429:
                if rate_retries >= len(self.RATE_BACKOFFS):
                    raise RateLimitExceeded("三次退避后仍被限流，已停止")
                delay = self._retry_after_seconds(response)
                if delay is None:
                    delay = self.RATE_BACKOFFS[rate_retries]
                rate_retries += 1
                self._sleep(delay)
                continue
            if 500 <= response.status_code < 600:
                if temporary_retries >= len(self.TEMP_BACKOFFS):
                    raise TemporaryApiError(
                        "服务端临时错误达到重试上限：HTTP {}".format(response.status_code)
                    )
                delay = self.TEMP_BACKOFFS[temporary_retries]
                temporary_retries += 1
                self._sleep(delay)
                continue
            if response.status_code >= 400:
                raise CicpaError("接口请求失败：HTTP {}".format(response.status_code))
            return response

    def request_json(
        self,
        method: str,
        url: str,
        *,
        cache_key: Optional[Hashable] = None,
        kind: str = "normal",
        **kwargs,
    ) -> Any:
        """发送 JSON 请求并应用本次任务缓存。"""
        if cache_key is not None and cache_key in self._cache:
            return self._cache[cache_key]

        response = self._request_response(method, url, kind=kind, **kwargs)

        try:
            payload = response.json()
        except (TypeError, ValueError) as exc:
            raise CicpaError("接口返回的不是有效 JSON") from exc

        if cache_key is not None:
            self._cache[cache_key] = payload
        return payload

    def request_bytes(
        self,
        method: str,
        url: str,
        *,
        kind: str = "normal",
        **kwargs,
    ) -> bytes:
        """通过同一认证与重试入口下载二进制内容。"""
        response = self._request_response(method, url, kind=kind, **kwargs)
        return bytes(response.content)

    @staticmethod
    def _text(item: Dict[str, Any], *keys: str) -> str:
        for key in keys:
            value = item.get(key)
            if value is not None and str(value).strip():
                return str(value).strip()
        return ""

    @classmethod
    def _ratio(cls, item: Dict[str, Any], *keys: str) -> float:
        raw = cls._text(item, *keys).replace("%", "").replace(",", "")
        try:
            return float(raw)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _business_data(payload: Any, alias: str) -> Any:
        if not isinstance(payload, dict):
            raise CicpaError("{}返回结构无效".format(alias))
        if payload.get("status_code") != 0:
            message = str(payload.get("status_msg") or "未知业务错误")
            raise CicpaError("{}失败：{}".format(alias, message))
        return payload.get("data", {})

    @staticmethod
    def _items(data: Any, *keys: str) -> List[Dict[str, Any]]:
        if isinstance(data, list):
            return [item for item in data if isinstance(item, dict)]
        if not isinstance(data, dict):
            return []
        for key in keys:
            value = data.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
        return []

    @classmethod
    def _company_match(cls, item: Dict[str, Any]) -> CompanyMatch:
        return CompanyMatch(
            org_id=cls._text(item, "org_id", "orgId", "orgid", "id", "entId"),
            name=cls._text(item, "name", "entName", "company_name", "enterpriseName"),
            credit_code=cls._text(item, "credit_code", "creditCode", "uniscid", "uniCode"),
            legal_person=cls._text(
                item,
                "legal_person",
                "legalPerson",
                "legalRepresentative",
                "frName",
            ),
            status=cls._text(item, "status", "entStatus", "regStatus"),
        )

    def search_companies(self, name: str, page_size: int = 10) -> List[CompanyMatch]:
        query = str(name).strip()
        if not query:
            raise ValueError("企业名称不能为空")
        payload = self.request_json(
            "POST",
            self.ZSK_BASE + "/open/industry_chain_api/v1/search/home_search",
            cache_key=("search", query, page_size),
            json={
                "condition": [],
                "query": query,
                "from": 0,
                "page": 1,
                "page_size": page_size,
                "size": page_size,
                "type": "company",
                "ranges": [],
            },
        )
        data = self._business_data(payload, "企业搜索")
        items = self._items(data, "list", "searchResultList", "items")
        matches = [self._company_match(item) for item in items]
        matches = [item for item in matches if item.name]
        return sorted(matches, key=lambda item: 0 if item.name == query else 1)

    def get_company_detail(self, company_id: str) -> CompanyDetail:
        org_id = str(company_id).strip()
        if not org_id:
            raise ValueError("企业标识不能为空")
        payload = self.request_json(
            "GET",
            self.ZSK_BASE + "/open/enterprise_info_api/v3/find_company_basic_info",
            cache_key=("detail", org_id),
            params={"orgid": org_id},
        )
        data = self._business_data(payload, "企业详情")
        if not isinstance(data, dict):
            data = {}
        return CompanyDetail(
            org_id=self._text(data, "org_id", "orgId", "orgid") or org_id,
            name=self._text(data, "name", "entName", "company_name", "enterpriseName"),
            credit_code=self._text(data, "credit_code", "creditCode", "uniscid", "uniCode"),
            legal_person=self._text(
                data,
                "legal_person",
                "legalPerson",
                "legalRepresentative",
                "frName",
            ),
            status=self._text(data, "status", "entStatus", "regStatus"),
            address=self._text(data, "address", "regAddress", "dom"),
        )

    def get_shareholders(self, company_id: str) -> List[Shareholder]:
        org_id = str(company_id).strip()
        payload = self.request_json(
            "GET",
            self.ZSK_BASE + "/open/enterprise_info_api/v4/stock_holder_newest",
            cache_key=("shareholders", org_id),
            params={"orgid": org_id, "page": 1, "pagesize": 50, "pageSize": 50},
        )
        data = self._business_data(payload, "股东查询")
        items = self._items(data, "list", "stockHolderList", "items")
        result = []
        for item in items:
            name = self._text(item, "shareholderName", "holderName", "name", "entName")
            if name:
                result.append(
                    Shareholder(
                        name=name,
                        ratio=self._ratio(
                            item,
                            "fundedRatio",
                            "stockPercent",
                            "ratio",
                            "investRatio",
                        ),
                        org_id=self._text(item, "orgId", "org_id", "orgid", "entId"),
                    )
                )
        return result

    def get_key_personnel(self, company_id: str) -> List[KeyPerson]:
        org_id = str(company_id).strip()
        primary_url = self.ZSK_BASE + "/open/enterprise_info_api/v3/main_person_list"
        payload = self.request_json(
            "GET",
            primary_url,
            cache_key=("key_personnel", org_id, "primary"),
            params={"orgid": org_id, "page": 1, "pageSize": 50},
        )
        try:
            data = self._business_data(payload, "主要人员查询")
        except CicpaError:
            payload = self.request_json(
                "GET",
                self.ZSK_BASE + "/open/enterprise_info_api/v1/main_person_tab",
                cache_key=("key_personnel", org_id, "fallback"),
                params={"orgid": org_id},
            )
            data = self._business_data(payload, "主要人员查询")
        items = self._items(data, "list", "mainPersonList", "items")
        result = []
        for item in items:
            name = self._text(item, "name", "personName", "ryName")
            if name:
                result.append(
                    KeyPerson(
                        name=name,
                        role=self._text(item, "position", "job", "positionName"),
                    )
                )
        return result

    def get_equity_relations(self, company_id: str) -> List[EquityEdge]:
        org_id = str(company_id).strip()
        payload = self.request_json(
            "GET",
            self.ZSK_BASE + "/open/enterprise_info_api/v3/atlas/enterprise_equity",
            cache_key=("equity", org_id),
            params={"orgid": org_id},
        )
        data = self._business_data(payload, "股权关系查询")
        if not isinstance(data, dict):
            return []
        result = []
        for container_name, direction in (("invests", "investment"), ("holders", "holder")):
            container = data.get(container_name, {})
            items = self._items(container, "children", "list", "items")
            for item in items:
                company = self._company_match(item)
                if company.name:
                    result.append(
                        EquityEdge(
                            company_id=company.org_id,
                            company_name=company.name,
                            ratio=self._ratio(item, "investRatio", "ratio", "czbl"),
                            direction=direction,
                        )
                    )
        return result

    def discover_subsidiaries(self, company_id: str) -> List[CompanyMatch]:
        return [
            CompanyMatch(org_id=edge.company_id, name=edge.company_name)
            for edge in self.get_equity_relations(company_id)
            if edge.direction == "investment"
        ]

    def clear_cache(self) -> None:
        """清空本次运行缓存。"""
        self._cache.clear()
