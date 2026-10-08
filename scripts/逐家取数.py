# -*- coding: utf-8 -*-
# 关联方识别与核查 — 逐家取数（保底通道）
#
# 背景：取得 20 个关联方核查维度的常规做法是"上传企业名单→拿批次号→批量导出"。
# 当这条批量通道不可用时（例如注协的批量上传接口持续返回业务错误），任务不能就此
# 停摆，改为逐个企业调用单企业接口取数，并把结果写成与批量导出同名的维度文件，
# 使下游的核查规则与七表报告仍能照常运行。
#
# 逐家取数能覆盖的维度：基础工商信息、最新公示股东、实际控制人、最终受益人、对外投资（新）。
# 取不到的维度（客户、供应商、发票信息、变更记录、主要人员等）如实记为数据缺口，
# 在取数说明中标注，绝不假装已取得。

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from scripts.state_io import atomic_write_json


SEARCH_PATH = "/open/industry_chain_api/v1/search/home_search"
DETAIL_PATH = "/open/enterprise_info_api/v3/find_company_basic_info"
HOLDER_PATH = "/open/enterprise_info_api/v4/stock_holder_newest"
EQUITY_PATH = "/open/enterprise_info_api/v3/atlas/enterprise_equity"
HOST = "https://zsk-cmis.cicpa.org.cn"

# 逐家取数能覆盖的维度：维度编码与文件名（文件名与批量导出保持一致）
SINGLE_COMPANY_DIMENSIONS = (
    ("S0000002", "基础工商信息.xlsx"),
    ("S0000103", "最新公示股东.xlsx"),
    ("S0000032", "实际控制人.xlsx"),
    ("S0000020", "最终受益人.xlsx"),
    ("S0000104", "对外投资（新）.xlsx"),
)

# 20 个维度中，逐家取数取不到的（如实记为缺口）
UNAVAILABLE_DIMENSIONS = (
    "股东信息", "主要人员（高管）", "核心团队", "参控股企业", "发票信息",
    "客户", "供应商", "变更记录", "法定代表人变更", "经营异常",
    "股权质押", "动产抵押", "商标", "软件著作权", "微信公众号",
)

# 各维度表头：列位置必须与核查规则读取的位置一致，不可随意调整
BASIC_HEADERS = (
    "公司名称", "企业标识", "经营状态", "法定代表人", "注册资本", "成立日期",
    "省份", "城市", "区县", "联系电话", "企业官网", "电子邮箱", "统一社会信用代码",
    "工商注册号", "组织机构代码", "参保人数", "企业类型", "一级行业", "二级行业",
    "三级行业", "曾用名", "注册地址", "经营范围",
)
HOLDER_HEADERS = (
    "序号", "公司名称", "股东类型", "股东名称", "是否机构", "持股比例",
    "认缴出资额", "股东标识",
)
CONTROLLER_HEADERS = ("序号", "公司名称", "实际控制人", "取得方式", "来源")
BENEFICIARY_HEADERS = (
    "序号", "公司名称", "受益人类型", "层级", "持股路径", "受益人名称", "受益比例",
)
INVEST_HEADERS = ("序号", "公司名称", "被投资企业名称", "被投资企业标识", "持股比例")


class SingleCompanyError(RuntimeError):
    """逐家取数无法继续。"""


@dataclass
class CompanyProfile:
    """单个企业逐家取到的数据。"""

    name: str = ""
    company_id: str = ""
    basic: List[Any] = field(default_factory=list)
    holders: List[List[Any]] = field(default_factory=list)
    controllers: List[str] = field(default_factory=list)
    beneficiaries: List[Tuple[str, str]] = field(default_factory=list)
    investments: List[Tuple[str, str, str]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


def _text(item: Any, *keys: str) -> str:
    """按候选键名取第一个非空文本值。"""
    if not isinstance(item, dict):
        return ""
    for key in keys:
        value = item.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def _clean_ratio(value: Any) -> str:
    """把持股比例统一成带百分号的文本。"""
    raw = str(value or "").strip().replace(",", "")
    if not raw:
        return ""
    if raw.endswith("%"):
        return raw
    return raw + "%"


def _equity_children(container: Any) -> List[Dict[str, Any]]:
    if isinstance(container, list):
        return [item for item in container if isinstance(item, dict)]
    if not isinstance(container, dict):
        return []
    for key in ("children", "list", "items"):
        value = container.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def _business_data(payload: Any, alias: str) -> Any:
    if not isinstance(payload, dict):
        raise SingleCompanyError("{}返回结构无效".format(alias))
    if payload.get("status_code") != 0:
        message = str(payload.get("status_msg") or payload.get("msg") or "未知业务错误")
        raise SingleCompanyError("{}失败：{}".format(alias, message))
    return payload.get("data", {})


def resolve_company_id(client, name: str, warnings: Optional[List[str]] = None) -> str:
    """按工商全称匹配企业标识，匹配不到就明确报错；单条模糊匹配照常采用但留痕。"""
    matches = client.search_companies(name)
    exact = [match for match in matches if match.name == name]
    if len(exact) == 1:
        return exact[0].org_id
    if len(matches) == 1:
        if warnings is not None:
            warnings.append("模糊匹配：查询 {} 采用 {}，主体待核实".format(name, matches[0].name))
        return matches[0].org_id
    raise SingleCompanyError("企业名称无法唯一匹配，需要确认具体主体：" + str(name))


def fetch_company(client, name: str) -> CompanyProfile:
    """逐家取一个企业的可取得数据。"""
    profile = CompanyProfile(name=str(name).strip())
    if not profile.name:
        raise SingleCompanyError("企业名称不能为空")
    profile.company_id = resolve_company_id(client, profile.name, warnings=profile.warnings)

    detail = _business_data(
        client.request_json(
            "GET",
            HOST + DETAIL_PATH,
            params={"orgid": profile.company_id},
            cache_key=("single-detail", profile.company_id),
        ),
        "企业详情",
    )
    if not isinstance(detail, dict):
        detail = {}

    holders_raw = _business_data(
        client.request_json(
            "GET",
            HOST + HOLDER_PATH,
            params={"orgid": profile.company_id, "page": 1, "pagesize": 50, "pageSize": 50},
            cache_key=("single-holder", profile.company_id),
        ),
        "股东查询",
    )
    holder_items = holder_rows = []
    if isinstance(holders_raw, dict):
        holder_rows = [item for item in holders_raw.get("list", []) if isinstance(item, dict)]
    holder_items = holder_rows

    equity = _business_data(
        client.request_json(
            "GET",
            HOST + EQUITY_PATH,
            params={"orgid": profile.company_id},
            cache_key=("single-equity", profile.company_id),
        ),
        "股权关系查询",
    )
    if not isinstance(equity, dict):
        equity = {}

    # 基础工商信息
    profile.basic = [
        _text(detail, "name") or profile.name,
        _text(detail, "org_encode") or profile.company_id,
        _text(detail, "operating_status"),
        _text(detail, "legal_representative"),
        _text(detail, "reg_capital"),
        _text(detail, "established_date"),
        _text(detail, "province"),
        _text(detail, "city"),
        _text(detail, "county", "district"),
        _text(detail, "telphone", "telephone"),
        _text(detail, "web_site", "website"),
        _text(detail, "email"),
        _text(detail, "unified_social_credit_code"),
        _text(detail, "reg_num"),
        _text(detail, "org_code"),
        _text(detail, "into_insurance_num"),
        _text(detail, "corp_type"),
        _text(detail, "one_industry"),
        _text(detail, "second_level_industry"),
        _text(detail, "third_level_industry"),
        _text(detail, "used_name"),
        _text(detail, "corp_address"),
        _text(detail, "operating_scope"),
    ]

    # 最新公示股东
    for index, item in enumerate(holder_items, 1):
        holder_name = _text(item, "holder_name", "shareholderName", "name")
        if not holder_name:
            continue
        is_org = _text(item, "is_holder_org")
        profile.holders.append([
            index,
            profile.name,
            _text(item, "holder_ctgry", "share_ctgry"),
            holder_name,
            "是" if is_org in ("1", "1.0", "true", "True", "是") else "否",
            _clean_ratio(_text(item, "held_ratio", "fundedRatio", "ratio")),
            _text(item, "invest_amount", "held_num"),
            _text(item, "holder_id", "orgId", "orgid"),
        ])

    # 实际控制人
    for item in equity.get("controllers") or []:
        person = _text(item, "name")
        if person:
            profile.controllers.append(person)

    # 最终受益人（比例与股东同口径，避免报告里受益比例恒为空）
    for item in equity.get("final_beneficiaries") or []:
        person = _text(item, "name")
        if person:
            profile.beneficiaries.append((
                person,
                _clean_ratio(_text(item, "czbl", "benefitRatio", "investRatio", "held_ratio", "ratio")),
            ))

    # 对外投资
    for item in _equity_children(equity.get("invests")):
        invest_name = _text(item, "name")
        if not invest_name:
            continue
        profile.investments.append((
            invest_name,
            _text(item, "orgid", "org_id", "id"),
            _clean_ratio(_text(item, "czbl", "investRatio", "ratio")),
        ))

    if not profile.holders:
        profile.warnings.append("最新公示股东未取得：" + profile.name)
    if not profile.investments:
        profile.warnings.append("对外投资未取得（可能确实没有对外投资）：" + profile.name)
    return profile


def _write_sheet(path: Path, headers, rows) -> None:
    import openpyxl
    from openpyxl.styles import Alignment, Font

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(list(headers))
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for row in rows:
        sheet.append(list(row))
    for index, header in enumerate(headers, 1):
        width = max(10, min(46, len(str(header)) * 2 + 6))
        sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def write_dimension_files(profiles: List[CompanyProfile], data_dir: Path, warnings: Optional[List[str]] = None) -> Dict[str, str]:
    """把逐家取到的数据写成与批量导出同名的维度文件，返回各维度的完成状态。

    每个维度统计实际覆盖的公司集合；少于成功取数公司时，把未覆盖名单写进 warnings 留痕。
    """
    data_dir = Path(data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    dimension_results: Dict[str, str] = {}

    basic_rows = [profile.basic for profile in profiles if profile.basic]
    _write_sheet(data_dir / "基础工商信息.xlsx", BASIC_HEADERS, basic_rows)
    dimension_results["S0000002"] = "completed" if basic_rows else "no_data"

    holder_rows = [row for profile in profiles for row in profile.holders]
    _write_sheet(data_dir / "最新公示股东.xlsx", HOLDER_HEADERS, holder_rows)
    dimension_results["S0000103"] = "completed" if holder_rows else "no_data"

    controller_rows = [
        [index, profile.name, person, "股权穿透", "注协企业股权关系接口"]
        for index, profile in enumerate(profiles, 1)
        for person in profile.controllers
    ]
    _write_sheet(data_dir / "实际控制人.xlsx", CONTROLLER_HEADERS, controller_rows)
    dimension_results["S0000032"] = "completed" if controller_rows else "no_data"

    beneficiary_rows = [
        [index, profile.name, "最终受益人", 1, "", person, ratio]
        for index, profile in enumerate(profiles, 1)
        for person, ratio in profile.beneficiaries
    ]
    _write_sheet(data_dir / "最终受益人.xlsx", BENEFICIARY_HEADERS, beneficiary_rows)
    dimension_results["S0000020"] = "completed" if beneficiary_rows else "no_data"

    invest_rows = [
        [index, profile.name, name, company_id, ratio]
        for index, profile in enumerate(profiles, 1)
        for name, company_id, ratio in profile.investments
    ]
    _write_sheet(data_dir / "对外投资（新）.xlsx", INVEST_HEADERS, invest_rows)
    dimension_results["S0000104"] = "completed" if invest_rows else "no_data"

    if warnings is not None:
        succeeded = {profile.name for profile in profiles}
        coverage = (
            ("基础工商信息", {profile.name for profile in profiles if profile.basic}),
            ("最新公示股东", {profile.name for profile in profiles if profile.holders}),
            ("实际控制人", {profile.name for profile in profiles if profile.controllers}),
            ("最终受益人", {profile.name for profile in profiles if profile.beneficiaries}),
            ("对外投资（新）", {profile.name for profile in profiles if profile.investments}),
        )
        for label, covered in coverage:
            uncovered = succeeded - covered
            if uncovered:
                warnings.append("维度{}未覆盖公司：{}（这些公司该维度未取得，报告中如实标注为缺口）".format(
                    label, "、".join(sorted(uncovered))))
    return dimension_results


def write_scope_note(
    data_dir: Path,
    company_names: List[str],
    dimension_results: Dict[str, str],
    *,
    fallback_reason: str,
    warnings: List[str],
) -> Path:
    """写取数说明，交代为什么走逐家取数、覆盖了哪些维度、缺了哪些。"""
    payload = {
        "company_names": list(company_names),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "data_route": "逐家取数（批量上传通道不可用时的保底做法）",
        "fallback_reason": str(fallback_reason),
        "required_dimensions": [code for code, _file in SINGLE_COMPANY_DIMENSIONS],
        "dimension_results": dict(dimension_results),
        "batch_history": [],
        "unavailable_dimensions": list(UNAVAILABLE_DIMENSIONS),
        "warnings": list(dict.fromkeys(warnings)),
    }
    path = Path(data_dir) / "取数说明.json"
    atomic_write_json(path, payload)
    return path


def collect_by_company(
    client,
    company_names: List[str],
    data_dir: Path,
    *,
    fallback_reason: str = "",
) -> Tuple[Dict[str, str], List[str]]:
    """逐个企业取数并落盘，返回（维度完成情况，告警）。"""
    names = [str(name).strip() for name in company_names if str(name).strip()]
    names = list(dict.fromkeys(names))
    if not names:
        raise SingleCompanyError("逐家取数的企业名单为空")

    profiles: List[CompanyProfile] = []
    warnings: List[str] = []
    for name in names:
        try:
            profiles.append(fetch_company(client, name))
        except Exception as exc:  # 单个企业失败不阻断其余企业
            warnings.append("逐家取数失败：{}（{}）".format(name, exc))

    if not profiles:
        raise SingleCompanyError("全部企业逐家取数均失败，未取得任何可用资料")

    dimension_results = write_dimension_files(profiles, data_dir, warnings=warnings)
    for profile in profiles:
        warnings.extend(profile.warnings)
    write_scope_note(
        data_dir,
        [profile.name for profile in profiles],
        dimension_results,
        fallback_reason=fallback_reason,
        warnings=warnings,
    )
    return dimension_results, warnings
