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
##!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
关联方核查引擎 / Related Party Identification Engine
=====================================================
读取 cicpa-company-query 完整导出的 _files 目录（52个维度 xlsx），
对其中所有公司做七类有效核查证据比对，输出多 sheet Excel 核查报告。

规则集对齐证监会/财政部近年处罚案例（蓝山科技、天沃科技、卓朗科技、
爱康科技、合纵科技等），把"监管认定的最低核查动作"固化成自动判定。

用法:
    python3 related_party_check.py \
        --data-dir "完整维度导出_files/" \
        --target "被审计单位全称" \
        -o "关联方核查报告.xlsx"
"""
import argparse
from datetime import date
import os
from pathlib import Path
import re
import sys
import hashlib
import json
import html
from collections import defaultdict, deque
from dataclasses import dataclass, field
from itertools import combinations
from typing import Optional

import openpyxl

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.证据追溯 import SourceRow, trace, unique_sources, website_domain

# ============================================================
# 常量
# ============================================================

# 基础工商信息列映射（A:企业名称 ... W:经营范围）
BASIC_COLS = {
    "name": 1, "company_id": 2, "status": 3, "legal_person": 4,
    "capital": 5, "found_date": 6, "province": 7, "city": 8, "district": 9,
    "phone": 10, "website": 11, "email": 12, "credit_code": 13,
    "reg_no": 14, "org_code": 15, "insured": 16, "company_type": 17,
    "industry1": 18, "industry2": 19, "industry3": 20,
    "former_name": 21, "address": 22, "business_scope": 23,
}

# 政府/公共实体名（作为实控人/受益人/关键人员出现时排除，避免所有国企误报关联）
EXCLUDED_ENTITIES = {
    "国务院国有资产监督管理委员会", "国务院", "财政部", "国家发改委",
    "中华人民共和国财政部", "中华人民共和国国务院",
    "地方国资委", "省国资委", "市国资委",
    "上海市国有资产监督管理委员会", "北京市国有资产监督管理委员会",
    "广东省人民政府国有资产监督管理委员会", "浙江省人民政府国有资产监督管理委员会",
    "江苏省人民政府国有资产监督管理委员会", "山东省人民政府国有资产监督管理委员会",
    "四川省政府国有资产监督管理委员会", "湖北省人民政府国有资产监督管理委员会",
    "湖南省人民政府国有资产监督管理委员会", "河南省人民政府国有资产监督管理委员会",
    "河北省人民政府国有资产监督管理委员会", "福建省人民政府国有资产监督管理委员会",
    "安徽省人民政府国有资产监督管理委员会", "辽宁省人民政府国有资产监督管理委员会",
    "陕西省人民政府国有资产监督管理委员会", "山西省人民政府国有资产监督管理委员会",
    "重庆市国有资产监督管理委员会", "天津市人民政府国有资产监督管理委员会",
    "江西省国有资产监督管理委员会", "云南省人民政府国有资产监督管理委员会",
    "广西壮族自治区人民政府国有资产监督管理委员会", "贵州省人民政府国有资产监督管理委员会",
    "新疆维吾尔自治区人民政府国有资产监督管理委员会",
}
# 公共邮箱域名（同这些域名不算关联信号）
PUBLIC_EMAIL_DOMAINS = {
    "qq.com", "163.com", "126.com", "sina.com", "sohu.com", "hotmail.com",
    "gmail.com", "outlook.com", "foxmail.com", "yeah.net", "139.com", "189.cn",
    "aliyun.com", "wo.cn", "vip.qq.com", "vip.163.com", "tom.com", "21cn.com",
    "188.com", "2980.com", "263.net", "mail.com", "yahoo.com", "live.com",
    "icloud.com", "me.com", "msn.com",
}

# 维度文件名（容错：缺文件跳过）
FILES = {
    "basic": "基础工商信息.xlsx",
    "shareholder": "股东信息.xlsx",
    "shareholder_new": "最新公示股东.xlsx",
    "actual_controller": "实际控制人.xlsx",
    "ultimate_beneficiary": "最终受益人.xlsx",
    "invest_new": "对外投资（新）.xlsx",
    "invest": "对外投资.xlsx",
    "holding": "参控股企业.xlsx",
    "branch": "分支机构.xlsx",
    "core_team": "核心团队.xlsx",
    "main_persons": "主要人员（高管）.xlsx",
    "legal_change": "法定代表人变更.xlsx",
    "customer": "客户.xlsx",
    "supplier": "供应商.xlsx",
    "guarantee": "对外担保.xlsx",
    "pledge": "股权质押.xlsx",
    "pledge2": "股权出质.xlsx",
    "mortgage": "动产抵押.xlsx",
    "change": "变更记录.xlsx",
    "trademark": "商标.xlsx",
    "software": "软件著作权.xlsx",
    "patent": "专利信息.xlsx",
    "wechat": "微信公众号.xlsx",
    "invoice": "发票信息.xlsx",
    "abnormal": "经营异常.xlsx",
}

# 风险等级
HARD = "🔴硬关联"      # high
MEDIUM = "🟡可疑红旗"   # medium
LOW = "🟢轻微异常"      # low

# ============================================================
# 数据加载（容错：文件缺失/空文件返回空表）
# ============================================================

def load_workbook_safe(path, errors=None):
    """加载 xlsx，返回 (rows_list, ncols)；文件缺失或异常返回 ([], 0)。"""
    if not os.path.exists(path):
        return [], 0
    wb = None
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = []
        headers = tuple(cell.value for cell in ws[1])
        for number, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            if any(c is not None and str(c).strip() != "" for c in r):
                rows.append(SourceRow(r, path, ws.title, number, headers))
        return rows, ws.max_column
    except Exception as e:
        print(f"  ⚠️ 读取失败 {os.path.basename(path)}: {e}", file=sys.stderr)
        if errors is not None:
            errors.append({
                "category": "读取失败",
                "source": os.path.basename(path),
                "message": str(e),
            })
        return [], 0
    finally:
        if wb is not None:
            wb.close()


def build_dim_index(data_dir, errors=None, limitations=None, file_info=None):
    """加载所有维度文件，返回 {key: {公司名: [行...]}} 索引。"""
    index = {}
    metadata_path = Path(data_dir) / "取数说明.json"
    metadata = {}
    if metadata_path.is_file():
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8-sig"))
            if not isinstance(metadata, dict):
                raise ValueError("取数说明不是对象")
            if not isinstance(metadata.get("dimension_results", {}), dict):
                raise ValueError("维度完成情况不是对象")
            names = metadata.get("company_names", [])
            if not isinstance(names, list) or any(not isinstance(name, str) or not name.strip() for name in names):
                raise ValueError("取数范围不是有效的公司名称列表")
        except (ValueError, OSError) as exc:
            if errors is not None:
                errors.append({"category": "取数说明读取失败", "source": str(metadata_path), "message": str(exc)})
            metadata = {}
    from scripts.cicpa.exporter import DIMENSION_NAMES
    codes = {name + ".xlsx": code for code, name in DIMENSION_NAMES.items()}
    for key, fname in FILES.items():
        path = os.path.join(data_dir, fname)
        terminal = metadata.get("dimension_results", {}).get(codes.get(fname), "")
        if not os.path.exists(path) and limitations is not None and fname in codes and terminal != "no_data":
            limitations.append({
                "category": "缺失维度",
                "source": fname,
                "message": "本次导出未提供该维度，相关规则可能受限",
            })
        rows, ncols = load_workbook_safe(path, errors=errors)
        minimum = {"basic": 23, "shareholder": 7, "shareholder_new": 6, "actual_controller": 3,
                   "ultimate_beneficiary": 6, "main_persons": 4, "core_team": 5, "invest_new": 5,
                   "holding": 4, "invoice": 7, "change": 5, "legal_change": 3, "pledge": 6,
                   "mortgage": 6, "trademark": 3, "software": 5, "wechat": 4}.get(key, 1)
        malformed = bool(rows and ncols < minimum)
        if malformed and errors is not None:
            errors.append({"category": "资料字段不足", "source": fname, "message": f"应至少有{minimum}列，实际只有{ncols}列，相关判断受限"})
        by_company = defaultdict(list)
        for r in rows:
            # 公司名称一般在第2列（B），基础工商在第1列（A），分支机构/担保/法人变更在第1列
            if key == "basic":
                name = r[0] if r else None
            elif key in ("branch", "guarantee", "legal_change", "invoice"):
                name = r[0] if r else None  # 这些表公司名在A列
            else:
                name = r[1] if len(r) > 1 else None
            if name and str(name).strip():
                by_company[str(name).strip()].append(r)
        index[key] = dict(by_company)
        if file_info is not None and fname in codes:
            file_info.append({"key": key, "file": str(Path(path).resolve()), "exists": Path(path).is_file(),
                "readable": ncols > 0, "companies": list(by_company), "scope": metadata.get("company_names", []),
                "terminal": terminal, "created_at": metadata.get("created_at", ""), "malformed": malformed})
        cnt = sum(len(v) for v in by_company.values())
        companies = len(by_company)
        if companies:
            print(f"  ✓ {fname}: {companies} 家公司, {cnt} 条记录")
        else:
            print(f"  · {fname}: 无数据")
    return index


# ============================================================
# 数据清洗
# ============================================================

def strip_html(s):
    if s is None:
        return ""
    value = html.unescape(re.sub(r"<[^>]+>", "", str(s))).strip()
    return "" if value.lower() in {"-", "--", "—", "暂无", "暂无数据", "无", "未知", "不详", "未披露", "未公示", "null", "none", "n/a", "nan"} else value


def normalize_address(addr):
    """地址归一化：去HTML、去空格标点、去邮编。"""
    if not addr:
        return ""
    s = strip_html(addr).strip()
    s = s.replace(" ", "").replace("\u3000", "").replace("\t", "")
    s = re.sub(r"[，,。；;:：、（）()\[\]【】「」“”‘’\-—_~`'\"#]", "", s)
    # 只去明确标为邮编的内容，保留行政区和门牌数字。
    s = re.sub(r"(?:邮政编码|邮编)\s*\d{6}", "", s)
    return s


def address_relationship(a1, a2):
    """比较两个地址。返回 (类型, 说明) 或 None。
    类型: exact(完全相同) / contains(前缀包含，同栋楼不同房间)"""
    n1, n2 = normalize_address(a1), normalize_address(a2)
    if not n1 or not n2 or len(n1) < 6 or len(n2) < 6:
        return None
    if n1 == n2:
        return ("exact", f"完全相同: {n1[:30]}")
    # 前缀包含（较短的地址是较长地址的前缀 → 同栋楼不同房间）
    short, long_ = (n1, n2) if len(n1) <= len(n2) else (n2, n1)
    if len(short) >= 8 and long_.startswith(short):
        return ("contains", f"地址前缀重合，完整地点待核实: {short}")
    # 提取"路/街/号+楼栋"核心段比对（同栋楼）
    c1 = re.sub(r"\d+(?:层|室|房|单元).*$", "", n1)
    c2 = re.sub(r"\d+(?:层|室|房|单元).*$", "", n2)
    if c1 == c2 and len(c1) >= 8 and c1 != n1 and c2 != n2 and any(mark in c1 for mark in ("号", "座", "栋", "大厦", "大楼")):
        return ("same_building", f"同一完整地址前缀、不同房间，待核实: {c1}")
    return None


def normalize_phones(phone_str):
    """拆分多值电话，归一化，返回电话集合。"""
    if not phone_str:
        return set()
    s = strip_html(phone_str)
    parts = re.split(r"[,，;；、/\s]+", s)
    result = set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        p = re.sub(r"[\-—()（）\s]", "", p)
        if p.startswith("+86"):
            p = p[3:]
        elif p.startswith("86") and len(p) > 11:
            p = p[2:]
        # 手机
        if re.fullmatch(r"1[3-9]\d{9}", p):
            result.add(p)
        # 座机 0XX-XXXXXXXX
        elif re.fullmatch(r"0\d{10,11}", p):
            result.add(p)  # 含区号
        elif re.fullmatch(r"\d{7,8}", p):
            result.add(p)
        # 4位区号+8位
        elif re.fullmatch(r"0\d{9,10}", p):
            result.add(p)
    # 过滤掉太短的
    return {x for x in result if len(x) >= 7}


def phone_segment_adjacent(phones_a, phones_b):
    """只返回完整座机号相差一位数值的弱线索，不据此推定同一办公地点。"""
    for pa in sorted(phones_a):
        for pb in sorted(phones_b):
            if pa.startswith("0") and pb.startswith("0") and len(pa) >= 10 and len(pb) >= 10:
                if pa[:-1] == pb[:-1] and abs(int(pa) - int(pb)) == 1:
                    return (pa, pb)
    return None


def normalize_emails(email_str):
    """返回 (完整邮箱列表, 非公共域名列表)。"""
    if not email_str:
        return [], []
    s = strip_html(email_str).lower()
    found = re.findall(r"[\w.+-]+@[\w.-]+\.\w{2,}", s)
    emails, domains = [], []
    for em in found:
        emails.append(em)
        domain = em.split("@")[-1]
        if domain not in PUBLIC_EMAIL_DOMAINS:
            domains.append(domain)
    return emails, list(set(domains))


def normalize_name(n):
    """姓名归一化（去HTML、去空格）。"""
    if not n:
        return ""
    s = strip_html(n).strip()
    return s.replace(" ", "").replace("\u3000", "")


def parse_capital(s):
    """解析注册资本字符串为元（float）。'1.5亿'→1.5e8, '2000万'→2e7。"""
    if s is None:
        return None
    s2 = re.sub(r"人民币|RMB|CNY|[元,，\s()（）]", "", strip_html(s), flags=re.I)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)(亿|万)?", s2)
    if m:
        return float(m.group(1)) * {"亿": 1e8, "万": 1e4, None: 1}[m.group(2)]
    return None


def parse_insured(s):
    """参保人数 → int。"""
    if s is None:
        return None
    s2 = re.sub(r"[,，\s]", "", strip_html(s))
    m = re.fullmatch(r"(\d+)(?:\.0+)?(?:人)?", s2)
    return int(m.group(1)) if m else None


def parse_found_date(s):
    """成立日期 → 'YYYY-MM-DD' 或 'YYYY' 字符串。"""
    if not s:
        return None
    s2 = strip_html(s).split(" ")[0].replace("年", "-").replace("月", "-").rstrip("日-").replace("/", "-")
    m = re.fullmatch(r"(\d{4})(?:-(\d{1,2})(?:-(\d{1,2}))?)?", s2)
    if m:
        year, month, day = (int(value) if value else None for value in m.groups())
        try:
            date(year, month or 1, day or 1)
        except ValueError:
            return None
        return f"{year:04d}-{month:02d}-{day:02d}" if day else (f"{year:04d}-{month:02d}" if month else str(year))
    return None


def row_value(row, index, default=None):
    """安全读取零基列号，短行返回默认值。"""
    return row[index] if index < len(row) else default


# ============================================================
# 公司实体（指纹）
# ============================================================

@dataclass
class Company:
    name: str
    legal_person: str = ""
    phones: set = field(default_factory=set)
    emails: list = field(default_factory=list)
    email_domains: list = field(default_factory=list)
    addresses: list = field(default_factory=list)
    websites: list = field(default_factory=list)
    capital: Optional[float] = None
    found_date: Optional[str] = None
    insured: Optional[int] = None
    former_names: list = field(default_factory=list)
    business_scope: str = ""
    industry: str = ""
    status: str = ""
    # 关系数据
    shareholders: list = field(default_factory=list)        # [(股东名, 比例str, 是否机构)]
    actual_controllers: list = field(default_factory=list)  # [实控人名]
    beneficiaries: list = field(default_factory=list)       # [(受益人名, 比例)]
    core_persons: list = field(default_factory=list)        # [(姓名, 职务)]
    main_persons: list = field(default_factory=list)        # [(姓名, 职务)]
    investments: list = field(default_factory=list)         # [(被投资企业, 比例)]
    holdings: list = field(default_factory=list)            # [(参控股企业, 比例)]
    trademarks: list = field(default_factory=list)
    softwares: list = field(default_factory=list)
    wechats: list = field(default_factory=list)
    pledges_out: list = field(default_factory=list)   # [(出质人, 质权人)]
    mortgages: list = field(default_factory=list)     # [(抵押人, 抵押权人)]
    changes: list = field(default_factory=list)       # [(日期, 项目, 变更前, 变更后)]
    legal_changes: list = field(default_factory=list) # [(变更前法人, 变更后法人)]
    invoice_addr: str = ""
    invoice_phone: str = ""
    source_rows: dict = field(default_factory=dict)


def build_company(name, dim, basic_row):
    """从基础工商行 + 各维度构建 Company 指纹。"""
    c = Company(name=name)
    c.source_rows = {key: values.get(name, []) for key, values in dim.items()}
    if basic_row:
        c.source_rows["basic"] = [basic_row]
    if basic_row:
        c.legal_person = normalize_name(basic_row[BASIC_COLS["legal_person"] - 1])
        c.phones = normalize_phones(basic_row[BASIC_COLS["phone"] - 1])
        c.emails, c.email_domains = normalize_emails(basic_row[BASIC_COLS["email"] - 1])
        addr = basic_row[BASIC_COLS["address"] - 1]
        if addr:
            c.addresses.append(strip_html(addr))
        web = basic_row[BASIC_COLS["website"] - 1]
        if web:
            c.websites.append(strip_html(str(web)))
        c.capital = parse_capital(basic_row[BASIC_COLS["capital"] - 1])
        c.found_date = parse_found_date(basic_row[BASIC_COLS["found_date"] - 1])
        c.insured = parse_insured(basic_row[BASIC_COLS["insured"] - 1])
        fn = basic_row[BASIC_COLS["former_name"] - 1]
        if fn and str(fn).strip():
            c.former_names = [normalize_name(x) for x in re.split(r"[,，;；、]", strip_html(fn)) if x.strip()]
        c.business_scope = strip_html(basic_row[BASIC_COLS["business_scope"] - 1])
        c.status = strip_html(basic_row[BASIC_COLS["status"] - 1])
        ind = basic_row[BASIC_COLS["industry2"] - 1] or basic_row[BASIC_COLS["industry1"] - 1]
        c.industry = strip_html(ind) if ind else ""

    # 股东
    for r in dim.get("shareholder", {}).get(name, []):
        if len(r) >= 7:
            sh = normalize_name(r[2])
            ratio = strip_html(str(r[4])) if r[4] else ""
            is_inst = strip_html(str(r[6])) if r[6] else ""
            if sh:
                c.shareholders.append((sh, ratio, is_inst.lower() in ("是", "true", "1", "1.0", "机构", "法人") or is_organization(sh)))
    # 最新公示股东
    for r in dim.get("shareholder_new", {}).get(name, []):
        if len(r) >= 6:
            sh = normalize_name(r[3])
            ratio = strip_html(str(r[5])) if r[5] else ""
            if sh:
                c.shareholders.append((sh, ratio, strip_html(r[4]).lower() in ("是", "true", "1", "1.0", "机构", "法人") or is_organization(sh)))
    # 实控人 / 最终受益人
    for r in dim.get("actual_controller", {}).get(name, []):
        if len(r) >= 3 and r[2]:
            c.actual_controllers.append(normalize_name(r[2]))
    for r in dim.get("ultimate_beneficiary", {}).get(name, []):
        if len(r) >= 6 and r[5]:
            c.beneficiaries.append((
                normalize_name(r[5]),
                strip_html(str(row_value(r, 6, "") or "")),
            ))
    # 核心团队 / 主要人员
    for r in dim.get("core_team", {}).get(name, []):
        if len(r) >= 5 and r[3]:
            c.core_persons.append((normalize_name(r[3]), strip_html(str(r[4] or ""))))
    for r in dim.get("main_persons", {}).get(name, []):
        if len(r) >= 4 and r[2]:
            c.main_persons.append((normalize_name(r[2]), strip_html(str(r[3] or ""))))
    # 对外投资 / 参控股
    for r in dim.get("invest_new", {}).get(name, []):
        if len(r) >= 5 and r[2]:
            c.investments.append((normalize_name(r[2]), strip_html(str(r[4] or ""))))
    for r in dim.get("holding", {}).get(name, []):
        if len(r) >= 4 and r[2]:
            c.holdings.append((normalize_name(r[2]), strip_html(str(r[3] or ""))))
    # 商标 / 软件 / 公众号
    for r in dim.get("trademark", {}).get(name, []):
        if len(r) >= 3 and r[2]:
            c.trademarks.append(strip_html(str(r[2])))
    for r in dim.get("software", {}).get(name, []):
        if len(r) >= 5 and r[4]:
            c.softwares.append(strip_html(str(r[4])))
    for r in dim.get("wechat", {}).get(name, []):
        if len(r) >= 4 and r[3]:
            c.wechats.append(strip_html(str(r[3])))
    # 股权质押 / 动产抵押
    for r in dim.get("pledge", {}).get(name, []):
        if len(r) >= 6:
            c.pledges_out.append((normalize_name(r[3] or ""), normalize_name(r[5] or "")))
    for r in dim.get("mortgage", {}).get(name, []):
        if len(r) >= 6:
            c.mortgages.append((normalize_name(r[4] or ""), normalize_name(r[5] or "")))
    # 变更记录 / 法人变更
    for r in dim.get("change", {}).get(name, []):
        if len(r) >= 5:
            c.changes.append((strip_html(str(r[2] or "")), strip_html(str(r[3] or "")),
                              strip_html(str(r[4] or "")),
                              strip_html(str(row_value(r, 5, "") or ""))))
    for r in dim.get("legal_change", {}).get(name, []):
        if len(r) >= 3:
            c.legal_changes.append((normalize_name(r[1] or ""), normalize_name(r[2] or "")))
    # 发票信息（补充地址电话）
    for r in dim.get("invoice", {}).get(name, []):
        if len(r) >= 7:
            if r[4]:
                c.invoice_addr = strip_html(str(r[4]))
            if r[6]:
                c.invoice_phone = strip_html(str(r[6]))
                c.phones.update(normalize_phones(r[6]))
    return c


def is_public_authority(name):
    """只排除政府部门同受国家控制的线索，国有企业母公司仍正常核查。"""
    return name in EXCLUDED_ENTITIES or str(name).endswith(("人民政府", "国有资产监督管理委员会", "国资委", "财政厅", "财政局"))


def is_organization(name):
    return is_public_authority(name) or any(word in str(name) for word in ("公司", "企业", "合伙", "合作社", "委员会"))


def all_person_names(c: Company):
    """公司所有关键人员姓名集合（法人+股东中的自然人+核心团队+主要人员+实控人+受益人）。"""
    names = set()
    if c.legal_person:
        names.add(c.legal_person)
    for sh, _, is_inst in c.shareholders:
        if not is_inst and sh:  # 使用机构标志，不以姓名长度判断自然人。
            names.add(sh)
    for n, _ in c.core_persons:
        names.add(n)
    for n, _ in c.main_persons:
        names.add(n)
    for n in c.actual_controllers:
        names.add(n)
    for n, _ in c.beneficiaries:
        names.add(n)
    return {n for n in names if strip_html(n) and len(n) >= 2 and not is_organization(n)}


def all_addresses(c: Company):
    """当前已取得的工商和发票地址；历史地址由历史规则单独检查。"""
    addrs = list(c.addresses)
    if c.invoice_addr:
        addrs.append(c.invoice_addr)
    addrs.extend(strip_html(row[4]) for row in c.source_rows.get("invoice", []) if len(row) >= 5 and row[4])
    return list(dict.fromkeys(a for a in addrs if a))


# ============================================================
# 七类有效核查证据
# ============================================================

def rule1_fingerprint(ca: Company, cb: Company):
    """维度1: 工商指纹重合（地址/电话/邮箱/网址）。"""
    hits = []
    # 电话
    common_phones = ca.phones & cb.phones
    if common_phones:
        complete = any(len(number) >= 10 for number in common_phones)
        hits.append(("phone", HARD if complete else MEDIUM, f"联系电话相同: {','.join(sorted(common_phones))}" + ("" if complete else "（缺区号，地域待核实）"), "蓝山/卓朗/达志科技案",
                     trace(ca, "phone", common_phones, normalize=normalize_phones) + trace(cb, "phone", common_phones, normalize=normalize_phones)))
    else:
        # 电话号段相邻（座机同一区号+局向，末位不同 → 同一办公地点的连续号码）
        seg = phone_segment_adjacent(ca.phones, cb.phones)
        if seg:
            hits.append(("phone_segment", MEDIUM, f"座机号段相邻: {seg[0]}↔{seg[1]}（待核实，不能据此认定同一地点）", "蓝山科技案(号段相邻)",
                         trace(ca, "phone", [seg[0]], normalize=normalize_phones) + trace(cb, "phone", [seg[1]], normalize=normalize_phones)))
    # 邮箱完全相同
    common_emails = set(ca.emails) & set(cb.emails)
    if common_emails:
        hits.append(("email", HARD, f"邮箱完全相同: {','.join(sorted(common_emails))}", "爱康科技案",
                     trace(ca, "email", common_emails, normalize=lambda x: normalize_emails(x)[0]) + trace(cb, "email", common_emails, normalize=lambda x: normalize_emails(x)[0])))
    # 邮箱同域名（非公共）
    common_domains = set(ca.email_domains) & set(cb.email_domains)
    if common_domains:
        hits.append(("email_domain", HARD, f"企业邮箱同域名: {','.join(sorted(common_domains))}", "天沃/爱康/志高机械案",
                     trace(ca, "email", common_domains, normalize=lambda x: normalize_emails(x)[1]) + trace(cb, "email", common_domains, normalize=lambda x: normalize_emails(x)[1])))
    # 地址
    for a1 in all_addresses(ca):
        for a2 in all_addresses(cb):
            rel = address_relationship(a1, a2)
            if rel:
                kind, desc = rel
                level = HARD if kind == "exact" else (MEDIUM if kind == "same_building" else MEDIUM)
                case = "天沃科技案(同楼同座)" if kind != "exact" else "达志科技案(地址相同)"
                hits.append(("address", level, f"{desc}；甲方地址：{a1}；乙方地址：{a2}", case,
                             trace(ca, "address", [a1], normalize=strip_html) + trace(cb, "address", [a2], normalize=strip_html)))
                break
    # 网址同域名
    def web_domain(w):
        return re.sub(r"https?://", "", str(w)).split("/")[0].replace("www.", "").lower() if w else ""
    wa = {web_domain(w) for w in ca.websites if w}
    wb = {web_domain(w) for w in cb.websites if w}
    common_web = (wa & wb) - {""}
    if common_web:
        hits.append(("website", MEDIUM, f"官网同域名: {','.join(sorted(common_web))}", "实务红旗",
                     trace(ca, "website", common_web, normalize=website_domain) + trace(cb, "website", common_web, normalize=website_domain)))
    return hits


def rule2_personnel(ca: Company, cb: Company):
    """维度2: 关键人员重合。"""
    hits = []
    pa, pb = all_person_names(ca), all_person_names(cb)
    common = pa & pb
    # 法人重合是强信号
    if ca.legal_person in common and ca.legal_person == cb.legal_person:
        hits.append(("legal_person", HARD, f"法定代表人同名: {ca.legal_person}", "爱康/合纵科技案",
                     trace(ca, "legal_person") + trace(cb, "legal_person")))
    # 其他人员重合
    other_common = common - ({ca.legal_person, cb.legal_person} if ca.legal_person == cb.legal_person else set())
    if other_common:
        # 区分：是否任董监高（强）/ 仅股东（中）
        is_exec = any(n in {x[0] for x in ca.core_persons + ca.main_persons + cb.core_persons + cb.main_persons} for n in other_common)
        level = HARD if is_exec else MEDIUM
        case = "卓朗科技案(对方监事为本公司员工)" if is_exec else "人员重合"
        hits.append(("person", level, f"关键人员同名: {','.join(sorted(other_common))}", case,
                     trace(ca, "person", other_common, normalize=normalize_name) + trace(cb, "person", other_common, normalize=normalize_name)))
    # 实控人/最终受益人一致（排除政府/公共实体）
    common_ctrl = {n for n in set(ca.actual_controllers) & set(cb.actual_controllers) if strip_html(n) and not is_public_authority(n)}
    if common_ctrl:
        hits.append(("controller", HARD, f"实际控制人同名: {','.join(sorted(common_ctrl))}", "股权控制类",
                     trace(ca, "controller", common_ctrl, normalize=normalize_name) + trace(cb, "controller", common_ctrl, normalize=normalize_name)))
    common_ben = {n for n in {b[0] for b in ca.beneficiaries} & {b[0] for b in cb.beneficiaries} if strip_html(n) and not is_public_authority(n)}
    if common_ben:
        hits.append(("beneficiary", MEDIUM, f"最终受益人同名: {','.join(sorted(common_ben))}（受益比例及实际控制另行核实）", "受益关系线索",
                     trace(ca, "beneficiary", common_ben, normalize=normalize_name) + trace(cb, "beneficiary", common_ben, normalize=normalize_name)))
    return hits


def months_since_found(found_date, as_of_date):
    """按年月计算成立至核查基准日的完整月数。"""
    if not found_date:
        return None
    parsed = parse_found_date(found_date)
    if not parsed or len(parsed) < 7:
        return None
    year, month = map(int, parsed[:7].split("-"))
    months = (as_of_date.year - year) * 12 + (as_of_date.month - month)
    if len(parsed) == 10 and as_of_date.day < int(parsed[-2:]):
        months -= 1
    return months


def rule3_counterparty_profile(ca: Company, cb: Company, target_set, as_of_date=None):
    """维度3: 客商异常画像（对非审计对象群的公司做单方向画像）。

    审计对象群 = 被审计单位 + 其重要子公司（--target 逗号分隔）。
    只对「审计对象 ↔ 对手方」做画像；双方都是/都不是审计对象则跳过。
    """
    hits = []
    ca_is_target = ca.name in target_set
    cb_is_target = cb.name in target_set
    if ca_is_target and not cb_is_target:
        party = cb
    elif cb_is_target and not ca_is_target:
        party = ca
    else:
        return hits  # 同属审计对象群（集团内）或都是外部对手方，不做画像
    flags = []
    as_of = as_of_date or date.today()
    age_months = months_since_found(party.found_date, as_of)
    if age_months is not None:
        if 0 <= age_months < 12:
            flags.append("成立不足12个月")
        elif 12 <= age_months < 24:
            flags.append("成立不足24个月")
    # 注册资本 vs 典型交易额：注册资本极小
    if party.capital is not None and party.capital < 1e6:  # <100万
        flags.append(f"注册资本仅{party.capital/1e4:.0f}万(过小)")
    # 参保人数
    if party.insured == 0:
        flags.append("参保人数0(空壳特征)")
    elif party.insured is not None and party.insured < 5:
        flags.append(f"参保人数{party.insured}(疑似空壳)")
    # 地址居民楼
    for a in party.addresses:
        if any(k in a for k in ["小区", "花园", "公寓", "苑", "村", "组"]):
            flags.append(f"注册地址疑似居民楼")
            break
    # 缺失经营范围列为数据缺口，不构造风险命中。
    if party.business_scope and len(party.business_scope) < 10:
        flags.append("经营范围极简")
    if flags:
        hits.append(("profile", MEDIUM if len(flags) >= 2 else LOW,
                     "；".join(flags), "专网通信/爱康案", trace(party, "profile")))
    return hits


def parse_percentage(value):
    """导出比例按百分数读取；不把未知比例当作零，也不猜小数的单位。"""
    text = strip_html(value).replace("%", "").replace("％", "").strip()
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    number = float(text)
    return number if 0 <= number <= 100 else None


def rule5_equity(ca: Company, cb: Company):
    """识别双向直接投资和共同股东；持股证据与最终控制认定分开。"""
    hits = []
    for parent, child in ((ca, cb), (cb, ca)):
        ratios = [ratio for name, ratio in parent.investments + parent.holdings if name == child.name]
        ratios += [ratio for name, ratio, _ in child.shareholders if name == parent.name]
        if ratios:
            parsed = [parse_percentage(ratio) for ratio in ratios]
            level = HARD if all(ratio is not None and ratio >= 20 for ratio in parsed) else MEDIUM
            hits.append(("invest", level, f"{parent.name} 对 {child.name} 的持股线索；记录比例：{'、'.join(sorted(set(str(r) or '未披露' for r in ratios)))}（控制或重大影响待核实）", "投资关系",
                trace(parent, "invest", [child.name], normalize=normalize_name) + trace(child, "shareholder", [parent.name], normalize=normalize_name)
                + trace(parent, "names", [parent.name], normalize=normalize_name) + trace(child, "names", [child.name], normalize=normalize_name)))
    common_sh = {name for name in {s[0] for s in ca.shareholders} & {s[0] for s in cb.shareholders}
                 if strip_html(name) and name not in {"国有企业", "自然人", "法人"} and not is_public_authority(name)}
    if common_sh:
        hits.append(("common_shareholder", MEDIUM,
            f"共同股东同名: {','.join(sorted(common_sh))}（结合持股比例及其他安排核实影响）", "股权穿透",
            trace(ca, "shareholder", common_sh, normalize=normalize_name) + trace(cb, "shareholder", common_sh, normalize=normalize_name)))
    return hits


def rule6_historical(ca: Company, cb: Company):
    """维度6: 历史关联痕迹。"""
    hits = []
    # 曾用名命中
    cb_names = {cb.name} | set(cb.former_names)
    ca_names = {ca.name} | set(ca.former_names)
    former_hit = (set(ca.former_names) & cb_names) | (set(cb.former_names) & ca_names)
    if former_hit:
        split_names = lambda x: [normalize_name(v) for v in re.split(r"[,，;；、]", strip_html(x))]
        hits.append(("former_name", MEDIUM, f"曾用名匹配: {','.join(sorted(former_hit))}", "历史关联",
                     trace(ca, "names", former_hit, normalize=split_names) + trace(cb, "names", former_hit, normalize=split_names)))
    # 曾任法人：A 现法人曾是 B 的法人（法人变更记录）
    ca_past_legals = {lc[0] for lc in ca.legal_changes if lc[0]} | {lc[1] for lc in ca.legal_changes if lc[1]}
    cb_past_legals = {lc[0] for lc in cb.legal_changes if lc[0]} | {lc[1] for lc in cb.legal_changes if lc[1]}
    # A 现法人 在 B 的法人变更历史里
    if ca.legal_person and ca.legal_person in cb_past_legals and ca.legal_person != cb.legal_person:
        hits.append(("past_legal", MEDIUM,
                     f"{ca.name}现法人{ca.legal_person}与{cb.name}历史法人同名", "华道生物案(代持痕迹)",
                     trace(ca, "legal_person") + trace(cb, "past_legal", [ca.legal_person], normalize=normalize_name)))
    # 变更记录里的历史地址撞对方现地址
    cb_addr_set = {normalize_address(a) for a in all_addresses(cb)}
    for _, item, before, after in ca.changes:
        if "地址" in item and before:
            if normalize_address(before) in cb_addr_set and len(normalize_address(before)) >= 8:
                hits.append(("past_address", MEDIUM,
                             f"{ca.name}曾用地址与{cb.name}已取得地址重合：{before}", "历史关联",
                             trace(ca, "past_address", [before], normalize=strip_html) + trace(cb, "address", [normalize_address(before)], normalize=normalize_address)))
    return hits


def equity_path_hits(companies):
    """在已取得资料内寻找最多五层的多数持股路径，不传递少数股权。"""
    records = defaultdict(list)
    for company in companies.values():
        for child, ratio in company.investments + company.holdings:
            records[(company.name, child)].append((parse_percentage(ratio), trace(company, "invest", [child], normalize=normalize_name)))
        for parent, ratio, _ in company.shareholders:
            records[(parent, company.name)].append((parse_percentage(ratio), trace(company, "shareholder", [parent], normalize=normalize_name)))
    graph = defaultdict(dict)
    for (parent, child), values in records.items():
        if parent and child and parent != child and not is_public_authority(parent) and all(ratio is not None and ratio > 50 for ratio, _ in values):
            graph[parent][child] = (min(ratio for ratio, _ in values), unique_sources([s for _, sources in values for s in sources]))
    paths = {}
    for root in sorted(graph):
        found = {}
        queue = deque([(root, [root], [], [])])
        visited = {root}
        while queue:
            node, nodes, ratios, sources = queue.popleft()
            if len(ratios) >= 5:
                continue
            for child, (ratio, evidence) in sorted(graph.get(node, {}).items()):
                if child in visited:
                    continue
                visited.add(child)
                entry = (nodes + [child], ratios + [ratio], sources + evidence)
                found[child] = entry
                queue.append((child, *entry))
        paths[root] = found
    hits = []
    for a, b in combinations(sorted(companies), 2):
        chains = []
        if b in paths.get(a, {}) and len(paths[a][b][1]) > 1:
            chains = [paths[a][b]]
        elif a in paths.get(b, {}) and len(paths[b][a][1]) > 1:
            chains = [paths[b][a]]
        elif b not in graph.get(a, {}) and a not in graph.get(b, {}):
            ancestors = [root for root, found in paths.items() if a in found and b in found]
            if ancestors:
                root = min(ancestors, key=lambda root: (len(paths[root][a][1]) + len(paths[root][b][1]), root))
                chains = [paths[root][a], paths[root][b]]
        if not chains:
            continue
        descriptions = [nodes[0] + "".join(f" --{ratio:g}%--> {child}" for ratio, child in zip(ratios, nodes[1:])) for nodes, ratios, _ in chains]
        hits.append({"company_a": a, "company_b": b, "dimension": "股权控制穿透", "field": "equity_path", "level": HARD,
            "evidence": "逐层持股均超过50%的路径：" + "；".join(descriptions) + "（按持股资料形成线索，实际控制及记录时点待核实）",
            "case_ref": "持股路径核查", "sources": unique_sources([source for _, _, sources in chains for source in sources])})
    return hits


def rule7_guarantee(ca: Company, cb: Company):
    """维度7: 担保/资金链。"""
    hits = []
    cb_names = {cb.name} | set(cb.former_names)
    ca_names = {ca.name} | set(ca.former_names)
    # 股权质押：出质人或质权人是对方
    for pledgor, pledgee in ca.pledges_out:
        if pledgor in cb_names or pledgee in cb_names:
            hits.append(("pledge", MEDIUM,
                         f"股权质押线索(出质:{pledgor}/质权:{pledgee})", "康得新案",
                         trace(ca, "pledge", {pledgor, pledgee}, normalize=normalize_name,
                               row_match=lambda row: (normalize_name(row_value(row, 3)), normalize_name(row_value(row, 5))) == (pledgor, pledgee))
                         + trace(cb, "names", {pledgor, pledgee}, normalize=normalize_name)))
    # 动产抵押
    for mortgagor, mortgagee in ca.mortgages:
        if mortgagor in cb_names or mortgagee in cb_names:
            hits.append(("mortgage", MEDIUM,
                         f"动产抵押线索(抵押:{mortgagor}/抵押权:{mortgagee})", "担保链",
                         trace(ca, "mortgage", {mortgagor, mortgagee}, normalize=normalize_name,
                               row_match=lambda row: (normalize_name(row_value(row, 4)), normalize_name(row_value(row, 5))) == (mortgagor, mortgagee))
                         + trace(cb, "names", {mortgagor, mortgagee}, normalize=normalize_name)))
    return hits


def rule8_intangible(ca: Company, cb: Company):
    """维度8: 无形资产/资源共用。"""
    hits = []
    # 商标同名（排除"图形"等纯商标类别标签）
    common_tm = (set(ca.trademarks) & set(cb.trademarks)) - {"", "图形", "文字", "字母", "数字", "颜色"}
    if common_tm:
        hits.append(("trademark", MEDIUM, f"同名商标: {','.join(sorted(common_tm))}", "无形资产共用",
                     trace(ca, "trademark", common_tm, normalize=strip_html) + trace(cb, "trademark", common_tm, normalize=strip_html)))
    # 软件著作权同名
    common_sw = (set(ca.softwares) & set(cb.softwares)) - {""}
    if common_sw:
        hits.append(("software", MEDIUM, f"同名软件著作权: {','.join(sorted(common_sw))}", "无形资产共用",
                     trace(ca, "software", common_sw, normalize=strip_html) + trace(cb, "software", common_sw, normalize=strip_html)))
    # 微信公众号同名
    common_wc = (set(ca.wechats) & set(cb.wechats)) - {""}
    if common_wc:
        hits.append(("wechat", LOW, f"同名微信公众号: {','.join(sorted(common_wc))}", "资源共用",
                     trace(ca, "wechat", common_wc, normalize=strip_html) + trace(cb, "wechat", common_wc, normalize=strip_html)))
    return hits


# ============================================================
# 主比对 & 汇总
# ============================================================

def compare_pair(ca, cb, dim, target_set, as_of_date=None, errors=None):
    """对一对公司跑全部规则，返回 [hit_dict...]。target_set 为审计对象群（含子公司）。"""
    results = []
    # rule3 需要方向参数（判断谁是审计主体、谁是对手方），单独调用；其余规则统一双参数
    rules_2arg = [
        (rule1_fingerprint, "工商指纹重合"),
        (rule2_personnel, "关键人员重合"),
        (rule5_equity, "股权控制穿透"),
        (rule6_historical, "历史关联痕迹"),
        (rule7_guarantee, "担保资金链"),
        (rule8_intangible, "无形资产共用"),
    ]
    for rule_fn, dim_name in rules_2arg:
        try:
            rule_hits = rule_fn(ca, cb)
            if rule_fn in (rule6_historical, rule7_guarantee):
                rule_hits += rule_fn(cb, ca)
            for field_key, level, evidence, case, sources in rule_hits:
                results.append({
                    "company_a": ca.name, "company_b": cb.name,
                    "dimension": dim_name, "field": field_key,
                    "level": level, "evidence": evidence, "case_ref": case, "sources": unique_sources(sources),
                })
        except Exception as e:
            print(f"  ⚠️ 规则异常 {dim_name} {ca.name}↔{cb.name}: {e}", file=sys.stderr)
            if errors is not None:
                errors.append({
                    "category": "规则异常",
                    "source": dim_name,
                    "message": f"{ca.name}↔{cb.name}: {e}",
                })
    # rule3: 客商异常画像（需 target_set 判定方向，只对 审计对象↔对手方 做画像）
    try:
        for field_key, level, evidence, case, sources in rule3_counterparty_profile(
            ca,
            cb,
            target_set,
            as_of_date=as_of_date,
        ):
            results.append({
                "company_a": ca.name, "company_b": cb.name,
                "dimension": "客商异常画像", "field": field_key,
                "level": level, "evidence": evidence, "case_ref": case, "sources": unique_sources(sources),
            })
    except Exception as e:
        print(f"  ⚠️ 规则异常 客商异常画像 {ca.name}↔{cb.name}: {e}", file=sys.stderr)
        if errors is not None:
            errors.append({
                "category": "规则异常",
                "source": "客商异常画像",
                "message": f"{ca.name}↔{cb.name}: {e}",
            })
    return results


def level_rank(level):
    return {HARD: 3, MEDIUM: 2, LOW: 1}.get(level, 0)


def aggregate(all_hits, target_set):
    """按公司对汇总。返回 [{pair, levels, dimensions, max_level, is_related, ...}]。target_set 为审计对象群。"""
    by_pair = defaultdict(list)
    for h in all_hits:
        pair = tuple(sorted([h["company_a"], h["company_b"]]))
        by_pair[pair].append(h)
    summary = []
    for pair, hits in by_pair.items():
        levels = [h["level"] for h in hits]
        dims = sorted({h["dimension"] for h in hits})
        max_level = max(levels, key=level_rank)
        # 按公司对最高风险分级，最终认定由人工复核填写。
        has_hard = any(l == HARD for l in levels)
        has_medium = any(l == MEDIUM for l in levels)
        if has_hard:
            is_related = "高风险线索，待核实"
            suggestion = "实施函证、实地走访、资金流水核对；询问管理层，核实关系并评估披露"
        elif has_medium:
            is_related = "中风险线索，待核实"
            suggestion = "结合交易背景进一步核查；关注交易商业合理性"
        else:
            is_related = "低风险线索，待核实"
            suggestion = "记录备查，必要时跟进"
        ca, cb = pair
        ca_t, cb_t = ca in target_set, cb in target_set
        if ca_t and cb_t:
            relation = "审计对象之间"
        elif ca_t or cb_t:
            relation = "审计对象与核查对象"
        else:
            relation = "其他核查对象之间（不推定与审计对象关联）"
        summary.append({
            "company_a": ca, "company_b": cb, "relation_type": relation,
            "is_related": is_related, "max_level": max_level,
            "hit_count": len(hits), "dimensions": "、".join(dims),
            "evidence": " | ".join(sorted({h["evidence"] for h in hits})),
            "suggestion": suggestion,
            "case_ref": "、".join(sorted({h["case_ref"] for h in hits})),
            "hits": hits,
        })
    summary.sort(key=lambda x: (-level_rank(x["max_level"]), -x["hit_count"]))
    return summary


# ============================================================
# Excel 输出
# ============================================================

def write_report(*args, **kwargs):
    from scripts.报告输出 import write_report as write_excel_report
    return write_excel_report(*args, **kwargs)


# ============================================================
# main
# ============================================================

@dataclass
class CheckResult:
    output_path: Path
    summary: list
    hits: list
    errors: list
    limitations: list
    companies: dict


def run_check(
    *,
    data_dir,
    target_names,
    output_path,
    as_of_date=None,
    disclosed_parties=None,
    task_mode="existing_export",
    scope_metadata=None,
    additional_data_dirs=None,
    object_records=None,
    discovery_warnings=None,
):
    """运行核查并返回结构化结果，命令行只负责参数转换。"""
    data_path = Path(data_dir)
    if not data_path.is_dir():
        raise FileNotFoundError(f"数据目录不存在: {data_path}")
    if isinstance(target_names, str):
        target_set = {name.strip() for name in target_names.split(",") if name.strip()}
    else:
        target_set = {str(name).strip() for name in target_names if str(name).strip()}
    if not target_set:
        raise ValueError("被审计单位名称不能为空")

    check_date = as_of_date or date.today()
    if isinstance(check_date, str):
        check_date = date.fromisoformat(check_date)
    errors = []
    limitations = []

    print(f"📂 加载数据: {data_path}")
    file_info = []
    dim = {}
    data_paths = [data_path] + [Path(path) for path in (additional_data_dirs or [])]
    # 离线重读同一交付目录时，也纳入已经完成的候选批次。
    candidate_dir = data_path / "候选公司原始导出"
    if candidate_dir.is_dir() and candidate_dir not in data_paths:
        data_paths.append(candidate_dir)
    for folder in dict.fromkeys(data_paths):
        if not folder.is_dir():
            raise FileNotFoundError(f"数据目录不存在: {folder}")
        part = build_dim_index(folder, errors=errors, limitations=limitations, file_info=file_info)
        for key, by_name in part.items():
            for name, rows in by_name.items():
                dim.setdefault(key, {}).setdefault(name, []).extend(rows)
    company_names = list(dim.get("basic", {}).keys())
    if not company_names:
        raise ValueError("基础工商信息.xlsx 无数据或未找到，无法核查")
    missing = target_set - set(company_names)
    if missing:
        limitations.append({
            "category": "名称未匹配",
            "source": "基础工商信息.xlsx",
            "message": "以下审计对象未精确匹配：{}".format("、".join(sorted(missing))),
        })

    target_display = " / ".join(sorted(target_set))
    companies = {}
    data_completeness = {}
    for name in company_names:
        basic_rows = dim.get("basic", {}).get(name, [])
        basic_row = basic_rows[0] if basic_rows else None
        if len(basic_row or ()) < 23:
            errors.append({"category": "基础资料字段不足", "source": name, "message": "基础工商记录不足23列，未纳入比对"})
            continue
        if len(basic_rows) > 1:
            limitations.append({"category": "重复基础记录", "source": name, "message": "存在多条基础工商记录，本次使用第一条，需复核记录日期"})
        company = build_company(name, dim, basic_row)
        if company.found_date and len(company.found_date) == 10 and date.fromisoformat(company.found_date) > check_date:
            errors.append({"category": "数据字段异常", "source": name, "message": "成立日期晚于核查基准日，不能用于成立年限风险判断"})
        companies[name] = company
        filled = sum([
            bool(company.legal_person),
            bool(company.phones),
            bool(company.emails),
            bool(company.addresses),
            company.capital is not None,
            company.insured is not None,
        ])
        data_completeness[name] = f"{filled}/6"

    all_hits = []
    names = sorted(companies)
    for company_a, company_b in combinations(names, 2):
        all_hits.extend(
            compare_pair(
                companies[company_a],
                companies[company_b],
                dim,
                target_set,
                as_of_date=check_date,
                errors=errors,
            )
        )
    all_hits.extend(equity_path_hits(companies))
    # 同一事实重复命中只保留一条；编号不依赖明细的显示行号。
    identified = {}
    for hit in all_hits:
        identity = [sorted([hit["company_a"], hit["company_b"]]), hit["dimension"], hit["field"], hit["evidence"],
                    [(s["company"], Path(s["file"]).name, s["sheet"], s["cell"], s["value"]) for s in hit["sources"]]]
        hit["evidence_id"] = "ZJ-" + hashlib.sha256(json.dumps(identity, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()[:10].upper()
        identified[hit["evidence_id"]] = hit
    all_hits = sorted(identified.values(), key=lambda h: (h["company_a"], h["company_b"], h["dimension"], h["evidence_id"]))
    summary = aggregate(all_hits, target_set)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    scope = {
        "核查基准日": check_date.isoformat(),
        "原始资料公司数量": len(company_names),
        "是否提供用户自报名单": "是" if disclosed_parties is not None else "否",
    }
    scope.update(scope_metadata or {})
    scope["实际尝试比对公司对数"] = len(companies) * (len(companies) - 1) // 2
    scope["原始资料目录"] = "\n".join(os.path.relpath(path, output.parent) for path in data_paths)
    scope["自动结论含义"] = "仅为线索初判；未命中不代表不存在关联关系；其他核查对象之间的命中不推定与被审计单位关联"
    scope["规则范围"] = "工商资料七类线索筛查；同名需核验身份。多数持股路径最多五层，每段均超过50%才自动串联；不把少数股权或共同合营关系直接传递为控制。"
    scope["需要补充的业务资料"] = "亲属关系、表决权与一致行动协议、集团联营合营安排、交易实质及管理层披露，需要结合相应资料人工核实。"
    records = list(object_records or [])
    snapshot = data_path / "候选发现记录.json"
    if not records and snapshot.is_file():
        try:
            saved = json.loads(snapshot.read_text(encoding="utf-8-sig"))
            if not isinstance(saved, dict) or not isinstance(saved.get("warnings", []), list):
                raise ValueError("候选发现记录格式无效")
            records = saved["candidates"]
            if not isinstance(records, list) or any(not isinstance(item, dict) or not isinstance(item.get("name"), str) or not item["name"].strip() for item in records):
                raise ValueError("候选发现记录缺少有效公司名称")
            for item in records:
                for key in ("reasons", "paths", "notes", "source_files"):
                    if not isinstance(item.get(key, []), list) or any(not isinstance(value, str) for value in item.get(key, [])):
                        raise ValueError(f"候选发现记录的{key}应为文本列表")
                if not isinstance(item.get("sources", []), list) or any(not isinstance(source, dict) for source in item.get("sources", [])):
                    raise ValueError("候选来源位置格式无效")
            discovery_warnings = list(discovery_warnings or []) + saved.get("warnings", [])
        except (ValueError, KeyError, OSError) as exc:
            errors.append({"category": "候选发现记录读取失败", "source": str(snapshot), "message": str(exc)})
            records = []
    if not records:
        # 离线和直接核查同样保留公开交易对手来源，但不据此生成风险命中。
        from scripts.related_party_workflow import extract_seed_export_candidates
        for folder in dict.fromkeys(data_paths):
            records.extend({"name": item.name, "relation_type": item.relation_type,
                            "reasons": [item.relation_type], "paths": [], "notes": [], "sources": list(item.sources)}
                           for item in extract_seed_export_candidates(folder))
    for warning in discovery_warnings or []:
        limitations.append({"category": "候选发现不完整", "source": "主动发现", "message": warning})
    write_report(
        output,
        summary,
        all_hits,
        companies,
        dim,
        target_display,
        data_completeness,
        errors=errors,
        limitations=limitations,
        task_mode=task_mode,
        scope_metadata=scope,
        file_info=file_info,
        object_records=records,
        target_set=target_set,
    )
    return CheckResult(
        output_path=output,
        summary=summary,
        hits=all_hits,
        errors=errors,
        limitations=limitations,
        companies=companies,
    )


def main(argv=None):
    parser = argparse.ArgumentParser(description="关联方核查引擎")
    parser.add_argument("--data-dir", required=True, help="注协完整导出解压后的 _files 目录")
    parser.add_argument("--target", required=True, help="被审计单位全称，多个名称用逗号分隔")
    parser.add_argument("-o", "--output", default="关联方核查报告.xlsx", help="输出 Excel 路径")
    parser.add_argument("--as-of-date", help="核查基准日，格式 YYYY-MM-DD")
    args = parser.parse_args(argv)
    try:
        run_check(
            data_dir=args.data_dir,
            target_names=args.target,
            output_path=args.output,
            as_of_date=date.fromisoformat(args.as_of_date) if args.as_of_date else None,
        )
    except Exception as exc:
        print(f"❌ 核查失败: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
