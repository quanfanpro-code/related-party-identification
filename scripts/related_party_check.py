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
from collections import Counter, defaultdict, deque
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
# 居民楼特征只认"小区/花园/公寓/苑"和带门牌的"×村N组/N号"，
# 裸"村"会把中关村、亚运村这类正常地名误判成居民楼。
RESIDENTIAL_ADDRESS = re.compile(r"小区|花园|公寓|苑|村\s*\d+\s*(?:组|队|排|栋|幢|号)")

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

# rule1 对 ≥5 家共用指纹的降级注记（"该电话/地址/邮箱/域名为N家共用"）。
# 汇总层据此识别"全部命中都来自高共用度指纹"的公司对并收敛，修改注记措辞时必须同步修改本正则。
SHARED_FINGERPRINT_NOTE = re.compile(r"该(?:电话|地址|邮箱|域名)为\d+家共用")

# 投资关系候选阈值：披露比例均不低于该值才列高优先级；全部低于该值且没有其他红旗时降为轻微。
EQUITY_RATIO_THRESHOLD = 20

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


def build_dim_index(data_dir, errors=None, limitations=None, file_info=None, dimension_scan=None):
    """加载所有维度文件，返回 {key: {公司名: [行...]}} 索引。

    dimension_scan 传入时，逐维度登记"是否在任一目录出现 + 最后看到的完成状态"，
    由调用方在读完所有目录后统一判缺失（同名维度任一目录提供即不算缺失，且只报一次）；
    不传则沿用旧的即时追加 limitations 行为。
    """
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
        if fname in codes:
            exists = os.path.exists(path)
            if dimension_scan is not None:
                entry = dimension_scan.setdefault(fname, {"present": False, "terminal": terminal})
                entry["present"] = entry["present"] or exists
                entry["terminal"] = terminal
            elif not exists and terminal != "no_data" and limitations is not None:
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


def phone_suffix_match(phones_a, phones_b):
    """一方登记完整座机（含区号）、另一方只写 7-8 位本地号，且本地号是完整号码的尾段。

    工商登记里"一方带区号、一方不带"是常见写法差异，只列待核实线索，不判相同。
    完整号码必须是以 0 开头的座机：手机号尾 8 位与本地号同形，不参与尾段比对，避免误报。
    返回 (完整号码, 本地号, 完整号码是否属于甲方) 或 None。
    """
    completes_a = {p for p in phones_a if p.startswith("0") and len(p) >= 10}
    locals_a = {p for p in phones_a if not p.startswith("0") and 7 <= len(p) <= 8}
    completes_b = {p for p in phones_b if p.startswith("0") and len(p) >= 10}
    locals_b = {p for p in phones_b if not p.startswith("0") and 7 <= len(p) <= 8}
    for full in sorted(completes_a):
        for local in sorted(locals_b):
            if full.endswith(local):
                return (full, local, True)
    for full in sorted(completes_b):
        for local in sorted(locals_a):
            if full.endswith(local):
                return (full, local, False)
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

def rule1_fingerprint(ca: Company, cb: Company, scarcity=None):
    """维度1: 工商指纹重合（地址/电话/邮箱/网址）。

    scarcity 为全样本共用计数 {"phone"/"email"/"address"/"email_domain": Counter}；
    同一指纹被 5 家及以上共用时视为代理记账/集中注册特征，降为中级并注明。
    """
    hits = []
    counters = scarcity or {}
    phone_count = counters.get("phone", {})
    address_count = counters.get("address", {})
    domain_count = counters.get("email_domain", {})
    email_count = counters.get("email", {})
    # 电话
    common_phones = ca.phones & cb.phones
    if common_phones:
        complete = any(len(number) >= 10 for number in common_phones)
        shared = max((phone_count.get(number, 0) for number in common_phones), default=0)
        notes = []
        if not complete:
            notes.append("缺区号，地域待核实")
        if shared >= 5:
            notes.append(f"该电话为{shared}家共用，疑似代理记账或集中注册")
        level = HARD if (complete and shared < 5) else MEDIUM
        suffix = f"（{'；'.join(notes)}）" if notes else ""
        hits.append(("phone", level, f"联系电话相同: {','.join(sorted(common_phones))}" + suffix, "蓝山/卓朗/达志科技案",
                     trace(ca, "phone", common_phones, normalize=normalize_phones) + trace(cb, "phone", common_phones, normalize=normalize_phones)))
    else:
        # 一方带区号、一方只写本地号：尾段一致列待核实线索（工商登记常见写法差异）
        suffix = phone_suffix_match(ca.phones, cb.phones)
        if suffix:
            full, local, full_from_a = suffix
            shared = phone_count.get(local, 0)
            notes = ["一方缺区号，地域待核实"]
            if shared >= 5:
                notes.append(f"该电话为{shared}家共用，疑似代理记账或集中注册")
            if full_from_a:
                sources = trace(ca, "phone", [full], normalize=normalize_phones) + trace(cb, "phone", [local], normalize=normalize_phones)
            else:
                sources = trace(ca, "phone", [local], normalize=normalize_phones) + trace(cb, "phone", [full], normalize=normalize_phones)
            hits.append(("phone", MEDIUM, f"联系电话尾段一致: {full}↔{local}（{'；'.join(notes)}）", "蓝山/卓朗/达志科技案", sources))
        else:
            # 电话号段相邻（座机同一区号+局向，末位不同 → 同一办公地点的连续号码）
            seg = phone_segment_adjacent(ca.phones, cb.phones)
            if seg:
                hits.append(("phone_segment", MEDIUM, f"座机号段相邻: {seg[0]}↔{seg[1]}（待核实，不能据此认定同一地点）", "蓝山科技案(号段相邻)",
                             trace(ca, "phone", [seg[0]], normalize=normalize_phones) + trace(cb, "phone", [seg[1]], normalize=normalize_phones)))
    # 邮箱完全相同（同一邮箱被 5 家及以上共用时同样视为代理记账特征）
    common_emails = set(ca.emails) & set(cb.emails)
    if common_emails:
        shared = max((email_count.get(address, 0) for address in common_emails), default=0)
        note = f"（该邮箱为{shared}家共用，疑似代理记账或集中注册）" if shared >= 5 else ""
        level = MEDIUM if shared >= 5 else HARD
        hits.append(("email", level, f"邮箱完全相同: {','.join(sorted(common_emails))}" + note, "爱康科技案",
                     trace(ca, "email", common_emails, normalize=lambda x: normalize_emails(x)[0]) + trace(cb, "email", common_emails, normalize=lambda x: normalize_emails(x)[0])))
    # 邮箱同域名（非公共）
    common_domains = set(ca.email_domains) & set(cb.email_domains)
    if common_domains:
        shared = max((domain_count.get(domain, 0) for domain in common_domains), default=0)
        if shared >= 5:
            hits.append(("email_domain", MEDIUM, f"企业邮箱同域名: {','.join(sorted(common_domains))}（该域名为{shared}家共用，疑似代理记账或集中注册）", "天沃/爱康/志高机械案",
                         trace(ca, "email", common_domains, normalize=lambda x: normalize_emails(x)[1]) + trace(cb, "email", common_domains, normalize=lambda x: normalize_emails(x)[1])))
        else:
            hits.append(("email_domain", HARD, f"企业邮箱同域名: {','.join(sorted(common_domains))}", "天沃/爱康/志高机械案",
                         trace(ca, "email", common_domains, normalize=lambda x: normalize_emails(x)[1]) + trace(cb, "email", common_domains, normalize=lambda x: normalize_emails(x)[1])))
    # 地址
    for a1 in all_addresses(ca):
        for a2 in all_addresses(cb):
            rel = address_relationship(a1, a2)
            if rel:
                kind, desc = rel
                level = MEDIUM
                if kind == "exact":
                    shared = address_count.get(normalize_address(a1), 0)
                    if shared >= 5:
                        desc += f"；该地址为{shared}家共用，疑似集中注册或商务秘书地址"
                    else:
                        level = HARD
                case = "天沃科技案(同楼同座)" if kind != "exact" else "达志科技案(地址相同)"
                hits.append(("address", level, f"{desc}；{ca.name}地址：{a1}；{cb.name}地址：{a2}", case,
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


def rule3_counterparty_profile(ca: Company, cb: Company, target_set, as_of_date=None, counterparties=None):
    """维度3: 客商异常画像（对非审计对象群的公司做单方向画像）。

    审计对象群 = 被审计单位 + 其重要子公司（--target 逗号分隔）。
    只对「审计对象 ↔ 对手方」做画像；双方都是/都不是审计对象则跳过。
    counterparties 传入审计对象群已知的交易对手方名称时，只有确实出现在客商名单里的公司才做画像；
    传 None 表示本次没有客商资料，维持原行为，避免静默丢掉线索。
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
    if counterparties is not None:
        party_names = {normalize_name(n) for n in {party.name} | set(party.former_names)}
        if not party_names & counterparties:
            return hits  # 名单里没有这家，不是客商，谈不上交易真实性画像
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
    # 地址居民楼（认"×村N组/N号"这类门牌，不认中关村、亚运村等正常地名）
    for a in party.addresses:
        if RESIDENTIAL_ADDRESS.search(a):
            flags.append("注册地址疑似居民楼")
            break
    # 缺失经营范围列为数据缺口，不构造风险命中。
    if party.business_scope and len(party.business_scope) < 10:
        flags.append("经营范围极简")
    if flags:
        hits.append(("profile", MEDIUM if len(flags) >= 2 else LOW,
                     f"对手方{party.name}自身特征异常：{'；'.join(flags)}"
                     f"（依据对手方自身公开资料，与被审计单位是否存在真实交易需核实）",
                     "专网通信/爱康案", trace(party, "profile")))
    return hits


def parse_percentage(value):
    """导出比例按百分数读取；不把未知比例当作零，也不猜小数的单位。"""
    text = strip_html(value).replace("%", "").replace("％", "").strip()
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    number = float(text)
    return number if 0 <= number <= 100 else None


def rule5_equity(ca: Company, cb: Company):
    """识别双向直接投资和共同股东；持股证据与最终控制认定分开。曾用名视同现名接通。"""
    hits = []
    for parent, child in ((ca, cb), (cb, ca)):
        child_names = {normalize_name(n) for n in ({child.name} | set(child.former_names))}
        parent_names = {normalize_name(n) for n in ({parent.name} | set(parent.former_names))}
        invest_names = [name for name, _ in parent.investments + parent.holdings if normalize_name(name) in child_names]
        sh_names = [name for name, _, _ in child.shareholders if normalize_name(name) in parent_names]
        ratios = [ratio for name, ratio in parent.investments + parent.holdings if normalize_name(name) in child_names]
        ratios += [ratio for name, ratio, _ in child.shareholders if normalize_name(name) in parent_names]
        if ratios:
            parsed = [parse_percentage(ratio) for ratio in ratios]
            level = HARD if all(ratio is not None and ratio >= EQUITY_RATIO_THRESHOLD for ratio in parsed) else MEDIUM
            hits.append(("invest", level, f"{parent.name} 对 {child.name} 的持股线索；记录比例：{'、'.join(sorted(set(str(r) or '未披露' for r in ratios)))}（控制或重大影响待核实）", "投资关系",
                trace(parent, "invest", invest_names, normalize=normalize_name) + trace(child, "shareholder", sh_names, normalize=normalize_name)
                + trace(parent, "names", [parent.name], normalize=normalize_name) + trace(child, "names", [child.name], normalize=normalize_name)))
    common_sh = {name for name in {s[0] for s in ca.shareholders} & {s[0] for s in cb.shareholders}
                 if strip_html(name) and name not in {"国有企业", "自然人", "法人"} and not is_public_authority(name)}
    if common_sh:
        hits.append(("common_shareholder", MEDIUM,
            f"共同股东同名: {','.join(sorted(common_sh))}（结合持股比例及其他安排核实影响）", "股权穿透",
            trace(ca, "shareholder", common_sh, normalize=normalize_name) + trace(cb, "shareholder", common_sh, normalize=normalize_name)))
    return hits


def low_ratio_only(ca: Company, cb: Company):
    """双向已披露持股比例全部低于候选阈值时返回 True。

    只按已披露比例判断：比例未知（None）不当作低比例，不据此降级。
    """
    ratios = []
    for parent, child in ((ca, cb), (cb, ca)):
        child_names = {normalize_name(n) for n in ({child.name} | set(child.former_names))}
        parent_names = {normalize_name(n) for n in ({parent.name} | set(parent.former_names))}
        ratios += [parse_percentage(ratio) for name, ratio in parent.investments + parent.holdings if normalize_name(name) in child_names]
        ratios += [parse_percentage(ratio) for name, ratio, _ in child.shareholders if normalize_name(name) in parent_names]
    known = [ratio for ratio in ratios if ratio is not None]
    return bool(known) and all(ratio < EQUITY_RATIO_THRESHOLD for ratio in known)


def rule6_historical(ca: Company, cb: Company):
    """维度6: 历史关联痕迹。"""
    hits = []
    split_names = lambda x: [normalize_name(v) for v in re.split(r"[,，;；、]", strip_html(x))]
    # 曾用名命中
    cb_names = {cb.name} | set(cb.former_names)
    ca_names = {ca.name} | set(ca.former_names)
    former_hit = (set(ca.former_names) & cb_names) | (set(cb.former_names) & ca_names)
    if former_hit:
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
    # 历史股东痕迹：变更记录"投资人/股东"项的变更前文本出现对方现名或曾用名
    for _, item, before, _after in ca.changes:
        if not before or ("投资人" not in item and "股东" not in item):
            continue
        before_compact = re.sub(r"\s+", "", before)
        for name in sorted({cb.name} | set(cb.former_names)):
            normalized = normalize_name(name)
            if len(normalized) < 4 or normalized not in before_compact:
                continue
            hits.append(("past_investor", MEDIUM,
                         f"{ca.name}历史股东变更记录中出现{name}", "非关联化线索（ST新亿/扬子新材案）",
                         trace(ca, "past_investor", None,
                               row_match=lambda row, n=normalized: n in re.sub(r"\s+", "", str(row_value(row, 4, ""))))
                         + trace(cb, "names", [normalized], normalize=split_names)))
            break
    return hits


def equity_path_hits(companies, target_set):
    """在已取得资料内寻找最多五层的多数持股路径，不传递少数股权。曾用名归一到现名建边。"""
    alias = {}
    for name, company in companies.items():
        alias[normalize_name(name)] = name
        for former in company.former_names:
            alias.setdefault(normalize_name(former), name)
    records = defaultdict(list)
    for company in companies.values():
        for child, ratio in company.investments + company.holdings:
            canonical_child = alias.get(normalize_name(child), child)
            records[(company.name, canonical_child)].append((parse_percentage(ratio), trace(company, "invest", [child], normalize=normalize_name)))
        for parent, ratio, _ in company.shareholders:
            canonical_parent = alias.get(normalize_name(parent), parent)
            records[(canonical_parent, company.name)].append((parse_percentage(ratio), trace(company, "shareholder", [parent], normalize=normalize_name)))
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
        # 与主比对同一口径：只保留被审计单位↔各公司的对子。
        if (a in target_set) == (b in target_set):
            continue
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

def rule9_dual_role(ca, cb, dim):
    """维度9: 既客又供 — 同一对手方同时出现在客户表和供应商表（交易真实性线索）。"""
    hits = []
    for owner, other in ((ca, cb), (cb, ca)):
        other_names = {normalize_name(n) for n in ({other.name} | set(other.former_names))}
        customers = {strip_html(str(r[7])) for r in dim.get("customer", {}).get(owner.name, [])
                     if len(r) > 7 and strip_html(str(r[7] or ""))}
        suppliers = {strip_html(str(r[7])) for r in dim.get("supplier", {}).get(owner.name, [])
                     if len(r) > 7 and strip_html(str(r[7] or ""))}
        for name in sorted(customers & suppliers):
            if normalize_name(name) not in other_names:
                continue
            hits.append(("dual_role", MEDIUM,
                         f"{owner.name}的客户与供应商名单同时出现{name}（既客又供）", "交易真实性线索（资金/货物双向流动）",
                         trace(owner, "dual_role", [name], normalize=strip_html)
                         + trace(other, "names", [normalize_name(name)], normalize=normalize_name)))
    return hits


def compare_pair(ca, cb, dim, target_set, as_of_date=None, errors=None, scarcity=None, counterparties=None):
    """对一对公司跑全部规则，返回 [hit_dict...]。target_set 为审计对象群（含子公司）。"""
    results = []
    # rule3 需要方向参数、rule1 需要全样本稀缺性计数、rule9 需要客户/供应商明细，均单独调用；其余规则统一双参数
    rules_2arg = [
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
            counterparties=counterparties,
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
    # rule1: 工商指纹重合（需要全样本稀缺性计数，共用度高的指纹降级）
    try:
        for field_key, level, evidence, case, sources in rule1_fingerprint(ca, cb, scarcity):
            results.append({
                "company_a": ca.name, "company_b": cb.name,
                "dimension": "工商指纹重合", "field": field_key,
                "level": level, "evidence": evidence, "case_ref": case, "sources": unique_sources(sources),
            })
    except Exception as e:
        print(f"  ⚠️ 规则异常 工商指纹重合 {ca.name}↔{cb.name}: {e}", file=sys.stderr)
        if errors is not None:
            errors.append({
                "category": "规则异常",
                "source": "工商指纹重合",
                "message": f"{ca.name}↔{cb.name}: {e}",
            })
    # rule9: 既客又供（需要 dim 里的客户/供应商明细）
    try:
        for field_key, level, evidence, case, sources in rule9_dual_role(ca, cb, dim):
            results.append({
                "company_a": ca.name, "company_b": cb.name,
                "dimension": "客商异常画像", "field": field_key,
                "level": level, "evidence": evidence, "case_ref": case, "sources": unique_sources(sources),
            })
    except Exception as e:
        print(f"  ⚠️ 规则异常 既客又供 {ca.name}↔{cb.name}: {e}", file=sys.stderr)
        if errors is not None:
            errors.append({
                "category": "规则异常",
                "source": "既客又供",
                "message": f"{ca.name}↔{cb.name}: {e}",
            })
    # 低比例持股按文档口径收紧：已披露比例全部低于候选阈值的投资线索，
    # 只有同时命中共同关键人员等其他中/高红旗时才保留中级；单独出现降为轻微。
    other_flags = [h for h in results if h["field"] != "invest" and level_rank(h["level"]) >= level_rank(MEDIUM)]
    if not other_flags and low_ratio_only(ca, cb):
        for h in results:
            if h["field"] == "invest" and h["level"] == MEDIUM:
                h["level"] = LOW
                h["evidence"] += "；已披露比例均低于20%，单独出现降为轻微（与共同关键人员等其他红旗同时出现时保留为可疑）"
    return results


def level_rank(level):
    return {HARD: 3, MEDIUM: 2, LOW: 1}.get(level, 0)


# 命中类别 → 报告里"与被审计单位的疑似关系"一列直接写给读者看的白话。
# 比对只发生在被审计单位与每一家公司之间，所以每类命中都能说成一种疑似关系。
DIMENSION_PHRASES = {
    "股权控制穿透": "存在持股或被持股关系",
    "关键人员重合": "关键人员与被审计单位重合",
    "历史关联痕迹": "与被审计单位有历史股权或人员痕迹",
    "担保资金链": "与被审计单位存在资产抵质押等资金往来",
    "无形资产共用": "与被审计单位使用相同品牌或无形资产",
    "客商异常画像": "是被审计单位的客户或供应商，且自身特征异常",
    "工商指纹重合": "注册地址、电话或邮箱与被审计单位相同",
}
# 个别命中按具体内容给更准的说法，覆盖上面的类别级表述。
FIELD_PHRASES = {
    "dual_role": "既是被审计单位的客户又是其供应商",
    "equity_path": "通过多层持股与被审计单位形成控制路径",
    "phone_segment": "座机号段与被审计单位相邻",
}


def relation_phrases(hits):
    """把一家公司的全部命中翻成"与被审计单位的疑似关系"，风险高的说法排前面。"""
    phrases = {}
    for h in hits:
        phrase = FIELD_PHRASES.get(h["field"]) or DIMENSION_PHRASES.get(h["dimension"], h["dimension"])
        phrases[phrase] = max(phrases.get(phrase, 0), level_rank(h["level"]))
    ranked = sorted(phrases.items(), key=lambda pair: (-pair[1], pair[0]))
    return "；".join(phrase for phrase, _rank in ranked)


def aggregate(all_hits, target_set):
    """按公司汇总命中。比对只含被审计单位↔各公司，故每个 pair 必有一方是被审计单位。"""
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
        corroborated = False
        if has_hard:
            is_related = "高风险线索，待核实"
            suggestion = "实施函证、实地走访、资金流水核对；询问管理层，核实关系并评估披露"
        elif has_medium:
            is_related = "中风险线索，待核实"
            # 三条以上中级线索且横跨三个以上不同维度 → 独立线索相互印证，建议程序升级为高风险档（结论仍保持中级）
            medium_hits = [h for h in hits if h["level"] == MEDIUM]
            corroborated = len(medium_hits) >= 3 and len({h["dimension"] for h in medium_hits}) >= 3
            if corroborated:
                suggestion = "实施函证、实地走访、资金流水核对；询问管理层，核实关系并评估披露"
            else:
                suggestion = "结合交易背景进一步核查；关注交易商业合理性"
        else:
            corroborated = False
            is_related = "低风险线索，待核实"
            suggestion = "记录备查，必要时跟进"
        ca, cb = pair
        # 比对只保留被审计单位↔各公司，pair 里必有一方是被审计单位，另一方就是要报告的公司。
        company = next((name for name in pair if name not in target_set), cb)
        # 全部命中都来自 ≥5 家共用指纹（疑似集中注册/代理记账）的公司对：审计价值低，
        # 收敛处理——概览折叠为一行、汇总排在后段；命中本身保留，等级与证据不变。
        converged = all(SHARED_FINGERPRINT_NOTE.search(h["evidence"]) for h in hits)
        summary.append({
            "company_a": ca, "company_b": cb, "company": company,
            "relation_to_target": relation_phrases(hits),
            "is_related": is_related, "max_level": max_level,
            "hit_count": len(hits), "dimensions": "、".join(dims),
            "evidence": ("多条独立线索相互印证；" if corroborated else "") + " | ".join(sorted({h["evidence"] for h in hits})),
            # evidence_items 供报告按整条取舍：同一级别里短的先列，摘要不把一条证据从中间切断。
            "evidence_items": [text for text, _rank in sorted(
                {h["evidence"]: level_rank(h["level"]) for h in hits}.items(),
                key=lambda pair: (-pair[1], len(pair[0]), pair[0]))],
            "corroborated": corroborated,
            "converged": converged,
            "suggestion": suggestion,
            "case_ref": "、".join(sorted({h["case_ref"] for h in hits})),
            "hits": hits,
        })
    # 收敛对排在后段，其余按风险与命中条数排。
    summary.sort(key=lambda x: (
        1 if x["converged"] else 0,
        -level_rank(x["max_level"]), -x["hit_count"]))
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
    channel_warnings=None,
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
    dimension_scan = {}
    data_paths = [data_path] + [Path(path) for path in (additional_data_dirs or [])]
    # 离线重读同一交付目录时，也纳入已经完成的候选批次。
    candidate_dir = data_path / "候选公司原始导出"
    if candidate_dir.is_dir() and candidate_dir not in data_paths:
        data_paths.append(candidate_dir)
    # 逐家取数保底文件固定在"逐家取数补采"子目录，同样自动纳入比对。
    fallback_dir = data_path / "逐家取数补采"
    if fallback_dir.is_dir() and fallback_dir not in data_paths:
        data_paths.append(fallback_dir)
    for folder in dict.fromkeys(data_paths):
        if not folder.is_dir():
            raise FileNotFoundError(f"数据目录不存在: {folder}")
        part = build_dim_index(folder, errors=errors, limitations=limitations, file_info=file_info,
                               dimension_scan=dimension_scan)
        for key, by_name in part.items():
            for name, rows in by_name.items():
                dim.setdefault(key, {}).setdefault(name, []).extend(rows)
    # 全部目录读完后统一报缺失（任一目录提供该维度即不算缺失，同名只报一次）。
    for fname, scanned in dimension_scan.items():
        if scanned["present"] or scanned["terminal"] == "no_data":
            continue
        limitations.append({
            "category": "缺失维度",
            "source": fname,
            "message": "本次导出未提供该维度，相关规则可能受限",
        })
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

    # 全样本指纹稀缺性计数：同一指纹被多家公司共用时降级（代理记账/集中注册特征）
    scarcity = {"phone": Counter(), "email": Counter(), "address": Counter(), "email_domain": Counter()}
    for company in companies.values():
        for number in company.phones:
            scarcity["phone"][number] += 1
        for address in set(company.emails):
            scarcity["email"][address] += 1
        for domain in set(company.email_domains):
            scarcity["email_domain"][domain] += 1
        for addr in {normalize_address(a) for a in all_addresses(company)} - {""}:
            scarcity["address"][addr] += 1

    # 客商名单：有客商资料时，画像只对真的出现在名单里的对手方做，
    # 否则同一份导出里的任何公司都会被配成"客商"出现在重点线索里。
    counterparty_names = set()
    counterparty_rows = 0
    for key in ("customer", "supplier"):
        for target in target_set:
            for row in dim.get(key, {}).get(target, []):
                counterparty_rows += 1
                value = strip_html(str(row_value(row, 7, "") or ""))
                if value:
                    counterparty_names.add(normalize_name(value))
    counterparties = counterparty_names if counterparty_rows else None

    all_hits = []
    # 只比"被审计单位 ↔ 每一家公司"。两家都与被审计单位无关的公司之间即使互相撞上，
    # 也回答不了"是不是/有哪些关联方"，属于本技能用途之外的内容，不比也不报。
    target_names_in_data = sorted(target_set & set(companies))
    other_names = sorted(set(companies) - target_set)
    compared_pairs = [(a, b) for a in target_names_in_data for b in other_names]
    for company_a, company_b in compared_pairs:
        all_hits.extend(
            compare_pair(
                companies[company_a],
                companies[company_b],
                dim,
                target_set,
                as_of_date=check_date,
                errors=errors,
                scarcity=scarcity,
                counterparties=counterparties,
            )
        )
    all_hits.extend(equity_path_hits(companies, target_set))
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
    scope["被审计单位与各公司的实际比对数"] = len(compared_pairs)
    scope["原始资料目录"] = "\n".join(os.path.relpath(path, output.parent) for path in data_paths)
    scope["本次核查结论（必读）"] = "本报告只是自动初判线索，不是审计结论：列出的疑似关联方需经审计程序核实；未发现不代表不存在关联关系，关联方是否完整须结合其他审计程序确认"
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
        extract_warnings = []
        for folder in dict.fromkeys(data_paths):
            records.extend({"name": item.name, "relation_type": item.relation_type,
                            "reasons": [item.relation_type], "paths": [], "notes": [], "sources": list(item.sources)}
                           for item in extract_seed_export_candidates(folder, warnings=extract_warnings))
        for warning in extract_warnings:
            limitations.append({"category": "候选发现不完整", "source": "客户/供应商表", "message": warning})
    for warning in channel_warnings or []:
        limitations.append({"category": "取数通道降级", "source": "逐家取数保底", "message": warning})
    for warning in discovery_warnings or []:
        limitations.append({"category": "候选发现不完整", "source": "主动发现", "message": warning})
    write_report(
        output,
        summary,
        all_hits,
        companies,
        target_display,
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
