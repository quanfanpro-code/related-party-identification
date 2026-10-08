# 七张工作表共享同一核查结果，使用 Excel 原生导航、筛选和相对文件链接。
from collections import Counter
from datetime import datetime
import hashlib
import math
import os
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.workbook.defined_name import DefinedName


SHEETS = {
    "目录": "先看概览，再按公司关系查看证据；所有原始资料保持原样。",
    "核查概览": "先看结论与重点线索，再查范围、口径和数据缺口。",
    "核查对象与来源": "每家公司为何纳入、取得什么资料、实际查到哪一步。",
    "关系核查汇总": "每对有命中的公司一行；自动初判与人工复核意见分开。",
    "证据明细": "每条证据的判断依据及实际参与匹配的原始字段，可按证据编号筛选。",
    "数据覆盖与缺口": "每家公司、每个维度的取得情况及其对核查的影响。",
    "任务说明": "任务入口、范围、日期、原始目录和判断口径。",
}
NAVY = "1F4E79"
PALE = "D6E4F0"
LIGHT = "F3F7FB"
RISK_COLORS = {"高": "FCE4D6", "中": "FFF2CC", "低": "E2F0D9"}
# 概览与汇总的"主要线索摘要"长度上限：按整条证据取舍，不切断单条证据。
SUMMARY_LIMIT = 240
DIMENSIONS = {
    "basic": ("基础工商信息", "工商指纹、人员、客商画像、历史痕迹"),
    "shareholder": ("股东信息", "人员、共同股东"),
    "shareholder_new": ("最新公示股东", "人员、共同股东"),
    "actual_controller": ("实际控制人", "共同实际控制人"),
    "ultimate_beneficiary": ("最终受益人", "共同最终受益人"),
    "main_persons": ("主要人员（高管）", "关键人员"),
    "core_team": ("核心团队", "关键人员"),
    "invest_new": ("对外投资（新）", "直接投资"),
    "holding": ("参控股企业", "参控股关系"),
    "invoice": ("发票信息", "地址补充"),
    "customer": ("客户", "仅用于候选来源"),
    "supplier": ("供应商", "仅用于候选来源"),
    "change": ("变更记录", "历史地址"),
    "legal_change": ("法定代表人变更", "历史法人"),
    "abnormal": ("经营异常", "背景资料，本次规则不直接使用"),
    "pledge": ("股权质押", "质押关系"),
    "mortgage": ("动产抵押", "抵押关系"),
    "trademark": ("商标", "同名商标"),
    "software": ("软件著作权", "同名软件著作权"),
    "wechat": ("微信公众号", "同名公众号"),
}


def risk_text(value):
    text = str(value)
    if "硬" in text or text.startswith("高"):
        return "高"
    if "可疑" in text or text.startswith("中"):
        return "中"
    return "低"


def pair_key(a, b):
    return tuple(sorted((a, b)))


def pair_id(a, b):
    return "GX-" + hashlib.sha256("\n".join(pair_key(a, b)).encode("utf-8")).hexdigest()[:8].upper()


def put(ws, row, column, value):
    cell = ws.cell(row, column)
    cell.value = value
    if isinstance(value, str):
        # 原始企业名称和字段是数据，不能被 Excel 当成公式执行。
        cell.data_type = "s"
    return cell


def append_row(ws, values):
    """超过 Excel 单元格上限时续行保留全文，编号和短字段在续行重复。"""
    start = max(ws.max_row + 1, 4)
    parts = max([1] + [math.ceil(len(v) / 32767) for v in values if isinstance(v, str)])
    for offset in range(parts):
        for column, value in enumerate(values, 1):
            part = value[offset * 32767:(offset + 1) * 32767] if isinstance(value, str) and len(value) > 32767 else value
            put(ws, start + offset, column, part)
    return start


def link(cell, sheet, row):
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet.replace(chr(39), chr(39)*2)}'!A{row}", display=str(cell.value))
    cell.font = Font(name="微软雅黑", size=10, color=NAVY, underline="single")


def source_link(cell, source, output):
    relative = os.path.relpath(source["file"], Path(output).parent).replace("\\", "/")
    cell.hyperlink = Hyperlink(ref=cell.coordinate, target=relative,
        location=f"'{source['sheet'].replace(chr(39), chr(39)*2)}'!{source['cell']}", display=str(cell.value))
    cell.font = Font(name="微软雅黑", size=10, color=NAVY, underline="single")


def keyed_link(cell, sheet, key):
    """按唯一编号定位，明细或汇总排序以后仍指向同一条记录。"""
    wb = cell.parent.parent
    name = "_位置_" + key.replace("-", "_")
    quoted = "'" + sheet.replace("'", "''") + "'"
    if name not in wb.defined_names:
        wb.defined_names.add(DefinedName(name, attr_text=f'INDEX({quoted}!$A:$A,MATCH("{key}",{quoted}!$A:$A,0))'))
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=name, display=str(cell.value))
    cell.font = Font(name="微软雅黑", size=10, color=NAVY, underline="single")


def create_sheet(wb, title, headers, widths):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = NAVY
    put(ws, 1, 1, "返回目录" if title != "目录" else "关联方识别与核查")
    if title != "目录":
        link(ws["A1"], "目录", 1)
    else:
        ws["A1"].font = Font(name="微软雅黑", size=14, bold=True, color=NAVY)
    for column in range(1, len(headers) + 1):
        ws.cell(1, column).fill = PatternFill("solid", fgColor=PALE)
    if len(headers) > 1:
        put(ws, 1, 2, title)
        ws["B1"].font = Font(name="微软雅黑", size=16, bold=True, color=NAVY)
    put(ws, 2, 1, SHEETS[title])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws["A2"].font = Font(name="微软雅黑", size=10, color="526579")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30
    for column, header in enumerate(headers, 1):
        cell = put(ws, 3, column, header)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="微软雅黑", size=10, color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(column)].width = widths[column - 1]
    ws.freeze_panes = "C4" if len(headers) > 3 else "A4"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if len(headers) > 8 else ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    return ws


def finish_sheet(ws, risk_column=None):
    border = Border(bottom=Side(style="hair", color="D9E2F3"))
    for row in ws.iter_rows(min_row=4):
        for cell in row:
            if not cell.hyperlink:
                cell.font = Font(name="微软雅黑", size=10, color="243746")
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
            cell.fill = PatternFill("solid", fgColor=LIGHT if cell.row % 2 == 0 else "FFFFFF")
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00" if isinstance(cell.value, float) else "#,##0"
        if risk_column:
            cell = row[risk_column - 1]
            if cell.value in RISK_COLORS:
                cell.fill = PatternFill("solid", fgColor=RISK_COLORS[cell.value])
                cell.font = Font(name="微软雅黑", size=10, bold=True, color="243746")
        lines = max([1] + [sum(max(1, math.ceil(len(part) * 1.6 / max(10, ws.column_dimensions[cell.column_letter].width)))
                              for part in str(cell.value or "").split("\n")) for cell in row])
        ws.row_dimensions[row[0].row].height = min(210, max(30, lines * 15))
    ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{max(3, ws.max_row)}"
    ws.print_options.horizontalCentered = True
    ws.print_area = f"A1:{get_column_letter(ws.max_column)}{max(4, ws.max_row)}"


def group_evidence_blocks(ws, risk_column=5):
    """同一证据编号的连续多行视为一块：相邻块交替浅底色、块首行加粗上边框，肉眼可分辨块界。"""
    if ws.max_row < 4:
        return
    previous = None
    block = -1
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        if row[0].value != previous:
            previous = row[0].value
            block += 1
            for cell in row:
                cell.border = Border(top=Side(style="medium", color=NAVY), bottom=Side(style="hair", color="D9E2F3"))
        fill = PALE if block % 2 else "FFFFFF"
        for cell in row:
            if cell.column != risk_column:  # 风险列保留红黄绿底色双通道
                cell.fill = PatternFill("solid", fgColor=fill)


def clue_summary(item, limit=SUMMARY_LIMIT):
    """概览与汇总共用的一份摘要：按整条证据取舍，宁少列一条也不切断一条。"""
    items = list(item.get("evidence_items") or []) or [item.get("evidence", "")]
    kept, total = [], 0
    for text in items:
        if kept and total + len(text) + 3 > limit:
            break
        kept.append(text)
        total += len(text) + 3
    omitted = len(items) - len(kept)
    joined = " | ".join(kept)
    if omitted:
        joined += f"…（另 {omitted} 条见明细）"
    return ("多条独立线索相互印证；" if item.get("corroborated") else "") + joined


def coverage_rows(names, companies, file_info, errors):
    rows = []
    incomplete = set()
    gap_dimensions = Counter()
    for name in names:
        for key, (label, impact) in DIMENSIONS.items():
            infos = [item for item in file_info if item["key"] == key]
            own = [item for item in infos if name in item["companies"] or name in item["scope"]]
            record = next((item for item in own if name in item["companies"]), None)
            if any(item.get("malformed") for item in own or infos):
                status = "资料字段不足"
            elif record:
                status = "已取得记录"
            elif any(item["exists"] and not item["readable"] for item in own or infos):
                status = "读取失败"
            elif any(item["terminal"] == "no_data" and name in item["scope"] for item in own):
                status = "明确无数据"
            elif any(item["readable"] and item["terminal"] == "completed" and name in item["scope"] for item in own):
                status = "已核对范围，未见记录"
            elif any(item["readable"] for item in infos):
                status = "未见该企业记录，范围待核实"
            else:
                status = "未取得"
            detail = ""
            if key == "basic" and name in companies:
                company = companies[name]
                fields = {"法定代表人": company.legal_person, "联系电话": company.phones,
                    "邮箱": company.emails, "地址": company.addresses, "注册资本": company.capital,
                    "成立日期": company.found_date, "参保人数": company.insured, "经营范围": company.business_scope}
                blanks = [label for label, value in fields.items() if value is None or value == "" or value == [] or value == set()]
                if blanks:
                    status = "已取得记录，部分字段缺失或不可用"
                    detail = "缺失或不能按本次口径解析的字段：" + "、".join(blanks)
            if status not in {"已取得记录", "明确无数据", "已核对范围，未见记录"} and key not in {"customer", "supplier", "abnormal"}:
                incomplete.add(name)
                gap_dimensions[name] += 1
            chosen = record or next(iter(own or infos), {})
            rows.append([name, label, status, impact, detail, chosen.get("file", ""), chosen.get("created_at", "未记录") or "未记录"])
        if name not in companies:
            incomplete.add(name)
        if any(name in str(item.get("message", "")) or name == item.get("source") for item in errors):
            incomplete.add(name)
    # 重度缺口：未取得基础工商记录，或缺失维度达 10 个及以上；其余缺口公司为轻度。
    severe = {name for name in names if name not in companies or gap_dimensions.get(name, 0) >= 10}
    return rows, incomplete, severe


def write_report(out_path, summary, all_hits, companies, dim, target_display, data_completeness,
                 *, errors=None, limitations=None, task_mode="", scope_metadata=None,
                 file_info=None, object_records=None, target_set=None):
    errors, limitations = list(errors or []), list(limitations or [])
    scope = dict(scope_metadata or {})
    target_set = set(target_set or target_display.split(" / "))
    file_info = list(file_info or [])
    objects = {}
    for record in object_records or []:
        name = record["name"]
        if name not in objects:
            objects[name] = dict(record)
        else:
            for key in ("reasons", "paths", "notes", "source_files"):
                objects[name][key] = list(dict.fromkeys(objects[name].get(key, []) + record.get(key, [])))
            objects[name]["sources"] = objects[name].get("sources", []) + record.get("sources", [])
    names = sorted(set(companies) | target_set | set(objects) | {name for info in file_info for name in info["scope"]}
                   | {name for info in file_info if info["key"] == "basic" for name in info["companies"]})
    coverage, incomplete, severe = coverage_rows(names, companies, file_info, errors)
    counts = Counter(risk_text(item["max_level"]) for item in summary)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = create_sheet(wb, "目录", ["工作表", "阅读用途", "数据行数"], [26, 84, 16])
    for title, description in list(SHEETS.items())[1:]:
        row = append_row(ws, [title, description, None])
        link(ws.cell(row, 1), title, 1)
    overview = create_sheet(wb, "核查概览", ["项目", "数量或情况", "与审计对象的关系", "统计口径 / 下一步"], [32, 48, 33, 76])
    # 结论先行：先给读懂报告所需的前提、降级提示和分级结果，再列重点线索，范围与口径注解殿后。
    append_row(overview, ["自动结论含义",
        scope.get("自动结论含义", "仅为线索初判；未命中不代表不存在关联关系；其他核查对象之间的命中不推定与被审计单位关联"),
        "", "理解本报告全部数字与线索的前提"])
    for item in limitations:
        if item.get("category") == "取数通道降级":
            append_row(overview, ["取数通道降级提示", item.get("message", ""), "",
                f"来源：{item.get('source', '')}；受影响范围详见数据覆盖与缺口"])
    for record in [
        ["高风险公司对数", counts["高"], "", "按该公司对最高风险等级统计，待核实"],
        ["中风险公司对数", counts["中"], "", "与高、低风险公司对互不重复"],
        ["低风险公司对数", counts["低"], "", "低风险线索仍是命中，不表示没有问题"],
    ]:
        append_row(overview, record)
    if not counts["高"] and not counts["中"]:
        append_row(overview, ["重要提示", "本结果不提供关联方完整性保证", "",
            "零命中或全部为低风险线索；未命中不代表不存在关联关系，关联方完整性须结合其他审计程序确认"])
    append_row(overview, ["重点线索", "点击公司对进入汇总", "与审计对象的关系", "下列按风险排序；完整原文见证据明细"])
    overview_links = []
    for item in summary:
        row = append_row(overview, [f"{item['company_a']} ↔ {item['company_b']}", risk_text(item["max_level"]) + "风险线索",
                                   item["relation_type"],
                                   clue_summary(item)])
        overview_links.append((row, pair_key(item["company_a"], item["company_b"])))
    if not summary:
        append_row(overview, ["本次未形成命中", "请结合数据缺口阅读", "", "只说明实际取得资料中的规则结果"])
    append_row(overview, ["范围与口径指标", "以下为范围、进度与统计口径注解", "", ""])
    metrics = [
        ["被审计单位", target_display, "", "自动初判须结合审计程序核实"],
        ["纳入范围公司数", len(names), "", "包括被审计单位、候选及资料未取得的核查对象，按名称去重"],
        ["主动发现候选公司数", len(set(objects) - target_set) if task_mode == "discovery" else 0, "", "只有主动发现模式计入；候选不等于关联方"],
        ["取得基础资料公司数", len(companies), "", "基础工商表可读取并已构建为核查对象"],
        ["实际尝试比对公司对数", scope.get("实际尝试比对公司对数", len(companies) * (len(companies) - 1) // 2), "", "两家公司组成一对，任一规则失败另列数据缺口"],
        ["命中公司对数", len(summary), "", "同一公司对命中多条证据，只计一对"],
        ["证据条数", len({hit.get("evidence_id", str(i)) for i, hit in enumerate(all_hits)}), "", "按证据编号去重；明细的多个来源行不重复计数"],
        ["存在数据缺口公司数", len(incomplete), "", "缺资料、空字段或规则失败；具体影响见数据覆盖与缺口"],
        ["其中重度缺口公司数", len(severe), "", "未取得基础工商记录，或缺失维度达 10 个及以上；缺口实质削弱核查结论"],
        ["其中轻度缺口公司数", len(incomplete - severe), "", "其余存在缺口的公司；个别维度或字段缺失，影响相对有限"],
        ["读取、字段或规则错误条数", len(errors), "", "错误不能当作未命中"],
        ["发现过程提示条数", sum(item.get("category") == "候选发现不完整" for item in limitations), "", "发现过程中取数失败会影响候选范围完整性"],
    ]
    for record in metrics:
        append_row(overview, record)

    roster = create_sheet(wb, "核查对象与来源",
        ["公司名称", "对象来源", "发现理由 / 名单来源", "层级", "直接来源", "关系路径", "资料情况", "比对结果", "法定代表人", "成立日期", "注册资本（人民币元）", "参保人数", "来源记录 / 说明"],
        [30, 21, 42, 9, 28, 54, 23, 28, 17, 15, 20, 12, 52])
    involved = {value for item in summary for value in (item["company_a"], item["company_b"])}
    for name in names:
        record = objects.get(name, {})
        company = companies.get(name)
        origin = "被审计单位" if name in target_set else record.get("relation_type", "已有数据中的企业")
        reason = "；".join(record.get("reasons", []))
        if record.get("source_files"):
            reason += "\n名单文件：" + "；".join(Path(p).name for p in record["source_files"])
        outcome = "未纳入比对，缺少基础资料" if not company else ("未形成公司对" if len(companies) < 2 else ("已尝试比对，发现线索" if name in involved else "已尝试比对，未命中"))
        if any(item.get("category") == "规则异常" and name in item.get("message", "") for item in errors):
            outcome = "部分规则执行失败；" + ("已有命中线索" if name in involved else "其余已执行规则未命中")
        source_notes = []
        for source in record.get("sources", []):
            if source.get("file"):
                source_notes.append(f"{Path(source['file']).name} / {source['sheet']}!{source['cell']}：{source['value']}")
            else:
                ratio_text = "比例未取得" if source.get("ratio") is None else str(source["ratio"]) + "%"
                source_notes.append(f"股权查询：{source.get('company', '')}（{source.get('company_id', '')}） / {source.get('relation_type', '')} / {source.get('counterparty', '')} / {ratio_text}")
        append_row(roster, [name, origin, reason, record.get("depth", ""), record.get("parent_name", ""),
            "\n".join(record.get("paths", [])), "存在数据缺口" if name in incomplete else "已取得本次核查资料", outcome,
            company.legal_person if company else "", company.found_date if company else "", company.capital if company else None,
            company.insured if company else None, "\n".join(source_notes + record.get("notes", []))])

    summary_ws = create_sheet(wb, "关系核查汇总",
        ["公司对编号", "公司A", "公司B", "与审计对象的关系", "风险初判", "主要线索摘要", "命中类别", "证据条数", "查看全部证据", "建议审计程序", "资料限制", "人工复核意见", "复核人 / 日期", "序号"],
        [20, 28, 28, 33, 12, 48, 30, 12, 22, 45, 30, 40, 24, 10])
    summary_rows = {}
    for index, item in enumerate(summary, 1):
        key = pair_key(item["company_a"], item["company_b"])
        # GX- 编号保留作链接锚点；末尾"对-XX"序号供底稿引用和口头沟通。
        row = append_row(summary_ws, [pair_id(*key), item["company_a"], item["company_b"], item["relation_type"],
            risk_text(item["max_level"]), clue_summary(item),
            item["dimensions"], item["hit_count"], "查看全部证据", item["suggestion"],
            "存在数据缺口，详见缺口页" if set(key) & incomplete else "", "", "", f"对-{index:02d}"])
        summary_rows[key] = row
    for row, key in overview_links:
        keyed_link(overview.cell(row, 1), "关系核查汇总", pair_id(*key))

    evidence_ws = create_sheet(wb, "证据明细",
        ["证据编号", "公司A", "公司B", "核查类别", "风险初判", "判断依据", "来源公司", "原始字段", "原始值", "原始文件 / 工作表 / 单元格", "返回汇总", "案例参考（非本次证据）", "序号"],
        [24, 27, 27, 22, 12, 55, 27, 23, 44, 48, 18, 33, 10])
    first_evidence = {}
    evidence_seq = {}
    for number, hit in enumerate(all_hits, 1):
        key = pair_key(hit["company_a"], hit["company_b"])
        evidence_id = hit.get("evidence_id", f"ZJ-{number:05d}")
        if evidence_id not in evidence_seq:
            # 同一证据编号的多行来源共享同一个"证-XX"序号，按首次出现顺序编号。
            evidence_seq[evidence_id] = f"证-{len(evidence_seq) + 1:02d}"
        sources = hit.get("sources") or [None]
        for source in sources:
            source_text = (os.path.relpath(source["file"], Path(out_path).parent) + "\n" + source["sheet"] + "!" + source["cell"]) if source else "来源位置未记录，须人工核对"
            row = append_row(evidence_ws, [evidence_id, hit["company_a"], hit["company_b"], hit["dimension"],
                risk_text(hit["level"]), hit["evidence"], source["company"] if source else "", source["field"] if source else "",
                source["value"] if source else "", source_text, "返回对应汇总", hit["case_ref"], evidence_seq[evidence_id]])
            first_evidence.setdefault(key, evidence_id)
            if source:
                source_link(evidence_ws.cell(row, 10), source, out_path)
            if key in summary_rows:
                keyed_link(evidence_ws.cell(row, 11), "关系核查汇总", pair_id(*key))
    for key, row in summary_rows.items():
        if key in first_evidence:
            keyed_link(summary_ws.cell(row, 9), "证据明细", first_evidence[key])
    gaps_ws = create_sheet(wb, "数据覆盖与缺口",
        ["公司 / 项目", "资料维度 / 类型", "取得或处理情况", "影响的核查", "说明", "来源文件", "取数任务创建时间"],
        [30, 25, 36, 38, 60, 48, 25])
    # 维度级汇总置顶：全部公司都未取得的维度先总起一笔，说清影响面。
    label_key = {label: key for key, (label, _impact) in DIMENSIONS.items()}
    all_missing = []
    for key, (label, _impact) in DIMENSIONS.items():
        dim_rows = [record for record in coverage if record[1] == label]
        if dim_rows and all(record[2] == "未取得" for record in dim_rows):
            all_missing.append(label)
    degraded_mode = any(item.get("category") == "取数通道降级" for item in limitations)
    if all_missing:
        impacts = "、".join(dict.fromkeys(DIMENSIONS[label_key[label]][1] for label in all_missing))
        append_row(gaps_ws, ["全部核查对象", "维度级汇总",
            f"本次未取得维度：{'、'.join(all_missing)}，影响全部 {len(names)} 家", impacts,
            "相关核查规则整体受限" + ("；逐公司重复行已折叠，不再逐家列出" if degraded_mode else "；逐公司明细见下文"), "", ""])
    # 涉及命中公司的缺口行排在前面并标记，读者先看到削弱结论的缺口。
    ok_status = {"已取得记录", "明确无数据", "已核对范围，未见记录"}
    candidate_keys = {"customer", "supplier", "abnormal"}
    front, rest = [], []
    for record in coverage:
        if degraded_mode and record[1] in all_missing:
            continue
        values = list(record)
        is_gap = record[2] not in ok_status and label_key.get(record[1]) not in candidate_keys
        if record[0] in involved and is_gap:
            values[4] = ("涉及命中；" + values[4]) if values[4] else "涉及命中，缺口直接削弱命中线索的强度"
            front.append(values)
        else:
            rest.append(values)
    for values in front + rest:
        if values[5]:
            values[5] = os.path.relpath(values[5], Path(out_path).parent)
        append_row(gaps_ws, values)
    for item in errors + limitations:
        append_row(gaps_ws, [item.get("source", ""), item.get("category", ""), "需复核", "按说明确认受影响范围", item.get("message", ""), "", ""])
    explanation = create_sheet(wb, "任务说明", ["项目", "内容"], [32, 115])
    mode_names = {"discovery": "主动发现", "list_check": "名单核查", "existing_export": "已有数据离线核查"}
    for values in [["任务模式", mode_names.get(task_mode, task_mode or "未说明")], ["被审计单位", target_display],
                   ["报告生成时间", datetime.now().astimezone().isoformat(timespec="seconds")]] + [[key, value] for key, value in scope.items()]:
        append_row(explanation, values)
    append_row(explanation, ["阅读方法", "概览 → 关系汇总 → 证据明细 → 原始单元格。证据明细同一编号的多行是同一命中的不同来源。"])
    append_row(explanation, ["筛选与链接", "汇总与证据之间按编号定位，排序后仍可往返。查看全部证据时，按公司A和公司B筛选；来源文件同时标明工作表和单元格。"])
    append_row(explanation, ["原始资料", "报告与原始资料目录应一起移动。链接使用相对路径；若客户端只打开文件，按同时显示的工作表和单元格定位。"])
    append_row(explanation, ["日期口径", "核查基准日用于计算成立时间等规则；原始记录日期见对应原表字段。取数任务创建时间不等同于记录生效日期。"])
    for ws in wb.worksheets:
        finish_sheet(ws, 5 if ws.title in {"关系核查汇总", "证据明细"} else None)
    group_evidence_blocks(evidence_ws)
    for row, title in enumerate(list(SHEETS)[1:], 4):
        put(wb["目录"], row, 3, max(0, wb[title].max_row - 3))
    wb.active = 0
    wb.save(out_path)
    wb.close()
