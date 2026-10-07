# 原始记录随数据保留；只引用实际参与本条匹配的字段。
from pathlib import Path
import re

from openpyxl.utils import get_column_letter


class SourceRow(tuple):
    """保持原有行的下标接口，同时保留原文件位置和表头。"""
    def __new__(cls, values, path, sheet, row, headers):
        item = super().__new__(cls, values)
        item.file = str(Path(path).resolve())
        item.sheet = sheet
        item.row = row
        item.headers = headers
        return item


# 每项为维度、用于匹配的列（1 基）、需要一同展示的上下文字段。
SOURCE_COLUMNS = {
    "phone": [("basic", (10,), ()), ("invoice", (7,), ())],
    "email": [("basic", (12,), ())],
    "website": [("basic", (11,), ())],
    "legal_person": [("basic", (4,), ())],
    "person": [("basic", (4,), ()), ("shareholder", (3,), (5, 7)),
               ("shareholder_new", (4,), (5, 6)), ("core_team", (4,), (5,)),
               ("main_persons", (3,), (4,)), ("actual_controller", (3,), ()),
               ("ultimate_beneficiary", (6,), (7,))],
    "controller": [("actual_controller", (3,), ())],
    "beneficiary": [("ultimate_beneficiary", (6,), (7,))],
    "address": [("basic", (22,), ()), ("invoice", (5,), ()), ("change", (5,), (3, 4, 6))],
    "invest": [("invest_new", (3,), (5,)), ("holding", (3,), (4,))],
    "shareholder": [("shareholder", (3,), (5, 7)), ("shareholder_new", (4,), (5, 6))],
    "names": [("basic", (1, 21), ())],
    "past_legal": [("legal_change", (2, 3), ())],
    "past_address": [("change", (5,), (3, 4, 6))],
    "pledge": [("pledge", (4, 6), ())],
    "mortgage": [("mortgage", (5, 6), ())],
    "trademark": [("trademark", (3,), ())],
    "software": [("software", (5,), ())],
    "wechat": [("wechat", (4,), ())],
    "profile": [("basic", (5, 6, 16, 22, 23), ())],
}


def trace(company, group, values=None, *, normalize=str, row_match=None):
    """values 为本次命中的值；未传时保留该单方判断使用的全部字段。"""
    wanted = {str(value) for value in values} if values is not None else None
    result = []
    for dimension, columns, context in SOURCE_COLUMNS[group]:
        for row in company.source_rows.get(dimension, []):
            if not isinstance(row, SourceRow):
                continue
            if row_match is not None and not row_match(row):
                continue
            if dimension == "change" and "地址" not in str(row[3] if len(row) > 3 else ""):
                continue
            matched = []
            for column in columns:
                raw = row[column - 1] if len(row) >= column else None
                if wanted is not None:
                    parsed = normalize(raw)
                    tokens = parsed if isinstance(parsed, (list, tuple, set)) else [parsed]
                    if not wanted.intersection(str(value) for value in tokens):
                        continue
                matched.append(column)
            if not matched:
                continue
            for column in dict.fromkeys(matched + list(context)):
                if column > len(row):
                    continue
                result.append({
                    "company": company.name, "file": row.file, "sheet": row.sheet,
                    "cell": f"{get_column_letter(column)}{row.row}", "value": row[column - 1],
                    "field": str(row.headers[column - 1] or f"第{column}列") if column <= len(row.headers) else f"第{column}列",
                    "dimension": dimension,
                })
    return unique_sources(result)


def unique_sources(sources):
    result = {}
    for source in sources:
        key = (source["company"], source["file"], source["sheet"], source["cell"])
        result[key] = source
    return [result[key] for key in sorted(result)]


def website_domain(value):
    return re.sub(r"https?://", "", str(value or "")).split("/")[0].replace("www.", "").lower()
