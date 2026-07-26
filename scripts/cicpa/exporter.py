# 关联方识别与核查 — 注协完整维度导出器
# 本文件改编自 nigo/nigo-skills/cicpa-company-query(MIT)
# 上游作者: nigo(涂佳兵) | 原始仓库: https://github.com/nigo81/nigo-skills
# 上游协议: MIT
#
# 本地修改: CPA-Q(quanfanpro-code)
# 本文件的修改部分同样以 MIT 协议发布,与上游保持一致。
#"""中注协完整维度导出的安全批次、任务关联与文件验证。"""

from dataclasses import asdict, dataclass, field
from datetime import datetime
import json
import os
from pathlib import Path
import tempfile
from typing import Any, Callable, Dict, List, Optional
import unicodedata
import zipfile

from .client import CicpaClient, CicpaError


class ExportError(CicpaError):
    """完整维度导出基础错误。"""


class ExportTaskNotCorrelated(ExportError):
    """下载中心任务无法与本次批次唯一对应。"""


class ExportValidationError(ExportError):
    """下载文件未通过结构或企业名称验证。"""


@dataclass
class ExportState:
    batch_no: str
    company_names: List[str]
    created_at: str
    triggered_at: str
    preexisting_task_ids: List[str] = field(default_factory=list)
    staging_dir: str = ""
    task_id: str = ""
    task_name: str = ""
    status: str = "waiting"
    archive_path: str = ""
    extract_dir: str = ""
    message_zh: str = ""


def save_export_state(path: Path, state: ExportState) -> None:
    """以不带 BOM 的 UTF-8 原子保存非敏感任务状态。"""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    text = json.dumps(asdict(state), ensure_ascii=False, indent=2)
    with open(temporary, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(text)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, destination)


def load_export_state(path: Path) -> ExportState:
    with open(path, "r", encoding="utf-8") as stream:
        payload = json.load(stream)
    return ExportState(**payload)


class CicpaExporter:
    ZSK_BASE = "https://zsk-cmis.cicpa.org.cn"

    def __init__(
        self,
        client: CicpaClient,
        *,
        staging_root: Optional[Path] = None,
        now: Callable[[], str] = lambda: datetime.now().astimezone().strftime(
            "%Y-%m-%d %H:%M"
        ),
    ):
        self.client = client
        self.staging_root = Path(staging_root) if staging_root is not None else None
        self._now = now

    @staticmethod
    def _business_data(payload: Any, alias: str) -> Any:
        if not isinstance(payload, dict):
            raise ExportError("{}返回结构无效".format(alias))
        if payload.get("status_code") not in (None, 0):
            raise ExportError(
                "{}失败：{}".format(alias, payload.get("status_msg") or "未知业务错误")
            )
        return payload.get("data", {})

    def _new_staging_dir(self) -> Path:
        if self.staging_root is not None:
            self.staging_root.mkdir(parents=True, exist_ok=True)
        return Path(
            tempfile.mkdtemp(
                prefix="rpi_export_",
                dir=str(self.staging_root) if self.staging_root is not None else None,
            )
        )

    @staticmethod
    def _normalize_names(company_names: List[str]) -> List[str]:
        result = []
        seen = set()
        for value in company_names:
            name = str(value).strip()
            if name and name not in seen:
                seen.add(name)
                result.append(name)
        if not result:
            raise ValueError("企业名单不能为空")
        return result

    def _upload_companies(self, company_names: List[str], staging_dir: Path) -> str:
        try:
            import openpyxl
        except ImportError as exc:
            raise ExportError("缺少 openpyxl，需先完成依赖预检") from exc

        workbook_path = staging_dir / "upload-companies.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["企业名称"])
        for name in company_names:
            sheet.append([name])
        workbook.save(workbook_path)
        content = workbook_path.read_bytes()
        payload = self.client.request_json(
            "POST",
            self.ZSK_BASE + "/open/industry_chain_api/v1/batch_search/upload_local",
            files={
                "file": (
                    "companies.xlsx",
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            timeout=60,
        )
        data = self._business_data(payload, "上传企业名单")
        batch_no = str(data.get("batch_id") or "")
        if not batch_no:
            raise ExportError("上传成功响应中缺少批次号")
        return batch_no

    def _trigger_search(self, batch_no: str) -> None:
        payload = self.client.request_json(
            "POST",
            self.ZSK_BASE + "/open/industry_chain_api/v1/batch_search/enter_search",
            json={
                "batch_id": batch_no,
                "page": 1,
                "page_size": 10,
                "type": "credit",
            },
        )
        self._business_data(payload, "触发批量查询")

    def _get_dimensions(self) -> List[str]:
        payload = self.client.request_json(
            "GET",
            self.ZSK_BASE
            + "/open/industry_chain_api/v1/search/export/get_dimension_class",
            params={"internal": "false"},
            cache_key=("export", "dimensions"),
        )
        data = self._business_data(payload, "获取导出维度")
        dimensions = []
        for group in data if isinstance(data, list) else []:
            for child in group.get("children", []) if isinstance(group, dict) else []:
                code = child.get("second_dimension_code") if isinstance(child, dict) else None
                if code:
                    dimensions.append(str(code))
        if not dimensions:
            raise ExportError("未取得可用导出维度")
        return dimensions

    def _list_tasks(self, kind: str = "normal") -> List[Dict[str, Any]]:
        payload = self.client.request_json(
            "GET",
            self.ZSK_BASE + "/open/industry_chain_api/v1/download/task_list",
            kind=kind,
            params={"page": 1, "pageSize": 20},
        )
        data = self._business_data(payload, "读取下载中心")
        if not isinstance(data, dict):
            return []
        return [item for item in data.get("list", []) if isinstance(item, dict)]

    def _trigger_export(self, batch_no: str, dimensions: List[str]) -> str:
        payload = self.client.request_json(
            "POST",
            self.ZSK_BASE
            + "/open/industry_chain_api/v1/batch_search/enter_search_out_type",
            json={"batch_id": batch_no, "dimensions": dimensions},
        )
        data = self._business_data(payload, "触发完整维度导出")
        if isinstance(data, dict):
            return str(data.get("task_id") or "")
        return ""

    def start_export(self, company_names: List[str], state_path: Path) -> ExportState:
        """上传名单并触发导出，返回可恢复状态。"""
        names = self._normalize_names(company_names)
        staging_dir = self._new_staging_dir()
        batch_no = self._upload_companies(names, staging_dir)
        self._trigger_search(batch_no)
        dimensions = self._get_dimensions()
        before = self._list_tasks()
        triggered_at = self._now()
        task_id = self._trigger_export(batch_no, dimensions)
        state = ExportState(
            batch_no=batch_no,
            company_names=names,
            created_at=triggered_at,
            triggered_at=triggered_at,
            preexisting_task_ids=[
                str(task.get("task_id"))
                for task in before
                if task.get("task_id") is not None
            ],
            staging_dir=str(staging_dir),
            task_id=task_id,
            status="waiting",
            message_zh="完整维度导出已触发，正在等待对应下载任务",
        )
        save_export_state(state_path, state)
        return state

    @staticmethod
    def correlate_task(
        state: ExportState,
        tasks: List[Dict[str, Any]],
        *,
        allow_missing: bool = False,
    ):
        """只接受能与本次批次唯一对应的下载任务。"""
        batch_tasks = [task for task in tasks if task.get("type") == "批量查询"]
        if state.task_id:
            exact = [
                task
                for task in batch_tasks
                if str(task.get("task_id") or "") == state.task_id
            ]
            if len(exact) == 1:
                return exact[0]
            if allow_missing:
                return None
            raise ExportTaskNotCorrelated("下载中心没有找到已记录的任务号")

        candidates = []
        for task in batch_tasks:
            task_id = str(task.get("task_id") or "")
            if not task_id or task_id in state.preexisting_task_ids:
                continue
            task_batch = str(task.get("batch_id") or "")
            if task_batch and task_batch != state.batch_no:
                continue
            task_date = str(task.get("date") or "")
            if not task_date or task_date < state.triggered_at:
                continue
            candidates.append(task)

        matching_batch = [
            task for task in candidates if str(task.get("batch_id") or "") == state.batch_no
        ]
        if len(matching_batch) == 1:
            return matching_batch[0]
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            raise ExportTaskNotCorrelated("同时出现多个新下载任务，无法唯一确认本次任务")
        if allow_missing:
            return None
        raise ExportTaskNotCorrelated("尚未发现能与本次批次对应的下载任务")

    def wait_for_task(
        self,
        state: ExportState,
        state_path: Path,
        *,
        max_polls: int = 18,
    ) -> Dict[str, Any]:
        """按轮询随机等待策略寻找并跟踪唯一任务。"""
        for _attempt in range(max_polls):
            task = self.correlate_task(
                state,
                self._list_tasks(kind="poll"),
                allow_missing=True,
            )
            if task is None:
                continue
            state.task_id = str(task.get("task_id") or "")
            state.task_name = str(task.get("name") or "")
            state.message_zh = "已关联下载任务，正在等待文件可下载"
            save_export_state(state_path, state)
            if task.get("status") == 1:
                return task
        state.message_zh = "轮询结束，未发现可唯一对应且已完成的下载任务"
        save_export_state(state_path, state)
        raise ExportTaskNotCorrelated(state.message_zh)

    @staticmethod
    def _safe_extract(archive_path: Path, extract_dir: Path) -> None:
        extract_root = extract_dir.resolve()
        with zipfile.ZipFile(archive_path, "r") as archive:
            members = archive.infolist()
            for member in members:
                destination = (extract_dir / member.filename).resolve()
                if os.path.commonpath([str(extract_root), str(destination)]) != str(
                    extract_root
                ):
                    raise ExportValidationError("ZIP 中存在越出暂存目录的路径")
            for member in members:
                destination = extract_dir / member.filename
                if member.is_dir():
                    destination.mkdir(parents=True, exist_ok=True)
                    continue
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, open(destination, "wb") as target:
                    target.write(source.read())

    @staticmethod
    def _normalize_company_name(value: str) -> str:
        normalized = unicodedata.normalize("NFKC", str(value))
        return "".join(normalized.split())

    @classmethod
    def _validate_companies(cls, extract_dir: Path, expected_names: List[str]) -> None:
        try:
            import openpyxl
        except ImportError as exc:
            raise ExportValidationError("缺少 openpyxl，无法验证导出文件") from exc

        workbooks = list(extract_dir.rglob("*.xlsx"))
        base_file = next(
            (path for path in workbooks if "基础工商信息" in path.stem),
            None,
        )
        if base_file is None:
            base_file = next(
                (
                    path
                    for path in workbooks
                    if "基础" in path.stem and "工商" in path.stem
                ),
                None,
            )
        if base_file is None:
            raise ExportValidationError("完整导出中缺少基础工商信息工作簿")

        workbook = openpyxl.load_workbook(base_file, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            actual_names = [
                str(row[0]).strip()
                for row in sheet.iter_rows(min_row=2, values_only=True)
                if row and row[0]
            ]
        finally:
            workbook.close()

        expected = {cls._normalize_company_name(name) for name in expected_names}
        actual = {cls._normalize_company_name(name) for name in actual_names}
        if expected != actual:
            raise ExportValidationError(
                "下载文件企业集合与本次上传名单不一致：预期 {} 家，实际 {} 家".format(
                    len(expected), len(actual)
                )
            )

    def download_and_validate(
        self,
        state: ExportState,
        task: Dict[str, Any],
        state_path: Path,
    ) -> Path:
        """下载到暂存目录并核对 ZIP、路径和公司集合。"""
        url = str(task.get("url") or "")
        if not url:
            raise ExportValidationError("对应下载任务没有下载地址")
        if url.startswith("/"):
            url = self.ZSK_BASE + url
        staging_dir = Path(state.staging_dir)
        staging_dir.mkdir(parents=True, exist_ok=True)
        archive_path = staging_dir / "complete-dimensions.zip"
        archive_path.write_bytes(self.client.request_bytes("GET", url, timeout=120))
        if not zipfile.is_zipfile(archive_path):
            state.status = "failed"
            state.message_zh = "下载内容不是有效 ZIP"
            save_export_state(state_path, state)
            raise ExportValidationError(state.message_zh)

        extract_dir = staging_dir / "complete-dimensions_files"
        extract_dir.mkdir(parents=True, exist_ok=True)
        try:
            self._safe_extract(archive_path, extract_dir)
            self._validate_companies(extract_dir, state.company_names)
        except ExportValidationError as exc:
            state.status = "failed"
            state.message_zh = str(exc)
            save_export_state(state_path, state)
            raise

        state.task_id = str(task.get("task_id") or state.task_id)
        state.task_name = str(task.get("name") or state.task_name)
        state.status = "completed"
        state.archive_path = str(archive_path)
        state.extract_dir = str(extract_dir)
        state.message_zh = "完整维度导出已下载并通过企业名称验证"
        save_export_state(state_path, state)
        return extract_dir
