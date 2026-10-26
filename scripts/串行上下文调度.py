# -*- coding: utf-8 -*-
"""128K 串行上下文窗口调度外壳（关联方识别技能）。

只管理调度占用、哈希与短状态、候选完整分页，不是业务真相：
任务与候选的完成与否以原任务状态、原始导出和八表报告为准。
本模块不得导入任何业务工作流模块。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import time
from pathlib import Path
from typing import Any, Sequence


class DispatchError(RuntimeError):
    """调度状态不允许当前动作。"""


STATE_NAME = "上下文调度状态.json"
LOG_NAME = "上下文调度记录.jsonl"
CHUNK_SIZE = 1 << 20
CONTROLLER_DISPATCH_LIMIT = 20


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


def _state_file(state_path: Path) -> Path:
    return Path(state_path).parent / STATE_NAME


def _log_file(state_path: Path) -> Path:
    return Path(state_path).parent / LOG_NAME


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _backup_file(state_path: Path, path: Path, backup_root: Path) -> None:
    """覆盖前把既有文件复制到备份目录并做字节比较，失败即停止。"""
    if not path.is_file():
        return
    target = Path(backup_root) / path.relative_to(Path(state_path).parent)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, target)
    if target.read_bytes() != path.read_bytes():
        raise DispatchError(f"备份字节校验失败：{path.name}")


def _os_replace_retry(tmp: Path, path: Path) -> None:
    # Windows 下新文件可能被杀毒扫描短暂锁定，replace 偶发拒绝访问，需重试
    for attempt in range(5):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(0.05 * (attempt + 1))


def _load_state_if_any(state_path: Path) -> dict[str, Any] | None:
    path = _state_file(state_path)
    if not path.is_file():
        return None
    state = json.loads(path.read_text(encoding="utf-8-sig"))
    if state.get("context_mode") not in ("1M", "128K"):
        raise DispatchError("调度状态文件损坏：context_mode 缺失或非法")
    return state


def _load_128k_state(state_path: Path) -> dict[str, Any]:
    state = _load_state_if_any(state_path)
    if state is None:
        raise DispatchError("尚未初始化 128K 调度状态")
    if state["context_mode"] != "128K":
        raise DispatchError("当前运行不是 128K 模式，禁止调度")
    return state


def _save_state_with_backup(state_path: Path, state: dict[str, Any], backup_root: Path) -> None:
    path = _state_file(state_path)
    _backup_file(state_path, path, backup_root)
    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8-sig")
    _os_replace_retry(tmp, path)


def _append_log_with_backup(state_path: Path, event: str, detail: dict[str, Any], backup_root: Path) -> None:
    path = _log_file(state_path)
    _backup_file(state_path, path, backup_root)
    line = json.dumps({"event": event, "time": _now(), "detail": detail}, ensure_ascii=False)
    data = line.encode("utf-8") + b"\n"
    if not path.is_file() or path.stat().st_size == 0:
        data = b"\xef\xbb\xbf" + data
    with open(path, "ab") as handle:
        handle.write(data)


def initialize(state_path: Path, mode: str, model_binding: str, backup_root: Path) -> dict[str, Any]:
    mode = str(mode).strip()
    if mode == "1M":
        return {"status": "bypassed"}
    if mode != "128K":
        raise DispatchError(f"不支持的上下文模式：{mode}")
    state_path = Path(state_path)
    existing = _load_state_if_any(state_path)
    if existing is not None:
        return existing
    if not state_path.is_file():
        raise DispatchError(f"原任务状态文件不存在：{state_path.name}")
    state: dict[str, Any] = {
        "context_mode": "128K",
        "model_binding": model_binding,
        "state_anchor": {"path": str(state_path), "sha256": sha256_file(state_path)},
        "active_dispatch": None,
        "controller_dispatch_count": 0,
        "created_at": _now(),
    }
    _save_state_with_backup(state_path, state, backup_root)
    _append_log_with_backup(state_path, "初始化", {"model_binding": model_binding}, backup_root)
    return state


def start_dispatch(state_path: Path, task_id: str, packet: Path, window_id: str, backup_root: Path) -> dict[str, Any]:
    state = _load_128k_state(state_path)
    if state["active_dispatch"] is not None:
        raise DispatchError("存在活动任务，禁止开始第二个任务")
    packet = Path(packet)
    if not packet.is_file():
        raise DispatchError("原任务文件不存在")
    state["active_dispatch"] = {
        "task_id": task_id,
        "window_id": window_id,
        "packet_path": str(packet),
        "packet_sha256": sha256_file(packet),
        "started_at": _now(),
    }
    _save_state_with_backup(state_path, state, backup_root)
    _append_log_with_backup(state_path, "start", state["active_dispatch"], backup_root)
    return state


def complete_dispatch(state_path: Path, task_id: str, result: Path, backup_root: Path) -> dict[str, Any]:
    """原工作流或报告程序接受并保存结果后才允许调用；这里只登记，不认定业务完成。"""
    state = _load_128k_state(state_path)
    active = state["active_dispatch"]
    if active is None:
        raise DispatchError("没有活动任务可登记完成")
    if active["task_id"] != task_id:
        raise DispatchError(f"任务编号与活动任务不匹配：{task_id} ≠ {active['task_id']}")
    result = Path(result)
    if not result.is_file():
        raise DispatchError("结果文件不存在")
    state_path = Path(state_path)
    detail = {
        "task_id": task_id,
        "window_id": active["window_id"],
        "result_path": str(result),
        "result_sha256": sha256_file(result),
        "state_anchor_path": str(state_path),
    }
    if state_path.is_file():
        detail["state_anchor_sha256"] = sha256_file(state_path)
        state["state_anchor"] = {"path": str(state_path), "sha256": detail["state_anchor_sha256"]}
    state["active_dispatch"] = None
    state["controller_dispatch_count"] = int(state.get("controller_dispatch_count", 0)) + 1
    _save_state_with_backup(state_path, state, backup_root)
    _append_log_with_backup(state_path, "完成", detail, backup_root)
    return state


def recover_dispatch(state_path: Path, task_id: str, reason: str, backup_root: Path) -> dict[str, Any]:
    """只解除调度占用并留痕，不改变候选、导出、报告或任务状态。"""
    state = _load_128k_state(state_path)
    active = state["active_dispatch"]
    if active is None:
        raise DispatchError("没有活动任务占用，无需恢复")
    if active["task_id"] != task_id:
        raise DispatchError(f"任务编号与活动任务不匹配：{task_id} ≠ {active['task_id']}")
    state["active_dispatch"] = None
    _save_state_with_backup(state_path, state, backup_root)
    _append_log_with_backup(state_path, "恢复", {"task_id": task_id, "window_id": active["window_id"], "reason": reason}, backup_root)
    return state


def status(state_path: Path) -> dict[str, Any]:
    """只输出调度所需简要字段，不输出完整候选记录。"""
    state = _load_128k_state(state_path)
    state_path = Path(state_path)
    summary: dict[str, Any] = {
        "context_mode": state["context_mode"],
        "model_binding": state.get("model_binding"),
        "stage": None,
        "status": None,
        "message_zh": None,
        "candidate_count": None,
        "report_paths": [],
    }
    if state_path.is_file():
        task_state = json.loads(state_path.read_text(encoding="utf-8-sig"))
        for key in ("stage", "status", "message_zh", "candidate_count", "report_paths"):
            summary[key] = task_state.get(key)
    count = int(state.get("controller_dispatch_count", 0))
    summary["active_dispatch"] = state.get("active_dispatch")
    summary["controller_dispatch_count"] = count
    summary["rotation_required"] = count >= CONTROLLER_DISPATCH_LIMIT
    return summary


def candidate_page(source: Path, page: int, page_size: int = 50) -> dict[str, Any]:
    """按原集合顺序输出一页完整候选；只改变展示，不改变集合。"""
    if not isinstance(page, int) or page < 1:
        raise DispatchError("页码必须从 1 开始")
    if not isinstance(page_size, int) or page_size < 1:
        raise DispatchError("页大小必须大于 0")
    source = Path(source)
    if not source.is_file():
        raise DispatchError("候选记录文件不存在")
    payload = json.loads(source.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict):
        records = payload.get("candidate_records", payload.get("candidates"))
    else:
        records = None
    if not isinstance(records, list):
        raise DispatchError("候选记录文件缺少 candidates 或 candidate_records 数组")
    total = len(records)
    last_page = max(1, -(-total // page_size))
    if page > last_page:
        raise DispatchError(f"页码越界：{page} 超过最后一页 {last_page}（共 {total} 条）")
    start = (page - 1) * page_size
    return {
        "page": page,
        "page_size": page_size,
        "total": total,
        "items": records[start : start + page_size],
    }


def handoff(state_path: Path, backup_root: Path | None = None) -> dict[str, Any]:
    """生成新主控窗口的简短续作信息，并重置主控调度计数。"""
    state = _load_128k_state(state_path)
    info = {
        "context_mode": state["context_mode"],
        "model_binding": state.get("model_binding"),
        "active_dispatch": state.get("active_dispatch"),
        "state_anchor": state.get("state_anchor"),
        "rotation_required": int(state.get("controller_dispatch_count", 0)) >= CONTROLLER_DISPATCH_LIMIT,
    }
    if backup_root is None:
        backup_root = Path(state_path).parent / "backup"
    if state.get("controller_dispatch_count"):
        state["controller_dispatch_count"] = 0
        _save_state_with_backup(state_path, state, backup_root)
        _append_log_with_backup(state_path, "主控轮换", {"handed_off": True}, backup_root)
    return info


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="128K 串行上下文窗口调度（关联方识别）")
    sub = parser.add_subparsers(dest="command", required=True)

    def with_state_path(name: str) -> argparse.ArgumentParser:
        child = sub.add_parser(name)
        child.add_argument("--state-path", required=True)
        return child

    child = with_state_path("初始化")
    child.add_argument("--mode", required=True)
    child.add_argument("--model-binding", default="")
    child.add_argument("--backup-root", required=True)
    child = with_state_path("开始")
    child.add_argument("--task-id", required=True)
    child.add_argument("--packet", required=True)
    child.add_argument("--window-id", required=True)
    child.add_argument("--backup-root", required=True)
    child = with_state_path("完成")
    child.add_argument("--task-id", required=True)
    child.add_argument("--result", required=True)
    child.add_argument("--backup-root", required=True)
    with_state_path("状态")
    child = sub.add_parser("候选页")
    child.add_argument("--source", required=True)
    child.add_argument("--page", type=int, required=True)
    child.add_argument("--page-size", type=int, default=50)
    with_state_path("交接")
    child = with_state_path("恢复")
    child.add_argument("--task-id", required=True)
    child.add_argument("--reason", required=True)
    child.add_argument("--backup-root", required=True)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    state_path = Path(args.state_path) if hasattr(args, "state_path") else None
    try:
        if args.command == "初始化":
            initialize(state_path, args.mode, args.model_binding, Path(args.backup_root))
        elif args.command == "开始":
            start_dispatch(state_path, args.task_id, Path(args.packet), args.window_id, Path(args.backup_root))
        elif args.command == "完成":
            complete_dispatch(state_path, args.task_id, Path(args.result), Path(args.backup_root))
        elif args.command == "状态":
            print(json.dumps(status(state_path), ensure_ascii=False, indent=2))
        elif args.command == "候选页":
            print(json.dumps(candidate_page(Path(args.source), args.page, args.page_size), ensure_ascii=False, indent=2))
        elif args.command == "交接":
            print(json.dumps(handoff(state_path), ensure_ascii=False, indent=2))
        elif args.command == "恢复":
            recover_dispatch(state_path, args.task_id, args.reason, Path(args.backup_root))
    except DispatchError as exc:
        print(f"ERROR={exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
