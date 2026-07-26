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
#"""受控比较并登记注协查询上游变化，不自动覆盖正式技能。"""

from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
from datetime import datetime
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Callable, Iterable, Optional


SYNC_FILE = "sync-stage.json"
VERIFY_FILE = "verification.json"
EXPECTED_CHECKS = ("unit_tests", "runtime_paths", "encoding", "secret_leaks")


@dataclass(frozen=True)
class FileComparison:
    path: str
    upstream_sha256: str
    recorded_sha256: str
    sha256_status: str
    upstream_git_blob: str
    recorded_git_blob: str
    git_blob_status: str


@dataclass(frozen=True)
class CheckReport:
    repository: str
    source_path: str
    upstream_commit: str
    recorded_commit: str
    commit_status: str
    files: list[FileComparison]
    source_mapping_present: bool

    @property
    def is_current(self) -> bool:
        return (
            self.commit_status == "same"
            and self.source_mapping_present
            and all(
                item.sha256_status == "same" and item.git_blob_status == "same"
                for item in self.files
            )
        )

    def to_dict(self) -> dict:
        payload = asdict(self)
        payload["is_current"] = self.is_current
        return payload


@dataclass(frozen=True)
class StageReport:
    staging_dir: str
    manifest_path: str
    changed_files: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass(frozen=True)
class VerifyReport:
    blockers: list[str]
    checks: dict[str, bool]
    confirmations: list[str]
    merged_files: list[str]
    local_hashes: dict[str, str]
    can_record: bool

    def to_dict(self) -> dict:
        return asdict(self)


def _read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _git_blob(path: Path) -> str:
    data = path.read_bytes()
    if b"\0" not in data:
        data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()


def _is_within(path: Path, parent: Path) -> bool:
    resolved = path.resolve()
    root = parent.resolve()
    return resolved == root or root in resolved.parents


def _source_record(sources_path: Path, component: str) -> tuple[dict, dict]:
    payload = _read_json(sources_path)
    for source in payload.get("sources", []):
        if source.get("component") == component:
            return payload, source
    raise ValueError(f"SOURCES.json 未登记来源组件：{component}")


def _relative_source_path(recorded_path: str, source_path: str) -> Path:
    normalized = recorded_path.replace("\\", "/")
    prefix = source_path.strip("/") + "/"
    if normalized.startswith(prefix):
        normalized = normalized[len(prefix) :]
    path = Path(normalized)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError(f"非法来源路径：{recorded_path}")
    return path


def check_upstream(
    snapshot_dir: Path | str,
    sources_path: Path | str,
    *,
    commit: str,
    component: str = "cicpa-company-query",
) -> CheckReport:
    snapshot_dir = Path(snapshot_dir)
    sources_path = Path(sources_path)
    _payload, source = _source_record(sources_path, component)
    comparisons = []
    for recorded in source.get("source_snapshot", []):
        relative = _relative_source_path(recorded["path"], source["source_path"])
        upstream_path = snapshot_dir / relative
        if upstream_path.exists():
            upstream_sha256 = _sha256(upstream_path)
            upstream_blob = _git_blob(upstream_path)
        else:
            upstream_sha256 = ""
            upstream_blob = ""
        comparisons.append(
            FileComparison(
                path=recorded["path"],
                upstream_sha256=upstream_sha256,
                recorded_sha256=recorded.get("sha256", ""),
                sha256_status=(
                    "same"
                    if upstream_sha256 and upstream_sha256 == recorded.get("sha256")
                    else "missing"
                    if not upstream_sha256
                    else "changed"
                ),
                upstream_git_blob=upstream_blob,
                recorded_git_blob=recorded.get("git_blob", ""),
                git_blob_status=(
                    "same"
                    if upstream_blob and upstream_blob == recorded.get("git_blob")
                    else "missing"
                    if not upstream_blob
                    else "changed"
                ),
            )
        )
    recorded_commit = source.get("baseline_commit", "")
    return CheckReport(
        repository=source.get("repository", ""),
        source_path=source.get("source_path", ""),
        upstream_commit=commit,
        recorded_commit=recorded_commit,
        commit_status="same" if commit == recorded_commit else "changed",
        files=comparisons,
        source_mapping_present=bool(source.get("local_modules")),
    )


def stage_upstream(
    snapshot_dir: Path | str,
    staging_dir: Path | str,
    sources_path: Path | str,
    *,
    skill_root: Path | str,
    commit: str,
    component: str = "cicpa-company-query",
) -> StageReport:
    snapshot_dir = Path(snapshot_dir)
    staging_dir = Path(staging_dir)
    sources_path = Path(sources_path)
    skill_root = Path(skill_root)
    if _is_within(staging_dir, skill_root):
        raise ValueError("暂存目录必须位于正式 skill 目录之外")
    if staging_dir.exists():
        raise FileExistsError(f"暂存目录已存在，为避免覆盖已停止：{staging_dir}")
    staging_dir.mkdir(parents=True)

    report = check_upstream(
        snapshot_dir,
        sources_path,
        commit=commit,
        component=component,
    )
    _payload, source = _source_record(sources_path, component)
    changed_files = [
        item.path
        for item in report.files
        if item.sha256_status != "same" or item.git_blob_status != "same"
    ]
    source_snapshot = []
    for item in report.files:
        relative = _relative_source_path(item.path, source["source_path"])
        upstream_path = snapshot_dir / relative
        if upstream_path.exists():
            source_snapshot.append(
                {
                    "path": item.path,
                    "sha256": _sha256(upstream_path),
                    "git_blob": _git_blob(upstream_path),
                }
            )
        if item.path in changed_files and upstream_path.exists():
            destination = staging_dir / "upstream" / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(upstream_path, destination)

    local_hashes = {}
    for relative in source.get("local_modules", []):
        local_path = skill_root / Path(relative)
        if local_path.is_file():
            local_hashes[relative] = _sha256(local_path)
    manifest = {
        "schema_version": 1,
        "component": component,
        "repository": source.get("repository", ""),
        "source_path": source.get("source_path", ""),
        "commit": commit,
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "changed_files": changed_files,
        "source_snapshot": source_snapshot,
        "check": report.to_dict(),
        "local_modules": list(source.get("local_modules", [])),
        "local_hashes_at_stage": local_hashes,
    }
    manifest_path = staging_dir / SYNC_FILE
    _write_json(manifest_path, manifest)
    return StageReport(
        staging_dir=str(staging_dir),
        manifest_path=str(manifest_path),
        changed_files=changed_files,
    )


def _classify_sensitive_changes(
    staging_dir: Path,
    changed_files: Iterable[str],
    source_path: str = "",
) -> set[str]:
    categories = set()
    for recorded_path in changed_files:
        relative = Path(recorded_path.replace("\\", "/"))
        source_parts = Path(source_path.replace("\\", "/")).parts
        if source_parts and relative.parts[: len(source_parts)] == source_parts:
            relative = Path(*relative.parts[len(source_parts) :])
        candidate = staging_dir / "upstream" / relative
        text = candidate.read_text(encoding="utf-8-sig").lower() if candidate.exists() else ""
        if re.search(r"login|cookie|auth|password|验证码|登录|会话", text):
            categories.add("auth")
        if re.search(r"/api/|https?://|endpoint|requests?\\.", text):
            categories.add("api")
        if re.search(r"export|dimension|download|\\.xlsx|维度|导出|文件名", text):
            categories.add("export")
    # ponytail: 无旧源码正文时采用保守关键词分类；未来若保存规范化基线，可升级为逐段语义差异。
    return categories


def _default_checks(skill_root: Path) -> dict[str, bool]:
    completed = subprocess.run(
        [
            sys.executable,
            "-X",
            "utf8",
            "-m",
            "unittest",
            "discover",
            "-s",
            "scripts/tests",
        ],
        cwd=skill_root,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    runtime_patterns = (
        "~/" + ".claude/skills/",
        "~/" + ".agents/skills/",
        "workbuddy skills/" + "cicpa-company-query",
        "workbuddy skills\\" + "cicpa-company-query",
    )
    runtime_paths_ok = True
    for path in (skill_root / "scripts").rglob("*.py"):
        if "tests" in path.parts:
            continue
        text = path.read_text(encoding="utf-8-sig")
        if any(pattern in text for pattern in runtime_patterns):
            runtime_paths_ok = False
            break

    encoding_ok = True
    text_suffixes = {".py", ".md", ".json", ".yaml", ".yml", ".txt"}
    for path in skill_root.rglob("*"):
        if not path.is_file() or "__pycache__" in path.parts:
            continue
        if path.suffix.lower() not in text_suffixes and path.name != "NOTICE":
            continue
        try:
            path.read_bytes().decode("utf-8-sig", errors="strict")
        except UnicodeDecodeError:
            encoding_ok = False
            break

    secret_patterns = (
        re.compile(r"authorization\\s*[:=]\\s*bearer\\s+[a-z0-9._-]{20,}", re.I),
        re.compile(r"(?:sessionid|cicpa_token)\\s*[:=]\\s*[\"']?[a-z0-9._-]{20,}", re.I),
        re.compile(r"[\"']password[\"']\\s*:\\s*[\"'][^\"']+[\"']", re.I),
    )
    secret_leaks_ok = True
    scan_suffixes = {".json", ".yaml", ".yml", ".txt", ".log", ".md"}
    for path in skill_root.rglob("*"):
        if not path.is_file() or "__pycache__" in path.parts:
            continue
        if path.suffix.lower() not in scan_suffixes and path.name != "NOTICE":
            continue
        text = path.read_text(encoding="utf-8-sig")
        if any(pattern.search(text) for pattern in secret_patterns):
            secret_leaks_ok = False
            break
    return {
        "unit_tests": completed.returncode == 0,
        "runtime_paths": runtime_paths_ok,
        "encoding": encoding_ok,
        "secret_leaks": secret_leaks_ok,
    }


def verify_staged_update(
    staging_dir: Path | str,
    skill_root: Path | str,
    sources_path: Path | str,
    *,
    confirmations: Optional[Iterable[str]] = None,
    merged_files: Optional[Iterable[str]] = None,
    checks_runner: Optional[Callable[[Path], dict[str, bool]]] = None,
) -> VerifyReport:
    staging_dir = Path(staging_dir)
    skill_root = Path(skill_root)
    sources_path = Path(sources_path)
    if _is_within(staging_dir, skill_root):
        raise ValueError("暂存目录必须位于正式 skill 目录之外")
    manifest = _read_json(staging_dir / SYNC_FILE)
    _source_record(sources_path, manifest["component"])
    confirmations_set = set(confirmations or [])
    merged = list(dict.fromkeys(merged_files or []))
    blockers = []

    categories = _classify_sensitive_changes(
        staging_dir,
        manifest.get("changed_files", []),
        str(manifest.get("source_path", "")),
    )
    messages = {
        "auth": "登录与凭据接口发生变化，需要人工复核",
        "api": "API 端点或请求协议发生变化，需要人工复核",
        "export": "导出维度或文件命名发生变化，需要人工复核",
    }
    for category in sorted(categories):
        if category not in confirmations_set:
            blockers.append(messages[category])
    if manifest.get("changed_files") and (
        "local-protections" not in confirmations_set or not merged
    ):
        blockers.append("本地用户交互、安全、限速和报告语义需确认保持不变")

    local_hashes = {}
    for relative in merged:
        candidate = skill_root / Path(relative)
        if not _is_within(candidate, skill_root) or not candidate.is_file():
            blockers.append(f"登记的本地合并文件无效：{relative}")
            continue
        local_hashes[relative] = _sha256(candidate)

    runner = checks_runner or _default_checks
    raw_checks = runner(skill_root)
    checks = {name: bool(raw_checks.get(name, False)) for name in EXPECTED_CHECKS}
    check_messages = {
        "unit_tests": "单元测试未通过",
        "runtime_paths": "发现外部运行时 skill 路径",
        "encoding": "编码检查未通过",
        "secret_leaks": "秘密泄漏检查未通过",
    }
    for name, passed in checks.items():
        if not passed:
            blockers.append(check_messages[name])
    blockers = list(dict.fromkeys(blockers))
    report = VerifyReport(
        blockers=blockers,
        checks=checks,
        confirmations=sorted(confirmations_set),
        merged_files=merged,
        local_hashes=local_hashes,
        can_record=not blockers,
    )
    _write_json(staging_dir / VERIFY_FILE, report.to_dict())
    return report


def record_verified_update(
    staging_dir: Path | str,
    sources_path: Path | str,
    *,
    skill_root: Path | str,
    backup_root: Optional[Path | str] = None,
) -> dict:
    staging_dir = Path(staging_dir)
    sources_path = Path(sources_path)
    skill_root = Path(skill_root)
    verification_path = staging_dir / VERIFY_FILE
    if not verification_path.exists():
        raise RuntimeError("尚未生成验证记录，不能登记")
    verification = _read_json(verification_path)
    if not verification.get("can_record"):
        raise RuntimeError("验证仍有阻断项，不能登记")
    for relative, expected_hash in verification.get("local_hashes", {}).items():
        candidate = skill_root / Path(relative)
        if not candidate.is_file() or _sha256(candidate) != expected_hash:
            raise RuntimeError(f"正式文件在验证后发生变化，需重新验证：{relative}")

    manifest = _read_json(staging_dir / SYNC_FILE)
    payload, source = _source_record(sources_path, manifest["component"])
    stamp = datetime.now().astimezone()
    backup_base = (
        Path(backup_root)
        if backup_root is not None
        else Path.home() / "BackUp"
    )
    backup_dir = backup_base / (
        "related-party-identification_sync_" + stamp.strftime("%Y%m%d_%H%M%S")
    )
    backup_dir.mkdir(parents=True, exist_ok=False)
    backup_path = backup_dir / sources_path.name
    shutil.copy2(sources_path, backup_path)
    if _sha256(backup_path) != _sha256(sources_path):
        raise RuntimeError("SOURCES.json 备份校验失败，已停止登记")

    source["baseline_commit"] = manifest["commit"]
    source["source_snapshot"] = manifest.get("source_snapshot", [])
    source["last_synced_commit"] = manifest["commit"]
    source["last_synced_at"] = stamp.isoformat(timespec="seconds")
    source["last_sync"] = {
        "changed_files": manifest.get("changed_files", []),
        "confirmed_categories": verification.get("confirmations", []),
        "merged_local_hashes": verification.get("local_hashes", {}),
    }
    _write_json(sources_path, payload)
    return {
        "status": "recorded",
        "sources_path": str(sources_path),
        "backup_path": str(backup_path),
        "commit": manifest["commit"],
    }


def _print_json(payload: dict) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def build_parser() -> argparse.ArgumentParser:
    skill_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description="关联方识别上游受控同步工具")
    subparsers = parser.add_subparsers(dest="command", required=True)

    check_parser = subparsers.add_parser("check", help="只读比较上游快照")
    check_parser.add_argument("--snapshot", required=True)
    check_parser.add_argument("--commit", required=True)
    check_parser.add_argument(
        "--component",
        default="cicpa-company-query",
        choices=["cicpa-company-query", "OpenCLI Browser Bridge extension"],
    )
    check_parser.add_argument(
        "--sources",
        default=str(skill_root / "references" / "SOURCES.json"),
    )

    stage_parser = subparsers.add_parser("stage", help="复制候选变化到技能外暂存目录")
    stage_parser.add_argument("--snapshot", required=True)
    stage_parser.add_argument("--staging", required=True)
    stage_parser.add_argument("--commit", required=True)
    stage_parser.add_argument(
        "--component",
        default="cicpa-company-query",
        choices=["cicpa-company-query", "OpenCLI Browser Bridge extension"],
    )
    stage_parser.add_argument("--skill-root", default=str(skill_root))
    stage_parser.add_argument(
        "--sources",
        default=str(skill_root / "references" / "SOURCES.json"),
    )

    verify_parser = subparsers.add_parser("verify", help="验证暂存变化和正式本地保护层")
    verify_parser.add_argument("--staging", required=True)
    verify_parser.add_argument("--skill-root", default=str(skill_root))
    verify_parser.add_argument(
        "--sources",
        default=str(skill_root / "references" / "SOURCES.json"),
    )
    verify_parser.add_argument("--confirm", action="append", default=[])
    verify_parser.add_argument("--merged-file", action="append", default=[])

    record_parser = subparsers.add_parser("record", help="登记已经合并并验证通过的来源状态")
    record_parser.add_argument("--staging", required=True)
    record_parser.add_argument("--skill-root", default=str(skill_root))
    record_parser.add_argument(
        "--sources",
        default=str(skill_root / "references" / "SOURCES.json"),
    )
    record_parser.add_argument("--backup-root")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "check":
        report = check_upstream(
            args.snapshot,
            args.sources,
            commit=args.commit,
            component=args.component,
        )
        _print_json(report.to_dict())
        return 0
    if args.command == "stage":
        report = stage_upstream(
            args.snapshot,
            args.staging,
            args.sources,
            skill_root=args.skill_root,
            commit=args.commit,
            component=args.component,
        )
        _print_json(report.to_dict())
        return 0
    if args.command == "verify":
        report = verify_staged_update(
            args.staging,
            args.skill_root,
            args.sources,
            confirmations=args.confirm,
            merged_files=args.merged_file,
        )
        _print_json(report.to_dict())
        return 0 if report.can_record else 2
    result = record_verified_update(
        args.staging,
        args.sources,
        skill_root=args.skill_root,
        backup_root=args.backup_root,
    )
    _print_json(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
