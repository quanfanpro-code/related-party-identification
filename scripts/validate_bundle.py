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
#"""一键校验关联方识别技能是否可安全独立分发。"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Callable, Optional


REQUIRED_FILES = (
    "SKILL.md",
    "NOTICE",
    "requirements.txt",
    "agents/openai.yaml",
    "references/SOURCES.json",
    "references/user-flow.md",
    "references/cases.md",
    "references/dimensions.md",
    "references/rules.md",
    "scripts/discovery.py",
    "scripts/discover_scope.py",
    "scripts/related_party_check.py",
    "scripts/related_party_workflow.py",
    "scripts/sync_upstream.py",
    "scripts/validate_bundle.py",
    "scripts/cicpa/auth.py",
    "scripts/cicpa/client.py",
    "scripts/cicpa/edge_bridge.py",
    "scripts/cicpa/exporter.py",
    "scripts/cicpa/opencli_setup.py",
)
CHECK_NAMES = (
    "required_files",
    "python_syntax",
    "unit_tests",
    "runtime_paths",
    "secret_leaks",
    "encoding",
    "metadata",
    "provenance",
    "dialog_entries",
)
TEXT_SUFFIXES = {".py", ".md", ".json", ".jsonl", ".yaml", ".yml", ".txt", ".log"}
IGNORED_DIRECTORY_NAMES = {".git", ".pytest_cache", "__pycache__"}


@dataclass(frozen=True)
class BundleValidationReport:
    checks: dict[str, bool]
    errors: list[str]
    test_count: int
    test_output: str

    @property
    def is_valid(self) -> bool:
        return all(self.checks.get(name, False) for name in CHECK_NAMES)


def _relative(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def _text_files(root: Path):
    for path in root.rglob("*"):
        if not path.is_file() or IGNORED_DIRECTORY_NAMES.intersection(path.parts):
            continue
        if path.suffix.lower() in TEXT_SUFFIXES or path.name == "NOTICE":
            yield path


def _default_test_runner(root: Path) -> tuple[bool, int, str]:
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
            "-p",
            "test_*.py",
            "-v",
        ],
        cwd=root,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    output = completed.stdout + completed.stderr
    match = re.search(r"Ran\s+(\d+)\s+tests?", output)
    count = int(match.group(1)) if match else 0
    return completed.returncode == 0, count, output


def _check_required_files(root: Path, errors: list[str]) -> bool:
    missing = [relative for relative in REQUIRED_FILES if not (root / relative).is_file()]
    for relative in missing:
        errors.append(f"缺少必需文件：{relative}")
    return not missing


def _check_python_syntax(root: Path, errors: list[str]) -> bool:
    ok = True
    for path in (root / "scripts").rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        try:
            source = path.read_text(encoding="utf-8-sig")
            compile(source, str(path), "exec")
        except (OSError, UnicodeError, SyntaxError) as exc:
            ok = False
            errors.append(f"Python 语法检查失败：{_relative(path, root)}：{exc}")
    return ok


def _check_runtime_paths(root: Path, errors: list[str]) -> bool:
    patterns = (
        "~/" + ".claude/skills/",
        "~/" + ".agents/skills/",
        "workbuddy skills/" + "cicpa-company-query",
        "workbuddy skills\\" + "cicpa-company-query",
    )
    ok = True
    for path in (root / "scripts").rglob("*.py"):
        if "tests" in path.parts or "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8-sig")
        for pattern in patterns:
            if pattern in text:
                ok = False
                errors.append(
                    f"发现外部运行时 skill 路径：{_relative(path, root)}：{pattern}"
                )
    return ok


def _check_secret_leaks(root: Path, errors: list[str]) -> bool:
    patterns = (
        (
            "授权头",
            re.compile(r"authorization\s*[:=]\s*bearer\s+[a-z0-9._-]{20,}", re.I),
        ),
        (
            "会话凭据",
            re.compile(
                r"(?:sessionid|cicpa_token)\s*[:=]\s*[\"']?[a-z0-9._-]{20,}",
                re.I,
            ),
        ),
        (
            "密码字段",
            re.compile(r"[\"']password[\"']\s*:\s*[\"'][^\"']+[\"']", re.I),
        ),
    )
    ok = True
    # 只枚举传入的正式 skill 根目录，不访问其父目录或技能库根目录。
    for path in _text_files(root):
        text = path.read_text(encoding="utf-8-sig")
        for label, pattern in patterns:
            if pattern.search(text):
                ok = False
                errors.append(f"发现模拟或明文{label}：{_relative(path, root)}")
    return ok


def _expected_bom(path: Path, root: Path, text: str) -> Optional[bool]:
    relative = _relative(path, root)
    if relative == "SKILL.md":
        return False
    if path.suffix.lower() in {".json", ".jsonl", ".yaml", ".yml"}:
        return False
    if relative == "requirements.txt":
        return False
    if path.name == "NOTICE":
        return True
    if path.suffix.lower() == ".md":
        return True
    if path.suffix.lower() == ".py" and re.search(r"[\u4e00-\u9fff]", text):
        return True
    return None


def _check_encoding(root: Path, errors: list[str]) -> bool:
    ok = True
    for path in _text_files(root):
        relative = _relative(path, root)
        raw = path.read_bytes()
        has_bom = raw.startswith(b"\xef\xbb\xbf")
        try:
            text = raw.decode("utf-8-sig", errors="strict")
        except UnicodeDecodeError as exc:
            ok = False
            errors.append(f"UTF-8 解码失败：{relative}：{exc}")
            continue
        if "\ufffd" in text:
            ok = False
            errors.append(f"出现 Unicode 替换字符：{relative}")
        if re.search(r"\?{3,}", text):
            ok = False
            errors.append(f"出现连续问号疑似乱码：{relative}")
        expected = _expected_bom(path, root, text)
        if expected is not None and has_bom != expected:
            ok = False
            expectation = "应带 BOM" if expected else "不应带 BOM"
            errors.append(f"BOM 策略不符：{relative}：{expectation}")
        if path.suffix.lower() in {".json", ".jsonl"}:
            try:
                if path.suffix.lower() == ".json":
                    json.loads(text)
                else:
                    for line_number, line in enumerate(text.splitlines(), 1):
                        if line.strip():
                            json.loads(line)
            except json.JSONDecodeError as exc:
                ok = False
                errors.append(f"JSON 解析失败：{relative}：{exc}")
    return ok


def _check_metadata(root: Path, errors: list[str]) -> bool:
    skill_path = root / "SKILL.md"
    metadata_path = root / "agents" / "openai.yaml"
    if not skill_path.is_file() or not metadata_path.is_file():
        return False
    skill = skill_path.read_text(encoding="utf-8")
    frontmatter = re.match(r"^---\r?\n(.*?)\r?\n---\r?\n", skill, re.S)
    ok = bool(frontmatter)
    if not frontmatter:
        errors.append("SKILL.md 缺少可识别的 YAML frontmatter")
    else:
        header = frontmatter.group(1)
        if not re.search(r"^name:\s*related-party-identification\s*$", header, re.M):
            ok = False
            errors.append("SKILL.md 的 name 不正确")
        description = re.search(r"^description:\s*(.+)$", header, re.M)
        if not description or not description.group(1).startswith("Use when"):
            ok = False
            errors.append("SKILL.md 的 description 必须以 Use when 开头")

    metadata = metadata_path.read_text(encoding="utf-8")
    short = re.search(r'short_description:\s*"([^"]+)"', metadata)
    if 'display_name: "关联方识别与核查"' not in metadata:
        ok = False
        errors.append("openai.yaml 的显示名称不正确")
    if not short or not 25 <= len(short.group(1)) <= 64:
        ok = False
        errors.append("openai.yaml 的简述长度必须为 25—64 个字符")
    if "$related-party-identification" not in metadata:
        ok = False
        errors.append("openai.yaml 的默认提示未引用技能名称")
    return ok


def _check_provenance(root: Path, errors: list[str]) -> bool:
    path = root / "references" / "SOURCES.json"
    notice_path = root / "NOTICE"
    if not path.is_file() or not notice_path.is_file():
        return False
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        sources = payload["sources"]
    except (json.JSONDecodeError, KeyError, IndexError, TypeError) as exc:
        errors.append(f"来源追踪文件不可用：{exc}")
        return False
    required = (
        "component",
        "author",
        "repository",
        "source_path",
        "baseline_commit",
        "source_snapshot",
        "local_modules",
        "source_mapping",
        "local_adaptation_points",
        "sync_contract",
    )
    ok = True
    if not isinstance(sources, list) or not sources:
        errors.append("SOURCES.json 没有有效来源记录")
        return False
    for source in sources:
        component = str(source.get("component", "未命名来源"))
        for field in required:
            if not source.get(field):
                ok = False
                errors.append(
                    f"SOURCES.json 的 {component} 缺少来源字段：{field}"
                )
        for item in source.get("source_snapshot", []):
            for field in ("path", "sha256", "git_blob"):
                if not item.get(field):
                    ok = False
                    errors.append(
                        f"SOURCES.json 的 {component} 来源文件缺少字段：{field}"
                    )
    by_component = {
        str(source.get("component")): source
        for source in sources
        if isinstance(source, dict)
    }
    cicpa = by_component.get("cicpa-company-query", {})
    opencli = by_component.get("OpenCLI Browser Bridge extension", {})
    if cicpa.get("license_evidence", {}).get("value") != "MIT":
        ok = False
        errors.append("SOURCES.json 未记录 cicpa-company-query 的 MIT 声明")
    if cicpa.get("license_evidence", {}).get("license_file_verified") is not False:
        ok = False
        errors.append("SOURCES.json 不得虚构 cicpa-company-query 的独立 LICENSE 文件")
    if opencli.get("license_evidence", {}).get("value") != "Apache-2.0":
        ok = False
        errors.append("SOURCES.json 未记录 OpenCLI 的 Apache-2.0 许可")
    if opencli.get("license_evidence", {}).get("license_file_verified") is not True:
        ok = False
        errors.append("SOURCES.json 未记录 OpenCLI 已核实的 LICENSE 文件")
    notice = notice_path.read_text(encoding="utf-8-sig")
    for phrase in (
        "nigo",
        "nigo81/nigo-skills",
        "cicpa-company-query",
        "MIT",
        "OpenCLI",
        "gitee.com/github_dep/opencli",
        "Apache-2.0",
        "本地改造",
    ):
        if phrase not in notice:
            ok = False
            errors.append(f"NOTICE 缺少来源说明：{phrase}")
    return ok


def _check_dialog_entries(root: Path, errors: list[str]) -> bool:
    skill_path = root / "SKILL.md"
    flow_path = root / "references" / "user-flow.md"
    if not skill_path.is_file() or not flow_path.is_file():
        return False
    combined = (
        skill_path.read_text(encoding="utf-8")
        + "\n"
        + flow_path.read_text(encoding="utf-8-sig")
    )
    required = (
        "主动发现",
        "名单核查",
        "已有数据",
        "Windows 文件选择窗口",
        "登录好了",
        "候选超过 100 家",
        "每分钟最多 15 次",
        "每小时最多 300 次",
        "不等于最终关联方结论",
        "Firefox",
        "OpenCLI",
        "Gitee 国内镜像",
    )
    forbidden = (
        "F12",
        "开发者工具",
        "复制 Cookie",
        "复制Cookie",
        "手填绝对路径",
        "python3 ",
        "python ",
        "~/" + ".claude/skills/",
        "~/" + ".agents/skills/",
        "extension://",
    )
    ok = True
    for phrase in required:
        if phrase not in combined:
            ok = False
            errors.append(f"对话入口缺少关键说明：{phrase}")
    for phrase in forbidden:
        if phrase in combined:
            ok = False
            errors.append(f"对话入口暴露了不应要求用户执行的操作：{phrase}")
    return ok


def validate_bundle(
    skill_root: Path | str,
    *,
    test_runner: Optional[Callable[[Path], tuple[bool, int, str]]] = None,
) -> BundleValidationReport:
    root = Path(skill_root).resolve()
    errors = []
    checks = {
        "required_files": _check_required_files(root, errors),
        "python_syntax": _check_python_syntax(root, errors),
        "runtime_paths": _check_runtime_paths(root, errors),
        "secret_leaks": _check_secret_leaks(root, errors),
        "encoding": _check_encoding(root, errors),
        "metadata": _check_metadata(root, errors),
        "provenance": _check_provenance(root, errors),
        "dialog_entries": _check_dialog_entries(root, errors),
    }
    runner = test_runner or _default_test_runner
    tests_ok, test_count, test_output = runner(root)
    checks["unit_tests"] = bool(tests_ok)
    if not tests_ok:
        errors.append("全量单元测试未通过")
    ordered_checks = {name: checks.get(name, False) for name in CHECK_NAMES}
    return BundleValidationReport(
        checks=ordered_checks,
        errors=errors,
        test_count=test_count,
        test_output=test_output,
    )


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="关联方识别独立分发校验")
    parser.add_argument(
        "--skill-root",
        default=str(Path(__file__).resolve().parents[1]),
    )
    args = parser.parse_args(argv)
    report = validate_bundle(args.skill_root)
    for name, passed in report.checks.items():
        print(f"[{'PASS' if passed else 'FAIL'}] {name}")
    for error in report.errors:
        print(f"  - {error}")
    print(f"TEST_COUNT={report.test_count}")
    print(f"BUNDLE_VALID={report.is_valid}")
    return 0 if report.is_valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
