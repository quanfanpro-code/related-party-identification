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
#import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from scripts.sync_upstream import (
    check_upstream,
    record_verified_update,
    stage_upstream,
    verify_staged_update,
)


def sha256_text(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def git_blob_text(text):
    data = text.encode("utf-8")
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()


def write_sources(path, source_text, commit="commit-old"):
    payload = {
        "schema_version": 1,
        "runtime_self_contained": True,
        "sources": [
            {
                "component": "cicpa-company-query",
                "author": "nigo",
                "repository": "https://github.com/nigo81/nigo-skills",
                "source_path": "cicpa-company-query",
                "baseline_commit": commit,
                "source_snapshot": [
                    {
                        "path": "cicpa-company-query/scripts/cicpa_query.py",
                        "sha256": sha256_text(source_text),
                        "git_blob": git_blob_text(source_text),
                    }
                ],
                "local_modules": [
                    "scripts/cicpa/client.py",
                    "scripts/cicpa/auth.py",
                    "scripts/cicpa/exporter.py",
                ],
                "adaptation_note_zh": "本地改造",
            }
        ],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def make_snapshot(root, source_text):
    source_dir = root / "snapshot"
    script = source_dir / "scripts" / "cicpa_query.py"
    script.parent.mkdir(parents=True)
    script.write_bytes(source_text.encode("utf-8"))
    return source_dir


class UpstreamCheckTests(unittest.TestCase):
    def test_check_比较提交_git_blob_sha256_和来源映射(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_check_"))
        source_text = "print('upstream')\n"
        snapshot = make_snapshot(artifact_dir, source_text)
        sources_path = artifact_dir / "SOURCES.json"
        write_sources(sources_path, source_text)

        current = check_upstream(snapshot, sources_path, commit="commit-old")
        changed = check_upstream(snapshot, sources_path, commit="commit-new")

        self.assertTrue(current.is_current)
        self.assertEqual(current.files[0].sha256_status, "same")
        self.assertEqual(current.files[0].git_blob_status, "same")
        self.assertTrue(current.source_mapping_present)
        self.assertFalse(changed.is_current)
        self.assertEqual(changed.commit_status, "changed")

    def test_check_可以单独比较_opencli_来源(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_opencli_"))
        source_text = '{"name":"OpenCLI"}\n'
        snapshot = artifact_dir / "extension"
        snapshot.mkdir()
        (snapshot / "manifest.json").write_bytes(source_text.encode("utf-8"))
        sources_path = artifact_dir / "SOURCES.json"
        payload = {
            "schema_version": 1,
            "sources": [
                {
                    "component": "OpenCLI Browser Bridge extension",
                    "author": "jackwener",
                    "repository": "https://gitee.com/github_dep/opencli",
                    "source_path": "extension",
                    "baseline_commit": "opencli-old",
                    "source_snapshot": [
                        {
                            "path": "extension/manifest.json",
                            "sha256": sha256_text(source_text),
                            "git_blob": git_blob_text(source_text),
                        }
                    ],
                    "local_modules": ["scripts/cicpa/edge_bridge.py"],
                }
            ],
        }
        sources_path.write_text(
            json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )

        report = check_upstream(
            snapshot,
            sources_path,
            commit="opencli-old",
            component="OpenCLI Browser Bridge extension",
        )

        self.assertTrue(report.is_current)
        self.assertEqual(report.repository, "https://gitee.com/github_dep/opencli")


class UpstreamStageTests(unittest.TestCase):
    def test_stage_只复制到技能外暂存目录且不覆盖正式文件(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_stage_"))
        skill_root = artifact_dir / "skill"
        official = skill_root / "scripts" / "cicpa" / "client.py"
        official.parent.mkdir(parents=True)
        official.write_text("official\n", encoding="utf-8")
        source_text = "print('new upstream')\n"
        snapshot = make_snapshot(artifact_dir, source_text)
        sources_path = skill_root / "references" / "SOURCES.json"
        sources_path.parent.mkdir(parents=True)
        write_sources(sources_path, "print('old upstream')\n")
        staging_dir = artifact_dir / "outside-stage"

        stage = stage_upstream(
            snapshot,
            staging_dir,
            sources_path,
            skill_root=skill_root,
            commit="commit-new",
        )

        self.assertEqual(official.read_text(encoding="utf-8"), "official\n")
        self.assertTrue((staging_dir / "upstream" / "scripts" / "cicpa_query.py").exists())
        self.assertNotIn(skill_root.resolve(), staging_dir.resolve().parents)
        self.assertEqual(stage.changed_files, ["cicpa-company-query/scripts/cicpa_query.py"])

    def test_stage_拒绝把暂存目录放进正式_skill(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_inside_"))
        skill_root = artifact_dir / "skill"
        snapshot = make_snapshot(artifact_dir, "new\n")
        sources_path = skill_root / "references" / "SOURCES.json"
        sources_path.parent.mkdir(parents=True)
        write_sources(sources_path, "old\n")

        with self.assertRaises(ValueError):
            stage_upstream(
                snapshot,
                skill_root / "stage",
                sources_path,
                skill_root=skill_root,
                commit="commit-new",
            )


class UpstreamVerifyTests(unittest.TestCase):
    def _stage_sensitive_update(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_verify_"))
        skill_root = artifact_dir / "skill"
        for relative in [
            "scripts/cicpa/client.py",
            "scripts/cicpa/auth.py",
            "scripts/cicpa/exporter.py",
        ]:
            path = skill_root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(f"{relative}\n", encoding="utf-8")
        sources_path = skill_root / "references" / "SOURCES.json"
        sources_path.parent.mkdir(parents=True)
        write_sources(sources_path, "old\n")
        source_text = (
            "login_url = '/login'\n"
            "cookie_name = 'session'\n"
            "api_url = '/api/company/detail'\n"
            "dimensions = ['客户.xlsx']\n"
        )
        snapshot = make_snapshot(artifact_dir, source_text)
        staging_dir = artifact_dir / "outside-stage"
        stage_upstream(
            snapshot,
            staging_dir,
            sources_path,
            skill_root=skill_root,
            commit="commit-new",
        )
        return skill_root, sources_path, staging_dir

    def test_敏感接口变化会阻断登记(self):
        skill_root, sources_path, staging_dir = self._stage_sensitive_update()

        report = verify_staged_update(
            staging_dir,
            skill_root,
            sources_path,
            checks_runner=lambda _root: {
                "unit_tests": True,
                "runtime_paths": True,
                "encoding": True,
                "secret_leaks": True,
            },
        )

        self.assertIn("登录与凭据接口发生变化，需要人工复核", report.blockers)
        self.assertIn("API 端点或请求协议发生变化，需要人工复核", report.blockers)
        self.assertIn("导出维度或文件命名发生变化，需要人工复核", report.blockers)
        self.assertIn(
            "本地用户交互、安全、限速和报告语义需确认保持不变",
            report.blockers,
        )
        self.assertFalse(report.can_record)

    def test_验证过程不覆盖本地保护层(self):
        skill_root, sources_path, staging_dir = self._stage_sensitive_update()
        protected = [
            skill_root / "scripts" / "cicpa" / "client.py",
            skill_root / "scripts" / "cicpa" / "auth.py",
            skill_root / "scripts" / "cicpa" / "exporter.py",
        ]
        before = {path: path.read_bytes() for path in protected}

        verify_staged_update(
            staging_dir,
            skill_root,
            sources_path,
            checks_runner=lambda _root: {
                "unit_tests": True,
                "runtime_paths": True,
                "encoding": True,
                "secret_leaks": True,
            },
        )

        self.assertEqual(before, {path: path.read_bytes() for path in protected})


class UpstreamRecordTests(unittest.TestCase):
    def test_只有全部验证通过后才能登记(self):
        artifact_dir = Path(tempfile.mkdtemp(prefix="rpi_sync_record_"))
        skill_root = artifact_dir / "skill"
        local_module = skill_root / "scripts" / "cicpa" / "client.py"
        local_module.parent.mkdir(parents=True)
        local_module.write_text("official merged\n", encoding="utf-8")
        sources_path = skill_root / "references" / "SOURCES.json"
        sources_path.parent.mkdir(parents=True)
        source_text = "print('same')\n"
        write_sources(sources_path, source_text)
        snapshot = make_snapshot(artifact_dir, source_text)
        staging_dir = artifact_dir / "outside-stage"
        stage_upstream(
            snapshot,
            staging_dir,
            sources_path,
            skill_root=skill_root,
            commit="commit-old",
        )

        with self.assertRaises(RuntimeError):
            record_verified_update(staging_dir, sources_path, skill_root=skill_root)

        report = verify_staged_update(
            staging_dir,
            skill_root,
            sources_path,
            merged_files=["scripts/cicpa/client.py"],
            checks_runner=lambda _root: {
                "unit_tests": True,
                "runtime_paths": True,
                "encoding": True,
                "secret_leaks": True,
            },
        )
        result = record_verified_update(
            staging_dir,
            sources_path,
            skill_root=skill_root,
            backup_root=artifact_dir / "BackUp",
        )
        recorded = json.loads(sources_path.read_text(encoding="utf-8"))

        self.assertTrue(report.can_record)
        self.assertEqual(result["status"], "recorded")
        self.assertEqual(recorded["sources"][0]["last_synced_commit"], "commit-old")
        self.assertIn("last_synced_at", recorded["sources"][0])
        self.assertTrue(Path(result["backup_path"]).exists())


if __name__ == "__main__":
    unittest.main()
