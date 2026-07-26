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
#"""从被审计单位出发，分层发现可核查的关联方候选。"""

from collections import deque
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from scripts.cicpa.client import CompanyMatch, EquityEdge


@dataclass(frozen=True)
class DiscoveryPolicy:
    equity_threshold: float = 20.0
    initial_depth: int = 2
    maximum_depth: int = 5
    candidate_cap: int = 100

    def __post_init__(self):
        if self.equity_threshold < 0:
            raise ValueError("股权阈值不能小于 0")
        if self.initial_depth < 1 or self.maximum_depth < self.initial_depth:
            raise ValueError("穿透层数设置无效")
        if self.candidate_cap < 1:
            raise ValueError("候选上限必须大于 0")


@dataclass(frozen=True)
class SeedExportCandidate:
    name: str
    relation_type: str
    red_flags: tuple = ()


@dataclass
class Candidate:
    name: str
    company_id: str = ""
    depth: int = 1
    parent_name: str = ""
    relation_type: str = ""
    ratio: Optional[float] = None
    reasons: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)
    paths: List[str] = field(default_factory=list)

    def merge_evidence(self, reasons: List[str], notes: List[str], path: str) -> None:
        for value in reasons:
            if value and value not in self.reasons:
                self.reasons.append(value)
        for value in notes:
            if value and value not in self.notes:
                self.notes.append(value)
        if path and path not in self.paths:
            self.paths.append(path)


@dataclass
class DiscoveryResult:
    status: str
    seed_name: str
    candidates: Dict[str, Candidate] = field(default_factory=dict)
    company_matches: List[CompanyMatch] = field(default_factory=list)
    stopped_at_cap: bool = False
    explored_depth: int = 0
    message_zh: str = ""
    warnings: List[str] = field(default_factory=list)


def _company_key(company_id: str, name: str) -> str:
    return str(company_id).strip() or "".join(str(name).split())


def _relation_label(edge: EquityEdge) -> str:
    return "对外投资" if edge.direction == "investment" else "股东关系"


def _person_names(client, company_id: str, warnings: List[str]) -> set:
    if not company_id:
        return set()
    try:
        return {
            person.name.strip()
            for person in client.get_key_personnel(company_id)
            if person.name and person.name.strip()
        }
    except Exception:
        warnings.append("主要人员数据未完整取得：{}".format(company_id))
        return set()


def discover(
    seed_name: str,
    *,
    client,
    policy: DiscoveryPolicy = DiscoveryPolicy(),
    approved_depth: Optional[int] = None,
    seed_export_candidates: Optional[List[SeedExportCandidate]] = None,
) -> DiscoveryResult:
    """执行初始两层或用户批准后的最多五层广度优先发现。"""
    name = str(seed_name).strip()
    if not name:
        raise ValueError("被审计单位名称不能为空")
    depth_limit = policy.initial_depth if approved_depth is None else approved_depth
    if depth_limit < 1 or depth_limit > policy.maximum_depth:
        raise ValueError("批准的穿透层数超出允许范围")

    matches = client.search_companies(name)
    exact = [match for match in matches if match.name == name]
    if len(exact) == 1:
        seed = exact[0]
    elif len(matches) == 1:
        seed = matches[0]
    else:
        return DiscoveryResult(
            status="needs_company_confirmation",
            seed_name=name,
            company_matches=matches,
            message_zh="企业名称无法唯一匹配，需要用户确认具体主体",
        )

    result = DiscoveryResult(
        status="completed",
        seed_name=seed.name,
        company_matches=[seed],
        explored_depth=depth_limit,
        message_zh="主动发现已完成",
    )

    def add_candidate(candidate: Candidate, reasons: List[str], notes: List[str], path: str):
        existing = result.candidates.get(candidate.name)
        if existing is not None:
            existing.merge_evidence(reasons, notes, path)
            existing.depth = min(existing.depth, candidate.depth)
            return existing
        if len(result.candidates) >= policy.candidate_cap:
            result.status = "needs_user_confirmation"
            result.stopped_at_cap = True
            result.message_zh = "候选已达到 {} 家，需用户确认后继续".format(
                policy.candidate_cap
            )
            return None
        candidate.merge_evidence(reasons, notes, path)
        result.candidates[candidate.name] = candidate
        return candidate

    for item in seed_export_candidates or []:
        candidate_name = str(item.name).strip()
        if not candidate_name or candidate_name == seed.name:
            continue
        reasons = ["被审计单位完整维度导出：{}".format(item.relation_type)]
        reasons.extend(str(flag) for flag in item.red_flags if str(flag).strip())
        added = add_candidate(
            Candidate(
                name=candidate_name,
                depth=1,
                parent_name=seed.name,
                relation_type=item.relation_type,
            ),
            reasons,
            [],
            "{} --{}--> {}".format(seed.name, item.relation_type, candidate_name),
        )
        if added is None:
            return result

    seed_people = _person_names(client, seed.org_id, result.warnings)
    visited = {_company_key(seed.org_id, seed.name)}
    queue = deque([(seed, 0, [seed.name])])

    while queue and result.status == "completed":
        parent, parent_depth, parent_path = queue.popleft()
        if parent_depth >= depth_limit:
            continue
        try:
            edges = client.get_equity_relations(parent.org_id)
        except Exception:
            result.warnings.append("股权关系未完整取得：{}".format(parent.name))
            continue

        for edge in edges:
            candidate_name = str(edge.company_name).strip()
            if not candidate_name or candidate_name == seed.name:
                continue
            candidate_depth = parent_depth + 1
            label = _relation_label(edge)
            path = "{} --{} {}%--> {}".format(
                " → ".join(parent_path),
                label,
                edge.ratio,
                candidate_name,
            )
            reasons = ["{}持股比例 {}%".format(label, edge.ratio)]
            notes = []
            meets_threshold = edge.ratio >= policy.equity_threshold
            if not meets_threshold:
                candidate_people = _person_names(
                    client,
                    edge.company_id,
                    result.warnings,
                )
                shared_people = sorted(seed_people & candidate_people)
                if not shared_people:
                    continue
                reasons.extend(
                    "共同关键人员：{}".format(person) for person in shared_people
                )
                notes.append("低持股比例但存在其他红旗")

            added = add_candidate(
                Candidate(
                    name=candidate_name,
                    company_id=edge.company_id,
                    depth=candidate_depth,
                    parent_name=parent.name,
                    relation_type=label,
                    ratio=edge.ratio,
                ),
                reasons,
                notes,
                path,
            )
            if added is None:
                break

            key = _company_key(edge.company_id, candidate_name)
            if meets_threshold and key not in visited:
                visited.add(key)
                queue.append(
                    (
                        CompanyMatch(org_id=edge.company_id, name=candidate_name),
                        candidate_depth,
                        parent_path + [candidate_name],
                    )
                )

    return result
