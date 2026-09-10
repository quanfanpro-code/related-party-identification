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
"""旧股权穿透入口的自包含兼容包装。"""

import argparse
from pathlib import Path
import sys

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.cicpa.auth import auth_status
from scripts.cicpa.browser_transport import OpenCliTransport
from scripts.cicpa.client import CicpaClient
from scripts.discovery import DiscoveryPolicy, discover


def penetrate(target, threshold=20.0, max_depth=2):
    """调用统一发现引擎并返回旧入口需要的列表结构。"""
    status = auth_status()
    if status.get("status") != "authenticated":
        raise RuntimeError("尚未登录，请先通过 AI 对话启动注协官方登录流程")
    result = discover(
        target,
        client=CicpaClient(session=OpenCliTransport()),
        policy=DiscoveryPolicy(equity_threshold=threshold),
        approved_depth=max_depth,
    )
    if result.status == "needs_company_confirmation":
        raise RuntimeError("企业名称无法唯一匹配，需要先确认具体主体")
    if result.status == "needs_user_confirmation":
        raise RuntimeError("候选达到 100 家，需要用户确认后继续")
    return [
        {
            "name": candidate.name,
            "ratio": candidate.ratio,
            "depth": candidate.depth,
            "parent": candidate.parent_name,
            "org_id": candidate.company_id,
            "reasons": list(candidate.reasons),
        }
        for candidate in result.candidates.values()
    ]


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="股权穿透：发现被审计单位的关联方候选"
    )
    parser.add_argument("--target", required=True, help="被审计单位全称")
    parser.add_argument(
        "--depth",
        type=int,
        default=2,
        help="初始穿透深度（默认 2 层，明确确认后最多 5 层）",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=20.0,
        help="股权扩展阈值（默认 20%%）",
    )
    parser.add_argument("-o", "--output", default="核查名单.txt", help="输出核查名单文件")
    args = parser.parse_args(argv)
    if args.depth < 1 or args.depth > 5:
        parser.error("穿透深度只能是 1 至 5 层")

    candidates = penetrate(args.target, args.threshold, args.depth)
    names = [args.target] + [item["name"] for item in candidates]
    unique = list(dict.fromkeys(names))
    output = Path(args.output)
    with open(output, "w", encoding="utf-8-sig", newline="\n") as stream:
        for name in unique:
            stream.write(name + "\n")
    print("发现 {} 家候选，核查名单已输出：{}".format(len(candidates), output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
