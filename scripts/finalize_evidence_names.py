#!/usr/bin/env python3
"""Rename final evidence screenshots to concise checklist evidence names."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


def norm(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\u3000", " ")
    return re.sub(r"\s+", "", text).strip()


def safe_name(text: str, fallback: str, max_len: int = 80) -> str:
    value = norm(text) or fallback
    value = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", value)
    value = value.strip("._ ")
    return (value or fallback)[:max_len].rstrip("._ ")


SHORT_LABEL_BY_CHECK_ID = {
    "row05": "身份鉴别",
    "row06": "口令策略",
    "row10": "失败锁定",
    "row11": "远程管理",
    "row12": "用户唯一性",
    "row14": "文件权限",
    "row17": "默认账户",
    "row18": "多余账户",
    "row19": "审计服务",
    "row20": "审计事件",
    "row21": "审计记录",
    "row22": "服务端口",
    "row23": "登录限制",
    "row24": "超时锁定",
}


def base_label(item: dict[str, Any]) -> str:
    for key in ("evidence_label", "evidenceLabel", "short_label", "shortName"):
        value = safe_name(item.get(key), "", max_len=32)
        if value:
            return value
    check_id = safe_name(item.get("check_id"), "", max_len=32)
    if check_id in SHORT_LABEL_BY_CHECK_ID:
        return SHORT_LABEL_BY_CHECK_ID[check_id]
    source = item.get("source") or {}
    for key in ("item", "expected", "operation", "category"):
        value = safe_name(source.get(key), "", max_len=32)
        if value:
            return value
    return safe_name(item.get("check_id") or item.get("gui_action") or "截图", "截图", max_len=32)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    index = 2
    while True:
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def rename_evidence(plan: dict[str, Any], evidence_dir: Path) -> dict[str, Any]:
    used_names: set[str] = set()
    for item in plan.get("items", []):
        evidence = item.get("evidence") or []
        if not evidence:
            continue
        row = int(item["row"])
        label = base_label(item)
        new_evidence: list[str] = []
        for index, old_name in enumerate(evidence):
            old_path = evidence_dir / Path(str(old_name)).name
            if not old_path.exists():
                new_evidence.append(Path(str(old_name)).name)
                continue
            suffix = "" if index == 0 else f"_补充{index}"
            new_name = f"row{row:02d}_{label}{suffix}{old_path.suffix.lower()}"
            new_name = safe_name(new_name, f"row{row:02d}_截图{suffix}{old_path.suffix.lower()}", max_len=64)
            if not new_name.lower().endswith(old_path.suffix.lower()):
                new_name = f"{new_name}{old_path.suffix.lower()}"
            new_path = unique_path(evidence_dir / new_name)
            old_path.rename(new_path)
            used_names.add(new_path.name)
            new_evidence.append(new_path.name)
        item["evidence"] = new_evidence
    return plan


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", required=True, type=Path)
    parser.add_argument("--evidence-dir", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()

    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    updated = rename_evidence(plan, args.evidence_dir)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")
    renamed = sum(len(item.get("evidence") or []) for item in updated.get("items", []))
    print(json.dumps({"output_plan": str(args.out), "evidence_entries": renamed}, ensure_ascii=False))


if __name__ == "__main__":
    main()
