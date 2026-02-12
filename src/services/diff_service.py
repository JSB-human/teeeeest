from __future__ import annotations

from difflib import SequenceMatcher
from typing import Any, Dict, List


def build_text_diff_summary(before: str, after: str, max_items: int = 8) -> Dict[str, Any]:
    before = before or ""
    after = after or ""

    sm = SequenceMatcher(None, before, after)
    added = 0
    removed = 0
    changed_spans: List[Dict[str, Any]] = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        old_seg = before[i1:i2]
        new_seg = after[j1:j2]
        if tag in ("replace", "delete"):
            removed += len(old_seg)
        if tag in ("replace", "insert"):
            added += len(new_seg)

        if len(changed_spans) < max_items:
            changed_spans.append(
                {
                    "tag": tag,
                    "old": old_seg[:80],
                    "new": new_seg[:80],
                    "old_range": [i1, i2],
                    "new_range": [j1, j2],
                }
            )

    return {
        "kind": "text",
        "chars_before": len(before),
        "chars_after": len(after),
        "chars_added": added,
        "chars_removed": removed,
        "sample_spans": changed_spans,
    }


def build_table_diff_summary(cells: List[Dict[str, Any]], max_items: int = 12) -> Dict[str, Any]:
    cells = cells or []
    sample = []
    for c in cells[:max_items]:
        sample.append(
            {
                "row": int(c.get("row", 0)),
                "col": int(c.get("col", 0)),
                "old": str(c.get("old", ""))[:80],
                "new": str(c.get("new", ""))[:80],
            }
        )

    return {
        "kind": "table",
        "changed_cells": len(cells),
        "sample_cells": sample,
    }
