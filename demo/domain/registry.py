from __future__ import annotations

from typing import Any


def human_fill(field_name: str) -> str:
    return ""


def resolve_candidate(candidates: list[dict[str, Any]], priority: list[str], field_name: str) -> Any:
    usable = [item for item in candidates if item.get("value") not in (None, "", [])]
    ranks = {name: index for index, name in enumerate(priority)}
    usable.sort(key=lambda item: ranks.get(item.get("source", ""), len(ranks)))
    return usable[0]["value"] if usable else human_fill(field_name)


def resolve_all(specs: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        item["field_key"]: resolve_candidate(item.get("candidates", []), item.get("priority", []), item["field_name"])
        for item in specs
    }
