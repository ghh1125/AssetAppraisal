from __future__ import annotations

from typing import Any


def validate_mapping(data: dict[str, Any]) -> list[dict[str, Any]]:
    locations = data.get("locations", [])
    ids = [item.get("location_id") for item in locations]
    if len(ids) != len(set(ids)) or any(not item for item in ids):
        raise ValueError("映射位置编号缺失或重复")
    return locations
