from __future__ import annotations

import json
from pathlib import Path


def write_json(path: Path, payload) -> Path:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return path
