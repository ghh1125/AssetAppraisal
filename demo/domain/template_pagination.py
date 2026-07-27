from __future__ import annotations

import re
from typing import Any


def _normal(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def _score(context: str, page: str) -> float:
    if not context:
        return 0.0
    if context in page:
        return 1000.0 + min(len(context), 300)
    width = min(10, max(4, len(context) // 8))
    grams = {context[index : index + width] for index in range(max(1, len(context) - width + 1))}
    if not grams:
        return 0.0
    overlap = sum(gram in page for gram in grams) / len(grams)
    prefix = context[: min(24, len(context))]
    return overlap * 100.0 + (25.0 if prefix and prefix in page else 0.0)


def map_location_pages(
    locations: list[dict[str, Any]],
    page_texts: list[str],
    paragraph_texts: list[tuple[int, str]] | None = None,
) -> dict[str, int | str]:
    """将模板位置映射到原 Word 渲染页，重复文本按文档顺序消歧。"""

    if not page_texts:
        return {}
    result: dict[str, int | str] = {}
    document_locations = []
    for location in locations:
        location_id = str(location["location_id"])
        if not location_id.startswith("DOCUMENT-"):
            result[location_id] = "多页页脚" if location_id.startswith("FOOTER") else "多页"
        else:
            document_locations.append(location)

    pages = [_normal(text) for text in page_texts]
    ordered_texts = (
        [(index, text) for index, text in paragraph_texts if _normal(text)]
        if paragraph_texts
        else [
            (int(re.search(r"-P(\d+)-", str(item["location_id"])).group(1)), item.get("context", ""))
            for item in document_locations
        ]
    )
    if not ordered_texts:
        return result
    scores = [[_score(_normal(text), page) for page in pages] for _, text in ordered_texts]
    previous = scores[0][:]
    backtracks: list[list[int]] = [[0] * len(pages)]
    for row in scores[1:]:
        best_value = float("-inf")
        best_page = 0
        current = []
        links = []
        for page_index, value in enumerate(row):
            if previous[page_index] > best_value:
                best_value = previous[page_index]
                best_page = page_index
            current.append(best_value + value)
            links.append(best_page)
        previous = current
        backtracks.append(links)

    selected = max(range(len(pages)), key=previous.__getitem__)
    assignments = [selected]
    for index in range(len(ordered_texts) - 1, 0, -1):
        selected = backtracks[index][selected]
        assignments.append(selected)
    assignments.reverse()
    page_by_paragraph = {
        paragraph_index: page_index + 1
        for (paragraph_index, _), page_index in zip(ordered_texts, assignments, strict=True)
    }
    known_indices = sorted(page_by_paragraph)
    for item in document_locations:
        match = re.search(r"-P(\d+)-", str(item["location_id"]))
        paragraph_index = (
            int(match.group(1))
            if match
            else int(item.get("paragraph_index", 0))
        )
        if paragraph_index <= 0:
            continue
        if paragraph_index in page_by_paragraph:
            page = page_by_paragraph[paragraph_index]
        else:
            nearest = min(known_indices, key=lambda value: abs(value - paragraph_index))
            page = page_by_paragraph[nearest]
        result[str(item["location_id"])] = page
    return result
