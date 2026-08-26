from __future__ import annotations

import re
import tempfile
import zipfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.request import Request, urlopen

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
PLACEHOLDER = re.compile(r"X{2,}", re.I)
UNRESOLVED_MARKER = re.compile(r"20XX|X{2,}", re.I)
PART_RE = re.compile(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml")
COMMENTS_PART = "word/comments.xml"
DOCUMENT_RELS_PART = "word/_rels/document.xml.rels"
CONTENT_TYPES_PART = "[Content_Types].xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"


def _paragraph_text(paragraph) -> str:
    parts: list[str] = []
    for node in paragraph.iter():
        if node.tag in {f"{{{W}}}t", f"{{{W}}}delText"} and node.text:
            parts.append(node.text)
        elif node.tag == f"{{{W}}}tab":
            parts.append("\t")
        elif node.tag in {f"{{{W}}}br", f"{{{W}}}cr"}:
            parts.append("\n")
    return "".join(parts)


def _highlight_text(paragraph) -> str:
    values: list[str] = []
    for run in paragraph.xpath(".//w:r", namespaces=NS):
        highlight = run.find("w:rPr/w:highlight", namespaces=NS)
        if highlight is not None and highlight.get(f"{{{W}}}val") not in (None, "none", "auto"):
            values.append("".join(run.xpath(".//w:t/text()", namespaces=NS)))
    return "".join(values).strip()


def _paragraph_text_without_highlight(paragraph) -> str:
    parts: list[str] = []
    for run in paragraph.xpath(".//w:r", namespaces=NS):
        highlight = run.find("w:rPr/w:highlight", namespaces=NS)
        if highlight is not None and highlight.get(f"{{{W}}}val") not in (None, "none", "auto"):
            run_text = "".join(run.xpath(".//w:t/text()|.//w:delText/text()", namespaces=NS))
            parts.extend(match.group(0) for match in PLACEHOLDER.finditer(run_text))
            continue
        for node in run.iter():
            if node.tag in {f"{{{W}}}t", f"{{{W}}}delText"} and node.text:
                parts.append(node.text)
            elif node.tag == f"{{{W}}}tab":
                parts.append("\t")
            elif node.tag in {f"{{{W}}}br", f"{{{W}}}cr"}:
                parts.append("\n")
    text = "".join(parts)
    return re.sub(r"（\s*）|\(\s*\)", "", text)


def _comment_texts(archive: zipfile.ZipFile) -> dict[str, str]:
    """Read Word comments without requiring python-docx comment support.

    Word stores comments in a separate OOXML part and anchors them in the
    document with ``commentRangeStart``/``commentReference`` elements.  The
    communication template now uses those comments as the source-of-truth
    annotations, so inventory records retain the exact comment text and IDs.
    Older yellow-highlight templates simply return an empty mapping.
    """
    if COMMENTS_PART not in archive.namelist():
        return {}
    root = etree.fromstring(archive.read(COMMENTS_PART))
    return {
        str(comment.get(f"{{{W}}}id")): "".join(
            comment.xpath(".//w:t/text()", namespaces=NS)
        ).strip()
        for comment in root.xpath(".//w:comment", namespaces=NS)
        if comment.get(f"{{{W}}}id") is not None
    }


def extract_word_comments(path: Path) -> dict[str, str]:
    """Return all Word comments keyed by OOXML comment ID."""
    with zipfile.ZipFile(path) as archive:
        return _comment_texts(archive)


def _paragraph_comment_ids(paragraph) -> list[str]:
    ids: list[str] = []
    for node in paragraph.xpath(
        ".//w:commentRangeStart|.//w:commentReference",
        namespaces=NS,
    ):
        value = node.get(f"{{{W}}}id")
        if value is not None and value not in ids:
            ids.append(str(value))
    return ids


def inventory_template(path: Path) -> list[dict]:
    records: list[dict] = []
    with zipfile.ZipFile(path) as zf:
        comments = _comment_texts(zf)
        for part in sorted(name for name in zf.namelist() if PART_RE.fullmatch(name)):
            root = etree.fromstring(zf.read(part))
            short = Path(part).stem.upper()
            for p_index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS), 1):
                raw = _paragraph_text(paragraph)
                context = re.sub(r"\s+", " ", raw).strip()
                markers = list(PLACEHOLDER.finditer(context))
                highlight = _highlight_text(paragraph)
                in_table = bool(paragraph.xpath("ancestor::w:tbl", namespaces=NS))
                comment_ids = _paragraph_comment_ids(paragraph)
                comment_texts = [comments[item] for item in comment_ids if item in comments]
                for occurrence, match in enumerate(markers, 1):
                    records.append({
                        "location_id": f"{short}-P{p_index:04d}-X{occurrence:02d}",
                        "record_type": "占位符",
                        "part": part,
                        "paragraph_index": p_index,
                        "occurrence_index": occurrence,
                        "marker": match.group(0),
                        "context": context,
                        "in_table": in_table,
                        "comment_ids": comment_ids,
                        "comment_texts": comment_texts,
                    })
                if highlight and not markers:
                    records.append({
                        "location_id": f"{short}-P{p_index:04d}-H01",
                        "record_type": "黄色标注内容块",
                        "part": part,
                        "paragraph_index": p_index,
                        "occurrence_index": 1,
                        "marker": "黄色标注",
                        "context": context,
                        "in_table": in_table,
                        "comment_ids": comment_ids,
                        "comment_texts": comment_texts,
                    })
                # Newer communication templates use comments on headings and
                # table lead-ins that contain no literal XXX marker.  Keep
                # those anchors in the inventory so their source instruction
                # is still available for mapping/audit (the final clean
                # template simply ignores the synthetic C location during
                # replacement).
                if comment_ids and not markers and not highlight:
                    records.append({
                        "location_id": f"{short}-P{p_index:04d}-C01",
                        "record_type": "批注内容块",
                        "paragraph_index": p_index,
                        "occurrence_index": 1,
                        "marker": "",
                        "context": context,
                        "in_table": in_table,
                        "comment_ids": comment_ids,
                        "comment_texts": comment_texts,
                    })
    return records


def document_paragraph_texts(path: Path) -> list[tuple[int, str]]:
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    return [
        (index, _paragraph_text(paragraph))
        for index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS), 1)
    ]


def _is_highlighted(run) -> bool:
    highlight = run.find("w:rPr/w:highlight", namespaces=NS)
    return highlight is not None and highlight.get(f"{{{W}}}val") not in (None, "none", "auto")


def _run_text_nodes(run):
    return run.xpath(".//w:t|.//w:delText", namespaces=NS)


def _run_text(run) -> str:
    return "".join(node.text or "" for node in _run_text_nodes(run))


def _clear_run_text(run) -> None:
    for node in _run_text_nodes(run):
        node.text = ""


def _remove_highlight(run) -> None:
    for highlight in run.xpath("./w:rPr/w:highlight", namespaces=NS):
        highlight.getparent().remove(highlight)


def _set_yellow_highlight(run) -> None:
    properties = run.find("w:rPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(f"{{{W}}}rPr")
        run.insert(0, properties)
    for highlight in properties.findall("w:highlight", namespaces=NS):
        properties.remove(highlight)
    highlight = etree.SubElement(properties, f"{{{W}}}highlight")
    highlight.set(f"{{{W}}}val", "yellow")


def _remove_paragraph_highlights(paragraph) -> None:
    for highlight in paragraph.xpath(".//w:highlight", namespaces=NS):
        parent = highlight.getparent()
        if parent is not None:
            parent.remove(highlight)


def _set_run_text(run, value: str) -> None:
    # A literal newline inside ``w:t`` is rendered as whitespace by Word/WPS,
    # not as a visible line break.  The narrative body deliberately contains
    # one line per accepted module, so write real ``w:br`` nodes while keeping
    # the source run's font, size and indentation properties intact.
    for node in _run_text_nodes(run):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    for index, line in enumerate(str(value).split("\n")):
        if index:
            etree.SubElement(run, f"{{{W}}}br")
        text = etree.SubElement(run, f"{{{W}}}t")
        if line[:1].isspace() or line[-1:].isspace():
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = line


def _normal_runs(paragraph):
    return [run for run in paragraph.xpath("./w:r", namespaces=NS) if not _is_highlighted(run)]


def _set_paragraph_text(
    paragraph,
    value: str,
    remove_highlight: bool = False,
    *,
    style_run=None,
) -> None:
    runs = paragraph.xpath("./w:r", namespaces=NS)
    if not runs:
        style_run = etree.SubElement(paragraph, f"{{{W}}}r")
        etree.SubElement(style_run, f"{{{W}}}t")
        runs = [style_run]
    if style_run is None:
        normal = _normal_runs(paragraph)
        style_run = normal[0] if normal else runs[0]
    _set_run_text(style_run, value)
    for run in runs:
        if run is not style_run:
            _clear_run_text(run)
    if remove_highlight:
        _remove_paragraph_highlights(paragraph)


def _visible_chunks(paragraph):
    return [(_run_text(run), _is_highlighted(run), run) for run in paragraph.xpath("./w:r", namespaces=NS)]


def _strip_yellow_annotation(paragraph) -> str:
    """Remove yellow instructions and only their adjacent annotation parentheses."""
    chunks = _visible_chunks(paragraph)
    visible: list[str] = []
    for index, (text, highlighted, _run) in enumerate(chunks):
        if highlighted:
            continue
        if index > 0 and chunks[index - 1][1]:
            text = re.sub(r"^\s*[）)]", "", text)
        if index + 1 < len(chunks) and chunks[index + 1][1]:
            text = re.sub(r"[（(]\s*$", "", text)
        visible.append(text)
    return re.sub(r"（\s*）|\(\s*\)", "", "".join(visible))


def _replace_yellow_annotation(paragraph, value: str) -> None:
    chunks = _visible_chunks(paragraph)
    first_highlight = next((i for i, (_, highlighted, _) in enumerate(chunks) if highlighted), None)
    if first_highlight is None:
        _set_paragraph_text(paragraph, value, True)
        return
    prefix_parts: list[str] = []
    suffix_parts: list[str] = []
    for index, (text, highlighted, _run) in enumerate(chunks):
        if highlighted:
            continue
        if index < first_highlight:
            if index + 1 < len(chunks) and chunks[index + 1][1]:
                text = re.sub(r"[（(]\s*$", "", text)
            prefix_parts.append(text)
        else:
            if index > 0 and chunks[index - 1][1]:
                text = re.sub(r"^\s*[）)]", "", text)
            suffix_parts.append(text)
    prefix = "".join(prefix_parts)
    suffix = "".join(suffix_parts)
    if value and prefix and not prefix.endswith(("：", ":", "；", ";", "。")):
        prefix += "："
    target = next((run for _text, highlighted, run in chunks if highlighted), None)
    _set_paragraph_text(paragraph, prefix + str(value) + suffix, True, style_run=target)


def _replace_method_heading(paragraph, value: str) -> None:
    runs = paragraph.xpath("./w:r", namespaces=NS)
    normal = [run for run in runs if not _is_highlighted(run)]
    if len(normal) >= 2:
        _set_run_text(normal[1], str(value))
        for run in runs:
            if _is_highlighted(run) or run not in normal[:2]:
                _clear_run_text(run)
                _remove_highlight(run)
        _remove_paragraph_highlights(paragraph)
    else:
        _set_paragraph_text(paragraph, str(value), True)


def _collapse_duplicate_company_suffixes(text: str) -> str:
    for before, after in (
        ("有限公司有限责任公司", "有限公司"),
        ("有限责任公司有限公司", "有限责任公司"),
        ("有限责任公司有限责任公司", "有限责任公司"),
        ("有限公司有限公司", "有限公司"),
    ):
        while before in text:
            text = text.replace(before, after)
    return text


def _replace_placeholders_preserving_runs(paragraph, items, replacements) -> None:
    ordered = sorted(items, key=lambda item: item["occurrence_index"])
    runs = paragraph.xpath("./w:r", namespaces=NS)
    paragraph_text = "".join(_run_text(run) for run in runs)
    values = iter(str(replacements[item["location_id"]]) for item in ordered)
    yuan_uppercase = (
        "人民币" in paragraph_text
        and "万元整" in paragraph_text
        and any(str(replacements.get(item["location_id"], "")).endswith("元") for item in ordered)
    )
    for index, run in enumerate(runs):
        original = _run_text(run)
        if not original:
            continue
        matches = list(PLACEHOLDER.finditer(original))
        if not matches:
            continue
        parts: list[str] = []
        cursor = 0
        for match in matches:
            parts.extend((original[cursor:match.start()], next(values, match.group(0))))
            cursor = match.end()
        parts.append(original[cursor:])
        if index > 0 and _is_highlighted(runs[index - 1]):
            parts[0] = re.sub(r"^\s*[）)]", "", parts[0])
        if index + 1 < len(runs) and _is_highlighted(runs[index + 1]):
            parts[-1] = re.sub(r"[（(]\s*$", "", parts[-1])
        _set_run_text(run, _collapse_duplicate_company_suffixes("".join(parts)))
    for index, run in enumerate(runs):
        if _is_highlighted(run):
            continue
        text = _run_text(run)
        if index > 0 and _is_highlighted(runs[index - 1]):
            text = re.sub(r"^\s*[）)]", "", text)
        if index + 1 < len(runs) and _is_highlighted(runs[index + 1]):
            text = re.sub(r"[（(]\s*$", "", text)
        _set_run_text(run, _collapse_duplicate_company_suffixes(text))
    for run in runs:
        if _is_highlighted(run):
            _clear_run_text(run)
    if yuan_uppercase:
        replacement_values = [str(replacements.get(item["location_id"], "")) for item in ordered]
        suffix = "整" if any(value.endswith("元") for value in replacement_values) else "元整"
        for run in runs:
            if not _is_highlighted(run):
                _set_run_text(run, _run_text(run).replace("万元整", suffix, 1))
    _remove_paragraph_highlights(paragraph)


def _set_cell_text(cell, value: str) -> None:
    paragraphs = cell.xpath("./w:p", namespaces=NS)
    if not paragraphs:
        paragraphs = [etree.SubElement(cell, f"{{{W}}}p")]
    _set_paragraph_text(paragraphs[0], value, True)
    for paragraph in paragraphs[1:]:
        cell.remove(paragraph)


def _split_paragraph_and_highlight_markers(paragraph) -> bool:
    runs = paragraph.xpath(".//w:r", namespaces=NS)
    run_texts = [_run_text(run) for run in runs]
    combined = "".join(run_texts)
    matches = list(UNRESOLVED_MARKER.finditer(combined))
    if not matches:
        return False
    offset = 0
    for run, text in zip(runs, run_texts, strict=True):
        start = offset
        end = offset + len(text)
        offset = end
        if not text:
            continue
        boundaries = {0, len(text)}
        for match in matches:
            overlap_start = max(start, match.start())
            overlap_end = min(end, match.end())
            if overlap_start < overlap_end:
                boundaries.add(overlap_start - start)
                boundaries.add(overlap_end - start)
        ordered = sorted(boundaries)
        chunks = [
            (text[left:right], start + left, start + right)
            for left, right in zip(ordered, ordered[1:])
            if left < right
        ]
        parent = run.getparent()
        if parent is None:
            continue
        position = parent.index(run)
        for value, global_start, global_end in chunks:
            clone = deepcopy(run)
            _set_run_text(clone, value)
            highlighted = any(
                match.start() <= global_start and global_end <= match.end()
                for match in matches
            )
            if highlighted:
                _set_yellow_highlight(clone)
            else:
                _remove_highlight(clone)
            parent.insert(position, clone)
            position += 1
        parent.remove(run)
    return True


def _table_coordinates(root, paragraph) -> tuple[int, int, int] | None:
    cells = paragraph.xpath("ancestor::w:tc[1]", namespaces=NS)
    tables = paragraph.xpath("ancestor::w:tbl[1]", namespaces=NS)
    if not cells or not tables:
        return None
    cell = cells[0]
    table = tables[0]
    all_tables = root.xpath(".//w:tbl", namespaces=NS)
    table_index = next(
        (index for index, candidate in enumerate(all_tables, 1) if candidate is table),
        None,
    )
    if table_index is None:
        return None
    rows = table.xpath("./w:tr", namespaces=NS)
    for row_index, row in enumerate(rows, 1):
        row_cells = row.xpath("./w:tc", namespaces=NS)
        for column_index, candidate in enumerate(row_cells, 1):
            if candidate is cell:
                return table_index, row_index, column_index
    return None


def highlight_unresolved_placeholders(path: Path) -> list[dict[str, Any]]:
    """Highlight unresolved markers and return exact Word-part locations."""
    with zipfile.ZipFile(path) as archive:
        items = archive.infolist()
        contents = {info.filename: archive.read(info.filename) for info in items}

    findings: list[dict[str, Any]] = []
    changed = False
    for part in sorted(name for name in contents if PART_RE.fullmatch(name)):
        root = etree.fromstring(contents[part])
        short = Path(part).stem.upper()
        paragraphs = root.xpath(".//w:p", namespaces=NS)
        part_changed = False
        for paragraph_index, paragraph in enumerate(paragraphs, 1):
            context = re.sub(r"\s+", " ", _paragraph_text(paragraph)).strip()
            matches = list(UNRESOLVED_MARKER.finditer(context))
            if not matches:
                continue
            coordinates = _table_coordinates(root, paragraph)
            for occurrence, match in enumerate(matches, 1):
                if coordinates is None:
                    location_id = (
                        f"{short}-P{paragraph_index:04d}-X{occurrence:02d}"
                    )
                    location_type = "段落"
                    table_index = row_index = column_index = ""
                else:
                    table_index, row_index, column_index = coordinates
                    location_id = (
                        f"{short}-T{table_index:02d}-R{row_index:02d}"
                        f"-C{column_index:02d}-X{occurrence:02d}"
                    )
                    location_type = "表格单元格"
                findings.append(
                    {
                        "location_id": location_id,
                        "location_type": location_type,
                        "part": part,
                        "paragraph_index": paragraph_index,
                        "occurrence_index": occurrence,
                        "context": context,
                        "current_text": match.group(0),
                        "table_index": table_index,
                        "row_index": row_index,
                        "column_index": column_index,
                    }
                )
            part_changed = (
                _split_paragraph_and_highlight_markers(paragraph) or part_changed
            )
        if part_changed:
            contents[part] = etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            changed = True

    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as output:
            for info in items:
                output.writestr(info, contents[info.filename])
    return findings


def _set_red_font(run) -> None:
    properties = run.find("w:rPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(f"{{{W}}}rPr")
        run.insert(0, properties)
    color = properties.find("w:color", namespaces=NS)
    if color is None:
        color = etree.SubElement(properties, f"{{{W}}}color")
    color.set(f"{{{W}}}val", "C00000")


TRACE_STATUS_FILLS = {
    "verified": "E2F0D9",
    "fallback": "FFF2CC",
    "review": "F4CCCC",
}


def _set_run_shading(run, fill: str) -> None:
    """Apply a light custom background without changing the run typography."""
    properties = run.find("w:rPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(f"{{{W}}}rPr")
        run.insert(0, properties)
    for shading in properties.findall("w:shd", namespaces=NS):
        properties.remove(shading)
    shading = etree.SubElement(properties, f"{{{W}}}shd")
    shading.set(f"{{{W}}}val", "clear")
    shading.set(f"{{{W}}}color", "auto")
    shading.set(f"{{{W}}}fill", fill)


def _compact_context(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def _trace_target_runs(
    roots: dict[str, Any],
    annotation: dict[str, Any],
    used: set[tuple[str, int, str, int]],
) -> tuple[str, Any, list[Any]] | None:
    """Locate one generated value using its table coordinate or semantic context."""
    target = str(annotation.get("target", "")).strip()
    if not target:
        return None
    part_hint = str(annotation.get("part") or "word/document.xml")
    root = roots.get(part_hint)
    if root is None:
        return None
    table_index = annotation.get("table_index")
    row_index = annotation.get("row_index")
    column_index = annotation.get("column_index")
    if all(value not in (None, "") for value in (table_index, row_index, column_index)):
        tables = root.xpath(".//w:tbl", namespaces=NS)
        try:
            table = tables[int(table_index)]
            row = table.xpath("./w:tr", namespaces=NS)[int(row_index)]
            cell = row.xpath("./w:tc", namespaces=NS)[int(column_index)]
        except (IndexError, TypeError, ValueError):
            return None
        runs = [run for run in cell.xpath(".//w:r", namespaces=NS) if _run_text(run)]
        exact = [run for run in runs if _matched_run_value(run, target) is not None]
        if exact:
            isolated = _isolate_run_text(exact[0], _matched_run_value(exact[0], target) or target)
            return (part_hint, root, [isolated] if isolated is not None else exact[:1])
        if target in _paragraph_text(cell) or _compact_context(target) == _compact_context(_paragraph_text(cell)):
            return part_hint, root, runs
        return None

    context_hint = _compact_context(annotation.get("context_hint"))
    paragraph_index_hint = int(annotation.get("paragraph_index_hint") or 0)
    target_occurrence = max(1, int(annotation.get("target_occurrence") or 1))
    candidates: list[tuple[int, Any, list[Any]]] = []
    for paragraph_index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS), 1):
        paragraph_text = _paragraph_text(paragraph)
        compact_paragraph = _compact_context(paragraph_text)
        if target not in paragraph_text and _compact_context(target) not in compact_paragraph:
            continue
        if context_hint and context_hint not in compact_paragraph:
            continue
        runs = [run for run in paragraph.xpath(".//w:r", namespaces=NS) if _run_text(run)]
        exact = [run for run in runs if _matched_run_value(run, target) is not None]
        if exact:
            chosen_runs = [exact[min(target_occurrence - 1, len(exact) - 1)]]
        else:
            chosen_runs = runs
        key = (part_hint, paragraph_index, target, target_occurrence)
        if chosen_runs and key not in used:
            candidates.append((paragraph_index, paragraph, chosen_runs))
    if not candidates:
        return None
    if paragraph_index_hint:
        candidates.sort(key=lambda item: abs(item[0] - paragraph_index_hint))
    paragraph_index, _paragraph, runs = candidates[0]
    if len(runs) == 1:
        matched = _matched_run_value(runs[0], target)
        if matched is not None:
            isolated = _isolate_run_text(runs[0], matched)
            if isolated is not None:
                runs = [isolated]
    used.add((part_hint, paragraph_index, target, target_occurrence))
    return part_hint, root, runs


def _existing_comment_id_for_runs(runs: list[Any]) -> str | None:
    if not runs:
        return None
    paragraph_nodes = runs[0].xpath("ancestor::w:p[1]", namespaces=NS)
    if not paragraph_nodes:
        return None
    paragraph = paragraph_nodes[0]
    siblings = list(paragraph)
    try:
        first = siblings.index(runs[0])
        last = siblings.index(runs[-1])
    except ValueError:
        return None
    starts: list[str] = []
    for node in siblings[: first + 1]:
        if node.tag == f"{{{W}}}commentRangeStart":
            value = node.get(f"{{{W}}}id")
            if value is not None:
                starts.append(str(value))
    for comment_id in reversed(starts):
        if any(
            node.tag == f"{{{W}}}commentRangeEnd"
            and str(node.get(f"{{{W}}}id")) == comment_id
            for node in siblings[last + 1 :]
        ):
            return comment_id
    return None


def annotate_traceable_content(
    path: Path,
    annotations: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Give every generated value a state colour and a titled provenance comment.

    ``verified`` is light green, ``fallback`` is light amber, ``review`` is
    light red, and ``missing`` keeps the yellow placeholder while turning its
    text red.  Exact table coordinates are preferred; paragraph values use the
    configured Word context and original paragraph index only as tie-breakers.
    """
    actionable = [item for item in annotations if str(item.get("target", "")).strip()]
    if not actionable:
        return []
    with zipfile.ZipFile(path) as archive:
        items = archive.infolist()
        contents = {info.filename: archive.read(info.filename) for info in items}
    roots = {
        part: etree.fromstring(contents[part])
        for part in sorted(name for name in contents if PART_RE.fullmatch(name))
    }
    comments = _comment_root(contents)
    existing_ids = [
        int(item.get(f"{{{W}}}id"))
        for item in comments.xpath(".//w:comment", namespaces=NS)
        if str(item.get(f"{{{W}}}id", "")).isdigit()
    ]
    next_id = max(existing_ids, default=-1) + 1
    applied: list[dict[str, Any]] = []
    changed_parts: set[str] = set()
    used: set[tuple[str, int, str, int]] = set()
    for annotation in actionable:
        located = _trace_target_runs(roots, annotation, used)
        if located is None:
            continue
        part, root, runs = located
        # Word comments are valid only in the main document story.  A comment
        # reference written into a header/footer makes LibreOffice reject the
        # entire package and Word may repair it on open.  Footer/header values
        # therefore keep their template formatting and are traced through the
        # corresponding body field when one exists.
        if part != "word/document.xml":
            continue
        runs = [run for run in runs if run is not None]
        if not runs:
            continue
        status = str(annotation.get("status") or "review")
        existing_comment_id = _existing_comment_id_for_runs(runs)
        # A prior source-conflict/LLM-review pass already anchored a warning
        # on this exact generated value.  Preserve that stronger state instead
        # of repainting the value green or amber during provenance coverage.
        if existing_comment_id is not None and status != "missing":
            status = "review"
        if status == "missing":
            for run in runs:
                _set_yellow_highlight(run)
                _set_red_font(run)
        else:
            fill = TRACE_STATUS_FILLS.get(status, TRACE_STATUS_FILLS["review"])
            for run in runs:
                _remove_highlight(run)
                _set_run_shading(run, fill)
                if status == "review":
                    _set_red_font(run)
        comment_id = existing_comment_id
        add_comment = bool(annotation.get("add_comment", True))
        if comment_id is None and add_comment:
            paragraph_nodes = runs[0].xpath("ancestor::w:p[1]", namespaces=NS)
            if not paragraph_nodes:
                continue
            paragraph = paragraph_nodes[0]
            first_parent = runs[0].getparent()
            last_parent = runs[-1].getparent()
            if first_parent is not paragraph or last_parent is not paragraph:
                continue
            comment_id = str(next_id)
            next_id += 1
            start = etree.Element(f"{{{W}}}commentRangeStart")
            start.set(f"{{{W}}}id", comment_id)
            end = etree.Element(f"{{{W}}}commentRangeEnd")
            end.set(f"{{{W}}}id", comment_id)
            paragraph.insert(paragraph.index(runs[0]), start)
            paragraph.insert(paragraph.index(runs[-1]) + 1, end)
            reference_run = etree.Element(f"{{{W}}}r")
            reference = etree.SubElement(reference_run, f"{{{W}}}commentReference")
            reference.set(f"{{{W}}}id", comment_id)
            paragraph.insert(paragraph.index(end) + 1, reference_run)
            comment = etree.SubElement(comments, f"{{{W}}}comment")
            comment.set(f"{{{W}}}id", comment_id)
            author = {
                "verified": "来源已核验",
                "fallback": "来源待补充核验",
                "review": "数据复核",
                "missing": "缺失数据",
            }.get(status, "数据复核")
            comment.set(f"{{{W}}}author", author)
            comment.set(f"{{{W}}}initials", "溯源")
            _append_comment_paragraph(comment, str(annotation.get("comment", "")).strip())
        applied.append({**annotation, "comment_id": comment_id, "part": part})
        contents[part] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        changed_parts.add(part)
    if not applied:
        return []
    _ensure_comment_parts(contents)
    contents[COMMENTS_PART] = etree.tostring(
        comments, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    existing_names = {info.filename for info in items}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as output:
        for info in items:
            output.writestr(info, contents[info.filename])
        for name in (DOCUMENT_RELS_PART, CONTENT_TYPES_PART, COMMENTS_PART):
            if name not in existing_names:
                output.writestr(name, contents[name])
    return applied


def _comment_root(contents: dict[str, bytes]):
    if COMMENTS_PART in contents:
        return etree.fromstring(contents[COMMENTS_PART])
    return etree.Element(f"{{{W}}}comments", nsmap={"w": W})


def _ensure_comment_parts(contents: dict[str, bytes]) -> None:
    """Add the OOXML relationship/content type needed for Word comments."""
    rels = (
        etree.fromstring(contents[DOCUMENT_RELS_PART])
        if DOCUMENT_RELS_PART in contents
        else etree.Element(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
        )
    )
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    if not any(item.get("Type") == COMMENTS_REL_TYPE for item in rels):
        existing = {
            item.get("Id", "") for item in rels.findall(f"{{{rel_ns}}}Relationship")
        }
        index = 1
        while f"rId{index}" in existing:
            index += 1
        relation = etree.SubElement(rels, f"{{{rel_ns}}}Relationship")
        relation.set("Id", f"rId{index}")
        relation.set("Type", COMMENTS_REL_TYPE)
        relation.set("Target", "comments.xml")
    contents[DOCUMENT_RELS_PART] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    content_types = etree.fromstring(contents[CONTENT_TYPES_PART])
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    if not any(
        item.get("PartName") == "/word/comments.xml"
        for item in content_types.findall(f"{{{ct_ns}}}Override")
    ):
        override = etree.SubElement(content_types, f"{{{ct_ns}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", COMMENTS_CONTENT_TYPE)
    contents[CONTENT_TYPES_PART] = etree.tostring(
        content_types, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _isolate_run_text(run, target: str):
    """Return a run containing only ``target`` while preserving surrounding text.

    Word often keeps a whole table-cell value or even a short label and value
    in one run.  Splitting that run means the red mark and comment point only
    to the differing PDF value, not to the full sentence or cell label.
    """
    text = _run_text(run)
    start = text.find(target)
    if start < 0:
        return None
    end = start + len(target)
    if start == 0 and end == len(text):
        return run
    parent = run.getparent()
    if parent is None:
        return None
    position = parent.index(run)
    pieces = (text[:start], target, text[end:])
    middle = None
    for offset, value in enumerate(pieces):
        if not value:
            continue
        clone = deepcopy(run)
        _set_run_text(clone, value)
        parent.insert(position + offset, clone)
        if value == target:
            middle = clone
    parent.remove(run)
    return middle


def _compact_match_text(value: Any) -> str:
    return re.sub(r"[\s：:，,、\\/（）()\-]", "", str(value or ""))


def _contains_semantic_text(actual: Any, expected: Any) -> bool:
    left, right = _compact_match_text(actual), _compact_match_text(expected)
    return bool(left and right and (left in right or right in left))


def _word_table_context(table) -> str:
    """Return the nearest visible paragraph before a table."""
    node = table.getprevious()
    while node is not None:
        if node.tag == f"{{{W}}}p":
            text = re.sub(r"\s+", " ", _paragraph_text(node)).strip()
            if text:
                return text
        node = node.getprevious()
    return ""


def _word_numeric_candidates(
    roots: dict[str, Any],
    target: str,
) -> list[dict[str, Any]]:
    """Collect every real Word occurrence and its table/row/period metadata."""
    candidates: list[dict[str, Any]] = []
    occurrence_counts: dict[str, int] = {}
    for part in sorted(roots):
        root = roots[part]
        paragraphs = root.xpath(".//w:p", namespaces=NS)
        runs = root.xpath(".//w:r", namespaces=NS)
        for run in runs:
            run_target = _matched_run_value(run, target)
            if run_target is None:
                continue
            paragraph_nodes = run.xpath("ancestor::w:p[1]", namespaces=NS)
            if not paragraph_nodes:
                continue
            paragraph = paragraph_nodes[0]
            coordinates = _table_coordinates(root, paragraph)
            table_index = row_index = column_index = ""
            row_text = period_text = table_context = ""
            if coordinates is not None:
                table_index, row_index, column_index = coordinates
                table = paragraph.xpath("ancestor::w:tbl[1]", namespaces=NS)[0]
                table_rows = table.xpath("./w:tr", namespaces=NS)
                if row_index <= len(table_rows):
                    row_text = re.sub(
                        r"\s+",
                        " ",
                        "".join(table_rows[row_index - 1].xpath(".//w:t/text()", namespaces=NS)),
                    ).strip()
                if table_rows:
                    header_cells = table_rows[0].xpath("./w:tc", namespaces=NS)
                    if column_index <= len(header_cells):
                        period_text = re.sub(
                            r"\s+",
                            " ",
                            "".join(header_cells[column_index - 1].xpath(".//w:t/text()", namespaces=NS)),
                        ).strip()
                table_context = _word_table_context(table)
            paragraph_index = next(
                (index for index, item in enumerate(paragraphs, 1) if item is paragraph),
                0,
            )
            short = Path(part).stem.upper()
            if table_index:
                candidate_base = (
                    f"{short}-T{int(table_index):02d}-R{int(row_index):02d}"
                    f"-C{int(column_index):02d}"
                )
            else:
                candidate_base = f"{short}-P{paragraph_index:04d}"
            occurrence_counts[candidate_base] = occurrence_counts.get(candidate_base, 0) + 1
            candidates.append(
                {
                    "candidate_id": (
                        f"{candidate_base}-V{occurrence_counts[candidate_base]:02d}"
                    ),
                    "part": part,
                    "paragraph_index": paragraph_index,
                    "paragraph_text": re.sub(r"\s+", " ", _paragraph_text(paragraph)).strip(),
                    "table_index": table_index,
                    "row_index": row_index,
                    "column_index": column_index,
                    "row_text": row_text,
                    "period_text": period_text,
                    "table_context": table_context,
                    "matched_value": run_target,
                    "candidate_kind": "value",
                    "anchor_text": run_target,
                    "_run": run,
                    "_root": root,
                }
            )
    return candidates


def _semantic_word_candidates(
    conflict: dict[str, Any],
    candidates: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Filter with the stable Word table, row and period identity."""
    matched = list(candidates)
    expected_table = conflict.get("word_table_index")
    if expected_table not in (None, ""):
        try:
            expected_table = int(expected_table)
        except (TypeError, ValueError):
            expected_table = None
        if expected_table is not None:
            matched = [item for item in matched if item.get("table_index") == expected_table]
    row_hint = str(conflict.get("word_context_hint", "")).strip()
    if row_hint:
        matched = [item for item in matched if _contains_semantic_text(item.get("row_text"), row_hint)]
    period_hint = str(conflict.get("word_period_hint", "")).strip()
    if period_hint:
        matched = [item for item in matched if _contains_semantic_text(item.get("period_text"), period_hint)]
    anchor_hints = conflict.get("word_anchor_hints", [])
    if isinstance(anchor_hints, list) and anchor_hints:
        for anchor in anchor_hints:
            if not isinstance(anchor, dict):
                continue
            prefix = str(anchor.get("prefix", "")).strip()
            suffix = str(anchor.get("suffix", "")).strip()
            anchored = [
                item
                for item in matched
                if (not prefix or _contains_semantic_text(item.get("paragraph_text"), prefix))
                and (not suffix or _contains_semantic_text(item.get("paragraph_text"), suffix))
            ]
            if anchored:
                matched = anchored
                break
    paragraph_hint = str(conflict.get("word_paragraph_hint", "")).strip()
    if paragraph_hint and not anchor_hints:
        matched = [
            item
            for item in matched
            if _contains_semantic_text(item.get("paragraph_text"), paragraph_hint)
            or _contains_semantic_text(item.get("table_context"), paragraph_hint)
        ]
    return matched


def _matched_run_value(run, target: str) -> str | None:
    """Find a display-equivalent numeric token (``4598.16`` / ``4,598.16``)."""
    text = _run_text(run)
    if target in text:
        return target
    normalized_target = target.replace(",", "").replace("，", "")
    if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?%?", normalized_target):
        return None
    for match in re.finditer(r"[-+]?\d[\d,]*(?:\.\d+)?%?", text):
        if match.group(0).replace(",", "") == normalized_target:
            return match.group(0)
    return None


def _default_review_comment(conflict: dict[str, Any], target: str) -> str:
    """Service-unavailable fallback; normal online runs use the LLM draft."""
    is_fallback = conflict.get("review_kind") == "excel_fallback"
    is_llm_review = conflict.get("review_kind") == "llm_review"
    if conflict.get("llm_comment"):
        value = str(conflict["llm_comment"]).strip()
        value = re.sub(r"^【[^】]+】", "", value).strip()
        title = (
            "【数据冲突，需人工复核】"
            if conflict.get("review_status") in {"needs_review", "conflict", "missing"}
            else "【来源已核验】"
        )
        return f"{title}{value}"
    if is_llm_review:
        return (
            "【数据冲突，需人工复核】"
            f"{conflict.get('field_name') or conflict.get('field_key', '')}当前采用 {target}，"
            f"来源为《{conflict.get('excel_file', '')}》“{conflict.get('excel_locator', '')}”。"
            f"复核结论为 {conflict.get('review_status', '')}："
            f"{conflict.get('review_reason', '') or '请核对原始材料'}。"
        )
    if is_fallback:
        pdf_status = (
            "审计PDF已上传，但该字段未从PDF中可靠识别，尚未完成PDF对照。"
            if conflict.get("pdf_uploaded")
            else "本次未上传审计PDF，尚未完成PDF对照。"
        )
        return (
            "【来源待补充核验】"
            f"{conflict.get('field_name') or conflict.get('field_key', '')}暂采用"
            f"{_source_reference(conflict.get('excel_file'), conflict.get('excel_locator'))}中的 {target}。"
            f"{pdf_status}请核对期间、单体/合并口径、金额单位及科目定义。"
        )
    return (
        "【数据冲突，需人工复核】"
        f"{conflict.get('field_name') or conflict.get('field_key', '')}在审计PDF"
        f"{_source_reference(conflict.get('pdf_file'), conflict.get('pdf_locator'))}中的识别值为 {target}，"
        f"与{_source_reference(conflict.get('excel_file'), conflict.get('excel_locator'))}中的"
        f" {conflict.get('excel_value', '')} 不一致。报告按来源规则采用PDF值；"
        "请核对期间、单体/合并口径、金额单位及科目定义后确认。"
    )


def _source_reference(source_file: Any, locator: Any) -> str:
    """Render one reviewer-facing source without repeating the file name."""
    file_text = str(source_file or "").strip()
    locator_text = str(locator or "").strip()
    if locator_text.startswith(("《", "审计 PDF", "PDF")) or (
        file_text and file_text in locator_text
    ):
        return locator_text
    if file_text and locator_text:
        return f"《{file_text}》“{locator_text}”"
    if file_text:
        return f"《{file_text}》"
    return f"“{locator_text}”" if locator_text else "本次上传材料"


def _append_comment_paragraph(comment, text: str, *, item_number: int = 1) -> None:
    paragraph = etree.SubElement(comment, f"{{{W}}}p")
    run = etree.SubElement(paragraph, f"{{{W}}}r")
    value = etree.SubElement(run, f"{{{W}}}t")
    value.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    prefix = "" if item_number == 1 else f"复核项 {item_number}："
    value.text = f"{prefix}{text.strip()}"


def _candidate_location_key(candidate: dict[str, Any]) -> tuple[Any, ...]:
    matched_value = str(candidate.get("matched_value", ""))
    if candidate.get("table_index") not in (None, ""):
        return (
            candidate.get("part"),
            "table",
            candidate.get("table_index"),
            candidate.get("row_index"),
            candidate.get("column_index"),
            matched_value,
        )
    return (
        candidate.get("part"),
        "paragraph",
        candidate.get("paragraph_index"),
        matched_value,
    )


def annotate_source_conflicts(
    path: Path,
    conflicts: list[dict[str, Any]],
    *,
    location_selector: Any | None = None,
) -> list[dict[str, Any]]:
    """Attach concise reconciliation notes, marking true conflicts in red.

    The report retains the authoritative audit-PDF value.  A red font is used
    only when a different numeric value was located in a non-authoritative
    workbook.  The attached Word comment makes the review action explicit,
    instead of forcing the reviewer to reconstruct file lineage manually.
    """
    actionable = []
    for item in conflicts:
        is_nonconflict_note = item.get("review_kind") in {"excel_fallback", "llm_review"}
        adopted_value = item.get("excel_value") if is_nonconflict_note else item.get("pdf_value")
        if str(adopted_value or "").strip():
            actionable.append(item)
    if not actionable:
        return []
    with zipfile.ZipFile(path) as archive:
        items = archive.infolist()
        contents = {info.filename: archive.read(info.filename) for info in items}

    comments = _comment_root(contents)
    existing_ids = [
        int(item.get(f"{{{W}}}id"))
        for item in comments.xpath(".//w:comment", namespaces=NS)
        if str(item.get(f"{{{W}}}id", "")).isdigit()
    ]
    next_id = max(existing_ids, default=-1) + 1
    applied: list[dict[str, Any]] = []
    changed_parts: set[str] = set()
    roots = {
        part: etree.fromstring(contents[part])
        for part in sorted(name for name in contents if PART_RE.fullmatch(name))
    }
    comment_groups: dict[tuple[Any, ...], dict[str, Any]] = {}

    for conflict in actionable:
        is_fallback = conflict.get("review_kind") == "excel_fallback"
        is_llm_review = conflict.get("review_kind") == "llm_review"
        target = str(
            conflict.get("excel_value") if (is_fallback or is_llm_review) else conflict.get("pdf_value")
        ).strip()
        numeric_candidates = _word_numeric_candidates(roots, target)
        if not numeric_candidates:
            # A review comment may only be anchored to the adopted value
            # itself.  Row labels, period headers and section headings are
            # context for validation, never substitute comment targets.
            continue
        has_semantic_hints = any(
            conflict.get(key)
            for key in (
                "word_table_index",
                "word_context_hint",
                "word_period_hint",
                "word_paragraph_hint",
                "word_anchor_hints",
            )
        )
        semantic_numeric = (
            _semantic_word_candidates(conflict, numeric_candidates)
            if has_semantic_hints
            else []
        )
        recommended_candidates = (
            semantic_numeric
            if semantic_numeric
            else numeric_candidates if len(numeric_candidates) == 1 else []
        )
        chosen: dict[str, Any] | None = None
        llm_comment = ""
        location_review_status = ""
        location_review_reason = ""
        selection_attempted = False
        selection: dict[str, Any] | None = None
        if location_selector is not None and hasattr(
            location_selector, "locate_word_review_comment"
        ):
            selection_attempted = True
            try:
                selection = location_selector.locate_word_review_comment(
                    conflict,
                    [
                        {key: value for key, value in item.items() if not key.startswith("_")}
                        for item in numeric_candidates
                    ],
                    [item["candidate_id"] for item in recommended_candidates],
                )
            except Exception:
                selection = None
            if isinstance(selection, dict):
                selected_id = str(selection.get("candidate_id", ""))
                llm_comment = str(selection.get("comment", "")).strip()
                location_review_status = str(selection.get("location_status", ""))
                location_review_reason = str(selection.get("reason", "")).strip()
                chosen = next(
                    (item for item in numeric_candidates if item["candidate_id"] == selected_id),
                    None,
                )
        elif location_selector is not None and hasattr(
            location_selector, "locate_word_comment_target"
        ):
            selection_attempted = True
            try:
                selected_id = location_selector.locate_word_comment_target(
                    conflict,
                    [
                        {key: value for key, value in item.items() if not key.startswith("_")}
                        for item in numeric_candidates
                    ],
                    [item["candidate_id"] for item in recommended_candidates],
                )
            except Exception:
                selected_id = None
            chosen = next(
                (item for item in numeric_candidates if item["candidate_id"] == selected_id),
                None,
            )
        if chosen is None:
            # When the model service is unavailable, a uniquely identified
            # rule match remains usable.  An explicit model rejection or a
            # genuinely ambiguous repeated amount is skipped rather than
            # attaching the comment to a label/header or the first occurrence.
            if not selection_attempted and len(recommended_candidates) == 1:
                chosen = recommended_candidates[0]
            elif selection_attempted and selection is None and len(recommended_candidates) == 1:
                chosen = recommended_candidates[0]
        if chosen is None:
            continue
        part = str(chosen["part"])
        root = chosen["_root"]
        run = _isolate_run_text(chosen["_run"], str(chosen["matched_value"]))
        if run is None:
            continue
        paragraph = run.xpath("ancestor::w:p[1]", namespaces=NS)
        if not paragraph:
            continue
        paragraph = paragraph[0]
        parent = run.getparent()
        if parent is None or parent.tag != f"{{{W}}}p":
            continue
        location_key = _candidate_location_key(chosen)
        # Excel-only fallback is a source-availability note and stays black.
        has_llm_warning = bool(llm_comment or conflict.get("llm_comment")) and conflict.get(
            "review_status"
        ) in {"needs_review", "conflict"}
        location_warning = location_review_status == "needs_review"
        if (not is_fallback or has_llm_warning or location_warning) and not (
            is_llm_review and conflict.get("review_status") == "missing"
        ):
            _set_red_font(run)
        review_text = llm_comment or _default_review_comment(conflict, target)
        group = comment_groups.get(location_key)
        if group is None:
            comment_id = str(next_id)
            next_id += 1
            start = etree.Element(f"{{{W}}}commentRangeStart")
            start.set(f"{{{W}}}id", comment_id)
            end = etree.Element(f"{{{W}}}commentRangeEnd")
            end.set(f"{{{W}}}id", comment_id)
            parent.insert(parent.index(run), start)
            parent.insert(parent.index(run) + 1, end)
            reference_run = etree.Element(f"{{{W}}}r")
            reference = etree.SubElement(reference_run, f"{{{W}}}commentReference")
            reference.set(f"{{{W}}}id", comment_id)
            paragraph.insert(paragraph.index(end) + 1, reference_run)

            comment = etree.SubElement(comments, f"{{{W}}}comment")
            comment.set(f"{{{W}}}id", comment_id)
            comment.set(f"{{{W}}}author", "数据需人工核对")
            comment.set(f"{{{W}}}initials", "核对")
            _append_comment_paragraph(comment, review_text)
            group = {"comment_id": comment_id, "comment": comment, "count": 1}
            comment_groups[location_key] = group
        else:
            comment_id = str(group["comment_id"])
            group["count"] = int(group["count"]) + 1
            _append_comment_paragraph(
                group["comment"],
                review_text,
                item_number=int(group["count"]),
            )
        contents[part] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        changed_parts.add(part)
        applied.append(
            {
                **conflict,
                "part": part,
                "comment_id": comment_id,
                "word_candidate_id": chosen["candidate_id"],
                "word_table_index": chosen.get("table_index", ""),
                "word_row_index": chosen.get("row_index", ""),
                "word_column_index": chosen.get("column_index", ""),
                "word_candidate_kind": chosen.get("candidate_kind", "value"),
                "comment_group_item": int(group["count"]),
                "comment_generated_by_llm": bool(llm_comment),
                "location_review_status": location_review_status,
                "location_review_reason": location_review_reason,
            }
        )

    if not applied:
        return []
    _ensure_comment_parts(contents)
    contents[COMMENTS_PART] = etree.tostring(
        comments, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    existing_names = {info.filename for info in items}
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as output:
        for info in items:
            output.writestr(info, contents[info.filename])
        for name in (DOCUMENT_RELS_PART, CONTENT_TYPES_PART, COMMENTS_PART):
            if name not in existing_names:
                output.writestr(name, contents[name])
    return applied


def replace_image_markers(path: Path) -> None:
    """Replace QCC image markers with the actual trademark image."""
    from docx import Document
    from docx.shared import Inches

    document = Document(str(path))
    changed = False
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                marker = next((p.text.strip() for p in cell.paragraphs if "__QCC_IMAGE__" in p.text), "")
                if not marker:
                    continue
                url = marker.split("__QCC_IMAGE__", 1)[1].strip()
                try:
                    request = Request(url, headers={"User-Agent": "asset-appraisal/1.0"})
                    with urlopen(request, timeout=15) as response:
                        data = response.read()
                    paragraph = cell.paragraphs[0]
                    paragraph.text = ""
                    paragraph.add_run().add_picture(BytesIO(data), width=Inches(0.42))
                except Exception:
                    cell.text = "图样"
                changed = True
    if changed:
        document.save(str(path))


def unresolved_placeholders(path: Path) -> list[str]:
    """Return template placeholders still present after a fill operation."""
    markers: set[str] = set()
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not PART_RE.fullmatch(name):
                continue
            text = archive.read(name).decode("utf-8", errors="ignore")
            markers.update(re.findall(r"20XX|X{2,}", text))
    return sorted(markers)


def replace_report_number_year(path: Path, year: Any) -> None:
    """Synchronize every report-number year, including legacy literal years.

    Some template occurrences use ``20XX`` while older pages contain a literal
    ``2024``.  They are the same report-number field and must not diverge.
    """
    year_text = str(year or "").strip()
    if not re.fullmatch(r"(?:19|20)\d{2}", year_text):
        return
    with zipfile.ZipFile(path) as zin:
        items = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in items}
    changed = False
    for name, data in list(contents.items()):
        if not PART_RE.fullmatch(name):
            continue
        root = etree.fromstring(data)
        for paragraph in root.xpath(".//w:p", namespaces=NS):
            text = _paragraph_text(paragraph)
            if "银信评报字" not in text:
                continue
            updated = re.sub(
                r"(银信评报字[（(])(?:20XX|(?:19|20)\d{2})([）)]第\s*\d+号)",
                rf"\g<1>{year_text}\g<2>",
                text,
            )
            if updated != text:
                runs = paragraph.xpath("./w:r", namespaces=NS)
                if runs:
                    _set_paragraph_text(paragraph, updated, False, style_run=runs[0])
                    changed = True
        if changed:
            contents[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if not changed:
        return
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in items:
            zout.writestr(info, contents[info.filename])


def replace_transaction_type_literals(path: Path, transaction_type: Any) -> None:
    """Replace literal ``拟收购`` text annotated as a transaction-type slot.

    The communication template has several legacy title/prose occurrences
    where the transaction type is literal text rather than an ``XXX``
    placeholder.  The review comments identify those occurrences as manual
    transaction-type fields.  Replace only the grammatical ``拟收购`` phrase;
    ordinary explanatory uses of the word ``收购`` are left unchanged.
    """
    value = str(transaction_type or "").strip()
    if value not in {"转让", "收购", "增资", "减资"}:
        return
    old = "拟收购"
    new = f"拟{value}"
    if old == new:
        return
    with zipfile.ZipFile(path) as zin:
        items = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in items}
    changed = False
    for name, data in list(contents.items()):
        if not PART_RE.fullmatch(name):
            continue
        root = etree.fromstring(data)
        part_changed = False
        for paragraph in root.xpath(".//w:p", namespaces=NS):
            text = _paragraph_text(paragraph)
            if old not in text:
                continue
            # In the supplied templates the phrase is normally in one run,
            # so this preserves every other run's font and paragraph layout.
            for node in paragraph.xpath(".//w:t|.//w:delText", namespaces=NS):
                if node.text and old in node.text:
                    node.text = node.text.replace(old, new)
                    part_changed = True
            # Handle a phrase split across runs without silently leaving it
            # unchanged.  This fallback uses the first run's formatting.
            if old in _paragraph_text(paragraph):
                updated = _paragraph_text(paragraph).replace(old, new)
                runs = paragraph.xpath("./w:r", namespaces=NS)
                if runs:
                    _set_paragraph_text(paragraph, updated, False, style_run=runs[0])
                    part_changed = True
        if part_changed:
            contents[name] = etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            changed = True
    if not changed:
        return
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in items:
            zout.writestr(info, contents[info.filename])


def _fill_tables(root, table_replacements: dict[int, list[list[str]]]) -> None:
    tables = root.xpath(".//w:tbl", namespaces=NS)
    for table_index, matrix in table_replacements.items():
        if table_index >= len(tables):
            raise ValueError(f"Word 表格编号不存在：{table_index}")
        table = tables[table_index]
        rows = table.xpath("./w:tr", namespaces=NS)
        if len(matrix) > len(rows):
            if not rows:
                raise ValueError(f"Word 表格 {table_index} 没有可复制的行")
            template_row = rows[-1]
            for _ in range(len(matrix) - len(rows)):
                table.append(deepcopy(template_row))
            rows = table.xpath("./w:tr", namespaces=NS)
        for row in rows[len(matrix):]:
            table.remove(row)
        for row_index, values in enumerate(matrix):
            cells = rows[row_index].xpath("./w:tc", namespaces=NS)
            if len(values) != len(cells):
                raise ValueError(f"Word 表格 {table_index} 第 {row_index + 1} 行列数不匹配")
            for cell, value in zip(cells, values, strict=True):
                _set_cell_text(cell, str(value))


def _set_table_column_ratios(
    root,
    table_column_ratios: dict[int, list[float]],
) -> None:
    tables = root.xpath(".//w:tbl", namespaces=NS)
    for table_index, ratios in table_column_ratios.items():
        if table_index >= len(tables):
            raise ValueError(f"Word 表格编号不存在：{table_index}")
        table = tables[table_index]
        grid_columns = table.xpath("./w:tblGrid/w:gridCol", namespaces=NS)
        if not grid_columns or len(grid_columns) != len(ratios):
            raise ValueError(f"Word 表格 {table_index} 的列宽配置与实际列数不匹配")
        total_ratio = sum(ratios)
        if total_ratio <= 0 or any(value <= 0 for value in ratios):
            raise ValueError(f"Word 表格 {table_index} 的列宽比例必须为正数")
        current_widths = [
            int(column.get(f"{{{W}}}w") or 0)
            for column in grid_columns
        ]
        total_width = sum(current_widths) or 9000
        widths = [
            max(1, round(total_width * ratio / total_ratio))
            for ratio in ratios
        ]
        widths[-1] += total_width - sum(widths)
        for column, width in zip(grid_columns, widths, strict=True):
            column.set(f"{{{W}}}w", str(width))
        for row in table.xpath("./w:tr", namespaces=NS):
            cells = row.xpath("./w:tc", namespaces=NS)
            if len(cells) != len(widths):
                continue
            for cell, width in zip(cells, widths, strict=True):
                properties = cell.find("w:tcPr", namespaces=NS)
                if properties is None:
                    properties = etree.Element(f"{{{W}}}tcPr")
                    cell.insert(0, properties)
                cell_width = properties.find("w:tcW", namespaces=NS)
                if cell_width is None:
                    cell_width = etree.SubElement(properties, f"{{{W}}}tcW")
                cell_width.set(f"{{{W}}}w", str(width))
                cell_width.set(f"{{{W}}}type", "dxa")


def _paragraph_replacement_text(paragraph, items, replacements) -> str:
    """Apply placeholder values to the visible paragraph text before splitting."""
    text = _paragraph_text(paragraph)
    for item in sorted(items, key=lambda item: item["occurrence_index"]):
        marker = str(item.get("marker") or "")
        if marker:
            text = text.replace(marker, str(replacements[item["location_id"]]), 1)
    text = re.sub(r"([。；;！!？?])\1+$", r"\1", text)
    return text


def _replace_paragraph_with_lines(paragraph, value: str) -> None:
    """Replace one placeholder paragraph with one real ``w:p`` per line."""
    parent = paragraph.getparent()
    if parent is None:
        _set_paragraph_text(paragraph, value, True)
        return
    lines = [line for line in str(value).splitlines() if line.strip()]
    if not lines:
        lines = [""]
    position = parent.index(paragraph)
    for offset, line in enumerate(lines):
        clone = deepcopy(paragraph)
        _set_paragraph_text(clone, line, True)
        parent.insert(position + offset, clone)
    parent.remove(paragraph)


def fill_template(
    template: Path,
    output: Path,
    replacements: dict[str, str],
    *,
    table_replacements: dict[int, list[list[str]]] | None = None,
    table_column_ratios: dict[int, list[float]] | None = None,
    paragraph_replacements: dict[tuple[str, int], str] | None = None,
    replacement_modes: dict[str, str] | None = None,
    progress_callback: Any | None = None,
) -> Path:
    if template.resolve() == output.resolve():
        raise ValueError("输出 Word 不能覆盖模板")
    locations = inventory_template(template)
    by_part_para: dict[tuple[str, int], list[dict]] = {}
    for item in locations:
        if item["location_id"] in replacements:
            by_part_para.setdefault((item["part"], item["paragraph_index"]), []).append(item)
    output.parent.mkdir(parents=True, exist_ok=True)
    table_replacements = table_replacements or {}
    table_column_ratios = table_column_ratios or {}
    paragraph_replacements = paragraph_replacements or {}
    replacement_modes = replacement_modes or {}
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            relevant = [(key, value) for key, value in by_part_para.items() if key[0] == info.filename]
            static_replacements = [
                (key, value) for key, value in paragraph_replacements.items() if key[0] == info.filename
            ]
            has_tables = info.filename == "word/document.xml" and bool(
                table_replacements or table_column_ratios
            )
            if relevant or has_tables or static_replacements:
                root = etree.fromstring(data)
                paragraphs = root.xpath(".//w:p", namespaces=NS)
                relevant_by_index = {p_index: items for (_, p_index), items in relevant}
                static_by_index = {p_index: value for (_, p_index), value in static_replacements}
                paragraph_indices = sorted(
                    set(relevant_by_index) | set(static_by_index),
                    reverse=True,
                )
                for p_index in paragraph_indices:
                    items = relevant_by_index.get(p_index, [])
                    paragraph = paragraphs[p_index - 1]
                    split_into_lines = False
                    if items and len(items) == 1 and items[0]["record_type"] == "占位符":
                        if progress_callback is not None:
                            progress_callback("fill_fields", f"正在填写：{items[0].get('field_name', items[0]['location_id'])}")
                        value = str(replacements[items[0]["location_id"]])
                        if "\n" in value:
                            _replace_paragraph_with_lines(
                                paragraph,
                                _paragraph_replacement_text(paragraph, items, replacements),
                            )
                            split_into_lines = True
                    if items and not split_into_lines and items[0]["record_type"] == "黄色标注内容块":
                        if progress_callback is not None:
                            progress_callback("fill_fields", f"正在填写：{items[0].get('field_name', items[0]['location_id'])}")
                        value = str(replacements[items[0]["location_id"]])
                        mode = replacement_modes.get(items[0]["location_id"], "replace_paragraph")
                        if mode == "strip_yellow_annotation":
                            if value:
                                _replace_yellow_annotation(paragraph, value)
                            else:
                                _set_paragraph_text(paragraph, _strip_yellow_annotation(paragraph), True)
                        elif mode == "strip_yellow_only":
                            if UNRESOLVED_MARKER.fullmatch(value):
                                _replace_yellow_annotation(paragraph, value)
                            else:
                                _set_paragraph_text(
                                    paragraph,
                                    _strip_yellow_annotation(paragraph),
                                    True,
                                )
                        elif mode == "replace_yellow_annotation":
                            _replace_yellow_annotation(paragraph, value)
                        elif mode == "replace_method_heading":
                            _replace_method_heading(paragraph, value)
                        elif value:
                            _set_paragraph_text(paragraph, value, True)
                        else:
                            _set_paragraph_text(paragraph, _strip_yellow_annotation(paragraph), True)
                    elif items and not split_into_lines:
                        _replace_placeholders_preserving_runs(paragraph, items, replacements)
                    if p_index in static_by_index and not split_into_lines:
                        _set_paragraph_text(paragraph, str(static_by_index[p_index]), True)
                if has_tables:
                    if progress_callback is not None:
                        progress_callback("fill_tables", "正在填写财务表格和长期资产表")
                    _fill_tables(root, table_replacements)
                    _set_table_column_ratios(root, table_column_ratios)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(info, data)
    return output
