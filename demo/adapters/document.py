from __future__ import annotations

from pathlib import Path

from docx import Document


def read_paragraph_containing(path: Path, text: str) -> str:
    """Return the first body paragraph containing the configured evidence text."""
    for paragraph in Document(path).paragraphs:
        value = paragraph.text.strip()
        if text in value:
            return value
    return ""


def read_table_cell(
    path: Path,
    *,
    table_contains: str,
    row_contains: str,
    column: int,
) -> str:
    """Return one cell from the first matching DOCX table row."""
    for table in Document(path).tables:
        table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
        if table_contains not in table_text:
            continue
        for row in table.rows:
            if row_contains in "\t".join(cell.text for cell in row.cells):
                if column >= len(row.cells):
                    raise ValueError(f"Word 表格列号不存在：{column}")
                return row.cells[column].text.strip()
    return ""


def read_table_matrix(path: Path, table_index: int) -> list[list[str]]:
    """Read a DOCX table as plain text without changing the source document."""
    tables = Document(path).tables
    if table_index < 0 or table_index >= len(tables):
        raise ValueError(f"Word 表格编号不存在：{table_index}")
    return [[cell.text.strip() for cell in row.cells] for row in tables[table_index].rows]
