import json

from demo.run_material_intake import _cached_pages, _has_usable_extraction


def test_placeholder_pdf_pages_are_not_usable_extraction():
    pages = [
        {
            "page_number": 1,
            "page_count": 1,
            "blocks": [{"text": ""}],
            "tables": [],
        }
    ]

    assert not _has_usable_extraction(pages)


def test_cached_pages_rejects_placeholder_only_cache(tmp_path, monkeypatch):
    source = tmp_path / "scan.pdf"
    source.write_bytes(b"scan")
    pages = [
        {
            "page_number": 1,
            "page_count": 1,
            "blocks": [{"text": ""}],
            "tables": [],
        }
    ]
    monkeypatch.setattr("demo.run_material_intake._sha256", lambda _path: "digest")
    (tmp_path / "digest.json").write_text(json.dumps(pages), encoding="utf-8")

    assert _cached_pages(tmp_path, source) is None


def test_cached_pages_accepts_nonempty_table_cache(tmp_path, monkeypatch):
    source = tmp_path / "scan.pdf"
    source.write_bytes(b"scan")
    pages = [
        {
            "page_number": 1,
            "page_count": 1,
            "blocks": [],
            "tables": [{"cells": [{"text": "资产负债表"}]}],
        }
    ]
    monkeypatch.setattr("demo.run_material_intake._sha256", lambda _path: "digest")
    (tmp_path / "digest.json").write_text(json.dumps(pages), encoding="utf-8")

    assert _cached_pages(tmp_path, source) == pages
