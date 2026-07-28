from __future__ import annotations

from demo.adapters.aliyun_docmind_ocr import layouts_to_pages


def test_layouts_to_pages_preserves_text_table_and_evidence():
    layouts = [
        {
            "type": "text",
            "text": "资产总计",
            "pageNum": 0,
            "layoutConf": 0.98,
            "pos": [{"x": 10, "y": 20}, {"x": 90, "y": 40}],
            "uniqueId": "text-1",
        },
        {
            "type": "table",
            "pageNum": 1,
            "uniqueId": "table-1",
            "cells": [
                {
                    "ysc": 0,
                    "yec": 0,
                    "xsc": 0,
                    "xec": 1,
                    "layouts": [{"text": "项目"}],
                },
                {
                    "ysc": 1,
                    "yec": 1,
                    "xsc": 0,
                    "xec": 0,
                    "layouts": [{"text": "货币资金"}],
                },
            ],
        },
    ]

    pages = layouts_to_pages(layouts)

    assert pages[0]["page_number"] == 1
    assert pages[0]["page_count"] == 2
    assert pages[0]["blocks"] == [
        {
            "block_id": "text-1",
            "block_type": "text",
            "text": "资产总计",
            "confidence": 0.98,
            "bbox": [10.0, 20.0, 90.0, 40.0],
        }
    ]
    assert pages[1]["page_number"] == 2
    assert pages[1]["tables"][0]["table_id"] == "table-1"
    assert pages[1]["tables"][0]["cells"][0] == {
        "row": 1,
        "column": 1,
        "row_span": 1,
        "column_span": 2,
        "text": "项目",
        "confidence": None,
        "bbox": [],
    }


def test_layouts_to_pages_keeps_unknown_confidence_and_uses_markdown_text():
    pages = layouts_to_pages(
        [
            {
                "type": "title",
                "markdownContent": "# 财务报表  \n\n",
                "pageNum": 0,
                "uniqueId": "title-1",
            }
        ]
    )

    assert pages[0]["blocks"][0]["text"] == "财务报表"
    assert pages[0]["blocks"][0]["confidence"] is None
    assert pages[0]["blocks"][0]["bbox"] == []


def test_layouts_to_pages_returns_empty_list_for_empty_layouts():
    assert layouts_to_pages([]) == []
