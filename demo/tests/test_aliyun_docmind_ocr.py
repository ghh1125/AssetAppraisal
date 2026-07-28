from __future__ import annotations

from pathlib import Path

from demo.adapters.aliyun_docmind_ocr import (
    AliyunDocMindOcrAdapter,
    layouts_to_pages,
)


class FakeClient:
    def __init__(self, *, statuses=None, result_pages=None, error=None):
        self.statuses = list(statuses or [{"Status": "success"}])
        self.result_pages = list(result_pages or [{"layouts": []}])
        self.error = error
        self.result_offsets = []
        self.submissions = []

    def submit(self, pdf_path: Path, *, vlm: bool) -> str:
        self.submissions.append((pdf_path, vlm))
        if self.error:
            raise self.error
        return "task-1"

    def status(self, task_id: str):
        assert task_id == "task-1"
        return self.statuses.pop(0)

    def result(self, task_id: str, *, layout_num: int, layout_step_size: int):
        assert task_id == "task-1"
        self.result_offsets.append(layout_num)
        return self.result_pages.pop(0)


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


def test_adapter_submits_polls_and_collects_all_layout_pages(tmp_path):
    pdf_path = tmp_path / "audit.pdf"
    pdf_path.write_bytes(b"%PDF")
    client = FakeClient(
        statuses=[{"Status": "Processing"}, {"Status": "success"}],
        result_pages=[
            {"layouts": [{"type": "text", "text": "第一页", "pageNum": 0}]},
            {"layouts": []},
        ],
    )
    adapter = AliyunDocMindOcrAdapter(
        client,
        poll_interval_seconds=0,
        timeout_seconds=60,
        layout_step_size=1,
        sleep=lambda _: None,
    )

    pages, issues = adapter.extract(pdf_path)

    assert issues == []
    assert pages[0]["blocks"][0]["text"] == "第一页"
    assert client.submissions == [(pdf_path, False)]
    assert client.result_offsets == [0, 1]


def test_adapter_returns_failed_task_as_review_issue(tmp_path):
    client = FakeClient(
        statuses=[{"Status": "Fail", "Code": "QuotaExhausted", "Message": "no quota"}]
    )

    pages, issues = AliyunDocMindOcrAdapter(client, sleep=lambda _: None).extract(
        tmp_path / "audit.pdf"
    )

    assert pages == []
    assert issues == ["阿里云 OCR 任务失败（task-1，QuotaExhausted）：no quota"]


def test_adapter_stops_polling_after_timeout(tmp_path):
    times = iter([0.0, 61.0])
    client = FakeClient(statuses=[{"Status": "Processing"}])
    adapter = AliyunDocMindOcrAdapter(
        client,
        timeout_seconds=60,
        sleep=lambda _: None,
        clock=lambda: next(times),
    )

    pages, issues = adapter.extract(tmp_path / "audit.pdf")

    assert pages == []
    assert issues == ["阿里云 OCR 轮询超时（task-1，60 秒）"]


def test_adapter_redacts_credentials_from_sdk_errors(tmp_path):
    client = FakeClient(error=RuntimeError("request used secret-token"))
    adapter = AliyunDocMindOcrAdapter(client, redact_values=("secret-token",))

    pages, issues = adapter.extract(tmp_path / "audit.pdf")

    assert pages == []
    assert "secret-token" not in issues[0]
    assert "***" in issues[0]


def test_adapter_reports_successful_empty_result(tmp_path):
    pages, issues = AliyunDocMindOcrAdapter(FakeClient()).extract(
        tmp_path / "audit.pdf"
    )

    assert pages == []
    assert issues == ["阿里云 OCR 返回成功但没有可用文本或表格"]
