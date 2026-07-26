import json

import pytest
from pydantic import ValidationError

from demo.adapters.workflow_trace import WorkflowTraceRecorder
from demo.schemas import (
    InventoryInput,
    InventoryOutput,
)


def test_trace_recorder_validates_models_and_serializes_versions(tmp_path):
    recorder = WorkflowTraceRecorder(
        workflow_version="1.1.0",
        contract_version="workflow_contract.v1",
        versions={"prompt": "yellow_narratives.v1"},
    )

    recorder.record(
        node_name="inventory",
        input_model=InventoryInput,
        output_model=InventoryOutput,
        input_payload={"template_path": "template.docx"},
        output_payload={"locations": []},
        status="completed",
        evidence=[],
        issues=[],
        human_checkpoint=None,
    )
    path = recorder.export(tmp_path / "workflow_trace.json")
    payload = json.loads(path.read_text(encoding="utf-8"))

    assert payload["nodes"][0]["input_model"] == "InventoryInput"
    assert payload["nodes"][0]["output_model"] == "InventoryOutput"
    assert payload["nodes"][0]["status"] == "completed"
    assert payload["versions"]["prompt"] == "yellow_narratives.v1"
    assert payload["nodes"][0]["started_at"].endswith("+00:00")
    assert payload["nodes"][0]["finished_at"].endswith("+00:00")


def test_trace_recorder_rejects_invalid_output():
    recorder = WorkflowTraceRecorder(
        workflow_version="1.1.0",
        contract_version="workflow_contract.v1",
        versions={},
    )

    with pytest.raises(ValidationError):
        recorder.record(
            node_name="inventory",
            input_model=InventoryInput,
            output_model=InventoryOutput,
            input_payload={"template_path": "template.docx"},
            output_payload={},
            status="completed",
            evidence=[],
            issues=[],
            human_checkpoint=None,
        )


def test_trace_recorder_rejects_unknown_status():
    recorder = WorkflowTraceRecorder(
        workflow_version="1.1.0",
        contract_version="workflow_contract.v1",
        versions={},
    )

    with pytest.raises(ValidationError):
        recorder.record(
            node_name="inventory",
            input_model=InventoryInput,
            output_model=InventoryOutput,
            input_payload={"template_path": "template.docx"},
            output_payload={"locations": []},
            status="unknown",
            evidence=[],
            issues=[],
            human_checkpoint=None,
        )
