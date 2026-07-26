from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from demo.adapters.audit import write_json
from demo.schemas import DemoModel, NodeTrace, WorkflowTrace


def trace_evidence(source: dict[str, Any] | None) -> dict[str, str]:
    source = source or {}
    return {
        "source_kind": str(source.get("kind", "unknown")),
        "source_file": str(source.get("file", "")),
        "source_locator": str(source.get("locator", "")),
    }


def trace_candidates(
    values: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    field_names: dict[str, str],
) -> list[dict[str, Any]]:
    return [
        {
            "field_key": field_key,
            "field_name": field_names.get(field_key, field_key),
            "value": value,
            "evidence": trace_evidence(evidence.get(field_key)),
        }
        for field_key, value in values.items()
        if value not in (None, "", [], {})
    ]


def trace_resolved_fields(
    values: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    field_names: dict[str, str],
    *,
    keys: set[str] | None = None,
) -> list[dict[str, Any]]:
    return [
        {
            "field_key": field_key,
            "field_name": field_names.get(field_key, field_key),
            "value": value,
            "evidence": trace_evidence(evidence.get(field_key)),
            "candidates": [],
        }
        for field_key, value in values.items()
        if (keys is None or field_key in keys) and value not in (None, "", [], {})
    ]


def trace_mappings(locations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "location_id": item["location_id"],
            "field_key": item["field_key"],
            "field_name": item["field_name"],
            "record_type": item["record_type"],
            "source_priority": list(item.get("candidate_sources", [])),
        }
        for item in locations
    ]


class WorkflowTraceRecorder:
    """Validate node boundaries and retain a JSON-safe execution trace."""

    def __init__(
        self,
        workflow_version: str,
        contract_version: str,
        versions: dict[str, str],
    ) -> None:
        self.workflow_version = workflow_version
        self.contract_version = contract_version
        self.versions = dict(versions)
        self.nodes: list[NodeTrace] = []

    def record(
        self,
        *,
        node_name: str,
        input_model: type[DemoModel],
        output_model: type[DemoModel],
        input_payload: dict[str, Any],
        output_payload: dict[str, Any],
        status: str,
        evidence: list[dict[str, Any]],
        issues: list[str],
        human_checkpoint: str | None,
    ) -> DemoModel:
        started_at = datetime.now(timezone.utc).isoformat(timespec="microseconds")
        validated_input = input_model.model_validate(input_payload)
        validated_output = output_model.model_validate(output_payload)
        finished_at = datetime.now(timezone.utc).isoformat(timespec="microseconds")
        trace = NodeTrace.model_validate(
            {
                "node_name": node_name,
                "status": status,
                "started_at": started_at,
                "finished_at": finished_at,
                "input_model": input_model.__name__,
                "output_model": output_model.__name__,
                "input_data": validated_input.model_dump(mode="json"),
                "output_data": validated_output.model_dump(mode="json"),
                "evidence": evidence,
                "issues": issues,
                "human_checkpoint": human_checkpoint,
            }
        )
        self.nodes.append(trace)
        return validated_output

    def export(self, path: Path) -> Path:
        trace = WorkflowTrace(
            workflow_version=self.workflow_version,
            contract_version=self.contract_version,
            versions=self.versions,
            nodes=self.nodes,
        )
        return write_json(path, trace.model_dump(mode="json"))
