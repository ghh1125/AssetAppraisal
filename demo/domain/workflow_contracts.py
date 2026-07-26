from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from pydantic import ValidationError

from demo.schemas import WorkflowDefinition, WorkflowNodeDefinition


def _has_cycle(nodes: list[WorkflowNodeDefinition]) -> bool:
    dependencies = {node.name: set(node.depends_on) for node in nodes}
    visited: set[str] = set()
    active: set[str] = set()

    def visit(name: str) -> bool:
        if name in active:
            return True
        if name in visited:
            return False
        active.add(name)
        for dependency in dependencies.get(name, set()):
            if dependency in dependencies and visit(dependency):
                return True
        active.remove(name)
        visited.add(name)
        return False

    return any(visit(name) for name in dependencies)


def validate_workflow_contract(
    payload: Mapping[str, Any],
    schema_module: Any,
) -> dict[str, Any]:
    """Validate workflow structure and the business metadata of node models."""
    try:
        definition = WorkflowDefinition.model_validate(payload)
    except ValidationError as exc:
        return {
            "valid": False,
            "issues": [str(exc)],
            "node_count": 0,
        }

    issues: list[str] = []
    names = [node.name for node in definition.nodes]
    if len(names) != len(set(names)):
        issues.append("工作流节点名称重复")

    known_names = set(names)
    for node in definition.nodes:
        for dependency in node.depends_on:
            if dependency not in known_names:
                issues.append(f"{node.name}：前置节点不存在：{dependency}")
        for direction, model_name in (
            ("输入", node.input_model),
            ("输出", node.output_model),
        ):
            model = getattr(schema_module, model_name, None)
            if model is None:
                issues.append(f"{node.name}：{direction}模型不存在：{model_name}")
                continue
            for field_name, field in model.model_fields.items():
                description = field.description or ""
                if not any("\u4e00" <= character <= "\u9fff" for character in description):
                    issues.append(f"{model_name}.{field_name}：缺少中文业务说明")
                if not field.examples:
                    issues.append(f"{model_name}.{field_name}：缺少示例")

    if _has_cycle(definition.nodes):
        issues.append("工作流依赖存在环")

    return {
        "valid": not issues,
        "issues": issues,
        "node_count": len(definition.nodes),
    }
