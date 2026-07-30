import json
from pathlib import Path

from demo import schemas
from demo.domain.workflow_contracts import validate_workflow_contract


def test_validates_current_workflow_models_and_dependencies():
    definition = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))

    result = validate_workflow_contract(definition, schemas)

    assert result == {"valid": True, "issues": [], "node_count": 4}


def test_rejects_unknown_model_and_dependency_cycle():
    definition = {
        "version": "test",
        "contract_version": "workflow_contract.v1",
        "nodes": [
            {
                "name": "a",
                "input_model": "MissingInput",
                "output_model": "InventoryOutput",
                "depends_on": ["b"],
                "human_checkpoint": None,
            },
            {
                "name": "b",
                "input_model": "InventoryInput",
                "output_model": "InventoryOutput",
                "depends_on": ["a"],
                "human_checkpoint": None,
            },
        ],
    }

    result = validate_workflow_contract(definition, schemas)

    assert not result["valid"]
    assert "a：输入模型不存在：MissingInput" in result["issues"]
    assert "工作流依赖存在环" in result["issues"]


def test_rejects_duplicate_and_missing_dependency():
    definition = {
        "version": "test",
        "contract_version": "workflow_contract.v1",
        "nodes": [
            {
                "name": "inventory",
                "input_model": "InventoryInput",
                "output_model": "InventoryOutput",
                "depends_on": ["missing"],
                "human_checkpoint": None,
            },
            {
                "name": "inventory",
                "input_model": "InventoryInput",
                "output_model": "InventoryOutput",
                "depends_on": [],
                "human_checkpoint": None,
            },
        ],
    }

    result = validate_workflow_contract(definition, schemas)

    assert not result["valid"]
    assert "工作流节点名称重复" in result["issues"]
    assert "inventory：前置节点不存在：missing" in result["issues"]
