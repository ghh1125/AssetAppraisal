from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil

import pytest

from demo.template_bundle import (
    DEFAULT_TEMPLATE_BUNDLE_DIR,
    TEMPLATE_BUNDLE_ENV,
    load_workbook_template_bundle,
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_builtin_workbook_templates_are_versioned_application_resources(monkeypatch):
    monkeypatch.delenv(TEMPLATE_BUNDLE_ENV, raising=False)
    bundle = load_workbook_template_bundle()

    assert bundle.version == "generic_appraisal_workbooks.v1"
    assert bundle.root == DEFAULT_TEMPLATE_BUNDLE_DIR.resolve()
    assert bundle.asset.name == "asset_base_v1.xlsx"
    assert bundle.income.name == "income_approach_v1.xlsx"
    assert "资产评估工作流" not in str(bundle.asset)
    assert "资产评估工作流" not in str(bundle.income)


def test_environment_can_select_a_complete_custom_template_bundle(tmp_path, monkeypatch):
    asset = tmp_path / "asset.xlsx"
    income = tmp_path / "income.xlsx"
    shutil.copy2(DEFAULT_TEMPLATE_BUNDLE_DIR / "asset_base_v1.xlsx", asset)
    shutil.copy2(DEFAULT_TEMPLATE_BUNDLE_DIR / "income_approach_v1.xlsx", income)
    built_in = json.loads((DEFAULT_TEMPLATE_BUNDLE_DIR / "workbook_templates.json").read_text(encoding="utf-8"))
    (tmp_path / "workbook_templates.json").write_text(
        json.dumps(
            {
                "version": "customer-neutral.v2",
                "roles": {
                    "asset": {**built_in["roles"]["asset"], "file": asset.name, "sha256": _sha256(asset)},
                    "income": {**built_in["roles"]["income"], "file": income.name, "sha256": _sha256(income)},
                },
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv(TEMPLATE_BUNDLE_ENV, str(tmp_path))

    bundle = load_workbook_template_bundle()

    assert bundle.version == "customer-neutral.v2"
    assert bundle.asset == asset.resolve()
    assert bundle.income == income.resolve()


def test_template_bundle_rejects_untracked_or_modified_binary(tmp_path):
    asset = tmp_path / "asset.xlsx"
    income = tmp_path / "income.xlsx"
    shutil.copy2(DEFAULT_TEMPLATE_BUNDLE_DIR / "asset_base_v1.xlsx", asset)
    shutil.copy2(DEFAULT_TEMPLATE_BUNDLE_DIR / "income_approach_v1.xlsx", income)
    built_in = json.loads((DEFAULT_TEMPLATE_BUNDLE_DIR / "workbook_templates.json").read_text(encoding="utf-8"))
    (tmp_path / "workbook_templates.json").write_text(
        json.dumps(
            {
                "version": "generic.v1",
                "roles": {
                    "asset": {**built_in["roles"]["asset"], "file": asset.name, "sha256": "0" * 64},
                    "income": {**built_in["roles"]["income"], "file": income.name, "sha256": _sha256(income)},
                },
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(RuntimeError, match="完整性校验失败"):
        load_workbook_template_bundle(tmp_path)
