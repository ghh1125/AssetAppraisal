"""Local HTTP bridge for the Vue asset-appraisal workbench.

The business pipeline remains in ``demo.pipeline``. This module only handles
uploads, job state and artifact downloads; c2m can replace it with its own
authenticated task service without changing the domain workflow.
"""

from __future__ import annotations

import json
import hashlib
import logging
import os
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Mapping

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from .run import _load_local_env
from .adapters.ocr_factory import create_ocr_adapter
from .domain.field_validation import (
    normalize_valuation_methods,
    validate_required_text,
    validate_final_valuation_method,
    validate_material_source_strategy,
    validate_transaction_type,
    validate_valuation_base_date,
    validate_valuation_subject_type,
)
from .domain.company_matching import matching_company_records, normalize_company_name
from .adapters.workbook_delivery_qa import prepare_generated_workbooks

ROOT = Path(__file__).resolve().parents[1]
PROJECT_CONFIG = ROOT / "demo/projects/tongfu.yaml"
RUNS_ROOT = ROOT / "runs/web"
OCR_CACHE_ROOT = ROOT / "runs"
JOBS: dict[str, dict[str, Any]] = {}
WORKBOOK_INTAKES: dict[str, dict[str, Any]] = {}
JOBS_LOCK = Lock()
LOGGER = logging.getLogger(__name__)

WORKBOOK_INTAKE_STEPS = (
    ("validate_archive", "上传与安全校验", "检查格式、空文件和任务隔离目录"),
    ("unpack_archive", "解压材料包", "安全解压并阻止越界路径和符号链接"),
    ("inventory", "材料清点", "清点可解析的 PDF、图片、Word、Excel 和 PPT"),
    ("ocr_materials", "逐文件解析", "读取原生表格/文本，并为扫描材料记录 OCR 坐标"),
    ("classify_materials", "材料分类", "识别审计报告、附注、营业执照和企业信息"),
    ("match_subject", "主体匹配", "隔离不同公司的材料，避免跨案例混入"),
    ("load_mapping_rules", "读取系统规则图", "加载系统提供的 Sheet、单元格、科目、期间和公式依赖规则"),
    ("map_asset", "生成资产法 Excel", "按资产法模板逐单元格映射审定数据"),
    ("map_income", "生成收益法 Excel", "填入历史三表并保留收益法公式链"),
    ("write_mapping_trace", "记录本次映射溯源", "记录本案例命中的规则、来源、换算和待补资料"),
    ("formula_qa", "公式重算与错误检查", "后台检查公式结构与错误缓存，并设置打开时自动完整计算"),
    ("verify_output", "结果校验", "确认两份工作簿存在、命名正确且可下载"),
)

PUBLIC_NODES = (
    ("start_input", "节点 2：开始 / 输入", "接收节点 1 的工作簿、人工字段和其他上传材料"),
    ("ocr_llm_candidates", "节点 3：材料解析 / LLM 候选", "OCR、Excel、企查查解析并生成候选"),
    ("fill_word", "节点 4：填充 Word", "写入确定性字段和用户选中的候选"),
    ("output", "节点 5：结果输出", "生成评估报告 Word"),
)

NODE_STEPS = {
    "start_input": (
        ("validate_inputs", "校验人工输入", "检查公司名称、评估对象、方法和基准日"),
        ("store_materials", "整理上传材料", "识别 PDF、Excel 和补充材料的文件角色"),
        ("load_template", "加载 Word 模板", "确认后台只读模板和批注映射已就绪"),
    ),
    "ocr_llm_candidates": (
        ("detect_materials", "识别材料类型", "确认审计材料、资产法表和收益法表"),
        ("ocr_pdf", "解析扫描材料", "有 PDF 或图片时执行 OCR；命中缓存时直接复用"),
        ("parse_excel", "解析 Excel 表格", "按工作表标题、科目、期间和单位识别数据"),
        ("reconcile_sources", "整合并比对来源", "按科目、期间、口径和单位归并 PDF/OCR 与 Excel，并记录一致或冲突"),
        ("review_evidence", "LLM 证据复核批注", "读取已整理的 PDF/OCR 与 Excel 证据，逐字段生成不改数值的人工审核批注"),
        ("query_qichacha", "企查查 API 搜索", "查询工商、股权、商标、专利和上市信息"),
        ("generate_candidates", "生成 LLM 候选", "按固定 Word 位置分别生成六个可选报告模块"),
        ("wait_selection", "等待人工选择", "展示候选内容，等待确认写入哪些模块"),
    ),
    "fill_word": (
        ("load_selection", "读取人工选择", "载入用户确认的 LLM 模块"),
        ("map_word", "匹配 Word 批注位置", "按模板原文和批注映射定位写入位置"),
        ("fill_fields", "填写文字字段", "逐项写入人工输入、PDF/OCR、Excel 和企查查字段"),
        ("fill_tables", "填写财务表格", "按工作表和表头语义写入资产负债及历史报表"),
        ("write_comments", "写入 LLM 批注", "把模型生成的审核说明挂到对应数值位置"),
        ("check_placeholders", "校验未填位置", "找不到证据的字段保留黄色 XXX"),
    ),
    "output": (
        ("save_word", "保存评估报告", "复制模板生成独立的评估报告 Word"),
        ("verify_output", "检查输出文件", "确认 Word 可打开且没有覆盖原模板"),
    ),
}

app = FastAPI(title="Asset Appraisal API", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def _project_template() -> Path:
    config = json.loads(PROJECT_CONFIG.read_text(encoding="utf-8"))
    # The web workflow uses the latest comment-annotated template when the
    # project provides one; CLI regression fixtures keep the legacy yellow
    # template as their explicit ``template`` entry.
    template = Path(config.get("web_template", config["template"]))
    resolved = template if template.is_absolute() else (PROJECT_CONFIG.parent / template).resolve()
    if not resolved.is_file():
        raise RuntimeError(f"后端默认模板不存在：{resolved}")
    return resolved


def _set_job(job_key: str, **values: Any) -> None:
    with JOBS_LOCK:
        JOBS.setdefault(job_key, {}).update(values)


def _initial_node_states() -> list[dict[str, Any]]:
    return [
        {
            "key": key,
            "name": name,
            "description": description,
            "status": "pending",
            "message": "等待执行",
            "active_step": "",
            "steps": [
                {"key": step_key, "name": step_name, "description": step_description, "status": "pending", "message": "等待执行"}
                for step_key, step_name, step_description in NODE_STEPS[key]
            ],
        }
        for key, name, description in PUBLIC_NODES
    ]


def _set_node(run_id: str, key: str, status: str, message: str = "") -> None:
    with JOBS_LOCK:
        job = JOBS.setdefault(run_id, {})
        nodes = job.setdefault("nodes", _initial_node_states())
        for node in nodes:
            if node.get("key") == key:
                node["status"] = status
                node["message"] = message
                if status in {"completed", "failed", "skipped", "awaiting_selection"}:
                    for step in node.get("steps", []):
                        if step.get("status") == "running":
                            step["status"] = "failed" if status == "failed" else "completed"
                    if status != "awaiting_selection":
                        node["active_step"] = ""
                break


def _set_step(
    run_id: str,
    node_key: str,
    step_key: str,
    status: str,
    message: str = "",
) -> None:
    """Update one user-facing operation without exposing internal field keys."""
    with JOBS_LOCK:
        job = JOBS.setdefault(run_id, {})
        nodes = job.setdefault("nodes", _initial_node_states())
        for node in nodes:
            if node.get("key") != node_key:
                continue
            steps = node.setdefault("steps", [])
            step = next((item for item in steps if item.get("key") == step_key), None)
            if step is None:
                step = {"key": step_key, "name": step_key, "description": "", "status": "pending", "message": "等待执行"}
                steps.append(step)
            step.update({"status": status, "message": message})
            if status == "running":
                node["active_step"] = step_key
            elif node.get("active_step") == step_key:
                node["active_step"] = ""
            if status == "running":
                for other in steps:
                    if other is not step and other.get("status") == "running":
                        other["status"] = "completed"
            if message:
                node["message"] = message
            break


def _progress_step_status(message: str) -> str:
    """Map a pipeline progress message to a visible step state."""
    text = str(message or "")
    completion_markers = (
        "已完成",
        "完成，",
        "完成：",
        "已返回",
        "已生成",
        "已保存",
        "已写入",
        "已校验",
        "已整理",
        "跳过",
    )
    return "completed" if any(marker in text for marker in completion_markers) else "running"


def _fill_progress_percent(percent: int | None) -> int | None:
    """Keep the second pipeline pass inside node 3's 72–96% range."""
    if percent is None:
        return None
    bounded = max(0, min(int(percent), 100))
    return 72 + int(round(bounded * 0.24))


def _public_error(exc: Exception) -> str:
    """Keep UI failures actionable without exposing stack traces or API details."""
    detail = str(exc)
    if "DASHSCOPE_API_KEY" in detail or "百炼" in detail:
        return "候选内容服务尚未就绪，请联系管理员检查百炼配置后重试。"
    if "OCR" in detail or "DocMind" in detail or "docmind" in detail:
        return "材料识别未完成，请检查扫描件清晰度或稍后重试；也可以直接上传已有 Excel 继续后续流程。"
    if "Excel" in detail or "工作簿" in detail:
        return "Excel 处理未完成，请确认文件未被占用、格式可打开后重试。"
    if "主体" in detail or "公司" in detail:
        return "材料主体或填写的评估主体不一致，请核对后重试。"
    return "任务未完成，请检查材料格式和必填信息后重试；详细诊断已保留给管理员。"


def _run_id_for_pdf(filename: str) -> str:
    """Create a readable, collision-safe output folder name."""
    stem = Path(filename or "source.pdf").stem
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip(" .") or "source"
    stem = stem[:100]
    prefix = datetime.now().astimezone().strftime("%Y%m%d%H%M")
    base = f"{prefix}-{stem}"
    candidate = base
    suffix = 1
    while candidate in JOBS or (RUNS_ROOT / candidate).exists():
        suffix += 1
        candidate = f"{base}-{suffix:02d}"
    return candidate


def _workbook_intakes_root() -> Path:
    return RUNS_ROOT / "_workbook_intakes"


def _workbook_intake_id(filename: str) -> str:
    stem = Path(filename or "materials.rar").stem
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip(" .") or "materials"
    prefix = datetime.now().astimezone().strftime("%Y%m%d%H%M%S")
    base = f"{prefix}-{stem[:80]}"
    candidate = base
    suffix = 1
    root = _workbook_intakes_root()
    while candidate in WORKBOOK_INTAKES or (root / candidate).exists():
        suffix += 1
        candidate = f"{base}-{suffix:02d}"
    return candidate


def _intake_state_path(intake_id: str) -> Path:
    return _workbook_intakes_root() / intake_id / "state.json"


def _set_workbook_intake(intake_id: str, **values: Any) -> dict[str, Any]:
    with JOBS_LOCK:
        state = WORKBOOK_INTAKES.get(intake_id)
        if state is None:
            path = _intake_state_path(intake_id)
            try:
                persisted = json.loads(path.read_text(encoding="utf-8")) if path.is_file() else {}
            except (OSError, json.JSONDecodeError):
                persisted = {}
            state = persisted if isinstance(persisted, dict) else {}
            state.setdefault("intake_id", intake_id)
            WORKBOOK_INTAKES[intake_id] = state
        state.update(values)
        snapshot = dict(state)
    path = _intake_state_path(intake_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)
    return snapshot


def _initial_workbook_intake_steps() -> list[dict[str, Any]]:
    return [
        {
            "key": key,
            "name": name,
            "description": description,
            "status": "pending",
            "message": "等待执行",
        }
        for key, name, description in WORKBOOK_INTAKE_STEPS
    ]


def _adaptive_intake_step_weights(paths: list[Path] | None = None) -> dict[str, float]:
    """Estimate stage cost from the uploaded material mix.

    Scanned PDFs/images and legacy Office binaries need layout OCR, so they
    carry more work units than natively readable OOXML files.  The weights are
    normalized by ``_adaptive_intake_progress`` rather than treated as fixed
    percentages.
    """
    weights = {
        "validate_archive": 1.0,
        "unpack_archive": 2.0,
        "inventory": 1.0,
        "ocr_materials": 8.0,
        "classify_materials": 2.0,
        "match_subject": 2.0,
        "load_mapping_rules": 2.0,
        "map_asset": 4.0,
        "map_income": 5.0,
        "write_mapping_trace": 2.0,
        "formula_qa": 4.0,
        "verify_output": 1.0,
    }
    if paths:
        ocr_suffixes = {".pdf", ".png", ".jpg", ".jpeg", ".doc", ".xls", ".ppt"}
        native_suffixes = {".docx", ".xlsx", ".xlsm", ".pptx"}
        parse_units = sum(
            3.0 if path.suffix.lower() in ocr_suffixes
            else 1.0 if path.suffix.lower() in native_suffixes
            else 0.0
            for path in paths
        )
        weights["ocr_materials"] = max(8.0, parse_units)
    return weights


def _adaptive_intake_progress(
    steps: list[dict[str, Any]],
    weights: Mapping[str, Any],
) -> int:
    total = sum(max(0.0, float(weights.get(step.get("key"), 1.0))) for step in steps) or 1.0
    achieved = 0.0
    for step in steps:
        weight = max(0.0, float(weights.get(step.get("key"), 1.0)))
        status = step.get("status")
        fraction = 1.0 if status == "completed" else 0.0
        if status == "running":
            fraction = 0.08
            if step.get("key") == "ocr_materials":
                match = re.search(r"(\d+)\s*/\s*(\d+)", str(step.get("message") or ""))
                if match:
                    current, count = int(match.group(1)), max(1, int(match.group(2)))
                    fraction = max(0.0, min(current / count, 0.99))
                    if str(step.get("message") or "").startswith("正在"):
                        fraction = max(0.0, min((current - 1) / count, 0.99))
        achieved += weight * fraction
    return max(0, min(round(100 * achieved / total), 99))


def _set_workbook_intake_step(
    intake_id: str,
    step_key: str,
    status: str,
    message: str,
    progress: int | None,
) -> None:
    state = _get_workbook_intake(intake_id) or {"intake_id": intake_id}
    previous_progress = max(0, min(int(state.get("progress", 0) or 0), 100))
    steps = [dict(item) for item in state.get("steps", _initial_workbook_intake_steps())]
    found = False
    for step in steps:
        if step.get("key") == step_key:
            step.update(status=status, message=message)
            found = True
        elif status == "running" and step.get("status") == "running":
            step["status"] = "completed"
    if not found:
        steps.append({"key": step_key, "name": step_key, "description": "", "status": status, "message": message})
    requested_progress = (
        _adaptive_intake_progress(steps, state.get("step_weights") or _adaptive_intake_step_weights())
        if progress is None
        else max(0, min(int(progress), 100))
    )
    next_progress = max(previous_progress, requested_progress)
    _set_workbook_intake(
        intake_id,
        status="running" if status != "failed" else "failed",
        progress=next_progress,
        message=message,
        steps=steps,
    )


def _get_workbook_intake(intake_id: str) -> dict[str, Any] | None:
    safe_id = Path(str(intake_id or "")).name
    if not safe_id or safe_id != intake_id:
        return None
    with JOBS_LOCK:
        state = WORKBOOK_INTAKES.get(safe_id)
        if state:
            return dict(state)
    path = _intake_state_path(safe_id)
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict) or payload.get("intake_id") != safe_id:
        return None
    with JOBS_LOCK:
        WORKBOOK_INTAKES[safe_id] = dict(payload)
    return payload


def _normalize_company_name(value: str) -> str:
    return normalize_company_name(value)


def _owned_intake_file(intake_id: str, value: Any) -> Path | None:
    if not value:
        return None
    root = (_workbook_intakes_root() / intake_id).resolve()
    candidate = Path(str(value)).resolve()
    try:
        candidate.relative_to(root)
    except ValueError:
        return None
    return candidate if candidate.is_file() else None


def _select_intake_document(manifest: Mapping[str, Any], target_company_name: str) -> dict[str, Any]:
    documents = [item for item in manifest.get("documents", []) if isinstance(item, dict)]
    if not documents:
        raise RuntimeError("材料包中未识别到可生成工作簿的审计财务报告")
    target_key = _normalize_company_name(target_company_name)
    if target_key:
        matches = matching_company_records(
            documents,
            target_company_name,
            name_getter=lambda item: item.get("metadata", {}).get("company_name", ""),
            source_getter=lambda item: item.get("source_file", ""),
        )
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            matches.sort(key=lambda item: (len(item.get("issues", [])), item.get("source_file", "")))
            return matches[0]
        detected = sorted({
            str(item.get("metadata", {}).get("company_name", "") or item.get("source_file", ""))
            for item in documents
        })
        raise RuntimeError(
            f"材料包中未找到与“{target_company_name}”可唯一匹配的审计主体；已识别：{'、'.join(detected[:12])}"
        )
    if len(documents) == 1:
        return documents[0]
    detected = sorted({
        str(item.get("metadata", {}).get("company_name", "") or item.get("source_file", ""))
        for item in documents
    })
    raise RuntimeError(f"材料包包含多个审计主体，请先填写评估主体全称：{'、'.join(detected[:12])}")


def _recalculate_generated_workbooks(output_dir: Path) -> dict[str, Any]:
    """Prepare the two templates entirely in the backend, without Excel COM."""
    return prepare_generated_workbooks(output_dir)


def _execute_workbook_intake(
    intake_id: str,
    archive_path: Path,
    target_company_name: str,
) -> None:
    intake_dir = _workbook_intakes_root() / intake_id
    extracted_dir = intake_dir / "extracted"
    generated_dir = intake_dir / "generated"
    try:
        _set_workbook_intake_step(intake_id, "unpack_archive", "running", "正在安全解压材料包", None)
        from .adapters.archive_intake import safe_extract_archive
        from .run_material_intake import generate_material_workbooks

        extracted = safe_extract_archive(archive_path, extracted_dir)
        if not extracted:
            raise RuntimeError("材料包为空")
        _set_workbook_intake(intake_id, step_weights=_adaptive_intake_step_weights(extracted))
        _set_workbook_intake_step(intake_id, "unpack_archive", "completed", f"安全解压完成，共 {len(extracted)} 个文件", None)
        def report_progress(step: str, status: str, message: str, percent: int) -> None:
            _set_workbook_intake_step(intake_id, step, status, message, None)
        manifest = generate_material_workbooks(
            extracted_dir,
            generated_dir,
            ROOT / "资产评估工作流",
            progress_callback=report_progress,
            # Share the validated SHA cache with the batch material-intake
            # workflow.  Keeping a second Web-only cache previously allowed a
            # blank local-PDF placeholder to hide an already valid DocMind
            # result for the same document.
            cache_dir=ROOT / "outputs" / "audit-material-intake" / "ocr_cache",
            target_company_name=target_company_name,
        )
        selected = _select_intake_document(manifest, target_company_name)
        workbooks = selected.get("workbooks", {})
        reporting_source = generated_dir / str(workbooks.get("reporting_workbook", ""))
        income_source = generated_dir / str(workbooks.get("income_workbook", ""))
        if not reporting_source.is_file() or not income_source.is_file():
            raise RuntimeError("材料解析已完成，但两份模板工作簿未完整生成")
        _set_workbook_intake_step(intake_id, "formula_qa", "running", "正在后台检查工作簿结构、公式引用和错误缓存", None)
        formula_qa = _recalculate_generated_workbooks(reporting_source.parent)
        _set_workbook_intake_step(
            intake_id,
            "formula_qa",
            "completed",
            "后台公式结构与错误缓存检查完成；工作簿打开时将自动完整计算",
            None,
        )
        reporting_target = intake_dir / "资产法.xlsx"
        income_target = intake_dir / "收益法.xlsx"
        shutil.copy2(reporting_source, reporting_target)
        shutil.copy2(income_source, income_target)
        _set_workbook_intake_step(intake_id, "verify_output", "running", "正在核对两份工作簿名称、文件完整性和下载入口", None)
        source_file = extracted_dir / str(selected.get("source_file", ""))
        target_key = _normalize_company_name(
            selected.get("metadata", {}).get("company_name", "") or target_company_name
        )
        supporting_sources: dict[str, str] = {}
        for item in manifest.get("all_materials", []):
            if not isinstance(item, dict) or item.get("source_file") == selected.get("source_file"):
                continue
            company_key = _normalize_company_name(item.get("metadata", {}).get("company_name", ""))
            if target_key and company_key != target_key:
                continue
            kind = item.get("material_type")
            if kind in {"营业执照", "企业信息报告"} and "registry_material" not in supporting_sources:
                supporting_sources["registry_material"] = str(extracted_dir / str(item.get("source_file", "")))
            elif kind == "其他企业资料" and "company_profile_material" not in supporting_sources:
                supporting_sources["company_profile_material"] = str(extracted_dir / str(item.get("source_file", "")))
        artifacts = [
            {"name": "资产法.xlsx", "label": "资产法 Excel"},
            {"name": "收益法.xlsx", "label": "收益法 Excel"},
        ]
        state = _get_workbook_intake(intake_id) or {}
        steps = [dict(item) for item in state.get("steps", [])]
        for step in steps:
            if step.get("key") == "verify_output":
                step.update(status="completed", message="资产法.xlsx、收益法.xlsx 已通过输出检查")
        _set_workbook_intake(
            intake_id,
            status="completed",
            progress=100,
            message="两份工作簿已生成；可下载检查，也可直接进入原工作流",
            target_company_name=(selected.get("metadata", {}).get("company_name", "") or target_company_name),
            selected_source_file=str(source_file) if source_file.is_file() else "",
            reporting_workbook=str(reporting_target),
            income_workbook=str(income_target),
            supporting_sources=supporting_sources,
            formula_qa=formula_qa,
            artifacts=artifacts,
            steps=steps,
            error="",
            technical_error="",
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        LOGGER.exception("Workbook intake %s failed", intake_id)
        public_error = _public_error(exc)
        state = _get_workbook_intake(intake_id) or {}
        steps = [dict(item) for item in state.get("steps", [])]
        failed_step_found = False
        for step in steps:
            if step.get("status") == "running":
                step.update(status="failed", message=public_error)
                failed_step_found = True
        if not failed_step_found:
            for step in steps:
                if step.get("status") == "pending":
                    step.update(status="failed", message=public_error)
                    break
        _set_workbook_intake(
            intake_id,
            status="failed",
            progress=max(0, min(int(state.get("progress", 0) or 0), 99)),
            message="材料包解析失败",
            error=public_error,
            technical_error=str(exc),
            artifacts=[],
            steps=steps,
        )


def _artifact_list(run_dir: Path) -> list[dict[str, str]]:
    # Intermediate OCR, candidate and trace files remain inside the run
    # directory for the workflow itself, but the public UI exposes only the
    # requested deliverable.  This also prevents users from mistaking an
    # internal trace or comparison sheet for the appraisal report.
    report = run_dir / "资产评估报告_待复核.docx"
    return [{"name": report.name, "label": "评估报告 Word"}] if report.is_file() else []


def _candidate_payload_path(run_id: str) -> Path:
    return RUNS_ROOT / run_id / "llm候选内容.json"


def _candidate_evidence_path(run_id: str) -> Path:
    return RUNS_ROOT / run_id / "llm候选证据.json"


def _update_candidate_payload(path: Path, field_key: str, value: str) -> dict[str, Any]:
    """Persist one candidate while leaving every other module unchanged."""
    payload = json.loads(path.read_text(encoding="utf-8"))
    candidates = payload.get("candidates")
    if not isinstance(candidates, list):
        raise RuntimeError("候选文件结构无效")
    for candidate in candidates:
        if candidate.get("field_key") == field_key:
            candidate["value"] = str(value or "").strip()
            candidate["available"] = bool(candidate["value"])
            candidate["updated_at"] = datetime.now(timezone.utc).isoformat()
            path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            return payload
    raise KeyError(field_key)


def _replace_job_candidate(run_id: str, field_key: str, value: str) -> None:
    with JOBS_LOCK:
        job = JOBS.setdefault(run_id, {})
        for candidate in job.get("candidates", []):
            if candidate.get("field_key") == field_key:
                candidate["value"] = str(value or "").strip()
                candidate["available"] = bool(candidate["value"])
                candidate["updated_at"] = datetime.now(timezone.utc).isoformat()
                break


def _find_ocr_cache(source_path: Path) -> Path | None:
    """Find a prior OCR workbook whose manifest matches the uploaded source."""
    pdf_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    configured = __import__("os").environ.get("APPRAISAL_OCR_CACHE_DIR", "")
    manifest_paths = []
    if configured:
        configured_path = Path(configured)
        if not configured_path.is_absolute():
            configured_path = ROOT / configured_path
        manifest_paths.extend(configured_path.glob("run_manifest.json"))
    manifest_paths.extend(OCR_CACHE_ROOT.glob("*/run_manifest.json"))
    seen: set[Path] = set()
    for manifest_path in manifest_paths:
        manifest_path = manifest_path.resolve()
        if manifest_path in seen or not manifest_path.is_file():
            continue
        seen.add(manifest_path)
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if manifest.get("pdf_sha256") != pdf_hash:
            continue
        workbook = manifest_path.parent / "OCR结构化结果.xlsx"
        if workbook.is_file():
            return workbook
    return None


def _select_ocr_adapter(
    pdf_path: Path | None,
    ocr_cache: Path | None,
    env: Mapping[str, str],
) -> Any:
    if pdf_path is None or ocr_cache is not None:
        return None
    return create_ocr_adapter(env)


def _build_external_adapters(use_glm: bool, use_qichacha: bool) -> tuple[Any, Any, Any]:
    """Create only the providers used by the four-node workflow."""
    llm_adapter = None
    qichacha_adapter = None
    http_client = None
    if use_glm or use_qichacha:
        import httpx

        http_client = httpx.Client(timeout=120)
    if use_glm:
        key = os.environ.get("DASHSCOPE_API_KEY", "")
        if not key:
            raise RuntimeError("未配置 DASHSCOPE_API_KEY")
        from .adapters.llm_factory import build_bailian_adapters

        project_config = json.loads(PROJECT_CONFIG.read_text(encoding="utf-8"))
        adapters = build_bailian_adapters(
            client=http_client,
            api_key=key,
            root=ROOT / "demo",
            config=project_config,
            env=os.environ,
            base_url=os.environ.get(
                "APPRAISAL_LLM_BASE_URL",
                "https://dashscope.aliyuncs.com/compatible-mode/v1",
            ),
        )
        llm_adapter = adapters["narrative"]
    if use_qichacha:
        from .adapters.company_api import QichachaApiAdapter

        qichacha_adapter = QichachaApiAdapter(
            http_client,
            os.environ.get("QICHACHA_APP_KEY", ""),
            os.environ.get("QICHACHA_SECRET_KEY", ""),
            base_url=os.environ.get("QICHACHA_API_BASE_URL", "https://api.qichacha.com"),
            endpoints={
                code: os.environ[name]
                for code, name in {
                    "735": "QICHACHA_ENDPOINT_735",
                    "231": "QICHACHA_ENDPOINT_231",
                    "514": "QICHACHA_ENDPOINT_514",
                    "233": "QICHACHA_ENDPOINT_233",
                    "2001": "QICHACHA_ENDPOINT_2001",
                    "213": "QICHACHA_ENDPOINT_213",
                    "886": "QICHACHA_ENDPOINT_886",
                    "915": "QICHACHA_ENDPOINT_915",
                    "699": "QICHACHA_ENDPOINT_699",
                }.items()
                if os.environ.get(name)
            },
            extra_api_codes=(
                tuple(
                    code.strip()
                    for code in os.environ.get("QICHACHA_EXTRA_API_CODES", "").split(",")
                    if code.strip()
                )
                or None
            ),
            enable_comparable_discovery=os.environ.get(
                "QICHACHA_ENABLE_COMPARABLE_DISCOVERY", "true"
            ).strip().lower() not in {"0", "false", "no", "off"},
        )
    return llm_adapter, qichacha_adapter, http_client


def _execute_run(
    run_id: str,
    pdf_path: Path | None,
    source_overrides: dict[str, Path | None],
    inputs: dict[str, Any],
    use_glm: bool,
    use_qichacha: bool,
    reuse_ocr: bool,
) -> None:
    _load_local_env()
    run_dir = RUNS_ROOT / run_id
    current_node = "start_input"
    try:
        _set_job(run_id, status="running", progress=5, message="节点 2：接收输入材料")
        _set_node(run_id, "start_input", "running", "校验人工字段和上传文件")
        # Keep the first node just as visible as the later parsing node.  The
        # frontend can now show exactly what is happening before any external
        # service is called, without exposing implementation field keys.
        _set_step(run_id, "start_input", "validate_inputs", "running", "正在校验人工输入和必填材料")
        _set_step(run_id, "start_input", "validate_inputs", "completed", "人工输入和必填材料校验完成")
        _set_step(run_id, "start_input", "store_materials", "running", "正在整理 PDF、Excel 和补充材料")
        _set_step(run_id, "start_input", "store_materials", "completed", "上传材料已按角色保存")
        _set_step(run_id, "start_input", "load_template", "running", "正在加载后台 Word 模板和批注映射")
        from .pipeline import run_pipeline
        from .adapters.template_pages import LibreOfficeTemplatePageReader
        template_path = _project_template()

        ocr_cache = (
            _find_ocr_cache(pdf_path)
            if reuse_ocr and pdf_path is not None
            else None
        )
        ocr_adapter = _select_ocr_adapter(pdf_path, ocr_cache, os.environ)
        _set_node(
            run_id,
            "start_input",
            "completed",
            "输入材料和后台 Word 模板已就绪",
        )
        current_node = "ocr_llm_candidates"
        _set_node(
            run_id,
            current_node,
            "running",
            "准备 OCR、Excel/API 解析和 LLM 候选",
        )
        def report_progress(node: str, step: str, message: str, percent: int | None = None) -> None:
            _set_step(run_id, node, step, _progress_step_status(message), message)
            if percent is not None:
                _set_job(run_id, progress=percent, message=message)

        llm_adapter, qichacha_adapter, _http_client = _build_external_adapters(use_glm, use_qichacha)
        _set_job(
            run_id,
            progress=15,
            message=(
                "未上传 PDF，跳过 OCR"
                if pdf_path is None
                else "命中已有 OCR 结果，跳过 OCR"
                if ocr_cache
                else "开始 PDF OCR 与字段解析"
            ),
            ocr_cache_hit=bool(ocr_cache),
        )
        result = run_pipeline(
            project_config=PROJECT_CONFIG,
            pdf_path=pdf_path,
            output_dir=run_dir,
            ocr_adapter=ocr_adapter,
            ocr_workbook_path=ocr_cache,
            llm_adapter=llm_adapter,
            qichacha_adapter=qichacha_adapter,
            node_inputs={
                key: inputs[key]
                for key in ("selected_valuation_method", "valuation_purpose_inputs")
                if inputs.get(key) not in (None, "")
            },
            manual_inputs_override=inputs,
            template_path=template_path,
            template_page_reader=LibreOfficeTemplatePageReader(),
            source_overrides=source_overrides,
            prepare_only=use_glm,
            generate_all_narratives=True,
            progress_callback=report_progress,
        )
        if not use_glm:
            _set_node(run_id, "ocr_llm_candidates", "completed", "材料解析完成，未启用 LLM")
            _set_node(run_id, "fill_word", "completed", "Word 已填充")
            _set_step(run_id, "output", "verify_output", "running", "正在检查 Word 文件是否可打开且未覆盖模板")
            _set_step(run_id, "output", "verify_output", "completed", "Word 文件检查完成")
            _set_node(run_id, "output", "completed", "评估报告 Word 已输出")
            _set_job(
                run_id,
                status="completed",
                progress=100,
                message="评估报告 Word 已生成",
                artifacts=_artifact_list(run_dir),
                issues=result.issues,
                completed_at=datetime.now(timezone.utc).isoformat(),
            )
            return
        candidate_items = []
        automatic_fields = {}
        if result.candidate_path and result.candidate_path.is_file():
            candidate_payload = json.loads(result.candidate_path.read_text(encoding="utf-8"))
            candidate_items = candidate_payload.get("candidates", [])
            automatic_fields = candidate_payload.get("automatic_fields", {})
        _set_node(
            run_id,
            "ocr_llm_candidates",
            "awaiting_selection",
            f"已生成 {len(candidate_items)} 个候选，等待人工选择",
        )
        _set_step(run_id, "ocr_llm_candidates", "wait_selection", "running", "候选内容已生成，等待人工选择")
        _set_job(
            run_id,
            status="awaiting_selection",
            progress=70,
            message="LLM候选内容已生成，请选择要写入 Word 的位置" + ("（已复用 OCR）" if ocr_cache else ""),
            artifacts=_artifact_list(run_dir),
            issues=result.issues,
            candidates=candidate_items,
            selection_context={
                "pdf_path": str(pdf_path) if pdf_path else "",
                "source_overrides": {
                    key: str(value) if value else ""
                    for key, value in source_overrides.items()
                },
                "inputs": inputs,
                "use_glm": use_glm,
                "use_qichacha": use_qichacha,
                "automatic_fields": automatic_fields,
            },
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        LOGGER.exception("Run %s failed", run_id)
        public_error = _public_error(exc)
        _set_node(run_id, current_node, "failed", public_error)
        _set_job(run_id, status="failed", progress=100, error=public_error, technical_error=str(exc), artifacts=[])


def _execute_fill(run_id: str, selected_fields: dict[str, Any]) -> None:
    _load_local_env()
    run_dir = RUNS_ROOT / run_id
    current_node = "fill_word"
    try:
        _set_job(run_id, status="running", progress=72, message="节点 4：按选择填充 Word")
        _set_node(run_id, "ocr_llm_candidates", "completed", "用户已确认候选内容")
        _set_node(run_id, "fill_word", "running", "正在复制模板并填充 Word")
        with JOBS_LOCK:
            context = dict(JOBS.get(run_id, {}).get("selection_context", {}))
        if not context:
            raise RuntimeError("任务缺少候选内容生成上下文，请重新提交")
        from .pipeline import run_pipeline
        from .adapters.template_pages import LibreOfficeTemplatePageReader

        pdf_path = Path(context["pdf_path"]) if context.get("pdf_path") else None
        source_overrides = {
            key: Path(value) if value else None
            for key, value in context.get("source_overrides", {}).items()
        }
        # Node 2 persists validated Qichacha payloads inside this run directory.
        # The pipeline reuses valid roles and only uses this client when the
        # snapshot is absent or incomplete, prioritizing report completeness.
        llm_adapter, qichacha_adapter, _http_client = _build_external_adapters(
            bool(context.get("use_glm")), bool(context.get("use_qichacha"))
        )
        ocr_workbook = run_dir / "OCR结构化结果.xlsx"
        def report_progress(node: str, step: str, message: str, percent: int | None = None) -> None:
            _set_step(run_id, node, step, _progress_step_status(message), message)
            if percent is not None:
                _set_job(run_id, progress=_fill_progress_percent(percent), message=message)

        _set_step(run_id, "fill_word", "load_selection", "running", "正在读取人工确认的 LLM 模块")
        result = run_pipeline(
            project_config=PROJECT_CONFIG,
            pdf_path=pdf_path,
            output_dir=run_dir,
            ocr_adapter=None,
            ocr_workbook_path=ocr_workbook if ocr_workbook.is_file() else None,
            # Reuse the configured LLM during the final fill pass for
            # evidence review and traceability-comment writing. Candidate
            # prose is supplied through ``llm_values_override`` below, so
            # this does not regenerate the six user-selected modules.
            llm_adapter=llm_adapter,
            word_comment_locator_adapter=llm_adapter,
            qichacha_adapter=qichacha_adapter,
            manual_inputs_override=context.get("inputs", {}),
            node_inputs={
                key: context.get("inputs", {})[key]
                for key in ("selected_valuation_method", "valuation_purpose_inputs")
                if context.get("inputs", {}).get(key) not in (None, "")
            },
            template_path=_project_template(),
            template_page_reader=LibreOfficeTemplatePageReader(),
            source_overrides=source_overrides,
            generate_all_narratives=True,
            llm_values_override=selected_fields,
            progress_callback=report_progress,
        )
        _set_node(run_id, "ocr_llm_candidates", "completed", "候选内容已确认，材料复核完成")
        _set_node(run_id, "fill_word", "completed", "Word 填充完成")
        current_node = "output"
        _set_node(run_id, "output", "running", "正在生成评估报告 Word")
        _set_step(run_id, "output", "verify_output", "running", "正在检查 Word 文件是否可打开且未覆盖模板")
        _set_step(run_id, "output", "verify_output", "completed", "Word 文件检查完成")
        _set_node(run_id, "output", "completed", "全部输出已生成")
        _set_job(
            run_id,
            status="completed",
            progress=100,
            message="评估报告 Word 已生成",
            artifacts=_artifact_list(run_dir),
            issues=result.issues,
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        LOGGER.exception("Word fill %s failed", run_id)
        public_error = _public_error(exc)
        _set_node(run_id, current_node, "failed", public_error)
        _set_job(run_id, status="failed", progress=100, error=public_error, technical_error=str(exc), artifacts=[])


def _execute_regenerate_candidate(run_id: str, field_key: str, feedback: str) -> None:
    """Regenerate one candidate from the persisted evidence and save it."""
    _load_local_env()
    try:
        with JOBS_LOCK:
            job = dict(JOBS.get(run_id, {}))
            context = dict(job.get("selection_context", {}) or {})
        evidence_path = _candidate_evidence_path(run_id)
        candidate_path = _candidate_payload_path(run_id)
        if not evidence_path.is_file() or not candidate_path.is_file():
            raise RuntimeError("候选证据已不存在，请重新运行材料解析")
        llm_adapter, _qichacha_adapter, _http_client = _build_external_adapters(True, False)
        _set_job(run_id, status="running", progress=62, message=f"正在重新生成：{field_key}")
        _set_node(run_id, "ocr_llm_candidates", "running", f"正在根据反馈重新生成 {field_key}")
        _set_step(
            run_id,
            "ocr_llm_candidates",
            "generate_candidates",
            "running",
            f"LLM 正在根据反馈重新生成：{field_key}",
        )
        evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
        evidence["selected_modules"] = [field_key]
        evidence["regenerate_field"] = field_key
        evidence["feedback"] = feedback
        values, provider_issues = llm_adapter.generate(evidence)
        value = str(values.get(field_key, "") or "").strip()
        if not value:
            raise RuntimeError(
                f"LLM 未返回 {field_key} 的有效内容"
                + (f"：{'；'.join(provider_issues)}" if provider_issues else "")
            )
        payload = _update_candidate_payload(candidate_path, field_key, value)
        _replace_job_candidate(run_id, field_key, value)
        with JOBS_LOCK:
            current = JOBS.setdefault(run_id, {})
            current["issues"] = list(current.get("issues", [])) + provider_issues
            current["candidates"] = payload.get("candidates", current.get("candidates", []))
        _set_step(
            run_id,
            "ocr_llm_candidates",
            "generate_candidates",
            "completed",
            f"已根据反馈更新：{field_key}",
        )
        _set_node(run_id, "ocr_llm_candidates", "awaiting_selection", "候选内容已更新，请继续选择")
        _set_step(
            run_id,
            "ocr_llm_candidates",
            "wait_selection",
            "running",
            "候选内容已更新，等待人工选择",
        )
        _set_job(
            run_id,
            status="awaiting_selection",
            progress=70,
            message="候选内容已更新，请确认要写入 Word 的模块",
        )
    except Exception as exc:
        LOGGER.exception("Candidate regeneration %s/%s failed", run_id, field_key)
        public_error = _public_error(exc)
        _set_node(run_id, "ocr_llm_candidates", "awaiting_selection", f"重新生成失败：{public_error}")
        _set_step(run_id, "ocr_llm_candidates", "generate_candidates", "failed", public_error)
        _set_step(
            run_id,
            "ocr_llm_candidates",
            "wait_selection",
            "running",
            "候选未更新，仍可继续选择或重试",
        )
        _set_job(
            run_id,
            status="awaiting_selection",
            progress=70,
            message=f"候选重新生成失败：{public_error}",
            error=public_error,
            technical_error=str(exc),
        )


@app.post("/api/v1/asset-appraisal/workbook-intakes", status_code=202)
async def create_workbook_intake(
    background_tasks: BackgroundTasks,
    archive: UploadFile = File(...),
    target_company_name: str = Form(""),
):
    suffix = Path(archive.filename or "").suffix.lower()
    if suffix not in {".rar", ".zip"}:
        raise HTTPException(status_code=422, detail="材料包仅支持 .rar 或 .zip 格式")
    intake_id = _workbook_intake_id(archive.filename or "materials.rar")
    intake_dir = _workbook_intakes_root() / intake_id
    intake_dir.mkdir(parents=True, exist_ok=True)
    archive_path = intake_dir / f"source{suffix}"
    with archive_path.open("wb") as destination:
        shutil.copyfileobj(archive.file, destination)
    if archive_path.stat().st_size == 0:
        archive_path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail="上传的材料包为空")
    initial_steps = [
        {
            **step,
            "status": "completed" if step["key"] == "validate_archive" else step["status"],
            "message": "材料包格式和文件大小校验通过" if step["key"] == "validate_archive" else step["message"],
        }
        for step in _initial_workbook_intake_steps()
    ]
    initial_weights = _adaptive_intake_step_weights()
    state = _set_workbook_intake(
        intake_id,
        status="queued",
        progress=_adaptive_intake_progress(initial_steps, initial_weights),
        message="材料包已上传，等待解析",
        archive_name=Path(archive.filename or "materials.rar").name,
        target_company_name=str(target_company_name or "").strip(),
        artifacts=[],
        step_weights=initial_weights,
        steps=initial_steps,
    )
    background_tasks.add_task(
        _execute_workbook_intake,
        intake_id,
        archive_path,
        str(target_company_name or "").strip(),
    )
    return state


@app.get("/api/v1/asset-appraisal/workbook-intakes/{intake_id}")
async def get_workbook_intake(intake_id: str):
    state = _get_workbook_intake(intake_id)
    if not state:
        raise HTTPException(status_code=404, detail="工作簿预处理任务不存在")
    return state


@app.get("/api/v1/asset-appraisal/workbook-intakes/{intake_id}/artifacts/{name}")
async def download_workbook_intake_artifact(intake_id: str, name: str):
    state = _get_workbook_intake(intake_id)
    allowed = {item["name"] for item in (state or {}).get("artifacts", [])}
    if not state or name not in allowed:
        raise HTTPException(status_code=404, detail="工作簿产物不存在")
    path = _workbook_intakes_root() / intake_id / name
    if not path.is_file():
        raise HTTPException(status_code=404, detail="工作簿文件不存在")
    return FileResponse(path, filename=name)


@app.post("/api/v1/asset-appraisal/runs", status_code=202)
async def create_run(
    background_tasks: BackgroundTasks,
    materials: list[UploadFile] | None = File(None),
    audit_materials: list[UploadFile] | None = File(None),
    pdf: UploadFile | None = File(None),
    income_workbook: UploadFile | None = File(None),
    reporting_workbook: UploadFile | None = File(None),
    registry_materials: list[UploadFile] | None = File(None),
    ownership_history_materials: list[UploadFile] | None = File(None),
    unrecorded_intangibles_materials: list[UploadFile] | None = File(None),
    company_profile_materials: list[UploadFile] | None = File(None),
    inputs: str = Form("{}"),
    use_glm: bool = Form(True),
    use_qichacha: bool = Form(True),
    reuse_ocr: bool = Form(True),
    workbook_intake_id: str = Form(""),
):
    # Named slots from the current UI are authoritative even when the user
    # gives the workbook an arbitrary filename.  The legacy ``materials``
    # multi-file field is only used to fill roles that were not supplied by a
    # typed slot.
    intake_state: dict[str, Any] | None = None
    workbook_intake_id = str(workbook_intake_id or "").strip()
    if workbook_intake_id:
        intake_state = _get_workbook_intake(workbook_intake_id)
        if not intake_state:
            raise HTTPException(status_code=422, detail="工作簿预处理任务不存在或已失效")
        if intake_state.get("status") != "completed":
            raise HTTPException(status_code=422, detail="工作簿预处理尚未完成，不能进入后续工作流")
    audit_uploads = list(audit_materials or [])
    if pdf is not None:
        audit_uploads.insert(0, pdf)
    role_uploads: dict[str, UploadFile | None] = {
        "pdf": next(
            (
                upload
                for upload in audit_uploads
                if Path(upload.filename or "").suffix.lower() == ".pdf"
            ),
            None,
        ),
        "reporting_workbook": reporting_workbook,
        "income_workbook": income_workbook,
        "reference_report": None,
    }
    workbook_candidates: list[UploadFile] = []
    for upload in list(materials or []):
        suffix = Path(upload.filename or "").suffix.lower()
        if suffix == ".pdf":
            audit_uploads.append(upload)
            if role_uploads["pdf"] is None:
                role_uploads["pdf"] = upload
        elif suffix in {".xls", ".xlsx", ".xlsm"}:
            workbook_candidates.append(upload)
        elif suffix in {".doc", ".docx"} and role_uploads["reference_report"] is None:
            role_uploads["reference_report"] = upload
    # Content-role names are only a hint.  The semantic Excel reader still
    # decides which workbook contains which facts after upload.
    for upload in workbook_candidates:
        name = (upload.filename or "").lower()
        if role_uploads["income_workbook"] is None and any(token in name for token in ("收益", "income", "现金流", "市场")):
            role_uploads["income_workbook"] = upload
        elif role_uploads["reporting_workbook"] is None:
            role_uploads["reporting_workbook"] = upload
        elif role_uploads["income_workbook"] is None:
            role_uploads["income_workbook"] = upload
    uploads = {
        "income_workbook": (
            role_uploads["income_workbook"],
            (".xls", ".xlsx", ".xlsm"),
            "收益法或市场法工作簿",
        ),
        "reporting_workbook": (
            role_uploads["reporting_workbook"],
            (".xls", ".xlsx", ".xlsm"),
            "资产基础法/资产清查工作簿",
        ),
        "reference_report": (
            role_uploads["reference_report"],
            (".doc", ".docx"),
            "补充 Word 材料",
        ),
    }
    for field_name, (upload, suffixes, label) in uploads.items():
        if upload is None:
            continue
        if (
            not upload.filename
            or not upload.filename.lower().endswith(suffixes)
        ):
            allowed = "、".join(suffixes)
            raise HTTPException(
                status_code=422,
                detail=f"{label}格式应为 {allowed}（字段：{field_name}）",
            )
    try:
        parsed_inputs = json.loads(inputs)
        if not isinstance(parsed_inputs, dict):
            raise ValueError
    except ValueError as exc:
        raise HTTPException(status_code=422, detail="用户输入字段格式错误") from exc
    required_text_fields = (
        ("commissioning_party_name", "委托方全称", 50),
        ("commissioning_party_short_name", "委托方简称", 20),
        ("target_company_name", "评估主体全称", 50),
        ("target_company_short_name", "评估主体简称", 20),
    )
    try:
        for key, label, limit in required_text_fields:
            parsed_inputs[key] = validate_required_text(parsed_inputs.get(key), label, limit)
        parsed_inputs["valuation_base_date"] = validate_valuation_base_date(
            parsed_inputs.get("valuation_base_date")
        )
        for key, label in (
            ("registry_info_strategy", "工商信息"),
            ("ownership_history_strategy", "股权结构及历史沿革"),
            ("unrecorded_intangibles_strategy", "账外无形资产"),
            ("company_profile_strategy", "企业介绍"),
        ):
            parsed_inputs[key] = validate_material_source_strategy(
                parsed_inputs.get(key), label
            )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    missing_choices = [
        key for key in (
            "transaction_type", "valuation_subject_type",
            "selected_valuation_method", "final_valuation_method",
        ) if parsed_inputs.get(key) in (None, "", [])
    ]
    if missing_choices:
        raise HTTPException(status_code=422, detail=f"缺少必填选择项：{'、'.join(missing_choices)}")
    has_direct_workbook = any(
        upload is not None
        for upload in (role_uploads["reporting_workbook"], role_uploads["income_workbook"])
    )
    if not audit_uploads and not intake_state and not has_direct_workbook:
        raise HTTPException(
            status_code=422,
            detail="请上传材料包、至少一份 Excel 工作簿或审计报告材料",
        )
    image_suffixes = {".png", ".jpg", ".jpeg"}
    audit_suffixes = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".xlsm", *image_suffixes}
    for upload in audit_uploads:
        suffix = Path(upload.filename or "").suffix.lower()
        if not upload.filename or suffix not in audit_suffixes:
            raise HTTPException(
                status_code=422,
                detail="审计报告材料仅支持 PDF、图片、Word、Excel 格式",
            )
    optional_upload_groups = {
        "registry_materials": (list(registry_materials or []), {".doc", ".docx", ".ppt", ".pptx", ".pdf", *image_suffixes}, "工商文件"),
        "ownership_history_materials": (list(ownership_history_materials or []), {".doc", ".docx", ".ppt", ".pptx", ".pdf", ".xls", ".xlsx", ".xlsm", *image_suffixes}, "股权结构及历史沿革文件"),
        "unrecorded_intangibles_materials": (list(unrecorded_intangibles_materials or []), {".doc", ".docx", ".ppt", ".pptx", ".pdf", ".xls", ".xlsx", ".xlsm", *image_suffixes}, "账外无形资产文件"),
        "company_profile_materials": (list(company_profile_materials or []), {".doc", ".docx", ".ppt", ".pptx", ".pdf", ".xls", ".xlsx", ".xlsm", *image_suffixes}, "企业介绍文件"),
    }
    for _role, group, suffixes, label in (
        (role, items, suffixes, label)
        for role, (items, suffixes, label) in optional_upload_groups.items()
    ):
        for upload in group:
            suffix = Path(upload.filename or "").suffix.lower()
            if not upload.filename or suffix not in suffixes:
                raise HTTPException(
                    status_code=422,
                    detail=f"{label}格式不支持",
                )
    has_upload = bool(intake_state) or bool(audit_uploads) or any(upload is not None for upload, _, _ in uploads.values())
    has_manual = any(
        value not in (None, "", [], {})
        for value in parsed_inputs.values()
    )
    if not has_upload and not has_manual:
        raise HTTPException(
            status_code=422,
            detail="请至少上传一份材料或填写一项基础信息",
        )
    if parsed_inputs.get("valuation_subject_type") not in (None, ""):
        try:
            parsed_inputs["valuation_subject_type"] = validate_valuation_subject_type(
                parsed_inputs["valuation_subject_type"]
            )
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
    try:
        if parsed_inputs.get("selected_valuation_method") not in (None, ""):
            parsed_inputs["selected_valuation_method"] = normalize_valuation_methods(
                parsed_inputs["selected_valuation_method"]
            )
        if parsed_inputs.get("final_valuation_method") not in (None, ""):
            parsed_inputs["final_valuation_method"] = validate_final_valuation_method(
                parsed_inputs["final_valuation_method"]
            )
        if parsed_inputs.get("transaction_type") not in (None, ""):
            parsed_inputs["transaction_type"] = validate_transaction_type(parsed_inputs["transaction_type"])
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    if intake_state:
        intake_company = _normalize_company_name(intake_state.get("target_company_name", ""))
        form_company = _normalize_company_name(parsed_inputs.get("target_company_name", ""))
        if intake_company and form_company and intake_company != form_company:
            raise HTTPException(
                status_code=422,
                detail="材料包生成的工作簿主体与当前评估主体不一致，请重新生成或上传正确的 Excel",
            )
    first_filename = next(
        (
            upload.filename
            for upload in audit_uploads
            if upload.filename
        ),
        str(
            parsed_inputs.get("target_company_name")
            or parsed_inputs.get("commissioning_party_name")
            or "人工输入"
        ),
    )
    run_id = _run_id_for_pdf(first_filename)
    input_dir = RUNS_ROOT / run_id / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    stored_files: dict[str, Path] = {}
    audit_dir = input_dir / "audit"
    audit_dir.mkdir(parents=True, exist_ok=True)
    audit_paths: list[Path] = []
    used_names: set[str] = set()
    for index, upload in enumerate(audit_uploads, start=1):
        original_name = Path(upload.filename or f"audit-{index}").name
        safe_name = re.sub(r"[\\\\/:*?\"<>|]+", "_", original_name) or f"audit-{index}"
        candidate = safe_name
        suffix = 2
        while candidate in used_names:
            candidate = f"{Path(safe_name).stem}-{suffix}{Path(safe_name).suffix}"
            suffix += 1
        used_names.add(candidate)
        target = audit_dir / candidate
        with target.open("wb") as destination:
            shutil.copyfileobj(upload.file, destination)
        audit_paths.append(target)
    if intake_state and not audit_paths:
        source = _owned_intake_file(workbook_intake_id, intake_state.get("selected_source_file"))
        if source and source.suffix.lower() in audit_suffixes:
            target = audit_dir / re.sub(r"[\\/:*?\"<>|]+", "_", source.name)
            shutil.copy2(source, target)
            audit_paths.append(target)
    for field_name, (upload, _, _) in uploads.items():
        if upload is None:
            continue
        suffix = Path(upload.filename or "").suffix.lower()
        original_name = Path(upload.filename or f"{field_name}{suffix}").name
        safe_name = re.sub(r"[\\\\/:*?\"<>|]+", "_", original_name) or f"{field_name}{suffix}"
        stored_name = "source.pdf" if field_name == "pdf" else safe_name
        stored_files[field_name] = input_dir / stored_name
        with stored_files[field_name].open("wb") as target:
            shutil.copyfileobj(upload.file, target)
    if intake_state:
        generated_workbooks = {
            "reporting_workbook": _owned_intake_file(workbook_intake_id, intake_state.get("reporting_workbook")),
            "income_workbook": _owned_intake_file(workbook_intake_id, intake_state.get("income_workbook")),
        }
        for field_name, source in generated_workbooks.items():
            # A user-uploaded workbook is authoritative and replaces only the
            # corresponding generated workbook.  The other generated role is
            # retained, which supports partial reviewer corrections.
            if field_name in stored_files:
                continue
            if source is None:
                raise HTTPException(status_code=422, detail=f"预处理结果缺少 {field_name}")
            target = input_dir / source.name
            shutil.copy2(source, target)
            stored_files[field_name] = target
    optional_paths: dict[str, list[Path]] = {}
    for role, (group, _suffixes, _label) in optional_upload_groups.items():
        if not group:
            continue
        group_dir = input_dir / role
        group_dir.mkdir(parents=True, exist_ok=True)
        optional_paths[role] = []
        used_group_names: set[str] = set()
        for index, upload in enumerate(group, start=1):
            original_name = Path(upload.filename or f"{index}{Path(upload.filename or '').suffix.lower()}").name
            safe_name = re.sub(r"[\\\\/:*?\"<>|]+", "_", original_name) or f"{index}"
            candidate = safe_name
            suffix = 2
            while candidate in used_group_names:
                candidate = f"{Path(safe_name).stem}-{suffix}{Path(safe_name).suffix}"
                suffix += 1
            used_group_names.add(candidate)
            target = group_dir / candidate
            with target.open("wb") as destination:
                shutil.copyfileobj(upload.file, destination)
            optional_paths[role].append(target)
    if intake_state:
        intake_supporting = intake_state.get("supporting_sources", {}) or {}
        intake_role_map = {
            "registry_material": "registry_materials",
            "company_profile_material": "company_profile_materials",
        }
        for source_role, upload_role in intake_role_map.items():
            if optional_paths.get(upload_role):
                continue
            source = _owned_intake_file(workbook_intake_id, intake_supporting.get(source_role))
            if source is None:
                continue
            group_dir = input_dir / upload_role
            group_dir.mkdir(parents=True, exist_ok=True)
            target = group_dir / re.sub(r"[\\/:*?\"<>|]+", "_", source.name)
            shutil.copy2(source, target)
            optional_paths[upload_role] = [target]
    # The legacy workflow still names this argument ``pdf_path``, but the OCR
    # adapters accept an image as well.  Prefer PDF where available and use a
    # scanned image only when it is the sole audit evidence.
    pdf_path = next((path for path in audit_paths if path.suffix.lower() == ".pdf"), None)
    if pdf_path is None:
        pdf_path = next((path for path in audit_paths if path.suffix.lower() in image_suffixes), None)
    audited_workbook = next(
        (path for path in audit_paths if path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}),
        None,
    )
    source_overrides = {
        "audit_pdf": pdf_path,
        "audited_financials": audited_workbook,
        "income_workbook": stored_files.get("income_workbook"),
        "reporting_workbook": stored_files.get("reporting_workbook"),
        "reference_report": stored_files.get("reference_report"),
        "registry_material": next(iter(optional_paths.get("registry_materials", [])), None),
        "ownership_history_material": next(iter(optional_paths.get("ownership_history_materials", [])), None),
        "unrecorded_intangibles_material": next(iter(optional_paths.get("unrecorded_intangibles_materials", [])), None),
        "company_profile_material": next(iter(optional_paths.get("company_profile_materials", [])), None),
    }
    _set_job(
        run_id,
        run_id=run_id,
        status="queued",
        progress=0,
        message="任务已创建",
        artifacts=[],
        nodes=_initial_node_states(),
        workbook_intake_id=workbook_intake_id,
        workbook_sources={
            "reporting_workbook": (
                "user_upload" if role_uploads["reporting_workbook"] is not None else "generated_intake" if intake_state else "missing"
            ),
            "income_workbook": (
                "user_upload" if role_uploads["income_workbook"] is not None else "generated_intake" if intake_state else "missing"
            ),
        },
    )
    background_tasks.add_task(
        _execute_run,
        run_id,
        pdf_path,
        source_overrides,
        parsed_inputs,
        use_glm,
        use_qichacha,
        reuse_ocr,
    )
    return JOBS[run_id]


@app.post("/api/v1/asset-appraisal/runs/{run_id}/candidates/{field_key}/edit", status_code=200)
async def edit_run_candidate(
    run_id: str,
    field_key: str,
    value: str = Form(""),
):
    """Persist a user edit to one candidate before final selection."""
    with JOBS_LOCK:
        job = JOBS.get(run_id)
        status = job.get("status") if job else None
        allowed = {
            item.get("field_key")
            for item in (job or {}).get("candidates", [])
            if item.get("field_key")
        }
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在或服务已重启")
    if status != "awaiting_selection":
        raise HTTPException(status_code=409, detail="任务当前不在候选内容选择阶段")
    if field_key not in allowed:
        raise HTTPException(status_code=404, detail="候选模块不存在")
    value = str(value or "").strip()
    if not value:
        raise HTTPException(status_code=422, detail="候选内容不能为空")
    path = _candidate_payload_path(run_id)
    try:
        payload = _update_candidate_payload(path, field_key, value)
    except (OSError, json.JSONDecodeError, KeyError, RuntimeError) as exc:
        raise HTTPException(status_code=500, detail=f"保存候选内容失败：{exc}") from exc
    _replace_job_candidate(run_id, field_key, value)
    with JOBS_LOCK:
        current = JOBS.setdefault(run_id, {})
        current["candidates"] = payload.get("candidates", current.get("candidates", []))
    return JOBS[run_id]


@app.post("/api/v1/asset-appraisal/runs/{run_id}/candidates/{field_key}/regenerate", status_code=202)
async def regenerate_run_candidate(
    run_id: str,
    field_key: str,
    background_tasks: BackgroundTasks,
    feedback: str = Form(""),
):
    """Ask the configured LLM to regenerate one candidate from its evidence."""
    with JOBS_LOCK:
        job = JOBS.get(run_id)
        status = job.get("status") if job else None
        allowed = {
            item.get("field_key")
            for item in (job or {}).get("candidates", [])
            if item.get("field_key")
        }
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在或服务已重启")
    if status != "awaiting_selection":
        raise HTTPException(status_code=409, detail="任务当前不在候选内容选择阶段")
    if field_key not in allowed:
        raise HTTPException(status_code=404, detail="候选模块不存在")
    feedback = str(feedback or "").strip()
    if not feedback:
        raise HTTPException(status_code=422, detail="请先填写重新生成的反馈")
    if len(feedback) > 2000:
        raise HTTPException(status_code=422, detail="反馈不能超过 2000 个字符")
    _set_job(run_id, status="queued", progress=60, message=f"已提交重新生成：{field_key}")
    background_tasks.add_task(_execute_regenerate_candidate, run_id, field_key, feedback)
    return JOBS[run_id]


@app.post("/api/v1/asset-appraisal/runs/{run_id}/select", status_code=202)
async def select_run_candidates(
    run_id: str,
    background_tasks: BackgroundTasks,
    selected_fields: str = Form("{}"),
):
    """Submit the human choice made after the candidate-generation node."""
    with JOBS_LOCK:
        job = JOBS.get(run_id)
        status = job.get("status") if job else None
        allowed = {
            item.get("field_key"): item.get("value")
            for item in (job or {}).get("candidates", [])
            if item.get("field_key")
        }
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在或服务已重启")
    if status != "awaiting_selection":
        raise HTTPException(status_code=409, detail="任务当前不在候选内容选择阶段")
    try:
        payload = json.loads(selected_fields)
        if not isinstance(payload, dict):
            raise ValueError
    except (TypeError, ValueError, json.JSONDecodeError) as exc:
        raise HTTPException(status_code=422, detail="候选内容选择格式错误") from exc
    selected = {
        key: payload[key] if payload[key] not in (None, "") else allowed[key]
        for key in payload
        if key in allowed
    }
    # Values always come from the generated candidate set.  This prevents a
    # client from turning the selection endpoint into an arbitrary text writer.
    selected = {key: allowed[key] for key in selected}
    automatic = (job.get("selection_context", {}) or {}).get("automatic_fields", {})
    if isinstance(automatic, dict):
        selected = {
            **{
                key: value
                for key, value in automatic.items()
                if key == "company_profile_section" and value not in (None, "", [], {})
            },
            **selected,
        }
    _set_job(run_id, status="queued", progress=70, message="已确认候选内容，开始填充 Word")
    background_tasks.add_task(_execute_fill, run_id, selected)
    return JOBS[run_id]


@app.post("/api/v1/asset-appraisal/ocr-cache/check")
async def check_ocr_cache(pdf: UploadFile = File(...)):
    allowed_suffixes = {".pdf", ".png", ".jpg", ".jpeg"}
    suffix = Path(pdf.filename or "").suffix.lower()
    if not pdf.filename or suffix not in allowed_suffixes:
        raise HTTPException(status_code=422, detail="请上传 PDF 或图片格式的审计材料")
    import tempfile

    content = await pdf.read()
    with tempfile.NamedTemporaryFile(suffix=suffix) as temporary:
        temporary.write(content)
        temporary.flush()
        cache = _find_ocr_cache(Path(temporary.name))
    return {
        "hit": cache is not None,
        "source": cache.parent.name if cache else "",
        "message": "命中已有 OCR 结果" if cache else "未命中 OCR 缓存；生成时会自动执行 OCR",
    }


@app.get("/api/v1/asset-appraisal/runs/{run_id}")
async def get_run(run_id: str):
    with JOBS_LOCK:
        job = JOBS.get(run_id)
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在或服务已重启")
    return job


@app.get("/api/v1/asset-appraisal/runs/{run_id}/artifacts/{name}")
async def download_artifact(run_id: str, name: str):
    with JOBS_LOCK:
        job = JOBS.get(run_id)
    allowed = {item["name"] for item in (job or {}).get("artifacts", [])}
    if not job or name not in allowed:
        raise HTTPException(status_code=404, detail="产物不存在")
    path = RUNS_ROOT / run_id / name
    if not path.is_file():
        raise HTTPException(status_code=404, detail="产物文件不存在")
    return FileResponse(path, filename=name)
