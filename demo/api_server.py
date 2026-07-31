"""Local HTTP bridge for the Vue asset-appraisal workbench.

The business pipeline remains in ``demo.pipeline``. This module only handles
uploads, job state and artifact downloads; c2m can replace it with its own
authenticated task service without changing the domain workflow.
"""

from __future__ import annotations

import json
import hashlib
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
    validate_report_serial_input,
    validate_final_valuation_method,
    validate_transaction_type,
    validate_valuation_subject_type,
)

ROOT = Path(__file__).resolve().parents[1]
PROJECT_CONFIG = ROOT / "demo/projects/tongfu.yaml"
RUNS_ROOT = ROOT / "runs/web"
OCR_CACHE_ROOT = ROOT / "runs"
JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = Lock()

PUBLIC_NODES = (
    ("start_input", "节点 1：开始 / 输入", "接收人工字段和上传材料"),
    ("ocr_llm_candidates", "节点 2：材料解析 / LLM 候选", "OCR、Excel、企查查解析并生成候选"),
    ("fill_word", "节点 3：填充 Word", "写入确定性字段和用户选中的候选"),
    ("output", "节点 4：结果输出", "生成评估报告 Word"),
)

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


def _initial_node_states() -> list[dict[str, str]]:
    return [
        {"key": key, "name": name, "description": description, "status": "pending", "message": "等待执行"}
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
                break


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


def _artifact_list(run_dir: Path) -> list[dict[str, str]]:
    # Intermediate OCR, candidate and trace files remain inside the run
    # directory for the workflow itself, but the public UI exposes only the
    # requested deliverable.  This also prevents users from mistaking an
    # internal trace or comparison sheet for the appraisal report.
    report = run_dir / "资产评估报告_待复核.docx"
    return [{"name": report.name, "label": "评估报告 Word"}] if report.is_file() else []


def _find_ocr_cache(pdf_path: Path) -> Path | None:
    """Find a prior OCR workbook whose manifest matches this PDF hash."""
    pdf_hash = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
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
        _set_job(run_id, status="running", progress=5, message="节点 1：接收输入材料")
        _set_node(run_id, "start_input", "running", "校验人工字段和上传文件")
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
        )
        if not use_glm:
            _set_node(run_id, "ocr_llm_candidates", "completed", "材料解析完成，未启用 LLM")
            _set_node(run_id, "fill_word", "completed", "Word 已填充")
            _set_node(run_id, "output", "completed", "评估报告 Word 已输出")
            _set_job(
                run_id,
                status="completed",
                progress=100,
                message="评估报告 Word 已生成",
                artifacts=_artifact_list(run_dir),
                issues=[],
                completed_at=datetime.now(timezone.utc).isoformat(),
            )
            return
        candidate_items = []
        if result.candidate_path and result.candidate_path.is_file():
            candidate_items = json.loads(
                result.candidate_path.read_text(encoding="utf-8")
            ).get("candidates", [])
        _set_node(
            run_id,
            "ocr_llm_candidates",
            "awaiting_selection",
            f"已生成 {len(candidate_items)} 个候选，等待人工选择",
        )
        _set_job(
            run_id,
            status="awaiting_selection",
            progress=65,
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
            },
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        _set_node(run_id, current_node, "failed", str(exc))
        _set_job(run_id, status="failed", progress=100, error=str(exc), artifacts=[])


def _execute_fill(run_id: str, selected_fields: dict[str, Any]) -> None:
    _load_local_env()
    run_dir = RUNS_ROOT / run_id
    current_node = "fill_word"
    try:
        _set_job(run_id, status="running", progress=72, message="节点 3：按选择填充 Word")
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
        llm_adapter, qichacha_adapter, _http_client = _build_external_adapters(
            False, bool(context.get("use_qichacha"))
        )
        ocr_workbook = run_dir / "OCR结构化结果.xlsx"
        result = run_pipeline(
            project_config=PROJECT_CONFIG,
            pdf_path=pdf_path,
            output_dir=run_dir,
            ocr_adapter=None,
            ocr_workbook_path=ocr_workbook if ocr_workbook.is_file() else None,
            llm_adapter=None,
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
        )
        _set_node(run_id, "fill_word", "completed", "Word 填充完成")
        current_node = "output"
        _set_node(run_id, "output", "running", "正在生成评估报告 Word")
        _set_node(run_id, "output", "completed", "全部输出已生成")
        _set_job(
            run_id,
            status="completed",
            progress=100,
            message="评估报告 Word 已生成",
            artifacts=_artifact_list(run_dir),
            issues=[],
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        _set_node(run_id, current_node, "failed", str(exc))
        _set_job(run_id, status="failed", progress=100, error=str(exc), artifacts=[])


@app.post("/api/v1/asset-appraisal/runs", status_code=202)
async def create_run(
    background_tasks: BackgroundTasks,
    materials: list[UploadFile] | None = File(None),
    pdf: UploadFile | None = File(None),
    income_workbook: UploadFile | None = File(None),
    reporting_workbook: UploadFile | None = File(None),
    inputs: str = Form("{}"),
    use_glm: bool = Form(True),
    use_qichacha: bool = Form(True),
    reuse_ocr: bool = Form(True),
):
    # Named slots from the current UI are authoritative even when the user
    # gives the workbook an arbitrary filename.  The legacy ``materials``
    # multi-file field is only used to fill roles that were not supplied by a
    # typed slot.
    role_uploads: dict[str, UploadFile | None] = {
        "pdf": pdf,
        "reporting_workbook": reporting_workbook,
        "income_workbook": income_workbook,
        "reference_report": None,
    }
    workbook_candidates: list[UploadFile] = []
    for upload in list(materials or []):
        suffix = Path(upload.filename or "").suffix.lower()
        if suffix == ".pdf" and role_uploads["pdf"] is None:
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
        "pdf": (role_uploads["pdf"], (".pdf",), "审计报告 PDF"),
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
        parsed_inputs["report_serial"] = validate_report_serial_input(parsed_inputs.get("report_serial"))
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
    has_upload = any(upload is not None for upload, _, _ in uploads.values())
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
    first_filename = next(
        (
            upload.filename
            for upload, _, _ in uploads.values()
            if upload is not None and upload.filename
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
    for field_name, (upload, _, _) in uploads.items():
        if upload is None:
            continue
        suffix = Path(upload.filename or "").suffix.lower()
        stored_name = (
            "source.pdf"
            if field_name == "pdf"
            else f"{field_name}{suffix}"
        )
        stored_files[field_name] = input_dir / stored_name
        with stored_files[field_name].open("wb") as target:
            shutil.copyfileobj(upload.file, target)
    pdf_path = stored_files.get("pdf")
    source_overrides = {
        "audit_pdf": pdf_path,
        "income_workbook": stored_files.get("income_workbook"),
        "reporting_workbook": stored_files.get("reporting_workbook"),
        "reference_report": stored_files.get("reference_report"),
    }
    _set_job(
        run_id,
        run_id=run_id,
        status="queued",
        progress=0,
        message="任务已创建",
        artifacts=[],
        nodes=_initial_node_states(),
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
    _set_job(run_id, status="queued", progress=70, message="已确认候选内容，开始填充 Word")
    background_tasks.add_task(_execute_fill, run_id, selected)
    return JOBS[run_id]


@app.post("/api/v1/asset-appraisal/ocr-cache/check")
async def check_ocr_cache(pdf: UploadFile = File(...)):
    if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=422, detail="请上传 PDF 审计报告")
    import tempfile

    content = await pdf.read()
    with tempfile.NamedTemporaryFile(suffix=".pdf") as temporary:
        temporary.write(content)
        temporary.flush()
        cache = _find_ocr_cache(Path(temporary.name))
    return {
        "hit": cache is not None,
        "source": cache.parent.name if cache else "",
        "message": "命中已有 OCR 结果" if cache else "未命中 OCR 缓存，将执行 OCR",
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
