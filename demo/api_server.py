"""Local HTTP bridge for the Vue asset-appraisal workbench.

The business pipeline remains in ``demo.pipeline``. This module only handles
uploads, job state and artifact downloads; c2m can replace it with its own
authenticated task service without changing the domain workflow.
"""

from __future__ import annotations

import json
import hashlib
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from .run import _load_local_env
from .domain.field_validation import validate_valuation_subject_type

ROOT = Path(__file__).resolve().parents[1]
PROJECT_CONFIG = ROOT / "demo/projects/tongfu.yaml"
RUNS_ROOT = ROOT / "runs/web"
OCR_CACHE_ROOT = ROOT / "runs"
JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = Lock()

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
    template = Path(config["template"])
    resolved = template if template.is_absolute() else (PROJECT_CONFIG.parent / template).resolve()
    if not resolved.is_file():
        raise RuntimeError(f"后端默认模板不存在：{resolved}")
    return resolved


def _set_job(job_key: str, **values: Any) -> None:
    with JOBS_LOCK:
        JOBS.setdefault(job_key, {}).update(values)


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
    labels = {
        "资产评估报告_待复核.docx": "评估报告 Word",
        "OCR结构化结果.xlsx": "OCR 结构化 Excel",
        "字段审计清单.xlsx": "字段审计清单",
        "run_manifest.json": "运行清单",
        "issues.json": "复核事项",
        "normalized_fields.json": "标准字段结果",
    }
    return [
        {"name": name, "label": label}
        for name, label in labels.items()
        if (run_dir / name).exists()
    ]


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


def _execute_run(
    run_id: str,
    pdf_path: Path,
    source_overrides: dict[str, Path],
    inputs: dict[str, Any],
    use_glm: bool,
    use_qichacha: bool,
    reuse_ocr: bool,
) -> None:
    _load_local_env()
    run_dir = RUNS_ROOT / run_id
    try:
        _set_job(run_id, status="running", progress=5, message="准备运行环境")
        from .pipeline import run_pipeline
        from .adapters.template_pages import LibreOfficeTemplatePageReader
        template_path = _project_template()

        ocr_cache = _find_ocr_cache(pdf_path) if reuse_ocr else None
        ocr_adapter = None
        if ocr_cache is None:
            from .adapters.paddle_ocr import PaddleStructureOcrAdapter, create_local_pipeline

            ocr_adapter = PaddleStructureOcrAdapter(create_local_pipeline())
        llm_adapter = None
        qichacha_adapter = None
        http_client = None
        if use_glm or use_qichacha:
            import httpx

            http_client = httpx.Client(timeout=120)
        if use_glm:
            from .adapters.bailian_glm import BailianYellowNarrativeAdapter

            key = __import__("os").environ.get("DASHSCOPE_API_KEY", "")
            if not key:
                raise RuntimeError("未配置 DASHSCOPE_API_KEY")
            prompt = (ROOT / "demo/prompts/yellow_narratives.v1.txt").read_text(encoding="utf-8")
            llm_adapter = BailianYellowNarrativeAdapter(
                http_client,
                key,
                prompt,
                base_url=__import__("os").environ.get("APPRAISAL_LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1"),
                model=__import__("os").environ.get("APPRAISAL_LLM_MODEL", "glm-5.2"),
            )
        if use_qichacha:
            from .adapters.company_api import QichachaApiAdapter

            env = __import__("os").environ
            qichacha_adapter = QichachaApiAdapter(
                http_client,
                env.get("QICHACHA_APP_KEY", ""),
                env.get("QICHACHA_SECRET_KEY", ""),
                base_url=env.get("QICHACHA_API_BASE_URL", "https://api.qichacha.com"),
            )
        _set_job(
            run_id,
            progress=15,
            message=("命中已有 OCR 结果，跳过 OCR" if ocr_cache else "开始 PDF OCR 与字段解析"),
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
        )
        _set_job(
            run_id,
            status="completed",
            progress=100,
            message=f"生成完成，复核事项 {len(result.issues)} 条" + ("（已复用 OCR）" if ocr_cache else ""),
            artifacts=_artifact_list(run_dir),
            issues=result.issues,
            completed_at=datetime.now(timezone.utc).isoformat(),
        )
    except Exception as exc:
        _set_job(run_id, status="failed", progress=100, error=str(exc), artifacts=[])


@app.post("/api/v1/asset-appraisal/runs", status_code=202)
async def create_run(
    background_tasks: BackgroundTasks,
    pdf: UploadFile = File(...),
    reference_report: UploadFile = File(...),
    audited_financials: UploadFile = File(...),
    income_workbook: UploadFile = File(...),
    reporting_workbook: UploadFile = File(...),
    inputs: str = Form("{}"),
    use_glm: bool = Form(True),
    use_qichacha: bool = Form(True),
    reuse_ocr: bool = Form(True),
):
    required_files = {
        "pdf": (pdf, ".pdf", "审计报告 PDF"),
        "reference_report": (reference_report, ".docx", "参考评估报告 DOCX"),
        "audited_financials": (audited_financials, ".xlsx", "审计财务 XLSX"),
        "income_workbook": (income_workbook, ".xlsx", "收益法 XLSX"),
        "reporting_workbook": (reporting_workbook, ".xlsx", "上报表 XLSX"),
    }
    for field_name, (upload, suffix, label) in required_files.items():
        if not upload.filename or not upload.filename.lower().endswith(suffix):
            raise HTTPException(status_code=422, detail=f"请上传{label}（字段：{field_name}）")
    try:
        parsed_inputs = json.loads(inputs)
        if not isinstance(parsed_inputs, dict):
            raise ValueError
    except ValueError as exc:
        raise HTTPException(status_code=422, detail="用户输入字段格式错误") from exc
    if parsed_inputs.get("valuation_subject_type") not in (None, ""):
        try:
            parsed_inputs["valuation_subject_type"] = validate_valuation_subject_type(
                parsed_inputs["valuation_subject_type"]
            )
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
    run_id = _run_id_for_pdf(pdf.filename)
    input_dir = RUNS_ROOT / run_id / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    stored_files = {
        "pdf": input_dir / "source.pdf",
        "reference_report": input_dir / "reference_report.docx",
        "audited_financials": input_dir / "audited_financials.xlsx",
        "income_workbook": input_dir / "income_workbook.xlsx",
        "reporting_workbook": input_dir / "reporting_workbook.xlsx",
    }
    for field_name, (upload, _, _) in required_files.items():
        with stored_files[field_name].open("wb") as target:
            shutil.copyfileobj(upload.file, target)
    pdf_path = stored_files["pdf"]
    source_overrides = {
        name: stored_files[name]
        for name in ("reference_report", "audited_financials", "income_workbook", "reporting_workbook")
    }
    template_path = _project_template()
    _set_job(run_id, run_id=run_id, status="queued", progress=0, message="任务已创建", artifacts=[])
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
