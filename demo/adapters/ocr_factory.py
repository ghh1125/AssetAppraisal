from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable, Mapping

from .aliyun_docmind_ocr import AliyunDocMindOcrAdapter, _plain


DEFAULT_DOCMIND_ENDPOINT = "docmind-api.cn-hangzhou.aliyuncs.com"


def response_data(value: Any) -> dict[str, Any]:
    plain = _plain(value)
    for body_name in ("body", "Body"):
        body = plain.get(body_name) if isinstance(plain, dict) else None
        if isinstance(body, dict):
            plain = body
            break
    code = plain.get("Code") or plain.get("code") if isinstance(plain, dict) else None
    if code not in (None, "", 200, "200"):
        message = str(
            plain.get("Message") or plain.get("message") or "阿里云请求失败"
        )
        raise RuntimeError(f"{code}: {message}")
    for data_name in ("data", "Data"):
        data = plain.get(data_name) if isinstance(plain, dict) else None
        if isinstance(data, dict):
            return data
        if isinstance(data, str):
            try:
                parsed = json.loads(data)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict):
                return parsed
    return plain if isinstance(plain, dict) else {}


class AliyunDocMindSdkClient:
    def __init__(
        self,
        *,
        access_key_id: str,
        access_key_secret: str,
        endpoint: str = DEFAULT_DOCMIND_ENDPOINT,
    ):
        from alibabacloud_docmind_api20220711 import models
        from alibabacloud_docmind_api20220711.client import Client
        from alibabacloud_tea_openapi import models as open_api_models
        from alibabacloud_tea_util import models as util_models

        config = open_api_models.Config(
            access_key_id=access_key_id,
            access_key_secret=access_key_secret,
        )
        config.endpoint = endpoint
        self.client = Client(config)
        self.models = models
        self.runtime = util_models.RuntimeOptions()

    def submit(self, pdf_path: Path, *, vlm: bool) -> str:
        with pdf_path.open("rb") as file_object:
            request = self.models.SubmitDocParserJobAdvanceRequest(
                file_url_object=file_object,
                file_name=pdf_path.name,
                file_name_extension=pdf_path.suffix.lstrip(".") or "pdf",
                llm_enhancement=vlm,
                enhancement_mode="VLM" if vlm else None,
                output_format=["visualLayoutInfo"],
            )
            response = self.client.submit_doc_parser_job_advance(
                request,
                self.runtime,
            )
        data = response_data(response)
        task_id = data.get("id") or data.get("Id")
        if not task_id:
            raise RuntimeError("阿里云文档解析提交成功但未返回任务 ID")
        return str(task_id)

    def status(self, task_id: str) -> dict[str, Any]:
        request = self.models.QueryDocParserStatusRequest(id=task_id)
        response = self.client.query_doc_parser_status(request)
        return response_data(response)

    def result(
        self,
        task_id: str,
        *,
        layout_num: int,
        layout_step_size: int,
    ) -> dict[str, Any]:
        request = self.models.GetDocParserResultRequest(
            id=task_id,
            layout_num=layout_num,
            layout_step_size=layout_step_size,
        )
        response = self.client.get_doc_parser_result(request)
        return response_data(response)


class UnavailableOcrAdapter:
    def __init__(self, issue: str):
        self.issue = issue

    def extract(self, _pdf_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
        return [], [self.issue]


def _enabled(value: Any) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def _timeout(value: Any) -> float:
    try:
        return max(1.0, float(value))
    except (TypeError, ValueError):
        return 900.0


def _local_paddle_adapter() -> Any:
    from .paddle_ocr import PaddleStructureOcrAdapter, create_local_pipeline

    return PaddleStructureOcrAdapter(create_local_pipeline())


def create_ocr_adapter(
    env: Mapping[str, str],
    *,
    aliyun_client_factory: Callable[..., Any] = AliyunDocMindSdkClient,
    paddle_factory: Callable[[], Any] = _local_paddle_adapter,
) -> Any:
    provider = str(env.get("APPRAISAL_OCR_PROVIDER") or "aliyun").strip().lower()
    if provider == "none":
        return None
    if provider == "paddle":
        return paddle_factory()
    if provider != "aliyun":
        return UnavailableOcrAdapter(f"未知 OCR 提供方：{provider}")

    access_key_id = str(env.get("ALIBABA_CLOUD_ACCESS_KEY_ID") or "").strip()
    access_key_secret = str(
        env.get("ALIBABA_CLOUD_ACCESS_KEY_SECRET") or ""
    ).strip()
    if not access_key_id or not access_key_secret:
        return UnavailableOcrAdapter(
            "阿里云 OCR 凭证缺失：请配置 AccessKey ID 和 AccessKey Secret"
        )
    try:
        client = aliyun_client_factory(
            access_key_id=access_key_id,
            access_key_secret=access_key_secret,
            endpoint=str(
                env.get("APPRAISAL_OCR_ENDPOINT") or DEFAULT_DOCMIND_ENDPOINT
            ),
        )
    except Exception as exc:
        message = str(exc).replace(access_key_id, "***").replace(
            access_key_secret, "***"
        )
        return UnavailableOcrAdapter(f"阿里云 OCR 客户端初始化失败：{message}")
    return AliyunDocMindOcrAdapter(
        client,
        vlm=_enabled(env.get("APPRAISAL_OCR_VLM")),
        timeout_seconds=_timeout(env.get("APPRAISAL_OCR_TIMEOUT_SECONDS")),
        redact_values=(access_key_id, access_key_secret),
    )
