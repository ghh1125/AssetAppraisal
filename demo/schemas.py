from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class DemoModel(BaseModel):
    model_config = ConfigDict(extra="forbid")

    @classmethod
    def model_json_schema(cls, *args: Any, **kwargs: Any) -> dict[str, Any]:
        """Expose the business required/optional flag in generated contracts.

        Pydantic already tracks required fields through the object-level
        ``required`` list.  The explicit Chinese flag keeps the handoff
        contract readable to business owners and downstream c2m tooling.
        """
        schema = super().model_json_schema(*args, **kwargs)
        required = set(schema.get("required", []))
        for name, property_schema in schema.get("properties", {}).items():
            property_schema["x-是否必填"] = "是" if name in required else "否"
        return schema


class SourceEvidence(DemoModel):
    source_kind: str = Field(description="来源类别", examples=["income_workbook"])
    source_file: str = Field(description="来源文件或系统名称", examples=["收益法.xlsx"])
    source_locator: str = Field(description="页码、工作表或字段定位", examples=["项目信息!B5"])


class FieldCandidate(DemoModel):
    field_key: str = Field(description="标准业务字段键", examples=["assessed_entity.legal_name"])
    field_name: str = Field(description="字段中文名称", examples=["被评估单位全称"])
    value: Any = Field(description="候选字段值", examples=["示例有限公司"])
    evidence: SourceEvidence = Field(description="候选值来源证据", examples=[{"source_kind": "manual", "source_file": "manual.yaml", "source_locator": "company"}])


class ResolvedField(DemoModel):
    field_key: str = Field(description="标准业务字段键", examples=["assessed_entity.legal_name"])
    field_name: str = Field(description="字段中文名称", examples=["被评估单位全称"])
    value: Any = Field(description="最终采用的字段值", examples=["示例有限公司"])
    evidence: SourceEvidence | None = Field(default=None, description="最终值来源证据", examples=[None])
    candidates: list[FieldCandidate] = Field(default_factory=list, description="全部候选字段值", examples=[[]])


class WordLocation(DemoModel):
    location_id: str = Field(description="Word 稳定位置编号", examples=["DOCUMENT-P0001-X01"])
    record_type: str = Field(description="占位符或黄色说明类型", examples=["占位符"])
    context: str = Field(description="Word 原文上下文", examples=["XXX有限公司"])
    marker: str = Field(description="原占位符或黄色标记", examples=["XXX"])


class LocationMapping(DemoModel):
    location_id: str = Field(description="Word 稳定位置编号", examples=["DOCUMENT-P0001-X01"])
    field_key: str = Field(description="位置对应标准字段键", examples=["assessed_entity.legal_name"])
    field_name: str = Field(description="位置对应字段中文名称", examples=["被评估单位全称"])
    record_type: str = Field(description="占位符或黄色说明类型", examples=["占位符"])
    source_priority: list[str] = Field(default_factory=list, description="固定来源优先级", examples=[["income_workbook", "manual"]])


class InventoryInput(DemoModel):
    template_path: str = Field(description="只读 Word 模板路径", examples=["template.docx"])


class InventoryOutput(DemoModel):
    locations: list[WordLocation] = Field(description="识别到的全部 Word 位置", examples=[[]])


class ExtractSourcesInput(DemoModel):
    sources: dict[str, str] = Field(description="来源名称与文件路径", examples=[{"income_workbook": "income.xlsx"}])


class ExtractSourcesOutput(DemoModel):
    candidates: list[FieldCandidate] = Field(description="从全部来源提取的候选值", examples=[[]])
    issues: list[str] = Field(default_factory=list, description="提取过程中需人工关注的问题", examples=[[]])


class OcrBlock(DemoModel):
    block_id: str = Field(description="页内文本块唯一编号", examples=["p3-b1"])
    block_type: str = Field(description="文本块类型", examples=["text"])
    text: str = Field(description="OCR 识别文本", examples=["资产总计"])
    confidence: float | None = Field(description="OCR 置信度，无法取得时为空", examples=[0.98])
    bbox: list[float] = Field(description="文本块坐标框，顺序为左上右下", examples=[[1, 2, 3, 4]])


class OcrCell(DemoModel):
    row: int = Field(description="表格单元格行号，从一开始", examples=[1])
    column: int = Field(description="表格单元格列号，从一开始", examples=[2])
    text: str = Field(description="单元格 OCR 识别文本", examples=["1,234.50"])
    confidence: float | None = Field(description="单元格 OCR 置信度，无法取得时为空", examples=[0.97])
    bbox: list[float] = Field(description="单元格坐标框，顺序为左上右下", examples=[[5, 6, 7, 8]])


class OcrTable(DemoModel):
    table_id: str = Field(description="页内表格唯一编号", examples=["p3-t1"])
    cells: list[OcrCell] = Field(description="表格单元格列表", examples=[[]])


class OcrPage(DemoModel):
    page_number: int = Field(description="PDF 页码，从一开始", examples=[3])
    page_count: int = Field(description="PDF 总页数", examples=[8])
    blocks: list[OcrBlock] = Field(description="本页文本块", examples=[[]])
    tables: list[OcrTable] = Field(description="本页表格", examples=[[]])


class OcrDocument(DemoModel):
    source_file: str = Field(description="被 OCR 的 PDF 文件名", examples=["审计报告.pdf"])
    pages: list[OcrPage] = Field(description="归一化后的 PDF 页面", examples=[[]])
    issues: list[str] = Field(description="OCR 识别过程问题", examples=[[]])


class OcrPdfInput(DemoModel):
    pdf_path: str = Field(description="待 OCR 的 PDF 文件路径", examples=["审计报告.pdf"])


class OcrPdfOutput(DemoModel):
    document: OcrDocument = Field(
        description="不包含 OCR SDK 对象的结构化文档",
        examples=[{"source_file": "审计报告.pdf", "pages": [], "issues": []}],
    )


class ExportOcrWorkbookInput(DemoModel):
    normalized_ocr: dict[str, list[dict[str, Any]]] = Field(
        description="待写入 Excel 的归一化 OCR 数据", examples=[{"text_blocks": []}]
    )
    output_path: str = Field(description="独立 OCR 结构化 Excel 输出路径", examples=["OCR结构化结果.xlsx"])


class ExportOcrWorkbookOutput(DemoModel):
    workbook_path: str = Field(description="已生成的 OCR 结构化 Excel 路径", examples=["OCR结构化结果.xlsx"])


class ResolveFieldsInput(DemoModel):
    candidates: list[FieldCandidate] = Field(description="待选择的候选字段值", examples=[[]])
    mappings: list[LocationMapping] = Field(description="字段映射规则", examples=[[]])


class ResolveFieldsOutput(DemoModel):
    fields: list[ResolvedField] = Field(description="解析后的标准字段", examples=[[]])


class GenerateNarrativeInput(DemoModel):
    fields: list[ResolvedField] = Field(description="作为生成证据的标准字段", examples=[[]])


class GenerateNarrativeOutput(DemoModel):
    fields: list[ResolvedField] = Field(description="生成的叙述性标准字段", examples=[[]])
    prompt_version: str = Field(description="所用 Prompt 版本", examples=["yellow_narratives.v1"])


class FillWordInput(DemoModel):
    template_path: str = Field(description="只读 Word 模板路径", examples=["template.docx"])
    output_path: str = Field(description="新 Word 输出路径", examples=["report.docx"])
    fields: list[ResolvedField] = Field(description="用于填充的标准字段", examples=[[]])


class FillWordOutput(DemoModel):
    report_path: str = Field(description="生成的新 Word 路径", examples=["report.docx"])
    replacement_count: int = Field(description="已替换位置数量", examples=[147])


class ReviewInput(DemoModel):
    review_type: str = Field(description="审核类型", examples=["data_validation"])
    report_path: str = Field(description="待审核 Word 路径", examples=["report.docx"])
    evidence: dict[str, Any] = Field(description="供模型审核的结构化证据", examples=[{"fields": {}}])


class ReviewOutput(DemoModel):
    review_type: str = Field(description="审核类型", examples=["semantic_review"])
    status: str = Field(description="审核状态", examples=["completed_with_issues"])
    summary: str = Field(description="审核摘要", examples=["发现一项问题"])
    findings: list[dict[str, Any]] = Field(description="审核发现的问题列表", examples=[[{"severity": "medium"}]])
    model: str = Field(description="实际使用的模型", examples=["qwen3.7-flash"])
    prompt_version: str = Field(description="审核 Prompt 版本", examples=["review_data.v1"])


class ReviewAggregateInput(DemoModel):
    reviews: dict[str, ReviewOutput] = Field(description="三个审核节点的结果", examples=[{}])


class ReviewAggregateOutput(DemoModel):
    status: str = Field(description="汇总审核状态", examples=["completed_with_issues"])
    review_count: int = Field(description="已执行审核数量", examples=[3])
    finding_count: int = Field(description="问题总数", examples=[2])
    severity_counts: dict[str, int] = Field(description="按严重级别统计的问题数量", examples=[{"high": 1, "medium": 1, "low": 0}])
    failed_reviews: list[str] = Field(description="执行失败的审核节点", examples=[[]])
    findings: list[dict[str, Any]] = Field(description="汇总后的问题列表", examples=[[{"severity": "high"}]])


class ExportAuditInput(DemoModel):
    report_path: str = Field(description="待审核 Word 路径", examples=["report.docx"])
    fields: list[ResolvedField] = Field(description="需导出的标准字段", examples=[[]])


class ExportAuditOutput(DemoModel):
    audit_path: str = Field(description="字段审计清单路径", examples=["audit.xlsx"])
    manifest_path: str = Field(description="运行记录路径", examples=["run_manifest.json"])


class WorkflowInput(DemoModel):
    project_config: str = Field(description="项目配置路径", examples=["demo/projects/tongfu.yaml"])


class WorkflowOutput(DemoModel):
    report_path: str = Field(description="生成的新 Word 路径", examples=["report.docx"])
    audit_path: str = Field(description="字段审计清单路径", examples=["audit.xlsx"])
    issues: list[str] = Field(description="人工审核问题", examples=[[]])
