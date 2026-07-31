from __future__ import annotations

from typing import Any, Literal

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


class WorkflowNodeDefinition(DemoModel):
    name: str = Field(description="工作流节点名称", examples=["ocr_pdf"])
    description: str = Field(
        default="",
        description="节点职责和边界说明",
        examples=["读取输入材料并校验文件角色"],
    )
    input_model: str = Field(description="节点输入模型名称", examples=["OcrPdfInput"])
    output_model: str = Field(description="节点输出模型名称", examples=["OcrPdfOutput"])
    depends_on: list[str] = Field(description="前置节点名称", examples=[["inventory"]])
    human_checkpoint: str | None = Field(
        default=None,
        description="节点完成后的人工确认要求",
        examples=["人工下载并复核Word"],
    )


class WorkflowDefinition(DemoModel):
    version: str = Field(description="工作流业务版本", examples=["1.1.0"])
    contract_version: str = Field(
        description="节点契约规则版本",
        examples=["workflow_contract.v1"],
    )
    nodes: list[WorkflowNodeDefinition] = Field(
        description="有序工作流节点",
        examples=[[]],
    )


class NodeTrace(DemoModel):
    node_name: str = Field(description="实际执行的节点名称", examples=["fill_word"])
    status: Literal[
        "completed",
        "completed_with_issues",
        "skipped",
        "failed",
    ] = Field(description="节点执行状态", examples=["completed"])
    started_at: str = Field(
        description="节点开始时间",
        examples=["2026-07-26T12:00:00+00:00"],
    )
    finished_at: str = Field(
        description="节点结束时间",
        examples=["2026-07-26T12:00:01+00:00"],
    )
    input_model: str = Field(description="节点输入模型名称", examples=["FillWordInput"])
    output_model: str = Field(description="节点输出模型名称", examples=["FillWordOutput"])
    input_data: dict[str, Any] = Field(
        description="已校验的节点输入摘要",
        examples=[{}],
    )
    output_data: dict[str, Any] = Field(
        description="已校验的节点输出摘要",
        examples=[{}],
    )
    evidence: list[dict[str, Any]] = Field(
        description="节点使用的来源证据",
        examples=[[]],
    )
    issues: list[str] = Field(description="节点问题列表", examples=[[]])
    human_checkpoint: str | None = Field(
        default=None,
        description="节点人工确认要求",
        examples=["用户逐项选择LLM候选内容"],
    )


class WorkflowTrace(DemoModel):
    workflow_version: str = Field(description="工作流业务版本", examples=["1.1.0"])
    contract_version: str = Field(
        description="节点契约规则版本",
        examples=["workflow_contract.v1"],
    )
    versions: dict[str, str] = Field(
        description="规则、Prompt、模型和数据版本",
        examples=[{"prompt": "yellow_narratives.v1"}],
    )
    nodes: list[NodeTrace] = Field(description="实际执行节点轨迹", examples=[[]])


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
    comment_source_kind: str | None = Field(
        default=None,
        description="最新版 Word 批注声明的来源类别",
        examples=["qichacha_api"],
    )
    comment_source_instruction: str | None = Field(
        default=None,
        description="最新版 Word 批注原文",
        examples=["数据来源：通过企查查/天眼查等API获取"],
    )
    force_unresolved: bool = Field(
        default=False,
        description="批注是否明确要求保留并高亮原占位符",
        examples=[False],
    )


class LocationMapping(DemoModel):
    location_id: str = Field(description="Word 稳定位置编号", examples=["DOCUMENT-P0001-X01"])
    field_key: str = Field(description="位置对应标准字段键", examples=["assessed_entity.legal_name"])
    field_name: str = Field(description="位置对应字段中文名称", examples=["被评估单位全称"])
    record_type: str = Field(description="占位符或黄色说明类型", examples=["占位符"])
    source_priority: list[str] = Field(default_factory=list, description="固定来源优先级", examples=[["income_workbook", "manual"]])
    comment_source_kind: str | None = Field(default=None, description="Word 批注声明的来源类别", examples=["bailian_glm"])
    comment_source_instruction: str | None = Field(default=None, description="Word 批注原文", examples=["数据来源：大模型"])
    force_unresolved: bool = Field(default=False, description="是否禁止自动填充并保留黄色占位符", examples=[False])


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
    row_span: int = Field(default=1, description="单元格跨越的行数", examples=[1])
    column_span: int = Field(default=1, description="单元格跨越的列数", examples=[2])
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


class FillWordInput(DemoModel):
    template_path: str = Field(description="只读 Word 模板路径", examples=["template.docx"])
    output_path: str = Field(description="新 Word 输出路径", examples=["report.docx"])
    fields: list[ResolvedField] = Field(description="用于填充的标准字段", examples=[[]])


class FillWordOutput(DemoModel):
    report_path: str = Field(description="生成的新 Word 路径", examples=["report.docx"])
    replacement_count: int = Field(description="已替换位置数量", examples=[147])


class ManualBasicInputs(DemoModel):
    commissioning_party_name: str | None = Field(default=None, description="委托方公司名称", examples=["示例委托有限公司"])
    commissioning_party_short_name: str | None = Field(default=None, description="委托方公司简称", examples=["示例委托"])
    transaction_type: Literal["转让", "收购", "增资", "减资"] | None = Field(default=None, description="委托类型，单选", examples=["收购"])
    target_company_name: str | None = Field(default=None, description="评估主体（被评估公司）名称", examples=["示例被评估有限公司"])
    target_company_short_name: str | None = Field(default=None, description="评估主体公司简称", examples=["示例主体"])
    valuation_subject_type: Literal["股东全部权益价值", "股东部分权益价值", "企业整体价值", "资产组价值"] | None = Field(default=None, description="评估对象，单选", examples=["股东全部权益价值"])
    selected_valuation_method: list[Literal["资产基础法", "收益法", "市场法"]] | str | None = Field(default=None, description="评估方法，多选且至少一个", examples=[["收益法", "资产基础法"]])
    final_valuation_method: Literal["资产基础法", "收益法", "市场法"] | None = Field(default=None, description="评估结论采用方法，单选", examples=["收益法"])
    report_serial: str | int | None = Field(default=None, description="评估报告编号流水号，非负整数", examples=[1])


class StartInput(DemoModel):
    manual_inputs: ManualBasicInputs = Field(
        default_factory=ManualBasicInputs,
        description="最新图片定义的九项必填人工基础信息；其余模板字段由材料、API、系统时间或占位符规则处理",
        examples=[{"commissioning_party_name": "示例委托有限公司", "commissioning_party_short_name": "示例委托", "target_company_name": "示例被评估有限公司", "target_company_short_name": "示例主体", "transaction_type": "收购", "valuation_subject_type": "股东全部权益价值", "selected_valuation_method": ["收益法"], "final_valuation_method": "收益法", "report_serial": 1}],
    )
    materials: dict[str, str] = Field(
        default_factory=dict,
        description="上传材料角色与保存路径，PDF和Excel均为可选",
        examples=[{"pdf": "审计报告.pdf", "reporting_workbook": "资产清查.xlsx"}],
    )
    template_path: str = Field(
        description="后台提供的只读 Word 模板路径",
        examples=["评估报告版式-沟通标注版.docx"],
    )


class StartOutput(DemoModel):
    accepted: bool = Field(description="输入材料和人工信息是否通过基本校验", examples=[True])
    material_roles: list[str] = Field(description="已接收的材料角色", examples=[["pdf", "reporting_workbook"]])
    template_path: str = Field(description="实际使用的后台模板路径", examples=["template.docx"])
    issues: list[str] = Field(default_factory=list, description="输入节点问题列表", examples=[[]])


class NarrativeCandidate(DemoModel):
    field_key: str = Field(description="Word固定LLM位置对应的字段键", examples=["industry_overview"])
    field_name: str = Field(description="Word固定LLM位置的中文名称", examples=["行业情况"])
    value: str = Field(description="LLM生成的候选文本", examples=["行业候选内容"])
    location_ids: list[str] = Field(description="该候选文本对应的Word位置编号", examples=[["DOCUMENT-P0003-X01"]])
    selected: bool = Field(description="用户是否选择写入Word", examples=[False])


class CandidateGenerationInput(DemoModel):
    materials: dict[str, str] = Field(description="供OCR和LLM读取的材料角色与路径", examples=[{"pdf": "审计报告.pdf"}])
    pdf_present: bool = Field(description="是否存在待OCR的PDF", examples=[True])
    llm_enabled: bool = Field(description="是否调用百炼模型生成候选文本", examples=[True])


class CandidateGenerationOutput(DemoModel):
    ocr_performed: bool = Field(description="本次是否执行PDF OCR", examples=[False])
    ocr_workbook_path: str | None = Field(default=None, description="OCR结构化结果Excel路径", examples=["OCR结构化结果.xlsx"])
    candidates: list[NarrativeCandidate] = Field(description="全部Word固定LLM位置的候选内容", examples=[[]])
    selection_required: bool = Field(description="是否暂停等待用户选择候选内容", examples=[True])


class WordFillInput(DemoModel):
    template_path: str = Field(description="只读Word模板路径", examples=["template.docx"])
    selected_llm_fields: dict[str, str] = Field(description="用户选择写入的LLM候选字段和值", examples=[{"industry_overview": "行业候选内容"}])
    deterministic_sources: dict[str, str] = Field(description="PDF OCR、Excel、企查查和人工字段来源摘要", examples=[{"asset_total": "审计报告.pdf"}])


class WordFillOutput(DemoModel):
    report_path: str = Field(description="生成的新Word路径", examples=["资产评估报告_待复核.docx"])
    replacement_count: int = Field(description="已替换的位置和表格数量", examples=[147])
    unresolved_count: int = Field(description="仍保留并高亮XXX的位置数量", examples=[3])


class OutputInput(DemoModel):
    report_path: str = Field(description="待交付的Word路径", examples=["资产评估报告_待复核.docx"])


class OutputOutput(DemoModel):
    report_path: str = Field(description="唯一可供下载的评估报告Word路径", examples=["资产评估报告_待复核.docx"])


class WorkflowInput(DemoModel):
    project_config: str = Field(description="项目配置路径", examples=["demo/projects/tongfu.yaml"])


class WorkflowOutput(DemoModel):
    report_path: str = Field(description="生成的新 Word 路径", examples=["report.docx"])
