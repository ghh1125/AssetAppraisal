# 资产评估报告自动生成工作流设计

## 1. 目标

构建一个不依赖 Codex 的本地 Python 工作流。工作流读取 Word 模板、审计报告 PDF、收益法及上报表 Excel、人工配置、企业数据 API 和大模型接口，生成一份新的资产评估报告 Word，同时输出供人工审核的来源清单与运行记录。

原 Word 文件始终作为只读模板使用。工作流不得覆盖、重命名或修改原模板，每次运行必须在独立目录中创建新的 Word 文件。

首个可运行项目为“通富昆山”，但核心代码不得写死公司名称、文件名、模板路径或服务商。后续项目通过更换项目配置、模板映射和输入文件复用同一工作流。

## 2. 已确认的产品边界

- 采用端到端方案：PDF OCR、Excel 取数、人工参数、企业数据 API、LLM 文本生成和 Word 填充全部纳入工作流。
- 首版提供命令行入口，不开发网页界面。
- 采用通用可配置的服务适配器，不绑定单一 OCR、LLM 或企业数据服务商。
- 模板中的全部 `XX`、`XXX`、`XXXX` 占位符都必须被处理。
- 黄色标注说明也是待填位置，必须替换为实际段落、列表、表格或明确的待人工补充文本，不能原样保留。
- 无论是否存在缺失字段，工作流都应尽可能生成新的 Word。
- 工作流以“机器生成、人工审核”为原则，不判断 OCR 或业务内容是否真实准确。

## 3. 总体架构

工作流采用“业务内核可复用、生产外壳由 c2m 实现”的模式。Demo 负责发现、验证和表达业务规则，不作为 c2m 的运行时依赖。首版不引入用户、权限、数据库、后台任务、监控或发布能力。

数据流如下：

```text
项目配置 YAML
    ↓
输入文件登记与哈希记录
    ↓
PDF OCR／Excel／人工参数／企业数据 API
    ↓
标准字段库
    ↓
固定优先级选值与派生计算
    ↓
LLM 生成叙述性内容
    ↓
Word 模板定位与替换
    ↓
新报告.docx + 字段审计清单.xlsx + 运行记录.json
```

核心组件：

1. `workflow.yaml`：声明节点、顺序、输入输出模型和人工检查点。
2. `schemas.py`：为每个节点定义带中文说明、必填性和示例的 Pydantic 契约。
3. `domain/`：只包含接收普通 Python/Pydantic 数据并返回 JSON 可序列化结果的业务规则。
4. `Provider`：由编排层注入 OCR、LLM、企业数据、文件读取能力，业务函数不读取 `.env` 或创建全局客户端。
5. `Template Adapter`：识别并替换 Word 中的占位符和黄色说明。
6. `Audit Exporter`：输出机器填充值及来源，供评估师人工审核。

## 4. 项目结构

计划使用以下目录：

```text
AssetAppraisal/
├── pyproject.toml
└── demo/
    ├── README.md
    ├── workflow.yaml
    ├── schemas.py
    ├── run.py
    ├── adapters/
    │   ├── files.py
    │   ├── excel.py
    │   ├── ocr.py
    │   ├── company_api.py
    │   └── llm.py
    ├── domain/
    │   ├── registry.py
    │   ├── calculations.py
    │   ├── mapping.py
    │   └── replacement.py
    ├── prompts/
    │   ├── company_narrative.v1.txt
    │   └── narrative_output.v1.json
    ├── mappings/
    │   └── appraisal_report_v1.yaml
    ├── projects/
    │   └── tongfu.yaml
    ├── fixtures/
    ├── expected/
    ├── tests/
    ├── data_manifest.yaml
    └── CHANGELOG.md
```

模块边界以公开契约为准。`domain/` 不得引用 Streamlit、FastAPI、c2m 数据库模型、页面 session、绝对路径或环境变量。服务适配器只能返回结构化结果，不能直接写 Word；Word 替换规则只能接收标准字段和模板结构，不能直接访问 Excel、PDF 或网络接口。

## 5. 配置设计

每个项目使用一个 YAML 文件，例如：

```yaml
project:
  id: tongfu_20250630
  template: 资产评估工作流/评估报告版式-沟通标注版.docx
  output_root: runs

sources:
  audit_pdf: 资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf
  income_workbook: 资产评估工作流/通富热处理（昆山）有限公司-收益法-20250630.xlsx
  reporting_workbook: 资产评估工作流/上报表文件_通富昆山_已处理.xlsx
  manual_inputs: projects/tongfu.manual.yaml

providers:
  ocr:
    type: paddleocr
    language: ch
  llm:
    type: openai_compatible
    base_url_env: APPRAISAL_LLM_BASE_URL
    api_key_env: APPRAISAL_LLM_API_KEY
    model_env: APPRAISAL_LLM_MODEL
  company_data:
    type: generic_company_api
    base_url_env: APPRAISAL_COMPANY_API_BASE_URL
    api_key_env: APPRAISAL_COMPANY_API_KEY

template_mapping: mappings/appraisal_report_v1.yaml
```

服务密钥只允许通过 `.env` 或进程环境变量提供。项目配置、运行日志、审计文件和 Word 输出不得保存密钥。

## 6. 命令行入口

首版提供一条主要命令：

```bash
python -m demo.run demo/projects/tongfu.yaml
```

Demo 首版只提供项目配置路径和可选输出目录，避免把生产调度能力放入业务仓库：

```text
--output-dir PATH      指定本次输出目录
--offline              不调用外部 API 和 LLM，缺值写入待人工补充文本
```

命令返回非零状态只用于无法继续的程序错误，例如模板损坏、配置无法解析或输出 Word 无法创建。字段缺失、API 无返回和 OCR 空结果不得阻止 Word 生成。

## 7. 标准字段库

所有来源数据必须先转换为统一字段记录，再用于填充：

```json
{
  "field_key": "assessed_entity.legal_name",
  "value": "通富热处理（昆山）有限公司",
  "value_type": "string",
  "source_kind": "income_workbook",
  "source_file": "通富热处理（昆山）有限公司-收益法-20250630.xlsx",
  "source_locator": "项目信息!B5",
  "transform": "trim_company_name",
  "generated_at": "2026-07-22T10:30:00+08:00"
}
```

同一字段可以拥有多个候选值。配置为每个字段规定固定的来源优先级，工作流直接选取最高优先级的非空值，不因来源冲突停止运行。所有候选值仍写入审计记录，供人工比较。

派生字段使用代码函数计算，例如日期拆分、元与万元换算、评估增值额、增值率、报告有效期和金额大写。派生公式及输入字段记录在审计文件中。

## 8. 来源适配器

### 8.1 PDF OCR

默认实现使用本地 PaddleOCR。工作流将 PDF 页面渲染为图像，逐页识别文字和表格，并保存：

- PDF 文件名；
- 原页码；
- OCR 原始文字；
- 表格单元格；
- PaddleOCR 返回的原始置信度；
- 页面图像和结果缓存路径。

置信度只作为审计信息保存，不用于阻止取值。只要 OCR 返回内容，工作流即按映射规则使用。若某页完全无识别结果，则相关字段写入待人工补充文本。

### 8.2 Excel

Excel 适配器同时读取公式和缓存值，保留工作表、单元格、合并区域、公式、显示值和单位。字段映射优先使用明确单元格；无法固定单元格时使用标签匹配规则。

### 8.3 企业数据 API

企业数据使用通用 HTTP 适配器。配置决定请求地址、认证头、查询参数和返回字段映射。企查查、天眼查或其他服务商通过新增配置或适配器接入。

API 无返回或未配置密钥时，工作流不终止，相关字段写入 `【待人工补充：字段名称】`。

### 8.4 LLM

LLM 使用 OpenAI 兼容接口，不依赖 Codex。输入为已收集的企业信息、OCR 文字、人工参数和固定提示模板；输出必须符合定义的 JSON 结构。

LLM 仅负责行业介绍、业务概述、主要产品、客户供应商概述、盈利模式、SWOT 和可比公司分析等叙述性内容。财务数字只能来自 Excel、OCR、API、人工参数或确定性计算函数。

## 9. Word 模板定位与填充

模板中的位置使用稳定编号，例如：

```text
DOCUMENT-P0074-X02
DOCUMENT-P0088-H01
FOOTER4-P0001-X01
```

位置编号由 Word 部件、段落序号和段内 occurrence 组成。映射示例：

```yaml
- location_id: DOCUMENT-P0074-X02
  field: assessed_entity.legal_name
  replacement_type: inline

- location_id: DOCUMENT-P0088-H01
  field: valuation.selected_methods
  replacement_type: paragraph
```

填充要求：

1. 根据上下文和稳定位置编号定位，不做全局无差别字符串替换。
2. 在原有 run 中替换占位符，尽量保留字体、字号、颜色和段落格式。
3. 黄色说明按配置替换为内联文本、段落、列表或表格。
4. 同一标准字段在封面、摘要、正文和页脚中统一复用。
5. 字段无值时写入 `【待人工补充：字段名称】`。
6. 所有原黄色说明必须删除或替换。
7. 原模板只读，新 Word 写入本次运行目录。

## 10. 运行输出

每次运行建立独立目录：

```text
runs/tongfu_20250630/20260722_103000/
├── 资产评估报告_待复核.docx
├── 字段审计清单.xlsx
├── run_manifest.json
├── normalized_fields.json
├── issues.json
├── logs/
└── ocr/
```

`字段审计清单.xlsx` 至少包括：

- Word 原页码及位置编号；
- 原文上下文；
- 标准字段名；
- 最终填充值；
- 来源类型；
- 来源文件；
- PDF 页码或 Excel 工作表/单元格；
- 转换或计算规则；
- 其他候选值；
- 是否为待人工补充。

`run_manifest.json` 记录程序版本、运行编号、时间、项目配置、模板哈希、输入文件哈希、服务适配器类型、各任务状态和输出文件列表。

## 11. 异常与降级策略

本工作流不承担专业可靠性判断。异常处理只保证程序能够尽可能产出可人工审核的 Word。

### 终止运行

仅在以下情况终止：

- 项目配置无法解析；
- Word 模板不存在或结构损坏；
- 输出目录不可写；
- Word 输出文件无法创建。

### 继续运行

以下情况记录后继续：

- OCR 无结果；
- OCR 置信度低；
- Excel 与 PDF 数值不一致；
- 企业 API 未配置、超时或返回空值；
- LLM 接口失败或返回格式不合法；
- 字段未找到可用来源。

继续运行时使用固定来源优先级，或写入 `【待人工补充：字段名称】`。已有黄色说明不作为降级文本保留。

## 12. 缓存与可扩展性

Demo 可把 OCR、API 和 LLM 结果写入本地运行目录以便重复验证，但缓存、幂等、异步重试和持久化由 c2m 生产外壳实现。

新增服务商时实现编排层 Adapter 接口；新增模板时提供新的模板映射；新增数据来源时增加 Adapter 并向标准字段库写入候选字段。`domain/` 不包含具体公司、服务商、路径或客户端逻辑。

## 13. 测试策略

实现遵循测试先行。

### 单元测试

- 项目配置解析；
- 任务依赖与缓存键；
- 标准字段候选值及固定优先级选择；
- 日期、金额、单位和派生计算；
- Word 段落内多占位符精确替换；
- 黄色说明替换；
- 缺值降级文本；
- 原模板不被修改。

### 适配器契约测试

- OCR Provider 返回统一页面结果；
- LLM Provider 返回结构化叙述内容；
- 企业数据 Provider 返回统一企业字段；
- 服务失败时工作流产生待人工补充字段而非终止。

### 模板回归测试

当前模板必须识别：

- 127 个 `XX`、`XXX`、`XXXX` 占位符；
- 20 个黄色说明块；
- 2 个页脚占位符；
- 正文和表格中的所有稳定位置编号。

### 端到端测试

使用通富项目配置运行工作流，验证：

- 创建独立运行目录；
- 原模板哈希不变；
- 生成新的 Word、字段审计清单和运行记录；
- 输出 Word 不保留原 `XXX` 或黄色说明；
- 无值位置使用明确的待人工补充文本；
- Word 文件结构有效并可正常打开。

OCR 和业务内容的真实准确性不属于自动测试目标，由用户人工审核。

## 14. 首版完成标准

首版完成必须满足：

1. 可通过一条命令运行通富项目。
2. 不依赖 Codex。
3. 原 Word 模板保持不变。
4. 每次生成独立的新 Word。
5. 127 个占位符和 20 处黄色说明均被处理。
6. PDF、Excel、人工参数、企业 API 和 LLM 均有可配置适配器入口。
7. 外部服务缺失或失败时仍生成待人工审核版 Word。
8. 生成字段审计清单和运行记录。
9. 自动测试覆盖核心替换、降级和模板保护行为。

## 15. 本阶段不包含

- 不开发网页界面。
- 不覆盖或修改原模板。
- 不由程序判断 OCR 内容、财务数字或评估结论是否专业正确。
- 不自动签章、提交或发布正式评估报告。
- 不把任何服务密钥写入项目文件或输出文件。
- 不实现用户、权限、资源归属、上传隔离、数据库持久化、后台任务、并发控制、重试、限流和监控；这些由 c2m 接入。
- 不要求 c2m 直接依赖或运行整个 Demo 仓库；交接资产是 `domain/`、`schemas.py`、Prompt、规则、样例和验收测试。

## 16. Demo 交付规范

每个工作流节点必须在 `workflow.yaml` 中声明输入模型、输出模型、顺序和人工检查点。每个 Prompt 单独存放并带版本号，输出由 Pydantic 模型约束。关键规则至少提供一个正常样例和一个边界样例，整个 Demo 至少提供 10 个脱敏代表性样例及期望结果。

规则、Prompt 或数据版本变化时同步更新 `CHANGELOG.md`、`data_manifest.yaml` 和受影响样例。Demo 验收只证明业务路径可执行，不代表具备生产上线条件。
