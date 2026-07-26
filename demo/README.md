# 资产评估报告自动生成 Demo

## 定位

本目录是资产评估业务规则的可执行 Demo，不是生产系统。它负责表达和验证流程、输入输出契约、字段规则、Prompt、样例以及 Word 填充结果；用户、权限、任务队列、持久化、并发、重试、监控和发布由 c2m 实现。c2m 应移植 `domain/`、`schemas.py`、Prompt、映射和测试，不应把整个 Demo 仓库作为运行时依赖。

## 工作流节点

1. `inventory`：识别 Word 中的稳定位置。
2. `ocr_pdf`：用 PP-StructureV3 将扫描 PDF 识别为页、文本块、表格和单元格。
3. `export_ocr_workbook`：生成独立的 `OCR结构化结果.xlsx`。
4. `extract_sources`：读取 OCR Excel、其他结构化 Excel、节点参数和外部服务结果。
5. `resolve_fields`：按字段语义、期间、单位和固定黄色来源路由选择字段值。
6. `generate_narrative`：通过注入的百炼模型生成限定的叙述内容。
7. `fill_word`：复制模板并替换全部占位符和黄色说明。
8. `llm_format_review`、`llm_data_validation`、`llm_semantic_review`：对生成 Word 执行格式、数据和语义审核。
9. `review_aggregate`：汇总三类审核结果和问题。
10. `export_audit`：导出字段来源清单和运行记录。

人工检查点位于最后一步：评估师审核生成 Word、字段审计清单和三类 LLM 审核问题。

## 安装与运行

```bash
uv sync --extra dev
uv run python -m demo.run demo/projects/tongfu.yaml --offline
```

`--offline` 不调用企业 API 和 LLM；企查查/LLM/人工类字段在无材料时留空，不向 Word 写入“待人工补充”等提示语。项目配置在 `required_financial_fields` 中声明基础财务输入，在 `required_monetary_fields` 中声明最终 Word 绝不允许为空的金额和财务结果；后者会在黄色来源路由完成后再次校验，缺失时终止生成。可用 `--output-dir` 指定独立输出目录。原 Word 永远作为只读模板，不会被覆盖。

在 c2m 或其他宿主中，可直接调用 `run_project(...)` / `run_pipeline(...)` 并通过 `ocr_adapter`、`company_api_adapter`、`llm_adapter` 和 `review_adapters` 参数注入已有服务。Demo 不在 `domain/` 内创建客户端或读取密钥；注入结果只接受映射表中已经登记的字段键。

端到端 OCR 使用独立的 Python 3.11 环境（PaddleOCR 当前可选依赖限定 Python `<3.13`）。项目根目录的本地 `.env` 会在 Demo 入口自动加载；也可以由宿主进程自行注入同名环境变量，例如 `DASHSCOPE_API_KEY`、`QICHACHA_APP_KEY` 和 `QICHACHA_SECRET_KEY`：

```bash
uv python install 3.11
uv sync --python 3.11 --extra dev --extra ocr --extra services
uv run --python 3.11 python -m demo.run demo/projects/tongfu.yaml \
  --pdf '资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf' \
  --template 'templates/评估报告版式-沟通标注版.docx' \
  --output-dir runs/tongfu-ocr --ocr-engine paddle --use-glm --use-qichacha \
  --commissioning-party-name '委托方全称' \
  --commissioning-party-short-name '委托方简称' \
  --report-serial '报告流水号' \
  --valuation-purpose-inputs '评估目的具体说明' \
  --selected-valuation-method '收益法、资产基础法' \
  --valuation-subject-type '股东全部权益价值' \
  --transaction-type '收购' \
  --final-valuation-method '收益法' \
  --target-company-short-name '通富昆山'
```

如需调用已购买的企查查接口 735（工商详情）、231（商标）、514（专利）和 233（著作权软著），设置 `QICHACHA_APP_KEY`、`QICHACHA_SECRET_KEY` 并增加 `--use-qichacha`。默认使用企查查官方签名方式；如平台给 514 或 233 分配了不同的路径，可用 `QICHACHA_ENDPOINT_514`、`QICHACHA_ENDPOINT_233` 覆盖，基地址可用 `QICHACHA_API_BASE_URL` 覆盖。两个节点输入可放入 JSON 文件并用 `--node-inputs-json` 传入。凭证只在 `run.py` 这一组合入口读取，不会进入 `domain/`、运行清单、审计文件或 Word。

## 黄色字段硬路由

模板中的 20 个黄色位置固定为四组，配置漏项、重复或模板位置变化会直接报错：

- 百炼叙述：7 个——公司概况、行业、业务板块、主要产品、客户供应商、盈利模式/SWOT、可比公司。
- 百炼审核：格式、数据、语义 3 个独立任务；默认都使用 `qwen3.7-flash`，可在项目 YAML 的 `llm.tasks` 中分别改模型。
- 企查查 API：5 个——委托方概况、历史股权沿革、基准日股权、账外无形资产、软件著作权。
- PDF OCR/XLSX：6 个——历史资产负债表、历史利润表、税率、评估范围、主要长期资产、资产基础法结果。
- 节点输入：2 个——选用评估方法、评估目的输入。

运行时会再次按白名单过滤每个适配器的返回值。指定来源无值时留空，不使用参考 Word、其他 API、其他 Excel 或 LLM 跨路由补值，也不写入“待人工补充”。现有 `XX/XXX` 非黄色占位符继续按项目映射和确定性规则填写。

## 输出

- `OCR结构化结果.xlsx`
  - `OCR_表格`：逐单元格审计明细；`OCR_表格索引` 及 `表_<页码>_<表格编号>`：按 OCR 行列恢复的矩阵表，便于人工查看和后续映射。
- `资产评估报告_待复核.docx`
- `字段审计清单.xlsx`
- `normalized_fields.json`
- `issues.json`
- `格式审核.json`、`数据校验.json`、`语义审核.json`
- `审核汇总.json`
- `run_manifest.json`

## Vue 前端工作台

`frontend/` 将黄色提示中的人工输入、材料上传、后端固定 Word 模板、GLM/企查查开关和产物下载做成页面。每次任务需要上传审计 PDF、参考评估报告 DOCX、审计财务 XLSX、收益法 XLSX 和上报表 XLSX；Word 模板由后端固定提供，不需要上传。上传 PDF 后会先调用 `/api/v1/asset-appraisal/ocr-cache/check`，按 PDF SHA-256 查找已有 `OCR结构化结果.xlsx`；命中时生成任务复用 OCR，跳过 PaddleOCR。前端调用 `/api/v1/asset-appraisal/runs`；本地可用下面的 HTTP 桥接服务承接现有流水线：

上传框按材料角色接收文件，原文件名不需要与配置一致。审计财务 XLSX 应包含 `06N_资产负债表`、`07N_利润表`；收益法 XLSX 应包含主要产品及服务、所得税表、净现金流计算表；上报表 XLSX 应包含表 1、表 4-6、表 4-12。若工作表名称或表格布局变化，需要在项目 YAML 映射中配置新的定位规则。

Web 任务输出目录按“`YYYYMMDDHHMM-PDF文件名`”命名，例如 `runs/web/202607231144-通富2025.6.30合并及母公司审计报告/`；同一分钟重复提交会自动追加序号，不覆盖已有任务。

```bash
uv sync --extra services --extra ocr --extra web --python 3.11
uv run --python 3.11 uvicorn demo.api_server:app --host 127.0.0.1 --port 8000

cd frontend
npm install
npm run dev
```

页面提交的字段与命令行参数一一对应：委托方全称/简称、被评估单位全称（可选，用于企查查身份核验）/简称、报告编号流水号、评估目的、评估方法、评估对象（四选一）、交易类型和最终采用方法。报告编号可以输入完整编号，程序会提取模板所需的流水号并同步所有报告编号年份。企查查、百炼和 PDF OCR/XLSX 不要求用户在页面重复填写。

字段审计清单逐行记录原模板页码、稳定位置编号、原文上下文、程序填入内容、来源类别、来源文件和来源位置。原模板页码由 LibreOffice 只读渲染模板后按段落顺序匹配得到；页脚标记为“多页页脚”。OCR、API、LLM、财务数据及评估结论的正确性由业务人员人工审核。

`OCR结构化结果.xlsx` 固定包含 `OCR_文本`、`OCR_表格`、`标准财务数据`、`识别问题` 四个工作表，保存页码、行列、坐标、置信度和证据编号。`run_manifest.json` 保存模板/PDF 哈希、黄色路由版本、财务规则版本、Prompt 版本和全部输出路径。原 Word 仅作模板，输出路径与模板相同时程序拒绝运行。

## 失败与人工审核策略

- OCR 失败或某个黄色指定来源无结果：记录到 `issues.json`，对应黄色内容留空。
- 本机无 LibreOffice 或 PyMuPDF：Word 仍可生成，字段审计表的原模板页码留空并记录问题。
- 金额及财务结果必填字段缺失：停止生成 Word；同字段同期间出现冲突候选时不自动选择，由 c2m 或评估师处理。
- 百炼叙述返回越权字段、无证据字段或未知证据编号：丢弃该字段并记录问题。
- 任一 LLM 审核失败：报告仍保留，审核结果标记为 `failed`，并记录到 `issues.json`，由人工复核。
- 企查查 API 未配置或企业身份核验不一致：对应 API 字段留空并记录复核事项。
- 已有同一 PDF 的 OCR 结果：默认复用缓存，不重复执行 OCR；取消“复用已有 OCR 结果”后才会强制重新 OCR。
- 生成完成：评估师同时审核 Word、字段审计清单和 OCR Excel；Demo 验收不代表生产上线。

## 扩展方式

- 新项目：新增 `projects/<project>.yaml` 和人工参数文件。
- 新 PDF：端到端入口支持替换 PDF，但“任意 PDF 直接盲填”不是可靠承诺。OCR、Excel 导出和 Word 复制是通用步骤；字段含义、审计表行列、黄色字段来源和目标 Word 位置仍由项目 YAML/映射配置声明。新 PDF 与现有审计版式一致时可直接复用；版式或报告模板变化时先更新对应配置并做一次人工验收。
- 新财务表：在项目配置的 `financial_tables` 中声明来源工作簿、工作表、单元格矩阵和 Word 表格编号。
- 新财务指标：在 `financial_fields` 中声明来源单元格和换算比例，并用 `final_value_field` 配置最终采用的评估结果字段。
- PDF OCR：PP-StructureV3 先输出统一页/块/表格契约，再生成独立 OCR Excel；通用财务字段可按别名、期间和单位匹配，版式敏感的附注字段在项目配置中声明页码、表格、行列定位。
- 新材料叙述：在 `material_fields` 中组合 Excel 单元格、Excel 范围、文件名或参考 Word 段落/表格；用 `paragraph_replacements` 替换模板中没有占位符的静态旧项目文字。
- 新模板：新增映射文件并运行模板回归测试。
- 新服务商：在 `adapters/` 实现相同输入输出契约，通过 `run.py` 注入。
- 新规则：修改 `domain/` 纯函数，同时更新 fixture、expected、测试和 `CHANGELOG.md`。Word 生成完成后必须没有任何 `XX/XXX/20XX` 占位符残留。
- 新 Prompt：新增带版本号的 Prompt 和输出结构，不覆盖旧版本。

## c2m 接入资产

生产接入优先复用：`schemas.py`、`domain/`、`prompts/`、`mappings/`、`fixtures/`、`expected/`、`tests/`、`workflow.yaml` 和 `data_manifest.yaml`。`run.py` 与 `adapters/` 仅用于 Demo 运行和接口示例，可由 c2m 的 FastAPI、后台任务、存储和安全体系替换。
