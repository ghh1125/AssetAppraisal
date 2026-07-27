# 资产评估报告自动生成 Demo

## 定位

本目录是资产评估业务规则的可执行 Demo，不是生产系统。它负责表达和验证流程、输入输出契约、字段规则、Prompt、样例以及 Word 填充结果；用户、权限、任务队列、持久化、并发、重试、监控和发布由 c2m 实现。c2m 应移植 `domain/`、`schemas.py`、Prompt、映射和测试，不应把整个 Demo 仓库作为运行时依赖。

## 工作流节点

1. `inventory`：识别 Word 中的稳定位置。
2. `ocr_pdf`：有 PDF 时用 PP-StructureV3 识别页、文本块、表格和单元格；无 PDF 时标记为 `skipped`。
3. `export_ocr_workbook`：有 OCR 输入时生成独立的 `OCR结构化结果.xlsx`，否则跳过。
4. `extract_sources`：读取 OCR Excel、其他结构化 Excel、节点参数和外部服务结果。
5. `resolve_fields`：按字段语义、期间、单位和固定黄色来源路由选择字段值。
6. `select_narrative_modules`：校验用户勾选的六类主体概况模块。
7. `generate_narrative`：通过注入的百炼模型生成用户勾选的叙述内容。
8. `fill_word`：复制模板并填入可用值；未解析位置保留黄色占位符。
9. `generate_narrative`：从参考 Word、工商信息和结构化字段检索证据，按七个叙述字段分别调用模型并校验证据编号。
10. `llm_format_review`、`llm_data_validation`、`llm_semantic_review`：对生成 Word 执行格式、数据和语义审核。
10. `review_aggregate`：汇总三类审核结果和问题。
11. `export_audit`：导出字段来源清单和运行记录。

运行前会校验 `workflow.yaml` 中每个节点引用的 Pydantic 输入输出模型、依赖关系和字段业务说明；契约无效时在 OCR、API、LLM 和 Word 调用前终止。人工检查点包括主体概况模块选择、三类审核问题确认及最终 Word/审计清单复核。

## 安装与运行

```bash
uv sync --extra dev
uv run python -m demo.run demo/projects/tongfu.yaml --offline
```

`--offline` 不调用企业 API 和 LLM。所有业务材料均可缺省；找不到的字段在 Word 保留原 `XX/XXX/20XX`（没有原标记时使用 `XXX`），只标黄占位符本身，并记录到生成问题清单。可用 `--output-dir` 指定独立输出目录。原 Word 永远作为只读模板，不会被覆盖。

在 c2m 或其他宿主中，可直接调用 `run_project(...)` / `run_pipeline(...)` 并通过 `ocr_adapter`、`company_api_adapter`、`llm_adapter` 和 `review_adapters` 参数注入已有服务。Demo 不在 `domain/` 内创建客户端或读取密钥；注入结果只接受映射表中已经登记的字段键。

端到端 OCR 使用独立的 Python 3.11 环境（PaddleOCR 当前可选依赖限定 Python `<3.13`）。项目根目录的本地 `.env` 会在 Demo 入口自动加载；也可以由宿主进程自行注入同名环境变量，例如 `DASHSCOPE_API_KEY`、`QICHACHA_APP_KEY` 和 `QICHACHA_SECRET_KEY`：

命令中的 `--pdf` 可省略；省略后不会加载 PaddleOCR，工作流从其他可用材料和人工输入继续。

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
- 百炼审核：格式、数据、语义 3 个独立任务；默认都使用 `qwen3.7-max-2026-05-17`，可在项目 YAML 的 `llm.tasks` 中分别改模型。
- 企查查 API：5 个——委托方概况、历史股权沿革、基准日股权、账外无形资产、软件著作权。
- PDF OCR/XLSX：6 个——历史资产负债表、历史利润表、税率、评估范围、主要长期资产、资产基础法结果。
- 节点输入：2 个——选用评估方法、评估目的输入。

运行时会再次按白名单过滤每个适配器的返回值。指定来源无值时留空，不使用参考 Word、其他 API、其他 Excel 或 LLM 跨路由补值，也不写入“待人工补充”。现有 `XX/XXX` 非黄色占位符继续按项目映射和确定性规则填写。

## 输出

- `OCR结构化结果.xlsx`
  - `OCR_表格`：逐单元格审计明细；`OCR_表格索引` 及 `表_<页码>_<表格编号>`：按 OCR 行列恢复的矩阵表，便于人工查看和后续映射。
- `资产评估报告_待复核.docx`
- `资产评估报告_最终候选.docx`（仅在财务字段完整、无黄色占位符且至少一项审核完成时生成）
- `字段审计清单.xlsx`
- `生成问题清单.xlsx`、`生成问题清单.json`
- `normalized_fields.json`
- `issues.json`
- `格式审核.json`、`数据校验.json`、`语义审核.json`
- `审核汇总.json`
- `run_manifest.json`
- `workflow_trace.json`
  - 按 `workflow.yaml` 顺序记录 13 个节点的执行/跳过状态、输入输出模型、结构化摘要、人工检查点、来源证据、问题以及规则/Prompt/模型/数据版本。

## Vue 前端工作台

`frontend/` 将人工输入、可选材料上传、后端固定 Word 模板、GLM/企查查开关和产物下载做成页面。PDF、参考 DOCX 和三个项目工作簿均可选，只需至少上传一份材料或填写一项基础信息；工作簿支持 `.xlsx/.xlsm`。只有选择 PDF 后才检查 OCR 缓存。Word 模板由后端固定提供，不需要上传。

上传框按材料角色接收文件，原文件名不需要与配置一致。程序先使用已验证项目的精确坐标；坐标不存在、为空或与新版式明显冲突时，再按工作表特征、科目行、列标题和金额单位做语义定位。当前通用规则覆盖资产基础法/市场法的账面净资产和评估价值、收益法股东全部权益价值、资产负债汇总、历史资产负债表、历史利润表、主要长期资产、主要产品及增值税率。无法唯一匹配时保留黄色 `XXX` 并进入问题清单，不用相邻非零单元格猜值。

Web 任务输出目录按“`YYYYMMDDHHMM-PDF文件名`”命名，例如 `runs/web/202607231144-通富2025.6.30合并及母公司审计报告/`；同一分钟重复提交会自动追加序号，不覆盖已有任务。

```bash
uv sync --extra services --extra ocr --extra web --python 3.11
uv run --python 3.11 uvicorn demo.api_server:app --host 127.0.0.1 --port 8000

cd frontend
npm install
npm run dev
```

页面提交的字段与命令行参数一一对应：委托方全称/简称、评估主体全称/简称、报告编号流水号、评估目的、评估方法（资产基础法/收益法/市场法多选）、评估对象（四选一）、委托类型（转让/收购/增资/减资）、评估结论采用方法（单选）和六类主体概况模块。报告编号可以输入完整编号，程序会提取模板所需的流水号并同步所有报告编号年份。企查查、百炼和 PDF OCR/XLSX 不要求用户在页面重复填写。

字段审计清单逐行记录原模板页码、稳定位置编号、原文上下文、程序填入内容、来源类别、来源文件和来源位置。原模板页码由 LibreOffice 只读渲染模板后按段落顺序匹配得到；页脚标记为“多页页脚”。`workflow_trace.json` 记录节点级契约执行情况，不复制 OCR 全文，也不记录 API 密钥。OCR、API、LLM、财务数据及评估结论的正确性由业务人员人工审核。

`OCR结构化结果.xlsx` 固定包含 `OCR_文本`、`OCR_表格`、`标准财务数据`、`识别问题` 四个工作表，保存页码、行列、坐标、置信度和证据编号。`run_manifest.json` 保存模板/PDF 哈希、黄色路由版本、财务规则版本、Prompt 版本和全部输出路径。原 Word 仅作模板，输出路径与模板相同时程序拒绝运行。

## 失败与人工审核策略

- 未上传 PDF：OCR 和 OCR Excel 导出节点标记为 `skipped`，其余节点继续。
- OCR 失败或某个指定来源无结果：Word 保留黄色占位符，并写入 `issues.json` 和生成问题清单。
- `workflow.yaml` 节点、模型、字段说明或依赖不符合契约：在任何外部调用前停止运行。
- 本机无 LibreOffice 或 PyMuPDF：Word 仍可生成，字段审计表的原模板页码留空并记录问题。
- 金额及财务结果字段缺失：对应段落或表格单元格写黄色 `XXX`，生成待复核 Word，并在问题清单、字段审计、运行清单和工作流轨迹中标为高优先级；不生成最终候选 Word。
- 同字段同期间出现冲突候选：不自动选择，保留待复核 Word 和冲突记录，不生成最终候选 Word，由 c2m 或评估师处理。
- 百炼叙述返回越权字段、无证据字段或未知证据编号：丢弃该字段并记录问题。
- 任一 LLM 审核失败：报告仍保留，审核结果标记为 `failed`，其他审核继续；未启用的审核在轨迹中标记为 `skipped`。
- 企查查 API 未配置或企业身份核验不一致：对应 API 字段保留黄色占位符并记录复核事项。
- 已有同一 PDF 的 OCR 结果：默认复用缓存，不重复执行 OCR；取消“复用已有 OCR 结果”后才会强制重新 OCR。
- 生成完成：评估师先查看生成问题清单的“检查总览”，再按“问题明细”的 Word 页码、检查位置、来源文件和处理建议逐项审核，同时核对字段审计清单和 OCR Excel；Demo 验收不代表生产上线。

## 扩展方式

- 新项目：新增 `projects/<project>.yaml` 和人工参数文件。
- 新 PDF：端到端入口支持替换 PDF，但“任意 PDF 直接盲填”不是可靠承诺。OCR、Excel 导出和 Word 复制是通用步骤；字段含义、审计表行列、黄色字段来源和目标 Word 位置仍由项目 YAML/映射配置声明。新 PDF 与现有审计版式一致时可直接复用；版式或报告模板变化时先更新对应配置并做一次人工验收。
- 新财务表：通用科目先由 `workbook_semantics.v3` 自动识别，并根据工作簿内的收益法/市场法语义归类估值结果；语义结果是主路径，`financial_tables` 和固定坐标只在语义结果缺失时回退或补空。疑似尚未完成评估的全零评估列不会作为有效结果。
- 新财务指标：通用估值结果优先按标签和表头定位；项目特有指标在 `financial_fields` 中声明来源单元格和换算比例，并用 `final_value_field` 配置最终采用的评估结果字段。
- PDF OCR：PP-StructureV3 先输出统一页/块/表格契约，再生成独立 OCR Excel；通用财务字段可按别名、期间和单位匹配，版式敏感的附注字段在项目配置中声明页码、表格、行列定位。
- 新材料叙述：参考 Word 会按主题自动提取带编号证据；项目特有的确定性内容仍可在 `material_fields` 中组合 Excel 单元格、Excel 范围或参考 Word 段落/表格，并用 `paragraph_replacements` 替换模板中没有占位符的静态旧项目文字。
- 新模板：新增映射文件并运行模板回归测试。
- 新服务商：在 `adapters/` 实现相同输入输出契约，通过 `run.py` 注入。
- 新规则：修改 `domain/` 纯函数，同时更新 fixture、expected、测试和 `CHANGELOG.md`。未解决的 `XX/XXX/20XX` 必须标黄并逐项进入生成问题清单。
- 新 Prompt：新增带版本号的 Prompt 和输出结构，不覆盖旧版本。

## c2m 接入资产

生产接入优先复用：`schemas.py`、`domain/`、`prompts/`、`mappings/`、`fixtures/`、`expected/`、`tests/`、`workflow.yaml` 和 `data_manifest.yaml`。c2m 可以消费 `workflow_trace.json` 和字段审计清单完成任务状态及审计接入；`run.py`、`pipeline.py`、`api_server.py` 与 `adapters/` 仍是 Demo 外壳，可由 c2m 的用户权限、后台任务、持久化、对象存储、超时重试、监控和发布体系替换。
