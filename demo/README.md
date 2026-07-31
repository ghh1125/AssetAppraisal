# 资产评估报告自动生成 Demo

## 定位

本目录是资产评估业务规则的可执行 Demo，不是生产系统。它负责表达和验证流程、输入输出契约、字段规则、Prompt、样例以及 Word 填充结果；用户、权限、任务队列、持久化、并发、重试、监控和发布由 c2m 实现。c2m 应移植 `domain/`、`schemas.py`、Prompt、映射和测试，不应把整个 Demo 仓库作为运行时依赖。

## 四个工作流节点

1. `start_input`（开始/输入）：接收图片定义的九项必填人工基础信息——委托方全称/简称、委托类型、评估主体全称/简称、评估对象、评估方法（多选且至少一个）、评估结论采用方法、评估报告编号流水号；资料入口是三个可选的有类型文件：审计报告 PDF、资产基础法/资产清查 Excel、收益法/基础法 Excel。Word 模板始终由后端提供，不接受用户覆盖。
2. `ocr_llm_candidates`（材料解析与候选生成）：有 PDF 才执行 OCR（命中 SHA-256 缓存则复用）；解析 OCR/XLSX/API 的确定性字段；百炼模型一次为模板中七个固定 LLM 位置生成候选文本，并将候选内容和对应 Word 位置返回。该节点暂停，等待用户逐项选择。
3. `fill_word`（Word 填充）：只把用户选择的 LLM 候选，加上人工、OCR/XLSX 和企查查字段，按黄色路由和表格语义映射写入复制后的 Word。原模板不改；未命中内容不编造。
4. `output`（结果输出）：只输出独立的待复核 Word。找不到的内容保留并高亮 `XXX`，供人工在 Word 中复核。

运行前会校验 `workflow.yaml` 中四个节点引用的 Pydantic 输入输出模型、依赖关系、描述和字段业务说明。工作流只保留上述四类节点，不再执行格式审核、数据校验、语义审核或审核汇总。

## 安装与运行

```bash
uv sync --extra dev
uv run python -m demo.run demo/projects/tongfu.yaml --offline
```

`--offline` 不调用企业 API 和 LLM。所有业务材料均可缺省；找不到的字段在 Word 保留原 `XX/XXX/20XX`（没有原标记时使用 `XXX`），只标黄占位符本身。可用 `--output-dir` 指定独立输出目录。原 Word 永远作为只读模板，不会被覆盖。

在 c2m 或其他宿主中，可直接调用 `run_project(...)` / `run_pipeline(...)` 并通过 `ocr_adapter`、`company_api_adapter` 和 `llm_adapter` 参数注入已有服务。Demo 不在 `domain/` 内创建客户端或读取密钥；注入结果只接受映射表中已经登记的字段键。

项目继续使用 Python 3.11。端到端 OCR 默认使用阿里云文档智能，安装 `services` 依赖即可，不在本地加载 PaddleOCR。项目根目录的本地 `.env` 会在 Demo 入口自动加载；生产环境也可由宿主进程注入同名环境变量。阿里云 RAM 用户需授予 `AliyunDocmindFullAccess`：

```dotenv
APPRAISAL_OCR_PROVIDER=aliyun
ALIBABA_CLOUD_ACCESS_KEY_ID=你的RAM AccessKey ID
ALIBABA_CLOUD_ACCESS_KEY_SECRET=你的RAM AccessKey Secret
APPRAISAL_OCR_VLM=false
APPRAISAL_OCR_TIMEOUT_SECONDS=900
```

百炼和企查查仍分别使用 `DASHSCOPE_API_KEY`、`QICHACHA_APP_KEY` 与 `QICHACHA_SECRET_KEY`，三类服务凭证互不替代。

命令中的 `--pdf` 可省略；省略后跳过 OCR，工作流从其他可用材料和人工输入继续。同一 PDF 命中 SHA-256 缓存时不会再次消耗云端页数。云端失败时不自动回退本地模型，相关字段保留黄色占位符。

```bash
uv python install 3.11
uv sync --python 3.11 --extra dev --extra services
uv run --python 3.11 python -m demo.run demo/projects/tongfu.yaml \
  --pdf '资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf' \
  --template 'templates/评估报告版式_v2.docx' \
  --output-dir runs/tongfu-ocr --ocr-provider aliyun --use-glm --use-qichacha \
  --commissioning-party-name '委托方全称' \
  --commissioning-party-short-name '委托方简称' \
  --report-serial '报告流水号' \
  --selected-valuation-method '收益法、资产基础法' \
  --valuation-subject-type '股东全部权益价值' \
  --transaction-type '收购' \
  --final-valuation-method '收益法' \
  --target-company-short-name '通富昆山'
```

仅在需要显式启用本地高内存模式时执行 `uv sync --extra ocr` 并设置 `APPRAISAL_OCR_PROVIDER=paddle` 或命令行参数 `--ocr-provider paddle`。`APPRAISAL_OCR_PROVIDER=none` 会完全跳过 PDF OCR。

如需调用已购买的企查查接口 735（工商详情）、231（商标）、514（专利）和 233（著作权软著），设置 `QICHACHA_APP_KEY`、`QICHACHA_SECRET_KEY` 并增加 `--use-qichacha`。默认使用企查查官方签名方式；如平台给 514 或 233 分配了不同的路径，可用 `QICHACHA_ENDPOINT_514`、`QICHACHA_ENDPOINT_233` 覆盖，基地址可用 `QICHACHA_API_BASE_URL` 覆盖。两个节点输入可放入 JSON 文件并用 `--node-inputs-json` 传入。凭证只在 `run.py` 这一组合入口读取，不会进入 `domain/`、内部运行状态或 Word。

## 最新 Word 批注来源路由

`templates/评估报告版式-沟通标注版_批注版.docx` 是批注规则版，包含 100 条 Word 批注、131 个占位符和 11 个无占位符的批注锚点；`templates/评估报告版式_v2.docx` 是实际复制填充的干净模板。批注优先于旧的黄色坐标映射，运行时逐占位符写入 `template_comments.json` 和 `workflow_trace.json`，保留批注原文、字段键和来源类别。

批注来源类别如下：

- `node_input`：人工基础信息中的公司名称、简称、交易类型、评估对象和评估方法；
- `qichacha_api`：工商概况、股东及股权、商标、专利、软件著作权等；
- `pdf_ocr_xlsx`：审计报告日期、范围、历史财务表、税率、长期资产和资产基础法结果；
- `bailian_glm` / `mixed`：先用已核验的 PDF/XLSX/API 事实生成 Word 固定 LLM 位置候选，用户选择后才写入；
- `system` / `derived`：报告日期、有效期、评估增值和大写金额等确定性派生值；
- `unresolved_manual`：批注明确“获取不到/未确认/暂时不做填充”的位置，强制保留原 `XX/XXX` 并标黄，不会被同段旧坐标或人工字段误填。

运行时仍按来源白名单过滤适配器返回值。批注写“API 或大模型”的字段保留 `mixed` 证据链：API 用于事实，百炼只负责叙述生成；不会把 LLM 文本写入金额、股东或税率字段。指定来源无值时保留并标黄原占位符。

## 输出

- `资产评估报告_待复核.docx`：唯一面向用户的输出。模板始终复制后填充，不覆盖原模板；找不到的值保留黄色 `XXX`。

OCR 缓存、LLM 候选和运行轨迹属于流程内部状态，不作为用户下载项。

## Vue 前端工作台

`frontend/` 将图片定义的九项必填人工输入、三个可选的有类型材料入口、后端固定 Word 模板和产物下载做成页面。上传文件名不需要与配置一致，后端按扩展名、工作表标题、科目行和金额列语义识别资产基础法/清查 Excel 与收益法/基础法 Excel；只有存在 PDF 时才检查 OCR 缓存。Word 模板和 OCR 结构化 Excel 都由后端提供，不需要用户上传。

模板批注：仓库中的 `templates/评估报告版式-沟通标注版_批注版.docx`保留新版 Word 批注，`templates/评估报告版式_v2.docx`是最终输出版式。项目配置通过 `annotation_template` 读取批注，通过 `web_template`/`template` 填充并输出普通版 Word；运行时会将批注 ID、批注文本和对应占位符写入内部 `template_comments.json`，批注不会出现在最终报告正文。

任务结果区会实时显示四个节点的时间线和状态。节点 2 完成后状态为“等待人工选择”，用户确认候选后才继续节点 3 和节点 4；节点状态同时写入 API 返回的 `nodes` 和 `workflow_trace.json`。

上传框按材料角色接收文件，原文件名不需要与配置一致。程序先使用已验证项目的精确坐标；坐标不存在、为空或与新版式明显冲突时，再按工作表特征、科目行、列标题和金额单位做语义定位。当前通用规则覆盖资产基础法/市场法的账面净资产和评估价值、收益法股东全部权益价值、资产负债汇总、历史资产负债表、历史利润表、主要长期资产、主要产品及增值税率。无法唯一匹配时保留黄色 `XXX`，不用相邻非零单元格猜值。

Web 任务输出目录按“`YYYYMMDDHHMM-PDF文件名`”命名，例如 `runs/web/202607231144-通富2025.6.30合并及母公司审计报告/`；同一分钟重复提交会自动追加序号，不覆盖已有任务。

```bash
uv sync --extra services --extra web --python 3.11
uv run --python 3.11 uvicorn demo.api_server:app --host 127.0.0.1 --port 8000

cd frontend
npm install
npm run dev
```

页面第一节点只提交图片定义的九项人工基础信息和三个可选的类型化文件。任务创建后，第二节点先调用百炼为 Word 固定 LLM 位置生成候选，用户逐项勾选后再提交填充。企查查、百炼和 PDF OCR/XLSX 不要求用户重复填写；批注明确未确认的字段继续保留并高亮 `XXX`。

运行轨迹和标准化字段只用于流程内部，不向用户提供下载链接。原 Word 仅作模板，输出路径与模板相同时程序拒绝运行。

## 失败与人工审核策略

- 未上传 PDF：OCR 和 OCR Excel 导出节点标记为 `skipped`，其余节点继续。
- OCR 失败或某个指定来源无结果：Word 保留黄色占位符，流程继续生成报告。
- `workflow.yaml` 节点、模型、字段说明或依赖不符合契约：在任何外部调用前停止运行。
- 本机无 LibreOffice 或 PyMuPDF：Word 仍可生成，缺失内容保留黄色占位符。
- 金额及财务结果字段缺失：对应段落或表格单元格写黄色 `XXX`，仍生成待复核 Word。
- 同字段同期间出现冲突候选：不自动选择，保留待复核 Word 和冲突记录，由 c2m 或评估师处理。
- 百炼叙述返回越权字段、无证据字段或未知证据编号：丢弃该字段并保留黄色占位符。
- LLM 候选生成失败：对应固定位置保留黄色 `XXX`，继续生成 Word；不会执行额外的格式、数据或语义审核。
- 企查查 API 未配置或企业身份核验不一致：对应 API 字段保留黄色占位符，继续生成报告供人工复核。
- 已有同一 PDF 的 OCR 结果：默认复用缓存，不重复执行 OCR；取消“复用已有 OCR 结果”后才会强制重新 OCR。
- 生成完成：评估师直接打开评估报告 Word，按黄色 `XXX` 逐项人工复核；Demo 验收不代表生产上线。

## 扩展方式

- 新项目：新增 `projects/<project>.yaml` 和人工参数文件。
- 新 PDF：端到端入口支持替换 PDF，但“任意 PDF 直接盲填”不是可靠承诺。OCR、Excel 导出和 Word 复制是通用步骤；字段含义、黄色字段来源和目标 Word 位置仍由项目 YAML/映射配置声明。新 PDF 与现有审计版式一致时可直接复用；版式或报告模板变化时先更新对应配置并做一次人工验收。
- 新财务表：通用科目先由 `workbook_semantics.v3` 自动识别，并根据工作簿内的收益法/市场法语义归类估值结果；语义结果是主路径，`financial_tables` 和固定坐标只在语义结果缺失时回退或补空。疑似尚未完成评估的全零评估列不会作为有效结果。
- 新财务指标：通用估值结果优先按标签和表头定位；项目特有指标在 `financial_fields` 中声明来源单元格和换算比例，并用 `final_value_field` 配置最终采用的评估结果字段。
- PDF OCR：阿里云文档智能默认输出统一页/块/表格契约，再生成独立 OCR Excel；本地 PP-StructureV3 仅作为显式可选提供方。通用财务字段可按别名、期间和单位匹配，版式敏感的附注字段在项目配置中声明页码、表格、行列定位。
- 新材料叙述：参考 Word 会按主题自动提取带编号证据；项目特有的确定性内容仍可在 `material_fields` 中组合 Excel 单元格、Excel 范围或参考 Word 段落/表格，并用 `paragraph_replacements` 替换模板中没有占位符的静态旧项目文字。
- 新模板：新增映射文件并运行模板回归测试。
- 新服务商：在 `adapters/` 实现相同输入输出契约，通过 `run.py` 注入。
- 新规则：修改 `domain/` 纯函数，同时更新 fixture、expected、测试和 `CHANGELOG.md`。未解决的 `XX/XXX/20XX` 必须标黄。
- 新 Prompt：新增带版本号的 Prompt 和输出结构，不覆盖旧版本。

## c2m 接入资产

生产接入优先复用：`schemas.py`、`domain/`、`prompts/`、`mappings/`、`fixtures/`、`expected/`、`tests/`、`workflow.yaml` 和 `data_manifest.yaml`。`run.py`、`pipeline.py`、`api_server.py` 与 `adapters/` 仍是 Demo 外壳，可由 c2m 的用户权限、后台任务、持久化、对象存储、超时重试、监控和发布体系替换。
