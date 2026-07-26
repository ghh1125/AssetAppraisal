# AssetAppraisal

可配置的资产评估报告 Demo 工作流：把审计 PDF 识别为结构化 OCR Excel，再按模板中的字段来源规则，将 PDF/OCR/XLSX、企查查 API、百炼模型和用户输入分别填入 Word，并在生成后执行格式、数据和语义审核。输出始终复制模板生成，不覆盖原模板。

## 安全边界

仓库只提交业务内核、配置、Prompt、测试和前端示例。以下内容不会提交：

- `.env`、本地 API Key/SecretKey 和前端本地环境文件；
- `资产评估工作流/` 中的审计报告、Word 模板、Excel、OCR 材料等客户文件；
- `runs/`、`outputs/`、`uploads/`、`cache/` 运行产物和上传文件。

首次使用：

```bash
cp .env.example .env
# 在 .env 中填写本机凭证；不要把 .env 加入 Git
uv sync --extra dev
```

完整流程、配置格式、前端启动和 c2m 接入方式见 [demo/README.md](demo/README.md)。

仓库内的标准 Word 模板位于 `templates/评估报告版式-沟通标注版.docx`。审计 PDF、Excel 和参考报告仍需由本地用户自行提供，不随仓库提交。

## 需要的外部服务及获取方式

### 企查查

在企查查开放平台购买并创建以下接口，复制平台生成的 `Key` 和 `SecretKey` 到本地 `.env`：

| ApiCode | 用途 | 购买/查看地址 |
| --- | --- | --- |
| 735 | 企业工商详情（委托方、被评估单位工商表） | [企业工商详情](https://openapi.qcc.com/dataApi/735) |
| 231 | 全国商标查询 | [全国商标查询](https://openapi.qcc.com/dataApi/231) |
| 514 | 专利查询 | [专利查询](https://openapi.qcc.com/dataApi/514) |
| 233 | 著作权/软件著作权查询 | [著作权软著查询](https://openapi.qcc.com/dataApi/233) |

对应环境变量：

```dotenv
QICHACHA_APP_KEY=你的Key
QICHACHA_SECRET_KEY=你的SecretKey
```

接口地址和路径已经写入程序默认配置，通常不需要修改。接口无结果时，该来源字段按规则留空并写入 `issues.json`，不会用其他来源冒填。

### 百炼模型

在[阿里云百炼控制台](https://bailian.console.aliyun.com/)开通兼容 OpenAI 接口并创建 API Key。默认配置为 `qwen3.7-flash`，模型和网关都可以自行修改：

```dotenv
DASHSCOPE_API_KEY=你的百炼APIKey
APPRAISAL_LLM_MODEL=qwen3.7-flash
```

项目配置还支持为 `narrative`、`format_review`、`data_validation`、`semantic_review` 分别指定模型；未指定时统一使用 `default_model`，环境变量优先级更高。

前端的主体概况模块支持多选：行业介绍、业务及细分市场、主要产品、客户供应商、盈利模式/SWOT、对标上市公司。至少选择一个；未选择的叙述模块不会调用模型，也不会写入 Word。

例如切换到其他模型，只改 `APPRAISAL_LLM_MODEL`。Prompt 和结构化输出契约分别位于：

- `demo/prompts/yellow_narratives.v1.txt`
- `demo/prompts/yellow_narratives_output.v1.json`
- `demo/prompts/review_format.v1.txt`
- `demo/prompts/review_data.v1.txt`
- `demo/prompts/review_semantic.v1.txt`
- `demo/prompts/review_output.v1.json`

修改模型时应保持输出字段白名单和 JSON 结构；业务代码不会读取 `.env` 创建全局客户端，凭证只在 CLI/API 入口注入。

### OCR

PDF OCR 使用 PaddleOCR/PP-StructureV3。若同一 PDF 的 SHA-256 已命中 OCR 缓存，流程直接复用 `OCR结构化结果.xlsx`，跳过 OCR；缓存和运行目录均被 Git 忽略。

## 本地运行示例

```bash
uv run --python 3.11 python -m demo.run demo/projects/tongfu.yaml \
  --pdf "资产评估工作流/审计报告.pdf" \
  --template "templates/评估报告版式-沟通标注版.docx" \
  --output-dir runs/local \
  --use-glm --use-qichacha \
  --commissioning-party-name "委托方全称" \
  --commissioning-party-short-name "委托方简称" \
  --report-serial "001" \
  --valuation-purpose-inputs "评估目的" \
  --selected-valuation-method "收益法、资产基础法" \
  --valuation-subject-type "股东全部权益价值" \
  --transaction-type "收购" \
  --final-valuation-method "收益法" \
  --target-company-name "被评估单位全称" \
  --target-company-short-name "被评估单位简称"
```

API、OCR、LLM 和人工输入的字段边界由 `demo/projects/*.yaml` 与 `demo/workflow.yaml` 配置；新增项目优先复制配置并修改映射，不要把客户材料提交到仓库。

## 前后端启动

先在项目根目录准备 `.env`，然后启动后端：

```bash
uv sync --python 3.11 --extra dev --extra services --extra web
uv run --python 3.11 uvicorn demo.api_server:app --host 127.0.0.1 --port 8000
```

另开一个终端启动前端：

```bash
cd frontend
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

浏览器打开 <http://127.0.0.1:5173/>。前端默认把 `/api` 请求代理到 `http://127.0.0.1:8000`；如需修改，可在 `frontend/.env` 中设置 `VITE_API_PROXY_TARGET`，该文件不会提交到 Git。生产构建使用：

页面提交时需要同时选择以下材料；上传框按材料角色识别，文件名可以不同：

- 审计报告 PDF：应包含资产负债表、利润表及附注；系统会从中生成或复用 `OCR结构化结果.xlsx`。
- 参考评估报告 DOCX：建议上传已填/已复核报告，用于参考历史填报和税费表。
- 审计财务 XLSX：应包含 `06N_资产负债表`、`07N_利润表` 等审计财务表，可由 PDF OCR 后整理核验。
- 收益法 XLSX：应包含主要产品及服务、所得税表、净现金流计算表等收益法数据。
- 上报表 XLSX：应包含表 1、表 4-6、表 4-12 等资产基础法和长期资产数据。

Word 模板由后端固定提供。上传的 3 个 XLSX 会覆盖本次任务的项目材料路径，生成结果不会修改原文件；如果工作表名称或表格结构也不同，需要先在项目配置中增加对应映射。

生成 Word 后，若启用百炼模型，还会生成 `格式审核.json`、`数据校验.json`、`语义审核.json` 和 `审核汇总.json`。审核只提出问题和证据，不直接改写 Word；所有问题同时汇总到 `issues.json`。
审核完成后另存 `资产评估报告_最终候选.docx`，该文件是审核后的候选版本；需要人工根据问题清单确认后正式使用。

```bash
cd frontend
npm run build
```
