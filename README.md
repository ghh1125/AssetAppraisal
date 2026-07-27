# AssetAppraisal

可配置的资产评估报告 Demo 工作流：按现有材料读取 PDF/OCR/XLSX、企查查 API、百炼模型和用户输入，再依据模板字段来源规则填入 Word。PDF 和各类项目材料均可缺省；找不到的数据在 Word 保留黄色占位符，并输出带 Word 页码和具体位置的生成问题清单。问题清单包含“检查总览”和按页码排序的“问题明细”，便于评估师逐页复核。输出始终复制模板生成，不覆盖原模板。

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

接口地址和路径已经写入程序默认配置，通常不需要修改。接口无结果时，该来源字段在 Word 保留黄色占位符并写入问题清单，不会用其他来源冒填。

### 百炼模型

在[阿里云百炼控制台](https://bailian.console.aliyun.com/)开通兼容 OpenAI 接口并创建 API Key。默认配置为 `qwen3.7-max-2026-05-17`，模型和网关都可以自行修改：

```dotenv
DASHSCOPE_API_KEY=你的百炼APIKey
APPRAISAL_LLM_MODEL=qwen3.7-max-2026-05-17
```

项目配置还支持为 `narrative`、`format_review`、`data_validation`、`semantic_review` 分别指定模型；未指定时统一使用 `default_model`，环境变量优先级更高。

前端的主体概况模块支持多选：行业介绍、业务及细分市场、主要产品、客户供应商、盈利模式/SWOT、对标上市公司。至少选择一个；未选择的叙述模块不会调用模型，也不会写入 Word。

例如切换到其他模型，只改 `APPRAISAL_LLM_MODEL`。Prompt 和结构化输出契约分别位于：

- `demo/prompts/yellow_narratives.v2.txt`
- `demo/prompts/yellow_narratives_output.v2.json`
- `demo/prompts/review_format.v1.txt`
- `demo/prompts/review_data.v1.txt`
- `demo/prompts/review_semantic.v1.txt`
- `demo/prompts/review_output.v1.json`

修改模型时应保持输出字段白名单和 JSON 结构；叙述生成会先从上传的参考 Word 检索相关证据，再按七个字段分别调用模型并在本地校验证据编号。业务代码不会读取 `.env` 创建全局客户端，凭证只在 CLI/API 入口注入。

### OCR

PDF OCR 使用 PaddleOCR/PP-StructureV3。若同一 PDF 的 SHA-256 已命中 OCR 缓存，流程直接复用 `OCR结构化结果.xlsx`；未上传 PDF 时整个 OCR 节点明确跳过。缓存和运行目录均被 Git 忽略。

## 本地运行示例

`--pdf` 是可选参数；不提供时跳过 OCR，仍会根据人工输入、工作簿/API/LLM 的现有结果生成待复核 Word。

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

以下材料均为可选；上传框按材料角色识别，文件名可以不同，有多少就处理多少：

- 审计报告 PDF：应包含资产负债表、利润表及附注；系统会从中生成或复用 `OCR结构化结果.xlsx`。
- 参考评估报告 DOCX：建议上传已填/已复核报告，用于参考历史填报和税费表。
- 审计财务 XLSX/XLSM：通常包含资产负债表、利润表等审计财务表，可由 PDF OCR 后整理核验。
- 收益法 XLSX/XLSM：通常包含主要产品及服务、所得税表、净现金流计算表等收益法数据。
- 上报表或资产/市场法 XLSX/XLSM：通常包含资产基础法、长期资产或市场法数据。

Word 模板由后端固定提供。上传的 3 个 XLSX 会覆盖本次任务的项目材料路径，生成结果不会修改原文件；文件名可以任意。常见资产基础法、收益法和市场法表会按工作表特征、科目标签、列标题及元/万元自动定位，项目特有字段才需要补充项目映射。

生成 Word 后始终输出 `生成问题清单.xlsx` 和 `生成问题清单.json`。若启用百炼模型，还会生成 `格式审核.json`、`数据校验.json`、`语义审核.json` 和 `审核汇总.json`。工作流契约不合法、模板缺失或模板结构错误时停止生成；普通材料缺失、损坏、字段无结果或布局不匹配时继续生成待复核报告。
仅当财务字段完整、没有黄色占位符且至少一项审核完成时，才另存 `资产评估报告_最终候选.docx`。

```bash
cd frontend
npm run build
```
