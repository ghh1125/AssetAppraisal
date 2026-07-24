# AssetAppraisal

可配置的资产评估报告 Demo 工作流：把审计 PDF 识别为结构化 OCR Excel，再按模板中的字段来源规则，将 PDF/OCR/XLSX、企查查 API、百炼 GLM 和用户输入分别填入 Word。输出始终复制模板生成，不覆盖原模板。

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

### 百炼 GLM

在[阿里云百炼控制台](https://bailian.console.aliyun.com/)开通兼容 OpenAI 接口并创建 API Key。默认配置为 `glm-5.2`，模型和网关都可以自行修改：

```dotenv
DASHSCOPE_API_KEY=你的百炼APIKey
APPRAISAL_LLM_MODEL=glm-5.2
```

例如切换到其他模型，只改 `APPRAISAL_LLM_MODEL`。Prompt 和结构化输出契约分别位于：

- `demo/prompts/yellow_narratives.v1.txt`
- `demo/prompts/yellow_narratives_output.v1.json`

修改模型时应保持输出字段白名单和 JSON 结构；业务代码不会读取 `.env` 创建全局客户端，凭证只在 CLI/API 入口注入。

### OCR

PDF OCR 使用 PaddleOCR/PP-StructureV3。若同一 PDF 的 SHA-256 已命中 OCR 缓存，流程直接复用 `OCR结构化结果.xlsx`，跳过 OCR；缓存和运行目录均被 Git 忽略。

## 本地运行示例

```bash
uv run --python 3.11 python -m demo.run demo/projects/tongfu.yaml \
  --pdf "资产评估工作流/审计报告.pdf" \
  --template "资产评估工作流/评估报告模板.docx" \
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

```bash
cd frontend
npm run build
```
