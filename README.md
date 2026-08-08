# AssetAppraisal

可配置的资产评估报告 Demo 工作流：按现有材料读取 PDF/OCR/XLSX、企查查 API、百炼模型和用户输入，再依据模板字段来源规则填入 Word。PDF 和各类项目材料均可缺省；找不到的数据在 Word 保留黄色占位符。最终只输出复制模板生成的独立评估报告 Word，不覆盖原模板。

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

仓库内保留模板和标注两个版本：`templates/评估报告版式_v2.docx`是最终输出版式，`templates/评估报告版式-沟通标注版_批注版.docx`是批注规则版。项目配置用 `web_template`/`template` 指定输出模板、用 `annotation_template` 指定批注模板；批注只用于建立“Word位置→字段→来源”映射，不会出现在最终报告正文。运行时用户只需上传审计 PDF、资产基础法/资产清查 Excel 和收益法/基础法 Excel（均可选）；客户材料不随仓库提交。

## 需要的外部服务及获取方式

### 企查查

在企查查开放平台购买并创建以下接口，复制平台生成的 `Key` 和 `SecretKey` 到本地 `.env`：

| ApiCode | 用途 | 购买/查看地址 |
| --- | --- | --- |
| 735 | 企业工商详情（委托方、被评估单位工商表） | [企业工商详情](https://openapi.qcc.com/dataApi/735) |
| 231 | 全国商标查询 | [全国商标查询](https://openapi.qcc.com/dataApi/231) |
| 514 | 专利查询 | [专利查询](https://openapi.qcc.com/dataApi/514) |
| 233 | 著作权/软件著作权查询 | [著作权软著查询](https://openapi.qcc.com/dataApi/233) |
| 2001 | 企业信息核验，补充行业、经营范围和企业规模证据 | [企业信息核验](https://openapi.qcc.com/dataApi/2001) |
| 213 | 企业年报，补充经营及年报证据 | [企业年报](https://openapi.qcc.com/dataApi/213) |
| 886 | 按产品、经营范围等关键词检索同行业企业候选 | [企业模糊搜索](https://openapi.qcc.com/dataApi/886) |
| 915 | 按业务关键词检索上市公司公告及股票代码 | [上市公告搜索](https://openapi.qcc.com/dataApi/915) |
| 699 | 补全 915 返回的上市候选公司的简介和主要指标 | [上市企业](https://openapi.qcc.com/dataApi) |

对应环境变量：

```dotenv
QICHACHA_APP_KEY=你的Key
QICHACHA_SECRET_KEY=你的SecretKey
# 留空时调用 735/231/514/233/2001/213；可显式限制补充证据接口
QICHACHA_EXTRA_API_CODES=2001,213
```

接口地址和路径已经写入程序默认配置，通常不需要修改。启用 `--use-qichacha` 时，企查查优先在节点 2 调用：默认调用 735、231、514、233、2001、213；前四项填工商及知识产权表，后两项只作为百炼叙述证据。对被评估主体，节点 2 会从 API 返回的经营范围中提取短业务关键词，再调用 886→915→699，形成可追溯的上市公司候选及其公开指标；被评估主体自身绝不会作为对标上市公司写入。经主体名称校验的响应会保存为运行目录内部快照；节点 3 优先复用每个主体的有效快照，只对节点 2 未取得或身份校验失败的主体再次调用接口补充，以填充率优先。可用 `QICHACHA_ENABLE_COMPARABLE_DISCOVERY=false` 关闭对标公司发现。962、521、723、724、1124 等面议接口不调用。客户、供应商没有 PDF/XLSX 等可靠证据时保留黄色占位符，不由 LLM 编造。接口路径可分别用 `QICHACHA_ENDPOINT_735`、`QICHACHA_ENDPOINT_231`、`QICHACHA_ENDPOINT_514`、`QICHACHA_ENDPOINT_233`、`QICHACHA_ENDPOINT_2001`、`QICHACHA_ENDPOINT_213`、`QICHACHA_ENDPOINT_886`、`QICHACHA_ENDPOINT_915` 覆盖。

### 百炼模型

在[阿里云百炼控制台](https://bailian.console.aliyun.com/)开通兼容 OpenAI 接口并创建 API Key。默认先调用 `deepseek-v4-flash-0731`；该次调用失败时自动改用 `qwen3.8-max`，两个模型和网关都可以自行修改：

```dotenv
DASHSCOPE_API_KEY=你的百炼APIKey
APPRAISAL_LLM_MODEL=deepseek-v4-flash-0731
APPRAISAL_LLM_FALLBACK_MODEL=qwen3.8-max
```

项目配置可以为 `narrative` 指定模型；未指定时统一使用 `default_model`，环境变量优先级更高。`APPRAISAL_LLM_FALLBACK_MODEL` 仅在主模型请求失败时启用，不会混合两种模型的文本。

公司概况在有确定证据时自动写入。第二节点固定展示并让用户选择六个可选模块：所处行业及行业介绍、业务内容及细分市场、主要产品、主要客户及供应商、盈利模式和SWOT分析、对标上市公司（列表多维度展示）。候选没有可靠证据时会明确显示“暂无可用证据”，该位置不写入并保留黄色 `XXX`。用户勾选的是“哪些候选写入 Word”，不是重新选择 API；API 负责提供可核验事实，LLM 只负责基于这些事实综合成对应位置的候选文本。

例如切换到其他模型，只改 `APPRAISAL_LLM_MODEL`。Prompt 和结构化输出契约分别位于：

- `demo/prompts/yellow_narratives.v2.txt`
- `demo/prompts/yellow_narratives_output.v2.json`

修改模型时应保持输出字段白名单和 JSON 结构；如项目后端配置了参考资料，叙述生成会先检索相关证据，再按公司概况和六个可选模块分别调用模型并在本地校验证据编号。参考报告不是前端运行时上传项。业务代码不会读取 `.env` 创建全局客户端，凭证只在 CLI/API 入口注入。

### OCR

OCR 在本工作流中只负责把审计报告 PDF 转换为可以检索、匹配和追溯的结构化证据。它不会直接决定 Word 应填什么，也不会替代资产基础法/资产清查 Excel、收益法或市场法 Excel、企查查、百炼及黄色字段来源规则。

完整链路如下：

```mermaid
flowchart TD
    A["前端上传审计报告 PDF"] --> B["计算 PDF SHA-256"]
    B --> C{"命中 OCR 缓存？"}
    C -->|是| D["复用 OCR结构化结果.xlsx"]
    C -->|否| E["调用阿里云文档智能"]
    E --> F["异步提交、状态轮询、分页获取结果"]
    F --> G["归一化页、文本块、表格、单元格和坐标"]
    G --> H["导出 OCR结构化结果.xlsx"]
    D --> I["按科目、期间、表头和单位进行语义匹配"]
    H --> I
    I --> J["只进入 PDF OCR/XLSX 允许的字段"]
    J --> K["与人工输入、业务 Excel、企查查和百炼结果合并"]
    K --> L["复制模板并填充 Word"]
    L --> M["缺失保留黄色 XXX，输出评估报告 Word"]
```

#### 1. 开通与配置

PDF OCR 默认使用阿里云文档智能“文档解析（大模型版）”，本机不会加载 PaddleOCR。先在[文档智能控制台](https://docmind.console.aliyun.com/)开通“文档理解 → 文档解析（大模型版）”，再给持有 AccessKey 的专用 RAM 用户授权。

阿里云官方文档中的系统策略名称为 `AliyunDocmindFullAccess`。若当前 RAM 控制台没有显示该系统策略，可创建只覆盖文档智能产品的自定义策略并授予同一个 RAM 用户：

```json
{
  "Version": "1",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": ["docmind:*"],
      "Resource": "*"
    }
  ]
}
```

授权和自定义策略可分别参考[文档智能服务鉴权指南](https://help.aliyun.com/zh/document-mind/getting-started/service-authentication-guide)与[文档智能自定义权限策略参考](https://help.aliyun.com/zh/document-mind/security-and-compliance/document-smart-custom-permission-policy-reference)。

本地 `.env` 配置为：

```dotenv
APPRAISAL_OCR_PROVIDER=aliyun
ALIBABA_CLOUD_ACCESS_KEY_ID=你的RAM AccessKey ID
ALIBABA_CLOUD_ACCESS_KEY_SECRET=你的RAM AccessKey Secret
APPRAISAL_OCR_VLM=false
APPRAISAL_OCR_TIMEOUT_SECONDS=900
```

基础链路适用于普通审计报告；复杂扫描件可将 `APPRAISAL_OCR_VLM` 改为 `true`，但费用和处理时间会增加。AccessKey 只从 `.env` 或生产宿主环境注入，不进入业务代码、Word、内部运行状态或 Git。

#### 2. 缓存与 OCR 提供方选择

上传 PDF 后，前端会立即查询 OCR 缓存。缓存使用 PDF 内容的 SHA-256，而不是文件名判断：

1. 命中时直接读取已有 `OCR结构化结果.xlsx`，不会再次消耗云端页数；
2. 未命中时调用当前配置的 OCR 提供方；
3. 缓存文件损坏且仍有原 PDF 时，记录缓存问题并尝试调用当前提供方；
4. 未上传 PDF 时，`ocr_pdf` 和 `export_ocr_workbook` 节点标记为 `skipped`，其余材料仍继续处理。

提供方配置：

- `APPRAISAL_OCR_PROVIDER=aliyun`：默认模式，调用阿里云文档智能，本机不加载 PaddleOCR；
- `APPRAISAL_OCR_PROVIDER=paddle`：显式启用原本地高内存模式，需要安装 `--extra ocr`；
- `APPRAISAL_OCR_PROVIDER=none`：完全跳过 PDF OCR。

云端失败时不会自动回退本地 PaddleOCR，避免突然占用大量本机内存。

#### 3. 阿里云解析和统一结构

缓存未命中时，适配器按阿里云异步接口执行：

1. 使用 `SubmitDocParserJobAdvance` 上传本地 PDF；
2. 轮询 `QueryDocParserStatus`，直到成功、失败或超时；
3. 使用 `GetDocParserResult` 分页获取全部版面块；
4. 把阿里云返回转换为本项目统一的页、文本块、表格和单元格结构；
5. 在服务实际返回的范围内保留页码、块/表格编号、行列、跨行跨列、坐标、置信度和证据位置。

统一后会生成独立的 `OCR结构化结果.xlsx`，固定包含：

- `OCR_文本`：逐页文本块及位置；
- `OCR_表格`：逐单元格明细及页码、行列和证据编号；
- `标准财务数据`：经过财务别名、期间和单位规则归一化的候选值；
- `识别问题`：OCR、结构转换和字段解析的内部状态，不作为用户输出；
- 表格索引及按页恢复的矩阵表：保留原表格行列形状，便于人工查看和继续匹配。

#### 4. OCR Excel 与两类业务 Excel

前端涉及的三类工作簿不是同一个文件，也不能相互替代：

| 工作簿 | 来源 | 是否由用户上传 | 主要用途 |
| --- | --- | --- | --- |
| `OCR结构化结果.xlsx` | 审计报告 PDF 的 OCR 结果 | 否，系统自动生成或复用 | 审计报告文本、表格、标准财务候选值和 OCR 问题 |
| 资产基础法/资产清查 Excel | 项目业务材料 | 可选 | 资产基础法、资产清查、长期资产和相关上报数据 |
| 收益法或市场法 Excel | 项目业务材料 | 可选 | 收益法或市场法测算过程及评估结论 |

用户上传的工作簿文件名可以不同；前端按材料角色传递，后端再根据工作表内容判断具体版式。`OCR结构化结果.xlsx` 是运行中间件，不需要也不应由用户重复上传。

#### 5. 从 OCR/Excel 到 Word

OCR 或业务 Excel 中的非空数字不会直接写入 Word。系统先根据以下信息定位候选值：

- 工作表内容和表格标题；
- 科目名称及别名；
- 历史年度、评估基准日和预测期间；
- 账面价值、账面净值、评估价值等列类型；
- 元、千元、万元等金额单位；
- 资产基础法、收益法和市场法特征。

通用语义识别是唯一的网页运行路径：先扫描上传 Excel/OCR 工作簿的表格标题、科目、列类型、期间和单位，再选择带来源单元格的候选值；不会读取开发机上的旧工作表坐标。出现多个同等合理的候选值、疑似尚未完成的全零评估列或证据冲突时，系统不根据相邻非零数字猜测，相关字段保持未解决。后续可启用受控 LLM 定位器：它只能在已扫描的候选单元格中选择并返回证据位置，不能生成金额。

每个接受的值都会在内部保留“来源文件 + 工作表/单元格”或“PDF/OCR 页码 + 表格/单元格证据”。黄色字段还要经过来源白名单：PDF OCR/XLSX 字段只能接收这一来源允许的数据，不得由企查查、百炼或其他材料跨来源冒填。

#### 6. 缺失、失败与最终输出

以下情况都不会阻止生成 `资产评估报告_待复核.docx`：

- 没有上传 PDF；
- 阿里云凭证缺失、无权限、超时、额度不足或解析失败；
- OCR/Excel 没有找到字段；
- 同一字段存在无法自动裁决的冲突。

缺失位置保留或写入黄色 `XXX`；冲突候选不自动选择。系统继续输出评估报告 Word，供人工直接复核。

OCR 所处的完整资产评估工作流为：

```text
节点1：人工输入 + 审计 PDF/OCR + 两类业务 Excel
→ 节点2：OCR/XLSX/API解析 + 百炼生成全部候选，人工选择
→ 节点3：优先复用节点2结果，缺失的企查查主体允许补查，再填充 Word
→ 节点4：评估报告 Word
```

这就是本仓库最初定义的可复用资产评估报告业务内核。用户、权限、任务队列、持久化、并发、监控和发布仍由 c2m 或其他生产外壳实现，不属于 Demo 业务内核。

## 本地运行示例

`--pdf` 是可选参数；不提供时跳过 OCR，仍会根据人工输入、工作簿/API/LLM 的现有结果生成待复核 Word。

```bash
uv run --python 3.11 python -m demo.run demo/projects/tongfu.yaml \
  --pdf "资产评估工作流/审计报告.pdf" \
  --template "templates/评估报告版式_v2.docx" \
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

前端提供三个有类型的可选上传入口：审计报告 PDF、资产基础法/资产清查 Excel、收益法/基础法 Excel。文件名可以不同，后端按扩展名、工作表标题、科目标签、列标题和金额单位识别角色。审计 PDF 会生成或复用 `OCR结构化结果.xlsx`；两个业务 Excel 直接进入语义取数；Word 模板始终使用后台版本，不接受用户上传。

人工基础信息严格对应最新图片中的九项：委托方全称/简称、委托类型（转让/收购/增资/减资单选）、评估主体全称/简称、评估对象（四选一）、评估方法（资产基础法/收益法/市场法多选且至少一个）、评估结论采用方法（单选）、评估报告编号流水号（非负整数）。

Word 模板由后端固定提供，不需要用户上传。参考评估报告 DOCX 仅用于开发期比对或后台项目配置，不属于前端输入；`OCR结构化结果.xlsx` 由系统根据 PDF 生成或复用，也不需要用户上传。两个业务 Excel 会覆盖本次任务对应的项目材料路径，生成结果不会修改原文件。常见资产基础法、收益法和市场法表会按工作表特征、科目标签、列标题及元/万元自动定位，项目特有字段才需要补充项目映射。

生成 Word 后只提供评估报告 Word；LLM 候选仅在第二节点暂停时用于页面内选择，不作为最终下载文件。工作流契约不合法、模板缺失或模板结构错误时停止生成；普通材料缺失、损坏、字段无结果或布局不匹配时继续生成待复核报告。不会生成三类 LLM 审核文件，也不会另存“最终候选” Word。

```bash
cd frontend
npm run build
```
