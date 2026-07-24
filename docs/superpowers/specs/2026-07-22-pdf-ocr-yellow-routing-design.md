# PDF OCR 到 Word 黄色字段路由设计

## 1. 目标

实现一条可替换输入文件的资产评估报告生成链路：扫描 PDF 经 PaddleOCR 识别并还原为统一 XLSX，随后严格按照 Word 模板中的黄色标注选择数据来源，最后在不覆盖模板的前提下生成独立 Word 和字段审计清单。

本设计只把模板中实际存在的 20 处黄色标注作为黄色字段路由依据，不新增黄色字段，不扩大字段含义，不在来源类别之间自行回退。模板中既有的 `XX/XXX/XXXX` 普通占位符继续使用现有映射规则，不受本次黄色字段路由调整影响。

## 2. 复用边界和目录责任

本 pipeline 遵循“业务内核可复用、生产外壳由 c2m 实现”的边界。Demo 用来发现、验证和表达业务规则，不作为 c2m 的运行时依赖。c2m 后续只移植数据契约、纯业务规则、Prompt、配置、样例和验收测试，自行实现用户、权限、异步任务、存储、监控、重试和交互。

目录责任如下：

```text
demo/
├── README.md
├── workflow.yaml
├── schemas.py
├── domain/                  # 可移植纯业务逻辑
│   ├── ocr_normalization.py
│   ├── financial_matching.py
│   ├── yellow_routing.py
│   └── field_validation.py
├── adapters/                # Demo/基础设施适配器
│   ├── paddle_ocr.py
│   ├── bailian_glm.py
│   ├── qichacha.py
│   ├── ocr_workbook.py
│   └── word.py
├── prompts/                 # 单节点、带版本号的 Prompt
├── fixtures/                # 脱敏正常和异常样例
├── expected/                # 期望结构和关键断言
├── tests/
├── data_manifest.yaml
└── CHANGELOG.md
```

强制边界：

1. `domain/` 只接收普通 Python/Pydantic 数据，只返回可 JSON 序列化结果；
2. `domain/` 不读取文件、环境变量或网络，不导入 PaddleOCR、OpenAI SDK、FastAPI、数据库模型或页面 session；
3. OCR、GLM、企查查、文件系统和 Word/XLSX 读写均通过参数或接口注入；
4. 绝对路径、API Key、用户信息和生产地址不得写入业务函数或配置；
5. CLI 只负责组装依赖和运行节点，不承载业务判断；
6. 每个 Prompt 单独存放、带版本号，并由 Pydantic 输出模型校验；
7. 返回结构保留规则版本、Prompt 版本、数据版本和证据来源；
8. 每个节点可以脱离 CLI 单独调用和测试。

## 3. 黄色字段的唯一分类

系统启动时必须重新扫描模板，并验证黄色标注恰好覆盖以下 20 个位置。配置缺项、多项或出现未知黄色标注时停止运行。

### 3.1 GLM 生成：7 项

以下字段只调用百炼 `glm-5.2` 生成：

1. `company_profile_section`：被评估单位概述
2. `industry_overview`：所处行业及行业介绍
3. `business_and_segments`：业务内容及细分市场
4. `main_products`：主要产品
5. `customers_suppliers`：主要客户、供应商
6. `profit_model_swot`：盈利模式、SWOT 分析
7. `comparable_list`：对标上市公司列表及相似性

GLM 可以读取本次运行产生的 OCR 文本、OCR 表格、统一 XLSX 和节点基础信息，但只允许返回这 7 个字段。系统拒绝模型返回的其他字段。模型输出必须符合 JSON 结构，并为每个字段记录所使用的证据编号；没有证据时返回空字符串，不允许生成金额或修改规则字段。

百炼配置：

- Base URL：`https://dashscope.aliyuncs.com/compatible-mode/v1`
- Model：`glm-5.2`
- API Key：仅从 `DASHSCOPE_API_KEY` 环境变量读取
- 结构化输出时关闭思考模式

### 3.2 企查查 API：5 项

以下字段只调用企查查适配器：

1. `commissioning_party_profile`：委托人基本情况
2. `ownership_history`：被评估单位股权结构及历史沿革
3. `ownership_at_valuation_date`：评估基准日股东及股权情况
4. `unrecorded_intangibles`：账外无形资产
5. `software_copyrights`：软件著作权

系统不得用 GLM、PDF OCR 或现有参考报告代替企查查结果。未配置企查查或接口无数据时，对应字段留空并进入问题清单。

### 3.3 PDF OCR / XLSX 规则读取：6 项

以下字段只从 PDF OCR 结果或其结构化 XLSX 读取：

1. `historical_balance_sheet_table`：近年资产负债状况表
2. `historical_income_statement_table`：近年经营状况/利润表
3. `tax_rates`：主要税种及税率
4. `valuation_scope`：评估对象和评估范围
5. `major_long_term_assets`：主要长期资产账面记录
6. `asset_approach_result_section`：资产基础法评估结果段落

金额、日期、税率、数量和表格内容由确定性规则处理。允许根据字段别名、期间、单位和表格上下文匹配，不允许使用固定单元格位置作为唯一依据，也不允许 GLM 为这些字段生成内容。

### 3.4 节点输入：2 项

以下字段只使用运行参数或人工节点输入：

1. `selected_valuation_method`：评估方法
2. `valuation_purpose_inputs`：评估目的输入参数

## 4. 端到端数据流

```text
扫描 PDF
  -> PaddleOCR PP-StructureV3
  -> OCR 原始 JSON / Markdown
  -> 统一 OCR XLSX
  -> 黄色字段严格路由
       -> GLM 7 项
       -> 企查查 API 5 项
       -> OCR/XLSX 规则 6 项
       -> 节点输入 2 项
  -> 普通占位符沿用现有映射
  -> 复制 Word 模板并填充
  -> 新 Word + OCR XLSX + 字段审计 XLSX + JSON 清单
```

## 5. 工作流节点及契约

`workflow.yaml` 必须声明以下节点、顺序、输入模型、输出模型和失败策略：

| 节点 | 输入 | 输出 | 失败策略 |
|---|---|---|---|
| `inventory_template` | Word 模板路径、黄色路由配置 | 模板位置清单 | 黄色位置变化即终止 |
| `ocr_pdf` | PDF 路径、OCR 选项 | OCR 页面结果 | 任一财务相关页失败即终止 |
| `export_ocr_workbook` | OCR 页面结果 | OCR XLSX 路径、证据索引 | 导出失败即终止 |
| `normalize_evidence` | OCR 页面结果、字段别名和单位规则 | 标准文本、表格和财务候选 | 保留冲突，不静默择一 |
| `resolve_rule_fields` | 标准证据、OCR/XLSX 六字段配置 | 六个规则字段和证据 | 财务必填缺失即终止 |
| `fetch_qichacha_fields` | 公司主体、基准日、注入 API | 五个 API 字段和证据 | 未配置或失败则留空 |
| `generate_glm_fields` | 七字段列表、证据包、注入 LLM | 七个结构化叙述字段 | 失败或越权字段被拒绝 |
| `resolve_node_inputs` | 节点输入 | 两个节点字段 | 必填输入缺失即终止 |
| `validate_source_routes` | 全部字段和来源 | 通过/拒绝及问题清单 | 发现跨类别回退即终止 |
| `fill_word` | 模板、副本路径、字段结果 | 独立 Word | 模板哈希变化即终止 |
| `export_audit` | 字段、证据、问题和版本 | 审计 XLSX/JSON | 导出失败即终止 |
| `human_review` | Word、OCR XLSX、审计清单 | 人工通过/退回 | Demo 的唯一人工检查点 |

每个节点的输入输出在 `schemas.py` 中声明中文字段说明、是否必填和示例。运行器只按照 `workflow.yaml` 编排，不在节点之间传递页面 session、数据库对象或不可序列化客户端。

## 6. PDF OCR 与统一 XLSX

### 6.1 OCR 接口

新增 `DocumentOcrEngine` 接口，默认实现为本地 PaddleOCR 3.x `PPStructureV3`。输入为 PDF 路径，输出为与具体 OCR SDK 解耦的页面、文本块和表格单元格模型。

模型至少包含：

- PDF 页码和总页数
- 块类型、原始文字、置信度
- 坐标框
- 表格编号、行号、列号和单元格文字
- OCR 原始结果定位信息

### 6.2 XLSX 结构

每次运行单独生成 `OCR结构化结果.xlsx`，包含：

- `OCR_文本`：页码、块编号、类型、文字、置信度、坐标
- `OCR_表格`：页码、表格编号、行、列、文字、置信度、坐标
- `标准财务数据`：标准字段、期间、数值、单位、证据编号、匹配方法
- `识别问题`：缺失、冲突、低置信度和无法归类的记录

原始 OCR 文本不得被覆盖或静默修正。标准化结果必须保留到原 PDF 页码和 OCR 单元格的反向定位。

## 7. 跨版式财务匹配

不同 PDF 可以改变页码、表名、列顺序、表头写法和单位。匹配器依次使用：

1. 文本规范化：空格、全半角、常见 OCR 字符和括号处理；
2. 字段别名字典，例如“资产总计/资产合计/总资产”；
3. 期间识别，例如年度、期末、上年末和基准日；
4. 单位识别与换算，例如元、万元、千元和百分比；
5. 同表行列和上下文约束；
6. 跨表一致性检查与计算关系校验。

规则只能选择 OCR 中实际存在的值或执行配置声明的计算。财务必填字段缺失时终止 Word 生成，并在问题清单中列出缺失字段及已检查的候选证据。

## 8. Word 填充规则

1. 原 Word 始终只读，输出写入新的运行目录。
2. 20 个黄色位置只能使用其配置的唯一来源类别。
3. 有值时用结果替换黄色说明并移除黄色高亮。
4. 无值时清除黄色说明并留空，不写“待人工补充”。
5. 不新增模板中不存在的段落、标题或说明。
6. 普通占位符继续使用既有配置和系统计算。
7. 填充后扫描 OOXML，确保不存在黄色高亮和遗留提示语。

## 9. 审计与输出

每次运行输出：

- `资产评估报告_待复核.docx`
- `OCR结构化结果.xlsx`
- `字段审计清单.xlsx`
- `normalized_fields.json`
- `issues.json`
- `run_manifest.json`

字段审计至少记录：字段键、黄色原文、唯一来源类别、最终值、PDF 页码、OCR 块/表格单元格、API/模型版本、Prompt 版本和运行时间。

## 10. 失败处理

- OCR 单页失败：记录页码并终止财务提取，不跳过后继续生成报告。
- OCR 表格冲突：保留所有候选并标记冲突；财务必填项不自动择一。
- GLM 调用失败或结构不合法：7 个 GLM 字段留空，其他通道继续。
- 企查查未配置或失败：5 个 API 字段留空，其他通道继续。
- GLM 返回黄色范围外字段：拒绝该字段并记录问题。
- 来源类别不符：拒绝填入，不执行跨类别回退。
- 模板黄色位置发生变化：在生成 Word 前停止，要求更新配置。

## 11. 测试与验收

实现必须按测试驱动完成，至少覆盖：

1. 模板恰好识别出 20 个黄色位置；
2. 路由数量严格为 GLM 7、企查查 5、OCR/XLSX 6、节点输入 2；
3. GLM 只会被要求生成固定的 7 个字段；
4. GLM 返回金额字段或额外字段时被拒绝；
5. 企查查字段不会由 GLM 或 OCR 结果补齐；
6. OCR/XLSX 字段不会由 GLM 或企查查结果补齐；
7. 相同财务数据在不同页码、列顺序、表头别名和单位下映射到相同标准字段；
8. PDF OCR fixture 可以生成统一 XLSX；
9. 财务必填字段缺失时不生成 Word；
10. 原模板哈希保持不变；
11. 生成 Word 不含黄色高亮、旧提示或“待人工补充”；
12. 现有通富材料的历史财务表和关键评估金额继续与参考报告一致；
13. 最终 Word 完整渲染并逐页检查无截断、重叠或表格溢出；
14. OCR XLSX 经数值、来源定位和视觉检查后才能交付。

此外必须提供至少 10 个脱敏代表性样例，覆盖：正常审计报告、表头别名、列顺序变化、单位变化、跨页表格、OCR 字符错误、缺失财务必填项、GLM 越权输出、企查查无结果和模板黄色位置变化。非确定性 LLM 测试只断言结构、必备事实、证据编号和业务约束，不断言全文完全相同。

每次规则或 Prompt 变化必须同步更新 `CHANGELOG.md`、`fixtures/`、`expected/` 和相关测试。`data_manifest.yaml` 记录每份样例的数据来源类型、版本、更新时间和缺失处理方式。

Demo 验收只代表业务方案成立，不代表已经满足生产上线所需的用户隔离、权限、并发、异步、幂等、重试、限流、监控或恢复要求。

## 12. 非目标

- 不自动推测黄色以外的新业务字段；
- 不用参考报告正文代替企查查、OCR 或 GLM 的指定来源；
- 不允许模型修改或编造财务数值；
- 不把密钥写入仓库、日志、输出文档或审计清单；
- 不覆盖原始 PDF、原始 XLSX 或 Word 模板。
