# 阿里云文档智能 OCR 接入设计

## 目标

将 Web 和 CLI 流程的默认 PDF OCR 从本地 `PPStructureV3` 改为阿里云文档智能“文档解析（大模型版）”，避免本机加载 PaddleOCR/PaddlePaddle 模型。现有 OCR 缓存、统一 OCR 数据契约、`OCR结构化结果.xlsx`、Word 填充和问题清单保持不变。

## 范围

- 新增阿里云文档智能适配器，调用异步文档解析接口并转换为现有页、文本块、表格和单元格结构。
- Web 后端与 CLI 根据环境变量选择 OCR 提供方。
- 默认提供方改为阿里云；本地 PaddleOCR 仅在显式配置时加载。
- 保留 PDF SHA-256 缓存优先逻辑，缓存命中时不调用任何 OCR 服务。
- API 无凭证、超时、额度不足或解析失败时不启动本地模型，继续生成待复核 Word，相关字段保留黄色占位符并进入问题清单。
- 不修改 Excel 语义识别、Word 映射、企查查或 LLM 逻辑。
- 不在前端增加新的必填项；VLM 增强由部署配置控制。

## 配置

本地 `.env` 使用：

```dotenv
APPRAISAL_OCR_PROVIDER=aliyun
ALIBABA_CLOUD_ACCESS_KEY_ID=
ALIBABA_CLOUD_ACCESS_KEY_SECRET=
APPRAISAL_OCR_VLM=false
APPRAISAL_OCR_TIMEOUT_SECONDS=900
```

`APPRAISAL_OCR_PROVIDER` 支持：

- `aliyun`：默认，调用阿里云文档解析（大模型版）。
- `paddle`：显式启用原本地 `PPStructureV3`。
- `none`：完全跳过 OCR。

凭证只在 API/CLI 入口注入适配器，不进入 `domain/`，不写入运行清单、工作流轨迹、日志或 Git。

## 数据流

1. 用户上传 PDF。
2. 按 PDF SHA-256 查找已有 `OCR结构化结果.xlsx`。
3. 缓存命中：直接复用，不创建云端客户端。
4. 未命中且提供方为 `aliyun`：
   - 通过 `SubmitDocParserJobAdvance` 上传本地 PDF；
   - 轮询 `QueryDocParserStatus`；
   - 分页调用 `GetDocParserResult` 获取全部布局块；
   - 将文本、页码、坐标、表格行列和单元格内容转换为现有普通字典契约。
5. 现有节点导出 `OCR结构化结果.xlsx` 并执行财务字段识别。
6. Word、字段审计清单和问题清单继续使用现有标准字段及证据链。

## 阿里云结果转换

- 普通文字、标题、页眉页脚等布局块转换为 `blocks`。
- `type=table` 的布局转换为 `tables`。
- 表格单元格使用 `ysc/yec/xsc/xec` 恢复起止行列；合并单元格保留跨度，并将其左上角作为标准行列位置。
- 单元格文本从嵌套 `layouts[].text` 合并。
- 页码和坐标原样保留；缺少置信度时写 `null`，不得伪造置信度。
- 不依赖 Markdown 反推表格；结构化 `layouts/cells` 是主要路径，Markdown 仅作为文本缺失时的兼容回退。

## 错误处理

- 未配置 AccessKey：记录“阿里云 OCR 凭证缺失”，OCR 节点失败但报告继续。
- 提交、轮询或结果分页失败：记录阶段、任务 ID（如已有）和阿里云错误码，不记录密钥。
- 超过配置超时：停止轮询，报告继续。
- 返回成功但无文本/表格：记录空结果，报告继续。
- 不自动回退本地 PaddleOCR，防止因云端异常再次占满内存。

## 依赖边界

阿里云 SDK 放入独立的轻量可选依赖组，不放入 `domain/`。适配器只返回可 JSON 序列化的普通字典，符合当前业务内核可复用规范。本地 `ocr` 依赖组继续保留，仅供显式 `paddle` 模式使用。

## 测试与验收

- 使用固定 JSON 样例测试文字块、表格、合并单元格、页码和坐标转换。
- 使用假客户端测试提交、轮询、分页、超时和错误码，不调用真实服务。
- 测试 Web/CLI 的提供方选择，确保 `aliyun` 模式不会导入或创建 PaddleOCR。
- 测试缓存命中时云端客户端零调用。
- 测试 API 失败时仍生成待复核 Word 和问题清单。
- 使用一页无客户数据的合成 PDF 做可选真实连通性测试；未经额外确认，不上传现有客户审计报告。
- 全量后端测试、前端测试和前端构建必须通过。

## 非目标

- 不让 LLM 代替 OCR 结构化取数。
- 不把用户 AccessKey 提交到仓库。
- 不删除本地 PaddleOCR 实现。
- 不修改已经验证的 Excel 到 Word 语义匹配规则。
