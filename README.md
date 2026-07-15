# Dify AI 市场研究 PPT 智能撰写工作流

这是一个基于 Dify Workflow 搭建的 AI 产品工作流 Demo。它可以将概念测试 Excel 数据转化为可追溯的研究证据包，进一步生成分析矩阵、研究洞察、Slide Spec、内容 QA，并最终渲染为可编辑的 PowerPoint 报告。

本仓库用于作品集展示，所有样例数据均已脱敏并合成，不包含真实客户名称、真实品牌名称、真实产品名称、原始研究数值或内部业务结论。

## Demo 展示材料

| 展示材料 | 文件 |
|---|---|
| 1-2 分钟运行录屏 | [assets/demo-recording.mp4](assets/demo-recording.mp4) |
| Dify 工作流全景预览 | [assets/dify-workflow-overview.jpeg](assets/dify-workflow-overview.jpeg) |
| 脱敏示例输入文件 | [sample_input/synthetic_concept_test_data.xlsx](sample_input/synthetic_concept_test_data.xlsx) |
| 最终生成 PPT | [sample_output/generated_demo_market_research_report.pptx](sample_output/generated_demo_market_research_report.pptx) |


![Dify 工作流全景预览](assets/dify-workflow-overview.jpeg)

<details>
<summary>查看完整纵向 Dify 工作流截图</summary>

![完整 Dify 工作流截图](assets/dify-workflow-full.jpeg)

</details>

## 项目亮点

- 将市场研究报告撰写拆解为 8 个环节：Excel 解析、证据标准化、分析矩阵构建、洞察生成、Slide Spec 生成、内容 QA、PPTX 渲染、下载交付。
- 搭建 23 节点 Dify 工作流，覆盖 HTTP 工具调用、Code 节点校验、LLM 结构化生成、异常分支、内容 QA、PPTX 渲染和质量评分。
- 从 Excel 交叉表中抽取并标准化 600 条研究证据，形成 72 组概念、受众和指标维度的分析比较。
- 通过 JSON Schema 风格的数据契约规范节点通信，保证证据、洞察、页面规划和 PPT 渲染过程可追溯、可校验、可复用。
- 将 LLM 能力边界控制在洞察组织和叙事生成上，把 Excel 解析、证据 ID 校验、降级策略和 PPTX 渲染交给确定性代码处理。
- 端到端生成 12 页可编辑 PPT 报告，覆盖从原始数据输入到交付物下载的完整闭环。

## 测试运行结果

本仓库中的样例结果来自一次本地集成测试。

测试输入的业务关注点：

```text
重点关注 Concept A 与 Concept B 的购买意愿、理解度、价值感和优化方向。
```

关键输出指标：

| 指标 | 结果 |
|---|---:|
| 标准化研究证据 | 600 条 |
| 分析比较组合 | 72 组 |
| 生成 PPT 页数 | 12 页 |
| Dify 工作流节点 | 23 个 |
| Code 节点结构校验 | 通过 |
| PPTX 渲染与下载 | 通过 |

## 工作流架构

```mermaid
flowchart LR
    A["Excel + Business Focus"] --> B["HTTP 工具：Excel 解析"]
    B --> C["Research Evidence Package"]
    C --> D["Code 节点：证据标准化"]
    D --> E["Code 节点：分析矩阵"]
    E --> F["LLM 节点：研究洞察"]
    F --> G["LLM 节点：Slide Spec"]
    G --> H["LLM / Code QA 节点"]
    H --> I["HTTP 工具：PPTX 渲染"]
    I --> J["可编辑 PPTX 报告"]
```

## 仓库结构

```text
.
|-- README.md
|-- assets/
|   |-- demo-recording.mp4
|   |-- dify-workflow-overview.jpeg
|   `-- dify-workflow-full.jpeg
|-- sample_input/
|   `-- synthetic_concept_test_data.xlsx
|-- sample_output/
|   `-- generated_demo_market_research_report.pptx
|-- workflow/
|   `-- generic_market_research_ppt_workflow.yml
|-- services/
|   `-- pptx_author_tool/
|       |-- server.py
|       |-- extractor.py
|       |-- renderer.py
|       |-- requirements.txt
|       |-- validate_workflow.py
|       |-- workflow_integration_test.py
|       `-- smoke_test.py
`-- reference_template.pptx
```

## 如何查看这个 Demo

1. 观看 `assets/demo-recording.mp4`，了解从上传 Excel 到下载 PPT 的完整运行过程。
2. 查看 `assets/dify-workflow-overview.jpeg`，快速理解工作流全貌；如需看节点细节，可展开 README 中的完整纵向截图。
3. 打开 `sample_input/synthetic_concept_test_data.xlsx`，查看脱敏后的概念测试输入结构。
4. 打开 `sample_output/generated_demo_market_research_report.pptx`，查看最终生成的可编辑 PPT 报告。
5. 查看 `workflow/generic_market_research_ppt_workflow.yml`，了解 23 节点 Dify Workflow 的具体配置。
6. 查看 `services/pptx_author_tool/`，了解 Excel 解析和 PPTX 渲染工具的实现方式。

## 本地验证方式

安装依赖：

```powershell
python -m pip install -r services/pptx_author_tool/requirements.txt
```

校验 Dify DSL 文件：

```powershell
python services/pptx_author_tool/validate_workflow.py workflow/generic_market_research_ppt_workflow.yml
```

运行本地集成测试：

```powershell
python services/pptx_author_tool/workflow_integration_test.py workflow/generic_market_research_ppt_workflow.yml
```

启动本地工具服务：

```powershell
python services/pptx_author_tool/server.py --host 0.0.0.0 --port 8077
```

健康检查地址：

```text
http://localhost:8077/health
```

工作流中的 HTTP 工具节点默认面向本地 Dify / Docker 环境：

```text
http://host.docker.internal:8077/extract-market-data
http://host.docker.internal:8077/render-pptx
```

## 能力体现

这个项目重点体现了以下 AI 产品经理能力：

- AI 工作流设计：将复杂报告撰写任务拆成可执行、可监控、可复用的节点链路。
- 数据驱动分析：从原始 Excel 表格中提取证据，并基于统一口径构建分析矩阵。
- AI 能力边界判断：把不稳定的生成任务交给 LLM，把确定性解析、校验、渲染交给工具代码。
- 结构化输出设计：通过 evidence package、analysis matrix、insight JSON、Slide Spec 等结构化对象连接各节点。
- 质量保障机制：加入证据 ID 校验、内容 QA、渲染审计、异常分支和 fallback，降低幻觉和格式错误风险。
- 端到端落地：从输入文件、Dify 编排、工具服务、LLM 生成到 PPTX 下载形成完整闭环。

## 上传说明

请上传这个脱敏后的展示仓库，不要上传任何原始内部文件。

推荐上传目录：

```text
market-research-dify-ppt-showcase-public/
```

不要上传以下类型文件：

```text
原始客户项目文件夹
原始工作流文件
原始市场研究 Excel
原始参考 PPT
内部测试输出 outputs/
日志文件 *.log
Office 临时锁文件 ~$*
```

## 隐私与脱敏说明

本仓库中的 Excel 和 PPTX 均为展示用途的脱敏样例。它们保留了工作流结构、数据处理方式和技术验证路径，但不保留真实业务数值、内部产品文案或客户特定结论。
