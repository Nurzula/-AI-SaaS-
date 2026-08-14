# 招投标合规审查与 AI 比对 SaaS 系统

当前版本：**v3.1.0**

这是一个可部署到 Streamlit Community Cloud 的双 Word 文档文字核查应用。用户上传招标文件和投标文件后，系统通过 DeepSeek 的 OpenAI 兼容接口并发完成三类核查，在内存中生成带风险高亮的 Excel 报告。

默认模型为 `deepseek-v4-flash`。项目依赖 `openai` 仅用于调用兼容接口，不需要 OpenAI API Key。

## v3 核查流程

v3 不再采用 v2 的“逐块清点、全文补扫、失败递归二分”流程。

1. **完整文字解析**：按 Word 原始顺序提取正文、表格、页眉页脚、文本框和脚注，为每个段落或表格行保留稳定来源 ID。空表格单元格保留为 `(空)`。
2. **规则锚点**：Python 本地标记废标、报价、保证金、评分、金额、期限、技术和合同等重点来源。锚点只保留 ID、类别与来源定位，并要求结果行的招标来源与锚点来源一致；不会把每个普通段落扩写成模型台账。
3. **三路并发长上下文核查**：三个独立客户端同时读取两份完整可提取文字：
   - 资格、废标与形式核查；
   - 评分办法与预估得分；
   - 技术、商务与合同核查。
4. **有界调用**：正常路径仅 3 次 API 请求；每个通道最多重试一次，整单绝对上限 6 次。首次输出预算 24K tokens，唯一重试使用 32K tokens。禁止递归拆分来源块。
5. **双层结构与逐行验真**：顶层协议缺字段、错通道或空评分会触发该通道唯一一次重试；通过顶层门槛后，Python 再逐行验证来源 ID、逐字摘录、锚点—来源关联、风险依据和评分边界。坏行只转成“待人工复核”，不会让有效同级行作废。
6. **文字模式安全规则**：图片、扫描合同、公章、手写签名和证件照片不会被模型当成已确认事实；相关结论强制转人工复核。程序也会阻断“已废标”等过度确定表述，并按被引用局部来源优先核对项目编号，不会用其他页面的正确编号掩盖局部错误。
7. **经理可读 Excel**：内部证据仍完整保留，但下载报告会折叠重复、无明确问题和过程性记录；长原文与完整来源清单放入单元格批注。结果通过 `BytesIO` 生成两张工作表，不在服务器写绝对路径文件。

## 输出报告

- `缺陷核查记录`：序号、核查模块、检查要点、招标出处与要求、投标现状、问题、风险和建议。
- `预估打分表`：评分项、满分、评分规则、招标出处、当前预估得分和扣分说明。
- 致命/废标风险：红底白字加粗。
- 扣分/瑕疵：橙底黑字加粗。
- 正常/符合：绿底黑字。
- 待人工复核：仅风险单元格使用浅蓝中性色并加粗，正文保持正常字体，不使用斜体。
- 图片、签章和扫描证明材料合并为一条人工核验清单；重复报价结论、问题为“无”的待复核项和通道覆盖说明折叠为一条范围说明。
- 主表每项只展示简明要求、现状、问题与建议，完整原文和超出 3 个的来源 ID 可在单元格批注中查看。

## 功能边界

- 仅核查 `.docx` 中可提取的文字和表格内容。
- 不执行 OCR，不识别图片、扫描件、签字笔迹或印章真伪。
- 文档中的命令或提示词按不可信数据处理，不会改变系统角色。
- AI 结果是辅助审查意见，不构成法律意见、最终评分或投标决策。
- Streamlit Community Cloud 上的任务是同步任务；运行期间需保持页面连接，刷新、关页、重启或重新部署会中断当前任务。

## 项目文件

```text
.
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── tests/
    ├── test_requirement_workflow.py
    └── test_v3_performance_contracts.py
```

## 本地运行

推荐 Python 3.12。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

本地 Secrets 文件：`.streamlit/secrets.toml`

```toml
DEEPSEEK_API_KEY = "你的 DeepSeek API Key"
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-v4-flash"
```

`.streamlit/secrets.toml` 已被 `.gitignore` 排除。页面只显示服务端锁定后的 Base URL 和模型名，不显示密钥输入框。

## 测试

开发环境安装 `pytest` 后运行：

```powershell
python -m pytest -q
python -m py_compile app.py
```

离线测试不会调用真实模型 API，覆盖 DOCX 来源解析、三路并发、六次调用硬上限、输出长度唯一重试、顶层协议重试、锚点—来源绑定、逐行错误隔离、文字模式安全规则和 Excel 内存生成。

上线验收还应使用组织自己的 DeepSeek 账户对代表性真实文档连续执行端到端测试，记录总耗时、API 调用数、Token 用量和人工抽检结果。

## Streamlit Community Cloud 部署

1. 将 `app.py`、`requirements.txt`、`README.md`、`.gitignore` 和 `tests/` 推送到 GitHub 仓库。
2. 在 Streamlit Community Cloud 选择该仓库、部署分支和入口文件 `app.py`。
3. Python 版本选择 3.12。
4. 在 **App settings → Secrets** 配置：

   ```toml
   DEEPSEEK_API_KEY = "你的 DeepSeek API Key"
   DEEPSEEK_BASE_URL = "https://api.deepseek.com"
   DEEPSEEK_MODEL = "deepseek-v4-flash"
   ```

5. 保存后执行 Reboot。页面侧边栏应显示 `v3.1.0` 和 `deepseek-v4-flash`。
6. 上传两份 DOCX，日志应出现“三路长上下文并发核查”，且最终调用数不超过 6；如果仍出现“逐块清点”或递归二分日志，说明 Cloud 尚未部署到 v3。
