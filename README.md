# 招投标合规审查与 AI 比对 SaaS 系统

当前应用版本：**v2.0.0**

这是一个可部署到 Streamlit Community Cloud 的双 Word 文档审查应用。用户上传招标文件与投标文件后，系统调用 DeepSeek 的 OpenAI 兼容接口完成逐要求核查，并在内存中生成带风险高亮的 Excel 报告。

> 本项目使用 **DeepSeek API Key**，不需要 OpenAI 账户或 OpenAI API Key。依赖包 `openai` 仅作为调用 DeepSeek OpenAI 兼容接口的 Python SDK。密钥由 `st.secrets` 在服务端读取，不会显示在页面输入框中。

## v2 审查流程

v2 对短文档和长文档统一执行以下可追溯流程：

1. **逐要求建账**：将招标文件解析为有序文字来源块，逐块清点资格、实质性、废标、商务、技术、报价、合同、提交与评分要求，为每条原子要求分配稳定 ID。
2. **本地候选检索 + 重点全文补扫**：在投标文件文字来源块上建立本地索引，按每条要求检索可引用的候选证据；对强制/废标、评分和本地未命中要求，再让小组要求共享分批扫描全部可提取投标文字，避免按要求反复发送整份正文。候选未命中仍不等于全文不存在，补扫失败会显式转人工复核。
3. **小批核查**：按要求 ID 将要求与各自的候选证据组成小批次，请模型逐条判断，禁止合并、跳过、增加要求或引用候选包之外的出处；确定性结论的来源 ID 和原文摘录必须通过 Python 逐字核验。
4. **失败二分或人工占位**：批次过大、JSON 不完整或结构校验失败时，系统按来源块或要求 ID 二分后重试；单项仍失败、任务时限或模型调用预算接近上限时，会停止继续消耗 API，并为尚未可靠完成的来源/要求生成“待人工复核”占位记录，不接受部分返回冒充完整结果。
5. **Python 数量守恒**：由 Python 校验来源块清点、要求 ID、批次返回和最终结果。每条已建账要求必须且只能对应一条核查结果，评分要求还必须对应一条评分结果；数量或 ID 不一致时任务失败，不会静默生成缺项报告。

这里的“程序层不静默丢项”是指：来源块和已经进入要求台账的条目均接受完整性、唯一性与数量守恒校验，自动处理失败会显式留下人工复核记录。它不等于模型理解准确率为 100%，也不能保证模型从自然语言中识别出全部隐含要求。AI 仍可能误判、错误拆分或未识别语义，最终报告必须由专业人员对照原件复核。

## 功能与边界

- 提取 `.docx` 中可读取的正文、表格、嵌套表格、页眉页脚、文本框和脚注文字。
- 保留段落、表格行等来源标记，便于从 Excel 回查原文。
- 输出“缺陷核查记录”和“预估打分表”两张工作表。
- 对致命/废标、扣分/瑕疵、正常/符合等风险进行颜色标记。
- Excel 使用 `io.BytesIO` 在内存中生成，不依赖服务器绝对路径。
- 当前版本**只审查可提取文字，不执行 OCR，也不理解图片视觉内容**。扫描页、证书照片、印章、签字和图片中的文字必须人工复核。
- AI 结果是辅助意见，不构成法律意见、最终评分或投标决策。

## 项目文件

```text
.
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── tests/
    └── test_requirement_workflow.py
```

## 本地运行

推荐使用 Python 3.12。项目依赖已固定为当前验证版本，避免云端重新解析依赖时发生行为漂移。

Windows PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

开发环境安装 `pytest` 后可运行离线回归测试（测试不会调用真实模型 API；生产部署无需安装 `pytest`）：

```bash
python -m pip install pytest
python -m pytest -q
```

首次本地运行前，新建 `.streamlit/secrets.toml`：

```toml
DEEPSEEK_API_KEY = "你的 DeepSeek API Key"

# 可选；省略时使用 app.py 中的默认值
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-v4-flash"
```

`.streamlit/secrets.toml` 应由 `.gitignore` 排除。不要把真实密钥写进 `app.py`、README、截图、日志或 Git 提交。Base URL 和模型在页面只读显示，只能由管理员通过 Secrets 修改。

## 部署到 Streamlit Community Cloud

1. 在 GitHub 创建仓库，将 `app.py`、`requirements.txt`、`README.md`、`.gitignore` 和 `tests/` 放在仓库中。`.gitignore` 是正常的隐藏配置文件：GitHub 网页若不便新建或显示它，请在本地直接运行下面的 `git add .gitignore`，Git 会正常上传。不要提交真实招投标文件、生成的报告或 `.streamlit/secrets.toml`。
2. 提交并推送代码；提交前使用 `git status` 和 `git diff --cached` 核对暂存内容：

   ```bash
   git init
   git add app.py requirements.txt README.md .gitignore tests
   git commit -m "release: bid compliance reviewer v2.0.0"
   git branch -M main
   git remote add origin https://github.com/<你的用户名>/<仓库名>.git
   git push -u origin main
   ```

3. 打开 [Streamlit Community Cloud](https://share.streamlit.io/)，选择 **Create app**，指定 GitHub 仓库、`main` 分支和入口文件 `app.py`。
4. 在部署设置中选择 Python 3.12，并在 **Advanced settings → Secrets** 填入：

   ```toml
   DEEPSEEK_API_KEY = "你的 DeepSeek API Key"

   # 可选的服务端锁定配置
   DEEPSEEK_BASE_URL = "https://api.deepseek.com"
   DEEPSEEK_MODEL = "deepseek-v4-flash"
   ```

5. 保存并部署。启动后侧边栏应同时显示“DeepSeek API Key 已从 Cloud Secrets 安全加载”和“应用版本：v2.0.0”。
6. 更新 `app.py` 或依赖后，等待 Community Cloud 完成重新部署。若页面仍显示旧版本、依赖未刷新或应用状态异常，在应用管理菜单中执行 **Reboot app**，随后再次核对侧边栏必须为 **v2.0.0**。若仍不是 v2.0.0，请检查部署分支、入口文件及 GitHub 最新提交是否正确。

### Community Cloud 运行注意事项

- 当前实现是同步任务。点击“开始智能核查”后必须保持页面打开并维持网络连接，直到出现完成提示和下载按钮；刷新、关闭标签页、网络中断、应用重启或重新部署都可能中断本次任务，且不能断点恢复。
- 长文档会产生多次模型调用，处理时间和费用高于短文档。请根据页面日志等待，不要重复点击提交。
- 招投标文档通常包含敏感信息，并且部署者的 DeepSeek Key 由所有获准访问者共用。建议使用私有 GitHub 仓库，并将 Streamlit 应用设置为私有或仅允许受信任账号访问；不要将公共链接直接对外发布。
- Streamlit Community Cloud 的托管区域并不适合所有敏感招投标材料。若文档涉及商业秘密、境内存储或数据出境要求，应先完成组织的法务与安全评估；不满足条件时应改用境内或内网私有部署。
- Community Cloud 的会话状态和本地磁盘不是长期任务存储。生成报告后请及时下载；若需要关页后继续、任务队列或持久化审计，应改用外部任务服务与持久化存储。

## 安全与使用边界

- API Key 只通过 Streamlit Secrets 读取，不提供前端输入框，也不应写入 Excel、日志或仓库。
- 上传内容会发送到配置的模型服务商。部署前应确认服务商条款、数据存储位置以及所在组织的保密与数据合规要求。
- 应用会校验空文件、损坏 DOCX、异常 ZIP 结构和任务规模上限，但文件扩展名本身不是安全保证。
- 本地候选检索与重点全文补扫用于提高文字证据召回；即便补扫完成，模型仍可能未理解同义表达，因此未命中不能被宣传为数学意义上的“全文不存在”。
- 模型输出可能遗漏、误判或产生不准确结论。提交投标前必须由专业人员核对招标原文、投标原件、扫描材料及 Excel 中的人工复核项。
- 若出现 `model not found`，请由管理员在 Secrets 中将 `DEEPSEEK_MODEL` 改为当前 DeepSeek 账户实际可用的模型 ID，然后重新部署或 Reboot。

## 依赖版本

```text
streamlit==1.61.1
openai==2.53.0
python-docx==1.2.0
openpyxl==3.1.5
pandas==2.3.3
```

## 官方参考

- [DeepSeek API 文档](https://api-docs.deepseek.com/zh-cn/)
- [DeepSeek JSON Output](https://api-docs.deepseek.com/zh-cn/guides/json_mode/)
- [Streamlit Community Cloud 部署](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Streamlit Community Cloud 依赖](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies)
